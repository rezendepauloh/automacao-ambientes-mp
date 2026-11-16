using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo de leitura (somente consulta) dos chamados OTRS e CitSmart
/// diretamente no PostgreSQL da Bancada Docker, sem credenciais hardcoded.
/// </summary>
public class BancadaChamadosService
{
    public static bool TryBuildConnectionString(
        string? host,
        int port,
        string? database,
        string? username,
        string? password,
        out string connectionString,
        out string error)
    {
        connectionString = string.Empty;
        error = string.Empty;

        string cleanHost = string.IsNullOrWhiteSpace(host) ? "127.0.0.1" : host.Trim();
        int cleanPort = port > 0 && port <= 65535 ? port : 5432;
        string cleanDb = database?.Trim() ?? string.Empty;
        string cleanUser = username?.Trim() ?? string.Empty;
        string cleanPass = password ?? string.Empty;

        if (!InputValidator.IsValidHostName(cleanHost))
        {
            error = "BANCADA_DB_HOST inválido (fail-closed).";
            return false;
        }

        if (string.IsNullOrWhiteSpace(cleanDb) || string.IsNullOrWhiteSpace(cleanUser) || string.IsNullOrWhiteSpace(cleanPass))
        {
            error = "Configuração do PostgreSQL incompleta no .env (defina BANCADA_DB_NAME, BANCADA_DB_USER e BANCADA_DB_PASSWORD em Configurações).";
            return false;
        }

        var builder = new NpgsqlConnectionStringBuilder
        {
            Host = cleanHost,
            Port = cleanPort,
            Database = cleanDb,
            Username = cleanUser,
            Password = cleanPass,
            Timeout = 5,
            CommandTimeout = 10
        };

        connectionString = builder.ConnectionString;
        return true;
    }

    public static bool TryBuildConnectionStringFromConfig(out string connectionString, out string error)
    {
        return TryBuildConnectionString(
            ConfigService.BancadaDbHost,
            ConfigService.BancadaDbPort,
            ConfigService.BancadaDbName,
            ConfigService.BancadaDbUser,
            ConfigService.BancadaDbPassword,
            out connectionString,
            out error);
    }

    public async Task<List<string>> GetDistinctLocalidadesAsync()
    {
        var list = new List<string>();
        if (!TryBuildConnectionStringFromConfig(out string connStr, out _))
            return list;

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();
            const string sql = """
                SELECT DISTINCT TRIM(localidade_fisica) AS loc
                FROM chamados
                WHERE TRIM(COALESCE(localidade_fisica, '')) <> ''
                ORDER BY loc;
                """;
            await using var cmd = new NpgsqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                string val = reader.GetString(0);
                if (!string.IsNullOrWhiteSpace(val))
                    list.Add(val);
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao listar localidades físicas: {ex.Message}", nameof(BancadaChamadosService));
        }

        return list;
    }

    public async Task<(List<string> Tags, List<string> Localidades)> GetAvailableFiltersAsync(
        string? search = null,
        string? baseFilter = null,
        string? statusFilter = "Aberto")
    {
        var tags = new List<string>();
        var locs = new List<string>();
        if (!TryBuildConnectionStringFromConfig(out string connStr, out _))
            return (tags, locs);

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            string cleanSearch = string.IsNullOrWhiteSpace(search) ? string.Empty : $"%{search.Trim().ToLowerInvariant()}%";
            string cleanBase = string.IsNullOrWhiteSpace(baseFilter) || baseFilter.Equals("Todas as Bases", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : $"%{baseFilter.Trim().ToLowerInvariant()}%";
            string cleanStatus = string.IsNullOrWhiteSpace(statusFilter) || statusFilter.Equals("Todos", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : statusFilter.Trim().ToLowerInvariant();

            const string whereClause = """
                WHERE (@baseFilter = '' OR LOWER(COALESCE(base, '')) LIKE @baseFilter)
                  AND (@statusFilter = '' OR LOWER(COALESCE(status, '')) = @statusFilter)
                  AND (
                        @search = ''
                        OR LOWER(COALESCE(id, '')) LIKE @search
                        OR LOWER(COALESCE(titulo, '')) LIKE @search
                        OR LOWER(COALESCE(usuario, '')) LIKE @search
                        OR LOWER(COALESCE(unidade, '')) LIKE @search
                        OR LOWER(COALESCE(tag, '')) LIKE @search
                        OR LOWER(COALESCE(hostname, '')) LIKE @search
                      )
                """;

            string tagsSql = $"SELECT DISTINCT TRIM(tag) FROM chamados {whereClause} AND TRIM(COALESCE(tag, '')) <> '' ORDER BY 1;";
            await using (var cmd = new NpgsqlCommand(tagsSql, conn))
            {
                cmd.Parameters.AddWithValue("@baseFilter", cleanBase);
                cmd.Parameters.AddWithValue("@statusFilter", cleanStatus);
                cmd.Parameters.AddWithValue("@search", cleanSearch);
                await using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                    tags.Add(r.GetString(0));
            }

            string locsSql = $"SELECT DISTINCT TRIM(localidade_fisica) FROM chamados {whereClause} AND TRIM(COALESCE(localidade_fisica, '')) <> '' ORDER BY 1;";
            await using (var cmd = new NpgsqlCommand(locsSql, conn))
            {
                cmd.Parameters.AddWithValue("@baseFilter", cleanBase);
                cmd.Parameters.AddWithValue("@statusFilter", cleanStatus);
                cmd.Parameters.AddWithValue("@search", cleanSearch);
                await using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                    locs.Add(r.GetString(0));
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao atualizar filtros dinâmicos: {ex.Message}", nameof(BancadaChamadosService));
        }

        return (tags, locs);
    }

    public async Task<(bool Success, List<ChamadoItem> Items, ChamadosSummary Summary, string Message)> QueryChamadosAsync(
        string? search = null,
        string? baseFilter = null,
        string? statusFilter = "Aberto",
        int limit = 150,
        string? tagFilter = null,
        List<string>? localidadesFilter = null)
    {
        if (!TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
            return (false, [], new ChamadosSummary(), error);
        }

        var items = new List<ChamadoItem>();
        var summary = new ChamadosSummary();

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            // 1. Resumo agregado (KPIs)
            const string summarySql = """
                SELECT
                    COUNT(*)::int AS total,
                    COUNT(*) FILTER (WHERE LOWER(COALESCE(status, '')) = 'aberto')::int AS abertos,
                    COUNT(*) FILTER (WHERE LOWER(COALESCE(base, '')) LIKE '%otrs%')::int AS otrs_cnt,
                    COUNT(*) FILTER (WHERE LOWER(COALESCE(base, '')) LIKE '%citsmart%')::int AS citsmart_cnt
                FROM chamados;
                """;

            await using (var sumCmd = new NpgsqlCommand(summarySql, conn))
            await using (var sumReader = await sumCmd.ExecuteReaderAsync())
            {
                if (await sumReader.ReadAsync())
                {
                    summary.Total = sumReader.GetInt32(0);
                    summary.Abertos = sumReader.GetInt32(1);
                    summary.OtrsCount = sumReader.GetInt32(2);
                    summary.CitSmartCount = sumReader.GetInt32(3);
                }
            }

            // 2. Consulta parametrizada de chamados
            const string querySql = """
                SELECT
                    COALESCE(id, ''),
                    COALESCE(data_criacao, ''),
                    COALESCE(titulo, ''),
                    COALESCE(cidade_predio, ''),
                    COALESCE(unidade, ''),
                    COALESCE(localidade_fisica, ''),
                    COALESCE(usuario, ''),
                    COALESCE(id_cliente, ''),
                    COALESCE(descricao, ''),
                    COALESCE(tag, ''),
                    COALESCE(ip_origem, ''),
                    COALESCE(status, 'Aberto'),
                    COALESCE(data_atualizacao, ''),
                    COALESCE(base, ''),
                    COALESCE(andamento, ''),
                    COALESCE(link, ''),
                    COALESCE(hostname, '')
                FROM chamados
                WHERE (@baseFilter = '' OR LOWER(COALESCE(base, '')) LIKE @baseFilter)
                  AND (@statusFilter = '' OR LOWER(COALESCE(status, '')) = @statusFilter)
                  AND (@tagFilter = '' OR LOWER(COALESCE(tag, '')) = @tagFilter)
                  AND (@hasLocalidades = false OR LOWER(TRIM(COALESCE(localidade_fisica, ''))) = ANY(@localidades))
                  AND (
                        @search = ''
                        OR LOWER(COALESCE(id, '')) LIKE @search
                        OR LOWER(COALESCE(titulo, '')) LIKE @search
                        OR LOWER(COALESCE(usuario, '')) LIKE @search
                        OR LOWER(COALESCE(unidade, '')) LIKE @search
                        OR LOWER(COALESCE(tag, '')) LIKE @search
                        OR LOWER(COALESCE(hostname, '')) LIKE @search
                      )
                ORDER BY COALESCE(NULLIF(data_atualizacao, ''), data_criacao) DESC
                LIMIT @limit;
                """;

            string cleanSearch = string.IsNullOrWhiteSpace(search) ? string.Empty : $"%{search.Trim().ToLowerInvariant()}%";
            string cleanBase = string.IsNullOrWhiteSpace(baseFilter) || baseFilter.Equals("Todas as Bases", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : $"%{baseFilter.Trim().ToLowerInvariant()}%";
            string cleanStatus = string.IsNullOrWhiteSpace(statusFilter) || statusFilter.Equals("Todos", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : statusFilter.Trim().ToLowerInvariant();
            string cleanTag = string.IsNullOrWhiteSpace(tagFilter) || tagFilter.Equals("Todas as Tags", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : tagFilter.Trim().ToLowerInvariant();
            string[] cleanLocalidades = localidadesFilter?
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Select(l => l.Trim().ToLowerInvariant())
                .Distinct()
                .ToArray() ?? [];

            await using var cmd = new NpgsqlCommand(querySql, conn);
            cmd.Parameters.AddWithValue("@baseFilter", cleanBase);
            cmd.Parameters.AddWithValue("@statusFilter", cleanStatus);
            cmd.Parameters.AddWithValue("@tagFilter", cleanTag);
            cmd.Parameters.AddWithValue("@hasLocalidades", cleanLocalidades.Length > 0);
            cmd.Parameters.AddWithValue("@localidades", cleanLocalidades);
            cmd.Parameters.AddWithValue("@search", cleanSearch);
            cmd.Parameters.AddWithValue("@limit", Math.Clamp(limit, 1, 500));

            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                items.Add(new ChamadoItem
                {
                    Id = reader.GetString(0),
                    DataCriacao = reader.GetString(1),
                    Titulo = reader.GetString(2),
                    CidadePredio = reader.GetString(3),
                    Unidade = reader.GetString(4),
                    LocalidadeFisica = reader.GetString(5),
                    Usuario = reader.GetString(6),
                    IdCliente = reader.GetString(7),
                    Descricao = reader.GetString(8),
                    Tag = reader.GetString(9),
                    IpOrigem = reader.GetString(10),
                    Status = reader.GetString(11),
                    DataAtualizacao = reader.GetString(12),
                    Base = reader.GetString(13),
                    Andamento = reader.GetString(14),
                    Link = reader.GetString(15),
                    Hostname = reader.GetString(16)
                });
            }

            return (true, items, summary, $"{items.Count} chamado(s) carregado(s) do PostgreSQL da Bancada.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao consultar PostgreSQL da Bancada: {ex.Message}", nameof(BancadaChamadosService));
            return (false, items, summary, $"Erro ao conectar no PostgreSQL ({ConfigService.BancadaDbHost}:{ConfigService.BancadaDbPort}): {ex.Message}");
        }
    }

    public async Task<List<ChamadoComentario>> GetComentariosChamadoAsync(string chamadoId)
    {
        var list = new List<ChamadoComentario>();
        if (string.IsNullOrWhiteSpace(chamadoId) || !TryBuildConnectionStringFromConfig(out string connStr, out _))
            return list;

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            const string sql = """
                SELECT COALESCE(data, ''), COALESCE(autor, ''), COALESCE(texto, '')
                FROM comentarios
                WHERE chamado_id = @chamadoId
                ORDER BY id ASC;
                """;

            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@chamadoId", chamadoId.Trim());
            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                list.Add(new ChamadoComentario
                {
                    Data = reader.GetString(0),
                    Autor = reader.GetString(1),
                    Texto = reader.GetString(2)
                });
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao carregar comentários do chamado #{chamadoId}: {ex.Message}", nameof(BancadaChamadosService));
        }

        return list;
    }

    public async Task<Dictionary<string, List<ChamadoComentario>>> GetComentariosPorChamadosAsync(IEnumerable<string> chamadoIds)
    {
        var result = new Dictionary<string, List<ChamadoComentario>>(StringComparer.OrdinalIgnoreCase);
        var ids = chamadoIds?
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray() ?? [];

        if (ids.Length == 0 || !TryBuildConnectionStringFromConfig(out string connStr, out _))
            return result;

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            const string sql = """
                SELECT COALESCE(chamado_id, ''), COALESCE(data, ''), COALESCE(autor, ''), COALESCE(texto, '')
                FROM comentarios
                WHERE chamado_id = ANY(@ids)
                ORDER BY id ASC;
                """;

            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@ids", ids);
            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                string chamadoId = reader.GetString(0).Trim();
                if (!result.TryGetValue(chamadoId, out var list))
                {
                    list = [];
                    result[chamadoId] = list;
                }
                list.Add(new ChamadoComentario
                {
                    Data = reader.GetString(1),
                    Autor = reader.GetString(2),
                    Texto = reader.GetString(3)
                });
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao carregar comentários em lote: {ex.Message}", nameof(BancadaChamadosService));
        }

        return result;
    }

    public async Task<(bool Success, string Message)> UpdateChamadoAsync(
        string id,
        string status,
        string hostname,
        string tag,
        string cidadePredio,
        string unidade,
        string localidadeFisica,
        string descricao,
        string andamento)
    {
        if (string.IsNullOrWhiteSpace(id))
            return (false, "ID do chamado inválido.");

        if (!TryBuildConnectionStringFromConfig(out string connStr, out string error))
            return (false, error);

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            const string sql = """
                UPDATE chamados
                SET status = @status,
                    hostname = @hostname,
                    tag = @tag,
                    cidade_predio = @cidadePredio,
                    unidade = @unidade,
                    localidade_fisica = @localidadeFisica,
                    descricao = @descricao,
                    andamento = @andamento,
                    data_atualizacao = @dataAtualizacao,
                    tag_manual = 1,
                    dados_manuais = 1
                WHERE id = @id;
                """;

            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", id.Trim());
            cmd.Parameters.AddWithValue("@status", status?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@hostname", hostname?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@tag", tag?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@cidadePredio", cidadePredio?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@unidade", unidade?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@localidadeFisica", localidadeFisica?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@descricao", descricao?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@andamento", andamento?.Trim() ?? string.Empty);
            cmd.Parameters.AddWithValue("@dataAtualizacao", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            int rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return (false, $"Chamado #{id} não encontrado no PostgreSQL.");

            return (true, $"Chamado #{id} atualizado no PostgreSQL com sucesso.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao atualizar chamado #{id}: {ex.Message}", nameof(BancadaChamadosService));
            return (false, $"Erro ao salvar no PostgreSQL: {ex.Message}");
        }
    }
}
