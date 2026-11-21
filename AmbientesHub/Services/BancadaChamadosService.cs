using System.Diagnostics;
using System.Text.Json.Serialization;
using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo de leitura e edição dos chamados OTRS e CitSmart.
/// Prioriza a API REST segura da Bancada (WSL/RedHat) com fallback opcional para PostgreSQL direto.
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
        var sw = Stopwatch.StartNew();
        if (BancadaApiClient.IsApiConfigured())
        {
            var (ok, res, _, err) = await BancadaApiClient.GetAsync<ApiLocalidadesResponse>("/api/v1/chamados/localidades");
            if (ok && res?.Localidades != null)
            {
                AppLoggerService.Info($"GetDistinctLocalidadesAsync (API REST): {res.Localidades.Count} localidade(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
                return res.Localidades;
            }
            AppLoggerService.Warning($"Falha ao listar localidades via API ({AppLoggerService.FormatElapsed(sw)}): {err}", nameof(BancadaChamadosService));
            return [];
        }

        var list = new List<string>();
        if (!TryBuildConnectionStringFromConfig(out string connStr, out string cfgErr))
        {
            AppLoggerService.Warning($"GetDistinctLocalidadesAsync abortado (PostgreSQL): {cfgErr}", nameof(BancadaChamadosService));
            return list;
        }

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
            AppLoggerService.Info($"GetDistinctLocalidadesAsync (PostgreSQL {ConfigService.BancadaDbHost}:{ConfigService.BancadaDbPort}/{ConfigService.BancadaDbName}): {list.Count} localidade(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaChamadosService), "GetDistinctLocalidadesAsync (PostgreSQL)", ex, sw.Elapsed, LogLevel.WARNING);
        }

        return list;
    }

    public async Task<(List<string> Tags, List<string> Localidades)> GetAvailableFiltersAsync(
        string? search = null,
        string? baseFilter = null,
        string? statusFilter = "Aberto",
        string? tagFilter = null,
        List<string>? localidadesFilter = null)
    {
        var sw = Stopwatch.StartNew();
        if (BancadaApiClient.IsApiConfigured())
        {
            string qs = BuildChamadosQueryString(search, baseFilter, statusFilter, null, tagFilter, localidadesFilter);
            var (ok, res, _, err) = await BancadaApiClient.GetAsync<ApiFiltersResponse>($"/api/v1/chamados/filters{qs}");
            if (ok && res != null)
            {
                int tCnt = res.Tags?.Count ?? 0;
                int lCnt = res.Localidades?.Count ?? 0;
                AppLoggerService.Info($"GetAvailableFiltersAsync (API REST): tags={tCnt}, localidades={lCnt} em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
                return (res.Tags ?? [], res.Localidades ?? []);
            }
            AppLoggerService.Warning($"Falha ao atualizar filtros dinâmicos via API ({AppLoggerService.FormatElapsed(sw)}): {err}", nameof(BancadaChamadosService));
            return ([], []);
        }

        var tags = new List<string>();
        var locs = new List<string>();
        if (!TryBuildConnectionStringFromConfig(out string connStr, out string cfgErr))
        {
            AppLoggerService.Warning($"GetAvailableFiltersAsync abortado (PostgreSQL): {cfgErr}", nameof(BancadaChamadosService));
            return (tags, locs);
        }

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
            string cleanTag = string.IsNullOrWhiteSpace(tagFilter) || tagFilter.Equals("Todas as Tags", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : tagFilter.Trim().ToLowerInvariant();
            string[] cleanLocalidades = localidadesFilter?
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Select(l => l.Trim().ToLowerInvariant())
                .Distinct()
                .ToArray() ?? [];

            const string baseWhere = """
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

            string tagsSql = $"SELECT DISTINCT TRIM(tag) FROM chamados {baseWhere} AND (@hasLocalidades = false OR LOWER(TRIM(COALESCE(localidade_fisica, ''))) = ANY(@localidades)) AND TRIM(COALESCE(tag, '')) <> '' ORDER BY 1;";
            await using (var cmd = new NpgsqlCommand(tagsSql, conn))
            {
                cmd.Parameters.AddWithValue("@baseFilter", cleanBase);
                cmd.Parameters.AddWithValue("@statusFilter", cleanStatus);
                cmd.Parameters.AddWithValue("@search", cleanSearch);
                cmd.Parameters.AddWithValue("@hasLocalidades", cleanLocalidades.Length > 0);
                cmd.Parameters.AddWithValue("@localidades", cleanLocalidades);
                await using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                    tags.Add(r.GetString(0));
            }

            string locsSql = $"SELECT DISTINCT TRIM(localidade_fisica) FROM chamados {baseWhere} AND (@tagFilter = '' OR LOWER(COALESCE(tag, '')) = @tagFilter) AND TRIM(COALESCE(localidade_fisica, '')) <> '' ORDER BY 1;";
            await using (var cmd = new NpgsqlCommand(locsSql, conn))
            {
                cmd.Parameters.AddWithValue("@baseFilter", cleanBase);
                cmd.Parameters.AddWithValue("@statusFilter", cleanStatus);
                cmd.Parameters.AddWithValue("@search", cleanSearch);
                cmd.Parameters.AddWithValue("@tagFilter", cleanTag);
                await using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                    locs.Add(r.GetString(0));
            }
            AppLoggerService.Info($"GetAvailableFiltersAsync (PostgreSQL): tags={tags.Count}, localidades={locs.Count} em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaChamadosService), "GetAvailableFiltersAsync (PostgreSQL)", ex, sw.Elapsed, LogLevel.WARNING);
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
        var sw = Stopwatch.StartNew();
        bool useApi = BancadaApiClient.IsApiConfigured();
        AppLoggerService.Info(
            $"QueryChamadosAsync via {(useApi ? "API REST" : $"PostgreSQL direto ({ConfigService.BancadaDbHost}:{ConfigService.BancadaDbPort}/{ConfigService.BancadaDbName})")} | search='{search}', base='{baseFilter}', status='{statusFilter}', tag='{tagFilter}', locs={localidadesFilter?.Count ?? 0}, limit={limit}.",
            nameof(BancadaChamadosService));

        if (useApi)
        {
            string qs = BuildChamadosQueryString(search, baseFilter, statusFilter, Math.Clamp(limit, 1, 500), tagFilter, localidadesFilter);
            var (ok, res, statusCode, err) = await BancadaApiClient.GetAsync<ApiChamadosListResponse>($"/api/v1/chamados{qs}");
            if (!ok || res == null)
            {
                AppLoggerService.Warning($"QueryChamadosAsync via API REST falhou (HTTP {statusCode}) em {AppLoggerService.FormatElapsed(sw)}: {err}", nameof(BancadaChamadosService));
                return (false, [], new ChamadosSummary(), err);
            }

            var apiItems = (res.Items ?? []).Select(MapChamadoDto).ToList();
            var apiSummary = new ChamadosSummary
            {
                Total = res.Summary?.Total ?? 0,
                Abertos = res.Summary?.Abertos ?? 0,
                OtrsCount = res.Summary?.OtrsCount ?? 0,
                CitSmartCount = res.Summary?.CitSmartCount ?? 0
            };

            AppLoggerService.Success($"QueryChamadosAsync (API REST) retornou {apiItems.Count} chamado(s) (Total={apiSummary.Total}, Abertos={apiSummary.Abertos}) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
            return (true, apiItems, apiSummary, $"{apiItems.Count} chamado(s) carregado(s) via API REST da Bancada.");
        }

        if (!TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
            AppLoggerService.Warning($"QueryChamadosAsync abortado (fail-closed): {error}", nameof(BancadaChamadosService));
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

            AppLoggerService.Success(
                $"QueryChamadosAsync (PostgreSQL) retornou {items.Count} chamado(s) (Total={summary.Total}, Abertos={summary.Abertos}, OTRS={summary.OtrsCount}, CitSmart={summary.CitSmartCount}) em {AppLoggerService.FormatElapsed(sw)}.",
                nameof(BancadaChamadosService));
            return (true, items, summary, $"{items.Count} chamado(s) carregado(s) do PostgreSQL da Bancada.");
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaChamadosService), $"QueryChamadosAsync PostgreSQL ({ConfigService.BancadaDbHost}:{ConfigService.BancadaDbPort})", ex, sw.Elapsed, LogLevel.WARNING);
            return (false, items, summary, $"Erro ao conectar no PostgreSQL ({ConfigService.BancadaDbHost}:{ConfigService.BancadaDbPort}): {ex.Message}");
        }
    }

    public async Task<List<ChamadoComentario>> GetComentariosChamadoAsync(string chamadoId)
    {
        var sw = Stopwatch.StartNew();
        var list = new List<ChamadoComentario>();
        if (string.IsNullOrWhiteSpace(chamadoId))
            return list;

        if (BancadaApiClient.IsApiConfigured())
        {
            string encodedId = Uri.EscapeDataString(chamadoId.Trim());
            var (ok, res, _, err) = await BancadaApiClient.GetAsync<ApiComentariosResponse>($"/api/v1/chamados/{encodedId}/comentarios");
            if (ok && res?.Comentarios != null)
            {
                AppLoggerService.Info($"GetComentariosChamadoAsync (API REST) #{chamadoId}: {res.Comentarios.Count} comentário(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
                return res.Comentarios.Select(c => new ChamadoComentario
                {
                    Data = c.Data ?? string.Empty,
                    Autor = c.Autor ?? string.Empty,
                    Texto = c.Texto ?? string.Empty
                }).ToList();
            }
            AppLoggerService.Warning($"Falha ao carregar comentários via API do chamado #{chamadoId} ({AppLoggerService.FormatElapsed(sw)}): {err}", nameof(BancadaChamadosService));
            return list;
        }

        if (!TryBuildConnectionStringFromConfig(out string connStr, out string cfgErr))
        {
            AppLoggerService.Warning($"GetComentariosChamadoAsync #{chamadoId} abortado: {cfgErr}", nameof(BancadaChamadosService));
            return list;
        }

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
            AppLoggerService.Info($"GetComentariosChamadoAsync (PostgreSQL) #{chamadoId}: {list.Count} comentário(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaChamadosService), $"GetComentariosChamadoAsync #{chamadoId}", ex, sw.Elapsed, LogLevel.WARNING);
        }

        return list;
    }

    public async Task<Dictionary<string, List<ChamadoComentario>>> GetComentariosPorChamadosAsync(IEnumerable<string> chamadoIds)
    {
        var sw = Stopwatch.StartNew();
        var result = new Dictionary<string, List<ChamadoComentario>>(StringComparer.OrdinalIgnoreCase);
        var ids = chamadoIds?
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray() ?? [];

        if (ids.Length == 0)
            return result;

        if (BancadaApiClient.IsApiConfigured())
        {
            var body = new ApiComentariosBatchRequest { ChamadoIds = ids.Take(500).ToArray() };
            var (ok, res, _, err) = await BancadaApiClient.PostAsync<ApiComentariosBatchRequest, ApiComentariosBatchResponse>("/api/v1/chamados/comentarios/batch", body);
            if (ok && res?.Comentarios != null)
            {
                foreach (var kv in res.Comentarios)
                {
                    result[kv.Key] = (kv.Value ?? []).Select(c => new ChamadoComentario
                    {
                        Data = c.Data ?? string.Empty,
                        Autor = c.Autor ?? string.Empty,
                        Texto = c.Texto ?? string.Empty
                    }).ToList();
                }
                AppLoggerService.Info($"GetComentariosPorChamadosAsync (API REST): {result.Count} chamado(s) com comentários de {ids.Length} solicitado(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
            }
            else if (!ok)
            {
                AppLoggerService.Warning($"Falha ao carregar comentários em lote via API ({AppLoggerService.FormatElapsed(sw)}): {err}", nameof(BancadaChamadosService));
            }
            return result;
        }

        if (!TryBuildConnectionStringFromConfig(out string connStr, out string cfgErr))
        {
            AppLoggerService.Warning($"GetComentariosPorChamadosAsync abortado: {cfgErr}", nameof(BancadaChamadosService));
            return result;
        }

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
            AppLoggerService.Info($"GetComentariosPorChamadosAsync (PostgreSQL): {result.Count} chamado(s) com comentários de {ids.Length} solicitado(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaChamadosService), $"GetComentariosPorChamadosAsync ({ids.Length} IDs)", ex, sw.Elapsed, LogLevel.WARNING);
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

        var sw = Stopwatch.StartNew();
        bool useApi = BancadaApiClient.IsApiConfigured();
        AppLoggerService.Info(
            $"UpdateChamadoAsync #{id} via {(useApi ? "API REST" : "PostgreSQL direto")} (status='{status}', tag='{tag}', hostname='{hostname}').",
            nameof(BancadaChamadosService));

        if (useApi)
        {
            string encodedId = Uri.EscapeDataString(id.Trim());
            var payload = new ApiChamadoUpdateRequest
            {
                Status = status?.Trim() ?? string.Empty,
                Hostname = hostname?.Trim() ?? string.Empty,
                Tag = tag?.Trim() ?? string.Empty,
                CidadePredio = cidadePredio?.Trim() ?? string.Empty,
                Unidade = unidade?.Trim() ?? string.Empty,
                LocalidadeFisica = localidadeFisica?.Trim() ?? string.Empty,
                Descricao = descricao?.Trim() ?? string.Empty,
                Andamento = andamento?.Trim() ?? string.Empty,
                TagManual = true,
                DadosManuais = true
            };
            var (ok, res, statusCode, err) = await BancadaApiClient.PatchAsync<ApiChamadoUpdateRequest, ApiMutationResponse>($"/api/v1/chamados/{encodedId}", payload);
            if (ok)
                AppLoggerService.Success($"Chamado #{id} atualizado via API REST em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
            else
                AppLoggerService.Warning($"Falha ao atualizar chamado #{id} via API REST (HTTP {statusCode}) em {AppLoggerService.FormatElapsed(sw)}: {err}", nameof(BancadaChamadosService));
            return ok
                ? (true, res?.Message ?? $"Chamado #{id} atualizado via API REST com sucesso.")
                : (false, err);
        }

        if (!TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
            AppLoggerService.Warning($"UpdateChamadoAsync #{id} abortado: {error}", nameof(BancadaChamadosService));
            return (false, error);
        }

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
            {
                AppLoggerService.Warning($"UpdateChamadoAsync #{id} (PostgreSQL): nenhuma linha afetada em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
                return (false, $"Chamado #{id} não encontrado no PostgreSQL.");
            }

            AppLoggerService.Success($"Chamado #{id} atualizado no PostgreSQL ({rows} linha) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
            return (true, $"Chamado #{id} atualizado no PostgreSQL com sucesso.");
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaChamadosService), $"UpdateChamadoAsync #{id} PostgreSQL", ex, sw.Elapsed, LogLevel.WARNING);
            return (false, $"Erro ao salvar no PostgreSQL: {ex.Message}");
        }
    }

    public async Task<(bool Success, string Message)> AddComentarioChamadoAsync(string id, string autor, string texto)
    {
        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(autor) || string.IsNullOrWhiteSpace(texto))
            return (false, "Chamado, autor e texto são obrigatórios.");

        if (!BancadaApiClient.IsApiConfigured())
        {
            AppLoggerService.Warning($"AddComentarioChamadoAsync #{id} bloqueado: BANCADA_API_URL/KEY não configurados.", nameof(BancadaChamadosService));
            return (false, "Configure BANCADA_API_URL e BANCADA_API_KEY para adicionar comentários via API.");
        }

        var sw = Stopwatch.StartNew();
        string encodedId = Uri.EscapeDataString(id.Trim());
        var payload = new ApiComentarioCreateRequest
        {
            Autor = autor.Trim(),
            Texto = texto.Trim()
        };
        var (ok, _, statusCode, err) = await BancadaApiClient.PostAsync<ApiComentarioCreateRequest, ApiMutationResponse>($"/api/v1/chamados/{encodedId}/comentarios", payload);
        if (ok)
            AppLoggerService.Success($"Comentário adicionado ao chamado #{id} por '{autor.Trim()}' em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaChamadosService));
        else
            AppLoggerService.Warning($"Falha ao adicionar comentário ao chamado #{id} (HTTP {statusCode}) em {AppLoggerService.FormatElapsed(sw)}: {err}", nameof(BancadaChamadosService));
        return ok
            ? (true, $"Comentário registrado no chamado #{id}.")
            : (false, err);
    }

    private static string BuildChamadosQueryString(
        string? search,
        string? baseFilter,
        string? statusFilter,
        int? limit,
        string? tagFilter,
        List<string>? localidadesFilter)
    {
        var pairs = new List<string>();
        if (!string.IsNullOrWhiteSpace(search))
            pairs.Add($"search={Uri.EscapeDataString(search.Trim())}");
        if (!string.IsNullOrWhiteSpace(baseFilter) && !baseFilter.Equals("Todas as Bases", StringComparison.OrdinalIgnoreCase))
            pairs.Add($"base={Uri.EscapeDataString(baseFilter.Trim())}");
        if (!string.IsNullOrWhiteSpace(statusFilter))
            pairs.Add($"status={Uri.EscapeDataString(statusFilter.Trim())}");
        if (!string.IsNullOrWhiteSpace(tagFilter) && !tagFilter.Equals("Todas as Tags", StringComparison.OrdinalIgnoreCase))
            pairs.Add($"tag={Uri.EscapeDataString(tagFilter.Trim())}");
        if (localidadesFilter != null)
        {
            foreach (string loc in localidadesFilter.Where(l => !string.IsNullOrWhiteSpace(l)).Distinct())
                pairs.Add($"localidades={Uri.EscapeDataString(loc.Trim())}");
        }
        if (limit.HasValue)
            pairs.Add($"limit={limit.Value}");

        return pairs.Count > 0 ? "?" + string.Join("&", pairs) : string.Empty;
    }

    private static ChamadoItem MapChamadoDto(ApiChamadoDto d) => new()
    {
        Id = d.Id ?? string.Empty,
        DataCriacao = d.DataCriacao ?? string.Empty,
        Titulo = d.Titulo ?? string.Empty,
        CidadePredio = d.CidadePredio ?? string.Empty,
        Unidade = d.Unidade ?? string.Empty,
        LocalidadeFisica = d.LocalidadeFisica ?? string.Empty,
        Usuario = d.Usuario ?? string.Empty,
        IdCliente = d.IdCliente ?? string.Empty,
        Descricao = d.Descricao ?? string.Empty,
        Tag = d.Tag ?? string.Empty,
        IpOrigem = d.IpOrigem ?? string.Empty,
        Status = string.IsNullOrWhiteSpace(d.Status) ? "Aberto" : d.Status,
        DataAtualizacao = d.DataAtualizacao ?? string.Empty,
        Base = d.Base ?? string.Empty,
        Andamento = d.Andamento ?? string.Empty,
        Link = d.Link ?? string.Empty,
        Hostname = d.Hostname ?? string.Empty
    };

    private sealed class ApiLocalidadesResponse
    {
        [JsonPropertyName("localidades")] public List<string>? Localidades { get; set; }
    }

    private sealed class ApiFiltersResponse
    {
        [JsonPropertyName("tags")] public List<string>? Tags { get; set; }
        [JsonPropertyName("localidades")] public List<string>? Localidades { get; set; }
    }

    private sealed class ApiChamadosListResponse
    {
        [JsonPropertyName("items")] public List<ApiChamadoDto>? Items { get; set; }
        [JsonPropertyName("summary")] public ApiChamadosSummaryDto? Summary { get; set; }
    }

    private sealed class ApiChamadosSummaryDto
    {
        [JsonPropertyName("total")] public int Total { get; set; }
        [JsonPropertyName("abertos")] public int Abertos { get; set; }
        [JsonPropertyName("otrs_count")] public int OtrsCount { get; set; }
        [JsonPropertyName("citsmart_count")] public int CitSmartCount { get; set; }
    }

    private sealed class ApiChamadoDto
    {
        [JsonPropertyName("id")] public string? Id { get; set; }
        [JsonPropertyName("data_criacao")] public string? DataCriacao { get; set; }
        [JsonPropertyName("titulo")] public string? Titulo { get; set; }
        [JsonPropertyName("cidade_predio")] public string? CidadePredio { get; set; }
        [JsonPropertyName("unidade")] public string? Unidade { get; set; }
        [JsonPropertyName("localidade_fisica")] public string? LocalidadeFisica { get; set; }
        [JsonPropertyName("usuario")] public string? Usuario { get; set; }
        [JsonPropertyName("id_cliente")] public string? IdCliente { get; set; }
        [JsonPropertyName("descricao")] public string? Descricao { get; set; }
        [JsonPropertyName("tag")] public string? Tag { get; set; }
        [JsonPropertyName("ip_origem")] public string? IpOrigem { get; set; }
        [JsonPropertyName("status")] public string? Status { get; set; }
        [JsonPropertyName("data_atualizacao")] public string? DataAtualizacao { get; set; }
        [JsonPropertyName("base")] public string? Base { get; set; }
        [JsonPropertyName("andamento")] public string? Andamento { get; set; }
        [JsonPropertyName("link")] public string? Link { get; set; }
        [JsonPropertyName("hostname")] public string? Hostname { get; set; }
    }

    private sealed class ApiChamadoUpdateRequest
    {
        [JsonPropertyName("status")] public string? Status { get; set; }
        [JsonPropertyName("hostname")] public string? Hostname { get; set; }
        [JsonPropertyName("tag")] public string? Tag { get; set; }
        [JsonPropertyName("cidade_predio")] public string? CidadePredio { get; set; }
        [JsonPropertyName("unidade")] public string? Unidade { get; set; }
        [JsonPropertyName("localidade_fisica")] public string? LocalidadeFisica { get; set; }
        [JsonPropertyName("descricao")] public string? Descricao { get; set; }
        [JsonPropertyName("andamento")] public string? Andamento { get; set; }
        [JsonPropertyName("tag_manual")] public bool? TagManual { get; set; }
        [JsonPropertyName("dados_manuais")] public bool? DadosManuais { get; set; }
    }

    private sealed class ApiComentariosResponse
    {
        [JsonPropertyName("comentarios")] public List<ApiComentarioDto>? Comentarios { get; set; }
    }

    private sealed class ApiComentariosBatchRequest
    {
        [JsonPropertyName("chamado_ids")] public string[] ChamadoIds { get; set; } = [];
    }

    private sealed class ApiComentariosBatchResponse
    {
        [JsonPropertyName("comentarios")] public Dictionary<string, List<ApiComentarioDto>>? Comentarios { get; set; }
    }

    private sealed class ApiComentarioCreateRequest
    {
        [JsonPropertyName("autor")] public string Autor { get; set; } = string.Empty;
        [JsonPropertyName("texto")] public string Texto { get; set; } = string.Empty;
    }

    private sealed class ApiComentarioDto
    {
        [JsonPropertyName("data")] public string? Data { get; set; }
        [JsonPropertyName("autor")] public string? Autor { get; set; }
        [JsonPropertyName("texto")] public string? Texto { get; set; }
    }

    private sealed class ApiMutationResponse
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("message")] public string? Message { get; set; }
    }
}
