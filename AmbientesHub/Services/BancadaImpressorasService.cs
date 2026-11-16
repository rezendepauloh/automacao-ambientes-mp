using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

public class BancadaImpressorasService
{
    public async Task<(bool Success, List<ImpressoraItem> Items, ImpressorasSummary Summary, string Message)> QueryImpressorasAsync()
    {
        if (!BancadaChamadosService.TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
            return (false, [], new ImpressorasSummary(), error);
        }

        var items = new List<ImpressoraItem>();

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            const string sql = "SELECT * FROM impressoras ORDER BY COALESCE(nome, '') ASC;";
            await using var cmd = new NpgsqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    string colName = reader.GetName(i);
                    string val = reader.IsDBNull(i) ? string.Empty : (reader.GetValue(i)?.ToString() ?? string.Empty).Trim();
                    fields[colName] = val;
                }

                long.TryParse(FirstNonEmpty(fields, "total_paginas"), out long totalPaginas);

                items.Add(new ImpressoraItem
                {
                    Id = FirstNonEmpty(fields, "id"),
                    Nome = FirstNonEmpty(fields, "nome"),
                    Servidor = FirstNonEmpty(fields, "servidor"),
                    Tipo = FirstNonEmpty(fields, "tipo"),
                    Modelo = FirstNonEmpty(fields, "modelo"),
                    Localizacao = FirstNonEmpty(fields, "localizacao"),
                    IpHost = FirstNonEmpty(fields, "ip_host"),
                    Status = FirstNonEmpty(fields, "status"),
                    TotalPaginas = totalPaginas,
                    FilasRelacionadas = FirstNonEmpty(fields, "filas_relacionadas"),
                    DetalhesExtra = FirstNonEmpty(fields, "detalhes_extra"),
                    DataAtualizacao = FirstNonEmpty(fields, "data_atualizacao"),
                    AllFields = fields
                });
            }

            var summary = new ImpressorasSummary
            {
                Total = items.Count,
                OkCount = items.Count(x => x.IsStatusOk),
                AlertCount = items.Count(x => !x.IsStatusOk),
                TotalPaginasSoma = items.Sum(x => x.TotalPaginas)
            };

            return (true, items, summary, $"{items.Count} impressora(s)/dispositivo(s) carregado(s) do PaperCut.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao consultar tabela impressoras: {ex.Message}", nameof(BancadaImpressorasService));
            return (false, [], new ImpressorasSummary(), $"Erro ao consultar Impressoras no PostgreSQL: {ex.Message}");
        }
    }

    private static string FirstNonEmpty(Dictionary<string, string> fields, string key)
    {
        if (fields.TryGetValue(key, out string? v) && BancadaOxeService.HasRealValue(v))
            return v;
        return string.Empty;
    }
}
