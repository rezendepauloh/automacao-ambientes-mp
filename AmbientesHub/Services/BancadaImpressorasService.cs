using System.Text.Json.Serialization;
using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

public class BancadaImpressorasService
{
    public async Task<(bool Success, List<ImpressoraItem> Items, ImpressorasSummary Summary, string Message)> QueryImpressorasAsync()
    {
        if (BancadaApiClient.IsApiConfigured())
        {
            var (ok, res, _, err) = await BancadaApiClient.GetAsync<ApiImpressorasResponse>("/api/v1/impressoras");
            if (!ok || res == null)
                return (false, [], new ImpressorasSummary(), err);

            var apiItems = (res.Items ?? []).Select(x => new ImpressoraItem
            {
                Id = x.Id ?? string.Empty,
                Nome = x.Nome ?? string.Empty,
                Servidor = x.Servidor ?? string.Empty,
                Tipo = x.Tipo ?? string.Empty,
                Modelo = x.Modelo ?? string.Empty,
                Localizacao = x.Localizacao ?? string.Empty,
                IpHost = x.IpHost ?? string.Empty,
                Status = x.Status ?? string.Empty,
                TotalPaginas = x.TotalPaginas,
                FilasRelacionadas = x.FilasRelacionadas ?? string.Empty,
                DetalhesExtra = x.DetalhesExtra ?? string.Empty,
                DataAtualizacao = x.DataAtualizacao ?? string.Empty,
                AllFields = x.AllFields != null
                    ? new Dictionary<string, string>(x.AllFields, StringComparer.OrdinalIgnoreCase)
                    : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            }).ToList();

            var apiSummary = new ImpressorasSummary
            {
                Total = res.Summary?.Total ?? apiItems.Count,
                OkCount = res.Summary?.OkCount ?? apiItems.Count(x => x.IsStatusOk),
                AlertCount = res.Summary?.AlertCount ?? apiItems.Count(x => !x.IsStatusOk),
                TotalPaginasSoma = res.Summary?.TotalPaginasSoma ?? apiItems.Sum(x => x.TotalPaginas)
            };

            return (true, apiItems, apiSummary, $"{apiItems.Count} impressora(s)/dispositivo(s) carregado(s) do PaperCut via API REST.");
        }

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

    private sealed class ApiImpressorasResponse
    {
        [JsonPropertyName("items")] public List<ApiImpressoraDto>? Items { get; set; }
        [JsonPropertyName("summary")] public ApiImpressorasSummaryDto? Summary { get; set; }
    }

    private sealed class ApiImpressorasSummaryDto
    {
        [JsonPropertyName("total")] public int Total { get; set; }
        [JsonPropertyName("ok_count")] public int OkCount { get; set; }
        [JsonPropertyName("alert_count")] public int AlertCount { get; set; }
        [JsonPropertyName("total_paginas_soma")] public long TotalPaginasSoma { get; set; }
    }

    private sealed class ApiImpressoraDto
    {
        [JsonPropertyName("id")] public string? Id { get; set; }
        [JsonPropertyName("nome")] public string? Nome { get; set; }
        [JsonPropertyName("servidor")] public string? Servidor { get; set; }
        [JsonPropertyName("tipo")] public string? Tipo { get; set; }
        [JsonPropertyName("modelo")] public string? Modelo { get; set; }
        [JsonPropertyName("localizacao")] public string? Localizacao { get; set; }
        [JsonPropertyName("ip_host")] public string? IpHost { get; set; }
        [JsonPropertyName("status")] public string? Status { get; set; }
        [JsonPropertyName("total_paginas")] public long TotalPaginas { get; set; }
        [JsonPropertyName("filas_relacionadas")] public string? FilasRelacionadas { get; set; }
        [JsonPropertyName("detalhes_extra")] public string? DetalhesExtra { get; set; }
        [JsonPropertyName("data_atualizacao")] public string? DataAtualizacao { get; set; }
        [JsonPropertyName("all_fields")] public Dictionary<string, string>? AllFields { get; set; }
    }
}
