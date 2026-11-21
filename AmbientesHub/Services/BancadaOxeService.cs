using System.Diagnostics;
using System.Text.Json.Serialization;
using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

public class BancadaOxeService
{
    public async Task<(bool Success, List<OxeRamalItem> Items, OxeSummary Summary, string Message)> QueryRamaisAsync(string? search = null)
    {
        var sw = Stopwatch.StartNew();
        bool useApi = BancadaApiClient.IsApiConfigured();
        AppLoggerService.Info(
            $"QueryRamaisAsync via {(useApi ? "API REST" : $"PostgreSQL direto ({ConfigService.BancadaDbHost}:{ConfigService.BancadaDbPort}/{ConfigService.BancadaDbName})")} (search='{search}').",
            nameof(BancadaOxeService));

        if (useApi)
        {
            string qs = string.IsNullOrWhiteSpace(search) ? string.Empty : $"?search={Uri.EscapeDataString(search.Trim())}";
            var (ok, res, statusCode, err) = await BancadaApiClient.GetAsync<ApiOxeResponse>($"/api/v1/oxe{qs}");
            if (!ok || res == null)
            {
                AppLoggerService.Warning($"QueryRamaisAsync (API REST) falhou (HTTP {statusCode}) em {AppLoggerService.FormatElapsed(sw)}: {err}", nameof(BancadaOxeService));
                return (false, [], new OxeSummary(), err);
            }

            var apiItems = (res.Items ?? []).Select(x => new OxeRamalItem
            {
                Ramal = x.Ramal ?? string.Empty,
                Nome = x.Nome ?? string.Empty,
                Localidade = x.Localidade ?? string.Empty,
                Setor = x.Setor ?? string.Empty,
                Tipo = x.Tipo ?? string.Empty,
                Ip = x.Ip ?? string.Empty,
                Mac = x.Mac ?? string.Empty,
                GrupoCaptura = x.GrupoCaptura ?? string.Empty,
                CatRedePublica = x.CatRedePublica ?? string.Empty,
                CategoriaDispositivo = x.CategoriaDispositivo ?? string.Empty,
                TipoEstacao = x.TipoEstacao ?? string.Empty,
                AllFields = x.AllFields != null
                    ? new Dictionary<string, string>(x.AllFields, StringComparer.OrdinalIgnoreCase)
                    : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            }).ToList();

            var apiSummary = new OxeSummary
            {
                Total = res.Summary?.Total ?? apiItems.Count,
                ComNome = res.Summary?.ComNome ?? apiItems.Count(x => HasRealValue(x.Nome)),
                ComIpOuMac = res.Summary?.ComIpOuMac ?? apiItems.Count(x => HasRealValue(x.Ip) || HasRealValue(x.Mac))
            };

            AppLoggerService.Success($"QueryRamaisAsync (API REST): {apiItems.Count} ramal(is) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaOxeService));
            return (true, apiItems, apiSummary, $"{apiItems.Count} ramal(is) carregado(s) da Central OXE via API REST.");
        }

        if (!BancadaChamadosService.TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
            AppLoggerService.Warning($"QueryRamaisAsync abortado (fail-closed): {error}", nameof(BancadaOxeService));
            return (false, [], new OxeSummary(), error);
        }

        var allItems = new List<OxeRamalItem>();

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            const string sql = """
                SELECT *
                FROM central_telefonica
                ORDER BY CASE WHEN COALESCE(ramal::text, '') ~ '^[0-9]+$' THEN CAST(ramal AS INTEGER) ELSE 999999 END ASC,
                         COALESCE(ramal::text, '') ASC;
                """;

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

                var item = new OxeRamalItem
                {
                    Ramal = FirstNonEmpty(fields, "ramal", "directory_number", "numero", "extensao"),
                    Nome = FirstNonEmpty(fields, "nome", "usuario", "titular", "display_name", "nome_usuario", "name"),
                    Localidade = FirstNonEmpty(fields, "localidade", "entidade", "unidade", "cidade", "predio"),
                    Setor = FirstNonEmpty(fields, "setor", "departamento", "lotacao", "department"),
                    Tipo = FirstNonEmpty(fields, "tipo", "aparelho", "modelo", "set_type", "station_type"),
                    Ip = FirstNonEmpty(fields, "ip", "endereco_ip", "ip_address"),
                    Mac = FirstNonEmpty(fields, "mac", "mac_address", "endereco_mac"),
                    GrupoCaptura = FirstNonEmpty(fields, "grupo_de_captura", "grupo_captura", "pickup_group"),
                    CatRedePublica = FirstNonEmpty(fields, "cat_rede_pública", "cat_rede_publica", "categoria_rede_publica", "public_network_category"),
                    CategoriaDispositivo = FirstNonEmpty(fields, "categoria_dispositivo", "categoria_do_dispositivo", "categoria"),
                    TipoEstacao = FirstNonEmpty(fields, "tipo_de_estacao", "tipo_estacao", "station_type", "tipo"),
                    AllFields = fields
                };

                allItems.Add(item);
            }

            var summary = new OxeSummary
            {
                Total = allItems.Count,
                ComNome = allItems.Count(x => HasRealValue(x.Nome)),
                ComIpOuMac = allItems.Count(x => HasRealValue(x.Ip) || HasRealValue(x.Mac))
            };

            string q = search?.Trim() ?? string.Empty;
            List<OxeRamalItem> filtered = string.IsNullOrWhiteSpace(q)
                ? allItems
                : allItems.Where(item =>
                    item.AllFields.Values.Any(v => v.Contains(q, StringComparison.OrdinalIgnoreCase))
                ).ToList();

            AppLoggerService.Success(
                $"QueryRamaisAsync (PostgreSQL) concluído: {filtered.Count}/{allItems.Count} ramal(is) (comNome={summary.ComNome}, comIpOuMac={summary.ComIpOuMac}) em {AppLoggerService.FormatElapsed(sw)}.",
                nameof(BancadaOxeService));
            return (true, filtered, summary, $"{filtered.Count} ramal(is) carregado(s) da Central OXE.");
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaOxeService), "QueryRamaisAsync PostgreSQL (central_telefonica)", ex, sw.Elapsed, LogLevel.WARNING);
            return (false, [], new OxeSummary(), $"Erro ao consultar OXE no PostgreSQL: {ex.Message}");
        }
    }

    public static bool HasRealValue(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return false;
        string t = val.Trim();
        return t != "-" &&
               !t.Equals("N/A", StringComparison.OrdinalIgnoreCase) &&
               !t.Equals("nan", StringComparison.OrdinalIgnoreCase) &&
               !t.Equals("none", StringComparison.OrdinalIgnoreCase) &&
               !t.Equals("null", StringComparison.OrdinalIgnoreCase);
    }

    private static string FirstNonEmpty(Dictionary<string, string> fields, params string[] keys)
    {
        foreach (string k in keys)
        {
            if (fields.TryGetValue(k, out string? v) && HasRealValue(v))
                return v;
        }
        return string.Empty;
    }

    private sealed class ApiOxeResponse
    {
        [JsonPropertyName("items")] public List<ApiOxeRamalDto>? Items { get; set; }
        [JsonPropertyName("summary")] public ApiOxeSummaryDto? Summary { get; set; }
    }

    private sealed class ApiOxeSummaryDto
    {
        [JsonPropertyName("total")] public int Total { get; set; }
        [JsonPropertyName("com_nome")] public int ComNome { get; set; }
        [JsonPropertyName("com_ip_ou_mac")] public int ComIpOuMac { get; set; }
    }

    private sealed class ApiOxeRamalDto
    {
        [JsonPropertyName("ramal")] public string? Ramal { get; set; }
        [JsonPropertyName("nome")] public string? Nome { get; set; }
        [JsonPropertyName("localidade")] public string? Localidade { get; set; }
        [JsonPropertyName("setor")] public string? Setor { get; set; }
        [JsonPropertyName("tipo")] public string? Tipo { get; set; }
        [JsonPropertyName("ip")] public string? Ip { get; set; }
        [JsonPropertyName("mac")] public string? Mac { get; set; }
        [JsonPropertyName("grupo_captura")] public string? GrupoCaptura { get; set; }
        [JsonPropertyName("cat_rede_publica")] public string? CatRedePublica { get; set; }
        [JsonPropertyName("categoria_dispositivo")] public string? CategoriaDispositivo { get; set; }
        [JsonPropertyName("tipo_estacao")] public string? TipoEstacao { get; set; }
        [JsonPropertyName("all_fields")] public Dictionary<string, string>? AllFields { get; set; }
    }
}
