using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

public class BancadaOxeService
{
    public async Task<(bool Success, List<OxeRamalItem> Items, OxeSummary Summary, string Message)> QueryRamaisAsync(string? search = null)
    {
        if (!BancadaChamadosService.TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
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

            return (true, filtered, summary, $"{filtered.Count} ramal(is) carregado(s) da Central OXE.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao consultar tabela central_telefonica: {ex.Message}", nameof(BancadaOxeService));
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
}
