using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

public class ZabbixHostMatch
{
    public string IpAddress { get; set; } = string.Empty;
    public string HostId { get; set; } = string.Empty;
    public string Host { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public List<string> Groups { get; set; } = [];
}

/// <summary>
/// Consulta a API JSON-RPC do Zabbix (somente leitura) para enriquecer saltos de rede
/// com nome real do switch/roteador, grupos e localização de inventário.
/// Política: fail-closed — se ZABBIX_API_URL ou ZABBIX_API_TOKEN estiverem ausentes/inválidos, ignora silenciosamente.
/// </summary>
public class ZabbixEnrichmentService
{
    private static readonly HttpClient SharedHttp = new() { Timeout = TimeSpan.FromSeconds(5) };

    public static bool IsConfigured(string? apiUrl = null, string? apiToken = null)
    {
        string url = apiUrl ?? ConfigService.ZabbixApiUrl;
        string token = apiToken ?? ConfigService.ZabbixApiToken;
        return InputValidator.IsSafeHttpUrl(url) && !string.IsNullOrWhiteSpace(token);
    }

    public async Task<Dictionary<string, ZabbixHostMatch>> LookupByIpsAsync(
        IEnumerable<string> ipAddresses,
        string? apiUrl = null,
        string? apiToken = null,
        CancellationToken cancellationToken = default)
    {
        var result = new Dictionary<string, ZabbixHostMatch>(StringComparer.OrdinalIgnoreCase);
        string url = (apiUrl ?? ConfigService.ZabbixApiUrl).Trim();
        string token = (apiToken ?? ConfigService.ZabbixApiToken).Trim();

        if (!IsConfigured(url, token))
            return result;

        var uniqueIps = ipAddresses
            .Where(ip => !string.IsNullOrWhiteSpace(ip) && NetworkPathService.IsStrictIpAddress(ip, out _))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (uniqueIps.Count == 0)
            return result;

        var sw = Stopwatch.StartNew();
        try
        {
            AppLoggerService.Info($"Zabbix API LookupByIpsAsync iniciando para {uniqueIps.Count} IP(s) em '{url}'.", nameof(ZabbixEnrichmentService));
            var ifacePayload = new
            {
                jsonrpc = "2.0",
                method = "hostinterface.get",
                @params = new
                {
                    output = new[] { "hostid", "ip" },
                    filter = new { ip = uniqueIps }
                },
                id = 1
            };

            string ifaceJson = await PostJsonRpcAsync(url, token, ifacePayload, cancellationToken);
            var hostIdToIps = ParseHostInterfaceResponse(ifaceJson);
            if (hostIdToIps.Count == 0)
            {
                AppLoggerService.Info($"Zabbix API: 0 host(s) correspondente(s) para {uniqueIps.Count} IP(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(ZabbixEnrichmentService));
                return result;
            }

            var hostPayload = new
            {
                jsonrpc = "2.0",
                method = "host.get",
                @params = new
                {
                    hostids = hostIdToIps.Keys.ToArray(),
                    output = new[] { "hostid", "host", "name" },
                    selectHostGroups = new[] { "name" },
                    selectInventory = new[] { "location", "model", "type" }
                },
                id = 2
            };

            string hostJson = await PostJsonRpcAsync(url, token, hostPayload, cancellationToken);
            var parsed = ParseHostDetailsResponse(hostJson, hostIdToIps);
            AppLoggerService.Success($"Zabbix API: {uniqueIps.Count} IP(s) consultado(s), {parsed.Count} equipamento(s) identificado(s) em {AppLoggerService.FormatElapsed(sw)}.", nameof(ZabbixEnrichmentService));
            return parsed;
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(ZabbixEnrichmentService), $"Enriquecimento Zabbix ({uniqueIps.Count} IPs)", ex, sw.Elapsed, LogLevel.WARNING);
            return result;
        }
    }

    public static void ApplyEnrichment(
        IList<NetworkHopNode> hops,
        IReadOnlyDictionary<string, ZabbixHostMatch> matches)
    {
        if (hops == null || matches == null || matches.Count == 0)
            return;

        foreach (var hop in hops)
        {
            if (string.IsNullOrWhiteSpace(hop.IpAddress))
                continue;

            if (!matches.TryGetValue(hop.IpAddress.Trim(), out var match))
                continue;

            string zabbixName = !string.IsNullOrWhiteSpace(match.DisplayName) ? match.DisplayName : match.Host;
            if (!string.IsNullOrWhiteSpace(zabbixName) &&
                (string.IsNullOrWhiteSpace(hop.Hostname) || string.Equals(hop.Hostname, hop.IpAddress, StringComparison.OrdinalIgnoreCase)))
            {
                hop.Hostname = zabbixName;
            }

            if (!string.IsNullOrWhiteSpace(zabbixName))
            {
                hop.ProbableDeviceLabel = string.IsNullOrWhiteSpace(match.Model)
                    ? $"Zabbix: {zabbixName}"
                    : $"Zabbix: {zabbixName} ({match.Model})";
            }

            string loc = !string.IsNullOrWhiteSpace(match.Location)
                ? match.Location
                : (match.Groups.Count > 0 ? string.Join(" / ", match.Groups.Take(2)) : string.Empty);

            if (!string.IsNullOrWhiteSpace(loc))
                hop.InferredLocation = loc;

            if (!string.IsNullOrWhiteSpace(zabbixName))
                hop.Details = $"{hop.Details} | Zabbix: {zabbixName}";
        }
    }

    internal static Dictionary<string, List<string>> ParseHostInterfaceResponse(string json)
    {
        var map = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(json))
            return map;

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("result", out var arr) || arr.ValueKind != JsonValueKind.Array)
            return map;

        foreach (var item in arr.EnumerateArray())
        {
            string hostId = item.TryGetProperty("hostid", out var hProp) ? (hProp.GetString() ?? string.Empty) : string.Empty;
            string ip = item.TryGetProperty("ip", out var ipProp) ? (ipProp.GetString() ?? string.Empty) : string.Empty;
            if (string.IsNullOrWhiteSpace(hostId) || string.IsNullOrWhiteSpace(ip))
                continue;

            if (!map.TryGetValue(hostId, out var ips))
            {
                ips = [];
                map[hostId] = ips;
            }
            if (!ips.Contains(ip, StringComparer.OrdinalIgnoreCase))
                ips.Add(ip);
        }

        return map;
    }

    internal static Dictionary<string, ZabbixHostMatch> ParseHostDetailsResponse(
        string json,
        IReadOnlyDictionary<string, List<string>> hostIdToIps)
    {
        var result = new Dictionary<string, ZabbixHostMatch>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(json))
            return result;

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("result", out var arr) || arr.ValueKind != JsonValueKind.Array)
            return result;

        foreach (var item in arr.EnumerateArray())
        {
            string hostId = item.TryGetProperty("hostid", out var idProp) ? (idProp.GetString() ?? string.Empty) : string.Empty;
            if (string.IsNullOrWhiteSpace(hostId) || !hostIdToIps.TryGetValue(hostId, out var ips))
                continue;

            string host = item.TryGetProperty("host", out var hProp) ? (hProp.GetString() ?? string.Empty) : string.Empty;
            string name = item.TryGetProperty("name", out var nProp) ? (nProp.GetString() ?? string.Empty) : string.Empty;

            var groups = new List<string>();
            if (item.TryGetProperty("hostgroups", out var gArr) && gArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var g in gArr.EnumerateArray())
                {
                    if (g.TryGetProperty("name", out var gn) && !string.IsNullOrWhiteSpace(gn.GetString()))
                        groups.Add(gn.GetString()!);
                }
            }

            string location = string.Empty;
            string model = string.Empty;
            if (item.TryGetProperty("inventory", out var inv) && inv.ValueKind == JsonValueKind.Object)
            {
                if (inv.TryGetProperty("location", out var lProp))
                    location = lProp.GetString() ?? string.Empty;
                if (inv.TryGetProperty("model", out var mProp))
                    model = mProp.GetString() ?? string.Empty;
            }

            foreach (var ip in ips)
            {
                result[ip] = new ZabbixHostMatch
                {
                    IpAddress = ip,
                    HostId = hostId,
                    Host = host,
                    DisplayName = !string.IsNullOrWhiteSpace(name) ? name : host,
                    Location = location,
                    Model = model,
                    Groups = groups
                };
            }
        }

        return result;
    }

    private static async Task<string> PostJsonRpcAsync(string url, string token, object payload, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        string body = JsonSerializer.Serialize(payload);
        req.Content = new StringContent(body, Encoding.UTF8, "application/json");

        using var resp = await SharedHttp.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();
        return await resp.Content.ReadAsStringAsync(ct);
    }
}
