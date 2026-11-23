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

public class ZabbixAlertItem
{
    public string TriggerId { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public int Priority { get; set; }
    public long LastChangeUnix { get; set; }
    public string HostName { get; set; } = string.Empty;
    public string GroupSummary { get; set; } = string.Empty;

    public string LastChangeBr => LastChangeUnix > 0
        ? DateTimeOffset.FromUnixTimeSeconds(LastChangeUnix).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss")
        : "-";

    public string PriorityLabel => Priority switch
    {
        5 => "DESASTRE",
        4 => "ALTA",
        3 => "MÉDIA",
        2 => "AVISO",
        1 => "INFO",
        _ => "N/A"
    };

    public string PriorityColorHex => Priority switch
    {
        5 => "#ef4444",
        4 => "#f97316",
        3 => "#f59e0b",
        2 => "#eab308",
        _ => "#38bdf8"
    };
}

public enum ZabbixHostAvailability
{
    Unknown = 0,
    Online = 1,
    Offline = 2
}

public class ZabbixMonitoredHost
{
    public string HostId { get; set; } = string.Empty;
    public string Host { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string IpAddress { get; set; } = string.Empty;
    public string GroupName { get; set; } = "Sem Grupo";
    public ZabbixHostAvailability Availability { get; set; } = ZabbixHostAvailability.Unknown;

    public string StatusText => Availability switch
    {
        ZabbixHostAvailability.Online => "ONLINE",
        ZabbixHostAvailability.Offline => "OFFLINE",
        _ => "DESCONHECIDO"
    };

    public string StatusColorHex => Availability switch
    {
        ZabbixHostAvailability.Online => "#34d399",
        ZabbixHostAvailability.Offline => "#f87171",
        _ => "#94a3b8"
    };

    public string BadgeBgHex => Availability switch
    {
        ZabbixHostAvailability.Online => "#064e3b",
        ZabbixHostAvailability.Offline => "#7f1d1d",
        _ => "#1e293b"
    };

    public string BorderHex => Availability switch
    {
        ZabbixHostAvailability.Online => "#10b981",
        ZabbixHostAvailability.Offline => "#ef4444",
        _ => "#334155"
    };

    public string TargetForTrace => !string.IsNullOrWhiteSpace(IpAddress) ? IpAddress : Host;
}

public class ZabbixHostDetails
{
    public string HostId { get; set; } = string.Empty;
    public string Host { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string IpAddress { get; set; } = string.Empty;
    public List<string> Groups { get; set; } = [];
    public Dictionary<string, string> Inventory { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<ZabbixAlertItem> ActiveAlerts { get; set; } = [];
    public List<(string Name, string Value, string Units)> Metrics { get; set; } = [];
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

    /// <summary>
    /// Consulta alertas/triggers ativos no Zabbix (ex.: túneis/links caídos nas comarcas).
    /// </summary>
    public async Task<List<ZabbixAlertItem>> GetActiveAlertsAsync(
        int minSeverity = 2,
        int limit = 50,
        string? apiUrl = null,
        string? apiToken = null,
        CancellationToken cancellationToken = default)
    {
        string url = (apiUrl ?? ConfigService.ZabbixApiUrl).Trim();
        string token = (apiToken ?? ConfigService.ZabbixApiToken).Trim();

        if (!IsConfigured(url, token))
            return [];

        try
        {
            var payload = new
            {
                jsonrpc = "2.0",
                method = "trigger.get",
                @params = new
                {
                    output = new[] { "triggerid", "description", "priority", "lastchange" },
                    only_true = 1,
                    active = 1,
                    monitored = 1,
                    skipDependent = 1,
                    min_severity = minSeverity,
                    sortfield = "priority",
                    sortorder = "DESC",
                    limit = Math.Clamp(limit, 1, 100),
                    selectHosts = new[] { "hostid", "host", "name" },
                    selectGroups = new[] { "name" }
                },
                id = 3
            };

            string json = await PostJsonRpcAsync(url, token, payload, cancellationToken);
            var alerts = ParseActiveTriggersResponse(json);
            AppLoggerService.Info($"Zabbix API trigger.get: {alerts.Count} alerta(s) ativo(s) (min_severity={minSeverity}).", nameof(ZabbixEnrichmentService));
            return alerts;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao consultar alertas ativos no Zabbix: {ex.Message}", nameof(ZabbixEnrichmentService));
            return [];
        }
    }

    public static List<ZabbixAlertItem> ParseActiveTriggersResponse(string json)
    {
        var list = new List<ZabbixAlertItem>();
        if (string.IsNullOrWhiteSpace(json))
            return list;

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("result", out var arr) || arr.ValueKind != JsonValueKind.Array)
            return list;

        foreach (var item in arr.EnumerateArray())
        {
            string triggerId = item.TryGetProperty("triggerid", out var tid) ? (tid.GetString() ?? string.Empty) : string.Empty;
            string desc = item.TryGetProperty("description", out var dProp) ? (dProp.GetString() ?? string.Empty) : string.Empty;
            int priority = item.TryGetProperty("priority", out var pProp) && int.TryParse(pProp.GetString(), out int pVal) ? pVal : 0;
            long lastChange = item.TryGetProperty("lastchange", out var lcProp) && long.TryParse(lcProp.GetString(), out long lcVal) ? lcVal : 0;

            string hostName = string.Empty;
            if (item.TryGetProperty("hosts", out var hArr) && hArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var h in hArr.EnumerateArray())
                {
                    string n = h.TryGetProperty("name", out var np) ? (np.GetString() ?? string.Empty) : string.Empty;
                    string hn = h.TryGetProperty("host", out var hp) ? (hp.GetString() ?? string.Empty) : string.Empty;
                    hostName = !string.IsNullOrWhiteSpace(n) ? n : hn;
                    if (!string.IsNullOrWhiteSpace(hostName)) break;
                }
            }

            var groups = new List<string>();
            if (item.TryGetProperty("groups", out var gArr) && gArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var g in gArr.EnumerateArray())
                {
                    if (g.TryGetProperty("name", out var gn) && !string.IsNullOrWhiteSpace(gn.GetString()))
                        groups.Add(gn.GetString()!);
                }
            }

            list.Add(new ZabbixAlertItem
            {
                TriggerId = triggerId,
                Description = desc,
                Priority = priority,
                LastChangeUnix = lastChange,
                HostName = hostName,
                GroupSummary = groups.Count > 0 ? string.Join(" / ", groups.Take(2)) : string.Empty
            });
        }

        return list;
    }

    /// <summary>
    /// Lista todos os hosts monitorados no Zabbix com status de disponibilidade e grupo (cidade/prédio).
    /// </summary>
    public async Task<List<ZabbixMonitoredHost>> GetMonitoredHostsAsync(
        string? apiUrl = null,
        string? apiToken = null,
        CancellationToken cancellationToken = default)
    {
        string url = (apiUrl ?? ConfigService.ZabbixApiUrl).Trim();
        string token = (apiToken ?? ConfigService.ZabbixApiToken).Trim();

        if (!IsConfigured(url, token))
            return [];

        try
        {
            var payload = new
            {
                jsonrpc = "2.0",
                method = "host.get",
                @params = new
                {
                    output = new[] { "hostid", "host", "name", "status" },
                    filter = new { status = 0 },
                    selectInterfaces = new[] { "ip", "available" },
                    selectHostGroups = new[] { "name" }
                },
                id = 4
            };

            string json = await PostJsonRpcAsync(url, token, payload, cancellationToken);
            var hosts = ParseMonitoredHostsResponse(json);
            AppLoggerService.Info($"Zabbix API host.get (monitor): {hosts.Count} host(s) carregado(s).", nameof(ZabbixEnrichmentService));
            return hosts;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao listar hosts do Zabbix: {ex.Message}", nameof(ZabbixEnrichmentService));
            return [];
        }
    }

    public static List<ZabbixMonitoredHost> ParseMonitoredHostsResponse(string json)
    {
        var list = new List<ZabbixMonitoredHost>();
        if (string.IsNullOrWhiteSpace(json))
            return list;

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("result", out var arr) || arr.ValueKind != JsonValueKind.Array)
            return list;

        foreach (var item in arr.EnumerateArray())
        {
            string hostId = item.TryGetProperty("hostid", out var idProp) ? (idProp.GetString() ?? string.Empty) : string.Empty;
            string host = item.TryGetProperty("host", out var hProp) ? (hProp.GetString() ?? string.Empty) : string.Empty;
            string name = item.TryGetProperty("name", out var nProp) ? (nProp.GetString() ?? string.Empty) : string.Empty;

            string ip = string.Empty;
            var availability = ZabbixHostAvailability.Unknown;
            if (item.TryGetProperty("interfaces", out var ifArr) && ifArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var iface in ifArr.EnumerateArray())
                {
                    if (string.IsNullOrWhiteSpace(ip) && iface.TryGetProperty("ip", out var ipProp))
                        ip = ipProp.GetString() ?? string.Empty;

                    if (iface.TryGetProperty("available", out var avProp))
                    {
                        string avStr = avProp.ValueKind == JsonValueKind.Number
                            ? avProp.GetInt32().ToString()
                            : (avProp.GetString() ?? "0");
                        if (avStr == "2")
                        {
                            availability = ZabbixHostAvailability.Offline;
                            break;
                        }
                        if (avStr == "1")
                            availability = ZabbixHostAvailability.Online;
                    }
                }
            }

            string groupName = "Sem Grupo";
            if (item.TryGetProperty("hostgroups", out var gArr) && gArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var g in gArr.EnumerateArray())
                {
                    if (g.TryGetProperty("name", out var gn) && !string.IsNullOrWhiteSpace(gn.GetString()))
                    {
                        groupName = gn.GetString()!;
                        break;
                    }
                }
            }

            list.Add(new ZabbixMonitoredHost
            {
                HostId = hostId,
                Host = host,
                DisplayName = !string.IsNullOrWhiteSpace(name) ? name : host,
                IpAddress = ip,
                GroupName = groupName,
                Availability = availability
            });
        }

        return list
            .OrderByDescending(h => h.Availability == ZabbixHostAvailability.Offline)
            .ThenBy(h => h.GroupName)
            .ThenBy(h => h.DisplayName)
            .ToList();
    }

    /// <summary>
    /// Busca detalhes completos de um host específico no Zabbix (inventário, alertas ativos e métricas recentes).
    /// </summary>
    public async Task<ZabbixHostDetails?> GetHostFullDetailsAsync(
        string hostId,
        string? apiUrl = null,
        string? apiToken = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(hostId))
            return null;

        string url = (apiUrl ?? ConfigService.ZabbixApiUrl).Trim();
        string token = (apiToken ?? ConfigService.ZabbixApiToken).Trim();
        if (!IsConfigured(url, token))
            return null;

        try
        {
            var hostPayload = new
            {
                jsonrpc = "2.0",
                method = "host.get",
                @params = new
                {
                    hostids = new[] { hostId },
                    output = new[] { "hostid", "host", "name" },
                    selectInterfaces = new[] { "ip" },
                    selectHostGroups = new[] { "name" },
                    selectInventory = "extend"
                },
                id = 5
            };

            string hostJson = await PostJsonRpcAsync(url, token, hostPayload, cancellationToken);
            using var doc = JsonDocument.Parse(hostJson);
            if (!doc.RootElement.TryGetProperty("result", out var arr) || arr.ValueKind != JsonValueKind.Array || arr.GetArrayLength() == 0)
                return null;

            var item = arr[0];
            var details = new ZabbixHostDetails
            {
                HostId = hostId,
                Host = item.TryGetProperty("host", out var hp) ? (hp.GetString() ?? string.Empty) : string.Empty,
                DisplayName = item.TryGetProperty("name", out var np) ? (np.GetString() ?? string.Empty) : string.Empty
            };

            if (item.TryGetProperty("interfaces", out var ifArr) && ifArr.ValueKind == JsonValueKind.Array && ifArr.GetArrayLength() > 0)
                details.IpAddress = ifArr[0].TryGetProperty("ip", out var ipP) ? (ipP.GetString() ?? string.Empty) : string.Empty;

            if (item.TryGetProperty("hostgroups", out var gArr) && gArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var g in gArr.EnumerateArray())
                {
                    if (g.TryGetProperty("name", out var gn) && !string.IsNullOrWhiteSpace(gn.GetString()))
                        details.Groups.Add(gn.GetString()!);
                }
            }

            if (item.TryGetProperty("inventory", out var inv) && inv.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in inv.EnumerateObject())
                {
                    string val = prop.Value.ValueKind == JsonValueKind.String ? (prop.Value.GetString() ?? string.Empty) : prop.Value.ToString();
                    if (!string.IsNullOrWhiteSpace(val))
                        details.Inventory[prop.Name] = val;
                }
            }

            // Alertas ativos do host
            var trigPayload = new
            {
                jsonrpc = "2.0",
                method = "trigger.get",
                @params = new
                {
                    hostids = new[] { hostId },
                    only_true = 1,
                    active = 1,
                    output = new[] { "triggerid", "description", "priority", "lastchange" }
                },
                id = 6
            };
            string trigJson = await PostJsonRpcAsync(url, token, trigPayload, cancellationToken);
            details.ActiveAlerts = ParseActiveTriggersResponse(trigJson);

            // Métricas recentes (até 15 itens)
            var itemPayload = new
            {
                jsonrpc = "2.0",
                method = "item.get",
                @params = new
                {
                    hostids = new[] { hostId },
                    output = new[] { "name", "lastvalue", "units" },
                    filter = new { status = 0 },
                    limit = 15
                },
                id = 7
            };
            string itemJson = await PostJsonRpcAsync(url, token, itemPayload, cancellationToken);
            using var itemDoc = JsonDocument.Parse(itemJson);
            if (itemDoc.RootElement.TryGetProperty("result", out var mArr) && mArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var m in mArr.EnumerateArray())
                {
                    string mName = m.TryGetProperty("name", out var mn) ? (mn.GetString() ?? string.Empty) : string.Empty;
                    string mVal = m.TryGetProperty("lastvalue", out var mv) ? (mv.GetString() ?? string.Empty) : string.Empty;
                    string mUnits = m.TryGetProperty("units", out var mu) ? (mu.GetString() ?? string.Empty) : string.Empty;
                    if (!string.IsNullOrWhiteSpace(mName) && !string.IsNullOrWhiteSpace(mVal))
                        details.Metrics.Add((mName, mVal, mUnits));
                }
            }

            return details;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao carregar detalhes do host {hostId} no Zabbix: {ex.Message}", nameof(ZabbixEnrichmentService));
            return null;
        }
    }
}
