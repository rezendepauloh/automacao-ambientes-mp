using System.Diagnostics;
using System.DirectoryServices;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Descoberta de vizinhos de rede na mesma sub-rede /24, tabela ARP remota (MSFT_NetNeighbor),
/// inventário de mesma OU no Active Directory e enriquecimento com Zabbix/SCCM.
/// </summary>
public class SubnetNeighborsService
{
    public static (string Cidr, string GatewayIp, string PrefixDots) BuildSubnetPrefix24(IPAddress ip)
    {
        if (ip.AddressFamily != AddressFamily.InterNetwork)
            throw new ArgumentException("A varredura de vizinhos requer um endereço IPv4.");

        byte[] b = ip.GetAddressBytes();
        string prefix = $"{b[0]}.{b[1]}.{b[2]}.";
        return ($"{prefix}0/24", $"{prefix}1", prefix);
    }

    public static string ExtractParentOuDn(string? distinguishedName)
    {
        if (string.IsNullOrWhiteSpace(distinguishedName))
            return string.Empty;

        int firstComma = distinguishedName.IndexOf(',');
        if (firstComma < 0 || firstComma + 1 >= distinguishedName.Length)
            return string.Empty;

        string parent = distinguishedName[(firstComma + 1)..].Trim();
        return parent.StartsWith("OU=", StringComparison.OrdinalIgnoreCase) ? parent : string.Empty;
    }

    public static void MergeNeighborNode(List<SubnetNeighborNode> list, SubnetNeighborNode incoming)
    {
        var existing = list.FirstOrDefault(n =>
            (!string.IsNullOrWhiteSpace(incoming.IpAddress) &&
             string.Equals(n.IpAddress, incoming.IpAddress, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(incoming.Hostname) &&
             string.Equals(n.Hostname.Split('.')[0], incoming.Hostname.Split('.')[0], StringComparison.OrdinalIgnoreCase)));

        if (existing == null)
        {
            list.Add(incoming);
            return;
        }

        if (incoming.IsReferenceHost) existing.IsReferenceHost = true;
        if (incoming.IsOnline) existing.IsOnline = true;
        if (existing.LatencyMs == null && incoming.LatencyMs != null) existing.LatencyMs = incoming.LatencyMs;
        if (string.IsNullOrWhiteSpace(existing.IpAddress) && !string.IsNullOrWhiteSpace(incoming.IpAddress)) existing.IpAddress = incoming.IpAddress;
        if (string.IsNullOrWhiteSpace(existing.Hostname) && !string.IsNullOrWhiteSpace(incoming.Hostname)) existing.Hostname = incoming.Hostname;
        if (string.IsNullOrWhiteSpace(existing.MacAddress) && !string.IsNullOrWhiteSpace(incoming.MacAddress)) existing.MacAddress = incoming.MacAddress;
        if (string.IsNullOrWhiteSpace(existing.ConnectedUser) && !string.IsNullOrWhiteSpace(incoming.ConnectedUser)) existing.ConnectedUser = incoming.ConnectedUser;
        if (string.IsNullOrWhiteSpace(existing.LocationOrOu) && !string.IsNullOrWhiteSpace(incoming.LocationOrOu)) existing.LocationOrOu = incoming.LocationOrOu;
        if (!existing.SourceTag.Contains(incoming.SourceTag, StringComparison.OrdinalIgnoreCase))
            existing.SourceTag = $"{existing.SourceTag}+{incoming.SourceTag}";
    }

    public Task<SubnetNeighborsResult> DiscoverNeighborsAsync(
        string rawTarget,
        bool includeRemoteArp = true,
        bool includeSameAdOu = true,
        IProgress<SubnetNeighborsResult>? onProgress = null,
        CancellationToken cancellationToken = default)
    {
        string target = NetworkPathService.ValidateTarget(rawTarget);
        return Task.Run(() => DiscoverNeighborsCoreAsync(target, includeRemoteArp, includeSameAdOu, onProgress, cancellationToken), cancellationToken);
    }

    private async Task<SubnetNeighborsResult> DiscoverNeighborsCoreAsync(
        string target,
        bool includeRemoteArp,
        bool includeSameAdOu,
        IProgress<SubnetNeighborsResult>? onProgress,
        CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();

        var result = new SubnetNeighborsResult
        {
            ReferenceInput = target,
            ScannedAt = DateTime.Now
        };

        // 1. Resolve IP e FQDN do host de referência
        IPAddress? refIp = null;
        string refHost = target;

        if (NetworkPathService.IsStrictIpAddress(target, out var directIp) && directIp != null)
        {
            refIp = directIp;
            refHost = await ResolvePtrSafeAsync(directIp, 600, cancellationToken);
            if (string.IsNullOrWhiteSpace(refHost))
                refHost = target;
        }
        else
        {
            try
            {
                var addrTask = Dns.GetHostAddressesAsync(target);
                if (await Task.WhenAny(addrTask, Task.Delay(1500, cancellationToken)) == addrTask)
                {
                    var addrs = await addrTask;
                    refIp = addrs.FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);
                }
            }
            catch { }
        }

        string shortRefHost = refHost.Split('.')[0].Trim();
        var (adDn, adLoc) = await Task.Run(() => LookupComputerDnAndLocationInAd(shortRefHost), cancellationToken);
        string parentOuDn = ExtractParentOuDn(adDn);

        result.ReferenceHostname = refHost;
        result.ReferenceIp = refIp?.ToString() ?? string.Empty;
        result.AdOuPath = !string.IsNullOrWhiteSpace(adLoc) ? adLoc : FormatOuSummary(adDn);

        if (refIp != null && refIp.AddressFamily == AddressFamily.InterNetwork)
        {
            var (cidr, gwIp, prefixDots) = BuildSubnetPrefix24(refIp);
            result.SubnetCidr = cidr;
            result.GatewayIp = gwIp;

            // Tenta identificar o Switch/Gateway via Zabbix ou regra local CIDR
            var subnetRules = NetworkPathService.LoadSubnetRules();
            var matchedRule = NetworkPathService.MatchSubnetRule(gwIp, subnetRules)
                           ?? NetworkPathService.MatchSubnetRule(refIp.ToString(), subnetRules);
            if (matchedRule != null)
            {
                result.SwitchOrZabbixLabel = !string.IsNullOrWhiteSpace(matchedRule.ProbableSwitch)
                    ? $"{matchedRule.ProbableSwitch} ({matchedRule.LocationName})"
                    : matchedRule.LocationName;
            }

            if (ZabbixEnrichmentService.IsConfigured())
            {
                var zabbix = new ZabbixEnrichmentService();
                var zMatches = await zabbix.LookupByIpsAsync([gwIp, refIp.ToString()], cancellationToken: cancellationToken);
                if (zMatches.TryGetValue(gwIp, out var gwMatch))
                    result.SwitchOrZabbixLabel = $"{gwMatch.DisplayName} ({string.Join(" / ", gwMatch.Groups.Take(2))})".Trim(' ', '(', ')');
                else if (zMatches.TryGetValue(refIp.ToString(), out var hostMatch))
                    result.SwitchOrZabbixLabel = hostMatch.DisplayName;
            }

            // Adiciona nó de referência
            var refNode = new SubnetNeighborNode
            {
                Hostname = refHost,
                IpAddress = refIp.ToString(),
                MacAddress = NetworkPathService.TryResolveMacViaArp(refIp.ToString()),
                LocationOrOu = result.AdOuPath,
                IsReferenceHost = true,
                IsOnline = false,
                SourceTag = "REFERÊNCIA"
            };
            MergeNeighborNode(result.Neighbors, refNode);
            ReportSnapshot(result, onProgress);

            // 2. Coleta ARP remoto (MSFT_NetNeighbor) se solicitado e acessível
            if (includeRemoteArp)
            {
                var arpEntries = await TryGetRemoteArpEntriesAsync(refIp.ToString(), cancellationToken).ConfigureAwait(false);
                foreach (var entry in arpEntries)
                {
                    if (!entry.Ip.StartsWith(prefixDots, StringComparison.Ordinal))
                        continue;

                    string ptr = await ResolvePtrSafeAsync(IPAddress.Parse(entry.Ip), 350, cancellationToken).ConfigureAwait(false);
                    MergeNeighborNode(result.Neighbors, new SubnetNeighborNode
                    {
                        IpAddress = entry.Ip,
                        Hostname = ptr,
                        MacAddress = entry.Mac,
                        IsOnline = true,
                        SourceTag = "ARP-REMOTO"
                    });
                }
                ReportSnapshot(result, onProgress);
            }

            // 3. Varredura paralela da sub-rede /24 (Ping 350ms + PTR + ARP local)
            using var gate = new SemaphoreSlim(64);
            var scanTasks = Enumerable.Range(1, 254).Select(async octet =>
            {
                await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    string candidateIp = $"{prefixDots}{octet}";
                    using var ping = new Ping();
                    var reply = await ping.SendPingAsync(candidateIp, 350).ConfigureAwait(false);
                    if (reply.Status == IPStatus.Success)
                    {
                        string ptr = await ResolvePtrSafeAsync(reply.Address, 400, cancellationToken).ConfigureAwait(false);
                        string mac = NetworkPathService.TryResolveMacViaArp(candidateIp);
                        var node = new SubnetNeighborNode
                        {
                            IpAddress = candidateIp,
                            Hostname = ptr,
                            MacAddress = mac,
                            LatencyMs = reply.RoundtripTime,
                            IsOnline = true,
                            IsReferenceHost = string.Equals(candidateIp, result.ReferenceIp, StringComparison.OrdinalIgnoreCase),
                            SourceTag = string.IsNullOrWhiteSpace(mac) ? "ICMP+DNS" : "ICMP+ARP"
                        };

                        lock (result.Neighbors)
                        {
                            MergeNeighborNode(result.Neighbors, node);
                        }
                    }
                }
                catch { }
                finally
                {
                    gate.Release();
                }
            });

            await Task.WhenAll(scanTasks).ConfigureAwait(false);
            ReportSnapshot(result, onProgress);
        }
        else
        {
            MergeNeighborNode(result.Neighbors, new SubnetNeighborNode
            {
                Hostname = refHost,
                LocationOrOu = result.AdOuPath,
                IsReferenceHost = true,
                IsOnline = false,
                SourceTag = "AD"
            });
        }

        // 4. Complementa com máquinas da mesma OU no Active Directory (útil para identificar vizinhos mesmo com Ping bloqueado)
        if (includeSameAdOu && !string.IsNullOrWhiteSpace(parentOuDn))
        {
            var ouPeers = await Task.Run(() => ListSameOuComputersFromAd(parentOuDn, maxResults: 40), cancellationToken);
            lock (result.Neighbors)
            {
                foreach (var peer in ouPeers)
                    MergeNeighborNode(result.Neighbors, peer);
            }
            ReportSnapshot(result, onProgress);
        }

        // Ordena: Referência primeiro, depois Online, depois por IP/Hostname
        result.Neighbors = result.Neighbors
            .OrderByDescending(n => n.IsReferenceHost)
            .ThenByDescending(n => n.IsOnline)
            .ThenBy(n => ParseIpSortKey(n.IpAddress))
            .ThenBy(n => n.TitleDisplay)
            .ToList();

        sw.Stop();
        AppLoggerService.Info(
            $"Descoberta de vizinhos '{target}' ({result.SubnetCidr}): {result.Neighbors.Count} host(s) em {sw.ElapsedMilliseconds} ms.",
            nameof(SubnetNeighborsService));

        return result;
    }

    private static void ReportSnapshot(SubnetNeighborsResult source, IProgress<SubnetNeighborsResult>? onProgress)
    {
        if (onProgress == null) return;
        List<SubnetNeighborNode> copy;
        lock (source.Neighbors)
        {
            copy = [.. source.Neighbors];
        }

        onProgress.Report(new SubnetNeighborsResult
        {
            ReferenceInput = source.ReferenceInput,
            ReferenceIp = source.ReferenceIp,
            ReferenceHostname = source.ReferenceHostname,
            SubnetCidr = source.SubnetCidr,
            GatewayIp = source.GatewayIp,
            SwitchOrZabbixLabel = source.SwitchOrZabbixLabel,
            AdOuPath = source.AdOuPath,
            ScannedAt = source.ScannedAt,
            Neighbors = copy
        });
    }

    private static uint ParseIpSortKey(string ip)
    {
        if (IPAddress.TryParse(ip, out var addr) && addr.AddressFamily == AddressFamily.InterNetwork)
        {
            byte[] b = addr.GetAddressBytes();
            return ((uint)b[0] << 24) | ((uint)b[1] << 16) | ((uint)b[2] << 8) | b[3];
        }
        return uint.MaxValue;
    }

    private static async Task<string> ResolvePtrSafeAsync(IPAddress ip, int timeoutMs, CancellationToken ct)
    {
        try
        {
            var task = Dns.GetHostEntryAsync(ip);
            if (await Task.WhenAny(task, Task.Delay(timeoutMs, ct)) == task)
                return (await task).HostName ?? string.Empty;
        }
        catch { }
        return string.Empty;
    }

    private static async Task<List<(string Ip, string Mac)>> TryGetRemoteArpEntriesAsync(string targetIp, CancellationToken ct)
    {
        var entries = new List<(string Ip, string Mac)>();
        try
        {
            var task = Task.Run(() =>
            {
                var options = new ConnectionOptions { Timeout = TimeSpan.FromSeconds(2) };
                var scope = new ManagementScope($@"\\{targetIp}\root\StandardCimv2", options);
                scope.Connect();
                using var searcher = new ManagementObjectSearcher(scope,
                    new ObjectQuery("SELECT IPAddress, LinkLayerAddress, State FROM MSFT_NetNeighbor WHERE AddressFamily = 2"));
                using var results = searcher.Get();
                foreach (ManagementObject obj in results)
                {
                    string ip = obj["IPAddress"]?.ToString() ?? string.Empty;
                    string mac = obj["LinkLayerAddress"]?.ToString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(ip) && !string.IsNullOrWhiteSpace(mac) &&
                        !mac.Equals("00-00-00-00-00-00", StringComparison.OrdinalIgnoreCase) &&
                        !mac.Equals("FF-FF-FF-FF-FF-FF", StringComparison.OrdinalIgnoreCase))
                    {
                        entries.Add((ip, mac.Replace('-', ':').ToUpperInvariant()));
                    }
                }
            }, ct);

            await Task.WhenAny(task, Task.Delay(2200, ct));
        }
        catch { }
        return entries;
    }

    private static (string Dn, string Location) LookupComputerDnAndLocationInAd(string shortHost)
    {
        try
        {
            if (!InputValidator.IsValidHostName(shortHost))
                return (string.Empty, string.Empty);

            using var searcher = new DirectorySearcher();
            string escaped = InputValidator.EscapeLdap(shortHost);
            searcher.Filter = $"(&(objectCategory=computer)(cn={escaped}))";
            searcher.PropertiesToLoad.AddRange(["distinguishedName", "location"]);
            searcher.SizeLimit = 1;
            searcher.ServerTimeLimit = TimeSpan.FromSeconds(2);

            var res = searcher.FindOne();
            if (res == null) return (string.Empty, string.Empty);

            string dn = res.Properties.Contains("distinguishedName") && res.Properties["distinguishedName"].Count > 0
                ? res.Properties["distinguishedName"][0]?.ToString() ?? string.Empty
                : string.Empty;
            string loc = res.Properties.Contains("location") && res.Properties["location"].Count > 0
                ? res.Properties["location"][0]?.ToString() ?? string.Empty
                : string.Empty;

            return (dn, loc);
        }
        catch
        {
            return (string.Empty, string.Empty);
        }
    }

    private static List<SubnetNeighborNode> ListSameOuComputersFromAd(string parentOuDn, int maxResults)
    {
        var list = new List<SubnetNeighborNode>();
        try
        {
            using var entry = new DirectoryEntry($"LDAP://{parentOuDn}");
            using var searcher = new DirectorySearcher(entry)
            {
                Filter = "(objectCategory=computer)",
                SearchScope = SearchScope.OneLevel,
                SizeLimit = maxResults,
                ServerTimeLimit = TimeSpan.FromSeconds(3)
            };
            searcher.PropertiesToLoad.AddRange(["cn", "dNSHostName", "location", "distinguishedName"]);

            foreach (SearchResult r in searcher.FindAll())
            {
                string cn = r.Properties.Contains("cn") && r.Properties["cn"].Count > 0 ? r.Properties["cn"][0]?.ToString() ?? "" : "";
                string dns = r.Properties.Contains("dNSHostName") && r.Properties["dNSHostName"].Count > 0 ? r.Properties["dNSHostName"][0]?.ToString() ?? "" : "";
                string loc = r.Properties.Contains("location") && r.Properties["location"].Count > 0 ? r.Properties["location"][0]?.ToString() ?? "" : "";
                string dn = r.Properties.Contains("distinguishedName") && r.Properties["distinguishedName"].Count > 0 ? r.Properties["distinguishedName"][0]?.ToString() ?? "" : "";

                if (string.IsNullOrWhiteSpace(cn)) continue;
                list.Add(new SubnetNeighborNode
                {
                    Hostname = !string.IsNullOrWhiteSpace(dns) ? dns : cn,
                    LocationOrOu = !string.IsNullOrWhiteSpace(loc) ? loc : FormatOuSummary(dn),
                    IsOnline = false,
                    SourceTag = "AD-OU"
                });
            }
        }
        catch { }
        return list;
    }

    private static string FormatOuSummary(string dn)
    {
        if (string.IsNullOrWhiteSpace(dn)) return string.Empty;
        var ous = dn.Split(',')
            .Select(p => p.Trim())
            .Where(p => p.StartsWith("OU=", StringComparison.OrdinalIgnoreCase))
            .Select(p => p[3..].Trim())
            .Take(2)
            .ToList();
        return ous.Count > 0 ? $"OU: {string.Join(" / ", ous)}" : string.Empty;
    }
}
