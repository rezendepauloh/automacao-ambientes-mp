using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo de diagnóstico de rede (Ping contínuo, Traceroute L3 assíncrono,
/// resolução DNS/PTR/AD e inferência de localização por CIDR e nomenclatura).
/// </summary>
public class NetworkPathService
{
    private const int DefaultHighLatencyThresholdMs = 80;
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static string GetDefaultSubnetMapPath()
    {
        string envDir = Path.GetDirectoryName(ConfigService.GetEnvFilePath()) ?? AppDomain.CurrentDomain.BaseDirectory;
        return Path.Combine(envDir, "subnet_map.json");
    }

    /// <summary>
    /// Retorna o diretório configurado para relatórios (mesmo de Chamados), com fallback seguro.
    /// </summary>
    public static string GetReportsDirectory()
    {
        string dir = ConfigService.ChamadosReportsDir;
        if (InputValidator.IsSafeDirectoryPath(dir))
            return dir;

        return ConfigService.GetRecommendedDefault("CHAMADOS_REPORTS_DIR");
    }

    /// <summary>
    /// Valida o destino (hostname NetBIOS/FQDN ou endereço IPv4/IPv6) em modo fail-closed.
    /// </summary>
    public static string ValidateTarget(string? rawTarget)
    {
        if (string.IsNullOrWhiteSpace(rawTarget))
            throw new ArgumentException("Informe o hostname ou endereço IP de destino.");

        string trimmed = rawTarget.Trim();
        if (trimmed.Length > 253)
            throw new ArgumentException("Destino excede o tamanho máximo permitido (253 caracteres).");

        if (IsStrictIpAddress(trimmed, out _))
            return trimmed;

        // Rejeita strings puramente numéricas ou IPv4 malformados (ex: 10.0.0.999)
        if (trimmed.All(char.IsDigit) ||
            (trimmed.Contains('.', StringComparison.Ordinal) && trimmed.All(c => char.IsDigit(c) || c == '.')))
        {
            throw new ArgumentException("Endereço IP inválido.");
        }

        if (!InputValidator.IsValidHostName(trimmed))
            throw new ArgumentException("Destino inválido. Informe um hostname DNS/NetBIOS ou IP válido.");

        return trimmed;
    }

    public static bool IsStrictIpAddress(string? value, out IPAddress? parsed)
    {
        parsed = null;
        if (string.IsNullOrWhiteSpace(value))
            return false;

        string trimmed = value.Trim();
        bool looksIpv4 = trimmed.Count(c => c == '.') == 3;
        bool looksIpv6 = trimmed.Contains(':', StringComparison.Ordinal);

        if (!looksIpv4 && !looksIpv6)
            return false;

        return IPAddress.TryParse(trimmed, out parsed);
    }

    /// <summary>
    /// Carrega regras CIDR -> Localização/Switch de um arquivo JSON local.
    /// </summary>
    public static List<SubnetLocationRule> LoadSubnetRules(string? filePath = null)
    {
        string path = string.IsNullOrWhiteSpace(filePath) ? GetDefaultSubnetMapPath() : filePath;
        try
        {
            if (!File.Exists(path))
            {
                AppLoggerService.Info($"Mapa de sub-redes não encontrado em '{path}'; 0 regras carregadas.", nameof(NetworkPathService));
                return [];
            }

            string json = File.ReadAllText(path, Encoding.UTF8);
            var raw = JsonSerializer.Deserialize<List<SubnetLocationRule>>(json, JsonOptions) ?? [];
            var valid = raw
                .Where(r => SubnetLocationRule.TryParseCidr(r.Cidr, out _, out _, out _))
                .OrderByDescending(r => r.PrefixLength)
                .ToList();
            AppLoggerService.Info($"Mapa de sub-redes carregado de '{path}': {valid.Count} regra(s) válida(s).", nameof(NetworkPathService));
            return valid;
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(NetworkPathService), $"Ler mapa de sub-redes '{path}'", ex, level: LogLevel.WARNING);
            return [];
        }
    }

    /// <summary>
    /// Salva regras CIDR válidas no arquivo JSON local.
    /// </summary>
    public static void SaveSubnetRules(IEnumerable<SubnetLocationRule> rules, string? filePath = null)
    {
        string path = string.IsNullOrWhiteSpace(filePath) ? GetDefaultSubnetMapPath() : filePath;
        var validRules = rules
            .Where(r => SubnetLocationRule.TryParseCidr(r.Cidr, out _, out _, out _))
            .Select(r => new SubnetLocationRule
            {
                Cidr = r.Cidr.Trim(),
                LocationName = (r.LocationName ?? string.Empty).Trim(),
                ProbableSwitch = (r.ProbableSwitch ?? string.Empty).Trim(),
                Notes = (r.Notes ?? string.Empty).Trim()
            })
            .OrderByDescending(r => r.PrefixLength)
            .ToList();

        string? dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        string json = JsonSerializer.Serialize(validRules, JsonOptions);
        File.WriteAllText(path, json, Encoding.UTF8);
        AppLoggerService.Success($"Mapa de sub-redes salvo em '{path}' ({validRules.Count} regra(s)).", nameof(NetworkPathService));
    }

    /// <summary>
    /// Obtém os dados da estação local (hostname, IP principal e gateway padrão) para o nó de origem.
    /// </summary>
    public NetworkHopNode GetLocalOriginNode(IReadOnlyList<SubnetLocationRule>? subnetRules = null)
    {
        string localHost = Environment.MachineName;
        string localIp = "127.0.0.1";
        string localMac = string.Empty;
        string gatewayIp = string.Empty;
        string nicDesc = string.Empty;

        try
        {
            var activeNic = NetworkInterface.GetAllNetworkInterfaces()
                .Where(n => n.OperationalStatus == OperationalStatus.Up &&
                            n.NetworkInterfaceType != NetworkInterfaceType.Loopback &&
                            n.NetworkInterfaceType != NetworkInterfaceType.Tunnel)
                .OrderByDescending(n => n.GetIPProperties().GatewayAddresses.Any(g => g.Address.AddressFamily == AddressFamily.InterNetwork))
                .ThenBy(n => n.NetworkInterfaceType == NetworkInterfaceType.Ethernet ? 0 : 1)
                .FirstOrDefault();

            if (activeNic != null)
            {
                nicDesc = activeNic.Name;
                var props = activeNic.GetIPProperties();
                var ipv4 = props.UnicastAddresses
                    .FirstOrDefault(u => u.Address.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(u.Address));
                if (ipv4 != null)
                    localIp = ipv4.Address.ToString();

                var gw = props.GatewayAddresses
                    .FirstOrDefault(g => g.Address.AddressFamily == AddressFamily.InterNetwork);
                if (gw != null)
                    gatewayIp = gw.Address.ToString();

                localMac = FormatPhysicalAddress(activeNic.GetPhysicalAddress());
            }
        }
        catch { }

        var matched = MatchSubnetRule(localIp, subnetRules);
        string loc = matched != null && !string.IsNullOrWhiteSpace(matched.LocationName)
            ? matched.LocationName
            : (!string.IsNullOrWhiteSpace(nicDesc) ? $"Interface: {nicDesc}" : "Estação de Trabalho Local");

        string localUser = string.IsNullOrWhiteSpace(Environment.UserDomainName)
            ? Environment.UserName
            : $"{Environment.UserDomainName}\\{Environment.UserName}";

        return new NetworkHopNode
        {
            HopNumber = 0,
            Hostname = localHost,
            IpAddress = localIp,
            LatencyMs = 0,
            Kind = NetworkHopKind.LocalStation,
            Status = NetworkHopStatus.Online,
            InferredLocation = loc,
            ConnectedUser = localUser,
            MacAddress = localMac,
            ProbableDeviceLabel = matched != null && !string.IsNullOrWhiteSpace(matched.ProbableSwitch)
                ? $"Estação Local • Switch: {matched.ProbableSwitch}"
                : "Estação Local (Origem)",
            MatchedCidr = matched?.Cidr ?? string.Empty,
            Details = string.IsNullOrWhiteSpace(gatewayIp)
                ? $"Origem local: {localHost} ({localIp})"
                : $"Origem local: {localHost} ({localIp}) | Gateway: {gatewayIp}"
        };
    }

    /// <summary>
    /// Executa um único ping assíncrono validado.
    /// </summary>
    public async Task<(bool Success, long? RttMs, string ResolvedIp, int? Ttl, IPStatus Status)> PingOnceAsync(
        string rawTarget,
        int timeoutMs = 1500,
        CancellationToken cancellationToken = default)
    {
        string target = ValidateTarget(rawTarget);
        using var ping = new Ping();
        var sw = Stopwatch.StartNew();

        try
        {
            var pingTask = ping.SendPingAsync(target, timeoutMs);
            var completed = await Task.WhenAny(pingTask, Task.Delay(timeoutMs + 300, cancellationToken));
            if (completed != pingTask)
                return (false, null, string.Empty, null, IPStatus.TimedOut);

            var reply = await pingTask;
            sw.Stop();
            bool ok = reply.Status == IPStatus.Success;
            long? rtt = ok ? (reply.RoundtripTime > 0 ? reply.RoundtripTime : Math.Max(1, sw.ElapsedMilliseconds)) : null;
            string ip = reply.Address?.ToString() ?? string.Empty;
            int? ttl = reply.Options?.Ttl;
            return (ok, rtt, ip, ttl, reply.Status);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            return (false, null, string.Empty, null, IPStatus.Unknown);
        }
    }

    /// <summary>
    /// Executa Traceroute L3 nativo com resolução PTR, detecção de destino offline e inferência de localização.
    /// </summary>
    public async Task<NetworkTraceResult> TraceRouteAsync(
        string rawTarget,
        int maxHops = 25,
        int timeoutMs = 1200,
        IReadOnlyList<SubnetLocationRule>? subnetRules = null,
        IProgress<NetworkHopNode>? onHopDiscovered = null,
        CancellationToken cancellationToken = default)
    {
        var swTotal = Stopwatch.StartNew();
        string target = ValidateTarget(rawTarget);
        subnetRules ??= LoadSubnetRules();
        maxHops = Math.Clamp(maxHops, 1, 40);
        timeoutMs = Math.Clamp(timeoutMs, 300, 5000);

        AppLoggerService.Info($"TraceRouteAsync iniciado: alvo='{target}', maxHops={maxHops}, timeoutMs={timeoutMs}, subnetRules={subnetRules.Count}.", nameof(NetworkPathService));

        var result = new NetworkTraceResult
        {
            TargetInput = target,
            StartedAt = DateTime.Now
        };

        var origin = GetLocalOriginNode(subnetRules);
        result.Hops.Add(origin);
        onHopDiscovered?.Report(origin);

        // 1. Resolve destino (IP + FQDN + OU do AD caso offline)
        var swDns = Stopwatch.StartNew();
        var (targetIp, resolvedFqdn, adOuLocation) = await ResolveTargetMetadataAsync(target, cancellationToken);
        result.ResolvedTargetIp = targetIp?.ToString() ?? string.Empty;
        result.ResolvedTargetFqdn = resolvedFqdn;
        result.AdOuLocation = adOuLocation;
        AppLoggerService.Info(
            $"Resolução destino '{target}' em {AppLoggerService.FormatElapsed(swDns)}: IP='{(targetIp?.ToString() ?? "N/D")}', FQDN='{resolvedFqdn}', AD_OU='{adOuLocation}'.",
            nameof(NetworkPathService));

        if (targetIp == null)
        {
            var unresolvedDest = new NetworkHopNode
            {
                HopNumber = 1,
                Hostname = target,
                IpAddress = string.Empty,
                Kind = NetworkHopKind.Destination,
                Status = NetworkHopStatus.Unreachable,
                InferredLocation = !string.IsNullOrWhiteSpace(adOuLocation) ? adOuLocation : "DNS não resolvido (sem IP ativo)",
                ProbableDeviceLabel = "Destino Offline / DNS sem registro",
                Details = string.IsNullOrWhiteSpace(adOuLocation)
                    ? $"Não foi possível resolver '{target}' via DNS."
                    : $"DNS não resolveu '{target}'. Localização estimada via AD: {adOuLocation}"
            };
            result.Hops.Add(unresolvedDest);
            onHopDiscovered?.Report(unresolvedDest);
            result.CompletedAt = DateTime.Now;
            AppLoggerService.Warning($"TraceRouteAsync '{target}' encerrado: DNS não resolveu ({AppLoggerService.FormatElapsed(swTotal)}).", nameof(NetworkPathService));
            return result;
        }

        // 2. Varredura TTL 1..maxHops
        byte[] payload = Encoding.ASCII.GetBytes("AMBIENTESHUB-TRACE-PROBE-32BYTES");
        using var ping = new Ping();
        var pendingTimeouts = new List<NetworkHopNode>();
        bool reached = false;
        int consecutiveTimeouts = 0;
        int responsiveIntermediateHops = 0;

        for (int ttl = 1; ttl <= maxHops; ttl++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var options = new PingOptions(ttl, dontFragment: true);
            var sw = Stopwatch.StartNew();
            PingReply? reply = null;

            try
            {
                var pingTask = ping.SendPingAsync(targetIp, timeoutMs, payload, options);
                var done = await Task.WhenAny(pingTask, Task.Delay(timeoutMs + 250, cancellationToken));
                if (done == pingTask)
                    reply = await pingTask;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                reply = null;
            }

            sw.Stop();

            if (reply != null && (reply.Status == IPStatus.TtlExpired || reply.Status == IPStatus.Success))
            {
                // Descarrega eventuais timeouts intermediários (máx. 2) antes do próximo salto válido
                foreach (var tHop in pendingTimeouts.Take(2))
                {
                    result.Hops.Add(tHop);
                    onHopDiscovered?.Report(tHop);
                }
                pendingTimeouts.Clear();
                consecutiveTimeouts = 0;

                string hopIp = reply.Address?.ToString() ?? string.Empty;
                long rtt = reply.RoundtripTime > 0 ? reply.RoundtripTime : Math.Max(1, sw.ElapsedMilliseconds);
                bool isDestination = reply.Status == IPStatus.Success ||
                                     (reply.Address != null && reply.Address.Equals(targetIp));

                string ptr = await ResolvePtrWithTimeoutAsync(reply.Address, 550, cancellationToken);
                if (isDestination && string.IsNullOrWhiteSpace(ptr))
                    ptr = !string.IsNullOrWhiteSpace(resolvedFqdn) ? resolvedFqdn : target;

                var node = BuildHopNode(
                    hopNumber: ttl,
                    ipAddress: hopIp,
                    ptrHostname: ptr,
                    latencyMs: rtt,
                    ttlRemaining: reply.Options?.Ttl,
                    isDestination: isDestination,
                    isFirstHop: ttl == 1,
                    adOuLocation: isDestination ? adOuLocation : string.Empty,
                    subnetRules: subnetRules);

                if (isDestination)
                    node.ConnectedUser = await TryGetRemoteLoggedUserAsync(hopIp, cancellationToken);

                node.MacAddress = TryResolveMacViaArp(hopIp);

                result.Hops.Add(node);
                onHopDiscovered?.Report(node);

                if (isDestination)
                {
                    reached = true;
                    result.TargetOnline = true;
                    break;
                }

                responsiveIntermediateHops++;
            }
            else
            {
                consecutiveTimeouts++;
                pendingTimeouts.Add(new NetworkHopNode
                {
                    HopNumber = ttl,
                    IpAddress = string.Empty,
                    Hostname = "* * *",
                    Kind = NetworkHopKind.Timeout,
                    Status = NetworkHopStatus.Timeout,
                    InferredLocation = "Salto sem resposta ICMP (firewall ou host offline)",
                    ProbableDeviceLabel = "Timeout ICMP (* * *)",
                    Details = $"Salto {ttl}: sem resposta dentro de {timeoutMs} ms."
                });

                // Se já passamos por pelo menos 1 salto e acumulamos 3 timeouts seguidos (ou 4 no início),
                // consideramos que a rota parou antes do destino (ex.: máquina desligada após o gateway).
                int limit = responsiveIntermediateHops > 0 ? 3 : 4;
                if (consecutiveTimeouts >= limit)
                {
                    AppLoggerService.Info($"TraceRouteAsync '{target}' interrompido no TTL={ttl} após {consecutiveTimeouts} timeouts consecutivos.", nameof(NetworkPathService));
                    break;
                }
            }
        }

        if (!reached)
        {
            // Mostra até 1 salto de timeout antes do nó final de destino offline para sinalizar a ruptura da rota
            if (pendingTimeouts.Count > 0)
            {
                var lastBreak = pendingTimeouts[0];
                result.Hops.Add(lastBreak);
                onHopDiscovered?.Report(lastBreak);
            }

            string targetIpStr = targetIp.ToString();
            var matchedDest = MatchSubnetRule(targetIpStr, subnetRules);
            string destLoc = !string.IsNullOrWhiteSpace(matchedDest?.LocationName)
                ? matchedDest!.LocationName
                : (!string.IsNullOrWhiteSpace(adOuLocation)
                    ? adOuLocation
                    : InferIpScopeLabel(targetIp));

            var offlineDest = new NetworkHopNode
            {
                HopNumber = result.Hops.Count,
                IpAddress = targetIpStr,
                Hostname = !string.IsNullOrWhiteSpace(resolvedFqdn) ? resolvedFqdn : target,
                Kind = NetworkHopKind.Destination,
                Status = NetworkHopStatus.Unreachable,
                InferredLocation = destLoc,
                ProbableDeviceLabel = matchedDest != null && !string.IsNullOrWhiteSpace(matchedDest.ProbableSwitch)
                    ? $"Destino Offline • Switch provável: {matchedDest.ProbableSwitch}"
                    : "Destino Offline / Sem Resposta ICMP",
                MatchedCidr = matchedDest?.Cidr ?? string.Empty,
                Details = $"Último salto válido antes do destino: {(result.Hops.LastOrDefault(h => h.Kind != NetworkHopKind.Timeout)?.SubtitleDisplay ?? "nenhum")}. Destino {targetIpStr} não respondeu."
            };
            result.Hops.Add(offlineDest);
            onHopDiscovered?.Report(offlineDest);
        }

        // 3. Enriquecimento opcional via Zabbix API (se ZABBIX_API_URL e ZABBIX_API_TOKEN estiverem no .env)
        if (ZabbixEnrichmentService.IsConfigured())
        {
            var zabbix = new ZabbixEnrichmentService();
            var hopIps = result.Hops
                .Where(h => !string.IsNullOrWhiteSpace(h.IpAddress))
                .Select(h => h.IpAddress);
            var matches = await zabbix.LookupByIpsAsync(hopIps, cancellationToken: cancellationToken);
            ZabbixEnrichmentService.ApplyEnrichment(result.Hops, matches);
        }
        else
        {
            AppLoggerService.Info("Zabbix API não consultada: ZABBIX_API_URL ou ZABBIX_API_TOKEN ausente/inválido no .env.", nameof(ZabbixEnrichmentService));
        }

        result.CompletedAt = DateTime.Now;
        AppLoggerService.Success(
            $"TraceRouteAsync '{target}' finalizado: online={result.TargetOnline}, nós={result.Hops.Count}, ip='{result.ResolvedTargetIp}', duração={AppLoggerService.FormatElapsed(swTotal)}.",
            nameof(NetworkPathService));
        return result;
    }

    public static NetworkHopNode BuildHopNode(
        int hopNumber,
        string ipAddress,
        string ptrHostname,
        long latencyMs,
        int? ttlRemaining,
        bool isDestination,
        bool isFirstHop,
        string adOuLocation = "",
        IReadOnlyList<SubnetLocationRule>? subnetRules = null)
    {
        var matched = MatchSubnetRule(ipAddress, subnetRules);
        var kind = ClassifyHop(hopNumber, ipAddress, ptrHostname, isDestination, isFirstHop);
        var status = latencyMs > DefaultHighLatencyThresholdMs ? NetworkHopStatus.HighLatency : NetworkHopStatus.Online;

        string inferredLoc = matched != null && !string.IsNullOrWhiteSpace(matched.LocationName)
            ? matched.LocationName
            : InferLocationFromHostname(ptrHostname, adOuLocation, ipAddress);

        string deviceLabel = BuildDeviceLabel(kind, matched, isDestination);

        return new NetworkHopNode
        {
            HopNumber = hopNumber,
            IpAddress = ipAddress,
            Hostname = ptrHostname,
            LatencyMs = latencyMs,
            TtlRemaining = ttlRemaining,
            Kind = kind,
            Status = status,
            InferredLocation = inferredLoc,
            ProbableDeviceLabel = deviceLabel,
            MatchedCidr = matched?.Cidr ?? string.Empty,
            Details = $"Salto {hopNumber}: {(string.IsNullOrWhiteSpace(ptrHostname) ? ipAddress : ptrHostname)} ({ipAddress}) • {latencyMs} ms" +
                      (string.IsNullOrWhiteSpace(inferredLoc) ? string.Empty : $" • {inferredLoc}")
        };
    }

    public static NetworkHopKind ClassifyHop(int hopNumber, string ipAddress, string ptrHostname, bool isDestination, bool isFirstHop)
    {
        if (isDestination)
            return NetworkHopKind.Destination;

        string lower = (ptrHostname ?? string.Empty).ToLowerInvariant();
        if (lower.Contains("core") || lower.Contains("backbone") || lower.Contains("bb-") || lower.Contains("mpls"))
            return NetworkHopKind.Backbone;

        if (lower.Contains("sw") || lower.Contains("switch") || lower.Contains("l3"))
            return NetworkHopKind.SwitchL3;

        if (isFirstHop || hopNumber == 1 || lower.Contains("gw") || lower.Contains("gateway"))
            return NetworkHopKind.Gateway;

        return NetworkHopKind.Router;
    }

    public static string InferLocationFromHostname(string? ptrHostname, string? adOuLocation, string? ipAddress)
    {
        if (!string.IsNullOrWhiteSpace(adOuLocation))
            return adOuLocation;

        if (!string.IsNullOrWhiteSpace(ptrHostname))
        {
            string lower = ptrHostname.ToLowerInvariant();
            if (lower.Contains("wan") || lower.Contains("mpls") || lower.Contains("isp") || lower.Contains("operadora"))
                return "Enlace WAN / Operadora";
            if (lower.Contains("vpn"))
                return "Concentrador VPN";
            if (lower.Contains("core"))
                return "Núcleo / Datacenter";
            if (lower.Contains("fw") || lower.Contains("firewall"))
                return "Perímetro / Firewall";
        }

        if (IPAddress.TryParse(ipAddress, out var parsed))
            return InferIpScopeLabel(parsed);

        return string.Empty;
    }

    public static SubnetLocationRule? MatchSubnetRule(string? ipAddress, IReadOnlyList<SubnetLocationRule>? subnetRules)
    {
        if (subnetRules == null || subnetRules.Count == 0 || string.IsNullOrWhiteSpace(ipAddress))
            return null;

        if (!IPAddress.TryParse(ipAddress.Trim(), out var ip) || ip.AddressFamily != AddressFamily.InterNetwork)
            return null;

        return subnetRules
            .OrderByDescending(r => r.PrefixLength)
            .FirstOrDefault(r => r.Contains(ip));
    }

    public static string FormatTraceReport(NetworkTraceResult trace)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"=== RELATÓRIO DE ROTA DE REDE — {trace.TargetInput} ===");
        sb.AppendLine($"Início: {trace.StartedAt:yyyy-MM-dd HH:mm:ss}");
        if (!string.IsNullOrWhiteSpace(trace.ResolvedTargetIp))
            sb.AppendLine($"IP Resolvido: {trace.ResolvedTargetIp} ({trace.ResolvedTargetFqdn})");
        if (!string.IsNullOrWhiteSpace(trace.AdOuLocation))
            sb.AppendLine($"Localização AD: {trace.AdOuLocation}");
        sb.AppendLine($"Status Final: {(trace.TargetOnline ? "ONLINE" : "OFFLINE / INACESSÍVEL")}");
        sb.AppendLine();

        foreach (var hop in trace.Hops)
        {
            string lat = hop.LatencyDisplay;
            string host = string.IsNullOrWhiteSpace(hop.Hostname) ? "-" : hop.Hostname;
            string ip = string.IsNullOrWhiteSpace(hop.IpAddress) ? "* * *" : hop.IpAddress;
            string loc = string.IsNullOrWhiteSpace(hop.InferredLocation) ? "-" : hop.InferredLocation;
            sb.AppendLine($"[{hop.HopNumber,2}] {ip,-16} | {lat,-9} | {host,-32} | {hop.KindLabel} | {loc}");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Gera um relatório HTML5 responsivo (desktop e celular) com datas no padrão brasileiro (dd/MM/yyyy HH:mm:ss).
    /// </summary>
    public static string FormatTraceHtmlReport(NetworkTraceResult trace, NetworkPingStats? pingStats = null)
    {
        string E(string? val) => WebUtility.HtmlEncode(val ?? string.Empty);

        string statusText = trace.TargetOnline ? "ONLINE" : "OFFLINE / INACESSÍVEL";
        string statusBg = trace.TargetOnline ? "#064e3b" : "#7f1d1d";
        string statusFg = trace.TargetOnline ? "#34d399" : "#f87171";
        string startedBr = trace.StartedAt.ToString("dd/MM/yyyy HH:mm:ss");
        string completedBr = trace.CompletedAt?.ToString("dd/MM/yyyy HH:mm:ss") ?? "-";

        var hopsHtml = new StringBuilder();
        for (int i = 0; i < trace.Hops.Count; i++)
        {
            var hop = trace.Hops[i];
            bool isBroken = hop.Status is NetworkHopStatus.Timeout or NetworkHopStatus.Unreachable;
            string lineStyle = isBroken ? "border-left: 2px dashed #ef4444;" : "border-left: 2px solid #10b981;";

            hopsHtml.AppendLine($@"
            <div class=""hop-item"">
                <div class=""hop-card"" style=""border-color:{E(hop.BorderBrushHex)}"">
                    <div class=""hop-top"">
                        <span class=""hop-badge"">{E(hop.IconGlyph)} {E(hop.HopBadgeText)}</span>
                        <span class=""lat-badge"" style=""background:{E(hop.BadgeBackgroundHex)};color:{E(hop.StatusColorHex)}"">{E(hop.LatencyDisplay)}</span>
                    </div>
                    <div class=""hop-title"">{E(hop.TitleDisplay)}</div>
                    <div class=""hop-ip"">{E(hop.SubtitleDisplay)}{(!string.IsNullOrWhiteSpace(hop.MacAddress) ? $" • {E(hop.MacAddress)}" : "")}{(hop.TtlRemaining.HasValue ? $" • TTL {hop.TtlRemaining}" : "")}</div>
                    {(!string.IsNullOrWhiteSpace(hop.ConnectedUser) ? $@"<div class=""hop-user"">👤 {E(hop.ConnectedUser)}</div>" : "")}
                    <div class=""hop-loc"">📍 {E(hop.LocationDisplay)}</div>
                    <div class=""hop-kind"">{E(hop.KindLabel)}</div>
                </div>
                {(i < trace.Hops.Count - 1 ? $@"<div class=""connector"" style=""{lineStyle}""></div>" : "")}
            </div>");
        }

        string pingSection = string.Empty;
        if (pingStats != null && pingStats.Sent > 0)
        {
            pingSection = $@"
            <div class=""kpi-grid"">
                <div class=""kpi""><div class=""kpi-lbl"">Ping Enviados</div><div class=""kpi-val"">{pingStats.Sent}</div></div>
                <div class=""kpi""><div class=""kpi-lbl"">Recebidos</div><div class=""kpi-val"" style=""color:#34d399"">{pingStats.Received}</div></div>
                <div class=""kpi""><div class=""kpi-lbl"">Perda</div><div class=""kpi-val"" style=""color:{(pingStats.LossPercent > 0 ? "#f87171" : "#34d399")}"">{pingStats.LossPercent:0.#}%</div></div>
                <div class=""kpi""><div class=""kpi-lbl"">Média / Máx</div><div class=""kpi-val"" style=""color:#38bdf8"">{pingStats.AvgRttMs:0.#} / {(pingStats.MaxRttMs?.ToString() ?? "-")} ms</div></div>
            </div>";
        }

        return $@"<!DOCTYPE html>
<html lang=""pt-BR"">
<head>
<meta charset=""utf-8"" />
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />
<title>Relatório de Rota — {E(trace.TargetInput)}</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ margin:0; padding:16px; background:#0b0f19; color:#f8fafc; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; }}
  .container {{ max-width:760px; margin:0 auto; }}
  .header {{ background:#0f172a; border:1px solid #1e293b; border-radius:12px; padding:16px; margin-bottom:14px; }}
  .header-top {{ display:flex; justify-content:space-between; align-items:center; gap:10px; flex-wrap:wrap; }}
  h1 {{ margin:0; font-size:18px; font-weight:700; }}
  .status-pill {{ padding:5px 12px; border-radius:999px; font-size:12px; font-weight:700; background:{statusBg}; color:{statusFg}; }}
  .meta {{ margin-top:10px; font-size:12.5px; color:#94a3b8; line-height:1.55; }}
  .kpi-grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin-bottom:14px; }}
  .kpi {{ background:#0f172a; border:1px solid #1e293b; border-radius:10px; padding:10px 12px; }}
  .kpi-lbl {{ font-size:10.5px; color:#94a3b8; text-transform:uppercase; font-weight:600; }}
  .kpi-val {{ font-size:16px; font-weight:700; margin-top:2px; }}
  .hop-item {{ display:flex; flex-direction:column; align-items:center; }}
  .hop-card {{ width:100%; background:#0f172a; border:1.5px solid #334155; border-radius:10px; padding:12px 14px; }}
  .hop-top {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; }}
  .hop-badge {{ font-size:11px; font-weight:700; background:#1e293b; color:#cbd5e1; padding:3px 8px; border-radius:5px; }}
  .lat-badge {{ font-size:11px; font-weight:700; padding:3px 8px; border-radius:5px; }}
  .hop-title {{ font-size:14px; font-weight:600; word-break:break-word; }}
  .hop-ip {{ font-family:Consolas,monospace; font-size:12px; color:#94a3b8; margin-top:2px; }}
  .hop-user {{ font-size:12px; color:#34d399; font-weight:600; margin-top:4px; }}
  .hop-loc {{ font-size:12px; color:#38bdf8; margin-top:6px; }}
  .hop-kind {{ font-size:11px; color:#64748b; margin-top:2px; }}
  .connector {{ height:22px; width:0; margin:2px 0; }}
  @media (max-width:640px) {{
    body {{ padding:10px; }}
    .kpi-grid {{ grid-template-columns:repeat(2,1fr); }}
    h1 {{ font-size:16px; }}
  }}
</style>
</head>
<body>
  <div class=""container"">
    <div class=""header"">
      <div class=""header-top"">
        <h1>🌐 Rota de Rede: {E(trace.TargetInput)}</h1>
        <span class=""status-pill"">{statusText}</span>
      </div>
      <div class=""meta"">
        <div><strong>Início:</strong> {E(startedBr)} &nbsp;|&nbsp; <strong>Fim:</strong> {E(completedBr)}</div>
        <div><strong>IP Resolvido:</strong> {E(string.IsNullOrWhiteSpace(trace.ResolvedTargetIp) ? "-" : trace.ResolvedTargetIp)} {(!string.IsNullOrWhiteSpace(trace.ResolvedTargetFqdn) ? $"({E(trace.ResolvedTargetFqdn)})" : "")}</div>
        {(!string.IsNullOrWhiteSpace(trace.AdOuLocation) ? $"<div><strong>Localização AD:</strong> {E(trace.AdOuLocation)}</div>" : "")}
      </div>
    </div>
    {pingSection}
    <div class=""hops-list"">
      {hopsHtml}
    </div>
  </div>
</body>
</html>";
    }

    private static string BuildDeviceLabel(NetworkHopKind kind, SubnetLocationRule? matched, bool isDestination)
    {
        if (matched != null && !string.IsNullOrWhiteSpace(matched.ProbableSwitch))
        {
            return isDestination
                ? $"Destino • Switch: {matched.ProbableSwitch}"
                : $"Salto L3 • Switch: {matched.ProbableSwitch}";
        }

        return kind switch
        {
            NetworkHopKind.Gateway => "Gateway / SVI L3",
            NetworkHopKind.SwitchL3 => "Switch L3 / Distribuição",
            NetworkHopKind.Backbone => "Core / Backbone",
            NetworkHopKind.Destination => "Estação Destino",
            _ => "Roteador / Salto L3"
        };
    }

    private static string InferIpScopeLabel(IPAddress ip)
    {
        if (ip.AddressFamily != AddressFamily.InterNetwork)
            return "Rede IPv6";

        byte[] b = ip.GetAddressBytes();
        if (b[0] == 10 ||
            (b[0] == 172 && b[1] >= 16 && b[1] <= 31) ||
            (b[0] == 192 && b[1] == 168))
        {
            return $"Rede Privada ({b[0]}.{b[1]}.{b[2]}.0/24)";
        }

        if (b[0] == 100 && b[1] >= 64 && b[1] <= 127)
            return "CGNAT / VPN";

        if (IPAddress.IsLoopback(ip))
            return "Loopback Local";

        return "Rede Externa / Pública";
    }

    private static async Task<(IPAddress? Ip, string Fqdn, string AdOuLocation)> ResolveTargetMetadataAsync(
        string target,
        CancellationToken ct)
    {
        IPAddress? resolvedIp = null;
        string fqdn = string.Empty;
        string adLocation = string.Empty;

        if (IsStrictIpAddress(target, out var directIp) && directIp != null)
        {
            resolvedIp = directIp;
            fqdn = await ResolvePtrWithTimeoutAsync(directIp, 700, ct);
        }
        else
        {
            try
            {
                var dnsTask = Dns.GetHostEntryAsync(target);
                if (await Task.WhenAny(dnsTask, Task.Delay(1500, ct)) == dnsTask)
                {
                    var entry = await dnsTask;
                    fqdn = entry.HostName;
                    resolvedIp = entry.AddressList.FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork)
                              ?? entry.AddressList.FirstOrDefault();
                }
            }
            catch { }

            if (resolvedIp == null)
            {
                try
                {
                    var addrTask = Dns.GetHostAddressesAsync(target);
                    if (await Task.WhenAny(addrTask, Task.Delay(1200, ct)) == addrTask)
                    {
                        var addrs = await addrTask;
                        resolvedIp = addrs.FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork)
                                  ?? addrs.FirstOrDefault();
                    }
                }
                catch { }
            }

            // Fallback opcional ao Active Directory para descobrir OU/Localização mesmo se a estação estiver desligada
            adLocation = await Task.Run(() => TryLookupComputerLocationInAd(target), ct);
        }

        return (resolvedIp, fqdn, adLocation);
    }

    private static async Task<string> ResolvePtrWithTimeoutAsync(IPAddress? address, int timeoutMs, CancellationToken ct)
    {
        if (address == null)
            return string.Empty;

        try
        {
            var task = Dns.GetHostEntryAsync(address);
            if (await Task.WhenAny(task, Task.Delay(timeoutMs, ct)) == task)
            {
                var entry = await task;
                return entry.HostName ?? string.Empty;
            }
        }
        catch { }

        return string.Empty;
    }

    private static string TryLookupComputerLocationInAd(string hostName)
    {
        try
        {
            string shortName = hostName.Split('.')[0];
            if (!InputValidator.IsValidHostName(shortName))
                return string.Empty;

            using var searcher = new DirectorySearcher();
            string escaped = InputValidator.EscapeLdap(shortName);
            searcher.Filter = $"(&(objectCategory=computer)(|(cn={escaped})(dNSHostName={InputValidator.EscapeLdap(hostName)})))";
            searcher.PropertiesToLoad.AddRange(["distinguishedName", "location"]);
            searcher.SizeLimit = 1;
            searcher.ServerTimeLimit = TimeSpan.FromSeconds(2);

            var result = searcher.FindOne();
            if (result == null)
                return string.Empty;

            string loc = result.Properties.Contains("location") && result.Properties["location"].Count > 0
                ? result.Properties["location"][0]?.ToString() ?? string.Empty
                : string.Empty;

            string dn = result.Properties.Contains("distinguishedName") && result.Properties["distinguishedName"].Count > 0
                ? result.Properties["distinguishedName"][0]?.ToString() ?? string.Empty
                : string.Empty;

            var ous = dn.Split(',')
                .Select(p => p.Trim())
                .Where(p => p.StartsWith("OU=", StringComparison.OrdinalIgnoreCase))
                .Select(p => p[3..].Trim())
                .Take(2)
                .ToList();

            string ouSummary = ous.Count > 0 ? $"OU: {string.Join(" / ", ous)}" : string.Empty;

            if (!string.IsNullOrWhiteSpace(loc) && !string.IsNullOrWhiteSpace(ouSummary))
                return $"{loc} ({ouSummary})";

            return !string.IsNullOrWhiteSpace(loc) ? loc : ouSummary;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"TryLookupComputerLocationInAd falhou para '{hostName}': {ex.Message}", nameof(NetworkPathService));
            return string.Empty;
        }
    }

    private static async Task<string> TryGetRemoteLoggedUserAsync(string targetIpOrHost, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(targetIpOrHost))
            return string.Empty;

        try
        {
            var wmiTask = Task.Run(() =>
            {
                var options = new ConnectionOptions { Timeout = TimeSpan.FromSeconds(2) };
                var scope = new ManagementScope($@"\\{targetIpOrHost}\root\cimv2", options);
                scope.Connect();
                using var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT UserName FROM Win32_ComputerSystem"));
                using var results = searcher.Get();
                foreach (ManagementObject obj in results)
                {
                    string user = obj["UserName"]?.ToString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(user))
                        return user;
                }
                return string.Empty;
            }, ct);

            if (await Task.WhenAny(wmiTask, Task.Delay(2200, ct)) == wmiTask)
                return await wmiTask;

            AppLoggerService.Warning($"Consulta WMI de usuário em '{targetIpOrHost}' excedeu o timeout (2.2s).", nameof(NetworkPathService));
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Não foi possível obter usuário conectado em '{targetIpOrHost}' via WMI: {ex.Message}", nameof(NetworkPathService));
        }

        return string.Empty;
    }

    [DllImport("iphlpapi.dll", ExactSpelling = true)]
    private static extern int SendARP(uint destIp, uint srcIp, byte[] macAddr, ref int phyAddrLen);

    /// <summary>
    /// Resolve o MAC Address via tabela ARP do Windows (apenas para IPs IPv4 na mesma sub-rede L2).
    /// </summary>
    public static string TryResolveMacViaArp(string? ipAddress)
    {
        if (string.IsNullOrWhiteSpace(ipAddress))
            return string.Empty;

        if (!IPAddress.TryParse(ipAddress.Trim(), out var ip) || ip.AddressFamily != AddressFamily.InterNetwork)
            return string.Empty;

        try
        {
            byte[] ipBytes = ip.GetAddressBytes();
            uint destIp = BitConverter.ToUInt32(ipBytes, 0);
            byte[] macAddr = new byte[6];
            int len = macAddr.Length;
            int res = SendARP(destIp, 0, macAddr, ref len);
            if (res == 0 && len == 6)
                return string.Join(":", macAddr.Select(b => b.ToString("X2")));
        }
        catch { }

        return string.Empty;
    }

    private static string FormatPhysicalAddress(PhysicalAddress? address)
    {
        if (address == null)
            return string.Empty;

        byte[] bytes = address.GetAddressBytes();
        return bytes.Length == 6
            ? string.Join(":", bytes.Select(b => b.ToString("X2")))
            : string.Empty;
    }

    public static async Task<bool> TryEnrichDestinationFromAdSccmAsync(
        string target,
        NetworkHopNode destNode,
        CancellationToken ct)
    {
        try
        {
            string shortHost = target.Split('.')[0].Trim();
            if (!InputValidator.IsValidHostName(shortHost))
                return false;

            var adSccm = new AdSccmService();
            var searchTask = adSccm.SearchMachinesAsync(shortHost);
            if (await Task.WhenAny(searchTask, Task.Delay(15000, ct)) != searchTask)
                return false;

            var list = await searchTask;
            var match = list.FirstOrDefault(m =>
                string.Equals(m.ComputerName, shortHost, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(m.DnsHostName, target, StringComparison.OrdinalIgnoreCase));

            if (match == null)
                return false;

            if (string.IsNullOrWhiteSpace(destNode.ConnectedUser))
            {
                string user = !string.IsNullOrWhiteSpace(match.CurrentLogonUser)
                    ? match.CurrentLogonUser
                    : match.PrimaryUser;
                if (!string.IsNullOrWhiteSpace(user))
                    destNode.ConnectedUser = $"{user} (SCCM)";
            }

            if (string.IsNullOrWhiteSpace(destNode.MacAddress) && !string.IsNullOrWhiteSpace(match.MacAddress))
                destNode.MacAddress = match.MacAddress;

            if (!string.IsNullOrWhiteSpace(match.EstimatedLocationDisplay) &&
                !match.EstimatedLocationDisplay.Equals("Localidade não identificada", StringComparison.OrdinalIgnoreCase))
            {
                destNode.InferredLocation = match.EstimatedLocationDisplay;
            }
            AppLoggerService.Info($"Destino '{target}' enriquecido via AD/SCCM fallback (Usuário='{destNode.ConnectedUser}', MAC='{destNode.MacAddress}').", nameof(NetworkPathService));
            return !string.IsNullOrWhiteSpace(destNode.ConnectedUser) || !string.IsNullOrWhiteSpace(destNode.MacAddress);
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"TryEnrichDestinationFromAdSccmAsync falhou para '{target}': {ex.Message}", nameof(NetworkPathService));
            return false;
        }
    }
}
