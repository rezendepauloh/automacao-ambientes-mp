using System.Net;
using System.Net.Sockets;
using System.Text.Json.Serialization;

namespace AmbientesHub.Models;

public enum NetworkHopKind
{
    LocalStation,
    Gateway,
    SwitchL3,
    Router,
    Backbone,
    Destination,
    Timeout
}

public enum NetworkHopStatus
{
    Online,
    HighLatency,
    Timeout,
    Unreachable
}

/// <summary>
/// Regra local que associa uma faixa CIDR IPv4 a uma localização física e switch provável.
/// </summary>
public class SubnetLocationRule
{
    public string Cidr { get; set; } = string.Empty;
    public string LocationName { get; set; } = string.Empty;
    public string ProbableSwitch { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;

    [JsonIgnore]
    public int PrefixLength => TryParseCidr(Cidr, out _, out _, out int prefix) ? prefix : -1;

    public bool Contains(IPAddress? address)
    {
        if (address == null || address.AddressFamily != AddressFamily.InterNetwork)
            return false;

        if (!TryParseCidr(Cidr, out uint network, out uint mask, out _))
            return false;

        uint ipUint = ToUInt32(address);
        return (ipUint & mask) == network;
    }

    public static bool TryParseCidr(string? cidr, out uint network, out uint mask, out int prefixLength)
    {
        network = 0;
        mask = 0;
        prefixLength = 0;

        if (string.IsNullOrWhiteSpace(cidr))
            return false;

        var parts = cidr.Trim().Split('/');
        if (parts.Length != 2)
            return false;

        if (!parts[0].Contains('.', StringComparison.Ordinal))
            return false;

        if (!IPAddress.TryParse(parts[0], out var ip) || ip.AddressFamily != AddressFamily.InterNetwork)
            return false;

        if (!int.TryParse(parts[1], out prefixLength) || prefixLength < 0 || prefixLength > 32)
            return false;

        mask = prefixLength == 0 ? 0u : uint.MaxValue << (32 - prefixLength);
        network = ToUInt32(ip) & mask;
        return true;
    }

    private static uint ToUInt32(IPAddress ip)
    {
        byte[] bytes = ip.GetAddressBytes();
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToUInt32(bytes, 0);
    }
}

/// <summary>
/// Representa um nó (estação local, salto L3 ou destino) na árvore visual de rota.
/// </summary>
public class NetworkHopNode
{
    public int HopNumber { get; set; }
    public string IpAddress { get; set; } = string.Empty;
    public string Hostname { get; set; } = string.Empty;
    public long? LatencyMs { get; set; }
    public int? TtlRemaining { get; set; }
    public NetworkHopKind Kind { get; set; } = NetworkHopKind.Router;
    public NetworkHopStatus Status { get; set; } = NetworkHopStatus.Online;
    public string InferredLocation { get; set; } = string.Empty;
    public string ProbableDeviceLabel { get; set; } = string.Empty;
    public string MatchedCidr { get; set; } = string.Empty;
    public string ConnectedUser { get; set; } = string.Empty;
    public string Details { get; set; } = string.Empty;

    public string HopBadgeText => Kind switch
    {
        NetworkHopKind.LocalStation => "ORIGEM",
        NetworkHopKind.Destination => "DESTINO",
        _ => $"SALTO {HopNumber}"
    };

    public string IconGlyph => Kind switch
    {
        NetworkHopKind.LocalStation => "💻",
        NetworkHopKind.Gateway => "🚪",
        NetworkHopKind.SwitchL3 => "🔀",
        NetworkHopKind.Router => "🌐",
        NetworkHopKind.Backbone => "🛰️",
        NetworkHopKind.Destination => Status == NetworkHopStatus.Online || Status == NetworkHopStatus.HighLatency ? "🖥️" : "🔌",
        NetworkHopKind.Timeout => "❓",
        _ => "🔹"
    };

    public string KindLabel => !string.IsNullOrWhiteSpace(ProbableDeviceLabel)
        ? ProbableDeviceLabel
        : Kind switch
        {
            NetworkHopKind.LocalStation => "Estação Local",
            NetworkHopKind.Gateway => "Gateway Padrão",
            NetworkHopKind.SwitchL3 => "Switch L3 / Core",
            NetworkHopKind.Router => "Roteador / Salto L3",
            NetworkHopKind.Backbone => "Backbone / WAN",
            NetworkHopKind.Destination => Status is NetworkHopStatus.Online or NetworkHopStatus.HighLatency
                ? "Estação Destino (Online)"
                : "Estação Destino (Offline)",
            NetworkHopKind.Timeout => "Sem Resposta ICMP",
            _ => "Nó de Rede"
        };

    public string TitleDisplay => !string.IsNullOrWhiteSpace(Hostname)
        ? Hostname
        : (!string.IsNullOrWhiteSpace(IpAddress) ? IpAddress : "* * * (Timeout)");

    public string SubtitleDisplay => !string.IsNullOrWhiteSpace(IpAddress)
        ? IpAddress
        : "Sem IP / Timeout ICMP";

    public string LatencyDisplay => Status switch
    {
        NetworkHopStatus.Timeout => "* * *",
        NetworkHopStatus.Unreachable => "OFFLINE",
        _ => LatencyMs.HasValue ? $"{LatencyMs.Value} ms" : "OK"
    };

    public string LocationDisplay => !string.IsNullOrWhiteSpace(InferredLocation)
        ? InferredLocation
        : (Kind == NetworkHopKind.Timeout ? "Rota sem resposta ICMP" : "Localização não mapeada");

    public string StatusColorHex => Status switch
    {
        NetworkHopStatus.Online => "#34d399",
        NetworkHopStatus.HighLatency => "#fbbf24",
        NetworkHopStatus.Timeout => "#f87171",
        NetworkHopStatus.Unreachable => "#ef4444",
        _ => "#38bdf8"
    };

    public string BadgeBackgroundHex => Status switch
    {
        NetworkHopStatus.Online => "#064e3b",
        NetworkHopStatus.HighLatency => "#78350f",
        NetworkHopStatus.Timeout => "#450a0a",
        NetworkHopStatus.Unreachable => "#7f1d1d",
        _ => "#1e293b"
    };

    public string BorderBrushHex => Status switch
    {
        NetworkHopStatus.Online => "#10b981",
        NetworkHopStatus.HighLatency => "#f59e0b",
        NetworkHopStatus.Timeout => "#b91c1c",
        NetworkHopStatus.Unreachable => "#ef4444",
        _ => "#334155"
    };
}

/// <summary>
/// Estatísticas acumuladas do monitoramento contínuo ou pontual de Ping.
/// </summary>
public class NetworkPingStats
{
    private const int MaxHistoryItems = 24;
    private long _rttSum;

    public int Sent { get; private set; }
    public int Received { get; private set; }
    public int Lost => Math.Max(0, Sent - Received);
    public long? LastRttMs { get; private set; }
    public long? MinRttMs { get; private set; }
    public long? MaxRttMs { get; private set; }
    public double AvgRttMs => Received > 0 ? Math.Round((double)_rttSum / Received, 1) : 0;
    public double LossPercent => Sent > 0 ? Math.Round((double)Lost * 100.0 / Sent, 1) : 0;
    public bool IsReachable => LastRttMs.HasValue;
    public List<long?> History { get; } = [];

    public void RecordSample(long? rttMs)
    {
        Sent++;
        LastRttMs = rttMs;

        if (rttMs.HasValue)
        {
            long val = Math.Max(0, rttMs.Value);
            Received++;
            _rttSum += val;
            MinRttMs = MinRttMs.HasValue ? Math.Min(MinRttMs.Value, val) : val;
            MaxRttMs = MaxRttMs.HasValue ? Math.Max(MaxRttMs.Value, val) : val;
        }

        History.Add(rttMs);
        if (History.Count > MaxHistoryItems)
            History.RemoveAt(0);
    }

    public void Reset()
    {
        Sent = 0;
        Received = 0;
        _rttSum = 0;
        LastRttMs = null;
        MinRttMs = null;
        MaxRttMs = null;
        History.Clear();
    }
}

public class NetworkTraceResult
{
    public string TargetInput { get; set; } = string.Empty;
    public string ResolvedTargetIp { get; set; } = string.Empty;
    public string ResolvedTargetFqdn { get; set; } = string.Empty;
    public string AdOuLocation { get; set; } = string.Empty;
    public bool TargetOnline { get; set; }
    public DateTime StartedAt { get; set; } = DateTime.Now;
    public DateTime? CompletedAt { get; set; }
    public List<NetworkHopNode> Hops { get; set; } = [];
}
