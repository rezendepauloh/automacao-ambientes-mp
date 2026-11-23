namespace AmbientesHub.Models;

public class SubnetNeighborNode
{
    public string Hostname { get; set; } = string.Empty;
    public string IpAddress { get; set; } = string.Empty;
    public string MacAddress { get; set; } = string.Empty;
    public string ConnectedUser { get; set; } = string.Empty;
    public string LocationOrOu { get; set; } = string.Empty;
    public long? LatencyMs { get; set; }
    public bool IsOnline { get; set; }
    public bool IsReferenceHost { get; set; }
    public string SourceTag { get; set; } = "ICMP+DNS";

    public string TitleDisplay => !string.IsNullOrWhiteSpace(Hostname)
        ? Hostname
        : (!string.IsNullOrWhiteSpace(IpAddress) ? IpAddress : "Host Desconhecido");

    public string SubtitleDisplay
    {
        get
        {
            if (string.IsNullOrWhiteSpace(IpAddress))
                return string.IsNullOrWhiteSpace(MacAddress) ? "Sem IP ativo" : MacAddress;
            return string.IsNullOrWhiteSpace(MacAddress) ? IpAddress : $"{IpAddress} • {MacAddress}";
        }
    }

    public string LatencyDisplay => IsOnline
        ? (LatencyMs.HasValue ? $"{LatencyMs.Value} ms" : "ONLINE")
        : "SEM PING";

    public string BadgeText => IsReferenceHost ? "⭐ REFERÊNCIA" : SourceTag;

    public string BorderBrushHex => IsReferenceHost
        ? "#38bdf8"
        : (IsOnline ? "#10b981" : "#475569");

    public string StatusColorHex => IsOnline ? "#34d399" : "#94a3b8";

    public string BadgeBackgroundHex => IsReferenceHost
        ? "#0c4a6e"
        : (IsOnline ? "#064e3b" : "#1e293b");
}

public class SubnetNeighborsResult
{
    public string ReferenceInput { get; set; } = string.Empty;
    public string ReferenceIp { get; set; } = string.Empty;
    public string ReferenceHostname { get; set; } = string.Empty;
    public string SubnetCidr { get; set; } = string.Empty;
    public string GatewayIp { get; set; } = string.Empty;
    public string SwitchOrZabbixLabel { get; set; } = string.Empty;
    public string AdOuPath { get; set; } = string.Empty;
    public DateTime ScannedAt { get; set; } = DateTime.Now;
    public List<SubnetNeighborNode> Neighbors { get; set; } = [];
}
