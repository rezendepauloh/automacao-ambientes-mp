namespace AmbientesHub.Models;

public class MachineDetails
{
    public string ComputerName { get; set; } = string.Empty;
    public string DnsHostName { get; set; } = string.Empty;
    public string OperatingSystem { get; set; } = string.Empty;
    public string DistinguishedName { get; set; } = string.Empty;
    public string WhenCreated { get; set; } = string.Empty;

    public string PrimaryUser { get; set; } = string.Empty;
    public string CurrentLogonUser { get; set; } = string.Empty;
    public string MacAddress { get; set; } = string.Empty;

    public bool FoundInAd { get; set; } = false;
    public bool FoundInSccm { get; set; } = false;
    public string ResourceId { get; set; } = string.Empty;
    public string SmsId { get; set; } = string.Empty;
    public string ClientStatus { get; set; } = "Não localizado";
    public string SiteCode { get; set; } = string.Empty;

    public bool IsOnline { get; set; } = false;
    public string IpAddress { get; set; } = string.Empty;
    public string PingResponseTime { get; set; } = string.Empty;

    public string OnlineBadgeText => IsOnline 
        ? $"🟢 ONLINE • {IpAddress} ({PingResponseTime})" 
        : (string.IsNullOrEmpty(IpAddress) ? "🔴 INACESSÍVEL / OFFLINE" : $"🔴 OFFLINE • IP: {IpAddress}");

    public string CurrentLogonUserDisplay => !string.IsNullOrWhiteSpace(CurrentLogonUser)
        ? CurrentLogonUser
        : (!string.IsNullOrWhiteSpace(PrimaryUser) ? $"{PrimaryUser} (Primário)" : "Nenhum usuário ativo");

    public string PrimaryUserDisplay => !string.IsNullOrWhiteSpace(PrimaryUser) ? PrimaryUser : "Não definido";

    public string IpWithLatency => !string.IsNullOrEmpty(PingResponseTime)
        ? $"{IpAddress} ({PingResponseTime})"
        : (!string.IsNullOrEmpty(IpAddress) ? IpAddress : "Sem resposta ICMP");

    public string MacAddressDisplay => !string.IsNullOrWhiteSpace(MacAddress) ? MacAddress : "Não informado";
}
