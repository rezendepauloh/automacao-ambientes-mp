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

    public string Location { get; set; } = string.Empty;
    public string UserDepartment { get; set; } = string.Empty;
    public string UserOffice { get; set; } = string.Empty;

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

    public string OnlineBadgeBackground => IsOnline ? "#064e3b" : "#7f1d1d";

    public string OnlineBadgeForeground => IsOnline ? "#34d399" : "#f87171";

    public string CardBorderBrush => IsOnline ? "#10b981" : "#ef4444";

    public string CurrentLogonUserDisplay => !string.IsNullOrWhiteSpace(CurrentLogonUser)
        ? CurrentLogonUser
        : (!string.IsNullOrWhiteSpace(PrimaryUser) ? $"{PrimaryUser} (Primário)" : "Nenhum usuário ativo");

    public string PrimaryUserDisplay => !string.IsNullOrWhiteSpace(PrimaryUser) ? PrimaryUser : "Não definido";

    public string IpWithLatency => !string.IsNullOrEmpty(PingResponseTime)
        ? $"{IpAddress} ({PingResponseTime})"
        : (!string.IsNullOrEmpty(IpAddress) ? IpAddress : "Sem resposta ICMP");

    public string MacAddressDisplay => !string.IsNullOrWhiteSpace(MacAddress) ? MacAddress : "Não informado";

    public string EstimatedLocationDisplay
    {
        get
        {
            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(Location))
                parts.Add(Location.Trim());

            if (!string.IsNullOrWhiteSpace(UserDepartment))
                parts.Add(UserDepartment.Trim());

            if (!string.IsNullOrWhiteSpace(UserOffice) &&
                !string.Equals(UserOffice.Trim(), UserDepartment.Trim(), StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(UserOffice.Trim(), Location.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                parts.Add(UserOffice.Trim());
            }

            if (parts.Count > 0)
                return string.Join(" • ", parts);

            // Fallback por Unidade Organizacional (OU) do computador no AD
            if (!string.IsNullOrWhiteSpace(DistinguishedName))
            {
                string ou = ExtractOuSummary(DistinguishedName);
                if (!string.IsNullOrWhiteSpace(ou))
                    return $"OU: {ou}";
            }

            return "Localidade não identificada";
        }
    }

    private static string ExtractOuSummary(string dn)
    {
        var ous = dn.Split(',')
            .Select(p => p.Trim())
            .Where(p => p.StartsWith("OU=", StringComparison.OrdinalIgnoreCase))
            .Select(p => p[3..].Trim())
            .Take(2)
            .ToList();

        return ous.Count > 0 ? string.Join(" / ", ous) : string.Empty;
    }
}
