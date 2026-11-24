namespace AmbientesHub.Models;

public enum RdpQualityProfile
{
    Balanced = 0,
    MaxStability = 1,
    BestVisual = 2
}

public readonly record struct RdpQualityParameters(
    int ColorDepth,
    uint NetworkConnectionType,
    int PerformanceFlags,
    bool BitmapCaching,
    bool Compression);

public class RdpServerEntry
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = string.Empty;
    public string HostOrUrl { get; set; } = string.Empty;
    public int Port { get; set; } = 3389;
    public string UserName { get; set; } = string.Empty;
    public string Domain { get; set; } = string.Empty;
    public string Group { get; set; } = "Geral";
    public bool OpenEmbedded { get; set; } = true;
    public bool PromptCredentials { get; set; } = true;
    public int AuthenticationLevel { get; set; } = 0;
    public bool UsesHardDrives { get; set; } = false;
    public bool AudioCaptureRedirection { get; set; } = true;
    public RdpQualityProfile QualityProfile { get; set; } = RdpQualityProfile.Balanced;
    public bool AutoAdaptResolution { get; set; } = true;
    public bool UseMultiMonitorInFullscreen { get; set; } = false;
    public string Notes { get; set; } = string.Empty;

    public static RdpQualityParameters GetProfileParameters(RdpQualityProfile profile) => profile switch
    {
        // 16-bit, Broadband Low (2), desativa wallpaper/drag/anim/themes/cursor shadow (0x2F)
        RdpQualityProfile.MaxStability => new(16, 2u, 0x0000002F, BitmapCaching: true, Compression: true),
        // 32-bit, LAN (6), habilita ClearType (0x80) + Desktop Composition (0x100)
        RdpQualityProfile.BestVisual => new(32, 6u, 0x00000180, BitmapCaching: true, Compression: true),
        // 24-bit, Broadband High/WAN (4), desativa wallpaper/drag/anim (0x07) mas mantém ClearType nítido (0x80)
        _ => new(24, 4u, 0x00000087, BitmapCaching: true, Compression: true)
    };

    public bool IsSeguraConfigured =>
        (!string.IsNullOrWhiteSpace(UserName) && UserName.Contains('[', StringComparison.Ordinal)) ||
        HostOrUrl.Contains("segura", StringComparison.OrdinalIgnoreCase);

    public string BadgeText => IsSeguraConfigured ? "🔐 SEGURA" : "🖥️ DIRETO";

    public string BadgeColorHex => IsSeguraConfigured ? "#34d399" : "#38bdf8";

    public string BadgeBgHex => IsSeguraConfigured ? "#064e3b" : "#0c4a6e";

    public string BorderHex => IsSeguraConfigured ? "#10b981" : "#334155";

    public string SubtitleDisplay => Port != 3389
        ? $"{HostOrUrl}:{Port}"
        : HostOrUrl;

    public string UserDisplay => !string.IsNullOrWhiteSpace(UserName)
        ? UserName
        : (PromptCredentials ? "Solicitar credenciais ao conectar" : "Usuário não definido");
}

public class RdpCatalogDocument
{
    public string DatabaseId { get; set; } = Guid.NewGuid().ToString();
    public int Version { get; set; } = 2;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
    public List<RdpServerEntry> Servers { get; set; } = [];
}
