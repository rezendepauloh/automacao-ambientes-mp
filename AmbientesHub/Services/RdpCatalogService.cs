using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Gerencia o catálogo local de servidores RDP (compatível com exportação/importação JSON do RDM),
/// persistido com criptografia Windows DPAPI (CurrentUser) + ACL 600 dentro de AmbientesHub\.
/// </summary>
public class RdpCatalogService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static string GetDefaultCatalogFilePath()
    {
        string envDir = Path.GetDirectoryName(ConfigService.GetEnvFilePath()) ?? AppDomain.CurrentDomain.BaseDirectory;
        return Path.Combine(envDir, "rdp_servers.dat");
    }

    public static RdpCatalogDocument LoadEncryptedCatalog(string? filePath = null)
    {
        string path = string.IsNullOrWhiteSpace(filePath) ? GetDefaultCatalogFilePath() : filePath;
        try
        {
            if (!File.Exists(path))
                return new RdpCatalogDocument();

            byte[] encrypted = File.ReadAllBytes(path);
            if (encrypted.Length == 0)
                return new RdpCatalogDocument();

            byte[] plainBytes = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
            string json = Encoding.UTF8.GetString(plainBytes);
            return JsonSerializer.Deserialize<RdpCatalogDocument>(json, JsonOptions) ?? new RdpCatalogDocument();
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao carregar catálogo RDP criptografado '{path}': {ex.Message}", nameof(RdpCatalogService));
            return new RdpCatalogDocument();
        }
    }

    public static void SaveEncryptedCatalog(RdpCatalogDocument document, string? filePath = null)
    {
        string path = string.IsNullOrWhiteSpace(filePath) ? GetDefaultCatalogFilePath() : filePath;
        string? dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        document.UpdatedAt = DateTime.Now;
        string json = JsonSerializer.Serialize(document, JsonOptions);
        byte[] plainBytes = Encoding.UTF8.GetBytes(json);
        byte[] encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);

        File.WriteAllBytes(path, encrypted);
        ConfigService.HardenFilePermissions(path);
        AppLoggerService.Info($"Catálogo RDP salvo com DPAPI ({document.Servers.Count} servidor(es)) em '{path}'.", nameof(RdpCatalogService));
    }

    /// <summary>
    /// Importa conexões a partir do JSON exportado pelo Devolutions Remote Desktop Manager (RDM).
    /// </summary>
    public static RdpCatalogDocument ImportFromRdmJson(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("Arquivo JSON vazio.");

        using var doc = JsonDocument.Parse(json);
        var catalog = new RdpCatalogDocument();

        if (doc.RootElement.TryGetProperty("DatabaseID", out var dbIdProp) && dbIdProp.ValueKind == JsonValueKind.String)
            catalog.DatabaseId = dbIdProp.GetString() ?? catalog.DatabaseId;

        if (doc.RootElement.TryGetProperty("Version", out var verProp) && verProp.ValueKind == JsonValueKind.Number)
            catalog.Version = verProp.GetInt32();

        if (!doc.RootElement.TryGetProperty("Connections", out var connections) || connections.ValueKind != JsonValueKind.Array)
            throw new ArgumentException("JSON não contém o array 'Connections' esperado do RDM.");

        foreach (var item in connections.EnumerateArray())
        {
            string name = GetStringProp(item, "Name");
            string url = GetStringProp(item, "Url");
            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(url))
                continue;

            string userName = string.Empty;
            string domain = string.Empty;
            int port = 3389;
            bool audioCapture = true;

            if (item.TryGetProperty("RDP", out var rdpObj) && rdpObj.ValueKind == JsonValueKind.Object)
            {
                userName = GetStringProp(rdpObj, "UserName");
                domain = GetStringProp(rdpObj, "Domain");
                if (rdpObj.TryGetProperty("Port", out var pProp) && pProp.ValueKind == JsonValueKind.Number)
                    port = Math.Clamp(pProp.GetInt32(), 1, 65535);
                if (rdpObj.TryGetProperty("AudioCaptureRedirectionMode", out var acProp) &&
                    (acProp.ValueKind == JsonValueKind.True || acProp.ValueKind == JsonValueKind.False))
                    audioCapture = acProp.GetBoolean();
            }

            string group = GetStringProp(item, "Group");
            if (string.IsNullOrWhiteSpace(group))
            {
                group = (!string.IsNullOrWhiteSpace(userName) && userName.Contains('[', StringComparison.Ordinal)) ||
                        url.Contains("segura", StringComparison.OrdinalIgnoreCase)
                    ? "Segura PAM"
                    : "Servidores";
            }

            int screenColor = GetIntProp(item, "ScreenColor", 2);
            var quality = item.TryGetProperty("QualityProfile", out var qpProp) && qpProp.ValueKind == JsonValueKind.Number
                ? (RdpQualityProfile)Math.Clamp(qpProp.GetInt32(), 0, 2)
                : screenColor switch
                {
                    0 or 1 => RdpQualityProfile.MaxStability,
                    3 or 4 => RdpQualityProfile.Balanced,
                    _ => RdpQualityProfile.Balanced
                };

            catalog.Servers.Add(new RdpServerEntry
            {
                Id = GetStringProp(item, "ID", Guid.NewGuid().ToString()),
                Name = string.IsNullOrWhiteSpace(name) ? url : name,
                HostOrUrl = url,
                Port = port,
                UserName = userName,
                Domain = domain,
                Group = group,
                OpenEmbedded = GetBoolProp(item, "OpenEmbedded", true),
                PromptCredentials = GetBoolProp(item, "PromptCredentials", true),
                AuthenticationLevel = GetIntProp(item, "AuthentificationLevel", 0),
                UsesHardDrives = GetBoolProp(item, "UsesHardDrives", false),
                AudioCaptureRedirection = audioCapture,
                QualityProfile = quality,
                AutoAdaptResolution = GetBoolProp(item, "AutoAdaptResolution", true),
                UseMultiMonitorInFullscreen = GetBoolProp(item, "UseMultiMonitorInFullscreen", false)
            });
        }

        catalog.Servers = catalog.Servers
            .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return catalog;
    }

    /// <summary>
    /// Exporta o catálogo no mesmo formato JSON estruturado compatível com o RDM (sem senhas).
    /// </summary>
    public static string ExportToRdmJson(RdpCatalogDocument catalog)
    {
        var connections = catalog.Servers.Select(s => new Dictionary<string, object?>
        {
            ["AppVersion"] = "2026.2",
            ["ConnectionType"] = 1,
            ["ID"] = s.Id,
            ["Name"] = s.Name,
            ["Group"] = s.Group,
            ["OpenEmbedded"] = s.OpenEmbedded,
            ["PromptCredentials"] = s.PromptCredentials,
            ["AuthentificationLevel"] = s.AuthenticationLevel,
            ["ScreenColor"] = s.QualityProfile switch
            {
                RdpQualityProfile.MaxStability => 1,
                RdpQualityProfile.BestVisual => 3,
                _ => 2
            },
            ["QualityProfile"] = (int)s.QualityProfile,
            ["AutoAdaptResolution"] = s.AutoAdaptResolution,
            ["UseMultiMonitorInFullscreen"] = s.UseMultiMonitorInFullscreen,
            ["Url"] = s.HostOrUrl,
            ["UsesHardDrives"] = s.UsesHardDrives,
            ["RDP"] = new Dictionary<string, object?>
            {
                ["AudioCaptureRedirectionMode"] = s.AudioCaptureRedirection,
                ["UserName"] = string.IsNullOrWhiteSpace(s.UserName) ? null : s.UserName,
                ["Domain"] = string.IsNullOrWhiteSpace(s.Domain) ? null : s.Domain,
                ["Port"] = s.Port != 3389 ? s.Port : null
            }
            .Where(kv => kv.Value != null)
            .ToDictionary(kv => kv.Key, kv => kv.Value)
        }).ToList();

        var exportObj = new
        {
            Connections = connections,
            DatabaseID = catalog.DatabaseId,
            Version = catalog.Version
        };

        return JsonSerializer.Serialize(exportObj, JsonOptions);
    }

    /// <summary>
    /// Preenche Host/Proxy e UserName no formato PAM Segura®: pamUser[domain\adAdminUser@SERVIDOR].
    /// </summary>
    public static void ApplySeguraTemplate(
        RdpServerEntry entry,
        string pamProxyHost,
        string pamUser,
        string domain,
        string adAdminUser)
    {
        string validatedProxy = NetworkPathService.ValidateTarget(pamProxyHost);
        string serverToken = ExtractServerToken(entry);

        entry.HostOrUrl = validatedProxy;
        entry.UserName = $"{pamUser.Trim()}[{domain.Trim()}\\{adAdminUser.Trim()}@{serverToken}]";
        entry.Group = "Segura PAM";
        entry.PromptCredentials = true;
    }

    public static string ExtractServerToken(RdpServerEntry entry)
    {
        var match = Regex.Match(entry.Name ?? string.Empty, @"\b(SRV-[A-Za-z0-9\-]+)\b", RegexOptions.IgnoreCase);
        if (match.Success)
            return match.Groups[1].Value.ToUpperInvariant();

        if (!string.IsNullOrWhiteSpace(entry.HostOrUrl) &&
            !entry.HostOrUrl.Contains("segura", StringComparison.OrdinalIgnoreCase))
        {
            return entry.HostOrUrl.Split('.')[0].Trim().ToUpperInvariant();
        }

        string firstPart = (entry.Name ?? string.Empty).Split(' ')[0].Trim();
        return string.IsNullOrWhiteSpace(firstPart) ? "SERVIDOR" : firstPart.ToUpperInvariant();
    }

    public static string BuildRdpFileContent(RdpServerEntry entry)
    {
        string safeHost = NetworkPathService.ValidateTarget(entry.HostOrUrl);
        int safePort = Math.Clamp(entry.Port, 1, 65535);
        string fullAddress = safePort == 3389 ? safeHost : $"{safeHost}:{safePort}";

        static string CleanLine(string? v) => (v ?? string.Empty).Replace("\r", "").Replace("\n", "").Replace("\0", "").Trim();

        var qp = RdpServerEntry.GetProfileParameters(entry.QualityProfile);
        bool disableWallpaper = (qp.PerformanceFlags & 0x01) != 0;
        bool disableFullWinDrag = (qp.PerformanceFlags & 0x02) != 0;
        bool disableMenuAnims = (qp.PerformanceFlags & 0x04) != 0;
        bool disableThemes = (qp.PerformanceFlags & 0x08) != 0;
        bool fontSmoothing = (qp.PerformanceFlags & 0x80) != 0;
        bool desktopComp = (qp.PerformanceFlags & 0x100) != 0;

        var sb = new StringBuilder();
        sb.AppendLine($"full address:s:{fullAddress}");
        if (!string.IsNullOrWhiteSpace(entry.UserName))
            sb.AppendLine($"username:s:{CleanLine(entry.UserName)}");
        if (!string.IsNullOrWhiteSpace(entry.Domain))
            sb.AppendLine($"domain:s:{CleanLine(entry.Domain)}");
        sb.AppendLine($"session bpp:i:{qp.ColorDepth}");
        sb.AppendLine($"connection type:i:{qp.NetworkConnectionType}");
        sb.AppendLine($"compression:i:{(qp.Compression ? 1 : 0)}");
        sb.AppendLine($"bitmapcachepersistenable:i:{(qp.BitmapCaching ? 1 : 0)}");
        sb.AppendLine($"disable wallpaper:i:{(disableWallpaper ? 1 : 0)}");
        sb.AppendLine($"disable full window drag:i:{(disableFullWinDrag ? 1 : 0)}");
        sb.AppendLine($"disable menu anims:i:{(disableMenuAnims ? 1 : 0)}");
        sb.AppendLine($"disable themes:i:{(disableThemes ? 1 : 0)}");
        sb.AppendLine($"allow font smoothing:i:{(fontSmoothing ? 1 : 0)}");
        sb.AppendLine($"allow desktop composition:i:{(desktopComp ? 1 : 0)}");
        sb.AppendLine($"authentication level:i:{Math.Clamp(entry.AuthenticationLevel, 0, 2)}");
        sb.AppendLine($"prompt for credentials:i:{(entry.PromptCredentials ? 1 : 0)}");
        sb.AppendLine($"redirectdrives:i:{(entry.UsesHardDrives ? 1 : 0)}");
        sb.AppendLine($"audiocapturemode:i:{(entry.AudioCaptureRedirection ? 1 : 0)}");
        sb.AppendLine("smart sizing:i:1");
        sb.AppendLine($"dynamic resolution:i:{(entry.AutoAdaptResolution ? 1 : 0)}");
        sb.AppendLine($"use multimon:i:{(entry.UseMultiMonitorInFullscreen ? 1 : 0)}");
        sb.AppendLine("screen mode id:i:2");
        return sb.ToString();
    }

    public static void LaunchExternalMstsc(RdpServerEntry entry)
    {
        string content = BuildRdpFileContent(entry);
        string tempFile = Path.Combine(Path.GetTempPath(), $"ambhub_{Guid.NewGuid():N}.rdp");
        File.WriteAllText(tempFile, content, Encoding.UTF8);

        var psi = new ProcessStartInfo
        {
            FileName = "mstsc.exe",
            UseShellExecute = true
        };
        psi.ArgumentList.Add(tempFile);
        Process.Start(psi);
        AppLoggerService.Info($"Sessão RDP externa (mstsc.exe) iniciada para '{entry.Name}' ({entry.HostOrUrl}).", nameof(RdpCatalogService));
    }

    private static string GetStringProp(JsonElement obj, string propName, string defaultValue = "")
    {
        if (obj.TryGetProperty(propName, out var prop) && prop.ValueKind == JsonValueKind.String)
            return prop.GetString() ?? defaultValue;
        return defaultValue;
    }

    private static bool GetBoolProp(JsonElement obj, string propName, bool defaultValue)
    {
        if (obj.TryGetProperty(propName, out var prop) &&
            (prop.ValueKind == JsonValueKind.True || prop.ValueKind == JsonValueKind.False))
            return prop.GetBoolean();
        return defaultValue;
    }

    private static int GetIntProp(JsonElement obj, string propName, int defaultValue)
    {
        if (obj.TryGetProperty(propName, out var prop) && prop.ValueKind == JsonValueKind.Number)
            return prop.GetInt32();
        return defaultValue;
    }
}
