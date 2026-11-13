using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;

namespace AmbientesHub.Services;

public static class ConfigService
{
    private static readonly Dictionary<string, string> Configs = new(StringComparer.OrdinalIgnoreCase);

    static ConfigService()
    {
        LoadEnv();
    }

    public static string GetEnvFilePath()
    {
        string[] possiblePaths =
        [
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".env"),
            @"C:\AmbientesVirtuais\AmbientesHub\.env",
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".env")
        ];

        return possiblePaths.FirstOrDefault(File.Exists) ?? @"C:\AmbientesVirtuais\AmbientesHub\.env";
    }

    public static void LoadEnv()
    {
        Configs.Clear();
        string envFile = GetEnvFilePath();

        if (File.Exists(envFile))
        {
            HardenFilePermissions(envFile);
            int rejected = 0;
            foreach (var line in File.ReadAllLines(envFile))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#"))
                    continue;

                int idx = trimmed.IndexOf('=');
                if (idx > 0)
                {
                    string key = trimmed[..idx].Trim();
                    string val = InputValidator.SanitizeEnvValue(trimmed[(idx + 1)..]);
                    if (!InputValidator.IsValidEnvKey(key))
                    {
                        rejected++;
                        continue;
                    }
                    Configs[key] = val;
                }
            }
            AppLoggerService.Info($"Configurações carregadas do .env ({Configs.Count} chaves, {rejected} rejeitadas)", nameof(ConfigService));
        }
        else
        {
            AppLoggerService.Warning("Arquivo .env não localizado; funcionalidades dependentes de configuração permanecem desabilitadas (fail-closed).", nameof(ConfigService));
        }
    }

    /// <summary>
    /// Restringe o ACL do arquivo ao usuário atual, Administradores e SYSTEM (equivalente a chmod 600 no Windows).
    /// Melhor esforço: falhas são registradas, nunca propagadas.
    /// </summary>
    public static bool HardenFilePermissions(string path)
    {
        try
        {
            if (!File.Exists(path)) return false;

            var security = new FileSecurity();
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);

            var current = WindowsIdentity.GetCurrent().User;
            if (current != null)
                security.AddAccessRule(new FileSystemAccessRule(current, FileSystemRights.FullControl, AccessControlType.Allow));
            security.AddAccessRule(new FileSystemAccessRule(
                new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null), FileSystemRights.FullControl, AccessControlType.Allow));
            security.AddAccessRule(new FileSystemAccessRule(
                new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null), FileSystemRights.FullControl, AccessControlType.Allow));

            new FileInfo(path).SetAccessControl(security);
            return true;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Não foi possível restringir permissões do arquivo de configuração: {ex.GetType().Name}", nameof(ConfigService));
            return false;
        }
    }

    public static void SaveConfig(Dictionary<string, string> newConfigs)
    {
        string envFile = GetEnvFilePath();
        var lines = new List<string>
        {
            "# ==========================================",
            "# CONFIGURAÇÕES DO AMBIENTES HUB",
            $"# Atualizado em: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
            "# =========================================="
        };

        foreach (var kvp in newConfigs)
        {
            // Fail-closed: chave inválida é descartada; valor perde CR/LF (impede injeção de novas chaves).
            if (!InputValidator.IsValidEnvKey(kvp.Key))
            {
                AppLoggerService.Warning("Chave de configuração inválida descartada ao salvar .env.", nameof(ConfigService));
                continue;
            }

            string safeValue = InputValidator.SanitizeEnvValue(kvp.Value);
            Configs[kvp.Key] = safeValue;
            lines.Add($"{kvp.Key}={safeValue}");
        }

        File.WriteAllLines(envFile, lines);
        HardenFilePermissions(envFile);
        AppLoggerService.Success("Arquivo .env atualizado com sucesso.", nameof(ConfigService));
    }

    public static string Get(string key, string defaultValue = "")
    {
        return Configs.TryGetValue(key, out var val) ? val : defaultValue;
    }

    public static string AdminUser => Get("ADMIN_USER", "admin_local");
    public static string AdminAdUser => Get("ADMIN_AD_USER", "admin_dominio");
    public static string UrlOtrs => Get("URL_OTRS", "");
    public static string UrlCitsmart => Get("URL_CITSMART", "");
    public static string UrlBancadaLocal => Get("URL_BANCADA_LOCAL", "http://localhost:8501/");
    public static string BancadaWindowsDir => Get("BANCADA_WINDOWS_DIR", "");
    public static string BancadaWslDir => Get("BANCADA_WSL_DIR", "");
    public static string BancadaWslDistro => Get("BANCADA_WSL_DISTRO", "");
    public static string WhatsappPwaArgs => Get("WHATSAPP_PWA_ARGS", "");
    public static string SccmSiteServer => Get("SCCM_SITE_SERVER", "");
    public static string SccmSiteCode => Get("SCCM_SITE_CODE", "");
    public static string RsatMsuNetworkPath => Get("RSAT_MSU_NETWORK_PATH", "");

    public static string DeviceReportsDir => Get("DEVICE_REPORTS_DIR", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "DeviceReports"));
    public static string CmRcViewerCustomPath => Get("CMRCVIEWER_PATH", "");
    public static string FirefoxBinaryPath => Get("FIREFOX_PATH", "");
    public static string ProtectedProfiles => Get("PROTECTED_PROFILES", "Public, Default, Default User, All Users, Administrator, Administrador");
    public static int PingTimeoutMs => int.TryParse(Get("PING_TIMEOUT_MS", "2500"), out int t) && t > 100 ? t : 2500;
    public static bool MinimizeToTray => bool.TryParse(Get("MINIMIZE_TO_TRAY", "true"), out bool min) ? min : true;
}
