using System.Diagnostics;
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
        var sw = Stopwatch.StartNew();
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
            AppLoggerService.Info(
                $"Configurações carregadas de '{envFile}' ({Configs.Count} chaves válidas, {rejected} rejeitadas) em {AppLoggerService.FormatElapsed(sw)}.",
                nameof(ConfigService));
        }
        else
        {
            AppLoggerService.Warning(
                $"Arquivo .env não localizado (caminho esperado: '{envFile}'); funcionalidades dependentes permanecem desabilitadas (fail-closed) [{AppLoggerService.FormatElapsed(sw)}].",
                nameof(ConfigService));
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
            AppLoggerService.Warning($"Não foi possível restringir permissões de '{Path.GetFileName(path)}' ({ex.GetType().Name}): {ex.Message}", nameof(ConfigService));
            return false;
        }
    }

    public static void SaveConfig(Dictionary<string, string> newConfigs)
    {
        var sw = Stopwatch.StartNew();
        string envFile = GetEnvFilePath();
        var updatedKeys = new List<string>();
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
                AppLoggerService.Warning($"Chave de configuração inválida '{kvp.Key}' descartada ao salvar .env.", nameof(ConfigService));
                continue;
            }

            string safeValue = InputValidator.SanitizeEnvValue(kvp.Value);
            Configs[kvp.Key] = safeValue;
            updatedKeys.Add(kvp.Key);
        }

        // Persiste o dicionário consolidado para evitar perda de outras configurações ao salvar parcialmente.
        foreach (var kvp in Configs)
        {
            if (!InputValidator.IsValidEnvKey(kvp.Key))
                continue;

            lines.Add($"{kvp.Key}={kvp.Value}");
        }

        File.WriteAllLines(envFile, lines);
        bool aclOk = HardenFilePermissions(envFile);
        AppLoggerService.Success(
            $"Arquivo .env atualizado em '{envFile}': {updatedKeys.Count} chave(s) alterada(s) [{string.Join(", ", updatedKeys)}], total={Configs.Count}, ACL600={aclOk}, duração={AppLoggerService.FormatElapsed(sw)}.",
            nameof(ConfigService));
    }

    /// <summary>Atualiza e persiste uma única chave no .env mantendo todas as demais configurações intactas.</summary>
    public static void SetAndSave(string key, string value)
    {
        SaveConfig(new Dictionary<string, string> { [key] = value });
    }

    public static string Get(string key, string defaultValue = "")
    {
        return Configs.TryGetValue(key, out var val) ? val : defaultValue;
    }

    public static string AdminUser => Get("ADMIN_USER", "");
    public static string AdminAdUser => Get("ADMIN_AD_USER", "");
    public static string UrlOtrs => Get("URL_OTRS", "");
    public static string UrlCitsmart => Get("URL_CITSMART", "");
    public static string UrlOxe => Get("URL_OXE", "");
    public static string UrlPapercut => Get("URL_PAPERCUT", "");
    public static string UrlBancadaLocal => Get("URL_BANCADA_LOCAL", "");
    public static string BancadaApiUrl => Get("BANCADA_API_URL", "");
    public static string BancadaApiKey => Get("BANCADA_API_KEY", "");
    public static string BancadaWindowsDir => Get("BANCADA_WINDOWS_DIR", "");
    public static string BancadaWslDir => Get("BANCADA_WSL_DIR", "");
    public static string BancadaWslDistro => Get("BANCADA_WSL_DISTRO", "");
    public static string WhatsappPwaArgs => Get("WHATSAPP_PWA_ARGS", "");
    public static string SccmSiteServer => Get("SCCM_SITE_SERVER", "");
    public static string SccmSiteCode => Get("SCCM_SITE_CODE", "");
    public static string RsatMsuNetworkPath => Get("RSAT_MSU_NETWORK_PATH", "");

    public static string DeviceReportsDir => Get("DEVICE_REPORTS_DIR", "");
    public static string ChamadosReportsDir => Get("CHAMADOS_REPORTS_DIR", "");
    public static string CmRcViewerCustomPath => Get("CMRCVIEWER_PATH", "");
    public static string FirefoxBinaryPath => Get("FIREFOX_PATH", "");
    public static string UsmtPath => Get("USMT_PATH", "");
    public static string BancadaDbHost => Get("BANCADA_DB_HOST", "127.0.0.1");
    public static int BancadaDbPort => int.TryParse(Get("BANCADA_DB_PORT", "5432"), out int p) && p > 0 ? p : 5432;
    public static string BancadaDbName => Get("BANCADA_DB_NAME", "");
    public static string BancadaDbUser => Get("BANCADA_DB_USER", "");
    public static string BancadaDbPassword => Get("BANCADA_DB_PASSWORD", "");
    public static string ProtectedProfiles => Get("PROTECTED_PROFILES", "");
    public static int PingTimeoutMs => int.TryParse(Get("PING_TIMEOUT_MS", "2500"), out int t) && t > 100 ? t : 2500;
    public static bool MinimizeToTray => bool.TryParse(Get("MINIMIZE_TO_TRAY", "true"), out bool min) ? min : true;
    public static string ZabbixApiUrl => Get("ZABBIX_API_URL", "");
    public static string ZabbixApiToken => Get("ZABBIX_API_TOKEN", "");
    public static string SajmpExePath => Get("SAJMP_EXE_PATH", "");
    public static string SajmpCredPath => Get("SAJMP_CRED_PATH", "");
    public static string SajmpWindowTitle => Get("SAJMP_WINDOW_TITLE", "");

    /// <summary>
    /// Retorna valores padrão recomendados e dinâmicos (ex: perfil do usuário atual do Windows)
    /// sem expor dados pessoais fixos no código-fonte.
    /// </summary>
    public static string GetRecommendedDefault(string key)
    {
        string user = Environment.UserName;
        return key switch
        {
            "APP_TERMINAL_FONT" => "Consolas",
            "ADMIN_USER" => !string.IsNullOrWhiteSpace(user) ? user : "admin_local",
            "ADMIN_AD_USER" => !string.IsNullOrWhiteSpace(user) ? $"{user}_ad" : "admin_ad",
            "SCCM_SITE_SERVER" => "srv-sccm.dominio.local",
            "SCCM_SITE_CODE" => "XXX",
            "URL_OTRS" => "https://central.empresa.gov.br/",
            "URL_CITSMART" => "https://suporte.empresa.gov.br/",
            "URL_OXE" => "https://oxe.empresa.local/",
            "URL_PAPERCUT" => "http://impressora.empresa.local:9191/admin",
            "URL_BANCADA_LOCAL" => "http://localhost:8501/",
            "BANCADA_API_URL" => "http://127.0.0.1:8510",
            "BANCADA_WSL_DIR" => $"/home/{(!string.IsNullOrWhiteSpace(user) ? user.ToLowerInvariant() : "usuario")}/projetos/bancada",
            "BANCADA_WSL_DISTRO" => "Ubuntu-26.04",
            "BANCADA_WINDOWS_DIR" => $@"\\wsl.localhost\Ubuntu-26.04\home\{(!string.IsNullOrWhiteSpace(user) ? user.ToLowerInvariant() : "usuario")}\projetos\bancada",
            "WHATSAPP_PWA_ARGS" => "\"-taskbar-tab\" \"uuid\" \"-new-window\" \"https://web.whatsapp.com/\" \"-container\" \"0\"",
            "DEVICE_REPORTS_DIR" => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "DeviceReports"),
            "CHAMADOS_REPORTS_DIR" => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "ChamadosReports"),
            "CMRCVIEWER_PATH" => @"C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\i386\CmRcViewer.exe",
            "FIREFOX_PATH" => @"C:\Program Files\Mozilla Firefox\firefox.exe",
            "USMT_PATH" => @"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\User State Migration Tool\amd64",
            "BANCADA_DB_HOST" => "127.0.0.1",
            "BANCADA_DB_PORT" => "5432",
            "PROTECTED_PROFILES" => "Public, Default, Default User, All Users, Administrator, Administrador",
            "PING_TIMEOUT_MS" => "2500",
            "ZABBIX_API_URL" => "https://zabbix.empresa.local/api_jsonrpc.php",
            _ => string.Empty
        };
    }
}
