using System.IO;

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
            foreach (var line in File.ReadAllLines(envFile))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#"))
                    continue;

                int idx = trimmed.IndexOf('=');
                if (idx > 0)
                {
                    string key = trimmed[..idx].Trim();
                    string val = trimmed[(idx + 1)..].Trim();
                    Configs[key] = val;
                }
            }
            AppLoggerService.Info($"Configurações carregadas do .env ({Configs.Count} chaves)", nameof(ConfigService));
        }
        else
        {
            AppLoggerService.Warning("Arquivo .env não localizado; usando valores padrão.", nameof(ConfigService));
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
            Configs[kvp.Key] = kvp.Value;
            lines.Add($"{kvp.Key}={kvp.Value}");
        }

        File.WriteAllLines(envFile, lines);
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
    public static string WhatsappPwaArgs => Get("WHATSAPP_PWA_ARGS", "");
    public static string SccmSiteServer => Get("SCCM_SITE_SERVER", "");
    public static string SccmSiteCode => Get("SCCM_SITE_CODE", "PGJ");
    public static string RsatMsuNetworkPath => Get("RSAT_MSU_NETWORK_PATH", @"\\manutencao\Manutencao\Dados\_UTILITARIOS\AD Interface\WIN 11\WindowsTH-RSAT_WS2016-x64 - funciona.msu");
}
