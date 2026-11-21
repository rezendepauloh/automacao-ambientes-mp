using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class BrowserLauncherService
{
    private static readonly string[] FirefoxLocations =
    [
        @"C:\Program Files\Mozilla Firefox\firefox.exe",
        @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Mozilla Firefox\firefox.exe")
    ];

    public string GetFirefoxPath()
    {
        string custom = ConfigService.FirefoxBinaryPath;
        if (InputValidator.IsSafeLocalExecutablePath(custom) && File.Exists(custom))
            return custom;

        foreach (var path in FirefoxLocations)
        {
            if (File.Exists(path))
                return path;
        }
        return "firefox.exe";
    }

    public void ClearFirefoxSession()
    {
        var sw = Stopwatch.StartNew();
        try
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string profilesDir = Path.Combine(appData, @"Mozilla\Firefox\Profiles");

            if (Directory.Exists(profilesDir))
            {
                int profilesCount = 0;
                int deletedFiles = 0;
                foreach (var profile in Directory.GetDirectories(profilesDir))
                {
                    profilesCount++;
                    // Apaga sessionstore.jsonlz4 e a pasta sessionstore-backups
                    string sessionFile = Path.Combine(profile, "sessionstore.jsonlz4");
                    if (File.Exists(sessionFile))
                    {
                        File.Delete(sessionFile);
                        deletedFiles++;
                    }

                    string backupsDir = Path.Combine(profile, "sessionstore-backups");
                    if (Directory.Exists(backupsDir))
                    {
                        foreach (var f in Directory.GetFiles(backupsDir))
                        {
                            try { File.Delete(f); deletedFiles++; } catch { }
                        }
                    }
                }
                AppLoggerService.Info(
                    $"Sessão anterior do Firefox limpa ({profilesCount} perfil(is), {deletedFiles} arquivo(s) removido(s)) em {AppLoggerService.FormatElapsed(sw)}.",
                    nameof(BrowserLauncherService));
            }
            else
            {
                AppLoggerService.Info($"Diretório de perfis do Firefox não encontrado ('{profilesDir}'); limpeza de sessão ignorada.", nameof(BrowserLauncherService));
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BrowserLauncherService), "Limpar sessão do Firefox", ex, sw.Elapsed, LogLevel.WARNING);
        }
    }

    public void OpenUrlsInFirefox(IEnumerable<string> urls)
    {
        var sw = Stopwatch.StartNew();
        // Fail-closed: somente http/https sem credenciais embutidas; demais esquemas (file:, javascript:, caminhos) são descartados.
        var urlList = new List<string>();
        int rejectedCount = 0;
        foreach (var raw in urls.Where(u => !string.IsNullOrWhiteSpace(u)))
        {
            string candidate = raw.Trim();
            if (InputValidator.IsSafeHttpUrl(candidate))
                urlList.Add(candidate);
            else
            {
                rejectedCount++;
                AppLoggerService.Warning($"URL rejeitada pela validação (apenas http/https são permitidos): '{candidate}'.", nameof(BrowserLauncherService));
            }
        }
        if (urlList.Count == 0)
        {
            AppLoggerService.Warning($"OpenUrlsInFirefox sem URLs válidas para abrir (rejeitadas={rejectedCount}).", nameof(BrowserLauncherService));
            return;
        }

        string firefoxPath = GetFirefoxPath();
        AppLoggerService.Info($"Abrindo {urlList.Count} URL(s) no Firefox ('{firefoxPath}'): {string.Join(", ", urlList)}", nameof(BrowserLauncherService));

        // Abre cada URL como uma nova aba na janela já existente do Firefox (sem forçar nova janela isolada)
        foreach (var url in urlList)
        {
            try
            {
                var info = new ProcessStartInfo { FileName = firefoxPath, UseShellExecute = true };
                info.ArgumentList.Add(url);
                Process.Start(info);
            }
            catch (Exception ex)
            {
                AppLoggerService.Warning($"Firefox falhou ao abrir '{url}' ({ex.Message}); usando fallback explorer.exe.", nameof(BrowserLauncherService));
                var info = new ProcessStartInfo { FileName = "explorer.exe", UseShellExecute = true };
                info.ArgumentList.Add(url);
                Process.Start(info);
            }
            Thread.Sleep(200);
        }

        AppLoggerService.Success($"OpenUrlsInFirefox concluído ({urlList.Count} URL(s)) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BrowserLauncherService));
    }

    public void OpenWhatsappPwa()
    {
        var sw = Stopwatch.StartNew();
        string firefoxPath = GetFirefoxPath();
        string args = ConfigService.WhatsappPwaArgs;

        if (!IsSafeBrowserArguments(args))
        {
            AppLoggerService.Warning("WHATSAPP_PWA_ARGS ausente ou contém caracteres não permitidos. Execução bloqueada (fail-closed).", nameof(BrowserLauncherService));
            return;
        }

        // Argumentos (que podem conter caminhos de perfil) não são gravados no log.
        AppLoggerService.Info($"Iniciando WhatsApp PWA via Firefox ('{firefoxPath}').", nameof(BrowserLauncherService));

        try
        {
            var proc = Process.Start(new ProcessStartInfo
            {
                FileName = firefoxPath,
                Arguments = args,
                UseShellExecute = true
            });
            AppLoggerService.Success($"WhatsApp PWA disparado (PID={proc?.Id}) em {AppLoggerService.FormatElapsed(sw)}.", nameof(BrowserLauncherService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BrowserLauncherService), "Iniciar WhatsApp PWA", ex, sw.Elapsed);
            throw;
        }
    }

    /// <summary>Whitelist de argumentos de navegador: rejeita metacaracteres de shell/controle.</summary>
    public static bool IsSafeBrowserArguments(string? args)
    {
        if (string.IsNullOrWhiteSpace(args) || args.Length > 1024) return false;
        return args.IndexOfAny(['&', '|', ';', '<', '>', '^', '%', '`', '$', '\r', '\n', '\0']) < 0;
    }
}
