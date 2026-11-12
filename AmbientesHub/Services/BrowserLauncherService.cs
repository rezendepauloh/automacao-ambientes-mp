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
        if (!string.IsNullOrWhiteSpace(custom) && File.Exists(custom))
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
        try
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string profilesDir = Path.Combine(appData, @"Mozilla\Firefox\Profiles");

            if (Directory.Exists(profilesDir))
            {
                foreach (var profile in Directory.GetDirectories(profilesDir))
                {
                    // Apaga sessionstore.jsonlz4 e a pasta sessionstore-backups
                    string sessionFile = Path.Combine(profile, "sessionstore.jsonlz4");
                    if (File.Exists(sessionFile))
                    {
                        File.Delete(sessionFile);
                    }

                    string backupsDir = Path.Combine(profile, "sessionstore-backups");
                    if (Directory.Exists(backupsDir))
                    {
                        foreach (var f in Directory.GetFiles(backupsDir))
                        {
                            try { File.Delete(f); } catch { }
                        }
                    }
                }
                AppLoggerService.Info("Arquivos de sessão anterior do Firefox limpos com sucesso.", nameof(BrowserLauncherService));
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Aviso ao limpar sessão do Firefox: {ex.Message}", nameof(BrowserLauncherService));
        }
    }

    public void OpenUrlsInFirefox(IEnumerable<string> urls)
    {
        var urlList = urls.Where(u => !string.IsNullOrWhiteSpace(u)).ToList();
        if (urlList.Count == 0) return;

        string firefoxPath = GetFirefoxPath();

        // Abre cada URL como uma nova aba na janela já existente do Firefox (sem forçar nova janela isolada)
        foreach (var url in urlList)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = firefoxPath,
                    Arguments = $"\"{url}\"",
                    UseShellExecute = true
                });
            }
            catch
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{url}\"",
                    UseShellExecute = true
                });
            }
            Thread.Sleep(200);
        }
    }

    public void OpenWhatsappPwa()
    {
        string firefoxPath = GetFirefoxPath();
        string args = ConfigService.WhatsappPwaArgs;

        AppLoggerService.Info($"Iniciando WhatsApp PWA: {firefoxPath} {args}", nameof(BrowserLauncherService));

        Process.Start(new ProcessStartInfo
        {
            FileName = firefoxPath,
            Arguments = args,
            UseShellExecute = true
        });
    }
}
