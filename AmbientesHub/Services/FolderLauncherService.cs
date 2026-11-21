using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class FolderLauncherService
{
    public void OpenFolders(IEnumerable<string> folderPaths)
    {
        var paths = folderPaths?.Where(p => !string.IsNullOrWhiteSpace(p)).ToList() ?? [];
        if (paths.Count == 0)
        {
            AppLoggerService.Info("Nenhuma pasta informada para abertura no Explorer.", nameof(FolderLauncherService));
            return;
        }

        var sw = Stopwatch.StartNew();
        int opened = 0;
        int missing = 0;
        int failed = 0;
        AppLoggerService.Info($"Iniciando abertura de {paths.Count} pasta(s) no Explorer.", nameof(FolderLauncherService));

        foreach (var path in paths)
        {
            string expanded = Environment.ExpandEnvironmentVariables(path.Trim());
            if (!Directory.Exists(expanded))
            {
                missing++;
                AppLoggerService.Warning($"Diretório não encontrado: '{expanded}'.", nameof(FolderLauncherService));
                continue;
            }

            try
            {
                var proc = Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{expanded}\"",
                    UseShellExecute = true
                });
                opened++;
                AppLoggerService.Info($"Pasta aberta no Explorer: '{expanded}' (PID={proc?.Id.ToString() ?? "shell"}).", nameof(FolderLauncherService));
                Thread.Sleep(350);
            }
            catch (Exception ex)
            {
                failed++;
                AppLoggerService.LogException(nameof(FolderLauncherService), $"Abrir pasta '{expanded}' no Explorer", ex, level: LogLevel.WARNING);
            }
        }

        var level = failed > 0 ? LogLevel.WARNING : LogLevel.SUCCESS;
        AppLoggerService.Log(
            level,
            $"Abertura de pastas concluída: Abertas={opened}, Ausentes={missing}, Falhas={failed}, Duração={AppLoggerService.FormatElapsed(sw)}.",
            nameof(FolderLauncherService));
    }
}

