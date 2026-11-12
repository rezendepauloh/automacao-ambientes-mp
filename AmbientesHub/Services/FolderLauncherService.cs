using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class FolderLauncherService
{
    public void OpenFolders(IEnumerable<string> folderPaths)
    {
        foreach (var path in folderPaths)
        {
            if (string.IsNullOrWhiteSpace(path)) continue;

            string expanded = Environment.ExpandEnvironmentVariables(path.Trim());
            if (Directory.Exists(expanded))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{expanded}\"",
                    UseShellExecute = true
                });
                AppLoggerService.Info($"Pasta aberta no Explorer: {expanded}", nameof(FolderLauncherService));
                Thread.Sleep(350);
            }
            else
            {
                AppLoggerService.Warning($"Diretório não encontrado: {expanded}", nameof(FolderLauncherService));
            }
        }
    }
}

