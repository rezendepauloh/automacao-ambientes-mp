using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class RemoteToolsService
{
    private static readonly string[] SccmStandardPaths =
    [
        @"C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe",
        @"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"
    ];

    public string? GetCmRcViewerPath()
    {
        // 1. Procura na pasta local da aplicação
        string localAppPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CmRcViewer.exe");
        if (File.Exists(localAppPath)) return localAppPath;

        // 2. Procura nos caminhos padrão do console SCCM oficial instalado no Windows
        foreach (var path in SccmStandardPaths)
        {
            if (File.Exists(path))
                return path;
        }

        // 3. Fallback opcional no diretório D:\ArrumarPC caso exista
        string legacyPath = @"D:\ArrumarPC\CmRcViewer.exe";
        if (File.Exists(legacyPath)) return legacyPath;

        return null;
    }

    public void OpenSccmRemoteControl(string computerName, string? siteServer = null)
    {
        if (string.IsNullOrWhiteSpace(computerName))
            throw new ArgumentException("Nome do computador não informado.");

        siteServer ??= ConfigService.SccmSiteServer;
        var cmPath = GetCmRcViewerPath();

        if (cmPath != null)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = cmPath,
                Arguments = $"{computerName.Trim()} \\\\{siteServer.Trim()}",
                UseShellExecute = true
            });
        }
        else
        {
            // Se a máquina nova não tiver o executável do SCCM, cai graciosamente para o RDP oficial
            OpenRdp(computerName);
        }
    }

    public void OpenRdp(string computerName)
    {
        if (string.IsNullOrWhiteSpace(computerName))
            throw new ArgumentException("Nome do computador não informado.");

        Process.Start(new ProcessStartInfo
        {
            FileName = "mstsc.exe",
            Arguments = $"/v:{computerName.Trim()}",
            UseShellExecute = true
        });
    }
}
