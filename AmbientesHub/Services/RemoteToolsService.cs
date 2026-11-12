using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class RemoteToolsService
{
    private static readonly string[] SccmStandardPaths =
    [
        @"C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\i386\CmRcViewer.exe",
        @"C:\Program Files\Microsoft Endpoint Manager\AdminConsole\bin\i386\CmRcViewer.exe",
        @"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe",
        @"C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe",
        @"C:\Program Files (x86)\ConfigMgr Console\bin\i386\CmRcViewer.exe",
        @"C:\Program Files\ConfigMgr Console\bin\i386\CmRcViewer.exe",
        @"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\CmRcViewer.exe",
        @"C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\CmRcViewer.exe"
    ];

    public string? GetCmRcViewerPath()
    {
        // 0. Caminho customizado configurado no .env
        string custom = ConfigService.CmRcViewerCustomPath;
        if (!string.IsNullOrWhiteSpace(custom) && File.Exists(custom))
            return custom;

        // 1. Procura pela variável de ambiente do SCCM Console (SMS_ADMIN_UI_PATH)
        string? smsUi = Environment.GetEnvironmentVariable("SMS_ADMIN_UI_PATH");
        if (!string.IsNullOrEmpty(smsUi))
        {
            string dir = Path.GetDirectoryName(smsUi) ?? smsUi;
            string test1 = Path.Combine(dir, "i386", "CmRcViewer.exe");
            if (File.Exists(test1)) return test1;
            string test2 = Path.Combine(dir, "CmRcViewer.exe");
            if (File.Exists(test2)) return test2;
        }

        // 2. Procura nos caminhos padrão do console SCCM oficial instalado no Windows
        foreach (var path in SccmStandardPaths)
        {
            if (File.Exists(path))
                return path;
        }

        // 3. Procura na pasta local da aplicação
        string localAppPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CmRcViewer.exe");
        if (File.Exists(localAppPath)) return localAppPath;

        // 4. Procura no PATH do sistema
        string? pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrEmpty(pathEnv))
        {
            foreach (var p in pathEnv.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                try
                {
                    string candidate = Path.Combine(p, "CmRcViewer.exe");
                    if (File.Exists(candidate)) return candidate;
                }
                catch { }
            }
        }

        // 5. Fallback opcional no diretório D:\ArrumarPC caso exista
        string legacyPath = @"D:\ArrumarPC\CmRcViewer.exe";
        if (File.Exists(legacyPath)) return legacyPath;

        return null;
    }

    public void OpenRemoteExplorer(string computerName)
    {
        if (string.IsNullOrWhiteSpace(computerName))
            throw new ArgumentException("Nome do computador não informado.");

        string host = computerName.Trim();
        bool isLocal = RemoteManagementHelper.IsLocalMachine(host);
        string uncPath = isLocal ? @"C:\" : $@"\\{host}\C$";

        if (!isLocal)
        {
            // Autentica previamente no compartilhamento administrativo com a credencial segura DPAPI
            RemoteManagementHelper.EnsureAdminShareConnected(host);
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{uncPath}\"",
            UseShellExecute = true
        });
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
            // Se o executável do CmRcViewer não foi encontrado em nenhum diretório padrão, lança erro para avisar o usuário
            throw new FileNotFoundException(
                "O executável 'CmRcViewer.exe' não foi encontrado. Certifique-se de que o Console do Configuration Manager (SCCM) está instalado nesta máquina.");
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
