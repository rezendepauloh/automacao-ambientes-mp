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
        if (InputValidator.IsSafeLocalExecutablePath(custom) && File.Exists(custom))
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
        string host = RequireValidHost(computerName);
        bool isLocal = RemoteManagementHelper.IsLocalMachine(host);
        string uncPath = isLocal ? @"C:\" : $@"\\{host}\C$";

        if (!isLocal)
        {
            // Autentica previamente no compartilhamento administrativo com a credencial segura DPAPI
            RemoteManagementHelper.EnsureAdminShareConnected(host);
        }

        var explorerInfo = new ProcessStartInfo { FileName = "explorer.exe", UseShellExecute = true };
        explorerInfo.ArgumentList.Add(uncPath);
        Process.Start(explorerInfo);
    }

    /// <summary>Fail-closed: aceita somente hostname/FQDN whitelisted; caso contrário lança ArgumentException.</summary>
    internal static string RequireValidHost(string? computerName)
    {
        if (string.IsNullOrWhiteSpace(computerName))
            throw new ArgumentException("Nome do computador não informado.");

        string host = computerName.Trim();
        if (!InputValidator.IsValidHostName(host))
            throw new ArgumentException("Nome do computador inválido. Use apenas letras, números, hífen e ponto.");

        return host;
    }

    public void OpenSccmRemoteControl(string computerName, string? siteServer = null)
    {
        string host = RequireValidHost(computerName);

        siteServer ??= ConfigService.SccmSiteServer;
        if (!InputValidator.IsValidHostName(siteServer?.Trim()))
            throw new ArgumentException("SCCM_SITE_SERVER ausente ou inválido. Configure-o na aba Configurações.");
        siteServer = siteServer!.Trim();

        var cmPath = GetCmRcViewerPath();

        if (cmPath != null)
        {
            var info = new ProcessStartInfo { FileName = cmPath, UseShellExecute = true };
            info.ArgumentList.Add(host);
            info.ArgumentList.Add($@"\\{siteServer}");
            Process.Start(info);
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
        string host = RequireValidHost(computerName);

        var info = new ProcessStartInfo { FileName = "mstsc.exe", UseShellExecute = true };
        info.ArgumentList.Add($"/v:{host}");
        Process.Start(info);
    }
}
