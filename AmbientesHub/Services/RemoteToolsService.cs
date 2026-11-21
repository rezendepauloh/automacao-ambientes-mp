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
        {
            AppLoggerService.Info($"CmRcViewer localizado via CMRCVIEWER_PATH (.env): '{custom}'.", nameof(RemoteToolsService));
            return custom;
        }

        // 1. Procura pela variável de ambiente do SCCM Console (SMS_ADMIN_UI_PATH)
        string? smsUi = Environment.GetEnvironmentVariable("SMS_ADMIN_UI_PATH");
        if (!string.IsNullOrEmpty(smsUi))
        {
            string dir = Path.GetDirectoryName(smsUi) ?? smsUi;
            string test1 = Path.Combine(dir, "i386", "CmRcViewer.exe");
            if (File.Exists(test1))
            {
                AppLoggerService.Info($"CmRcViewer localizado via SMS_ADMIN_UI_PATH: '{test1}'.", nameof(RemoteToolsService));
                return test1;
            }
            string test2 = Path.Combine(dir, "CmRcViewer.exe");
            if (File.Exists(test2))
            {
                AppLoggerService.Info($"CmRcViewer localizado via SMS_ADMIN_UI_PATH: '{test2}'.", nameof(RemoteToolsService));
                return test2;
            }
        }

        // 2. Procura nos caminhos padrão do console SCCM oficial instalado no Windows
        foreach (var path in SccmStandardPaths)
        {
            if (File.Exists(path))
            {
                AppLoggerService.Info($"CmRcViewer localizado em caminho padrão: '{path}'.", nameof(RemoteToolsService));
                return path;
            }
        }

        // 3. Procura na pasta local da aplicação
        string localAppPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CmRcViewer.exe");
        if (File.Exists(localAppPath))
        {
            AppLoggerService.Info($"CmRcViewer localizado na pasta local do Hub: '{localAppPath}'.", nameof(RemoteToolsService));
            return localAppPath;
        }

        // 4. Procura no PATH do sistema
        string? pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrEmpty(pathEnv))
        {
            foreach (var p in pathEnv.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                try
                {
                    string candidate = Path.Combine(p, "CmRcViewer.exe");
                    if (File.Exists(candidate))
                    {
                        AppLoggerService.Info($"CmRcViewer localizado no PATH: '{candidate}'.", nameof(RemoteToolsService));
                        return candidate;
                    }
                }
                catch { }
            }
        }

        // 5. Fallback opcional no diretório D:\ArrumarPC caso exista
        string legacyPath = @"D:\ArrumarPC\CmRcViewer.exe";
        if (File.Exists(legacyPath))
        {
            AppLoggerService.Info($"CmRcViewer localizado via fallback legado: '{legacyPath}'.", nameof(RemoteToolsService));
            return legacyPath;
        }

        AppLoggerService.Warning("CmRcViewer.exe não foi localizado em nenhum caminho conhecido.", nameof(RemoteToolsService));
        return null;
    }

    public void OpenRemoteExplorer(string computerName)
    {
        var sw = Stopwatch.StartNew();
        string host = RequireValidHost(computerName);
        bool isLocal = RemoteManagementHelper.IsLocalMachine(host);
        string uncPath = isLocal ? @"C:\" : $@"\\{host}\C$";
        AppLoggerService.Info($"Abrindo Explorador remoto em '{uncPath}' (local={isLocal}).", nameof(RemoteToolsService));

        try
        {
            if (!isLocal)
            {
                // Autentica previamente no compartilhamento administrativo com a credencial segura DPAPI
                bool connected = RemoteManagementHelper.EnsureAdminShareConnected(host);
                if (!connected)
                    AppLoggerService.Warning($"Não foi possível pré-autenticar em '{uncPath}'; tentando abrir Explorer mesmo assim.", nameof(RemoteToolsService));
            }

            var explorerInfo = new ProcessStartInfo { FileName = "explorer.exe", UseShellExecute = true };
            explorerInfo.ArgumentList.Add(uncPath);
            Process.Start(explorerInfo);
            AppLoggerService.Success($"Explorer iniciado para '{uncPath}' em {AppLoggerService.FormatElapsed(sw)}.", nameof(RemoteToolsService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(RemoteToolsService), $"Abrir Explorer em '{uncPath}'", ex, sw.Elapsed);
            throw;
        }
    }

    /// <summary>Fail-closed: aceita somente hostname/FQDN whitelisted; caso contrário lança ArgumentException.</summary>
    internal static string RequireValidHost(string? computerName)
    {
        if (string.IsNullOrWhiteSpace(computerName))
        {
            AppLoggerService.Warning("Operação remota rejeitada: nome do computador não informado.", nameof(RemoteToolsService));
            throw new ArgumentException("Nome do computador não informado.");
        }

        string host = computerName.Trim();
        if (!InputValidator.IsValidHostName(host))
        {
            AppLoggerService.Warning($"Operação remota rejeitada pela whitelist (host inválido: '{host}').", nameof(RemoteToolsService));
            throw new ArgumentException("Nome do computador inválido. Use apenas letras, números, hífen e ponto.");
        }

        return host;
    }

    public void OpenSccmRemoteControl(string computerName, string? siteServer = null)
    {
        var sw = Stopwatch.StartNew();
        string host = RequireValidHost(computerName);

        siteServer ??= ConfigService.SccmSiteServer;
        if (!InputValidator.IsValidHostName(siteServer?.Trim()))
        {
            AppLoggerService.Warning($"SCCM Remote abortado para '{host}': SCCM_SITE_SERVER ausente ou inválido.", nameof(RemoteToolsService));
            throw new ArgumentException("SCCM_SITE_SERVER ausente ou inválido. Configure-o na aba Configurações.");
        }
        siteServer = siteServer!.Trim();

        var cmPath = GetCmRcViewerPath();

        if (cmPath != null)
        {
            try
            {
                var info = new ProcessStartInfo { FileName = cmPath, UseShellExecute = true };
                info.ArgumentList.Add(host);
                info.ArgumentList.Add($@"\\{siteServer}");
                var proc = Process.Start(info);
                AppLoggerService.Success(
                    $"CmRcViewer iniciado para '{host}' via '{siteServer}' (PID={proc?.Id}, bin='{cmPath}') em {AppLoggerService.FormatElapsed(sw)}.",
                    nameof(RemoteToolsService));
            }
            catch (Exception ex)
            {
                AppLoggerService.LogException(nameof(RemoteToolsService), $"Iniciar CmRcViewer para '{host}'", ex, sw.Elapsed);
                throw;
            }
        }
        else
        {
            AppLoggerService.Danger($"CmRcViewer.exe não localizado ao tentar conectar em '{host}'.", nameof(RemoteToolsService));
            // Se o executável do CmRcViewer não foi encontrado em nenhum diretório padrão, lança erro para avisar o usuário
            throw new FileNotFoundException(
                "O executável 'CmRcViewer.exe' não foi encontrado. Certifique-se de que o Console do Configuration Manager (SCCM) está instalado nesta máquina.");
        }
    }

    public void OpenRdp(string computerName)
    {
        var sw = Stopwatch.StartNew();
        string host = RequireValidHost(computerName);

        try
        {
            var info = new ProcessStartInfo { FileName = "mstsc.exe", UseShellExecute = true };
            info.ArgumentList.Add($"/v:{host}");
            var proc = Process.Start(info);
            AppLoggerService.Success(
                $"mstsc.exe (RDP) iniciado para '{host}' (PID={proc?.Id}) em {AppLoggerService.FormatElapsed(sw)}.",
                nameof(RemoteToolsService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(RemoteToolsService), $"Iniciar RDP para '{host}'", ex, sw.Elapsed);
            throw;
        }
    }
}
