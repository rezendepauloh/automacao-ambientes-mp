using System.IO;

namespace AmbientesHub.Services;

public class EnvironmentLauncherService
{
    private readonly BrowserLauncherService _browserService = new();
    private readonly CleanupService _cleanupService = new();

    public void LaunchEnvironment(string modeName, bool cleanFirst = true)
    {
        AppLoggerService.Info($"Acionamento do ambiente: {modeName} (Limpeza prévia: {cleanFirst})", nameof(EnvironmentLauncherService));

        if (cleanFirst)
        {
            _cleanupService.CleanEnvironment(new Models.CleanupOptions
            {
                CloseOffice = true,
                CloseBrowsers = true,
                CloseFirefox = false,
                CloseTeams = true,
                CloseAdministrativeTools = true,
                CloseUtilities = true,
                StopTelemetry = true
            });

            _browserService.ClearFirefoxSession();
        }

        switch (modeName)
        {
            case "Chamados":
                LaunchModoChamados();
                break;

            case "Contratos":
                LaunchModoContratos();
                break;

            case "Impressoras":
                LaunchModoImpressoras();
                break;

            case "PXE":
                LaunchModoPxe();
                break;

            case "Estudo":
                LaunchModoEstudo();
                break;
        }
    }

    private void LaunchModoChamados()
    {
        AppLoggerService.Info("Carregando Modo Chamados...", nameof(EnvironmentLauncherService));
        _browserService.OpenWhatsappPwa();
        _browserService.OpenUrlsInFirefox([
            ConfigService.UrlOtrs,
            ConfigService.UrlCitsmart,
            ConfigService.UrlBancadaLocal
        ]);
        AppLoggerService.Success("Modo Chamados carregado.", nameof(EnvironmentLauncherService));
    }

    private void LaunchModoContratos()
    {
        AppLoggerService.Info("Carregando Modo Contratos...", nameof(EnvironmentLauncherService));
        _browserService.OpenUrlsInFirefox([ConfigService.UrlCitsmart]);
        AppLoggerService.Success("Modo Contratos carregado.", nameof(EnvironmentLauncherService));
    }

    private void LaunchModoImpressoras()
    {
        AppLoggerService.Info("Carregando Modo Impressoras...", nameof(EnvironmentLauncherService));
        _browserService.OpenUrlsInFirefox([ConfigService.UrlCitsmart]);
        AppLoggerService.Success("Modo Impressoras carregado.", nameof(EnvironmentLauncherService));
    }

    private void LaunchModoPxe()
    {
        AppLoggerService.Info("Carregando Modo PXE...", nameof(EnvironmentLauncherService));
        _browserService.OpenUrlsInFirefox([ConfigService.UrlCitsmart]);
        AppLoggerService.Success("Modo PXE carregado.", nameof(EnvironmentLauncherService));
    }

    private void LaunchModoEstudo()
    {
        AppLoggerService.Info("Carregando Modo Estudo...", nameof(EnvironmentLauncherService));
        _browserService.OpenUrlsInFirefox(["https://google.com"]);
        AppLoggerService.Success("Modo Estudo carregado.", nameof(EnvironmentLauncherService));
    }
}
