using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class EnvironmentLauncherService
{
    private readonly BrowserLauncherService _browserService = new();
    private readonly CleanupService _cleanupService = new();
    private readonly SajmpLauncherService _sajmpService = new();

    public void LaunchEnvironment(string modeName, bool cleanFirst = true)
    {
        var sw = Stopwatch.StartNew();
        AppLoggerService.Info($"Acionamento do ambiente: '{modeName}' (Limpeza prévia: {cleanFirst}).", nameof(EnvironmentLauncherService));

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

            case "SAJMP":
                _sajmpService.LaunchAndLogin();
                break;

            default:
                AppLoggerService.Warning($"Modo de ambiente desconhecido: '{modeName}'.", nameof(EnvironmentLauncherService));
                return;
        }

        AppLoggerService.Success($"Ambiente '{modeName}' finalizado em {AppLoggerService.FormatElapsed(sw)}.", nameof(EnvironmentLauncherService));
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

    public static List<Models.EnvironmentModeItem> CreateDefaultModes()
    {
        static List<Models.EnvironmentStepItem> CleanupSteps() =>
        [
            new() { Id = "clean_office", Category = "Limpeza Prévia", Label = "Salvar e fechar Office (Word / Excel via COM)" },
            new() { Id = "clean_browsers", Category = "Limpeza Prévia", Label = "Fechar Chrome & Edge" },
            new() { Id = "clean_teams", Category = "Limpeza Prévia", Label = "Fechar Microsoft Teams" },
            new() { Id = "clean_admin_tools", Category = "Limpeza Prévia", Label = "Fechar Consoles AD, SCCM e SAJMP" },
            new() { Id = "clean_utilities", Category = "Limpeza Prévia", Label = "Fechar utilitários (Foxit, Bloco de Notas, Relógio)" },
            new() { Id = "clean_telemetry", Category = "Limpeza Prévia", Label = "Parar Telemetria (CompatTelRunner)" },
            new() { Id = "clear_firefox_session", Category = "Limpeza Prévia", Label = "Limpar sessão anterior do Firefox" }
        ];

        var chamados = new Models.EnvironmentModeItem
        {
            ModeKey = "Chamados",
            Icon = "🎫",
            Title = "Modo Chamados",
            BadgeText = "SUPORTE TI",
            Subtitle = "Encerra distrações, abre WhatsApp PWA e carrega OTRS, CitSmart e Bancada no Firefox.",
            Steps = [.. CleanupSteps(),
                new() { Id = "open_whatsapp", Category = "Abertura", Label = "Abrir WhatsApp Web PWA (Firefox Container)" },
                new() { Id = "open_otrs", Category = "Abertura", Label = "Abrir OTRS (Central)" },
                new() { Id = "open_citsmart", Category = "Abertura", Label = "Abrir CitSmart (Suporte)" },
                new() { Id = "open_bancada", Category = "Abertura", Label = "Abrir Bancada Local (Streamlit 8501)" }
            ]
        };

        var contratos = new Models.EnvironmentModeItem
        {
            ModeKey = "Contratos",
            Icon = "📑",
            Title = "Modo Contratos",
            BadgeText = "GESTÃO",
            Subtitle = "Prepara ambiente limpo para fiscalização de contratos, termos e CitSmart.",
            Steps = [.. CleanupSteps(),
                new() { Id = "open_citsmart", Category = "Abertura", Label = "Abrir CitSmart (Suporte)" }
            ]
        };

        var impressoras = new Models.EnvironmentModeItem
        {
            ModeKey = "Impressoras",
            Icon = "🖨️",
            Title = "Modo Impressoras",
            BadgeText = "PAPERCUT",
            Subtitle = "Ambiente focado em servidores de impressão, filas e suporte.",
            Steps = [.. CleanupSteps(),
                new() { Id = "open_citsmart", Category = "Abertura", Label = "Abrir CitSmart (Suporte)" },
                new() { Id = "open_papercut", Category = "Abertura", Label = "Abrir Admin PaperCut (se configurado)" }
            ]
        };

        var pxe = new Models.EnvironmentModeItem
        {
            ModeKey = "PXE",
            Icon = "💿",
            Title = "Modo PXE",
            BadgeText = "BANCADA PXE",
            Subtitle = "Prepara bancada de formatação, deploys de imagem e chamados CitSmart.",
            Steps = [.. CleanupSteps(),
                new() { Id = "open_citsmart", Category = "Abertura", Label = "Abrir CitSmart (Suporte)" }
            ]
        };

        var estudo = new Models.EnvironmentModeItem
        {
            ModeKey = "Estudo",
            Icon = "📚",
            Title = "Modo Estudo",
            BadgeText = "FOCO",
            Subtitle = "Ambiente blindado livre de distrações para leitura, estudos e pesquisa.",
            Steps = [.. CleanupSteps(),
                new() { Id = "open_google", Category = "Abertura", Label = "Abrir navegador limpo (Google)" }
            ]
        };

        var sajmp = new Models.EnvironmentModeItem
        {
            ModeKey = "SAJMP",
            Icon = "⚖️",
            Title = "Abrir SAJMP",
            BadgeText = "SAJMP",
            Subtitle = "Inicia o SAJMP, lê a senha do cofre DPAPI e efetua login automático sem SendKeys.",
            Steps =
            [
                new() { Id = "open_sajmp", Category = "Abertura", Label = "Abrir SAJMP e efetuar login automático" }
            ]
        };

        return [chamados, contratos, impressoras, pxe, estudo, sajmp];
    }

    public void ExecuteCustomEnvironment(Models.EnvironmentModeItem mode, Action<string> logCallback)
    {
        var sw = Stopwatch.StartNew();
        void Log(string msg)
        {
            logCallback?.Invoke($"[{DateTime.Now:HH:mm:ss}] {msg}");
        }

        var enabledIds = mode.Steps
            .Where(s => s.IsEnabled)
            .Select(s => s.Id)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        int skippedCount = mode.Steps.Count - enabledIds.Count;
        Log($"=== Iniciando {mode.Title} ===");
        AppLoggerService.Info($"Acionamento personalizado do ambiente '{mode.ModeKey}': {enabledIds.Count} etapa(s) ativa(s), {skippedCount} pulada(s).", nameof(EnvironmentLauncherService));

        foreach (var step in mode.Steps.Where(s => !s.IsEnabled))
            Log($"[PULADO] {step.Label}");

        bool anyCleanup = enabledIds.Any(id => id.StartsWith("clean_", StringComparison.OrdinalIgnoreCase));
        if (anyCleanup)
        {
            Log("[▶] Executando limpeza de processos selecionados...");
            _cleanupService.CleanEnvironment(new Models.CleanupOptions
            {
                CloseOffice = enabledIds.Contains("clean_office"),
                CloseBrowsers = enabledIds.Contains("clean_browsers"),
                CloseFirefox = false,
                CloseTeams = enabledIds.Contains("clean_teams"),
                CloseAdministrativeTools = enabledIds.Contains("clean_admin_tools"),
                CloseUtilities = enabledIds.Contains("clean_utilities"),
                StopTelemetry = enabledIds.Contains("clean_telemetry")
            });
            Log("[OK] Limpeza de processos concluída.");
        }

        if (enabledIds.Contains("clear_firefox_session"))
        {
            Log("[▶] Limpando sessão anterior do Firefox...");
            _browserService.ClearFirefoxSession();
            Log("[OK] Sessão do Firefox limpa.");
        }

        if (enabledIds.Contains("open_sajmp"))
        {
            var sajResult = _sajmpService.LaunchAndLogin(Log);
            if (!sajResult.Success)
                return;
        }

        if (enabledIds.Contains("open_whatsapp"))
        {
            Log("[▶] Abrindo WhatsApp Web PWA...");
            _browserService.OpenWhatsappPwa();
            Log("[OK] WhatsApp Web PWA acionado.");
        }

        var urlsToOpen = new List<string>();
        if (enabledIds.Contains("open_otrs") && InputValidator.IsSafeHttpUrl(ConfigService.UrlOtrs))
            urlsToOpen.Add(ConfigService.UrlOtrs);
        if (enabledIds.Contains("open_citsmart") && InputValidator.IsSafeHttpUrl(ConfigService.UrlCitsmart))
            urlsToOpen.Add(ConfigService.UrlCitsmart);
        if (enabledIds.Contains("open_bancada") && InputValidator.IsSafeHttpUrl(ConfigService.UrlBancadaLocal))
            urlsToOpen.Add(ConfigService.UrlBancadaLocal);
        if (enabledIds.Contains("open_papercut") && InputValidator.IsSafeHttpUrl(ConfigService.UrlPapercut))
            urlsToOpen.Add(ConfigService.UrlPapercut);
        if (enabledIds.Contains("open_google"))
            urlsToOpen.Add("https://google.com");

        if (urlsToOpen.Count > 0)
        {
            Log($"[▶] Abrindo {urlsToOpen.Count} aba(s) no Firefox...");
            _browserService.OpenUrlsInFirefox(urlsToOpen.ToArray());
            Log("[OK] Abas abertas no Firefox.");
        }

        Log($"[SUCESSO] {mode.Title} carregado com sucesso!");
        AppLoggerService.Success($"Ambiente personalizado '{mode.Title}' carregado ({urlsToOpen.Count} URL(s)) em {AppLoggerService.FormatElapsed(sw)}.", nameof(EnvironmentLauncherService));
    }
}
