using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Diagnostics;
using System.IO;
using AmbientesHub.Models;
using AmbientesHub.Services;
using AmbientesHub.Views;
using Wpf.Ui.Controls;

namespace AmbientesHub;

public partial class MainWindow : FluentWindow
{
    private readonly RemoteToolsService _remoteToolsService = new();
    private readonly CleanupService _cleanupService = new();
    private readonly EnvironmentLauncherService _envService = new();
    private readonly AdSccmService _adSccmService = new();
    private readonly SystemPerformanceService _perfService = new();
    private readonly SoftwareProvisioningService _softwareService = new();
    private readonly DeviceAnalyzerService _deviceAnalyzerService = new();
    private readonly RemoteMaintenanceService _remoteMaintenanceService = new();
    private readonly RemoteUserProfileService _remoteUserProfileService = new();
    private readonly ObservableCollection<SoftwarePackageItem> _softwareCatalog = new(SoftwareProvisioningService.CreateDefaultCatalog());
    private readonly ObservableCollection<RemoteUserProfileItem> _remoteProfilesList = [];
    private readonly StartupService _startupService = new();
    private readonly BancadaDockerService _bancadaService = new();
    private readonly BrowserLauncherService _browserLauncherService = new();

    private readonly List<AdUserDetails> _allAdUsersFound = [];
    private readonly List<AdUserDetails> _filteredAdUsers = [];
    private int _adUserCurrentPage = 1;
    private const int AdUsersPageSize = 10;
    private bool _isPopulatingDeptFilter = false;

    private TrayIconService? _trayIconService;
    private bool _isRealExit = false;

    private DeviceAnalysisResult? _lastAnalyzerResult;
    private bool _isNavExpanded = true;
    private bool _appsCatalogLoadedOnce;
    private readonly List<string> _allAppLogsInMemory = [];
    private readonly RegistryOptimizerService _registryOptimizerService = new();
    private readonly ObservableCollection<RegistryTweakItem> _allRegistryTweaks = [];
    private readonly ObservableCollection<RegistryTweakItem> _filteredRegistryTweaks = [];
    private bool _registryTweaksScannedOnce = false;

    public MainWindow()
    {
        InitializeComponent();
        AppLoggerService.OnLogLineWritten += OnNewLogLineReceived;
        AppLoggerService.Info("Ambientes Hub iniciado com sucesso.", nameof(MainWindow));
        LstSoftwareCatalog.ItemsSource = _softwareCatalog;
        ItemsRemoteProfiles.ItemsSource = _remoteProfilesList;
        LstRegistryTweaks.ItemsSource = _filteredRegistryTweaks;
        InitializeRegistryTweaks();
        PopulateFontList();
        LoadSettingsIntoUi();
        InitializeSystemTray();
        RefreshPerformanceAdminStatus();
        RefreshAppsAdminStatus();
        if (CmbAdUserStatusFilter != null && CmbAdUserStatusFilter.Items.Count > 0)
        {
            CmbAdUserStatusFilter.SelectedIndex = 0;
        }
    }

    #region Navegação Lateral Vertical & Hambúrguer
    private void BtnToggleNav_Click(object sender, RoutedEventArgs e)
    {
        _isNavExpanded = !_isNavExpanded;
        if (_isNavExpanded)
        {
            ColNavPane.Width = new GridLength(240);
            TxtNavTitle.Visibility = Visibility.Visible;
            NavLbl0.Visibility = Visibility.Visible;
            NavLbl1.Visibility = Visibility.Visible;
            NavLbl2.Visibility = Visibility.Visible;
            NavLbl3.Visibility = Visibility.Visible;
            NavLbl4.Visibility = Visibility.Visible;
            NavLbl5.Visibility = Visibility.Visible;
            NavLbl6.Visibility = Visibility.Visible;
            NavLbl7.Visibility = Visibility.Visible;
            NavLbl8.Visibility = Visibility.Visible;
            NavLbl9.Visibility = Visibility.Visible;
            NavLbl10.Visibility = Visibility.Visible;
            NavLbl11.Visibility = Visibility.Visible;
            NavLbl12.Visibility = Visibility.Visible;
            NavLbl13.Visibility = Visibility.Visible;
        }
        else
        {
            ColNavPane.Width = new GridLength(60);
            TxtNavTitle.Visibility = Visibility.Collapsed;
            NavLbl0.Visibility = Visibility.Collapsed;
            NavLbl1.Visibility = Visibility.Collapsed;
            NavLbl2.Visibility = Visibility.Collapsed;
            NavLbl3.Visibility = Visibility.Collapsed;
            NavLbl4.Visibility = Visibility.Collapsed;
            NavLbl5.Visibility = Visibility.Collapsed;
            NavLbl6.Visibility = Visibility.Collapsed;
            NavLbl7.Visibility = Visibility.Collapsed;
            NavLbl8.Visibility = Visibility.Collapsed;
            NavLbl9.Visibility = Visibility.Collapsed;
            NavLbl10.Visibility = Visibility.Collapsed;
            NavLbl11.Visibility = Visibility.Collapsed;
            NavLbl12.Visibility = Visibility.Collapsed;
            NavLbl13.Visibility = Visibility.Collapsed;
        }
    }

    private void TxtGlobalSearch_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            ExecutarPesquisaGlobal();
        }
        else if (e.Key == Key.Escape)
        {
            TxtGlobalSearch.Clear();
            TxtFooterStatus.Text = "Ambientes Hub ativo e pronto.";
        }
    }

    private void ExecutarPesquisaGlobal()
    {
        string query = TxtGlobalSearch.Text.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            TxtFooterStatus.Text = "Digite um termo para pesquisar (máquina, usuário ou módulo).";
            return;
        }

        string lower = query.ToLowerInvariant();

        // 1. Navegação rápida para abas/módulos
        if (lower is "ambientes" or "ambiente" or "virtual" or "virtuais" or "chamados" or "contratos" or "pastas")
        {
            NavListBox.SelectedIndex = 0;
            TxtFooterStatus.Text = "Navegado para: Ambientes Virtuais";
            return;
        }

        if (lower is "usuarios ad" or "usuários ad" or "consulta ad" or "usuario ad" or "usuário ad" or "grupos ad")
        {
            NavListBox.SelectedIndex = 1;
            TxtFooterStatus.Text = "Navegado para: Consulta de Usuários AD";
            return;
        }

        if (lower.StartsWith("usuario ") || lower.StartsWith("usuário "))
        {
            string targetUser = query[(query.IndexOf(' ') + 1)..].Trim();
            NavListBox.SelectedIndex = 1;
            if (!string.IsNullOrEmpty(targetUser))
            {
                TxtAdUserQuery.Text = targetUser;
                BtnConsultarUsuarioAd_Click(this, new RoutedEventArgs());
            }
            TxtFooterStatus.Text = $"Navegado para Consulta de Usuários AD: {targetUser}";
            return;
        }

        if (lower is "acesso remoto" or "remoto" or "rdp" or "cmrcviewer" or "vnc" or "mstsc")
        {
            NavListBox.SelectedIndex = 2;
            TxtFooterStatus.Text = "Navegado para: Acesso Remoto";
            return;
        }

        if (lower is "saneamento" or "limpar cache" or "teams" or "recursos" or "cache" or "explorer")
        {
            NavListBox.SelectedIndex = 4;
            TxtFooterStatus.Text = "Navegado para: Saneamento de Recursos";
            return;
        }

        if (lower is "performance" or "auditoria" or "debloat" or "otimizar" or "otimização" or "otimizacao" or "lentidão" or "lentidao" or "lento" or "wsl" or "trim" or "hibernação" or "hibernacao" or "bloat" or "checkup")
        {
            NavListBox.SelectedIndex = 5;
            TxtFooterStatus.Text = "Navegado para: Performance & Auditoria";
            if (lower is "auditoria" or "checkup")
            {
                BtnExecutarAuditoria_Click(this, new RoutedEventArgs());
            }
            return;
        }

        if (lower is "apps" or "programas" or "instalar" or "instalador" or "rsat" or "dsa" or "console" or "winget" or "provisionamento")
        {
            NavListBox.SelectedIndex = 6;
            TxtFooterStatus.Text = "Navegado para: Apps & Provisionamento";
            return;
        }

        if (lower is "analisador" or "analisar" or "relatorio" or "relatório" or "inventario" or "inventário" or "html" or "pdf" or "excel")
        {
            NavListBox.SelectedIndex = 7;
            TxtFooterStatus.Text = "Navegado para: Analisador de Máquina";
            return;
        }

        if (lower is "manutencao" or "manutenção" or "limpeza remota" or "manutencao remota")
        {
            NavListBox.SelectedIndex = 8;
            TxtFooterStatus.Text = "Navegado para: Manutenção & Limpeza Remota";
            return;
        }

        if (lower is "perfis" or "perfil" or "usuarios remotos" or "usuários remotos" or "remover usuario" or "remover usuário")
        {
            NavListBox.SelectedIndex = 9;
            TxtFooterStatus.Text = "Navegado para: Limpeza de Perfis de Usuários";
            return;
        }

        if (lower is "bancada" or "docker" or "wsl" or "streamlit" or "8501" or "chamados bancada")
        {
            NavListBox.SelectedIndex = 10;
            TxtFooterStatus.Text = "Navegado para: Bancada Streamlit (Docker WSL)";
            return;
        }

        if (lower is "regedit" or "registro" or "tweak" or "tweaks" or "otimizar registro" or "otimização registro" or "rollback" or "ajustes registro")
        {
            NavListBox.SelectedIndex = 11;
            TxtFooterStatus.Text = "Navegado para: Otimizações de Registro (Tweaks & Rollback)";
            return;
        }

        if (lower is "logs" or "log" or "eventos" or "historico" or "histórico" or "console" or "app.log")
        {
            NavListBox.SelectedIndex = 12;
            TxtFooterStatus.Text = "Navegado para: Logs da Aplicação";
            return;
        }

        if (lower is "config" or "configuração" or "configurações" or "cofre" or "credenciais" or "ajustes" or "fonte")
        {
            NavListBox.SelectedIndex = 13;
            TxtFooterStatus.Text = "Navegado para: Configurações & Credenciais";
            return;
        }

        if (lower.StartsWith("limpar ") || lower.StartsWith("desprovisionar ") || lower.StartsWith("ad "))
        {
            string target = query.Substring(query.IndexOf(' ') + 1).Trim();
            NavListBox.SelectedIndex = 3;
            if (!string.IsNullOrEmpty(target))
            {
                TxtAdTargetComputer.Text = target;
                BtnAdQuery_Click(this, new RoutedEventArgs());
            }
            TxtFooterStatus.Text = $"Navegado para Limpeza AD/SCCM: {target}";
            return;
        }

        if (lower is "limpeza" or "desprovisionar" or "ad" or "sccm")
        {
            NavListBox.SelectedIndex = 3;
            TxtFooterStatus.Text = "Navegado para: Limpeza AD & SCCM";
            return;
        }

        // 2. Consulta Global no Acesso Remoto & Diagnóstico (Máquinas e Usuários)
        NavListBox.SelectedIndex = 2;
        TxtComputerName.Text = query;
        TxtFooterStatus.Text = $"Pesquisa global acionada para: '{query}'...";
        BtnConsultarRemoto_Click(this, new RoutedEventArgs());
    }

    private void TxtGlobalSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        string query = TxtGlobalSearch.Text.Trim();
        if (string.IsNullOrEmpty(query)) return;

        // Feedback visual na barra de status indicando o filtro da pesquisa
        TxtFooterStatus.Text = $"Filtro ativo: '{query}' (Pressione Enter para executar)";
    }

    private void NavListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (PanelAmbientes == null) return;

        PanelAmbientes.Visibility = Visibility.Collapsed;
        PanelRemoto.Visibility = Visibility.Collapsed;
        PanelAdSccm.Visibility = Visibility.Collapsed;
        PanelSaneamento.Visibility = Visibility.Collapsed;
        PanelPerformance.Visibility = Visibility.Collapsed;
        PanelApps.Visibility = Visibility.Collapsed;
        PanelAnalisador.Visibility = Visibility.Collapsed;
        PanelLimpezaRemota.Visibility = Visibility.Collapsed;
        PanelPerfisRemotos.Visibility = Visibility.Collapsed;
        PanelUsuariosAd.Visibility = Visibility.Collapsed;
        PanelBancada.Visibility = Visibility.Collapsed;
        PanelRegistryTweaks.Visibility = Visibility.Collapsed;
        PanelAppLogs.Visibility = Visibility.Collapsed;
        PanelConfig.Visibility = Visibility.Collapsed;

        switch (NavListBox.SelectedIndex)
        {
            case 0:
                PanelAmbientes.Visibility = Visibility.Visible;
                break;
            case 1:
                PanelUsuariosAd.Visibility = Visibility.Visible;
                break;
            case 2:
                PanelRemoto.Visibility = Visibility.Visible;
                break;
            case 3:
                PanelAdSccm.Visibility = Visibility.Visible;
                break;
            case 4:
                PanelSaneamento.Visibility = Visibility.Visible;
                break;
            case 5:
                PanelPerformance.Visibility = Visibility.Visible;
                RefreshPerformanceAdminStatus();
                break;
            case 6:
                PanelApps.Visibility = Visibility.Visible;
                RefreshAppsAdminStatus();
                if (!_appsCatalogLoadedOnce)
                {
                    _appsCatalogLoadedOnce = true;
                    _ = RefreshAppsCatalogAsync();
                }
                break;
            case 7:
                PanelAnalisador.Visibility = Visibility.Visible;
                break;
            case 8:
                PanelLimpezaRemota.Visibility = Visibility.Visible;
                break;
            case 9:
                PanelPerfisRemotos.Visibility = Visibility.Visible;
                break;
            case 10:
                PanelBancada.Visibility = Visibility.Visible;
                _ = CheckBancadaStatusAsync();
                break;
            case 11:
                PanelRegistryTweaks.Visibility = Visibility.Visible;
                RefreshRegistryTweaksStatus();
                break;
            case 12:
                PanelAppLogs.Visibility = Visibility.Visible;
                LoadAppLogsToUi();
                break;
            case 13:
                PanelConfig.Visibility = Visibility.Visible;
                break;
        }
    }
    #endregion

    #region Configurações UI & Seletor de Fontes
    private void PopulateFontList()
    {
        try
        {
            var allFonts = Fonts.SystemFontFamilies
                .Select(f => f.Source)
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            CmbFonts.ItemsSource = allFonts;

            string savedFont = ConfigService.Get("APP_TERMINAL_FONT", "");

            string selected = "";
            if (!string.IsNullOrEmpty(savedFont) && allFonts.Contains(savedFont))
            {
                selected = savedFont;
            }
            else
            {
                selected = allFonts.FirstOrDefault(f => f.StartsWith("Caskaydia", StringComparison.OrdinalIgnoreCase))
                        ?? allFonts.FirstOrDefault(f => f.Contains("Nerd Font", StringComparison.OrdinalIgnoreCase))
                        ?? allFonts.FirstOrDefault(f => f.Contains("Cascadia Code", StringComparison.OrdinalIgnoreCase))
                        ?? allFonts.FirstOrDefault(f => f.Contains("Cascadia Mono", StringComparison.OrdinalIgnoreCase))
                        ?? allFonts.FirstOrDefault(f => f.Equals("Consolas", StringComparison.OrdinalIgnoreCase))
                        ?? "Consolas";
            }

            CmbFonts.SelectedItem = selected;
            ApplyTerminalFont(selected);
        }
        catch { }
    }

    private void CmbFonts_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (CmbFonts.SelectedItem is string selectedFont)
        {
            ApplyTerminalFont(selectedFont);
        }
    }

    private void ApplyTerminalFont(string fontName)
    {
        try
        {
            var font = new FontFamily(fontName);
            RichAdLogs.FontFamily = font;
            RichCleanLogs.FontFamily = font;
            RichPerformanceLogs.FontFamily = font;
            RichAppsLogs.FontFamily = font;
            RichAnalyzerLogs.FontFamily = font;
            RichRemoteMaintLogs.FontFamily = font;
            RichRemoteProfilesLogs.FontFamily = font;
            RichAdUsersLogs.FontFamily = font;
            RichBancadaLogs.FontFamily = font;
            RichScriptsLogs.FontFamily = font;
            RichRemoteLogs.FontFamily = font;
            RichAppLogs.FontFamily = font;
            RichRegistryLogs.FontFamily = font;

            foreach (var block in RichAdLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichCleanLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichPerformanceLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichAppsLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichAnalyzerLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichRemoteMaintLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichRemoteProfilesLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichAdUsersLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichBancadaLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichScriptsLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichRemoteLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichAppLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichRegistryLogs.Document.Blocks) block.FontFamily = font;
        }
        catch { }
    }

    private void BtnSalvarFonte_Click(object sender, RoutedEventArgs e)
    {
        if (CmbFonts.SelectedItem is string selectedFont)
        {
            var dict = new Dictionary<string, string>
            {
                ["APP_TERMINAL_FONT"] = selectedFont
            };
            ConfigService.SaveConfig(dict);
            ApplyTerminalFont(selectedFont);
            TxtFooterStatus.Text = $"Fonte '{selectedFont}' gravada com sucesso!";
            System.Windows.MessageBox.Show($"Fonte '{selectedFont}' gravada com sucesso!", "Tipografia", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void LoadSettingsIntoUi()
    {
        ToggleStartWithWindows.IsChecked = _startupService.IsRunAtStartupEnabled();
        ToggleMinimizeToTray.IsChecked = ConfigService.MinimizeToTray;
        ToggleStartMenuShortcut.IsChecked = _startupService.IsStartMenuShortcutCreated();

        TxtCfgAdminUser.Text = ConfigService.AdminUser;
        TxtCfgAdminAdUser.Text = ConfigService.AdminAdUser;
        TxtCfgSccmServer.Text = ConfigService.SccmSiteServer;
        TxtCfgSccmCode.Text = ConfigService.SccmSiteCode;
        TxtCfgOtrs.Text = ConfigService.UrlOtrs;
        TxtCfgCitsmart.Text = ConfigService.UrlCitsmart;
        TxtCfgBancada.Text = ConfigService.UrlBancadaLocal;
        TxtCfgBancadaWslDir.Text = ConfigService.BancadaWslDir;
        TxtCfgBancadaWslDistro.Text = ConfigService.BancadaWslDistro;
        TxtCfgWhatsapp.Text = ConfigService.WhatsappPwaArgs;
        TxtCfgDeviceReportsDir.Text = ConfigService.DeviceReportsDir;
        TxtCfgCmRcViewer.Text = ConfigService.CmRcViewerCustomPath;
        TxtCfgFirefoxPath.Text = ConfigService.FirefoxBinaryPath;
        TxtCfgProtectedProfiles.Text = ConfigService.ProtectedProfiles;
        TxtCfgPingTimeout.Text = ConfigService.PingTimeoutMs.ToString();
    }

    private void ToggleStartWithWindows_Click(object sender, RoutedEventArgs e)
    {
        bool enable = ToggleStartWithWindows.IsChecked == true;
        bool ok = _startupService.SetRunAtStartup(enable);
        if (ok)
        {
            TxtFooterStatus.Text = enable
                ? "Inicialização com o Windows ATIVADA."
                : "Inicialização com o Windows DESATIVADA.";
        }
        else
        {
            ToggleStartWithWindows.IsChecked = !enable;
            TxtFooterStatus.Text = "Erro ao configurar inicialização no registro.";
        }
    }

    private void ToggleMinimizeToTray_Click(object sender, RoutedEventArgs e)
    {
        bool min = ToggleMinimizeToTray.IsChecked == true;
        ConfigService.SaveConfig(new Dictionary<string, string>
        {
            ["MINIMIZE_TO_TRAY"] = min ? "true" : "false"
        });
        TxtFooterStatus.Text = min
            ? "Minimizar para bandeja ATIVADO."
            : "Minimizar para bandeja DESATIVADO.";
    }

    private void ToggleStartMenuShortcut_Click(object sender, RoutedEventArgs e)
    {
        bool enable = ToggleStartMenuShortcut.IsChecked == true;
        bool ok = enable
            ? _startupService.CreateStartMenuShortcut()
            : _startupService.RemoveStartMenuShortcut();

        if (ok)
        {
            TxtFooterStatus.Text = enable
                ? "Atalho no Menu Iniciar CRIADO com sucesso."
                : "Atalho no Menu Iniciar REMOVIDO com sucesso.";
        }
        else
        {
            ToggleStartMenuShortcut.IsChecked = !enable;
            TxtFooterStatus.Text = "Falha ao atualizar atalho no Menu Iniciar.";
        }
    }

    #region Bandeja do Sistema (Tray Icon)
    private void InitializeSystemTray()
    {
        try
        {
            string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "app.ico");
            if (!File.Exists(iconPath))
            {
                iconPath = @"C:\AmbientesVirtuais\AmbientesHub\Assets\app.ico";
            }

            _trayIconService = new TrayIconService(this);
            _trayIconService.IsAdminModeProvider = () => SystemPerformanceService.IsUserAdministrator();
            _trayIconService.OnOpenRequested += RestoreFromTray;
            _trayIconService.OnExitRequested += ExitApplication;
            _trayIconService.OnToggleAdminModeRequested += ToggleAdminMode;
            _trayIconService.Initialize(iconPath, "Ambientes Hub — Produtividade & Performance");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Não foi possível inicializar o ícone de bandeja: {ex.Message}", nameof(MainWindow));
        }
    }

    private void RestoreFromTray()
    {
        Show();
        WindowState = WindowState.Normal;
        Activate();
        Focus();
    }

    private void ExitApplication()
    {
        _isRealExit = true;
        AppLoggerService.OnLogLineWritten -= OnNewLogLineReceived;
        _trayIconService?.Dispose();
        _trayIconService = null;
        System.Windows.Application.Current.Shutdown();
    }

    protected override void OnStateChanged(EventArgs e)
    {
        base.OnStateChanged(e);
        if (WindowState == WindowState.Minimized && ConfigService.MinimizeToTray)
        {
            Hide();
            _trayIconService?.ShowBalloon("Ambientes Hub", "O aplicativo continua ativo na bandeja do sistema.");
        }
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (!_isRealExit && ConfigService.MinimizeToTray)
        {
            e.Cancel = true;
            Hide();
            _trayIconService?.ShowBalloon("Ambientes Hub", "Ambientes Hub foi minimizado para a área de notificação.");
            return;
        }

        AppLoggerService.OnLogLineWritten -= OnNewLogLineReceived;
        _trayIconService?.Dispose();
        _trayIconService = null;
        base.OnClosing(e);
    }
    #endregion

    private void BtnSalvarConfiguracoes_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string chosenFont = CmbFonts.SelectedItem as string ?? "Consolas";
            var dict = new Dictionary<string, string>
            {
                ["APP_TERMINAL_FONT"] = chosenFont,
                ["ADMIN_USER"] = TxtCfgAdminUser.Text.Trim(),
                ["ADMIN_AD_USER"] = TxtCfgAdminAdUser.Text.Trim(),
                ["SCCM_SITE_SERVER"] = TxtCfgSccmServer.Text.Trim(),
                ["SCCM_SITE_CODE"] = TxtCfgSccmCode.Text.Trim(),
                ["URL_OTRS"] = TxtCfgOtrs.Text.Trim(),
                ["URL_CITSMART"] = TxtCfgCitsmart.Text.Trim(),
                ["URL_BANCADA_LOCAL"] = TxtCfgBancada.Text.Trim(),
                ["BANCADA_WSL_DIR"] = TxtCfgBancadaWslDir.Text.Trim(),
                ["BANCADA_WSL_DISTRO"] = TxtCfgBancadaWslDistro.Text.Trim(),
                ["WHATSAPP_PWA_ARGS"] = TxtCfgWhatsapp.Text.Trim(),
                ["DEVICE_REPORTS_DIR"] = TxtCfgDeviceReportsDir.Text.Trim(),
                ["CMRCVIEWER_PATH"] = TxtCfgCmRcViewer.Text.Trim(),
                ["FIREFOX_PATH"] = TxtCfgFirefoxPath.Text.Trim(),
                ["PROTECTED_PROFILES"] = TxtCfgProtectedProfiles.Text.Trim(),
                ["PING_TIMEOUT_MS"] = TxtCfgPingTimeout.Text.Trim()
            };

            ConfigService.SaveConfig(dict);
            ApplyTerminalFont(chosenFont);

            TxtFooterStatus.Text = "Configurações salvas no arquivo .env com sucesso!";
            System.Windows.MessageBox.Show("Configurações atualizadas com sucesso!", "Sucesso", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao salvar configurações: {ex.Message}", nameof(MainWindow));
            System.Windows.MessageBox.Show($"Falha ao salvar: {ex.Message}", "Erro", System.Windows.MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnResetConfigField_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button btn && btn.Tag is string key)
        {
            string defVal = ConfigService.GetRecommendedDefault(key);
            switch (key)
            {
                case "ADMIN_USER": TxtCfgAdminUser.Text = defVal; break;
                case "ADMIN_AD_USER": TxtCfgAdminAdUser.Text = defVal; break;
                case "SCCM_SITE_SERVER": TxtCfgSccmServer.Text = defVal; break;
                case "SCCM_SITE_CODE": TxtCfgSccmCode.Text = defVal; break;
                case "URL_OTRS": TxtCfgOtrs.Text = defVal; break;
                case "URL_CITSMART": TxtCfgCitsmart.Text = defVal; break;
                case "URL_BANCADA_LOCAL": TxtCfgBancada.Text = defVal; break;
                case "BANCADA_WSL_DIR": TxtCfgBancadaWslDir.Text = defVal; break;
                case "BANCADA_WSL_DISTRO": TxtCfgBancadaWslDistro.Text = defVal; break;
                case "WHATSAPP_PWA_ARGS": TxtCfgWhatsapp.Text = defVal; break;
                case "DEVICE_REPORTS_DIR": TxtCfgDeviceReportsDir.Text = defVal; break;
                case "CMRCVIEWER_PATH": TxtCfgCmRcViewer.Text = defVal; break;
                case "FIREFOX_PATH": TxtCfgFirefoxPath.Text = defVal; break;
                case "PROTECTED_PROFILES": TxtCfgProtectedProfiles.Text = defVal; break;
                case "PING_TIMEOUT_MS": TxtCfgPingTimeout.Text = defVal; break;
            }
            TxtFooterStatus.Text = $"Campo '{key}' preenchido com o valor recomendado padrão.";
        }
    }
    #endregion

    #region Scripts & Cofre de Credenciais
    private void BtnCheckAdminXml_Click(object sender, RoutedEventArgs e)
    {
        string xmlPath = @"C:\AmbientesVirtuais\Credenciais\cred_admin.xml";
        RichScriptsLogs.Document.Blocks.Clear();
        AppendColoredLog(RichScriptsLogs, "=== VERIFICAÇÃO DO COFRE LOCAL ===");

        var (success, msg, user) = CredentialService.VerifyCredential(xmlPath);
        if (success)
        {
            AppendColoredLog(RichScriptsLogs, $"[OK] Cofre localizado em: {xmlPath}");
            AppendColoredLog(RichScriptsLogs, $"     Usuário no cofre: {user}");
            AppendColoredLog(RichScriptsLogs, $"[SUCESSO] {msg}");
            AppLoggerService.Success($"Cofre local {xmlPath} verificado com sucesso para {user}", nameof(MainWindow));
        }
        else
        {
            AppendColoredLog(RichScriptsLogs, $"[ERRO] {msg}");
            AppLoggerService.Warning($"Verificação do cofre local falhou: {msg}", nameof(MainWindow));
        }
    }

    private void BtnCheckAdminAdXml_Click(object sender, RoutedEventArgs e)
    {
        string xmlPath = @"C:\AmbientesVirtuais\Credenciais\cred_admin_ad.xml";
        RichScriptsLogs.Document.Blocks.Clear();
        AppendColoredLog(RichScriptsLogs, "=== VERIFICAÇÃO DO COFRE AD ===");

        var (success, msg, user) = CredentialService.VerifyCredential(xmlPath);
        if (success)
        {
            AppendColoredLog(RichScriptsLogs, $"[OK] Cofre localizado em: {xmlPath}");
            AppendColoredLog(RichScriptsLogs, $"     Usuário no cofre: {user}");
            AppendColoredLog(RichScriptsLogs, $"[SUCESSO] {msg}");
            AppLoggerService.Success($"Cofre AD {xmlPath} verificado com sucesso para {user}", nameof(MainWindow));
        }
        else
        {
            AppendColoredLog(RichScriptsLogs, $"[ERRO] {msg}");
            AppLoggerService.Warning($"Verificação do cofre AD falhou: {msg}", nameof(MainWindow));
        }
    }

    private void BtnCriarAdminAdXml_Click(object sender, RoutedEventArgs e)
    {
        string targetUser = ConfigService.AdminAdUser;
        string destPath = @"C:\AmbientesVirtuais\Credenciais\cred_admin_ad.xml";

        var dialog = new CredentialDialog(targetUser)
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true)
        {
            RichScriptsLogs.Document.Blocks.Clear();
            AppendColoredLog(RichScriptsLogs, $"=== GERANDO COFRE SEGURO DPAPI PARA: {dialog.Username} ===");
            AppLoggerService.Info($"Iniciando gravação de cofre seguro para {dialog.Username}", nameof(MainWindow));

            var (success, msg) = CredentialService.SaveCredential(destPath, dialog.Username, dialog.Password);
            if (success)
            {
                AppendColoredLog(RichScriptsLogs, $"[OK] Credencial DPAPI gerada com sucesso!");
                AppendColoredLog(RichScriptsLogs, $"     Arquivo : {destPath}");
                AppendColoredLog(RichScriptsLogs, $"     Usuário : {dialog.Username}");
                AppendColoredLog(RichScriptsLogs, $"[SUCESSO] {msg}");
                TxtFooterStatus.Text = $"Cofre para {dialog.Username} criado com sucesso.";
                AppLoggerService.Success($"Cofre {destPath} gerado e validado para {dialog.Username}.", nameof(MainWindow));
            }
            else
            {
                AppendColoredLog(RichScriptsLogs, $"[ERRO] {msg}");
                AppLoggerService.Danger($"Falha ao gravar cofre: {msg}", nameof(MainWindow));
            }
        }
        else
        {
            RichScriptsLogs.Document.Blocks.Clear();
            AppendColoredLog(RichScriptsLogs, "[i] Operação cancelada pelo usuário.");
            AppLoggerService.Info("Geração de cofre cancelada pelo usuário no diálogo.", nameof(MainWindow));
        }
    }
    #endregion

    #region Color Terminal Helper
    private void AppendColoredLog(System.Windows.Controls.RichTextBox rtb, string text)
    {
        if (Dispatcher.CheckAccess())
        {
            AddLogBlock(rtb, text);
        }
        else
        {
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
            {
                AddLogBlock(rtb, text);
            }));
        }
    }

    private void AddLogBlock(System.Windows.Controls.RichTextBox rtb, string text)
    {
        var brush = new SolidColorBrush(Color.FromRgb(203, 213, 225)); // Slate 300 padrão

        if (text.Contains("[OK]") || text.Contains("[SUCESSO]") || text.Contains("concluída"))
        {
            brush = new SolidColorBrush(Color.FromRgb(52, 211, 153)); // Emerald 400
        }
        else if (text.Contains("[ERRO]") || text.Contains("Falha") || text.Contains("Erro SCCM") || text.Contains("EXCLUSÃO"))
        {
            brush = new SolidColorBrush(Color.FromRgb(248, 113, 113)); // Red 400
        }
        else if (text.Contains("[i]") || text.Contains("Aviso", StringComparison.OrdinalIgnoreCase) || text.Contains("não localizado") || text.Contains("Dica:"))
        {
            brush = new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Amber 400
        }
        else if (text.StartsWith("===") || (text.StartsWith("[") && text.Contains("/")))
        {
            brush = new SolidColorBrush(Color.FromRgb(56, 189, 248)); // Sky 400
        }

        var currentFont = rtb.FontFamily;
        string cleanText = text.TrimEnd('\r', '\n');

        var run = new Run(cleanText)
        {
            Foreground = brush,
            FontFamily = currentFont
        };

        var para = new Paragraph(run)
        {
            Margin = new Thickness(0, 0, 0, 0),
            LineHeight = 18,
            FontFamily = currentFont
        };

        rtb.Document.Blocks.Add(para);
        rtb.ScrollToEnd();
    }
    #endregion

    #region Lançador de Ambientes
    private void BtnModoChamados_Click(object sender, RoutedEventArgs e) => LaunchEnv("Chamados");
    private void BtnModoContratos_Click(object sender, RoutedEventArgs e) => LaunchEnv("Contratos");
    private void BtnModoImpressoras_Click(object sender, RoutedEventArgs e) => LaunchEnv("Impressoras");
    private void BtnModoPxe_Click(object sender, RoutedEventArgs e) => LaunchEnv("PXE");
    private void BtnModoEstudo_Click(object sender, RoutedEventArgs e) => LaunchEnv("Estudo");

    private async void LaunchEnv(string modeName)
    {
        try
        {
            AppLoggerService.Info($"Usuário clicou para acionar: Modo {modeName}", nameof(MainWindow));
            TxtFooterStatus.Text = $"Preparando ambiente: Modo {modeName}...";

            await Task.Run(() =>
            {
                _envService.LaunchEnvironment(modeName, cleanFirst: true);
            });

            TxtFooterStatus.Text = $"Ambiente ativo: Modo {modeName}";
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Falha ao carregar ambiente {modeName}: {ex.Message}", nameof(MainWindow));
            System.Windows.MessageBox.Show($"Erro ao carregar o ambiente: {ex.Message}", "Falha", System.Windows.MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    #endregion

    #region Acesso Remoto & Diagnóstico
    private async void BtnConsultarRemoto_Click(object sender, RoutedEventArgs e)
    {
        string query = TxtComputerName.Text.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            AppendColoredLog(RichRemoteLogs, "[Aviso] Informe parte do nome da máquina (ex: 80600) ou usuário para pesquisar.");
            return;
        }

        BtnConsultarRemoto.IsEnabled = false;
        PlaceholderRemoteCard.Visibility = Visibility.Collapsed;
        ItemsMachineResults.ItemsSource = null;
        LoadingRemoteCard.Visibility = Visibility.Visible;
        RichRemoteLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Pesquisando máquinas e usuários para '{query}'...";

        try
        {
            var results = await _adSccmService.SearchMachinesAsync(query, msg =>
            {
                AppendColoredLog(RichRemoteLogs, msg);
            });

            if (results.Count > 0)
            {
                ItemsMachineResults.ItemsSource = results;
                PlaceholderRemoteCard.Visibility = Visibility.Collapsed;
                TxtFooterStatus.Text = $"{results.Count} máquina(s) encontrada(s) para '{query}'.";
            }
            else
            {
                PlaceholderRemoteCard.Visibility = Visibility.Visible;
                AppendColoredLog(RichRemoteLogs, $"[i] Nenhum registro encontrado para '{query}' no AD ou SCCM.");
                TxtFooterStatus.Text = $"Nenhum resultado para '{query}'.";
            }
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRemoteLogs, $"[ERRO] Falha na pesquisa: {ex.Message}");
            AppLoggerService.Danger($"Erro na pesquisa de '{query}': {ex.Message}", nameof(MainWindow));
            PlaceholderRemoteCard.Visibility = Visibility.Visible;
        }
        finally
        {
            LoadingRemoteCard.Visibility = Visibility.Collapsed;
            BtnConsultarRemoto.IsEnabled = true;
        }
    }

    private void BtnCardConnectSccm_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details && !string.IsNullOrWhiteSpace(details.ComputerName))
        {
            ConectarSccm(details.ComputerName);
        }
    }

    private void BtnCardConnectRdp_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details && !string.IsNullOrWhiteSpace(details.ComputerName))
        {
            ConectarRdp(details.ComputerName);
        }
    }

    private void BtnCardOpenExplorer_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details && !string.IsNullOrWhiteSpace(details.ComputerName))
        {
            AbrirExploradorRemoto(details.ComputerName);
        }
    }

    private void AbrirExploradorRemoto(string host)
    {
        try
        {
            AppLoggerService.Info($"Abrindo Explorador de Arquivos para {host}", nameof(MainWindow));
            _remoteToolsService.OpenRemoteExplorer(host);
            AppendColoredLog(RichRemoteLogs, $"[SUCESSO] Explorador de arquivos aberto em \\\\{host}\\c$ (sessão autenticada)");
            TxtFooterStatus.Text = $"Explorador de Arquivos aberto em \\\\{host}\\c$";
            AppLoggerService.Success($"Explorador aberto para \\\\{host}\\c$", nameof(MainWindow));
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao abrir Explorador para {host}: {ex.Message}", nameof(MainWindow));
            AppendColoredLog(RichRemoteLogs, $"[ERRO] Falha ao abrir Explorador para {host}: {ex.Message}");
            TxtFooterStatus.Text = $"Erro ao acessar \\\\{host}\\c$";
        }
    }

    private void BtnCardCopyMachine_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details && !string.IsNullOrWhiteSpace(details.ComputerName))
        {
            if (ClipboardHelper.Copy(details.ComputerName))
            {
                TxtFooterStatus.Text = $"Nome da máquina '{details.ComputerName}' copiado!";
            }
        }
    }

    private void BtnCardCopyUser_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details)
        {
            string user = !string.IsNullOrWhiteSpace(details.CurrentLogonUser) ? details.CurrentLogonUser : details.PrimaryUser;
            if (!string.IsNullOrWhiteSpace(user) && ClipboardHelper.Copy(user))
            {
                TxtFooterStatus.Text = $"Usuário '{user}' copiado!";
            }
            else
            {
                TxtFooterStatus.Text = "Nenhum usuário disponível para cópia nesta máquina.";
            }
        }
    }

    private void BtnCardCopyMac_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details && !string.IsNullOrWhiteSpace(details.MacAddress))
        {
            if (ClipboardHelper.Copy(details.MacAddress))
            {
                TxtFooterStatus.Text = $"MAC '{details.MacAddress}' copiado!";
            }
        }
    }

    private void BtnCardCopyIp_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is MachineDetails details && !string.IsNullOrWhiteSpace(details.IpAddress))
        {
            if (ClipboardHelper.Copy(details.IpAddress))
            {
                TxtFooterStatus.Text = $"IP '{details.IpAddress}' copiado!";
            }
        }
    }

    private void ConectarSccm(string host)
    {
        try
        {
            AppLoggerService.Info($"Iniciando SCCM Remote para {host}", nameof(MainWindow));
            _remoteToolsService.OpenSccmRemoteControl(host);
            AppendColoredLog(RichRemoteLogs, $"[SUCESSO] Conexão SCCM disparada para: {host}");
            TxtFooterStatus.Text = $"Conexão SCCM: {host}";
            AppLoggerService.Success($"SCCM Remote disparado para {host}", nameof(MainWindow));
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao abrir SCCM Remote para {host}: {ex.Message}", nameof(MainWindow));
            AppendColoredLog(RichRemoteLogs, $"[ERRO] Falha ao abrir SCCM: {ex.Message}");
        }
    }

    private void ConectarRdp(string host)
    {
        try
        {
            AppLoggerService.Info($"Iniciando RDP para {host}", nameof(MainWindow));
            _remoteToolsService.OpenRdp(host);
            AppendColoredLog(RichRemoteLogs, $"[SUCESSO] Conexão RDP (mstsc) disparada para: {host}");
            TxtFooterStatus.Text = $"Conexão RDP: {host}";
            AppLoggerService.Success($"RDP disparado para {host}", nameof(MainWindow));
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao abrir RDP para {host}: {ex.Message}", nameof(MainWindow));
            AppendColoredLog(RichRemoteLogs, $"[ERRO] Falha ao abrir RDP: {ex.Message}");
        }
    }

    private void TxtComputerName_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnConsultarRemoto_Click(sender, e);
        }
    }
    #endregion

    #region Limpeza AD & SCCM (Color Terminal)
    private async void BtnAdQuery_Click(object sender, RoutedEventArgs e)
    {
        string host = TxtAdTargetComputer.Text;
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendColoredLog(RichAdLogs, "[Aviso] Informe o nome do computador para consultar.");
            return;
        }

        BtnAdQuery.IsEnabled = false;
        BtnAdClean.IsEnabled = false;
        RichAdLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Consultando máquina {host} no AD e SCCM...";

        await Task.Run(() =>
        {
            _adSccmService.QueryComputer(host, msg =>
            {
                AppendColoredLog(RichAdLogs, msg);
            });
        });

        BtnAdQuery.IsEnabled = true;
        BtnAdClean.IsEnabled = true;
        TxtFooterStatus.Text = $"Consulta concluída para {host}";
    }

    private async void BtnAdClean_Click(object sender, RoutedEventArgs e)
    {
        string host = TxtAdTargetComputer.Text;
        if (string.IsNullOrWhiteSpace(host))
        {
            AppLoggerService.Warning("Desprovisionamento acionado com campo em branco.", nameof(MainWindow));
            AppendColoredLog(RichAdLogs, "[Aviso] Informe o nome do computador para desprovisionar.");
            return;
        }

        var result = System.Windows.MessageBox.Show(
            $"Deseja realmente remover o registro de '{host}' do Active Directory e do SCCM?",
            "Confirmação de Exclusão",
            System.Windows.MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != System.Windows.MessageBoxResult.Yes) return;

        AppLoggerService.Info($"Iniciando rotina de exclusão AD/SCCM para {host}", nameof(MainWindow));
        BtnAdQuery.IsEnabled = false;
        BtnAdClean.IsEnabled = false;
        RichAdLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Removendo máquina {host} do AD e SCCM...";

        await Task.Run(() =>
        {
            _adSccmService.CleanComputer(host, msg =>
            {
                AppendColoredLog(RichAdLogs, msg);
            });
        });

        BtnAdQuery.IsEnabled = true;
        BtnAdClean.IsEnabled = true;
        TxtFooterStatus.Text = $"Exclusão concluída para {host}";
        AppLoggerService.Success($"Rotina de exclusão concluída para {host}", nameof(MainWindow));
    }

    private void TxtAdTargetComputer_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnAdQuery_Click(sender, e);
        }
    }
    #endregion

    #region Performance & Limpeza (Color Terminal)
    private async void BtnLimparAmbiente_Click(object sender, RoutedEventArgs e)
    {
        var options = new CleanupOptions
        {
            CloseOffice = ChkOffice.IsChecked ?? true,
            CloseBrowsers = ChkBrowsers.IsChecked ?? true,
            CloseFirefox = ChkFirefox.IsChecked ?? false,
            CloseTeams = ChkTeams.IsChecked ?? true,
            CloseAdministrativeTools = ChkAdminTools.IsChecked ?? true,
            CloseUtilities = ChkUtilities.IsChecked ?? true,
            StopTelemetry = ChkTelemetry.IsChecked ?? true,
            RestartExplorer = ChkRestartExplorer.IsChecked ?? false
        };

        AppLoggerService.Info("Limpeza manual com opções personalizadas iniciada.", nameof(MainWindow));
        BtnLimparAmbiente.IsEnabled = false;
        RichCleanLogs.Document.Blocks.Clear();

        await Task.Run(() =>
        {
            _cleanupService.CleanEnvironment(options, msg =>
            {
                AppendColoredLog(RichCleanLogs, $"[{DateTime.Now:HH:mm:ss}] {msg}");
            });
        });

        BtnLimparAmbiente.IsEnabled = true;
        TxtFooterStatus.Text = "Limpeza seletiva concluída com sucesso!";
        AppLoggerService.Success("Limpeza seletiva concluída com sucesso.", nameof(MainWindow));
    }
    #endregion

    #region Performance & Auditoria (ArrumarPC Nativo)
    private void RefreshPerformanceAdminStatus()
    {
        bool isAdmin = SystemPerformanceService.IsUserAdministrator();

        // Atualização da Badge da Barra de Título (ao lado dos botões de controle da janela)
        if (TitleBarAdminBadge != null && TxtTitleBarAdmin != null && IconTitleBarAdmin != null)
        {
            if (isAdmin)
            {
                TxtTitleBarAdmin.Text = "Admin";
                TxtTitleBarAdmin.Foreground = new SolidColorBrush(Color.FromRgb(52, 211, 153)); // Emerald 400
                IconTitleBarAdmin.Foreground = new SolidColorBrush(Color.FromRgb(52, 211, 153));
                TitleBarAdminBadge.Background = new SolidColorBrush(Color.FromArgb(40, 16, 185, 129));
                TitleBarAdminBadge.BorderBrush = new SolidColorBrush(Color.FromArgb(120, 16, 185, 129));
                TitleBarAdminBadge.ToolTip = "Executando com Privilégios de Administrador. Clique para alternar para Usuário Padrão.";
            }
            else
            {
                TxtTitleBarAdmin.Text = "Usuário";
                TxtTitleBarAdmin.Foreground = new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Amber 400
                IconTitleBarAdmin.Foreground = new SolidColorBrush(Color.FromRgb(251, 191, 36));
                TitleBarAdminBadge.Background = new SolidColorBrush(Color.FromArgb(40, 245, 158, 11));
                TitleBarAdminBadge.BorderBrush = new SolidColorBrush(Color.FromArgb(120, 245, 158, 11));
                TitleBarAdminBadge.ToolTip = "Executando em Modo Usuário Padrão. Clique para reiniciar como Administrador (UAC).";
            }
        }

        if (TxtPerfAdminStatus == null || BadgeAdminStatus == null || BtnReiniciarComoAdmin == null)
            return;

        if (isAdmin)
        {
            TxtPerfAdminStatus.Text = "Modo Administrador";
            TxtPerfAdminStatus.Foreground = new SolidColorBrush(Color.FromRgb(52, 211, 153));
            BadgeAdminStatus.Background = new SolidColorBrush(Color.FromArgb(40, 16, 185, 129));
            BadgeAdminStatus.BorderBrush = new SolidColorBrush(Color.FromArgb(100, 16, 185, 129));
            BtnReiniciarComoAdmin.Visibility = Visibility.Collapsed;
        }
        else
        {
            TxtPerfAdminStatus.Text = "Modo Usuário Padrão";
            TxtPerfAdminStatus.Foreground = new SolidColorBrush(Color.FromRgb(251, 191, 36));
            BadgeAdminStatus.Background = new SolidColorBrush(Color.FromArgb(40, 245, 158, 11));
            BadgeAdminStatus.BorderBrush = new SolidColorBrush(Color.FromArgb(100, 245, 158, 11));
            BtnReiniciarComoAdmin.Visibility = Visibility.Visible;
        }
    }

    private void TitleBarAdminBadge_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
        {
            ToggleAdminMode();
        }
    }

    private void ToggleAdminMode()
    {
        bool isAdmin = SystemPerformanceService.IsUserAdministrator();
        if (isAdmin)
        {
            // Se já for admin, relança em modo normal via explorer.exe (sem herdar token elevado)
            try
            {
                string? exePath = Environment.ProcessPath;
                if (!string.IsNullOrWhiteSpace(exePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = $"\"{exePath}\"",
                        UseShellExecute = true
                    });
                    ExitApplication();
                }
            }
            catch (Exception ex)
            {
                AppLoggerService.Warning($"Erro ao reiniciar em modo padrão: {ex.Message}", nameof(MainWindow));
            }
        }
        else
        {
            BtnReiniciarComoAdmin_Click(this, new RoutedEventArgs());
        }
    }

    private void BtnReiniciarComoAdmin_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string? exePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exePath))
                return;

            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true,
                Verb = "runas"
            });

            Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichPerformanceLogs, $"[AVISO] Elevação UAC cancelada ou indisponível: {ex.Message}");
        }
    }

    private async void BtnExecutarAuditoria_Click(object sender, RoutedEventArgs e)
    {
        BtnExecutarAuditoria.IsEnabled = false;
        BtnAplicarOtimizacao.IsEnabled = false;
        BtnReverterServicos.IsEnabled = false;
        RichPerformanceLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = "Executando auditoria completa de performance...";

        try
        {
            var report = await _perfService.RunFullAuditAsync(msg => AppendColoredLog(RichPerformanceLogs, msg));
            UpdatePerformanceKpis(report);
            TxtFooterStatus.Text = "Auditoria de performance concluída com sucesso!";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichPerformanceLogs, $"[ERRO] Falha durante a auditoria: {ex.Message}");
            TxtFooterStatus.Text = "Falha durante a auditoria de performance.";
        }
        finally
        {
            BtnExecutarAuditoria.IsEnabled = true;
            BtnAplicarOtimizacao.IsEnabled = true;
            BtnReverterServicos.IsEnabled = true;
        }
    }

    private async void BtnAplicarOtimizacao_Click(object sender, RoutedEventArgs e)
    {
        string selectedWslMem = (CmbWslMemory.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "3GB";

        var options = new OptimizationOptions
        {
            DisableTelemetryServices = ChkOptTelemetryServices.IsChecked == true,
            ApplyRegistryTelemetryTweaks = ChkOptTelemetryReg.IsChecked == true,
            RemoveBloatwareApps = ChkOptBloatApps.IsChecked == true,
            DisableHibernation = ChkOptHibernation.IsChecked == true,
            EnableSsdTrim = ChkOptTrim.IsChecked == true,
            DisableBingSearch = ChkOptBingSearch.IsChecked == true,
            AccelerateUiMenus = ChkOptUiSpeed.IsChecked == true,
            ConfigureWslLimit = ChkOptWslConfig.IsChecked == true,
            WslMemoryLimit = selectedWslMem,
            CleanDockerBuilderCache = ChkOptDockerCache.IsChecked == true
        };

        var confirm = System.Windows.MessageBox.Show(
            "Deseja aplicar as otimizações selecionadas nesta estação de trabalho?",
            "Aplicar Otimizações de Performance",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        BtnExecutarAuditoria.IsEnabled = false;
        BtnAplicarOtimizacao.IsEnabled = false;
        BtnReverterServicos.IsEnabled = false;
        RichPerformanceLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = "Aplicando otimizações e debloat no sistema...";

        try
        {
            await _perfService.ApplyOptimizationsAsync(options, msg => AppendColoredLog(RichPerformanceLogs, msg));
            AppendColoredLog(RichPerformanceLogs, "\n[i] Executando check-up e atualização dos KPIs...");
            var postAudit = await _perfService.RunFullAuditAsync(msg => { });
            UpdatePerformanceKpis(postAudit);
            TxtFooterStatus.Text = "Otimizações aplicadas com sucesso!";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichPerformanceLogs, $"[ERRO] Falha ao aplicar otimizações: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao aplicar otimizações.";
        }
        finally
        {
            BtnExecutarAuditoria.IsEnabled = true;
            BtnAplicarOtimizacao.IsEnabled = true;
            BtnReverterServicos.IsEnabled = true;
        }
    }

    private async void BtnReverterServicos_Click(object sender, RoutedEventArgs e)
    {
        var confirm = System.Windows.MessageBox.Show(
            "Deseja restaurar os serviços de telemetria/diagnóstico para o padrão de inicialização do Windows?",
            "Rollback de Serviços",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        BtnExecutarAuditoria.IsEnabled = false;
        BtnAplicarOtimizacao.IsEnabled = false;
        BtnReverterServicos.IsEnabled = false;
        RichPerformanceLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = "Revertendo serviços do Windows...";

        try
        {
            await _perfService.RollbackTelemetryServicesAsync(msg => AppendColoredLog(RichPerformanceLogs, msg));
            TxtFooterStatus.Text = "Rollback de serviços concluído.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichPerformanceLogs, $"[ERRO] Falha no rollback: {ex.Message}");
        }
        finally
        {
            BtnExecutarAuditoria.IsEnabled = true;
            BtnAplicarOtimizacao.IsEnabled = true;
            BtnReverterServicos.IsEnabled = true;
        }
    }

    private void UpdatePerformanceKpis(SystemAuditReport report)
    {
        RefreshPerformanceAdminStatus();

        TxtPerfRamKpi.Text = report.TotalRamGb > 0
            ? $"{report.FreeRamGb:F2} / {report.TotalRamGb:F2} GB ({report.FreeRamPercent:F0}%)"
            : "Indisponível";

        var cDrive = report.Disks.FirstOrDefault(d => d.DriveName.StartsWith("C", StringComparison.OrdinalIgnoreCase));
        string hiberTxt = report.HiberfilExists ? $"hiberfil {report.HiberfilSizeGb:F1}GB" : "sem hibernação";
        TxtPerfDiskKpi.Text = cDrive != null
            ? $"C: {cDrive.FreeGb:F1} GB livres ({hiberTxt})"
            : hiberTxt;

        TxtPerfTrimKpi.Text = report.IsTrimEnabled switch
        {
            true => "Ativo (SSD otimizado)",
            false => "Desativado (atenção)",
            null => "Não verificado"
        };

        TxtPerfBloatKpi.Text = report.DetectedBloatApps.Count == 0
            ? "Nenhum detectado"
            : $"{report.DetectedBloatApps.Count} app(s) encontrado(s)";
    }
    #endregion

    #region Apps Instalados & Provisionamento
    private void RefreshAppsAdminStatus()
    {
        if (TxtAppsAdminStatus == null || BadgeAppsAdminStatus == null || BtnAppsReiniciarAdmin == null)
            return;

        bool isAdmin = SystemPerformanceService.IsUserAdministrator();
        if (isAdmin)
        {
            TxtAppsAdminStatus.Text = "Modo Administrador";
            TxtAppsAdminStatus.Foreground = new SolidColorBrush(Color.FromRgb(52, 211, 153));
            BadgeAppsAdminStatus.Background = new SolidColorBrush(Color.FromArgb(40, 16, 185, 129));
            BadgeAppsAdminStatus.BorderBrush = new SolidColorBrush(Color.FromArgb(100, 16, 185, 129));
            BtnAppsReiniciarAdmin.Visibility = Visibility.Collapsed;
        }
        else
        {
            TxtAppsAdminStatus.Text = "Modo Usuário Padrão";
            TxtAppsAdminStatus.Foreground = new SolidColorBrush(Color.FromRgb(251, 191, 36));
            BadgeAppsAdminStatus.Background = new SolidColorBrush(Color.FromArgb(40, 245, 158, 11));
            BadgeAppsAdminStatus.BorderBrush = new SolidColorBrush(Color.FromArgb(100, 245, 158, 11));
            BtnAppsReiniciarAdmin.Visibility = Visibility.Visible;
        }
    }

    private async Task RefreshAppsCatalogAsync()
    {
        if (BtnVerificarApps != null)
            BtnVerificarApps.IsEnabled = false;

        try
        {
            await _softwareService.RefreshStatusAsync(_softwareCatalog);
            TxtFooterStatus.Text = "Status de aplicativos instalados atualizado.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAppsLogs, $"[ERRO] Falha ao verificar aplicativos: {ex.Message}");
        }
        finally
        {
            if (BtnVerificarApps != null)
                BtnVerificarApps.IsEnabled = true;
        }
    }

    private async void BtnVerificarApps_Click(object sender, RoutedEventArgs e)
    {
        RichAppsLogs.Document.Blocks.Clear();
        AppendColoredLog(RichAppsLogs, "[i] Verificando presença dos softwares administrativos na estação...");
        await RefreshAppsCatalogAsync();
        foreach (var item in _softwareCatalog)
        {
            AppendColoredLog(RichAppsLogs, item.IsInstalled
                ? $"  [OK] {item.Name} — {item.StatusDetail}"
                : $"  [AVISO] {item.Name} — {item.StatusDetail}");
        }
    }

    private async void BtnInstallSoftware_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not SoftwarePackageItem pkg)
            return;

        if (pkg.IsInstalling)
            return;

        bool silent = ChkInstallSilent.IsChecked == true;
        pkg.IsInstalling = true;
        if (BtnVerificarApps != null) BtnVerificarApps.IsEnabled = false;
        RichAppsLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Instalando {pkg.Name}...";

        try
        {
            bool success = await _softwareService.InstallSoftwareAsync(
                pkg,
                silent,
                msg => AppendColoredLog(RichAppsLogs, msg));

            TxtFooterStatus.Text = success
                ? $"{pkg.Name} pronto para uso!"
                : $"Instalação de {pkg.Name} finalizada com avisos.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAppsLogs, $"[ERRO] Exceção ao instalar {pkg.Name}: {ex.Message}");
            TxtFooterStatus.Text = $"Erro ao instalar {pkg.Name}.";
        }
        finally
        {
            pkg.IsInstalling = false;
            if (BtnVerificarApps != null) BtnVerificarApps.IsEnabled = true;
        }
    }

    private async void BtnUninstallSoftware_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not SoftwarePackageItem pkg)
            return;

        var confirm = System.Windows.MessageBox.Show(
            $"Deseja realmente desinstalar '{pkg.Name}' desta máquina?",
            "Confirmar Desinstalação",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        bool silent = ChkInstallSilent.IsChecked == true;
        pkg.IsInstalling = true;
        if (BtnVerificarApps != null) BtnVerificarApps.IsEnabled = false;
        RichAppsLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Desinstalando {pkg.Name}...";

        try
        {
            bool ok = await _softwareService.UninstallSoftwareAsync(
                pkg,
                silent,
                msg => AppendColoredLog(RichAppsLogs, msg));

            TxtFooterStatus.Text = ok
                ? $"{pkg.Name} desinstalado com sucesso!"
                : $"Desinstalação de {pkg.Name} concluída com avisos.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAppsLogs, $"[ERRO] Falha ao desinstalar {pkg.Name}: {ex.Message}");
            TxtFooterStatus.Text = $"Erro ao desinstalar {pkg.Name}.";
        }
        finally
        {
            pkg.IsInstalling = false;
            if (BtnVerificarApps != null) BtnVerificarApps.IsEnabled = true;
        }
    }
    #endregion

    #region Painel 6: Analisador de Dispositivos Remotos
    private void TxtAnalyzerComputer_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnRunAnalyzer_Click(this, new RoutedEventArgs());
        }
    }

    private async void BtnRunAnalyzer_Click(object sender, RoutedEventArgs e)
    {
        string target = string.IsNullOrWhiteSpace(TxtAnalyzerComputer.Text) ? "." : TxtAnalyzerComputer.Text.Trim();
        bool skipMajor = ChkAnalyzerSkipMajor.IsChecked == true;
        bool openFirefox = ChkAnalyzerOpenFirefox.IsChecked == true;

        BtnRunAnalyzer.IsEnabled = false;
        BtnOpenReportHtml.IsEnabled = false;
        BtnOpenReportPdf.IsEnabled = false;
        BtnOpenReportExcel.IsEnabled = false;
        RichAnalyzerLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Analisando dispositivo '{target}' e gerando relatórios HTML, PDF e Excel...";

        try
        {
            var res = await Task.Run(() => _deviceAnalyzerService.AnalyzeDevice(
                target,
                skipMajorData: skipMajor,
                autoOpenHtmlInFirefox: openFirefox,
                outputFolder: null,
                onLog: msg => Dispatcher.Invoke(() => AppendColoredLog(RichAnalyzerLogs, msg))));

            _lastAnalyzerResult = res;

            if (res.Data != null)
            {
                var d = res.Data;
                TxtAnlKpiModel.Text = $"{d.HardwareInfo.Computador} ({d.OS.NomeMaquina})";
                TxtAnlKpiUser.Text = string.IsNullOrWhiteSpace(d.CurrentUser)
                    ? "Nenhum usuário interativo"
                    : (string.IsNullOrWhiteSpace(d.CurrentUserDisplayName) ? d.CurrentUser : $"{d.CurrentUser} ({d.CurrentUserDisplayName})");
                TxtAnlKpiOs.Text = $"{d.OS.VersaoWindowsHnH} | RAM: {d.HardwareInfo.MemoriaFisicaGb}";
                var firstDisk = d.LogicalDisks.FirstOrDefault(ld => ld.DeviceId.StartsWith("C", StringComparison.OrdinalIgnoreCase)) ?? d.LogicalDisks.FirstOrDefault();
                string diskSummary = firstDisk != null ? $"C: {firstDisk.EspacoLivre} livres" : "Disco N/D";
                TxtAnlKpiDisk.Text = $"{diskSummary} | {d.Monitors.Count} monitor(es)";
            }

            BtnOpenReportHtml.IsEnabled = !string.IsNullOrWhiteSpace(res.HtmlPath) && File.Exists(res.HtmlPath);
            BtnOpenReportPdf.IsEnabled = !string.IsNullOrWhiteSpace(res.PdfPath) && File.Exists(res.PdfPath);
            BtnOpenReportExcel.IsEnabled = !string.IsNullOrWhiteSpace(res.ExcelPath) && File.Exists(res.ExcelPath);

            TxtFooterStatus.Text = res.Success
                ? $"Análise de '{target}' concluída! Relatórios salvos em {res.OutputFolder}"
                : $"Análise de '{target}' finalizada com alertas.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAnalyzerLogs, $"[ERRO] Exceção no analisador: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao executar análise de dispositivo.";
        }
        finally
        {
            BtnRunAnalyzer.IsEnabled = true;
        }
    }

    private void BtnOpenReportFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string folder = _lastAnalyzerResult?.OutputFolder ?? DeviceAnalyzerService.DefaultOutputFolder;
            Directory.CreateDirectory(folder);
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{folder}\"",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAnalyzerLogs, $"[ERRO] Não foi possível abrir a pasta: {ex.Message}");
        }
    }

    private void BtnOpenReportHtml_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastAnalyzerResult?.HtmlPath) || !File.Exists(_lastAnalyzerResult.HtmlPath))
            return;

        try
        {
            var browser = new BrowserLauncherService();
            browser.OpenUrlsInFirefox([_lastAnalyzerResult.HtmlPath]);
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAnalyzerLogs, $"[ERRO] Falha ao abrir HTML no Firefox: {ex.Message}");
        }
    }

    private void BtnOpenReportPdf_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastAnalyzerResult?.PdfPath) || !File.Exists(_lastAnalyzerResult.PdfPath))
            return;

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _lastAnalyzerResult.PdfPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAnalyzerLogs, $"[ERRO] Falha ao abrir PDF: {ex.Message}");
        }
    }

    private void BtnOpenReportExcel_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastAnalyzerResult?.ExcelPath) || !File.Exists(_lastAnalyzerResult.ExcelPath))
            return;

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _lastAnalyzerResult.ExcelPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAnalyzerLogs, $"[ERRO] Falha ao abrir planilha Excel: {ex.Message}");
        }
    }
    #endregion

    #region Painel 7: Manutenção & Limpeza Remota
    private void TxtRemoteMaintComputer_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnAnalyzeRemoteMaint_Click(this, new RoutedEventArgs());
        }
    }

    private RemoteMaintenanceOptions BuildRemoteMaintOptionsFromUi()
    {
        return new RemoteMaintenanceOptions
        {
            CleanTempFiles = ChkRmTemp.IsChecked == true,
            CleanPrefetch = ChkRmPrefetch.IsChecked == true,
            CleanRecycleBin = ChkRmRecycle.IsChecked == true,
            CleanWindowsUpdateCache = ChkRmWuCache.IsChecked == true,
            CleanCbsLogs = ChkRmCbs.IsChecked == true,
            CleanThumbnailCache = ChkRmThumbnails.IsChecked == true,
            CleanDeliveryOptimization = ChkRmDeliveryOpt.IsChecked == true,
            CleanCrashDumps = ChkRmCrashDumps.IsChecked == true,
            CleanStoreCache = ChkRmStore.IsChecked == true,
            CleanEventLogs = ChkRmEventLogs.IsChecked == true,
            FlushDnsCache = ChkRmDns.IsChecked == true,
            CleanWinSxS = ChkRmWinSxS.IsChecked == true,
            CleanWindowsOld = ChkRmWindowsOld.IsChecked == true,
            RunRepairVolume = ChkRmRepairVol.IsChecked == true,
            RunDefragIfHdd = ChkRmDefragHdd.IsChecked == true,
            RunSfcIfCorrupted = ChkRmSfc.IsChecked == true
        };
    }

    private async void BtnAnalyzeRemoteMaint_Click(object sender, RoutedEventArgs e)
    {
        string target = string.IsNullOrWhiteSpace(TxtRemoteMaintComputer.Text) ? "." : TxtRemoteMaintComputer.Text.Trim();
        var options = BuildRemoteMaintOptionsFromUi();

        BtnAnalyzeRemoteMaint.IsEnabled = false;
        BtnExecuteRemoteMaint.IsEnabled = false;
        RichRemoteMaintLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Inspecionando caches e espaço livre em '{target}'...";

        try
        {
            var res = await Task.Run(() => _remoteMaintenanceService.AnalyzeRemoteCleanup(
                target,
                options,
                msg => Dispatcher.Invoke(() => AppendColoredLog(RichRemoteMaintLogs, msg))));

            if (res.Success)
            {
                TxtRmFreeBefore.Text = $"{res.FreeSpaceBeforeGb:N2} GB livres (de {res.TotalSpaceGb:N2} GB)";
                TxtRmCleanableEstimate.Text = $"~{res.EstimatedCleanableGb:N2} GB em arquivos de cache";
                TxtRmFreeAfter.Text = "Aguardando execução...";
                TxtFooterStatus.Text = $"Análise de '{res.ComputerName}' concluída (~{res.EstimatedCleanableGb:N2} GB limpáveis).";
            }
            else
            {
                TxtFooterStatus.Text = $"Falha ao inspecionar '{target}'.";
            }
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRemoteMaintLogs, $"[ERRO] Exceção na inspeção: {ex.Message}");
            TxtFooterStatus.Text = "Erro durante análise de limpeza.";
        }
        finally
        {
            BtnAnalyzeRemoteMaint.IsEnabled = true;
            BtnExecuteRemoteMaint.IsEnabled = true;
        }
    }

    private async void BtnExecuteRemoteMaint_Click(object sender, RoutedEventArgs e)
    {
        string target = string.IsNullOrWhiteSpace(TxtRemoteMaintComputer.Text) ? "." : TxtRemoteMaintComputer.Text.Trim();
        var options = BuildRemoteMaintOptionsFromUi();

        var confirm = System.Windows.MessageBox.Show(
            $"Deseja executar as rotinas de limpeza e manutenção selecionadas na máquina '{target}'?",
            "Confirmar Limpeza & Manutenção",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        BtnAnalyzeRemoteMaint.IsEnabled = false;
        BtnExecuteRemoteMaint.IsEnabled = false;
        RichRemoteMaintLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Executando limpeza e manutenção em '{target}'...";

        try
        {
            var res = await Task.Run(() => _remoteMaintenanceService.ExecuteMaintenance(
                target,
                options,
                msg => Dispatcher.Invoke(() => AppendColoredLog(RichRemoteMaintLogs, msg))));

            if (res.Success)
            {
                TxtRmFreeBefore.Text = $"{res.FreeSpaceBeforeGb:N2} GB livres (de {res.TotalSpaceGb:N2} GB)";
                TxtRmFreeAfter.Text = $"{res.FreeSpaceAfterGb:N2} GB livres (+{res.FreedGb:N2} GB recuperados)";
                TxtFooterStatus.Text = $"Manutenção em '{res.ComputerName}' concluída! (+{res.FreedGb:N2} GB livres)";
            }
            else
            {
                TxtFooterStatus.Text = $"Manutenção em '{target}' encerrada com alertas.";
            }
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRemoteMaintLogs, $"[ERRO] Exceção na manutenção: {ex.Message}");
            TxtFooterStatus.Text = "Erro durante manutenção remota.";
        }
        finally
        {
            BtnAnalyzeRemoteMaint.IsEnabled = true;
            BtnExecuteRemoteMaint.IsEnabled = true;
        }
    }
    #endregion

    #region Painel 8: Limpeza de Perfis de Usuários Remotos
    private void TxtRemoteProfilesComputer_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnQueryRemoteProfiles_Click(this, new RoutedEventArgs());
        }
    }

    private async void BtnQueryRemoteProfiles_Click(object sender, RoutedEventArgs e)
    {
        string target = string.IsNullOrWhiteSpace(TxtRemoteProfilesComputer.Text) ? "." : TxtRemoteProfilesComputer.Text.Trim();

        BtnQueryRemoteProfiles.IsEnabled = false;
        BtnPurgeRemoteProfiles.IsEnabled = false;
        RichRemoteProfilesLogs.Document.Blocks.Clear();
        _remoteProfilesList.Clear();
        TxtFooterStatus.Text = $"Consultando perfis de usuários em '{target}'...";

        try
        {
            var profiles = await Task.Run(() => _remoteUserProfileService.QueryUserProfiles(
                target,
                msg => Dispatcher.Invoke(() => AppendColoredLog(RichRemoteProfilesLogs, msg)),
                calculateDiskSize: true));

            foreach (var p in profiles)
            {
                _remoteProfilesList.Add(p);
            }

            double totalGb = profiles.Sum(p => p.EspacoDiscoGb);
            int removable = profiles.Count(p => p.CanBeRemoved);
            TxtProfilesListHeader.Text = $"Perfis em '{RemoteManagementHelper.NormalizeTargetHost(target)}': {profiles.Count} encontrado(s) ({removable} inativos removíveis | Total: {totalGb:N2} GB)";
            TxtFooterStatus.Text = $"Consulta de perfis concluída ({profiles.Count} perfis encontrados).";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRemoteProfilesLogs, $"[ERRO] Exceção ao listar perfis: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao listar perfis de usuários.";
        }
        finally
        {
            BtnQueryRemoteProfiles.IsEnabled = true;
            BtnPurgeRemoteProfiles.IsEnabled = true;
        }
    }

    private void BtnSelectAllInactiveProfiles_Click(object sender, RoutedEventArgs e)
    {
        foreach (var p in _remoteProfilesList)
        {
            if (p.CanBeRemoved)
                p.IsSelected = true;
        }
    }

    private void BtnClearProfileSelection_Click(object sender, RoutedEventArgs e)
    {
        foreach (var p in _remoteProfilesList)
        {
            p.IsSelected = false;
        }
    }

    private async void BtnPurgeRemoteProfiles_Click(object sender, RoutedEventArgs e)
    {
        string target = string.IsNullOrWhiteSpace(TxtRemoteProfilesComputer.Text) ? "." : TxtRemoteProfilesComputer.Text.Trim();

        var selectedUsers = _remoteProfilesList
            .Where(p => p.IsSelected && p.CanBeRemoved)
            .Select(p => p.Usuario)
            .ToList();

        if (!string.IsNullOrWhiteSpace(TxtManualProfilesToPurge.Text))
        {
            var manual = TxtManualProfilesToPurge.Text
                .Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            selectedUsers.AddRange(manual);
        }

        selectedUsers = selectedUsers.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        if (selectedUsers.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "Marque pelo menos um perfil inativo na lista ou digite os nomes de usuários para remover.",
                "Nenhum Perfil Selecionado",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Deseja realmente EXCLUIR {selectedUsers.Count} perfil(is) na máquina '{target}'?\n\nUsuários: {string.Join(", ", selectedUsers)}\n\n(Perfis logados no momento serão protegidos e ignorados automaticamente.)",
            "Confirmar Exclusão de Perfis",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        BtnQueryRemoteProfiles.IsEnabled = false;
        BtnPurgeRemoteProfiles.IsEnabled = false;
        RichRemoteProfilesLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Removendo {selectedUsers.Count} perfil(is) em '{target}'...";

        try
        {
            var res = await Task.Run(() => _remoteUserProfileService.PurgeUserProfiles(
                target,
                selectedUsers,
                msg => Dispatcher.Invoke(() => AppendColoredLog(RichRemoteProfilesLogs, msg))));

            TxtManualProfilesToPurge.Clear();

            // Atualiza a lista após remoção
            var updated = await Task.Run(() => _remoteUserProfileService.QueryUserProfiles(target, onLog: null, calculateDiskSize: true));
            _remoteProfilesList.Clear();
            foreach (var p in updated)
                _remoteProfilesList.Add(p);

            TxtProfilesListHeader.Text = $"Perfis em '{res.ComputerName}': {updated.Count} restante(s) ({res.RemovedCount} removido(s), ~{res.FreedEstimateGb:N2} GB liberados)";
            TxtFooterStatus.Text = $"Remoção concluída: {res.RemovedCount} perfil(is) excluído(s) (~{res.FreedEstimateGb:N2} GB liberados).";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRemoteProfilesLogs, $"[ERRO] Exceção ao remover perfis: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao remover perfis de usuários.";
        }
        finally
        {
            BtnQueryRemoteProfiles.IsEnabled = true;
            BtnPurgeRemoteProfiles.IsEnabled = true;
        }
    }
    #endregion

    #region Painel 9: Consulta de Usuários Active Directory
    private void TxtAdUserQuery_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnConsultarUsuarioAd_Click(sender, e);
        }
    }

    private async void BtnConsultarUsuarioAd_Click(object sender, RoutedEventArgs e)
    {
        string query = TxtAdUserQuery.Text.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            AppendColoredLog(RichAdUsersLogs, "[Aviso] Informe login, nome completo, e-mail, setor ou matrícula para consultar no Active Directory.");
            return;
        }

        BtnConsultarUsuarioAd.IsEnabled = false;
        PlaceholderAdUserCard.Visibility = Visibility.Collapsed;
        BorderAdUserPagination.Visibility = Visibility.Collapsed;
        ItemsAdUserResults.ItemsSource = null;
        _allAdUsersFound.Clear();
        _adUserCurrentPage = 1;

        LoadingAdUserCard.Visibility = Visibility.Visible;
        RichAdUsersLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Consultando usuários no Active Directory para '{query}'...";

        try
        {
            var results = await _adSccmService.SearchUsersAsync(query, msg =>
            {
                AppendColoredLog(RichAdUsersLogs, msg);
            }, maxResults: 150);

            _allAdUsersFound.AddRange(results);

            if (_allAdUsersFound.Count > 0)
            {
                PopulateAdUserDepartments();
                ApplyAdUserFilters();
                PlaceholderAdUserCard.Visibility = Visibility.Collapsed;
                TxtFooterStatus.Text = $"{_allAdUsersFound.Count} usuário(s) encontrado(s) no Active Directory para '{query}'.";
            }
            else
            {
                PlaceholderAdUserCard.Visibility = Visibility.Visible;
                BorderAdUserPagination.Visibility = Visibility.Collapsed;
                AppendColoredLog(RichAdUsersLogs, $"[i] Nenhum usuário encontrado no Active Directory correspondente a '{query}'.");
                TxtFooterStatus.Text = $"Nenhum usuário AD encontrado para '{query}'.";
            }
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichAdUsersLogs, $"[ERRO] Falha na consulta de usuários AD: {ex.Message}");
            AppLoggerService.Danger($"Erro na consulta de usuários AD '{query}': {ex.Message}", nameof(MainWindow));
            PlaceholderAdUserCard.Visibility = Visibility.Visible;
            BorderAdUserPagination.Visibility = Visibility.Collapsed;
        }
        finally
        {
            LoadingAdUserCard.Visibility = Visibility.Collapsed;
            BtnConsultarUsuarioAd.IsEnabled = true;
        }
    }

    private void PopulateAdUserDepartments()
    {
        _isPopulatingDeptFilter = true;
        try
        {
            var depts = _allAdUsersFound
                .Select(u => !string.IsNullOrWhiteSpace(u.Department) ? u.Department.Trim() : (!string.IsNullOrWhiteSpace(u.Office) ? u.Office.Trim() : ""))
                .Where(d => !string.IsNullOrWhiteSpace(d))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(d => d)
                .ToList();

            var list = new List<string> { "Todos os Departamentos" };
            list.AddRange(depts);

            CmbAdUserDeptFilter.ItemsSource = list;
            CmbAdUserDeptFilter.SelectedIndex = 0;
        }
        finally
        {
            _isPopulatingDeptFilter = false;
        }
    }

    private void TxtAdUserLocalFilter_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (!IsLoaded) return;
        ApplyAdUserFilters();
    }

    private void CmbAdUserFilter_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (!IsLoaded || _isPopulatingDeptFilter) return;
        ApplyAdUserFilters();
    }

    private void BtnClearAdUserFilters_Click(object sender, RoutedEventArgs e)
    {
        _isPopulatingDeptFilter = true;
        try
        {
            TxtAdUserLocalFilter.Clear();
            if (CmbAdUserDeptFilter.Items.Count > 0)
                CmbAdUserDeptFilter.SelectedIndex = 0;
            CmbAdUserStatusFilter.SelectedIndex = 0;
        }
        finally
        {
            _isPopulatingDeptFilter = false;
        }

        ApplyAdUserFilters();
    }

    private void ApplyAdUserFilters()
    {
        if (ItemsAdUserResults == null || BorderAdUserPagination == null)
            return;

        if (_allAdUsersFound.Count == 0)
        {
            _filteredAdUsers.Clear();
            RenderAdUsersPage();
            return;
        }

        string textFilter = TxtAdUserLocalFilter?.Text.Trim().ToLowerInvariant() ?? "";
        string selectedDept = CmbAdUserDeptFilter?.SelectedItem as string ?? "";
        int statusIndex = CmbAdUserStatusFilter?.SelectedIndex ?? 0;

        var query = _allAdUsersFound.AsEnumerable();

        // 1. Filtro textual
        if (!string.IsNullOrEmpty(textFilter))
        {
            query = query.Where(u =>
                (u.SamAccountName != null && u.SamAccountName.ToLowerInvariant().Contains(textFilter)) ||
                (u.DisplayName != null && u.DisplayName.ToLowerInvariant().Contains(textFilter)) ||
                (u.CommonName != null && u.CommonName.ToLowerInvariant().Contains(textFilter)) ||
                (u.Mail != null && u.Mail.ToLowerInvariant().Contains(textFilter)) ||
                (u.Title != null && u.Title.ToLowerInvariant().Contains(textFilter)) ||
                (u.Department != null && u.Department.ToLowerInvariant().Contains(textFilter)) ||
                (u.Office != null && u.Office.ToLowerInvariant().Contains(textFilter)) ||
                (u.EmployeeId != null && u.EmployeeId.ToLowerInvariant().Contains(textFilter)) ||
                (u.Description != null && u.Description.ToLowerInvariant().Contains(textFilter))
            );
        }

        // 2. Filtro por Departamento
        if (!string.IsNullOrEmpty(selectedDept) && !selectedDept.StartsWith("Todos", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(u =>
                string.Equals(u.Department?.Trim(), selectedDept, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(u.Office?.Trim(), selectedDept, StringComparison.OrdinalIgnoreCase));
        }

        // 3. Filtro por Status da Conta
        // 0: Todos, 1: Ativas, 2: Bloqueadas, 3: Desativadas, 4: Troca de Senha
        if (statusIndex == 1)
        {
            query = query.Where(u => u.IsEnabled && !u.IsLockedOut);
        }
        else if (statusIndex == 2)
        {
            query = query.Where(u => u.IsLockedOut);
        }
        else if (statusIndex == 3)
        {
            query = query.Where(u => !u.IsEnabled);
        }
        else if (statusIndex == 4)
        {
            query = query.Where(u => u.MustChangePasswordNextLogon);
        }

        _filteredAdUsers.Clear();
        _filteredAdUsers.AddRange(query);

        _adUserCurrentPage = 1;
        RenderAdUsersPage();
    }

    private void RenderAdUsersPage()
    {
        if (ItemsAdUserResults == null || BorderAdUserPagination == null ||
            TxtAdUserPaginationSummary == null || TxtAdUserCurrentPageNumber == null ||
            BtnAdUserPrevPage == null || BtnAdUserNextPage == null)
            return;

        if (_filteredAdUsers.Count == 0)
        {
            ItemsAdUserResults.ItemsSource = null;
            TxtAdUserPaginationSummary.Text = _allAdUsersFound.Count > 0
                ? $"Nenhum usuário corresponde aos filtros aplicados (de {_allAdUsersFound.Count} carregados)"
                : "Nenhum usuário localizado";
            TxtAdUserCurrentPageNumber.Text = "Página 0 / 0";
            BtnAdUserPrevPage.IsEnabled = false;
            BtnAdUserNextPage.IsEnabled = false;
            BorderAdUserPagination.Visibility = _allAdUsersFound.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            return;
        }

        int totalUsers = _filteredAdUsers.Count;
        int totalPages = (int)Math.Ceiling((double)totalUsers / AdUsersPageSize);
        if (_adUserCurrentPage > totalPages) _adUserCurrentPage = totalPages;
        if (_adUserCurrentPage < 1) _adUserCurrentPage = 1;

        int startIndex = (_adUserCurrentPage - 1) * AdUsersPageSize;
        var pageItems = _filteredAdUsers.Skip(startIndex).Take(AdUsersPageSize).ToList();

        ItemsAdUserResults.ItemsSource = pageItems;

        int endIndex = Math.Min(startIndex + pageItems.Count, totalUsers);
        string filterNote = _filteredAdUsers.Count < _allAdUsersFound.Count
            ? $" (filtrados de {_allAdUsersFound.Count})"
            : "";

        TxtAdUserPaginationSummary.Text = $"Exibindo {startIndex + 1}–{endIndex} de {totalUsers} usuário(s){filterNote}";
        TxtAdUserCurrentPageNumber.Text = $"Página {_adUserCurrentPage} / {totalPages}";

        BtnAdUserPrevPage.IsEnabled = _adUserCurrentPage > 1;
        BtnAdUserNextPage.IsEnabled = _adUserCurrentPage < totalPages;

        BorderAdUserPagination.Visibility = Visibility.Visible;
    }

    private void BtnAdUserPrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_adUserCurrentPage > 1)
        {
            _adUserCurrentPage--;
            RenderAdUsersPage();
        }
    }

    private void BtnAdUserNextPage_Click(object sender, RoutedEventArgs e)
    {
        int totalPages = (int)Math.Ceiling((double)_filteredAdUsers.Count / AdUsersPageSize);
        if (_adUserCurrentPage < totalPages)
        {
            _adUserCurrentPage++;
            RenderAdUsersPage();
        }
    }

    private void BtnAdUserCopyLogin_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is AdUserDetails user && !string.IsNullOrWhiteSpace(user.SamAccountName))
        {
            if (ClipboardHelper.Copy(user.SamAccountName))
            {
                TxtFooterStatus.Text = $"Login '{user.SamAccountName}' copiado para a Área de Transferência!";
            }
        }
    }

    private void BtnAdUserCopyMail_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is AdUserDetails user)
        {
            string mail = !string.IsNullOrWhiteSpace(user.Mail) ? user.Mail : user.UserPrincipalName;
            if (!string.IsNullOrWhiteSpace(mail) && ClipboardHelper.Copy(mail))
            {
                TxtFooterStatus.Text = $"E-mail '{mail}' copiado para a Área de Transferência!";
            }
            else
            {
                TxtFooterStatus.Text = "Nenhum endereço de e-mail cadastrado para este usuário.";
            }
        }
    }

    private void BtnAdUserCopyDn_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is AdUserDetails user && !string.IsNullOrWhiteSpace(user.DistinguishedName))
        {
            if (ClipboardHelper.Copy(user.DistinguishedName))
            {
                TxtFooterStatus.Text = $"Distinguished Name (DN) de '{user.SamAccountName}' copiado!";
            }
        }
    }

    private void BtnAdUserCopyGroups_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is AdUserDetails user)
        {
            string groupsText = user.Groups.Count > 0
                ? string.Join(Environment.NewLine, user.Groups)
                : "Nenhum grupo listado.";

            if (ClipboardHelper.Copy(groupsText))
            {
                TxtFooterStatus.Text = $"Lista de {user.Groups.Count} grupo(s) de '{user.SamAccountName}' copiada!";
            }
        }
    }

    private void BtnAdUserCopyFullDossier_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is AdUserDetails user)
        {
            if (ClipboardHelper.Copy(user.FormatFullDossier()))
            {
                TxtFooterStatus.Text = $"Ficha completa do usuário '{user.SamAccountName}' copiada para a Área de Transferência!";
            }
        }
    }

    private void BtnAdUserFindMachines_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is AdUserDetails user && !string.IsNullOrWhiteSpace(user.SamAccountName))
        {
            NavListBox.SelectedIndex = 2;
            TxtComputerName.Text = user.SamAccountName;
            TxtFooterStatus.Text = $"Localizando máquinas conectadas para o usuário '{user.SamAccountName}'...";
            BtnConsultarRemoto_Click(this, new RoutedEventArgs());
        }
    }
    #endregion

    #region Painel 11: Bancada Streamlit (Docker WSL)
    private async Task CheckBancadaStatusAsync()
    {
        TxtBancadaDetails.Text = $"URL: {ConfigService.UrlBancadaLocal} | WSL: {ConfigService.BancadaWslDistro} ({ConfigService.BancadaWslDir})";
        TxtBancadaStatus.Text = "⚪ Verificando...";
        BadgeBancadaStatus.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1e293b"));

        bool online = await _bancadaService.IsBancadaOnlineAsync();
        if (online)
        {
            TxtBancadaStatus.Text = "🟢 ONLINE (Streamlit 8501)";
            TxtBancadaStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#34d399"));
            BadgeBancadaStatus.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#064e3b"));
            TxtFooterStatus.Text = "Sistema Bancada STI está online em http://localhost:8501/";
        }
        else
        {
            TxtBancadaStatus.Text = "🔴 OFFLINE / PARADO";
            TxtBancadaStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f87171"));
            BadgeBancadaStatus.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7f1d1d"));
            TxtFooterStatus.Text = "Sistema Bancada STI está parado.";
        }
    }

    private async void BtnStartBancada_Click(object sender, RoutedEventArgs e)
    {
        BtnStartBancada.IsEnabled = false;
        BtnStopBancada.IsEnabled = false;
        BtnRestartBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Iniciando containers Docker da Bancada via WSL...";
        AppendColoredLog(RichBancadaLogs, "=== INICIALIZANDO SISTEMA BANCADA STI (DOCKER WSL) ===");
        AppendColoredLog(RichBancadaLogs, $"WSL Distro: {ConfigService.BancadaWslDistro}");
        AppendColoredLog(RichBancadaLogs, $"Diretório: {ConfigService.BancadaWslDir}");
        AppendColoredLog(RichBancadaLogs, $"Comando: docker compose up -d");

        _ = Task.Run(async () =>
        {
            await _bancadaService.StartDockerBancadaAsync(msg =>
            {
                Dispatcher.Invoke(() => AppendColoredLog(RichBancadaLogs, msg));
            });
        });

        // Espera 3 segundos para reativar botões e checar status
        await Task.Delay(3500);
        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        await CheckBancadaStatusAsync();
    }

    private async void BtnStopBancada_Click(object sender, RoutedEventArgs e)
    {
        BtnStartBancada.IsEnabled = false;
        BtnStopBancada.IsEnabled = false;
        BtnRestartBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Parando containers Docker da Bancada...";
        AppendColoredLog(RichBancadaLogs, "=== PARANDO CONTAINERS (DOCKER COMPOSE DOWN) ===");

        await _bancadaService.StopDockerBancadaAsync(msg =>
        {
            Dispatcher.Invoke(() => AppendColoredLog(RichBancadaLogs, msg));
        });

        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        await CheckBancadaStatusAsync();
        TxtFooterStatus.Text = "Containers da Bancada encerrados.";
    }

    private async void BtnRestartBancada_Click(object sender, RoutedEventArgs e)
    {
        BtnStartBancada.IsEnabled = false;
        BtnStopBancada.IsEnabled = false;
        BtnRestartBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Reiniciando stack Docker da Bancada...";
        AppendColoredLog(RichBancadaLogs, "=== REINICIANDO CONTAINERS DOCKER ===");

        await _bancadaService.RestartDockerBancadaAsync(msg =>
        {
            Dispatcher.Invoke(() => AppendColoredLog(RichBancadaLogs, msg));
        });

        await Task.Delay(3000);
        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        await CheckBancadaStatusAsync();
    }

    private void BtnOpenBancadaBrowser_Click(object sender, RoutedEventArgs e)
    {
        string url = ConfigService.UrlBancadaLocal;
        if (string.IsNullOrWhiteSpace(url)) url = "http://localhost:8501/";
        _browserLauncherService.OpenUrlsInFirefox([url]);
        TxtFooterStatus.Text = $"Abrindo Bancada no navegador: {url}";
    }

    private async void BtnCheckBancadaHealth_Click(object sender, RoutedEventArgs e)
    {
        await CheckBancadaStatusAsync();
    }

    private void BtnClearBancadaLogs_Click(object sender, RoutedEventArgs e)
    {
        RichBancadaLogs.Document.Blocks.Clear();
    }

    private void BtnOpenBancadaFolder_Click(object sender, RoutedEventArgs e)
    {
        string path = ConfigService.BancadaWindowsDir;
        if (Directory.Exists(path))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true
            });
            TxtFooterStatus.Text = $"Pasta da Bancada aberta no Windows Explorer: {path}";
        }
        else
        {
            System.Windows.MessageBox.Show($"O caminho da pasta da Bancada não foi encontrado:\n{path}", "Pasta Não Localizada", System.Windows.MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
    #endregion

    #region Central de Logs da Aplicação
    private void OnNewLogLineReceived(string logLine)
    {
        lock (_allAppLogsInMemory)
        {
            _allAppLogsInMemory.Add(logLine);
            if (_allAppLogsInMemory.Count > 3000)
            {
                _allAppLogsInMemory.RemoveAt(0);
            }
        }

        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
        {
            if (PanelAppLogs != null && PanelAppLogs.Visibility == Visibility.Visible)
            {
                if (PassesAppLogFilter(logLine))
                {
                    AddAppLogParagraph(logLine);
                    UpdateAppLogsCountBadge();
                    if (ChkAutoScrollAppLogs?.IsChecked == true)
                    {
                        RichAppLogs.ScrollToEnd();
                    }
                }
            }
        }));
    }

    private void LoadAppLogsToUi()
    {
        try
        {
            var logs = AppLoggerService.ReadRecentLogs(2000);
            lock (_allAppLogsInMemory)
            {
                _allAppLogsInMemory.Clear();
                _allAppLogsInMemory.AddRange(logs);
            }
            RenderFilteredAppLogs();
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Erro ao carregar logs da aplicação: {ex.Message}", nameof(MainWindow));
        }
    }

    private void RenderFilteredAppLogs()
    {
        if (RichAppLogs == null) return;

        RichAppLogs.Document.Blocks.Clear();

        List<string> snapshot;
        lock (_allAppLogsInMemory)
        {
            snapshot = new List<string>(_allAppLogsInMemory);
        }

        int count = 0;
        foreach (var line in snapshot)
        {
            if (PassesAppLogFilter(line))
            {
                AddAppLogParagraph(line);
                count++;
            }
        }

        UpdateAppLogsCountBadge(count);

        if (ChkAutoScrollAppLogs?.IsChecked == true)
        {
            RichAppLogs.ScrollToEnd();
        }
    }

    private bool PassesAppLogFilter(string line)
    {
        if (string.IsNullOrWhiteSpace(line)) return false;

        // 1. Filtro por Nível
        if (CmbAppLogLevelFilter != null && CmbAppLogLevelFilter.SelectedIndex > 0)
        {
            string requiredLevel = CmbAppLogLevelFilter.SelectedIndex switch
            {
                1 => "[INFO]",
                2 => "[WARNING]",
                3 => "[DANGER]",
                4 => "[SUCCESS]",
                _ => ""
            };

            if (!string.IsNullOrEmpty(requiredLevel) && !line.Contains(requiredLevel, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        // 2. Filtro por Busca de Texto
        string search = TxtAppLogSearch?.Text?.Trim() ?? string.Empty;
        if (!string.IsNullOrEmpty(search) && !line.Contains(search, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private void AddAppLogParagraph(string line)
    {
        var para = new Paragraph
        {
            Margin = new Thickness(0, 1, 0, 1),
            LineHeight = 18,
            FontFamily = RichAppLogs.FontFamily
        };

        int idxFirstClose = line.IndexOf(']');
        if (line.StartsWith("[") && idxFirstClose > 0)
        {
            // Timestamp
            string timestamp = line[..(idxFirstClose + 1)];
            para.Inlines.Add(new Run(timestamp)
            {
                Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)) // Slate 400
            });

            string rest = line[(idxFirstClose + 1)..].TrimStart();
            int idxSecondClose = rest.IndexOf(']');
            if (rest.StartsWith("[") && idxSecondClose > 0)
            {
                string levelTag = rest[..(idxSecondClose + 1)];
                var levelBrush = levelTag.ToUpperInvariant() switch
                {
                    "[INFO]" => new SolidColorBrush(Color.FromRgb(56, 189, 248)),       // Sky 400
                    "[WARNING]" => new SolidColorBrush(Color.FromRgb(251, 191, 36)),    // Amber 400
                    "[DANGER]" => new SolidColorBrush(Color.FromRgb(248, 113, 113)),    // Red 400
                    "[SUCCESS]" => new SolidColorBrush(Color.FromRgb(52, 211, 153)),    // Emerald 400
                    _ => new SolidColorBrush(Color.FromRgb(203, 213, 225))              // Slate 300
                };

                para.Inlines.Add(new Run(" " + levelTag)
                {
                    Foreground = levelBrush,
                    FontWeight = FontWeights.Bold
                });

                string afterLevel = rest[(idxSecondClose + 1)..].TrimStart();
                int idxThirdClose = afterLevel.IndexOf(']');
                if (afterLevel.StartsWith("[") && idxThirdClose > 0)
                {
                    // Caller
                    string callerTag = afterLevel[..(idxThirdClose + 1)];
                    para.Inlines.Add(new Run(" " + callerTag)
                    {
                        Foreground = new SolidColorBrush(Color.FromRgb(167, 139, 250)), // Purple 400
                        FontWeight = FontWeights.Medium
                    });

                    string msg = afterLevel[(idxThirdClose + 1)..];
                    para.Inlines.Add(new Run(" " + msg.TrimStart())
                    {
                        Foreground = GetMessageBrush(levelTag)
                    });
                }
                else
                {
                    para.Inlines.Add(new Run(" " + afterLevel)
                    {
                        Foreground = GetMessageBrush(levelTag)
                    });
                }
            }
            else
            {
                para.Inlines.Add(new Run(" " + rest)
                {
                    Foreground = new SolidColorBrush(Color.FromRgb(226, 232, 240))
                });
            }
        }
        else
        {
            para.Inlines.Add(new Run(line)
            {
                Foreground = new SolidColorBrush(Color.FromRgb(226, 232, 240))
            });
        }

        RichAppLogs.Document.Blocks.Add(para);
    }

    private static SolidColorBrush GetMessageBrush(string levelTag)
    {
        return levelTag.ToUpperInvariant() switch
        {
            "[DANGER]" => new SolidColorBrush(Color.FromRgb(254, 202, 202)),   // Red 200
            "[WARNING]" => new SolidColorBrush(Color.FromRgb(254, 240, 138)),  // Yellow 200
            "[SUCCESS]" => new SolidColorBrush(Color.FromRgb(167, 243, 208)),  // Emerald 200
            _ => new SolidColorBrush(Color.FromRgb(241, 245, 249))             // Slate 100
        };
    }

    private void UpdateAppLogsCountBadge(int? specificCount = null)
    {
        if (TxtAppLogsCount != null && RichAppLogs != null)
        {
            int count = specificCount ?? RichAppLogs.Document.Blocks.Count;
            TxtAppLogsCount.Text = $"{count} registro(s)";
        }
    }

    private void BtnReloadAppLogs_Click(object sender, RoutedEventArgs e)
    {
        LoadAppLogsToUi();
        TxtFooterStatus.Text = "Logs da aplicação recarregados do disco.";
    }

    private void BtnClearAppLogsView_Click(object sender, RoutedEventArgs e)
    {
        RichAppLogs?.Document.Blocks.Clear();
        UpdateAppLogsCountBadge(0);
        TxtFooterStatus.Text = "Visualização de logs limpa na tela.";
    }

    private void BtnOpenAppLogsFolder_Click(object sender, RoutedEventArgs e)
    {
        string dir = AppLoggerService.LogDirectoryPath;
        if (Directory.Exists(dir))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = dir,
                UseShellExecute = true
            });
            TxtFooterStatus.Text = $"Pasta de logs aberta: {dir}";
        }
        else
        {
            System.Windows.MessageBox.Show($"A pasta de logs não foi encontrada em:\n{dir}", "Logs", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void CmbAppLogLevelFilter_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        RenderFilteredAppLogs();
    }

    private void TxtAppLogSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        RenderFilteredAppLogs();
    }
    #endregion

    #region Painel 11: Otimizações de Registro (Tweaks & Rollback)
    private void InitializeRegistryTweaks()
    {
        _allRegistryTweaks.Clear();
        var defaults = _registryOptimizerService.GetDefaultTweaks();
        foreach (var t in defaults)
        {
            _allRegistryTweaks.Add(t);
        }
        FilterRegistryTweaksByCategory();
    }

    private void RefreshRegistryTweaksStatus()
    {
        if (!_registryTweaksScannedOnce)
        {
            _registryTweaksScannedOnce = true;
            _ = ScanRegistryTweaksAsync();
        }
    }

    private async Task ScanRegistryTweaksAsync()
    {
        BtnScanRegistryTweaks.IsEnabled = false;
        BtnApplySelectedRegistryTweaks.IsEnabled = false;
        BtnRollbackSelectedRegistryTweaks.IsEnabled = false;
        RichRegistryLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = "Escaneando chaves do Registro do Windows...";

        try
        {
            await Task.Run(() =>
            {
                _registryOptimizerService.ScanAll(_allRegistryTweaks, msg =>
                {
                    Dispatcher.Invoke(() => AppendColoredLog(RichRegistryLogs, msg));
                });
            });

            UpdateRegistryOptBadge();
            TxtFooterStatus.Text = "Escaneamento de registro concluído.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRegistryLogs, $"[ERRO] Falha no escaneamento: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao escanear o registro.";
        }
        finally
        {
            BtnScanRegistryTweaks.IsEnabled = true;
            BtnApplySelectedRegistryTweaks.IsEnabled = true;
            BtnRollbackSelectedRegistryTweaks.IsEnabled = true;
        }
    }

    private void UpdateRegistryOptBadge()
    {
        int total = _allRegistryTweaks.Count;
        int opt = _allRegistryTweaks.Count(t => t.IsOptimized);
        if (TxtRegistryOptBadge != null)
        {
            TxtRegistryOptBadge.Text = $"{opt} / {total} Otimizados";
        }
    }

    private void FilterRegistryTweaksByCategory()
    {
        _filteredRegistryTweaks.Clear();
        string category = (CmbRegistryCategoryFilter?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Todas as Categorias";

        foreach (var t in _allRegistryTweaks)
        {
            if (category == "Todas as Categorias" || t.Category.Equals(category, StringComparison.OrdinalIgnoreCase))
            {
                _filteredRegistryTweaks.Add(t);
            }
        }
    }

    private void CmbRegistryCategoryFilter_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        FilterRegistryTweaksByCategory();
    }

    private void BtnSelectAllRegistryTweaks_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _filteredRegistryTweaks) t.IsSelected = true;
    }

    private void BtnDeselectAllRegistryTweaks_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _filteredRegistryTweaks) t.IsSelected = false;
    }

    private async void BtnScanRegistryTweaks_Click(object sender, RoutedEventArgs e)
    {
        await ScanRegistryTweaksAsync();
    }

    private async void BtnApplySelectedRegistryTweaks_Click(object sender, RoutedEventArgs e)
    {
        var selected = _allRegistryTweaks.Where(t => t.IsSelected).ToList();
        if (selected.Count == 0)
        {
            System.Windows.MessageBox.Show("Marque pelo menos um tweak de registro para aplicar.", "Nenhum Tweak Selecionado", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Deseja aplicar {selected.Count} tweak(s) no Registro do Windows?\n\nUm ponto de restauração (.reg e snapshot) será gerado automaticamente antes de qualquer alteração.",
            "Confirmar Aplicação de Tweaks",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);

        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        BtnScanRegistryTweaks.IsEnabled = false;
        BtnApplySelectedRegistryTweaks.IsEnabled = false;
        BtnRollbackSelectedRegistryTweaks.IsEnabled = false;
        TxtFooterStatus.Text = $"Aplicando {selected.Count} otimização(ões) no registro...";

        try
        {
            var (applied, failed) = await Task.Run(() => _registryOptimizerService.ApplyBatch(selected, msg =>
            {
                Dispatcher.Invoke(() => AppendColoredLog(RichRegistryLogs, msg));
            }));

            UpdateRegistryOptBadge();
            TxtFooterStatus.Text = $"Tweaks aplicados: {applied} sucesso(s), {failed} falha(s).";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRegistryLogs, $"[ERRO] Exceção durante aplicação: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao aplicar tweaks de registro.";
        }
        finally
        {
            BtnScanRegistryTweaks.IsEnabled = true;
            BtnApplySelectedRegistryTweaks.IsEnabled = true;
            BtnRollbackSelectedRegistryTweaks.IsEnabled = true;
        }
    }

    private async void BtnRollbackSelectedRegistryTweaks_Click(object sender, RoutedEventArgs e)
    {
        var selected = _allRegistryTweaks.Where(t => t.IsSelected).ToList();
        if (selected.Count == 0)
        {
            System.Windows.MessageBox.Show("Marque pelo menos um tweak para restauração (rollback).", "Nenhum Tweak Selecionado", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Deseja restaurar {selected.Count} tweak(s) aos valores padrão/anteriores do Windows?",
            "Confirmar Rollback",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        BtnScanRegistryTweaks.IsEnabled = false;
        BtnApplySelectedRegistryTweaks.IsEnabled = false;
        BtnRollbackSelectedRegistryTweaks.IsEnabled = false;
        TxtFooterStatus.Text = $"Restaurando {selected.Count} tweak(s)...";

        try
        {
            var (restored, failed) = await Task.Run(() => _registryOptimizerService.RollbackBatch(selected, msg =>
            {
                Dispatcher.Invoke(() => AppendColoredLog(RichRegistryLogs, msg));
            }));

            UpdateRegistryOptBadge();
            TxtFooterStatus.Text = $"Rollback finalizado: {restored} restaurado(s), {failed} falha(s).";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRegistryLogs, $"[ERRO] Exceção durante rollback: {ex.Message}");
            TxtFooterStatus.Text = "Erro no rollback de registro.";
        }
        finally
        {
            BtnScanRegistryTweaks.IsEnabled = true;
            BtnApplySelectedRegistryTweaks.IsEnabled = true;
            BtnRollbackSelectedRegistryTweaks.IsEnabled = true;
        }
    }

    private void BtnApplySingleTweak_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button btn && btn.Tag is RegistryTweakItem tweak)
        {
            try
            {
                var (success, _) = _registryOptimizerService.ApplyTweak(tweak, msg =>
                {
                    Dispatcher.Invoke(() => AppendColoredLog(RichRegistryLogs, msg));
                });
                UpdateRegistryOptBadge();
                TxtFooterStatus.Text = success ? $"Tweak '{tweak.Title}' aplicado." : $"Falha no tweak '{tweak.Title}'.";
            }
            catch (Exception ex)
            {
                AppendColoredLog(RichRegistryLogs, $"[ERRO] {ex.Message}");
            }
        }
    }

    private void BtnRollbackSingleTweak_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button btn && btn.Tag is RegistryTweakItem tweak)
        {
            try
            {
                var (success, _) = _registryOptimizerService.RollbackTweak(tweak, msg =>
                {
                    Dispatcher.Invoke(() => AppendColoredLog(RichRegistryLogs, msg));
                });
                UpdateRegistryOptBadge();
                TxtFooterStatus.Text = success ? $"Tweak '{tweak.Title}' restaurado." : $"Falha no rollback de '{tweak.Title}'.";
            }
            catch (Exception ex)
            {
                AppendColoredLog(RichRegistryLogs, $"[ERRO] {ex.Message}");
            }
        }
    }

    private void BtnOpenRegistryBackupFolder_Click(object sender, RoutedEventArgs e)
    {
        string dir = RegistryOptimizerService.BackupFolderPath;
        if (Directory.Exists(dir))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = dir,
                UseShellExecute = true
            });
            TxtFooterStatus.Text = $"Pasta de backups aberta: {dir}";
        }
        else
        {
            System.Windows.MessageBox.Show($"Nenhum backup gerado ainda em:\n{dir}", "Backups de Registro", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
    #endregion
}
