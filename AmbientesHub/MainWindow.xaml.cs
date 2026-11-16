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
    private readonly UsmtBackupService _usmtService = new();
    private readonly ObservableCollection<RemoteUserProfileItem> _usmtUsersList = [];
    private readonly BancadaChamadosService _chamadosService = new();
    private readonly ObservableCollection<ChamadoItem> _chamadosList = [];
    private readonly List<ChamadoItem> _allChamadosFound = [];
    private readonly HashSet<string> _selectedLocalidades = new(StringComparer.OrdinalIgnoreCase);
    private bool _isUpdatingChamadosFilters = false;
    private int _chamadosCurrentPage = 1;
    private const int ChamadosPageSize = 10;
    private readonly BancadaOxeService _oxeService = new();
    private readonly ObservableCollection<OxeRamalItem> _oxeList = [];
    private readonly List<OxeRamalItem> _allOxeRaw = [];
    private readonly List<OxeRamalItem> _allOxeFound = [];
    private bool _isUpdatingOxeFilters = false;
    private int _oxeCurrentPage = 1;
    private const int OxePageSize = 15;
    private readonly BancadaImpressorasService _impressorasService = new();
    private readonly ObservableCollection<ImpressoraItem> _impressorasList = [];
    private readonly List<ImpressoraItem> _allImpressorasRaw = [];
    private readonly List<ImpressoraItem> _allImpressorasFound = [];
    private bool _isUpdatingImpressorasFilters = false;
    private int _impressorasCurrentPage = 1;
    private const int ImpressorasPageSize = 15;

    public MainWindow()
    {
        InitializeComponent();
        AppLoggerService.OnLogLineWritten += OnNewLogLineReceived;
        AppLoggerService.Info("Ambientes Hub iniciado com sucesso.", nameof(MainWindow));
        LstSoftwareCatalog.ItemsSource = _softwareCatalog;
        ItemsRemoteProfiles.ItemsSource = _remoteProfilesList;
        LoadMachineInfoIntoSettings();
        ItemsUsmtUsers.ItemsSource = _usmtUsersList;
        ItemsChamados.ItemsSource = _chamadosList;
        ItemsOxe.ItemsSource = _oxeList;
        ItemsImpressoras.ItemsSource = _impressorasList;
        LstRegistryTweaks.ItemsSource = _filteredRegistryTweaks;
        InitializeRegistryTweaks();
        PopulateFontList();
        LoadSettingsIntoUi();
        InitializeSystemTray();
        RefreshPerformanceAdminStatus();
        RefreshUsmtBinStatusBadge();
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
            NavLbl14.Visibility = Visibility.Visible;
            NavLbl15.Visibility = Visibility.Visible;
            NavLbl16.Visibility = Visibility.Visible;
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
            NavLbl14.Visibility = Visibility.Collapsed;
            NavLbl15.Visibility = Visibility.Collapsed;
            NavLbl16.Visibility = Visibility.Collapsed;
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

        if (lower is "usmt" or "backup" or "restore" or "backup usuarios" or "backup usuários" or "migracao" or "migração")
        {
            NavListBox.SelectedIndex = 10;
            TxtFooterStatus.Text = "Navegado para: Backup & Restore de Perfis (USMT)";
            return;
        }

        if (lower is "bancada" or "docker" or "wsl" or "streamlit" or "8501" or "chamados bancada")
        {
            NavListBox.SelectedIndex = 11;
            TxtFooterStatus.Text = "Navegado para: Bancada Streamlit (Docker WSL)";
            return;
        }

        if (lower is "chamados" or "otrs" or "citsmart" or "tickets" or "fila")
        {
            NavListBox.SelectedIndex = 12;
            TxtFooterStatus.Text = "Navegado para: Chamados (OTRS & CitSmart)";
            return;
        }

        if (lower is "oxe" or "telefone" or "telefones" or "ramal" or "ramais" or "central" or "telefonia")
        {
            NavListBox.SelectedIndex = 13;
            TxtFooterStatus.Text = "Navegado para: Telefonia OXE (Ramais)";
            return;
        }

        if (lower is "impressora" or "impressoras" or "papercut" or "print" or "multifuncional")
        {
            NavListBox.SelectedIndex = 17;
            TxtFooterStatus.Text = "Navegado para: Impressoras PaperCut";
            return;
        }

        if (lower is "regedit" or "registro" or "tweak" or "tweaks" or "otimizar registro" or "otimização registro" or "rollback" or "ajustes registro")
        {
            NavListBox.SelectedIndex = 14;
            TxtFooterStatus.Text = "Navegado para: Otimizações de Registro (Tweaks & Rollback)";
            return;
        }

        if (lower is "logs" or "log" or "eventos" or "historico" or "histórico" or "console" or "app.log")
        {
            NavListBox.SelectedIndex = 15;
            TxtFooterStatus.Text = "Navegado para: Logs da Aplicação";
            return;
        }

        if (lower is "config" or "configuração" or "configurações" or "cofre" or "credenciais" or "ajustes" or "fonte")
        {
            NavListBox.SelectedIndex = 16;
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

    private void NavGroupHeader_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not string groupName) return;
        if (FindName(groupName) is System.Windows.Controls.StackPanel panel)
        {
            bool willOpen = panel.Visibility != Visibility.Visible;
            panel.Visibility = willOpen ? Visibility.Visible : Visibility.Collapsed;
            string chevName = groupName.Replace("Grp", "GrpChev");
            if (FindName(chevName) is System.Windows.Controls.TextBlock chev)
                chev.Text = willOpen ? "⌃" : "⌄";
        }
    }

    private void NavMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.FrameworkElement fe && fe.Tag is string tagStr && int.TryParse(tagStr, out int idx))
        {
            NavListBox.SelectedIndex = idx;
        }
    }

    private void HighlightActiveNavMenuItem(int activeIndex)
    {
        for (int i = 0; i <= 17; i++)
        {
            if (FindName($"NavBtn{i}") is Wpf.Ui.Controls.Button btn)
                btn.Appearance = (i == activeIndex) ? ControlAppearance.Secondary : ControlAppearance.Transparent;
        }
    }

    private void NavListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (PanelAmbientes == null) return;
        HighlightActiveNavMenuItem(NavListBox.SelectedIndex);

        PanelAmbientes.Visibility = Visibility.Collapsed;
        PanelRemoto.Visibility = Visibility.Collapsed;
        PanelAdSccm.Visibility = Visibility.Collapsed;
        PanelSaneamento.Visibility = Visibility.Collapsed;
        PanelPerformance.Visibility = Visibility.Collapsed;
        PanelApps.Visibility = Visibility.Collapsed;
        PanelAnalisador.Visibility = Visibility.Collapsed;
        PanelLimpezaRemota.Visibility = Visibility.Collapsed;
        PanelPerfisRemotos.Visibility = Visibility.Collapsed;
        PanelUsmtBackup.Visibility = Visibility.Collapsed;
        PanelUsuariosAd.Visibility = Visibility.Collapsed;
        PanelBancada.Visibility = Visibility.Collapsed;
        PanelChamados.Visibility = Visibility.Collapsed;
        PanelOxe.Visibility = Visibility.Collapsed;
        PanelImpressoras.Visibility = Visibility.Collapsed;
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
                PanelUsmtBackup.Visibility = Visibility.Visible;
                RefreshUsmtBinStatusBadge();
                break;
            case 11:
                PanelBancada.Visibility = Visibility.Visible;
                _ = CheckBancadaStatusAsync();
                break;
            case 12:
                PanelChamados.Visibility = Visibility.Visible;
                _ = LoadChamadosAsync();
                break;
            case 13:
                PanelOxe.Visibility = Visibility.Visible;
                _ = LoadOxeAsync();
                break;
            case 14:
                PanelRegistryTweaks.Visibility = Visibility.Visible;
                RefreshRegistryTweaksStatus();
                break;
            case 15:
                PanelAppLogs.Visibility = Visibility.Visible;
                LoadAppLogsToUi();
                break;
            case 16:
                PanelConfig.Visibility = Visibility.Visible;
                break;
            case 17:
                PanelImpressoras.Visibility = Visibility.Visible;
                _ = LoadImpressorasAsync();
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
            RichUsmtLogs.FontFamily = font;

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
            foreach (var block in RichUsmtLogs.Document.Blocks) block.FontFamily = font;
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
        TxtCfgOxe.Text = ConfigService.UrlOxe;
        TxtCfgPapercut.Text = ConfigService.UrlPapercut;
        TxtCfgBancada.Text = ConfigService.UrlBancadaLocal;
        TxtCfgBancadaWslDir.Text = ConfigService.BancadaWslDir;
        TxtCfgBancadaWslDistro.Text = ConfigService.BancadaWslDistro;
        TxtCfgWhatsapp.Text = ConfigService.WhatsappPwaArgs;
        TxtCfgDeviceReportsDir.Text = ConfigService.DeviceReportsDir;
        TxtCfgChamadosReportsDir.Text = ConfigService.ChamadosReportsDir;
        TxtCfgCmRcViewer.Text = ConfigService.CmRcViewerCustomPath;
        TxtCfgFirefoxPath.Text = ConfigService.FirefoxBinaryPath;
        TxtCfgUsmtPath.Text = ConfigService.UsmtPath;
        TxtCfgDbHost.Text = ConfigService.BancadaDbHost;
        TxtCfgDbPort.Text = ConfigService.BancadaDbPort.ToString();
        TxtCfgDbName.Text = ConfigService.BancadaDbName;
        TxtCfgDbUser.Text = ConfigService.BancadaDbUser;
        TxtCfgDbPassword.Password = ConfigService.BancadaDbPassword;
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
            _trayIconService.ClipboardTextProvider = () =>
            {
                try
                {
                    return System.Windows.Clipboard.ContainsText() ? System.Windows.Clipboard.GetText() : string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            };
            _trayIconService.OnOpenRequested += RestoreFromTray;
            _trayIconService.OnExitRequested += ExitApplication;
            _trayIconService.OnToggleAdminModeRequested += ToggleAdminMode;
            _trayIconService.OnLaunchEnvironmentRequested += TrayLaunchEnvironment;
            _trayIconService.OnQuickCleanRequested += TrayQuickClean;
            _trayIconService.OnStartDockerRequested += () => BtnStartBancada_Click(this, new RoutedEventArgs());
            _trayIconService.OnStopDockerRequested += () => BtnStopBancada_Click(this, new RoutedEventArgs());
            _trayIconService.OnRebuildDockerRequested += () => BtnRebuildBancada_Click(this, new RoutedEventArgs());
            _trayIconService.OnOpenBancadaBrowserRequested += () => BtnOpenBancadaBrowser_Click(this, new RoutedEventArgs());
            _trayIconService.OnOpenBancadaFolderRequested += () => BtnOpenBancadaFolder_Click(this, new RoutedEventArgs());
            _trayIconService.OnCheckBancadaHealthRequested += async () =>
            {
                await CheckBancadaStatusAsync();
                bool online = await _bancadaService.IsBancadaOnlineAsync();
                _trayIconService?.ShowBalloon("Bancada Streamlit", online ? "O sistema está ONLINE (porta 8501)." : "O sistema está OFFLINE / PARADO.");
            };
            _trayIconService.OnRemoteConnectClipboardRequested += TrayRemoteConnect;
            _trayIconService.OnOpenAppLogsRequested += () =>
            {
                RestoreFromTray();
                NavListBox.SelectedIndex = 15;
            };
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

    private void TrayLaunchEnvironment(string modeName)
    {
        try
        {
            _envService.LaunchEnvironment(modeName, cleanFirst: false);
            _trayIconService?.ShowBalloon("Ambientes Hub", $"Ambiente 'Modo {modeName}' disparado com sucesso.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao alternar para modo {modeName} via bandeja: {ex.Message}", nameof(MainWindow));
        }
    }

    private async void TrayQuickClean()
    {
        _trayIconService?.ShowBalloon("Saneamento", "Iniciando saneamento de processos e liberação de memória...");
        try
        {
            var options = new CleanupOptions
            {
                CloseOffice = true,
                CloseBrowsers = true,
                CloseTeams = true,
                StopTelemetry = true,
                RestartExplorer = false
            };

            await Task.Run(() => _cleanupService.CleanEnvironment(options));
            _trayIconService?.ShowBalloon("Saneamento Concluído", "Memória RAM liberada e processos do Office/Browsers saneados.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro no saneamento rápido via bandeja: {ex.Message}", nameof(MainWindow));
        }
    }

    private void TrayRemoteConnect(bool useSccm)
    {
        string raw = System.Windows.Clipboard.ContainsText() ? System.Windows.Clipboard.GetText().Trim() : string.Empty;
        if (string.IsNullOrWhiteSpace(raw) || !InputValidator.IsValidHostName(raw))
        {
            RestoreFromTray();
            NavListBox.SelectedIndex = 2; // Painel Remoto
            System.Windows.MessageBox.Show(
                "Nenhum nome de computador ou endereço IP válido foi detectado na sua Área de Transferência.\n\nDigite o alvo diretamente no campo de Acesso Remoto.",
                "Acesso Remoto",
                System.Windows.MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        try
        {
            if (useSccm)
            {
                _remoteToolsService.OpenSccmRemoteControl(raw, ConfigService.SccmSiteServer);
                _trayIconService?.ShowBalloon("SCCM Remoto", $"Conectando ao console SCCM na máquina '{raw}'...");
            }
            else
            {
                _remoteToolsService.OpenRdp(raw);
                _trayIconService?.ShowBalloon("RDP Padrão", $"Conectando via mstsc na máquina '{raw}'...");
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Falha ao iniciar conexão remota: {ex.Message}", "Erro de Conexão", System.Windows.MessageBoxButton.OK, MessageBoxImage.Error);
        }
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
                ["URL_OXE"] = TxtCfgOxe.Text.Trim(),
                ["URL_PAPERCUT"] = TxtCfgPapercut.Text.Trim(),
                ["URL_BANCADA_LOCAL"] = TxtCfgBancada.Text.Trim(),
                ["BANCADA_WSL_DIR"] = TxtCfgBancadaWslDir.Text.Trim(),
                ["BANCADA_WSL_DISTRO"] = TxtCfgBancadaWslDistro.Text.Trim(),
                ["WHATSAPP_PWA_ARGS"] = TxtCfgWhatsapp.Text.Trim(),
                ["DEVICE_REPORTS_DIR"] = TxtCfgDeviceReportsDir.Text.Trim(),
                ["CHAMADOS_REPORTS_DIR"] = TxtCfgChamadosReportsDir.Text.Trim(),
                ["CMRCVIEWER_PATH"] = TxtCfgCmRcViewer.Text.Trim(),
                ["FIREFOX_PATH"] = TxtCfgFirefoxPath.Text.Trim(),
                ["USMT_PATH"] = TxtCfgUsmtPath.Text.Trim(),
                ["PROTECTED_PROFILES"] = TxtCfgProtectedProfiles.Text.Trim(),
                ["PING_TIMEOUT_MS"] = TxtCfgPingTimeout.Text.Trim(),
                ["BANCADA_DB_HOST"] = TxtCfgDbHost.Text.Trim(),
                ["BANCADA_DB_PORT"] = TxtCfgDbPort.Text.Trim(),
                ["BANCADA_DB_NAME"] = TxtCfgDbName.Text.Trim(),
                ["BANCADA_DB_USER"] = TxtCfgDbUser.Text.Trim(),
                ["BANCADA_DB_PASSWORD"] = TxtCfgDbPassword.Password.Trim()
            };

            ConfigService.SaveConfig(dict);
            ApplyTerminalFont(chosenFont);
            RefreshUsmtBinStatusBadge();

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
                case "URL_OXE": TxtCfgOxe.Text = defVal; break;
                case "URL_PAPERCUT": TxtCfgPapercut.Text = defVal; break;
                case "URL_BANCADA_LOCAL": TxtCfgBancada.Text = defVal; break;
                case "BANCADA_WSL_DIR": TxtCfgBancadaWslDir.Text = defVal; break;
                case "BANCADA_WSL_DISTRO": TxtCfgBancadaWslDistro.Text = defVal; break;
                case "WHATSAPP_PWA_ARGS": TxtCfgWhatsapp.Text = defVal; break;
                case "DEVICE_REPORTS_DIR": TxtCfgDeviceReportsDir.Text = defVal; break;
                case "CHAMADOS_REPORTS_DIR": TxtCfgChamadosReportsDir.Text = defVal; break;
                case "CMRCVIEWER_PATH": TxtCfgCmRcViewer.Text = defVal; break;
                case "FIREFOX_PATH": TxtCfgFirefoxPath.Text = defVal; break;
                case "USMT_PATH": TxtCfgUsmtPath.Text = defVal; break;
                case "PROTECTED_PROFILES": TxtCfgProtectedProfiles.Text = defVal; break;
                case "BANCADA_DB_HOST": TxtCfgDbHost.Text = defVal; break;
                case "BANCADA_DB_PORT": TxtCfgDbPort.Text = defVal; break;
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
        _trayIconService?.ShowBalloon("Desprovisionamento AD & SCCM", $"Rotina finalizada para a máquina '{host}'.", TrayIconService.NIIF_INFO);
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

            if (success)
            {
                if (pkg.Method == SoftwareInstallMethod.WindowsAdkUsmt)
                {
                    TxtCfgUsmtPath.Text = ConfigService.UsmtPath;
                    RefreshUsmtBinStatusBadge();
                }
                _trayIconService?.ShowBalloon("Apps & Provisionamento", $"{pkg.Name} instalado com sucesso!", TrayIconService.NIIF_INFO);
            }
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

            if (res.Success)
            {
                _trayIconService?.ShowBalloon("Analisador de Dispositivo", $"Diagnóstico de '{target}' concluído com sucesso. Relatórios gerados em DeviceReports.", TrayIconService.NIIF_INFO);
            }
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
                _trayIconService?.ShowBalloon("Manutenção Remota Concluída", $"Máquina '{res.ComputerName}': +{res.FreedGb:N2} GB de espaço recuperados.", TrayIconService.NIIF_INFO);
            }
            else
            {
                TxtFooterStatus.Text = $"Manutenção em '{target}' encerrada com alertas.";
                _trayIconService?.ShowBalloon("Manutenção Remota", $"Operação em '{target}' finalizada com alertas.", TrayIconService.NIIF_WARNING);
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
            _trayIconService?.ShowBalloon("Perfis Remotos Excluídos", $"{res.RemovedCount} perfil(is) removido(s) na máquina '{res.ComputerName}' (~{res.FreedEstimateGb:N2} GB liberados).", TrayIconService.NIIF_INFO);
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

            // Se os containers já estão de pé mas o streaming de logs ainda não está conectado no terminal:
            if (!_bancadaService.IsProcessRunning)
            {
                _ = Task.Run(async () =>
                {
                    await _bancadaService.AttachLogsAsync(msg =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            AppendColoredLog(RichBancadaLogs, msg);
                            InspectBancadaLogForSystemNotification(msg);
                        });
                    });
                });
            }
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
        BtnRebuildBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Iniciando containers Docker da Bancada via WSL...";
        AppendColoredLog(RichBancadaLogs, "=== INICIALIZANDO SISTEMA BANCADA STI (DOCKER WSL) ===");
        AppendColoredLog(RichBancadaLogs, $"WSL Distro: {ConfigService.BancadaWslDistro}");
        AppendColoredLog(RichBancadaLogs, $"Diretório: {ConfigService.BancadaWslDir}");
        AppendColoredLog(RichBancadaLogs, $"Comando: docker compose up -d");

        _ = Task.Run(async () =>
        {
            await _bancadaService.StartDockerBancadaAsync(msg =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendColoredLog(RichBancadaLogs, msg);
                    InspectBancadaLogForSystemNotification(msg);
                });
            });
        });

        // Espera 3 segundos para reativar botões e checar status
        await Task.Delay(3500);
        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        BtnRebuildBancada.IsEnabled = true;
        await CheckBancadaStatusAsync();
    }

    private async void BtnStopBancada_Click(object sender, RoutedEventArgs e)
    {
        BtnStartBancada.IsEnabled = false;
        BtnStopBancada.IsEnabled = false;
        BtnRestartBancada.IsEnabled = false;
        BtnRebuildBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Parando containers Docker da Bancada...";
        AppendColoredLog(RichBancadaLogs, "=== PARANDO CONTAINERS (DOCKER COMPOSE DOWN) ===");

        await _bancadaService.StopDockerBancadaAsync(msg =>
        {
            Dispatcher.Invoke(() => AppendColoredLog(RichBancadaLogs, msg));
        });

        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        BtnRebuildBancada.IsEnabled = true;
        await CheckBancadaStatusAsync();
        TxtFooterStatus.Text = "Containers da Bancada encerrados.";
    }

    private async void BtnRestartBancada_Click(object sender, RoutedEventArgs e)
    {
        BtnStartBancada.IsEnabled = false;
        BtnStopBancada.IsEnabled = false;
        BtnRestartBancada.IsEnabled = false;
        BtnRebuildBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Reiniciando stack Docker da Bancada...";
        AppendColoredLog(RichBancadaLogs, "=== REINICIANDO CONTAINERS DOCKER ===");

        await _bancadaService.RestartDockerBancadaAsync(msg =>
        {
            Dispatcher.Invoke(() =>
            {
                AppendColoredLog(RichBancadaLogs, msg);
                InspectBancadaLogForSystemNotification(msg);
            });
        });

        await Task.Delay(3000);
        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        BtnRebuildBancada.IsEnabled = true;
        await CheckBancadaStatusAsync();
    }

    private async void BtnRebuildBancada_Click(object sender, RoutedEventArgs e)
    {
        BtnStartBancada.IsEnabled = false;
        BtnStopBancada.IsEnabled = false;
        BtnRestartBancada.IsEnabled = false;
        BtnRebuildBancada.IsEnabled = false;
        TxtFooterStatus.Text = "Reconstruindo imagens Docker da Bancada (build --no-cache)...";
        AppendColoredLog(RichBancadaLogs, "=== RECONSTRUINDO CONTAINERS (BUILD --NO-CACHE) ===");

        _ = Task.Run(async () =>
        {
            await _bancadaService.RebuildDockerBancadaAsync(msg =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendColoredLog(RichBancadaLogs, msg);
                    InspectBancadaLogForSystemNotification(msg);
                });
            });
        });

        await Task.Delay(4000);
        BtnStartBancada.IsEnabled = true;
        BtnStopBancada.IsEnabled = true;
        BtnRestartBancada.IsEnabled = true;
        BtnRebuildBancada.IsEnabled = true;
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

    private async void BtnAttachBancadaLogs_Click(object sender, RoutedEventArgs e)
    {
        TxtFooterStatus.Text = "Conectando fluxo de logs do Docker Compose...";
        AppendColoredLog(RichBancadaLogs, "=== CONECTANDO AO FLUXO DE LOGS (DOCKER COMPOSE LOGS -F) ===");

        await Task.Run(async () =>
        {
            await _bancadaService.AttachLogsAsync(msg =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendColoredLog(RichBancadaLogs, msg);
                    InspectBancadaLogForSystemNotification(msg);
                });
            });
        });
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

    private void InspectBancadaLogForSystemNotification(string logLine)
    {
        if (string.IsNullOrWhiteSpace(logLine)) return;

        // Remove códigos de escape ANSI de cor se presentes (ex: [36m, [0m)
        string cleanText = System.Text.RegularExpressions.Regex.Replace(logLine, @"\x1B\[[^@-~]*[@-~]", "");

        // 1. Notificação de Tarefa Concluída com Sucesso (ex: orquestrador_chamados)
        if (cleanText.Contains("Tarefa", StringComparison.OrdinalIgnoreCase) && 
            (cleanText.Contains("finalizada com sucesso", StringComparison.OrdinalIgnoreCase) || cleanText.Contains("executado com sucesso", StringComparison.OrdinalIgnoreCase)))
        {
            string summary = cleanText;
            int idx = cleanText.IndexOf("Tarefa", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) summary = cleanText[idx..].Trim();

            _trayIconService?.ShowBalloon("Bancada STI — Tarefa Concluída", summary, TrayIconService.NIIF_INFO);
            return;
        }

        // 2. Notificação de Erro Crítico ou Falha em Tarefa
        if (cleanText.Contains("Falha na tarefa", StringComparison.OrdinalIgnoreCase) || 
            cleanText.Contains("Erro crítico na execução", StringComparison.OrdinalIgnoreCase))
        {
            string summary = cleanText;
            int idx = cleanText.IndexOf("Falha", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = cleanText.IndexOf("Erro", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) summary = cleanText[idx..].Trim();

            _trayIconService?.ShowBalloon("Bancada STI — Alerta / Falha", summary, TrayIconService.NIIF_WARNING);
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
            _trayIconService?.ShowBalloon("Otimizações de Registro", $"{applied} tweak(s) aplicados com sucesso! Ponto de restauração salvo em Backups.", TrayIconService.NIIF_INFO);
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRegistryLogs, $"[ERRO] Exceção durante aplicação: {ex.Message}");
            TxtFooterStatus.Text = "Erro ao aplicar tweaks de registro.";
            _trayIconService?.ShowBalloon("Ajustes de Registro", "Falha ao aplicar tweaks de registro.", TrayIconService.NIIF_WARNING);
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
            _trayIconService?.ShowBalloon("Restauração de Registro", $"{restored} tweak(s) restaurados com sucesso aos valores originais.", TrayIconService.NIIF_INFO);
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichRegistryLogs, $"[ERRO] Exceção durante rollback: {ex.Message}");
            TxtFooterStatus.Text = "Erro no rollback de registro.";
            _trayIconService?.ShowBalloon("Restauração de Registro", "Falha durante o rollback de registro.", TrayIconService.NIIF_WARNING);
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

    #region Painel 10: Backup & Restore de Usuários (USMT)
    private void RefreshUsmtBinStatusBadge()
    {
        if (TxtUsmtBinStatus == null || BadgeUsmtBinStatus == null)
            return;

        string? usmtDir = UsmtBackupService.FindUsmtDirectory();
        if (usmtDir != null)
        {
            TxtUsmtBinStatus.Text = "🟢 USMT Pronto";
            TxtUsmtBinStatus.Foreground = new SolidColorBrush(Color.FromRgb(52, 211, 153));
            BadgeUsmtBinStatus.Background = new SolidColorBrush(Color.FromRgb(6, 78, 59));
            BadgeUsmtBinStatus.ToolTip = $"Detectado em: {usmtDir}";

            if (string.IsNullOrWhiteSpace(ConfigService.UsmtPath) && TxtCfgUsmtPath != null)
            {
                ConfigService.SetAndSave("USMT_PATH", usmtDir);
                TxtCfgUsmtPath.Text = usmtDir;
            }
        }
        else
        {
            TxtUsmtBinStatus.Text = "🟡 USMT Não Detectado";
            TxtUsmtBinStatus.Foreground = new SolidColorBrush(Color.FromRgb(251, 191, 36));
            BadgeUsmtBinStatus.Background = new SolidColorBrush(Color.FromRgb(69, 26, 3));
            BadgeUsmtBinStatus.ToolTip = "Configure USMT_PATH ou clique em 'Baixar / Detectar USMT'";
        }
    }

    private void BtnCfgBrowseUsmt_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "Selecione a pasta contendo scanstate.exe e loadstate.exe (USMT)"
        };
        if (dlg.ShowDialog() == true)
        {
            TxtCfgUsmtPath.Text = dlg.FolderName;
        }
    }

    private async void BtnCfgDownloadUsmt_Click(object sender, RoutedEventArgs e)
    {
        BtnCfgDownloadUsmt.IsEnabled = false;
        BtnUsmtDownloadBin.IsEnabled = false;
        TxtFooterStatus.Text = "Verificando / baixando Windows ADK USMT...";

        try
        {
            var (success, usmtPath, message) = await _usmtService.DownloadAndInstallUsmtAsync(
                silent: true,
                msg => Dispatcher.Invoke(() => AppendColoredLog(RichUsmtLogs, msg)));

            if (success && !string.IsNullOrWhiteSpace(usmtPath))
            {
                TxtCfgUsmtPath.Text = usmtPath;
                RefreshUsmtBinStatusBadge();
                TxtFooterStatus.Text = $"USMT configurado em: {usmtPath}";
                _trayIconService?.ShowBalloon("USMT Pronto", $"USMT detectado/instalado e salvo em USMT_PATH: {usmtPath}", TrayIconService.NIIF_INFO);
                System.Windows.MessageBox.Show($"USMT pronto e salvo automaticamente no .env:\n{usmtPath}", "USMT Configurado", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                TxtFooterStatus.Text = "Falha ao baixar/detectar USMT.";
                System.Windows.MessageBox.Show(message, "USMT", System.Windows.MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        finally
        {
            BtnCfgDownloadUsmt.IsEnabled = true;
            BtnUsmtDownloadBin.IsEnabled = true;
        }
    }

    private void BtnUsmtDownloadBin_Click(object sender, RoutedEventArgs e)
    {
        RichUsmtLogs.Document.Blocks.Clear();
        BtnCfgDownloadUsmt_Click(sender, e);
    }

    private void TxtUsmtTargetComputer_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
            BtnUsmtQueryUsers_Click(this, new RoutedEventArgs());
    }

    private async void BtnUsmtQueryUsers_Click(object sender, RoutedEventArgs e)
    {
        string target = string.IsNullOrWhiteSpace(TxtUsmtTargetComputer.Text) ? "." : TxtUsmtTargetComputer.Text.Trim();
        BtnUsmtQueryUsers.IsEnabled = false;
        _usmtUsersList.Clear();
        RichUsmtLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Listando perfis de usuário em '{target}'...";

        try
        {
            var profiles = await Task.Run(() => _usmtService.QueryUsersForBackup(
                target,
                msg => Dispatcher.Invoke(() => AppendColoredLog(RichUsmtLogs, msg)),
                calculateDiskSize: true));

            foreach (var p in profiles)
                _usmtUsersList.Add(p);

            TxtUsmtUsersHeader.Text = $"Perfis em '{RemoteManagementHelper.NormalizeTargetHost(target)}': {profiles.Count} usuário(s) — marque os perfis desejados (nenhum será excluído)";
            TxtFooterStatus.Text = $"{profiles.Count} perfil(is) listado(s) para Backup/Restore USMT.";
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichUsmtLogs, $"[ERRO] Falha ao listar perfis: {ex.Message}");
        }
        finally
        {
            BtnUsmtQueryUsers.IsEnabled = true;
        }
    }

    private void BtnUsmtBrowseStore_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "Selecione a pasta de Destino do Backup ou Origem do Restore USMT"
        };
        if (dlg.ShowDialog() == true)
        {
            TxtUsmtStorePath.Text = dlg.FolderName;
        }
    }

    private void BtnUsmtSelectAllUsers_Click(object sender, RoutedEventArgs e)
    {
        foreach (var p in _usmtUsersList)
            p.IsSelected = true;
    }

    private void BtnUsmtClearUsers_Click(object sender, RoutedEventArgs e)
    {
        foreach (var p in _usmtUsersList)
            p.IsSelected = false;
    }

    private UsmtBackupOptions BuildUsmtOptionsFromUi()
    {
        return new UsmtBackupOptions
        {
            TargetComputer = string.IsNullOrWhiteSpace(TxtUsmtTargetComputer.Text) ? "." : TxtUsmtTargetComputer.Text.Trim(),
            StorePath = TxtUsmtStorePath.Text.Trim(),
            SelectedUsers = _usmtUsersList.Where(u => u.IsSelected).Select(u => u.Usuario).ToList(),
            IncludeMigApp = ChkUsmtMigApp.IsChecked == true,
            IncludeMigDocs = ChkUsmtMigDocs.IsChecked == true,
            IncludeMigUser = ChkUsmtMigUser.IsChecked == true,
            Compress = ChkUsmtCompress.IsChecked == true,
            CreateLocalAccountsOnRestore = ChkUsmtCreateLocalAccounts.IsChecked == true,
            EncryptionKey = string.IsNullOrWhiteSpace(TxtUsmtEncryptKey.Text) ? null : TxtUsmtEncryptKey.Text.Trim(),
            CustomUsmtPath = ConfigService.UsmtPath
        };
    }

    private async void BtnUsmtRunBackup_Click(object sender, RoutedEventArgs e)
    {
        var opts = BuildUsmtOptionsFromUi();
        if (string.IsNullOrWhiteSpace(opts.StorePath))
        {
            System.Windows.MessageBox.Show("Informe a pasta de destino para salvar o pacote de backup USMT.", "Destino Obrigatório", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (opts.SelectedUsers.Count == 0)
        {
            System.Windows.MessageBox.Show("Clique em 'Listar Perfis' e marque pelo menos um usuário para fazer backup.", "Nenhum Usuário Marcado", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Iniciar Backup USMT (ScanState) na máquina '{opts.TargetComputer}'?\n\nUsuários selecionados ({opts.SelectedUsers.Count}): {string.Join(", ", opts.SelectedUsers)}\nDestino: {opts.StorePath}\n\n(Nenhum perfil ou arquivo será excluído da máquina.)",
            "Confirmar Backup USMT",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);

        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        SetUsmtButtonsEnabled(false);
        RichUsmtLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Executando Backup USMT em '{opts.TargetComputer}'...";

        try
        {
            var res = await _usmtService.ExecuteBackupAsync(opts, msg => Dispatcher.Invoke(() => AppendColoredLog(RichUsmtLogs, msg)));
            TxtFooterStatus.Text = res.Message;
            _trayIconService?.ShowBalloon(
                res.Success ? "Backup USMT Concluído" : "Backup USMT — Alerta",
                res.Message,
                res.Success ? TrayIconService.NIIF_INFO : TrayIconService.NIIF_WARNING);
        }
        finally
        {
            SetUsmtButtonsEnabled(true);
        }
    }

    private async void BtnUsmtRunRestore_Click(object sender, RoutedEventArgs e)
    {
        var opts = BuildUsmtOptionsFromUi();
        if (string.IsNullOrWhiteSpace(opts.StorePath))
        {
            System.Windows.MessageBox.Show("Informe a pasta onde o pacote de backup USMT está armazenado para restaurar.", "Origem Obrigatória", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string userInfo = opts.SelectedUsers.Count > 0
            ? $"Usuários filtrados ({opts.SelectedUsers.Count}): {string.Join(", ", opts.SelectedUsers)}"
            : "Todos os usuários presentes no pacote USMT";

        var confirm = System.Windows.MessageBox.Show(
            $"Iniciar Restore USMT (LoadState) para a máquina '{opts.TargetComputer}'?\n\nOrigem: {opts.StorePath}\nEscopo: {userInfo}",
            "Confirmar Restore USMT",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);

        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        SetUsmtButtonsEnabled(false);
        RichUsmtLogs.Document.Blocks.Clear();
        TxtFooterStatus.Text = $"Executando Restore USMT em '{opts.TargetComputer}'...";

        try
        {
            var res = await _usmtService.ExecuteRestoreAsync(opts, msg => Dispatcher.Invoke(() => AppendColoredLog(RichUsmtLogs, msg)));
            TxtFooterStatus.Text = res.Message;
            _trayIconService?.ShowBalloon(
                res.Success ? "Restore USMT Concluído" : "Restore USMT — Alerta",
                res.Message,
                res.Success ? TrayIconService.NIIF_INFO : TrayIconService.NIIF_WARNING);
        }
        finally
        {
            SetUsmtButtonsEnabled(true);
        }
    }

    private void BtnUsmtOpenStore_Click(object sender, RoutedEventArgs e)
    {
        string path = TxtUsmtStorePath.Text.Trim();
        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
        {
            Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
        else
        {
            System.Windows.MessageBox.Show("Informe um caminho existente para abrir.", "USMT", System.Windows.MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void SetUsmtButtonsEnabled(bool enabled)
    {
        BtnUsmtQueryUsers.IsEnabled = enabled;
        BtnUsmtRunBackup.IsEnabled = enabled;
        BtnUsmtRunRestore.IsEnabled = enabled;
    }
    #endregion

    #region Painel 12: Chamados (OTRS & CitSmart)

    private async System.Threading.Tasks.Task LoadChamadosAsync()
    {
        string baseFilter = (CmbChamadosBase?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Todas as Bases";
        string statusFilter = (CmbChamadosStatus?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Aberto";
        string tagFilter = (CmbChamadosTag?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Todas as Tags";
        string search = TxtChamadosSearch?.Text?.Trim() ?? string.Empty;

        if (TxtFooterStatus != null) TxtFooterStatus.Text = "Consultando PostgreSQL da Bancada...";
        if (BtnRefreshChamados != null) BtnRefreshChamados.IsEnabled = false;

        try
        {
            var (availTags, availLocs) = await _chamadosService.GetAvailableFiltersAsync(search, baseFilter, statusFilter);
            UpdateChamadosFilterOptions(availTags, availLocs);
            tagFilter = (CmbChamadosTag?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Todas as Tags";

            var (success, items, summary, error) = await _chamadosService.QueryChamadosAsync(
                search, baseFilter, statusFilter, 150, tagFilter, _selectedLocalidades.ToList());

            if (success)
            {
                _allChamadosFound.Clear();
                _allChamadosFound.AddRange(items);
                _chamadosCurrentPage = 1;
                RenderChamadosPage();

                if (TxtChamadosKpiTotal != null) TxtChamadosKpiTotal.Text = summary.Total.ToString();
                if (TxtChamadosKpiAbertos != null) TxtChamadosKpiAbertos.Text = summary.Abertos.ToString();
                if (TxtChamadosKpiOtrs != null) TxtChamadosKpiOtrs.Text = summary.OtrsCount.ToString();
                if (TxtChamadosKpiCitSmart != null) TxtChamadosKpiCitSmart.Text = summary.CitSmartCount.ToString();
            }

            if (TxtFooterStatus != null)
            {
                TxtFooterStatus.Text = !string.IsNullOrEmpty(error)
                    ? error
                    : (items.Count == 0
                        ? "Nenhum chamado encontrado para os filtros aplicados."
                        : $"Exibindo {items.Count} chamado(s) mais recentes.");
            }
        }
        catch (Exception ex)
        {
            if (TxtFooterStatus != null) TxtFooterStatus.Text = $"Erro ao carregar chamados: {ex.Message}";
        }
        finally
        {
            if (BtnRefreshChamados != null) BtnRefreshChamados.IsEnabled = true;
        }
    }

    private void UpdateChamadosFilterOptions(List<string> availTags, List<string> availLocs)
    {
        _isUpdatingChamadosFilters = true;
        try
        {
            if (CmbChamadosTag != null)
            {
                string currentTag = (CmbChamadosTag.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Todas as Tags";
                CmbChamadosTag.Items.Clear();
                var allItem = new System.Windows.Controls.ComboBoxItem { Content = "Todas as Tags" };
                CmbChamadosTag.Items.Add(allItem);
                System.Windows.Controls.ComboBoxItem? toSelect = allItem;

                foreach (var t in availTags)
                {
                    var item = new System.Windows.Controls.ComboBoxItem { Content = t };
                    CmbChamadosTag.Items.Add(item);
                    if (string.Equals(t, currentTag, StringComparison.OrdinalIgnoreCase))
                        toSelect = item;
                }
                CmbChamadosTag.SelectedItem = toSelect;
            }

            if (PanelLocalidadesCheckboxes != null)
            {
                _selectedLocalidades.IntersectWith(availLocs);
                PanelLocalidadesCheckboxes.Children.Clear();
                foreach (var loc in availLocs)
                {
                    var cb = new System.Windows.Controls.CheckBox
                    {
                        Content = loc,
                        Foreground = System.Windows.Media.Brushes.White,
                        Margin = new Thickness(0, 2, 0, 4),
                        IsChecked = _selectedLocalidades.Contains(loc)
                    };
                    cb.Checked += (_, _) =>
                    {
                        _selectedLocalidades.Add(loc);
                        UpdateLocalidadesButtonText();
                        _ = LoadChamadosAsync();
                    };
                    cb.Unchecked += (_, _) =>
                    {
                        _selectedLocalidades.Remove(loc);
                        UpdateLocalidadesButtonText();
                        _ = LoadChamadosAsync();
                    };
                    PanelLocalidadesCheckboxes.Children.Add(cb);
                }
                UpdateLocalidadesButtonText();
            }
        }
        finally
        {
            _isUpdatingChamadosFilters = false;
        }
    }

    private void UpdateLocalidadesButtonText()
    {
        if (BtnChamadosLocalidades == null) return;
        BtnChamadosLocalidades.Content = _selectedLocalidades.Count == 0
            ? "Localidades (Todas)"
            : $"Localidades ({_selectedLocalidades.Count})";
    }

    private void BtnChamadosLocalidades_Click(object sender, RoutedEventArgs e)
    {
        if (PopupChamadosLocalidades != null)
            PopupChamadosLocalidades.IsOpen = !PopupChamadosLocalidades.IsOpen;
    }

    private void BtnClearLocalidades_Click(object sender, RoutedEventArgs e)
    {
        _selectedLocalidades.Clear();
        if (PanelLocalidadesCheckboxes != null)
        {
            foreach (var child in PanelLocalidadesCheckboxes.Children)
            {
                if (child is System.Windows.Controls.CheckBox cb)
                    cb.IsChecked = false;
            }
        }
        UpdateLocalidadesButtonText();
        _ = LoadChamadosAsync();
    }

    private void RenderChamadosPage()
    {
        _chamadosList.Clear();
        int total = _allChamadosFound.Count;
        int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)ChamadosPageSize));
        if (_chamadosCurrentPage > totalPages) _chamadosCurrentPage = totalPages;
        if (_chamadosCurrentPage < 1) _chamadosCurrentPage = 1;

        var pageItems = _allChamadosFound
            .Skip((_chamadosCurrentPage - 1) * ChamadosPageSize)
            .Take(ChamadosPageSize);

        foreach (var item in pageItems)
            _chamadosList.Add(item);

        if (TxtChamadosPageInfo != null)
            TxtChamadosPageInfo.Text = $"Página {_chamadosCurrentPage} de {totalPages} ({total} chamados)";
        if (BtnChamadosPrevPage != null)
            BtnChamadosPrevPage.IsEnabled = _chamadosCurrentPage > 1;
        if (BtnChamadosNextPage != null)
            BtnChamadosNextPage.IsEnabled = _chamadosCurrentPage < totalPages;
    }

    private void BtnChamadosPrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_chamadosCurrentPage > 1)
        {
            _chamadosCurrentPage--;
            RenderChamadosPage();
        }
    }

    private void BtnChamadosNextPage_Click(object sender, RoutedEventArgs e)
    {
        int totalPages = Math.Max(1, (int)Math.Ceiling(_allChamadosFound.Count / (double)ChamadosPageSize));
        if (_chamadosCurrentPage < totalPages)
        {
            _chamadosCurrentPage++;
            RenderChamadosPage();
        }
    }

    private void BtnRefreshChamados_Click(object sender, RoutedEventArgs e)
    {
        _ = LoadChamadosAsync();
    }

    private void TxtChamadosSearch_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            _ = LoadChamadosAsync();
        }
    }

    private void CmbChamadosFilter_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingChamadosFilters) return;
        _ = LoadChamadosAsync();
    }

    private async void BtnExportChamados_Click(object sender, RoutedEventArgs e)
    {
        if (_allChamadosFound.Count == 0)
        {
            System.Windows.MessageBox.Show("Nenhum chamado filtrado para exportar.", "Exportar Chamados", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            return;
        }

        if (TxtFooterStatus != null) TxtFooterStatus.Text = "Gerando relatório HTML responsivo com comentários...";
        var comentariosPorChamado = await _chamadosService.GetComentariosPorChamadosAsync(_allChamadosFound.Select(c => c.Id));

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("""
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Relatório de Chamados - AmbientesHub</title>
            <style>
              :root { --bg: #f1f5f9; --card: #ffffff; --border: #cbd5e1; --text: #0f172a; --muted: #475569; --accent: #0284c7; --loc-bg: #fef3c7; --loc-border: #f59e0b; --loc-text: #92400e; }
              * { box-sizing: border-box; }
              body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 14px; line-height: 1.45; }
              .container { max-width: 1020px; margin: 0 auto; }
              .header { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 16px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
              h1 { font-size: 1.25rem; margin: 0 0 4px 0; }
              .meta { color: var(--muted); font-size: 0.85rem; }
              .toolbar { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
              .toolbar input { flex: 1 1 220px; padding: 8px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.9rem; }
              .btn { border: 1px solid var(--border); background: #fff; color: var(--text); padding: 8px 12px; border-radius: 8px; font-size: 0.85rem; font-weight: 600; cursor: pointer; }
              .btn-primary { background: var(--accent); color: #fff; border-color: var(--accent); }
              details.card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; margin-bottom: 12px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.04); transition: opacity 0.2s ease; }
              details.card.done { opacity: 0.58; border-left: 5px solid #10b981; }
              summary.card-head { list-style: none; padding: 14px 16px; cursor: pointer; display: flex; flex-direction: column; gap: 8px; user-select: none; }
              summary.card-head::-webkit-details-marker { display: none; }
              .head-top { display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }
              .badges { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }
              .badge { font-size: 0.72rem; font-weight: 700; padding: 3px 8px; border-radius: 6px; background: #e2e8f0; color: #1e293b; text-transform: uppercase; }
              .badge-otrs { background: #dbeafe; color: #1e40af; }
              .badge-cit { background: #ccfbf1; color: #0f766e; }
              .badge-open { background: #dcfce7; color: #166534; }
              .id-tag { font-weight: 800; color: var(--accent); font-size: 0.95rem; }
              .done-wrap { display: inline-flex; align-items: center; gap: 5px; font-size: 0.8rem; font-weight: 600; color: #047857; background: #ecfdf5; border: 1px solid #a7f3d0; padding: 4px 8px; border-radius: 8px; white-space: nowrap; }
              .title { font-size: 1.05rem; font-weight: 700; color: var(--text); word-break: break-word; }
              .loc-banner { background: var(--loc-bg); border: 1px solid var(--loc-border); color: var(--loc-text); padding: 7px 11px; border-radius: 8px; font-size: 0.88rem; font-weight: 700; display: flex; align-items: center; gap: 6px; word-break: break-word; }
              .card-body { padding: 0 16px 16px 16px; border-top: 1px solid #e2e8f0; }
              .info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 8px 14px; margin: 12px 0; font-size: 0.84rem; color: #334155; }
              .info-item b { color: #0f172a; }
              .sec { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; margin-top: 10px; font-size: 0.86rem; white-space: pre-wrap; word-break: break-word; }
              .sec-title { font-size: 0.78rem; font-weight: 700; text-transform: uppercase; color: var(--accent); margin-bottom: 4px; letter-spacing: 0.03em; }
              .comment { background: #ffffff; border: 1px solid #cbd5e1; border-left: 4px solid #10b981; border-radius: 8px; padding: 9px 12px; margin-top: 8px; font-size: 0.85rem; white-space: pre-wrap; word-break: break-word; }
              .comment-meta { font-size: 0.76rem; font-weight: 700; color: #047857; margin-bottom: 4px; }
              @media (max-width: 600px) {
                body { padding: 8px; }
                summary.card-head { padding: 12px; }
                .card-body { padding: 0 12px 12px 12px; }
                .head-top { flex-direction: column; }
                .title { font-size: 1rem; }
              }
              @media print {
                @page { margin: 10mm; }
                body { background: #fff; padding: 0; font-size: 11px; }
                .container { max-width: 100%; }
                .no-print { display: none !important; }
                details.card { box-shadow: none; break-inside: avoid; page-break-inside: avoid; border: 1px solid #94a3b8; margin-bottom: 8px; }
                details.card.done { display: none; }
                .info-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 4px 10px; }
                .loc-banner { background: #fffbeb !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                .badge { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
              }
            </style>
            </head>
            <body>
            <div class="container">
            """);

        sb.AppendLine("<div class='header'>");
        sb.AppendLine($"<h1>Relatório de Chamados (<span id='visibleCount'>{_allChamadosFound.Count}</span> de {_allChamadosFound.Count})</h1>");
        sb.AppendLine($"<div class='meta'>Gerado em {DateTime.Now:dd/MM/yyyy HH:mm:ss} • Toque em qualquer chamado para expandir/recolher</div>");
        sb.AppendLine("""
              <div class="toolbar no-print">
                <input id="quickSearch" type="search" placeholder="Filtrar rápido por localidade, nome, ID ou título..." oninput="filterCards()" />
                <button class="btn" onclick="toggleAll(true)">Expandir Todos</button>
                <button class="btn" onclick="toggleAll(false)">Recolher Todos</button>
                <button class="btn btn-primary" onclick="window.print()">🖨️ Imprimir / PDF</button>
              </div>
            </div>
            """);

        static string H(string? s) => System.Net.WebUtility.HtmlEncode(s ?? string.Empty);

        foreach (var c in _allChamadosFound)
        {
            string baseCls = c.BaseDisplay.Contains("OTRS", StringComparison.OrdinalIgnoreCase) ? "badge badge-otrs" : "badge badge-cit";
            string statusCls = c.Status.Equals("Aberto", StringComparison.OrdinalIgnoreCase) ? "badge badge-open" : "badge";
            string safeIdAttr = H(c.Id);

            sb.AppendLine($"<details class='card' open data-id='{safeIdAttr}' data-search='{H((c.Id + " " + c.TituloDisplay + " " + c.LocalizacaoDisplay + " " + c.SolicitanteDisplay + " " + c.TagDisplay + " " + c.Hostname).ToLowerInvariant())}'>");
            sb.AppendLine("<summary class='card-head'>");
            sb.AppendLine("<div class='head-top'>");
            sb.AppendLine($"<div class='badges'><span class='{baseCls}'>{H(c.BaseDisplay)}</span><span class='{statusCls}'>{H(c.Status)}</span><span class='badge'>{H(c.TagDisplay)}</span><span class='id-tag'>#{safeIdAttr}</span></div>");
            sb.AppendLine($"<label class='done-wrap no-print' onclick='event.stopPropagation()'><input type='checkbox' onchange='toggleDone(\"{safeIdAttr}\", this)' /> Concluído</label>");
            sb.AppendLine("</div>");
            sb.AppendLine($"<div class='title'>{H(c.TituloDisplay)}</div>");
            sb.AppendLine($"<div class='loc-banner'>📍 Localidade: {H(c.LocalizacaoDisplay)}</div>");
            sb.AppendLine("</summary>");

            sb.AppendLine("<div class='card-body'>");
            sb.AppendLine("<div class='info-grid'>");
            sb.AppendLine($"<div class='info-item'><b>👤 Solicitante:</b> {H(c.SolicitanteDisplay)}</div>");
            sb.AppendLine($"<div class='info-item'><b>🕒 Criado:</b> {H(c.DataCriacaoBr)}</div>");
            if (!string.IsNullOrWhiteSpace(c.DataAtualizacao))
                sb.AppendLine($"<div class='info-item'><b>🔄 Atualizado:</b> {H(FormatDateBr(c.DataAtualizacao))}</div>");
            if (!string.IsNullOrWhiteSpace(c.Hostname))
                sb.AppendLine($"<div class='info-item'><b>💻 Hostname:</b> {H(c.Hostname)}</div>");
            if (!string.IsNullOrWhiteSpace(c.IpOrigem))
                sb.AppendLine($"<div class='info-item'><b>🌐 IP:</b> {H(c.IpOrigem)}</div>");
            if (InputValidator.IsSafeHttpUrl(c.Link))
                sb.AppendLine($"<div class='info-item no-print'><b>🔗 Link:</b> <a href='{H(c.Link)}' target='_blank' rel='noopener'>Abrir no sistema</a></div>");
            sb.AppendLine("</div>");

            if (!string.IsNullOrWhiteSpace(c.Descricao))
                sb.AppendLine($"<div class='sec'><div class='sec-title'>Descrição</div>{H(c.Descricao.Trim())}</div>");
            if (!string.IsNullOrWhiteSpace(c.Andamento))
                sb.AppendLine($"<div class='sec'><div class='sec-title'>Andamento</div>{H(c.Andamento.Trim())}</div>");

            if (comentariosPorChamado.TryGetValue(c.Id.Trim(), out var coms) && coms.Count > 0)
            {
                sb.AppendLine("<div class='sec'><div class='sec-title'>Comentários / Histórico</div>");
                foreach (var cm in coms)
                {
                    string autor = string.IsNullOrWhiteSpace(cm.Autor) ? "Sistema" : cm.Autor;
                    sb.AppendLine($"<div class='comment'><div class='comment-meta'>{H(autor)} • {H(FormatDateBr(cm.Data))}</div>{H(cm.Texto.Trim())}</div>");
                }
                sb.AppendLine("</div>");
            }

            sb.AppendLine("</div></details>");
        }

        sb.AppendLine("""
            </div>
            <script>
              const KEY = 'ambienteshub_chamados_done_v1';
              function loadDone() {
                try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch { return {}; }
              }
              function saveDone(map) {
                try { localStorage.setItem(KEY, JSON.stringify(map)); } catch {}
              }
              function toggleDone(id, cb) {
                const map = loadDone();
                const card = cb.closest('details.card');
                if (cb.checked) {
                  map[id] = true;
                  card.classList.add('done');
                  card.open = false;
                } else {
                  delete map[id];
                  card.classList.remove('done');
                  card.open = true;
                }
                saveDone(map);
              }
              function toggleAll(openState) {
                document.querySelectorAll('details.card').forEach(d => {
                  if (d.style.display !== 'none') d.open = openState;
                });
              }
              function filterCards() {
                const q = (document.getElementById('quickSearch').value || '').toLowerCase().trim();
                let count = 0;
                document.querySelectorAll('details.card').forEach(d => {
                  const hay = d.getAttribute('data-search') || '';
                  const show = !q || hay.includes(q);
                  d.style.display = show ? '' : 'none';
                  if (show) count++;
                });
                document.getElementById('visibleCount').textContent = count;
              }
              window.addEventListener('DOMContentLoaded', () => {
                const map = loadDone();
                document.querySelectorAll('details.card').forEach(d => {
                  const id = d.getAttribute('data-id');
                  if (id && map[id]) {
                    d.classList.add('done');
                    d.open = false;
                    const cb = d.querySelector('input[type="checkbox"]');
                    if (cb) cb.checked = true;
                  }
                });
              });
            </script>
            </body></html>
            """);

        string outDir = string.IsNullOrWhiteSpace(ConfigService.ChamadosReportsDir)
            ? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "ChamadosReports")
            : ConfigService.ChamadosReportsDir;
        System.IO.Directory.CreateDirectory(outDir);
        string outPath = System.IO.Path.Combine(outDir, $"chamados_relatorio_{DateTime.Now:yyyyMMdd_HHmmss}.html");
        System.IO.File.WriteAllText(outPath, sb.ToString(), System.Text.Encoding.UTF8);
        if (TxtFooterStatus != null) TxtFooterStatus.Text = $"Relatório HTML gerado com {_allChamadosFound.Count} chamado(s): {outPath}";
        Process.Start(new ProcessStartInfo(outPath) { UseShellExecute = true });
    }

    private void BtnOpenChamadosReportsFolder_Click(object sender, RoutedEventArgs e)
    {
        string outDir = string.IsNullOrWhiteSpace(ConfigService.ChamadosReportsDir)
            ? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "ChamadosReports")
            : ConfigService.ChamadosReportsDir;
        System.IO.Directory.CreateDirectory(outDir);
        Process.Start(new ProcessStartInfo("explorer.exe", outDir) { UseShellExecute = true });
    }

    private void BtnChamadosOpenOtrs_Click(object sender, RoutedEventArgs e)
    {
        string url = ConfigService.UrlOtrs;
        if (!InputValidator.IsSafeHttpUrl(url))
        {
            System.Windows.MessageBox.Show("Configure uma URL http/https válida para o OTRS nas Configurações.", "URL Inválida", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        _browserLauncherService.OpenUrlsInFirefox(new[] { url });
    }

    private void BtnChamadosOpenCitSmart_Click(object sender, RoutedEventArgs e)
    {
        string url = ConfigService.UrlCitsmart;
        if (!InputValidator.IsSafeHttpUrl(url))
        {
            System.Windows.MessageBox.Show("Configure uma URL http/https válida para o CitSmart nas Configurações.", "URL Inválida", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        _browserLauncherService.OpenUrlsInFirefox(new[] { url });
    }

    private void BtnChamadoDetails_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not ChamadoItem item)
            return;

        var win = new System.Windows.Window
        {
            Title = $"Chamado #{item.Id} ({item.BaseDisplay})",
            Width = 720,
            Height = 620,
            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
            Owner = this,
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
            Foreground = System.Windows.Media.Brushes.White
        };

        var root = new System.Windows.Controls.Grid { Margin = new Thickness(20) };
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var header = new System.Windows.Controls.StackPanel { Margin = new Thickness(0, 0, 0, 12) };
        header.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"#{item.Id} — {item.TituloDisplay}",
            FontSize = 18,
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            Foreground = System.Windows.Media.Brushes.White
        });
        header.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"{item.BaseDisplay}  |  Status: {item.Status}  |  Criado: {FormatDateBr(item.DataCriacao)}  |  Atualizado: {FormatDateBr(item.DataAtualizacao)}",
            FontSize = 12,
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        System.Windows.Controls.Grid.SetRow(header, 0);
        root.Children.Add(header);

        var scroll = new System.Windows.Controls.ScrollViewer
        {
            VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled
        };
        var contentStack = new System.Windows.Controls.StackPanel();

        // Card de Metadados (2 colunas)
        var metaCard = new System.Windows.Controls.Border
        {
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(14),
            Margin = new Thickness(0, 0, 0, 12)
        };
        var metaGrid = new System.Windows.Controls.Grid();
        metaGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        metaGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var leftMeta = new System.Windows.Controls.StackPanel { Margin = new Thickness(0, 0, 10, 0) };
        leftMeta.Children.Add(CreateMetaLine("Solicitante", $"{item.SolicitanteDisplay}" + (string.IsNullOrWhiteSpace(item.IdCliente) ? "" : $" ({item.IdCliente})")));
        var txtEditCidade = CreateEditableMetaBox("Localização (Cidade / Prédio)", item.CidadePredio, leftMeta, isReadOnly: true);
        var txtEditUnidade = CreateEditableMetaBox("Unidade", item.Unidade, leftMeta, isReadOnly: true);
        var txtEditLocalidade = CreateEditableMetaBox("Localidade", item.LocalidadeFisica, leftMeta, isReadOnly: true);
        System.Windows.Controls.Grid.SetColumn(leftMeta, 0);
        metaGrid.Children.Add(leftMeta);

        var rightMeta = new System.Windows.Controls.StackPanel { Margin = new Thickness(10, 0, 0, 0) };
        var cmbEditStatus = new System.Windows.Controls.ComboBox
        {
            IsEditable = true,
            IsEnabled = false,
            Text = item.Status,
            Margin = new Thickness(0, 2, 0, 6)
        };
        cmbEditStatus.Items.Add("Aberto");
        cmbEditStatus.Items.Add("Em Atendimento");
        cmbEditStatus.Items.Add("Pendente");
        cmbEditStatus.Items.Add("Fechado");
        rightMeta.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Status",
            FontSize = 11,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        rightMeta.Children.Add(cmbEditStatus);

        var txtEditHostname = CreateEditableMetaBox("Hostname", item.Hostname, rightMeta, isReadOnly: true);

        var cmbEditTag = new System.Windows.Controls.ComboBox
        {
            IsEditable = true,
            IsEnabled = false,
            Text = item.Tag,
            Margin = new Thickness(0, 2, 0, 6)
        };
        foreach (var tagOpt in new[]
        {
            "BACKUP", "EVENTO", "FORMATAÇÃO", "GARANTIA", "IMPRESSORA",
            "INSTALAÇÃO HARDWARE", "INSTALAÇÃO SOFTWARE", "MANUTENÇÃO", "MONITOR",
            "MUDANÇA", "PREPARAÇÃO COMPUTADORES", "REDE", "SOLICITAÇÃO SSD",
            "SUPORTE", "TELEFONIA FIXA", "VIAGEM"
        })
        {
            cmbEditTag.Items.Add(tagOpt);
        }
        rightMeta.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Tag",
            FontSize = 11,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        rightMeta.Children.Add(cmbEditTag);
        System.Windows.Controls.Grid.SetColumn(rightMeta, 1);
        metaGrid.Children.Add(rightMeta);

        metaCard.Child = metaGrid;
        contentStack.Children.Add(metaCard);

        // Card de Descrição
        var (descCard, txtEditDesc) = CreateEditableDetailSectionCard(
            "Descrição do Chamado",
            item.Descricao ?? string.Empty,
            isReadOnly: true);
        contentStack.Children.Add(descCard);

        // Card de Andamento
        var (andamentoCard, txtEditAndamento) = CreateEditableDetailSectionCard(
            "Andamento",
            item.Andamento ?? string.Empty,
            isReadOnly: true);
        contentStack.Children.Add(andamentoCard);

        // Seção de Comentários (tabela comentarios no PostgreSQL)
        var comentariosHeader = new System.Windows.Controls.TextBlock
        {
            Text = "Comentários do Chamado",
            FontSize = 14,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(56, 189, 248)),
            Margin = new Thickness(0, 4, 0, 8)
        };
        contentStack.Children.Add(comentariosHeader);

        var comentariosPanel = new System.Windows.Controls.StackPanel();
        comentariosPanel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Carregando comentários...",
            FontSize = 12,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184)),
            Margin = new Thickness(0, 0, 0, 8)
        });
        contentStack.Children.Add(comentariosPanel);

        scroll.Content = contentStack;
        System.Windows.Controls.Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        var footer = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };

        var btnSave = new Wpf.Ui.Controls.Button
        {
            Content = "Salvar Alterações",
            Appearance = ControlAppearance.Primary,
            Visibility = Visibility.Collapsed,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(12, 6, 12, 6)
        };
        var btnEnableEdit = new Wpf.Ui.Controls.Button
        {
            Content = "Habilitar Edição",
            Appearance = ControlAppearance.Secondary,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(12, 6, 12, 6)
        };
        btnEnableEdit.Click += (_, _) =>
        {
            cmbEditStatus.IsEnabled = true;
            txtEditCidade.IsReadOnly = false;
            txtEditUnidade.IsReadOnly = false;
            txtEditLocalidade.IsReadOnly = false;
            txtEditHostname.IsReadOnly = false;
            cmbEditTag.IsEnabled = true;
            txtEditDesc.IsReadOnly = false;
            txtEditAndamento.IsReadOnly = false;
            btnEnableEdit.Visibility = Visibility.Collapsed;
            btnSave.Visibility = Visibility.Visible;
        };
        btnSave.Click += async (_, _) =>
        {
            btnSave.IsEnabled = false;
            var (ok, msg) = await _chamadosService.UpdateChamadoAsync(
                item.Id,
                cmbEditStatus.Text,
                txtEditHostname.Text,
                cmbEditTag.Text,
                txtEditCidade.Text,
                txtEditUnidade.Text,
                txtEditLocalidade.Text,
                txtEditDesc.Text,
                txtEditAndamento.Text);
            btnSave.IsEnabled = true;

            if (ok)
            {
                item.Status = cmbEditStatus.Text.Trim();
                item.Hostname = txtEditHostname.Text.Trim();
                item.Tag = cmbEditTag.Text.Trim();
                item.CidadePredio = txtEditCidade.Text.Trim();
                item.Unidade = txtEditUnidade.Text.Trim();
                item.LocalidadeFisica = txtEditLocalidade.Text.Trim();
                item.Descricao = txtEditDesc.Text.Trim();
                item.Andamento = txtEditAndamento.Text.Trim();
                if (TxtFooterStatus != null) TxtFooterStatus.Text = msg;
                _ = LoadChamadosAsync();
                System.Windows.MessageBox.Show(msg, "Salvo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
            else
            {
                System.Windows.MessageBox.Show(msg, "Erro ao Salvar", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        };
        footer.Children.Add(btnEnableEdit);
        footer.Children.Add(btnSave);

        if (InputValidator.IsSafeHttpUrl(item.Link))
        {
            var btnLink = new Wpf.Ui.Controls.Button
            {
                Content = "Abrir no Firefox",
                Appearance = ControlAppearance.Primary,
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(12, 6, 12, 6)
            };
            btnLink.Click += (_, _) =>
            {
                var btnSave = new Wpf.Ui.Controls.Button
        {
            Content = "Salvar Alterações",
            Appearance = ControlAppearance.Primary,
            Visibility = Visibility.Collapsed,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(12, 6, 12, 6)
        };
        var btnEnableEdit = new Wpf.Ui.Controls.Button
        {
            Content = "Habilitar Edição",
            Appearance = ControlAppearance.Secondary,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(12, 6, 12, 6)
        };
        btnEnableEdit.Click += (_, _) =>
        {
            cmbEditStatus.IsEnabled = true;
            txtEditCidade.IsReadOnly = false;
            txtEditUnidade.IsReadOnly = false;
            txtEditLocalidade.IsReadOnly = false;
            txtEditHostname.IsReadOnly = false;
            cmbEditTag.IsEnabled = true;
            txtEditDesc.IsReadOnly = false;
            txtEditAndamento.IsReadOnly = false;
            btnEnableEdit.Visibility = Visibility.Collapsed;
            btnSave.Visibility = Visibility.Visible;
        };
        btnSave.Click += async (_, _) =>
        {
            btnSave.IsEnabled = false;
            var (ok, msg) = await _chamadosService.UpdateChamadoAsync(
                item.Id,
                cmbEditStatus.Text,
                txtEditHostname.Text,
                cmbEditTag.Text,
                txtEditCidade.Text,
                txtEditUnidade.Text,
                txtEditLocalidade.Text,
                txtEditDesc.Text,
                txtEditAndamento.Text);
            btnSave.IsEnabled = true;

            if (ok)
            {
                item.Status = cmbEditStatus.Text.Trim();
                item.Hostname = txtEditHostname.Text.Trim();
                item.Tag = cmbEditTag.Text.Trim();
                item.CidadePredio = txtEditCidade.Text.Trim();
                item.Unidade = txtEditUnidade.Text.Trim();
                item.LocalidadeFisica = txtEditLocalidade.Text.Trim();
                item.Descricao = txtEditDesc.Text.Trim();
                item.Andamento = txtEditAndamento.Text.Trim();
                if (TxtFooterStatus != null) TxtFooterStatus.Text = msg;
                _ = LoadChamadosAsync();
                System.Windows.MessageBox.Show(msg, "Salvo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
            else
            {
                System.Windows.MessageBox.Show(msg, "Erro ao Salvar", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        };
        footer.Children.Add(btnEnableEdit);
        footer.Children.Add(btnSave);

        if (InputValidator.IsSafeHttpUrl(item.Link))
                    _browserLauncherService.OpenUrlsInFirefox(new[] { item.Link });
            };
            footer.Children.Add(btnLink);
        }

        if (!string.IsNullOrWhiteSpace(item.Hostname))
        {
            var btnRemote = new Wpf.Ui.Controls.Button
            {
                Content = "Acesso Remoto",
                Appearance = ControlAppearance.Secondary,
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(12, 6, 12, 6)
            };
            btnRemote.Click += (_, _) =>
            {
                TxtComputerName.Text = item.Hostname.Trim();
                NavListBox.SelectedIndex = 2;
                win.Close();
            };
            footer.Children.Add(btnRemote);
        }

        var btnClose = new Wpf.Ui.Controls.Button
        {
            Content = "Fechar",
            Appearance = ControlAppearance.Secondary,
            Padding = new Thickness(12, 6, 12, 6)
        };
        btnClose.Click += (_, _) => win.Close();
        footer.Children.Add(btnClose);

        System.Windows.Controls.Grid.SetRow(footer, 2);
        root.Children.Add(footer);

        win.Content = root;
        _ = LoadComentariosIntoPanelAsync(item.Id, comentariosPanel);
        win.ShowDialog();
    }

    private void LoadMachineInfoIntoSettings()
    {
        if (TxtCfgMachineHostname != null) TxtCfgMachineHostname.Text = Environment.MachineName;
        if (TxtCfgMachineUser != null) TxtCfgMachineUser.Text = $"{Environment.UserDomainName}\\{Environment.UserName}";
        if (TxtCfgMachineOs != null) TxtCfgMachineOs.Text = System.Runtime.InteropServices.RuntimeInformation.OSDescription;
        if (TxtCfgMachineIp != null)
        {
            try
            {
                var ips = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName())
                    .AddressList
                    .Where(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork && !System.Net.IPAddress.IsLoopback(a))
                    .Select(a => a.ToString());
                string ipText = string.Join(", ", ips);
                TxtCfgMachineIp.Text = string.IsNullOrWhiteSpace(ipText) ? "127.0.0.1" : ipText;
            }
            catch
            {
                TxtCfgMachineIp.Text = "Indisponível";
            }
        }
    }

    private async System.Threading.Tasks.Task LoadComentariosIntoPanelAsync(string chamadoId, System.Windows.Controls.StackPanel panel)
    {
        var comentarios = await _chamadosService.GetComentariosChamadoAsync(chamadoId);
        panel.Children.Clear();

        if (comentarios.Count == 0)
        {
            panel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "Nenhum comentário registrado para este chamado.",
                FontSize = 12,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184)),
                Margin = new Thickness(0, 0, 0, 12)
            });
            return;
        }

        foreach (var c in comentarios)
        {
            var card = new System.Windows.Controls.Border
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 8)
            };
            var sp = new System.Windows.Controls.StackPanel();
            sp.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = $"{(string.IsNullOrWhiteSpace(c.Autor) ? "Sistema" : c.Autor)}  •  {FormatDateBr(c.Data)}",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 211, 153)),
                Margin = new Thickness(0, 0, 0, 4)
            });
            sp.Children.Add(new System.Windows.Controls.TextBox
            {
                Text = c.Texto,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = System.Windows.Media.Brushes.White,
                FontSize = 13,
                Padding = new Thickness(0)
            });
            card.Child = sp;
            panel.Children.Add(card);
        }
    }

    private static System.Windows.Controls.StackPanel CreateMetaLine(string label, string value)
    {
        var sp = new System.Windows.Controls.StackPanel { Margin = new Thickness(0, 0, 0, 6) };
        sp.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = label,
            FontSize = 11,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        sp.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = string.IsNullOrWhiteSpace(value) ? "-" : value,
            FontSize = 13,
            FontWeight = FontWeights.Medium,
            TextWrapping = TextWrapping.Wrap,
            Foreground = System.Windows.Media.Brushes.White
        });
        return sp;
    }

    private static string FormatDateBr(string? rawDate)
    {
        if (string.IsNullOrWhiteSpace(rawDate)) return "-";
        if (DateTime.TryParse(rawDate, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeLocal, out var dt) ||
            DateTime.TryParse(rawDate, new System.Globalization.CultureInfo("pt-BR"), System.Globalization.DateTimeStyles.AssumeLocal, out dt))
        {
            return dt.ToString("dd/MM/yyyy HH:mm:ss");
        }
        return rawDate.Trim();
    }

    private static System.Windows.Controls.TextBox CreateEditableMetaBox(string label, string value, System.Windows.Controls.StackPanel parent, bool isReadOnly = false)
    {
        parent.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = label,
            FontSize = 11,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        var tb = new System.Windows.Controls.TextBox
        {
            Text = value ?? string.Empty,
            IsReadOnly = isReadOnly,
            FontSize = 13,
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
            Foreground = System.Windows.Media.Brushes.White,
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(6, 4, 6, 4),
            Margin = new Thickness(0, 2, 0, 6)
        };
        parent.Children.Add(tb);
        return tb;
    }

    private static (System.Windows.Controls.Border Card, System.Windows.Controls.TextBox Box) CreateEditableDetailSectionCard(string title, string body, bool isReadOnly = false)
    {
        var card = new System.Windows.Controls.Border
        {
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(14),
            Margin = new Thickness(0, 0, 0, 12)
        };
        var sp = new System.Windows.Controls.StackPanel();
        sp.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = title,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(56, 189, 248)),
            Margin = new Thickness(0, 0, 0, 6)
        });
        var tb = new System.Windows.Controls.TextBox
        {
            Text = body,
            IsReadOnly = isReadOnly,
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 80,
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            Foreground = System.Windows.Media.Brushes.White,
            FontSize = 13,
            Padding = new Thickness(8)
        };
        sp.Children.Add(tb);
        card.Child = sp;
        return (card, tb);
    }

    private void BtnChamadoOpenLink_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.FrameworkElement fe && fe.Tag is ChamadoItem item)
        {
            if (!InputValidator.IsSafeHttpUrl(item.Link))
            {
                System.Windows.MessageBox.Show("O chamado não possui uma URL http/https válida.", "Link Inválido", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            _browserLauncherService.OpenUrlsInFirefox(new[] { item.Link });
        }
    }

    private void BtnChamadoRemoteHost_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.FrameworkElement fe && fe.Tag is ChamadoItem item && !string.IsNullOrWhiteSpace(item.Hostname))
        {
            TxtComputerName.Text = item.Hostname.Trim();
            NavListBox.SelectedIndex = 2; // Acesso Remoto
            if (TxtFooterStatus != null) TxtFooterStatus.Text = $"Computador '{item.Hostname.Trim()}' carregado na aba Acesso Remoto.";
        }
    }

    #endregion

    #region Painel 13: Telefonia OXE

    private async System.Threading.Tasks.Task LoadOxeAsync()
    {
        if (TxtFooterStatus != null) TxtFooterStatus.Text = "Consultando ramais da Central Telefônica OXE...";
        if (BtnRefreshOxe != null) BtnRefreshOxe.IsEnabled = false;

        try
        {
            var (success, items, summary, message) = await _oxeService.QueryRamaisAsync(null);
            if (success)
            {
                _allOxeRaw.Clear();
                _allOxeRaw.AddRange(items);
                PopulateOxeFilterCombos();
                ApplyOxeFilters();

                if (TxtOxeKpiTotal != null) TxtOxeKpiTotal.Text = summary.Total.ToString();
                if (TxtOxeKpiComNome != null) TxtOxeKpiComNome.Text = summary.ComNome.ToString();
                if (TxtOxeKpiComIpMac != null) TxtOxeKpiComIpMac.Text = summary.ComIpOuMac.ToString();
            }

            if (TxtFooterStatus != null) TxtFooterStatus.Text = message;
        }
        catch (Exception ex)
        {
            if (TxtFooterStatus != null) TxtFooterStatus.Text = $"Erro ao carregar ramais OXE: {ex.Message}";
        }
        finally
        {
            if (BtnRefreshOxe != null) BtnRefreshOxe.IsEnabled = true;
        }
    }

    private static string GetOxeField(OxeRamalItem item, params string[] keys)
    {
        foreach (string k in keys)
        {
            if (item.AllFields.TryGetValue(k, out string? v) && !string.IsNullOrWhiteSpace(v))
                return v;
        }
        return string.Empty;
    }

    private void PopulateOxeFilterCombos()
    {
        _isUpdatingOxeFilters = true;
        try
        {
            FillOxeCombo(CmbOxeCategoria, "Categoria: Todas", new[] { "categoria_do_dispositivo", "categoria_dispositivo", "categoria" });
            FillOxeCombo(CmbOxeTipoEstacao, "Estação: Todas", new[] { "tipo_de_estacao", "tipo_estacao", "station_type", "tipo" });
            FillOxeCombo(CmbOxeGrupoCaptura, "Captura: Todos", new[] { "grupo_de_captura", "grupo_captura", "pickup_group" });
            FillOxeCombo(CmbOxeRedePublica, "Rede Pública: Todas", new[] { "cat_rede_pública", "cat_rede_publica", "categoria_rede_publica", "rede_publica", "public_network_category" });
            FillOxeCombo(CmbOxeFuncao, "Função: Todas", new[] { "funcao_role", "funcao", "role" });
        }
        finally
        {
            _isUpdatingOxeFilters = false;
        }
    }

    private void FillOxeCombo(System.Windows.Controls.ComboBox? cmb, string allLabel, string[] fieldKeys)
    {
        if (cmb == null) return;
        string? prev = cmb.SelectedItem as string;
        cmb.Items.Clear();
        cmb.Items.Add(allLabel);

        var values = _allOxeRaw
            .Select(x => GetOxeField(x, fieldKeys))
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v, StringComparer.OrdinalIgnoreCase);

        foreach (var v in values)
            cmb.Items.Add(v);

        cmb.SelectedItem = prev != null && cmb.Items.Contains(prev) ? prev : allLabel;
    }

    private void ApplyOxeFilters()
    {
        string q = TxtOxeSearch?.Text?.Trim() ?? string.Empty;
        string cat = CmbOxeCategoria?.SelectedIndex > 0 ? CmbOxeCategoria.SelectedItem?.ToString() ?? "" : "";
        string est = CmbOxeTipoEstacao?.SelectedIndex > 0 ? CmbOxeTipoEstacao.SelectedItem?.ToString() ?? "" : "";
        string grp = CmbOxeGrupoCaptura?.SelectedIndex > 0 ? CmbOxeGrupoCaptura.SelectedItem?.ToString() ?? "" : "";
        string red = CmbOxeRedePublica?.SelectedIndex > 0 ? CmbOxeRedePublica.SelectedItem?.ToString() ?? "" : "";
        string fun = CmbOxeFuncao?.SelectedIndex > 0 ? CmbOxeFuncao.SelectedItem?.ToString() ?? "" : "";

        bool comMac = ChkOxeComMac?.IsChecked == true;
        bool comIp = ChkOxeComIp?.IsChecked == true;
        bool semMac = ChkOxeSemMac?.IsChecked == true;
        bool semIp = ChkOxeSemIp?.IsChecked == true;

        var filtered = _allOxeRaw.Where(item =>
        {
            if (!string.IsNullOrWhiteSpace(q) && !item.AllFields.Values.Any(v => v.Contains(q, StringComparison.OrdinalIgnoreCase)))
                return false;
            if (!string.IsNullOrEmpty(cat) && !string.Equals(GetOxeField(item, "categoria_do_dispositivo", "categoria_dispositivo", "categoria"), cat, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(est) && !string.Equals(GetOxeField(item, "tipo_de_estacao", "tipo_estacao", "station_type", "tipo"), est, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(grp) && !string.Equals(GetOxeField(item, "grupo_de_captura", "grupo_captura", "pickup_group"), grp, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(red) && !string.Equals(GetOxeField(item, "cat_rede_pública", "cat_rede_publica", "categoria_rede_publica", "rede_publica", "public_network_category"), red, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(fun) && !string.Equals(GetOxeField(item, "funcao_role", "funcao", "role"), fun, StringComparison.OrdinalIgnoreCase))
                return false;
            if (comMac && !BancadaOxeService.HasRealValue(item.Mac)) return false;
            if (comIp && !BancadaOxeService.HasRealValue(item.Ip)) return false;
            if (semMac && BancadaOxeService.HasRealValue(item.Mac)) return false;
            if (semIp && BancadaOxeService.HasRealValue(item.Ip)) return false;
            return true;
        });

        _allOxeFound.Clear();
        _allOxeFound.AddRange(filtered);
        _oxeCurrentPage = 1;
        RenderOxePage();
    }

    private void OxeFilter_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingOxeFilters) return;
        ApplyOxeFilters();
    }

    private void RenderOxePage()
    {
        _oxeList.Clear();
        int total = _allOxeFound.Count;
        int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)OxePageSize));
        if (_oxeCurrentPage > totalPages) _oxeCurrentPage = totalPages;
        if (_oxeCurrentPage < 1) _oxeCurrentPage = 1;

        var pageItems = _allOxeFound
            .Skip((_oxeCurrentPage - 1) * OxePageSize)
            .Take(OxePageSize);

        foreach (var item in pageItems)
            _oxeList.Add(item);

        if (TxtOxePageInfo != null)
            TxtOxePageInfo.Text = $"Página {_oxeCurrentPage} de {totalPages} ({total} ramais)";
        if (BtnOxePrevPage != null)
            BtnOxePrevPage.IsEnabled = _oxeCurrentPage > 1;
        if (BtnOxeNextPage != null)
            BtnOxeNextPage.IsEnabled = _oxeCurrentPage < totalPages;
    }

    private void BtnOxePrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_oxeCurrentPage > 1)
        {
            _oxeCurrentPage--;
            RenderOxePage();
        }
    }

    private void BtnOxeNextPage_Click(object sender, RoutedEventArgs e)
    {
        int totalPages = Math.Max(1, (int)Math.Ceiling(_allOxeFound.Count / (double)OxePageSize));
        if (_oxeCurrentPage < totalPages)
        {
            _oxeCurrentPage++;
            RenderOxePage();
        }
    }

    private void BtnRefreshOxe_Click(object sender, RoutedEventArgs e)
    {
        _ = LoadOxeAsync();
    }

    private void TxtOxeSearch_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            ApplyOxeFilters();
        }
    }

    private void BtnOxeOpenSite_Click(object sender, RoutedEventArgs e)
    {
        string url = ConfigService.UrlOxe;
        if (!InputValidator.IsSafeHttpUrl(url))
        {
            System.Windows.MessageBox.Show("Configure uma URL http/https válida para a Central Telefônica OXE nas Configurações.", "URL Inválida", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        _browserLauncherService.OpenUrlsInFirefox(new[] { url });
    }

    private void BtnOxeDetails_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not OxeRamalItem item)
            return;

        var win = new System.Windows.Window
        {
            Title = $"Detalhes Completos do Ramal — {item.RamalDisplay}",
            Width = 860,
            Height = 700,
            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
            Owner = this,
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
            Foreground = System.Windows.Media.Brushes.White
        };

        var root = new System.Windows.Controls.Grid { Margin = new Thickness(20) };
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var header = new System.Windows.Controls.StackPanel { Margin = new Thickness(0, 0, 0, 12) };
        header.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"📞 Ramal {item.RamalDisplay} — {item.NomeDisplay}",
            FontSize = 19,
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            Foreground = System.Windows.Media.Brushes.White
        });
        header.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Ficha técnica completa extraída da Central Telefônica Alcatel-Lucent OmniPCX Enterprise (OXE).",
            FontSize = 12,
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        System.Windows.Controls.Grid.SetRow(header, 0);
        root.Children.Add(header);

        var scroll = new System.Windows.Controls.ScrollViewer
        {
            VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        };
        var sectionsStack = new System.Windows.Controls.StackPanel();

        string G(params string[] keys)
        {
            string val = GetOxeField(item, keys);
            return string.IsNullOrWhiteSpace(val) ? "-" : val;
        }

        sectionsStack.Children.Add(CreateOxeSectionExpander("👤 Identificação do Assinante & Titular", true, new (string, string)[]
        {
            ("Ramal", G("ramal", "directory_number")),
            ("Login Externo", G("login_externo", "external_login")),
            ("Nome Titular", G("nome_titular", "name", "nome")),
            ("E-mail", G("email", "e_mail")),
            ("Complemento", G("complemento", "complement")),
            ("Idioma (Language_Id)", G("language_id", "idioma")),
            ("Nome Exibido", G("nome_exibido", "display_name")),
            ("Qtd. Usuários", G("qtd_usuarios", "nb_users"))
        }));

        sectionsStack.Children.Add(CreateOxeSectionExpander("🌐 Rede, IP & Hardware", true, new (string, string)[]
        {
            ("Endereço IP", G("endereco_ip", "ip", "ip_address")),
            ("Endereço PABX", $"{G("endereco_equipamento")} (Rack {G("rack", "equipment_address_rack")} / Placa {G("placa", "equipment_address_board")} / Terminal {G("terminal", "equipment_address_terminal")})"),
            ("MAC Address", G("mac_address", "mac")),
            ("Criptografia Nativa", G("criptografia_nativa", "native_encryption")),
            ("Categoria Dispositivo", G("categoria_dispositivo", "categoria_do_dispositivo", "categoria")),
            ("Suporte a Vídeo", G("suporte_video", "video_support"))
        }));

        sectionsStack.Children.Add(CreateOxeSectionExpander("📞 Telefonia, Grupo de Captura & Regras de Acesso", true, new (string, string)[]
        {
            ("Grupo de Captura (Pickup Group)", G("grupo_de_captura", "grupo_captura", "pickup_group")),
            ("Centro de Custo", G("centro_de_custo", "cost_center")),
            ("Cat. Rede Pública", G("cat_rede_pública", "cat_rede_publica", "categoria_rede_publica", "public_network_category")),
            ("Função Adm", G("funcao_adm", "admin_function")),
            ("Cat. Encaminhamento Externo", G("cat_encaminhamento_externo", "external_forwarding_category")),
            ("Cat. Chamador (Caller Category)", G("cat_chamador", "caller_category")),
            ("Cat. Medição / Tarifação", G("cat_medicao", "charging_category")),
            ("Cat. Conexão / Dados", $"{G("cat_conexao")} / {G("cat_dados")}")
        }));

        sectionsStack.Children.Add(CreateOxeSectionExpander("⚙️ Tipo de Estação, Perfil & Teclado", true, new (string, string)[]
        {
            ("Tipo de Estação", G("tipo_de_estacao", "tipo_estacao", "station_type")),
            ("Módulos de Extensão", $"{G("modulo_extensao_1")} / {G("modulo_extensao_2")}"),
            ("Subtipo de Estação", G("subtipo", "station_sub_type")),
            ("Teclado Interno", G("teclado_interno", "internal_keyboard")),
            ("Função / Role", G("funcao_role", "set_role", "funcao")),
            ("Entidade / Domínio", $"{G("entity_number", "entidade")} / {G("domain", "dominio")}"),
            ("Perfil DM / Tipo", $"{G("dm_profile")} | {G("dm_type")}"),
            ("Licença Opex", G("opex_license", "licenca_opex"))
        }));

        sectionsStack.Children.Add(CreateOxeSectionExpander("🎙️ Correio de Voz & Recursos Especiais", true, new (string, string)[]
        {
            ("Tipo Correio de Voz", G("tipo_correio_voz", "voice_mail_type")),
            ("Controle PIN", G("controle_pin", "pin_control")),
            ("Ramal Correio de Voz", G("ramal_correio_voz", "voice_mail_dir_nb")),
            ("Atendimento VIP", G("atendimento_vip", "vip")),
            ("Grupo ACD", G("grupo_acd", "acd_group")),
            ("Chamada Urgente", G("chamada_urgente", "urgent_call"))
        }));

        // Dicionário Completo de Dados Brutos (JSON)
        string rawJson = System.Text.Json.JsonSerializer.Serialize(item.AllFields);
        var rawBox = TryBuildFormattedJsonRichTextBox(rawJson);
        if (rawBox != null)
        {
            rawBox.MaxHeight = 320;
            var rawExpander = new System.Windows.Controls.Expander
            {
                Header = "🔍 Dicionário Completo de Dados Brutos (API OXE)",
                IsExpanded = false,
                Foreground = System.Windows.Media.Brushes.White,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(10),
                Content = rawBox
            };
            sectionsStack.Children.Add(rawExpander);
        }

        scroll.Content = sectionsStack;
        System.Windows.Controls.Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        var footer = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        var btnClose = new Wpf.Ui.Controls.Button
        {
            Content = "Fechar",
            Appearance = ControlAppearance.Primary,
            Padding = new Thickness(14, 6, 14, 6)
        };
        btnClose.Click += (_, _) => win.Close();
        footer.Children.Add(btnClose);
        System.Windows.Controls.Grid.SetRow(footer, 2);
        root.Children.Add(footer);

        win.Content = root;
        win.ShowDialog();
    }

    private static System.Windows.Controls.Expander CreateOxeSectionExpander(string title, bool isExpanded, (string Label, string Value)[] items)
    {
        var grid = new System.Windows.Controls.Grid { Margin = new Thickness(4, 8, 4, 4) };
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        int rows = (int)Math.Ceiling(items.Length / 2.0);
        for (int r = 0; r < rows; r++)
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        for (int i = 0; i < items.Length; i++)
        {
            int r = i / 2;
            int c = i % 2;

            var rowPanel = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                Margin = new Thickness(0, 0, 12, 8)
            };
            rowPanel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = $"{items[i].Label}: ",
                FontSize = 12.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)),
                VerticalAlignment = VerticalAlignment.Center
            });
            rowPanel.Children.Add(new System.Windows.Controls.Border
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 2, 6, 2),
                Child = new System.Windows.Controls.TextBlock
                {
                    Text = items[i].Value,
                    FontSize = 12,
                    FontFamily = new FontFamily("Consolas, Cascadia Code, Courier New"),
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(110, 231, 183)),
                    TextWrapping = TextWrapping.Wrap
                }
            });

            System.Windows.Controls.Grid.SetRow(rowPanel, r);
            System.Windows.Controls.Grid.SetColumn(rowPanel, c);
            grid.Children.Add(rowPanel);
        }

        return new System.Windows.Controls.Expander
        {
            Header = title,
            IsExpanded = isExpanded,
            Foreground = System.Windows.Media.Brushes.White,
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0, 0, 0, 10),
            Padding = new Thickness(10),
            Content = grid
        };
    }

    private static System.Windows.Controls.RichTextBox? TryBuildFormattedJsonRichTextBox(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        string trimmed = raw.Trim();
        if (!(trimmed.StartsWith('{') || trimmed.StartsWith('['))) return null;

        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(trimmed);
            string pretty = System.Text.Json.JsonSerializer.Serialize(
                doc.RootElement,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

            var rtb = new System.Windows.Controls.RichTextBox
            {
                IsReadOnly = true,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
                Padding = new Thickness(8),
                FontFamily = new FontFamily("Consolas, Cascadia Code, Courier New"),
                FontSize = 12,
                VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                MaxHeight = 240
            };

            var para = new Paragraph { Margin = new Thickness(0) };
            var keyBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(125, 211, 252));   // Sky
            var strBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(110, 231, 183));   // Emerald
            var numBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 191, 36));    // Amber
            var puncBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184));  // Slate

            foreach (string line in pretty.Split('\n'))
            {
                string cleanLine = line.TrimEnd('\r');
                int colonIdx = cleanLine.IndexOf("\": ", StringComparison.Ordinal);
                if (colonIdx > 0)
                {
                    string keyPart = cleanLine[..(colonIdx + 1)];
                    string sepPart = ": ";
                    string valPart = cleanLine[(colonIdx + 3)..];
                    para.Inlines.Add(new Run(keyPart) { Foreground = keyBrush });
                    para.Inlines.Add(new Run(sepPart) { Foreground = puncBrush });
                    var valBrush = valPart.TrimStart().StartsWith('"') ? strBrush
                        : (valPart.TrimStart().StartsWith('{') || valPart.TrimStart().StartsWith('[')) ? puncBrush
                        : numBrush;
                    para.Inlines.Add(new Run(valPart + "\n") { Foreground = valBrush });
                }
                else
                {
                    para.Inlines.Add(new Run(cleanLine + "\n") { Foreground = puncBrush });
                }
            }

            rtb.Document = new FlowDocument(para);
            return rtb;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region Painel 17: Impressoras PaperCut

    private async System.Threading.Tasks.Task LoadImpressorasAsync()
    {
        if (TxtFooterStatus != null) TxtFooterStatus.Text = "Consultando impressoras do PaperCut no PostgreSQL...";
        if (BtnRefreshImpressoras != null) BtnRefreshImpressoras.IsEnabled = false;

        try
        {
            var (success, items, summary, message) = await _impressorasService.QueryImpressorasAsync();
            if (success)
            {
                _allImpressorasRaw.Clear();
                _allImpressorasRaw.AddRange(items);
                PopulateImpressorasFilterCombos();
                ApplyImpressorasFilters();

                if (TxtImpressorasKpiTotal != null) TxtImpressorasKpiTotal.Text = summary.Total.ToString();
                if (TxtImpressorasKpiOk != null) TxtImpressorasKpiOk.Text = summary.OkCount.ToString();
                if (TxtImpressorasKpiAlert != null) TxtImpressorasKpiAlert.Text = summary.AlertCount.ToString();
                if (TxtImpressorasKpiPaginas != null) TxtImpressorasKpiPaginas.Text = summary.TotalPaginasSoma.ToString("N0", new System.Globalization.CultureInfo("pt-BR"));
            }

            if (TxtFooterStatus != null) TxtFooterStatus.Text = message;
        }
        catch (Exception ex)
        {
            if (TxtFooterStatus != null) TxtFooterStatus.Text = $"Erro ao carregar impressoras: {ex.Message}";
        }
        finally
        {
            if (BtnRefreshImpressoras != null) BtnRefreshImpressoras.IsEnabled = true;
        }
    }

    private void PopulateImpressorasFilterCombos()
    {
        _isUpdatingImpressorasFilters = true;
        try
        {
            FillStringCombo(CmbImpressorasTipo, "Tipo: Todos", _allImpressorasRaw.Select(x => x.Tipo));
            FillStringCombo(CmbImpressorasStatus, "Status: Todos", _allImpressorasRaw.Select(x => x.Status));
            FillStringCombo(CmbImpressorasServidor, "Servidor: Todos", _allImpressorasRaw.Select(x => x.Servidor));
            FillStringCombo(CmbImpressorasLocalizacao, "Localização: Todas", _allImpressorasRaw.Select(x => x.Localizacao));
        }
        finally
        {
            _isUpdatingImpressorasFilters = false;
        }
    }

    private static void FillStringCombo(System.Windows.Controls.ComboBox? cmb, string allLabel, IEnumerable<string> rawValues)
    {
        if (cmb == null) return;
        string? prev = cmb.SelectedItem as string;
        cmb.Items.Clear();
        cmb.Items.Add(allLabel);

        foreach (var v in rawValues
            .Where(BancadaOxeService.HasRealValue)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v, StringComparer.OrdinalIgnoreCase))
        {
            cmb.Items.Add(v);
        }

        cmb.SelectedItem = prev != null && cmb.Items.Contains(prev) ? prev : allLabel;
    }

    private void ApplyImpressorasFilters()
    {
        string q = TxtImpressorasSearch?.Text?.Trim() ?? string.Empty;
        string tipo = CmbImpressorasTipo?.SelectedIndex > 0 ? CmbImpressorasTipo.SelectedItem?.ToString() ?? "" : "";
        string status = CmbImpressorasStatus?.SelectedIndex > 0 ? CmbImpressorasStatus.SelectedItem?.ToString() ?? "" : "";
        string servidor = CmbImpressorasServidor?.SelectedIndex > 0 ? CmbImpressorasServidor.SelectedItem?.ToString() ?? "" : "";
        string loc = CmbImpressorasLocalizacao?.SelectedIndex > 0 ? CmbImpressorasLocalizacao.SelectedItem?.ToString() ?? "" : "";

        var filtered = _allImpressorasRaw.Where(item =>
        {
            if (!string.IsNullOrWhiteSpace(q) && !item.AllFields.Values.Any(v => v.Contains(q, StringComparison.OrdinalIgnoreCase)))
                return false;
            if (!string.IsNullOrEmpty(tipo) && !string.Equals(item.Tipo, tipo, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(status) && !string.Equals(item.Status, status, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(servidor) && !string.Equals(item.Servidor, servidor, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrEmpty(loc) && !string.Equals(item.Localizacao, loc, StringComparison.OrdinalIgnoreCase))
                return false;
            return true;
        });

        _allImpressorasFound.Clear();
        _allImpressorasFound.AddRange(filtered);
        _impressorasCurrentPage = 1;
        RenderImpressorasPage();
    }

    private void ImpressoraFilter_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingImpressorasFilters) return;
        ApplyImpressorasFilters();
    }

    private void RenderImpressorasPage()
    {
        _impressorasList.Clear();
        int total = _allImpressorasFound.Count;
        int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)ImpressorasPageSize));
        if (_impressorasCurrentPage > totalPages) _impressorasCurrentPage = totalPages;
        if (_impressorasCurrentPage < 1) _impressorasCurrentPage = 1;

        var pageItems = _allImpressorasFound
            .Skip((_impressorasCurrentPage - 1) * ImpressorasPageSize)
            .Take(ImpressorasPageSize);

        foreach (var item in pageItems)
            _impressorasList.Add(item);

        if (TxtImpressorasPageInfo != null)
            TxtImpressorasPageInfo.Text = $"Página {_impressorasCurrentPage} de {totalPages} ({total} impressoras)";
        if (BtnImpressorasPrevPage != null)
            BtnImpressorasPrevPage.IsEnabled = _impressorasCurrentPage > 1;
        if (BtnImpressorasNextPage != null)
            BtnImpressorasNextPage.IsEnabled = _impressorasCurrentPage < totalPages;
    }

    private void BtnImpressorasPrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_impressorasCurrentPage > 1)
        {
            _impressorasCurrentPage--;
            RenderImpressorasPage();
        }
    }

    private void BtnImpressorasNextPage_Click(object sender, RoutedEventArgs e)
    {
        int totalPages = Math.Max(1, (int)Math.Ceiling(_allImpressorasFound.Count / (double)ImpressorasPageSize));
        if (_impressorasCurrentPage < totalPages)
        {
            _impressorasCurrentPage++;
            RenderImpressorasPage();
        }
    }

    private void BtnRefreshImpressoras_Click(object sender, RoutedEventArgs e)
    {
        _ = LoadImpressorasAsync();
    }

    private void TxtImpressorasSearch_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            ApplyImpressorasFilters();
        }
    }

    private void BtnImpressorasOpenSite_Click(object sender, RoutedEventArgs e)
    {
        string url = ConfigService.UrlPapercut;
        if (!InputValidator.IsSafeHttpUrl(url))
        {
            System.Windows.MessageBox.Show("Configure uma URL http/https válida para o Admin PaperCut nas Configurações.", "URL Inválida", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        _browserLauncherService.OpenUrlsInFirefox(new[] { url });
    }

    private void BtnImpressoraOpenIp_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not ImpressoraItem item || !item.HasIpHost)
            return;

        string raw = item.IpHost.Trim();
        string url = raw.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || raw.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            ? raw
            : $"http://{raw}";

        if (!InputValidator.IsSafeHttpUrl(url))
        {
            System.Windows.MessageBox.Show($"Endereço '{item.IpHost}' inválido para abertura no navegador.", "Endereço Inválido", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        _browserLauncherService.OpenUrlsInFirefox(new[] { url });
    }

    private void BtnImpressoraDetails_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe || fe.Tag is not ImpressoraItem item)
            return;

        var win = new System.Windows.Window
        {
            Title = $"Detalhes da Impressora — {item.NomeDisplay}",
            Width = 840,
            Height = 660,
            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
            Owner = this,
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42)),
            Foreground = System.Windows.Media.Brushes.White
        };

        var root = new System.Windows.Controls.Grid { Margin = new Thickness(20) };
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var header = new System.Windows.Controls.StackPanel { Margin = new Thickness(0, 0, 0, 12) };
        header.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"🖨️ {item.NomeDisplay}",
            FontSize = 19,
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            Foreground = System.Windows.Media.Brushes.White
        });
        header.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"Status: {item.StatusDisplay}  |  Tipo: {item.TipoDisplay}  |  Atualizado em: {FormatDateBr(item.DataAtualizacao)}",
            FontSize = 12,
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(148, 163, 184))
        });
        System.Windows.Controls.Grid.SetRow(header, 0);
        root.Children.Add(header);

        var scroll = new System.Windows.Controls.ScrollViewer { VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto };
        var stack = new System.Windows.Controls.StackPanel();

        string D(string? v) => string.IsNullOrWhiteSpace(v) ? "-" : v;

        stack.Children.Add(CreateOxeSectionExpander("🖨️ Identificação & Localização", true, new (string, string)[]
        {
            ("Nome", D(item.Nome)),
            ("Tipo", D(item.Tipo)),
            ("Modelo", D(item.Modelo)),
            ("Localização", D(item.Localizacao)),
            ("Servidor", D(item.Servidor)),
            ("Status", D(item.Status))
        }));

        stack.Children.Add(CreateOxeSectionExpander("🌐 Rede, Filas & Contagem", true, new (string, string)[]
        {
            ("IP / Host", D(item.IpHost)),
            ("Total de Páginas", item.TotalPaginas.ToString("N0", new System.Globalization.CultureInfo("pt-BR"))),
            ("Filas Relacionadas", D(item.FilasRelacionadas)),
            ("Atualizado em", FormatDateBr(item.DataAtualizacao))
        }));

        if (!string.IsNullOrWhiteSpace(item.DetalhesExtra))
        {
            var extraJsonBox = TryBuildFormattedJsonRichTextBox(item.DetalhesExtra);
            object extraContent = extraJsonBox != null
                ? extraJsonBox
                : new System.Windows.Controls.TextBox
                {
                    Text = item.DetalhesExtra,
                    IsReadOnly = true,
                    TextWrapping = TextWrapping.Wrap,
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 12.5
                };

            stack.Children.Add(new System.Windows.Controls.Expander
            {
                Header = "📋 Detalhes Extras (PaperCut)",
                IsExpanded = true,
                Foreground = System.Windows.Media.Brushes.White,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(10),
                Content = extraContent
            });
        }

        string rawJson = System.Text.Json.JsonSerializer.Serialize(item.AllFields);
        var rawBox = TryBuildFormattedJsonRichTextBox(rawJson);
        if (rawBox != null)
        {
            rawBox.MaxHeight = 280;
            stack.Children.Add(new System.Windows.Controls.Expander
            {
                Header = "🔍 Dicionário Completo de Dados Brutos",
                IsExpanded = false,
                Foreground = System.Windows.Media.Brushes.White,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(10),
                Content = rawBox
            });
        }

        scroll.Content = stack;
        System.Windows.Controls.Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        var footer = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        if (item.HasIpHost)
        {
            var btnOpenIp = new Wpf.Ui.Controls.Button
            {
                Content = "Abrir Web IP",
                Appearance = ControlAppearance.Secondary,
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(12, 6, 12, 6)
            };
            btnOpenIp.Click += (_, _) => BtnImpressoraOpenIp_Click(new System.Windows.Controls.Button { Tag = item }, new RoutedEventArgs());
            footer.Children.Add(btnOpenIp);
        }
        var btnClose = new Wpf.Ui.Controls.Button
        {
            Content = "Fechar",
            Appearance = ControlAppearance.Primary,
            Padding = new Thickness(14, 6, 14, 6)
        };
        btnClose.Click += (_, _) => win.Close();
        footer.Children.Add(btnClose);
        System.Windows.Controls.Grid.SetRow(footer, 2);
        root.Children.Add(footer);

        win.Content = root;
        win.ShowDialog();
    }

    #endregion
}
