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
    private readonly ObservableCollection<SoftwarePackageItem> _softwareCatalog = new(SoftwareProvisioningService.CreateDefaultCatalog());

    private bool _isNavExpanded = true;
    private bool _appsCatalogLoadedOnce;

    public MainWindow()
    {
        InitializeComponent();
        AppLoggerService.Info("Ambientes Hub iniciado com sucesso.", nameof(MainWindow));
        LstSoftwareCatalog.ItemsSource = _softwareCatalog;
        PopulateFontList();
        LoadSettingsIntoUi();
        RefreshPerformanceAdminStatus();
        RefreshAppsAdminStatus();
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

        if (lower is "saneamento" or "limpar cache" or "teams" or "recursos" or "cache" or "explorer")
        {
            NavListBox.SelectedIndex = 3;
            TxtFooterStatus.Text = "Navegado para: Saneamento de Recursos";
            return;
        }

        if (lower is "performance" or "auditoria" or "debloat" or "otimizar" or "otimização" or "otimizacao" or "lentidão" or "lentidao" or "lento" or "wsl" or "trim" or "hibernação" or "hibernacao" or "bloat" or "checkup")
        {
            NavListBox.SelectedIndex = 4;
            TxtFooterStatus.Text = "Navegado para: Performance & Auditoria";
            if (lower is "auditoria" or "checkup")
            {
                BtnExecutarAuditoria_Click(this, new RoutedEventArgs());
            }
            return;
        }

        if (lower is "apps" or "programas" or "instalar" or "instalador" or "rsat" or "dsa" or "console" or "winget" or "provisionamento")
        {
            NavListBox.SelectedIndex = 5;
            TxtFooterStatus.Text = "Navegado para: Apps & Provisionamento";
            return;
        }

        if (lower is "config" or "configuração" or "configurações" or "cofre" or "credenciais" or "ajustes" or "fonte")
        {
            NavListBox.SelectedIndex = 6;
            TxtFooterStatus.Text = "Navegado para: Configurações & Credenciais";
            return;
        }

        if (lower.StartsWith("limpar ") || lower.StartsWith("desprovisionar ") || lower.StartsWith("ad "))
        {
            string target = query.Substring(query.IndexOf(' ') + 1).Trim();
            NavListBox.SelectedIndex = 2;
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
            NavListBox.SelectedIndex = 2;
            TxtFooterStatus.Text = "Navegado para: Limpeza AD & SCCM";
            return;
        }

        // 2. Consulta Global no Acesso Remoto & Diagnóstico (Máquinas e Usuários)
        NavListBox.SelectedIndex = 1;
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
        PanelConfig.Visibility = Visibility.Collapsed;

        switch (NavListBox.SelectedIndex)
        {
            case 0:
                PanelAmbientes.Visibility = Visibility.Visible;
                break;
            case 1:
                PanelRemoto.Visibility = Visibility.Visible;
                break;
            case 2:
                PanelAdSccm.Visibility = Visibility.Visible;
                break;
            case 3:
                PanelSaneamento.Visibility = Visibility.Visible;
                break;
            case 4:
                PanelPerformance.Visibility = Visibility.Visible;
                RefreshPerformanceAdminStatus();
                break;
            case 5:
                PanelApps.Visibility = Visibility.Visible;
                RefreshAppsAdminStatus();
                if (!_appsCatalogLoadedOnce)
                {
                    _appsCatalogLoadedOnce = true;
                    _ = RefreshAppsCatalogAsync();
                }
                break;
            case 6:
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
            RichScriptsLogs.FontFamily = font;
            RichRemoteLogs.FontFamily = font;

            foreach (var block in RichAdLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichCleanLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichPerformanceLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichAppsLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichScriptsLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichRemoteLogs.Document.Blocks) block.FontFamily = font;
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
        TxtCfgAdminUser.Text = ConfigService.AdminUser;
        TxtCfgAdminAdUser.Text = ConfigService.AdminAdUser;
        TxtCfgSccmServer.Text = ConfigService.SccmSiteServer;
        TxtCfgSccmCode.Text = ConfigService.SccmSiteCode;
        TxtCfgOtrs.Text = ConfigService.UrlOtrs;
        TxtCfgCitsmart.Text = ConfigService.UrlCitsmart;
        TxtCfgBancada.Text = ConfigService.UrlBancadaLocal;
        TxtCfgWhatsapp.Text = ConfigService.WhatsappPwaArgs;
    }

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
                ["WHATSAPP_PWA_ARGS"] = TxtCfgWhatsapp.Text.Trim()
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
        if (TxtPerfAdminStatus == null || BadgeAdminStatus == null || BtnReiniciarComoAdmin == null)
            return;

        bool isAdmin = SystemPerformanceService.IsUserAdministrator();
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
}
