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

    private bool _isNavExpanded = true;

    public MainWindow()
    {
        InitializeComponent();
        AppLoggerService.Info("Ambientes Hub iniciado com sucesso.", nameof(MainWindow));
        PopulateFontList();
        LoadSettingsIntoUi();
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

        if (lower is "config" or "configuração" or "configurações" or "cofre" or "credenciais" or "ajustes" or "fonte")
        {
            NavListBox.SelectedIndex = 4;
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
            RichScriptsLogs.FontFamily = font;
            RichRemoteLogs.FontFamily = font;

            foreach (var block in RichAdLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichCleanLogs.Document.Blocks) block.FontFamily = font;
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
        else if (text.Contains("[i]") || text.Contains("Aviso") || text.Contains("não localizado") || text.Contains("Dica:"))
        {
            brush = new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Amber 400
        }
        else if (text.StartsWith("===") || text.StartsWith("[1/") || text.StartsWith("[2/"))
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
}
