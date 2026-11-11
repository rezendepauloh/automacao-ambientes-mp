using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Diagnostics;
using System.IO;
using AmbientesHub.Models;
using AmbientesHub.Services;
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

    private void TxtGlobalSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        string query = TxtGlobalSearch.Text.Trim();
        if (string.IsNullOrEmpty(query)) return;

        // Feedback visual na barra de status indicando o filtro da pesquisa
        TxtFooterStatus.Text = $"Filtro ativo: '{query}'";
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
            TxtRemoteStatus.FontFamily = font;

            foreach (var block in RichAdLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichCleanLogs.Document.Blocks) block.FontFamily = font;
            foreach (var block in RichScriptsLogs.Document.Blocks) block.FontFamily = font;
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

        if (File.Exists(xmlPath))
        {
            AppendColoredLog(RichScriptsLogs, "=== VERIFICAÇÃO DO COFRE LOCAL ===");
            AppendColoredLog(RichScriptsLogs, $"[OK] Cofre localizado em: {xmlPath}");
            AppendColoredLog(RichScriptsLogs, $"     Usuário configurado: {ConfigService.AdminUser}");
            AppendColoredLog(RichScriptsLogs, "[SUCESSO] Credencial DPAPI ativa e pronta para uso nas automações!");
        }
        else
        {
            AppendColoredLog(RichScriptsLogs, $"[ERRO] Arquivo não localizado: {xmlPath}");
        }
    }

    private void BtnCriarAdminAdXml_Click(object sender, RoutedEventArgs e)
    {
        string targetUser = ConfigService.AdminAdUser;
        string destDir = @"C:\AmbientesVirtuais\Credenciais";
        string destPath = Path.Combine(destDir, "cred_admin_ad.xml");

        var result = System.Windows.MessageBox.Show(
            $"Deseja abrir o prompt de credencial para salvar com segurança a senha de '{targetUser}' em {destPath}?",
            "Cofre de Credenciais",
            System.Windows.MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != System.Windows.MessageBoxResult.Yes) return;

        RichScriptsLogs.Document.Blocks.Clear();
        AppendColoredLog(RichScriptsLogs, $"=== GERANDO COFRE SEGURO DPAPI PARA: {targetUser} ===");
        AppLoggerService.Info($"Iniciando geração de cofre seguro para {targetUser}", nameof(MainWindow));

        try
        {
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            string scriptTemp = Path.Combine(Path.GetTempPath(), $"save_cred_{Guid.NewGuid():N}.ps1");
            string psScript = $@"
Add-Type -AssemblyName System.Windows.Forms
$cred = Get-Credential -UserName '{targetUser}' -Message 'Digite a senha do Administrador AD para criptografia segura DPAPI:'
if ($cred) {{
    $cred | Export-Clixml -Path '{destPath}' -Force
    Write-Output '[OK] Credencial salva'
}} else {{
    Write-Output '[CANCELADO]'
}}
";
            File.WriteAllText(scriptTemp, psScript, System.Text.Encoding.UTF8);

            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptTemp}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var p = Process.Start(startInfo);
            string stdout = p?.StandardOutput.ReadToEnd() ?? "";
            string stderr = p?.StandardError.ReadToEnd() ?? "";
            p?.WaitForExit();

            try { File.Delete(scriptTemp); } catch { }

            if (!string.IsNullOrWhiteSpace(stderr))
            {
                AppLoggerService.Warning($"Prompt credenciais erro/stderr: {stderr}", nameof(MainWindow));
                AppendColoredLog(RichScriptsLogs, $"  -> Aviso: {stderr.Trim()}");
            }

            if (File.Exists(destPath))
            {
                AppendColoredLog(RichScriptsLogs, $"[SUCESSO] Cofre '{destPath}' criado com êxito!");
                TxtFooterStatus.Text = $"Cofre para {targetUser} criado com sucesso.";
                AppLoggerService.Success($"Cofre {destPath} gerado com sucesso.", nameof(MainWindow));
            }
            else
            {
                AppendColoredLog(RichScriptsLogs, "[i] Operação cancelada ou não concluída.");
                AppLoggerService.Warning($"Geração de cofre para {targetUser} não concluída pelo usuário.", nameof(MainWindow));
            }
        }
        catch (Exception ex)
        {
            AppendColoredLog(RichScriptsLogs, $"[ERRO] Falha ao criar cofre: {ex.Message}");
            AppLoggerService.Danger($"Falha ao gerar cofre: {ex.Message}", nameof(MainWindow));
        }
    }
    #endregion

    #region Color Terminal Helper
    private void AppendColoredLog(System.Windows.Controls.RichTextBox rtb, string text)
    {
        Dispatcher.Invoke(() =>
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
        });
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

    #region Acesso Remoto
    private void BtnSccmRemote_Click(object sender, RoutedEventArgs e)
    {
        string host = TxtComputerName.Text;
        if (string.IsNullOrWhiteSpace(host))
        {
            AppLoggerService.Warning("Tentativa de conexão SCCM com campo vazio.", nameof(MainWindow));
            TxtRemoteStatus.Text = "Aviso: Digite o nome do computador ou IP.";
            return;
        }

        try
        {
            AppLoggerService.Info($"Iniciando SCCM Remote para {host}", nameof(MainWindow));
            _remoteToolsService.OpenSccmRemoteControl(host);
            TxtRemoteStatus.Text = $"Conexão SCCM disparada para: {host}";
            TxtFooterStatus.Text = $"Conexão SCCM: {host}";
            AppLoggerService.Success($"SCCM Remote disparado para {host}", nameof(MainWindow));
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao abrir SCCM Remote para {host}: {ex.Message}", nameof(MainWindow));
            TxtRemoteStatus.Text = $"Erro ao abrir SCCM: {ex.Message}";
        }
    }

    private void BtnRdpRemote_Click(object sender, RoutedEventArgs e)
    {
        string host = TxtComputerName.Text;
        if (string.IsNullOrWhiteSpace(host))
        {
            AppLoggerService.Warning("Tentativa de conexão RDP com campo vazio.", nameof(MainWindow));
            TxtRemoteStatus.Text = "Aviso: Digite o nome do computador ou IP.";
            return;
        }

        try
        {
            AppLoggerService.Info($"Iniciando RDP para {host}", nameof(MainWindow));
            _remoteToolsService.OpenRdp(host);
            TxtRemoteStatus.Text = $"Conexão RDP (mstsc) disparada para: {host}";
            TxtFooterStatus.Text = $"Conexão RDP: {host}";
            AppLoggerService.Success($"RDP disparado para {host}", nameof(MainWindow));
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao abrir RDP para {host}: {ex.Message}", nameof(MainWindow));
            TxtRemoteStatus.Text = $"Erro ao abrir RDP: {ex.Message}";
        }
    }

    private void TxtComputerName_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnSccmRemote_Click(sender, e);
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
            StopTelemetry = ChkTelemetry.IsChecked ?? true
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
