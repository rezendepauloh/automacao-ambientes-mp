using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
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

    public MainWindow()
    {
        InitializeComponent();
        AppLoggerService.Info("Ambientes Hub iniciado com sucesso.", nameof(MainWindow));
        PopulateFontList();
        LoadSettingsIntoUi();
    }

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
            TxtRemoteStatus.FontFamily = font;

            foreach (var block in RichAdLogs.Document.Blocks)
            {
                block.FontFamily = font;
            }
            foreach (var block in RichCleanLogs.Document.Blocks)
            {
                block.FontFamily = font;
            }
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
            // Remove quebras extras de linha: o Paragraph já é um bloco de linha no WPF
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
