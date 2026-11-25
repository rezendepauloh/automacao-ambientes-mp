using System.Diagnostics;
using System.Windows;
using AmbientesHub.Models;
using AmbientesHub.Services;
using WControls = System.Windows.Controls;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Views;

public partial class PrinterMigrationView : WControls.UserControl
{
    private readonly PrinterMigrationService _migrationService = new();

    public event Action<string>? OnOpenServerInRdpRequested;

    public PrinterMigrationView()
    {
        InitializeComponent();
        UpdateScriptPreview();
    }

    private PrinterMigrationRequest BuildRequestFromUi()
    {
        int port = int.TryParse(TxtPortNumber.Text.Trim(), out int p) ? Math.Clamp(p, 1, 65535) : 9100;
        return new PrinterMigrationRequest
        {
            PrintServer = string.IsNullOrWhiteSpace(TxtPrintServer.Text) ? "srv-print01" : TxtPrintServer.Text.Trim(),
            PrinterName = string.IsNullOrWhiteSpace(TxtPrinterName.Text) ? "PRT-0001" : TxtPrinterName.Text.Trim(),
            NewIpAddress = string.IsNullOrWhiteSpace(TxtNewIp.Text) ? "10.0.0.50" : TxtNewIp.Text.Trim(),
            PortNumber = port,
            PortPrefix = string.IsNullOrWhiteSpace(TxtPortPrefix.Text) ? "PAPERCUT_" : TxtPortPrefix.Text.Trim(),
            DriverName = string.IsNullOrWhiteSpace(TxtDriverName.Text) ? "HP Universal Printing PCL 6" : TxtDriverName.Text.Trim(),
            Location = TxtLocation.Text.Trim(),
            Comment = TxtDriverName.Text.Trim(),
            MacAddress = TxtMacAddress.Text.Trim(),
            SendTestPage = ChkSendTestPage.IsChecked == true,
            OpenPrintQueue = ChkOpenQueue.IsChecked == true
        };
    }

    private void FormInput_TextChanged(object sender, WControls.TextChangedEventArgs e)
    {
        if (!IsLoaded) return;
        UpdateScriptPreview();
    }

    private void FormCheck_Clicked(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded) return;
        UpdateScriptPreview();
    }

    private void BtnRefreshScriptPreview_Click(object sender, RoutedEventArgs e)
    {
        UpdateScriptPreview();
    }

    private void UpdateScriptPreview()
    {
        string prefix = string.IsNullOrWhiteSpace(TxtPortPrefix?.Text) ? "PAPERCUT_" : TxtPortPrefix.Text.Trim();
        string ip = string.IsNullOrWhiteSpace(TxtNewIp?.Text) ? "10.0.0.50" : TxtNewIp.Text.Trim();
        if (TxtPortPreview != null)
            TxtPortPreview.Text = $"{prefix}{ip}";

        if (TxtScriptConsole == null) return;

        try
        {
            var req = BuildRequestFromUi();
            TxtScriptConsole.Text = PrinterMigrationService.BuildPrintServerPowerShellScript(req, dryRunOnly: false);
            if (TxtConsoleHeader != null)
                TxtConsoleHeader.Text = $"Script PowerShell Gerado — {req.PrinterName} → {req.ExpectedPortName}";
        }
        catch (ArgumentException ex)
        {
            TxtScriptConsole.Text = $"# Aguardando preenchimento válido: {ex.Message}";
        }
    }

    private async void BtnCheckPrinterNet_Click(object sender, RoutedEventArgs e)
    {
        string ip = TxtNewIp.Text.Trim();
        int port = int.TryParse(TxtPortNumber.Text.Trim(), out int p) ? Math.Clamp(p, 1, 65535) : 9100;

        NetCheckChipsPanel.Children.Clear();
        BtnCheckPrinterNet.IsEnabled = false;
        SetStatus($"🔄 Testando conectividade da impressora ({ip})...", "#1e3a8a", "#93c5fd", "#3b82f6");

        try
        {
            var res = await _migrationService.CheckPrinterNetworkAsync(ip, port);

            NetCheckChipsPanel.Children.Add(CreateStatusChip(
                res.PingSuccess ? $"🟢 Ping OK ({res.PingRttMs} ms)" : "🔴 Sem Ping ICMP",
                res.PingSuccess));

            NetCheckChipsPanel.Children.Add(CreateStatusChip(
                res.RawPort9100Open ? $"🟢 Porta {port} RAW Aberta" : $"🔴 Porta {port} Fechada",
                res.RawPort9100Open));

            NetCheckChipsPanel.Children.Add(CreateStatusChip(
                res.WebInterfaceOpen ? "🟢 Painel Web (80/443)" : "⚪ Web Fechado",
                res.WebInterfaceOpen));

            if (!string.IsNullOrWhiteSpace(res.MacAddressArp))
            {
                NetCheckChipsPanel.Children.Add(CreateStatusChip($"🏷️ ARP: {res.MacAddressArp}", true));
                if (string.IsNullOrWhiteSpace(TxtMacAddress.Text))
                    TxtMacAddress.Text = res.MacAddressArp;
            }

            if (res.IsReadyForInstall)
                SetStatus($"🟢 Impressora acessível em {ip}!", "#064e3b", "#34d399", "#10b981");
            else
                SetStatus($"⚠️ Impressora {ip} não respondeu na rede.", "#78350f", "#fbbf24", "#f59e0b");
        }
        catch (ArgumentException ex)
        {
            SetStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
        }
        finally
        {
            BtnCheckPrinterNet.IsEnabled = true;
        }
    }

    private void BtnCopyServerScript_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var req = BuildRequestFromUi();
            PrinterMigrationService.ValidateRequest(req);
            string script = PrinterMigrationService.BuildPrintServerPowerShellScript(req, dryRunOnly: false);
            TxtScriptConsole.Text = script;

            if (ClipboardHelper.Copy(script))
            {
                SetStatus($"📋 Script de '{req.PrinterName}' copiado! Cole no PowerShell do servidor '{req.PrintServer}'.", "#064e3b", "#34d399", "#10b981");
            }
        }
        catch (ArgumentException ex)
        {
            SetStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
        }
    }

    private async void BtnDryRunRemote_Click(object sender, RoutedEventArgs e)
    {
        await RunRemoteExecutionAsync(dryRunOnly: true);
    }

    private async void BtnExecuteRemote_Click(object sender, RoutedEventArgs e)
    {
        await RunRemoteExecutionAsync(dryRunOnly: false);
    }

    private async Task RunRemoteExecutionAsync(bool dryRunOnly)
    {
        PrinterMigrationRequest req;
        try
        {
            req = BuildRequestFromUi();
            PrinterMigrationService.ValidateRequest(req);
        }
        catch (ArgumentException ex)
        {
            SetStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        BtnDryRunRemote.IsEnabled = false;
        BtnExecuteRemote.IsEnabled = false;
        SetStatus($"🔄 Executando {(dryRunOnly ? "simulação" : "atualização")} em '{req.PrintServer}'...", "#1e3a8a", "#93c5fd", "#3b82f6");

        try
        {
            var (ok, output) = await _migrationService.ExecuteScriptAsync(req, dryRunOnly);
            TxtConsoleHeader.Text = $"Saída da Execução Remota em '{req.PrintServer}' ({(dryRunOnly ? "Simulação" : "Aplicação Real")})";
            TxtScriptConsole.Text = output;

            if (ok)
                SetStatus($"✅ Concluído com sucesso em '{req.PrintServer}'!", "#064e3b", "#34d399", "#10b981");
            else
                SetStatus($"⚠️ Execução remota retornou aviso — veja o console ou use 'Copiar Script'", "#78350f", "#fbbf24", "#f59e0b");
        }
        finally
        {
            BtnDryRunRemote.IsEnabled = true;
            BtnExecuteRemote.IsEnabled = true;
        }
    }

    private void BtnOpenServerInRdpTab_Click(object sender, RoutedEventArgs e)
    {
        string server = TxtPrintServer.Text.Trim();
        OnOpenServerInRdpRequested?.Invoke(server);
    }

    private void BtnOpenPaperCutWeb_Click(object sender, RoutedEventArgs e)
    {
        string url = ConfigService.UrlPapercut;
        if (!InputValidator.IsSafeHttpUrl(url))
        {
            SetStatus("⚠️ Configure URL_PAPERCUT na aba Configurações.", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
        SetStatus("🌐 PaperCut Admin aberto no navegador!", "#064e3b", "#34d399", "#10b981");
    }

    private static WControls.Border CreateStatusChip(string text, bool ok)
    {
        var border = new WControls.Border
        {
            Background = new WMedia.SolidColorBrush(ok
                ? WMedia.Color.FromRgb(6, 78, 59)
                : WMedia.Color.FromRgb(69, 10, 10)),
            BorderBrush = new WMedia.SolidColorBrush(ok
                ? WMedia.Color.FromRgb(16, 185, 129)
                : WMedia.Color.FromRgb(185, 28, 28)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(7, 3, 7, 3),
            Margin = new Thickness(0, 0, 6, 5)
        };
        border.Child = new WControls.TextBlock
        {
            Text = text,
            FontSize = 11,
            FontWeight = FontWeights.SemiBold,
            Foreground = new WMedia.SolidColorBrush(ok
                ? WMedia.Color.FromRgb(52, 211, 153)
                : WMedia.Color.FromRgb(248, 113, 113))
        };
        return border;
    }

    private void SetStatus(string text, string bgHex, string fgHex, string borderHex)
    {
        TxtPrinterMigStatus.Text = text;
        TxtPrinterMigStatus.Foreground = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(fgHex));
        BadgePrinterMigStatus.Background = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(bgHex));
        BadgePrinterMigStatus.BorderBrush = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(borderHex));
    }
}
