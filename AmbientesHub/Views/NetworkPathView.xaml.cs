using System.Diagnostics;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Effects;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Wpf.Ui.Controls;
using WControls = System.Windows.Controls;
using WMedia = System.Windows.Media;
using WShapes = System.Windows.Shapes;

namespace AmbientesHub.Views;

public partial class NetworkPathView : WControls.UserControl
{
    private readonly NetworkPathService _networkService = new();
    private readonly NetworkPingStats _pingStats = new();
    private readonly List<NetworkHopNode> _currentHops = [];
    private List<SubnetLocationRule> _subnetRules = [];

    private NetworkTraceResult? _lastTraceResult;
    private NetworkHopNode? _selectedNode;
    private CancellationTokenSource? _traceCts;
    private CancellationTokenSource? _continuousPingCts;

    private bool _isTracing;
    private bool _isContinuousPingRunning;
    private bool _isPanning;
    private Point _panStartPoint;
    private double _panStartHorizontalOffset;
    private double _panStartVerticalOffset;
    private double _zoomScale = 1.0;

    public NetworkPathView()
    {
        InitializeComponent();
        _subnetRules = NetworkPathService.LoadSubnetRules();
    }

    private void TxtTargetHost_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && !_isTracing)
        {
            BtnRunTrace_Click(this, new RoutedEventArgs());
        }
    }

    private async void BtnRunTrace_Click(object sender, RoutedEventArgs e)
    {
        string rawTarget = TxtTargetHost.Text;
        string validatedTarget;
        try
        {
            validatedTarget = NetworkPathService.ValidateTarget(rawTarget);
        }
        catch (ArgumentException ex)
        {
            SetTraceStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        StopActiveOperations();

        _traceCts = new CancellationTokenSource();
        _isTracing = true;
        _currentHops.Clear();
        _selectedNode = null;
        _pingStats.Reset();
        UpdatePingTelemetryUi();
        ClearInspector();
        RenderNetworkGraph();

        UpdateButtonsState();
        SetTraceStatus($"🔄 Rastreando rota para {validatedTarget}...", "#1e3a8a", "#93c5fd", "#3b82f6");

        int maxHops = GetSelectedMaxHops();
        var ct = _traceCts.Token;

        // Dispara um ping inicial rápido em paralelo ao traceroute para alimentar a barra de telemetria
        _ = RunInitialPingBurstAsync(validatedTarget, ct);

        var progress = new Progress<NetworkHopNode>(hop =>
        {
            _currentHops.Add(hop);
            _selectedNode ??= hop;
            RenderNetworkGraph();
            if (_selectedNode == hop)
                UpdateInspector(hop);
        });

        try
        {
            var result = await _networkService.TraceRouteAsync(
                validatedTarget,
                maxHops: maxHops,
                timeoutMs: 1200,
                subnetRules: _subnetRules,
                onHopDiscovered: progress,
                cancellationToken: ct);

            _lastTraceResult = result;
            if (result.Hops.Count > 0)
            {
                var last = result.Hops[^1];
                _selectedNode = last;
                UpdateInspector(last);
                RenderNetworkGraph();
            }

            if (result.TargetOnline)
            {
                SetTraceStatus(
                    $"🟢 Online • {result.Hops.Count - 1} salto(s) até {result.ResolvedTargetIp}",
                    "#064e3b", "#34d399", "#10b981");
            }
            else
            {
                var lastValid = result.Hops.LastOrDefault(h => h.Kind != NetworkHopKind.Destination && h.Kind != NetworkHopKind.Timeout);
                string viaInfo = lastValid != null ? $"último salto: {lastValid.SubtitleDisplay}" : "sem rota ativa";
                SetTraceStatus(
                    $"🔴 Destino Offline • {viaInfo}",
                    "#7f1d1d", "#f87171", "#ef4444");
            }

            double elapsedSec = result.CompletedAt.HasValue
                ? (result.CompletedAt.Value - result.StartedAt).TotalSeconds
                : 0;
            string destUser = result.Hops.LastOrDefault()?.ConnectedUser ?? "-";
            AppLoggerService.Info(
                $"Traceroute '{validatedTarget}' ({(string.IsNullOrWhiteSpace(result.ResolvedTargetIp) ? "sem IP" : result.ResolvedTargetIp)}): " +
                $"{(result.TargetOnline ? "ONLINE" : "OFFLINE")} | Nós={result.Hops.Count} | Usuário={destUser} | Duração={elapsedSec:0.0}s",
                nameof(NetworkPathView));
        }
        catch (OperationCanceledException)
        {
            SetTraceStatus("⏹️ Rastreamento interrompido", "#1e293b", "#cbd5e1", "#475569");
        }
        catch (Exception ex)
        {
            SetTraceStatus($"❌ Falha no rastreamento: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444");
            AppLoggerService.Warning($"Erro no traceroute para {validatedTarget}: {ex.Message}", nameof(NetworkPathView));
        }
        finally
        {
            _isTracing = false;
            UpdateButtonsState();
        }
    }

    private async Task RunInitialPingBurstAsync(string target, CancellationToken ct)
    {
        for (int i = 0; i < 4; i++)
        {
            if (ct.IsCancellationRequested) break;
            try
            {
                var (ok, rtt, _, _, _) = await _networkService.PingOnceAsync(target, 1200, ct);
                _pingStats.RecordSample(ok ? rtt : null);
                UpdatePingTelemetryUi();
                await Task.Delay(350, ct);
            }
            catch (OperationCanceledException) { break; }
            catch { break; }
        }
    }

    private async void BtnToggleContinuousPing_Click(object sender, RoutedEventArgs e)
    {
        if (_isContinuousPingRunning)
        {
            _continuousPingCts?.Cancel();
            _isContinuousPingRunning = false;
            UpdateButtonsState();
            return;
        }

        string validatedTarget;
        try
        {
            validatedTarget = NetworkPathService.ValidateTarget(TxtTargetHost.Text);
        }
        catch (ArgumentException ex)
        {
            SetTraceStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        _continuousPingCts?.Cancel();
        _continuousPingCts = new CancellationTokenSource();
        _isContinuousPingRunning = true;
        _pingStats.Reset();
        UpdatePingTelemetryUi();
        UpdateButtonsState();

        var ct = _continuousPingCts.Token;
        SetTraceStatus($"📡 Ping contínuo ativo em {validatedTarget}...", "#1e3a8a", "#93c5fd", "#3b82f6");

        try
        {
            while (!ct.IsCancellationRequested)
            {
                var (ok, rtt, resolvedIp, _, _) = await _networkService.PingOnceAsync(validatedTarget, 1200, ct);
                _pingStats.RecordSample(ok ? rtt : null);
                UpdatePingTelemetryUi();

                if (ok)
                {
                    SetTraceStatus($"🟢 Ping OK • {resolvedIp} ({rtt} ms)", "#064e3b", "#34d399", "#10b981");
                }
                else
                {
                    SetTraceStatus($"🔴 Sem resposta ICMP de {validatedTarget}", "#7f1d1d", "#f87171", "#ef4444");
                }

                await Task.Delay(1000, ct);
            }
        }
        catch (OperationCanceledException)
        {
            SetTraceStatus("⏹️ Ping contínuo parado", "#1e293b", "#cbd5e1", "#475569");
        }
        finally
        {
            _isContinuousPingRunning = false;
            UpdateButtonsState();
        }
    }

    private void BtnStopTrace_Click(object sender, RoutedEventArgs e)
    {
        StopActiveOperations();
        UpdateButtonsState();
    }

    private void StopActiveOperations()
    {
        _traceCts?.Cancel();
        _continuousPingCts?.Cancel();
        _isTracing = false;
        _isContinuousPingRunning = false;
    }

    private void UpdateButtonsState()
    {
        BtnRunTrace.IsEnabled = !_isTracing;
        BtnStopTrace.IsEnabled = _isTracing || _isContinuousPingRunning;
        BtnToggleContinuousPing.Content = _isContinuousPingRunning ? "Parar Ping" : "Ping Contínuo";
    }

    private void BtnCopyReport_Click(object sender, RoutedEventArgs e)
    {
        if (_lastTraceResult == null && _currentHops.Count == 0)
        {
            SetTraceStatus("ℹ️ Execute um rastreamento antes de copiar o relatório.", "#1e293b", "#cbd5e1", "#475569");
            return;
        }

        var reportModel = _lastTraceResult ?? new NetworkTraceResult
        {
            TargetInput = TxtTargetHost.Text.Trim(),
            Hops = [.. _currentHops]
        };

        string report = NetworkPathService.FormatTraceReport(reportModel);
        if (ClipboardHelper.Copy(report))
        {
            SetTraceStatus("📋 Relatório da rota copiado para a área de transferência!", "#064e3b", "#34d399", "#10b981");
        }
    }

    private void BtnExportHtmlReport_Click(object sender, RoutedEventArgs e)
    {
        if (_lastTraceResult == null && _currentHops.Count == 0)
        {
            SetTraceStatus("ℹ️ Execute um rastreamento antes de exportar o HTML.", "#1e293b", "#cbd5e1", "#475569");
            return;
        }

        var reportModel = _lastTraceResult ?? new NetworkTraceResult
        {
            TargetInput = TxtTargetHost.Text.Trim(),
            Hops = [.. _currentHops],
            CompletedAt = DateTime.Now
        };

        string html = NetworkPathService.FormatTraceHtmlReport(reportModel, _pingStats);
        string safeName = new string(reportModel.TargetInput.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_').ToArray());
        if (string.IsNullOrWhiteSpace(safeName)) safeName = "rota";

        string reportsDir = NetworkPathService.GetReportsDirectory();
        if (!System.IO.Directory.Exists(reportsDir))
            System.IO.Directory.CreateDirectory(reportsDir);

        string filePath = System.IO.Path.Combine(reportsDir, $"rota_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.html");
        System.IO.File.WriteAllText(filePath, html, Encoding.UTF8);

        Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
        AppLoggerService.Success($"Relatório HTML de rota salvo e aberto: {filePath}", nameof(NetworkPathView));
        SetTraceStatus($"🌐 Relatório HTML salvo em {reportsDir}!", "#064e3b", "#34d399", "#10b981");
    }

    private void BtnSubnetMapEditor_Click(object sender, RoutedEventArgs e)
    {
        ShowSubnetMapEditorDialog();
    }

    private void CmbLayoutOrientation_SelectionChanged(object sender, WControls.SelectionChangedEventArgs e)
    {
        if (IsLoaded && _currentHops.Count > 0)
            RenderNetworkGraph();
    }

    #region Renderização da Árvore Bézier Nativa WPF

    private bool IsVerticalLayout()
    {
        return CmbLayoutOrientation.SelectedItem is WControls.ComboBoxItem item &&
               string.Equals(item.Tag?.ToString(), "V", StringComparison.OrdinalIgnoreCase);
    }

    private int GetSelectedMaxHops()
    {
        if (CmbMaxHops.SelectedItem is WControls.ComboBoxItem item &&
            int.TryParse(item.Tag?.ToString(), out int hops))
        {
            return hops;
        }
        return 25;
    }

    private void RenderNetworkGraph()
    {
        GraphCanvas.Children.Clear();
        GraphEmptyState.Visibility = _currentHops.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        if (_currentHops.Count == 0)
        {
            GraphCanvas.Width = 800;
            GraphCanvas.Height = 420;
            return;
        }

        bool vertical = IsVerticalLayout();
        double cardW = vertical ? 290 : 234;
        double cardH = vertical ? 104 : 114;
        double gap = vertical ? 56 : 74;
        double margin = 44;

        var positions = new List<Point>(_currentHops.Count);
        for (int i = 0; i < _currentHops.Count; i++)
        {
            if (vertical)
            {
                double x = margin + 40 + (i % 2 == 0 ? 0 : 42);
                double y = margin + i * (cardH + gap);
                positions.Add(new Point(x, y));
            }
            else
            {
                double x = margin + i * (cardW + gap);
                double y = margin + 56 + (i % 2 == 0 ? 0 : 44);
                positions.Add(new Point(x, y));
            }
        }

        double totalW = positions.Max(p => p.X) + cardW + margin + 40;
        double totalH = positions.Max(p => p.Y) + cardH + margin + 40;
        GraphCanvas.Width = Math.Max(820, totalW);
        GraphCanvas.Height = Math.Max(420, totalH);

        // 1. Desenha as conexões Bézier primeiro (camada inferior)
        for (int i = 0; i < _currentHops.Count - 1; i++)
        {
            var fromPos = positions[i];
            var toPos = positions[i + 1];
            var nextNode = _currentHops[i + 1];

            Point start = vertical
                ? new Point(fromPos.X + cardW / 2, fromPos.Y + cardH)
                : new Point(fromPos.X + cardW, fromPos.Y + cardH / 2);

            Point end = vertical
                ? new Point(toPos.X + cardW / 2, toPos.Y)
                : new Point(toPos.X, toPos.Y + cardH / 2);

            DrawBezierConnection(start, end, nextNode, vertical);
        }

        // 2. Desenha os Cards XAML de cada nó
        for (int i = 0; i < _currentHops.Count; i++)
        {
            var hop = _currentHops[i];
            var pos = positions[i];
            var card = CreateHopCard(hop, cardW, cardH);
            WControls.Canvas.SetLeft(card, pos.X);
            WControls.Canvas.SetTop(card, pos.Y);
            GraphCanvas.Children.Add(card);
        }
    }

    private void DrawBezierConnection(Point start, Point end, NetworkHopNode targetHop, bool vertical)
    {
        WMedia.Color strokeColor = ParseColor(targetHop.BorderBrushHex, WMedia.Color.FromRgb(56, 189, 248));
        var brush = new WMedia.SolidColorBrush(strokeColor);
        bool isBroken = targetHop.Status is NetworkHopStatus.Timeout or NetworkHopStatus.Unreachable;

        Point ctrl1;
        Point ctrl2;
        if (vertical)
        {
            double dy = Math.Max(28, (end.Y - start.Y) * 0.5);
            ctrl1 = new Point(start.X, start.Y + dy);
            ctrl2 = new Point(end.X, end.Y - dy);
        }
        else
        {
            double dx = Math.Max(32, (end.X - start.X) * 0.5);
            ctrl1 = new Point(start.X + dx, start.Y);
            ctrl2 = new Point(end.X - dx, end.Y);
        }

        var figure = new WMedia.PathFigure { StartPoint = start, IsClosed = false };
        figure.Segments.Add(new WMedia.BezierSegment(ctrl1, ctrl2, end, isStroked: true));

        var geometry = new WMedia.PathGeometry();
        geometry.Figures.Add(figure);

        var path = new WShapes.Path
        {
            Data = geometry,
            Stroke = brush,
            StrokeThickness = isBroken ? 2.2 : 2.6,
            StrokeStartLineCap = WMedia.PenLineCap.Round,
            StrokeEndLineCap = WMedia.PenLineCap.Round,
            Opacity = isBroken ? 0.85 : 0.95
        };

        if (isBroken)
            path.StrokeDashArray = new WMedia.DoubleCollection { 6, 4 };

        GraphCanvas.Children.Add(path);

        // Seta discreta na entrada do card destino
        var arrow = new WShapes.Polygon
        {
            Fill = brush,
            Points = vertical
                ? new WMedia.PointCollection
                {
                    new Point(end.X, end.Y),
                    new Point(end.X - 5, end.Y - 8),
                    new Point(end.X + 5, end.Y - 8)
                }
                : new WMedia.PointCollection
                {
                    new Point(end.X, end.Y),
                    new Point(end.X - 8, end.Y - 5),
                    new Point(end.X - 8, end.Y + 5)
                }
        };
        GraphCanvas.Children.Add(arrow);
    }

    private WControls.Border CreateHopCard(NetworkHopNode hop, double width, double height)
    {
        bool isSelected = ReferenceEquals(hop, _selectedNode);
        WMedia.Color borderColor = isSelected
            ? WMedia.Color.FromRgb(56, 189, 248)
            : ParseColor(hop.BorderBrushHex, WMedia.Color.FromRgb(51, 65, 85));

        var card = new WControls.Border
        {
            Width = width,
            Height = height,
            CornerRadius = new CornerRadius(10),
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(15, 23, 42)),
            BorderBrush = new WMedia.SolidColorBrush(borderColor),
            BorderThickness = new Thickness(isSelected ? 2.2 : 1.3),
            Padding = new Thickness(12, 10, 12, 10),
            Cursor = Cursors.Hand,
            Effect = new DropShadowEffect
            {
                Color = isSelected ? WMedia.Color.FromRgb(56, 189, 248) : WMedia.Colors.Black,
                BlurRadius = isSelected ? 14 : 10,
                ShadowDepth = 2,
                Opacity = isSelected ? 0.45 : 0.4
            }
        };

        card.MouseLeftButtonDown += (_, args) =>
        {
            args.Handled = true;
            _selectedNode = hop;
            UpdateInspector(hop);
            RenderNetworkGraph();
        };

        var root = new WControls.Grid();
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });

        // Linha 0: Badge do salto + ícone + pill de latência
        var topRow = new WControls.Grid();
        topRow.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });
        topRow.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topRow.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        var hopBadge = new WControls.Border
        {
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(30, 41, 59)),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(6, 2, 6, 2)
        };
        hopBadge.Child = new WControls.TextBlock
        {
            Text = $"{hop.IconGlyph} {hop.HopBadgeText}",
            FontSize = 10,
            FontWeight = FontWeights.Bold,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(203, 213, 225))
        };
        WControls.Grid.SetColumn(hopBadge, 0);
        topRow.Children.Add(hopBadge);

        var latencyBadge = new WControls.Border
        {
            Background = new WMedia.SolidColorBrush(ParseColor(hop.BadgeBackgroundHex, WMedia.Color.FromRgb(30, 41, 59))),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(7, 2, 7, 2)
        };
        latencyBadge.Child = new WControls.TextBlock
        {
            Text = hop.LatencyDisplay,
            FontSize = 10.5,
            FontWeight = FontWeights.Bold,
            Foreground = new WMedia.SolidColorBrush(ParseColor(hop.StatusColorHex, WMedia.Color.FromRgb(52, 211, 153)))
        };
        WControls.Grid.SetColumn(latencyBadge, 2);
        topRow.Children.Add(latencyBadge);

        WControls.Grid.SetRow(topRow, 0);
        root.Children.Add(topRow);

        // Linha 1: Hostname / PTR + IP
        var centerStack = new WControls.StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 4, 0, 4) };
        centerStack.Children.Add(new WControls.TextBlock
        {
            Text = hop.TitleDisplay,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(248, 250, 252)),
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = hop.TitleDisplay
        });
        string ipLine = hop.SubtitleDisplay;
        if (!string.IsNullOrWhiteSpace(hop.MacAddress))
            ipLine = $"{ipLine} • {hop.MacAddress}";
        if (hop.TtlRemaining.HasValue)
            ipLine = $"{ipLine} • TTL {hop.TtlRemaining}";

        centerStack.Children.Add(new WControls.TextBlock
        {
            Text = ipLine,
            FontSize = 10.5,
            FontFamily = new WMedia.FontFamily("Consolas"),
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(148, 163, 184)),
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = ipLine
        });
        WControls.Grid.SetRow(centerStack, 1);
        root.Children.Add(centerStack);

        // Linha 2: Localização / Dispositivo provável (e usuário conectado na origem/destino)
        string footerText = !string.IsNullOrWhiteSpace(hop.ConnectedUser)
            ? $"👤 {hop.ConnectedUser}"
            : (!string.IsNullOrWhiteSpace(hop.InferredLocation) ? $"📍 {hop.InferredLocation}" : hop.KindLabel);

        var footerStack = new WControls.TextBlock
        {
            Text = footerText,
            FontSize = 10.5,
            Foreground = new WMedia.SolidColorBrush(!string.IsNullOrWhiteSpace(hop.ConnectedUser)
                ? WMedia.Color.FromRgb(52, 211, 153)
                : WMedia.Color.FromRgb(56, 189, 248)),
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = $"{hop.KindLabel} — {hop.LocationDisplay}"
        };
        WControls.Grid.SetRow(footerStack, 2);
        root.Children.Add(footerStack);

        card.Child = root;
        return card;
    }

    #endregion

    #region Telemetria de Ping, Inspetor e Zoom/Pan

    private void UpdatePingTelemetryUi()
    {
        TxtPingSent.Text = _pingStats.Sent.ToString();
        TxtPingRecv.Text = _pingStats.Received.ToString();
        TxtPingLoss.Text = $"{_pingStats.LossPercent:0.#}%";
        TxtPingLoss.Foreground = new WMedia.SolidColorBrush(_pingStats.LossPercent > 0
            ? WMedia.Color.FromRgb(248, 113, 113)
            : WMedia.Color.FromRgb(52, 211, 153));

        string lastStr = _pingStats.LastRttMs.HasValue ? $"{_pingStats.LastRttMs} ms" : "Timeout";
        string avgStr = _pingStats.Received > 0 ? $"{_pingStats.AvgRttMs:0.#} ms" : "-";
        TxtPingLastAvg.Text = $"{lastStr} / {avgStr}";

        string minStr = _pingStats.MinRttMs.HasValue ? $"{_pingStats.MinRttMs}" : "-";
        string maxStr = _pingStats.MaxRttMs.HasValue ? $"{_pingStats.MaxRttMs} ms" : "-";
        TxtPingMinMax.Text = $"{minStr} / {maxStr}";

        PingHistoryBarsPanel.Children.Clear();
        long scaleMax = Math.Max(30, _pingStats.MaxRttMs ?? 30);
        foreach (var sample in _pingStats.History)
        {
            double barHeight = sample.HasValue
                ? Math.Clamp((double)Math.Max(2, sample.Value) / scaleMax * 22.0, 4.0, 22.0)
                : 22.0;

            WMedia.Color barColor = !sample.HasValue
                ? WMedia.Color.FromRgb(239, 68, 68)
                : (sample.Value > 80 ? WMedia.Color.FromRgb(245, 158, 11) : WMedia.Color.FromRgb(16, 185, 129));

            PingHistoryBarsPanel.Children.Add(new WControls.Border
            {
                Width = 6,
                Height = barHeight,
                Margin = new Thickness(2, 0, 0, 0),
                CornerRadius = new CornerRadius(2),
                VerticalAlignment = VerticalAlignment.Bottom,
                Background = new WMedia.SolidColorBrush(barColor),
                ToolTip = sample.HasValue ? $"{sample.Value} ms" : "Timeout"
            });
        }
    }

    private void UpdateInspector(NetworkHopNode node)
    {
        TxtInspectorBadge.Text = $"{node.IconGlyph} {node.HopBadgeText} • {node.Status.ToString().ToUpperInvariant()}";
        TxtInspectorBadge.Foreground = new WMedia.SolidColorBrush(ParseColor(node.StatusColorHex, WMedia.Color.FromRgb(56, 189, 248)));
        TxtInspectorHost.Text = node.TitleDisplay;
        TxtInspectorIp.Text = node.SubtitleDisplay;
        TxtInspectorKind.Text = node.KindLabel;
        TxtInspectorLatency.Text = node.TtlRemaining.HasValue
            ? $"{node.LatencyDisplay} (TTL retornado: {node.TtlRemaining})"
            : node.LatencyDisplay;
        TxtInspectorMac.Text = string.IsNullOrWhiteSpace(node.MacAddress) ? "-" : node.MacAddress;
        TxtInspectorUser.Text = string.IsNullOrWhiteSpace(node.ConnectedUser) ? "-" : node.ConnectedUser;
        TxtInspectorLocation.Text = node.LocationDisplay;
        TxtInspectorCidr.Text = string.IsNullOrWhiteSpace(node.MatchedCidr) ? "Nenhuma regra CIDR correspondente" : node.MatchedCidr;
        TxtInspectorDetails.Text = string.IsNullOrWhiteSpace(node.Details) ? "-" : node.Details;
    }

    private void ClearInspector()
    {
        TxtInspectorBadge.Text = "RASTREANDO...";
        TxtInspectorHost.Text = "Aguardando saltos...";
        TxtInspectorIp.Text = "-";
        TxtInspectorKind.Text = "-";
        TxtInspectorLatency.Text = "-";
        TxtInspectorMac.Text = "-";
        TxtInspectorUser.Text = "-";
        TxtInspectorLocation.Text = "-";
        TxtInspectorCidr.Text = "-";
        TxtInspectorDetails.Text = "-";
    }

    private void SetTraceStatus(string text, string bgHex, string fgHex, string borderHex)
    {
        TxtTraceStatus.Text = text;
        TxtTraceStatus.Foreground = new WMedia.SolidColorBrush(ParseColor(fgHex, WMedia.Colors.White));
        BadgeTraceStatus.Background = new WMedia.SolidColorBrush(ParseColor(bgHex, WMedia.Color.FromRgb(30, 41, 59)));
        BadgeTraceStatus.BorderBrush = new WMedia.SolidColorBrush(ParseColor(borderHex, WMedia.Color.FromRgb(51, 65, 85)));
    }

    private void ApplyZoom(double newScale)
    {
        _zoomScale = Math.Clamp(newScale, 0.45, 2.0);
        GraphScaleTransform.ScaleX = _zoomScale;
        GraphScaleTransform.ScaleY = _zoomScale;
        TxtZoomPercent.Text = $"{Math.Round(_zoomScale * 100)}%";
    }

    private void BtnZoomIn_Click(object sender, RoutedEventArgs e) => ApplyZoom(_zoomScale + 0.15);
    private void BtnZoomOut_Click(object sender, RoutedEventArgs e) => ApplyZoom(_zoomScale - 0.15);

    private void BtnZoomFit_Click(object sender, RoutedEventArgs e)
    {
        if (GraphCanvas.Width <= 0 || GraphScrollViewer.ViewportWidth <= 0)
        {
            ApplyZoom(1.0);
            return;
        }

        double ratioX = (GraphScrollViewer.ViewportWidth - 20) / GraphCanvas.Width;
        double ratioY = (GraphScrollViewer.ViewportHeight - 20) / GraphCanvas.Height;
        ApplyZoom(Math.Min(1.0, Math.Min(ratioX, ratioY)));
    }

    private void GraphScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            e.Handled = true;
            ApplyZoom(_zoomScale + (e.Delta > 0 ? 0.1 : -0.1));
        }
    }

    private void GraphScrollViewer_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isPanning = true;
        _panStartPoint = e.GetPosition(GraphScrollViewer);
        _panStartHorizontalOffset = GraphScrollViewer.HorizontalOffset;
        _panStartVerticalOffset = GraphScrollViewer.VerticalOffset;
        GraphScrollViewer.Cursor = Cursors.SizeAll;
    }

    private void GraphScrollViewer_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (!_isPanning || e.LeftButton != MouseButtonState.Pressed) return;
        Point current = e.GetPosition(GraphScrollViewer);
        GraphScrollViewer.ScrollToHorizontalOffset(_panStartHorizontalOffset - (current.X - _panStartPoint.X));
        GraphScrollViewer.ScrollToVerticalOffset(_panStartVerticalOffset - (current.Y - _panStartPoint.Y));
    }

    private void GraphScrollViewer_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        _isPanning = false;
        GraphScrollViewer.Cursor = Cursors.Arrow;
    }

    #endregion

    #region Editor Modal de Mapa de Sub-redes (CIDR -> Localização / Switch)

    private void ShowSubnetMapEditorDialog()
    {
        var owner = Window.GetWindow(this);
        var win = new Window
        {
            Title = "Mapa Local de Sub-redes (CIDR → Localização / Switch)",
            Width = 760,
            Height = 520,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = owner,
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(15, 23, 42)),
            Foreground = WMedia.Brushes.White
        };

        var root = new WControls.Grid { Margin = new Thickness(18) };
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });

        var header = new WControls.StackPanel { Margin = new Thickness(0, 0, 0, 10) };
        header.Children.Add(new WControls.TextBlock
        {
            Text = "Mapeamento Local de Sub-redes (subnet_map.json)",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold
        });
        header.Children.Add(new WControls.TextBlock
        {
            Text = "Informe uma regra por linha no formato: CIDR ; Localização / Prédio ; Switch Provável\nExemplo: 10.20.10.0/24 ; Sede - Bloco A (2º Andar) ; SW-BLOA-02",
            FontSize = 12,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(148, 163, 184)),
            Margin = new Thickness(0, 4, 0, 0),
            TextWrapping = TextWrapping.Wrap
        });
        WControls.Grid.SetRow(header, 0);
        root.Children.Add(header);

        var sb = new StringBuilder();
        foreach (var r in _subnetRules)
            sb.AppendLine($"{r.Cidr} ; {r.LocationName} ; {r.ProbableSwitch}");

        var editor = new WControls.TextBox
        {
            Text = sb.ToString(),
            AcceptsReturn = true,
            TextWrapping = TextWrapping.NoWrap,
            VerticalScrollBarVisibility = WControls.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = WControls.ScrollBarVisibility.Auto,
            FontFamily = new WMedia.FontFamily("Consolas, Cascadia Code, Courier New"),
            FontSize = 12.5,
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(11, 15, 25)),
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(248, 250, 252)),
            BorderBrush = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(51, 65, 85)),
            Padding = new Thickness(10)
        };
        WControls.Grid.SetRow(editor, 1);
        root.Children.Add(editor);

        var footer = new WControls.Grid { Margin = new Thickness(0, 12, 0, 0) };
        footer.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        footer.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        var lblInfo = new WControls.TextBlock
        {
            Text = $"Arquivo: {NetworkPathService.GetDefaultSubnetMapPath()}",
            FontSize = 11,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(148, 163, 184)),
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(0, 0, 12, 0)
        };
        WControls.Grid.SetColumn(lblInfo, 0);
        footer.Children.Add(lblInfo);

        var btnPanel = new WControls.StackPanel { Orientation = WControls.Orientation.Horizontal };
        var btnCancel = new Wpf.Ui.Controls.Button
        {
            Content = "Cancelar",
            Appearance = ControlAppearance.Secondary,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(14, 7, 14, 7)
        };
        btnCancel.Click += (_, _) => win.Close();

        var btnSave = new Wpf.Ui.Controls.Button
        {
            Content = "Salvar Regras",
            Appearance = ControlAppearance.Primary,
            Padding = new Thickness(14, 7, 14, 7)
        };
        btnSave.Click += (_, _) =>
        {
            var parsed = new List<SubnetLocationRule>();
            var invalidLines = new List<int>();
            var lines = editor.Text.Split(["\r\n", "\n"], StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith('#'))
                    continue;

                var parts = line.Split(';');
                string cidr = parts[0].Trim();
                if (!SubnetLocationRule.TryParseCidr(cidr, out _, out _, out _))
                {
                    invalidLines.Add(i + 1);
                    continue;
                }

                parsed.Add(new SubnetLocationRule
                {
                    Cidr = cidr,
                    LocationName = parts.Length > 1 ? parts[1].Trim() : string.Empty,
                    ProbableSwitch = parts.Length > 2 ? parts[2].Trim() : string.Empty
                });
            }

            if (invalidLines.Count > 0)
            {
                lblInfo.Text = $"⚠️ CIDR inválido na(s) linha(s): {string.Join(", ", invalidLines)}";
                lblInfo.Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(251, 191, 36));
                return;
            }

            NetworkPathService.SaveSubnetRules(parsed);
            _subnetRules = NetworkPathService.LoadSubnetRules();
            SetTraceStatus($"✅ {_subnetRules.Count} regra(s) CIDR salvas em subnet_map.json", "#064e3b", "#34d399", "#10b981");
            win.Close();
        };

        btnPanel.Children.Add(btnCancel);
        btnPanel.Children.Add(btnSave);
        WControls.Grid.SetColumn(btnPanel, 1);
        footer.Children.Add(btnPanel);

        WControls.Grid.SetRow(footer, 2);
        root.Children.Add(footer);

        win.Content = root;
        win.ShowDialog();
    }

    #endregion

    private static WMedia.Color ParseColor(string hex, WMedia.Color fallback)
    {
        try
        {
            return (WMedia.Color)WMedia.ColorConverter.ConvertFromString(hex);
        }
        catch
        {
            return fallback;
        }
    }
}
