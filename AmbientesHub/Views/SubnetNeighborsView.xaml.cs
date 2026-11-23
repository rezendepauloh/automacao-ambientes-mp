using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Effects;
using AmbientesHub.Models;
using AmbientesHub.Services;
using WControls = System.Windows.Controls;
using WMedia = System.Windows.Media;
using WShapes = System.Windows.Shapes;

namespace AmbientesHub.Views;

public partial class SubnetNeighborsView : WControls.UserControl
{
    private readonly SubnetNeighborsService _neighborsService = new();
    private readonly NetworkPathService _networkService = new();
    private readonly RemoteToolsService _remoteToolsService = new();

    private SubnetNeighborsResult? _lastResult;
    private SubnetNeighborNode? _selectedNode;
    private CancellationTokenSource? _scanCts;
    private bool _isScanning;
    private bool _isPanning;
    private Point _panStartPoint;
    private double _panStartH;
    private double _panStartV;
    private double _zoomScale = 1.0;
    private bool _isFullWindow;

    public event Action<string>? OnTraceRouteRequested;

    public SubnetNeighborsView()
    {
        InitializeComponent();
    }

    private void TxtRefTarget_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && !_isScanning)
            BtnScanNeighbors_Click(this, new RoutedEventArgs());
    }

    private void TxtFilterNeighbors_TextChanged(object sender, WControls.TextChangedEventArgs e)
    {
        if (IsLoaded && _lastResult != null)
            RenderGraph();
    }

    private async void BtnScanNeighbors_Click(object sender, RoutedEventArgs e)
    {
        string validated;
        try
        {
            validated = NetworkPathService.ValidateTarget(TxtRefTarget.Text);
        }
        catch (ArgumentException ex)
        {
            SetStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        _scanCts?.Cancel();
        _scanCts = new CancellationTokenSource();
        var ct = _scanCts.Token;

        _isScanning = true;
        _selectedNode = null;
        BtnScanNeighbors.IsEnabled = false;
        BtnStopScan.IsEnabled = true;
        ClearInspector();
        SetStatus($"🔄 Mapeando vizinhos de {validated}...", "#1e3a8a", "#93c5fd", "#3b82f6");

        bool includeArp = ChkIncludeRemoteArp.IsChecked == true;
        bool includeOu = ChkIncludeSameOu.IsChecked == true;

        var progress = new Progress<SubnetNeighborsResult>(partial =>
        {
            _lastResult = partial;
            RenderGraph();
        });

        try
        {
            var result = await _neighborsService.DiscoverNeighborsAsync(
                validated,
                includeRemoteArp: includeArp,
                includeSameAdOu: includeOu,
                onProgress: progress,
                cancellationToken: ct);

            _lastResult = result;
            _selectedNode = result.Neighbors.FirstOrDefault();
            if (_selectedNode != null)
            {
                UpdateInspector(_selectedNode);
                _ = EnrichSelectedNeighborInBackgroundAsync(_selectedNode, ct);
            }
            RenderGraph();

            int onlineCount = result.Neighbors.Count(n => n.IsOnline);
            SetStatus(
                $"🟢 {onlineCount} online / {result.Neighbors.Count} total em {result.SubnetCidr}",
                "#064e3b", "#34d399", "#10b981");
        }
        catch (OperationCanceledException)
        {
            SetStatus("⏹️ Varredura interrompida", "#1e293b", "#cbd5e1", "#475569");
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Falha: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444");
        }
        finally
        {
            _isScanning = false;
            BtnScanNeighbors.IsEnabled = true;
            BtnStopScan.IsEnabled = false;
        }
    }

    private void BtnStopScan_Click(object sender, RoutedEventArgs e)
    {
        _scanCts?.Cancel();
    }

    private void BtnCopyNeighbors_Click(object sender, RoutedEventArgs e)
    {
        if (_lastResult == null || _lastResult.Neighbors.Count == 0)
            return;

        var sb = new StringBuilder();
        sb.AppendLine($"=== VIZINHOS DE REDE — {_lastResult.ReferenceInput} ({_lastResult.SubnetCidr}) ===");
        sb.AppendLine($"Data: {_lastResult.ScannedAt:dd/MM/yyyy HH:mm:ss}");
        if (!string.IsNullOrWhiteSpace(_lastResult.SwitchOrZabbixLabel))
            sb.AppendLine($"Switch/Gateway: {_lastResult.SwitchOrZabbixLabel}");
        sb.AppendLine();

        foreach (var n in GetFilteredNeighbors())
        {
            sb.AppendLine($"{(n.IsOnline ? "ONLINE " : "OFFLINE"),-8} | {n.IpAddress,-15} | {n.MacAddress,-17} | {n.TitleDisplay,-30} | {n.ConnectedUser}");
        }

        if (ClipboardHelper.Copy(sb.ToString()))
            SetStatus("📋 Lista de vizinhos copiada!", "#064e3b", "#34d399", "#10b981");
    }

    private List<SubnetNeighborNode> GetFilteredNeighbors()
    {
        if (_lastResult == null) return [];
        string filter = TxtFilterNeighbors.Text.Trim();
        IEnumerable<SubnetNeighborNode> query = _lastResult.Neighbors;

        if (!string.IsNullOrWhiteSpace(filter))
        {
            query = query.Where(n =>
                n.Hostname.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                n.IpAddress.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                n.MacAddress.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                n.ConnectedUser.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                n.LocationOrOu.Contains(filter, StringComparison.OrdinalIgnoreCase));
        }

        // Limita a 60 nós renderizados simultaneamente no Canvas para manter 60 FPS (use o filtro para refinar)
        return query.Take(60).ToList();
    }

    #region Renderização da Árvore Estrela Bézier (Raiz -> Vizinhos)

    private void RenderGraph()
    {
        GraphCanvas.Children.Clear();
        var neighbors = GetFilteredNeighbors();
        GraphEmptyState.Visibility = (_lastResult == null) ? Visibility.Visible : Visibility.Collapsed;

        if (_lastResult == null)
        {
            GraphCanvas.Width = 800;
            GraphCanvas.Height = 420;
            return;
        }

        const double rootW = 248;
        const double rootH = 128;
        const double nodeW = 240;
        const double nodeH = 92;
        const double colGap = 100;
        const double cardGap = 12;
        const double groupPad = 14;
        const double groupHeaderH = 28;
        const double groupGap = 18;
        const double margin = 36;

        // Agrupa os vizinhos por Localização / Setor / OU
        var groups = neighbors
            .GroupBy(n => string.IsNullOrWhiteSpace(n.LocationOrOu) ? "Sem Setor / Aguardando Consulta" : n.LocationOrOu.Trim())
            .OrderByDescending(g => g.Any(n => n.IsReferenceHost))
            .ThenBy(g => g.Key.StartsWith("Sem Setor", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
            .ThenBy(g => g.Key)
            .ToList();

        double currentY = margin;
        double groupX = margin + rootW + colGap;
        double maxGroupWidth = 0;

        var groupLayouts = new List<(string Title, List<SubnetNeighborNode> Items, double X, double Y, double W, double H, int Cols)>();

        foreach (var grp in groups)
        {
            var items = grp.ToList();
            int cols = items.Count > 4 ? 2 : 1;
            int rows = (int)Math.Ceiling((double)items.Count / cols);
            double gw = groupPad * 2 + cols * nodeW + (cols - 1) * cardGap;
            double gh = groupPad + groupHeaderH + rows * nodeH + Math.Max(0, rows - 1) * cardGap + groupPad;

            groupLayouts.Add((grp.Key, items, groupX, currentY, gw, gh, cols));
            maxGroupWidth = Math.Max(maxGroupWidth, gw);
            currentY += gh + groupGap;
        }

        double totalChildrenH = Math.Max(rootH, currentY - margin - groupGap);
        double rootX = margin;
        double rootY = margin + Math.Max(0, (totalChildrenH - rootH) / 2);

        GraphCanvas.Width = Math.Max(820, groupX + maxGroupWidth + margin);
        GraphCanvas.Height = Math.Max(420, margin + totalChildrenH + margin);

        var rootPos = new Point(rootX, rootY);
        var rootStart = new Point(rootPos.X + rootW, rootPos.Y + rootH / 2);

        // 1. Desenha curvas Bézier da Raiz para cada retângulo de Setor/OU
        foreach (var g in groupLayouts)
        {
            var end = new Point(g.X, g.Y + g.H / 2);
            bool hasRef = g.Items.Any(i => i.IsReferenceHost);
            bool hasOnline = g.Items.Any(i => i.IsOnline);
            DrawBezierLink(rootStart, end, new SubnetNeighborNode { IsReferenceHost = hasRef, IsOnline = hasOnline });
        }

        // 2. Nó Raiz (Sub-rede / Gateway / Switch)
        var rootCard = CreateRootCard(_lastResult, rootW, rootH);
        WControls.Canvas.SetLeft(rootCard, rootPos.X);
        WControls.Canvas.SetTop(rootCard, rootPos.Y);
        GraphCanvas.Children.Add(rootCard);

        // 3. Retângulos de Setor/OU + Cards internos
        foreach (var g in groupLayouts)
        {
            var groupBox = new WControls.Border
            {
                Width = g.W,
                Height = g.H,
                CornerRadius = new CornerRadius(10),
                Background = new WMedia.SolidColorBrush(WMedia.Color.FromArgb(90, 15, 23, 42)),
                BorderBrush = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1.2),
                Padding = new Thickness(groupPad, 10, groupPad, groupPad)
            };

            var titleTb = new WControls.TextBlock
            {
                Text = $"📁 {g.Title} ({g.Items.Count})",
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(56, 189, 248)),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            groupBox.Child = titleTb;
            WControls.Canvas.SetLeft(groupBox, g.X);
            WControls.Canvas.SetTop(groupBox, g.Y);
            GraphCanvas.Children.Add(groupBox);

            for (int i = 0; i < g.Items.Count; i++)
            {
                int r = i / g.Cols;
                int c = i % g.Cols;
                double cx = g.X + groupPad + c * (nodeW + cardGap);
                double cy = g.Y + groupPad + groupHeaderH + r * (nodeH + cardGap);

                var card = CreateNeighborCard(g.Items[i], nodeW, nodeH);
                WControls.Canvas.SetLeft(card, cx);
                WControls.Canvas.SetTop(card, cy);
                GraphCanvas.Children.Add(card);
            }
        }
    }

    private void DrawBezierLink(Point start, Point end, SubnetNeighborNode target)
    {
        var color = target.IsReferenceHost
            ? WMedia.Color.FromRgb(56, 189, 248)
            : (target.IsOnline ? WMedia.Color.FromRgb(16, 185, 129) : WMedia.Color.FromRgb(100, 116, 139));
        var brush = new WMedia.SolidColorBrush(color);

        double dx = Math.Max(40, (end.X - start.X) * 0.55);
        var fig = new WMedia.PathFigure { StartPoint = start, IsClosed = false };
        fig.Segments.Add(new WMedia.BezierSegment(
            new Point(start.X + dx, start.Y),
            new Point(end.X - dx, end.Y),
            end,
            isStroked: true));

        var geom = new WMedia.PathGeometry();
        geom.Figures.Add(fig);

        var path = new WShapes.Path
        {
            Data = geom,
            Stroke = brush,
            StrokeThickness = target.IsReferenceHost ? 2.6 : 1.8,
            Opacity = target.IsOnline ? 0.9 : 0.55
        };
        if (!target.IsOnline)
            path.StrokeDashArray = new WMedia.DoubleCollection { 5, 4 };

        GraphCanvas.Children.Add(path);
    }

    private WControls.Border CreateRootCard(SubnetNeighborsResult res, double width, double height)
    {
        var card = new WControls.Border
        {
            Width = width,
            Height = height,
            CornerRadius = new CornerRadius(10),
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(15, 23, 42)),
            BorderBrush = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(56, 189, 248)),
            BorderThickness = new Thickness(2),
            Padding = new Thickness(12, 10, 12, 10),
            Effect = new DropShadowEffect { Color = WMedia.Color.FromRgb(56, 189, 248), BlurRadius = 14, ShadowDepth = 2, Opacity = 0.35 }
        };

        var sp = new WControls.StackPanel { VerticalAlignment = VerticalAlignment.Center };
        sp.Children.Add(new WControls.TextBlock
        {
            Text = "🌐 SUB-REDE / GATEWAY",
            FontSize = 10,
            FontWeight = FontWeights.Bold,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(56, 189, 248))
        });
        sp.Children.Add(new WControls.TextBlock
        {
            Text = string.IsNullOrWhiteSpace(res.SubnetCidr) ? res.ReferenceInput : res.SubnetCidr,
            FontSize = 15,
            FontWeight = FontWeights.Bold,
            Foreground = WMedia.Brushes.White,
            Margin = new Thickness(0, 4, 0, 2)
        });
        sp.Children.Add(new WControls.TextBlock
        {
            Text = string.IsNullOrWhiteSpace(res.GatewayIp) ? "Sem IPv4 ativo" : $"Gateway: {res.GatewayIp}",
            FontSize = 11.5,
            FontFamily = new WMedia.FontFamily("Consolas"),
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(148, 163, 184))
        });

        string extra = !string.IsNullOrWhiteSpace(res.SwitchOrZabbixLabel) ? $"🔀 {res.SwitchOrZabbixLabel}" : res.AdOuPath;
        if (!string.IsNullOrWhiteSpace(extra))
        {
            sp.Children.Add(new WControls.TextBlock
            {
                Text = extra,
                FontSize = 11,
                Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(52, 211, 153)),
                Margin = new Thickness(0, 4, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = extra
            });
        }

        card.Child = sp;
        return card;
    }

    private WControls.Border CreateNeighborCard(SubnetNeighborNode node, double width, double height)
    {
        bool isSelected = ReferenceEquals(node, _selectedNode);
        var borderCol = isSelected
            ? WMedia.Color.FromRgb(56, 189, 248)
            : (WMedia.Color)WMedia.ColorConverter.ConvertFromString(node.BorderBrushHex);

        var card = new WControls.Border
        {
            Width = width,
            Height = height,
            CornerRadius = new CornerRadius(8),
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(15, 23, 42)),
            BorderBrush = new WMedia.SolidColorBrush(borderCol),
            BorderThickness = new Thickness(isSelected ? 2 : 1.2),
            Padding = new Thickness(10, 8, 10, 8),
            Cursor = Cursors.Hand
        };

        card.MouseLeftButtonDown += (_, args) =>
        {
            args.Handled = true;
            _selectedNode = node;
            UpdateInspector(node);
            RenderGraph();
            _ = EnrichSelectedNeighborInBackgroundAsync(node, CancellationToken.None);
        };

        var root = new WControls.Grid();
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });

        var top = new WControls.Grid();
        top.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        top.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        var badge = new WControls.Border
        {
            Background = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(node.BadgeBackgroundHex)),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(6, 1, 6, 1),
            HorizontalAlignment = HorizontalAlignment.Left
        };
        badge.Child = new WControls.TextBlock
        {
            Text = node.BadgeText,
            FontSize = 9.5,
            FontWeight = FontWeights.Bold,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(226, 232, 240))
        };
        top.Children.Add(badge);

        var lat = new WControls.TextBlock
        {
            Text = node.LatencyDisplay,
            FontSize = 10,
            FontWeight = FontWeights.Bold,
            Foreground = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(node.StatusColorHex)),
            VerticalAlignment = VerticalAlignment.Center
        };
        WControls.Grid.SetColumn(lat, 1);
        top.Children.Add(lat);
        root.Children.Add(top);

        var mid = new WControls.StackPanel { VerticalAlignment = VerticalAlignment.Center };
        mid.Children.Add(new WControls.TextBlock
        {
            Text = node.TitleDisplay,
            FontSize = 12.5,
            FontWeight = FontWeights.SemiBold,
            Foreground = WMedia.Brushes.White,
            TextTrimming = TextTrimming.CharacterEllipsis
        });
        mid.Children.Add(new WControls.TextBlock
        {
            Text = node.SubtitleDisplay,
            FontSize = 10.5,
            FontFamily = new WMedia.FontFamily("Consolas"),
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(148, 163, 184)),
            TextTrimming = TextTrimming.CharacterEllipsis
        });
        WControls.Grid.SetRow(mid, 1);
        root.Children.Add(mid);

        string bottomText = !string.IsNullOrWhiteSpace(node.ConnectedUser)
            ? $"👤 {node.ConnectedUser}"
            : (!string.IsNullOrWhiteSpace(node.LocationOrOu) ? $"📍 {node.LocationOrOu}" : "Clique para detalhes e usuário");
        var bot = new WControls.TextBlock
        {
            Text = bottomText,
            FontSize = 10,
            Foreground = new WMedia.SolidColorBrush(!string.IsNullOrWhiteSpace(node.ConnectedUser)
                ? WMedia.Color.FromRgb(52, 211, 153)
                : WMedia.Color.FromRgb(56, 189, 248)),
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        WControls.Grid.SetRow(bot, 2);
        root.Children.Add(bot);

        card.Child = root;
        return card;
    }

    #endregion

    #region Inspetor, Enriquecimento sob Demanda & Ações

    private async Task EnrichSelectedNeighborInBackgroundAsync(SubnetNeighborNode node, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(node.ConnectedUser) && !string.IsNullOrWhiteSpace(node.MacAddress))
            return;

        string target = !string.IsNullOrWhiteSpace(node.Hostname) ? node.Hostname : node.IpAddress;
        if (string.IsNullOrWhiteSpace(target))
            return;

        var tempHop = new NetworkHopNode
        {
            Hostname = node.Hostname,
            IpAddress = node.IpAddress,
            MacAddress = node.MacAddress,
            ConnectedUser = node.ConnectedUser
        };

        bool updated = await NetworkPathService.TryEnrichDestinationFromAdSccmAsync(target, tempHop, ct);
        if (updated)
        {
            node.ConnectedUser = tempHop.ConnectedUser;
            node.MacAddress = string.IsNullOrWhiteSpace(node.MacAddress) ? tempHop.MacAddress : node.MacAddress;
            node.LocationOrOu = string.IsNullOrWhiteSpace(node.LocationOrOu) ? tempHop.InferredLocation : node.LocationOrOu;

            if (ReferenceEquals(_selectedNode, node))
                UpdateInspector(node);
            RenderGraph();
        }
    }

    private void UpdateInspector(SubnetNeighborNode node)
    {
        TxtInspectorBadge.Text = $"{node.BadgeText} • {(node.IsOnline ? "ONLINE" : "SEM PING")}";
        TxtInspectorHost.Text = node.TitleDisplay;
        TxtInspectorIp.Text = string.IsNullOrWhiteSpace(node.IpAddress) ? "-" : node.IpAddress;
        TxtInspectorMac.Text = string.IsNullOrWhiteSpace(node.MacAddress) ? "-" : node.MacAddress;
        TxtInspectorUser.Text = string.IsNullOrWhiteSpace(node.ConnectedUser) ? "(consultando ou não informado)" : node.ConnectedUser;
        TxtInspectorLocation.Text = string.IsNullOrWhiteSpace(node.LocationOrOu) ? "-" : node.LocationOrOu;
        TxtInspectorLatency.Text = $"{node.LatencyDisplay} ({node.SourceTag})";

        bool canAct = !string.IsNullOrWhiteSpace(node.Hostname) || !string.IsNullOrWhiteSpace(node.IpAddress);
        BtnQuickSccm.IsEnabled = canAct;
        BtnQuickRdp.IsEnabled = canAct;
        BtnQuickExplorer.IsEnabled = canAct;
        BtnQuickTrace.IsEnabled = canAct;
        BtnTestPorts.IsEnabled = canAct;
        PortsResultPanel.Children.Clear();
    }

    private void ClearInspector()
    {
        TxtInspectorBadge.Text = "MAPEANDO...";
        TxtInspectorHost.Text = "Aguardando hosts...";
        TxtInspectorIp.Text = "-";
        TxtInspectorMac.Text = "-";
        TxtInspectorUser.Text = "-";
        TxtInspectorLocation.Text = "-";
        TxtInspectorLatency.Text = "-";
        BtnQuickSccm.IsEnabled = false;
        BtnQuickRdp.IsEnabled = false;
        BtnQuickExplorer.IsEnabled = false;
        BtnQuickTrace.IsEnabled = false;
        BtnTestPorts.IsEnabled = false;
        PortsResultPanel.Children.Clear();
    }

    private string GetSelectedTarget()
    {
        if (_selectedNode == null) return string.Empty;
        if (!string.IsNullOrWhiteSpace(_selectedNode.Hostname))
            return _selectedNode.Hostname.Split('.')[0].Trim();
        return _selectedNode.IpAddress.Trim();
    }

    private void BtnQuickSccm_Click(object sender, RoutedEventArgs e)
    {
        string target = GetSelectedTarget();
        if (string.IsNullOrWhiteSpace(target)) return;
        try { _remoteToolsService.OpenSccmRemoteControl(target, ConfigService.SccmSiteServer); }
        catch (Exception ex) { SetStatus($"❌ SCCM: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444"); }
    }

    private void BtnQuickRdp_Click(object sender, RoutedEventArgs e)
    {
        string target = GetSelectedTarget();
        if (string.IsNullOrWhiteSpace(target)) return;
        try { _remoteToolsService.OpenRdp(target); }
        catch (Exception ex) { SetStatus($"❌ RDP: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444"); }
    }

    private void BtnQuickExplorer_Click(object sender, RoutedEventArgs e)
    {
        string target = GetSelectedTarget();
        if (string.IsNullOrWhiteSpace(target)) return;
        try { _remoteToolsService.OpenRemoteExplorer(target); }
        catch (Exception ex) { SetStatus($"❌ C$: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444"); }
    }

    private void BtnQuickTrace_Click(object sender, RoutedEventArgs e)
    {
        string target = GetSelectedTarget();
        if (!string.IsNullOrWhiteSpace(target))
            OnTraceRouteRequested?.Invoke(target);
    }

    private async void BtnTestPorts_Click(object sender, RoutedEventArgs e)
    {
        string target = GetSelectedTarget();
        if (string.IsNullOrWhiteSpace(target)) return;

        BtnTestPorts.IsEnabled = false;
        PortsResultPanel.Children.Clear();
        try
        {
            var results = await _networkService.CheckCommonPortsAsync(target, timeoutMs: 1000);
            foreach (var r in results)
            {
                var chip = new WControls.Border
                {
                    Background = new WMedia.SolidColorBrush(r.IsOpen ? WMedia.Color.FromRgb(6, 78, 59) : WMedia.Color.FromRgb(69, 10, 10)),
                    BorderBrush = new WMedia.SolidColorBrush(r.IsOpen ? WMedia.Color.FromRgb(16, 185, 129) : WMedia.Color.FromRgb(185, 28, 28)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(7, 3, 7, 3),
                    Margin = new Thickness(0, 0, 5, 5)
                };
                chip.Child = new WControls.TextBlock
                {
                    Text = $"{(r.IsOpen ? "🟢" : "🔴")} {r.Port} {r.ServiceName}",
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new WMedia.SolidColorBrush(r.IsOpen ? WMedia.Color.FromRgb(52, 211, 153) : WMedia.Color.FromRgb(248, 113, 113))
                };
                PortsResultPanel.Children.Add(chip);
            }
        }
        finally
        {
            BtnTestPorts.IsEnabled = true;
        }
    }

    #endregion

    #region Zoom, Pan & Status

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
        if (GraphCanvas.Width <= 0 || GraphScrollViewer.ViewportWidth <= 0) { ApplyZoom(1.0); return; }
        double rx = (GraphScrollViewer.ViewportWidth - 20) / GraphCanvas.Width;
        double ry = (GraphScrollViewer.ViewportHeight - 20) / GraphCanvas.Height;
        ApplyZoom(Math.Min(1.0, Math.Min(rx, ry)));
    }

    private void BtnToggleFullWindow_Click(object sender, RoutedEventArgs e)
    {
        _isFullWindow = !_isFullWindow;
        HeaderRow.Visibility = _isFullWindow ? Visibility.Collapsed : Visibility.Visible;
        ControlsRow.Visibility = _isFullWindow ? Visibility.Collapsed : Visibility.Visible;
        ColInspector.Width = _isFullWindow ? new GridLength(0) : new GridLength(300);
        BtnToggleFullWindow.Content = _isFullWindow ? "⛶ Restaurar" : "⛶ Expandir";
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
        _panStartH = GraphScrollViewer.HorizontalOffset;
        _panStartV = GraphScrollViewer.VerticalOffset;
        GraphScrollViewer.Cursor = Cursors.SizeAll;
    }

    private void GraphScrollViewer_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (!_isPanning || e.LeftButton != MouseButtonState.Pressed) return;
        Point cur = e.GetPosition(GraphScrollViewer);
        GraphScrollViewer.ScrollToHorizontalOffset(_panStartH - (cur.X - _panStartPoint.X));
        GraphScrollViewer.ScrollToVerticalOffset(_panStartV - (cur.Y - _panStartPoint.Y));
    }

    private void GraphScrollViewer_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        _isPanning = false;
        GraphScrollViewer.Cursor = Cursors.Arrow;
    }

    private void SetStatus(string text, string bgHex, string fgHex, string borderHex)
    {
        TxtScanStatus.Text = text;
        TxtScanStatus.Foreground = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(fgHex));
        BadgeScanStatus.Background = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(bgHex));
        BadgeScanStatus.BorderBrush = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(borderHex));
    }

    #endregion
}
