using System.Collections.ObjectModel;
using System.Windows;
using AmbientesHub.Services;
using WControls = System.Windows.Controls;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Views;

public partial class ZabbixMonitorView : WControls.UserControl
{
    private readonly ZabbixEnrichmentService _zabbixService = new();
    private readonly ObservableCollection<ZabbixMonitoredHost> _filteredHosts = [];
    private List<ZabbixMonitoredHost> _allHosts = [];
    private List<ZabbixMonitoredHost> _filteredAll = [];
    private int _currentPage = 1;
    private const int PageSize = 25;
    private bool _isUpdatingFilters;
    private bool _loadedOnce;

    public event Action<string>? OnTraceRouteRequested;

    public ZabbixMonitorView()
    {
        InitializeComponent();
        ItemsHosts.ItemsSource = _filteredHosts;
    }

    public async Task EnsureLoadedOnceAsync()
    {
        if (_loadedOnce) return;
        _loadedOnce = true;
        await RefreshHostsAsync();
    }

    private async void BtnRefresh_Click(object sender, RoutedEventArgs e)
    {
        await RefreshHostsAsync();
    }

    private async Task RefreshHostsAsync()
    {
        if (!ZabbixEnrichmentService.IsConfigured())
        {
            SetStatus("⚠️ Configure ZABBIX_API_URL e ZABBIX_API_TOKEN", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        BtnRefresh.IsEnabled = false;
        SetStatus("🔄 Carregando hosts do Zabbix...", "#1e3a8a", "#93c5fd", "#3b82f6");

        try
        {
            _allHosts = await _zabbixService.GetMonitoredHostsAsync();
            PopulateGroupFilter();
            ApplyFilters();

            int online = _allHosts.Count(h => h.Availability == ZabbixHostAvailability.Online);
            int offline = _allHosts.Count(h => h.Availability == ZabbixHostAvailability.Offline);
            SetStatus($"🟢 {online} Online  •  🔴 {offline} Offline", "#064e3b", "#34d399", "#10b981");
        }
        finally
        {
            BtnRefresh.IsEnabled = true;
        }
    }

    private void PopulateGroupFilter()
    {
        _isUpdatingFilters = true;
        try
        {
            string? previous = CmbGroupFilter.SelectedItem as string;
            var groups = _allHosts
                .Select(h => h.GroupName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g)
                .ToList();

            groups.Insert(0, "Todos os Grupos");
            CmbGroupFilter.ItemsSource = groups;
            CmbGroupFilter.SelectedItem = (!string.IsNullOrEmpty(previous) && groups.Contains(previous)) ? previous : "Todos os Grupos";
        }
        finally
        {
            _isUpdatingFilters = false;
        }
    }

    private void Filter_Changed(object sender, RoutedEventArgs e)
    {
        if (!_isUpdatingFilters && IsLoaded)
            ApplyFilters();
    }

    private void ApplyFilters()
    {
        string search = TxtSearch.Text.Trim();
        string selectedGroup = CmbGroupFilter.SelectedItem as string ?? "Todos os Grupos";
        string statusTag = (CmbStatusFilter.SelectedItem as WControls.ComboBoxItem)?.Tag?.ToString() ?? "ALL";

        IEnumerable<ZabbixMonitoredHost> query = _allHosts;

        if (!string.Equals(selectedGroup, "Todos os Grupos", StringComparison.OrdinalIgnoreCase))
            query = query.Where(h => string.Equals(h.GroupName, selectedGroup, StringComparison.OrdinalIgnoreCase));

        if (statusTag == "OFFLINE")
            query = query.Where(h => h.Availability == ZabbixHostAvailability.Offline);
        else if (statusTag == "ONLINE")
            query = query.Where(h => h.Availability == ZabbixHostAvailability.Online);

        if (!string.IsNullOrWhiteSpace(search))
        {
            query = query.Where(h =>
                h.DisplayName.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                h.IpAddress.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                h.GroupName.Contains(search, StringComparison.OrdinalIgnoreCase));
        }

        _filteredAll = query.ToList();
        _currentPage = 1;
        RenderCurrentPage();
    }

    private void RenderCurrentPage()
    {
        int totalItems = _filteredAll.Count;
        int totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));
        _currentPage = Math.Clamp(_currentPage, 1, totalPages);

        _filteredHosts.Clear();
        foreach (var item in _filteredAll.Skip((_currentPage - 1) * PageSize).Take(PageSize))
            _filteredHosts.Add(item);

        TxtKpiTotal.Text = _allHosts.Count.ToString();
        TxtKpiOnline.Text = _allHosts.Count(h => h.Availability == ZabbixHostAvailability.Online).ToString();
        TxtKpiOffline.Text = _allHosts.Count(h => h.Availability == ZabbixHostAvailability.Offline).ToString();
        TxtKpiFiltered.Text = totalItems.ToString();

        TxtPageInfo.Text = $"Página {_currentPage} de {totalPages} ({totalItems} equipamento(s))";
        BtnPrevPage.IsEnabled = _currentPage > 1;
        BtnNextPage.IsEnabled = _currentPage < totalPages;
    }

    private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_currentPage > 1)
        {
            _currentPage--;
            RenderCurrentPage();
        }
    }

    private void BtnNextPage_Click(object sender, RoutedEventArgs e)
    {
        int totalPages = Math.Max(1, (int)Math.Ceiling((double)_filteredAll.Count / PageSize));
        if (_currentPage < totalPages)
        {
            _currentPage++;
            RenderCurrentPage();
        }
    }

    private async void BtnHostDetails_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement fe || fe.Tag is not ZabbixMonitoredHost host)
            return;

        SetStatus($"🔄 Carregando detalhes de {host.DisplayName}...", "#1e3a8a", "#93c5fd", "#3b82f6");
        var details = await _zabbixService.GetHostFullDetailsAsync(host.HostId);

        int online = _allHosts.Count(h => h.Availability == ZabbixHostAvailability.Online);
        int offline = _allHosts.Count(h => h.Availability == ZabbixHostAvailability.Offline);
        SetStatus($"🟢 {online} Online  •  🔴 {offline} Offline", "#064e3b", "#34d399", "#10b981");

        if (details == null)
            return;

        var win = new Window
        {
            Title = $"Detalhes Zabbix — {details.DisplayName}",
            Width = 820,
            Height = 600,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Window.GetWindow(this),
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(15, 23, 42)),
            Foreground = WMedia.Brushes.White
        };

        var root = new WControls.Grid { Margin = new Thickness(22) };
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });

        var header = new WControls.StackPanel { Margin = new Thickness(0, 0, 0, 14) };
        header.Children.Add(new WControls.TextBlock { Text = details.DisplayName, FontSize = 20, FontWeight = FontWeights.SemiBold });
        header.Children.Add(new WControls.TextBlock
        {
            Text = $"IP: {details.IpAddress}   •   Grupos: {string.Join(", ", details.Groups)}",
            FontSize = 13.5,
            Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(56, 189, 248)),
            Margin = new Thickness(0, 4, 0, 0),
            TextWrapping = TextWrapping.Wrap
        });
        WControls.Grid.SetRow(header, 0);
        root.Children.Add(header);

        var scroll = new WControls.ScrollViewer { VerticalScrollBarVisibility = WControls.ScrollBarVisibility.Auto };
        var stack = new WControls.StackPanel();

        if (details.ActiveAlerts.Count > 0)
        {
            stack.Children.Add(new WControls.TextBlock { Text = $"🚨 Alertas Ativos ({details.ActiveAlerts.Count})", FontSize = 15, FontWeight = FontWeights.Bold, Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(248, 113, 113)), Margin = new Thickness(0, 0, 0, 8) });
            foreach (var a in details.ActiveAlerts)
            {
                var alertCard = new WControls.Border
                {
                    Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(11, 15, 25)),
                    BorderBrush = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(a.PriorityColorHex)),
                    BorderThickness = new Thickness(1.2),
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(12, 8, 12, 8),
                    Margin = new Thickness(0, 0, 0, 6)
                };
                alertCard.Child = new WControls.TextBlock
                {
                    Text = $"[{a.PriorityLabel}]  {a.Description}  ({a.LastChangeBr})",
                    FontSize = 13.5,
                    TextWrapping = TextWrapping.Wrap
                };
                stack.Children.Add(alertCard);
            }
        }

        if (details.Inventory.Count > 0)
        {
            stack.Children.Add(new WControls.TextBlock { Text = "📋 Inventário", FontSize = 15, FontWeight = FontWeights.Bold, Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(56, 189, 248)), Margin = new Thickness(0, 14, 0, 8) });
            stack.Children.Add(BuildKeyValueTable(details.Inventory.Select(kv => (kv.Key, kv.Value))));
        }

        if (details.Metrics.Count > 0)
        {
            stack.Children.Add(new WControls.TextBlock { Text = "📊 Métricas Recentes", FontSize = 15, FontWeight = FontWeights.Bold, Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(52, 211, 153)), Margin = new Thickness(0, 14, 0, 8) });
            stack.Children.Add(BuildKeyValueTable(details.Metrics.Select(m => (m.Name, $"{m.Value} {m.Units}".Trim()))));
        }

        scroll.Content = stack;
        WControls.Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        var btnClose = new Wpf.Ui.Controls.Button { Content = "Fechar", Appearance = Wpf.Ui.Controls.ControlAppearance.Primary, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0), Padding = new Thickness(14, 6, 14, 6) };
        btnClose.Click += (_, _) => win.Close();
        WControls.Grid.SetRow(btnClose, 2);
        root.Children.Add(btnClose);

        win.Content = root;
        win.ShowDialog();
    }

    private static WControls.Border BuildKeyValueTable(IEnumerable<(string Key, string Value)> rows)
    {
        var tablePanel = new WControls.StackPanel();
        int idx = 0;
        foreach (var (key, value) in rows)
        {
            var rowBg = idx % 2 == 0
                ? WMedia.Color.FromRgb(11, 15, 25)
                : WMedia.Color.FromRgb(19, 28, 46);

            var rowGrid = new WControls.Grid();
            rowGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1.3, GridUnitType.Star) });
            rowGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var lblKey = new WControls.TextBlock
            {
                Text = key,
                FontSize = 13.5,
                Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(203, 213, 225)),
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 12, 0)
            };
            WControls.Grid.SetColumn(lblKey, 0);
            rowGrid.Children.Add(lblKey);

            var lblVal = new WControls.TextBlock
            {
                Text = value,
                FontSize = 13.5,
                FontWeight = FontWeights.SemiBold,
                FontFamily = new WMedia.FontFamily("Consolas, Segoe UI"),
                Foreground = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(248, 250, 252)),
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            WControls.Grid.SetColumn(lblVal, 1);
            rowGrid.Children.Add(lblVal);

            tablePanel.Children.Add(new WControls.Border
            {
                Background = new WMedia.SolidColorBrush(rowBg),
                Padding = new Thickness(12, 8, 12, 8),
                Child = rowGrid
            });
            idx++;
        }

        return new WControls.Border
        {
            BorderBrush = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            ClipToBounds = true,
            Child = tablePanel
        };
    }

    private void BtnTraceHost_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is ZabbixMonitoredHost host && !string.IsNullOrWhiteSpace(host.TargetForTrace))
        {
            OnTraceRouteRequested?.Invoke(host.TargetForTrace);
        }
    }

    private void SetStatus(string text, string bgHex, string fgHex, string borderHex)
    {
        TxtZabbixStatus.Text = text;
        TxtZabbixStatus.Foreground = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(fgHex));
        BadgeZabbixStatus.Background = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(bgHex));
        BadgeZabbixStatus.BorderBrush = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(borderHex));
    }
}
