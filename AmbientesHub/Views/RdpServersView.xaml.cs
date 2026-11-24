using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using AmbientesHub.Controls;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Microsoft.Win32;
using WControls = System.Windows.Controls;
using WFormsHost = System.Windows.Forms.Integration.WindowsFormsHost;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Views;

public partial class RdpServersView : WControls.UserControl
{
    private readonly ObservableCollection<RdpServerEntry> _pagedServers = [];
    private RdpCatalogDocument _catalog = new();
    private List<RdpServerEntry> _filteredAll = [];
    private RdpServerEntry? _selectedServer;
    private int _currentPage = 1;
    private const int PageSize = 15;
    private bool _isUpdatingFilters;

    public RdpServersView()
    {
        InitializeComponent();
        ItemsServers.ItemsSource = _pagedServers;
        LoadCatalogFromDisk();
    }

    private void LoadCatalogFromDisk()
    {
        _catalog = RdpCatalogService.LoadEncryptedCatalog();
        PopulateGroupFilter();
        ApplyFilters();
        if (_filteredAll.Count > 0)
            SelectServer(_filteredAll[0]);
    }

    private void PopulateGroupFilter()
    {
        _isUpdatingFilters = true;
        try
        {
            string? previous = CmbGroupFilter.SelectedItem as string;
            var groups = _catalog.Servers
                .Select(s => string.IsNullOrWhiteSpace(s.Group) ? "Geral" : s.Group)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g)
                .ToList();
            groups.Insert(0, "Todos os Grupos");
            CmbGroupFilter.ItemsSource = groups;
            CmbGroupFilter.SelectedItem = (!string.IsNullOrEmpty(previous) && groups.Contains(previous))
                ? previous
                : "Todos os Grupos";
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
        string group = CmbGroupFilter.SelectedItem as string ?? "Todos os Grupos";
        string seguraTag = (CmbSeguraFilter.SelectedItem as WControls.ComboBoxItem)?.Tag?.ToString() ?? "ALL";

        IEnumerable<RdpServerEntry> query = _catalog.Servers;

        if (!string.Equals(group, "Todos os Grupos", StringComparison.OrdinalIgnoreCase))
            query = query.Where(s => string.Equals(s.Group, group, StringComparison.OrdinalIgnoreCase));

        if (seguraTag == "SEGURA")
            query = query.Where(s => s.IsSeguraConfigured);
        else if (seguraTag == "DIRECT")
            query = query.Where(s => !s.IsSeguraConfigured);

        if (!string.IsNullOrWhiteSpace(search))
        {
            query = query.Where(s =>
                s.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                s.HostOrUrl.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                s.UserName.Contains(search, StringComparison.OrdinalIgnoreCase));
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

        _pagedServers.Clear();
        foreach (var item in _filteredAll.Skip((_currentPage - 1) * PageSize).Take(PageSize))
            _pagedServers.Add(item);

        int totalAll = _catalog.Servers.Count;
        int totalSegura = _catalog.Servers.Count(s => s.IsSeguraConfigured);
        SetStatus($"{totalAll} servidor(es) • {totalSegura} Segura®", "#1e293b", "#38bdf8", "#334155");

        TxtPageInfo.Text = $"Página {_currentPage} de {totalPages} ({totalItems} exibido(s))";
        BtnPrevPage.IsEnabled = _currentPage > 1;
        BtnNextPage.IsEnabled = _currentPage < totalPages;
    }

    private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
    {
        if (_currentPage > 1) { _currentPage--; RenderCurrentPage(); }
    }

    private void BtnNextPage_Click(object sender, RoutedEventArgs e)
    {
        int totalPages = Math.Max(1, (int)Math.Ceiling((double)_filteredAll.Count / PageSize));
        if (_currentPage < totalPages) { _currentPage++; RenderCurrentPage(); }
    }

    private void ServerCard_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is RdpServerEntry entry)
            SelectServer(entry);
    }

    private void SelectServer(RdpServerEntry entry)
    {
        _selectedServer = entry;
        TxtEditName.Text = entry.Name;
        TxtEditHost.Text = entry.HostOrUrl;
        TxtEditPort.Text = entry.Port.ToString();
        TxtEditUser.Text = entry.UserName;
        TxtEditDomain.Text = entry.Domain;
        TxtEditGroup.Text = entry.Group;
        ChkEditPrompt.IsChecked = entry.PromptCredentials;
        ChkEditDrives.IsChecked = entry.UsesHardDrives;
        ChkEditEmbedded.IsChecked = entry.OpenEmbedded;
        ChkEditAutoAdapt.IsChecked = entry.AutoAdaptResolution;
        ChkEditMultiMon.IsChecked = entry.UseMultiMonitorInFullscreen;
        CmbEditQualityProfile.SelectedIndex = Math.Clamp((int)entry.QualityProfile, 0, 2);
        CmbEditAuthLevel.SelectedIndex = Math.Clamp(entry.AuthenticationLevel, 0, 2);
    }

    private void BtnNewServer_Click(object sender, RoutedEventArgs e)
    {
        var novo = new RdpServerEntry
        {
            Name = "NOVO-SERVIDOR",
            HostOrUrl = "10.0.0.10",
            Group = "Servidores"
        };
        _catalog.Servers.Insert(0, novo);
        PopulateGroupFilter();
        ApplyFilters();
        SelectServer(novo);
    }

    private void BtnSaveServer_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedServer == null)
        {
            _selectedServer = new RdpServerEntry();
            _catalog.Servers.Insert(0, _selectedServer);
        }

        string host = TxtEditHost.Text.Trim();
        try
        {
            host = NetworkPathService.ValidateTarget(host);
        }
        catch (ArgumentException ex)
        {
            SetStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        int port = int.TryParse(TxtEditPort.Text.Trim(), out int p) ? Math.Clamp(p, 1, 65535) : 3389;

        _selectedServer.Name = string.IsNullOrWhiteSpace(TxtEditName.Text) ? host : TxtEditName.Text.Trim();
        _selectedServer.HostOrUrl = host;
        _selectedServer.Port = port;
        _selectedServer.UserName = TxtEditUser.Text.Trim();
        _selectedServer.Domain = TxtEditDomain.Text.Trim();
        _selectedServer.Group = string.IsNullOrWhiteSpace(TxtEditGroup.Text) ? "Geral" : TxtEditGroup.Text.Trim();
        _selectedServer.PromptCredentials = ChkEditPrompt.IsChecked == true;
        _selectedServer.UsesHardDrives = ChkEditDrives.IsChecked == true;
        _selectedServer.OpenEmbedded = ChkEditEmbedded.IsChecked == true;
        _selectedServer.AutoAdaptResolution = ChkEditAutoAdapt.IsChecked == true;
        _selectedServer.UseMultiMonitorInFullscreen = ChkEditMultiMon.IsChecked == true;
        _selectedServer.QualityProfile = (RdpQualityProfile)Math.Clamp(CmbEditQualityProfile.SelectedIndex, 0, 2);
        _selectedServer.AuthenticationLevel = CmbEditAuthLevel.SelectedIndex >= 0 ? CmbEditAuthLevel.SelectedIndex : 0;

        RdpCatalogService.SaveEncryptedCatalog(_catalog);
        PopulateGroupFilter();
        ApplyFilters();
        SetStatus($"💾 '{_selectedServer.Name}' salvo no cofre DPAPI!", "#064e3b", "#34d399", "#10b981");
    }

    private void BtnDeleteServer_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedServer == null) return;
        var confirm = System.Windows.MessageBox.Show(
            $"Remover '{_selectedServer.Name}' do catálogo local?",
            "Excluir Servidor",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes) return;

        _catalog.Servers.Remove(_selectedServer);
        _selectedServer = null;
        RdpCatalogService.SaveEncryptedCatalog(_catalog);
        PopulateGroupFilter();
        ApplyFilters();
        SetStatus("🗑️ Servidor removido do cofre.", "#1e293b", "#cbd5e1", "#475569");
    }

    private void BtnApplySegura_Click(object sender, RoutedEventArgs e)
    {
        var temp = _selectedServer ?? new RdpServerEntry();
        temp.Name = TxtEditName.Text.Trim();
        temp.HostOrUrl = TxtEditHost.Text.Trim();

        // Tenta inferir proxy, usuário e domínio a partir de servidores Segura já configurados no catálogo
        var existingSegura = _catalog.Servers.FirstOrDefault(s => s.IsSeguraConfigured && !string.IsNullOrWhiteSpace(s.UserName));
        string defaultProxy = existingSegura?.HostOrUrl ?? ConfigService.Get("SEGURA_PROXY_HOST", "segura.dominio.local");
        string defaultUser = Environment.UserName;
        string defaultDom = "dominio.local";
        string defaultAdm = ConfigService.AdminAdUser;
        if (string.IsNullOrWhiteSpace(defaultAdm))
            defaultAdm = $"{defaultUser}_adm";

        if (existingSegura != null && existingSegura.UserName.Contains('[') && existingSegura.UserName.Contains('\\'))
        {
            int b1 = existingSegura.UserName.IndexOf('[');
            int slash = existingSegura.UserName.IndexOf('\\', b1 + 1);
            int at = existingSegura.UserName.IndexOf('@', slash + 1);
            if (b1 > 0 && slash > b1 && at > slash)
            {
                defaultDom = existingSegura.UserName[(b1 + 1)..slash];
            }
        }

        RdpCatalogService.ApplySeguraTemplate(temp, defaultProxy, defaultUser, defaultDom, defaultAdm);
        TxtEditHost.Text = temp.HostOrUrl;
        TxtEditUser.Text = temp.UserName;
        TxtEditGroup.Text = temp.Group;
        ChkEditPrompt.IsChecked = true;
        SetStatus("🔐 Padrão Segura® aplicado! Clique em Salvar para gravar.", "#064e3b", "#34d399", "#10b981");
    }

    private void BtnImportRdmJson_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Importar JSON do Remote Desktop Manager",
            Filter = "Arquivos JSON (*.json)|*.json|Todos (*.*)|*.*"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            string json = File.ReadAllText(dlg.FileName);
            var imported = RdpCatalogService.ImportFromRdmJson(json);
            _catalog = imported;
            RdpCatalogService.SaveEncryptedCatalog(_catalog);
            PopulateGroupFilter();
            ApplyFilters();
            if (_filteredAll.Count > 0) SelectServer(_filteredAll[0]);
            SetStatus($"📥 {imported.Servers.Count} servidor(es) importado(s) e criptografados!", "#064e3b", "#34d399", "#10b981");
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Erro ao importar: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444");
        }
    }

    private void BtnExportRdmJson_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Title = "Exportar Catálogo RDP (Formato RDM JSON)",
            Filter = "Arquivos JSON (*.json)|*.json",
            FileName = $"RDP_Export_{DateTime.Now:yyyyMMdd}.json"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            string json = RdpCatalogService.ExportToRdmJson(_catalog);
            File.WriteAllText(dlg.FileName, json);
            SetStatus($"📤 JSON exportado com {_catalog.Servers.Count} servidor(es)!", "#064e3b", "#34d399", "#10b981");
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Erro ao exportar: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444");
        }
    }

    private void BtnCardConnectEmbedded_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is RdpServerEntry entry)
            OpenEmbeddedRdpTab(entry);
    }

    private void BtnCardConnectMstsc_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is RdpServerEntry entry)
            LaunchExternalRdp(entry);
    }

    private void BtnEditorConnectEmbedded_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedServer != null)
            OpenEmbeddedRdpTab(_selectedServer);
    }

    private void BtnEditorConnectMstsc_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedServer != null)
            LaunchExternalRdp(_selectedServer);
    }

    private void LaunchExternalRdp(RdpServerEntry entry)
    {
        try
        {
            RdpCatalogService.LaunchExternalMstsc(entry);
            SetStatus($"🪟 mstsc.exe iniciado para {entry.Name}", "#064e3b", "#34d399", "#10b981");
        }
        catch (Exception ex)
        {
            SetStatus($"❌ RDP: {ex.Message}", "#7f1d1d", "#f87171", "#ef4444");
        }
    }

    private void OpenEmbeddedRdpTab(RdpServerEntry entry)
    {
        try
        {
            NetworkPathService.ValidateTarget(entry.HostOrUrl);
        }
        catch (ArgumentException ex)
        {
            SetStatus($"⚠️ {ex.Message}", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        var rdpControl = new EmbeddedRdpHost();
        var wfHost = new WFormsHost { Child = rdpControl };

        var tabGrid = new WControls.Grid { Margin = new Thickness(0, 6, 0, 0) };
        tabGrid.RowDefinitions.Add(new WControls.RowDefinition { Height = GridLength.Auto });
        tabGrid.RowDefinitions.Add(new WControls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var toolbar = new WControls.Border
        {
            Background = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(15, 23, 42)),
            BorderBrush = new WMedia.SolidColorBrush(WMedia.Color.FromRgb(51, 65, 85)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(10, 6, 10, 6),
            Margin = new Thickness(0, 0, 0, 6)
        };

        var tbGrid = new WControls.Grid();
        tbGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        tbGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        tbGrid.Children.Add(new WControls.TextBlock
        {
            Text = $"🖥️ {entry.Name}  ({entry.SubtitleDisplay})",
            ToolTip = $"{entry.Name} • {entry.SubtitleDisplay} • {entry.UserDisplay}",
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
            Foreground = WMedia.Brushes.White,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(0, 0, 12, 0)
        });

        var actions = new WControls.WrapPanel
        {
            Orientation = WControls.Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center
        };

        var cmbProfile = new WControls.ComboBox
        {
            Width = 172,
            Height = 30,
            FontSize = 11.5,
            Margin = new Thickness(0, 0, 6, 0),
            ToolTip = "Perfil de qualidade visual e estabilidade de rede (clique em Reconectar para aplicar)"
        };
        cmbProfile.Items.Add("⚖️ Balanceado (24-bit)");
        cmbProfile.Items.Add("🚀 Estabilidade (16-bit)");
        cmbProfile.Items.Add("✨ Qualidade (32-bit)");
        cmbProfile.SelectedIndex = Math.Clamp((int)entry.QualityProfile, 0, 2);
        cmbProfile.SelectionChanged += (_, _) =>
        {
            if (cmbProfile.SelectedIndex >= 0)
            {
                entry.QualityProfile = (RdpQualityProfile)cmbProfile.SelectedIndex;
                RdpCatalogService.SaveEncryptedCatalog(_catalog);
                if (ReferenceEquals(_selectedServer, entry))
                    CmbEditQualityProfile.SelectedIndex = cmbProfile.SelectedIndex;
            }
        };

        var chkScale = new WControls.CheckBox
        {
            Content = "Auto-Ajuste",
            IsChecked = entry.AutoAdaptResolution,
            FontSize = 11.5,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(2, 0, 8, 0),
            ToolTip = "Adapta automaticamente a escala (SmartSizing) e a resolução ao redimensionar a janela/monitor"
        };
        chkScale.Click += (_, _) =>
        {
            bool enabled = chkScale.IsChecked == true;
            entry.AutoAdaptResolution = enabled;
            rdpControl.AutoAdaptResolution = enabled;
            rdpControl.SetSmartSizing(enabled);
            RdpCatalogService.SaveEncryptedCatalog(_catalog);
        };

        var btnFitMonitor = new Wpf.Ui.Controls.Button
        {
            Content = "📐 Ajustar",
            Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary,
            Padding = new Thickness(9, 4, 9, 4),
            Margin = new Thickness(0, 0, 6, 0),
            ToolTip = "Adapta imediatamente a resolução remota ao painel e DPI do monitor atual"
        };
        btnFitMonitor.Click += (_, _) => rdpControl.AdaptResolutionToContainer(useFullMonitorBounds: false);

        var btnFullScreen = new Wpf.Ui.Controls.Button
        {
            Content = "⛶ Tela Cheia",
            Appearance = Wpf.Ui.Controls.ControlAppearance.Primary,
            Padding = new Thickness(10, 4, 10, 4),
            Margin = new Thickness(0, 0, 6, 0),
            ToolTip = "Abre em Tela Cheia nativa adaptada ao monitor onde o AmbientesHub está (Ctrl+Alt+Break para voltar à aba)"
        };
        btnFullScreen.Click += (_, _) => rdpControl.EnterMonitorFullScreen();

        var btnReconnect = new Wpf.Ui.Controls.Button
        {
            Content = "🔄 Reconectar",
            Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary,
            Padding = new Thickness(9, 4, 9, 4),
            Margin = new Thickness(0, 0, 6, 0),
            ToolTip = "Reconecta aplicando o perfil de qualidade e resolução selecionados"
        };
        btnReconnect.Click += (_, _) => rdpControl.Reconnect(entry);

        var btnExternal = new Wpf.Ui.Controls.Button
        {
            Content = "🪟 mstsc",
            Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary,
            Padding = new Thickness(9, 4, 9, 4),
            Margin = new Thickness(0, 0, 6, 0),
            ToolTip = "Abrir sessão externa no mstsc.exe com o perfil configurado"
        };
        btnExternal.Click += (_, _) => LaunchExternalRdp(entry);

        var btnClose = new Wpf.Ui.Controls.Button
        {
            Content = "✕ Fechar",
            Appearance = Wpf.Ui.Controls.ControlAppearance.Danger,
            Padding = new Thickness(10, 4, 10, 4),
            ToolTip = "Desconectar e fechar esta aba"
        };

        actions.Children.Add(cmbProfile);
        actions.Children.Add(chkScale);
        actions.Children.Add(btnFitMonitor);
        actions.Children.Add(btnFullScreen);
        actions.Children.Add(btnReconnect);
        actions.Children.Add(btnExternal);
        actions.Children.Add(btnClose);
        WControls.Grid.SetColumn(actions, 1);
        tbGrid.Children.Add(actions);

        toolbar.Child = tbGrid;
        WControls.Grid.SetRow(toolbar, 0);
        tabGrid.Children.Add(toolbar);

        WControls.Grid.SetRow(wfHost, 1);
        tabGrid.Children.Add(wfHost);

        var tab = new WControls.TabItem
        {
            Header = $"🖥️ {entry.Name}",
            Content = tabGrid
        };

        btnClose.Click += (_, _) =>
        {
            rdpControl.SafeDisconnect();
            wfHost.Dispose();
            MainRdpTabs.Items.Remove(tab);
            MainRdpTabs.SelectedIndex = 0;
        };

        MainRdpTabs.Items.Add(tab);
        MainRdpTabs.SelectedItem = tab;

        int w = (int)Math.Max(1024, ActualWidth - 40);
        int h = (int)Math.Max(700, ActualHeight - 100);
        rdpControl.Connect(entry, w, h);
    }

    private void SetStatus(string text, string bgHex, string fgHex, string borderHex)
    {
        TxtRdpStatus.Text = text;
        TxtRdpStatus.Foreground = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(fgHex));
        BadgeRdpStatus.Background = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(bgHex));
        BadgeRdpStatus.BorderBrush = new WMedia.SolidColorBrush((WMedia.Color)WMedia.ColorConverter.ConvertFromString(borderHex));
    }
}
