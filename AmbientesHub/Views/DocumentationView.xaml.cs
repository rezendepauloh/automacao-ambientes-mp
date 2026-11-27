using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Input;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Microsoft.Win32;
using WControls = System.Windows.Controls;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Views;

public partial class DocumentationView : WControls.UserControl
{
    private const double DefaultFontSize = 13.5;
    private const double MinFontSize = 11.0;
    private const double MaxFontSize = 20.0;

    private readonly DocumentationService _docsService = new();
    private readonly SharePointFaqSyncService _spSyncService = new();
    private readonly BrowserLauncherService _browserLauncher = new();
    private readonly ObservableCollection<DocEntryItem> _visibleDocs = [];
    private readonly Stack<string> _backHistory = new();
    private readonly Stack<string> _forwardHistory = new();

    private List<DocEntryItem> _allDocs = [];
    private DocEntryItem? _currentDoc;
    private string _currentRawMarkdown = string.Empty;
    private string _docsRootPath = DocumentationService.ResolveDefaultDocsDirectory();
    private string _codeFontName = "Consolas";
    private double _baseFontSize = DefaultFontSize;
    private double _savedSidebarWidth = 345;
    private bool _isSidebarVisible = true;
    private bool _loadedOnce;
    private bool _isUpdatingCombos;

    public DocumentationView()
    {
        InitializeComponent();
        ItemsDocs.ItemsSource = _visibleDocs;
        _codeFontName = ConfigService.Get("APP_TERMINAL_FONT", "Consolas");
    }

    public void EnsureLoadedOnce()
    {
        if (_loadedOnce) return;
        _loadedOnce = true;
        ReloadCatalog();
    }

    public void ReloadFromConfig()
    {
        if (!_loadedOnce) return;
        ReloadCatalog();
    }

    public void SetSearchFilter(string query)
    {
        EnsureLoadedOnce();
        CmbDocSource.SelectedItem = "Todas as Origens";
        CmbDocCategory.SelectedItem = "Todas as Categorias";
        TxtSearchDocs.Text = (query ?? string.Empty).Trim();
        ApplyFilter();
        if (_visibleDocs.Count > 0)
        {
            OpenDocument(_visibleDocs[0], recordHistory: true);
        }
    }

    public void ApplyCodeFont(string fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName)) return;
        _codeFontName = fontName;
        if (_currentDoc != null && !string.IsNullOrEmpty(_currentRawMarkdown))
        {
            RenderCurrentMarkdown();
        }
    }

    private void ReloadCatalog()
    {
        string? previousFullPath = _currentDoc?.FullPath;
        _docsRootPath = DocumentationService.ResolveDefaultDocsDirectory();
        _allDocs = _docsService.DiscoverAllDocuments(_docsRootPath);

        PopulateSourceDropdown();
        PopulateCategoryDropdown();
        ApplyFilter();

        if (_allDocs.Count == 0)
        {
            _currentDoc = null;
            _currentRawMarkdown = string.Empty;
            TxtCurrentDocTitle.Text = "Nenhum arquivo .md encontrado";
            TxtCurrentDocMeta.Text = _docsRootPath;
            RenderPlaceholder($"Nenhum documento Markdown (.md) foi localizado em `{_docsRootPath}`.");
            SetStatus("⚠️ Nenhum documento encontrado", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        var target = (!string.IsNullOrEmpty(previousFullPath)
            ? _allDocs.FirstOrDefault(d => d.FullPath.Equals(previousFullPath, StringComparison.OrdinalIgnoreCase))
            : null)
            ?? _visibleDocs.FirstOrDefault()
            ?? _allDocs[0];

        OpenDocument(target, recordHistory: false);
    }

    private void PopulateSourceDropdown()
    {
        _isUpdatingCombos = true;
        try
        {
            string? prev = CmbDocSource.SelectedItem as string;
            var sources = new List<string> { "Todas as Origens", "Docs do Hub" };
            if (_allDocs.Any(d => d.SourceKind == DocSourceKind.SharePointFaqs))
            {
                sources.Add("Tutoriais SharePoint");
            }

            CmbDocSource.ItemsSource = sources;
            CmbDocSource.SelectedItem = (!string.IsNullOrEmpty(prev) && sources.Contains(prev))
                ? prev
                : "Todas as Origens";
        }
        finally
        {
            _isUpdatingCombos = false;
        }
    }

    private void PopulateCategoryDropdown()
    {
        _isUpdatingCombos = true;
        try
        {
            string source = CmbDocSource.SelectedItem as string ?? "Todas as Origens";
            string? prev = CmbDocCategory.SelectedItem as string;

            IEnumerable<DocEntryItem> scoped = _allDocs;
            if (!source.Equals("Todas as Origens", StringComparison.OrdinalIgnoreCase))
            {
                scoped = scoped.Where(d => d.SourceLabel.Equals(source, StringComparison.OrdinalIgnoreCase));
            }

            var categories = scoped
                .OrderBy(d => d.CategorySortOrder)
                .ThenBy(d => d.CategoryLabel, StringComparer.CurrentCultureIgnoreCase)
                .Select(d => d.CategoryLabel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            categories.Insert(0, "Todas as Categorias");
            CmbDocCategory.ItemsSource = categories;
            CmbDocCategory.SelectedItem = (!string.IsNullOrEmpty(prev) && categories.Contains(prev))
                ? prev
                : "Todas as Categorias";
        }
        finally
        {
            _isUpdatingCombos = false;
        }
    }

    private void ApplyFilter()
    {
        string search = TxtSearchDocs.Text.Trim();
        string source = CmbDocSource.SelectedItem as string ?? "Todas as Origens";
        string category = CmbDocCategory.SelectedItem as string ?? "Todas as Categorias";

        var filtered = _docsService.FilterDocuments(_allDocs, search, category, _docsRootPath, source);

        _visibleDocs.Clear();
        foreach (var doc in filtered)
        {
            doc.IsSelected = _currentDoc != null &&
                             doc.FullPath.Equals(_currentDoc.FullPath, StringComparison.OrdinalIgnoreCase);
            _visibleDocs.Add(doc);
        }

        int spCount = _allDocs.Count(d => d.SourceKind == DocSourceKind.SharePointFaqs);
        if (string.IsNullOrWhiteSpace(search) && source == "Todas as Origens" && category == "Todas as Categorias")
        {
            string spSuffix = spCount > 0 ? $" • {spCount} SharePoint" : string.Empty;
            SetStatus($"📚 {_allDocs.Count} doc(s){spSuffix}", "#1e293b", "#38bdf8", "#334155");
        }
        else
        {
            SetStatus($"🔎 {_visibleDocs.Count} de {_allDocs.Count} doc(s)", "#1e3a8a", "#93c5fd", "#3b82f6");
        }
    }

    private void OpenDocument(DocEntryItem item, bool recordHistory)
    {
        try
        {
            string sourceRoot = !string.IsNullOrWhiteSpace(item.SourceRootPath) ? item.SourceRootPath : _docsRootPath;
            string rawMd = _docsService.ReadDocumentContent(sourceRoot, item.FullPath);

            if (recordHistory && _currentDoc != null &&
                !_currentDoc.FullPath.Equals(item.FullPath, StringComparison.OrdinalIgnoreCase))
            {
                _backHistory.Push(_currentDoc.FullPath);
                _forwardHistory.Clear();
            }

            _currentDoc = item;
            _currentRawMarkdown = rawMd;

            foreach (var d in _allDocs)
            {
                d.IsSelected = d.FullPath.Equals(item.FullPath, StringComparison.OrdinalIgnoreCase);
            }

            TxtCurrentDocTitle.Text = $"{item.Emoji} {item.Title}";
            TxtCurrentDocMeta.Text = $"{item.DisplayPathPrefix}/{item.RelativePath}  •  Atualizado em {item.LastUpdatedDisplay}";
            UpdateHistoryButtons();

            RenderCurrentMarkdown();
            ScrollReaderToTop();
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Erro ao abrir documento '{item.RelativePath}': {ex.Message}", nameof(DocumentationView));
            SetStatus($"⚠️ Falha ao ler '{item.FileName}'", "#78350f", "#fbbf24", "#f59e0b");
        }
    }

    private void RenderCurrentMarkdown()
    {
        var blocks = DocumentationService.ParseMarkdown(_currentRawMarkdown);
        string sourceRoot = _currentDoc != null && !string.IsNullOrWhiteSpace(_currentDoc.SourceRootPath)
            ? _currentDoc.SourceRootPath
            : _docsRootPath;

        MarkdownViewer.Document = MarkdownFlowDocumentBuilder.Build(
            blocks,
            _baseFontSize,
            _codeFontName,
            onLinkClicked: HandleMarkdownLinkClicked,
            onCodeCopied: lang =>
            {
                SetStatus($"📋 Bloco {lang} copiado!", "#064e3b", "#34d399", "#10b981");
            },
            sourceRootPath: sourceRoot,
            currentDocFullPath: _currentDoc?.FullPath);

        int zoomPct = (int)Math.Round((_baseFontSize / DefaultFontSize) * 100.0);
        BtnZoomReset.Content = $"{zoomPct}%";
    }

    private void RenderPlaceholder(string markdownMessage)
    {
        var blocks = DocumentationService.ParseMarkdown(markdownMessage);
        MarkdownViewer.Document = MarkdownFlowDocumentBuilder.Build(blocks, _baseFontSize, _codeFontName);
    }

    private void HandleMarkdownLinkClicked(string rawHref)
    {
        if (_currentDoc == null) return;

        string sourceRoot = !string.IsNullOrWhiteSpace(_currentDoc.SourceRootPath)
            ? _currentDoc.SourceRootPath
            : _docsRootPath;

        var res = DocumentationService.ResolveMarkdownLinkTarget(sourceRoot, _currentDoc.FullPath, rawHref);
        switch (res.Kind)
        {
            case MarkdownLinkKind.InternalDocument:
                var targetDoc = _allDocs.FirstOrDefault(
                    d => d.FullPath.Equals(res.Target, StringComparison.OrdinalIgnoreCase));
                if (targetDoc != null)
                {
                    OpenDocument(targetDoc, recordHistory: true);
                    SetStatus($"📖 Aberto: {targetDoc.RelativePath}", "#1e293b", "#38bdf8", "#334155");
                }
                break;

            case MarkdownLinkKind.LocalImage:
            case MarkdownLinkKind.LocalAttachment:
                OpenLocalMediaOnDemand(res.Target);
                break;

            case MarkdownLinkKind.ExternalHttpUrl:
                _browserLauncher.OpenUrlsInFirefox([res.Target]);
                SetStatus("🌐 Link externo aberto no navegador", "#1e3a8a", "#93c5fd", "#3b82f6");
                break;

            case MarkdownLinkKind.UnsupportedReference:
                SetStatus($"ℹ️ Referência externa: {res.Target}", "#1e293b", "#94a3b8", "#334155");
                break;
        }
    }

    private void OpenLocalMediaOnDemand(string validatedFullPath)
    {
        try
        {
            // Dispara pelo aplicativo associado do Windows; no OneDrive Files On-Demand,
            // o Windows Cloud Filter baixa este único arquivo automaticamente ao abrir.
            Process.Start(new ProcessStartInfo
            {
                FileName = validatedFullPath,
                UseShellExecute = true
            });
            SetStatus($"📂 Abrindo sob demanda: {Path.GetFileName(validatedFullPath)}", "#064e3b", "#34d399", "#10b981");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao abrir mídia local '{Path.GetFileName(validatedFullPath)}': {ex.Message}", nameof(DocumentationView));
            SetStatus($"⚠️ Não foi possível abrir '{Path.GetFileName(validatedFullPath)}'", "#78350f", "#fbbf24", "#f59e0b");
        }
    }

    private void UpdateHistoryButtons()
    {
        BtnDocBack.IsEnabled = _backHistory.Count > 0;
        BtnDocForward.IsEnabled = _forwardHistory.Count > 0;
    }

    private void ScrollReaderToTop()
    {
        try
        {
            if (FindDescendant<WControls.ScrollViewer>(MarkdownViewer) is { } sv)
            {
                sv.ScrollToHome();
            }
        }
        catch
        {
            // Ignora caso a árvore visual ainda não esteja materializada
        }
    }

    private static T? FindDescendant<T>(DependencyObject root) where T : DependencyObject
    {
        int count = WMedia.VisualTreeHelper.GetChildrenCount(root);
        for (int i = 0; i < count; i++)
        {
            var child = WMedia.VisualTreeHelper.GetChild(root, i);
            if (child is T typed) return typed;
            if (FindDescendant<T>(child) is { } nested) return nested;
        }
        return null;
    }

    private void SetStatus(string text, string bgHex, string fgHex, string borderHex)
    {
        var conv = new WMedia.BrushConverter();
        TxtDocsStatus.Text = text;
        BadgeDocsStatus.Background = (WMedia.Brush)conv.ConvertFromString(bgHex)!;
        TxtDocsStatus.Foreground = (WMedia.Brush)conv.ConvertFromString(fgHex)!;
        BadgeDocsStatus.BorderBrush = (WMedia.Brush)conv.ConvertFromString(borderHex)!;
    }

    #region Eventos de UI

    private void TxtSearchDocs_TextChanged(object sender, WControls.TextChangedEventArgs e)
    {
        if (!IsLoaded) return;
        ApplyFilter();
    }

    private void CmbDocSource_SelectionChanged(object sender, WControls.SelectionChangedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingCombos) return;
        PopulateCategoryDropdown();
        ApplyFilter();
    }

    private void CmbDocCategory_SelectionChanged(object sender, WControls.SelectionChangedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingCombos) return;
        ApplyFilter();
    }

    private void BtnReloadDocs_Click(object sender, RoutedEventArgs e)
    {
        ReloadCatalog();
    }

    private void DocCard_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement { Tag: DocEntryItem doc })
        {
            OpenDocument(doc, recordHistory: true);
        }
    }

    private void BtnToggleDocsSidebar_Click(object sender, RoutedEventArgs e)
    {
        _isSidebarVisible = !_isSidebarVisible;
        if (_isSidebarVisible)
        {
            ColDocsCatalog.MinWidth = 250;
            ColDocsCatalog.Width = new GridLength(Math.Clamp(_savedSidebarWidth, 250, 500));
            ColDocsSplitter.Width = new GridLength(8);
            PanelDocsCatalog.Visibility = Visibility.Visible;
            SplitterDocsCatalog.Visibility = Visibility.Visible;
            BtnToggleDocsSidebar.Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary;
            BtnToggleDocsSidebar.ToolTip = "Ocultar lista de documentos (Modo Leitura)";
        }
        else
        {
            if (ColDocsCatalog.ActualWidth >= 250)
                _savedSidebarWidth = ColDocsCatalog.ActualWidth;

            ColDocsCatalog.MinWidth = 0;
            ColDocsCatalog.Width = new GridLength(0);
            ColDocsSplitter.Width = new GridLength(0);
            PanelDocsCatalog.Visibility = Visibility.Collapsed;
            SplitterDocsCatalog.Visibility = Visibility.Collapsed;
            BtnToggleDocsSidebar.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;
            BtnToggleDocsSidebar.ToolTip = "Mostrar lista de documentos";
        }
    }

    private void BtnDocBack_Click(object sender, RoutedEventArgs e)
    {
        if (_backHistory.Count == 0) return;
        string targetPath = _backHistory.Pop();
        var target = _allDocs.FirstOrDefault(d => d.FullPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase));
        if (target != null)
        {
            if (_currentDoc != null)
                _forwardHistory.Push(_currentDoc.FullPath);
            OpenDocument(target, recordHistory: false);
        }
        UpdateHistoryButtons();
    }

    private void BtnDocForward_Click(object sender, RoutedEventArgs e)
    {
        if (_forwardHistory.Count == 0) return;
        string targetPath = _forwardHistory.Pop();
        var target = _allDocs.FirstOrDefault(d => d.FullPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase));
        if (target != null)
        {
            if (_currentDoc != null)
                _backHistory.Push(_currentDoc.FullPath);
            OpenDocument(target, recordHistory: false);
        }
        UpdateHistoryButtons();
    }

    private void BtnZoomOut_Click(object sender, RoutedEventArgs e)
    {
        _baseFontSize = Math.Max(MinFontSize, Math.Round(_baseFontSize - 1.0, 1));
        if (_currentDoc != null) RenderCurrentMarkdown();
    }

    private void BtnZoomReset_Click(object sender, RoutedEventArgs e)
    {
        _baseFontSize = DefaultFontSize;
        if (_currentDoc != null) RenderCurrentMarkdown();
    }

    private void BtnZoomIn_Click(object sender, RoutedEventArgs e)
    {
        _baseFontSize = Math.Min(MaxFontSize, Math.Round(_baseFontSize + 1.0, 1));
        if (_currentDoc != null) RenderCurrentMarkdown();
    }

    private void BtnSyncSharePoint_Click(object sender, RoutedEventArgs e)
    {
        string? spDir = DocumentationService.ResolveConfiguredSharePointDirectory();
        if (string.IsNullOrWhiteSpace(spDir))
        {
            MessageBox.Show(
                "Configure primeiro SHAREPOINT_FAQS_DIR na aba Configurações & Credenciais.",
                "Sincronização SharePoint",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        string? apiUrl = SharePointFaqSyncService.BuildSitePagesApiUrl();
        if (!string.IsNullOrWhiteSpace(apiUrl))
        {
            var answer = MessageBox.Show(
                "Deseja abrir agora o feed de SitePages do SharePoint no Firefox?\n\n" +
                "• Sim: abre no Firefox já logado. O Firefox baixa automaticamente o arquivo 'items' (ou 'items.xml') na pasta Downloads; se exibir o XML na tela, pressione Ctrl+S.\n" +
                "• Não: pula direto para selecionar um arquivo já salvo.",
                "Sincronizar FAQs do SharePoint",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (answer == MessageBoxResult.Cancel) return;
            if (answer == MessageBoxResult.Yes)
            {
                _browserLauncher.OpenUrlsInFirefox([apiUrl]);
                SetStatus("🌐 Feed aberto no Firefox — selecione o arquivo baixado em Downloads", "#1e3a8a", "#93c5fd", "#3b82f6");
            }
        }

        string downloadsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        var dlg = new OpenFileDialog
        {
            Title = "Selecione o arquivo de SitePages baixado pelo navegador (ex: items ou items.xml)",
            InitialDirectory = Directory.Exists(downloadsDir) ? downloadsDir : spDir,
            Filter = "Feed SharePoint (items*;*.xml;*.txt)|items*;*.xml;*.txt|Todos os arquivos (*.*)|*.*",
            CheckFileExists = true
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            var res = _spSyncService.SyncFromFeedFile(dlg.FileName, spDir);
            ReloadCatalog();
            SetStatus($"✅ {res.PagesConverted} FAQs sincronizados ({res.LocalImagesMapped} imgs)", "#064e3b", "#34d399", "#10b981");
            MessageBox.Show(
                $"Sincronização concluída com sucesso!\n\n" +
                $"• FAQs atualizados: {res.PagesConverted}\n" +
                $"• Imagens locais mapeadas: {res.LocalImagesMapped}\n" +
                $"• Vídeos/anexos sob demanda: {res.LocalFilesMapped}\n" +
                $"• Pasta: {res.OutputDirectory}\n\n" +
                "O OneDrive sincronizará os .md atualizados automaticamente para toda a equipe.",
                "Sincronização SharePoint",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao sincronizar FAQs do SharePoint: {ex.Message}", nameof(DocumentationView));
            SetStatus("⚠️ Falha na sincronização SharePoint", "#78350f", "#fbbf24", "#f59e0b");
            MessageBox.Show($"Não foi possível sincronizar o XML:\n{ex.Message}", "Erro na Sincronização", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnExportHtml_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDoc == null || string.IsNullOrWhiteSpace(_currentRawMarkdown)) return;

        string suggestedName = Path.GetFileNameWithoutExtension(_currentDoc.FileName) + "_offline.html";
        var dlg = new SaveFileDialog
        {
            Title = "Exportar Tutorial Offline (HTML Autocontido)",
            Filter = "Página HTML Offline (*.html)|*.html",
            FileName = suggestedName,
            DefaultExt = ".html"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            string html = DocumentationService.ExportStandaloneHtml(_currentDoc, _currentRawMarkdown);
            File.WriteAllText(dlg.FileName, html, Encoding.UTF8);
            SetStatus($"📦 Exportado: {Path.GetFileName(dlg.FileName)}", "#064e3b", "#34d399", "#10b981");
            AppLoggerService.Success($"Tutorial '{_currentDoc.RelativePath}' exportado para HTML offline.", nameof(DocumentationView));
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao exportar HTML offline: {ex.Message}", nameof(DocumentationView));
            SetStatus("⚠️ Falha ao exportar HTML", "#78350f", "#fbbf24", "#f59e0b");
        }
    }

    private void BtnCopyRawMd_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentRawMarkdown)) return;
        if (ClipboardHelper.Copy(_currentRawMarkdown))
        {
            SetStatus($"📋 Markdown de '{_currentDoc?.FileName}' copiado!", "#064e3b", "#34d399", "#10b981");
        }
    }

    private void BtnOpenDocsFolder_Click(object sender, RoutedEventArgs e)
    {
        string targetDir = _currentDoc != null && !string.IsNullOrWhiteSpace(_currentDoc.FullPath)
            ? (Path.GetDirectoryName(_currentDoc.FullPath) ?? _docsRootPath)
            : _docsRootPath;

        if (!InputValidator.IsSafeDirectoryPath(targetDir) || !Directory.Exists(targetDir))
        {
            SetStatus("⚠️ Diretório não encontrado.", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                UseShellExecute = false
            };
            psi.ArgumentList.Add(targetDir);
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao abrir pasta no Explorer: {ex.Message}", nameof(DocumentationView));
        }
    }

    #endregion
}
