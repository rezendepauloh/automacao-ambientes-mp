using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using AmbientesHub.Models;
using AmbientesHub.Services;
using WControls = System.Windows.Controls;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Views;

public partial class DocumentationView : WControls.UserControl
{
    private const double DefaultFontSize = 13.5;
    private const double MinFontSize = 11.0;
    private const double MaxFontSize = 20.0;

    private readonly DocumentationService _docsService = new();
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
    private bool _loadedOnce;
    private bool _isUpdatingCategoryCombo;

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

    public void SetSearchFilter(string query)
    {
        EnsureLoadedOnce();
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
        _allDocs = _docsService.DiscoverDocuments(_docsRootPath);

        PopulateCategoryDropdown();
        ApplyFilter();

        if (_allDocs.Count == 0)
        {
            _currentDoc = null;
            _currentRawMarkdown = string.Empty;
            TxtCurrentDocTitle.Text = "Nenhum arquivo .md encontrado";
            TxtCurrentDocMeta.Text = _docsRootPath;
            RenderPlaceholder($"Nenhum documento Markdown (.md) foi localizado em `{_docsRootPath}`.");
            SetStatus("⚠️ Pasta Docs vazia ou não localizada", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        var target = (!string.IsNullOrEmpty(previousFullPath)
            ? _allDocs.FirstOrDefault(d => d.FullPath.Equals(previousFullPath, StringComparison.OrdinalIgnoreCase))
            : null)
            ?? _visibleDocs.FirstOrDefault()
            ?? _allDocs[0];

        OpenDocument(target, recordHistory: false);
    }

    private void PopulateCategoryDropdown()
    {
        _isUpdatingCategoryCombo = true;
        try
        {
            string? prev = CmbDocCategory.SelectedItem as string;
            var categories = _allDocs
                .OrderBy(d => d.CategorySortOrder)
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
            _isUpdatingCategoryCombo = false;
        }
    }

    private void ApplyFilter()
    {
        string search = TxtSearchDocs.Text.Trim();
        string category = CmbDocCategory.SelectedItem as string ?? "Todas as Categorias";

        var filtered = _docsService.FilterDocuments(_allDocs, search, category, _docsRootPath);

        _visibleDocs.Clear();
        foreach (var doc in filtered)
        {
            doc.IsSelected = _currentDoc != null &&
                             doc.FullPath.Equals(_currentDoc.FullPath, StringComparison.OrdinalIgnoreCase);
            _visibleDocs.Add(doc);
        }

        if (string.IsNullOrWhiteSpace(search) && category == "Todas as Categorias")
        {
            SetStatus($"📚 {_allDocs.Count} documento(s) disponíveis", "#1e293b", "#38bdf8", "#334155");
        }
        else
        {
            SetStatus($"🔎 {_visibleDocs.Count} de {_allDocs.Count} documento(s)", "#1e3a8a", "#93c5fd", "#3b82f6");
        }
    }

    private void OpenDocument(DocEntryItem item, bool recordHistory)
    {
        try
        {
            string rawMd = _docsService.ReadDocumentContent(_docsRootPath, item.FullPath);

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
            TxtCurrentDocMeta.Text = $"Docs/{item.RelativePath}  •  Atualizado em {item.LastUpdatedDisplay}";
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
        MarkdownViewer.Document = MarkdownFlowDocumentBuilder.Build(
            blocks,
            _baseFontSize,
            _codeFontName,
            onLinkClicked: HandleMarkdownLinkClicked,
            onCodeCopied: lang =>
            {
                SetStatus($"📋 Bloco {lang} copiado!", "#064e3b", "#34d399", "#10b981");
            });

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

        var res = DocumentationService.ResolveMarkdownLinkTarget(_docsRootPath, _currentDoc.FullPath, rawHref);
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

            case MarkdownLinkKind.ExternalHttpUrl:
                _browserLauncher.OpenUrlsInFirefox([res.Target]);
                SetStatus("🌐 Link externo aberto no navegador", "#1e3a8a", "#93c5fd", "#3b82f6");
                break;

            case MarkdownLinkKind.UnsupportedReference:
                SetStatus($"ℹ️ Referência externa à pasta Docs: {res.Target}", "#1e293b", "#94a3b8", "#334155");
                break;
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

    private void CmbDocCategory_SelectionChanged(object sender, WControls.SelectionChangedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingCategoryCombo) return;
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
        if (!InputValidator.IsSafeDirectoryPath(_docsRootPath) || !Directory.Exists(_docsRootPath))
        {
            SetStatus("⚠️ Diretório Docs não encontrado.", "#78350f", "#fbbf24", "#f59e0b");
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                UseShellExecute = false
            };
            psi.ArgumentList.Add(_docsRootPath);
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao abrir pasta Docs no Explorer: {ex.Message}", nameof(DocumentationView));
        }
    }

    #endregion
}
