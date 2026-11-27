using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AmbientesHub.Models;

public enum DocSourceKind
{
    ProjectDocs,
    SharePointFaqs
}

public class DocEntryItem : INotifyPropertyChanged
{
    private bool _isSelected;

    public string Title { get; init; } = string.Empty;
    public string Emoji { get; init; } = "📄";
    public string FileName { get; init; } = string.Empty;
    public string RelativePath { get; init; } = string.Empty;
    public string FullPath { get; init; } = string.Empty;
    public string SourceRootPath { get; init; } = string.Empty;
    public DocSourceKind SourceKind { get; init; } = DocSourceKind.ProjectDocs;
    public string CategoryKey { get; init; } = string.Empty;
    public string CategoryLabel { get; init; } = "Geral";
    public string BadgeText { get; init; } = "DOC";
    public string BadgeBgHex { get; init; } = "#1e293b";
    public string BadgeFgHex { get; init; } = "#94a3b8";
    public string Summary { get; init; } = string.Empty;
    public DateTime LastModifiedLocal { get; init; }
    public int CategorySortOrder { get; init; } = 99;
    public bool IsIndexReadme { get; init; }

    public string SourceLabel => SourceKind == DocSourceKind.SharePointFaqs
        ? "Tutoriais SharePoint"
        : "Docs do Hub";

    public string SourceBadgeText => SourceKind == DocSourceKind.SharePointFaqs
        ? "SHAREPOINT"
        : "HUB";

    public string SourceBadgeBgHex => SourceKind == DocSourceKind.SharePointFaqs
        ? "#0c4a6e"
        : "#1e293b";

    public string SourceBadgeFgHex => SourceKind == DocSourceKind.SharePointFaqs
        ? "#7dd3fc"
        : "#94a3b8";

    public string DisplayPathPrefix => SourceKind == DocSourceKind.SharePointFaqs
        ? "SharePoint"
        : "Docs";

    public string LastUpdatedDisplay => LastModifiedLocal == default
        ? string.Empty
        : LastModifiedLocal.ToString("dd/MM/yyyy HH:mm");

    public bool HasSummary => !string.IsNullOrWhiteSpace(Summary);

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(BorderHex));
            OnPropertyChanged(nameof(CardBgHex));
        }
    }

    public string BorderHex => _isSelected ? "#38bdf8" : "#1e293b";
    public string CardBgHex => _isSelected ? "#172554" : "#0f172a";

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
}

public enum MarkdownBlockKind
{
    Heading,
    Paragraph,
    CodeBlock,
    BlockQuote,
    List,
    Table,
    HorizontalRule,
    Image,
    AttachmentCard
}

public enum MarkdownInlineKind
{
    Text,
    Bold,
    Italic,
    Code,
    Link,
    Strikethrough,
    Image
}

public sealed record MarkdownInlineSpan(
    MarkdownInlineKind Kind,
    string Text,
    string? TargetUrl = null);

public sealed record MarkdownListItem(
    int IndentLevel,
    string Marker,
    IReadOnlyList<MarkdownInlineSpan> Inlines);

public sealed class MarkdownBlock
{
    public MarkdownBlockKind Kind { get; init; }
    public int HeadingLevel { get; init; }
    public string Text { get; init; } = string.Empty;
    public string CodeLanguage { get; init; } = string.Empty;
    public string AlertKind { get; init; } = string.Empty;
    public string ImageAlt { get; init; } = string.Empty;
    public string ImageSource { get; init; } = string.Empty;
    public string AttachmentLabel { get; init; } = string.Empty;
    public string AttachmentSource { get; init; } = string.Empty;
    public IReadOnlyList<MarkdownInlineSpan> Inlines { get; init; } = [];
    public IReadOnlyList<MarkdownListItem> ListItems { get; init; } = [];
    public IReadOnlyList<IReadOnlyList<MarkdownInlineSpan>> TableHeaders { get; init; } = [];
    public IReadOnlyList<IReadOnlyList<IReadOnlyList<MarkdownInlineSpan>>> TableRows { get; init; } = [];
}

public enum MarkdownLinkKind
{
    InternalDocument,
    ExternalHttpUrl,
    LocalImage,
    LocalAttachment,
    UnsupportedReference
}

public sealed record MarkdownLinkResolution(
    MarkdownLinkKind Kind,
    string Target,
    string? Anchor = null);
