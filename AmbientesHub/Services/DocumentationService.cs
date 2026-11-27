using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo para descoberta multi-origem (Docs/ do projeto + Tutoriais SharePoint via .env),
/// leitura segura (path-containment + bloqueio de symlinks/junctions compatível com OneDrive Files On-Demand),
/// filtragem, parsing Markdown (incluindo imagens e cards de mídia sob demanda) e exportação HTML offline.
/// </summary>
public class DocumentationService
{
    private const long MaxMarkdownFileBytes = 2 * 1024 * 1024; // 2 MB
    private const long MaxEmbeddedImageBytes = 15 * 1024 * 1024; // 15 MB
    private static readonly TimeSpan RxTimeout = TimeSpan.FromMilliseconds(250);

    private static readonly HashSet<string> AllowedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"
    };

    private static readonly HashSet<string> AllowedAttachmentExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mp4", ".mkv", ".webm", ".mov", ".avi", ".pdf", ".docx", ".xlsx", ".pptx", ".txt"
    };

    private static readonly Regex HeadingRegex = new(
        @"^(#{1,6})\s+(.+?)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex HrRegex = new(
        @"^\s*(?:-\s*-\s*-[-\s]*|\*\s*\*\s*\*[\*\s]*|_\s*_\s*_[_\s]*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex FenceOpenRegex = new(
        @"^\s*(`{3,}|~{3,})\s*([^\s`~]*).*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex StandaloneImageRegex = new(
        @"^\s*!\[(?<alt>[^\]\r\n]*)\]\((?<src><[^>\r\n]+>|(?:[^\)\r\n(]|\([^\)\r\n]*\))+)\)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex StandaloneLinkRegex = new(
        @"^\s*\[(?<label>[^\]\r\n]+)\]\((?<src><[^>\r\n]+>|(?:[^\)\r\n(]|\([^\)\r\n]*\))+)\)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex AlertTagRegex = new(
        @"^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*(.*)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex ListItemRegex = new(
        @"^(?<indent>[ \t]*)(?<marker>(?:[-*+])|(?:\d+\.))\s+(?<content>.+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex TaskBoxRegex = new(
        @"^\[(?<state>[ xX])\]\s+(?<rest>.+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex TableSeparatorCellRegex = new(
        @"^:?-{2,}:?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    // Tokenizador inline: código em crases, imagens ![alt](src), links [txt](url) (com suporte a <url> e parênteses balanceados), **negrito**, __negrito__, ~~riscado~~, *itálico* e _itálico_ seguro (preserva snake_case).
    private static readonly Regex InlineTokenRegex = new(
        @"(?<code>`[^`\r\n]+`)" +
        @"|(?<img>!\[(?<ialt>[^\]\r\n]*)\]\((?<isrc><[^>\r\n]+>|(?:[^\)\r\n(]|\([^\)\r\n]*\))+)\))" +
        @"|(?<link>\[(?<ltext>[^\]\r\n]+)\]\((?<lurl><[^>\r\n]+>|(?:[^\)\r\n(]|\([^\)\r\n]*\))+)\))" +
        @"|(?<bold1>\*\*(?<b1text>.+?)\*\*)" +
        @"|(?<bold2>(?<![\p{L}\p{N}])__(?<b2text>.+?)__(?![\p{L}\p{N}]))" +
        @"|(?<strike>~~(?<stext>.+?)~~)" +
        @"|(?<ital1>(?<!\*)\*(?<i1text>[^\*\s](?:.*?[^\*\s])?)\*(?!\*))" +
        @"|(?<ital2>(?<![\p{L}\p{N}_])_(?<i2text>[^_\s](?:.*?[^_\s])?)_(?![\p{L}\p{N}_]))",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    public static string ResolveDefaultDocsDirectory()
    {
        try
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            for (int depth = 0; depth <= 6 && dir != null; depth++)
            {
                string candidate = Path.Combine(dir.FullName, "Docs");
                if (Directory.Exists(candidate))
                    return Path.GetFullPath(candidate);

                dir = dir.Parent;
            }
        }
        catch
        {
            // Prossegue para o fallback padrão
        }

        const string fallback = @"C:\AmbientesVirtuais\Docs";
        return Path.GetFullPath(fallback);
    }

    /// <summary>
    /// Retorna o diretório sincronizado de FAQs do SharePoint configurado no .env (fail-closed: null se ausente ou inválido).
    /// </summary>
    public static string? ResolveConfiguredSharePointDirectory()
    {
        string raw = ConfigService.SharePointFaqsDir;
        if (!InputValidator.IsSafeDirectoryPath(raw))
            return null;

        try
        {
            string full = Path.GetFullPath(raw.Trim());
            if (!Directory.Exists(full))
                return null;

            if (IsSymbolicLinkOrJunction(new DirectoryInfo(full)))
                return null;

            return full;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Garante que o arquivo seja estritamente um .md contido dentro de docsRoot, sem caracteres de controle
    /// e sem atravessar symlinks/junctions (compatível com OneDrive Files On-Demand).
    /// </summary>
    public static bool IsSafeDocPath(string docsRoot, string? candidatePath)
        => IsSafeContainedPath(docsRoot, candidatePath, ext => ext.Equals(".md", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Garante que a imagem referenciada esteja contida na raiz permitida e tenha extensão de imagem na whitelist.
    /// </summary>
    public static bool IsSafeMediaPath(string allowedRoot, string? candidatePath)
        => IsSafeContainedPath(allowedRoot, candidatePath, AllowedImageExtensions.Contains);

    /// <summary>
    /// Garante que vídeo/documento anexo esteja contido na raiz permitida e tenha extensão na whitelist.
    /// </summary>
    public static bool IsSafeAttachmentPath(string allowedRoot, string? candidatePath)
        => IsSafeContainedPath(allowedRoot, candidatePath, AllowedAttachmentExtensions.Contains);

    public List<DocEntryItem> DiscoverAllDocuments(string? customDocsRoot = null, string? customSharePointRoot = null)
    {
        var combined = new List<DocEntryItem>();

        string docsRoot = string.IsNullOrWhiteSpace(customDocsRoot)
            ? ResolveDefaultDocsDirectory()
            : Path.GetFullPath(customDocsRoot);

        combined.AddRange(DiscoverDocuments(docsRoot, DocSourceKind.ProjectDocs));

        string? spRoot = !string.IsNullOrWhiteSpace(customSharePointRoot)
            ? Path.GetFullPath(customSharePointRoot)
            : ResolveConfiguredSharePointDirectory();

        if (!string.IsNullOrWhiteSpace(spRoot) && Directory.Exists(spRoot))
        {
            combined.AddRange(DiscoverDocuments(spRoot, DocSourceKind.SharePointFaqs));
        }

        return combined;
    }

    public List<DocEntryItem> DiscoverDocuments(string? customDocsRoot = null, DocSourceKind sourceKind = DocSourceKind.ProjectDocs)
    {
        string sourceRoot = string.IsNullOrWhiteSpace(customDocsRoot)
            ? ResolveDefaultDocsDirectory()
            : Path.GetFullPath(customDocsRoot);

        var results = new List<DocEntryItem>();
        if (!Directory.Exists(sourceRoot))
            return results;

        var rootDirInfo = new DirectoryInfo(sourceRoot);
        if (IsSymbolicLinkOrJunction(rootDirInfo))
            return results;

        // Para SharePoint, prioriza a subpasta "Manuais Markdown" para não varrer pastas de mídia de 14 GB.
        DirectoryInfo scanRoot = rootDirInfo;
        if (sourceKind == DocSourceKind.SharePointFaqs)
        {
            string mdSubDir = Path.Combine(sourceRoot, "Manuais Markdown");
            if (Directory.Exists(mdSubDir) && !IsSymbolicLinkOrJunction(new DirectoryInfo(mdSubDir)))
            {
                scanRoot = new DirectoryInfo(mdSubDir);
            }
        }

        foreach (string filePath in EnumerateMarkdownFilesSafe(scanRoot, sourceRoot))
        {
            try
            {
                var fi = new FileInfo(filePath);
                if (fi.Length > MaxMarkdownFileBytes)
                    continue;

                string content = File.ReadAllText(filePath, Encoding.UTF8);
                string relPathFromScan = Path.GetRelativePath(scanRoot.FullName, filePath).Replace('\\', '/');
                bool isRootIndex = relPathFromScan.Equals("README.md", StringComparison.OrdinalIgnoreCase);

                string folderKey = ExtractTopFolderKey(relPathFromScan);
                var (catLabel, badgeText, badgeBg, badgeFg, catOrder) = GetCategoryMetadata(folderKey, isRootIndex, sourceKind);
                var (emoji, title) = ExtractDocumentTitle(content, fi.Name, isRootIndex, folderKey, sourceKind);
                string summary = ExtractDocumentSummary(content);

                results.Add(new DocEntryItem
                {
                    Title = title,
                    Emoji = emoji,
                    FileName = fi.Name,
                    RelativePath = relPathFromScan,
                    FullPath = fi.FullName,
                    SourceRootPath = sourceRoot,
                    SourceKind = sourceKind,
                    CategoryKey = folderKey,
                    CategoryLabel = catLabel,
                    BadgeText = badgeText,
                    BadgeBgHex = badgeBg,
                    BadgeFgHex = badgeFg,
                    Summary = summary,
                    LastModifiedLocal = fi.LastWriteTime,
                    CategorySortOrder = catOrder,
                    IsIndexReadme = isRootIndex
                });
            }
            catch (Exception ex)
            {
                AppLoggerService.Warning(
                    $"Falha ao indexar documento '{Path.GetFileName(filePath)}': {ex.Message}",
                    nameof(DocumentationService));
            }
        }

        return results
            .OrderBy(d => d.CategorySortOrder)
            .ThenByDescending(d => d.IsIndexReadme)
            .ThenBy(d => d.Title, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public string ReadDocumentContent(string docsRoot, string fullPath)
    {
        if (!IsSafeDocPath(docsRoot, fullPath))
            throw new UnauthorizedAccessException("Caminho de documento fora do diretório permitido.");

        var fi = new FileInfo(fullPath);
        if (!fi.Exists)
            throw new FileNotFoundException("Documento Markdown não encontrado.", fullPath);

        if (fi.Length > MaxMarkdownFileBytes)
            throw new InvalidOperationException("O documento excede o tamanho máximo permitido de 2 MB.");

        return File.ReadAllText(fi.FullName, Encoding.UTF8);
    }

    public List<DocEntryItem> FilterDocuments(
        IEnumerable<DocEntryItem> allDocs,
        string? searchQuery,
        string? categoryLabel,
        string? fallbackDocsRoot = null,
        string? sourceFilter = null)
    {
        IEnumerable<DocEntryItem> query = allDocs;

        if (!string.IsNullOrWhiteSpace(sourceFilter) &&
            !sourceFilter.Equals("Todas as Origens", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(d => d.SourceLabel.Equals(sourceFilter, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(categoryLabel) &&
            !categoryLabel.Equals("Todas as Categorias", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(d => d.CategoryLabel.Equals(categoryLabel, StringComparison.OrdinalIgnoreCase));
        }

        string trimmedSearch = (searchQuery ?? string.Empty).Trim();
        if (!string.IsNullOrEmpty(trimmedSearch))
        {
            query = query.Where(d => MatchesSearch(d, trimmedSearch, fallbackDocsRoot));
        }

        return query.ToList();
    }

    public static MarkdownLinkResolution ResolveMarkdownLinkTarget(
        string sourceRoot,
        string currentDocFullPath,
        string? rawHref)
    {
        if (string.IsNullOrWhiteSpace(rawHref))
            return new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, string.Empty);

        string href = NormalizeMarkdownHref(rawHref);
        if (string.IsNullOrEmpty(href) || href.IndexOfAny(['\0', '\r', '\n']) >= 0)
            return new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, string.Empty);

        // Link externo HTTP/HTTPS seguro
        if (href.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            href.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            return InputValidator.IsSafeHttpUrl(href)
                ? new MarkdownLinkResolution(MarkdownLinkKind.ExternalHttpUrl, href)
                : new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, href);
        }

        string? anchor = null;
        int hashIdx = href.IndexOf('#');
        if (hashIdx >= 0)
        {
            anchor = href[(hashIdx + 1)..].Trim();
            href = href[..hashIdx].Trim();
        }

        // Âncora no próprio documento
        if (string.IsNullOrEmpty(href))
        {
            return IsSafeDocPath(sourceRoot, currentDocFullPath)
                ? new MarkdownLinkResolution(MarkdownLinkKind.InternalDocument, Path.GetFullPath(currentDocFullPath), anchor)
                : new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, rawHref);
        }

        try
        {
            string candidateFullPath;
            if (href.StartsWith("file://", StringComparison.OrdinalIgnoreCase) &&
                Uri.TryCreate(href, UriKind.Absolute, out var fileUri) &&
                fileUri.IsFile)
            {
                candidateFullPath = Path.GetFullPath(fileUri.LocalPath);
            }
            else
            {
                string baseDir = Path.GetDirectoryName(Path.GetFullPath(currentDocFullPath)) ?? sourceRoot;
                string normalizedRel = Uri.UnescapeDataString(href).Replace('/', Path.DirectorySeparatorChar);
                candidateFullPath = Path.GetFullPath(Path.Combine(baseDir, normalizedRel));
            }

            if (IsSafeDocPath(sourceRoot, candidateFullPath) && File.Exists(candidateFullPath))
                return new MarkdownLinkResolution(MarkdownLinkKind.InternalDocument, candidateFullPath, anchor);

            if (IsSafeMediaPath(sourceRoot, candidateFullPath) && File.Exists(candidateFullPath))
                return new MarkdownLinkResolution(MarkdownLinkKind.LocalImage, candidateFullPath, anchor);

            if (IsSafeAttachmentPath(sourceRoot, candidateFullPath) && File.Exists(candidateFullPath))
                return new MarkdownLinkResolution(MarkdownLinkKind.LocalAttachment, candidateFullPath, anchor);

            return new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, href, anchor);
        }
        catch
        {
            return new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, href, anchor);
        }
    }

    public static IReadOnlyList<MarkdownBlock> ParseMarkdown(string? markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return [];

        string[] lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var blocks = new List<MarkdownBlock>();
        int i = 0;

        while (i < lines.Length)
        {
            string line = lines[i];

            if (string.IsNullOrWhiteSpace(line))
            {
                i++;
                continue;
            }

            // 1. Bloco de código cercado (``` ou ~~~, suportando 3 ou mais delimitadores)
            var fenceMatch = SafeMatch(FenceOpenRegex, line);
            if (fenceMatch.Success)
            {
                string fenceToken = fenceMatch.Groups[1].Value;
                char fenceChar = fenceToken[0];
                int fenceLen = fenceToken.Length;
                string lang = fenceMatch.Groups[2].Value.Trim();

                var codeSb = new StringBuilder();
                i++;
                while (i < lines.Length)
                {
                    string cl = lines[i];
                    if (IsClosingFence(cl, fenceChar, fenceLen))
                    {
                        i++;
                        break;
                    }
                    if (codeSb.Length > 0) codeSb.Append('\n');
                    codeSb.Append(cl);
                    i++;
                }

                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.CodeBlock,
                    CodeLanguage = lang,
                    Text = codeSb.ToString()
                });
                continue;
            }

            // 2. Linha horizontal (---, ***, ___)
            if (SafeMatch(HrRegex, line).Success)
            {
                blocks.Add(new MarkdownBlock { Kind = MarkdownBlockKind.HorizontalRule });
                i++;
                continue;
            }

            // 3. Imagem em linha isolada (![alt](src))
            var imgMatch = SafeMatch(StandaloneImageRegex, line);
            if (imgMatch.Success)
            {
                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.Image,
                    ImageAlt = imgMatch.Groups["alt"].Value.Trim(),
                    ImageSource = NormalizeMarkdownHref(imgMatch.Groups["src"].Value)
                });
                i++;
                continue;
            }

            // 4. Card de vídeo/anexo em linha isolada ([Rótulo](caminho.mp4|pdf|docx...))
            var attachMatch = SafeMatch(StandaloneLinkRegex, line);
            if (attachMatch.Success && IsAttachmentHref(attachMatch.Groups["src"].Value))
            {
                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.AttachmentCard,
                    AttachmentLabel = StripInlineFormatting(attachMatch.Groups["label"].Value),
                    AttachmentSource = NormalizeMarkdownHref(attachMatch.Groups["src"].Value)
                });
                i++;
                continue;
            }

            // 5. Títulos (# a ######)
            var headingMatch = SafeMatch(HeadingRegex, line);
            if (headingMatch.Success)
            {
                int level = headingMatch.Groups[1].Value.Length;
                string rawHeading = headingMatch.Groups[2].Value.Trim().TrimEnd('#').Trim();
                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.Heading,
                    HeadingLevel = level,
                    Text = StripInlineFormatting(rawHeading),
                    Inlines = ParseInlines(rawHeading)
                });
                i++;
                continue;
            }

            // 6. Blockquote / Alertas GitHub (> ...)
            if (line.TrimStart().StartsWith('>'))
            {
                var quoteLines = new List<string>();
                while (i < lines.Length && lines[i].TrimStart().StartsWith('>'))
                {
                    string q = lines[i].TrimStart()[1..];
                    if (q.StartsWith(' ')) q = q[1..];
                    quoteLines.Add(q);
                    i++;
                }

                string alertKind = string.Empty;
                if (quoteLines.Count > 0)
                {
                    var alertMatch = SafeMatch(AlertTagRegex, quoteLines[0].Trim());
                    if (alertMatch.Success)
                    {
                        alertKind = alertMatch.Groups[1].Value.ToUpperInvariant();
                        string remainder = alertMatch.Groups[2].Value.Trim();
                        if (string.IsNullOrEmpty(remainder))
                            quoteLines.RemoveAt(0);
                        else
                            quoteLines[0] = remainder;
                    }
                }

                string joinedQuote = string.Join(" ", quoteLines.Where(l => !string.IsNullOrWhiteSpace(l))).Trim();
                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.BlockQuote,
                    AlertKind = alertKind,
                    Text = StripInlineFormatting(joinedQuote),
                    Inlines = ParseInlines(joinedQuote)
                });
                continue;
            }

            // 7. Tabela Markdown (| Col1 | Col2 | seguida por linha separadora | --- | --- |)
            if (IsPotentialTableRow(line) && i + 1 < lines.Length && IsTableSeparatorLine(lines[i + 1]))
            {
                var headerCells = SplitTableCells(line)
                    .Select(ParseInlines)
                    .ToList();

                i += 2; // pula cabeçalho + separador
                var rows = new List<IReadOnlyList<IReadOnlyList<MarkdownInlineSpan>>>();

                while (i < lines.Length && IsPotentialTableRow(lines[i]) && !string.IsNullOrWhiteSpace(lines[i]))
                {
                    var rowCells = SplitTableCells(lines[i])
                        .Select(ParseInlines)
                        .ToList();
                    rows.Add(rowCells);
                    i++;
                }

                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.Table,
                    TableHeaders = headerCells,
                    TableRows = rows
                });
                continue;
            }

            // 8. Listas ordenadas, não ordenadas e checklists
            var listMatch = SafeMatch(ListItemRegex, line);
            if (listMatch.Success)
            {
                var items = new List<MarkdownListItem>();
                while (i < lines.Length)
                {
                    if (string.IsNullOrWhiteSpace(lines[i]))
                    {
                        if (i + 1 < lines.Length && SafeMatch(ListItemRegex, lines[i + 1]).Success)
                        {
                            i++;
                            continue;
                        }
                        break;
                    }

                    var m = SafeMatch(ListItemRegex, lines[i]);
                    if (!m.Success)
                        break;

                    string indentRaw = m.Groups["indent"].Value.Replace("\t", "  ");
                    int indentLevel = Math.Clamp(indentRaw.Length / 2, 0, 5);
                    string rawMarker = m.Groups["marker"].Value;
                    string itemContent = m.Groups["content"].Value.Trim();

                    string displayMarker = rawMarker.EndsWith('.') ? rawMarker : "•";
                    var taskMatch = SafeMatch(TaskBoxRegex, itemContent);
                    if (taskMatch.Success)
                    {
                        bool isChecked = !taskMatch.Groups["state"].Value.Equals(" ", StringComparison.Ordinal);
                        displayMarker = isChecked ? "☑️" : "⬜";
                        itemContent = taskMatch.Groups["rest"].Value.Trim();
                    }

                    items.Add(new MarkdownListItem(indentLevel, displayMarker, ParseInlines(itemContent)));
                    i++;
                }

                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.List,
                    ListItems = items
                });
                continue;
            }

            // 9. Parágrafo comum
            var paraLines = new List<string>();
            while (i < lines.Length &&
                   !string.IsNullOrWhiteSpace(lines[i]) &&
                   !SafeMatch(FenceOpenRegex, lines[i]).Success &&
                   !SafeMatch(HrRegex, lines[i]).Success &&
                   !SafeMatch(StandaloneImageRegex, lines[i]).Success &&
                   !(SafeMatch(StandaloneLinkRegex, lines[i]) is { Success: true } slm && IsAttachmentHref(slm.Groups["src"].Value)) &&
                   !SafeMatch(HeadingRegex, lines[i]).Success &&
                   !lines[i].TrimStart().StartsWith('>') &&
                   !SafeMatch(ListItemRegex, lines[i]).Success &&
                   !(IsPotentialTableRow(lines[i]) && i + 1 < lines.Length && IsTableSeparatorLine(lines[i + 1])))
            {
                paraLines.Add(lines[i].Trim());
                i++;
            }

            string paragraphText = string.Join(" ", paraLines);
            if (!string.IsNullOrWhiteSpace(paragraphText))
            {
                blocks.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.Paragraph,
                    Text = StripInlineFormatting(paragraphText),
                    Inlines = ParseInlines(paragraphText)
                });
            }
        }

        return blocks;
    }

    public static IReadOnlyList<MarkdownInlineSpan> ParseInlines(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return [];

        var spans = new List<MarkdownInlineSpan>();
        int lastIndex = 0;

        try
        {
            foreach (Match m in InlineTokenRegex.Matches(text))
            {
                if (m.Index > lastIndex)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Text, text[lastIndex..m.Index]));
                }

                if (m.Groups["code"].Success)
                {
                    string rawCode = m.Groups["code"].Value;
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Code, rawCode[1..^1]));
                }
                else if (m.Groups["img"].Success)
                {
                    string alt = m.Groups["ialt"].Value.Trim();
                    string src = NormalizeMarkdownHref(m.Groups["isrc"].Value);
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Image, string.IsNullOrEmpty(alt) ? Path.GetFileName(src) : alt, src));
                }
                else if (m.Groups["link"].Success)
                {
                    string linkText = StripInlineFormatting(m.Groups["ltext"].Value);
                    string linkUrl = NormalizeMarkdownHref(m.Groups["lurl"].Value);
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Link, linkText, linkUrl));
                }
                else if (m.Groups["bold1"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Bold, m.Groups["b1text"].Value.Replace("`", string.Empty)));
                }
                else if (m.Groups["bold2"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Bold, m.Groups["b2text"].Value.Replace("`", string.Empty)));
                }
                else if (m.Groups["strike"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Strikethrough, m.Groups["stext"].Value.Replace("`", string.Empty)));
                }
                else if (m.Groups["ital1"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Italic, m.Groups["i1text"].Value.Replace("`", string.Empty)));
                }
                else if (m.Groups["ital2"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Italic, m.Groups["i2text"].Value.Replace("`", string.Empty)));
                }

                lastIndex = m.Index + m.Length;
            }
        }
        catch (RegexMatchTimeoutException)
        {
            return [new MarkdownInlineSpan(MarkdownInlineKind.Text, text)];
        }

        if (lastIndex < text.Length)
        {
            spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Text, text[lastIndex..]));
        }

        return spans;
    }

    /// <summary>
    /// Monta a URL direta no SharePoint para um vídeo ou anexo relativo dentro de SHAREPOINT_FAQS_DIR,
    /// permitindo que técnicos em campo abram o vídeo diretamente no SharePoint a partir do HTML exportado.
    /// </summary>
    public static string? BuildSharePointWebUrl(
        DocEntryItem doc,
        string? attachmentSource,
        string? sharePointWebBaseUrl = null)
    {
        if (doc.SourceKind != DocSourceKind.SharePointFaqs || string.IsNullOrWhiteSpace(attachmentSource))
            return null;

        string baseUrl = (sharePointWebBaseUrl ?? ConfigService.SharePointFaqsWebUrl).Trim();
        if (!InputValidator.IsSafeHttpUrl(baseUrl))
            return null;

        string sourceRoot = !string.IsNullOrWhiteSpace(doc.SourceRootPath)
            ? doc.SourceRootPath
            : (ResolveConfiguredSharePointDirectory() ?? string.Empty);

        if (string.IsNullOrWhiteSpace(sourceRoot) || string.IsNullOrWhiteSpace(doc.FullPath))
            return null;

        try
        {
            string href = NormalizeMarkdownHref(attachmentSource);
            int hashIdx = href.IndexOf('#');
            if (hashIdx >= 0) href = href[..hashIdx];

            string baseDir = Path.GetDirectoryName(Path.GetFullPath(doc.FullPath)) ?? sourceRoot;
            string normalizedRel = Uri.UnescapeDataString(href).Replace('/', Path.DirectorySeparatorChar);
            string candidateFull = Path.GetFullPath(Path.Combine(baseDir, normalizedRel));

            if (!IsSafeAttachmentPath(sourceRoot, candidateFull) && !IsSafeMediaPath(sourceRoot, candidateFull))
                return null;

            string relFromRoot = Path.GetRelativePath(Path.GetFullPath(sourceRoot), candidateFull).Replace('\\', '/');
            if (relFromRoot.StartsWith("..", StringComparison.Ordinal))
                return null;

            string encodedPath = string.Join("/",
                relFromRoot.Split('/', StringSplitOptions.RemoveEmptyEntries)
                    .Select(Uri.EscapeDataString));

            string finalUrl = $"{baseUrl.TrimEnd('/')}/{encodedPath}";
            return InputValidator.IsSafeHttpUrl(finalUrl) ? finalUrl : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Gera um fragmento HTML semântico e limpo (sem fundo escuro nem estilos invasivos),
    /// ideal para colar diretamente na Web Part "Texto" do SharePoint via CF_HTML.
    /// Remove o H1 inicial e o cabeçalho "Origem SharePoint:" que já existem na página do SharePoint.
    /// </summary>
    public static string ExportSharePointRichHtml(DocEntryItem doc, string rawMarkdown, string? sharePointWebBaseUrl = null)
    {
        var blocks = ParseMarkdown(rawMarkdown);
        var sb = new StringBuilder();
        bool skippedLeadingH1 = false;
        bool skippedOriginBanner = false;

        for (int i = 0; i < blocks.Count; i++)
        {
            var block = blocks[i];

            if (!skippedLeadingH1 && block.Kind == MarkdownBlockKind.Heading && block.HeadingLevel == 1)
            {
                skippedLeadingH1 = true;
                continue;
            }

            if (!skippedOriginBanner && block.Kind == MarkdownBlockKind.BlockQuote &&
                block.Text.Contains("Origem SharePoint:", StringComparison.OrdinalIgnoreCase))
            {
                skippedOriginBanner = true;
                if (i + 1 < blocks.Count && blocks[i + 1].Kind == MarkdownBlockKind.HorizontalRule)
                    i++;
                continue;
            }

            switch (block.Kind)
            {
                case MarkdownBlockKind.Heading:
                    int lvl = Math.Clamp(block.HeadingLevel, 2, 4);
                    sb.AppendLine($"<h{lvl}>{RenderHtmlInlines(block.Inlines)}</h{lvl}>");
                    break;

                case MarkdownBlockKind.Paragraph:
                    sb.AppendLine($"<p>{RenderHtmlInlines(block.Inlines)}</p>");
                    break;

                case MarkdownBlockKind.HorizontalRule:
                    sb.AppendLine("<hr />");
                    break;

                case MarkdownBlockKind.BlockQuote:
                    string alertLabel = block.AlertKind switch
                    {
                        "NOTE" => "ℹ️ NOTA",
                        "TIP" => "💡 DICA",
                        "IMPORTANT" => "📌 IMPORTANTE",
                        "WARNING" => "⚠️ ATENÇÃO",
                        "CAUTION" => "🛑 CUIDADO",
                        _ => string.Empty
                    };
                    string prefix = string.IsNullOrEmpty(alertLabel)
                        ? string.Empty
                        : $"<strong>{WebUtility.HtmlEncode(alertLabel)}:</strong> ";
                    sb.AppendLine($"<blockquote><p>{prefix}{RenderHtmlInlines(block.Inlines)}</p></blockquote>");
                    break;

                case MarkdownBlockKind.CodeBlock:
                    sb.AppendLine($"<pre><code>{WebUtility.HtmlEncode(block.Text)}</code></pre>");
                    break;

                case MarkdownBlockKind.List:
                    RenderSharePointList(sb, block.ListItems);
                    break;

                case MarkdownBlockKind.Table:
                    sb.AppendLine("<table><thead><tr>");
                    foreach (var h in block.TableHeaders)
                        sb.Append($"<th>{RenderHtmlInlines(h)}</th>");
                    sb.AppendLine("</tr></thead><tbody>");
                    foreach (var row in block.TableRows)
                    {
                        sb.Append("<tr>");
                        foreach (var cell in row)
                            sb.Append($"<td>{RenderHtmlInlines(cell)}</td>");
                        sb.AppendLine("</tr>");
                    }
                    sb.AppendLine("</tbody></table>");
                    break;

                case MarkdownBlockKind.Image:
                    string? imgSrc = ResolveSharePointRichImageSrc(doc, block.ImageSource, sharePointWebBaseUrl);
                    if (!string.IsNullOrWhiteSpace(imgSrc))
                    {
                        string safeAlt = WebUtility.HtmlEncode(string.IsNullOrWhiteSpace(block.ImageAlt) ? Path.GetFileName(block.ImageSource) : block.ImageAlt);
                        sb.AppendLine($"<p><img src=\"{WebUtility.HtmlEncode(imgSrc)}\" alt=\"{safeAlt}\" /></p>");
                    }
                    break;

                case MarkdownBlockKind.AttachmentCard:
                    string? spUrl = BuildSharePointWebUrl(doc, block.AttachmentSource, sharePointWebBaseUrl);
                    string safeLabel = WebUtility.HtmlEncode(block.AttachmentLabel);
                    if (!string.IsNullOrWhiteSpace(spUrl))
                        sb.AppendLine($"<p>🎬 <a href=\"{WebUtility.HtmlEncode(spUrl)}\">{safeLabel}</a></p>");
                    else
                        sb.AppendLine($"<p>📎 {safeLabel}</p>");
                    break;
            }
        }

        return sb.ToString().Trim();
    }

    private static void RenderSharePointList(StringBuilder sb, IReadOnlyList<MarkdownListItem> items)
    {
        if (items.Count == 0) return;

        var openTags = new Stack<string>();
        int currentLevel = -1;

        for (int idx = 0; idx < items.Count; idx++)
        {
            var item = items[idx];
            int level = Math.Max(0, item.IndentLevel);
            string tag = item.Marker.EndsWith('.') ? "ol" : "ul";

            while (level > currentLevel)
            {
                sb.AppendLine($"<{tag}>");
                openTags.Push(tag);
                currentLevel++;
            }

            while (level < currentLevel && openTags.Count > 0)
            {
                sb.AppendLine($"</{openTags.Pop()}>");
                currentLevel--;
            }

            sb.AppendLine($"<li>{RenderHtmlInlines(item.Inlines)}</li>");
        }

        while (openTags.Count > 0)
        {
            sb.AppendLine($"</{openTags.Pop()}>");
        }
    }

    private static string? ResolveSharePointRichImageSrc(DocEntryItem doc, string imageSource, string? sharePointWebBaseUrl)
    {
        if (string.IsNullOrWhiteSpace(imageSource))
            return null;

        if (InputValidator.IsSafeHttpUrl(imageSource))
            return imageSource;

        string sourceRoot = !string.IsNullOrWhiteSpace(doc.SourceRootPath)
            ? doc.SourceRootPath
            : (ResolveConfiguredSharePointDirectory() ?? ResolveDefaultDocsDirectory());

        string? resolvedLocal = null;
        if (!string.IsNullOrWhiteSpace(sourceRoot) && !string.IsNullOrWhiteSpace(doc.FullPath))
        {
            var res = ResolveMarkdownLinkTarget(sourceRoot, doc.FullPath, imageSource);
            if (res.Kind == MarkdownLinkKind.LocalImage)
                resolvedLocal = res.Target;
        }

        if (doc.SourceKind == DocSourceKind.SharePointFaqs)
        {
            string baseUrl = (sharePointWebBaseUrl ?? ConfigService.SharePointFaqsWebUrl).Trim();
            if (!string.IsNullOrWhiteSpace(resolvedLocal) && InputValidator.IsSafeHttpUrl(baseUrl))
            {
                string rel = Path.GetRelativePath(Path.GetFullPath(sourceRoot), resolvedLocal).Replace('\\', '/');
                const string siteAssetsPrefix = "Imagens FAQ/SiteAssets/";
                if (rel.StartsWith(siteAssetsPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    int sharedIdx = baseUrl.IndexOf("/Shared%20Documents", StringComparison.OrdinalIgnoreCase);
                    if (sharedIdx < 0)
                        sharedIdx = baseUrl.IndexOf("/Shared Documents", StringComparison.OrdinalIgnoreCase);

                    if (sharedIdx > 0)
                    {
                        string siteRoot = baseUrl[..sharedIdx].TrimEnd('/');
                        string subPath = rel[siteAssetsPrefix.Length..];
                        string encodedSub = string.Join("/", subPath.Split('/', StringSplitOptions.RemoveEmptyEntries).Select(Uri.EscapeDataString));
                        string candidate = $"{siteRoot}/SiteAssets/SitePages/{encodedSub}";
                        if (InputValidator.IsSafeHttpUrl(candidate))
                            return candidate;
                    }
                }
            }

            string? spDocUrl = BuildSharePointWebUrl(doc, imageSource, sharePointWebBaseUrl);
            if (!string.IsNullOrWhiteSpace(spDocUrl))
                return spDocUrl;
        }

        if (!string.IsNullOrWhiteSpace(resolvedLocal) && File.Exists(resolvedLocal))
        {
            try
            {
                var fi = new FileInfo(resolvedLocal);
                if (fi.Length > 0 && fi.Length <= 5 * 1024 * 1024)
                {
                    string mime = GetImageMimeType(Path.GetExtension(resolvedLocal));
                    return $"data:{mime};base64,{Convert.ToBase64String(File.ReadAllBytes(resolvedLocal))}";
                }
            }
            catch
            {
                // Ignora imagem ilegível
            }
        }

        return null;
    }

    /// <summary>
    /// Exporta o documento atual como HTML5 standalone (autocontido) com imagens locais embutidas em Base64
    /// e links diretos do SharePoint para vídeos/anexos, ideal para envio à equipe de campo.
    /// </summary>
    public static string ExportStandaloneHtml(DocEntryItem doc, string rawMarkdown, string? sharePointWebBaseUrl = null)
    {
        var blocks = ParseMarkdown(rawMarkdown);
        string sourceRoot = !string.IsNullOrWhiteSpace(doc.SourceRootPath)
            ? doc.SourceRootPath
            : ResolveDefaultDocsDirectory();

        var body = new StringBuilder();
        var toc = new StringBuilder();
        int headingIndex = 0;

        foreach (var block in blocks)
        {
            switch (block.Kind)
            {
                case MarkdownBlockKind.Heading:
                    int lvl = Math.Clamp(block.HeadingLevel, 1, 6);
                    headingIndex++;
                    string hid = $"sec-{headingIndex}";
                    body.AppendLine($"<h{lvl} id=\"{hid}\">{RenderHtmlInlines(block.Inlines)}</h{lvl}>");
                    if (lvl <= 3)
                    {
                        string tocText = WebUtility.HtmlEncode(block.Text);
                        toc.AppendLine($"<a class=\"toc-item toc-l{lvl}\" href=\"#{hid}\">{tocText}</a>");
                    }
                    break;

                case MarkdownBlockKind.Paragraph:
                    body.AppendLine($"<p>{RenderHtmlInlines(block.Inlines)}</p>");
                    break;

                case MarkdownBlockKind.HorizontalRule:
                    body.AppendLine("<hr />");
                    break;

                case MarkdownBlockKind.BlockQuote:
                    string alertCls = string.IsNullOrEmpty(block.AlertKind) ? "quote" : $"quote quote-{block.AlertKind.ToLowerInvariant()}";
                    string alertLabel = block.AlertKind switch
                    {
                        "NOTE" => "ℹ️ NOTA",
                        "TIP" => "💡 DICA",
                        "IMPORTANT" => "📌 IMPORTANTE",
                        "WARNING" => "⚠️ ATENÇÃO",
                        "CAUTION" => "🛑 CUIDADO",
                        _ => block.AlertKind
                    };
                    string alertHeader = string.IsNullOrEmpty(alertLabel)
                        ? string.Empty
                        : $"<div class=\"quote-badge\">{WebUtility.HtmlEncode(alertLabel)}</div>";
                    body.AppendLine($"<blockquote class=\"{alertCls}\">{alertHeader}{RenderHtmlInlines(block.Inlines)}</blockquote>");
                    break;

                case MarkdownBlockKind.CodeBlock:
                    string lang = string.IsNullOrWhiteSpace(block.CodeLanguage) ? "TEXT" : block.CodeLanguage.ToUpperInvariant();
                    body.AppendLine($"<div class=\"code-card\"><div class=\"code-head\"><span class=\"code-lang\">{WebUtility.HtmlEncode(lang)}</span><button type=\"button\" class=\"copy-btn\" onclick=\"copyCode(this)\">📋 Copiar</button></div><pre><code>{WebUtility.HtmlEncode(block.Text)}</code></pre></div>");
                    break;

                case MarkdownBlockKind.List:
                    body.AppendLine("<ul class=\"md-list\">");
                    foreach (var item in block.ListItems)
                    {
                        int pad = item.IndentLevel * 14;
                        body.AppendLine($"<li style=\"margin-left:{pad}px\"><span class=\"marker\">{WebUtility.HtmlEncode(item.Marker)}</span><span>{RenderHtmlInlines(item.Inlines)}</span></li>");
                    }
                    body.AppendLine("</ul>");
                    break;

                case MarkdownBlockKind.Table:
                    body.AppendLine("<div class=\"table-wrap\"><table><thead><tr>");
                    foreach (var h in block.TableHeaders)
                        body.Append($"<th>{RenderHtmlInlines(h)}</th>");
                    body.AppendLine("</tr></thead><tbody>");
                    foreach (var row in block.TableRows)
                    {
                        body.Append("<tr>");
                        foreach (var cell in row)
                            body.Append($"<td>{RenderHtmlInlines(cell)}</td>");
                        body.AppendLine("</tr>");
                    }
                    body.AppendLine("</tbody></table></div>");
                    break;

                case MarkdownBlockKind.Image:
                    body.AppendLine(RenderHtmlImageBlock(sourceRoot, doc.FullPath, block.ImageAlt, block.ImageSource));
                    break;

                case MarkdownBlockKind.AttachmentCard:
                    body.AppendLine(RenderHtmlAttachmentCard(doc, block.AttachmentLabel, block.AttachmentSource, sharePointWebBaseUrl));
                    break;
            }
        }

        string safeTitle = WebUtility.HtmlEncode(doc.Title);
        string safePath = WebUtility.HtmlEncode($"{doc.DisplayPathPrefix}/{doc.RelativePath}");
        string safeDate = WebUtility.HtmlEncode($"{DateTime.Now:dd/MM/yyyy HH:mm}");
        string tocHtml = toc.Length > 0
            ? toc.ToString()
            : "<div class=\"toc-empty\">Sem seções numeradas</div>";

        return $$"""
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
              <meta charset="utf-8" />
              <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover" />
              <title>{{safeTitle}} — Ambientes Hub</title>
              <style>
                :root { color-scheme: dark; --bg: #0b0f19; --card: #0f172a; --border: #1e293b; --accent: #38bdf8; --text: #e2e8f0; --muted: #94a3b8; }
                * { box-sizing: border-box; }
                html { scroll-behavior: smooth; }
                body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); margin: 0; line-height: 1.6; overflow-x: hidden; }
                .topbar { position: sticky; top: 0; z-index: 30; display: flex; align-items: center; gap: 10px; padding: 10px 14px; background: rgba(15,23,42,0.95); backdrop-filter: blur(8px); border-bottom: 1px solid var(--border); }
                .topbar-btn { background: #1e293b; color: #f8fafc; border: 1px solid #334155; border-radius: 6px; padding: 6px 11px; font-size: 13px; font-weight: 600; cursor: pointer; white-space: nowrap; }
                .topbar-btn:hover { background: #334155; }
                .topbar-title { flex: 1; min-width: 0; font-size: 14px; font-weight: 600; color: #f8fafc; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
                .layout { display: flex; width: 100%; min-height: calc(100vh - 52px); }
                .toc-drawer { width: 280px; flex-shrink: 0; background: #0d1424; border-right: 1px solid var(--border); padding: 14px 12px; position: sticky; top: 52px; height: calc(100vh - 52px); overflow-y: auto; }
                body.toc-collapsed .toc-drawer { display: none; }
                .toc-header { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: var(--muted); padding: 0 8px 10px; border-bottom: 1px solid var(--border); margin-bottom: 8px; }
                .toc-item { display: block; padding: 6px 8px; margin: 2px 0; border-radius: 6px; border-left: 3px solid transparent; color: #cbd5e1; text-decoration: none; font-size: 13px; line-height: 1.35; overflow-wrap: anywhere; transition: background 0.15s ease, color 0.15s ease, border-color 0.15s ease; }
                .toc-item:hover { background: #162238; color: var(--accent); text-decoration: none; }
                .toc-item.active { background: #1e293b; color: var(--accent); border-left-color: var(--accent); font-weight: 600; text-decoration: none; }
                .toc-l1 { font-weight: 700; color: #f8fafc; }
                .toc-l2 { padding-left: 14px; }
                .toc-l3 { padding-left: 24px; font-size: 12px; color: var(--muted); }
                .toc-empty { font-size: 12px; color: var(--muted); padding: 8px; }
                .toc-backdrop { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.55); z-index: 35; }
                .content-wrap { flex: 1; min-width: 0; padding: 20px 18px 40px; }
                main { max-width: 880px; margin: 0 auto; background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 24px 28px; overflow-wrap: anywhere; }
                .meta { display: flex; flex-wrap: wrap; gap: 6px 14px; font-family: Consolas, monospace; font-size: 11.5px; color: var(--muted); border-bottom: 1px solid var(--border); padding-bottom: 12px; margin-bottom: 18px; overflow-wrap: anywhere; word-break: break-word; }
                h1, h2, h3, h4 { color: #f8fafc; margin-top: 1.4em; scroll-margin-top: 68px; line-height: 1.3; }
                h1 { font-size: 1.65rem; border-bottom: 1px solid #334155; padding-bottom: 8px; margin-top: 0; }
                h2 { font-size: 1.3rem; color: var(--accent); border-bottom: 1px solid var(--border); padding-bottom: 4px; }
                h3 { font-size: 1.1rem; color: #93c5fd; }
                a { color: var(--accent); text-decoration: none; overflow-wrap: anywhere; }
                a:hover { text-decoration: underline; }
                code { font-family: Consolas, monospace; background: #1e293b; color: #7dd3fc; padding: 2px 6px; border-radius: 4px; font-size: 0.88em; overflow-wrap: anywhere; }
                .code-card { background: #070b14; border: 1px solid var(--border); border-radius: 8px; margin: 14px 0; overflow: hidden; }
                .code-head { display: flex; justify-content: space-between; align-items: center; background: #1e293b; padding: 5px 10px; }
                .code-lang { color: var(--accent); font-family: Consolas, monospace; font-size: 11px; font-weight: 700; }
                .copy-btn { background: #0f172a; color: #e2e8f0; border: 1px solid #334155; border-radius: 4px; padding: 2px 8px; font-size: 11px; cursor: pointer; }
                pre { margin: 0; padding: 12px 14px; overflow-x: auto; -webkit-overflow-scrolling: touch; }
                pre code { background: transparent; color: var(--text); padding: 0; overflow-wrap: normal; }
                blockquote.quote { --q-accent: #38bdf8; --q-bg: #0f172a; margin: 14px 0; padding: 10px 14px; border: 1px solid var(--q-accent); border-left: 4px solid var(--q-accent); background: var(--q-bg); border-radius: 6px; color: #cbd5e1; }
                blockquote.quote-note { --q-accent: #38bdf8; --q-bg: #0c192c; }
                blockquote.quote-tip { --q-accent: #34d399; --q-bg: #071f1a; }
                blockquote.quote-important { --q-accent: #a78bfa; --q-bg: #1a1433; }
                blockquote.quote-warning { --q-accent: #fbbf24; --q-bg: #271c09; }
                blockquote.quote-caution { --q-accent: #f87171; --q-bg: #2a1215; }
                .quote-badge { font-size: 11.5px; font-weight: 700; margin-bottom: 4px; color: var(--q-accent); letter-spacing: .02em; }
                ul.md-list { list-style: none; padding-left: 4px; }
                ul.md-list li { display: flex; gap: 6px; margin: 6px 0; }
                .marker { color: var(--accent); font-weight: 700; flex-shrink: 0; }
                .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; margin: 14px 0; border-radius: 6px; }
                table { width: 100%; border-collapse: collapse; font-size: 0.9rem; }
                th, td { border: 1px solid #334155; padding: 8px 10px; text-align: left; }
                th { background: #1e293b; color: #f8fafc; }
                tr:nth-child(even) td { background: #111827; }
                figure.img-card { margin: 16px 0; padding: 10px; background: #070b14; border: 1px solid var(--border); border-radius: 8px; text-align: center; }
                figure.img-card img { max-width: 100%; height: auto; border-radius: 4px; cursor: zoom-in; }
                figure.img-card figcaption { margin-top: 8px; font-size: 12px; color: var(--muted); }
                .attach-card { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 12px; margin: 14px 0; padding: 12px 14px; background: #111c32; border: 1px solid #1e3a8a; border-radius: 8px; }
                .attach-info { flex: 1 1 220px; min-width: 0; }
                .attach-badge { display: inline-block; background: #1e3a8a; color: #93c5fd; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle; }
                .attach-path { font-family: Consolas, monospace; font-size: 11.5px; color: var(--muted); margin-top: 4px; overflow-wrap: anywhere; }
                .attach-btn { display: inline-block; background: #0284c7; color: #f8fafc !important; font-weight: 600; font-size: 13px; padding: 9px 14px; border-radius: 6px; text-decoration: none !important; text-align: center; }
                .attach-btn:hover { background: #0369a1; }
                hr { border: 0; border-top: 1px solid var(--border); margin: 18px 0; }
                .lightbox { display: none; position: fixed; inset: 0; z-index: 60; background: rgba(0,0,0,0.9); align-items: center; justify-content: center; padding: 16px; cursor: zoom-out; }
                .lightbox.open { display: flex; }
                .lightbox img { max-width: 96vw; max-height: 92vh; border-radius: 6px; }
                @media (max-width: 899px) {
                  .toc-drawer { position: fixed; top: 0; left: 0; bottom: 0; height: 100vh; z-index: 40; width: 82vw; max-width: 310px; margin-left: 0 !important; transform: translateX(-105%); box-shadow: 4px 0 24px rgba(0,0,0,0.5); }
                  body.toc-mobile-open .toc-drawer { transform: translateX(0); }
                  body.toc-mobile-open .toc-backdrop { display: block; }
                  .content-wrap { padding: 12px 10px 32px; }
                  main { padding: 16px 14px; border-radius: 10px; }
                  h1 { font-size: 1.38rem; }
                  h2 { font-size: 1.15rem; }
                  .attach-btn { width: 100%; }
                }
                @media print { .topbar, .toc-drawer, .toc-backdrop, .copy-btn, .lightbox { display: none !important; } body { background: #fff; color: #111; } main { background: #fff; border: 0; padding: 0; } }
              </style>
            </head>
            <body>
              <header class="topbar">
                <button type="button" class="topbar-btn" id="btnToggleToc" aria-label="Abrir ou ocultar índice">☰ Índice</button>
                <div class="topbar-title">{{safeTitle}}</div>
                <button type="button" class="topbar-btn" onclick="window.scrollTo({top:0,behavior:'smooth'})">↑ Topo</button>
              </header>
              <div class="toc-backdrop" id="tocBackdrop"></div>
              <div class="layout">
                <aside class="toc-drawer" id="tocDrawer" aria-label="Índice do documento">
                  <div class="toc-header">📑 Seções do Tutorial</div>
                  {{tocHtml}}
                </aside>
                <div class="content-wrap">
                  <main>
                    <div class="meta"><span>📄 {{safePath}}</span><span>🕒 Exportado em {{safeDate}}</span></div>
                    {{body}}
                  </main>
                </div>
              </div>
              <div class="lightbox" id="imgLightbox" onclick="this.classList.remove('open')"><img id="imgLightboxTarget" alt="" /></div>
              <script>
                (function () {
                  const btn = document.getElementById('btnToggleToc');
                  const backdrop = document.getElementById('tocBackdrop');
                  const drawer = document.getElementById('tocDrawer');
                  const isMobile = () => window.innerWidth < 900;
                  function toggleToc() {
                    if (isMobile()) document.body.classList.toggle('toc-mobile-open');
                    else document.body.classList.toggle('toc-collapsed');
                  }
                  function closeMobileToc() { document.body.classList.remove('toc-mobile-open'); }
                  btn.addEventListener('click', toggleToc);
                  backdrop.addEventListener('click', closeMobileToc);
                  drawer.querySelectorAll('a.toc-item').forEach(a => a.addEventListener('click', () => {
                    const targetId = (a.getAttribute('href') || '').replace(/^#/, '');
                    if (targetId) setActiveToc(targetId, false);
                    if (isMobile()) closeMobileToc();
                  }));

                  // ScrollSpy inteligente com destaque (.active) no índice
                  const headings = Array.from(document.querySelectorAll('main h1[id], main h2[id], main h3[id]'));
                  const tocLinks = Array.from(drawer.querySelectorAll('a.toc-item[href^="#"]'));
                  let lastActiveId = '';
                  function setActiveToc(id, autoScroll) {
                    if (!id || id === lastActiveId) return;
                    lastActiveId = id;
                    tocLinks.forEach(link => {
                      const isMatch = link.getAttribute('href') === '#' + id;
                      link.classList.toggle('active', isMatch);
                      if (isMatch && autoScroll && !document.body.classList.contains('toc-collapsed')) {
                        const top = link.offsetTop;
                        if (top < drawer.scrollTop || top + link.offsetHeight > drawer.scrollTop + drawer.clientHeight) {
                          link.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                        }
                      }
                    });
                  }
                  function updateActiveSection() {
                    if (headings.length === 0 || tocLinks.length === 0) return;
                    const scrollY = window.scrollY || document.documentElement.scrollTop;
                    const docHeight = document.documentElement.scrollHeight;
                    let currentId = headings[0].id;
                    if (scrollY + window.innerHeight >= docHeight - 15) {
                      currentId = headings[headings.length - 1].id;
                    } else {
                      for (let i = 0; i < headings.length; i++) {
                        if (headings[i].getBoundingClientRect().top <= 110) {
                          currentId = headings[i].id;
                        } else {
                          break;
                        }
                      }
                    }
                    setActiveToc(currentId, true);
                  }
                  window.addEventListener('scroll', updateActiveSection, { passive: true });
                  window.addEventListener('resize', updateActiveSection, { passive: true });
                  updateActiveSection();
                  document.querySelectorAll('figure.img-card img').forEach(img => {
                    img.addEventListener('click', () => {
                      document.getElementById('imgLightboxTarget').src = img.src;
                      document.getElementById('imgLightbox').classList.add('open');
                    });
                  });
                })();
                function copyCode(btn) {
                  const code = btn.closest('.code-card').querySelector('code').innerText;
                  if (navigator.clipboard) navigator.clipboard.writeText(code).then(() => { btn.textContent = '✅ Copiado'; setTimeout(() => btn.textContent = '📋 Copiar', 1500); });
                }
              </script>
            </body>
            </html>
            """;
    }

    #region Auxiliares Internos

    private static bool IsSafeContainedPath(string allowedRoot, string? candidatePath, Func<string, bool> extensionPredicate)
    {
        if (string.IsNullOrWhiteSpace(allowedRoot) || string.IsNullOrWhiteSpace(candidatePath))
            return false;

        if (allowedRoot.IndexOfAny(['\0', '\r', '\n']) >= 0 || candidatePath.IndexOfAny(['\0', '\r', '\n']) >= 0)
            return false;

        try
        {
            string fullRoot = Path.GetFullPath(allowedRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string rootWithSep = fullRoot + Path.DirectorySeparatorChar;

            string fullCandidate = Path.GetFullPath(candidatePath);
            if (!fullCandidate.StartsWith(rootWithSep, StringComparison.OrdinalIgnoreCase))
                return false;

            string ext = Path.GetExtension(fullCandidate);
            if (string.IsNullOrEmpty(ext) || !extensionPredicate(ext))
                return false;

            var rootInfo = new DirectoryInfo(fullRoot);
            if (rootInfo.Exists && IsSymbolicLinkOrJunction(rootInfo))
                return false;

            var fileInfo = new FileInfo(fullCandidate);
            if (fileInfo.Exists)
            {
                if (IsSymbolicLinkOrJunction(fileInfo))
                    return false;

                var parent = fileInfo.Directory;
                while (parent != null && parent.FullName.Length > fullRoot.Length)
                {
                    if (IsSymbolicLinkOrJunction(parent))
                        return false;
                    parent = parent.Parent;
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsSymbolicLinkOrJunction(FileSystemInfo info)
    {
        // Em .NET 8 no Windows, LinkTarget é preenchido para symlinks e junctions NTFS,
        // mas permanece nulo para Cloud Files / OneDrive Files On-Demand (IO_REPARSE_TAG_CLOUD).
        try
        {
            return !string.IsNullOrEmpty(info.LinkTarget);
        }
        catch
        {
            return true;
        }
    }

    private static string NormalizeMarkdownHref(string rawHref)
    {
        string trimmed = (rawHref ?? string.Empty).Trim();
        if (trimmed.StartsWith('<') && trimmed.EndsWith('>') && trimmed.Length > 2)
            trimmed = trimmed[1..^1].Trim();
        return trimmed;
    }

    private static bool IsAttachmentHref(string rawHref)
    {
        string href = NormalizeMarkdownHref(rawHref);
        int hashIdx = href.IndexOf('#');
        if (hashIdx >= 0) href = href[..hashIdx];
        string ext = Path.GetExtension(href);
        return !string.IsNullOrEmpty(ext) && AllowedAttachmentExtensions.Contains(ext);
    }

    private static string RenderHtmlInlines(IReadOnlyList<MarkdownInlineSpan> spans)
    {
        var sb = new StringBuilder();
        foreach (var s in spans)
        {
            string enc = WebUtility.HtmlEncode(s.Text);
            switch (s.Kind)
            {
                case MarkdownInlineKind.Text:
                    sb.Append(enc);
                    break;
                case MarkdownInlineKind.Bold:
                    sb.Append($"<strong>{enc}</strong>");
                    break;
                case MarkdownInlineKind.Italic:
                    sb.Append($"<em>{enc}</em>");
                    break;
                case MarkdownInlineKind.Strikethrough:
                    sb.Append($"<del>{enc}</del>");
                    break;
                case MarkdownInlineKind.Code:
                    sb.Append($"<code>{enc}</code>");
                    break;
                case MarkdownInlineKind.Link:
                    if (InputValidator.IsSafeHttpUrl(s.TargetUrl))
                        sb.Append($"<a href=\"{WebUtility.HtmlEncode(s.TargetUrl)}\" target=\"_blank\" rel=\"noopener\">{enc}</a>");
                    else
                        sb.Append($"<span class=\"local-ref\">{enc}</span>");
                    break;
                case MarkdownInlineKind.Image:
                    sb.Append($"<span>🖼️ {enc}</span>");
                    break;
            }
        }
        return sb.ToString();
    }

    private static string RenderHtmlAttachmentCard(
        DocEntryItem doc,
        string label,
        string source,
        string? sharePointWebBaseUrl)
    {
        string ext = Path.GetExtension(source).ToLowerInvariant();
        var (emoji, badge, btnText) = ext switch
        {
            ".mp4" or ".mkv" or ".webm" or ".mov" or ".avi" => ("🎬", $"VÍDEO {ext.TrimStart('.').ToUpperInvariant()}", "▶ Assistir no SharePoint ↗"),
            ".pdf" => ("📄", "PDF", "📖 Abrir PDF no SharePoint ↗"),
            ".docx" => ("📝", "WORD", "📝 Abrir Documento no SharePoint ↗"),
            ".xlsx" => ("📊", "EXCEL", "📊 Abrir Planilha no SharePoint ↗"),
            _ => ("📎", "ANEXO", "📂 Abrir no SharePoint ↗")
        };

        string? spWebUrl = BuildSharePointWebUrl(doc, source, sharePointWebBaseUrl);
        string safeLabel = WebUtility.HtmlEncode(label);
        string safeSource = WebUtility.HtmlEncode(source);

        if (!string.IsNullOrEmpty(spWebUrl))
        {
            string safeHref = WebUtility.HtmlEncode(spWebUrl);
            return $"<div class=\"attach-card\"><div class=\"attach-info\"><span class=\"attach-badge\">{badge}</span><strong>{emoji} {safeLabel}</strong><div class=\"attach-path\">Disponível no SharePoint para visualização direta em campo</div></div><a class=\"attach-btn\" href=\"{safeHref}\" target=\"_blank\" rel=\"noopener\">{btnText}</a></div>";
        }

        return $"<div class=\"attach-card\"><div class=\"attach-info\"><span class=\"attach-badge\">{badge}</span><strong>{emoji} {safeLabel}</strong><div class=\"attach-path\">{safeSource}</div></div></div>";
    }

    private static string RenderHtmlImageBlock(string sourceRoot, string currentDocFullPath, string alt, string src)
    {
        string safeAlt = WebUtility.HtmlEncode(string.IsNullOrWhiteSpace(alt) ? Path.GetFileName(src) : alt);
        var resolved = ResolveMarkdownLinkTarget(sourceRoot, currentDocFullPath, src);

        if (resolved.Kind == MarkdownLinkKind.LocalImage && File.Exists(resolved.Target))
        {
            try
            {
                var fi = new FileInfo(resolved.Target);
                if (fi.Length <= MaxEmbeddedImageBytes)
                {
                    byte[] bytes = File.ReadAllBytes(resolved.Target);
                    string mime = GetImageMimeType(Path.GetExtension(resolved.Target));
                    string b64 = Convert.ToBase64String(bytes);
                    return $"<figure class=\"img-card\"><img src=\"data:{mime};base64,{b64}\" alt=\"{safeAlt}\" /><figcaption>{safeAlt}</figcaption></figure>";
                }
            }
            catch
            {
                // Fallback abaixo
            }
        }

        return $"<figure class=\"img-card\"><figcaption>🖼️ Imagem não embutida: {safeAlt} ({WebUtility.HtmlEncode(src)})</figcaption></figure>";
    }

    private static string GetImageMimeType(string ext) => ext.ToLowerInvariant() switch
    {
        ".png" => "image/png",
        ".jpg" or ".jpeg" => "image/jpeg",
        ".gif" => "image/gif",
        ".webp" => "image/webp",
        ".bmp" => "image/bmp",
        _ => "application/octet-stream"
    };

    private static IEnumerable<string> EnumerateMarkdownFilesSafe(DirectoryInfo currentDir, string allowedRoot)
    {
        FileInfo[] files;
        try
        {
            files = currentDir.GetFiles("*.md", SearchOption.TopDirectoryOnly);
        }
        catch
        {
            yield break;
        }

        foreach (var f in files)
        {
            if (IsSymbolicLinkOrJunction(f))
                continue;
            if (IsSafeDocPath(allowedRoot, f.FullName))
                yield return f.FullName;
        }

        DirectoryInfo[] subDirs;
        try
        {
            subDirs = currentDir.GetDirectories("*", SearchOption.TopDirectoryOnly);
        }
        catch
        {
            yield break;
        }

        foreach (var sub in subDirs)
        {
            if (IsSymbolicLinkOrJunction(sub))
                continue;
            foreach (string nested in EnumerateMarkdownFilesSafe(sub, allowedRoot))
                yield return nested;
        }
    }

    private static bool MatchesSearch(DocEntryItem doc, string query, string? fallbackDocsRoot)
    {
        if (doc.Title.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.RelativePath.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.CategoryLabel.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.SourceLabel.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.Summary.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        string effectiveRoot = !string.IsNullOrWhiteSpace(doc.SourceRootPath)
            ? doc.SourceRootPath
            : (fallbackDocsRoot ?? ResolveDefaultDocsDirectory());

        try
        {
            if (!IsSafeDocPath(effectiveRoot, doc.FullPath) || !File.Exists(doc.FullPath))
                return false;

            var fi = new FileInfo(doc.FullPath);
            if (fi.Length > MaxMarkdownFileBytes)
                return false;

            string content = File.ReadAllText(doc.FullPath, Encoding.UTF8);
            return content.Contains(query, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string ExtractTopFolderKey(string relativePath)
    {
        int slashIdx = relativePath.IndexOf('/');
        return slashIdx > 0 ? relativePath[..slashIdx].ToLowerInvariant() : "indice";
    }

    private static (string Label, string Badge, string BgHex, string FgHex, int Order) GetCategoryMetadata(
        string folderKey,
        bool isRootIndex,
        DocSourceKind sourceKind)
    {
        if (sourceKind == DocSourceKind.SharePointFaqs)
        {
            if (isRootIndex || folderKey == "indice")
                return ("Índice FAQs (SharePoint)", "FAQ ÍNDICE", "#0c4a6e", "#7dd3fc", 10);

            return (HumanizeIdentifier(folderKey), folderKey.ToUpperInvariant(), "#065f46", "#6ee7b7", 11);
        }

        if (isRootIndex || folderKey == "indice")
            return ("Índice Geral", "ÍNDICE", "#1e3a8a", "#93c5fd", 0);

        return folderKey switch
        {
            "guias" => ("Guias Operacionais", "GUIA", "#064e3b", "#34d399", 1),
            "arquitetura" => ("Arquitetura & SecOps", "ARQUITETURA", "#312e81", "#c7d2fe", 2),
            "issues" => ("Issues & Homologação", "ISSUES", "#78350f", "#fbbf24", 3),
            _ => (HumanizeIdentifier(folderKey), folderKey.ToUpperInvariant(), "#1e293b", "#94a3b8", 4)
        };
    }

    private static (string Emoji, string Title) ExtractDocumentTitle(
        string content,
        string fileName,
        bool isRootIndex,
        string folderKey,
        DocSourceKind sourceKind)
    {
        string defaultEmoji = sourceKind == DocSourceKind.SharePointFaqs
            ? (isRootIndex ? "🏢" : "📘")
            : isRootIndex ? "📚" : folderKey switch
            {
                "guias" => "📖",
                "arquitetura" => "🏛️",
                "issues" => "🚨",
                _ => "📄"
            };

        bool inFence = false;
        foreach (string rawLine in content.Replace("\r\n", "\n").Split('\n'))
        {
            string trimmed = rawLine.Trim();
            if (trimmed.StartsWith("```", StringComparison.Ordinal) || trimmed.StartsWith("~~~", StringComparison.Ordinal))
            {
                inFence = !inFence;
                continue;
            }
            if (inFence) continue;

            if (trimmed.StartsWith("# ", StringComparison.Ordinal))
            {
                string h1 = StripInlineFormatting(trimmed[2..].Trim());
                var (leadEmoji, cleanText) = SplitLeadingEmoji(h1, defaultEmoji);
                if (!string.IsNullOrWhiteSpace(cleanText))
                    return (leadEmoji, cleanText);
            }
        }

        string stem = Path.GetFileNameWithoutExtension(fileName);
        return (defaultEmoji, HumanizeIdentifier(stem));
    }

    private static (string Emoji, string CleanTitle) SplitLeadingEmoji(string heading, string fallbackEmoji)
    {
        if (string.IsNullOrWhiteSpace(heading))
            return (fallbackEmoji, heading);

        int firstSpace = heading.IndexOf(' ');
        if (firstSpace > 0 && firstSpace <= 8)
        {
            string token = heading[..firstSpace];
            bool hasAsciiAlphaNum = token.Any(char.IsLetterOrDigit);
            if (!hasAsciiAlphaNum)
            {
                return (token, heading[(firstSpace + 1)..].Trim());
            }
        }

        return (fallbackEmoji, heading);
    }

    private static string ExtractDocumentSummary(string content)
    {
        bool inFence = false;
        foreach (string rawLine in content.Replace("\r\n", "\n").Split('\n'))
        {
            string trimmed = rawLine.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            if (trimmed.StartsWith("```", StringComparison.Ordinal) || trimmed.StartsWith("~~~", StringComparison.Ordinal))
            {
                inFence = !inFence;
                continue;
            }
            if (inFence) continue;

            if (trimmed.StartsWith('#') ||
                SafeMatch(HrRegex, trimmed).Success ||
                SafeMatch(StandaloneImageRegex, trimmed).Success ||
                trimmed.StartsWith('|') ||
                SafeMatch(ListItemRegex, trimmed).Success)
            {
                continue;
            }

            if (trimmed.StartsWith('>'))
            {
                trimmed = trimmed.TrimStart('>', ' ');
                var alertMatch = SafeMatch(AlertTagRegex, trimmed);
                if (alertMatch.Success)
                    trimmed = alertMatch.Groups[2].Value.Trim();
            }

            string clean = StripInlineFormatting(trimmed);
            if (clean.Length < 8) continue;

            return clean.Length > 135 ? clean[..132].TrimEnd() + "..." : clean;
        }

        return string.Empty;
    }

    private static string HumanizeIdentifier(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "Documento";
        string spaced = raw.Replace('_', ' ').Replace('-', ' ').Trim();
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(spaced.ToLowerInvariant());
    }

    private static string StripInlineFormatting(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var sb = new StringBuilder();
        foreach (var span in ParseInlines(text))
        {
            sb.Append(span.Text);
        }
        return sb.ToString().Trim();
    }

    private static bool IsClosingFence(string line, char fenceChar, int minLen)
    {
        string trimmed = line.Trim();
        if (trimmed.Length < minLen) return false;
        foreach (char c in trimmed)
        {
            if (c != fenceChar) return false;
        }
        return true;
    }

    private static bool IsPotentialTableRow(string line)
    {
        string t = line.Trim();
        return t.Contains('|') && !t.StartsWith("```", StringComparison.Ordinal);
    }

    private static bool IsTableSeparatorLine(string line)
    {
        var cells = SplitTableCells(line);
        if (cells.Count == 0) return false;
        return cells.All(c => SafeMatch(TableSeparatorCellRegex, c.Trim()).Success);
    }

    private static List<string> SplitTableCells(string line)
    {
        string trimmed = line.Trim();
        if (trimmed.StartsWith('|')) trimmed = trimmed[1..];
        if (trimmed.EndsWith('|') && !trimmed.EndsWith(@"\|", StringComparison.Ordinal))
            trimmed = trimmed[..^1];

        var cells = new List<string>();
        var current = new StringBuilder();
        bool inInlineCode = false;

        for (int i = 0; i < trimmed.Length; i++)
        {
            char c = trimmed[i];
            if (c == '`')
            {
                inInlineCode = !inInlineCode;
                current.Append(c);
            }
            else if (c == '\\' && i + 1 < trimmed.Length && trimmed[i + 1] == '|')
            {
                current.Append('|');
                i++;
            }
            else if (c == '|' && !inInlineCode)
            {
                cells.Add(current.ToString().Trim());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        cells.Add(current.ToString().Trim());
        return cells;
    }

    private static Match SafeMatch(Regex rx, string input)
    {
        try
        {
            return rx.Match(input);
        }
        catch (RegexMatchTimeoutException)
        {
            return Match.Empty;
        }
    }

    #endregion
}
