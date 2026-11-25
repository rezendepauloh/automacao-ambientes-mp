using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo para descoberta, leitura segura (path-containment / anti-reparse-point),
/// filtragem e parsing de documentos Markdown (.md) da pasta Docs/.
/// </summary>
public class DocumentationService
{
    private const long MaxMarkdownFileBytes = 2 * 1024 * 1024; // 2 MB
    private static readonly TimeSpan RxTimeout = TimeSpan.FromMilliseconds(250);

    private static readonly Regex HeadingRegex = new(
        @"^(#{1,6})\s+(.+?)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex HrRegex = new(
        @"^\s*(?:-\s*-\s*-[-\s]*|\*\s*\*\s*\*[\*\s]*|_\s*_\s*_[_\s]*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RxTimeout);

    private static readonly Regex FenceOpenRegex = new(
        @"^\s*(`{3,}|~{3,})\s*([^\s`~]*).*$",
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

    // Tokenizador inline: código em crases, links [txt](url), **negrito**, __negrito__, ~~riscado~~, *itálico* e _itálico_ seguro (preserva snake_case).
    private static readonly Regex InlineTokenRegex = new(
        @"(?<code>`[^`\r\n]+`)" +
        @"|(?<link>\[(?<ltext>[^\]\r\n]+)\]\((?<lurl>[^\)\r\n]+)\))" +
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
    /// Garante que o arquivo seja estritamente um .md contido dentro de docsRoot, sem caracteres de controle
    /// e sem atravessar junctions/symlinks (ReparsePoint).
    /// </summary>
    public static bool IsSafeDocPath(string docsRoot, string? candidatePath)
    {
        if (string.IsNullOrWhiteSpace(docsRoot) || string.IsNullOrWhiteSpace(candidatePath))
            return false;

        if (docsRoot.IndexOfAny(['\0', '\r', '\n']) >= 0 || candidatePath.IndexOfAny(['\0', '\r', '\n']) >= 0)
            return false;

        try
        {
            string fullRoot = Path.GetFullPath(docsRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string rootWithSep = fullRoot + Path.DirectorySeparatorChar;

            string fullCandidate = Path.GetFullPath(candidatePath);
            if (!fullCandidate.StartsWith(rootWithSep, StringComparison.OrdinalIgnoreCase))
                return false;

            if (!fullCandidate.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
                return false;

            var rootInfo = new DirectoryInfo(fullRoot);
            if (rootInfo.Exists && (rootInfo.Attributes & FileAttributes.ReparsePoint) != 0)
                return false;

            var fileInfo = new FileInfo(fullCandidate);
            if (fileInfo.Exists)
            {
                if ((fileInfo.Attributes & FileAttributes.ReparsePoint) != 0)
                    return false;

                var parent = fileInfo.Directory;
                while (parent != null && parent.FullName.Length > fullRoot.Length)
                {
                    if ((parent.Attributes & FileAttributes.ReparsePoint) != 0)
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

    public List<DocEntryItem> DiscoverDocuments(string? customDocsRoot = null)
    {
        string docsRoot = string.IsNullOrWhiteSpace(customDocsRoot)
            ? ResolveDefaultDocsDirectory()
            : Path.GetFullPath(customDocsRoot);

        var results = new List<DocEntryItem>();
        if (!Directory.Exists(docsRoot))
            return results;

        var rootDirInfo = new DirectoryInfo(docsRoot);
        if ((rootDirInfo.Attributes & FileAttributes.ReparsePoint) != 0)
            return results;

        foreach (string filePath in EnumerateMarkdownFilesSafe(rootDirInfo, docsRoot))
        {
            try
            {
                var fi = new FileInfo(filePath);
                if (fi.Length > MaxMarkdownFileBytes)
                    continue;

                string content = File.ReadAllText(filePath, Encoding.UTF8);
                string relPath = Path.GetRelativePath(docsRoot, filePath).Replace('\\', '/');
                bool isRootIndex = relPath.Equals("README.md", StringComparison.OrdinalIgnoreCase);

                string folderKey = ExtractTopFolderKey(relPath);
                var (catLabel, badgeText, badgeBg, badgeFg, catOrder) = GetCategoryMetadata(folderKey, isRootIndex);
                var (emoji, title) = ExtractDocumentTitle(content, fi.Name, isRootIndex, folderKey);
                string summary = ExtractDocumentSummary(content);

                results.Add(new DocEntryItem
                {
                    Title = title,
                    Emoji = emoji,
                    FileName = fi.Name,
                    RelativePath = relPath,
                    FullPath = fi.FullName,
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
            throw new UnauthorizedAccessException("Caminho de documento fora do diretório permitido (Docs/).");

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
        string docsRoot)
    {
        IEnumerable<DocEntryItem> query = allDocs;

        if (!string.IsNullOrWhiteSpace(categoryLabel) &&
            !categoryLabel.Equals("Todas as Categorias", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(d => d.CategoryLabel.Equals(categoryLabel, StringComparison.OrdinalIgnoreCase));
        }

        string trimmedSearch = (searchQuery ?? string.Empty).Trim();
        if (!string.IsNullOrEmpty(trimmedSearch))
        {
            query = query.Where(d => MatchesSearch(d, trimmedSearch, docsRoot));
        }

        return query.ToList();
    }

    public static MarkdownLinkResolution ResolveMarkdownLinkTarget(
        string docsRoot,
        string currentDocFullPath,
        string? rawHref)
    {
        if (string.IsNullOrWhiteSpace(rawHref))
            return new MarkdownLinkResolution(MarkdownLinkKind.UnsupportedReference, string.Empty);

        string href = rawHref.Trim();
        if (href.IndexOfAny(['\0', '\r', '\n']) >= 0)
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
            return IsSafeDocPath(docsRoot, currentDocFullPath)
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
                string baseDir = Path.GetDirectoryName(Path.GetFullPath(currentDocFullPath)) ?? docsRoot;
                string normalizedRel = Uri.UnescapeDataString(href).Replace('/', Path.DirectorySeparatorChar);
                candidateFullPath = Path.GetFullPath(Path.Combine(baseDir, normalizedRel));
            }

            if (IsSafeDocPath(docsRoot, candidateFullPath) && File.Exists(candidateFullPath))
            {
                return new MarkdownLinkResolution(MarkdownLinkKind.InternalDocument, candidateFullPath, anchor);
            }

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

            // 3. Títulos (# a ######)
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

            // 4. Blockquote / Alertas GitHub (> ...)
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

            // 5. Tabela Markdown (| Col1 | Col2 | seguida por linha separadora | --- | --- |)
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

            // 6. Listas ordenadas, não ordenadas e checklists
            var listMatch = SafeMatch(ListItemRegex, line);
            if (listMatch.Success)
            {
                var items = new List<MarkdownListItem>();
                while (i < lines.Length)
                {
                    if (string.IsNullOrWhiteSpace(lines[i]))
                    {
                        // Se a próxima linha não vazia continuar a lista, consome a linha em branco
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

            // 7. Parágrafo comum (agrupa linhas adjacentes simples)
            var paraLines = new List<string>();
            while (i < lines.Length &&
                   !string.IsNullOrWhiteSpace(lines[i]) &&
                   !SafeMatch(FenceOpenRegex, lines[i]).Success &&
                   !SafeMatch(HrRegex, lines[i]).Success &&
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
                else if (m.Groups["link"].Success)
                {
                    string linkText = StripInlineFormatting(m.Groups["ltext"].Value);
                    string linkUrl = m.Groups["lurl"].Value.Trim();
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Link, linkText, linkUrl));
                }
                else if (m.Groups["bold1"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Bold, m.Groups["b1text"].Value));
                }
                else if (m.Groups["bold2"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Bold, m.Groups["b2text"].Value));
                }
                else if (m.Groups["strike"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Strikethrough, m.Groups["stext"].Value));
                }
                else if (m.Groups["ital1"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Italic, m.Groups["i1text"].Value));
                }
                else if (m.Groups["ital2"].Success)
                {
                    spans.Add(new MarkdownInlineSpan(MarkdownInlineKind.Italic, m.Groups["i2text"].Value));
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

    #region Auxiliares Internos

    private static IEnumerable<string> EnumerateMarkdownFilesSafe(DirectoryInfo currentDir, string docsRoot)
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
            if ((f.Attributes & FileAttributes.ReparsePoint) != 0)
                continue;
            if (IsSafeDocPath(docsRoot, f.FullName))
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
            if ((sub.Attributes & FileAttributes.ReparsePoint) != 0)
                continue;
            foreach (string nested in EnumerateMarkdownFilesSafe(sub, docsRoot))
                yield return nested;
        }
    }

    private static bool MatchesSearch(DocEntryItem doc, string query, string docsRoot)
    {
        if (doc.Title.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.RelativePath.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.CategoryLabel.Contains(query, StringComparison.OrdinalIgnoreCase) ||
            doc.Summary.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        try
        {
            if (!IsSafeDocPath(docsRoot, doc.FullPath) || !File.Exists(doc.FullPath))
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
        bool isRootIndex)
    {
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
        string folderKey)
    {
        string defaultEmoji = isRootIndex ? "📚" : folderKey switch
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
