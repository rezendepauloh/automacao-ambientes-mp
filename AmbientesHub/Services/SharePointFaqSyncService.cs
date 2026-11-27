using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace AmbientesHub.Services;

public sealed record SharePointSyncResult(
    int PagesConverted,
    int LocalImagesMapped,
    int LocalFilesMapped,
    int SharePointFallbackLinks,
    string OutputDirectory);

/// <summary>
/// Sincronizador nativo C# de páginas SitePages do SharePoint (feed Atom XML) para Markdown
/// em SHAREPOINT_FAQS_DIR\Manuais Markdown\, preservando imagens locais (Imagens FAQ + SiteAssets),
/// cards de vídeo/anexos sob demanda e links cruzados entre FAQs.
/// </summary>
public class SharePointFaqSyncService
{
    private const long MaxFeedBytes = 25 * 1024 * 1024; // 25 MB
    private static readonly TimeSpan RxTimeout = TimeSpan.FromMilliseconds(350);

    private static readonly HashSet<string> IgnoredSystemPages = new(StringComparer.OrdinalIgnoreCase)
    {
        "Home.aspx",
        "CollabHome.aspx",
        "Frequently-asked-questions.aspx",
        "Links-para-FAQs.aspx"
    };

    private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"
    };

    /// <summary>
    /// Deriva a URL da REST API de SitePages a partir de SHAREPOINT_FAQS_WEB_URL (fail-closed se ausente/inválida).
    /// </summary>
    public static string? BuildSitePagesApiUrl(string? sharePointFaqsWebUrl = null)
    {
        string raw = (sharePointFaqsWebUrl ?? ConfigService.SharePointFaqsWebUrl).Trim();
        if (!InputValidator.IsSafeHttpUrl(raw))
            return null;

        if (!Uri.TryCreate(raw, UriKind.Absolute, out var uri))
            return null;

        string path = Uri.UnescapeDataString(uri.AbsolutePath);
        string sitePath = ExtractSiteServerRelativePath(path);
        if (string.IsNullOrEmpty(sitePath))
            return null;

        string origin = uri.GetLeftPart(UriPartial.Authority);
        string apiUrl = $"{origin}{sitePath}/_api/web/GetList('{sitePath}/SitePages')/items?$select=Id,Title,FileLeafRef,FileRef,CanvasContent1,Description,Modified&$top=500";
        return InputValidator.IsSafeHttpUrl(apiUrl) ? apiUrl : null;
    }

    public SharePointSyncResult SyncFromFeedFile(string feedFilePath, string sharePointFaqsDir, string? sharePointFaqsWebUrl = null)
    {
        if (string.IsNullOrWhiteSpace(feedFilePath) || !File.Exists(feedFilePath))
            throw new FileNotFoundException("Arquivo XML do SharePoint não encontrado.", feedFilePath);

        var fi = new FileInfo(feedFilePath);
        if (fi.Length > MaxFeedBytes)
            throw new InvalidOperationException("O arquivo de feed excede o limite de 25 MB.");

        string content = File.ReadAllText(feedFilePath, Encoding.UTF8);
        return SyncFromFeedXml(content, sharePointFaqsDir, sharePointFaqsWebUrl);
    }

    public SharePointSyncResult SyncFromFeedXml(string rawFeedContent, string sharePointFaqsDir, string? sharePointFaqsWebUrl = null)
    {
        if (!InputValidator.IsSafeDirectoryPath(sharePointFaqsDir) || !Directory.Exists(sharePointFaqsDir))
            throw new DirectoryNotFoundException("Diretório SHAREPOINT_FAQS_DIR inválido ou não encontrado.");

        if (string.IsNullOrWhiteSpace(rawFeedContent))
            throw new InvalidDataException("Conteúdo XML vazio.");

        string xmlSlice = ExtractXmlFeedSlice(rawFeedContent);
        var entries = ParseAtomFeedEntries(xmlSlice);
        if (entries.Count == 0)
            throw new InvalidDataException("Nenhuma página SitePages encontrada no feed XML.");

        string origin = ResolveSiteOrigin(sharePointFaqsWebUrl);
        string outRoot = Path.Combine(Path.GetFullPath(sharePointFaqsDir), "Manuais Markdown");
        Directory.CreateDirectory(outRoot);

        // Indexa apenas nomes dos arquivos locais (sem ler/hidratar mídia do OneDrive)
        var (localFilesByRel, localImgsByName) = IndexLocalMediaMetadata(sharePointFaqsDir);

        // Descobre categorias a partir de Links-para-FAQs.aspx (se presente no feed)
        var linksPage = entries.FirstOrDefault(p => p.Leaf.Equals("Links-para-FAQs.aspx", StringComparison.OrdinalIgnoreCase));
        var leafToCategory = linksPage != null
            ? ExtractCategoriesFromLinksPage(linksPage.Canvas)
            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var faqEntries = entries
            .Where(p => !IgnoredSystemPages.Contains(p.Leaf) && !string.IsNullOrWhiteSpace(p.Title))
            .ToList();

        // Monta mapa leaf -> cat/slug.md para reescrever links cruzados entre FAQs
        var leafToTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var entryMeta = new List<(SitePageRecord Page, string Category, string Slug)>();

        foreach (var page in faqEntries)
        {
            string cat = leafToCategory.TryGetValue(page.Leaf, out var mappedCat)
                ? mappedCat
                : InferCategoryFromTitle(page.Title);
            string slug = Slugify(page.Title);
            leafToTarget[page.Leaf] = $"{cat}/{slug}.md";
            entryMeta.Add((page, cat, slug));
        }

        int totalLocalImages = 0;
        int totalLocalFiles = 0;
        int totalSpFallback = 0;
        var readmeByCat = new Dictionary<string, List<(string Title, string RelPath)>>(StringComparer.OrdinalIgnoreCase);

        foreach (var (page, cat, slug) in entryMeta)
        {
            string catDir = Path.Combine(outRoot, cat);
            Directory.CreateDirectory(catDir);

            var stats = new ConversionStats();
            string body = ConvertCanvasToMarkdown(page.Canvas, cat, origin, leafToTarget, localFilesByRel, localImgsByName, stats);

            string spUrl = string.IsNullOrWhiteSpace(origin) || string.IsNullOrWhiteSpace(page.FileRef)
                ? string.Empty
                : origin + page.FileRef.Replace(" ", "%20");
            string modShort = page.Modified.Length >= 10 ? page.Modified[..10] : string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine($"# 📘 {page.Title}");
            sb.AppendLine();
            if (!string.IsNullOrEmpty(spUrl))
            {
                string modPart = string.IsNullOrEmpty(modShort) ? string.Empty : $" • Atualizado no SharePoint em `{modShort}`";
                sb.AppendLine($"> **Origem SharePoint:** [Abrir página original]({spUrl}){modPart}");
                sb.AppendLine();
            }
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine(string.IsNullOrWhiteSpace(body)
                ? "_Página sem conteúdo textual extraível — consulte o link original acima._"
                : body);

            string outFile = Path.Combine(catDir, $"{slug}.md");
            File.WriteAllText(outFile, sb.ToString(), Encoding.UTF8);

            if (!readmeByCat.TryGetValue(cat, out var list))
            {
                list = [];
                readmeByCat[cat] = list;
            }
            list.Add((page.Title, $"{cat}/{slug}.md"));

            totalLocalImages += stats.LocalImages;
            totalLocalFiles += stats.LocalFiles;
            totalSpFallback += stats.SharePointFallback;
        }

        WriteReadmeIndex(outRoot, readmeByCat, origin, linksPage?.FileRef);

        AppLoggerService.Success(
            $"Sincronização SharePoint concluída: {entryMeta.Count} FAQs, {totalLocalImages} imagens locais, {totalLocalFiles} vídeos/anexos.",
            nameof(SharePointFaqSyncService));

        return new SharePointSyncResult(entryMeta.Count, totalLocalImages, totalLocalFiles, totalSpFallback, outRoot);
    }

    #region XML & WebPart Parsing

    private sealed record SitePageRecord(int Id, string Title, string Leaf, string FileRef, string Modified, string Canvas);

    private sealed class ConversionStats
    {
        public int LocalImages;
        public int LocalFiles;
        public int SharePointFallback;
    }

    private static string ExtractXmlFeedSlice(string raw)
    {
        int start = raw.IndexOf("<?xml", StringComparison.OrdinalIgnoreCase);
        if (start < 0) start = raw.IndexOf("<feed", StringComparison.OrdinalIgnoreCase);
        int end = raw.LastIndexOf("</feed>", StringComparison.OrdinalIgnoreCase);
        if (start < 0 || end < 0 || end < start)
            throw new InvalidDataException("O arquivo não contém um feed Atom XML válido (<feed>...</feed>).");
        return raw[start..(end + "</feed>".Length)];
    }

    private static List<SitePageRecord> ParseAtomFeedEntries(string xmlContent)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true
        };

        using var sr = new StringReader(xmlContent);
        using var reader = XmlReader.Create(sr, settings);
        var xdoc = XDocument.Load(reader);

        XNamespace atom = "http://www.w3.org/2005/Atom";
        XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";
        XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";

        var list = new List<SitePageRecord>();
        foreach (var entry in xdoc.Descendants(atom + "entry"))
        {
            var props = entry.Element(atom + "content")?.Element(m + "properties");
            if (props == null) continue;

            int.TryParse(props.Element(d + "Id")?.Value, out int id);
            string title = (props.Element(d + "Title")?.Value ?? string.Empty).Trim();
            string leaf = (props.Element(d + "FileLeafRef")?.Value ?? string.Empty).Trim();
            string fileRef = (props.Element(d + "FileRef")?.Value ?? string.Empty).Trim();
            string modified = (props.Element(d + "Modified")?.Value ?? string.Empty).Trim();
            string canvas = props.Element(d + "CanvasContent1")?.Value ?? string.Empty;

            if (!string.IsNullOrEmpty(leaf))
                list.Add(new SitePageRecord(id, title, leaf, fileRef, modified, canvas));
        }

        return list;
    }

    private static Dictionary<string, string> ExtractCategoriesFromLinksPage(string canvas)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string currentCat = "geral";

        foreach (var (kind, payload) in ExtractBalancedControls(canvas))
        {
            if (kind != "rte") continue;

            // Detecta cabeçalhos de seção e links SitePages/*.aspx na ordem do HTML
            foreach (Match m in Regex.Matches(payload, @"<h[1-4][^>]*>(?<h>.*?)</h[1-4]>|href=""(?<href>[^""]*SitePages/[^""]+\.aspx)""", RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout))
            {
                if (m.Groups["h"].Success)
                {
                    string hText = StripHtmlTags(WebUtility.HtmlDecode(m.Groups["h"].Value)).Trim();
                    if (!string.IsNullOrWhiteSpace(hText))
                        currentCat = Slugify(hText);
                }
                else if (m.Groups["href"].Success)
                {
                    string decoded = Uri.UnescapeDataString(WebUtility.HtmlDecode(m.Groups["href"].Value));
                    int idx = decoded.LastIndexOf('/');
                    string leaf = idx >= 0 ? decoded[(idx + 1)..] : decoded;
                    if (!string.IsNullOrWhiteSpace(leaf) && !map.ContainsKey(leaf))
                        map[leaf] = currentCat;
                }
            }
        }

        return map;
    }

    private static List<(string Kind, string Payload)> ExtractBalancedControls(string canvas)
    {
        var results = new List<(string, string)>();
        if (string.IsNullOrEmpty(canvas)) return results;

        var startRx = new Regex(@"<div\b[^>]*(?:data-sp-rte|data-sp-webpartdata=""[^""]+"")[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout);
        var divRx = new Regex(@"<div\b[^>]*>|</div>", RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout);
        var wpAttrRx = new Regex(@"data-sp-webpartdata=""([^""]+)""", RegexOptions.IgnoreCase, RxTimeout);

        int pos = 0;
        while (pos < canvas.Length)
        {
            var m = startRx.Match(canvas, pos);
            if (!m.Success) break;

            string openTag = m.Value;
            int tagEnd = m.Index + m.Length;
            int depth = 1;
            int closeStart = canvas.Length;
            int closeEnd = canvas.Length;

            foreach (Match tm in divRx.Matches(canvas, tagEnd))
            {
                if (tm.Value.StartsWith("</div", StringComparison.OrdinalIgnoreCase))
                {
                    depth--;
                    if (depth == 0)
                    {
                        closeStart = tm.Index;
                        closeEnd = tm.Index + tm.Length;
                        break;
                    }
                }
                else
                {
                    depth++;
                }
            }

            var wpMatch = wpAttrRx.Match(openTag);
            if (wpMatch.Success)
                results.Add(("wp", wpMatch.Groups[1].Value));
            else
                results.Add(("rte", canvas[tagEnd..closeStart]));

            pos = closeEnd;
        }

        return results;
    }

    private static string ConvertCanvasToMarkdown(
        string canvas,
        string currentCat,
        string origin,
        Dictionary<string, string> leafToTarget,
        Dictionary<string, string> localFilesByRel,
        Dictionary<string, string> localImgsByName,
        ConversionStats stats)
    {
        var sections = new List<string>();

        foreach (var (kind, payload) in ExtractBalancedControls(canvas))
        {
            if (kind == "rte")
            {
                string md = ConvertRteHtmlToMarkdown(payload, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName, stats);
                if (!string.IsNullOrWhiteSpace(md))
                    sections.Add(md);
            }
            else if (kind == "wp")
            {
                string wpMd = ConvertWebPartToMarkdown(payload, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName, stats);
                if (!string.IsNullOrWhiteSpace(wpMd))
                    sections.Add(wpMd);
            }
        }

        string joined = string.Join("\n\n", sections);
        return Regex.Replace(joined, @"\n{3,}", "\n\n").Trim();
    }

    private static string ConvertWebPartToMarkdown(
        string encodedJson,
        string currentCat,
        string origin,
        Dictionary<string, string> leafToTarget,
        Dictionary<string, string> localFilesByRel,
        Dictionary<string, string> localImgsByName,
        ConversionStats stats)
    {
        try
        {
            string json = WebUtility.HtmlDecode(encodedJson);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            string wid = (root.TryGetProperty("id", out var idEl) ? idEl.GetString() : string.Empty)?.ToLowerInvariant() ?? string.Empty;

            root.TryGetProperty("serverProcessedContent", out var spc);
            root.TryGetProperty("properties", out var props);

            // Markdown WebPart
            if (wid == "1ef5ed11-ce7b-44be-bc5e-4abd55101d16")
            {
                if (props.ValueKind == JsonValueKind.Object && props.TryGetProperty("markdown", out var mdProp))
                {
                    string rawMd = mdProp.GetString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(rawMd))
                        return rawMd.Trim();
                }
                string htmlFallback = TryGetNestedString(spc, "htmlStrings", "html");
                if (!string.IsNullOrWhiteSpace(htmlFallback))
                    return ConvertRteHtmlToMarkdown(htmlFallback, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName, stats);
            }

            // Image WebPart
            if (wid == "d1d91016-032f-456d-98a4-721247c305e8")
            {
                string src = TryGetNestedString(spc, "imageSources", "imageSource");
                if (string.IsNullOrWhiteSpace(src) && props.ValueKind == JsonValueKind.Object && props.TryGetProperty("imageSourceUrl", out var isu))
                    src = isu.GetString() ?? string.Empty;

                string caption = TryGetNestedString(spc, "searchablePlainTexts", "captionText");
                if (string.IsNullOrWhiteSpace(caption))
                    caption = Path.GetFileName(Uri.UnescapeDataString(src.Split('?')[0]));

                return FormatResolvedImage(src, caption, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName, stats);
            }

            // File & Media WebPart (vídeos MP4, PDFs, DOCX)
            if (wid == "b7dd04e1-19ce-4b24-9132-b60a1c2b910d")
            {
                string src = TryGetNestedString(spc, "links", "serverRelativeUrl");
                if (string.IsNullOrWhiteSpace(src)) src = TryGetNestedString(spc, "links", "wopiurl");
                if (string.IsNullOrWhiteSpace(src) && props.ValueKind == JsonValueKind.Object && props.TryGetProperty("file", out var fEl))
                    src = fEl.GetString() ?? string.Empty;

                string title = TryGetNestedString(spc, "searchablePlainTexts", "title");
                var (rk, target) = ResolveSharePointUrl(src, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName);
                if (rk == "local_file")
                {
                    stats.LocalFiles++;
                    string niceLabel = string.IsNullOrWhiteSpace(title) || Regex.IsMatch(title, @"_[0-9a-f]{8,}$")
                        ? Path.GetFileNameWithoutExtension(target)
                        : title;
                    return $"[{niceLabel}](<{target}>)";
                }
                if (rk == "local_image")
                {
                    stats.LocalImages++;
                    return $"![{(string.IsNullOrWhiteSpace(title) ? Path.GetFileName(target) : title)}](<{target}>)";
                }
                if (!string.IsNullOrWhiteSpace(target))
                    return $"- 📎 [{(string.IsNullOrWhiteSpace(title) ? "Arquivo" : title)}]({target})";
            }

            // YouTube WebPart
            if (wid == "544dd15b-cf3c-441b-96da-004d5a8cea1d")
            {
                string ytUrl = props.ValueKind == JsonValueKind.Object && props.TryGetProperty("embedCode", out var ec) ? ec.GetString() ?? string.Empty : string.Empty;
                string ytTitle = TryGetNestedString(spc, "searchablePlainTexts", "title");
                if (!string.IsNullOrWhiteSpace(ytUrl))
                    return $"- ▶️ **YouTube:** [{(string.IsNullOrWhiteSpace(ytTitle) ? "Vídeo no YouTube" : ytTitle)}]({ytUrl})";
            }

            // Link WebPart
            if (wid == "6410b3b6-d440-4663-8744-378976dc041e")
            {
                string url = TryGetNestedString(spc, "links", "url");
                if (!string.IsNullOrWhiteSpace(url))
                    return $"- 🔗 [{url}]({url})";
            }

            // Divisor
            if (wid == "2161a1c6-db61-4731-b97c-3cdb303f7cbb")
                return "---";
        }
        catch
        {
            // Ignora webpart malformado sem abortar os demais
        }

        return string.Empty;
    }

    private static string ConvertRteHtmlToMarkdown(
        string rteHtml,
        string currentCat,
        string origin,
        Dictionary<string, string> leafToTarget,
        Dictionary<string, string> localFilesByRel,
        Dictionary<string, string> localImgsByName,
        ConversionStats stats)
    {
        if (string.IsNullOrWhiteSpace(rteHtml)) return string.Empty;

        string s = rteHtml;

        // 1. Imagens inline do SharePoint Modern RTE (<div ... data-imageurl="...">)
        s = Regex.Replace(s, @"<div\b[^>]*data-imageurl=""(?<src>[^""]+)""[^>]*>", m =>
        {
            string src = WebUtility.HtmlDecode(m.Groups["src"].Value);
            string cap = Path.GetFileName(Uri.UnescapeDataString(src.Split('?')[0]));
            return "\n\n" + FormatResolvedImage(src, cap, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName, stats, escapeAnglesForHtmlPass: true) + "\n\n";
        }, RegexOptions.IgnoreCase, RxTimeout);

        // 2. Tags <img src="...">
        s = Regex.Replace(s, @"<img\b[^>]*src=""(?<src>[^""]+)""[^>]*>", m =>
        {
            string src = WebUtility.HtmlDecode(m.Groups["src"].Value);
            string cap = Path.GetFileName(Uri.UnescapeDataString(src.Split('?')[0]));
            return "\n\n" + FormatResolvedImage(src, cap, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName, stats, escapeAnglesForHtmlPass: true) + "\n\n";
        }, RegexOptions.IgnoreCase, RxTimeout);

        // 3. Blocos <pre>...</pre>
        s = Regex.Replace(s, @"<pre\b[^>]*>(?<inner>.*?)</pre>", m =>
        {
            string inner = Regex.Replace(m.Groups["inner"].Value, @"<br\s*/?>|</p>|</div>", "\n", RegexOptions.IgnoreCase);
            inner = WebUtility.HtmlDecode(StripHtmlTags(inner)).Replace("\u00a0", " ").Trim();
            return $"\n\n```text\n{inner}\n```\n\n";
        }, RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout);

        // 4. Títulos h1-h4
        s = Regex.Replace(s, @"<h(?<lvl>[1-4])\b[^>]*>(?<inner>.*?)</h\1>", m =>
        {
            int lvl = Math.Clamp(int.Parse(m.Groups["lvl"].Value), 2, 4);
            string txt = WebUtility.HtmlDecode(StripHtmlTags(m.Groups["inner"].Value)).Trim();
            return string.IsNullOrWhiteSpace(txt) ? string.Empty : $"\n\n{new string('#', lvl)} {txt}\n\n";
        }, RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout);

        // 5. Links <a href="...">texto</a>
        s = Regex.Replace(s, @"<a\b[^>]*href=""(?<href>[^""]+)""[^>]*>(?<inner>.*?)</a>", m =>
        {
            string href = WebUtility.HtmlDecode(m.Groups["href"].Value);
            string label = WebUtility.HtmlDecode(StripHtmlTags(m.Groups["inner"].Value)).Trim();
            if (string.IsNullOrWhiteSpace(label)) label = href;

            var (rk, target) = ResolveSharePointUrl(href, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName);
            if (rk == "local_file") { stats.LocalFiles++; return $"\n\n[{label}](\uE001{target}\uE002)\n\n"; }
            if (rk == "local_image") { stats.LocalImages++; return $"\n\n![{label}](\uE001{target}\uE002)\n\n"; }
            if (rk == "internal_md") return $"[{label}]({target})";
            if (rk == "anchor") return label;
            return string.IsNullOrWhiteSpace(target) ? label : $"[{label}](\uE001{target}\uE002)";
        }, RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout);

        // 6. Listas e parágrafos
        s = Regex.Replace(s, @"<li\b[^>]*>(?<inner>.*?)</li>", m =>
        {
            string txt = WebUtility.HtmlDecode(StripHtmlTags(m.Groups["inner"].Value)).Trim();
            return string.IsNullOrWhiteSpace(txt) ? string.Empty : $"\n- {txt}";
        }, RegexOptions.IgnoreCase | RegexOptions.Singleline, RxTimeout);

        s = Regex.Replace(s, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"</p>|</div>|</tr>", "\n\n", RegexOptions.IgnoreCase);
        s = WebUtility.HtmlDecode(StripHtmlTags(s))
            .Replace("\uE001", "<")
            .Replace("\uE002", ">")
            .Replace("\u00a0", " ")
            .Replace("\u200b", "");

        var lines = s.Split('\n').Select(l => l.TrimEnd());
        return Regex.Replace(string.Join("\n", lines), @"\n{3,}", "\n\n").Trim();
    }

    private static string FormatResolvedImage(
        string src,
        string caption,
        string currentCat,
        string origin,
        Dictionary<string, string> leafToTarget,
        Dictionary<string, string> localFilesByRel,
        Dictionary<string, string> localImgsByName,
        ConversionStats stats,
        bool escapeAnglesForHtmlPass = false)
    {
        string lt = escapeAnglesForHtmlPass ? "\uE001" : "<";
        string gt = escapeAnglesForHtmlPass ? "\uE002" : ">";
        var (rk, target) = ResolveSharePointUrl(src, currentCat, origin, leafToTarget, localFilesByRel, localImgsByName);
        string safeCap = string.IsNullOrWhiteSpace(caption) ? "Captura de tela" : caption;
        if (rk == "local_image")
        {
            stats.LocalImages++;
            return $"![{safeCap}]({lt}{target}{gt})";
        }
        if (!string.IsNullOrWhiteSpace(target))
        {
            stats.SharePointFallback++;
            return $"> 🖼️ **Captura no SharePoint:** [{safeCap}]({lt}{target}{gt})";
        }
        return string.Empty;
    }

    private static (string Kind, string Target) ResolveSharePointUrl(
        string rawUrl,
        string currentCat,
        string origin,
        Dictionary<string, string> leafToTarget,
        Dictionary<string, string> localFilesByRel,
        Dictionary<string, string> localImgsByName)
    {
        if (string.IsNullOrWhiteSpace(rawUrl)) return ("none", string.Empty);

        string u = rawUrl.Trim();
        if (u.StartsWith('#')) return ("anchor", u);
        string uDec = Uri.UnescapeDataString(u);

        // 1. Outro FAQ em SitePages/*.aspx
        var spMatch = Regex.Match(uDec, @"/SitePages/([^?#]+\.aspx)", RegexOptions.IgnoreCase);
        if (spMatch.Success)
        {
            string leaf = spMatch.Groups[1].Value.Trim();
            if (leafToTarget.TryGetValue(leaf, out var targetRel))
            {
                string[] parts = targetRel.Split('/', 2);
                string relMd = parts[0].Equals(currentCat, StringComparison.OrdinalIgnoreCase)
                    ? parts[1]
                    : $"../{targetRel}";
                return ("internal_md", relMd);
            }
        }

        // 2. Mídia dentro de Tutoriais-FAQs/
        const string marker = "Tutoriais-FAQs/";
        int idx = uDec.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
        {
            string subRel = uDec[(idx + marker.Length)..].Split('?')[0].Split('#')[0].TrimStart('/');
            if (localFilesByRel.TryGetValue(subRel, out var exactRel))
            {
                string ext = Path.GetExtension(exactRel);
                return ImageExtensions.Contains(ext)
                    ? ("local_image", $"../../{exactRel}")
                    : ("local_file", $"../../{exactRel}");
            }
        }

        // 3. Mídia em SiteAssets/SitePages/<subpath> -> Imagens FAQ/SiteAssets/<subpath>
        var saMatch = Regex.Match(uDec, @"/SiteAssets/SitePages/([^?#]+)", RegexOptions.IgnoreCase);
        if (saMatch.Success)
        {
            string saSub = saMatch.Groups[1].Value.TrimStart('/');
            string saRel = $"Imagens FAQ/SiteAssets/{saSub}";
            if (localFilesByRel.TryGetValue(saRel, out var exactSa))
                return ("local_image", $"../../{exactSa}");

            string fname = Path.GetFileName(saSub);
            if (localImgsByName.TryGetValue(fname, out var exactByName))
                return ("local_image", $"../../{exactByName}");
        }

        if (u.StartsWith('/') && !string.IsNullOrEmpty(origin))
            return ("external", origin + u);

        return ("external", u);
    }

    private static (Dictionary<string, string> ByRel, Dictionary<string, string> ByName) IndexLocalMediaMetadata(string sharePointFaqsDir)
    {
        var byRel = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var byName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string rootFull = Path.GetFullPath(sharePointFaqsDir);

        foreach (string sub in new[] { "Imagens FAQ", "Vídeos FAQ", "Manuais em texto" })
        {
            string subDir = Path.Combine(rootFull, sub);
            if (!Directory.Exists(subDir)) continue;

            try
            {
                foreach (string f in Directory.EnumerateFiles(subDir, "*", SearchOption.AllDirectories))
                {
                    string rel = Path.GetRelativePath(rootFull, f).Replace('\\', '/');
                    byRel[rel] = rel;
                    if (sub == "Imagens FAQ")
                        byName.TryAdd(Path.GetFileName(f), rel);
                }
            }
            catch
            {
                // Ignora subpasta inacessível
            }
        }

        return (byRel, byName);
    }

    private static void WriteReadmeIndex(
        string outRoot,
        Dictionary<string, List<(string Title, string RelPath)>> readmeByCat,
        string origin,
        string? linksFileRef)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# 🏢 Central de Tutoriais & FAQs — SharePoint");
        sb.AppendLine();
        sb.AppendLine($"> Espelho sincronizado em Markdown ({DateTime.Now:dd/MM/yyyy HH:mm}). Vídeos e anexos são abertos sob demanda pelo OneDrive.");
        sb.AppendLine();
        if (!string.IsNullOrEmpty(origin) && !string.IsNullOrEmpty(linksFileRef))
        {
            sb.AppendLine($"- 🌐 [Abrir Links-para-FAQs no SharePoint]({origin}{linksFileRef.Replace(" ", "%20")})");
            sb.AppendLine();
        }
        sb.AppendLine("---");
        sb.AppendLine();

        foreach (var kvp in readmeByCat.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        {
            string heading = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(kvp.Key.Replace('_', ' '));
            sb.AppendLine($"## 📁 {heading}");
            sb.AppendLine();
            foreach (var (title, rel) in kvp.Value.OrderBy(x => x.Title, StringComparer.CurrentCultureIgnoreCase))
                sb.AppendLine($"- 📘 [{title}]({rel})");
            sb.AppendLine();
        }

        File.WriteAllText(Path.Combine(outRoot, "README.md"), sb.ToString(), Encoding.UTF8);
    }

    private static string ExtractSiteServerRelativePath(string absolutePath)
    {
        var m = Regex.Match(absolutePath, @"^(/sites/[^/]+|/teams/[^/]+)", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : string.Empty;
    }

    private static string ResolveSiteOrigin(string? sharePointFaqsWebUrl)
    {
        string raw = (sharePointFaqsWebUrl ?? ConfigService.SharePointFaqsWebUrl).Trim();
        return Uri.TryCreate(raw, UriKind.Absolute, out var uri) ? uri.GetLeftPart(UriPartial.Authority) : string.Empty;
    }

    private static string InferCategoryFromTitle(string title)
    {
        string lower = title.ToLowerInvariant();
        if (lower.Contains("impressora") || lower.Contains("papercut") || lower.Contains("scan")) return "impressoras";
        if (lower.Contains("telefone") || lower.Contains("ramal") || lower.Contains("oxe") || lower.Contains("vlan") || lower.Contains("discagem")) return "telefonia";
        if (lower.Contains("formata")) return "formatacao";
        if (lower.Contains("garantia")) return "garantia";
        if (lower.Contains("instala") || lower.Contains("intune") || lower.Contains("sccm") || lower.Contains("lenovo")) return "instalacao";
        return "geral";
    }

    private static string Slugify(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "documento";
        string norm = text.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();
        foreach (char c in norm)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark)
                continue;
            sb.Append(char.IsLetterOrDigit(c) ? char.ToLowerInvariant(c) : '_');
        }
        string clean = Regex.Replace(sb.ToString(), @"_+", "_").Trim('_');
        return string.IsNullOrEmpty(clean) ? "documento" : clean[..Math.Min(clean.Length, 70)];
    }

    private static string StripHtmlTags(string html)
        => Regex.Replace(html, @"<[^>]+>", string.Empty);

    private static string TryGetNestedString(JsonElement parent, string objKey, string propKey)
    {
        if (parent.ValueKind == JsonValueKind.Object &&
            parent.TryGetProperty(objKey, out var obj) &&
            obj.ValueKind == JsonValueKind.Object &&
            obj.TryGetProperty(propKey, out var val) &&
            val.ValueKind == JsonValueKind.String)
        {
            return val.GetString() ?? string.Empty;
        }
        return string.Empty;
    }

    #endregion
}
