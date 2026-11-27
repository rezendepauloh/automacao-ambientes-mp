using System.IO;
using System.Runtime.ExceptionServices;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class DocumentationServiceTests
{
    // PNG 1x1 transparente válido (67 bytes) para testes de imagem inline e Base64
    private static readonly byte[] TinyValidPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    [Fact]
    public void ResolveDefaultDocsDirectory_LocatesProjectDocs_AndDiscoversMarkdownFiles()
    {
        string docsDir = DocumentationService.ResolveDefaultDocsDirectory();
        Assert.True(Directory.Exists(docsDir), $"Pasta Docs esperada em '{docsDir}'.");

        var svc = new DocumentationService();
        var docs = svc.DiscoverDocuments(docsDir);

        Assert.NotEmpty(docs);
        Assert.True(docs[0].IsIndexReadme);
        Assert.Equal("README.md", docs[0].RelativePath);
        Assert.Contains(docs, d => d.CategoryKey == "guias");
        Assert.Contains(docs, d => d.CategoryKey == "arquitetura");
    }

    [Fact]
    public void DiscoverDocuments_ExtractsEmojiTitleSummary_AndOrdersCategories()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(tempRoot, "guias"));
            Directory.CreateDirectory(Path.Combine(tempRoot, "arquitetura"));

            File.WriteAllText(
                Path.Combine(tempRoot, "README.md"),
                "# 📚 Central de Documentação\n\n> Visão unificada de todos os módulos.\n");

            File.WriteAllText(
                Path.Combine(tempRoot, "guias", "manual_rede.md"),
                "# 🌐 Manual de Diagnóstico de Rede\n\nProcedimento completo para análise de rotas L3.\n");

            File.WriteAllText(
                Path.Combine(tempRoot, "arquitetura", "seguranca.md"),
                "```powershell\n# Comentário dentro de bloco não é título\n```\n# 🔒 Modelo SecOps\n\n> [!IMPORTANT] Validação estrita por whitelist.\n");

            File.WriteAllText(
                Path.Combine(tempRoot, "guias", "sem_titulo_h1.md"),
                "Apenas um parágrafo descritivo sem cabeçalho H1 inicial.\n");

            var svc = new DocumentationService();
            var docs = svc.DiscoverDocuments(tempRoot);

            Assert.Equal(4, docs.Count);
            Assert.Equal("README.md", docs[0].RelativePath);
            Assert.Equal("📚", docs[0].Emoji);
            Assert.Equal("Central de Documentação", docs[0].Title);
            Assert.Equal("Visão unificada de todos os módulos.", docs[0].Summary);

            var secDoc = docs.Single(d => d.RelativePath == "arquitetura/seguranca.md");
            Assert.Equal("🔒", secDoc.Emoji);
            Assert.Equal("Modelo SecOps", secDoc.Title);
            Assert.Equal("Validação estrita por whitelist.", secDoc.Summary);

            var fallbackDoc = docs.Single(d => d.RelativePath == "guias/sem_titulo_h1.md");
            Assert.Equal("Sem Titulo H1", fallbackDoc.Title);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DiscoverAllDocuments_CombinesProjectDocsAndSharePointManuaisMarkdown_WithSourceFiltering()
    {
        string tempHubDocs = CreateTempDocsDir();
        string tempSpRoot = CreateTempDocsDir();
        try
        {
            File.WriteAllText(Path.Combine(tempHubDocs, "README.md"), "# 📚 Hub Docs\n\nDocumentação de engenharia.");

            string mdDir = Path.Combine(tempSpRoot, "Manuais Markdown", "impressoras");
            Directory.CreateDirectory(mdDir);
            Directory.CreateDirectory(Path.Combine(tempSpRoot, "Vídeos FAQ"));
            File.WriteAllText(Path.Combine(tempSpRoot, "Manuais Markdown", "README.md"), "# 🏢 FAQs SharePoint\n\nÍndice institucional.");
            File.WriteAllText(Path.Combine(mdDir, "papercut.md"), "# 🖨️ PaperCut\n\nConfiguração de fila.");

            var svc = new DocumentationService();
            var all = svc.DiscoverAllDocuments(tempHubDocs, tempSpRoot);

            Assert.Equal(3, all.Count);
            Assert.Single(all, d => d.SourceKind == DocSourceKind.ProjectDocs);
            Assert.Equal(2, all.Count(d => d.SourceKind == DocSourceKind.SharePointFaqs));

            // SourceRootPath do SharePoint preserva a raiz pai (Tutoriais-FAQs) para resolver ../Imagens FAQ e ../Vídeos FAQ
            var spGuide = all.Single(d => d.RelativePath == "impressoras/papercut.md");
            Assert.Equal(Path.GetFullPath(tempSpRoot), spGuide.SourceRootPath);

            var onlySp = svc.FilterDocuments(all, "", "Todas as Categorias", tempHubDocs, "Tutoriais SharePoint");
            Assert.Equal(2, onlySp.Count);

            var onlyHub = svc.FilterDocuments(all, "", "Todas as Categorias", tempHubDocs, "Docs do Hub");
            Assert.Single(onlyHub);
        }
        finally
        {
            Directory.Delete(tempHubDocs, recursive: true);
            Directory.Delete(tempSpRoot, recursive: true);
        }
    }

    [Fact]
    public void IsSafeDocMediaAndAttachmentPaths_EnforceContainmentAndExtensionWhitelists()
    {
        string tempRoot = CreateTempDocsDir();
        string siblingDir = tempRoot + "_externo";
        Directory.CreateDirectory(siblingDir);

        try
        {
            string validMd = Path.Combine(tempRoot, "valido.md");
            string validImg = Path.Combine(tempRoot, "tela.png");
            string validVideo = Path.Combine(tempRoot, "aula.mp4");
            string validPdf = Path.Combine(tempRoot, "manual.pdf");
            File.WriteAllText(validMd, "# Documento Válido");
            File.WriteAllBytes(validImg, TinyValidPng);
            File.WriteAllText(validVideo, "stub");
            File.WriteAllText(validPdf, "stub");

            string siblingMd = Path.Combine(siblingDir, "escape_prefixo.md");
            File.WriteAllText(siblingMd, "# Fora da Raiz");

            Assert.True(DocumentationService.IsSafeDocPath(tempRoot, validMd));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, Path.Combine(tempRoot, "..", "AGENTS.md")));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, siblingMd));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, Path.Combine(tempRoot, ".env")));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, Path.Combine(tempRoot, "script.ps1")));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, validMd + "\0.md"));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, ""));

            Assert.True(DocumentationService.IsSafeMediaPath(tempRoot, validImg));
            Assert.False(DocumentationService.IsSafeMediaPath(tempRoot, Path.Combine(tempRoot, "malware.exe")));
            Assert.False(DocumentationService.IsSafeMediaPath(tempRoot, Path.Combine(siblingDir, "fora.png")));

            Assert.True(DocumentationService.IsSafeAttachmentPath(tempRoot, validVideo));
            Assert.True(DocumentationService.IsSafeAttachmentPath(tempRoot, validPdf));
            Assert.False(DocumentationService.IsSafeAttachmentPath(tempRoot, Path.Combine(tempRoot, "run.exe")));
            Assert.False(DocumentationService.IsSafeAttachmentPath(tempRoot, Path.Combine(siblingDir, "fora.mp4")));

            var svc = new DocumentationService();
            Assert.Contains("Documento Válido", svc.ReadDocumentContent(tempRoot, validMd));
            Assert.Throws<UnauthorizedAccessException>(() =>
                svc.ReadDocumentContent(tempRoot, Path.Combine(tempRoot, "..", "segredo.md")));
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, recursive: true);
            if (Directory.Exists(siblingDir)) Directory.Delete(siblingDir, recursive: true);
        }
    }

    [Fact]
    public void ResolveMarkdownLinkTarget_HandlesInternalDocsImagesAttachmentsExternalAndBlocksTraversal()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(tempRoot, "Manuais Markdown", "guias"));
            Directory.CreateDirectory(Path.Combine(tempRoot, "Imagens FAQ", "Rede"));
            Directory.CreateDirectory(Path.Combine(tempRoot, "Vídeos FAQ", "Rede"));

            string readme = Path.Combine(tempRoot, "Manuais Markdown", "README.md");
            string guia = Path.Combine(tempRoot, "Manuais Markdown", "guias", "uso.md");
            string img = Path.Combine(tempRoot, "Imagens FAQ", "Rede", "passo 01.png");
            string video = Path.Combine(tempRoot, "Vídeos FAQ", "Rede", "demo rede.mp4");
            File.WriteAllText(readme, "# Índice");
            File.WriteAllText(guia, "# Uso");
            File.WriteAllBytes(img, TinyValidPng);
            File.WriteAllText(video, "stub");

            var relChild = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, readme, "guias/uso.md");
            Assert.Equal(MarkdownLinkKind.InternalDocument, relChild.Kind);
            Assert.Equal(Path.GetFullPath(guia), relChild.Target);

            var relImg = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, guia, "<../../Imagens FAQ/Rede/passo 01.png>");
            Assert.Equal(MarkdownLinkKind.LocalImage, relImg.Kind);
            Assert.Equal(Path.GetFullPath(img), relImg.Target);

            var relVideo = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, guia, "<../../Vídeos FAQ/Rede/demo rede.mp4>");
            Assert.Equal(MarkdownLinkKind.LocalAttachment, relVideo.Kind);
            Assert.Equal(Path.GetFullPath(video), relVideo.Target);

            var externalOk = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, readme, "https://learn.microsoft.com/dotnet/");
            Assert.Equal(MarkdownLinkKind.ExternalHttpUrl, externalOk.Kind);

            var traversalBlocked = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, guia, "../../../../AmbientesHub/.env");
            Assert.Equal(MarkdownLinkKind.UnsupportedReference, traversalBlocked.Kind);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void FilterDocuments_SearchesByTitleCategoryAndFileBody()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(tempRoot, "guias"));
            File.WriteAllText(Path.Combine(tempRoot, "README.md"), "# Índice Geral\nResumo inicial.");
            File.WriteAllText(Path.Combine(tempRoot, "guias", "migracao.md"), "# Guia de Migração\nContém a palavra-chave CHAVE_UNICA_PORTPROXY_8501 no corpo.");

            var svc = new DocumentationService();
            var all = svc.DiscoverDocuments(tempRoot);

            var byBody = svc.FilterDocuments(all, "chave_unica_portproxy_8501", "Todas as Categorias", tempRoot);
            Assert.Single(byBody);
            Assert.Equal("guias/migracao.md", byBody[0].RelativePath);

            var byCat = svc.FilterDocuments(all, "", "Índice Geral", tempRoot);
            Assert.Single(byCat);
            Assert.True(byCat[0].IsIndexReadme);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void ParseMarkdown_And_ParseInlines_HandleImagesAttachmentsTablesAlertsCodeListsAndPreserveSnakeCase()
    {
        const string sampleMd = """
            # 🚀 Título Principal

            > [!WARNING] Cuidado ao editar `BANCADA_DB_HOST` e `tag_manual` em snake_case_longo.

            ![Tela Inicial](<../Imagens FAQ/tela 1.png>)

            [Videoaula FortiNet](<../Vídeos FAQ/Firewall/Factory Reset FortiNet.mp4>)

            ---

            - [x] Validar `InputValidator.cs`
            - [ ] Revisar **hardening** e _itálico real_ com [Guia](guias/uso.md)
              - Subitem indentado

            | Camada | Regra |
            | :--- | :--- |
            | `ConfigService` | Escapa `a|b` sem quebrar coluna |

            ````markdown
            ```powershell
            dotnet test
            ```
            ````
            """;

        var blocks = DocumentationService.ParseMarkdown(sampleMd);

        Assert.Contains(blocks, b => b.Kind == MarkdownBlockKind.Heading && b.HeadingLevel == 1);
        Assert.Contains(blocks, b => b.Kind == MarkdownBlockKind.HorizontalRule);

        var img = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.Image);
        Assert.Equal("Tela Inicial", img.ImageAlt);
        Assert.Equal("../Imagens FAQ/tela 1.png", img.ImageSource);

        var attach = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.AttachmentCard);
        Assert.Equal("Videoaula FortiNet", attach.AttachmentLabel);
        Assert.Equal("../Vídeos FAQ/Firewall/Factory Reset FortiNet.mp4", attach.AttachmentSource);

        var alert = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.BlockQuote);
        Assert.Equal("WARNING", alert.AlertKind);
        Assert.Contains(alert.Inlines, s => s.Kind == MarkdownInlineKind.Text && s.Text.Contains("snake_case_longo"));

        var list = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.List);
        Assert.Equal(3, list.ListItems.Count);
        Assert.Equal("☑️", list.ListItems[0].Marker);
        Assert.Equal("⬜", list.ListItems[1].Marker);
        Assert.Equal(1, list.ListItems[2].IndentLevel);
        Assert.Contains(list.ListItems[1].Inlines, s => s.Kind == MarkdownInlineKind.Bold && s.Text == "hardening");
        Assert.Contains(list.ListItems[1].Inlines, s => s.Kind == MarkdownInlineKind.Italic && s.Text == "itálico real");
        Assert.Contains(list.ListItems[1].Inlines, s => s.Kind == MarkdownInlineKind.Link && s.TargetUrl == "guias/uso.md");

        var table = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.Table);
        Assert.Equal(2, table.TableHeaders.Count);
        Assert.Single(table.TableRows);
        Assert.Equal(2, table.TableRows[0].Count);
        Assert.Contains(table.TableRows[0][1], s => s.Kind == MarkdownInlineKind.Code && s.Text == "a|b");

        var code = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.CodeBlock);
        Assert.Equal("markdown", code.CodeLanguage);
        Assert.Contains("```powershell", code.Text);

        // Garante suporte a URLs com parênteses balanceados (ex: SharePoint SitePages/Pasta(1)(2)/img.png)
        var parenSpans = DocumentationService.ParseInlines(
            "Veja [3919206663.png](https://empresa.sharepoint.com/SiteAssets/Pasta(1)(2)/3919206663.png)");
        var parenLink = Assert.Single(parenSpans, s => s.Kind == MarkdownInlineKind.Link);
        Assert.Equal("https://empresa.sharepoint.com/SiteAssets/Pasta(1)(2)/3919206663.png", parenLink.TargetUrl);
    }

    [Fact]
    public void ExportStandaloneHtml_EmbedsLocalImagesAsBase64_AndEncodesHtmlSafely()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(tempRoot, "Manuais Markdown"));
            Directory.CreateDirectory(Path.Combine(tempRoot, "Imagens FAQ"));

            string imgPath = Path.Combine(tempRoot, "Imagens FAQ", "passo1.png");
            string mdPath = Path.Combine(tempRoot, "Manuais Markdown", "guia_campo.md");
            File.WriteAllBytes(imgPath, TinyValidPng);

            const string md = "# Guia <script>alert(1)</script>\n\n## Seção 1\n\n> [!TIP] Use o atalho.\n\n> [!CAUTION] Cuidado redobrado.\n\n![Passo 1](<../Imagens FAQ/passo1.png>)\n\n[Vídeo](<../Vídeos FAQ/demo.mp4>)\n";
            File.WriteAllText(mdPath, md);

            var doc = new DocEntryItem
            {
                Title = "Guia de Campo",
                FileName = "guia_campo.md",
                RelativePath = "guia_campo.md",
                FullPath = mdPath,
                SourceRootPath = tempRoot,
                SourceKind = DocSourceKind.SharePointFaqs
            };

            string html = DocumentationService.ExportStandaloneHtml(
                doc,
                md,
                sharePointWebBaseUrl: "https://empresa.sharepoint.com/sites/equipe/Shared%20Documents/Tutoriais-FAQs");

            Assert.Contains("<!DOCTYPE html>", html);
            Assert.Contains("data:image/png;base64,", html);
            Assert.DoesNotContain("<script>alert(1)</script>", html);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html);
            Assert.Contains("attach-card", html);
            Assert.Contains("toc-drawer", html);
            Assert.Contains("id=\"sec-1\"", html);
            Assert.Contains("scroll-behavior: smooth", html);
            Assert.Contains("quote quote-tip", html);
            Assert.Contains(System.Net.WebUtility.HtmlEncode("💡 DICA"), html);
            Assert.Contains("quote quote-caution", html);
            Assert.Contains(System.Net.WebUtility.HtmlEncode("🛑 CUIDADO"), html);
            Assert.Contains("blockquote.quote-tip", html);
            Assert.Contains("updateActiveSection", html);
            Assert.Contains("link.classList.toggle('active'", html);
            Assert.Contains("https://empresa.sharepoint.com/sites/equipe/Shared%20Documents/Tutoriais-FAQs/V%C3%ADdeos%20FAQ/demo.mp4", html);
            Assert.Contains("Assistir no SharePoint", html);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void BuildSharePointWebUrl_ReturnsNull_WhenProjectDocsOrUnsafeUrlOrTraversal()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            string mdPath = Path.Combine(tempRoot, "Manuais Markdown", "guia.md");
            var spDoc = new DocEntryItem
            {
                FullPath = mdPath,
                SourceRootPath = tempRoot,
                SourceKind = DocSourceKind.SharePointFaqs
            };
            var hubDoc = new DocEntryItem
            {
                FullPath = mdPath,
                SourceRootPath = tempRoot,
                SourceKind = DocSourceKind.ProjectDocs
            };

            Assert.Null(DocumentationService.BuildSharePointWebUrl(hubDoc, "../Vídeos FAQ/demo.mp4", "https://empresa.sharepoint.com/docs"));
            Assert.Null(DocumentationService.BuildSharePointWebUrl(spDoc, "../Vídeos FAQ/demo.mp4", "javascript:alert(1)"));
            Assert.Null(DocumentationService.BuildSharePointWebUrl(spDoc, "../../../Windows/win.ini", "https://empresa.sharepoint.com/docs"));
            Assert.Null(DocumentationService.BuildSharePointWebUrl(spDoc, "../Vídeos FAQ/evil.exe", "https://empresa.sharepoint.com/docs"));
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void MarkdownFlowDocumentBuilder_BuildsDocumentWithImagesAndAttachmentsOnStaThread_WithoutErrors()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            string imgPath = Path.Combine(tempRoot, "sample.png");
            string mdPath = Path.Combine(tempRoot, "doc.md");
            File.WriteAllBytes(imgPath, TinyValidPng);
            File.WriteAllText(mdPath, "# Demo");

            var blocks = DocumentationService.ParseMarkdown(
                "# Título\n\n> [!NOTE] Alerta.\n\n![Exemplo](sample.png)\n\n[Vídeo Demo](demo.mp4)\n\n- Item `código`\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n```powershell\ndotnet build\n```");

            Exception? threadError = null;
            int blockCount = 0;

            var staThread = new Thread(() =>
            {
                try
                {
                    var doc = MarkdownFlowDocumentBuilder.Build(
                        blocks,
                        baseFontSize: 14.0,
                        codeFontName: "Consolas",
                        sourceRootPath: tempRoot,
                        currentDocFullPath: mdPath);
                    blockCount = doc.Blocks.Count;
                }
                catch (Exception ex)
                {
                    threadError = ex;
                }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            if (threadError != null)
                ExceptionDispatchInfo.Capture(threadError).Throw();

            Assert.True(blockCount >= 7);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    private static string CreateTempDocsDir()
    {
        string path = Path.Combine(Path.GetTempPath(), $"ambhub_docs_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }
}
