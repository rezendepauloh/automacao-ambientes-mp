using System.Runtime.ExceptionServices;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class DocumentationServiceTests
{
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
    public void IsSafeDocPath_And_ReadDocumentContent_EnforceContainmentAndExtension()
    {
        string tempRoot = CreateTempDocsDir();
        string siblingDir = tempRoot + "_externo";
        Directory.CreateDirectory(siblingDir);

        try
        {
            string validMd = Path.Combine(tempRoot, "valido.md");
            File.WriteAllText(validMd, "# Documento Válido");

            string siblingMd = Path.Combine(siblingDir, "escape_prefixo.md");
            File.WriteAllText(siblingMd, "# Fora da Raiz");

            Assert.True(DocumentationService.IsSafeDocPath(tempRoot, validMd));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, Path.Combine(tempRoot, "..", "AGENTS.md")));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, siblingMd));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, Path.Combine(tempRoot, ".env")));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, Path.Combine(tempRoot, "script.ps1")));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, validMd + "\0.md"));
            Assert.False(DocumentationService.IsSafeDocPath(tempRoot, ""));

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
    public void ResolveMarkdownLinkTarget_HandlesInternalRelativeExternalAndBlocksTraversal()
    {
        string tempRoot = CreateTempDocsDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(tempRoot, "guias"));
            Directory.CreateDirectory(Path.Combine(tempRoot, "arquitetura"));

            string readme = Path.Combine(tempRoot, "README.md");
            string guia = Path.Combine(tempRoot, "guias", "uso.md");
            string arq = Path.Combine(tempRoot, "arquitetura", "sec.md");
            File.WriteAllText(readme, "# Índice");
            File.WriteAllText(guia, "# Uso");
            File.WriteAllText(arq, "# Segurança");

            var relChild = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, readme, "guias/uso.md");
            Assert.Equal(MarkdownLinkKind.InternalDocument, relChild.Kind);
            Assert.Equal(Path.GetFullPath(guia), relChild.Target);

            var relSibling = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, guia, "../arquitetura/sec.md#camadas");
            Assert.Equal(MarkdownLinkKind.InternalDocument, relSibling.Kind);
            Assert.Equal(Path.GetFullPath(arq), relSibling.Target);
            Assert.Equal("camadas", relSibling.Anchor);

            var externalOk = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, readme, "https://learn.microsoft.com/dotnet/");
            Assert.Equal(MarkdownLinkKind.ExternalHttpUrl, externalOk.Kind);

            var traversalBlocked = DocumentationService.ResolveMarkdownLinkTarget(tempRoot, guia, "../../AmbientesHub/.env");
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
    public void ParseMarkdown_And_ParseInlines_HandleTablesAlertsCodeListsAndPreserveSnakeCase()
    {
        const string sampleMd = """
            # 🚀 Título Principal

            > [!WARNING] Cuidado ao editar `BANCADA_DB_HOST` e `tag_manual` em snake_case_longo.

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

        var alert = Assert.Single(blocks, b => b.Kind == MarkdownBlockKind.BlockQuote);
        Assert.Equal("WARNING", alert.AlertKind);
        // Garante que snake_case_longo permaneceu texto simples e não virou itálico quebrado
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
    }

    [Fact]
    public void MarkdownFlowDocumentBuilder_BuildsDocumentOnStaThread_WithoutErrors()
    {
        var blocks = DocumentationService.ParseMarkdown(
            "# Título\n\n> [!NOTE] Alerta.\n\n- Item `código`\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n```powershell\ndotnet build\n```");

        Exception? threadError = null;
        int blockCount = 0;

        var staThread = new Thread(() =>
        {
            try
            {
                var doc = MarkdownFlowDocumentBuilder.Build(blocks, baseFontSize: 14.0, codeFontName: "Consolas");
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

        Assert.True(blockCount >= 5);
    }

    private static string CreateTempDocsDir()
    {
        string path = Path.Combine(Path.GetTempPath(), $"ambhub_docs_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }
}
