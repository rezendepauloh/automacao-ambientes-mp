using System.IO;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class SharePointFaqSyncServiceTests
{
    [Fact]
    public void BuildSitePagesApiUrl_DerivesValidRestEndpoint_AndFailsClosedOnInvalidInput()
    {
        string? url = SharePointFaqSyncService.BuildSitePagesApiUrl(
            "https://empresa.sharepoint.com/sites/equipe/Shared%20Documents/Tutoriais-FAQs");

        Assert.NotNull(url);
        Assert.Equal(
            "https://empresa.sharepoint.com/sites/equipe/_api/web/GetList('/sites/equipe/SitePages')/items?$select=Id,Title,FileLeafRef,FileRef,CanvasContent1,Description,Modified&$top=500",
            url);

        Assert.Null(SharePointFaqSyncService.BuildSitePagesApiUrl(""));
        Assert.Null(SharePointFaqSyncService.BuildSitePagesApiUrl("javascript:alert(1)"));
        Assert.Null(SharePointFaqSyncService.BuildSitePagesApiUrl("https://empresa.sharepoint.com/sem-site"));
    }

    [Fact]
    public void SyncFromFeedXml_ConvertsPagesMapsLocalMediaAndGeneratesReadme()
    {
        string tempRoot = Path.Combine(Path.GetTempPath(), $"sp_sync_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path.Combine(tempRoot, "Imagens FAQ", "SiteAssets", "Teste"));
        Directory.CreateDirectory(Path.Combine(tempRoot, "Vídeos FAQ", "Rede"));

        try
        {
            File.WriteAllBytes(Path.Combine(tempRoot, "Imagens FAQ", "SiteAssets", "Teste", "foto1.png"), [137, 80, 78, 71]);
            File.WriteAllText(Path.Combine(tempRoot, "Vídeos FAQ", "Rede", "aula.mp4"), "stub");

            const string sampleAtom = """
                <?xml version="1.0" encoding="utf-8"?>
                <feed xmlns="http://www.w3.org/2005/Atom"
                      xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices"
                      xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">
                  <entry>
                    <content type="application/xml">
                      <m:properties>
                        <d:Id>5</d:Id>
                        <d:Title>Links para FAQs</d:Title>
                        <d:FileLeafRef>Links-para-FAQs.aspx</d:FileLeafRef>
                        <d:FileRef>/sites/equipe/SitePages/Links-para-FAQs.aspx</d:FileRef>
                        <d:Modified>2026-11-26T10:00:00Z</d:Modified>
                        <d:CanvasContent1>&lt;div data-sp-rte=""&gt;&lt;h2&gt;Redes&lt;/h2&gt;&lt;a href="/sites/equipe/SitePages/Guia-Rede.aspx"&gt;Guia Rede&lt;/a&gt;&lt;/div&gt;</d:CanvasContent1>
                      </m:properties>
                    </content>
                  </entry>
                  <entry>
                    <content type="application/xml">
                      <m:properties>
                        <d:Id>10</d:Id>
                        <d:Title>Guia de Rede</d:Title>
                        <d:FileLeafRef>Guia-Rede.aspx</d:FileLeafRef>
                        <d:FileRef>/sites/equipe/SitePages/Guia-Rede.aspx</d:FileRef>
                        <d:Modified>2026-11-26T12:00:00Z</d:Modified>
                        <d:CanvasContent1>&lt;div data-sp-rte=""&gt;&lt;p&gt;Passo inicial.&lt;/p&gt;&lt;div data-imageurl="/sites/equipe/SiteAssets/SitePages/Teste/foto1.png"&gt;&lt;/div&gt;&lt;a href="/sites/equipe/Shared%20Documents/Tutoriais-FAQs/V%C3%ADdeos%20FAQ/Rede/aula.mp4"&gt;Ver aula&lt;/a&gt;&lt;/div&gt;</d:CanvasContent1>
                      </m:properties>
                    </content>
                  </entry>
                </feed>
                """;

            var svc = new SharePointFaqSyncService();
            var res = svc.SyncFromFeedXml(
                sampleAtom,
                tempRoot,
                "https://empresa.sharepoint.com/sites/equipe/Shared%20Documents/Tutoriais-FAQs");

            Assert.Equal(1, res.PagesConverted);
            Assert.Equal(1, res.LocalImagesMapped);
            Assert.Equal(1, res.LocalFilesMapped);

            string mdPath = Path.Combine(tempRoot, "Manuais Markdown", "redes", "guia_de_rede.md");
            Assert.True(File.Exists(mdPath));

            string md = File.ReadAllText(mdPath);
            Assert.Contains("# 📘 Guia de Rede", md);
            Assert.Contains("![foto1.png](<../../Imagens FAQ/SiteAssets/Teste/foto1.png>)", md);
            Assert.Contains("[Ver aula](<../../Vídeos FAQ/Rede/aula.mp4>)", md);
            Assert.True(File.Exists(Path.Combine(tempRoot, "Manuais Markdown", "README.md")));
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void SyncFromFeedXml_RejectsDtdXxePayload_FailClosed()
    {
        string tempRoot = Path.Combine(Path.GetTempPath(), $"sp_xxe_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        try
        {
            const string xxe = """
                <?xml version="1.0"?>
                <!DOCTYPE feed [<!ENTITY xxe SYSTEM "file:///C:/Windows/win.ini">]>
                <feed xmlns="http://www.w3.org/2005/Atom"></feed>
                """;

            var svc = new SharePointFaqSyncService();
            Assert.ThrowsAny<Exception>(() => svc.SyncFromFeedXml(xxe, tempRoot));
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, recursive: true);
        }
    }
}
