using System.IO;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class RdpCatalogServiceTests
{
    [Fact]
    public void ImportAndExportRdmJson_RoundTripsConnectionsCorrectly()
    {
        string rdmJson = """
        {
          "Connections": [
            {
              "ConnectionType": 1,
              "ID": "11111111-1111-1111-1111-111111111111",
              "Name": "SRV-0001 - TESTE DIRETO",
              "OpenEmbedded": true,
              "PromptCredentials": true,
              "AuthentificationLevel": 0,
              "Url": "10.0.0.10",
              "UsesHardDrives": false,
              "RDP": { "AudioCaptureRedirectionMode": true }
            },
            {
              "ConnectionType": 1,
              "ID": "22222222-2222-2222-2222-222222222222",
              "Name": "UNIDADE - SRV-0002",
              "OpenEmbedded": true,
              "Url": "segura.dominio.local",
              "RDP": { "UserName": "usuario.teste[dominio.local\\usuario_adm@SRV-0002]" }
            }
          ],
          "DatabaseID": "33333333-3333-3333-3333-333333333333",
          "Version": 2
        }
        """;

        var imported = RdpCatalogService.ImportFromRdmJson(rdmJson);
        Assert.Equal(2, imported.Servers.Count);

        var s1 = imported.Servers.First(s => s.Name.Contains("SRV-0001"));
        Assert.False(s1.IsSeguraConfigured);
        Assert.Equal("10.0.0.10", s1.HostOrUrl);

        var s2 = imported.Servers.First(s => s.Name.Contains("SRV-0002"));
        Assert.True(s2.IsSeguraConfigured);
        Assert.Equal("Segura PAM", s2.Group);

        string exported = RdpCatalogService.ExportToRdmJson(imported);
        var reimported = RdpCatalogService.ImportFromRdmJson(exported);
        Assert.Equal(2, reimported.Servers.Count);
    }

    [Fact]
    public void SaveAndLoadEncryptedCatalog_PersistsWithDpapiRoundTrip()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"rdp_cat_test_{Guid.NewGuid():N}.dat");
        try
        {
            var doc = new RdpCatalogDocument();
            doc.Servers.Add(new RdpServerEntry
            {
                Name = "SRV-TESTE",
                HostOrUrl = "10.0.0.50",
                Port = 3389,
                UserName = "usuario.teste"
            });

            RdpCatalogService.SaveEncryptedCatalog(doc, tempFile);
            Assert.True(File.Exists(tempFile));

            // Garante que o arquivo gravado não está em JSON texto claro
            string rawText = File.ReadAllText(tempFile);
            Assert.DoesNotContain("SRV-TESTE", rawText);

            var loaded = RdpCatalogService.LoadEncryptedCatalog(tempFile);
            Assert.Single(loaded.Servers);
            Assert.Equal("SRV-TESTE", loaded.Servers[0].Name);
            Assert.Equal("10.0.0.50", loaded.Servers[0].HostOrUrl);
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Fact]
    public void ApplySeguraTemplate_FormatsPamUserCorrectly()
    {
        var entry = new RdpServerEntry { Name = "CIDADE - SRV-9999", HostOrUrl = "10.0.0.99" };
        RdpCatalogService.ApplySeguraTemplate(entry, "segura.dominio.local", "usuario.teste", "dominio.local", "usuario_adm");

        Assert.Equal("segura.dominio.local", entry.HostOrUrl);
        Assert.Equal(@"usuario.teste[dominio.local\usuario_adm@SRV-9999]", entry.UserName);
        Assert.True(entry.IsSeguraConfigured);
    }

    [Fact]
    public void BuildRdpFileContent_SanitizesLineBreaksAndRejectsUnsafeHost()
    {
        var unsafeEntry = new RdpServerEntry { HostOrUrl = "10.0.0.1 && calc" };
        Assert.Throws<ArgumentException>(() => RdpCatalogService.BuildRdpFileContent(unsafeEntry));

        var validEntry = new RdpServerEntry
        {
            HostOrUrl = "10.0.0.10",
            Port = 3390,
            UserName = "usuario\r\ninjected:s:1"
        };
        string rdp = RdpCatalogService.BuildRdpFileContent(validEntry);
        Assert.Contains("full address:s:10.0.0.10:3390", rdp);
        Assert.DoesNotContain("\ninjected:s:1", rdp);
    }

    [Theory]
    [InlineData(RdpQualityProfile.Balanced, 24, 4u, "allow font smoothing:i:1", "disable wallpaper:i:1")]
    [InlineData(RdpQualityProfile.MaxStability, 16, 2u, "disable themes:i:1", "allow font smoothing:i:0")]
    [InlineData(RdpQualityProfile.BestVisual, 32, 6u, "allow desktop composition:i:1", "disable wallpaper:i:0")]
    public void QualityProfiles_ConfigureColorDepthAndRdpDirectivesCorrectly(
        RdpQualityProfile profile,
        int expectedBpp,
        uint expectedConnType,
        string expectedFlagA,
        string expectedFlagB)
    {
        var qp = RdpServerEntry.GetProfileParameters(profile);
        Assert.Equal(expectedBpp, qp.ColorDepth);
        Assert.Equal(expectedConnType, qp.NetworkConnectionType);

        var entry = new RdpServerEntry
        {
            HostOrUrl = "10.0.0.20",
            QualityProfile = profile,
            AutoAdaptResolution = true
        };

        string rdp = RdpCatalogService.BuildRdpFileContent(entry);
        Assert.Contains($"session bpp:i:{expectedBpp}", rdp);
        Assert.Contains($"connection type:i:{expectedConnType}", rdp);
        Assert.Contains("dynamic resolution:i:1", rdp);
        Assert.Contains(expectedFlagA, rdp);
        Assert.Contains(expectedFlagB, rdp);
    }
}
