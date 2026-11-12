using System.IO;
using System.IO.Compression;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesVirtuais.Tests.Services;

public class BancadaServicesTests
{
    [Theory]
    [InlineData("LENOVO", "Lenovo")]
    [InlineData("DELL INC.", "Dell")]
    [InlineData("HEWLETT-PACKARD", "HP")]
    [InlineData("", "N/D")]
    public void GetFriendlyManufacturer_ShouldNormalizeKnownBrands(string raw, string expected)
    {
        Assert.Equal(expected, HardwareMappings.GetFriendlyManufacturer(raw));
    }

    [Theory]
    [InlineData("0x0198", "Kingston")]
    [InlineData("80AD", "Hynix")]
    [InlineData("0x802C", "Micron")]
    [InlineData("0x029E", "ADATA")]
    public void FormatRamManufacturer_ShouldTranslateJedecHexCodes(string rawHex, string expectedBrand)
    {
        Assert.Equal(expectedBrand, HardwareMappings.FormatRamManufacturer(rawHex));
    }

    [Fact]
    public void GetParsedDiskInfo_ShouldMapKnownModelsAndSmartPrefixes()
    {
        var mapped = HardwareMappings.GetParsedDiskInfo("KINGSTON SA400S37240G");
        Assert.Equal("Kingston", mapped.Fabricante);
        Assert.Equal("A400 SATA 240GB", mapped.Modelo);

        var smart = HardwareMappings.GetParsedDiskInfo("SAMSUNG MZVL2512HCJQ");
        Assert.Equal("Samsung", smart.Fabricante);
        Assert.Equal("MZVL2512HCJQ", smart.Modelo);
    }

    [Theory]
    [InlineData("20260313", "13/03/2026")]
    [InlineData("2025-11-04", "04/11/2025")]
    public void FormatInstallDate_ShouldFormatToBrazilianDate(string raw, string expected)
    {
        Assert.Equal(expected, HardwareMappings.FormatInstallDate(raw));
    }

    [Fact]
    public void RemoteUserProfileItem_ShouldProtectLoadedAndSpecialProfiles()
    {
        var activeProfile = new RemoteUserProfileItem
        {
            Usuario = "usuario.sobrenome",
            Loaded = true,
            Special = false
        };

        Assert.False(activeProfile.CanBeRemoved);
        activeProfile.IsSelected = true;
        Assert.False(activeProfile.IsSelected);

        var inactiveProfile = new RemoteUserProfileItem
        {
            Usuario = "usuario.teste",
            Loaded = false,
            Special = false
        };

        Assert.True(inactiveProfile.CanBeRemoved);
        inactiveProfile.IsSelected = true;
        Assert.True(inactiveProfile.IsSelected);
    }

    [Fact]
    public void RemoteManagementHelper_ConvertLocalPathToTargetPath_ShouldBuildAdminShareUnc()
    {
        string unc = RemoteManagementHelper.ConvertLocalPathToTargetPath("DOMINIO-80600", @"C:\Users\usuario.sobrenome");
        Assert.Equal(@"\\DOMINIO-80600\C$\Users\usuario.sobrenome", unc);

        string local = RemoteManagementHelper.ConvertLocalPathToTargetPath(".", @"C:\Users\usuario.sobrenome");
        Assert.Equal(@"C:\Users\usuario.sobrenome", local);
    }

    [Fact]
    public void DeviceReportExportService_ShouldBuildHtmlAndValidOpenXmlExcel()
    {
        var sample = new DeviceReportData
        {
            ComputerName = "DOMINIO-80600",
            CurrentUser = "usuario.sobrenome",
            CurrentUserDisplayName = "Usuario Sobrenome Teste",
            CurrentUserDomain = "DOMINIO",
            HardwareInfo = new HardwareSummaryItem
            {
                Computador = "Lenovo ThinkCentre M70q",
                Modelo = "11DUSD3R00",
                Fabricante = "Lenovo",
                MemoriaFisicaGb = "16 GB",
                ProcessadoresLogicos = 12,
                ProcessadoresFisicos = 1,
                SerialNumber = "SN123456"
            },
            OS = new OsSummaryItem
            {
                SO = "Microsoft Windows 11 Pro",
                NomeMaquina = "DOMINIO-80600",
                VersaoWindowsHnH = "24H2",
                NumeroBuild = "26100"
            }
        };
        sample.MemoryModules.Add(new RamModuleItem { Fabricante = "Kingston", CapacidadeGb = "16 GB", VelocidadeMhz = "3200 MHz" });
        sample.UserProfiles.Add(new RemoteUserProfileItem { Usuario = "usuario.sobrenome", Nome = "Usuario Sobrenome", EspacoDiscoGb = 4.25 });

        string html = DeviceReportExportService.BuildHtmlReport(sample);
        Assert.Contains("Lenovo ThinkCentre M70q", html);
        Assert.Contains("usuario.sobrenome", html);
        Assert.Contains("sec-users-all", html);

        string tempExcel = Path.Combine(Path.GetTempPath(), $"TestReport_{Guid.NewGuid():N}.xlsx");
        try
        {
            DeviceReportExportService.GenerateExcelReport(sample, tempExcel);
            Assert.True(File.Exists(tempExcel));

            using var zip = ZipFile.OpenRead(tempExcel);
            Assert.NotNull(zip.GetEntry("[Content_Types].xml"));
            Assert.NotNull(zip.GetEntry("xl/workbook.xml"));
            Assert.NotNull(zip.GetEntry("xl/worksheets/sheet1.xml"));
        }
        finally
        {
            if (File.Exists(tempExcel)) File.Delete(tempExcel);
        }
    }

    [Theory]
    [InlineData(0, "0,00 MB")]
    [InlineData(500L * 1024 * 1024, "500,00 MB")]
    [InlineData(11991L * 1024 * 1024, "11,71 GB")]
    public void RemoteMaintenanceService_FormatDataSize_ShouldFormatMbAndGb(long bytes, string expected)
    {
        Assert.Equal(expected, RemoteMaintenanceService.FormatDataSize(bytes));
    }

    [Fact]
    public void ConfigService_DynamicSettings_ShouldHaveValidDefaults()
    {
        Assert.False(string.IsNullOrWhiteSpace(ConfigService.DeviceReportsDir));
        Assert.True(ConfigService.PingTimeoutMs >= 100);
        Assert.Contains("Administrator", ConfigService.ProtectedProfiles);
    }
}

