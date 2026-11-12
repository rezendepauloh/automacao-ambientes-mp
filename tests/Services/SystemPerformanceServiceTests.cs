using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class SystemPerformanceServiceTests
{
    [Theory]
    [InlineData("Microsoft.XboxGamingOverlay_5.822.11281.0_x64__8wekyb3d8bbwe", true)]
    [InlineData("Microsoft.BingNews_4.55.62231.0_x64__8wekyb3d8bbwe", true)]
    [InlineData("Microsoft.549981C3F5F10_4.2308.1005.0_x64__8wekyb3d8bbwe", true)]
    [InlineData("Microsoft.WindowsCalculator_11.2401.0.0_x64__8wekyb3d8bbwe", false)]
    [InlineData("", false)]
    public void MatchesAnyToken_DetectsExpectedBloatRemovalPatterns(string packageName, bool expectedMatch)
    {
        bool actual = SystemPerformanceService.MatchesAnyToken(packageName, SystemPerformanceService.BloatRemovalTokens);
        Assert.Equal(expectedMatch, actual);
    }

    [Theory]
    [InlineData("Microsoft.ZuneVideo", true)]
    [InlineData("MicrosoftCorporationII.QuickAssist", true)]
    [InlineData("Microsoft.SolitaireCollection", true)]
    [InlineData("Mozilla.Firefox", false)]
    public void MatchesAnyToken_DetectsExpectedAuditTokens(string appName, bool expectedMatch)
    {
        bool actual = SystemPerformanceService.MatchesAnyToken(appName, SystemPerformanceService.BloatAuditTokens);
        Assert.Equal(expectedMatch, actual);
    }

    [Theory]
    [InlineData(null, "3GB")]
    [InlineData("", "3GB")]
    [InlineData("   ", "3GB")]
    [InlineData("4gb", "4GB")]
    [InlineData("6", "6GB")]
    [InlineData("2048MB", "2048MB")]
    [InlineData("invalido", "3GB")]
    public void SanitizeWslMemoryLimit_NormalizesValidAndInvalidInputs(string? input, string expected)
    {
        string actual = SystemPerformanceService.SanitizeWslMemoryLimit(input);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void BuildWslConfigContent_ContainsExpectedSectionsAndLimits()
    {
        string content = SystemPerformanceService.BuildWslConfigContent("4gb", processors: 4, swap: "2GB");

        Assert.Contains("[wsl2]", content);
        Assert.Contains("memory=4GB", content);
        Assert.Contains("processors=4", content);
        Assert.Contains("swap=2GB", content);
        Assert.Contains("networkingMode=nat", content);
        Assert.Contains("autoProxy=true", content);
        Assert.Contains("dnsTunneling=false", content);
        Assert.Contains("[experimental]", content);
        Assert.Contains("autoMemoryReclaim=gradual", content);
    }

    [Theory]
    [InlineData("NTFS DisableDeleteNotify = 0  (Disabled)", true)]
    [InlineData("NTFS DisableDeleteNotify = 1  (Enabled)", false)]
    [InlineData("", null)]
    [InlineData("Saída desconhecida", null)]
    public void ParseTrimStatus_InterpretsFsutilOutputCorrectly(string fsutilOutput, bool? expectedEnabled)
    {
        var (isEnabled, summary) = SystemPerformanceService.ParseTrimStatus(fsutilOutput);
        Assert.Equal(expectedEnabled, isEnabled);
        Assert.False(string.IsNullOrWhiteSpace(summary));
    }

    [Fact]
    public void NormalizeCliOutput_RemovesNullBytesFromUnicodeStreams()
    {
        string raw = "  \0N\0A\0M\0E\0 \0 \r\n";
        string normalized = SystemPerformanceService.NormalizeCliOutput(raw);
        Assert.Equal("NAME", normalized);
    }

    [Fact]
    public void OptimizationOptions_HasExpectedSafeDefaults()
    {
        var opts = new OptimizationOptions();

        Assert.True(opts.DisableTelemetryServices);
        Assert.True(opts.ApplyRegistryTelemetryTweaks);
        Assert.True(opts.RemoveBloatwareApps);
        Assert.True(opts.DisableHibernation);
        Assert.True(opts.EnableSsdTrim);
        Assert.True(opts.DisableBingSearch);
        Assert.True(opts.AccelerateUiMenus);
        Assert.True(opts.ConfigureWslLimit);
        Assert.Equal("4GB", opts.WslMemoryLimit);
        Assert.False(opts.CleanDockerBuilderCache);
    }

    [Fact]
    public void IsUserAdministrator_DoesNotThrow()
    {
        var ex = Record.Exception(() => _ = SystemPerformanceService.IsUserAdministrator());
        Assert.Null(ex);
    }

    [Fact]
    public async Task RunFullAuditAsync_CollectsReadOnlyMetricsWithoutThrowing()
    {
        var service = new SystemPerformanceService();
        var logs = new List<string>();

        var report = await service.RunFullAuditAsync(msg => logs.Add(msg));

        Assert.NotNull(report);
        Assert.NotEmpty(logs);
        Assert.NotEmpty(report.Disks);
        Assert.True(report.TotalRamGb > 0);
        Assert.Contains(logs, l => l.Contains("Auditoria completa concluída", StringComparison.OrdinalIgnoreCase));
    }
}

