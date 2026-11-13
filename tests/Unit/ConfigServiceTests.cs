using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class ConfigServiceTests
{
    [Fact]
    public void ConfigService_GetRecommendedDefault_ReturnsValidValuesWithoutHardcodingPersonalData()
    {
        string wslDir = ConfigService.GetRecommendedDefault("BANCADA_WSL_DIR");
        string sccm = ConfigService.GetRecommendedDefault("SCCM_SITE_SERVER");
        string ping = ConfigService.GetRecommendedDefault("PING_TIMEOUT_MS");

        Assert.False(string.IsNullOrWhiteSpace(wslDir));
        Assert.False(string.IsNullOrWhiteSpace(sccm));
        Assert.Equal("2500", ping);
        Assert.DoesNotContain("paulogoncalves", wslDir, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConfigService_Get_ReturnsDefaultValue_WhenKeyNotFound()
    {
        string missing = ConfigService.Get("CHAVE_INEXISTENTE_XYZ_123", "fallback_default");
        Assert.Equal("fallback_default", missing);
    }
}

