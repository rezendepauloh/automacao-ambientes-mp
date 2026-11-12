using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class ConfigServiceTests
{
    [Fact]
    public void ConfigService_HasDefaultsForCriticalKeys()
    {
        Assert.False(string.IsNullOrWhiteSpace(ConfigService.AdminUser));
        Assert.False(string.IsNullOrWhiteSpace(ConfigService.AdminAdUser));
    }

    [Fact]
    public void ConfigService_Get_ReturnsDefaultValue_WhenKeyNotFound()
    {
        string missing = ConfigService.Get("CHAVE_INEXISTENTE_XYZ_123", "fallback_default");
        Assert.Equal("fallback_default", missing);
    }
}

