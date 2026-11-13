using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class BancadaDockerServiceTests
{
    [Fact]
    public void BancadaConfiguration_DefaultUrlUsesHttpScheme_AndNoPersonalPathsHardcoded()
    {
        Assert.StartsWith("http", BancadaDockerService.BancadaUrl);
        Assert.True(InputValidator.IsSafeHttpUrl(BancadaDockerService.BancadaUrl));

        // Zero dados pessoais hardcoded: quando .env não define caminhos WSL, o fallback é vazio (fail-closed)
        // e nunca deve conter nomes de usuários reais no código.
        Assert.DoesNotContain("paulogoncalves", BancadaDockerService.BancadaWslDir, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("paulogoncalves", BancadaDockerService.BancadaWindowsDir, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("", "/home/usuario/app", false)]
    [InlineData("Ubuntu-24.04", "", false)]
    [InlineData("Ubuntu; rm -rf /", "/home/usuario/app", false)]
    [InlineData("Ubuntu-24.04", "/home/usuario/$(whoami)", false)]
    [InlineData("Ubuntu-24.04", "/home/usuario/`id`", false)]
    [InlineData("Ubuntu-24.04", "/home/usuario/../../etc", false)]
    [InlineData("Ubuntu-24.04", "/home/usuario/app", true)]
    public void ValidateTarget_EnforcesFailClosedAndWhitelist(string distro, string wslDir, bool expectedValid)
    {
        bool ok = BancadaDockerService.ValidateTarget(distro, wslDir, out string error);
        Assert.Equal(expectedValid, ok);
        if (!expectedValid)
            Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void BancadaDockerService_InitialStateIsNotRunning()
    {
        var service = new BancadaDockerService();
        Assert.False(service.IsProcessRunning);
    }
}
