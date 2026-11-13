using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class BancadaDockerServiceTests
{
    [Fact]
    public void BancadaConfiguration_HasValidDefaults()
    {
        Assert.False(string.IsNullOrWhiteSpace(BancadaDockerService.BancadaWslDistro));
        Assert.False(string.IsNullOrWhiteSpace(BancadaDockerService.BancadaWslDir));
        Assert.False(string.IsNullOrWhiteSpace(BancadaDockerService.BancadaWindowsDir));
        Assert.StartsWith("http", BancadaDockerService.BancadaUrl);
    }

    [Fact]
    public void BancadaDockerService_InitialStateIsNotRunning()
    {
        var service = new BancadaDockerService();
        Assert.False(service.IsProcessRunning);
    }
}

