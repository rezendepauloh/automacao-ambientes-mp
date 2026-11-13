using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class RemoteToolsServiceTests
{
    private readonly RemoteToolsService _service = new();

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("pc01 /admin")]
    [InlineData("pc01&calc")]
    [InlineData("pc01;calc")]
    [InlineData("..\\evil")]
    public void OpenRdp_ThrowsArgumentException_WhenComputerNameIsEmptyOrInjected(string badHost)
    {
        Assert.Throws<ArgumentException>(() => _service.OpenRdp(badHost));
    }

    [Theory]
    [InlineData("", "srv-sccm.dominio.local")]
    [InlineData("   ", "srv-sccm.dominio.local")]
    [InlineData("pc01 & calc", "srv-sccm.dominio.local")]
    [InlineData("DOMINIO-01", "")]
    [InlineData("DOMINIO-01", "srv\" /evil")]
    public void OpenSccmRemoteControl_ThrowsArgumentException_WhenHostOrServerInvalid(string host, string server)
    {
        Assert.Throws<ArgumentException>(() => _service.OpenSccmRemoteControl(host, server));
    }

    [Theory]
    [InlineData("")]
    [InlineData("pc01\" & calc")]
    [InlineData("\\\\evil\\share")]
    public void OpenRemoteExplorer_ThrowsArgumentException_WhenComputerNameInvalid(string badHost)
    {
        Assert.Throws<ArgumentException>(() => _service.OpenRemoteExplorer(badHost));
    }

    [Fact]
    public void GetCmRcViewerPath_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
        {
            _ = _service.GetCmRcViewerPath();
        });

        Assert.Null(exception);
    }
}
