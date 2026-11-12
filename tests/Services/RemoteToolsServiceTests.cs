using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class RemoteToolsServiceTests
{
    private readonly RemoteToolsService _service = new();

    [Fact]
    public void OpenRdp_ThrowsArgumentException_WhenComputerNameIsEmpty()
    {
        Assert.Throws<ArgumentException>(() => _service.OpenRdp(""));
        Assert.Throws<ArgumentException>(() => _service.OpenRdp("   "));
    }

    [Fact]
    public void OpenSccmRemoteControl_ThrowsArgumentException_WhenComputerNameIsEmpty()
    {
        Assert.Throws<ArgumentException>(() => _service.OpenSccmRemoteControl(""));
        Assert.Throws<ArgumentException>(() => _service.OpenSccmRemoteControl("   "));
    }

    [Fact]
    public void GetCmRcViewerPath_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
        {
            var path = _service.GetCmRcViewerPath();
            // Path pode ser nulo se SCCM não estiver instalado, mas não deve lançar exceção
        });

        Assert.Null(exception);
    }
}

