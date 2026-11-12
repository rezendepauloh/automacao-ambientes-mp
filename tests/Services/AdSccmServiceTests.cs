using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class AdSccmServiceTests
{
    private readonly AdSccmService _service = new();

    [Fact]
    public void FindSccmModulePath_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
        {
            var path = AdSccmService.FindSccmModulePath();
            // Path pode ser nulo se o console não estiver instalado no ambiente de teste
        });

        Assert.Null(exception);
    }

    [Fact]
    public void QueryComputer_HandlesEmptyInput_WithoutThrowing()
    {
        var logs = new List<string>();
        _service.QueryComputer("", logs.Add);

        Assert.NotEmpty(logs);
        Assert.Contains(logs, l => l.Contains("Informe o nome"));
    }

    [Fact]
    public void CleanComputer_HandlesEmptyInput_WithoutThrowing()
    {
        var logs = new List<string>();
        _service.CleanComputer("   ", logs.Add);

        Assert.NotEmpty(logs);
        Assert.Contains(logs, l => l.Contains("Informe o nome"));
    }

    [Fact]
    public async Task SearchMachinesAsync_ReturnsEmptyList_OnWhitespaceQuery()
    {
        var result = await _service.SearchMachinesAsync("   ");
        Assert.NotNull(result);
        Assert.Empty(result);
    }
}

