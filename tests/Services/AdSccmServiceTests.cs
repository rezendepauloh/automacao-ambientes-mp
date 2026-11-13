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

    [Theory]
    [InlineData("*")]
    [InlineData("pc01' OR '1'='1")]
    [InlineData("pc01; Remove-Item C:\\ -Recurse")]
    [InlineData("pc01)(objectClass=*")]
    public void CleanComputer_BlocksWildcardsAndInjectedPayloads_FailClosed(string maliciousName)
    {
        var logs = new List<string>();
        _service.CleanComputer(maliciousName, logs.Add);

        Assert.Contains(logs, l => l.Contains("[ERRO]") && l.Contains("hostname exato"));
        Assert.DoesNotContain(logs, l => l.Contains("Consultando Active Directory"));
    }

    [Theory]
    [InlineData("a'; Start-Process calc #")]
    [InlineData("*)(objectClass=*")]
    [InlineData("a%")]
    public async Task SearchMachinesAsync_RejectsInjectedTerms_FailClosed(string maliciousQuery)
    {
        var logs = new List<string>();
        var result = await _service.SearchMachinesAsync(maliciousQuery, logs.Add);

        Assert.Empty(result);
        Assert.Contains(logs, l => l.Contains("[ERRO]"));
    }

    [Fact]
    public void BuildAdUserLdapFilter_EscapesSpecialCharacters()
    {
        string filter = AdSccmService.BuildAdUserLdapFilter("admin*)(|(cn=*");
        Assert.DoesNotContain("admin*)(|", filter);
        Assert.Contains(@"\2a\29\28", filter);
    }
}

