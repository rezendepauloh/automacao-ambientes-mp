using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class ZabbixEnrichmentServiceTests
{
    [Theory]
    [InlineData("", "token123", false)]
    [InlineData("https://zabbix.empresa.local/api_jsonrpc.php", "", false)]
    [InlineData("ftp://zabbix.empresa.local/api_jsonrpc.php", "token123", false)]
    [InlineData("https://user:pass@zabbix.empresa.local/api_jsonrpc.php", "token123", false)]
    [InlineData("https://zabbix.empresa.local/api_jsonrpc.php", "meu-token-valido", true)]
    public void IsConfigured_ValidatesUrlAndTokenFailClosed(string url, string token, bool expected)
    {
        Assert.Equal(expected, ZabbixEnrichmentService.IsConfigured(url, token));
    }

    [Fact]
    public void ParseHostInterfaceAndDetails_EnrichesHopsCorrectly()
    {
        string ifaceJson = """
        {
          "jsonrpc": "2.0",
          "result": [
            { "hostid": "10169", "ip": "10.12.99.12" }
          ],
          "id": 1
        }
        """;

        string hostJson = """
        {
          "jsonrpc": "2.0",
          "result": [
            {
              "hostid": "10169",
              "host": "10.12.99.12",
              "name": "SWE-GALPAO - SWE-1012",
              "hostgroups": [ { "name": "Switches/Acesso" } ],
              "inventory": { "location": "Galpao Central - Rack 01", "model": "Switch L2/L3" }
            }
          ],
          "id": 2
        }
        """;

        var map = ZabbixEnrichmentService.ParseHostInterfaceResponse(ifaceJson);
        var matches = ZabbixEnrichmentService.ParseHostDetailsResponse(hostJson, map);

        Assert.True(matches.ContainsKey("10.12.99.12"));

        var hops = new List<NetworkHopNode>
        {
            new() { HopNumber = 1, IpAddress = "10.12.99.12", Hostname = "", Kind = NetworkHopKind.Router }
        };

        ZabbixEnrichmentService.ApplyEnrichment(hops, matches);

        Assert.Equal("SWE-GALPAO - SWE-1012", hops[0].Hostname);
        Assert.Equal("Galpao Central - Rack 01", hops[0].InferredLocation);
        Assert.Contains("Switch L2/L3", hops[0].ProbableDeviceLabel);
    }

    [Fact]
    public void ParseHostDetails_HandlesEmptyInventoryArrayWithoutError()
    {
        var map = new Dictionary<string, List<string>> { ["10170"] = ["10.12.99.23"] };
        string hostJson = """
        {
          "jsonrpc": "2.0",
          "result": [
            {
              "hostid": "10170",
              "host": "10.12.99.23",
              "name": "SWE-ANEXO - SWE-1023",
              "hostgroups": [ { "name": "Switches/Anexo" } ],
              "inventory": []
            }
          ],
          "id": 2
        }
        """;

        var matches = ZabbixEnrichmentService.ParseHostDetailsResponse(hostJson, map);
        Assert.True(matches.TryGetValue("10.12.99.23", out var match));
        Assert.Equal("SWE-ANEXO - SWE-1023", match!.DisplayName);
        Assert.Contains("Switches/Anexo", match.Groups);
    }
}
