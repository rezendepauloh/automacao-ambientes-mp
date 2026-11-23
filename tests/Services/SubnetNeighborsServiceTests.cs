using System.Net;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class SubnetNeighborsServiceTests
{
    [Fact]
    public void BuildSubnetPrefix24_ExtractsCidrGatewayAndPrefix()
    {
        var ip = IPAddress.Parse("10.111.129.29");
        var (cidr, gw, dots) = SubnetNeighborsService.BuildSubnetPrefix24(ip);

        Assert.Equal("10.111.129.0/24", cidr);
        Assert.Equal("10.111.129.1", gw);
        Assert.Equal("10.111.129.", dots);
    }

    [Theory]
    [InlineData("CN=PC-01,OU=Sala1,OU=Predio,DC=dominio,DC=local", "OU=Sala1,OU=Predio,DC=dominio,DC=local")]
    [InlineData("CN=PC-01,CN=Computers,DC=dominio,DC=local", "")]
    [InlineData("", "")]
    public void ExtractParentOuDn_ReturnsParentOuOnly(string dn, string expected)
    {
        Assert.Equal(expected, SubnetNeighborsService.ExtractParentOuDn(dn));
    }

    [Fact]
    public void MergeNeighborNode_DeduplicatesByIpAndMergesMetadata()
    {
        var list = new List<SubnetNeighborNode>
        {
            new()
            {
                IpAddress = "10.10.20.15",
                Hostname = "",
                MacAddress = "AA:BB:CC:DD:EE:FF",
                IsOnline = true,
                SourceTag = "ARP-REMOTO"
            }
        };

        SubnetNeighborsService.MergeNeighborNode(list, new SubnetNeighborNode
        {
            IpAddress = "10.10.20.15",
            Hostname = "estacao-15.dominio.local",
            LatencyMs = 2,
            IsOnline = true,
            IsReferenceHost = true,
            SourceTag = "ICMP+DNS"
        });

        Assert.Single(list);
        Assert.Equal("estacao-15.dominio.local", list[0].Hostname);
        Assert.Equal("AA:BB:CC:DD:EE:FF", list[0].MacAddress);
        Assert.Equal(2, list[0].LatencyMs);
        Assert.True(list[0].IsReferenceHost);
        Assert.Contains("ARP-REMOTO", list[0].SourceTag);
        Assert.Contains("ICMP+DNS", list[0].SourceTag);
    }

    [Fact]
    public async Task DiscoverNeighborsAsync_RejectsUnsafeTarget_FailClosed()
    {
        var svc = new SubnetNeighborsService();
        await Assert.ThrowsAsync<ArgumentException>(() => svc.DiscoverNeighborsAsync("10.0.0.1 && whoami"));
    }
}
