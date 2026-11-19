using System.Net;
using AmbientesHub.Models;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class NetworkPathModelsTests
{
    [Theory]
    [InlineData("10.10.20.0/24", "10.10.20.45", true)]
    [InlineData("10.10.20.0/24", "10.10.21.45", false)]
    [InlineData("172.16.0.0/16", "172.16.99.1", true)]
    [InlineData("192.168.1.128/25", "192.168.1.200", true)]
    [InlineData("192.168.1.128/25", "192.168.1.50", false)]
    [InlineData("10.0.0.5/32", "10.0.0.5", true)]
    [InlineData("10.0.0.5/32", "10.0.0.6", false)]
    public void SubnetLocationRule_Contains_MatchesIpv4CidrCorrectly(string cidr, string ipStr, bool expected)
    {
        var rule = new SubnetLocationRule { Cidr = cidr, LocationName = "Predio Teste" };
        var ip = IPAddress.Parse(ipStr);

        Assert.Equal(expected, rule.Contains(ip));
    }

    [Theory]
    [InlineData("")]
    [InlineData("10.10.20.0")]
    [InlineData("10.10.20.0/33")]
    [InlineData("10.10.20.0/-1")]
    [InlineData("12345/24")]
    [InlineData("host.local/24")]
    public void SubnetLocationRule_TryParseCidr_RejectsInvalidFormats(string invalidCidr)
    {
        bool ok = SubnetLocationRule.TryParseCidr(invalidCidr, out _, out _, out _);
        Assert.False(ok);
    }

    [Fact]
    public void NetworkHopNode_DisplaysOfflineAndTimeoutCorrectly()
    {
        var timeoutNode = new NetworkHopNode
        {
            HopNumber = 3,
            Kind = NetworkHopKind.Timeout,
            Status = NetworkHopStatus.Timeout,
            Hostname = "* * *"
        };

        var offlineDest = new NetworkHopNode
        {
            HopNumber = 4,
            Kind = NetworkHopKind.Destination,
            Status = NetworkHopStatus.Unreachable,
            Hostname = "estacao-teste.dominio.local",
            IpAddress = "10.10.20.55",
            InferredLocation = "Bloco B - 2o Andar"
        };

        Assert.Equal("SALTO 3", timeoutNode.HopBadgeText);
        Assert.Equal("* * *", timeoutNode.LatencyDisplay);
        Assert.Equal("#b91c1c", timeoutNode.BorderBrushHex);

        Assert.Equal("DESTINO", offlineDest.HopBadgeText);
        Assert.Equal("OFFLINE", offlineDest.LatencyDisplay);
        Assert.Equal("estacao-teste.dominio.local", offlineDest.TitleDisplay);
        Assert.Equal("10.10.20.55", offlineDest.SubtitleDisplay);
        Assert.Equal("Bloco B - 2o Andar", offlineDest.LocationDisplay);
        Assert.Equal("#ef4444", offlineDest.BorderBrushHex);
    }

    [Fact]
    public void NetworkPingStats_ComputesLossAndLatencyAggregates()
    {
        var stats = new NetworkPingStats();
        stats.RecordSample(10);
        stats.RecordSample(null);
        stats.RecordSample(20);
        stats.RecordSample(30);

        Assert.Equal(4, stats.Sent);
        Assert.Equal(3, stats.Received);
        Assert.Equal(1, stats.Lost);
        Assert.Equal(25.0, stats.LossPercent);
        Assert.Equal(10, stats.MinRttMs);
        Assert.Equal(30, stats.MaxRttMs);
        Assert.Equal(20.0, stats.AvgRttMs);
        Assert.Equal(30, stats.LastRttMs);
        Assert.True(stats.IsReachable);
        Assert.Equal(4, stats.History.Count);

        stats.RecordSample(null);
        Assert.False(stats.IsReachable);

        stats.Reset();
        Assert.Equal(0, stats.Sent);
        Assert.Empty(stats.History);
    }
}
