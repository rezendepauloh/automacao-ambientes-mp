using System.IO;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class NetworkPathServiceTests
{
    [Theory]
    [InlineData("estacao-01")]
    [InlineData("servidor.dominio.local")]
    [InlineData("10.10.20.5")]
    [InlineData("192.168.100.254")]
    [InlineData("::1")]
    public void ValidateTarget_AcceptsValidHostsAndIps(string target)
    {
        string validated = NetworkPathService.ValidateTarget($"  {target}  ");
        Assert.Equal(target, validated);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("10.0.0.1; whoami")]
    [InlineData("host & calc")]
    [InlineData("host..local")]
    [InlineData("123456")]
    [InlineData("10.0.0.999")]
    public void ValidateTarget_RejectsInvalidOrUnsafeInputs(string unsafeTarget)
    {
        Assert.Throws<ArgumentException>(() => NetworkPathService.ValidateTarget(unsafeTarget));
    }

    [Fact]
    public void IsStrictIpAddress_DoesNotTreatNumericStringAsIpv4()
    {
        Assert.False(NetworkPathService.IsStrictIpAddress("1234", out _));
        Assert.True(NetworkPathService.IsStrictIpAddress("10.20.30.40", out var ip));
        Assert.Equal("10.20.30.40", ip?.ToString());
    }

    [Fact]
    public void MatchSubnetRule_PrefersMostSpecificCidr()
    {
        var rules = new List<SubnetLocationRule>
        {
            new() { Cidr = "10.10.0.0/16", LocationName = "Campus Geral", ProbableSwitch = "SW-CORE" },
            new() { Cidr = "10.10.20.0/24", LocationName = "Predio B", ProbableSwitch = "SW-BLB" },
            new() { Cidr = "10.10.20.128/25", LocationName = "Predio B - Ala Norte", ProbableSwitch = "SW-BLB-N" }
        };

        var matched = NetworkPathService.MatchSubnetRule("10.10.20.200", rules);
        Assert.NotNull(matched);
        Assert.Equal("10.10.20.128/25", matched!.Cidr);
        Assert.Equal("Predio B - Ala Norte", matched.LocationName);
    }

    [Fact]
    public void SaveAndLoadSubnetRules_PersistsValidRulesAndIgnoresInvalid()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), $"subnet_map_test_{Guid.NewGuid():N}.json");
        try
        {
            var input = new List<SubnetLocationRule>
            {
                new() { Cidr = "10.50.10.0/24", LocationName = "Anexo I", ProbableSwitch = "SW-ANX-01" },
                new() { Cidr = "cidr-invalido", LocationName = "Ignorar" }
            };

            NetworkPathService.SaveSubnetRules(input, tempFile);
            var loaded = NetworkPathService.LoadSubnetRules(tempFile);

            Assert.Single(loaded);
            Assert.Equal("10.50.10.0/24", loaded[0].Cidr);
            Assert.Equal("Anexo I", loaded[0].LocationName);
            Assert.Equal("SW-ANX-01", loaded[0].ProbableSwitch);
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Theory]
    [InlineData(1, "10.0.0.1", "gw-predio1.dominio.local", false, true, NetworkHopKind.Gateway)]
    [InlineData(2, "10.0.0.2", "sw-l3-dist.dominio.local", false, false, NetworkHopKind.SwitchL3)]
    [InlineData(3, "10.0.0.3", "core-dc01.dominio.local", false, false, NetworkHopKind.Backbone)]
    [InlineData(4, "10.0.0.4", "estacao-01.dominio.local", true, false, NetworkHopKind.Destination)]
    public void ClassifyHop_IdentifiesDeviceKindFromHostnameAndRole(
        int hop, string ip, string ptr, bool isDest, bool isFirst, NetworkHopKind expected)
    {
        var kind = NetworkPathService.ClassifyHop(hop, ip, ptr, isDest, isFirst);
        Assert.Equal(expected, kind);
    }

    [Fact]
    public void BuildHopNode_MarksHighLatencyAboveThreshold()
    {
        var rules = new List<SubnetLocationRule>
        {
            new() { Cidr = "10.20.0.0/16", LocationName = "Unidade Remota", ProbableSwitch = "SW-REM-01" }
        };

        var node = NetworkPathService.BuildHopNode(
            hopNumber: 2,
            ipAddress: "10.20.5.1",
            ptrHostname: "rtr-remoto.dominio.local",
            latencyMs: 120,
            ttlRemaining: 62,
            isDestination: false,
            isFirstHop: false,
            subnetRules: rules);

        Assert.Equal(NetworkHopStatus.HighLatency, node.Status);
        Assert.Equal("Unidade Remota", node.InferredLocation);
        Assert.Contains("SW-REM-01", node.ProbableDeviceLabel);
        Assert.Equal("10.20.0.0/16", node.MatchedCidr);
    }

    [Fact]
    public void FormatTraceReport_IncludesHeaderAndAllHops()
    {
        var trace = new NetworkTraceResult
        {
            TargetInput = "estacao-teste",
            ResolvedTargetIp = "10.10.20.30",
            ResolvedTargetFqdn = "estacao-teste.dominio.local",
            TargetOnline = false,
            Hops =
            [
                new NetworkHopNode { HopNumber = 0, IpAddress = "10.0.0.10", Hostname = "ORIGEM-PC", Kind = NetworkHopKind.LocalStation, LatencyMs = 0 },
                new NetworkHopNode { HopNumber = 1, IpAddress = "10.0.0.1", Hostname = "gw.dominio.local", Kind = NetworkHopKind.Gateway, LatencyMs = 2, InferredLocation = "Sede" },
                new NetworkHopNode { HopNumber = 2, IpAddress = "10.10.20.30", Hostname = "estacao-teste.dominio.local", Kind = NetworkHopKind.Destination, Status = NetworkHopStatus.Unreachable }
            ]
        };

        string report = NetworkPathService.FormatTraceReport(trace);
        Assert.Contains("estacao-teste", report);
        Assert.Contains("OFFLINE", report);
        Assert.Contains("10.0.0.1", report);
        Assert.Contains("Sede", report);
    }

    [Fact]
    public void FormatTraceHtmlReport_GeneratesResponsiveHtmlWithBrazilianDatesAndEscaping()
    {
        var trace = new NetworkTraceResult
        {
            TargetInput = "estacao-teste<script>",
            ResolvedTargetIp = "10.10.20.30",
            TargetOnline = true,
            StartedAt = new DateTime(2026, 11, 19, 13, 25, 0),
            CompletedAt = new DateTime(2026, 11, 19, 13, 25, 4),
            Hops =
            [
                new NetworkHopNode { HopNumber = 0, IpAddress = "10.0.0.10", Hostname = "ORIGEM-PC", Kind = NetworkHopKind.LocalStation, LatencyMs = 0 },
                new NetworkHopNode { HopNumber = 1, IpAddress = "10.10.20.30", Hostname = "estacao-teste", Kind = NetworkHopKind.Destination, LatencyMs = 2 }
            ]
        };

        var ping = new NetworkPingStats();
        ping.RecordSample(2);

        string html = NetworkPathService.FormatTraceHtmlReport(trace, ping);

        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("name=\"viewport\"", html);
        Assert.Contains("19/11/2026 13:25:00", html);
        Assert.Contains("19/11/2026 13:25:04", html);
        Assert.DoesNotContain("<script>", html);
        Assert.Contains("&lt;script&gt;", html);
        Assert.Contains("ONLINE", html);
    }

    [Fact]
    public async Task PingOnceAsync_RejectsInvalidTarget_FailClosed()
    {
        var svc = new NetworkPathService();
        await Assert.ThrowsAsync<ArgumentException>(() => svc.PingOnceAsync("10.0.0.1 && dir"));
    }

    [Fact]
    public async Task PingOnceAsync_LoopbackCompletesWithoutException()
    {
        var svc = new NetworkPathService();
        var result = await svc.PingOnceAsync("127.0.0.1", timeoutMs: 1000);
        Assert.True(result.Success);
        Assert.NotNull(result.RttMs);
    }

    [Fact]
    public async Task CheckCommonPortsAsync_RejectsInvalidTarget_FailClosed()
    {
        var svc = new NetworkPathService();
        await Assert.ThrowsAsync<ArgumentException>(() => svc.CheckCommonPortsAsync("10.0.0.1; calc"));
    }

    [Fact]
    public async Task CheckCommonPortsAsync_ReturnsAllFourStandardPorts()
    {
        var svc = new NetworkPathService();
        var results = await svc.CheckCommonPortsAsync("127.0.0.1", timeoutMs: 300);
        Assert.Equal(4, results.Count);
        Assert.Contains(results, r => r.Port == 3389 && r.ServiceName == "RDP");
        Assert.Contains(results, r => r.Port == 445 && r.ServiceName == "SMB/C$");
        Assert.Contains(results, r => r.Port == 135 && r.ServiceName == "WMI/RPC");
        Assert.Contains(results, r => r.Port == 5985 && r.ServiceName == "WinRM");
    }

    [Fact]
    public void BuildWakeOnLanMagicPacket_Builds102BytesCorrectly()
    {
        byte[] packet = NetworkPathService.BuildWakeOnLanMagicPacket("AA:BB:CC:DD:EE:FF");
        Assert.Equal(102, packet.Length);
        for (int i = 0; i < 6; i++)
            Assert.Equal(0xFF, packet[i]);

        Assert.Equal(0xAA, packet[6]);
        Assert.Equal(0xFF, packet[11]);
        Assert.Equal(0xAA, packet[96]);
        Assert.Equal(0xFF, packet[101]);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("ZZ:ZZ:ZZ:ZZ:ZZ:ZZ")]
    [InlineData("AA:BB:CC")]
    public void BuildWakeOnLanMagicPacket_RejectsInvalidMac_FailClosed(string? invalidMac)
    {
        Assert.Throws<ArgumentException>(() => NetworkPathService.BuildWakeOnLanMagicPacket(invalidMac));
    }
}
