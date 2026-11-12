using AmbientesHub.Models;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class MachineDetailsTests
{
    [Theory]
    [InlineData(@"DOMINIO\usuario.conectado", @"DOMINIO\usuario.primario", @"DOMINIO\usuario.conectado")]
    [InlineData("", @"DOMINIO\usuario.primario", @"DOMINIO\usuario.primario (Primário)")]
    [InlineData("   ", @"DOMINIO\usuario.primario", @"DOMINIO\usuario.primario (Primário)")]
    [InlineData("", "", "Nenhum usuário ativo")]
    [InlineData("   ", "   ", "Nenhum usuário ativo")]
    public void CurrentLogonUserDisplay_ResolvesHierarchyCorrectly(string currentLogon, string primary, string expectedDisplay)
    {
        var model = new MachineDetails
        {
            CurrentLogonUser = currentLogon,
            PrimaryUser = primary
        };

        Assert.Equal(expectedDisplay, model.CurrentLogonUserDisplay);
    }

    [Theory]
    [InlineData(@"DOMINIO\usuario.primario", @"DOMINIO\usuario.primario")]
    [InlineData("", "Não definido")]
    [InlineData("   ", "Não definido")]
    public void PrimaryUserDisplay_ResolvesCorrectly(string primary, string expectedDisplay)
    {
        var model = new MachineDetails { PrimaryUser = primary };
        Assert.Equal(expectedDisplay, model.PrimaryUserDisplay);
    }

    [Theory]
    [InlineData("10.0.0.1", "1 ms", "10.0.0.1 (1 ms)")]
    [InlineData("10.0.0.1", "", "10.0.0.1")]
    [InlineData("", "", "Sem resposta ICMP")]
    [InlineData("   ", "   ", "Sem resposta ICMP")]
    public void IpWithLatency_FormatsCorrectly(string ip, string ping, string expected)
    {
        var model = new MachineDetails
        {
            IpAddress = ip,
            PingResponseTime = ping
        };

        Assert.Equal(expected, model.IpWithLatency);
    }

    [Theory]
    [InlineData("AA:BB:CC:DD:EE:FF", "AA:BB:CC:DD:EE:FF")]
    [InlineData("", "Não informado")]
    [InlineData("   ", "Não informado")]
    public void MacAddressDisplay_ReturnsMac_OrFallback(string mac, string expected)
    {
        var model = new MachineDetails { MacAddress = mac };
        Assert.Equal(expected, model.MacAddressDisplay);
    }

    [Fact]
    public void OnlineBadgeText_RendersCorrectStatus()
    {
        var online = new MachineDetails
        {
            IsOnline = true,
            IpAddress = "10.0.0.1",
            PingResponseTime = "1 ms"
        };

        var offlineWithIp = new MachineDetails
        {
            IsOnline = false,
            IpAddress = "10.0.0.1"
        };

        var inaccessible = new MachineDetails
        {
            IsOnline = false,
            IpAddress = ""
        };

        Assert.Contains("ONLINE", online.OnlineBadgeText);
        Assert.Contains("10.0.0.1", online.OnlineBadgeText);
        Assert.Contains("OFFLINE", offlineWithIp.OnlineBadgeText);
        Assert.Contains("INACESSÍVEL", inaccessible.OnlineBadgeText);
    }
}
