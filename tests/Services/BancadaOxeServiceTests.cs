using AmbientesHub.Models;
using Xunit;

namespace AmbientesHub.Tests.Services
{
    public class BancadaOxeServiceTests
    {
        [Fact]
        public void OxeRamalItem_ComputesDisplayPropertiesCorrectly()
        {
            var item = new OxeRamalItem
            {
                Ramal = "4100",
                Nome = "Suporte Teste",
                Localidade = "Sede",
                Setor = "TI",
                Tipo = "IP Touch",
                Ip = "10.0.0.10",
                Mac = "00:11:22:33:44:55"
            };

            Assert.Equal("4100", item.RamalDisplay);
            Assert.Equal("Suporte Teste", item.NomeDisplay);
            Assert.Contains("Sede", item.SubtituloDisplay);
            Assert.Contains("IP: 10.0.0.10", item.RedeDisplay);
            Assert.Contains("MAC: 00:11:22:33:44:55", item.RedeDisplay);
        }

        [Fact]
        public void OxeRamalItem_FallbacksWhenEmpty()
        {
            var item = new OxeRamalItem();

            Assert.Equal("S/N", item.RamalDisplay);
            Assert.Equal("Sem nome associado", item.NomeDisplay);
            Assert.Equal("Sem IP/MAC", item.RedeDisplay);
        }
    }
}
