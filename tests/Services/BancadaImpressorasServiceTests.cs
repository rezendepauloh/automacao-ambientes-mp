using AmbientesHub.Models;
using Xunit;

namespace AmbientesHub.Tests.Services
{
    public class BancadaImpressorasServiceTests
    {
        [Fact]
        public void ImpressoraItem_ComputesDisplayAndStatusCorrectly()
        {
            var item = new ImpressoraItem
            {
                Nome = "IMP-TI-01",
                Servidor = "srv-print",
                Tipo = "Multifuncional",
                Modelo = "LaserJet M428",
                Localizacao = "Sede - Sala TI",
                IpHost = "10.0.0.50",
                Status = "OK",
                TotalPaginas = 12345,
                FilasRelacionadas = "FILA_TI"
            };

            Assert.True(item.IsStatusOk);
            Assert.True(item.HasIpHost);
            Assert.Contains("Sede - Sala TI", item.SubtituloDisplay);
            Assert.Contains("10.0.0.50", item.RedeEPaginasDisplay);
            Assert.Contains("FILA_TI", item.RedeEPaginasDisplay);
        }

        [Fact]
        public void ImpressoraItem_AlertStatusAndEmptyFallbacks()
        {
            var item = new ImpressoraItem
            {
                Status = "Sem Papel",
                IpHost = "-"
            };

            Assert.False(item.IsStatusOk);
            Assert.False(item.HasIpHost);
            Assert.Equal("Impressora sem nome", item.NomeDisplay);
        }
    }
}
