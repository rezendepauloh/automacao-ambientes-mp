using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services
{
    public class BancadaChamadosServiceTests
    {
        [Theory]
        [InlineData("", "5432", "db", "user", "pass")]
        [InlineData("127.0.0.1", "5432", "", "user", "pass")]
        [InlineData("127.0.0.1", "5432", "db", "", "pass")]
        [InlineData("127.0.0.1", "5432", "db", "user", "")]
        [InlineData("127.0.0.1;drop", "5432", "db", "user", "pass")]
        [InlineData("127.0.0.1", "abc", "db", "user", "pass")]
        [InlineData("127.0.0.1", "70000", "db", "user", "pass")]
        public void TryBuildConnectionString_FailClosed_WhenMissingOrInvalid(
            string host, string port, string db, string user, string pass)
        {
            bool ok = BancadaChamadosService.TryBuildConnectionString(host, port, db, user, pass, out string connStr, out string error);

            Assert.False(ok);
            Assert.Equal(string.Empty, connStr);
            Assert.False(string.IsNullOrWhiteSpace(error));
        }

        [Fact]
        public void TryBuildConnectionString_Succeeds_WithValidParameters()
        {
            bool ok = BancadaChamadosService.TryBuildConnectionString("127.0.0.1", "5432", "bancada_db", "usuario_teste", "senha_teste", out string connStr, out string error);

            Assert.True(ok);
            Assert.Equal(string.Empty, error);
            Assert.Contains("Host=127.0.0.1", connStr);
            Assert.Contains("Port=5432", connStr);
            Assert.Contains("Database=bancada_db", connStr);
            Assert.Contains("Username=usuario_teste", connStr);
        }

        [Fact]
        public void ChamadoItem_ComputesDisplayPropertiesCorrectly()
        {
            var item = new ChamadoItem
            {
                Id = "INC-1001",
                Base = "citsmart",
                Status = "Em Atendimento",
                Link = "https://exemplo.local/ticket/1001",
                Hostname = "ESTACAO-01"
            };

            Assert.Equal("CitSmart", item.BaseBadge);
            Assert.True(item.HasLink);
            Assert.True(item.HasHostname);
            Assert.Contains("ESTACAO-01", item.LocationDisplay);
        }
    }
}
