using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services
{
    public class BancadaApiClientTests
    {
        [Theory]
        [InlineData("", "chave_super_segura_com_mais_de_16_chars")]
        [InlineData("ftp://127.0.0.1:8502", "chave_super_segura_com_mais_de_16_chars")]
        [InlineData("http://user:pass@127.0.0.1:8502", "chave_super_segura_com_mais_de_16_chars")]
        [InlineData("http://127.0.0.1:8502", "")]
        [InlineData("http://127.0.0.1:8502", "curta")]
        public void TryBuildApiEndpoint_FailClosed_WhenUrlOrKeyInvalid(string url, string key)
        {
            bool ok = BancadaApiClient.TryBuildApiEndpoint(url, key, out string normalized, out string cleanKey, out string error);

            Assert.False(ok);
            Assert.Equal(string.Empty, normalized);
            Assert.Equal(string.Empty, cleanKey);
            Assert.False(string.IsNullOrWhiteSpace(error));
        }

        [Fact]
        public void TryBuildApiEndpoint_Succeeds_AndNormalizesTrailingSlash()
        {
            bool ok = BancadaApiClient.TryBuildApiEndpoint(
                "https://bancada.mpms.mp.br/api/",
                "token_seguro_teste_123456",
                out string normalized,
                out string cleanKey,
                out string error);

            Assert.True(ok);
            Assert.Equal("https://bancada.mpms.mp.br/api", normalized);
            Assert.Equal("token_seguro_teste_123456", cleanKey);
            Assert.Equal(string.Empty, error);
        }
    }
}
