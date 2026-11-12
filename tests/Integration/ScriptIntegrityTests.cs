using System.IO;
using Xunit;

namespace AmbientesHub.Tests.Integration;

public class ScriptIntegrityTests
{
    private const string RepoRoot = @"C:\AmbientesVirtuais";

    [Fact]
    public void LegacyBibliotecaScript_ExistsAndHasValidEncoding()
    {
        string bibPath = Path.Combine(RepoRoot, "Scripts", "biblioteca.ps1");
        Assert.True(File.Exists(bibPath), $"O script legado em {bibPath} deve existir.");

        string content = File.ReadAllText(bibPath);
        Assert.NotEmpty(content);

        // Funções vitais legadas catalogadas
        Assert.Contains("function Fechar-AmbientePrevio", content);
        Assert.Contains("function Salvar-E-Fechar-Office", content);
        Assert.Contains("function Parar-Telemetria", content);
    }

    [Fact]
    public void CredentialsDirectory_Exists()
    {
        string credDir = Path.Combine(RepoRoot, "Credenciais");
        Assert.True(Directory.Exists(credDir), $"O diretório de credenciais {credDir} deve existir.");
    }
}

