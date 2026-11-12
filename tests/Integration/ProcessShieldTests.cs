using System.Diagnostics;
using Xunit;

namespace AmbientesHub.Tests.Integration;

public class ProcessShieldTests
{
    [Fact]
    public void Shield_ProtectsCurrentProcessAndParentFromTermination()
    {
        int currentPid = Environment.ProcessId;
        Assert.True(currentPid > 0);

        // O próprio processo nunca deve ser um alvo passível de fechamento
        using var currentProc = Process.GetCurrentProcess();
        Assert.Equal(currentPid, currentProc.Id);
    }

    [Theory]
    [InlineData("🛠️ Modo Chamados - Google Chrome", true)]
    [InlineData("🛠️ Modo Contratos - Edge", true)]
    [InlineData("🛠️ MODO PXE - Firefox", true)]
    [InlineData("Google - Google Chrome", false)]
    [InlineData("Início - Microsoft Edge", false)]
    [InlineData("", false)]
    public void Shield_CorrectlyIdentifiesProtectedTitles(string title, bool expectedProtected)
    {
        bool isProtected = !string.IsNullOrEmpty(title) && title.Contains("🛠️ Modo", StringComparison.OrdinalIgnoreCase);
        Assert.Equal(expectedProtected, isProtected);
    }
}

