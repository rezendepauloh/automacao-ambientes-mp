using System.IO;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class SajmpLauncherServiceTests
{
    [Fact]
    public void CreateDefaultModes_IncludesSajmpCardWithSingleStep()
    {
        var modes = EnvironmentLauncherService.CreateDefaultModes();
        var sajmp = Assert.Single(modes, m => m.ModeKey == "SAJMP");

        Assert.Equal("Abrir SAJMP", sajmp.Title);
        var step = Assert.Single(sajmp.Steps);
        Assert.Equal("open_sajmp", step.Id);
        Assert.True(step.IsEnabled);
    }

    [Fact]
    public void LaunchAndLogin_WhenConfigMissing_ReturnsFailClosed()
    {
        var service = new SajmpLauncherService();

        var result = service.LaunchAndLogin(exePath: "", credPath: "", windowTitleHint: "");

        Assert.False(result.Success);
        Assert.Contains("SAJMP_EXE_PATH", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LaunchAndLogin_WhenExeDoesNotExist_ReturnsFailure()
    {
        var service = new SajmpLauncherService();
        string missingExe = Path.Combine(Path.GetTempPath(), $"saj_missing_{Guid.NewGuid():N}.exe");

        var result = service.LaunchAndLogin(exePath: missingExe, credPath: "nonexistent.xml", windowTitleHint: "SAJ - Teste");

        Assert.False(result.Success);
        Assert.Contains("não encontrado", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LaunchAndLogin_WhenCredFileDoesNotExist_ReturnsFailure()
    {
        var service = new SajmpLauncherService();
        string tempExe = Path.Combine(Path.GetTempPath(), $"saj_dummy_{Guid.NewGuid():N}.exe");
        File.WriteAllText(tempExe, string.Empty);

        try
        {
            string missingCred = Path.Combine(Path.GetTempPath(), $"cred_missing_{Guid.NewGuid():N}.xml");
            var result = service.LaunchAndLogin(exePath: tempExe, credPath: missingCred, windowTitleHint: "SAJ - Teste");

            Assert.False(result.Success);
            Assert.Contains("cofre", result.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (File.Exists(tempExe))
                File.Delete(tempExe);
        }
    }
}
