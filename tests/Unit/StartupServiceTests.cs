using System.IO;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class StartupServiceTests
{
    private class TestableStartupService : StartupService
    {
        private readonly string _mockExePath;
        private readonly string _mockShortcutPath;

        public TestableStartupService(string mockExePath, string mockShortcutPath)
        {
            _mockExePath = mockExePath;
            _mockShortcutPath = mockShortcutPath;
        }

        public override string GetExecutablePath() => _mockExePath;
        public override string GetStartMenuShortcutPath() => _mockShortcutPath;
    }

    [Fact]
    public void StartupService_GetExecutablePath_ReturnsValidString()
    {
        var service = new StartupService();
        string path = service.GetExecutablePath();

        Assert.False(string.IsNullOrWhiteSpace(path));
        Assert.EndsWith(".exe", path, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void StartupService_EnableAndDisable_CyclesCorrectly()
    {
        var tempShortcut = Path.Combine(Path.GetTempPath(), "test_ambientes_hub.lnk");
        var service = new TestableStartupService(@"C:\AmbientesVirtuais\AmbientesHub\bin\Debug\net8.0-windows\AmbientesHub.exe", tempShortcut);

        bool setEnabled = service.SetRunAtStartup(true);
        Assert.True(setEnabled);
        Assert.True(service.IsRunAtStartupEnabled());

        bool setDisabled = service.SetRunAtStartup(false);
        Assert.True(setDisabled);
        Assert.False(service.IsRunAtStartupEnabled());
    }

    [Fact]
    public void StartupService_StartMenuShortcut_CanCreateAndRemove()
    {
        var tempShortcut = Path.Combine(Path.GetTempPath(), $"hub_shortcut_{Guid.NewGuid():N}.lnk");
        var service = new TestableStartupService(@"C:\Windows\notepad.exe", tempShortcut);

        try
        {
            Assert.False(service.IsStartMenuShortcutCreated());

            bool created = service.CreateStartMenuShortcut();
            Assert.True(created);
            Assert.True(service.IsStartMenuShortcutCreated());
            Assert.True(File.Exists(tempShortcut));

            bool removed = service.RemoveStartMenuShortcut();
            Assert.True(removed);
            Assert.False(service.IsStartMenuShortcutCreated());
            Assert.False(File.Exists(tempShortcut));
        }
        finally
        {
            if (File.Exists(tempShortcut))
            {
                File.Delete(tempShortcut);
            }
        }
    }
}
