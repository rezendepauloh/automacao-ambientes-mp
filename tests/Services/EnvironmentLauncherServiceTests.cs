using AmbientesHub.Models;
using AmbientesHub.Services;
using System.Windows;
using Xunit;

namespace AmbientesHub.Tests.Services
{
    public class EnvironmentLauncherServiceTests
    {
        [Fact]
        public void CreateDefaultModes_ReturnsAllFiveModesWithEnabledSteps()
        {
            var modes = EnvironmentLauncherService.CreateDefaultModes();

            Assert.Equal(5, modes.Count);
            Assert.Contains(modes, m => m.ModeKey == "Chamados");
            Assert.Contains(modes, m => m.ModeKey == "Contratos");
            Assert.Contains(modes, m => m.ModeKey == "Impressoras");
            Assert.Contains(modes, m => m.ModeKey == "PXE");
            Assert.Contains(modes, m => m.ModeKey == "Estudo");

            foreach (var mode in modes)
            {
                Assert.NotEmpty(mode.Steps);
                Assert.All(mode.Steps, s => Assert.True(s.IsEnabled));
            }
        }

        [Fact]
        public void EnvironmentModeItem_ToggleExpanded_UpdatesVisibilityAndChevron()
        {
            var mode = new EnvironmentModeItem { ModeKey = "Chamados" };

            Assert.False(mode.IsExpanded);
            Assert.Equal(Visibility.Collapsed, mode.DetailsVisibility);
            Assert.Equal("⌄", mode.ChevronSymbol);

            mode.IsExpanded = true;

            Assert.Equal(Visibility.Visible, mode.DetailsVisibility);
            Assert.Equal("⌃", mode.ChevronSymbol);
        }
    }
}
