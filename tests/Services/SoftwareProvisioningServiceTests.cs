using System.Windows;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class SoftwareProvisioningServiceTests
{
    [Fact]
    public void CreateDefaultCatalog_ContainsAllRequestedApps()
    {
        var catalog = SoftwareProvisioningService.CreateDefaultCatalog();

        Assert.NotEmpty(catalog);
        Assert.Contains(catalog, p => p.Method == SoftwareInstallMethod.RsatAdConsole);
        Assert.Contains(catalog, p => p.Method == SoftwareInstallMethod.SccmConsole);
        Assert.Contains(catalog, p => p.WingetId == "Mozilla.Firefox");
        Assert.Contains(catalog, p => p.WingetId == "Microsoft.WindowsTerminal");
        Assert.Contains(catalog, p => p.WingetId == "Microsoft.PowerShell");
        Assert.Contains(catalog, p => p.WingetId == "Microsoft.PowerToys");
        Assert.Contains(catalog, p => p.WingetId == "Microsoft.WSL");
        Assert.Contains(catalog, p => p.WingetId == "voidtools.Everything");
        Assert.Contains(catalog, p => p.WingetId == "srwi.everythingtoolbar.launcher");
        Assert.Contains(catalog, p => p.WingetId == "hluk.CopyQ");
        Assert.Contains(catalog, p => p.WingetId == "Espanso.Espanso");
        Assert.Contains(catalog, p => p.WingetId == "KandoMenu.Kando");
        Assert.Contains(catalog, p => p.WingetId == "AltSnap.AltSnap");
        Assert.Contains(catalog, p => p.Method == SoftwareInstallMethod.DirectMsi && p.Id == "artcraft-launcher");

        // Todos os itens devem ter um ícone/emoji configurado
        Assert.All(catalog, p => Assert.False(string.IsNullOrWhiteSpace(p.IconEmoji)));
    }

    [Fact]
    public void SoftwarePackageItem_SupportsUpdateAndUninstallVisibilities()
    {
        var item = new SoftwarePackageItem
        {
            Id = "app-update-test",
            Name = "App Update Test",
            Category = "Teste",
            Description = "Descrição fictícia",
            Method = SoftwareInstallMethod.Winget,
            WingetId = "Vendor.App"
        };

        // Estado inicial: Não instalado
        Assert.Equal(Visibility.Visible, item.InstallButtonVisibility);
        Assert.Equal(Visibility.Collapsed, item.InstalledBadgeVisibility);
        Assert.Equal(Visibility.Collapsed, item.UpdateButtonVisibility);
        Assert.Equal(Visibility.Collapsed, item.UninstallButtonVisibility);

        // Estado instalado e atualizado
        item.IsInstalled = true;
        item.HasUpdate = false;
        Assert.Equal(Visibility.Collapsed, item.InstallButtonVisibility);
        Assert.Equal(Visibility.Visible, item.InstalledBadgeVisibility);
        Assert.Equal(Visibility.Collapsed, item.UpdateButtonVisibility);
        Assert.Equal(Visibility.Visible, item.UninstallButtonVisibility);

        // Estado com atualização disponível
        item.HasUpdate = true;
        Assert.Equal(Visibility.Collapsed, item.InstalledBadgeVisibility);
        Assert.Equal(Visibility.Visible, item.UpdateBadgeVisibility);
        Assert.Equal(Visibility.Visible, item.UpdateButtonVisibility);
        Assert.Equal(Visibility.Visible, item.UninstallButtonVisibility);
    }

    [Theory]
    [InlineData("7zip.7zip", true, "--silent")]
    [InlineData("Git.Git", false, "--source winget")]
    public void BuildWingetInstallArguments_AlwaysUsesWingetSourceAndAgreements(string pkgId, bool silent, string expectedToken)
    {
        string args = SoftwareProvisioningService.BuildWingetInstallArguments(pkgId, silent);

        Assert.Contains($"--id \"{pkgId}\"", args);
        Assert.Contains("--source winget", args);
        Assert.DoesNotContain("msstore", args, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--accept-package-agreements", args);
        Assert.Contains("--accept-source-agreements", args);
        Assert.Contains(expectedToken, args);

        if (!silent)
        {
            Assert.DoesNotContain("--silent", args);
        }
    }

    [Fact]
    public void BuildWingetInstallArguments_Throws_WhenPackageIdIsEmpty()
    {
        Assert.Throws<ArgumentException>(() => SoftwareProvisioningService.BuildWingetInstallArguments(""));
        Assert.Throws<ArgumentException>(() => SoftwareProvisioningService.BuildWingetInstallArguments("   "));
    }

    [Theory]
    [InlineData("srv-teste.corp.local", true, "/q ")]
    [InlineData("srv-teste.corp.local", false, "TargetDir=")]
    public void BuildSccmSetupArguments_FormatsTargetDirAndServerProperly(string server, bool silent, string expectedPrefix)
    {
        string args = SoftwareProvisioningService.BuildSccmSetupArguments(server, SoftwareProvisioningService.DefaultSccmTargetDir, silent);

        Assert.Contains($"DefaultSiteServerName={server}", args);
        Assert.Contains(expectedPrefix, args);
        if (!silent)
        {
            Assert.DoesNotContain("/q ", args);
        }
    }

    [Fact]
    public void ParseRsatCapabilityFromDism_DetectsInstalledAndNotPresentStates()
    {
        const string installedSample = """
            Capability Identity : Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
            State : Installed
            """;

        var (cap1, isInst1) = SoftwareProvisioningService.ParseRsatCapabilityFromDism(installedSample);
        Assert.Equal("Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0", cap1);
        Assert.True(isInst1);

        const string missingSample = """
            Identidade da Capacidade : Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
            Estado : Not Present
            """;

        var (cap2, isInst2) = SoftwareProvisioningService.ParseRsatCapabilityFromDism(missingSample);
        Assert.Equal("Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0", cap2);
        Assert.False(isInst2);
    }

    [Fact]
    public void SoftwarePackageItem_TogglesVisibilityAndActionTextCorrectly()
    {
        var item = new SoftwarePackageItem
        {
            Id = "app-teste",
            Name = "App Teste",
            Category = "Teste",
            Description = "Descrição fictícia",
            Method = SoftwareInstallMethod.Winget,
            WingetId = "Vendor.App"
        };

        Assert.Equal(Visibility.Collapsed, item.InstalledBadgeVisibility);
        Assert.Equal(Visibility.Visible, item.InstallButtonVisibility);
        Assert.Equal("Instalar", item.InstallButtonText);
        Assert.True(item.CanClickInstall);

        item.IsInstalling = true;
        Assert.Equal("Instalando...", item.InstallButtonText);
        Assert.False(item.CanClickInstall);

        item.IsInstalling = false;
        item.IsInstalled = true;
        Assert.Equal(Visibility.Visible, item.InstalledBadgeVisibility);
        Assert.Equal(Visibility.Collapsed, item.InstallButtonVisibility);
    }

    [Fact]
    public void DetectionHelpers_DoNotThrow()
    {
        var ex = Record.Exception(() =>
        {
            _ = SoftwareProvisioningService.FindAdConsoleMscPath();
            _ = SoftwareProvisioningService.FindSccmConsoleExecutablePath();
            _ = SoftwareProvisioningService.FindAltSnapInstalledPath();
            _ = SoftwareProvisioningService.FindEverythingToolbarInstalledPath();
        });

        Assert.Null(ex);
    }
}

