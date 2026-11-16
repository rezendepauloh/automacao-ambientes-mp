using System.IO;
using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class UsmtBackupServiceTests
{
    [Fact]
    public void BuildScanStateArgumentList_SelectsOnlyCheckedUsers_AndContainsNoDeleteFlags()
    {
        var opts = new UsmtBackupOptions
        {
            TargetComputer = ".",
            StorePath = @"C:\Backups\USMT_Estacao01",
            SelectedUsers = ["usuario.teste", @"DOMINIO\operador.suporte"],
            IncludeMigApp = true,
            IncludeMigDocs = true,
            IncludeMigUser = false,
            Compress = true,
            EncryptionKey = "ChaveForte123"
        };

        var args = UsmtBackupService.BuildScanStateArgumentList(@"C:\Tools\USMT\amd64", opts, @"C:\Logs\usmt.log");

        Assert.Contains(@"C:\Backups\USMT_Estacao01", args);
        Assert.Contains("/o", args);
        Assert.Contains("/c", args);
        Assert.Contains(@"/i:C:\Tools\USMT\amd64\MigApp.xml", args);
        Assert.Contains(@"/i:C:\Tools\USMT\amd64\MigDocs.xml", args);
        Assert.DoesNotContain(@"/i:C:\Tools\USMT\amd64\MigUser.xml", args);
        Assert.Contains("/encrypt", args);
        Assert.Contains("/key:ChaveForte123", args);
        Assert.Contains(@"/ue:*\*", args);
        Assert.Contains(@"/ui:*\usuario.teste", args);
        Assert.Contains(@"/ui:DOMINIO\operador.suporte", args);
        Assert.Contains(@"/l:C:\Logs\usmt.log", args);

        // Garante que nunca exista qualquer flag destrutiva de perfil
        Assert.DoesNotContain(args, a => a.Contains("/delete", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void BuildScanStateArgumentList_ThrowsFailClosed_WhenNoUsersSelected()
    {
        var opts = new UsmtBackupOptions
        {
            StorePath = @"C:\Backups\USMT",
            SelectedUsers = []
        };

        Assert.Throws<ArgumentException>(() =>
            UsmtBackupService.BuildScanStateArgumentList(@"C:\Tools\USMT", opts));
    }

    [Theory]
    [InlineData(@"relativo\usmt")]
    [InlineData(@"C:\Backups\..\..\Windows\System32")]
    [InlineData(@"C:\Backups & calc")]
    [InlineData("")]
    public void BuildScanStateArgumentList_ThrowsFailClosed_OnUnsafeStorePath(string badPath)
    {
        var opts = new UsmtBackupOptions
        {
            StorePath = badPath,
            SelectedUsers = ["usuario.teste"]
        };

        Assert.Throws<ArgumentException>(() =>
            UsmtBackupService.BuildScanStateArgumentList(@"C:\Tools\USMT", opts));
    }

    [Fact]
    public void BuildScanStateArgumentList_ThrowsFailClosed_OnInvalidUserName()
    {
        var opts = new UsmtBackupOptions
        {
            StorePath = @"C:\Backups\USMT",
            SelectedUsers = ["usuario.teste & calc"]
        };

        Assert.Throws<ArgumentException>(() =>
            UsmtBackupService.BuildScanStateArgumentList(@"C:\Tools\USMT", opts));
    }

    [Fact]
    public void BuildLoadStateArgumentList_SupportsRestoreAllAndOptionalDecryption()
    {
        var opts = new UsmtBackupOptions
        {
            StorePath = @"\\srv-arquivos\backups$\estacao01",
            SelectedUsers = [],
            IncludeMigApp = true,
            IncludeMigDocs = false,
            IncludeMigUser = true,
            Compress = false,
            CreateLocalAccountsOnRestore = true
        };

        var args = UsmtBackupService.BuildLoadStateArgumentList(@"C:\Tools\USMT", opts);

        Assert.Contains(@"\\srv-arquivos\backups$\estacao01", args);
        Assert.Contains("/lac", args);
        Assert.Contains("/lae", args);
        Assert.Contains("/nocompress", args);
        Assert.Contains(@"/i:C:\Tools\USMT\MigApp.xml", args);
        Assert.Contains(@"/i:C:\Tools\USMT\MigUser.xml", args);
        Assert.DoesNotContain(@"/ue:*\*", args);
    }

    [Fact]
    public void IsValidUsmtDirectory_And_FindUsmtDirectory_RequireBothExecutables()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), $"UsmtTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            Assert.False(UsmtBackupService.IsValidUsmtDirectory(tempDir));

            File.WriteAllText(Path.Combine(tempDir, "scanstate.exe"), "stub");
            Assert.False(UsmtBackupService.IsValidUsmtDirectory(tempDir));

            File.WriteAllText(Path.Combine(tempDir, "loadstate.exe"), "stub");
            Assert.True(UsmtBackupService.IsValidUsmtDirectory(tempDir));
            Assert.Equal(tempDir, UsmtBackupService.FindUsmtDirectory(tempDir));
        }
        finally
        {
            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }
}
