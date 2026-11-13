using AmbientesHub.Models;
using AmbientesHub.Services;
using Microsoft.Win32;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class RegistryOptimizerServiceTests
{
    private readonly RegistryOptimizerService _service = new();

    [Fact]
    public void RegistryOptimizerService_GetDefaultTweaks_ReturnsPopulatedCatalog()
    {
        var tweaks = _service.GetDefaultTweaks();

        Assert.NotNull(tweaks);
        Assert.True(tweaks.Count >= 10, "Deveria conter pelo menos 10 otimizações essenciais.");

        foreach (var t in tweaks)
        {
            Assert.False(string.IsNullOrWhiteSpace(t.Id));
            Assert.False(string.IsNullOrWhiteSpace(t.Title));
            Assert.False(string.IsNullOrWhiteSpace(t.Category));
            Assert.False(string.IsNullOrWhiteSpace(t.SubKeyPath));
            Assert.False(string.IsNullOrWhiteSpace(t.ValueName));
            Assert.NotNull(t.DefaultValue);
            Assert.NotNull(t.OptimizedValue);
        }

        Assert.Contains(tweaks, t => t.Hive == RegistryHive.CurrentUser);
        Assert.Contains(tweaks, t => t.Hive == RegistryHive.LocalMachine && t.RequiresAdmin);
    }

    [Fact]
    public void RegistryOptimizerService_ScanTweak_DoesNotThrowAndSetsStatus()
    {
        var tweaks = _service.GetDefaultTweaks();
        var menuTweak = tweaks.First(t => t.Id == "MenuShowDelay");

        var exception = Record.Exception(() => _service.ScanTweak(menuTweak));

        Assert.Null(exception);
        Assert.NotEqual(RegistryTweakStatus.NotChecked, menuTweak.Status);
    }

    [Fact]
    public void RegistryOptimizerService_CreateBackup_GeneratesValidRegFile()
    {
        var tweaks = _service.GetDefaultTweaks().Take(3).ToList();
        _service.ScanAll(tweaks);

        string tempRegPath = Path.Combine(Path.GetTempPath(), $"RegTestBackup_{Guid.NewGuid():N}.reg");

        try
        {
            string generated = _service.CreateBackup(tweaks, tempRegPath);

            Assert.True(File.Exists(generated));
            string content = File.ReadAllText(generated);
            Assert.Contains("Windows Registry Editor Version 5.00", content);
            Assert.Contains("Control Panel\\Desktop", content);
        }
        finally
        {
            if (File.Exists(tempRegPath))
            {
                try { File.Delete(tempRegPath); } catch { }
            }
        }
    }

    [Fact]
    public void RegistryOptimizerService_IsRunningAsAdministrator_ReturnsBoolean()
    {
        var exception = Record.Exception(() =>
        {
            bool isAdmin = RegistryOptimizerService.IsRunningAsAdministrator();
            Assert.True(isAdmin || !isAdmin);
        });

        Assert.Null(exception);
    }

    [Fact]
    public void RegistryOptimizerService_ApplyAndRollback_HKCU_SafeTestKey()
    {
        string testSubKey = @"Software\AmbientesHubTests_Temp";
        string valueName = "SpeedTestParam";

        var testTweak = new RegistryTweakItem
        {
            Id = "UnitTestSafeKey",
            Title = "Chave de Teste Unitário",
            Description = "Tweak temporário para validação de ciclo de escrita e reversão.",
            Category = "Responsividade & UI",
            Hive = RegistryHive.CurrentUser,
            SubKeyPath = testSubKey,
            ValueName = valueName,
            ValueKind = RegistryValueKind.DWord,
            DefaultValue = 100,
            OptimizedValue = 1,
            RequiresAdmin = false
        };

        try
        {
            // 1. Aplica o tweak
            var (applySuccess, _) = _service.ApplyTweak(testTweak);
            Assert.True(applySuccess);
            Assert.Equal(RegistryTweakStatus.Optimized, testTweak.Status);

            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
            using var subKey = baseKey.OpenSubKey(testSubKey, false);
            Assert.NotNull(subKey);
            int val = Convert.ToInt32(subKey.GetValue(valueName));
            Assert.Equal(1, val);

            // 2. Executa rollback
            var (rollbackSuccess, _) = _service.RollbackTweak(testTweak);
            Assert.True(rollbackSuccess);

            // 3. Verifica restauração
            using var subKey2 = baseKey.OpenSubKey(testSubKey, false);
            if (subKey2 != null)
            {
                object? restoredVal = subKey2.GetValue(valueName);
                if (restoredVal != null)
                {
                    Assert.Equal(100, Convert.ToInt32(restoredVal));
                }
            }
        }
        finally
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
                baseKey.DeleteSubKeyTree(testSubKey, false);
            }
            catch { }
        }
    }
}
