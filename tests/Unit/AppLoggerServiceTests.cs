using System.IO;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class AppLoggerServiceTests
{
    [Fact]
    public void AppLogger_LogsWithoutThrowingExceptions()
    {
        var exception = Record.Exception(() =>
        {
            AppLoggerService.Info("Teste de registro INFO unitário.", "AppLoggerServiceTests");
            AppLoggerService.Warning("Teste de registro WARNING unitário.", "AppLoggerServiceTests");
            AppLoggerService.Success("Teste de registro SUCCESS unitário.", "AppLoggerServiceTests");
            AppLoggerService.Danger("Teste de registro DANGER unitário.", "AppLoggerServiceTests");
        });

        Assert.Null(exception);
    }

    [Fact]
    public void AppLogger_IsThreadSafeUnderConcurrency()
    {
        var exception = Record.Exception(() =>
        {
            Parallel.For(0, 50, i =>
            {
                AppLoggerService.Info($"Mensagem concorrente #{i}", "ConcurrentTest");
            });
        });

        Assert.Null(exception);
    }

    [Fact]
    public void AppLogger_ReadRecentLogs_ReturnsEntriesAndFiresEvent()
    {
        string testToken = Guid.NewGuid().ToString("N");
        string? receivedLine = null;
        Action<string> handler = line => receivedLine = line;

        AppLoggerService.OnLogLineWritten += handler;
        try
        {
            AppLoggerService.Info($"Mensagem de teste {testToken}", "TestCaller");
            Assert.NotNull(receivedLine);
            Assert.Contains(testToken, receivedLine);
            Assert.Contains("[PID:", receivedLine);

            var recentLogs = AppLoggerService.ReadRecentLogs(50);
            Assert.NotEmpty(recentLogs);
            Assert.Contains(recentLogs, l => l.Contains(testToken));
        }
        finally
        {
            AppLoggerService.OnLogLineWritten -= handler;
        }
    }

    [Fact]
    public void RotateLogsInDirectory_KeepsAtMostThreeFilesAndDeletesOldest()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), $"log_rot_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        try
        {
            File.WriteAllText(Path.Combine(tempDir, "app.log"), "atual");
            File.WriteAllText(Path.Combine(tempDir, "app_1.log"), "backup1");
            File.WriteAllText(Path.Combine(tempDir, "app_2.log"), "backup2_antigo");

            AppLoggerService.RotateLogsInDirectory(tempDir, maxTotalFiles: 3);

            Assert.False(File.Exists(Path.Combine(tempDir, "app.log")));
            Assert.Equal("atual", File.ReadAllText(Path.Combine(tempDir, "app_1.log")));
            Assert.Equal("backup1", File.ReadAllText(Path.Combine(tempDir, "app_2.log")));
            Assert.False(File.Exists(Path.Combine(tempDir, "app_3.log")));
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }
}

