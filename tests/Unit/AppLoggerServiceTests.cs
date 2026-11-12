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
}

