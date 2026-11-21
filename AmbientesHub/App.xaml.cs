using System.Diagnostics;
using System.Windows;
using AmbientesHub.Services;

namespace AmbientesHub;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private readonly Stopwatch _uptimeSw = Stopwatch.StartNew();

    protected override void OnStartup(StartupEventArgs e)
    {
        string? jumpAction = TaskbarJumpListService.ExtractJumpAction(e.Args);
        if (!string.IsNullOrEmpty(jumpAction))
        {
            TaskbarJumpListService.TrySendActionToRunningInstance(jumpAction);
            Shutdown(0);
            return;
        }

        base.OnStartup(e);

        AppLoggerService.Info(
            $"Aplicação iniciada (OS={Environment.OSVersion.VersionString}, .NET={Environment.Version}, Admin={SystemPerformanceService.IsUserAdministrator()}, BaseDir='{AppContext.BaseDirectory}', Args={e.Args.Length}).",
            nameof(App));

        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                AppLoggerService.LogException(nameof(App), $"Crash fatal AppDomain (IsTerminating={args.IsTerminating})", ex, _uptimeSw.Elapsed);
            }
            else
            {
                AppLoggerService.Danger($"[CRASH FATAL] UnhandledException não-Exception: {args.ExceptionObject} (IsTerminating={args.IsTerminating})", nameof(App));
            }
        };

        DispatcherUnhandledException += (s, args) =>
        {
            var baseEx = args.Exception.GetBaseException() ?? args.Exception;
            AppLoggerService.LogException(nameof(App), "Exceção não tratada na Dispatcher UI", baseEx, _uptimeSw.Elapsed);
            MessageBox.Show($"Ocorreu um erro na aplicação:\n{baseEx.Message}", "Erro no Ambientes Hub", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };
    }

    protected override void OnExit(ExitEventArgs e)
    {
        AppLoggerService.Info($"Aplicação encerrando (ExitCode={e.ApplicationExitCode}, Uptime={AppLoggerService.FormatElapsed(_uptimeSw)}).", nameof(App));
        base.OnExit(e);
    }
}

