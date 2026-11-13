using System.Windows;
using AmbientesHub.Services;

namespace AmbientesHub;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                AppLoggerService.Danger($"[CRASH FATAL] UnhandledException: {ex.Message}\n{ex.StackTrace}", nameof(App));
            }
        };

        DispatcherUnhandledException += (s, args) =>
        {
            var baseEx = args.Exception.GetBaseException() ?? args.Exception;
            AppLoggerService.Danger($"[ERRO UI] BaseException: {baseEx.Message}\n{baseEx.StackTrace}", nameof(App));
            MessageBox.Show($"Ocorreu um erro na aplicação:\n{baseEx.Message}", "Erro no Ambientes Hub", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };
    }
}

