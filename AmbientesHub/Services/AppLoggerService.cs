using System.IO;

namespace AmbientesHub.Services;

public enum LogLevel
{
    INFO,
    WARNING,
    DANGER,
    SUCCESS
}

public class AppLoggerService
{
    private static readonly object LockObj = new();
    
    // Grava fixo em C:\AmbientesVirtuais\AmbientesHub\logs
    private static readonly string LogsDirectory = @"C:\AmbientesVirtuais\AmbientesHub\logs";
    private const long MaxFileSizeBytes = 3 * 1024 * 1024; // 3 MB
    private const int MaxLogFiles = 3;

    static AppLoggerService()
    {
        try
        {
            if (!Directory.Exists(LogsDirectory))
            {
                Directory.CreateDirectory(LogsDirectory);
            }
        }
        catch { }
    }

    public static void Log(LogLevel level, string message, string? caller = null)
    {
        lock (LockObj)
        {
            try
            {
                if (!Directory.Exists(LogsDirectory))
                    Directory.CreateDirectory(LogsDirectory);

                string currentLogPath = Path.Combine(LogsDirectory, "app.log");

                // Rotatividade de logs se atingir 3MB
                if (File.Exists(currentLogPath))
                {
                    var fileInfo = new FileInfo(currentLogPath);
                    if (fileInfo.Length >= MaxFileSizeBytes)
                    {
                        RotateLogs();
                    }
                }

                // Formato exigido: [2026-11-11 16:26:06] [INFO]
                string prefix = caller != null ? $" [{caller}]" : "";
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}]{prefix} {InputValidator.SanitizeLogText(message)}";

                File.AppendAllText(currentLogPath, line + Environment.NewLine);
            }
            catch
            {
                // Garante estabilidade
            }
        }
    }

    private static void RotateLogs()
    {
        try
        {
            string file3 = Path.Combine(LogsDirectory, "app_3.log");
            if (File.Exists(file3)) File.Delete(file3);

            string file2 = Path.Combine(LogsDirectory, "app_2.log");
            if (File.Exists(file2)) File.Move(file2, file3);

            string file1 = Path.Combine(LogsDirectory, "app_1.log");
            if (File.Exists(file1)) File.Move(file1, file2);

            string current = Path.Combine(LogsDirectory, "app.log");
            if (File.Exists(current)) File.Move(current, file1);
        }
        catch { }
    }

    public static void Info(string msg, string? caller = null) => Log(LogLevel.INFO, msg, caller);
    public static void Warning(string msg, string? caller = null) => Log(LogLevel.WARNING, msg, caller);
    public static void Danger(string msg, string? caller = null) => Log(LogLevel.DANGER, msg, caller);
    public static void Success(string msg, string? caller = null) => Log(LogLevel.SUCCESS, msg, caller);
}
