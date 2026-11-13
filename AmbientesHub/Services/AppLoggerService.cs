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

    public static string LogDirectoryPath => LogsDirectory;
    public static string CurrentLogFilePath => Path.Combine(LogsDirectory, "app.log");
    public static event Action<string>? OnLogLineWritten;

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

                try
                {
                    OnLogLineWritten?.Invoke(line);
                }
                catch { }
            }
            catch
            {
                // Garante estabilidade
            }
        }
    }

    public static List<string> ReadRecentLogs(int maxLines = 1000)
    {
        lock (LockObj)
        {
            try
            {
                string logFile = CurrentLogFilePath;
                if (!File.Exists(logFile)) return [];

                using var fs = new FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = new StreamReader(fs, System.Text.Encoding.UTF8);
                var lines = new List<string>();
                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        lines.Add(line);
                }

                if (lines.Count > maxLines)
                {
                    lines = lines.Skip(lines.Count - maxLines).ToList();
                }

                return lines;
            }
            catch
            {
                return [];
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
