using System.Diagnostics;
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
    private static readonly int CurrentPid = Environment.ProcessId;
    private static readonly Lazy<string> PrivilegeTag = new(() =>
        SystemPerformanceService.IsUserAdministrator() ? "Admin" : "User");

    // Grava fixo em C:\AmbientesVirtuais\AmbientesHub\logs
    private static readonly string LogsDirectory = @"C:\AmbientesVirtuais\AmbientesHub\logs";
    private const long MaxFileSizeBytes = 3 * 1024 * 1024; // 3 MB
    private const int MaxLogFiles = 3; // app.log + app_1.log + app_2.log (máx. 3 arquivos)

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

                // Formato enriquecido com milissegundos, PID, ThreadId e privilégio
                int threadId = Environment.CurrentManagedThreadId;
                string safeCaller = InputValidator.SanitizeLogText(caller);
                string prefix = !string.IsNullOrEmpty(safeCaller) ? $" [{safeCaller}]" : "";
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] [PID:{CurrentPid}|T:{threadId}|{PrivilegeTag.Value}]{prefix} {InputValidator.SanitizeLogText(message)}";

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

    public static void RotateLogsInDirectory(string directory, int maxTotalFiles = MaxLogFiles)
    {
        try
        {
            // Mantém no máximo `maxTotalFiles` (ex: app.log + app_1.log + app_2.log = 3 arquivos)
            int maxBackupIndex = Math.Max(1, maxTotalFiles - 1);
            string oldest = Path.Combine(directory, $"app_{maxBackupIndex}.log");
            if (File.Exists(oldest)) File.Delete(oldest);

            for (int i = maxBackupIndex - 1; i >= 1; i--)
            {
                string src = Path.Combine(directory, $"app_{i}.log");
                string dst = Path.Combine(directory, $"app_{i + 1}.log");
                if (File.Exists(src)) File.Move(src, dst);
            }

            string current = Path.Combine(directory, "app.log");
            if (File.Exists(current))
                File.Move(current, Path.Combine(directory, "app_1.log"));
        }
        catch { }
    }

    private static void RotateLogs() => RotateLogsInDirectory(LogsDirectory, MaxLogFiles);

    public static void Info(string msg, string? caller = null) => Log(LogLevel.INFO, msg, caller);
    public static void Warning(string msg, string? caller = null) => Log(LogLevel.WARNING, msg, caller);
    public static void Danger(string msg, string? caller = null) => Log(LogLevel.DANGER, msg, caller);
    public static void Success(string msg, string? caller = null) => Log(LogLevel.SUCCESS, msg, caller);

    /// <summary>Formata duração de operação de forma legível em ms/s.</summary>
    public static string FormatElapsed(Stopwatch sw) => FormatElapsed(sw.Elapsed);

    /// <summary>Formata duração de operação de forma legível em ms/s.</summary>
    public static string FormatElapsed(TimeSpan elapsed)
    {
        long ms = (long)elapsed.TotalMilliseconds;
        return ms < 1000
            ? $"{ms} ms"
            : $"{elapsed.TotalSeconds:F2}s ({ms} ms)";
    }

    /// <summary>Retorna apenas o status de preenchimento de um segredo sem nunca expor seu valor.</summary>
    public static string DescribeSecretPresence(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "(ausente)";
        return $"(configurado, {value.Trim().Length} chars)";
    }

    /// <summary>Registra exceção detalhada (tipo, HResult, mensagem, inner exception e duração opcional).</summary>
    public static void LogException(string caller, string operation, Exception ex, TimeSpan? elapsed = null, LogLevel level = LogLevel.DANGER)
    {
        string inner = ex.InnerException != null ? $" | Inner({ex.InnerException.GetType().Name}): {ex.InnerException.Message}" : string.Empty;
        string dur = elapsed.HasValue ? $" | Duração={FormatElapsed(elapsed.Value)}" : string.Empty;
        string hresult = ex.HResult != 0 ? $" | HRESULT=0x{ex.HResult:X8}" : string.Empty;
        Log(level, $"{operation} falhou ({ex.GetType().Name}{hresult}): {ex.Message}{inner}{dur}", caller);
    }
}
