using System.Diagnostics;
using System.IO;
using System.Net.Http;

namespace AmbientesHub.Services;

public class BancadaDockerService
{
    private static readonly HttpClient HttpClient = new() { Timeout = TimeSpan.FromSeconds(3) };
    private Process? _activeProcess;
    private readonly object _processLock = new();

    public static string BancadaWindowsDir => ConfigService.Get("BANCADA_WINDOWS_DIR", @"\\wsl.localhost\Ubuntu-26.04\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart");
    public static string BancadaWslDir => ConfigService.Get("BANCADA_WSL_DIR", "/home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart");
    public static string BancadaWslDistro => ConfigService.Get("BANCADA_WSL_DISTRO", "Ubuntu-26.04");
    public static string BancadaUrl => ConfigService.UrlBancadaLocal;

    public bool IsProcessRunning
    {
        get
        {
            lock (_processLock)
            {
                return _activeProcess != null && !_activeProcess.HasExited;
            }
        }
    }

    public async Task<bool> IsBancadaOnlineAsync()
    {
        try
        {
            var uri = new Uri(BancadaUrl.TrimEnd('/') + "/_stcore/health");
            var res = await HttpClient.GetAsync(uri);
            return res.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public async Task StartDockerBancadaAsync(Action<string> onOutput, bool streamLogs = true)
    {
        await Task.Run(() =>
        {
            lock (_processLock)
            {
                if (_activeProcess != null && !_activeProcess.HasExited)
                {
                    onOutput("[Aviso] O processo Docker já está em execução no Hub.");
                    return;
                }
            }

            onOutput("[WSL/Docker] Preparando ambiente da Bancada...");

            string wslDir = BancadaWslDir.Replace("'", @"'\''");
            string distro = BancadaWslDistro;

            // Comando a ser executado via WSL: cd para o diretório e executa docker compose up -d
            // E em seguida conecta aos logs se streamLogs for true
            string cmd = $"cd '{wslDir}' && docker compose up -d";
            if (streamLogs)
            {
                cmd += " && docker compose logs -f --tail=100 web";
            }

            RunProcess("wsl.exe", $"-d {distro} -- bash -lic \"{cmd}\"", onOutput);
        });
    }

    public async Task StopDockerBancadaAsync(Action<string> onOutput)
    {
        await Task.Run(() =>
        {
            onOutput("[WSL/Docker] Solicitando parada dos serviços da Bancada (docker compose down)...");

            lock (_processLock)
            {
                if (_activeProcess != null && !_activeProcess.HasExited)
                {
                    try
                    {
                        _activeProcess.Kill(true);
                    }
                    catch { }
                    _activeProcess = null;
                }
            }

            string wslDir = BancadaWslDir.Replace("'", @"'\''");
            string distro = BancadaWslDistro;
            string cmd = $"cd '{wslDir}' && docker compose down";

            RunProcess("wsl.exe", $"-d {distro} -- bash -lic \"{cmd}\"", onOutput, waitForExit: true);
            onOutput("[WSL/Docker] Containers encerrados.");
        });
    }

    public async Task RestartDockerBancadaAsync(Action<string> onOutput)
    {
        await StopDockerBancadaAsync(onOutput);
        await Task.Delay(1500);
        await StartDockerBancadaAsync(onOutput);
    }

    private void RunProcess(string filename, string arguments, Action<string> onOutput, bool waitForExit = false)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = filename,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8
            };

            var p = new Process { StartInfo = psi, EnableRaisingEvents = true };

            p.OutputDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    onOutput(e.Data);
            };

            p.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    onOutput(e.Data);
            };

            lock (_processLock)
            {
                _activeProcess = p;
            }

            p.Start();
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();

            if (waitForExit)
            {
                p.WaitForExit();
                lock (_processLock)
                {
                    if (_activeProcess == p) _activeProcess = null;
                }
            }
        }
        catch (Exception ex)
        {
            onOutput($"[ERRO] Falha ao executar {filename}: {ex.Message}");
            AppLoggerService.Danger($"Erro ao executar {filename}: {ex.Message}", nameof(BancadaDockerService));
        }
    }
}

