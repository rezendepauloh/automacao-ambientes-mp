using System.Diagnostics;
using System.IO;
using System.Net.Http;

namespace AmbientesHub.Services;

public class BancadaDockerService
{
    private static readonly HttpClient HttpClient = new() { Timeout = TimeSpan.FromSeconds(3) };
    private Process? _activeProcess;
    private readonly object _processLock = new();

    public static string BancadaWindowsDir => ConfigService.BancadaWindowsDir;
    public static string BancadaWslDir => ConfigService.BancadaWslDir;
    public static string BancadaWslDistro => ConfigService.BancadaWslDistro;
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

            if (!TryGetValidatedTarget(out string distro, out string wslDir, out string error))
            {
                onOutput($"[ERRO] {error}");
                return;
            }

            // Comando a ser executado via WSL: define PYTHONUNBUFFERED=1, sobe os containers e conecta aos logs em tempo real
            string cmd = $"cd '{wslDir}' && export PYTHONUNBUFFERED=1 && docker compose up -d";
            if (streamLogs)
            {
                cmd += " && docker compose logs -f --tail=100";
            }

            RunProcess("wsl.exe", $"-d {distro} -- bash -lic \"{cmd}\"", onOutput);
        });
    }

    /// <summary>
    /// Fail-closed: distro e diretório WSL precisam estar configurados no .env e casar com a whitelist
    /// (sem espaços, aspas, $, crase, ;, &amp;, |, "..") antes de qualquer processo ser iniciado.
    /// </summary>
    public static bool TryGetValidatedTarget(out string distro, out string wslDir, out string error)
    {
        distro = BancadaWslDistro?.Trim() ?? string.Empty;
        wslDir = BancadaWslDir?.Trim() ?? string.Empty;
        return ValidateTarget(distro, wslDir, out error);
    }

    /// <summary>Validação pura (sem I/O) de distro + diretório WSL; usada pelo serviço e pelos testes.</summary>
    public static bool ValidateTarget(string? distro, string? wslDir, out string error)
    {
        error = string.Empty;
        distro = distro?.Trim() ?? string.Empty;
        wslDir = wslDir?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(distro) || string.IsNullOrWhiteSpace(wslDir))
        {
            error = "BANCADA_WSL_DISTRO e BANCADA_WSL_DIR não configurados no .env. Operação bloqueada.";
            return false;
        }

        if (!InputValidator.IsValidWslDistro(distro))
        {
            error = "BANCADA_WSL_DISTRO inválido (permitido: letras, números, ponto, hífen e sublinhado).";
            return false;
        }

        if (!InputValidator.IsValidWslPath(wslDir))
        {
            error = "BANCADA_WSL_DIR inválido (caminho POSIX absoluto, sem espaços ou metacaracteres de shell).";
            return false;
        }

        return true;
    }

    public async Task StopDockerBancadaAsync(Action<string> onOutput)
    {
        await Task.Run(() =>
        {
            if (!TryGetValidatedTarget(out string distro, out string wslDir, out string error))
            {
                onOutput($"[ERRO] {error}");
                return;
            }

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

