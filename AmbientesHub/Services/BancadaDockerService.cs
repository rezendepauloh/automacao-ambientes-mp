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

    public async Task AttachLogsAsync(Action<string> onOutput)
    {
        await Task.Run(() =>
        {
            lock (_processLock)
            {
                if (_activeProcess != null && !_activeProcess.HasExited)
                {
                    AppLoggerService.Info("AttachLogsAsync ignorado: stream de logs Docker já ativo.", nameof(BancadaDockerService));
                    return; // Já conectado
                }
            }

            if (!TryGetValidatedTarget(out string distro, out string wslDir, out string error))
            {
                AppLoggerService.Warning($"AttachLogsAsync bloqueado: {error}", nameof(BancadaDockerService));
                return;
            }

            AppLoggerService.Info($"Conectando ao stream de logs Docker Compose (distro='{distro}', dir='{wslDir}').", nameof(BancadaDockerService));
            string cmd = $"cd '{wslDir}' && docker compose logs -f --tail=30";
            RunWslCommand(distro, cmd, onOutput);
        });
    }

    public async Task<bool> IsBancadaOnlineAsync()
    {
        var sw = Stopwatch.StartNew();
        try
        {
            if (InputValidator.IsSafeHttpUrl(BancadaUrl))
            {
                var uri = new Uri(BancadaUrl.TrimEnd('/') + "/_stcore/health");
                var res = await HttpClient.GetAsync(uri);
                if (res.IsSuccessStatusCode)
                {
                    AppLoggerService.Info($"Bancada Streamlit online em '{uri}' (HTTP {(int)res.StatusCode}) [{AppLoggerService.FormatElapsed(sw)}].", nameof(BancadaDockerService));
                    return true;
                }
                AppLoggerService.Info($"Healthcheck Streamlit '{uri}' retornou HTTP {(int)res.StatusCode}; tentando fallback para API REST.", nameof(BancadaDockerService));
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Info($"Healthcheck Streamlit '{BancadaUrl}' indisponível ({ex.Message}); tentando fallback para BancadaApiClient.CheckHealthAsync().", nameof(BancadaDockerService));
        }

        bool apiOk = await BancadaApiClient.CheckHealthAsync();
        AppLoggerService.Info($"Status final Bancada Online={apiOk} em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaDockerService));
        return apiOk;
    }

    public async Task StartDockerBancadaAsync(Action<string> onOutput, bool streamLogs = true)
    {
        var sw = Stopwatch.StartNew();
        await Task.Run(() =>
        {
            lock (_processLock)
            {
                if (_activeProcess != null && !_activeProcess.HasExited)
                {
                    AppLoggerService.Warning("StartDockerBancadaAsync ignorado: processo Docker já ativo no Hub.", nameof(BancadaDockerService));
                    onOutput("[Aviso] O processo Docker já está em execução no Hub.");
                    return;
                }
            }

            onOutput("[WSL/Docker] Preparando ambiente da Bancada...");

            if (!TryGetValidatedTarget(out string distro, out string wslDir, out string error))
            {
                AppLoggerService.Warning($"StartDockerBancadaAsync bloqueado (fail-closed): {error}", nameof(BancadaDockerService));
                onOutput($"[ERRO] {error}");
                return;
            }

            AppLoggerService.Info($"Iniciando Bancada Docker (distro='{distro}', dir='{wslDir}', streamLogs={streamLogs}).", nameof(BancadaDockerService));
            // Comando a ser executado via WSL: define PYTHONUNBUFFERED=1, sobe os containers e conecta aos logs em tempo real
            string cmd = $"cd '{wslDir}' && export PYTHONUNBUFFERED=1 && docker compose up -d";
            if (streamLogs)
            {
                cmd += " && docker compose logs -f --tail=30";
            }

            RunWslCommand(distro, cmd, onOutput);
            AppLoggerService.Info($"StartDockerBancadaAsync disparado em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaDockerService));
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
        var sw = Stopwatch.StartNew();
        await Task.Run(() =>
        {
            if (!TryGetValidatedTarget(out string distro, out string wslDir, out string error))
            {
                AppLoggerService.Warning($"StopDockerBancadaAsync bloqueado (fail-closed): {error}", nameof(BancadaDockerService));
                onOutput($"[ERRO] {error}");
                return;
            }

            AppLoggerService.Info($"Solicitando parada dos containers da Bancada (distro='{distro}', dir='{wslDir}').", nameof(BancadaDockerService));
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

            RunWslCommand(distro, cmd, onOutput, waitForExit: true);
            onOutput("[WSL/Docker] Containers encerrados.");
            AppLoggerService.Success($"Bancada Docker parada em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaDockerService));
        });
    }

    public async Task RestartDockerBancadaAsync(Action<string> onOutput)
    {
        var sw = Stopwatch.StartNew();
        AppLoggerService.Info("Reiniciando stack Docker da Bancada...", nameof(BancadaDockerService));
        await StopDockerBancadaAsync(onOutput);
        await Task.Delay(1500);
        await StartDockerBancadaAsync(onOutput);
        AppLoggerService.Info($"RestartDockerBancadaAsync concluído em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaDockerService));
    }

    public async Task RebuildDockerBancadaAsync(Action<string> onOutput, bool streamLogs = true)
    {
        var sw = Stopwatch.StartNew();
        await Task.Run(() =>
        {
            if (!TryGetValidatedTarget(out string distro, out string wslDir, out string error))
            {
                AppLoggerService.Warning($"RebuildDockerBancadaAsync bloqueado (fail-closed): {error}", nameof(BancadaDockerService));
                onOutput($"[ERRO] {error}");
                return;
            }

            AppLoggerService.Info($"Iniciando rebuild --no-cache da Bancada Docker (distro='{distro}', dir='{wslDir}').", nameof(BancadaDockerService));
            onOutput("[WSL/Docker] Iniciando reconstrução da stack (docker compose build --no-cache)...");

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

            // Para os containers e reconstrói as imagens
            string buildCmd = $"cd '{wslDir}' && docker compose down && export PYTHONUNBUFFERED=1 && docker compose build --no-cache && docker compose up -d";
            if (streamLogs)
            {
                buildCmd += " && docker compose logs -f --tail=30";
            }

            RunWslCommand(distro, buildCmd, onOutput);
            AppLoggerService.Info($"RebuildDockerBancadaAsync disparado em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaDockerService));
        });
    }

    /// <summary>Extrai o host remoto de <see cref="BancadaUrl"/> quando não for localhost/loopback.</summary>
    public static bool TryGetRemoteHost(string? url, out string remoteHost)
    {
        remoteHost = string.Empty;
        if (string.IsNullOrWhiteSpace(url) || !Uri.TryCreate(url.Trim(), UriKind.Absolute, out var uri))
            return false;

        string host = uri.Host;
        if (string.IsNullOrWhiteSpace(host) ||
            host.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
            host.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
            host.Equals("::1", StringComparison.OrdinalIgnoreCase))
            return false;

        if (!InputValidator.IsValidHostName(host))
            return false;

        remoteHost = host;
        return true;
    }

    private void RunWslCommand(string distro, string bashCmd, Action<string> onOutput, bool waitForExit = false)
    {
        if (TryGetRemoteHost(BancadaUrl, out string remoteHost))
        {
            string escapedBash = bashCmd.Replace("\"", "`\"");
            string psCommand = $"Invoke-Command -ComputerName '{remoteHost}' -ScriptBlock {{ wsl.exe -d {distro} -- bash -lic \"{escapedBash}\" }}";
            AppLoggerService.Info($"BancadaDockerService modo remoto via WinRM -> '{remoteHost}' (distro='{distro}').", nameof(BancadaDockerService));
            onOutput($"[WinRM -> {remoteHost}] Executando no WSL remoto ({distro})...");
            RunProcess("powershell.exe", $"-NoProfile -NonInteractive -Command \"{psCommand.Replace("\"", "\\\"")}\"", onOutput, waitForExit);
            return;
        }

        AppLoggerService.Info($"BancadaDockerService modo local via wsl.exe (distro='{distro}').", nameof(BancadaDockerService));
        RunProcess("wsl.exe", $"-d {distro} -- bash -lic \"{bashCmd}\"", onOutput, waitForExit);
    }

    private void RunProcess(string filename, string arguments, Action<string> onOutput, bool waitForExit = false)
    {
        var sw = Stopwatch.StartNew();
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
            AppLoggerService.Info($"Processo '{filename}' iniciado (PID={p.Id}, waitForExit={waitForExit}).", nameof(BancadaDockerService));

            if (waitForExit)
            {
                p.WaitForExit();
                AppLoggerService.Info($"Processo '{filename}' (PID={p.Id}) encerrou com ExitCode={p.ExitCode} em {AppLoggerService.FormatElapsed(sw)}.", nameof(BancadaDockerService));
                lock (_processLock)
                {
                    if (_activeProcess == p) _activeProcess = null;
                }
            }
        }
        catch (Exception ex)
        {
            onOutput($"[ERRO] Falha ao executar {filename}: {ex.Message}");
            AppLoggerService.LogException(nameof(BancadaDockerService), $"Executar '{filename}'", ex, sw.Elapsed);
        }
    }
}

