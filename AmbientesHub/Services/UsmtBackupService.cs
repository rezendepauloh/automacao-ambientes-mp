using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net.Http;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo em C# para Backup (ScanState) e Restore (LoadState) de perfis de usuários
/// via Windows USMT (User State Migration Tool), com suporte a execução local e remota.
/// Política estrita: nunca exclui perfis de usuários da máquina de origem ou destino.
/// </summary>
public class UsmtBackupService
{
    public const string AdkOfficialDownloadUrl = "https://go.microsoft.com/fwlink/?linkid=2271937";
    private const string RemoteStagingLocalDir = @"C:\Windows\Temp\AmbientesHub_USMT";
    private const long MaxAdkSetupBytes = 300L * 1024 * 1024;

    private readonly RemoteUserProfileService _profileService = new();

    public static bool IsValidUsmtDirectory(string? dir)
    {
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
            return false;

        return File.Exists(Path.Combine(dir, "scanstate.exe"))
            && File.Exists(Path.Combine(dir, "loadstate.exe"));
    }

    /// <summary>
    /// Localiza a pasta de binários do USMT em ordem de prioridade:
    /// caminho customizado/.env -> pasta Tools local do Hub -> instalações padrão do Windows ADK.
    /// </summary>
    public static string? FindUsmtDirectory(string? configuredPath = null)
    {
        string? envOrCustom = !string.IsNullOrWhiteSpace(configuredPath)
            ? configuredPath.Trim()
            : ConfigService.UsmtPath.Trim();

        string[] candidates =
        [
            envOrCustom,
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tools", "USMT", "amd64"),
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tools", "USMT"),
            @"C:\AmbientesVirtuais\AmbientesHub\Tools\USMT\amd64",
            @"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\User State Migration Tool\amd64",
            @"C:\Program Files\Windows Kits\10\Assessment and Deployment Kit\User State Migration Tool\amd64",
            @"C:\Program Files (x86)\Windows Kits\8.1\Assessment and Deployment Kit\User State Migration Tool\amd64",
            @"C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\User State Migration Tool\x86"
        ];

        foreach (string candidate in candidates)
        {
            if (IsValidUsmtDirectory(candidate))
                return candidate;
        }

        return null;
    }

    public List<RemoteUserProfileItem> QueryUsersForBackup(
        string computerName,
        Action<string>? onLog = null,
        bool calculateDiskSize = true)
    {
        return _profileService.QueryUserProfiles(computerName, onLog, calculateDiskSize);
    }

    /// <summary>
    /// Baixa e instala silenciosamente apenas o recurso USMT do Windows ADK oficial
    /// (ou via winget) e grava USMT_PATH automaticamente no .env.
    /// </summary>
    public async Task<(bool Success, string? UsmtPath, string Message)> DownloadAndInstallUsmtAsync(
        bool silent = true,
        Action<string>? onLog = null)
    {
        return await Task.Run(() => DownloadAndInstallUsmtCore(silent, onLog));
    }

    private static (bool Success, string? UsmtPath, string Message) DownloadAndInstallUsmtCore(
        bool silent,
        Action<string>? onLog)
    {
        var sw = Stopwatch.StartNew();
        void Log(string msg)
        {
            onLog?.Invoke(msg);
            AppLoggerService.Info(msg, nameof(UsmtBackupService));
        }

        string? existing = FindUsmtDirectory();
        if (existing != null)
        {
            ConfigService.SetAndSave("USMT_PATH", existing);
            Log($"[OK] USMT já detectado em: {existing} (salvo em USMT_PATH).");
            return (true, existing, $"USMT localizado em {existing}");
        }

        if (!SystemPerformanceService.IsUserAdministrator())
        {
            string err = "Instalar o Windows ADK / USMT exige privilégios de Administrador. Reinicie o Hub como Admin e tente novamente.";
            Log($"[ERRO] {err}");
            AppLoggerService.Warning("DownloadAndInstallUsmtCore abortado: requer privilégios de Administrador.", nameof(UsmtBackupService));
            return (false, null, err);
        }

        AppLoggerService.Info($"Iniciando download e instalação do Windows ADK USMT (silent={silent}).", nameof(UsmtBackupService));
        Log("==========================================================");
        Log("=== DOWNLOAD & INSTALAÇÃO DO WINDOWS USMT (ADK) ===");
        Log("==========================================================");

        string tempSetup = Path.Combine(Path.GetTempPath(), $"adksetup_{Guid.NewGuid():N}.exe");
        bool setupOk = false;

        try
        {
            Log($"[1/2] Baixando instalador oficial Microsoft ADK ({AdkOfficialDownloadUrl})...");
            using (var httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(5) })
            {
                using var response = httpClient.GetAsync(AdkOfficialDownloadUrl, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
                response.EnsureSuccessStatusCode();
                if (response.Content.Headers.ContentLength is long len && len > MaxAdkSetupBytes)
                    throw new InvalidDataException("Instalador ADK excedeu limite máximo permitido (300 MB).");

                using var netStream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
                using var fileStream = new FileStream(tempSetup, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                byte[] buffer = new byte[81920];
                long total = 0;
                int read;
                while ((read = netStream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += read;
                    if (total > MaxAdkSetupBytes)
                        throw new InvalidDataException("Download excedeu limite de segurança (300 MB).");
                    fileStream.Write(buffer, 0, read);
                }
            }

            Log($"  -> Download concluído ({new FileInfo(tempSetup).Length / (1024.0 * 1024.0):F2} MB).");
            Log("[2/2] Instalando recurso OptionId.UserStateMigrationTool (USMT)...");

            var psi = new ProcessStartInfo
            {
                FileName = tempSetup,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            if (silent)
            {
                psi.ArgumentList.Add("/quiet");
                psi.ArgumentList.Add("/norestart");
            }
            psi.ArgumentList.Add("/features");
            psi.ArgumentList.Add("OptionId.UserStateMigrationTool");

            using var proc = Process.Start(psi);
            if (proc != null)
            {
                proc.WaitForExit(600000);
                setupOk = proc.ExitCode == 0 || proc.ExitCode == 3010;
                Log($"  -> adksetup.exe encerrou com código {proc.ExitCode}.");
            }
        }
        catch (Exception ex)
        {
            Log($"  -> [AVISO] Falha via adksetup.exe direto: {ex.Message}. Tentando fallback via winget...");
            AppLoggerService.LogException(nameof(UsmtBackupService), "Download/instalação adksetup.exe -> fallback winget", ex, sw.Elapsed, LogLevel.WARNING);
        }
        finally
        {
            try { if (File.Exists(tempSetup)) File.Delete(tempSetup); } catch { }
        }

        if (!setupOk && FindUsmtDirectory() == null)
        {
            Log("[Fallback] Executando: winget install --id Microsoft.WindowsADK --exact --source winget...");
            AppLoggerService.Info("Acionando fallback winget para Microsoft.WindowsADK.", nameof(UsmtBackupService));
            var wingetPsi = new ProcessStartInfo
            {
                FileName = "winget.exe",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            foreach (var arg in new[] { "install", "--id", "Microsoft.WindowsADK", "--exact", "--source", "winget", "--accept-package-agreements", "--accept-source-agreements" })
                wingetPsi.ArgumentList.Add(arg);
            if (silent) wingetPsi.ArgumentList.Add("--silent");

            try
            {
                using var wProc = Process.Start(wingetPsi);
                if (wProc != null)
                {
                    wProc.OutputDataReceived += (_, e) => { if (!string.IsNullOrWhiteSpace(e.Data)) Log($"  {e.Data.Trim()}"); };
                    wProc.BeginOutputReadLine();
                    wProc.WaitForExit(600000);
                    AppLoggerService.Info($"Fallback winget Microsoft.WindowsADK encerrou com ExitCode={wProc.ExitCode}.", nameof(UsmtBackupService));
                }
            }
            catch (Exception ex)
            {
                Log($"  -> [ERRO] Falha no fallback winget: {ex.Message}");
                AppLoggerService.LogException(nameof(UsmtBackupService), "Fallback winget Microsoft.WindowsADK", ex, sw.Elapsed);
            }
        }

        string? installedDir = FindUsmtDirectory();
        if (installedDir != null)
        {
            ConfigService.SetAndSave("USMT_PATH", installedDir);
            Log($"[SUCESSO] USMT instalado e USMT_PATH preenchido automaticamente: {installedDir}");
            AppLoggerService.Success($"USMT instalado e configurado em USMT_PATH='{installedDir}' em {AppLoggerService.FormatElapsed(sw)}.", nameof(UsmtBackupService));
            return (true, installedDir, $"USMT instalado em {installedDir}");
        }

        string failMsg = "Não foi possível confirmar os executáveis scanstate.exe / loadstate.exe após a instalação.";
        Log($"[ERRO] {failMsg}");
        AppLoggerService.Danger($"DownloadAndInstallUsmtCore falhou após {AppLoggerService.FormatElapsed(sw)}: {failMsg}", nameof(UsmtBackupService));
        return (false, null, failMsg);
    }

    public static List<string> BuildScanStateArgumentList(
        string usmtExecutionDir,
        UsmtBackupOptions options,
        string? logFilePath = null)
    {
        if (string.IsNullOrWhiteSpace(usmtExecutionDir))
            throw new ArgumentException("Diretório do USMT não informado.", nameof(usmtExecutionDir));
        if (!InputValidator.IsSafeDirectoryPath(options.StorePath))
            throw new ArgumentException("Caminho de destino do backup inválido ou inseguro (fail-closed).", nameof(options));

        var validUsers = ValidateAndNormalizeUsers(options.SelectedUsers, requireAtLeastOne: true);

        var args = new List<string>
        {
            options.StorePath.Trim(),
            "/o",
            "/c",
            "/v:13"
        };

        AppendXmlConfigArgs(args, usmtExecutionDir, options);

        if (!options.Compress)
        {
            args.Add("/nocompress");
        }
        else if (!string.IsNullOrWhiteSpace(options.EncryptionKey))
        {
            args.Add("/encrypt");
            args.Add($"/key:{options.EncryptionKey.Trim()}");
        }

        // Seleciona SOMENTE os usuários escolhidos (sem nenhuma exclusão de perfil no Windows)
        args.Add(@"/ue:*\*");
        foreach (string user in validUsers)
        {
            string pattern = user.Contains('\\') ? user : $@"*\{user}";
            args.Add($"/ui:{pattern}");
        }

        if (!string.IsNullOrWhiteSpace(logFilePath))
            args.Add($"/l:{logFilePath.Trim()}");

        return args;
    }

    public static List<string> BuildLoadStateArgumentList(
        string usmtExecutionDir,
        UsmtBackupOptions options,
        string? logFilePath = null)
    {
        if (string.IsNullOrWhiteSpace(usmtExecutionDir))
            throw new ArgumentException("Diretório do USMT não informado.", nameof(usmtExecutionDir));
        if (!InputValidator.IsSafeDirectoryPath(options.StorePath))
            throw new ArgumentException("Caminho de origem do restore inválido ou inseguro (fail-closed).", nameof(options));

        var validUsers = ValidateAndNormalizeUsers(options.SelectedUsers, requireAtLeastOne: false);

        var args = new List<string>
        {
            options.StorePath.Trim(),
            "/c",
            "/v:13"
        };

        AppendXmlConfigArgs(args, usmtExecutionDir, options);

        if (options.CreateLocalAccountsOnRestore)
        {
            args.Add("/lac");
            args.Add("/lae");
        }

        if (!options.Compress)
        {
            args.Add("/nocompress");
        }
        else if (!string.IsNullOrWhiteSpace(options.EncryptionKey))
        {
            args.Add("/decrypt");
            args.Add($"/key:{options.EncryptionKey.Trim()}");
        }

        if (validUsers.Count > 0)
        {
            args.Add(@"/ue:*\*");
            foreach (string user in validUsers)
            {
                string pattern = user.Contains('\\') ? user : $@"*\{user}";
                args.Add($"/ui:{pattern}");
            }
        }

        if (!string.IsNullOrWhiteSpace(logFilePath))
            args.Add($"/l:{logFilePath.Trim()}");

        return args;
    }

    private static List<string> ValidateAndNormalizeUsers(IEnumerable<string> users, bool requireAtLeastOne)
    {
        var list = (users ?? [])
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .Select(u => u.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (string user in list)
        {
            if (!InputValidator.IsValidUserName(user))
                throw new ArgumentException($"Identificador de usuário inválido (fail-closed): '{user}'.");
        }

        if (requireAtLeastOne && list.Count == 0)
            throw new ArgumentException("Selecione pelo menos um usuário para realizar o backup USMT.");

        return list;
    }

    private static void AppendXmlConfigArgs(List<string> args, string usmtDir, UsmtBackupOptions options)
    {
        if (options.IncludeMigApp) args.Add($"/i:{Path.Combine(usmtDir, "MigApp.xml")}");
        if (options.IncludeMigDocs) args.Add($"/i:{Path.Combine(usmtDir, "MigDocs.xml")}");
        if (options.IncludeMigUser) args.Add($"/i:{Path.Combine(usmtDir, "MigUser.xml")}");
    }

    public Task<UsmtOperationResult> ExecuteBackupAsync(UsmtBackupOptions options, Action<string>? onLog = null)
        => Task.Run(() => ExecuteUsmtOperationCore(isRestore: false, options, onLog));

    public Task<UsmtOperationResult> ExecuteRestoreAsync(UsmtBackupOptions options, Action<string>? onLog = null)
        => Task.Run(() => ExecuteUsmtOperationCore(isRestore: true, options, onLog));

    private UsmtOperationResult ExecuteUsmtOperationCore(
        bool isRestore,
        UsmtBackupOptions options,
        Action<string>? onLog)
    {
        string target = RemoteManagementHelper.NormalizeTargetHost(options.TargetComputer);
        bool isLocal = RemoteManagementHelper.IsLocalMachine(options.TargetComputer);
        string modeLabel = isRestore ? "Restore (LoadState)" : "Backup (ScanState)";

        var result = new UsmtOperationResult
        {
            ComputerName = target,
            Mode = modeLabel,
            StorePath = options.StorePath?.Trim() ?? string.Empty
        };

        void Log(string msg) => onLog?.Invoke(msg);

        // 1. Fail-closed: validar host, caminho do store e binários USMT
        if (!isLocal && !InputValidator.IsValidHostName(target))
        {
            result.Message = "Nome de máquina de destino inválido (fail-closed).";
            Log($"[ERRO] {result.Message}");
            AppLoggerService.Warning($"USMT {modeLabel} rejeitado: host inválido '{target}'.", nameof(UsmtBackupService));
            return result;
        }

        if (!InputValidator.IsSafeDirectoryPath(options.StorePath))
        {
            result.Message = "Caminho de origem/destino inválido ou inseguro (fail-closed).";
            Log($"[ERRO] {result.Message}");
            AppLoggerService.Warning($"USMT {modeLabel} rejeitado: StorePath inválido '{options.StorePath}'.", nameof(UsmtBackupService));
            return result;
        }

        string? usmtDir = FindUsmtDirectory(options.CustomUsmtPath);
        if (usmtDir == null)
        {
            result.Message = "USMT (scanstate.exe / loadstate.exe) não encontrado. Configure USMT_PATH ou clique em 'Baixar / Instalar USMT'.";
            Log($"[ERRO] {result.Message}");
            AppLoggerService.Warning($"USMT {modeLabel} abortado: binários USMT não encontrados.", nameof(UsmtBackupService));
            return result;
        }

        List<string> args;
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string exeName = isRestore ? "loadstate.exe" : "scanstate.exe";
        string localLogPath = Path.Combine(AppLoggerService.LogDirectoryPath, $"usmt_{(isRestore ? "restore" : "backup")}_{target}_{timestamp}.log");
        result.LogFilePath = localLogPath;

        try
        {
            string execDir = isLocal ? usmtDir : RemoteStagingLocalDir;
            string remoteLogPath = Path.Combine(RemoteStagingLocalDir, "usmt_operation.log");
            args = isRestore
                ? BuildLoadStateArgumentList(execDir, options, isLocal ? localLogPath : remoteLogPath)
                : BuildScanStateArgumentList(execDir, options, isLocal ? localLogPath : remoteLogPath);
            result.ProcessedUsersCount = options.SelectedUsers.Count;
        }
        catch (Exception ex)
        {
            result.Message = ex.Message;
            Log($"[ERRO] {ex.Message}");
            AppLoggerService.Warning($"USMT {modeLabel} abortado na construção de argumentos: {ex.Message}", nameof(UsmtBackupService));
            return result;
        }

        AppLoggerService.Info(
            $"Iniciando USMT {modeLabel}: alvo='{target}' (local={isLocal}), store='{result.StorePath}', usuarios={options.SelectedUsers.Count}, compress={options.Compress}, encrypted={!string.IsNullOrWhiteSpace(options.EncryptionKey)}, log='{localLogPath}'.",
            nameof(UsmtBackupService));

        Log("==========================================================");
        Log($"=== USMT {modeLabel.ToUpperInvariant()} EM: {target} ===");
        Log($"  USMT Dir : {usmtDir}");
        Log($"  Store    : {result.StorePath}");
        Log($"  Usuários : {(options.SelectedUsers.Count > 0 ? string.Join(", ", options.SelectedUsers) : "Todos do pacote (Restore)")}");
        Log("  Política : Exclusão de usuários DESATIVADA (somente leitura/cópia)");
        Log("==========================================================");

        if (isLocal)
        {
            return ExecuteLocalUsmt(usmtDir, exeName, args, localLogPath, result, Log);
        }

        return ExecuteRemoteUsmt(target, usmtDir, exeName, args, localLogPath, isRestore, result.StorePath, result, Log);
    }

    private static UsmtOperationResult ExecuteLocalUsmt(
        string usmtDir,
        string exeName,
        List<string> args,
        string logPath,
        UsmtOperationResult result,
        Action<string> log)
    {
        if (!SystemPerformanceService.IsUserAdministrator())
        {
            result.Message = "Executar ScanState/LoadState localmente exige privilégios de Administrador.";
            log($"[ERRO] {result.Message}");
            AppLoggerService.Warning($"USMT Local ({result.Mode}) abortado: requer privilégios de Administrador.", nameof(UsmtBackupService));
            return result;
        }

        var sw = Stopwatch.StartNew();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
            if (!result.Mode.Contains("Restore", StringComparison.OrdinalIgnoreCase))
                Directory.CreateDirectory(result.StorePath);

            string exePath = Path.Combine(usmtDir, exeName);
            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                WorkingDirectory = usmtDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            foreach (string a in args)
                psi.ArgumentList.Add(a);

            using var proc = new Process { StartInfo = psi };
            proc.OutputDataReceived += (_, e) => { if (!string.IsNullOrWhiteSpace(e.Data)) log($"  {e.Data.Trim()}"); };
            proc.ErrorDataReceived += (_, e) => { if (!string.IsNullOrWhiteSpace(e.Data)) log($"  [AVISO] {e.Data.Trim()}"); };

            log($"-> Executando {exeName} localmente...");
            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            proc.WaitForExit();

            result.ExitCode = proc.ExitCode;
            result.Success = proc.ExitCode == 0 || proc.ExitCode == 3;
            result.Message = result.Success
                ? $"{result.Mode} concluído com sucesso! (ExitCode {proc.ExitCode})"
                : $"{result.Mode} encerrado com código {proc.ExitCode}. Consulte o log em {logPath}.";

            log(result.Success ? $"[SUCESSO] {result.Message}" : $"[ERRO] {result.Message}");
            AppLoggerService.Log(
                result.Success ? LogLevel.SUCCESS : LogLevel.DANGER,
                $"USMT Local ({result.Mode}) finalizado em '{result.ComputerName}': ExitCode={proc.ExitCode}, duração={AppLoggerService.FormatElapsed(sw)}.",
                nameof(UsmtBackupService));
        }
        catch (Exception ex)
        {
            result.Message = $"Falha ao executar USMT localmente: {ex.Message}";
            log($"[ERRO] {result.Message}");
            AppLoggerService.LogException(nameof(UsmtBackupService), $"USMT Local ({result.Mode})", ex, sw.Elapsed);
        }

        return result;
    }

    private static UsmtOperationResult ExecuteRemoteUsmt(
        string target,
        string localUsmtDir,
        string exeName,
        List<string> args,
        string localLogCopyPath,
        bool isRestore,
        string storePath,
        UsmtOperationResult result,
        Action<string> log)
    {
        var sw = Stopwatch.StartNew();
        if (!RemoteManagementHelper.PingHost(target))
        {
            result.Message = $"Máquina remota '{target}' não respondeu ao ping (ICMP).";
            log($"[ERRO] {result.Message}");
            AppLoggerService.Warning($"USMT Remoto ({result.Mode}) abortado em '{target}': sem resposta ICMP.", nameof(UsmtBackupService));
            return result;
        }

        if (!RemoteManagementHelper.EnsureAdminShareConnected(target))
        {
            result.Message = $"Não foi possível conectar ao compartilhamento administrativo \\\\{target}\\C$.";
            log($"[ERRO] {result.Message}");
            AppLoggerService.Warning($"USMT Remoto ({result.Mode}) abortado em '{target}': falha ao conectar em \\\\{target}\\C$.", nameof(UsmtBackupService));
            return result;
        }

        string remoteUncStaging = $@"\\{target}\C$\Windows\Temp\AmbientesHub_USMT";
        string remoteUncLog = Path.Combine(remoteUncStaging, "usmt_operation.log");

        try
        {
            if (!isRestore && storePath.StartsWith(@"\\", StringComparison.Ordinal))
            {
                Directory.CreateDirectory(storePath);
            }

            log($"[1/3] Copiando binários USMT para staging temporário em '{remoteUncStaging}'...");
            CopyDirectoryRecursive(localUsmtDir, remoteUncStaging);
            log("  -> [OK] Binários USMT transferidos para a estação remota.");

            string remoteExePath = Path.Combine(RemoteStagingLocalDir, exeName);
            string commandLine = $"\"{remoteExePath}\" " + string.Join(" ", args.Select(QuoteWin32ProcessArg));

            log($"[2/3] Disparando {exeName} remotamente via WMI (Win32_Process.Create)...");
            var scope = RemoteManagementHelper.CreateScope(target, @"root\cimv2");
            uint pid;
            using (var procClass = new ManagementClass(scope, new ManagementPath("Win32_Process"), null))
            {
                using var inParams = procClass.GetMethodParameters("Create");
                inParams["CommandLine"] = commandLine;
                inParams["CurrentDirectory"] = RemoteStagingLocalDir;
                using var outParams = procClass.InvokeMethod("Create", inParams, null);
                uint ret = outParams != null ? Convert.ToUInt32(outParams["ReturnValue"]) : 99;
                if (ret != 0)
                {
                    result.Message = $"Win32_Process.Create falhou na máquina remota (ReturnValue={ret}).";
                    log($"[ERRO] {result.Message}");
                    return result;
                }
                pid = Convert.ToUInt32(outParams!["ProcessId"]);
            }

            log($"  -> Processo remoto iniciado (PID {pid}). Acompanhando log em tempo real...");
            long lastLogPos = 0;

            while (true)
            {
                Thread.Sleep(1500);
                lastLogPos = StreamNewRemoteLogLines(remoteUncLog, lastLogPos, log);

                using var searcher = new ManagementObjectSearcher(scope, new ObjectQuery($"SELECT ProcessId FROM Win32_Process WHERE ProcessId = {pid}"));
                if (searcher.Get().Count == 0)
                    break;
            }

            // Captura linhas finais após término do processo
            Thread.Sleep(400);
            StreamNewRemoteLogLines(remoteUncLog, lastLogPos, log);

            if (File.Exists(remoteUncLog))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(localLogCopyPath)!);
                File.Copy(remoteUncLog, localLogCopyPath, overwrite: true);
                string tail = File.ReadAllText(remoteUncLog);
                bool noErrors = tail.Contains("Total errors: 0", StringComparison.OrdinalIgnoreCase)
                             || tail.Contains("Success", StringComparison.OrdinalIgnoreCase);
                bool hasFatal = tail.Contains("Fatal error", StringComparison.OrdinalIgnoreCase)
                             || tail.Contains("An error occurred", StringComparison.OrdinalIgnoreCase);

                result.Success = noErrors && !hasFatal;
                result.ExitCode = result.Success ? 0 : 1;
            }
            else
            {
                result.Success = false;
                result.Message = "Processo remoto encerrou sem gerar arquivo de log USMT.";
                log($"[ERRO] {result.Message}");
                return result;
            }

            result.Message = result.Success
                ? $"{result.Mode} remoto em '{target}' concluído com sucesso!"
                : $"{result.Mode} remoto em '{target}' finalizou com alertas. Verifique {localLogCopyPath}.";

            log(result.Success ? $"[SUCESSO] {result.Message}" : $"[AVISO] {result.Message}");
            AppLoggerService.Log(
                result.Success ? LogLevel.SUCCESS : LogLevel.WARNING,
                $"USMT Remoto ({result.Mode}) em '{target}': Success={result.Success}, ExitCode={result.ExitCode}, duração={AppLoggerService.FormatElapsed(sw)}.",
                nameof(UsmtBackupService));
        }
        catch (Exception ex)
        {
            result.Message = $"Erro durante execução remota do USMT em '{target}': {ex.Message}";
            log($"[ERRO] {result.Message}");
            AppLoggerService.LogException(nameof(UsmtBackupService), $"USMT Remoto ({result.Mode}) em '{target}'", ex, sw.Elapsed);
        }
        finally
        {
            log("[3/3] Limpando pasta temporária de staging USMT (perfis de usuário intactos)...");
            try
            {
                if (Directory.Exists(remoteUncStaging))
                    Directory.Delete(remoteUncStaging, recursive: true);
            }
            catch (Exception exClean)
            {
                AppLoggerService.Warning($"Falha ao remover staging temporário '{remoteUncStaging}': {exClean.Message}", nameof(UsmtBackupService));
            }
        }

        return result;
    }

    private static long StreamNewRemoteLogLines(string logFile, long lastPos, Action<string> log)
    {
        try
        {
            if (!File.Exists(logFile)) return lastPos;
            using var fs = new FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (fs.Length <= lastPos) return lastPos;
            fs.Seek(lastPos, SeekOrigin.Begin);
            using var reader = new StreamReader(fs);
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (line.Contains("Error", StringComparison.OrdinalIgnoreCase) ||
                    line.Contains("Warning", StringComparison.OrdinalIgnoreCase) ||
                    line.Contains("Success", StringComparison.OrdinalIgnoreCase) ||
                    line.Contains("Processing", StringComparison.OrdinalIgnoreCase) ||
                    line.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    log($"  {line.Trim()}");
                }
            }
            return fs.Position;
        }
        catch
        {
            return lastPos;
        }
    }

    private static void CopyDirectoryRecursive(string sourceDir, string destDir)
    {
        Directory.CreateDirectory(destDir);
        foreach (string file in Directory.GetFiles(sourceDir))
        {
            File.Copy(file, Path.Combine(destDir, Path.GetFileName(file)), overwrite: true);
        }
        foreach (string subDir in Directory.GetDirectories(sourceDir))
        {
            CopyDirectoryRecursive(subDir, Path.Combine(destDir, Path.GetFileName(subDir)));
        }
    }

    private static string QuoteWin32ProcessArg(string arg)
    {
        if (string.IsNullOrEmpty(arg)) return "\"\"";
        return arg.Contains(' ') ? $"\"{arg}\"" : arg;
    }
}
