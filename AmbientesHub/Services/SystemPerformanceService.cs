using System.Diagnostics;
using System.IO;
using System.Management;
using System.Security.Principal;
using System.Text;
using AmbientesHub.Models;
using Microsoft.Win32;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo em C# responsável pela auditoria completa de lentidão, recursos,
/// bloatware, WSL/Docker e aplicação de otimizações de desempenho no Windows.
/// </summary>
public class SystemPerformanceService
{
    public static readonly string[] TelemetryAuditServiceNames =
    [
        "DiagTrack", "dmwappushservice", "WerSvc", "RetailDemo", "MapsBroker", "SysMain", "WSearch"
    ];

    public static readonly Dictionary<string, string> DebloatTargetServices = new(StringComparer.OrdinalIgnoreCase)
    {
        ["DiagTrack"] = "Automatic",
        ["dmwappushservice"] = "Manual",
        ["WerSvc"] = "Manual",
        ["RetailDemo"] = "Manual",
        ["MapsBroker"] = "Automatic",
        ["SysMain"] = "Automatic"
    };

    public static readonly string[] BloatAuditTokens =
    [
        "Xbox", "Zune", "Cortana", "FeedbackHub", "GetHelp",
        "Getstarted", "Solitaire", "BingNews", "BingWeather",
        "YourPhone", "QuickAssist"
    ];

    public static readonly string[] BloatRemovalTokens =
    [
        "Microsoft.549981C3F5F10",
        "Microsoft.BingNews",
        "Microsoft.BingWeather",
        "Microsoft.GetHelp",
        "Microsoft.Getstarted",
        "Microsoft.MicrosoftFeedbackHub",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.XboxApp",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.YourPhone"
    ];

    public static bool IsUserAdministrator()
    {
        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    public static bool MatchesAnyToken(string candidate, IEnumerable<string> tokens)
    {
        if (string.IsNullOrWhiteSpace(candidate))
            return false;

        foreach (var token in tokens)
        {
            if (!string.IsNullOrWhiteSpace(token) &&
                candidate.Contains(token, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    public static string SanitizeWslMemoryLimit(string? memoryLimit)
    {
        if (string.IsNullOrWhiteSpace(memoryLimit))
            return "3GB";

        string trimmed = memoryLimit.Trim().ToUpperInvariant().Replace(" ", "");
        if (trimmed.EndsWith("GB") || trimmed.EndsWith("MB"))
            return trimmed;

        if (int.TryParse(trimmed, out int numeric) && numeric > 0)
            return $"{numeric}GB";

        return "3GB";
    }

    public static string BuildWslConfigContent(string? memoryLimit = "4GB", int processors = 4, string swap = "2GB")
    {
        string safeMem = SanitizeWslMemoryLimit(memoryLimit);
        int safeProc = processors > 0 ? processors : 4;
        string safeSwap = string.IsNullOrWhiteSpace(swap) ? "2GB" : swap.Trim().ToUpperInvariant();

        return string.Join(Environment.NewLine,
        [
            "[wsl2]",
            $"memory={safeMem}",
            $"processors={safeProc}",
            $"swap={safeSwap}",
            "networkingMode=nat",
            "autoProxy=true",
            "dnsTunneling=false",
            "[experimental]",
            "autoMemoryReclaim=gradual",
            string.Empty
        ]);
    }

    public static (bool? isTrimEnabled, string summary) ParseTrimStatus(string fsutilOutput)
    {
        if (string.IsNullOrWhiteSpace(fsutilOutput))
            return (null, "Não foi possível consultar DisableDeleteNotify.");

        foreach (string rawLine in fsutilOutput.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            string line = rawLine.Trim();
            if (line.Contains("NTFS", StringComparison.OrdinalIgnoreCase) &&
                line.Contains("DisableDeleteNotify", StringComparison.OrdinalIgnoreCase))
            {
                if (line.Contains("= 0"))
                    return (true, $"TRIM ATIVO ({line})");
                if (line.Contains("= 1"))
                    return (false, $"TRIM DESATIVADO ({line})");
            }
        }

        if (fsutilOutput.Contains("= 0"))
            return (true, "TRIM ATIVO (DisableDeleteNotify = 0)");
        if (fsutilOutput.Contains("= 1"))
            return (false, "TRIM DESATIVADO (DisableDeleteNotify = 1)");

        return (null, NormalizeCliOutput(fsutilOutput));
    }

    public static string NormalizeCliOutput(string raw)
    {
        if (string.IsNullOrEmpty(raw))
            return string.Empty;

        string cleaned = raw.Replace("\0", string.Empty).Trim();
        return cleaned;
    }

    public async Task<SystemAuditReport> RunFullAuditAsync(Action<string>? onLog = null)
    {
        return await Task.Run(() => RunFullAuditCore(onLog));
    }

    private SystemAuditReport RunFullAuditCore(Action<string>? onLog)
    {
        void Log(string msg)
        {
            onLog?.Invoke(msg);
            AppLoggerService.Info(msg, nameof(SystemPerformanceService));
        }

        var report = new SystemAuditReport
        {
            IsAdministrator = IsUserAdministrator()
        };

        Log("==========================================================");
        Log("=== AUDITORIA DE PERFORMANCE, BLOAT & SAÚDE DO SISTEMA ===");
        Log("==========================================================");
        Log(report.IsAdministrator
            ? "[OK] Privilégios: Executando como Administrador."
            : "[AVISO] Privilégios: Modo usuário padrão (auditoria funciona 100%; otimizações de sistema requerem Admin).");

        // [1/9] Discos Disponíveis
        Log("[1/9] Verificando espaço em disco (FileSystem)...");
        try
        {
            foreach (var drive in DriveInfo.GetDrives().Where(d => d.IsReady && d.DriveType == DriveType.Fixed))
            {
                double freeGb = Math.Round(drive.AvailableFreeSpace / (1024.0 * 1024.0 * 1024.0), 2);
                double totalGb = Math.Round(drive.TotalSize / (1024.0 * 1024.0 * 1024.0), 2);
                double usedGb = Math.Round(totalGb - freeGb, 2);
                var item = new DiskStatusItem(drive.Name.TrimEnd('\\'), freeGb, usedGb, totalGb);
                report.Disks.Add(item);
                Log($"  -> {item.Summary}");
            }
        }
        catch (Exception ex)
        {
            Log($"  [ERRO] Falha ao ler discos: {ex.Message}");
        }

        // [2/9] Memória RAM Física (WMI)
        Log("[2/9] Consultando Memória RAM (Win32_OperatingSystem)...");
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem");
            foreach (ManagementObject obj in searcher.Get())
            {
                using (obj)
                {
                    double totalKb = Convert.ToDouble(obj["TotalVisibleMemorySize"] ?? 0);
                    double freeKb = Convert.ToDouble(obj["FreePhysicalMemory"] ?? 0);
                    report.TotalRamGb = Math.Round(totalKb / (1024.0 * 1024.0), 2);
                    report.FreeRamGb = Math.Round(freeKb / (1024.0 * 1024.0), 2);
                    report.FreeRamPercent = totalKb > 0 ? Math.Round((freeKb / totalKb) * 100.0, 1) : 0;
                    Log($"  -> Total: {report.TotalRamGb:F2} GB  |  Livre: {report.FreeRamGb:F2} GB ({report.FreeRamPercent:F1}%)");
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            Log($"  [ERRO] Falha ao consultar RAM via WMI: {ex.Message}");
        }

        // [3/9] Top 10 Processos por Memória + SearchHost
        Log("[3/9] Analisando Top 10 processos por consumo de RAM e SearchHost...");
        try
        {
            var grouped = new Dictionary<string, (long bytes, int count)>(StringComparer.OrdinalIgnoreCase);
            bool searchHostRunning = false;

            foreach (var proc in Process.GetProcesses())
            {
                try
                {
                    string name = proc.ProcessName;
                    if (name.Equals("SearchHost", StringComparison.OrdinalIgnoreCase))
                        searchHostRunning = true;

                    long ws = proc.WorkingSet64;
                    if (grouped.TryGetValue(name, out var existing))
                        grouped[name] = (existing.bytes + ws, existing.count + 1);
                    else
                        grouped[name] = (ws, 1);
                }
                catch
                {
                    // Processos protegidos de sistema podem negar leitura de WorkingSet64
                }
                finally
                {
                    proc.Dispose();
                }
            }

            report.IsSearchHostRunning = searchHostRunning;
            report.TopProcesses = grouped
                .Select(kv => new ProcessMemoryItem(kv.Key, Math.Round(kv.Value.bytes / (1024.0 * 1024.0), 2), kv.Value.count))
                .OrderByDescending(p => p.TotalRamMb)
                .Take(10)
                .ToList();

            foreach (var p in report.TopProcesses)
            {
                Log($"  -> {p.ProcessName,-28} {p.TotalRamMb,9:F2} MB  ({p.InstanceCount} inst.)");
            }

            Log(searchHostRunning
                ? "  -> [AVISO] SearchHost está em execução em segundo plano."
                : "  -> [OK] SearchHost inativo.");
        }
        catch (Exception ex)
        {
            Log($"  [ERRO] Falha ao listar processos: {ex.Message}");
        }

        // [4/9] Hibernação & Cache Windows Update
        Log("[4/9] Inspecionando hiberfil.sys e cache do Windows Update...");
        try
        {
            var hiberInfo = new FileInfo(@"C:\hiberfil.sys");
            if (hiberInfo.Exists)
            {
                report.HiberfilExists = true;
                report.HiberfilSizeGb = Math.Round(hiberInfo.Length / (1024.0 * 1024.0 * 1024.0), 2);
                Log($"  -> [AVISO] C:\\hiberfil.sys EXISTE ocupando {report.HiberfilSizeGb:F2} GB no drive C:");
            }
            else
            {
                report.HiberfilExists = false;
                report.HiberfilSizeGb = 0;
                Log("  -> [OK] Hibernação já está desativada (economizando espaço em C:).");
            }

            string softDist = @"C:\Windows\SoftwareDistribution\Download";
            if (Directory.Exists(softDist))
            {
                long totalBytes = GetDirectorySizeSafe(softDist);
                report.SoftwareDistributionDownloadMb = Math.Round(totalBytes / (1024.0 * 1024.0), 2);
                Log($"  -> Cache Windows Update (SoftwareDistribution\\Download): {report.SoftwareDistributionDownloadMb:F2} MB");
            }
        }
        catch (Exception ex)
        {
            Log($"  -> Aviso ao inspecionar arquivos de armazenamento: {ex.Message}");
        }

        // [5/9] Plano de Energia & TRIM SSD
        Log("[5/9] Verificando Plano de Energia ativo e status TRIM (fsutil)...");
        try
        {
            var powerRes = RunCliTool("powercfg.exe", "/getactivescheme", 5000);
            report.ActivePowerPlan = string.IsNullOrWhiteSpace(powerRes.stdout)
                ? "Indisponível"
                : NormalizeCliOutput(powerRes.stdout);
            Log($"  -> Plano de Energia: {report.ActivePowerPlan}");

            var trimRes = RunCliTool("fsutil.exe", "behavior query DisableDeleteNotify", 5000);
            var (trimEnabled, trimSummary) = ParseTrimStatus(trimRes.stdout);
            report.IsTrimEnabled = trimEnabled;
            report.TrimStatusText = trimSummary;
            Log(trimEnabled == true
                ? $"  -> [OK] {trimSummary}"
                : $"  -> [AVISO] {trimSummary}");
        }
        catch (Exception ex)
        {
            Log($"  -> Aviso ao verificar energia/TRIM: {ex.Message}");
        }

        // [6/9] Serviços de Telemetria e Diagnóstico
        Log("[6/9] Auditando serviços nativos de telemetria e diagnóstico...");
        try
        {
            string filter = string.Join(" OR ", TelemetryAuditServiceNames.Select(s => $"Name = '{s}'"));
            using var searcher = new ManagementObjectSearcher($"SELECT Name, DisplayName, State, StartMode FROM Win32_Service WHERE {filter}");
            foreach (ManagementObject svc in searcher.Get())
            {
                using (svc)
                {
                    var item = new TelemetryServiceStatus(
                        svc["Name"]?.ToString() ?? "",
                        svc["DisplayName"]?.ToString() ?? "",
                        svc["State"]?.ToString() ?? "Unknown",
                        svc["StartMode"]?.ToString() ?? "Unknown");
                    report.TelemetryServices.Add(item);
                    Log($"  -> {item.Name,-18} | Estado: {item.State,-10} | Inicialização: {item.StartMode}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"  [ERRO] Falha ao consultar serviços via WMI: {ex.Message}");
        }

        // [7/9] Aplicativos Bloatware / UWP Detectados (via Registro nativo HKCU)
        Log("[7/9] Buscando pacotes Bloatware UWP instalados para o usuário...");
        try
        {
            var foundApps = DetectInstalledBloatApps();
            report.DetectedBloatApps = foundApps;
            if (foundApps.Count == 0)
            {
                Log("  -> [OK] Nenhum bloatware UWP da lista alvo encontrado.");
            }
            else
            {
                foreach (var app in foundApps)
                {
                    Log($"  -> [AVISO] Detectado: {app}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"  -> Aviso ao ler pacotes UWP no registro: {ex.Message}");
        }

        // [8/9] Itens na Inicialização Automática (Startup Registry)
        Log("[8/9] Auditando aplicativos na inicialização automática (HKCU/HKLM Run)...");
        try
        {
            report.StartupItems = ReadStartupEntries();
            if (report.StartupItems.Count == 0)
            {
                Log("  -> Nenhum item de Run encontrado.");
            }
            else
            {
                foreach (var st in report.StartupItems)
                {
                    Log($"  -> [{st.Origin}] {st.Name}: {st.Command}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"  -> Aviso ao ler inicialização: {ex.Message}");
        }

        // [9/9] WSL, VHDX e Docker
        Log("[9/9] Verificando WSL 2, discos virtuais VHDX e Docker...");
        try
        {
            var wslRes = RunCliTool("wsl.exe", "--list -v", 6000, Encoding.Unicode);
            report.WslStatusOutput = NormalizeCliOutput(wslRes.stdout);
            if (!string.IsNullOrWhiteSpace(report.WslStatusOutput))
            {
                foreach (var line in report.WslStatusOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    Log($"  [WSL] {line.Trim()}");
                }
            }
            else
            {
                Log("  -> WSL não instalado ou sem distros registradas.");
            }

            report.VhdxFiles = FindVhdxFiles();
            if (report.VhdxFiles.Count > 0)
            {
                foreach (var vhdx in report.VhdxFiles)
                {
                    Log($"  -> VHDX: {vhdx.FullPath} ({vhdx.SizeGb:F2} GB)");
                }
            }
            else
            {
                Log("  -> Nenhum disco virtual .vhdx localizado nos diretórios padrão.");
            }

            var dockerRes = RunDockerCommand("system df", 5000);
            if (dockerRes.exitCode == 0 && !string.IsNullOrWhiteSpace(dockerRes.stdout))
            {
                report.IsDockerAvailable = true;
                report.DockerStatusOutput = NormalizeCliOutput(dockerRes.stdout);
                foreach (var line in report.DockerStatusOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    Log($"  [Docker] {line.Trim()}");
                }
            }
            else
            {
                report.IsDockerAvailable = false;
                Log("  -> Docker daemon não está rodando no Windows ou no WSL padrão.");
            }
        }
        catch (Exception ex)
        {
            Log($"  -> Aviso ao verificar WSL/Docker: {ex.Message}");
        }

        Log("==========================================================");
        Log("[SUCESSO] Auditoria completa concluída!");
        return report;
    }

    public async Task ApplyOptimizationsAsync(OptimizationOptions options, Action<string>? onLog = null)
    {
        options ??= new OptimizationOptions();
        await Task.Run(() => ApplyOptimizationsCore(options, onLog));
    }

    private void ApplyOptimizationsCore(OptimizationOptions options, Action<string>? onLog)
    {
        void Log(string msg)
        {
            onLog?.Invoke(msg);
            AppLoggerService.Info(msg, nameof(SystemPerformanceService));
        }

        bool isAdmin = IsUserAdministrator();
        Log("==========================================================");
        Log("=== APLICANDO OTIMIZAÇÃO DE DESEMPENHO & DEBLOAT (C#) ===");
        Log("==========================================================");

        if (!isAdmin)
        {
            Log("[AVISO] Aplicação não está elevada como Administrador. Ajustes de usuário (HKCU/.wslconfig) serão aplicados; ações de sistema/HKLM podem ser ignoradas pelo Windows.");
        }

        // [1/8] Serviços de Telemetria
        if (options.DisableTelemetryServices)
        {
            Log("[1/8] Desativando serviços de telemetria e diagnóstico...");
            DisableServicesNative(Log);
        }

        // [2/8] Registro: Telemetria & Sugestões
        if (options.ApplyRegistryTelemetryTweaks)
        {
            Log("[2/8] Aplicando políticas de telemetria, Cortana e anúncios no Registro...");
            ApplyRegistryTelemetryTweaks(isAdmin, Log);
        }

        // [3/8] Remover Bloatwares UWP + Provisionados
        if (options.RemoveBloatwareApps)
        {
            Log("[3/8] Removendo pacotes Bloatware UWP e provisionamentos DISM...");
            RemoveBloatwareApps(isAdmin, Log);
        }

        // [4/8] Desativar Hibernação (hiberfil.sys)
        if (options.DisableHibernation)
        {
            Log("[4/8] Desativando hibernação (powercfg -h off)...");
            DisableHibernationNative(isAdmin, Log);
        }

        // [5/8] Ativar TRIM em SSDs
        if (options.EnableSsdTrim)
        {
            Log("[5/8] Garantindo TRIM ativo em SSDs (fsutil behavior set DisableDeleteNotify 0)...");
            EnableTrimNative(isAdmin, Log);
        }

        // [6/9] Desativar Busca Web (Bing) no Menu Iniciar
        if (options.DisableBingSearch)
        {
            Log("[6/9] Desativando Pesquisa Bing na Barra de Tarefas e Menu Iniciar...");
            DisableBingSearchRegistry(Log);
        }

        // [7/9] Acelerar Responsividade de Menus e Timeouts
        if (options.AccelerateUiMenus)
        {
            Log("[7/9] Acelerando tempo de resposta da interface gráfica (MenuShowDelay = 50ms)...");
            AccelerateUiRegistry(Log);
        }

        // [8/9] Configurar limite de memória do WSL 2 (.wslconfig)
        if (options.ConfigureWslLimit)
        {
            Log("[8/9] Configurando arquivo de limites do WSL 2 (.wslconfig)...");
            WriteWslConfigNative(options.WslMemoryLimit, Log);
        }

        // [9/9] Limpeza de Cache do Docker (Build Cache e Objetos Órfãos / sem apagar containers/volumes ativos)
        if (options.CleanDockerBuilderCache)
        {
            Log("[9/9] Executando limpeza segura de caches do Docker (WSL)...");
            CleanDockerCacheNative(Log);
        }

        Log("==========================================================");
        Log("[SUCESSO] Rotina de otimização e debloat concluída!");
    }

    public async Task RollbackTelemetryServicesAsync(Action<string>? onLog = null)
    {
        await Task.Run(() =>
        {
            void Log(string msg)
            {
                onLog?.Invoke(msg);
                AppLoggerService.Info(msg, nameof(SystemPerformanceService));
            }

            Log("=== REVERTENDO SERVIÇOS DE TELEMETRIA PARA PADRÃO WINDOWS ===");
            if (!IsUserAdministrator())
            {
                Log("[ERRO] Reverter serviços do Windows exige privilégios de Administrador.");
                return;
            }

            foreach (var kv in DebloatTargetServices)
            {
                try
                {
                    using var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Service WHERE Name = '{kv.Key}'");
                    foreach (ManagementObject svc in searcher.Get())
                    {
                        using (svc)
                        {
                            svc.InvokeMethod("ChangeStartMode", [kv.Value]);
                            svc.InvokeMethod("StartService", null);
                            Log($"  -> [OK] Serviço {kv.Key} restaurado ({kv.Value}).");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"  -> Aviso ao restaurar {kv.Key}: {ex.Message}");
                }
            }

            Log("[SUCESSO] Rollback de serviços finalizado.");
        });
    }

    private void DisableServicesNative(Action<string> log)
    {
        foreach (string svcName in DebloatTargetServices.Keys)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Service WHERE Name = '{svcName}'");
                bool found = false;
                foreach (ManagementObject svc in searcher.Get())
                {
                    using (svc)
                    {
                        found = true;
                        string state = svc["State"]?.ToString() ?? "";
                        if (state.Equals("Running", StringComparison.OrdinalIgnoreCase))
                        {
                            var stopRes = svc.InvokeMethod("StopService", null);
                            uint code = Convert.ToUInt32(stopRes ?? 1);
                            if (code != 0 && code != 5) // 5 = already stopped / not active
                                log($"  -> Aviso ao parar {svcName} (código WMI {code})");
                        }

                        var modeRes = svc.InvokeMethod("ChangeStartMode", ["Disabled"]);
                        uint modeCode = Convert.ToUInt32(modeRes ?? 1);
                        if (modeCode == 0)
                            log($"  -> [OK] Serviço {svcName} desativado.");
                        else
                            log($"  -> [AVISO] Serviço {svcName} retornou código WMI {modeCode} (requer Admin ou protegido pelo sistema).");
                    }
                }

                if (!found)
                    log($"  -> Serviço {svcName} não instalado neste host.");
            }
            catch (Exception ex)
            {
                log($"  -> [AVISO] Falha ao desativar {svcName}: {ex.Message}");
            }
        }
    }

    private void ApplyRegistryTelemetryTweaks(bool isAdmin, Action<string> log)
    {
        // HKLM tweaks (requer Admin)
        if (isAdmin)
        {
            SetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Policies\Microsoft\Windows\DataCollection", "AllowTelemetry", 0, log);
            SetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Policies\Microsoft\Windows\Windows Search", "AllowCortana", 0, log);
        }
        else
        {
            log("  -> [AVISO] Políticas HKLM (AllowTelemetry / AllowCortana) ignoradas sem permissão de Administrador.");
        }

        // HKCU tweaks (usuário atual)
        const string cdmPath = @"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";
        SetRegistryDword(Registry.CurrentUser, cdmPath, "SystemPaneSuggestionsEnabled", 0, log);
        SetRegistryDword(Registry.CurrentUser, cdmPath, "SubscribedContent-338388Enabled", 0, log);
        SetRegistryDword(Registry.CurrentUser, cdmPath, "SoftLandingEnabled", 0, log);
        log("  -> [OK] Políticas de telemetria e sugestões do usuário aplicadas.");
    }

    private void RemoveBloatwareApps(bool isAdmin, Action<string> log)
    {
        // 1. Remover pacotes instalados para o usuário atual via Registry + PowerShell 5.1 AppX fallback (sanitizado)
        var detected = DetectInstalledBloatApps();
        if (detected.Count == 0)
        {
            log("  -> [OK] Nenhum pacote UWP bloatware ativo para o usuário atual.");
        }
        else
        {
            string patternList = string.Join(", ", BloatRemovalTokens.Select(t => $"'*{t}*'"));
            string psScript = $@"
$ErrorActionPreference = 'SilentlyContinue'
$patterns = @({patternList})
foreach ($pattern in $patterns) {{
    Get-AppxPackage -Name $pattern | ForEach-Object {{
        try {{
            Remove-AppxPackage -Package $_.PackageFullName -ErrorAction Stop
            Write-Output ""[OK] Removido do usuario: $($_.Name)""
        }} catch {{}}
    }}
}}";
            var res = RunSanitizedWindowsPowerShell(psScript, 45000);
            foreach (string line in res.stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
            {
                log($"  -> {line.Trim()}");
            }
        }

        // 2. Remover pacotes provisionados via DISM.exe nativo (requer Admin)
        if (!isAdmin)
        {
            log("  -> [AVISO] Remoção de pacotes provisionados (DISM) requer execução como Administrador.");
            return;
        }

        try
        {
            var dismList = RunCliTool("dism.exe", "/Online /Get-ProvisionedAppxPackages", 25000);
            if (dismList.exitCode == 0 && !string.IsNullOrWhiteSpace(dismList.stdout))
            {
                int removedCount = 0;
                foreach (string rawLine in dismList.stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    string line = rawLine.Trim();
                    if (!line.StartsWith("PackageName", StringComparison.OrdinalIgnoreCase))
                        continue;

                    int colon = line.IndexOf(':');
                    if (colon < 0) continue;

                    string pkgName = line[(colon + 1)..].Trim();
                    if (!MatchesAnyToken(pkgName, BloatRemovalTokens))
                        continue;

                    var rem = RunCliTool("dism.exe", $"/Online /Remove-ProvisionedAppxPackage /PackageName:{pkgName}", 20000);
                    if (rem.exitCode == 0)
                    {
                        removedCount++;
                        log($"  -> [OK] Provisionamento DISM removido: {pkgName}");
                    }
                }

                if (removedCount == 0)
                    log("  -> Nenhum pacote provisionado da lista alvo encontrado no DISM.");
            }
        }
        catch (Exception ex)
        {
            log($"  -> Aviso ao executar DISM: {ex.Message}");
        }
    }

    private void DisableHibernationNative(bool isAdmin, Action<string> log)
    {
        if (!isAdmin)
        {
            log("  -> [AVISO] Desativar hibernação (powercfg -h off) requer execução como Administrador.");
            return;
        }

        try
        {
            var hiberInfo = new FileInfo(@"C:\hiberfil.sys");
            double beforeGb = hiberInfo.Exists ? Math.Round(hiberInfo.Length / (1024.0 * 1024.0 * 1024.0), 2) : 0;

            var res = RunCliTool("powercfg.exe", "-h off", 10000);
            if (res.exitCode == 0)
            {
                log("  -> [OK] Hibernação DESATIVADA com sucesso!");
                if (beforeGb > 0)
                    log($"  -> [OK] Espaço em disco C: liberado: ~{beforeGb:F2} GB.");
            }
            else
            {
                log($"  -> [AVISO] powercfg retornou código {res.exitCode}: {NormalizeCliOutput(res.stderr)}");
            }
        }
        catch (Exception ex)
        {
            log($"  -> [ERRO] Falha ao desativar hibernação: {ex.Message}");
        }
    }

    private void EnableTrimNative(bool isAdmin, Action<string> log)
    {
        if (!isAdmin)
        {
            log("  -> [AVISO] Ativar TRIM via fsutil requer execução como Administrador.");
            return;
        }

        try
        {
            var res = RunCliTool("fsutil.exe", "behavior set DisableDeleteNotify 0", 10000);
            if (res.exitCode == 0)
                log($"  -> [OK] TRIM forçado como ativo: {NormalizeCliOutput(res.stdout)}");
            else
                log($"  -> [AVISO] fsutil retornou código {res.exitCode}: {NormalizeCliOutput(res.stderr)}");
        }
        catch (Exception ex)
        {
            log($"  -> [ERRO] Falha ao ativar TRIM: {ex.Message}");
        }
    }

    private void DisableBingSearchRegistry(Action<string> log)
    {
        SetRegistryDword(Registry.CurrentUser, @"Software\Policies\Microsoft\Windows\Explorer", "DisableSearchBoxSuggestions", 1, log);
        SetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\Search", "BingSearchEnabled", 0, log);
        SetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\Search", "CortanaConsent", 0, log);
        log("  -> [OK] Pesquisa do Menu Iniciar configurada para busca 100% local.");
    }

    private void AccelerateUiRegistry(Action<string> log)
    {
        SetRegistryString(Registry.CurrentUser, @"Control Panel\Desktop", "MenuShowDelay", "50", log);
        SetRegistryString(Registry.CurrentUser, @"Control Panel\Desktop", "WaitToKillAppTimeout", "5000", log);
        SetRegistryString(Registry.CurrentUser, @"Control Panel\Desktop", "HungAppTimeout", "4000", log);
        log("  -> [OK] Atraso de abertura de menus reduzido para 50ms (MenuShowDelay).");
    }

    private void WriteWslConfigNative(string memoryLimit, Action<string> log)
    {
        try
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string wslConfigPath = Path.Combine(userProfile, ".wslconfig");
            string content = BuildWslConfigContent(memoryLimit);
            File.WriteAllText(wslConfigPath, content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            log($"  -> [OK] Arquivo de limites criado em: {wslConfigPath}");
            log($"  -> [OK] WSL 2 configurado ({SanitizeWslMemoryLimit(memoryLimit)} RAM, 4 CPUs, NAT, gradual auto-reclaim).");
        }
        catch (Exception ex)
        {
            log($"  -> [ERRO] Falha ao gravar .wslconfig: {ex.Message}");
        }
    }

    private void CleanDockerCacheNative(Action<string> log)
    {
        try
        {
            // 1. Limpar Build Cache do Docker (camadas intermediárias de build antigas)
            log("  -> Executando 'docker builder prune -f' (preservando imagens ativas, volumes e containers)...");
            var buildPrune = RunDockerCommand("builder prune -f", 45000);
            if (buildPrune.exitCode == 0 && !string.IsNullOrWhiteSpace(buildPrune.stdout))
            {
                foreach (string line in buildPrune.stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    log($"     [Docker Build] {line.Trim()}");
                }
            }
            else if (!string.IsNullOrWhiteSpace(buildPrune.stderr))
            {
                log($"  -> Aviso docker builder prune: {NormalizeCliOutput(buildPrune.stderr)}");
            }

            // 2. Limpar containers parados e redes não utilizadas (docker system prune sem -a para preservar imagens e volumes)
            log("  -> Executando 'docker system prune -f' (limpeza de containers finalizados e redes órfãs)...");
            var sysPrune = RunDockerCommand("system prune -f", 45000);
            if (sysPrune.exitCode == 0 && !string.IsNullOrWhiteSpace(sysPrune.stdout))
            {
                foreach (string line in sysPrune.stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    log($"     [Docker System] {line.Trim()}");
                }
                log("  -> [OK] Cache e resíduos do Docker liberados com sucesso!");
            }
            else if (!string.IsNullOrWhiteSpace(sysPrune.stderr))
            {
                log($"  -> Aviso docker system prune: {NormalizeCliOutput(sysPrune.stderr)}");
            }
        }
        catch (Exception ex)
        {
            log($"  -> [AVISO] Falha ao executar limpeza do Docker: {ex.Message}");
        }
    }

    private static (int exitCode, string stdout, string stderr) RunDockerCommand(string dockerArgs, int timeoutMs)
    {
        // 1. Tenta executar via docker.exe nativo do Windows (Docker Desktop / Docker CLI Windows)
        var winRes = RunCliTool("docker.exe", dockerArgs, timeoutMs);
        if (winRes.exitCode == 0 && !string.IsNullOrWhiteSpace(winRes.stdout))
        {
            return winRes;
        }

        // 2. Se falhar ou não encontrar, executa dentro do WSL 2 nativo (onde os containers de desenvolvimento rodam)
        try
        {
            var wslRes = RunCliTool("wsl.exe", $"docker {dockerArgs}", timeoutMs);
            if (wslRes.exitCode == 0 && !string.IsNullOrWhiteSpace(wslRes.stdout))
            {
                return wslRes;
            }
        }
        catch { }

        return winRes;
    }

    private List<string> DetectInstalledBloatApps()
    {
        var matches = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        const string regPath = @"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages";

        using var key = Registry.CurrentUser.OpenSubKey(regPath);
        if (key == null)
            return [];

        foreach (string subKeyName in key.GetSubKeyNames())
        {
            if (MatchesAnyToken(subKeyName, BloatAuditTokens))
            {
                string friendly = subKeyName.Split('_')[0];
                matches.Add(friendly);
            }
        }

        return matches.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private List<StartupAppItem> ReadStartupEntries()
    {
        var items = new List<StartupAppItem>();
        const string subPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

        ReadRunKey(Registry.CurrentUser, "HKCU", subPath, items);
        ReadRunKey(Registry.LocalMachine, "HKLM", subPath, items);
        return items;
    }

    private static void ReadRunKey(RegistryKey root, string originLabel, string subPath, List<StartupAppItem> items)
    {
        try
        {
            using var key = root.OpenSubKey(subPath);
            if (key == null) return;

            foreach (string valueName in key.GetValueNames())
            {
                if (string.IsNullOrEmpty(valueName)) continue;
                string val = key.GetValue(valueName)?.ToString() ?? "";
                items.Add(new StartupAppItem(originLabel, valueName, val));
            }
        }
        catch { }
    }

    private List<VhdxFileItem> FindVhdxFiles()
    {
        var results = new List<VhdxFileItem>();
        var candidateRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            @"D:\WSL",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Docker\wsl"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".wsl")
        };

        var seenFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string root in candidateRoots)
        {
            if (!Directory.Exists(root))
                continue;

            try
            {
                var opts = new System.IO.EnumerationOptions
                {
                    RecurseSubdirectories = true,
                    IgnoreInaccessible = true,
                    AttributesToSkip = FileAttributes.ReparsePoint
                };

                foreach (string file in Directory.EnumerateFiles(root, "*.vhdx", opts))
                {
                    if (!seenFiles.Add(file))
                        continue;

                    try
                    {
                        var fi = new FileInfo(file);
                        double gb = Math.Round(fi.Length / (1024.0 * 1024.0 * 1024.0), 2);
                        results.Add(new VhdxFileItem(file, gb));
                    }
                    catch { }
                }
            }
            catch { }
        }

        return results.OrderByDescending(v => v.SizeGb).ToList();
    }

    private static long GetDirectorySizeSafe(string dirPath)
    {
        long total = 0;
        try
        {
            var opts = new System.IO.EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible = true,
                AttributesToSkip = FileAttributes.ReparsePoint
            };

            foreach (string file in Directory.EnumerateFiles(dirPath, "*", opts))
            {
                try { total += new FileInfo(file).Length; } catch { }
            }
        }
        catch { }

        return total;
    }

    private static void SetRegistryDword(RegistryKey root, string subPath, string name, int value, Action<string> log)
    {
        try
        {
            using var key = root.CreateSubKey(subPath, writable: true);
            key?.SetValue(name, value, RegistryValueKind.DWord);
        }
        catch (Exception ex)
        {
            log($"  -> [AVISO] Falha ao gravar {subPath}\\{name}: {ex.Message}");
        }
    }

    private static void SetRegistryString(RegistryKey root, string subPath, string name, string value, Action<string> log)
    {
        try
        {
            using var key = root.CreateSubKey(subPath, writable: true);
            key?.SetValue(name, value, RegistryValueKind.String);
        }
        catch (Exception ex)
        {
            log($"  -> [AVISO] Falha ao gravar {subPath}\\{name}: {ex.Message}");
        }
    }

    private static (int exitCode, string stdout, string stderr) RunCliTool(
        string fileName,
        string arguments,
        int timeoutMs,
        Encoding? outputEncoding = null)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            if (outputEncoding != null)
            {
                psi.StandardOutputEncoding = outputEncoding;
                psi.StandardErrorEncoding = outputEncoding;
            }

            using var proc = new Process { StartInfo = psi };
            proc.Start();

            var stdoutTask = proc.StandardOutput.ReadToEndAsync();
            var stderrTask = proc.StandardError.ReadToEndAsync();

            if (!proc.WaitForExit(timeoutMs))
            {
                try { proc.Kill(entireProcessTree: true); } catch { }
                return (-1, string.Empty, "Tempo limite excedido.");
            }

            Task.WaitAll([stdoutTask, stderrTask], 1500);
            return (proc.ExitCode, stdoutTask.Result, stderrTask.Result);
        }
        catch (Exception ex)
        {
            return (-1, string.Empty, ex.Message);
        }
    }

    private static (int exitCode, string stdout, string stderr) RunSanitizedWindowsPowerShell(string scriptText, int timeoutMs)
    {
        try
        {
            string winPs = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            if (!File.Exists(winPs)) winPs = "powershell.exe";

            string encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(scriptText));
            var psi = new ProcessStartInfo
            {
                FileName = winPs,
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            // Higieniza PSModulePath herdado do PowerShell 7 (pwsh) para evitar conflito de binários no PS 5.1
            string sysModules = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\Modules");
            psi.EnvironmentVariables["PSModulePath"] = sysModules;

            using var proc = new Process { StartInfo = psi };
            proc.Start();
            var stdoutTask = proc.StandardOutput.ReadToEndAsync();
            var stderrTask = proc.StandardError.ReadToEndAsync();

            if (!proc.WaitForExit(timeoutMs))
            {
                try { proc.Kill(entireProcessTree: true); } catch { }
                return (-1, string.Empty, "Timeout ao executar remoção Appx.");
            }

            Task.WaitAll([stdoutTask, stderrTask], 1500);
            return (proc.ExitCode, stdoutTask.Result, stderrTask.Result);
        }
        catch (Exception ex)
        {
            return (-1, string.Empty, ex.Message);
        }
    }
}

