using System.IO;
using System.Management;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

public class RemoteMaintenanceService
{
    public (double FreeGb, double TotalGb) GetLogicalDiskCStats(string computerName)
    {
        try
        {
            var scope = RemoteManagementHelper.CreateScope(computerName, @"root\cimv2");
            using var searcher = new ManagementObjectSearcher(
                scope,
                new ObjectQuery("SELECT FreeSpace, Size FROM Win32_LogicalDisk WHERE DeviceID='C:'"));

            foreach (ManagementObject disk in searcher.Get())
            {
                using (disk)
                {
                    double free = disk["FreeSpace"] != null
                        ? Math.Round(Convert.ToDouble(disk["FreeSpace"]) / (1024.0 * 1024.0 * 1024.0), 2)
                        : 0;
                    double total = disk["Size"] != null
                        ? Math.Round(Convert.ToDouble(disk["Size"]) / (1024.0 * 1024.0 * 1024.0), 2)
                        : 0;
                    return (free, total);
                }
            }
        }
        catch { }

        return (0, 0);
    }

    public RemoteMaintenanceResult AnalyzeRemoteCleanup(
        string computerName,
        RemoteMaintenanceOptions options,
        Action<string>? onLog = null)
    {
        string target = RemoteManagementHelper.NormalizeTargetHost(computerName);
        var result = new RemoteMaintenanceResult { ComputerName = target };

        void Log(string msg) => onLog?.Invoke(msg);

        Log($"=== ANALISANDO ESPAÇO E CACHES EM: {target} (MODO DIAGNÓSTICO) ===");

        if (!RemoteManagementHelper.PingHost(target))
        {
            Log($"[ERRO] Máquina '{target}' inacessível via ICMP.");
            return result;
        }

        try
        {
            RemoteManagementHelper.EnsureAdminShareConnected(target);
            var (freeGb, totalGb) = GetLogicalDiskCStats(target);
            result.FreeSpaceBeforeGb = freeGb;
            result.FreeSpaceAfterGb = freeGb;
            result.TotalSpaceGb = totalGb;

            Log($"[DISCO C:] Livre: {freeGb:N2} GB de {totalGb:N2} GB totais");
            Log("------------------------------------------------------------------");

            string rootC = RemoteManagementHelper.ConvertLocalPathToTargetPath(target, @"C:\");
            string usersDir = Path.Combine(rootC, "Users");

            // 1. Temp de Usuários
            if (options.CleanTempFiles)
            {
                long tempBytes = 0;
                int tempFiles = 0;
                if (Directory.Exists(usersDir))
                {
                    foreach (var userFolder in SafeGetDirectories(usersDir))
                    {
                        var st = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(userFolder, @"AppData\Local\Temp"));
                        tempBytes += st.Bytes;
                        tempFiles += st.FileCount;
                    }
                }
                AddInspection(result, "Temp de Usuários (AppData\\Local\\Temp)", @"C:\Users\*\AppData\Local\Temp", tempBytes, tempFiles, Log);

                var sysTemp = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(rootC, @"Windows\Temp"));
                AddInspection(result, "Temp do Sistema (Windows\\Temp)", @"C:\Windows\Temp", sysTemp.Bytes, sysTemp.FileCount, Log);
            }

            // 2. Prefetch
            if (options.CleanPrefetch)
            {
                var pf = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(rootC, @"Windows\Prefetch"));
                AddInspection(result, "Cache Prefetch", @"C:\Windows\Prefetch", pf.Bytes, pf.FileCount, Log);
            }

            // 3. Lixeira
            if (options.CleanRecycleBin)
            {
                var rb = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(rootC, "$Recycle.Bin"));
                AddInspection(result, "Lixeira Global ($Recycle.Bin)", @"C:\$Recycle.Bin", rb.Bytes, rb.FileCount, Log);
            }

            // 4. Windows Update Cache
            if (options.CleanWindowsUpdateCache)
            {
                var wu = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(rootC, @"Windows\SoftwareDistribution\Download"));
                AddInspection(result, "Windows Update (SoftwareDistribution)", @"C:\Windows\SoftwareDistribution\Download", wu.Bytes, wu.FileCount, Log);
            }

            // 5. Logs CBS
            if (options.CleanCbsLogs)
            {
                var cbs = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(rootC, @"Windows\Logs\CBS"));
                AddInspection(result, "Logs CBS (Component-Based Servicing)", @"C:\Windows\Logs\CBS", cbs.Bytes, cbs.FileCount, Log);
            }

            // 6. Thumbnail Cache
            if (options.CleanThumbnailCache)
            {
                long thumbBytes = 0;
                int thumbFiles = 0;
                if (Directory.Exists(usersDir))
                {
                    foreach (var userFolder in SafeGetDirectories(usersDir))
                    {
                        string expDir = Path.Combine(userFolder, @"AppData\Local\Microsoft\Windows\Explorer");
                        if (Directory.Exists(expDir))
                        {
                            foreach (var f in SafeGetFiles(expDir, "thumbcache_*.db"))
                            {
                                try { thumbBytes += new FileInfo(f).Length; thumbFiles++; } catch { }
                            }
                        }
                    }
                }
                AddInspection(result, "Cache de Miniaturas (thumbcache_*.db)", @"C:\Users\*\Explorer\thumbcache_*.db", thumbBytes, thumbFiles, Log);
            }

            // 7. Delivery Optimization
            if (options.CleanDeliveryOptimization)
            {
                var doSt = RemoteManagementHelper.CalculateDirectoryStats(
                    Path.Combine(rootC, @"Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"));
                AddInspection(result, "Delivery Optimization Cache", @"C:\Windows\...\DeliveryOptimization\Cache", doSt.Bytes, doSt.FileCount, Log);
            }

            // 8. CrashDumps & WER
            if (options.CleanCrashDumps)
            {
                long dumpBytes = 0;
                int dumpFiles = 0;

                string[] dumpDirs =
                [
                    Path.Combine(rootC, @"ProgramData\Microsoft\Windows\WER\ReportQueue"),
                    Path.Combine(rootC, @"ProgramData\Microsoft\Windows\WER\ReportArchive"),
                    Path.Combine(rootC, @"Windows\Minidump")
                ];

                foreach (var d in dumpDirs)
                {
                    var st = RemoteManagementHelper.CalculateDirectoryStats(d);
                    dumpBytes += st.Bytes;
                    dumpFiles += st.FileCount;
                }

                string memDmp = Path.Combine(rootC, @"Windows\MEMORY.DMP");
                if (File.Exists(memDmp))
                {
                    try { dumpBytes += new FileInfo(memDmp).Length; dumpFiles++; } catch { }
                }

                if (Directory.Exists(usersDir))
                {
                    foreach (var u in SafeGetDirectories(usersDir))
                    {
                        var st = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(u, @"AppData\Local\CrashDumps"));
                        dumpBytes += st.Bytes;
                        dumpFiles += st.FileCount;
                    }
                }

                AddInspection(result, "Relatórios de Erro (WER) & CrashDumps", @"C:\ProgramData\...\WER & Minidump", dumpBytes, dumpFiles, Log);
            }

            // 9. Windows.old
            if (options.CleanWindowsOld)
            {
                long oldBytes = 0;
                int oldFiles = 0;
                foreach (string folder in new[] { "Windows.old", "$Windows.~BT", "$Windows.~WS" })
                {
                    var st = RemoteManagementHelper.CalculateDirectoryStats(Path.Combine(rootC, folder));
                    oldBytes += st.Bytes;
                    oldFiles += st.FileCount;
                }
                AddInspection(result, "Instalações Antigas (Windows.old / ~BT)", @"C:\Windows.old", oldBytes, oldFiles, Log);
            }

            double totalCleanableMb = result.InspectionItems.Sum(i => i.SizeMb);
            result.EstimatedCleanableGb = Math.Round(totalCleanableMb / 1024.0, 2);
            result.Success = true;

            string totalFormatted = totalCleanableMb >= 1024.0
                ? $"{result.EstimatedCleanableGb:N2} GB ({totalCleanableMb:N1} MB)"
                : $"{totalCleanableMb:N1} MB";

            Log("------------------------------------------------------------------");
            Log($"[DIAGNÓSTICO CONCLUÍDO] Potencial estimado de limpeza direta em arquivos: ~{totalFormatted}");
            Log("Clique em 'Executar Limpeza & Manutenção' para processar os itens selecionados.");
        }
        catch (Exception ex)
        {
            Log($"[ERRO] Falha durante a análise em '{target}': {ex.Message}");
        }

        return result;
    }

    public RemoteMaintenanceResult ExecuteMaintenance(
        string computerName,
        RemoteMaintenanceOptions options,
        Action<string>? onLog = null)
    {
        string target = RemoteManagementHelper.NormalizeTargetHost(computerName);
        var result = new RemoteMaintenanceResult { ComputerName = target };

        void Log(string msg) => onLog?.Invoke(msg);

        Log($"=== INICIANDO MANUTENÇÃO & LIMPEZA EM: {target} ===");

        if (!RemoteManagementHelper.PingHost(target))
        {
            Log($"[ERRO] Máquina '{target}' inacessível via ICMP. Abortando.");
            return result;
        }

        try
        {
            RemoteManagementHelper.EnableWinRmAndFirewallViaDcom(target, onLog);
            RemoteManagementHelper.EnsureAdminShareConnected(target);

            var (beforeGb, totalGb) = GetLogicalDiskCStats(target);
            result.FreeSpaceBeforeGb = beforeGb;
            result.TotalSpaceGb = totalGb;
            Log($"[ANTES] Espaço livre em C: {beforeGb:N2} GB (Total: {totalGb:N2} GB)");

            string rootC = RemoteManagementHelper.ConvertLocalPathToTargetPath(target, @"C:\");
            string usersDir = Path.Combine(rootC, "Users");

            // 1. Temp Files
            if (options.CleanTempFiles)
            {
                Log("-> [1] Limpando arquivos temporários de usuários e C:\\Windows\\Temp...");
                int deleted = 0;
                if (Directory.Exists(usersDir))
                {
                    foreach (var u in SafeGetDirectories(usersDir))
                    {
                        deleted += CleanDirectoryContentsSafe(Path.Combine(u, @"AppData\Local\Temp"));
                    }
                }
                deleted += CleanDirectoryContentsSafe(Path.Combine(rootC, @"Windows\Temp"));
                Log($"   • {deleted} arquivo(s)/pasta(s) temporários removidos.");
            }

            // 2. Prefetch
            if (options.CleanPrefetch)
            {
                Log("-> [2] Limpando C:\\Windows\\Prefetch...");
                int deleted = CleanDirectoryContentsSafe(Path.Combine(rootC, @"Windows\Prefetch"));
                Log($"   • {deleted} item(ns) removidos de Prefetch.");
            }

            // 3. Lixeira
            if (options.CleanRecycleBin)
            {
                Log("-> [3] Esvaziando Lixeira (C:\\$Recycle.Bin)...");
                CleanDirectoryContentsSafe(Path.Combine(rootC, "$Recycle.Bin"));
                RemoteManagementHelper.InvokeCommandOnTarget(target, @"rd /s /q C:\$Recycle.Bin", 30);
                Log("   • Lixeira esvaziada.");
            }

            // 4. Event Logs (Admin / Operational)
            if (options.CleanEventLogs)
            {
                Log("-> [4] Limpando logs de eventos Admin e Operational (wevtutil)...");
                RemoteManagementHelper.InvokeCommandOnTarget(
                    target,
                    @"powershell.exe -NoProfile -NonInteractive -Command ""wevtutil el | Where-Object { $_ -match 'Admin|Operational' } | ForEach-Object { wevtutil cl $_ 2>$null }""",
                    60);
                Log("   • Logs Admin/Operational limpos.");
            }

            // 5. Windows Update Cache
            if (options.CleanWindowsUpdateCache)
            {
                Log("-> [5] Limpando cache do Windows Update (SoftwareDistribution\\Download)...");
                ControlServiceViaWmi(target, "wuauserv", start: false);
                int deleted = CleanDirectoryContentsSafe(Path.Combine(rootC, @"Windows\SoftwareDistribution\Download"));
                ControlServiceViaWmi(target, "wuauserv", start: true);
                Log($"   • Cache do Windows Update limpo ({deleted} itens removidos) e serviço wuauserv reiniciado.");
            }

            // 6. DNS Cache
            if (options.FlushDnsCache)
            {
                Log("-> [6] Limpando cache DNS (ipconfig /flushdns)...");
                RemoteManagementHelper.InvokeCommandOnTarget(target, "ipconfig /flushdns", 20);
                Log("   • Cache DNS limpo.");
            }

            // 7. WinSxS (DISM StartComponentCleanup)
            if (options.CleanWinSxS)
            {
                Log("-> [7] Executando limpeza de componentes WinSxS (DISM /StartComponentCleanup)...");
                RemoteManagementHelper.InvokeCommandOnTarget(target, "dism.exe /Online /Cleanup-Image /StartComponentCleanup /Quiet", 180);
                Log("   • DISM StartComponentCleanup concluído.");
            }

            // 8. Thumbnail Cache
            if (options.CleanThumbnailCache)
            {
                Log("-> [8] Limpando cache de miniaturas (thumbcache_*.db)...");
                int count = 0;
                if (Directory.Exists(usersDir))
                {
                    foreach (var u in SafeGetDirectories(usersDir))
                    {
                        string exp = Path.Combine(u, @"AppData\Local\Microsoft\Windows\Explorer");
                        foreach (var f in SafeGetFiles(exp, "thumbcache_*.db"))
                        {
                            try { File.Delete(f); count++; } catch { }
                        }
                    }
                }
                Log($"   • {count} arquivo(s) thumbcache removidos.");
            }

            // 9. Store Cache
            if (options.CleanStoreCache)
            {
                Log("-> [9] Limpando cache da Microsoft Store...");
                if (Directory.Exists(usersDir))
                {
                    foreach (var u in SafeGetDirectories(usersDir))
                    {
                        string pkgDir = Path.Combine(u, @"AppData\Local\Packages");
                        if (Directory.Exists(pkgDir))
                        {
                            foreach (var storeDir in SafeGetDirectories(pkgDir, "Microsoft.WindowsStore*"))
                            {
                                CleanDirectoryContentsSafe(Path.Combine(storeDir, "LocalCache"));
                            }
                        }
                    }
                }
                Log("   • Cache da Microsoft Store limpo.");
            }

            // 10. CBS Logs
            if (options.CleanCbsLogs)
            {
                Log("-> [10] Truncando/limpando logs em C:\\Windows\\Logs\\CBS...");
                string cbsDir = Path.Combine(rootC, @"Windows\Logs\CBS");
                int count = 0;
                if (Directory.Exists(cbsDir))
                {
                    foreach (var f in SafeGetFiles(cbsDir, "*"))
                    {
                        try
                        {
                            File.WriteAllText(f, string.Empty);
                            count++;
                        }
                        catch
                        {
                            try { File.Delete(f); count++; } catch { }
                        }
                    }
                }
                Log($"   • {count} arquivo(s) de log CBS processados.");
            }

            // 11. Windows.old
            if (options.CleanWindowsOld)
            {
                Log("-> [11] Removendo pastas de instalações antigas (Windows.old / $Windows.~BT / $Windows.~WS)...");
                foreach (string folder in new[] { @"C:\Windows.old", @"C:\$Windows.~BT", @"C:\$Windows.~WS" })
                {
                    string uncFolder = RemoteManagementHelper.ConvertLocalPathToTargetPath(target, folder);
                    if (Directory.Exists(uncFolder))
                    {
                        RemoteManagementHelper.InvokeCommandOnTarget(
                            target,
                            $"takeown /F \"{folder}\" /R /A /D Y >nul 2>&1 & icacls \"{folder}\" /grant *S-1-5-32-544:F /T /C /Q >nul 2>&1 & rmdir /S /Q \"{folder}\"",
                            120);
                        Log($"   • Pasta '{folder}' removida.");
                    }
                }
            }

            // 12. Delivery Optimization
            if (options.CleanDeliveryOptimization)
            {
                Log("-> [12] Limpando cache do Delivery Optimization...");
                ControlServiceViaWmi(target, "DoSvc", start: false);
                CleanDirectoryContentsSafe(Path.Combine(rootC, @"Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"));
                ControlServiceViaWmi(target, "DoSvc", start: true);
                Log("   • Cache Delivery Optimization limpo.");
            }

            // 13. CrashDumps & WER
            if (options.CleanCrashDumps)
            {
                Log("-> [13] Limpando relatórios de erro (WER) e despejos de memória (CrashDumps)...");
                CleanDirectoryContentsSafe(Path.Combine(rootC, @"ProgramData\Microsoft\Windows\WER\ReportQueue"));
                CleanDirectoryContentsSafe(Path.Combine(rootC, @"ProgramData\Microsoft\Windows\WER\ReportArchive"));
                CleanDirectoryContentsSafe(Path.Combine(rootC, @"Windows\Minidump"));

                string memDmp = Path.Combine(rootC, @"Windows\MEMORY.DMP");
                if (File.Exists(memDmp))
                {
                    try { File.Delete(memDmp); } catch { }
                }

                if (Directory.Exists(usersDir))
                {
                    foreach (var u in SafeGetDirectories(usersDir))
                    {
                        CleanDirectoryContentsSafe(Path.Combine(u, @"AppData\Local\CrashDumps"));
                    }
                }
                Log("   • Despejos de memória e filas WER limpos.");
            }

            // 14. Repair-Volume / Chkdsk
            if (options.RunRepairVolume)
            {
                Log("-> [14] Verificando/reparando volume C: (Repair-Volume)...");
                RemoteManagementHelper.InvokeCommandOnTarget(
                    target,
                    @"powershell.exe -NoProfile -NonInteractive -Command ""Repair-Volume -DriveLetter C -Scan -ErrorAction SilentlyContinue""",
                    180);
                Log("   • Verificação Repair-Volume concluída.");
            }

            // 15. Defrag / Optimize-Volume (somente se HDD)
            if (options.RunDefragIfHdd)
            {
                Log("-> [15] Verificando tipo de mídia física (SSD vs HDD) para desfragmentação...");
                bool hasHdd = CheckIfSystemHasHdd(target);
                if (!hasHdd)
                {
                    Log("   • Todos os discos são SSD. Pulando desfragmentação (não recomendado em SSD).");
                }
                else
                {
                    Log("   • HDD detectado. Executando Optimize-Volume -DriveLetter C -Defrag...");
                    RemoteManagementHelper.InvokeCommandOnTarget(
                        target,
                        @"powershell.exe -NoProfile -NonInteractive -Command ""Optimize-Volume -DriveLetter C -Defrag -ErrorAction SilentlyContinue""",
                        300);
                    Log("   • Optimize-Volume (Defrag) concluído.");
                }
            }

            // 16. DISM CheckHealth + SFC /scannow
            if (options.RunSfcIfCorrupted)
            {
                Log("-> [16] Verificando integridade da imagem (DISM CheckHealth + SFC se necessário)...");
                RemoteManagementHelper.InvokeCommandOnTarget(
                    target,
                    @"powershell.exe -NoProfile -NonInteractive -Command ""$h = dism.exe /Online /Cleanup-Image /CheckHealth | Out-String; if ($h -notmatch 'No component store corruption|Nenhuma corrup') { sfc.exe /scannow }""",
                    300);
                Log("   • Verificação de integridade concluída.");
            }

            var (afterGb, _) = GetLogicalDiskCStats(target);
            result.FreeSpaceAfterGb = afterGb;
            result.Success = true;

            Log("==================================================================");
            Log($"[ANTES]      Espaço livre em C: : {result.FreeSpaceBeforeGb:N2} GB");
            Log($"[DEPOIS]     Espaço livre em C: : {result.FreeSpaceAfterGb:N2} GB");
            Log($"[RECUPERADO] Total liberado     : {result.FreedGb:N2} GB");
            Log($"[SUCESSO] Manutenção e limpeza em '{target}' concluídas com êxito!");
        }
        catch (Exception ex)
        {
            Log($"[ERRO] Falha durante a manutenção em '{target}': {ex.Message}");
        }

        return result;
    }

    public static string FormatDataSize(long bytes)
    {
        if (bytes <= 0) return "0,00 MB";
        double mb = bytes / (1024.0 * 1024.0);
        if (mb >= 1024.0)
        {
            double gb = mb / 1024.0;
            return $"{gb:N2} GB";
        }
        return $"{mb:N2} MB";
    }

    private static void AddInspection(
        RemoteMaintenanceResult res,
        string category,
        string displayPath,
        long bytes,
        int files,
        Action<string> log)
    {
        double mb = Math.Round(bytes / (1024.0 * 1024.0), 2);
        res.InspectionItems.Add(new RemoteCleanupInspectionItem
        {
            Category = category,
            TargetPath = displayPath,
            SizeMb = mb,
            FileCount = files,
            Exists = bytes > 0 || files > 0
        });

        string formattedSize = FormatDataSize(bytes);
        log($"  • {category,-44} : {formattedSize,10} ({files} arquivos)");
    }

    private static int CleanDirectoryContentsSafe(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
            return 0;

        int removed = 0;
        foreach (var file in SafeGetFiles(directoryPath, "*"))
        {
            try
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
                removed++;
            }
            catch { }
        }

        foreach (var subDir in SafeGetDirectories(directoryPath))
        {
            try
            {
                // Anti Symlink/Junction Traversal: se for ReparsePoint (junction/symlink), remove apenas o elo sem descer recursivamente no destino.
                var dirInfo = new DirectoryInfo(subDir);
                if ((dirInfo.Attributes & FileAttributes.ReparsePoint) != 0)
                {
                    dirInfo.Delete(recursive: false);
                }
                else
                {
                    Directory.Delete(subDir, recursive: true);
                }
                removed++;
            }
            catch
            {
                removed += CleanDirectoryContentsSafe(subDir);
            }
        }

        return removed;
    }

    private static void ControlServiceViaWmi(string computerName, string serviceName, bool start)
    {
        try
        {
            var scope = RemoteManagementHelper.CreateScope(computerName, @"root\cimv2");
            using var searcher = new ManagementObjectSearcher(
                scope,
                new ObjectQuery($"SELECT * FROM Win32_Service WHERE Name='{serviceName}'"));

            foreach (ManagementObject svc in searcher.Get())
            {
                using (svc)
                {
                    svc.InvokeMethod(start ? "StartService" : "StopService", null);
                }
            }
            Thread.Sleep(800);
        }
        catch { }
    }

    private static bool CheckIfSystemHasHdd(string computerName)
    {
        try
        {
            var storageScope = RemoteManagementHelper.CreateScope(computerName, @"root\Microsoft\Windows\Storage");
            using var searcher = new ManagementObjectSearcher(
                storageScope,
                new ObjectQuery("SELECT MediaType FROM MSFT_PhysicalDisk"));

            foreach (ManagementObject disk in searcher.Get())
            {
                using (disk)
                {
                    if (disk["MediaType"] != null && Convert.ToUInt16(disk["MediaType"]) == 3) // 3 = HDD, 4 = SSD
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    private static IEnumerable<string> SafeGetDirectories(string path, string pattern = "*")
    {
        try
        {
            return Directory.Exists(path) ? Directory.GetDirectories(path, pattern) : [];
        }
        catch
        {
            return [];
        }
    }

    private static IEnumerable<string> SafeGetFiles(string path, string pattern = "*")
    {
        try
        {
            return Directory.Exists(path) ? Directory.GetFiles(path, pattern) : [];
        }
        catch
        {
            return [];
        }
    }
}

