using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

public class CleanupService
{
    // Mapeamento idêntico ao da biblioteca.ps1 (Linhas 84-103)
    private static readonly HashSet<string> AllTargetProcesses = new(StringComparer.OrdinalIgnoreCase)
    {
        "msedge", "msedgewebview2", "msedgedriver", "firefox", "chrome",
        "msteams", "Teams", "ms-teams", "mmc", "Microsoft.ConfigurationManagement",
        "sajapp", "olk", "FoxitPDFEditor", "Notepad", "Time",
        "RemoteDesktopManager", "WindowsTerminal"
    };

    public (int closedCount, List<string> logs) CleanEnvironment(CleanupOptions? options = null, Action<string>? onLog = null)
    {
        options ??= new CleanupOptions();
        var logs = new List<string>();

        void Log(string msg)
        {
            logs.Add(msg);
            onLog?.Invoke(msg);
            AppLoggerService.Info(msg, nameof(CleanupService));
        }

        Log("Iniciando rotina de fechamento prévio e limpeza de ambiente...");

        // 1. Salvar e Fechar Office se selecionado
        if (options.CloseOffice)
        {
            CloseOfficeSafelyNative(Log);
        }

        // 2. Monta lista seletiva de processos baseada na biblioteca.ps1
        var targets = new HashSet<string>(AllTargetProcesses, StringComparer.OrdinalIgnoreCase);

        // Se o usuário não quiser fechar Firefox nessa rodada, remove da mira
        if (!options.CloseFirefox)
        {
            targets.Remove("firefox");
        }

        if (!options.CloseBrowsers)
        {
            targets.Remove("msedge");
            targets.Remove("msedgewebview2");
            targets.Remove("msedgedriver");
            targets.Remove("chrome");
        }

        if (!options.CloseTeams)
        {
            targets.Remove("msteams");
            targets.Remove("Teams");
            targets.Remove("ms-teams");
        }

        if (!options.CloseAdministrativeTools)
        {
            targets.Remove("mmc");
            targets.Remove("Microsoft.ConfigurationManagement");
            targets.Remove("sajapp");
            targets.Remove("olk");
            targets.Remove("RemoteDesktopManager");
        }

        if (!options.CloseUtilities)
        {
            targets.Remove("FoxitPDFEditor");
            targets.Remove("Notepad");
            targets.Remove("Time");
            targets.Remove("WindowsTerminal");
        }

        int closed = CloseProcesses(targets, Log);

        // 3. Telemetria
        if (options.StopTelemetry)
        {
            StopTelemetryNative(Log);
        }

        // 4. Reinício opcional do Explorer (Modo Nuclear nativo C#)
        if (options.RestartExplorer)
        {
            RestartExplorerNative(Log);
        }

        // 5. Garbage Collection
        GC.Collect();
        GC.WaitForPendingFinalizers();

        Log($"Fechamento concluído: {closed} processos finalizados.");
        return (closed, logs);
    }

    private int CloseProcesses(HashSet<string> targets, Action<string> log)
    {
        int myPid = Environment.ProcessId;
        int parentPid = GetParentProcessId(myPid);
        int closedCount = 0;

        foreach (var proc in Process.GetProcesses())
        {
            try
            {
                if (proc.Id == myPid || proc.Id == parentPid)
                    continue;

                string pName = proc.ProcessName;
                if (!targets.Contains(pName))
                    continue;

                string title = proc.MainWindowTitle;
                if (!string.IsNullOrEmpty(title) && title.Contains("🛠️ Modo", StringComparison.OrdinalIgnoreCase))
                {
                    log($"Escudo ativo: mantendo {pName} (Título: {title})");
                    continue;
                }

                proc.Kill();
                closedCount++;
            }
            catch { }
            finally
            {
                proc.Dispose();
            }
        }

        return closedCount;
    }

    private void CloseOfficeSafelyNative(Action<string> log)
    {
        log("Inspecionando instâncias ativas do Excel e Word via COM...");

        // Excel
        if (Process.GetProcessesByName("excel").Length > 0)
        {
            try
            {
                dynamic? excel = ComHelper.GetActiveObject("Excel.Application");
                if (excel != null)
                {
                    excel.DisplayAlerts = false;
                    foreach (dynamic wb in excel.Workbooks)
                    {
                        string path = wb.Path;
                        if (string.IsNullOrEmpty(path))
                        {
                            string downloads = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                            string fileName = $"Planilha_Salva_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                            wb.SaveAs(Path.Combine(downloads, fileName));
                            log($"Planilha nova salva em Downloads: {fileName}");
                        }
                        else
                        {
                            wb.Save();
                            log($"Alterações salvas na planilha: {wb.Name}");
                        }
                    }
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    log("Excel salvo e finalizado.");
                }
            }
            catch (Exception ex)
            {
                log($"Nota no Excel: {ex.Message}");
            }
        }

        // Word
        if (Process.GetProcessesByName("winword").Length > 0)
        {
            try
            {
                dynamic? word = ComHelper.GetActiveObject("Word.Application");
                if (word != null)
                {
                    word.DisplayAlerts = 0;
                    foreach (dynamic doc in word.Documents)
                    {
                        string path = doc.Path;
                        if (string.IsNullOrEmpty(path))
                        {
                            string downloads = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                            string fileName = $"Documento_Salvo_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                            doc.SaveAs2(Path.Combine(downloads, fileName), 16);
                            log($"Documento novo salvo em Downloads: {fileName}");
                        }
                        else
                        {
                            doc.Save();
                            log($"Alterações salvas no documento: {doc.Name}");
                        }
                    }
                    word.Quit();
                    Marshal.ReleaseComObject(word);
                    log("Word salvo e finalizado.");
                }
            }
            catch (Exception ex)
            {
                log($"Nota no Word: {ex.Message}");
            }
        }
    }

    private void StopTelemetryNative(Action<string> log)
    {
        try
        {
            foreach (var p in Process.GetProcessesByName("CompatTelRunner"))
            {
                p.Kill();
                p.Dispose();
            }
            log("Telemetria CompatTelRunner finalizada.");
        }
        catch { }
    }

    private void RestartExplorerNative(Action<string> log)
    {
        try
        {
            log("Reiniciando o Explorador de Arquivos (Modo Nuclear nativo C#)...");
            foreach (var p in Process.GetProcessesByName("explorer"))
            {
                try { p.Kill(); } catch { }
                finally { p.Dispose(); }
            }
            Thread.Sleep(2500);

            if (Process.GetProcessesByName("explorer").Length == 0)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    UseShellExecute = true
                });
            }
            log("Explorador de Arquivos reiniciado e memória do shell reciclada com sucesso.");
        }
        catch (Exception ex)
        {
            log($"Aviso ao reiniciar Explorer: {ex.Message}");
        }
    }

    private int GetParentProcessId(int pid)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher($"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {pid}");
            foreach (var item in searcher.Get())
            {
                return Convert.ToInt32(item["ParentProcessId"]);
            }
        }
        catch { }
        return 0;
    }
}
