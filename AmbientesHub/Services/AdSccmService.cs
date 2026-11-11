using System.DirectoryServices;
using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Services;

public class AdSccmService
{
    private static readonly string[] SccmModuleCandidates =
    [
        @"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1",
        @"C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1",
        @"C:\Program Files (x86)\ConfigMgr Console\bin\ConfigurationManager.psd1",
        @"C:\Program Files (x86)\ConfigMgr Console\bin\ConfigurationManager\ConfigurationManager.psd1",
        @"C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1",
        @"C:\Program Files\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
    ];

    public static string? FindSccmModulePath()
    {
        string? fromEnv = Environment.GetEnvironmentVariable("SMS_ADMIN_UI_PATH");
        if (!string.IsNullOrEmpty(fromEnv))
        {
            string parent = Path.GetDirectoryName(fromEnv) ?? fromEnv;
            string test1 = Path.Combine(parent, "ConfigurationManager.psd1");
            if (File.Exists(test1)) return test1;
            string test2 = Path.Combine(parent, @"ConfigurationManager\ConfigurationManager.psd1");
            if (File.Exists(test2)) return test2;
        }

        foreach (var p in SccmModuleCandidates)
        {
            if (File.Exists(p)) return p;
        }

        return null;
    }

    public void QueryComputer(string computerName, Action<string> onLog)
    {
        Execute(computerName, isDelete: false, onLog);
    }

    public void CleanComputer(string computerName, Action<string> onLog)
    {
        Execute(computerName, isDelete: true, onLog);
    }

    private void Execute(string computerName, bool isDelete, Action<string> onLog)
    {
        if (string.IsNullOrWhiteSpace(computerName))
        {
            onLog("[ERRO] Informe o nome do computador.");
            return;
        }

        string cleanName = computerName.Trim();
        string action = isDelete ? "EXCLUSÃO DIRETA" : "CONSULTA (-WhatIf)";
        onLog($"=== INICIANDO {action}: {cleanName} ===");
        AppLoggerService.Info($"Iniciando {action} para {cleanName}", nameof(AdSccmService));

        // 1. ACTIVE DIRECTORY VIA LDAP NATIVO
        onLog("[1/2] Consultando Active Directory (LDAP)...");
        DirectoryEntry? adEntry = null;
        try
        {
            using var searcher = new DirectorySearcher();
            searcher.Filter = $"(&(objectCategory=computer)(|(cn={cleanName})(sAMAccountName={cleanName}$)))";
            searcher.PropertiesToLoad.AddRange(["cn", "distinguishedName", "operatingSystem", "whenCreated"]);

            SearchResult? result = searcher.FindOne();
            if (result != null)
            {
                adEntry = result.GetDirectoryEntry();
                string dn = adEntry.Properties["distinguishedName"].Value?.ToString() ?? "";
                string os = adEntry.Properties["operatingSystem"].Value?.ToString() ?? "Não informado";
                string created = adEntry.Properties["whenCreated"].Value?.ToString() ?? "";

                onLog("  [OK] Registro encontrado no Active Directory!");
                onLog($"       Nome      : {adEntry.Properties["cn"].Value}");
                onLog($"       DN        : {dn}");
                onLog($"       SO        : {os}");
                onLog($"       Criado em : {created}");

                if (isDelete)
                {
                    onLog("  -> Removendo conta do Active Directory...");
                    adEntry.DeleteTree();
                    adEntry.CommitChanges();
                    onLog("  -> [SUCESSO] Computador removido do Active Directory!");
                    AppLoggerService.Success($"Conta {cleanName} removida do AD.", nameof(AdSccmService));
                }
            }
            else
            {
                onLog($"  [i] Computador '{cleanName}' não localizado no Active Directory.");
            }
        }
        catch (Exception ex)
        {
            onLog($"  -> Falha no Active Directory: {ex.Message}");
            AppLoggerService.Warning($"Erro no AD para {cleanName}: {ex.Message}", nameof(AdSccmService));
        }
        finally
        {
            adEntry?.Dispose();
        }

        // 2. CONFIGURATION MANAGER (SCCM) VIA PROVIDER OFICIAL (CMSite)
        onLog("[2/2] Consultando Configuration Manager (SCCM)...");
        string siteServer = ConfigService.SccmSiteServer;
        string siteCode = ConfigService.SccmSiteCode;

        if (string.IsNullOrWhiteSpace(siteServer) || string.IsNullOrWhiteSpace(siteCode))
        {
            onLog("  [Aviso] SCCM_SITE_SERVER ou SCCM_SITE_CODE não configurados na aba Configurações.");
            return;
        }

        string? modulePath = FindSccmModulePath();

        if (modulePath == null)
        {
            onLog("  [Aviso] Módulo 'ConfigurationManager.psd1' não encontrado na estação.");
            onLog("          O Console do Configuration Manager precisa estar instalado.");
            return;
        }

        try
        {
            string scriptTemp = Path.Combine(Path.GetTempPath(), $"sccm_action_{Guid.NewGuid():N}.ps1");

            string psCode = $@"
$ErrorActionPreference = 'Stop'
Import-Module '{modulePath}'
if (-not (Get-PSDrive -Name '{siteCode}' -ErrorAction SilentlyContinue)) {{
    New-PSDrive -Name '{siteCode}' -PSProvider 'CMSite' -Root '{siteServer}' -Description 'SCCM Drive' | Out-Null
}}
Set-Location '{siteCode}:'
$devices = @(Get-CMDevice -Name '{cleanName}' -ErrorAction SilentlyContinue)
if ($devices.Count -gt 0) {{
    Write-Output ""[OK] $($devices.Count) registro(s) encontrado(s) no SCCM!""
    foreach ($dev in $devices) {{
        Write-Output ""       ResourceID  : $($dev.ResourceID)""
        Write-Output ""       SMSID       : $($dev.SMSID)""
        Write-Output ""       Cliente Atv : $($dev.ClientActiveStatus)""
        Write-Output ""       SO          : $($dev.DeviceOS)""
        {(isDelete ? @"
        Write-Output ""  -> Excluindo ResourceID $($dev.ResourceID) do SCCM...""
        Remove-CMResource -ResourceId $dev.ResourceID -Force
        Write-Output ""  -> [SUCESSO] Registro removido do Configuration Manager!""
        " : "")}
    }}
}} else {{
    Write-Output ""[i] Computador '{cleanName}' não localizado no SCCM.""
}}
";
            File.WriteAllText(scriptTemp, psCode, System.Text.Encoding.UTF8);

            string winPs = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
            if (!File.Exists(winPs)) winPs = "powershell.exe";

            var startInfo = new ProcessStartInfo
            {
                FileName = winPs,
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptTemp}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var proc = new Process { StartInfo = startInfo };
            proc.OutputDataReceived += (_, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    onLog($"  {e.Data}");
            };
            proc.ErrorDataReceived += (_, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    onLog($"  -> Erro SCCM: {e.Data}");
            };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            proc.WaitForExit();

            try { File.Delete(scriptTemp); } catch { }

            if (proc.ExitCode == 0)
            {
                AppLoggerService.Success($"Operação SCCM para {cleanName} finalizada com sucesso.", nameof(AdSccmService));
            }
        }
        catch (Exception ex)
        {
            onLog($"  -> Falha ao comunicar com o SCCM: {ex.Message}");
            AppLoggerService.Danger($"Erro SCCM para {cleanName}: {ex.Message}", nameof(AdSccmService));
        }

        onLog($"=== OPERAÇÃO FINALIZADA PARA: {cleanName} ===");
    }
}
