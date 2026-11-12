using System.DirectoryServices;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using AmbientesHub.Models;

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

        // 1. ACTIVE DIRECTORY VIA LDAP NATIVO (Busca Ampla com Wildcard / LIKE)
        onLog("[1/2] Consultando Active Directory (LDAP)...");
        try
        {
            using var searcher = new DirectorySearcher();
            searcher.Filter = $"(&(objectCategory=computer)(|(cn=*{cleanName}*)(sAMAccountName=*{cleanName}*)))";
            searcher.PropertiesToLoad.AddRange(["cn", "distinguishedName", "operatingSystem", "whenCreated"]);
            searcher.SizeLimit = 15;

            var results = searcher.FindAll();
            if (results.Count > 0)
            {
                onLog($"  [OK] {results.Count} registro(s) encontrado(s) no Active Directory!");
                foreach (SearchResult result in results)
                {
                    using var adEntry = result.GetDirectoryEntry();
                    string name = adEntry.Properties["cn"].Value?.ToString() ?? cleanName;
                    string dn = adEntry.Properties["distinguishedName"].Value?.ToString() ?? "";
                    string os = adEntry.Properties["operatingSystem"].Value?.ToString() ?? "Não informado";
                    string created = adEntry.Properties["whenCreated"].Value?.ToString() ?? "";

                    onLog($"       Nome      : {name}");
                    onLog($"       DN        : {dn}");
                    onLog($"       SO        : {os}");
                    onLog($"       Criado em : {created}");

                    if (isDelete)
                    {
                        onLog($"  -> Removendo conta '{name}' do Active Directory...");
                        adEntry.DeleteTree();
                        adEntry.CommitChanges();
                        onLog($"  -> [SUCESSO] Computador '{name}' removido do Active Directory!");
                        AppLoggerService.Success($"Conta {name} removida do AD.", nameof(AdSccmService));
                    }
                }
            }
            else
            {
                onLog($"  [i] Nenhum computador correspondente a '{cleanName}' localizado no Active Directory.");
            }
        }
        catch (Exception ex)
        {
            onLog($"  -> Falha no Active Directory: {ex.Message}");
            AppLoggerService.Warning($"Erro no AD para {cleanName}: {ex.Message}", nameof(AdSccmService));
        }

        // 2. CONFIGURATION MANAGER (SCCM)
        onLog("[2/2] Consultando Configuration Manager (SCCM)...");
        string siteServer = ConfigService.SccmSiteServer;
        string siteCode = ConfigService.SccmSiteCode;

        if (string.IsNullOrWhiteSpace(siteServer) || string.IsNullOrWhiteSpace(siteCode))
        {
            onLog("  [Aviso] SCCM_SITE_SERVER ou SCCM_SITE_CODE não configurados na aba Configurações.");
            return;
        }

        // Tentativa 1: WMI C# Nativo com PacketPrivacy (Zero processos externos)
        bool nativeSuccess = TryExecuteSccmNative(siteServer, siteCode, cleanName, isDelete, onLog);
        if (!nativeSuccess)
        {
            // Tentativa 2: Provedor Oficial CMSite com Ambiente PowerShell Sanitizado
            ExecuteSccmProviderFallback(siteServer, siteCode, cleanName, isDelete, onLog);
        }

        onLog($"=== OPERAÇÃO FINALIZADA PARA: {cleanName} ===");
    }

    private bool TryExecuteSccmNative(string siteServer, string siteCode, string cleanName, bool isDelete, Action<string> onLog)
    {
        try
        {
            var options = new ConnectionOptions
            {
                Impersonation = ImpersonationLevel.Impersonate,
                Authentication = AuthenticationLevel.PacketPrivacy,
                EnablePrivileges = true,
                Timeout = TimeSpan.FromSeconds(5)
            };

            string scopePath = $@"\\{siteServer}\root\sms\site_{siteCode}";
            var scope = new ManagementScope(scopePath, options);
            scope.Connect();

            string query = $"SELECT ResourceID, SMSID, ClientActiveStatus, OperatingSystemNameandVersion, NetbiosName, Name FROM SMS_R_System WHERE NetbiosName LIKE '%{cleanName}%' OR Name LIKE '%{cleanName}%'";
            using var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(query));
            using var results = searcher.Get();

            var list = results.Cast<ManagementObject>().ToList();
            if (list.Count > 0)
            {
                onLog($"  [OK] {list.Count} registro(s) encontrado(s) no SCCM (WMI C# Nativo)!");
                foreach (var dev in list)
                {
                    string name = dev["NetbiosName"]?.ToString() ?? dev["Name"]?.ToString() ?? cleanName;
                    string resId = dev["ResourceID"]?.ToString() ?? "";
                    string smsId = dev["SMSID"]?.ToString() ?? "";
                    string clientActive = dev["ClientActiveStatus"]?.ToString() == "1" ? "Ativo" : "Inativo";
                    string os = dev["OperatingSystemNameandVersion"]?.ToString() ?? "";

                    onLog($"       Nome        : {name}");
                    onLog($"       ResourceID  : {resId}");
                    onLog($"       SMSID       : {smsId}");
                    onLog($"       Cliente Atv : {clientActive}");
                    onLog($"       SO          : {os}");

                    if (isDelete)
                    {
                        onLog($"  -> Excluindo ResourceID {resId} ({name}) via WMI C# Nativo...");
                        dev.Delete();
                        onLog($"  -> [SUCESSO] Registro '{name}' removido do Configuration Manager!");
                        AppLoggerService.Success($"ResourceID {resId} de {name} excluído via WMI Nativo.", nameof(AdSccmService));
                    }
                }
                return true;
            }
            else
            {
                onLog($"  [i] Nenhum computador correspondente a '{cleanName}' localizado no SCCM (WMI C# Nativo).");
                return true;
            }
        }
        catch (Exception ex)
        {
            AppLoggerService.Info($"WMI C# Nativo não estabeleceu conexão direta ({ex.Message}). Recorrendo ao provedor CMSite com ambiente sanitizado.", nameof(AdSccmService));
            return false;
        }
    }

    private void ExecuteSccmProviderFallback(string siteServer, string siteCode, string cleanName, bool isDelete, Action<string> onLog)
    {
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
$devices = @(Get-CMDevice | Where-Object {{ $_.Name -like '*{cleanName}*' }})
if ($devices.Count -gt 0) {{
    Write-Output ""[OK] $($devices.Count) registro(s) encontrado(s) no SCCM via CMSite!""
    foreach ($dev in $devices) {{
        Write-Output ""       Nome        : $($dev.Name)""
        Write-Output ""       ResourceID  : $($dev.ResourceID)""
        Write-Output ""       SMSID       : $($dev.SMSID)""
        Write-Output ""       Cliente Atv : $($dev.ClientActiveStatus)""
        Write-Output ""       SO          : $($dev.DeviceOS)""
        {(isDelete ? @"
        Write-Output ""  -> Excluindo ResourceID $($dev.ResourceID) ($($dev.Name)) do SCCM...""
        Remove-CMResource -ResourceId $dev.ResourceID -Force
        Write-Output ""  -> [SUCESSO] Registro '$($dev.Name)' removido do Configuration Manager!""
        " : "")}
    }}
}} else {{
    Write-Output ""[i] Nenhum computador correspondente a '{cleanName}' localizado no SCCM.""
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

            // Sanitização de ambiente para prevenir contaminação pelo PowerShell 7
            string defaultPsModules = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\Modules");
            startInfo.EnvironmentVariables["PSModulePath"] = defaultPsModules;
            if (startInfo.EnvironmentVariables.ContainsKey("DOTNET_ROOT"))
            {
                startInfo.EnvironmentVariables.Remove("DOTNET_ROOT");
            }

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
            if (!proc.WaitForExit(35000))
            {
                try { proc.Kill(); } catch { }
                onLog("  [Aviso] Tempo limite de 35s excedido na consulta SCCM CMSite.");
            }

            try { File.Delete(scriptTemp); } catch { }

            if (proc.ExitCode == 0)
            {
                AppLoggerService.Success($"Operação SCCM para {cleanName} finalizada com sucesso via CMSite.", nameof(AdSccmService));
            }
        }
        catch (Exception ex)
        {
            onLog($"  -> Falha ao comunicar com o SCCM: {ex.Message}");
            AppLoggerService.Danger($"Erro SCCM para {cleanName}: {ex.Message}", nameof(AdSccmService));
        }

        onLog($"=== OPERAÇÃO FINALIZADA PARA: {cleanName} ===");
    }

    public async Task<List<MachineDetails>> SearchMachinesAsync(string query, Action<string>? onLog = null)
    {
        return await Task.Run(async () =>
        {
            string cleanQuery = query.Trim();
            if (string.IsNullOrWhiteSpace(cleanQuery))
                return new List<MachineDetails>();

            void Log(string msg) => onLog?.Invoke(msg);

            Log($"Iniciando pesquisa ampla (máquina ou usuário) por: '{cleanQuery}'...");
            var resultsDict = new Dictionary<string, MachineDetails>(StringComparer.OrdinalIgnoreCase);

            // 1. Pesquisa no Active Directory (LDAP) com wildcard
            Log("[1/3] Buscando no Active Directory (LDAP)...");
            try
            {
                using var searcher = new DirectorySearcher();
                searcher.Filter = $"(&(objectCategory=computer)(|(cn=*{cleanQuery}*)(sAMAccountName=*{cleanQuery}*)(description=*{cleanQuery}*)))";
                searcher.PropertiesToLoad.AddRange(["cn", "dNSHostName", "operatingSystem", "distinguishedName", "whenCreated", "description"]);
                searcher.SizeLimit = 20;

                var allResults = searcher.FindAll();
                foreach (SearchResult res in allResults)
                {
                    using var entry = res.GetDirectoryEntry();
                    string cn = entry.Properties["cn"].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(cn)) continue;

                    string key = cn.ToUpperInvariant();
                    if (!resultsDict.TryGetValue(key, out var item))
                    {
                        item = new MachineDetails { ComputerName = cn };
                        resultsDict[key] = item;
                    }

                    item.FoundInAd = true;
                    item.DnsHostName = entry.Properties["dNSHostName"].Value?.ToString() ?? "";
                    item.OperatingSystem = entry.Properties["operatingSystem"].Value?.ToString() ?? "Windows";
                    item.DistinguishedName = entry.Properties["distinguishedName"].Value?.ToString() ?? "";
                    item.WhenCreated = entry.Properties["whenCreated"].Value?.ToString() ?? "";
                }
                Log($"  [OK] {allResults.Count} máquina(s) localizada(s) no Active Directory.");
            }
            catch (Exception ex)
            {
                Log($"  [Aviso] Falha na busca AD: {ex.Message}");
            }

            // 2. Pesquisa no Configuration Manager (SCCM) com LIKE
            Log("[2/3] Buscando no Configuration Manager (SCCM)...");
            string siteServer = ConfigService.SccmSiteServer;
            string siteCode = ConfigService.SccmSiteCode;

            if (!string.IsNullOrWhiteSpace(siteServer) && !string.IsNullOrWhiteSpace(siteCode))
            {
                bool sccmOk = false;

                // Tentativa via WMI C# Nativo com LIKE
                try
                {
                    var options = new ConnectionOptions
                    {
                        Impersonation = ImpersonationLevel.Impersonate,
                        Authentication = AuthenticationLevel.PacketPrivacy,
                        EnablePrivileges = true,
                        Timeout = TimeSpan.FromSeconds(5)
                    };

                    var scope = new ManagementScope($@"\\{siteServer}\root\sms\site_{siteCode}", options);
                    scope.Connect();

                    string wql = $@"SELECT ResourceID, SMSID, ClientActiveStatus, OperatingSystemNameandVersion, LastLogonUserDomain, LastLogonUserName, UserDomain, UserName, MACAddresses, NetbiosName, Name 
                                    FROM SMS_R_System 
                                    WHERE NetbiosName LIKE '%{cleanQuery}%' 
                                       OR Name LIKE '%{cleanQuery}%' 
                                       OR LastLogonUserName LIKE '%{cleanQuery}%' 
                                       OR UserName LIKE '%{cleanQuery}%'
                                       OR MACAddresses LIKE '%{cleanQuery}%'";

                    using var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(wql));
                    using var sccmResults = searcher.Get();

                    int wmiFoundCount = 0;
                    foreach (ManagementObject obj in sccmResults)
                    {
                        string name = obj["NetbiosName"]?.ToString() ?? obj["Name"]?.ToString() ?? "";
                        if (string.IsNullOrEmpty(name)) continue;

                        string key = name.ToUpperInvariant();
                        if (!resultsDict.TryGetValue(key, out var item))
                        {
                            item = new MachineDetails { ComputerName = name };
                            resultsDict[key] = item;
                        }

                        item.FoundInSccm = true;
                        item.ResourceId = obj["ResourceID"]?.ToString() ?? "";
                        item.SmsId = obj["SMSID"]?.ToString() ?? "";
                        string clientActive = obj["ClientActiveStatus"]?.ToString() ?? "";
                        item.ClientStatus = clientActive == "1" ? "Ativo" : "Inativo";
                        item.SiteCode = siteCode;

                        string lastUser = obj["LastLogonUserName"]?.ToString() ?? "";
                        string lastDomain = obj["LastLogonUserDomain"]?.ToString() ?? "";
                        string primaryUser = obj["UserName"]?.ToString() ?? "";
                        string primaryDomain = obj["UserDomain"]?.ToString() ?? "";

                        if (!string.IsNullOrEmpty(lastUser))
                        {
                            item.CurrentLogonUser = string.IsNullOrEmpty(lastDomain) ? lastUser : $"{lastDomain}\\{lastUser}";
                        }
                        else if (!string.IsNullOrEmpty(primaryUser))
                        {
                            item.CurrentLogonUser = string.IsNullOrEmpty(primaryDomain) ? primaryUser : $"{primaryDomain}\\{primaryUser}";
                        }

                        if (!string.IsNullOrEmpty(primaryUser))
                        {
                            item.PrimaryUser = string.IsNullOrEmpty(primaryDomain) ? primaryUser : $"{primaryDomain}\\{primaryUser}";
                        }

                        if (string.IsNullOrEmpty(item.OperatingSystem))
                        {
                            item.OperatingSystem = obj["OperatingSystemNameandVersion"]?.ToString() ?? "";
                        }

                        // MAC Address
                        if (obj["MACAddresses"] is string[] macs && macs.Length > 0)
                        {
                            item.MacAddress = string.Join(", ", macs);
                        }
                        else if (obj["MACAddresses"] is string singleMac)
                        {
                            item.MacAddress = singleMac;
                        }

                        wmiFoundCount++;
                    }

                    Log($"  [OK] {wmiFoundCount} registro(s) obtido(s) no SCCM via WMI Nativo.");
                    sccmOk = true;
                }
                catch (Exception ex)
                {
                    AppLoggerService.Info($"Busca WMI SCCM falhou ({ex.Message}), tentando provedor CMSite...", nameof(AdSccmService));
                }

                // Fallback via CMSite Provider
                if (!sccmOk)
                {
                    TrySearchSccmViaProvider(cleanQuery, resultsDict, siteServer, siteCode, onLog);
                }
            }

            // 3. Verificação de Conectividade de Rede Paralela (Ping com prioridade IPv4)
            Log($"[3/3] Diagnosticando status de rede de {resultsDict.Count} máquina(s)...");
            var pingTasks = resultsDict.Values.Select(async item =>
            {
                try
                {
                    using var ping = new Ping();
                    var pingTask = ping.SendPingAsync(item.ComputerName, 1000);
                    var completed = await Task.WhenAny(pingTask, Task.Delay(1200));

                    if (completed == pingTask && pingTask.Result.Status == IPStatus.Success)
                    {
                        var reply = pingTask.Result;
                        item.IsOnline = true;
                        item.PingResponseTime = $"{reply.RoundtripTime} ms";

                        if (reply.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            item.IpAddress = reply.Address.ToString();
                        }
                        else
                        {
                            try
                            {
                                var dnsTask = Dns.GetHostAddressesAsync(item.ComputerName);
                                if (await Task.WhenAny(dnsTask, Task.Delay(800)) == dnsTask)
                                {
                                    var addresses = dnsTask.Result;
                                    var ipv4 = addresses.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                                    item.IpAddress = ipv4 != null ? ipv4.ToString() : reply.Address.ToString();
                                }
                                else
                                {
                                    item.IpAddress = reply.Address.ToString();
                                }
                            }
                            catch
                            {
                                item.IpAddress = reply.Address.ToString();
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            var dnsTask = Dns.GetHostAddressesAsync(item.ComputerName);
                            if (await Task.WhenAny(dnsTask, Task.Delay(800)) == dnsTask)
                            {
                                var addresses = dnsTask.Result;
                                var ipv4 = addresses.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                                if (ipv4 != null) item.IpAddress = ipv4.ToString();
                                else if (addresses.Length > 0) item.IpAddress = addresses[0].ToString();
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            });

            await Task.WhenAll(pingTasks);

            var list = resultsDict.Values
                .OrderByDescending(m => m.IsOnline)
                .ThenBy(m => m.ComputerName)
                .ToList();

            Log($"Pesquisa concluída: {list.Count} máquina(s) encontrada(s).");
            return list;
        });
    }

    private void TrySearchSccmViaProvider(string cleanQuery, Dictionary<string, MachineDetails> resultsDict, string siteServer, string siteCode, Action<string>? onLog)
    {
        string? modulePath = FindSccmModulePath();
        if (modulePath == null) return;

        try
        {
            string scriptTemp = Path.Combine(Path.GetTempPath(), $"sccm_search_{Guid.NewGuid():N}.ps1");
            string psCode = $@"
$ErrorActionPreference = 'SilentlyContinue'
Import-Module '{modulePath}'
if (-not (Get-PSDrive -Name '{siteCode}' -ErrorAction SilentlyContinue)) {{
    New-PSDrive -Name '{siteCode}' -PSProvider 'CMSite' -Root '{siteServer}' -Description 'SCCM' | Out-Null
}}
Set-Location '{siteCode}:'
$all = Get-CMDevice | Where-Object {{
    $_.Name -like '*{cleanQuery}*' -or 
    $_.CurrentLogonUser -like '*{cleanQuery}*' -or 
    $_.LastLogonUserName -like '*{cleanQuery}*' -or 
    $_.UserName -like '*{cleanQuery}*' -or 
    $_.MACAddress -like '*{cleanQuery}*'
}} | Select-Object -First 15

foreach ($dev in $all) {{
    Write-Output '---DEV---'
    Write-Output ""Name=$($dev.Name)""
    Write-Output ""ResourceID=$($dev.ResourceID)""
    Write-Output ""SMSID=$($dev.SMSID)""
    Write-Output ""ClientActiveStatus=$($dev.ClientActiveStatus)""
    Write-Output ""CurrentLogonUser=$($dev.CurrentLogonUser)""
    Write-Output ""PrimaryUser=$($dev.PrimaryUser)""
    Write-Output ""LastLogonUserName=$($dev.LastLogonUserName)""
    Write-Output ""LastLogonUserDomain=$($dev.LastLogonUserDomain)""
    Write-Output ""UserName=$($dev.UserName)""
    Write-Output ""UserDomain=$($dev.UserDomain)""
    Write-Output ""MACAddress=$($dev.MACAddress)""
    Write-Output ""DeviceOS=$($dev.DeviceOS)""
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

            string defaultPsModules = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\Modules");
            startInfo.EnvironmentVariables["PSModulePath"] = defaultPsModules;
            if (startInfo.EnvironmentVariables.ContainsKey("DOTNET_ROOT"))
            {
                startInfo.EnvironmentVariables.Remove("DOTNET_ROOT");
            }

            using var proc = Process.Start(startInfo);
            string output = proc?.StandardOutput.ReadToEnd() ?? "";
            if (proc != null && !proc.WaitForExit(35000))
            {
                try { proc.Kill(); } catch { }
                onLog?.Invoke("  [Aviso] Tempo limite excedido ao consultar SCCM via CMSite.");
            }

            try { File.Delete(scriptTemp); } catch { }

            var devBlocks = output.Split(new[] { "---DEV---" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var block in devBlocks)
            {
                var lines = block.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string name = "";
                string resId = "";
                string smsId = "";
                string status = "";
                string curLogon = "";
                string primaryUser = "";
                string lastLogonUser = "";
                string lastLogonDomain = "";
                string devUser = "";
                string devDomain = "";
                string mac = "";
                string os = "";

                foreach (var line in lines)
                {
                    var parts = line.Split('=', 2);
                    if (parts.Length != 2) continue;
                    string k = parts[0].Trim();
                    string v = parts[1].Trim();

                    if (k == "Name") name = v;
                    else if (k == "ResourceID") resId = v;
                    else if (k == "SMSID") smsId = v;
                    else if (k == "ClientActiveStatus") status = v == "1" ? "Ativo" : "Inativo";
                    else if (k == "CurrentLogonUser") curLogon = v;
                    else if (k == "PrimaryUser") primaryUser = v;
                    else if (k == "LastLogonUserName") lastLogonUser = v;
                    else if (k == "LastLogonUserDomain") lastLogonDomain = v;
                    else if (k == "UserName") devUser = v;
                    else if (k == "UserDomain") devDomain = v;
                    else if (k == "MACAddress") mac = v;
                    else if (k == "DeviceOS") os = v;
                }

                if (string.IsNullOrEmpty(name)) continue;

                string key = name.ToUpperInvariant();
                if (!resultsDict.TryGetValue(key, out var item))
                {
                    item = new MachineDetails { ComputerName = name };
                    resultsDict[key] = item;
                }

                item.FoundInSccm = true;
                item.ResourceId = resId;
                item.SmsId = smsId;
                item.ClientStatus = status;
                item.SiteCode = siteCode;
                item.MacAddress = mac;

                if (!string.IsNullOrEmpty(curLogon))
                {
                    item.CurrentLogonUser = curLogon;
                }
                else if (!string.IsNullOrEmpty(lastLogonUser))
                {
                    item.CurrentLogonUser = string.IsNullOrEmpty(lastLogonDomain) ? lastLogonUser : $"{lastLogonDomain}\\{lastLogonUser}";
                }
                else if (!string.IsNullOrEmpty(primaryUser))
                {
                    item.CurrentLogonUser = primaryUser;
                }
                else if (!string.IsNullOrEmpty(devUser))
                {
                    item.CurrentLogonUser = string.IsNullOrEmpty(devDomain) ? devUser : $"{devDomain}\\{devUser}";
                }

                if (!string.IsNullOrEmpty(primaryUser))
                {
                    item.PrimaryUser = primaryUser;
                }
                else if (!string.IsNullOrEmpty(devUser))
                {
                    item.PrimaryUser = string.IsNullOrEmpty(devDomain) ? devUser : $"{devDomain}\\{devUser}";
                }

                if (string.IsNullOrEmpty(item.OperatingSystem))
                {
                    item.OperatingSystem = os;
                }
            }

            onLog?.Invoke($"  [OK] {devBlocks.Length} registro(s) SCCM via CMSite.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Erro ao buscar SCCM via CMSite: {ex.Message}", nameof(AdSccmService));
        }
    }
}
