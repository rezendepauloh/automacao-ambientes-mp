using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;

namespace AmbientesHub.Services;

public static class RemoteManagementHelper
{
    public const uint HKEY_CLASSES_ROOT = 0x80000000;
    public const uint HKEY_CURRENT_USER = 0x80000001;
    public const uint HKEY_LOCAL_MACHINE = 0x80000002;
    public const uint HKEY_USERS = 0x80000003;

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetAddConnection2(ref NETRESOURCE netResource, string? password, string? username, int flags);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NETRESOURCE
    {
        public int dwScope;
        public int dwType;
        public int dwDisplayType;
        public int dwUsage;
        public string? lpLocalName;
        public string? lpRemoteName;
        public string? lpComment;
        public string? lpProvider;
    }

    public static bool IsLocalMachine(string? computerName)
    {
        if (string.IsNullOrWhiteSpace(computerName)) return true;
        string clean = computerName.Trim();
        return clean == "."
            || clean.Equals("localhost", StringComparison.OrdinalIgnoreCase)
            || clean.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase)
            || clean.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase);
    }

    public static string NormalizeTargetHost(string? computerName)
    {
        return IsLocalMachine(computerName) ? Environment.MachineName : computerName!.Trim();
    }

    public static bool PingHost(string? computerName, int timeoutMs = -1)
    {
        if (IsLocalMachine(computerName)) return true;
        if (timeoutMs <= 0) timeoutMs = ConfigService.PingTimeoutMs;
        try
        {
            using var ping = new Ping();
            var reply = ping.Send(computerName!.Trim(), timeoutMs);
            return reply.Status == IPStatus.Success;
        }
        catch
        {
            return false;
        }
    }

    public static ManagementScope CreateScope(string? computerName, string wmiNamespace = @"root\cimv2", TimeSpan? timeout = null)
    {
        bool isLocal = IsLocalMachine(computerName);
        string target = isLocal ? "." : computerName!.Trim();

        var options = new ConnectionOptions
        {
            Impersonation = ImpersonationLevel.Impersonate,
            Authentication = AuthenticationLevel.PacketPrivacy,
            EnablePrivileges = true,
            Timeout = timeout ?? TimeSpan.FromSeconds(25)
        };

        if (!isLocal)
        {
            var cred = CredentialService.ReadPlaintextCredential();
            if (cred.Success && !string.IsNullOrWhiteSpace(cred.Username) && cred.Password != null)
            {
                // Se a credencial no cofre for diferente do usuário atual do processo, aplica na sessão WMI
                string currentIdentity = $"{Environment.UserDomainName}\\{Environment.UserName}";
                if (!cred.Username.Equals(currentIdentity, StringComparison.OrdinalIgnoreCase) &&
                    !cred.Username.Equals(Environment.UserName, StringComparison.OrdinalIgnoreCase))
                {
                    options.Username = cred.Username;
                    options.Password = cred.Password;
                }
            }
        }

        var scope = new ManagementScope($@"\\{target}\{wmiNamespace}", options);
        scope.Connect();
        return scope;
    }

    public static void EnableWinRmAndFirewallViaDcom(string computerName, Action<string>? onLog = null)
    {
        if (IsLocalMachine(computerName)) return;

        try
        {
            var scope = CreateScope(computerName, @"root\cimv2", TimeSpan.FromSeconds(15));

            // 1. Habilitar e iniciar serviço WinRM via DCOM
            using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_Service WHERE Name='WinRM'")))
            {
                foreach (ManagementObject svc in searcher.Get())
                {
                    using (svc)
                    {
                        svc.InvokeMethod("ChangeStartMode", ["Automatic"]);
                        string state = svc["State"]?.ToString() ?? "";
                        if (!state.Equals("Running", StringComparison.OrdinalIgnoreCase))
                        {
                            svc.InvokeMethod("StartService", null);
                        }
                    }
                }
            }

            // 2. Liberar regra de firewall do WinRM via Win32_Process
            using (var procClass = new ManagementClass(scope, new ManagementPath("Win32_Process"), null))
            {
                using var inParams = procClass.GetMethodParameters("Create");
                inParams["CommandLine"] = "cmd.exe /c netsh advfirewall firewall set rule group=\"Windows Remote Management\" new enable=yes & netsh advfirewall firewall set rule group=\"Gerenciamento Remoto do Windows\" new enable=yes";
                procClass.InvokeMethod("Create", inParams, null);
            }

            // 3. Configurar TrustedHosts via StdRegProv
            var regScope = CreateScope(computerName, @"root\default", TimeSpan.FromSeconds(10));
            SetStdRegStringValue(
                regScope,
                HKEY_LOCAL_MACHINE,
                @"SOFTWARE\Policies\Microsoft\Windows\WinRM\Client",
                "TrustedHosts",
                Environment.MachineName);

            onLog?.Invoke($"  -> [DCOM] Serviço WinRM, Firewall e TrustedHosts preparados em '{computerName}'.");
        }
        catch (Exception ex)
        {
            onLog?.Invoke($"  -> [AVISO DCOM] Preparação auxiliar WinRM em '{computerName}': {ex.Message}");
        }
    }

    public static bool EnsureAdminShareConnected(string computerName)
    {
        if (IsLocalMachine(computerName)) return true;

        string uncRoot = $@"\\{computerName.Trim()}\C$";
        if (Directory.Exists(uncRoot)) return true;

        try
        {
            var cred = CredentialService.ReadPlaintextCredential();
            if (!cred.Success || string.IsNullOrWhiteSpace(cred.Username) || cred.Password == null)
                return Directory.Exists(uncRoot);

            var nr = new NETRESOURCE
            {
                dwType = 1, // RESOURCETYPE_DISK
                lpRemoteName = uncRoot
            };

            int res = WNetAddConnection2(ref nr, cred.Password, cred.Username, 0);
            return res == 0 || res == 1219 || Directory.Exists(uncRoot);
        }
        catch
        {
            return Directory.Exists(uncRoot);
        }
    }

    public static string ConvertLocalPathToTargetPath(string computerName, string localPath)
    {
        if (IsLocalMachine(computerName) || string.IsNullOrWhiteSpace(localPath))
            return localPath;

        if (localPath.Length >= 2 && localPath[1] == ':')
        {
            char drive = char.ToUpperInvariant(localPath[0]);
            string rest = localPath.Length > 2 ? localPath.Substring(2).TrimStart('\\') : string.Empty;
            return string.IsNullOrEmpty(rest)
                ? $@"\\{computerName.Trim()}\{drive}$"
                : $@"\\{computerName.Trim()}\{drive}$\{rest}";
        }

        return localPath;
    }

    public static string[]? EnumStdRegSubKeys(ManagementScope regScope, uint hDefKey, string subKeyName)
    {
        try
        {
            using var reg = new ManagementClass(regScope, new ManagementPath("StdRegProv"), null);
            using var inParams = reg.GetMethodParameters("EnumKey");
            inParams["hDefKey"] = hDefKey;
            inParams["sSubKeyName"] = subKeyName;
            using var outParams = reg.InvokeMethod("EnumKey", inParams, null);
            if (outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0)
            {
                return outParams["sNames"] as string[];
            }
        }
        catch { }
        return null;
    }

    public static string? GetStdRegStringValue(ManagementScope regScope, uint hDefKey, string subKeyName, string valueName)
    {
        try
        {
            using var reg = new ManagementClass(regScope, new ManagementPath("StdRegProv"), null);
            using var inParams = reg.GetMethodParameters("GetStringValue");
            inParams["hDefKey"] = hDefKey;
            inParams["sSubKeyName"] = subKeyName;
            inParams["sValueName"] = valueName;
            using var outParams = reg.InvokeMethod("GetStringValue", inParams, null);
            if (outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0)
            {
                return outParams["sValue"]?.ToString();
            }
        }
        catch { }
        return null;
    }

    public static uint? GetStdRegDwordValue(ManagementScope regScope, uint hDefKey, string subKeyName, string valueName)
    {
        try
        {
            using var reg = new ManagementClass(regScope, new ManagementPath("StdRegProv"), null);
            using var inParams = reg.GetMethodParameters("GetDWORDValue");
            inParams["hDefKey"] = hDefKey;
            inParams["sSubKeyName"] = subKeyName;
            inParams["sValueName"] = valueName;
            using var outParams = reg.InvokeMethod("GetDWORDValue", inParams, null);
            if (outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0 && outParams["uValue"] != null)
            {
                return Convert.ToUInt32(outParams["uValue"]);
            }
        }
        catch { }
        return null;
    }

    public static bool SetStdRegStringValue(ManagementScope regScope, uint hDefKey, string subKeyName, string valueName, string value)
    {
        try
        {
            using var reg = new ManagementClass(regScope, new ManagementPath("StdRegProv"), null);
            using var inParamsCreate = reg.GetMethodParameters("CreateKey");
            inParamsCreate["hDefKey"] = hDefKey;
            inParamsCreate["sSubKeyName"] = subKeyName;
            reg.InvokeMethod("CreateKey", inParamsCreate, null);

            using var inParams = reg.GetMethodParameters("SetStringValue");
            inParams["hDefKey"] = hDefKey;
            inParams["sSubKeyName"] = subKeyName;
            inParams["sValueName"] = valueName;
            inParams["sValue"] = value;
            using var outParams = reg.InvokeMethod("SetStringValue", inParams, null);
            return outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0;
        }
        catch
        {
            return false;
        }
    }

    public static bool DeleteStdRegKey(ManagementScope regScope, uint hDefKey, string subKeyName)
    {
        try
        {
            using var reg = new ManagementClass(regScope, new ManagementPath("StdRegProv"), null);
            using var inParams = reg.GetMethodParameters("DeleteKey");
            inParams["hDefKey"] = hDefKey;
            inParams["sSubKeyName"] = subKeyName;
            using var outParams = reg.InvokeMethod("DeleteKey", inParams, null);
            return outParams != null && Convert.ToUInt32(outParams["ReturnValue"]) == 0;
        }
        catch
        {
            return false;
        }
    }

    public static bool InvokeCommandOnTarget(string computerName, string commandLine, int timeoutSeconds = 90)
    {
        if (IsLocalMachine(computerName))
        {
            try
            {
                using var proc = Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c {commandLine}",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                });
                if (proc == null) return false;
                proc.WaitForExit(timeoutSeconds * 1000);
                return proc.ExitCode == 0;
            }
            catch
            {
                return false;
            }
        }

        try
        {
            var scope = CreateScope(computerName, @"root\cimv2");
            using var procClass = new ManagementClass(scope, new ManagementPath("Win32_Process"), null);
            using var inParams = procClass.GetMethodParameters("Create");
            inParams["CommandLine"] = $"cmd.exe /c {commandLine}";
            using var outParams = procClass.InvokeMethod("Create", inParams, null);
            if (outParams == null || Convert.ToUInt32(outParams["ReturnValue"]) != 0)
                return false;

            uint pid = Convert.ToUInt32(outParams["ProcessId"]);
            var sw = Stopwatch.StartNew();
            while (sw.Elapsed.TotalSeconds < timeoutSeconds)
            {
                Thread.Sleep(600);
                using var searcher = new ManagementObjectSearcher(scope, new ObjectQuery($"SELECT ProcessId FROM Win32_Process WHERE ProcessId = {pid}"));
                if (searcher.Get().Count == 0)
                    return true;
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static (string DisplayName, string Description, string Office) LookupAdUser(string? samAccountName)
    {
        if (string.IsNullOrWhiteSpace(samAccountName))
            return (string.Empty, string.Empty, string.Empty);

        string cleanUser = samAccountName.Contains('\\')
            ? samAccountName.Split('\\').Last().Trim()
            : samAccountName.Trim();

        if (string.IsNullOrWhiteSpace(cleanUser))
            return (string.Empty, string.Empty, string.Empty);

        try
        {
            using var searcher = new DirectorySearcher
            {
                Filter = $"(&(objectCategory=person)(objectClass=user)(sAMAccountName={cleanUser}))",
                SizeLimit = 1
            };
            searcher.PropertiesToLoad.AddRange(["displayName", "name", "description", "physicalDeliveryOfficeName"]);

            var result = searcher.FindOne();
            if (result == null)
                return (string.Empty, string.Empty, string.Empty);

            string displayName = result.Properties["displayName"].Count > 0
                ? result.Properties["displayName"][0]?.ToString() ?? ""
                : (result.Properties["name"].Count > 0 ? result.Properties["name"][0]?.ToString() ?? "" : "");

            string desc = result.Properties["description"].Count > 0
                ? result.Properties["description"][0]?.ToString() ?? ""
                : "";

            string office = result.Properties["physicalDeliveryOfficeName"].Count > 0
                ? result.Properties["physicalDeliveryOfficeName"][0]?.ToString() ?? ""
                : "";

            return (displayName, desc, office);
        }
        catch
        {
            return (string.Empty, string.Empty, string.Empty);
        }
    }

    public static double CalculateDirectorySizeGb(string directoryPath)
    {
        return Math.Round(CalculateDirectorySizeBytes(directoryPath) / (1024.0 * 1024.0 * 1024.0), 2);
    }

    public static (long Bytes, int FileCount) CalculateDirectoryStats(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
            return (0, 0);

        long totalBytes = 0;
        int count = 0;

        try
        {
            var opts = new System.IO.EnumerationOptions
            {
                IgnoreInaccessible = true,
                RecurseSubdirectories = true,
                AttributesToSkip = FileAttributes.ReparsePoint
            };

            foreach (var file in new DirectoryInfo(directoryPath).EnumerateFiles("*", opts))
            {
                try
                {
                    totalBytes += file.Length;
                    count++;
                }
                catch { }
            }
        }
        catch { }

        return (totalBytes, count);
    }

    public static long CalculateDirectorySizeBytes(string directoryPath)
    {
        return CalculateDirectoryStats(directoryPath).Bytes;
    }
}

