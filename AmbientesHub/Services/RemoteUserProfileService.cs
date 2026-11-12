using System.IO;
using System.Management;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

public class RemoteUserProfileService
{
    private static bool IsProtectedUser(string user)
    {
        if (string.IsNullOrWhiteSpace(user)) return true;
        var protectedList = ConfigService.ProtectedProfiles
            .Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        return protectedList.Any(p => p.Equals(user, StringComparison.OrdinalIgnoreCase))
            || user.Equals("Public", StringComparison.OrdinalIgnoreCase)
            || user.Equals("Default", StringComparison.OrdinalIgnoreCase)
            || user.Equals("Default User", StringComparison.OrdinalIgnoreCase)
            || user.Equals("All Users", StringComparison.OrdinalIgnoreCase);
    }

    public List<RemoteUserProfileItem> QueryUserProfiles(
        string computerName,
        Action<string>? onLog = null,
        bool calculateDiskSize = true)
    {
        var profiles = new List<RemoteUserProfileItem>();
        string target = RemoteManagementHelper.NormalizeTargetHost(computerName);

        void Log(string msg) => onLog?.Invoke(msg);

        Log($"=== CONSULTANDO PERFIS DE USUÁRIOS EM: {target} ===");

        if (!RemoteManagementHelper.PingHost(target))
        {
            Log($"[ERRO] Máquina '{target}' não respondeu ao ping (ICMP).");
            return profiles;
        }

        try
        {
            RemoteManagementHelper.EnsureAdminShareConnected(target);
            var scope = RemoteManagementHelper.CreateScope(target, @"root\cimv2");
            var regScope = RemoteManagementHelper.CreateScope(target, @"root\default");

            using var searcher = new ManagementObjectSearcher(
                scope,
                new ObjectQuery("SELECT SID, LocalPath, Loaded, Special, LastUseTime FROM Win32_UserProfile WHERE Special = False"));

            foreach (ManagementObject prof in searcher.Get())
            {
                using (prof)
                {
                    string localPath = prof["LocalPath"]?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(localPath) ||
                        !localPath.Contains(@"\Users\", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string sid = prof["SID"]?.ToString() ?? string.Empty;
                    bool loaded = prof["Loaded"] != null && Convert.ToBoolean(prof["Loaded"]);
                    bool special = prof["Special"] != null && Convert.ToBoolean(prof["Special"]);
                    string userFolder = Path.GetFileName(localPath.TrimEnd('\\'));

                    if (string.IsNullOrWhiteSpace(userFolder) || IsProtectedUser(userFolder))
                    {
                        continue;
                    }

                    // 1. Obter data real de último uso via ProfileList (LoadTime/UnloadTime) ou WMI LastUseTime / ntuser.dat
                    DateTime? lastUseDt = GetRealProfileLastUse(regScope, target, sid, localPath, prof["LastUseTime"]?.ToString());
                    string lastUseStr = lastUseDt.HasValue
                        ? lastUseDt.Value.ToString("dd/MM/yyyy HH:mm:ss")
                        : "N/D";

                    // 2. Consultar Active Directory (Nome, Descrição, Escritório)
                    var (displayName, description, office) = RemoteManagementHelper.LookupAdUser(userFolder);

                    // 3. Calcular tamanho em disco (GB)
                    double sizeGb = 0;
                    if (calculateDiskSize)
                    {
                        string targetFolder = RemoteManagementHelper.ConvertLocalPathToTargetPath(target, localPath);
                        sizeGb = RemoteManagementHelper.CalculateDirectorySizeGb(targetFolder);
                    }

                    var item = new RemoteUserProfileItem
                    {
                        Nome = string.IsNullOrWhiteSpace(displayName) ? userFolder : displayName,
                        Usuario = userFolder,
                        Descricao = description,
                        Escritorio = office,
                        CaminhoPerfil = localPath,
                        Sid = sid,
                        UltimoUso = lastUseStr,
                        UltimoUsoDate = lastUseDt,
                        Loaded = loaded,
                        Special = special,
                        EspacoDiscoGb = sizeGb,
                        IsSelected = false
                    };

                    profiles.Add(item);
                }
            }

            profiles = profiles
                .OrderByDescending(p => p.Loaded)
                .ThenByDescending(p => p.UltimoUsoDate ?? DateTime.MinValue)
                .ToList();

            double totalGb = profiles.Sum(p => p.EspacoDiscoGb);
            int removable = profiles.Count(p => p.CanBeRemoved);
            Log($"[OK] {profiles.Count} perfil(is) encontrado(s) em '{target}' ({removable} elegível(is) para remoção | Total: {totalGb:N2} GB).");

            foreach (var p in profiles)
            {
                string statusTag = p.Loaded ? "[LOGADO - PROTEGIDO]" : "[Inativo]";
                Log($"  • {p.Usuario,-22} | {p.EspacoDiscoGb,6:N2} GB | Último uso: {p.UltimoUso,-19} | {statusTag} | {p.Nome}");
            }
        }
        catch (Exception ex)
        {
            Log($"[ERRO] Falha ao consultar perfis em '{target}': {ex.Message}");
        }

        return profiles;
    }

    public RemoteProfilePurgeResult PurgeUserProfiles(
        string computerName,
        IEnumerable<string> usersToPurge,
        Action<string>? onLog = null)
    {
        string target = RemoteManagementHelper.NormalizeTargetHost(computerName);
        var result = new RemoteProfilePurgeResult { ComputerName = target };

        void Log(string msg) => onLog?.Invoke(msg);

        var cleanUsers = usersToPurge
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .Select(u => u.Contains('\\') ? u.Split('\\').Last().Trim() : u.Trim())
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (cleanUsers.Count == 0)
        {
            Log("[AVISO] Nenhum perfil de usuário selecionado ou informado para remoção.");
            return result;
        }

        Log($"=== INICIANDO REMOÇÃO DE PERFIS EM: {target} ({cleanUsers.Count} usuário(s)) ===");

        if (!RemoteManagementHelper.PingHost(target))
        {
            Log($"[ERRO] Máquina '{target}' inacessível via ICMP. Operação abortada.");
            return result;
        }

        try
        {
            RemoteManagementHelper.EnableWinRmAndFirewallViaDcom(target, onLog);
            RemoteManagementHelper.EnsureAdminShareConnected(target);

            var scope = RemoteManagementHelper.CreateScope(target, @"root\cimv2");
            var regScope = RemoteManagementHelper.CreateScope(target, @"root\default");

            foreach (string user in cleanUsers)
            {
                Log($"--------------------------------------------------");
                if (IsProtectedUser(user))
                {
                    Log($"[AVISO] '{user}' é um perfil de sistema ou protegido pelas configurações. Operação ignorada.");
                    result.SkippedCount++;
                    continue;
                }

                Log($">>> Processando usuário '{user}' em '{target}'...");

                // 1) Tenta excluir conta local (se existir; contas de domínio ignoram sem erro)
                RemoteManagementHelper.InvokeCommandOnTarget(target, $"net user \"{user}\" /delete", 15);

                // 2) Busca Win32_UserProfile correspondente
                using var searcher = new ManagementObjectSearcher(
                    scope,
                    new ObjectQuery("SELECT * FROM Win32_UserProfile WHERE Special = False"));

                ManagementObject? matchedProfile = null;
                foreach (ManagementObject prof in searcher.Get())
                {
                    string lp = prof["LocalPath"]?.ToString() ?? string.Empty;
                    string leaf = Path.GetFileName(lp.TrimEnd('\\'));
                    if (leaf.Equals(user, StringComparison.OrdinalIgnoreCase))
                    {
                        matchedProfile = prof;
                        break;
                    }
                    prof.Dispose();
                }

                if (matchedProfile == null)
                {
                    Log($"  -> [AVISO] Nenhum perfil encontrado em C:\\Users para '{user}'.");
                    result.SkippedCount++;
                    continue;
                }

                using (matchedProfile)
                {
                    string localPath = matchedProfile["LocalPath"]?.ToString() ?? $@"C:\Users\{user}";
                    string sid = matchedProfile["SID"]?.ToString() ?? string.Empty;
                    bool loaded = matchedProfile["Loaded"] != null && Convert.ToBoolean(matchedProfile["Loaded"]);
                    bool special = matchedProfile["Special"] != null && Convert.ToBoolean(matchedProfile["Special"]);

                    if (loaded || special)
                    {
                        Log($"  -> [PROTEGIDO] O perfil '{localPath}' está carregado/ativo na sessão atual. Pulando por segurança.");
                        result.SkippedCount++;
                        continue;
                    }

                    string targetUncPath = RemoteManagementHelper.ConvertLocalPathToTargetPath(target, localPath);
                    double sizeBeforeGb = RemoteManagementHelper.CalculateDirectorySizeGb(targetUncPath);

                    // 3) Tenta remoção limpa via WMI Win32_UserProfile.Delete()
                    bool deletedViaWmi = false;
                    try
                    {
                        Log($"  -> Removendo perfil '{localPath}' (SID: {sid}) via WMI...");
                        matchedProfile.Delete();
                        deletedViaWmi = true;
                        Log($"  -> [SUCESSO] Perfil '{user}' excluído via Win32_UserProfile ({sizeBeforeGb:N2} GB liberados).");
                    }
                    catch (Exception exWmi)
                    {
                        Log($"  -> [AVISO] WMI Delete falhou ({exWmi.Message}). Acionando fallback (rmdir + ProfileList Registry)...");
                    }

                    // 4) Fallback: rmdir /S /Q + limpeza de ProfileList\<SID> no Registro
                    if (!deletedViaWmi)
                    {
                        bool cmdOk = RemoteManagementHelper.InvokeCommandOnTarget(target, $"rmdir /S /Q \"{localPath}\"", 60);
                        if (!string.IsNullOrWhiteSpace(sid))
                        {
                            RemoteManagementHelper.DeleteStdRegKey(
                                regScope,
                                RemoteManagementHelper.HKEY_LOCAL_MACHINE,
                                $@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{sid}");
                        }

                        if (cmdOk || !Directory.Exists(targetUncPath))
                        {
                            Log($"  -> [SUCESSO] Perfil '{user}' removido via fallback (rmdir + ProfileList) ({sizeBeforeGb:N2} GB).");
                            deletedViaWmi = true;
                        }
                        else
                        {
                            Log($"  -> [ERRO] Não foi possível remover completamente a pasta '{localPath}'.");
                            result.FailedCount++;
                        }
                    }

                    if (deletedViaWmi)
                    {
                        result.RemovedCount++;
                        result.FreedEstimateGb = Math.Round(result.FreedEstimateGb + sizeBeforeGb, 2);
                    }
                }
            }

            result.Success = result.FailedCount == 0;
            Log($"==================================================");
            Log($"[CONCLUÍDO] Removidos: {result.RemovedCount} | Ignorados/Ativos: {result.SkippedCount} | Falhas: {result.FailedCount} | Espaço Liberado: ~{result.FreedEstimateGb:N2} GB");
        }
        catch (Exception ex)
        {
            Log($"[ERRO] Falha geral na limpeza de perfis em '{target}': {ex.Message}");
        }

        return result;
    }

    private static DateTime? GetRealProfileLastUse(
        ManagementScope regScope,
        string targetHost,
        string sid,
        string localPath,
        string? wmiLastUseTime)
    {
        var candidates = new List<DateTime>();

        if (!string.IsNullOrWhiteSpace(sid))
        {
            string subKey = $@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{sid}";

            uint? loadHigh = RemoteManagementHelper.GetStdRegDwordValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, subKey, "LocalProfileLoadTimeHigh");
            uint? loadLow = RemoteManagementHelper.GetStdRegDwordValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, subKey, "LocalProfileLoadTimeLow");
            if (loadHigh.HasValue && loadLow.HasValue && (loadHigh.Value != 0 || loadLow.Value != 0))
            {
                try
                {
                    long ft = ((long)loadHigh.Value << 32) | loadLow.Value;
                    candidates.Add(DateTime.FromFileTime(ft));
                }
                catch { }
            }

            uint? unloadHigh = RemoteManagementHelper.GetStdRegDwordValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, subKey, "LocalProfileUnloadTimeHigh");
            uint? unloadLow = RemoteManagementHelper.GetStdRegDwordValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, subKey, "LocalProfileUnloadTimeLow");
            if (unloadHigh.HasValue && unloadLow.HasValue && (unloadHigh.Value != 0 || unloadLow.Value != 0))
            {
                try
                {
                    long ft = ((long)unloadHigh.Value << 32) | unloadLow.Value;
                    candidates.Add(DateTime.FromFileTime(ft));
                }
                catch { }
            }
        }

        if (candidates.Count == 0 && !string.IsNullOrWhiteSpace(wmiLastUseTime))
        {
            try
            {
                candidates.Add(ManagementDateTimeConverter.ToDateTime(wmiLastUseTime));
            }
            catch { }
        }

        if (candidates.Count == 0)
        {
            try
            {
                string targetFolder = RemoteManagementHelper.ConvertLocalPathToTargetPath(targetHost, localPath);
                string ntUserDat = Path.Combine(targetFolder, "NTUSER.DAT");
                if (File.Exists(ntUserDat))
                {
                    candidates.Add(File.GetLastWriteTime(ntUserDat));
                }
                else if (Directory.Exists(targetFolder))
                {
                    candidates.Add(Directory.GetLastWriteTime(targetFolder));
                }
            }
            catch { }
        }

        return candidates.Count > 0 ? candidates.Max() : null;
    }
}

