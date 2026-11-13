using System.IO;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using AmbientesHub.Models;
using Microsoft.Win32;

namespace AmbientesHub.Services;

public class RegistryOptimizerService
{
    private static readonly string BackupDirectory = @"C:\AmbientesVirtuais\Backups";
    private static readonly string LatestRollbackJsonPath = Path.Combine(BackupDirectory, "registry_rollback_latest.json");

    public static string BackupFolderPath => BackupDirectory;

    public static bool IsRunningAsAdministrator()
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

    public List<RegistryTweakItem> GetDefaultTweaks()
    {
        return
        [
            // 1. Responsividade & UI
            new RegistryTweakItem
            {
                Id = "MenuShowDelay",
                Title = "Latência de Menus de Contexto",
                Description = "Reduz o tempo de resposta e abertura de submenus e menus de contexto no Windows de 400ms para 50ms.",
                Category = "Responsividade & UI",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Control Panel\Desktop",
                ValueName = "MenuShowDelay",
                ValueKind = RegistryValueKind.String,
                DefaultValue = "400",
                OptimizedValue = "50",
                RequiresAdmin = false,
                Emoji = "⚡"
            },
            new RegistryTweakItem
            {
                Id = "DisableSearchBoxSuggestions",
                Title = "Desativar Bing no Menu Iniciar",
                Description = "Remove pesquisas na Web e sugestões do Bing da caixa de busca do Windows, acelerando a abertura de programas locais.",
                Category = "Responsividade & UI",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Software\Policies\Microsoft\Windows\Explorer",
                ValueName = "DisableSearchBoxSuggestions",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 0,
                OptimizedValue = 1,
                RequiresAdmin = false,
                Emoji = "🔍"
            },
            new RegistryTweakItem
            {
                Id = "HungAppTimeout",
                Title = "Detecção de Janelas Travadas",
                Description = "Reduz de 5s para 1s o tempo para o Windows considerar um aplicativo como 'Não Respondendo', liberando a interface gráfica.",
                Category = "Responsividade & UI",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Control Panel\Desktop",
                ValueName = "HungAppTimeout",
                ValueKind = RegistryValueKind.String,
                DefaultValue = "5000",
                OptimizedValue = "1000",
                RequiresAdmin = false,
                Emoji = "⏱️"
            },

            // 2. Inicialização & Encerramento
            new RegistryTweakItem
            {
                Id = "StartupDelayInMSec",
                Title = "Atraso de Inicialização de Programas",
                Description = "Elimina o atraso artificial de 4 a 10 segundos imposto pelo Windows para inicializar programas pós-login.",
                Category = "Inicialização & Encerramento",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize",
                ValueName = "StartupDelayInMSec",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 4000,
                OptimizedValue = 0,
                RequiresAdmin = false,
                Emoji = "🚀"
            },
            new RegistryTweakItem
            {
                Id = "WaitToKillAppTimeout",
                Title = "Tempo de Encerramento de Apps Travados",
                Description = "Reduz de 20s para 2s o tempo de espera do Windows para fechar processos travados durante o desligamento.",
                Category = "Inicialização & Encerramento",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Control Panel\Desktop",
                ValueName = "WaitToKillAppTimeout",
                ValueKind = RegistryValueKind.String,
                DefaultValue = "20000",
                OptimizedValue = "2000",
                RequiresAdmin = false,
                Emoji = "🛑"
            },
            new RegistryTweakItem
            {
                Id = "AutoEndTasks",
                Title = "Encerramento Forçado no Shutdown",
                Description = "Força o encerramento automático de tarefas pendentes no desligamento sem bloquear a tela com avisos interativos.",
                Category = "Inicialização & Encerramento",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Control Panel\Desktop",
                ValueName = "AutoEndTasks",
                ValueKind = RegistryValueKind.String,
                DefaultValue = "0",
                OptimizedValue = "1",
                RequiresAdmin = false,
                Emoji = "🔒"
            },
            new RegistryTweakItem
            {
                Id = "WaitToKillServiceTimeout",
                Title = "Tempo Limite de Serviços do Sistema",
                Description = "Reduz o tempo de espera para serviços do sistema responderem ao sinal de parada durante o desligamento.",
                Category = "Inicialização & Encerramento",
                Hive = RegistryHive.LocalMachine,
                SubKeyPath = @"SYSTEM\CurrentControlSet\Control",
                ValueName = "WaitToKillServiceTimeout",
                ValueKind = RegistryValueKind.String,
                DefaultValue = "5000",
                OptimizedValue = "2000",
                RequiresAdmin = true,
                Emoji = "⚙️"
            },

            // 3. Sistema & Background
            new RegistryTweakItem
            {
                Id = "SystemResponsiveness",
                Title = "Prioridade de CPU do Aplicativo Ativo",
                Description = "Reduz a reserva de ciclos de CPU para serviços secundários de 20% para 0%, priorizando o aplicativo em primeiro plano.",
                Category = "Sistema & Background",
                Hive = RegistryHive.LocalMachine,
                SubKeyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile",
                ValueName = "SystemResponsiveness",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 20,
                OptimizedValue = 0,
                RequiresAdmin = true,
                Emoji = "💻"
            },
            new RegistryTweakItem
            {
                Id = "DisableWindowsConsumerFeatures",
                Title = "Bloquear Apps Promocionais Automáticos",
                Description = "Impede que o Windows instale e sugira silenciosamente jogos e aplicativos de terceiros na inicialização.",
                Category = "Sistema & Background",
                Hive = RegistryHive.LocalMachine,
                SubKeyPath = @"SOFTWARE\Policies\Microsoft\Windows\CloudContent",
                ValueName = "DisableWindowsConsumerFeatures",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 0,
                OptimizedValue = 1,
                RequiresAdmin = true,
                Emoji = "🛡️"
            },
            new RegistryTweakItem
            {
                Id = "DisableTailoredExperiences",
                Title = "Desativar Experiências Personalizadas",
                Description = "Desativa anúncios, dicas e coletas personalizadas baseadas no uso de diagnóstico do Windows.",
                Category = "Sistema & Background",
                Hive = RegistryHive.CurrentUser,
                SubKeyPath = @"Software\Policies\Microsoft\Windows\CloudContent",
                ValueName = "DisableTailoredExperiencesWithDiagnosticData",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 0,
                OptimizedValue = 1,
                RequiresAdmin = false,
                Emoji = "🚫"
            },

            // 4. Rede & Multimídia
            new RegistryTweakItem
            {
                Id = "NetworkThrottlingIndex",
                Title = "Desativar Throttling de Rede TCP",
                Description = "Desativa a limitação artificial de pacotes de rede não multimídia (0xFFFFFFFF), melhorando a latência de RDP e rede corporativa.",
                Category = "Rede & Multimídia",
                Hive = RegistryHive.LocalMachine,
                SubKeyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile",
                ValueName = "NetworkThrottlingIndex",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 10,
                OptimizedValue = unchecked((int)0xFFFFFFFF),
                RequiresAdmin = true,
                Emoji = "🌐"
            },
            new RegistryTweakItem
            {
                Id = "DODownloadMode",
                Title = "Desativar Envio P2P do Windows Update",
                Description = "Desativa o compartilhamento P2P de atualizações pela Internet, poupando largura de banda e conexões TCP.",
                Category = "Rede & Multimídia",
                Hive = RegistryHive.LocalMachine,
                SubKeyPath = @"SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization",
                ValueName = "DODownloadMode",
                ValueKind = RegistryValueKind.DWord,
                DefaultValue = 1,
                OptimizedValue = 0,
                RequiresAdmin = true,
                Emoji = "📦"
            }
        ];
    }

    public void ScanTweak(RegistryTweakItem tweak)
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(tweak.Hive, RegistryView.Default);
            using var subKey = baseKey.OpenSubKey(tweak.SubKeyPath, false);

            if (subKey == null)
            {
                tweak.CurrentRawValue = null;
                tweak.CurrentValueDisplay = "(chave ausente)";
                tweak.Status = RegistryTweakStatus.Missing;
                return;
            }

            object? val = subKey.GetValue(tweak.ValueName);
            if (val == null)
            {
                tweak.CurrentRawValue = null;
                tweak.CurrentValueDisplay = "(valor não definido)";
                tweak.Status = RegistryTweakStatus.Missing;
                return;
            }

            tweak.CurrentRawValue = val;
            tweak.CurrentValueDisplay = val.ToString() ?? "(indefinido)";

            if (ValuesAreEqual(val, tweak.OptimizedValue))
            {
                tweak.Status = RegistryTweakStatus.Optimized;
            }
            else if (ValuesAreEqual(val, tweak.DefaultValue))
            {
                tweak.Status = RegistryTweakStatus.Default;
            }
            else
            {
                tweak.Status = RegistryTweakStatus.Custom;
            }
        }
        catch (Exception ex)
        {
            tweak.CurrentRawValue = null;
            tweak.CurrentValueDisplay = $"Erro de acesso: {ex.Message}";
            tweak.Status = RegistryTweakStatus.NotChecked;
        }
    }

    public void ScanAll(IEnumerable<RegistryTweakItem> tweaks, Action<string>? onLog = null)
    {
        onLog?.Invoke("[INFO] Iniciando escaneamento cirúrgico de chaves do registro...");
        int count = 0;
        int optCount = 0;

        foreach (var tweak in tweaks)
        {
            ScanTweak(tweak);
            count++;
            if (tweak.IsOptimized) optCount++;
            onLog?.Invoke($"[STATUS] {tweak.Title} [{tweak.FullRegistryPath}]: {tweak.CurrentValueDisplay} ({tweak.StatusBadgeText})");
        }

        onLog?.Invoke($"[SUCESSO] Escaneamento concluído: {optCount} de {count} chaves já otimizadas.");
    }

    public string CreateBackup(IEnumerable<RegistryTweakItem> tweaks, string? customPath = null)
    {
        EnsureBackupDirectory();

        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string regPath = customPath ?? Path.Combine(BackupDirectory, $"RegBackup_{timestamp}.reg");

        var snapshotDict = new Dictionary<string, object?>();

        var sb = new StringBuilder();
        sb.AppendLine("Windows Registry Editor Version 5.00");
        sb.AppendLine();
        sb.AppendLine($"; Backup de Registro gerado pelo Ambientes Hub em {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        // Agrupa por Chave para formato .reg limpo
        var grouped = tweaks.GroupBy(t => new { t.Hive, t.SubKeyPath });
        foreach (var group in grouped)
        {
            string hiveName = group.Key.Hive == RegistryHive.CurrentUser ? "HKEY_CURRENT_USER" : "HKEY_LOCAL_MACHINE";
            sb.AppendLine($"[{hiveName}\\{group.Key.SubKeyPath}]");

            foreach (var tweak in group)
            {
                // Armazena no snapshot JSON
                snapshotDict[tweak.Id] = tweak.CurrentRawValue;

                if (tweak.CurrentRawValue != null)
                {
                    if (tweak.ValueKind == RegistryValueKind.DWord)
                    {
                        int dwordVal = Convert.ToInt32(tweak.CurrentRawValue);
                        sb.AppendLine($"\"{tweak.ValueName}\"=dword:{dwordVal:x8}");
                    }
                    else
                    {
                        string strVal = (tweak.CurrentRawValue.ToString() ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
                        sb.AppendLine($"\"{tweak.ValueName}\"=\"{strVal}\"");
                    }
                }
                else
                {
                    // No formato .reg, "-" remove o valor
                    sb.AppendLine($"\"{tweak.ValueName}\"=-");
                }
            }
            sb.AppendLine();
        }

        File.WriteAllText(regPath, sb.ToString(), Encoding.Unicode);
        ConfigService.HardenFilePermissions(regPath);

        // Grava também snapshot JSON para rollback interno
        try
        {
            string json = JsonSerializer.Serialize(snapshotDict, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(LatestRollbackJsonPath, json, Encoding.UTF8);
            ConfigService.HardenFilePermissions(LatestRollbackJsonPath);
        }
        catch { }

        AppLoggerService.Success($"Backup de registro salvo com sucesso em: {regPath}", nameof(RegistryOptimizerService));
        return regPath;
    }

    public (bool Success, string Message) ApplyTweak(RegistryTweakItem tweak, Action<string>? onLog = null)
    {
        if (tweak.RequiresAdmin && !IsRunningAsAdministrator())
        {
            string err = $"A chave '{tweak.Title}' reside em HKLM e exige execução do aplicativo como Administrador.";
            onLog?.Invoke($"[ERRO] {err}");
            AppLoggerService.Warning(err, nameof(RegistryOptimizerService));
            return (false, err);
        }

        try
        {
            // Salva valor anterior para rollback se ainda não gravado
            ScanTweak(tweak);
            if (tweak.PreviousRawValue == null && tweak.CurrentRawValue != null)
            {
                tweak.PreviousRawValue = tweak.CurrentRawValue;
            }

            using var baseKey = RegistryKey.OpenBaseKey(tweak.Hive, RegistryView.Default);
            using var subKey = baseKey.CreateSubKey(tweak.SubKeyPath, true);

            if (subKey == null)
            {
                string err = $"Não foi possível abrir/criar a chave: {tweak.FullRegistryPath}";
                onLog?.Invoke($"[ERRO] {err}");
                return (false, err);
            }

            if (tweak.ValueKind == RegistryValueKind.DWord)
            {
                int val = Convert.ToInt32(tweak.OptimizedValue);
                subKey.SetValue(tweak.ValueName, val, RegistryValueKind.DWord);
            }
            else
            {
                string val = tweak.OptimizedValue.ToString() ?? "";
                subKey.SetValue(tweak.ValueName, val, RegistryValueKind.String);
            }

            ScanTweak(tweak);

            string msg = $"Tweak '{tweak.Title}' aplicado com sucesso! ({tweak.ValueName} = {tweak.OptimizedValue})";
            onLog?.Invoke($"[SUCESSO] {msg}");
            AppLoggerService.Success(msg, nameof(RegistryOptimizerService));
            return (true, msg);
        }
        catch (Exception ex)
        {
            string err = $"Falha ao aplicar tweak '{tweak.Title}': {ex.Message}";
            onLog?.Invoke($"[ERRO] {err}");
            AppLoggerService.Danger(err, nameof(RegistryOptimizerService));
            return (false, err);
        }
    }

    public (bool Success, string Message) RollbackTweak(RegistryTweakItem tweak, Action<string>? onLog = null)
    {
        if (tweak.RequiresAdmin && !IsRunningAsAdministrator())
        {
            string err = $"Rollback de '{tweak.Title}' exige privilégios de Administrador.";
            onLog?.Invoke($"[ERRO] {err}");
            return (false, err);
        }

        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(tweak.Hive, RegistryView.Default);
            using var subKey = baseKey.OpenSubKey(tweak.SubKeyPath, true);

            if (subKey == null)
            {
                ScanTweak(tweak);
                return (true, $"A chave {tweak.SubKeyPath} já não existe.");
            }

            object? targetRestore = tweak.PreviousRawValue ?? tweak.DefaultValue;

            if (targetRestore == null)
            {
                // Se era ausente, remove
                subKey.DeleteValue(tweak.ValueName, false);
                onLog?.Invoke($"[RESTOURADO] Valor '{tweak.ValueName}' removido (restaurado para estado ausente padrão).");
            }
            else
            {
                if (tweak.ValueKind == RegistryValueKind.DWord)
                {
                    int dwordVal = Convert.ToInt32(targetRestore);
                    subKey.SetValue(tweak.ValueName, dwordVal, RegistryValueKind.DWord);
                }
                else
                {
                    subKey.SetValue(tweak.ValueName, targetRestore.ToString() ?? "", RegistryValueKind.String);
                }
                onLog?.Invoke($"[RESTAURADO] '{tweak.Title}' restaurado para: {targetRestore}");
            }

            tweak.PreviousRawValue = null;
            ScanTweak(tweak);

            string msg = $"Rollback de '{tweak.Title}' concluído com sucesso!";
            AppLoggerService.Info(msg, nameof(RegistryOptimizerService));
            return (true, msg);
        }
        catch (Exception ex)
        {
            string err = $"Falha no rollback de '{tweak.Title}': {ex.Message}";
            onLog?.Invoke($"[ERRO] {err}");
            AppLoggerService.Danger(err, nameof(RegistryOptimizerService));
            return (false, err);
        }
    }

    public (int Applied, int Failed) ApplyBatch(IEnumerable<RegistryTweakItem> tweaks, Action<string>? onLog = null)
    {
        var list = tweaks.Where(t => t.IsSelected).ToList();
        if (list.Count == 0)
        {
            onLog?.Invoke("[AVISO] Nenhum tweak foi selecionado para aplicação.");
            return (0, 0);
        }

        // Gera backup de segurança obrigatório antes do lote
        try
        {
            string backupPath = CreateBackup(list);
            onLog?.Invoke($"[BACKUP] Ponto de restauração gerado em: {backupPath}");
        }
        catch (Exception ex)
        {
            onLog?.Invoke($"[ALERTA] Não foi possível gerar backup automático: {ex.Message}");
        }

        int applied = 0;
        int failed = 0;

        foreach (var tweak in list)
        {
            var (success, _) = ApplyTweak(tweak, onLog);
            if (success) applied++;
            else failed++;
        }

        onLog?.Invoke($"[CONCLUÍDO] Lote finalizado: {applied} aplicado(s), {failed} falha(s).");
        return (applied, failed);
    }

    public (int Restored, int Failed) RollbackBatch(IEnumerable<RegistryTweakItem> tweaks, Action<string>? onLog = null)
    {
        var list = tweaks.Where(t => t.IsSelected).ToList();
        if (list.Count == 0)
        {
            onLog?.Invoke("[AVISO] Nenhum tweak selecionado para restauração.");
            return (0, 0);
        }

        int restored = 0;
        int failed = 0;

        foreach (var tweak in list)
        {
            var (success, _) = RollbackTweak(tweak, onLog);
            if (success) restored++;
            else failed++;
        }

        onLog?.Invoke($"[CONCLUÍDO] Rollback finalizado: {restored} restaurado(s), {failed} falha(s).");
        return (restored, failed);
    }

    private static void EnsureBackupDirectory()
    {
        if (!Directory.Exists(BackupDirectory))
        {
            Directory.CreateDirectory(BackupDirectory);
            ConfigService.HardenFilePermissions(BackupDirectory);
        }
    }

    private static bool ValuesAreEqual(object a, object b)
    {
        if (a is int aInt && b is int bInt) return aInt == bInt;
        return string.Equals(a.ToString(), b.ToString(), StringComparison.OrdinalIgnoreCase);
    }
}
