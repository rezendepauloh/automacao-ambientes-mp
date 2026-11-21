using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Windows;
using System.Windows.Shell;

namespace AmbientesHub.Services;

/// <summary>
/// Gerencia a JumpList da barra de tarefas do Windows e o canal IPC (Named Pipe)
/// para encaminhar ações da JumpList para a instância já em execução.
/// </summary>
public static class TaskbarJumpListService
{
    public const string JumpArgPrefix = "--jump-action=";

    public static string GetPipeName()
    {
        int sessionId = Process.GetCurrentProcess().SessionId;
        string safeUser = new string((Environment.UserName ?? "user")
            .Where(c => char.IsLetterOrDigit(c) || c == '_' || c == '-')
            .ToArray());
        if (string.IsNullOrEmpty(safeUser)) safeUser = "user";
        return $"AmbientesHub_JumpPipe_{safeUser}_{sessionId}";
    }

    public static string? ExtractJumpAction(string[]? args)
    {
        if (args == null || args.Length == 0) return null;
        foreach (string arg in args)
        {
            if (!string.IsNullOrWhiteSpace(arg) &&
                arg.StartsWith(JumpArgPrefix, StringComparison.OrdinalIgnoreCase))
            {
                string action = arg[JumpArgPrefix.Length..].Trim();
                return string.IsNullOrEmpty(action) ? null : action;
            }
        }
        return null;
    }

    public static bool TrySendActionToRunningInstance(string action, string? pipeNameOverride = null, int timeoutMs = 1500)
    {
        if (string.IsNullOrWhiteSpace(action)) return false;

        string pipeName = string.IsNullOrWhiteSpace(pipeNameOverride) ? GetPipeName() : pipeNameOverride;
        try
        {
            using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.Out, PipeOptions.None);
            client.Connect(timeoutMs);
            byte[] payload = Encoding.UTF8.GetBytes(action.Trim() + "\n");
            client.Write(payload, 0, payload.Length);
            client.Flush();
            AppLoggerService.Info($"Ação JumpList '{action}' enviada para instância ativa via pipe '{pipeName}'.", nameof(TaskbarJumpListService));
            return true;
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(TaskbarJumpListService), $"Enviar ação JumpList '{action}' via pipe '{pipeName}'", ex, level: LogLevel.WARNING);
            return false;
        }
    }

    public static CancellationTokenSource StartPipeServer(Action<string> onActionReceived, string? pipeNameOverride = null)
    {
        string pipeName = string.IsNullOrWhiteSpace(pipeNameOverride) ? GetPipeName() : pipeNameOverride;
        var cts = new CancellationTokenSource();

        _ = Task.Run(async () =>
        {
            while (!cts.Token.IsCancellationRequested)
            {
                try
                {
                    using var server = new NamedPipeServerStream(
                        pipeName,
                        PipeDirection.In,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await server.WaitForConnectionAsync(cts.Token).ConfigureAwait(false);
                    using var reader = new StreamReader(server, Encoding.UTF8);
                    string? line = await reader.ReadLineAsync(cts.Token).ConfigureAwait(false);
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        string action = line.Trim();
                        AppLoggerService.Info($"Ação JumpList recebida via pipe: '{action}'.", nameof(TaskbarJumpListService));
                        onActionReceived(action);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    AppLoggerService.LogException(nameof(TaskbarJumpListService), $"Loop do servidor pipe '{pipeName}'", ex, level: LogLevel.WARNING);
                    try { await Task.Delay(300, cts.Token).ConfigureAwait(false); } catch { break; }
                }
            }
        }, cts.Token);

        return cts;
    }

    public static void RegisterJumpList()
    {
        try
        {
            string? exePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
            {
                AppLoggerService.Warning("RegisterJumpList ignorado: Environment.ProcessPath indisponível.", nameof(TaskbarJumpListService));
                return;
            }

            var jumpList = new JumpList
            {
                ShowFrequentCategory = false,
                ShowRecentCategory = false
            };

            JumpTask MakeTask(string category, string title, string description, string action) => new()
            {
                CustomCategory = category,
                Title = title,
                Description = description,
                ApplicationPath = exePath,
                Arguments = $"{JumpArgPrefix}{action}",
                IconResourcePath = exePath,
                IconResourceIndex = 0
            };

            // Ações rápidas de ambiente
            jumpList.JumpItems.Add(MakeTask("Ambientes", "Modo Chamados", "Abre o ambiente de Chamados (OTRS / CitSmart)", "env:Chamados"));
            jumpList.JumpItems.Add(MakeTask("Ambientes", "Modo TI / Admin", "Abre o ambiente de TI / Administração", "env:TI"));
            jumpList.JumpItems.Add(MakeTask("Ambientes", "Saneamento Rápido", "Limpa processos do Office, navegadores e telemetria", "quick-clean"));

            // Acesso Remoto & Bancada
            jumpList.JumpItems.Add(MakeTask("Ferramentas", "SCCM Remoto (Clipboard)", "Conecta via CmRcViewer no host da área de transferência", "sccm-clipboard"));
            jumpList.JumpItems.Add(MakeTask("Ferramentas", "RDP Padrão (Clipboard)", "Conecta via mstsc no host da área de transferência", "rdp-clipboard"));
            jumpList.JumpItems.Add(MakeTask("Ferramentas", "Abrir Bancada no Navegador", "Abre a URL da Bancada STI", "open-bancada"));
            jumpList.JumpItems.Add(MakeTask("Ferramentas", "Iniciar Docker da Bancada", "Sobe os containers Docker da Bancada via WSL", "start-docker"));
            jumpList.JumpItems.Add(MakeTask("Ferramentas", "Parar Docker da Bancada", "Encerra os containers Docker da Bancada", "stop-docker"));

            // Controle do aplicativo
            jumpList.JumpItems.Add(MakeTask("Aplicativo", "Abrir / Restaurar Janela", "Restaura a janela principal do AmbientesHub", "open"));
            jumpList.JumpItems.Add(MakeTask("Aplicativo", "Ver Logs do Aplicativo", "Abre a Central de Logs no AmbientesHub", "open-logs"));
            jumpList.JumpItems.Add(MakeTask("Aplicativo", "Fechar AmbientesHub (Definitivo)", "Encerra completamente o processo (libera para dotnet publish)", "exit"));

            JumpList.SetJumpList(Application.Current, jumpList);
            jumpList.Apply();

            AppLoggerService.Success($"JumpList da barra de tarefas registrada com {jumpList.JumpItems.Count} atalho(s).", nameof(TaskbarJumpListService));
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(TaskbarJumpListService), "Registrar JumpList da barra de tarefas", ex, level: LogLevel.WARNING);
        }
    }
}
