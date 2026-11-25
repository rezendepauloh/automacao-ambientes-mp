using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace AmbientesHub.Services;

/// <summary>
/// Inicia o SAJMP e realiza login automático via Win32 API (WM_SETTEXT / BM_CLICK) sem depender de SendKeys ou foco global.
/// </summary>
public class SajmpLauncherService
{
    private const int GwlStyle = -16;
    private const int EsPassword = 0x0020;
    private const uint WmSetText = 0x000C;
    private const uint WmKeyDown = 0x0100;
    private const uint WmKeyUp = 0x0101;
    private const int VkReturn = 0x0D;
    private const uint BmClick = 0x00F5;

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    public (bool Success, string Message) LaunchAndLogin(
        Action<string>? logCallback = null,
        string? exePath = null,
        string? credPath = null,
        string? windowTitleHint = null,
        int timeoutMs = 45000)
    {
        var sw = Stopwatch.StartNew();
        void Log(string msg) => logCallback?.Invoke(msg);

        exePath ??= ConfigService.SajmpExePath;
        credPath ??= ConfigService.SajmpCredPath;
        windowTitleHint ??= ConfigService.SajmpWindowTitle;

        if (string.IsNullOrWhiteSpace(exePath) || string.IsNullOrWhiteSpace(credPath) || string.IsNullOrWhiteSpace(windowTitleHint))
        {
            const string err = "Configuração SAJMP ausente no .env (SAJMP_EXE_PATH, SAJMP_CRED_PATH e SAJMP_WINDOW_TITLE são obrigatórios).";
            Log($"[ERRO] {err}");
            AppLoggerService.Warning(err, nameof(SajmpLauncherService));
            return (false, err);
        }

        if (!File.Exists(exePath))
        {
            string err = $"Executável do SAJMP não encontrado em '{exePath}'.";
            Log($"[ERRO] {err}");
            AppLoggerService.Warning(err, nameof(SajmpLauncherService));
            return (false, err);
        }

        var cred = CredentialService.ReadPlaintextCredential(credPath);
        if (!cred.Success || string.IsNullOrEmpty(cred.Password))
        {
            string err = $"Falha ao obter senha do cofre '{credPath}': {cred.Message}";
            Log($"[ERRO] {err}");
            AppLoggerService.Warning(err, nameof(SajmpLauncherService));
            return (false, err);
        }

        Log($"[▶] Iniciando '{exePath}'...");
        Process? proc;
        try
        {
            proc = Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                WorkingDirectory = Path.GetDirectoryName(exePath) ?? string.Empty,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            string err = $"Erro ao iniciar SAJMP: {ex.Message}";
            Log($"[ERRO] {err}");
            AppLoggerService.LogException(nameof(SajmpLauncherService), "Process.Start SAJMP", ex, sw.Elapsed);
            return (false, err);
        }

        if (proc == null)
        {
            const string err = "Não foi possível obter o processo do SAJMP.";
            Log($"[ERRO] {err}");
            return (false, err);
        }

        Log("[▶] Aguardando janela de login do SAJ...");
        IntPtr loginHwnd = WaitForLoginWindow((uint)proc.Id, windowTitleHint, timeoutMs);
        if (loginHwnd == IntPtr.Zero)
        {
            string err = $"Janela de login do SAJ não apareceu dentro de {timeoutMs / 1000}s.";
            Log($"[ERRO] {err}");
            AppLoggerService.Warning(err, nameof(SajmpLauncherService));
            return (false, err);
        }

        var (senhaHwnd, entrarHwnd) = FindLoginControls(loginHwnd);
        if (senhaHwnd == IntPtr.Zero)
        {
            const string err = "Campo de senha não localizado na janela de login do SAJ.";
            Log($"[ERRO] {err}");
            AppLoggerService.Warning(err, nameof(SajmpLauncherService));
            return (false, err);
        }

        Log("[▶] Preenchendo senha via Win32 (WM_SETTEXT)...");
        SendMessage(senhaHwnd, WmSetText, IntPtr.Zero, cred.Password);

        if (entrarHwnd != IntPtr.Zero)
        {
            Log("[▶] Acionando botão Entrar via Win32 (BM_CLICK)...");
            PostMessage(entrarHwnd, BmClick, IntPtr.Zero, IntPtr.Zero);
        }
        else
        {
            Log("[▶] Botão Entrar customizado; enviando Enter (WM_KEYDOWN VK_RETURN) direto ao campo de senha...");
            PostMessage(senhaHwnd, WmKeyDown, (IntPtr)VkReturn, IntPtr.Zero);
            PostMessage(senhaHwnd, WmKeyUp, (IntPtr)VkReturn, IntPtr.Zero);
        }

        Log($"[OK] Login automático do SAJMP acionado em {AppLoggerService.FormatElapsed(sw)}.");
        AppLoggerService.Success($"SAJMP iniciado e login acionado em {AppLoggerService.FormatElapsed(sw)}.", nameof(SajmpLauncherService));
        return (true, "SAJMP iniciado e login acionado com sucesso.");
    }

    private static IntPtr WaitForLoginWindow(uint processId, string windowTitleHint, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        bool loggedTitleMatchWithoutEdit = false;

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            IntPtr found = IntPtr.Zero;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;

                GetWindowThreadProcessId(hWnd, out uint winPid);
                string title = ReadWindowText(hWnd);

                if (winPid == processId || title.Contains(windowTitleHint, StringComparison.OrdinalIgnoreCase))
                {
                    var (senha, _) = FindLoginControls(hWnd);
                    if (senha != IntPtr.Zero)
                    {
                        found = hWnd;
                        return false;
                    }

                    if (!loggedTitleMatchWithoutEdit && title.Contains(windowTitleHint, StringComparison.OrdinalIgnoreCase))
                    {
                        loggedTitleMatchWithoutEdit = true;
                        AppLoggerService.Info($"Janela '{title}' detectada, mas nenhum controle Edit visível ainda.", nameof(SajmpLauncherService));
                    }
                }
                return true;
            }, IntPtr.Zero);

            if (found != IntPtr.Zero)
                return found;

            Thread.Sleep(250);
        }
        return IntPtr.Zero;
    }

    private static (IntPtr SenhaHwnd, IntPtr EntrarHwnd) FindLoginControls(IntPtr parentHwnd)
    {
        var editControls = new List<IntPtr>();
        IntPtr passwordEdit = IntPtr.Zero;
        IntPtr entrarButton = IntPtr.Zero;

        EnumChildWindows(parentHwnd, (child, _) =>
        {
            if (!IsWindowVisible(child) || !IsWindowEnabled(child))
                return true;

            string cls = ReadClassName(child);
            if (cls.Contains("Edit", StringComparison.OrdinalIgnoreCase))
            {
                int style = GetWindowLong(child, GwlStyle);
                if ((style & EsPassword) != 0 && passwordEdit == IntPtr.Zero)
                    passwordEdit = child;
                editControls.Add(child);
            }
            else
            {
                string text = ReadWindowText(child).Replace("&", string.Empty).Trim();
                if (text.Equals("Entrar", StringComparison.OrdinalIgnoreCase))
                    entrarButton = child;
            }
            return true;
        }, IntPtr.Zero);

        if (passwordEdit == IntPtr.Zero && editControls.Count >= 2)
            passwordEdit = editControls[1];

        return (passwordEdit, entrarButton);
    }

    private static string ReadWindowText(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        _ = GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string ReadClassName(IntPtr hWnd)
    {
        var sb = new StringBuilder(128);
        _ = GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}
