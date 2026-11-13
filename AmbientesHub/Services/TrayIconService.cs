using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace AmbientesHub.Services;

public sealed class TrayIconService : IDisposable
{
    private const int WM_USER = 0x0400;
    public const int WM_TRAYICON = WM_USER + 1024;
    private const uint NOTIFYICON_VERSION_4 = 4;

    private const uint NIM_ADD = 0x00000000;
    private const uint NIM_MODIFY = 0x00000001;
    private const uint NIM_DELETE = 0x00000002;
    private const uint NIM_SETVERSION = 0x00000004;

    private const uint NIF_MESSAGE = 0x00000001;
    private const uint NIF_ICON = 0x00000002;
    private const uint NIF_TIP = 0x00000004;
    private const uint NIF_INFO = 0x00000010;

    private const int WM_LBUTTONDBLCLK = 0x0203;
    private const int WM_RBUTTONUP = 0x0205;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NOTIFYICONDATA
    {
        public uint cbSize;
        public IntPtr hWnd;
        public uint uID;
        public uint uFlags;
        public uint uCallbackMessage;
        public IntPtr hIcon;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szTip;
        public uint dwState;
        public uint dwStateMask;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string szInfo;
        public uint uVersion;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string szInfoTitle;
        public uint dwInfoFlags;
        public Guid guidItem;
        public IntPtr hBalloonIcon;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern bool Shell_NotifyIcon(uint dwMessage, ref NOTIFYICONDATA lpData);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr LoadImage(IntPtr hinst, string lpszName, uint uType, int cxDesired, int cyDesired, uint fuLoad);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    private const uint IMAGE_ICON = 1;
    private const uint LR_LOADFROMFILE = 0x00000010;
    private const uint LR_DEFAULTSIZE = 0x00000040;

    private readonly Window _window;
    private IntPtr _hwnd;
    private IntPtr _hIcon = IntPtr.Zero;
    private bool _isAdded;
    private HwndSource? _hwndSource;

    public event Action? OnOpenRequested;
    public event Action? OnExitRequested;
    public event Action? OnToggleAdminModeRequested;
    public Func<bool>? IsAdminModeProvider;

    public TrayIconService(Window window)
    {
        _window = window;
    }

    public void Initialize(string iconPath, string tooltip)
    {
        var helper = new WindowInteropHelper(_window);
        _hwnd = helper.EnsureHandle();

        _hwndSource = HwndSource.FromHwnd(_hwnd);
        _hwndSource?.AddHook(WndProc);

        if (System.IO.File.Exists(iconPath))
        {
            _hIcon = LoadImage(IntPtr.Zero, iconPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE);
        }

        var data = new NOTIFYICONDATA
        {
            cbSize = (uint)Marshal.SizeOf<NOTIFYICONDATA>(),
            hWnd = _hwnd,
            uID = 1,
            uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP,
            uCallbackMessage = WM_TRAYICON,
            hIcon = _hIcon,
            szTip = tooltip.Length > 127 ? tooltip[..127] : tooltip
        };

        _isAdded = Shell_NotifyIcon(NIM_ADD, ref data);
        data.uVersion = NOTIFYICON_VERSION_4;
        Shell_NotifyIcon(NIM_SETVERSION, ref data);
    }

    public void ShowBalloon(string title, string message)
    {
        if (!_isAdded || _hwnd == IntPtr.Zero) return;

        var data = new NOTIFYICONDATA
        {
            cbSize = (uint)Marshal.SizeOf<NOTIFYICONDATA>(),
            hWnd = _hwnd,
            uID = 1,
            uFlags = NIF_INFO,
            szInfoTitle = title.Length > 63 ? title[..63] : title,
            szInfo = message.Length > 255 ? message[..255] : message,
            dwInfoFlags = 0x00000001 // NIIF_INFO
        };

        Shell_NotifyIcon(NIM_MODIFY, ref data);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_TRAYICON)
        {
            int eventId = lParam.ToInt32() & 0xFFFF;
            if (eventId == WM_LBUTTONDBLCLK)
            {
                OnOpenRequested?.Invoke();
                handled = true;
            }
            else if (eventId == WM_RBUTTONUP)
            {
                ShowContextMenu();
                handled = true;
            }
        }

        return IntPtr.Zero;
    }

    private void ShowContextMenu()
    {
        SetForegroundWindow(_hwnd);

        var menu = new System.Windows.Controls.ContextMenu();

        var itemOpen = new System.Windows.Controls.MenuItem
        {
            Header = "Abrir Ambientes Hub",
            FontWeight = FontWeights.Bold
        };
        itemOpen.Click += (s, e) => OnOpenRequested?.Invoke();

        bool isAdmin = IsAdminModeProvider?.Invoke() ?? false;
        var itemToggleAdmin = new System.Windows.Controls.MenuItem
        {
            Header = isAdmin ? "Reiniciar em Modo Padrão (Sem Admin)" : "Reiniciar como Administrador (Elevado)"
        };
        itemToggleAdmin.Click += (s, e) => OnToggleAdminModeRequested?.Invoke();

        var itemExit = new System.Windows.Controls.MenuItem
        {
            Header = "Sair do Aplicativo"
        };
        itemExit.Click += (s, e) => OnExitRequested?.Invoke();

        menu.Items.Add(itemOpen);
        menu.Items.Add(new System.Windows.Controls.Separator());
        menu.Items.Add(itemToggleAdmin);
        menu.Items.Add(new System.Windows.Controls.Separator());
        menu.Items.Add(itemExit);

        menu.IsOpen = true;
    }

    public void Remove()
    {
        if (_isAdded && _hwnd != IntPtr.Zero)
        {
            var data = new NOTIFYICONDATA
            {
                cbSize = (uint)Marshal.SizeOf<NOTIFYICONDATA>(),
                hWnd = _hwnd,
                uID = 1
            };
            Shell_NotifyIcon(NIM_DELETE, ref data);
            _isAdded = false;
        }

        if (_hIcon != IntPtr.Zero)
        {
            DestroyIcon(_hIcon);
            _hIcon = IntPtr.Zero;
        }

        if (_hwndSource != null)
        {
            _hwndSource.RemoveHook(WndProc);
            _hwndSource = null;
        }
    }

    public void Dispose()
    {
        Remove();
        GC.SuppressFinalize(this);
    }
}

