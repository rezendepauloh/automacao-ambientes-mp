using AmbientesHub.Models;
using AmbientesHub.Services;
using WForms = System.Windows.Forms;

namespace AmbientesHub.Controls;

/// <summary>
/// Wrapper WinForms AxHost para o controle ActiveX nativo do Windows RDP (mstscax.dll),
/// com perfis de qualidade/estabilidade, adaptação dinâmica de resolução ao monitor/painel e tela cheia nativa.
/// </summary>
public sealed class EmbeddedRdpHost : WForms.AxHost
{
    private static readonly string[] CandidateProgIds =
    [
        "MsTscAx.MsTscAx.12",
        "MsTscAx.MsTscAx.11",
        "MsTscAx.MsTscAx.10",
        "MsTscAx.MsTscAx.9",
        "MsTscAx.MsTscAx.8",
        "MsTscAx.MsTscAx.7",
        "MsTscAx.MsTscAx",
        "MsRdpClient.MsRdpClient"
    ];

    private static string ResolveRdpClsid()
    {
        foreach (string progId in CandidateProgIds)
        {
            var type = Type.GetTypeFromProgID(progId);
            if (type != null && type.GUID != Guid.Empty)
                return type.GUID.ToString();
        }
        return "a3bc03a0-041d-42e3-ad22-882b7865c9c5";
    }

    private readonly WForms.Timer _resizeDebounceTimer;
    private RdpServerEntry? _currentEntry;
    private RdpServerEntry? _pendingEntry;
    private int _pendingWidth;
    private int _pendingHeight;
    private bool _autoAdaptResolution = true;
    private bool _smartSizing = true;

    public EmbeddedRdpHost() : base(ResolveRdpClsid())
    {
        _resizeDebounceTimer = new WForms.Timer { Interval = 400 };
        _resizeDebounceTimer.Tick += (_, _) =>
        {
            _resizeDebounceTimer.Stop();
            if (_autoAdaptResolution)
                AdaptResolutionToContainer();
        };

        HandleCreated += (_, _) =>
        {
            if (_pendingEntry != null)
            {
                var entry = _pendingEntry;
                int w = _pendingWidth;
                int h = _pendingHeight;
                _pendingEntry = null;
                BeginInvoke(() => ConnectCore(entry, w, h));
            }
        };

        Resize += (_, _) =>
        {
            if (_autoAdaptResolution && IsHandleCreated)
            {
                _resizeDebounceTimer.Stop();
                _resizeDebounceTimer.Start();
            }
        };
    }

    public bool AutoAdaptResolution
    {
        get => _autoAdaptResolution;
        set
        {
            _autoAdaptResolution = value;
            if (value)
                AdaptResolutionToContainer();
        }
    }

    public void SetSmartSizing(bool enabled)
    {
        _smartSizing = enabled;
        try
        {
            object? raw = GetOcx();
            if (raw == null) return;
            dynamic ocx = raw;
            dynamic? adv = GetBestAdvancedSettings(ocx);
            if (adv != null)
                adv.SmartSizing = enabled;
        }
        catch { }
    }

    public void Connect(RdpServerEntry entry, int fallbackWidth, int fallbackHeight)
    {
        _currentEntry = entry;
        _autoAdaptResolution = entry.AutoAdaptResolution;

        if (!IsHandleCreated)
        {
            _pendingEntry = entry;
            _pendingWidth = fallbackWidth;
            _pendingHeight = fallbackHeight;
            return;
        }

        ConnectCore(entry, fallbackWidth, fallbackHeight);
    }

    public void Reconnect(RdpServerEntry updatedEntry)
    {
        _currentEntry = updatedEntry;
        _autoAdaptResolution = updatedEntry.AutoAdaptResolution;
        SafeDisconnect();

        var reconnectTimer = new WForms.Timer { Interval = 350 };
        reconnectTimer.Tick += (_, _) =>
        {
            reconnectTimer.Stop();
            reconnectTimer.Dispose();
            if (!IsDisposed)
                ConnectCore(updatedEntry, Width, Height);
        };
        reconnectTimer.Start();
    }

    /// <summary>
    /// Adapta a resolução remota e a escala DPI ao tamanho atual do painel ou do monitor onde a janela está.
    /// </summary>
    public bool AdaptResolutionToContainer(bool useFullMonitorBounds = false)
    {
        try
        {
            object? raw = GetOcx();
            if (raw == null) return false;
            dynamic ocx = raw;
            if (ocx.Connected == 0) return false;

            var screen = WForms.Screen.FromControl(this);
            int targetW = useFullMonitorBounds
                ? screen.Bounds.Width
                : (Width >= 640 ? Width : screen.WorkingArea.Width);
            int targetH = useFullMonitorBounds
                ? screen.Bounds.Height
                : (Height >= 480 ? Height : screen.WorkingArea.Height);

            uint w = (uint)(Math.Clamp(targetW, 640, 7680) & ~1);
            uint h = (uint)Math.Clamp(targetH, 480, 4320);
            uint desktopScale = (uint)Math.Clamp((int)Math.Round(DeviceDpi * 100.0 / 96.0), 100, 500);

            ocx.UpdateSessionDisplaySettings(w, h, w, h, 0u, desktopScale, 100u);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Entra em tela cheia nativa do RDP adaptada exatamente ao monitor onde o AmbientesHub está aberto.
    /// </summary>
    public void EnterMonitorFullScreen()
    {
        try
        {
            object? raw = GetOcx();
            if (raw == null) return;
            dynamic ocx = raw;

            AdaptResolutionToContainer(useFullMonitorBounds: true);

            try
            {
                dynamic? adv = GetBestAdvancedSettings(ocx);
                if (adv != null)
                    adv.ContainerHandledFullScreen = 0;
            }
            catch { }

            if (_currentEntry != null)
                ocx.FullScreenTitle = $"{_currentEntry.Name} ({_currentEntry.SubtitleDisplay}) — Ambientes Hub";

            ocx.FullScreen = true;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Não foi possível alternar para Tela Cheia RDP: {ex.Message}", nameof(EmbeddedRdpHost));
        }
    }

    private void ConnectCore(RdpServerEntry entry, int fallbackWidth, int fallbackHeight)
    {
        try
        {
            string safeHost = NetworkPathService.ValidateTarget(entry.HostOrUrl);
            object? raw = GetOcx();
            if (raw == null)
                throw new InvalidOperationException("Controle ActiveX mstscax.dll não inicializado.");

            var screen = WForms.Screen.FromControl(this);
            int initW = entry.AutoAdaptResolution
                ? (Width >= 640 ? Width : screen.WorkingArea.Width)
                : (fallbackWidth >= 640 ? fallbackWidth : screen.Bounds.Width);
            int initH = entry.AutoAdaptResolution
                ? (Height >= 480 ? Height : screen.WorkingArea.Height)
                : (fallbackHeight >= 480 ? fallbackHeight : screen.Bounds.Height);

            var qp = RdpServerEntry.GetProfileParameters(entry.QualityProfile);

            dynamic ocx = raw;
            ocx.Server = safeHost;
            if (!string.IsNullOrWhiteSpace(entry.UserName))
                ocx.UserName = entry.UserName.Trim();
            if (!string.IsNullOrWhiteSpace(entry.Domain))
                ocx.Domain = entry.Domain.Trim();

            ocx.ColorDepth = qp.ColorDepth;
            ocx.DesktopWidth = Math.Clamp(initW & ~1, 800, 7680);
            ocx.DesktopHeight = Math.Clamp(initH, 600, 4320);

            if (_currentEntry != null)
                ocx.FullScreenTitle = $"{entry.Name} ({entry.SubtitleDisplay}) — Ambientes Hub";

            dynamic? adv = GetBestAdvancedSettings(ocx);
            if (adv != null)
            {
                adv.RDPPort = Math.Clamp(entry.Port, 1, 65535);
                adv.EnableCredSspSupport = true;
                adv.AuthenticationLevel = (uint)Math.Clamp(entry.AuthenticationLevel, 0, 2);
                adv.RedirectDrives = entry.UsesHardDrives;
                adv.SmartSizing = _smartSizing;
                adv.Compress = qp.Compression ? 1 : 0;
                adv.BitmapPeristence = qp.BitmapCaching ? 1 : 0;
                adv.PerformanceFlags = qp.PerformanceFlags;

                try { adv.NetworkConnectionType = qp.NetworkConnectionType; } catch { }
                try { adv.EnableAutoReconnect = true; } catch { }
                try { adv.MaxReconnectAttempts = 10; } catch { }
            }

            ocx.Connect();
            AppLoggerService.Info(
                $"Sessão RDP embutida iniciada para '{entry.Name}' ({safeHost}, Perfil={entry.QualityProfile}, {ocx.DesktopWidth}x{ocx.DesktopHeight}, Monitor={screen.DeviceName}).",
                nameof(EmbeddedRdpHost));
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao iniciar RDP embutido para '{entry.Name}': {ex.Message}", nameof(EmbeddedRdpHost));
            throw;
        }
    }

    private static dynamic? GetBestAdvancedSettings(dynamic ocx)
    {
        try { return ocx.AdvancedSettings9; } catch { }
        try { return ocx.AdvancedSettings8; } catch { }
        try { return ocx.AdvancedSettings7; } catch { }
        try { return ocx.AdvancedSettings2; } catch { }
        return null;
    }

    public void SafeDisconnect()
    {
        try
        {
            _resizeDebounceTimer.Stop();
            object? raw = GetOcx();
            if (raw == null) return;
            dynamic ocx = raw;
            if (ocx.Connected != 0)
                ocx.Disconnect();
        }
        catch { }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _resizeDebounceTimer.Stop();
            _resizeDebounceTimer.Dispose();
        }
        base.Dispose(disposing);
    }
}
