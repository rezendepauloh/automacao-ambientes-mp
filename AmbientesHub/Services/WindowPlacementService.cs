using System.IO;
using System.Text.Json;
using System.Windows;

namespace AmbientesHub.Services;

/// <summary>
/// Persiste posição, tamanho e estado da janela principal em window-state.json
/// na pasta do executável (compartilhado entre modo Usuário e Administrador).
/// </summary>
public static class WindowPlacementService
{
    private static readonly string StateFilePath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "window-state.json");

    private sealed class WindowPlacementState
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public WindowState WindowState { get; set; } = WindowState.Normal;
    }

    public static void Restore(Window window)
    {
        try
        {
            if (!File.Exists(StateFilePath)) return;

            var json = File.ReadAllText(StateFilePath);
            var state = JsonSerializer.Deserialize<WindowPlacementState>(json);
            if (state is null || state.Width < 400 || state.Height < 300) return;

            // Garante que a janela não abra fora da área virtual dos monitores conectados
            double vLeft = SystemParameters.VirtualScreenLeft;
            double vTop = SystemParameters.VirtualScreenTop;
            double vRight = vLeft + SystemParameters.VirtualScreenWidth;
            double vBottom = vTop + SystemParameters.VirtualScreenHeight;

            if (state.Left + state.Width < vLeft + 80 || state.Left > vRight - 80 ||
                state.Top + state.Height < vTop + 80 || state.Top > vBottom - 80)
                return;

            window.WindowStartupLocation = WindowStartupLocation.Manual;
            window.Left = state.Left;
            window.Top = state.Top;
            window.Width = state.Width;
            window.Height = state.Height;

            if (state.WindowState == WindowState.Maximized)
            {
                window.Loaded += (_, _) => window.WindowState = WindowState.Maximized;
            }
        }
        catch
        {
            // Ignora arquivo corrompido ou sem permissão
        }
    }

    public static void Save(Window window)
    {
        try
        {
            Rect bounds = window.WindowState == WindowState.Normal
                ? new Rect(window.Left, window.Top, window.Width, window.Height)
                : window.RestoreBounds;

            if (bounds.Width < 400 || bounds.Height < 300 ||
                double.IsNaN(bounds.Left) || double.IsNaN(bounds.Top))
                return;

            var state = new WindowPlacementState
            {
                Left = bounds.Left,
                Top = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height,
                WindowState = window.WindowState == WindowState.Maximized
                    ? WindowState.Maximized
                    : WindowState.Normal
            };

            File.WriteAllText(StateFilePath, JsonSerializer.Serialize(state));
        }
        catch
        {
            // Falha silenciosa para não interromper uso da UI
        }
    }
}
