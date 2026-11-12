using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace AmbientesHub.Services;

public interface IStartupManager
{
    bool IsRunAtStartupEnabled();
    bool SetRunAtStartup(bool enable, string? customExecutablePath = null);
    bool IsStartMenuShortcutCreated();
    bool CreateStartMenuShortcut(string? customExecutablePath = null);
    bool RemoveStartMenuShortcut();
}

public class StartupService : IStartupManager
{
    public const string StartupRegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    public const string AppName = "AmbientesHub";
    public const string ShortcutDisplayName = "Ambientes Hub.lnk";

    public virtual string GetExecutablePath()
    {
        string? procPath = Process.GetCurrentProcess().MainModule?.FileName;
        if (!string.IsNullOrWhiteSpace(procPath) && procPath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            return procPath;
        }

        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string candidateExe = Path.Combine(baseDir, "AmbientesHub.exe");
        if (File.Exists(candidateExe))
        {
            return candidateExe;
        }

        return procPath ?? candidateExe;
    }

    public virtual string GetStartMenuShortcutPath()
    {
        string startMenuPrograms = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
        return Path.Combine(startMenuPrograms, ShortcutDisplayName);
    }

    public bool IsRunAtStartupEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKeyPath, writable: false);
            if (key == null) return false;

            object? val = key.GetValue(AppName);
            return val != null && !string.IsNullOrWhiteSpace(val.ToString());
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Erro ao verificar inicialização com Windows: {ex.Message}", nameof(StartupService));
            return false;
        }
    }

    public bool SetRunAtStartup(bool enable, string? customExecutablePath = null)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKeyPath, writable: true);
            if (key == null)
            {
                AppLoggerService.Warning("Chave de registro HKCU Run não encontrada.", nameof(StartupService));
                return false;
            }

            if (enable)
            {
                string exePath = customExecutablePath ?? GetExecutablePath();
                string formattedCmd = $"\"{exePath}\"";
                key.SetValue(AppName, formattedCmd, RegistryValueKind.String);
                AppLoggerService.Info($"Inicialização com Windows habilitada: {formattedCmd}", nameof(StartupService));
            }
            else
            {
                if (key.GetValue(AppName) != null)
                {
                    key.DeleteValue(AppName, throwOnMissingValue: false);
                    AppLoggerService.Info("Inicialização com Windows desabilitada.", nameof(StartupService));
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao alterar inicialização com Windows: {ex.Message}", nameof(StartupService));
            return false;
        }
    }

    public bool IsStartMenuShortcutCreated()
    {
        try
        {
            return File.Exists(GetStartMenuShortcutPath());
        }
        catch
        {
            return false;
        }
    }

    public bool CreateStartMenuShortcut(string? customExecutablePath = null)
    {
        try
        {
            string shortcutPath = GetStartMenuShortcutPath();
            string exePath = customExecutablePath ?? GetExecutablePath();
            string iconPath = Path.Combine(Path.GetDirectoryName(exePath) ?? "", "Assets", "app.ico");
            if (!File.Exists(iconPath))
            {
                iconPath = exePath;
            }

            var link = (IShellLinkW)new ShellLink();
            link.SetPath(exePath);
            link.SetWorkingDirectory(Path.GetDirectoryName(exePath) ?? "");
            link.SetDescription("Ambientes Hub — Central de Produtividade & Performance");
            link.SetIconLocation(iconPath, 0);

            var file = (IPersistFile)link;
            file.Save(shortcutPath, false);

            Marshal.ReleaseComObject(file);
            Marshal.ReleaseComObject(link);

            AppLoggerService.Info($"Atalho criado no Menu Iniciar: {shortcutPath}", nameof(StartupService));
            return true;
        }
        catch (Exception ex)
        {
            AppLoggerService.Danger($"Erro ao criar atalho no Menu Iniciar: {ex.Message}", nameof(StartupService));
            return false;
        }
    }

    public bool RemoveStartMenuShortcut()
    {
        try
        {
            string shortcutPath = GetStartMenuShortcutPath();
            if (File.Exists(shortcutPath))
            {
                File.Delete(shortcutPath);
                AppLoggerService.Info($"Atalho removido do Menu Iniciar: {shortcutPath}", nameof(StartupService));
            }
            return true;
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Erro ao remover atalho do Menu Iniciar: {ex.Message}", nameof(StartupService));
            return false;
        }
    }

    #region Win32 IShellLink COM Definition
    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink { }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    private interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszName, int cchMaxName);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszDir, int cchMaxPath);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszArgs, int cchMaxPath);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }
    #endregion
}
