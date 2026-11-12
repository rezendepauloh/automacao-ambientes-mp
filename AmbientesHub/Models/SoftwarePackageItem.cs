using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace AmbientesHub.Models;

public enum SoftwareInstallMethod
{
    RsatAdConsole,
    SccmConsole,
    Winget,
    DirectMsi
}

public class SoftwarePackageItem : INotifyPropertyChanged
{
    private bool _isInstalled;
    private bool _isInstalling;
    private bool _hasUpdate;
    private string _statusDetail = string.Empty;

    public required string Id { get; init; }
    public required string Name { get; init; }
    public required string Category { get; init; }
    public required string Description { get; init; }
    public required SoftwareInstallMethod Method { get; init; }
    public string? WingetId { get; init; }
    public string? DownloadUrl { get; init; }
    public string? LocalInstallPath { get; init; }
    public string? UninstallMatchName { get; init; }
    public string IconEmoji { get; init; } = "📦";
    public bool RequiresAdmin { get; init; } = true;

    public bool IsInstalled
    {
        get => _isInstalled;
        set
        {
            if (_isInstalled != value)
            {
                _isInstalled = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(InstalledBadgeVisibility));
                OnPropertyChanged(nameof(InstallButtonVisibility));
                OnPropertyChanged(nameof(UninstallButtonVisibility));
            }
        }
    }

    public bool HasUpdate
    {
        get => _hasUpdate;
        set
        {
            if (_hasUpdate != value)
            {
                _hasUpdate = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(UpdateBadgeVisibility));
                OnPropertyChanged(nameof(UpdateButtonVisibility));
            }
        }
    }

    public bool IsInstalling
    {
        get => _isInstalling;
        set
        {
            if (_isInstalling != value)
            {
                _isInstalling = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanClickInstall));
                OnPropertyChanged(nameof(InstallButtonText));
            }
        }
    }

    public string StatusDetail
    {
        get => _statusDetail;
        set
        {
            if (_statusDetail != value)
            {
                _statusDetail = value;
                OnPropertyChanged();
            }
        }
    }

    public Visibility InstalledBadgeVisibility => IsInstalled && !HasUpdate ? Visibility.Visible : Visibility.Collapsed;
    public Visibility UpdateBadgeVisibility => IsInstalled && HasUpdate ? Visibility.Visible : Visibility.Collapsed;
    public Visibility InstallButtonVisibility => !IsInstalled ? Visibility.Visible : Visibility.Collapsed;
    public Visibility UpdateButtonVisibility => IsInstalled && HasUpdate ? Visibility.Visible : Visibility.Collapsed;
    public Visibility UninstallButtonVisibility => IsInstalled ? Visibility.Visible : Visibility.Collapsed;
    public bool CanClickInstall => !IsInstalling;
    public string InstallButtonText => IsInstalling ? "Instalando..." : "Instalar";

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
