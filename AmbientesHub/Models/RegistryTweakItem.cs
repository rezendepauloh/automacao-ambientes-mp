using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.Win32;

namespace AmbientesHub.Models;

public enum RegistryTweakStatus
{
    NotChecked,
    Optimized,
    Default,
    Custom,
    Missing
}

public class RegistryTweakItem : INotifyPropertyChanged
{
    private bool _isSelected = true;
    private RegistryTweakStatus _status = RegistryTweakStatus.NotChecked;
    private string _currentValueDisplay = "Não verificado";
    private object? _currentRawValue;
    private object? _previousRawValue;
    private bool _isApplying;

    public required string Id { get; init; }
    public required string Title { get; init; }
    public required string Description { get; init; }
    public required string Category { get; init; }
    public required RegistryHive Hive { get; init; }
    public required string SubKeyPath { get; init; }
    public required string ValueName { get; init; }
    public required RegistryValueKind ValueKind { get; init; }
    public required object DefaultValue { get; init; }
    public required object OptimizedValue { get; init; }
    public bool RequiresAdmin { get; init; }
    public string Emoji { get; init; } = "⚡";

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected != value)
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }
    }

    public RegistryTweakStatus Status
    {
        get => _status;
        set
        {
            if (_status != value)
            {
                _status = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(StatusBadgeText));
                OnPropertyChanged(nameof(StatusBadgeBackground));
                OnPropertyChanged(nameof(StatusBadgeForeground));
                OnPropertyChanged(nameof(IsOptimized));
            }
        }
    }

    public string CurrentValueDisplay
    {
        get => _currentValueDisplay;
        set
        {
            if (_currentValueDisplay != value)
            {
                _currentValueDisplay = value;
                OnPropertyChanged();
            }
        }
    }

    public object? CurrentRawValue
    {
        get => _currentRawValue;
        set => _currentRawValue = value;
    }

    public object? PreviousRawValue
    {
        get => _previousRawValue;
        set
        {
            _previousRawValue = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasRollback));
        }
    }

    public bool IsApplying
    {
        get => _isApplying;
        set
        {
            if (_isApplying != value)
            {
                _isApplying = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsOptimized => Status == RegistryTweakStatus.Optimized;
    public bool HasRollback => PreviousRawValue != null;

    public string FullRegistryPath => $"{(Hive == RegistryHive.CurrentUser ? "HKCU" : "HKLM")}\\{SubKeyPath}\\{ValueName}";
    public string RecommendedValueDisplay => OptimizedValue?.ToString() ?? "(vazio)";
    public string DefaultValueDisplay => DefaultValue?.ToString() ?? "(padrão)";

    public string PrivilegeTag => RequiresAdmin ? "HKLM • Requer Admin" : "HKCU • Modo Usuário";
    public string PrivilegeBadgeBackground => RequiresAdmin ? "#3b1e54" : "#1e3a5f";
    public string PrivilegeBadgeForeground => RequiresAdmin ? "#d8b4fe" : "#7dd3fc";

    public string StatusBadgeText => Status switch
    {
        RegistryTweakStatus.Optimized => "✔ Otimizado",
        RegistryTweakStatus.Default => "Padrão Windows",
        RegistryTweakStatus.Custom => "⚠️ Personalizado",
        RegistryTweakStatus.Missing => "⚪ Ausente (Padrão)",
        _ => "Aguardando Leitura"
    };

    public string StatusBadgeBackground => Status switch
    {
        RegistryTweakStatus.Optimized => "#064e3b",
        RegistryTweakStatus.Default => "#1e293b",
        RegistryTweakStatus.Custom => "#451a03",
        RegistryTweakStatus.Missing => "#1e293b",
        _ => "#1e293b"
    };

    public string StatusBadgeForeground => Status switch
    {
        RegistryTweakStatus.Optimized => "#34d399",
        RegistryTweakStatus.Default => "#94a3b8",
        RegistryTweakStatus.Custom => "#fbbf24",
        RegistryTweakStatus.Missing => "#94a3b8",
        _ => "#94a3b8"
    };

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
