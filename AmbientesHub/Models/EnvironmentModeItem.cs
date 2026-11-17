using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace AmbientesHub.Models;

public class EnvironmentStepItem : INotifyPropertyChanged
{
    private bool _isEnabled = true;

    public string Id { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public string Category { get; set; } = "Execução";

    public bool IsEnabled
    {
        get => _isEnabled;
        set { _isEnabled = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? prop = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
}

public class EnvironmentModeItem : INotifyPropertyChanged
{
    private bool _isExpanded;
    private bool _isRunning;
    private string _consoleLog = string.Empty;

    public string ModeKey { get; set; } = string.Empty;
    public string Icon { get; set; } = "🛠️";
    public string Title { get; set; } = string.Empty;
    public string Subtitle { get; set; } = string.Empty;
    public string BadgeText { get; set; } = "MODO";
    public ObservableCollection<EnvironmentStepItem> Steps { get; set; } = [];

    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            _isExpanded = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(DetailsVisibility));
            OnPropertyChanged(nameof(ChevronSymbol));
        }
    }

    public bool IsRunning
    {
        get => _isRunning;
        set
        {
            _isRunning = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(CanRun));
            OnPropertyChanged(nameof(RunButtonText));
        }
    }

    public string ConsoleLog
    {
        get => _consoleLog;
        set
        {
            _consoleLog = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ConsoleVisibility));
        }
    }

    public Visibility DetailsVisibility => IsExpanded ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ConsoleVisibility => string.IsNullOrWhiteSpace(ConsoleLog) ? Visibility.Collapsed : Visibility.Visible;
    public string ChevronSymbol => IsExpanded ? "⌃" : "⌄";
    public bool CanRun => !IsRunning;
    public string RunButtonText => IsRunning ? "Executando..." : "Executar";

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? prop = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
}
