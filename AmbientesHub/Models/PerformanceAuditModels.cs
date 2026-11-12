namespace AmbientesHub.Models;

public record DiskStatusItem(string DriveName, double FreeGb, double UsedGb, double TotalGb)
{
    public string Summary => $"{DriveName}  |  Livre: {FreeGb:F2} GB  |  Usado: {UsedGb:F2} GB  |  Total: {TotalGb:F2} GB";
}

public record ProcessMemoryItem(string ProcessName, double TotalRamMb, int InstanceCount);

public record TelemetryServiceStatus(string Name, string DisplayName, string State, string StartMode);

public record StartupAppItem(string Origin, string Name, string Command);

public record VhdxFileItem(string FullPath, double SizeGb);

public class SystemAuditReport
{
    public bool IsAdministrator { get; set; }
    public List<DiskStatusItem> Disks { get; set; } = [];
    public double TotalRamGb { get; set; }
    public double FreeRamGb { get; set; }
    public double FreeRamPercent { get; set; }
    public List<ProcessMemoryItem> TopProcesses { get; set; } = [];
    public bool IsSearchHostRunning { get; set; }
    public bool HiberfilExists { get; set; }
    public double HiberfilSizeGb { get; set; }
    public double SoftwareDistributionDownloadMb { get; set; }
    public string ActivePowerPlan { get; set; } = "Não identificado";
    public bool? IsTrimEnabled { get; set; }
    public string TrimStatusText { get; set; } = "Não verificado";
    public List<TelemetryServiceStatus> TelemetryServices { get; set; } = [];
    public List<string> DetectedBloatApps { get; set; } = [];
    public List<StartupAppItem> StartupItems { get; set; } = [];
    public string WslStatusOutput { get; set; } = string.Empty;
    public List<VhdxFileItem> VhdxFiles { get; set; } = [];
    public bool IsDockerAvailable { get; set; }
    public string DockerStatusOutput { get; set; } = string.Empty;
}

public class OptimizationOptions
{
    public bool DisableTelemetryServices { get; set; } = true;
    public bool ApplyRegistryTelemetryTweaks { get; set; } = true;
    public bool RemoveBloatwareApps { get; set; } = true;
    public bool DisableHibernation { get; set; } = true;
    public bool EnableSsdTrim { get; set; } = true;
    public bool DisableBingSearch { get; set; } = true;
    public bool AccelerateUiMenus { get; set; } = true;
    public bool ConfigureWslLimit { get; set; } = true;
    public string WslMemoryLimit { get; set; } = "4GB";
    public bool CleanDockerBuilderCache { get; set; } = false;
}

