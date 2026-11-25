namespace AmbientesHub.Models;

public enum PrinterMigrationScenario
{
    Unknown = 0,
    AlreadyUpToDate = 1,
    PortNeedsUpdate = 2,
    PrinterNotFoundPendingCreation = 3,
    Error = 4
}

public class PrinterMigrationRequest
{
    public string PrintServer { get; set; } = string.Empty;
    public string PrinterName { get; set; } = string.Empty;
    public string NewIpAddress { get; set; } = string.Empty;
    public int PortNumber { get; set; } = 9100;
    public string PortPrefix { get; set; } = "PAPERCUT_";
    public string PortMonitorName { get; set; } = "PaperCut TCP/IP Port";
    public string DriverName { get; set; } = "HP Universal Printing PCL 6";
    public string Location { get; set; } = string.Empty;
    public string Comment { get; set; } = string.Empty;
    public string MacAddress { get; set; } = string.Empty;
    public bool SendTestPage { get; set; } = true;
    public bool OpenPrintQueue { get; set; } = true;

    public string ExpectedPortName => string.IsNullOrWhiteSpace(NewIpAddress)
        ? string.Empty
        : $"{PortPrefix.Trim()}{NewIpAddress.Trim()}";
}

public class PrinterNetworkCheckResult
{
    public string IpAddress { get; set; } = string.Empty;
    public bool PingSuccess { get; set; }
    public long? PingRttMs { get; set; }
    public bool RawPort9100Open { get; set; }
    public bool WebInterfaceOpen { get; set; }
    public string MacAddressArp { get; set; } = string.Empty;

    public bool IsReadyForInstall => PingSuccess || RawPort9100Open;
}

public class PrinterMigrationOutcome
{
    public PrinterMigrationScenario Scenario { get; set; } = PrinterMigrationScenario.Unknown;
    public bool Success { get; set; }
    public string PrinterName { get; set; } = string.Empty;
    public string CurrentPortName { get; set; } = string.Empty;
    public string TargetPortName { get; set; } = string.Empty;
    public bool PortCreated { get; set; }
    public bool PortUpdated { get; set; }
    public string Summary { get; set; } = string.Empty;
    public List<string> Logs { get; set; } = [];
}
