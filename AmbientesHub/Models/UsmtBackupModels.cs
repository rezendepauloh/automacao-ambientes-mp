namespace AmbientesHub.Models;

public class UsmtBackupOptions
{
    public string TargetComputer { get; set; } = ".";
    public string StorePath { get; set; } = string.Empty;
    public List<string> SelectedUsers { get; set; } = [];
    public bool IncludeMigApp { get; set; } = true;
    public bool IncludeMigDocs { get; set; } = true;
    public bool IncludeMigUser { get; set; } = false;
    public bool Compress { get; set; } = true;
    public bool CreateLocalAccountsOnRestore { get; set; } = false;
    public string? EncryptionKey { get; set; }
    public string? CustomUsmtPath { get; set; }
}

public class UsmtOperationResult
{
    public bool Success { get; set; }
    public string ComputerName { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string StorePath { get; set; } = string.Empty;
    public string LogFilePath { get; set; } = string.Empty;
    public int ExitCode { get; set; } = -1;
    public int ProcessedUsersCount { get; set; }
    public string Message { get; set; } = string.Empty;
}
