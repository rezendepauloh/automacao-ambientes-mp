namespace AmbientesHub.Models;

public class CleanupOptions
{
    public bool CloseOffice { get; set; } = true;
    public bool CloseBrowsers { get; set; } = true; // Chrome, Edge
    public bool CloseFirefox { get; set; } = false; // Manter Firefox aberto ou fechar
    public bool CloseTeams { get; set; } = true;
    public bool CloseAdministrativeTools { get; set; } = true; // MMC, SCCM, SAJMP
    public bool CloseUtilities { get; set; } = true; // Notepad, FoxitPDF, Calculadora
    public bool StopTelemetry { get; set; } = true;
}

