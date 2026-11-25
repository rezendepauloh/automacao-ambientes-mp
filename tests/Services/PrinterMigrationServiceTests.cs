using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class PrinterMigrationServiceTests
{
    [Fact]
    public void ValidateRequest_AcceptsValidPrintServerAndPrinterData()
    {
        var req = new PrinterMigrationRequest
        {
            PrintServer = "srv-print01",
            PrinterName = "PRT-0001",
            NewIpAddress = "10.0.0.50",
            PortNumber = 9100,
            PortPrefix = "PAPERCUT_"
        };

        var ex = Record.Exception(() => PrinterMigrationService.ValidateRequest(req));
        Assert.Null(ex);
        Assert.Equal("PAPERCUT_10.0.0.50", req.ExpectedPortName);
    }

    [Theory]
    [InlineData("", "PRT-0001", "10.0.0.50")]
    [InlineData("srv;whoami", "PRT-0001", "10.0.0.50")]
    [InlineData("srv-print01", "PRT-0001 && calc", "10.0.0.50")]
    [InlineData("srv-print01", "PRT-0001", "10.0.0.999")]
    [InlineData("srv-print01", "PRT-0001", "hostname-em-vez-de-ip")]
    public void ValidateRequest_RejectsInvalidOrUnsafeInputs_FailClosed(
        string server,
        string printer,
        string ip)
    {
        var req = new PrinterMigrationRequest
        {
            PrintServer = server,
            PrinterName = printer,
            NewIpAddress = ip
        };

        Assert.Throws<ArgumentException>(() => PrinterMigrationService.ValidateRequest(req));
    }

    [Theory]
    [InlineData(true, "PAPERCUT_10.0.0.50", "PAPERCUT_10.0.0.50", PrinterMigrationScenario.AlreadyUpToDate)]
    [InlineData(true, "papercut_10.0.0.50", "PAPERCUT_10.0.0.50", PrinterMigrationScenario.AlreadyUpToDate)]
    [InlineData(true, "PAPERCUT_10.0.0.12", "PAPERCUT_10.0.0.50", PrinterMigrationScenario.PortNeedsUpdate)]
    [InlineData(false, "", "PAPERCUT_10.0.0.50", PrinterMigrationScenario.PrinterNotFoundPendingCreation)]
    public void EvaluateScenario_IdentifiesAllThreeWorkflowCasesCorrectly(
        bool exists,
        string currentPort,
        string expectedPort,
        PrinterMigrationScenario expected)
    {
        var actual = PrinterMigrationService.EvaluateScenario(exists, currentPort, expectedPort);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void BuildPrintServerPowerShellScript_GeneratesExpectedStepsAndPaperCutPortName()
    {
        var req = new PrinterMigrationRequest
        {
            PrintServer = "srv-print01",
            PrinterName = "PRT-0001",
            NewIpAddress = "10.0.0.50",
            PortNumber = 9100,
            PortPrefix = "PAPERCUT_",
            SendTestPage = true,
            OpenPrintQueue = true
        };

        string script = PrinterMigrationService.BuildPrintServerPowerShellScript(req, dryRunOnly: false);

        Assert.Contains("$PrinterName      = 'PRT-0001'", script);
        Assert.Contains("$ExpectedPortName = 'PAPERCUT_10.0.0.50'", script);
        Assert.Contains("$PortNumber       = 9100", script);
        Assert.Contains("PaperCut TCP/IP Port", script);
        Assert.Contains("Ensure-PaperCutTcpPort", script);
        Assert.Contains("Provision-NewPrinterQueue", script);
        Assert.Contains("Set-Printer -Name $PrinterName -PortName $ExpectedPortName", script);
        Assert.Contains("PrintTestPage", script);
    }
}
