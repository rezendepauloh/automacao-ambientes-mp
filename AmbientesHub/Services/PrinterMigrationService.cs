using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço de automação para verificação e atualização de portas PaperCut TCP/IP Port
/// no Windows Print Management (Print Server), com pré-check de rede (ICMP + 9100 RAW),
/// avaliação dos cenários de migração e geração/execução de script PowerShell sanitizado.
/// </summary>
public class PrinterMigrationService
{
    private static readonly Regex PrinterQueueNameRegex = new(
        @"^[A-Za-z0-9][A-Za-z0-9._\- ]{0,63}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant,
        TimeSpan.FromMilliseconds(200));

    private static readonly Regex PortPrefixRegex = new(
        @"^[A-Za-z0-9_\-]{1,24}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant,
        TimeSpan.FromMilliseconds(200));

    /// <summary>
    /// Valida todos os campos da requisição em modo fail-closed.
    /// </summary>
    public static void ValidateRequest(PrinterMigrationRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        string server = (request.PrintServer ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(server) || (!server.Equals(".", StringComparison.Ordinal) && !InputValidator.IsValidHostName(server)))
            throw new ArgumentException("Servidor de impressão inválido. Informe o hostname (ex: srv-print01) ou '.' para local.");

        string printer = (request.PrinterName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(printer) || !PrinterQueueNameRegex.IsMatch(printer))
            throw new ArgumentException("Nome da impressora inválido. Use apenas letras, números, hífen, ponto ou sublinhado.");

        string ip = (request.NewIpAddress ?? string.Empty).Trim();
        if (!NetworkPathService.IsStrictIpAddress(ip, out var parsed) || parsed == null || parsed.AddressFamily != AddressFamily.InterNetwork)
            throw new ArgumentException("Novo endereço IP inválido. Informe um IPv4 válido (ex: 10.0.0.50).");

        if (request.PortNumber is < 1 or > 65535)
            throw new ArgumentException("Número da porta TCP deve estar entre 1 e 65535 (padrão 9100).");

        string prefix = (request.PortPrefix ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(prefix) || !PortPrefixRegex.IsMatch(prefix))
            throw new ArgumentException("Prefixo da porta inválido (padrão: PAPERCUT_).");
    }

    /// <summary>
    /// Avalia qual cenário do cronograma deve ser aplicado com base no estado atual da impressora no Print Server.
    /// </summary>
    public static PrinterMigrationScenario EvaluateScenario(
        bool printerExistsOnServer,
        string? currentPortName,
        string expectedPortName)
    {
        if (!printerExistsOnServer)
            return PrinterMigrationScenario.PrinterNotFoundPendingCreation;

        if (!string.IsNullOrWhiteSpace(currentPortName) &&
            string.Equals(currentPortName.Trim(), expectedPortName.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return PrinterMigrationScenario.AlreadyUpToDate;
        }

        return PrinterMigrationScenario.PortNeedsUpdate;
    }

    /// <summary>
    /// Verifica na rede se o novo IP da impressora responde a Ping, porta 9100 (JetDirect RAW) e Interface Web (80/443).
    /// </summary>
    public async Task<PrinterNetworkCheckResult> CheckPrinterNetworkAsync(
        string rawIpAddress,
        int rawPort = 9100,
        CancellationToken cancellationToken = default)
    {
        if (!NetworkPathService.IsStrictIpAddress(rawIpAddress, out var parsed) || parsed == null || parsed.AddressFamily != AddressFamily.InterNetwork)
            throw new ArgumentException("Endereço IPv4 da impressora inválido.");

        string ip = parsed.ToString();
        int port = Math.Clamp(rawPort, 1, 65535);
        var netSvc = new NetworkPathService();

        var pingTask = netSvc.PingOnceAsync(ip, 1200, cancellationToken);
        var port9100Task = CheckTcpPortOpenAsync(ip, port, 1000, cancellationToken);
        var port80Task = CheckTcpPortOpenAsync(ip, 80, 900, cancellationToken);
        var port443Task = CheckTcpPortOpenAsync(ip, 443, 900, cancellationToken);

        await Task.WhenAll(pingTask, port9100Task, port80Task, port443Task);
        var pingRes = await pingTask;

        return new PrinterNetworkCheckResult
        {
            IpAddress = ip,
            PingSuccess = pingRes.Success,
            PingRttMs = pingRes.RttMs,
            RawPort9100Open = await port9100Task,
            WebInterfaceOpen = await port80Task || await port443Task,
            MacAddressArp = NetworkPathService.TryResolveMacViaArp(ip)
        };
    }

    /// <summary>
    /// Gera o script PowerShell completo e sanitizado que executa os passos 3 a 10 no Print Server:
    /// - Verifica se a impressora já existe e qual sua porta atual;
    /// - Se já estiver com PAPERCUT_IP, para e oferece página de teste + abertura da fila (Passo 4);
    /// - Se estiver com IP antigo, cria a porta PaperCut TCP/IP Port (PAPERCUT_IP : 9100) caso não exista e vincula na impressora (Passos 5 a 10);
    /// - Deixa preparada a função Provision-NewPrinterQueue para quando a impressora ainda não existir no servidor (Passo 9).
    /// </summary>
    public static string BuildPrintServerPowerShellScript(PrinterMigrationRequest request, bool dryRunOnly = false)
    {
        ValidateRequest(request);

        string qServer = InputValidator.EscapePowerShellSingleQuoted(request.PrintServer.Trim());
        string qPrinter = InputValidator.EscapePowerShellSingleQuoted(request.PrinterName.Trim());
        string qIp = InputValidator.EscapePowerShellSingleQuoted(request.NewIpAddress.Trim());
        string qExpectedPort = InputValidator.EscapePowerShellSingleQuoted(request.ExpectedPortName);
        string qMonitor = InputValidator.EscapePowerShellSingleQuoted(
            string.IsNullOrWhiteSpace(request.PortMonitorName) ? "PaperCut TCP/IP Port" : request.PortMonitorName.Trim());
        string qDriver = InputValidator.EscapePowerShellSingleQuoted(
            string.IsNullOrWhiteSpace(request.DriverName) ? "HP Universal Printing PCL 6" : request.DriverName.Trim());
        string qLocation = InputValidator.EscapePowerShellSingleQuoted(request.Location.Trim());
        string qComment = InputValidator.EscapePowerShellSingleQuoted(request.Comment.Trim());
        int portNum = request.PortNumber;
        string sendTestBool = request.SendTestPage ? "$true" : "$false";
        string openQueueBool = request.OpenPrintQueue ? "$true" : "$false";
        string dryRunBool = dryRunOnly ? "$true" : "$false";

        return $$"""
# ==============================================================================
# AUTOMAÇÃO DE PRINT SERVER (PAPERCUT TCP/IP PORT) — AMBIENTES HUB
# Servidor Alvo : {{qServer}}
# Impressora    : {{qPrinter}}
# Novo IP       : {{qIp}} (Porta esperada: {{qExpectedPort}} / TCP {{portNum}})
# ==============================================================================
$ErrorActionPreference = 'Stop'
$TargetServer     = '{{qServer}}'
$PrinterName      = '{{qPrinter}}'
$NewIp            = '{{qIp}}'
$ExpectedPortName = '{{qExpectedPort}}'
$PortNumber       = {{portNum}}
$MonitorName      = '{{qMonitor}}'
$DriverName       = '{{qDriver}}'
$Location         = '{{qLocation}}'
$Comment          = '{{qComment}}'
$SendTestPage     = {{sendTestBool}}
$OpenPrintQueue   = {{openQueueBool}}
$DryRun           = {{dryRunBool}}

# Função completa para criação de impressora ainda inexistente no servidor (Passos 01 a 07)
function Provision-NewPrinterQueue {
    param([string]$Name, [string]$Port, [string]$Driver, [string]$Loc, [string]$Desc)

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "[ERRO DE PRIVILÉGIO] Para cadastrar uma nova impressora e criar a porta PaperCut, abra o PowerShell como Administrador e cole o script novamente." -ForegroundColor Red
        return
    }

    if ($DryRun) {
        Write-Host "[MODO SIMULAÇÃO] Impressora '$Name' seria criada na porta '$Port' com driver '$Driver' (Location='$Loc')." -ForegroundColor Cyan
        return
    }

    Ensure-PaperCutTcpPort -PortName $Port -IpAddress $NewIp -TcpPort $PortNumber -Monitor $MonitorName | Out-Null

    $drv = Get-PrinterDriver -Name $Driver -ErrorAction SilentlyContinue
    if ($null -eq $drv) {
        throw "Driver '$Driver' não encontrado no servidor. Instale o driver primeiro ou ajuste o campo 'Modelo / Driver Padrão'."
    }

    Write-Host "  [PASSOS 05-07] Criando fila '$Name' (Shared=$true, Publish=No, Driver='$Driver', Porta='$Port')..." -ForegroundColor Cyan
    $addParams = @{
        Name       = $Name
        DriverName = $Driver
        PortName   = $Port
        Shared     = $true
        ShareName  = $Name
        Published  = $false
    }
    if (-not [string]::IsNullOrWhiteSpace($Loc))  { $addParams['Location'] = $Loc }
    if (-not [string]::IsNullOrWhiteSpace($Desc)) { $addParams['Comment']  = $Desc }

    Add-Printer @addParams
    Write-Host "  [SUCESSO] Impressora '$Name' cadastrada no Print Management!" -ForegroundColor Green

    Invoke-TestAndQueueUI -Name $Name -DoTest $SendTestPage -DoQueue $OpenPrintQueue
    Write-Host ""
    Write-Host "=== [PASSO 11 - LEMBRETE] Lembre-se de conferir/atualizar '$Name' ($NewIp) no PaperCut Admin! ===" -ForegroundColor Magenta
}

# Função auxiliar: cria a porta PaperCut TCP/IP Port (PAPERCUT_<IP>) no monitor do PaperCut
function Ensure-PaperCutTcpPort {
    param([string]$PortName, [string]$IpAddress, [int]$TcpPort, [string]$Monitor)

    $existing = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        Write-Host "  [OK] Porta '$PortName' já existe no Print Management." -ForegroundColor Green
        return $true
    }

    Write-Host "  [PASSOS 5-8] Criando porta TCP/IP '$PortName' ($IpAddress : $TcpPort)..." -ForegroundColor Cyan
    Add-PrinterPort -Name $PortName -PrinterHostAddress $IpAddress -PortNumber $TcpPort

    $verify = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
    if ($null -eq $verify) {
        throw "Não foi possível confirmar a criação da porta '$PortName' no Spooler."
    }
    Write-Host "  [SUCESSO] Porta '$PortName' registrada no servidor!" -ForegroundColor Green
    return $true
}

function Invoke-TestAndQueueUI {
    param([string]$Name, [bool]$DoTest, [bool]$DoQueue)
    if ($DoTest) {
        Write-Host "  [PASSO 10] Enviando página de teste para '$Name'..." -ForegroundColor Cyan
        $wmiPrinter = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$Name'" -ErrorAction SilentlyContinue
        if ($null -ne $wmiPrinter) {
            Invoke-CimMethod -InputObject $wmiPrinter -MethodName PrintTestPage | Out-Null
            Write-Host "  [OK] Página de teste enviada!" -ForegroundColor Green
        }
    }
    if ($DoQueue) {
        Write-Host "  -> Abrindo janela da fila de impressão de '$Name'..." -ForegroundColor Cyan
        Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /o /n `"$Name`""
    }
}

Write-Host "=== [PASSO 3] Consultando impressora '$PrinterName' no Print Management ===" -ForegroundColor Cyan
$printer = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue

if ($null -eq $printer) {
    Write-Host "[CENÁRIO: NÃO ENCONTRADA] A impressora '$PrinterName' não existe neste servidor." -ForegroundColor Yellow
    Provision-NewPrinterQueue -Name $PrinterName -Port $ExpectedPortName -Driver $DriverName -Loc $Location -Desc $Comment
    return
}

$currentPort = [string]$printer.PortName
Write-Host "  Impressora localizada : $($printer.Name)"
Write-Host "  Porta atual           : $currentPort"
Write-Host "  Porta esperada        : $ExpectedPortName"

if ($currentPort -ieq $ExpectedPortName) {
    Write-Host ""
    Write-Host "=== [PASSO 4] TUDO CERTO! '$PrinterName' já está com a porta '$ExpectedPortName'. ===" -ForegroundColor Green
    if (-not $DryRun) {
        Invoke-TestAndQueueUI -Name $PrinterName -DoTest $SendTestPage -DoQueue $OpenPrintQueue
    }
    return
}

Write-Host ""
Write-Host "=== [PASSO 5-9] ATUALIZAÇÃO NECESSÁRIA: '$currentPort' -> '$ExpectedPortName' ===" -ForegroundColor Yellow
if ($DryRun) {
    Write-Host "[MODO SIMULAÇÃO] Nenhuma alteração aplicada (-DryRun)." -ForegroundColor Cyan
    return
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERRO DE PRIVILÉGIO] Para criar a porta PaperCut, reiniciar o Spooler e alterar a impressora, abra o PowerShell como Administrador e cole o script novamente." -ForegroundColor Red
    return
}

Ensure-PaperCutTcpPort -PortName $ExpectedPortName -IpAddress $NewIp -TcpPort $PortNumber -Monitor $MonitorName | Out-Null

Write-Host "  [PASSO 9] Vinculando nova porta '$ExpectedPortName' na impressora '$PrinterName'..." -ForegroundColor Cyan
Set-Printer -Name $PrinterName -PortName $ExpectedPortName

$updatedPrinter = Get-Printer -Name $PrinterName
Write-Host "  [SUCESSO] '$PrinterName' agora aponta para '$($updatedPrinter.PortName)'!" -ForegroundColor Green

Invoke-TestAndQueueUI -Name $PrinterName -DoTest $SendTestPage -DoQueue $OpenPrintQueue

Write-Host ""
Write-Host "=== [PASSO 11 - LEMBRETE] Lembre-se de atualizar o IP ($NewIp) da '$PrinterName' no PaperCut Admin! ===" -ForegroundColor Magenta
""";
    }

    /// <summary>
    /// Tenta executar o script no servidor (localmente se '.' ou via Invoke-Command WinRM se remoto).
    /// </summary>
    public async Task<(bool Success, string Output)> ExecuteScriptAsync(
        PrinterMigrationRequest request,
        bool dryRunOnly,
        CancellationToken cancellationToken = default)
    {
        ValidateRequest(request);
        string innerScript = BuildPrintServerPowerShellScript(request, dryRunOnly);
        string server = request.PrintServer.Trim();
        bool isLocal = server.Equals(".", StringComparison.Ordinal) ||
                       server.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase);

        string finalScript = isLocal
            ? innerScript
            : $"Invoke-Command -ComputerName '{InputValidator.EscapePowerShellSingleQuoted(server)}' -ScriptBlock {{\n{innerScript}\n}}";

        string tempPs1 = Path.Combine(Path.GetTempPath(), $"ambhub_prtmig_{Guid.NewGuid():N}.ps1");
        try
        {
            await File.WriteAllTextAsync(tempPs1, finalScript, Encoding.UTF8, cancellationToken);

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            psi.ArgumentList.Add("-NoProfile");
            psi.ArgumentList.Add("-NonInteractive");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-File");
            psi.ArgumentList.Add(tempPs1);

            using var proc = Process.Start(psi);
            if (proc == null)
                return (false, "Não foi possível iniciar o processo powershell.exe.");

            var outTask = proc.StandardOutput.ReadToEndAsync(cancellationToken);
            var errTask = proc.StandardError.ReadToEndAsync(cancellationToken);
            await proc.WaitForExitAsync(cancellationToken);

            string stdout = await outTask;
            string stderr = await errTask;
            string combined = string.IsNullOrWhiteSpace(stderr) ? stdout : $"{stdout}\n[STDERR]\n{stderr}";
            bool ok = proc.ExitCode == 0 && !combined.Contains("[STDERR]", StringComparison.Ordinal);

            AppLoggerService.Info(
                $"Execução PrinterMigration '{request.PrinterName}' em '{server}' (DryRun={dryRunOnly}, ExitCode={proc.ExitCode}, Ok={ok}).",
                nameof(PrinterMigrationService));

            return (ok, combined.Trim());
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao executar PrinterMigration em '{server}': {ex.Message}", nameof(PrinterMigrationService));
            return (false, $"Falha na execução remota: {ex.Message}\n\nDica: Se o servidor estiver atrás do Segura® PAM (sem WinRM direto), clique em 'Copiar Script PowerShell' e cole no PowerShell dentro da sessão RDP do servidor.");
        }
        finally
        {
            try { if (File.Exists(tempPs1)) File.Delete(tempPs1); } catch { }
        }
    }

    private static async Task<bool> CheckTcpPortOpenAsync(string ip, int port, int timeoutMs, CancellationToken ct)
    {
        try
        {
            using var tcp = new TcpClient();
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(timeoutMs);
            await tcp.ConnectAsync(ip, port, cts.Token);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
