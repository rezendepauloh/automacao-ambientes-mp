# ==============================================================================
# AMBIENTES HUB — RUNNER UNIFICADO DA SUÍTE DE TESTES AUTOMATIZADOS (C#)
# ==============================================================================

$Host.UI.RawUI.WindowTitle = "Ambientes Hub — Test Runner"

$esc = [char]27
$RESET = "$esc[0m"
$BOLD = "$esc[1m"
$GREEN = "$esc[32m"
$RED = "$esc[31m"
$YELLOW = "$esc[33m"
$CYAN = "$esc[36m"
$DARK_GRAY = "$esc[90m"

Write-Host ""
Write-Host "$YELLOW╔═══════════════════════════════════════════════════════════════╗$RESET"
Write-Host "$YELLOW║         AMBIENTES HUB — SUÍTE DE TESTES AUTOMATIZADOS         ║$RESET"
Write-Host "$YELLOW╚═══════════════════════════════════════════════════════════════╝$RESET"
Write-Host ""

$testsProj = Join-Path $PSScriptRoot "AmbientesHub.Tests.csproj"

if (-not (Test-Path $testsProj)) {
    Write-Host "$RED✖ Erro: Projeto de testes não encontrado em $testsProj$RESET"
    exit 1
}

Write-Host "$CYAN▶ Executando Suíte Completa: Unit, Services & Integration$RESET"
Write-Host "$DARK_GRAY$('─' * 65)$RESET"

$start = [System.Diagnostics.Stopwatch]::StartNew()

# Executa o runner nativo do .NET (xUnit)
dotnet test $testsProj --nologo --logger "console;verbosity=normal"

$start.Stop()
$elapsed = $start.ElapsedMilliseconds

Write-Host "$DARK_GRAY$('─' * 65)$RESET"
if ($LASTEXITCODE -eq 0) {
    Write-Host "$GREEN✔ TODOS OS TESTES PASSARAM COM SUCESSO!$RESET (⏱ $YELLOW$($elapsed)ms$RESET)"
} else {
    Write-Host "$RED✖ FALHA NA EXECUÇÃO DOS TESTES.$RESET (Exit Code: $LASTEXITCODE | ⏱ $YELLOW$($elapsed)ms$RESET)"
}
Write-Host ""
exit $LASTEXITCODE

