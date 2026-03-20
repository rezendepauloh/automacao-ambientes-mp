# --- AMBIENTE VIRTUAL: Login teste ---

Write-Host "Preparando o Modo Login..." -ForegroundColor Green -BackgroundColor Black

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1


Write-Host "Iniciando Terminais..." -ForegroundColor Cyan -BackgroundColor Black

#Write-Host "  -> Abrindo Terminal Padrão..." -ForegroundColor Gray
#Start-Process "wt.exe"

# Terminal Elevado com as 4 abas (Mágica do Script Temporário)
Write-Host "  -> Abrindo Terminal Elevado..." -ForegroundColor Gray

$caminhoScriptTemp = "$env:TEMP\IniciaTerminalAdmin.ps1"

# 1. A MÁGICA DO ARQUIVO TEMPORÁRIO AVANÇADO:
$conteudoTemp = @"
`$pastaRaiz = `$env:USERPROFILE
`$argWt = '-w new -d "' + `$pastaRaiz + '" ; new-tab -d "$pastaScripts1" ; new-tab -d "$pastaScripts2" ; new-tab -d "$pastaScripts3"'
Start-Process wt.exe -ArgumentList `$argWt -Verb RunAs

# A SUA IDEIA AQUI: O script mata o próprio processo (PID) instantaneamente!
# Stop-Process -Id `$PID -Force
"@

Set-Content -Path $caminhoScriptTemp -Value $conteudoTemp -Encoding UTF8


$cofreAdminComum = "C:\AmbientesVirtuais\Credenciais\cred_admin.xml"

Write-Host "Lendo credenciais de administrador do cofre..." -ForegroundColor DarkGray
$credencialAdmin = Import-Clixml -Path $cofreAdminComum

Write-Host "Iniciando Terminal Elevado de forma nativa e silenciosa..." -ForegroundColor Cyan
Start-Process -FilePath "pwsh.exe" -ArgumentList "-WindowStyle Hidden -NoProfile -NonInteractive -File `"$caminhoScriptTemp`"" -Credential $credencialAdmin -WindowStyle Hidden