# --- AMBIENTE VIRTUAL: PXE ---

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1

# 2. Chama a função que criamos lá dentro (Isso vai rodar tudo: Office, Programas e Pastas)
Limpar-Ambiente

# --- 3. Abrir o AD (Usando runas em cascata para forçar a elevação UAC) ---
Write-Host "Iniciando Active Directory..." -ForegroundColor Green -BackgroundColor Black

# Na PRIMEIRA execução, uma tela preta do CMD pedirá a senha. Nas próximas, vai abrir direto!
runas.exe /user:$($usuarioAdminAD) /savecred "powershell.exe -WindowStyle Hidden -Command `"Start-Process mmc dsa.msc -Verb RunAs`""

# 4. Iniciar o Configuration Manager
Write-Host "Iniciando Configuration Manager..." -ForegroundColor Green -BackgroundColor Black

$atalhoSCCM = $atalhoInternoSCCM
Start-Process $atalhoSCCM

Write-Host "Ambiente PXE carregado!" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

# Carrega a biblioteca gráfica do Windows
Add-Type -AssemblyName System.Windows.Forms

# Monta o "window.alert" com Botão OK e Ícone de Informação (Azulzinho)
[System.Windows.Forms.MessageBox]::Show(
    "O Modo PXE foi carregado com sucesso. Pode tirar a máquina e formatar!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null