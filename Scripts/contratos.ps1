# --- AMBIENTE VIRTUAL: FISCAL DE CONTRATOS ---

Write-Host "Iniciando limpeza para o Ambiente de Contratos..." -ForegroundColor Green -BackgroundColor Black

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1

# 2. Chama a função que criamos lá dentro (Isso vai rodar tudo: Office, Programas e Pastas)
Limpar-Ambiente

# Matar Teams
Matar-Teams

##########################
# Abrir pastas em abas
##########################

# Montamos o dicionário (Alias = Caminho). 
# O [ordered] garante que a primeira da lista sempre será a janela mãe!
$minhasPastas = [ordered]@{
    "Download"         = "$env:USERPROFILE\Downloads"
    "Contratos"        = $pastaContratos
}

# Chamamos a função passando o nosso cardápio
Abrir-PastasEmAbas -Pastas $minhasPastas

##########################
# MS Teams
##########################
Write-Host "Iniciando Microsoft Teams..." -ForegroundColor Green -BackgroundColor Black
Start-Process "msteams:"

##########################
# MS Edge Leve (Nativo e Modular)
##########################

# Montamos o dicionário (Alias = URL)
$meusSitesContratos = [ordered]@{
    "YouTube Music"   = $youtubeMusic
    "Google Gemini"   = $gemini
    "Google Keep"     = $keepChamados
    "Google Tasks"    = $tasks
    "Google Calendar" = $googleCalendar
}

# Chamamos a função da nossa biblioteca!
Abrir-SitesEdgeLeve -Sites $meusSitesContratos

##########################
# WhatsApp App
##########################
Write-Host "Aguardando o Edge principal estabilizar..." -ForegroundColor Yellow -BackgroundColor Black
Start-Sleep -Seconds 3 # 👇 ESSA PAUSA É O SEGREDO 👇

Write-Host "Iniciando WhatsApp App..." -ForegroundColor Green -BackgroundColor Black
$idWhatsApp = "--app-id=$($idWhatsAppConfig)" 
Start-Process "msedge.exe" -ArgumentList $idWhatsApp

# Dá um tempo para o Teams e o Edge "pularem" na tela e não atrapalharem depois
Start-Sleep -Seconds 4

##########################
# SAJMP
##########################
Write-Host "Iniciando SAJMP..." -ForegroundColor Green -BackgroundColor Black
Start-Process $sajmpAtalho

Write-Host "Aguardando o SAJMP carregar a tela de login..." -ForegroundColor Green -BackgroundColor Black
# IMPORTANTE: Se o seu computador for muito rápido, pode diminuir esse tempo. 
# Se for mais lento para abrir o SAJ, aumente para 6 ou 7 segundos.
Start-Sleep -Seconds 10 

# Cria o objeto que simula o teclado
$wshell = New-Object -ComObject wscript.shell

# --- 4. Forçar o foco na janela do SAJMP ---
$processoSAJ = Get-Process saj -ErrorAction SilentlyContinue
if ($processoSAJ) {
    Write-Host "Puxando a janela do SAJMP para frente..." -ForegroundColor Yellow -BackgroundColor Black
    $wshell.AppActivate($processoSAJ.Id)
    Start-Sleep -Milliseconds 500 # Meio segundo para o Windows trazer a janela
}

Write-Host "Buscando credencial segura..." -ForegroundColor Green -BackgroundColor Black
$credSAJ = Import-Clixml -Path $credenciais
$senhaDescriptografada = $credSAJ.GetNetworkCredential().Password

Write-Host "Digitando credenciais..." -ForegroundColor Green -BackgroundColor Black

# Aperta TAB para pular do campo de Usuário para o campo de Senha
$wshell.SendKeys("{TAB}")
Start-Sleep -Milliseconds 200 # Pausa rapidinha imitando um humano

# Digita a senha (Lembre de colocar a sua senha aqui)
$wshell.SendKeys($senhaDescriptografada)
Start-Sleep -Milliseconds 200

# Aperta o Enter para logar
$wshell.SendKeys("{ENTER}")

Write-Host "Ambiente Fiscal de Contratos carregado de forma 100% segura!" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

# --- DESPISTANDO O ANTIVÍRUS ---
Remover-CredenciaisMemoria

# Carrega a biblioteca gráfica do Windows
Add-Type -AssemblyName System.Windows.Forms

# Monta o "window.alert" com Botão OK e Ícone de Informação (Azulzinho)
[System.Windows.Forms.MessageBox]::Show(
    "O Modo Contratos foi carregado com sucesso. Manda bala nesse ETP!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null