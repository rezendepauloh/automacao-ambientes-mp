# --- AMBIENTE VIRTUAL: MODO ESTUDO (FOCO TOTAL) ---

Write-Host "Preparando o Modo Estudo..." -ForegroundColor Green -BackgroundColor Black

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1

# Chama a função que criamos lá dentro (Isso vai rodar tudo: Office, Programas e Pastas)
Limpar-Ambiente

# --- 1. Abrir Pastas em Abas (Modo Teclado Fantasma) ---
Write-Host "Abrindo pastas de material agrupadas em abas..." -ForegroundColor Cyan -BackgroundColor Black

$wshell = New-Object -ComObject WScript.Shell

# Passo 1: Abre a PRIMEIRA pasta normalmente (Isso cria a janela base)
Start-Process "explorer.exe" -ArgumentList "`"$env:USERPROFILE\Downloads`""

# Dá um tempo bem generoso para a janela do Windows abrir, carregar e ganhar o foco do mouse
Start-Sleep -Seconds 4 

# Passo 2: Abre a SEGUNDA pasta em uma nova aba
# Envia Ctrl + T (Nova Aba)
$wshell.SendKeys("^t")
Start-Sleep -Seconds 1

# Envia Ctrl + L (Focar na barra de endereço lá em cima)
$wshell.SendKeys("^l")
Start-Sleep -Milliseconds 500

# Copia o caminho da segunda pasta para a memória do Windows (evita erros de digitação do robô)
$caminhoAulas = $pastaAulas
Set-Clipboard -Value $caminhoAulas

# Envia Ctrl + V (Colar o caminho)
$wshell.SendKeys("^v")
Start-Sleep -Milliseconds 500

# Envia Enter
$wshell.SendKeys("~")
Start-Sleep -Seconds 2

# --- 2. Abrir o Edge com os sites de Estudo e Música ---
Write-Host "Iniciando Plataformas EAD e Ferramentas..." -ForegroundColor Green -BackgroundColor Black
$sitesEstudo = @(
    "--disable-extensions",
    "--disable-session-crashed-bubble", 
    $unigranEAD,
    $notebookLM,
    $gemini,
    $youtubeLofiGirl,
    $googleCalendar,
    $keepEstudos,
    $tasks,
    $estrategiaConcursos,
    $GranConcursos
)

Start-Process "msedge.exe" -ArgumentList $sitesEstudo

# --- 3. Chamar o Relógio do Windows (Sessões de Foco) ---
Write-Host "Abrindo o painel de Sessões de Foco..." -ForegroundColor Green -BackgroundColor Black
# Isso abre o aplicativo Relógio nativo do Windows 11. 
# Basta clicar em "Iniciar" lá dentro para o Windows ativar o Não Incomodar sozinho!
Start-Process "ms-clock:"

Write-Host "Modo Estudo ativado! Bons estudos e foco total." -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

Write-Host "Ambiente carregado! Exibindo aviso..." -ForegroundColor Green -BackgroundColor Black

# Carrega a biblioteca gráfica do Windows
Add-Type -AssemblyName System.Windows.Forms

# Monta o "window.alert" com Botão OK e Ícone de Informação (Azulzinho)
[System.Windows.Forms.MessageBox]::Show(
    "O Modo Estudo foi carregado com sucesso. Foco total e bons estudos!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null