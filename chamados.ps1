# --- AMBIENTE VIRTUAL: CHAMADOS ---

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
    "Provas"           = $pastaProvas
    "Pasta SharePoint" = $pastaSharePoint
}

# Chamamos a função passando o nosso cardápio
Abrir-PastasEmAbas -Pastas $minhasPastas

##########################
# Planilha Chamados
##########################
Write-Host "Abrindo Planilha de Chamados..." -ForegroundColor Green -BackgroundColor Black
Start-Process "excel.exe" -ArgumentList "`"$planilhaChamados`""

##########################
# MS Edge Leve
##########################
Write-Host "Abrindo sistemas web no Edge (Modo Leve)..." -ForegroundColor Green -BackgroundColor Black
# 1. Carrega a biblioteca do Selenium
Add-Type -Path $driverSelenium

# 2. Configurações do Edge
$opcoes = New-Object OpenQA.Selenium.Edge.EdgeOptions
$opcoes.AddArgument("user-data-dir=$env:LOCALAPPDATA\Microsoft\Edge\User Data")
$opcoes.AddArgument("--disable-extensions")

# --- A MÁGICA PARA ESCONDER A FAIXA DE TESTE AUTOMATIZADO ---
$opcoes.AddExcludedArgument("enable-automation")

# --- A MÁGICA ANTI-CRASH ---
# Impede que o Edge mostre aquele balão chato de "O Edge foi fechado inesperadamente. Restaurar páginas?"
$opcoes.AddArgument("--disable-session-crashed-bubble")
$opcoes.AddUserProfilePreference("profile.exit_type", "Normal")

# 3. Prepara o Motorista Invisível
$servico = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService("C:\AmbientesVirtuais")
$servico.HideCommandPromptWindow = $true

# 4. Abre o Navegador
Write-Host "Iniciando o Motor do Selenium..." -ForegroundColor Yellow
$driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($servico, $opcoes)

# 5. Navega para o SIMP
Write-Host "Acessando o SIMP..." -ForegroundColor Cyan -BackgroundColor Black
$driver.Navigate().GoToUrl($simp)

# Dá 3 segundos generosos para a página carregar, redirecionar ou o Edge preencher a senha
Start-Sleep -Seconds 3

# --- A MÁGICA DA VERIFICAÇÃO DE SESSÃO ---
$existeTelaLogin = $driver.FindElements([OpenQA.Selenium.By]::Id("username"))

if ($existeTelaLogin.Count -eq 0) {
    # Cenário 1: Já está logado!
    Write-Host "Sessão já está ativa no SIMP! Pulando etapa de login." -ForegroundColor Green -BackgroundColor Black
} 
else {
    Write-Host "Sessão expirada. Injetando credenciais do cofre XML..." -ForegroundColor Yellow -BackgroundColor Black
    
    $campoUsuario = $existeTelaLogin[0]
    $campoSenha = $driver.FindElement([OpenQA.Selenium.By]::Id("password"))
    
    if (Test-Path $credenciais) {
        $credencial = Import-Clixml -Path $credenciais
        $senhaDescriptografada = $credencial.GetNetworkCredential().Password
        
        # --- SOLUÇÃO 1: Limpar o domínio do usuário ---
        $usuarioLimpo = $credencial.UserName.Replace("$($dominioLocal)\", "")
        
        # --- SOLUÇÃO 2: O "Trator" do Ctrl + A ---
        # Prepara a combinação de teclas Ctrl + A
        $teclaCtrlA = [OpenQA.Selenium.Keys]::Control + "a"
        
        Write-Host "Forçando limpeza dos campos e injetando credenciais limpas..." -ForegroundColor Cyan -BackgroundColor Black
        
        # Campo de Usuário: Seleciona tudo (Ctrl+A) e digita o usuário limpo por cima
        $campoUsuario.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200 # Pausa rápida para a página respirar
        $campoUsuario.SendKeys($usuarioLimpo)
        
        # Campo de Senha: Seleciona tudo (Ctrl+A) e digita a senha por cima
        $campoSenha.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoSenha.SendKeys($senhaDescriptografada)
        
        Start-Sleep -Milliseconds 500
    } else {
        Write-Host "ERRO: Arquivo XML de senha não encontrado!" -ForegroundColor Red -BackgroundColor Black
        throw "Sem credenciais para continuar." 
    }

    # --- O CLIQUE FINAL ---
    Write-Host "Clicando no botão de Acessar..." -ForegroundColor Yellow -BackgroundColor Black
    $botaoEntrar = $driver.FindElement([OpenQA.Selenium.By]::XPath("//button[@type='submit']"))
    $botaoEntrar.Click()
    
    Write-Host "Login no SIMP efetuado com sucesso!" -ForegroundColor Green -BackgroundColor Black
}

# =======================================================================
# --- ESTÁGIO 2: CENTRAL (OTRS) ---
# =======================================================================

Write-Host "Criando nova aba para a Central (OTRS)..." -ForegroundColor Cyan -BackgroundColor Black
# O comando abaixo é exclusivo do Selenium 4: Ele cria uma aba nova e já foca nela!
$driver.SwitchTo().NewWindow([OpenQA.Selenium.WindowType]::Tab) | Out-Null

$driver.Navigate().GoToUrl($otrs)

# Dá 3 segundos para a página carregar ou redirecionar
Start-Sleep -Seconds 3 

# --- A MÁGICA DA VERIFICAÇÃO DE SESSÃO NO OTRS ---
$existeTelaLoginOTRS = $driver.FindElements([OpenQA.Selenium.By]::Id("User"))

if ($existeTelaLoginOTRS.Count -eq 0) {
    Write-Host "Sessão já está ativa no OTRS! Pulando etapa de login." -ForegroundColor Green -BackgroundColor Black
} 
else {
    Write-Host "Sessão do OTRS expirada. Injetando credenciais do cofre..." -ForegroundColor Yellow -BackgroundColor Black
    
    $campoUsuarioOTRS = $existeTelaLoginOTRS[0]
    $campoSenhaOTRS = $driver.FindElement([OpenQA.Selenium.By]::Id("Password"))
    
    if (Test-Path $credenciais) {
        $credencial = Import-Clixml -Path $credenciais
        $senhaDescriptografada = $credencial.GetNetworkCredential().Password
        
        # Reaproveitamos a lógica de limpar o domínio
        $usuarioLimpo = $credencial.UserName.Replace("$($dominioLocal)\", "")
        
        # O "Trator" do Ctrl + A
        $teclaCtrlA = [OpenQA.Selenium.Keys]::Control + "a"
        
        $campoUsuarioOTRS.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoUsuarioOTRS.SendKeys($usuarioLimpo)
        
        $campoSenhaOTRS.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoSenhaOTRS.SendKeys($senhaDescriptografada)
        
        Start-Sleep -Milliseconds 500
    } else {
        Write-Host "ERRO: Arquivo XML de senha não encontrado!" -ForegroundColor Red -BackgroundColor Black
        throw "Sem credenciais para continuar." 
    }

    # --- O CLIQUE FINAL NO OTRS ---
    Write-Host "Clicando no botão de Login do OTRS..." -ForegroundColor Yellow -BackgroundColor Black
    $botaoEntrarOTRS = $driver.FindElement([OpenQA.Selenium.By]::Id("LoginButton"))
    $botaoEntrarOTRS.Click()
    
    Write-Host "Login no OTRS efetuado com sucesso!" -ForegroundColor Green -BackgroundColor Black
}

# =======================================================================
# --- ESTÁGIO 3: ABRINDO AS FERRAMENTAS RESTANTES ---
# =======================================================================
Write-Host "Logins críticos concluídos! Abrindo o restante das abas..." -ForegroundColor Magenta -BackgroundColor Black

# A sua lista de sites que não precisam de login forçado
$outrasAbas = @(
    $citsmart,
    $sharePointSite,
    $gemini,
    $youtubeMusic,
    $keepChamados,
    $tasks,
    $googleCalendar
)

foreach ($url in $outrasAbas) {
    $driver.SwitchTo().NewWindow([OpenQA.Selenium.WindowType]::Tab) | Out-Null
    $driver.Navigate().GoToUrl($url)
    
    # Se for o Gemini ou NotebookLM, damos um tempo extra para o redirecionamento de conta
    if ($url -like "*gemini.google.com*" -or $url -like "*notebooklm*") {
        Write-Host "  -> Sincronizando conta acadêmica para ferramenta de IA..." -ForegroundColor Gray
        Start-Sleep -Seconds 2
    }
    
    Start-Sleep -Milliseconds 800
}

##########################
# WhatsApp App
##########################
Write-Host "Aguardando o Edge principal estabilizar..." -ForegroundColor Yellow -BackgroundColor Black
Start-Sleep -Seconds 3 # 👇 ESSA PAUSA É O SEGREDO 👇

Write-Host "Iniciando WhatsApp App..." -ForegroundColor Green -BackgroundColor Black
$idWhatsApp = "--app-id=$($idWhatsAppConfig)" 
Start-Process "msedge.exe" -ArgumentList $idWhatsApp

##########################
# MS Teams
##########################
Write-Host "Iniciando Microsoft Teams..." -ForegroundColor Green -BackgroundColor Black
# No Windows 11, a melhor forma de chamar o Teams novo é usando o protocolo URI dele
Start-Process "msteams:"

##########################
# Windows Terminal
##########################
Write-Host "  -> Abrindo Terminal Padrão..." -ForegroundColor Gray
Start-Process "wt.exe"

Start-Sleep -Seconds 2

# =======================================================================
# --- INVOCANDO TERMINAL ADMIN COM ABAS ESPECÍFICAS ---
# =======================================================================
Write-Host "  -> Abrindo Terminal Elevado..." -ForegroundColor Gray

$caminhoScriptTemp = "$env:TEMP\IniciaTerminalAdmin.ps1"

# 1. A MÁGICA DO ARQUIVO TEMPORÁRIO AVANÇADO:
# Usamos o 'Here-String' (@" "@) para montar o código com perfeição.
# A crase (`) antes do $ impede o PowerShell atual de ler a variável, forçando a leitura apenas no Admin!
$conteudoTemp = @"
`$pastaRaiz = `$env:USERPROFILE
`$argWt = '-w new -d "' + `$pastaRaiz + '" ; new-tab -d "$pastaScripts1" ; new-tab -d "$pastaScripts2" ; new-tab -d "$pastaScripts3"'
Start-Process wt.exe -ArgumentList `$argWt -Verb RunAs

# A SUA IDEIA AQUI: O script mata o próprio processo (PID) instantaneamente!
Stop-Process -Id `$PID -Force
"@

Set-Content -Path $caminhoScriptTemp -Value $conteudoTemp -Encoding UTF8

# 2. O TIRO FINAL (AGORA COM -NoProfile PARA MATAR A JANELA FANTASMA)
runas.exe /user:$($usuarioAdminComum) /savecred "pwsh.exe -WindowStyle Hidden -NoProfile -NonInteractive -WindowStyle Hidden -File `"$caminhoScriptTemp`""

Start-Sleep -Seconds 2

Write-Host "Ambiente Chamados carregado com sucesso!" -ForegroundColor Green -BackgroundColor Black

# Carrega a biblioteca gráfica do Windows
Add-Type -AssemblyName System.Windows.Forms

# Monta o "window.alert" com Botão OK e Ícone de Informação (Azulzinho)
[System.Windows.Forms.MessageBox]::Show(
    "O Modo Chamado foi carregado com sucesso. Foco total e excelente trabalho!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null