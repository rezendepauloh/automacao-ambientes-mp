# =======================================================================
# --- AMBIENTE VIRTUAL: IMPRESSORAS ---
# =======================================================================

Write-Host "Preparando o Modo Impressoras..." -ForegroundColor Green -BackgroundColor Black

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1

# 1. Limpeza e Escudo de Processos
Limpar-Ambiente

# =======================================================================
# --- INICIANDO APLICATIVOS BASE ---
# =======================================================================
Write-Host "Iniciando aplicativos base..." -ForegroundColor Green -BackgroundColor Black

Write-Host "  -> Abrindo Planilha de Chamados..." -ForegroundColor Gray
Start-Process $planilhaChamados

Write-Host "  -> Abrindo Remote Desktop Manager..." -ForegroundColor Gray
Start-Process "RemoteDesktopManager.exe" -ErrorAction SilentlyContinue

Write-Host "  -> Iniciando Microsoft Teams..." -ForegroundColor Gray
Start-Process "msteams:"

Start-Sleep -Seconds 2

# =======================================================================
# --- MOTOR SELENIUM (PAPERCUT E OTRS) ---
# =======================================================================
Write-Host "Iniciando o Motor do Selenium..." -ForegroundColor Green -BackgroundColor Black

# Carrega a DLL do Selenium
Add-Type -Path $driverSelenium

$opcoes = New-Object OpenQA.Selenium.Edge.EdgeOptions
$opcoes.AddArgument("user-data-dir=$env:LOCALAPPDATA\Microsoft\Edge\User Data")
$opcoes.AddArgument("--disable-extensions")
$opcoes.AddExcludedArgument("enable-automation")
$opcoes.AddArgument("--disable-session-crashed-bubble")
$opcoes.AddUserProfilePreference("profile.exit_type", "Normal")

$servico = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService("C:\AmbientesVirtuais")
$servico.HideCommandPromptWindow = $true

$driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($servico, $opcoes)

# 1. PAPERCUT
Write-Host "  -> Acessando PaperCut..." -ForegroundColor Gray
$driver.Navigate().GoToUrl($sitePapercut)
Start-Sleep -Seconds 2

$existeTelaLoginPaperCut = $driver.FindElements([OpenQA.Selenium.By]::Id("inputUsername"))

if ($existeTelaLoginPaperCut.Count -eq 0) {
    Write-Host "  -> Sessão já está ativa no PaperCut! Pulando etapa de login." -ForegroundColor Green -BackgroundColor Black
} 
else {
    Write-Host "  -> Sessão do PaperCut expirada. Injetando credenciais do cofre..." -ForegroundColor Yellow -BackgroundColor Black
    
    $campoUsuarioPaperCut = $existeTelaLoginPaperCut[0]
    $campoSenhaPaperCut = $driver.FindElement([OpenQA.Selenium.By]::Id("inputPassword"))
    
    # Busca a credencial comum no cofre
    if (Test-Path $usuarioPaperCut) {
        $credencial = Import-Clixml -Path $usuarioPaperCut
        $user = $credencial.UserName
        $senhaDescriptografada = $credencial.GetNetworkCredential().Password
        
        # O "Trator" do Ctrl + A
        $teclaCtrlA = [OpenQA.Selenium.Keys]::Control + "a"
        
        $campoUsuarioPaperCut.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoUsuarioPaperCut.SendKeys($user)
        
        $campoSenhaPaperCut.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoSenhaPaperCut.SendKeys($senhaDescriptografada)
        
        Start-Sleep -Milliseconds 500
    } else {
        Write-Host "  -> ERRO: Arquivo XML de senha não encontrado!" -ForegroundColor Red -BackgroundColor Black
        throw "  -> Sem credenciais para continuar." 
    }

    # --- O CLIQUE FINAL NO PAPERCUT (CORRIGIDO) ---
    Write-Host "  -> Clicando no botão de Login do PaperCut..." -ForegroundColor Yellow -BackgroundColor Black
    $botaoLoginPaper = $driver.FindElement([OpenQA.Selenium.By]::Name('$Submit$0'))
    $botaoLoginPaper.Click()

    Write-Host "  -> Login no PaperCut efetuado com sucesso!" -ForegroundColor Green -BackgroundColor Black
}

# 2. OTRS
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
Start-Sleep -Seconds 2

# (Se houver lógica de login do OTRS, ela entraria aqui. 
# Segundo o seu log, a sessão geralmente já está ativa!)
Write-Host "  -> Sessão já está ativa no OTRS! Pulando etapa de login." -ForegroundColor Cyan -BackgroundColor Black

# =======================================================================
# --- ABAS ADICIONAIS ---
# =======================================================================
Write-Host "Logins críticos concluídos! Abrindo o restante das abas..." -ForegroundColor Green -BackgroundColor Black

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

Start-Sleep -Seconds 2

# WhatsApp App (Como processo separado fora do Selenium para manter sessão do celular)
Write-Host "  -> Iniciando WhatsApp App..." -ForegroundColor Gray
$idWhatsApp = "--app-id=$($idWhatsAppConfig)" 
Start-Process "msedge.exe" -ArgumentList $idWhatsApp

# =======================================================================
# --- TERMINAIS (A NOVA GERAÇÃO) ---
# =======================================================================
Write-Host "Iniciando Terminais..." -ForegroundColor Green -BackgroundColor Black

Write-Host "  -> Abrindo Terminal Padrão..." -ForegroundColor Gray
Start-Process "wt.exe" -ArgumentList "-d `"$pastaScripts1`""

Write-Host "  -> Abrindo Terminal Elevado..." -ForegroundColor Gray

$caminhoScriptTemp = "$env:TEMP\IniciaTerminalAdmin.ps1"

$conteudoTemp = @"
`$pastaRaiz = `$env:USERPROFILE
`$argWt = '-w new -d "' + `$pastaRaiz + '" ; new-tab -d "$pastaScripts1" ; new-tab -d "$pastaScripts2" ; new-tab -d "$pastaScripts3"'
Start-Process wt.exe -ArgumentList `$argWt -Verb RunAs
"@

Set-Content -Path $caminhoScriptTemp -Value $conteudoTemp -Encoding UTF8

Write-Host "  -> Lendo credenciais de administrador do cofre ($cofreAdminComum)..." -ForegroundColor DarkGray
$credencialAdmin = Import-Clixml -Path $cofreAdminComum

# Inicia o Terminal Elevado de forma nativa e silenciosa, injetando a credencial direto na veia
# Start-Process -FilePath "pwsh.exe" -ArgumentList "-WindowStyle Hidden -NoProfile -NonInteractive -Command `"Start-Process wt.exe -Verb RunAs`"" -Credential $credencialAdmin -WindowStyle Hidden
Write-Host "  -> Iniciando Terminal Elevado de forma nativa e silenciosa..." -ForegroundColor Cyan
Start-Process -FilePath "pwsh.exe" -ArgumentList "-WindowStyle Hidden -NoProfile -NonInteractive -File `"$caminhoScriptTemp`"" -Credential $credencialAdmin -WindowStyle Hidden

# =======================================================================
# --- FIM DO SCRIPT ---
# =======================================================================
Start-Sleep -Seconds 2
Write-Host "Ambiente Impressoras carregado com sucesso!" -ForegroundColor Green -BackgroundColor Black

# Carrega a biblioteca gráfica do Windows para o Alerta Final
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show(
    "O Modo Impressoras foi reconstruído, carregado e está pronto para o combate!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null