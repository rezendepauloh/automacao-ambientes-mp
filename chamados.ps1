# --- AMBIENTE VIRTUAL: CHAMADOS ---

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1

# 2. Chama a função que criamos lá dentro (Isso vai rodar tudo: Office, Programas e Pastas)
Limpar-Ambiente

# --- 1. Abrir Pastas em Abas (Modo Teclado Fantasma) ---
Write-Host "Abrindo pastas de trabalho agrupadas em abas..." -ForegroundColor Cyan -BackgroundColor Black

$wshell = New-Object -ComObject WScript.Shell

# Passo 1: Abre a PRIMEIRA pasta normalmente (Isso cria a janela base)
Start-Process "explorer.exe" -ArgumentList "`"$env:USERPROFILE\Downloads`""
Write-Host "  -> Download aberto" -ForegroundColor Cyan -BackgroundColor Black

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
$caminhoAulas = $pastaProvas
Set-Clipboard -Value $caminhoAulas

# Envia Ctrl + V (Colar o caminho)
$wshell.SendKeys("^v")
Start-Sleep -Milliseconds 500

# Envia Enter
$wshell.SendKeys("~")
Write-Host "  -> Provas aberto" -ForegroundColor Cyan -BackgroundColor Black
Start-Sleep -Seconds 2


# Passo 3: Abre a TERCEIRA pasta em uma nova aba
# Envia Ctrl + T (Nova Aba)
$wshell.SendKeys("^t")
Start-Sleep -Seconds 1

# Envia Ctrl + L (Focar na barra de endereço lá em cima)
$wshell.SendKeys("^l")
Start-Sleep -Milliseconds 500

# Copia o caminho da segunda pasta para a memória do Windows (evita erros de digitação do robô)
$caminhoSharePoint = $pastaSharePoint
Set-Clipboard -Value $caminhoSharePoint

# Envia Ctrl + V (Colar o caminho)
$wshell.SendKeys("^v")
Start-Sleep -Milliseconds 500

# Envia Enter
$wshell.SendKeys("~")
Write-Host "  -> Pasta SharePoint aberta" -ForegroundColor Cyan -BackgroundColor Black
Start-Sleep -Seconds 2

# --- 2. Abrir a Planilha de Chamados no Excel ---
Write-Host "Abrindo Planilha de Chamados..." -ForegroundColor Green -BackgroundColor Black
$planilha = $planilhaChamados
Start-Process "excel.exe" -ArgumentList "`"$planilha`""

foreach ($pasta in $pastas) {
    # Usamos explorer.exe passando o caminho da pasta
    Start-Process "explorer.exe" -ArgumentList "`"$pasta`""
    Start-Sleep -Milliseconds 500 # Uma pequena pausa para o Windows não se engasgar
}

# --- 3. Abrir o Edge SEM Extensões e com as abas do setor ---
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
    
    $caminhoCofre = $credenciais
    
    if (Test-Path $caminhoCofre) {
        $credencial = Import-Clixml -Path $caminhoCofre
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
    
    $caminhoCofre = $credenciais 
    
    if (Test-Path $caminhoCofre) {
        $credencial = Import-Clixml -Path $caminhoCofre
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
    # Abre uma nova aba e acessa a URL
    $driver.SwitchTo().NewWindow([OpenQA.Selenium.WindowType]::Tab) | Out-Null
    $driver.Navigate().GoToUrl($url)
    Start-Sleep -Milliseconds 500 # Uma pausa bem curtinha só para o Edge não engasgar
}

# --- Abrir o WhatsApp (App do Edge) ---
Write-Host "Iniciando WhatsApp App..." -ForegroundColor Green -BackgroundColor Black
# Substitua o texto abaixo pelo ID que você copiou do seu atalho
$idWhatsApp = "--app-id=$($idWhatsAppConfig)" 
Start-Process "msedge.exe" -ArgumentList $idWhatsApp

Write-Host "Ambiente Chamados carregado com sucesso!" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

# Carrega a biblioteca gráfica do Windows
Add-Type -AssemblyName System.Windows.Forms

# Monta o "window.alert" com Botão OK e Ícone de Informação (Azulzinho)
[System.Windows.Forms.MessageBox]::Show(
    "O Modo Chamado foi carregado com sucesso. Foco total e excelente trabalho!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null