# --- AMBIENTE VIRTUAL: MODO ESTUDO (FOCO TOTAL) ---

Write-Host "Preparando o Modo Estudo..." -ForegroundColor Green -BackgroundColor Black

# Importa bibliotecas e configurações
. .\biblioteca.ps1
. .\config.ps1

# Chama a função que criamos lá dentro (Isso vai rodar tudo: Office, Programas e Pastas)
Limpar-Ambiente

# Matar Teams
Matar-Teams

##########################
# Abrir pastas em abas
##########################

# Montamos o dicionário (Alias = Caminho). 
# O [ordered] garante que a primeira da lista sempre será a janela mãe!
$minhasPastas = [ordered]@{
    "Download"   = "$env:USERPROFILE\Downloads"
    "Aulas"      = $pastaAulas
}

# Chamamos a função passando o nosso cardápio
Abrir-PastasEmAbas -Pastas $minhasPastas

##########################
# MS Edge
##########################
Write-Host "Iniciando Plataformas EAD e Ferramentas..." -ForegroundColor Green -BackgroundColor Black
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

# =======================================================================
# --- ESTÁGIO: ESTRATÉGIA CONCURSOS ---
# =======================================================================
Write-Host "Acessando a porta de autenticação do Estratégia..." -ForegroundColor Cyan -BackgroundColor Black

$driver.Navigate().GoToUrl("https://www.estrategiaconcursos.com.br/loja/entrar/")

# Pausa mínima de 2 segundos só para garantir que a tela carregou visualmente
Start-Sleep -Seconds 2 

$existeLoginEstrategia = $driver.FindElements([OpenQA.Selenium.By]::Name("loginField"))

if ($existeLoginEstrategia.Count -eq 0) {
    Write-Host "Sessão já está ativa no Estratégia! O site fará o redirecionamento sozinho..." -ForegroundColor Green -BackgroundColor Black
    Start-Sleep -Seconds 2
} 
else {
    Write-Host "Sessão expirada. Injetando credenciais..." -ForegroundColor Yellow -BackgroundColor Black
    
    if (Test-Path $usuarioEstrategia) {
        $credEstudo = Import-Clixml -Path $usuarioEstrategia
        $emailEstrategia = $credEstudo.UserName
        $senhaEstrategia = $credEstudo.GetNetworkCredential().Password
        
        $campoEmail = $existeLoginEstrategia[0]
        $campoSenha = $driver.FindElement([OpenQA.Selenium.By]::Name("passwordField"))
        $teclaCtrlA = [OpenQA.Selenium.Keys]::Control + "a"
        
        # --- SUA LÓGICA AQUI: O CLIQUE PARA ACORDAR O CAMPO ---
        
        # 1. Acorda e preenche o E-mail
        $campoEmail.Click()
        Start-Sleep -Milliseconds 200
        $campoEmail.SendKeys($teclaCtrlA)
        $campoEmail.SendKeys($emailEstrategia)
        
        # 2. Acorda e preenche a Senha
        $campoSenha.Click()
        Start-Sleep -Milliseconds 200
        $campoSenha.SendKeys($teclaCtrlA)
        $campoSenha.SendKeys($senhaEstrategia)
        
        Start-Sleep -Milliseconds 500
        
        Write-Host "Clicando no botão Entrar..." -ForegroundColor Yellow -BackgroundColor Black
        $botaoEntrar = $driver.FindElement([OpenQA.Selenium.By]::XPath("//button[@type='submit']"))
        $botaoEntrar.Click()
        
        Write-Host "Login efetuado! Aguardando carregamento do Dashboard..." -ForegroundColor Green -BackgroundColor Black
        Start-Sleep -Seconds 2 
        
    } else {
        Write-Host "ERRO: Arquivo de credenciais não encontrado!" -ForegroundColor Red -BackgroundColor Black
    }
}

# =======================================================================
# --- ESTÁGIO: GRAN CURSOS ---
# =======================================================================
Write-Host "Verificando porta de autenticação do Gran Cursos..." -ForegroundColor Cyan -BackgroundColor Black

# O comando abaixo é exclusivo do Selenium 4: Ele cria uma aba nova e já foca nela!
$driver.SwitchTo().NewWindow([OpenQA.Selenium.WindowType]::Tab) | Out-Null

# Tentamos ir direto para os meus cursos
$driver.Navigate().GoToUrl($GranConcursos)

# O Gran tem alguns scripts de proteção, vamos dar 4 segundos para a página estabilizar
Start-Sleep -Seconds 4

# Verificamos se o campo de e-mail (ID que você passou) está na tela
$existeLoginGran = $driver.FindElements([OpenQA.Selenium.By]::Id("login-email-site"))

if ($existeLoginGran.Count -eq 0) {
    Write-Host "Sessão já está ativa no Gran Cursos! Área do aluno carregada." -ForegroundColor Green -BackgroundColor Black
} 
else {
    Write-Host "Sessão expirada no Gran. Injetando credenciais de Rezende..." -ForegroundColor Yellow -BackgroundColor Black
    
    if (Test-Path $usuarioGran) {
        $credGran = Import-Clixml -Path $usuarioGran
        $emailGran = $credGran.UserName
        $senhaGran = $credGran.GetNetworkCredential().Password
        
        # Localiza os elementos pelos IDs fornecidos
        $campoEmailGran = $existeLoginGran[0]
        $campoSenhaGran = $driver.FindElement([OpenQA.Selenium.By]::Id("login-senha-site"))
        $botaoEntrarGran = $driver.FindElement([OpenQA.Selenium.By]::Id("login-entrar-site"))
        
        $teclaCtrlA = [OpenQA.Selenium.Keys]::Control + "a"
        
        # Injeta E-mail
        $campoEmailGran.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoEmailGran.SendKeys($emailGran)
        
        # Injeta Senha
        $campoSenhaGran.SendKeys($teclaCtrlA)
        Start-Sleep -Milliseconds 200
        $campoSenhaGran.SendKeys($senhaGran)
        
        Start-Sleep -Milliseconds 500
        
        # Clique no botão entrar
        Write-Host "Finalizando login no Gran..." -ForegroundColor Yellow -BackgroundColor Black
        $botaoEntrarGran.Click()
        
        Write-Host "Login no Gran efetuado com sucesso!" -ForegroundColor Green -BackgroundColor Black
        Start-Sleep -Seconds 5
        
    } else {
        Write-Host "ERRO: Arquivo de credencial não encontrado em $usuarioGran" -ForegroundColor Red -BackgroundColor Black
    }
}

# =======================================================================
# --- ESTÁGIO: UNIGRAN EAD ---
# =======================================================================
Write-Host "Abrindo aba da Unigran EAD..." -ForegroundColor Cyan -BackgroundColor Black

# Abre nova aba e navega
$driver.SwitchTo().NewWindow([OpenQA.Selenium.WindowType]::Tab) | Out-Null
$driver.Navigate().GoToUrl($unigranEAD)

# Aguarda o portal carregar
Start-Sleep -Seconds 4

# Verifica se o campo de login existe (se não existir, já estamos logados)
$existeLoginUnigran = $driver.FindElements([OpenQA.Selenium.By]::Id("login"))

if ($existeLoginUnigran.Count -eq 0) {
    Write-Host "Sessão já está ativa na Unigran!" -ForegroundColor Green -BackgroundColor Black
} 
else {
    Write-Host "Sessão expirada na Unigran. Injetando credenciais de RGA..." -ForegroundColor Yellow -BackgroundColor Black
    
    if (Test-Path $usuarioUnigran) {
        $credUnigran = Import-Clixml -Path $usuarioUnigran
        $rgaUnigran = $credUnigran.UserName
        $senhaUnigran = $credUnigran.GetNetworkCredential().Password
        
        $campoRga = $existeLoginUnigran[0]
        $campoSenhaUnigran = $driver.FindElement([OpenQA.Selenium.By]::Id("senha"))
        $teclaCtrlA = [OpenQA.Selenium.Keys]::Control + "a"
        
        # Injeção dos dados
        $campoRga.SendKeys($teclaCtrlA)
        $campoRga.SendKeys($rgaUnigran)
        
        $campoSenhaUnigran.SendKeys($teclaCtrlA)
        $campoSenhaUnigran.SendKeys($senhaUnigran)
        
        # Primeiro Clique: Login Principal
        Write-Host "Realizando login primário..." -ForegroundColor Yellow -BackgroundColor Black
        $driver.FindElement([OpenQA.Selenium.By]::CssSelector("button.entrar")).Click()
        
        # --- ETAPA 2: TELA DE DISCIPLINAS ---
        Write-Host "Aguardando tela de seleção de disciplina..." -ForegroundColor Cyan -BackgroundColor Black
        Start-Sleep -Seconds 3
        
        # Como o botão é um submit simples, vamos pegá-lo pelo tipo
        $botoesSubmit = $driver.FindElements([OpenQA.Selenium.By]::XPath("//button[@type='submit']"))
        
        if ($botoesSubmit.Count -gt 0) {
            Write-Host "Confirmando acesso à plataforma..." -ForegroundColor Yellow -BackgroundColor Black
            $botoesSubmit[0].Click()
        }
        
        Write-Host "Login Unigran concluído!" -ForegroundColor Green -BackgroundColor Black
        Start-Sleep -Seconds 2
        
    } else {
        Write-Host "ERRO: Arquivo credencial não encontrado em $usuarioUnigran" -ForegroundColor Red -BackgroundColor Black
    }
}

# =======================================================================
# --- ESTÁGIO 4: ABRINDO AS FERRAMENTAS RESTANTES ---
# =======================================================================
Write-Host "Logins críticos concluídos! Abrindo ferramentas Google..." -ForegroundColor Magenta -BackgroundColor Black

$outrasAbas = @($notebookLM, $gemini, $youtubeLofiGirl, $googleCalendar, $keepEstudos, $tasks)

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
# Relógio do Windows
##########################
Write-Host "Abrindo o painel de Sessões de Foco..." -ForegroundColor Green -BackgroundColor Black
# Isso abre o aplicativo Relógio nativo do Windows 11. 
# Basta clicar em "Iniciar" lá dentro para o Windows ativar o Não Incomodar sozinho!
Start-Process "ms-clock:"

Write-Host "Modo Estudo ativado! Bons estudos e foco total." -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

# Carrega a biblioteca gráfica do Windows
Add-Type -AssemblyName System.Windows.Forms

# Monta o "window.alert" com Botão OK e Ícone de Informação (Azulzinho)
[System.Windows.Forms.MessageBox]::Show(
    "O Modo Estudo foi carregado com sucesso. Foco total e bons estudos!", 
    "Automação Concluída", 
    [System.Windows.Forms.MessageBoxButtons]::OK, 
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null