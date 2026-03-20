# ==========================================
# BIBLIOTECA DE FUNÇÕES - AUTOMAÇÃO MP
# ==========================================

# Inicia a gravação de tudo que acontece no terminal
Start-Transcript -Path "$env:USERPROFILE\Desktop\Log_Automacao.txt" -Force

function Limpar-Ambiente {
    Write-Host "Iniciando limpeza do ambiente..." -ForegroundColor Green -BackgroundColor Black

    # --- 1. SALVAR E FECHAR OFFICE ---
    Write-Host "Salvando e fechando documentos do Office de forma segura..." -ForegroundColor Cyan -BackgroundColor Black
    
    $codigoSalvarOffice = {
        # --- FECHAR EXCEL ---
        # Só tenta fazer algo se o processo do Excel existir
        if (Get-Process -Name "excel" -ErrorAction SilentlyContinue) {
            try {
                $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                Write-Host "Entrou no Excel" -ForegroundColor Cyan -BackgroundColor Black
                
                $excel.DisplayAlerts = $false 
                
                foreach ($wb in $excel.Workbooks) {
                    Write-Host "Vendo a planilha: $($wb.Name)" -ForegroundColor Cyan -BackgroundColor Black

                    if ([string]::IsNullOrEmpty($wb.Path)) {
                        $nome = "Planilha_Salva_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
                        $caminho = Join-Path $env:USERPROFILE "Downloads\$nome"

                        Write-Host "Tentando salvar planilha nova em: $caminho" -ForegroundColor Cyan -BackgroundColor Black
                        $wb.SaveAs($caminho)
                        Write-Host "Planilha nova salva com sucesso!" -ForegroundColor Green -BackgroundColor Black
                    } else {
                        Write-Host "Salvando alterações na planilha existente..." -ForegroundColor Cyan -BackgroundColor Black
                        $wb.Save()
                        Write-Host "Planilha existente salva com sucesso!" -ForegroundColor Green -BackgroundColor Black
                    }
                }
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            } catch { 
                Write-Host "Erro ao manipular o Excel: $_" -ForegroundColor Red -BackgroundColor Black
            }
        }

        # --- FECHAR WORD ---
        # Só tenta fazer algo se o processo do Word existir
        if (Get-Process -Name "winword" -ErrorAction SilentlyContinue) {
            try {
                $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
                Write-Host "Entrou no Word" -ForegroundColor Cyan -BackgroundColor Black
                
                $word.DisplayAlerts = 0 
                
                foreach ($doc in $word.Documents) {
                    Write-Host "Vendo o documento: $($doc.Name)" -ForegroundColor Cyan -BackgroundColor Black

                    if ([string]::IsNullOrEmpty($doc.Path)) {
                        $nome = "Documento_Salvo_$(Get-Date -Format 'yyyyMMdd_HHmmss').docx"
                        $caminho = Join-Path $env:USERPROFILE "Downloads\$nome"
                        
                        Write-Host "Tentando salvar documento novo em: $caminho" -ForegroundColor Cyan -BackgroundColor Black
                        
                        # Salvando como DOCX (formato 16)
                        $doc.SaveAs2([string]$caminho, 16)
                        
                        Write-Host "Documento novo salvo com sucesso!" -ForegroundColor Green -BackgroundColor Black
                    } else {
                        Write-Host "Salvando alterações no documento existente..." -ForegroundColor Cyan -BackgroundColor Black
                        $doc.Save()
                        Write-Host "Documento existente salvo com sucesso!" -ForegroundColor Green -BackgroundColor Black
                    }
                }
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            } catch { 
                Write-Host "ERRO CRÍTICO NO WORD: $_" -ForegroundColor Red -BackgroundColor Black
            }
        }
    }

    # Executa o bloco acima usando o PowerShell nativo do Windows
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $codigoSalvarOffice

    # --- 2. FECHAR PROCESSOS (ALIASES) ---
    $processos = [ordered]@{
        "msedge" = "Microsoft Edge"
        "msedgewebview2" = "Edge WebView2"
        "msedgedriver" = "Motor Selenium Edge"
        "firefox" = "Mozilla Firefox"
        "chrome" = "Google Chrome"
        "msteams" = "Microsoft Teams 01"
        "Teams" = "Microsoft Teams 02"
        "ms-teams" = "Microsoft Teams 03"
        "mmc" = "Active Directory (MMC)"
        "Microsoft.ConfigurationManagement" = "SCCM"
        "sajapp" = "SAJMP"
        "olk" = "Novo Outlook"
        "FoxitPDFEditor" = "Foxit PDF"
        "Notepad" = "Bloco de Notas"
        "Time" = "Relógio do Windows"
        "RemoteDesktopManager" = "Remote Desktop Manager"
        "WindowsTerminal" = "Windows Terminal"
        #"pwsh" = "PowerShell 7"
    }

    Write-Host "Gerando escudo de proteção (Apenas para Terminais)..." -ForegroundColor DarkGray
    
    # Define a linha de corte: 15 segundos atrás
    $tempoLimite = (Get-Date).AddSeconds(-15)

    foreach ($chave in $processos.Keys) {
        $processosEncontrados = Get-Process -Name $chave -ErrorAction SilentlyContinue
        
        if ($processosEncontrados) {
            foreach ($proc in $processosEncontrados) {
                
                # REGRA DE OURO: Só poupamos se for o Terminal ou PowerShell, E se for novo.
                # Se for Edge, Teams, Firefox, ele mata sem dó, não importa a idade.
                if (($chave -match "WindowsTerminal|pwsh|powershell") -and ($proc.StartTime -gt $tempoLimite)) {
                    Write-Host "  -> Poupando $($processos[$chave]) (Terminal recém-nascido)." -ForegroundColor Cyan
                    continue
                }
                
                Write-Host "Processo $($processos[$chave]) encontrado. Fechando..." -ForegroundColor Yellow -BackgroundColor Black
                Stop-Process -Id $proc.Id -Force
            }
        }
    }

    Write-Host "Apagando a memória de abas do Edge da última sessão..." -ForegroundColor Cyan
    # O Edge guarda o histórico de abas abertas nesta pasta 'Sessions' do perfil Default
    $pastaSessoes = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Sessions"

    if (Test-Path $pastaSessoes) {
      # Apaga todos os arquivos de sessão restaurada sem dó
      Remove-Item -Path "$pastaSessoes\*" -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- 3. EXPLORADOR DE ARQUIVOS (A OPÇÃO NUCLEAR) ---
    Write-Host "Limpando janelas do Explorador de Arquivos (Modo Nuclear)..." -ForegroundColor Cyan -BackgroundColor Black
    
    if (Get-Process -Name explorer -ErrorAction SilentlyContinue) {
        Write-Host "  -> Reiniciando o processo raiz (explorer.exe)..." -ForegroundColor Yellow -BackgroundColor Black
        
        # O tiro de misericórdia. Derruba a interface inteira do Windows e todas as pastas junto.
        Stop-Process -Name explorer -Force
        
        # Dá um tempo para o Windows recarregar a Barra de Tarefas e a Área de Trabalho sozinho
        Write-Host "  -> Aguardando a interface do Windows voltar..." -ForegroundColor DarkGray -BackgroundColor Black
        Start-Sleep -Seconds 3
        
        # Trava de segurança: Se por acaso o Windows 11 for preguiçoso e não recarregar a barra sozinho, o script empurra.
        if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) {
            Start-Process "explorer.exe"
            Start-Sleep -Seconds 2
        }
        
        Write-Host "  -> Explorador reiniciado e memória limpa com sucesso!" -ForegroundColor Green -BackgroundColor Black
    }


    # --- 4. SNIPER ANTI-TELEMETRIA (Otimização de Performance) ---
    Write-Host "Abatendo processos de Telemetria da Microsoft..." -ForegroundColor Cyan -BackgroundColor Black
    
    # 1. Mata o processo que consome Disco e CPU (CompatTelRunner)
    if (Get-Process CompatTelRunner -ErrorAction SilentlyContinue) {
        Write-Host "  -> Matando CompatTelRunner.exe..." -ForegroundColor Yellow -BackgroundColor Black
        Stop-Process -Name "CompatTelRunner" -Force -ErrorAction SilentlyContinue
    }

    # 2. Tenta parar o Serviço de Telemetria (DiagTrack)
    # Nota: Parar serviços geralmente exige que o PowerShell esteja rodando como Administrador.
    # O "SilentlyContinue" garante que, se não estiver como Admin, o script apenas ignora e segue a vida sem dar erro vermelho na tela.
    Stop-Service -Name "DiagTrack" -Force -ErrorAction SilentlyContinue
    
    Write-Host "  -> Telemetria neutralizada (dentro dos privilégios atuais)." -ForegroundColor Green -BackgroundColor Black

    Write-Host "Aguardando estabilização do sistema..." -ForegroundColor Green -BackgroundColor Black
    Start-Sleep -Seconds 2
}

function Matar-Teams {
    Start-Sleep -Seconds 2

    Write-Host "Verificando se o Teams pegou carona no Explorer..." -ForegroundColor Yellow -BackgroundColor Black

    # Colocamos os três nomes possíveis do Teams (O Antigo, o Novo e o processo de Background)
    $fantasmasDoTeams = @("Teams", "ms-teams", "msteams")

    # Damos uma pausa de 3 segundos só para garantir que o Explorer já chamou o intruso
    Start-Sleep -Seconds 3 

    foreach ($fantasma in $fantasmasDoTeams) {
        if (Get-Process -Name $fantasma -ErrorAction SilentlyContinue) {
            Write-Host "  -> Abatendo $fantasma indesejado..." -ForegroundColor Gray
            Stop-Process -Name $fantasma -Force
        }
    }

    Write-Host "Área de estudos 100% blindada contra distrações!" -ForegroundColor Green -BackgroundColor Black

    Start-Sleep -Seconds 2
}

function Abrir-PastasEmAbas {
    param (
        # Aceita um dicionário (chave = Nome pro Log, valor = Caminho da pasta)
        [Parameter(Mandatory=$true)]
        [System.Collections.IDictionary]$Pastas
    )

    if ($Pastas.Count -eq 0) { return }

    Write-Host "Abrindo pastas de trabalho agrupadas em abas..." -ForegroundColor Cyan -BackgroundColor Black
    $wshell = New-Object -ComObject WScript.Shell

    # Extrai as chaves (os nomes/alias) para podermos usar um índice numérico
    $chaves = @($Pastas.Keys)

    # Passo 1: Abre a PRIMEIRA pasta (cria a janela base)
    $primeiraChave = $chaves[0]
    $primeiroCaminho = $Pastas[$primeiraChave]

    Start-Process "explorer.exe" -ArgumentList "`"$primeiroCaminho`""
    Write-Host "  -> $primeiraChave aberto" -ForegroundColor Cyan -BackgroundColor Black
    
    # Pausa generosa para a janela base carregar e ganhar o foco
    Start-Sleep -Seconds 4

    # Passo 2: Abre as pastas seguintes (se existirem)
    if ($chaves.Count -gt 1) {
        for ($i = 1; $i -lt $chaves.Count; $i++) {
            $chaveAtual = $chaves[$i]
            $caminhoAtual = $Pastas[$chaveAtual]

            # Envia Ctrl + T (Nova Aba)
            $wshell.SendKeys("^t")
            Start-Sleep -Seconds 2

            # Envia Ctrl + L (Focar na barra de endereço)
            $wshell.SendKeys("^l")
            Start-Sleep -Milliseconds 600

            # Copia o caminho para a memória (Zero erros de digitação!)
            Set-Clipboard -Value $caminhoAtual

            # Envia Ctrl + V (Colar o caminho)
            $wshell.SendKeys("^v")
            Start-Sleep -Milliseconds 600

            # Envia Enter
            $wshell.SendKeys("~")
            Write-Host "  -> $chaveAtual aberto" -ForegroundColor Cyan -BackgroundColor Black
            
            # Pausa antes da próxima aba
            Start-Sleep -Seconds 2
        }
    }
    
    # Pausa final para estabilização da interface
    Start-Sleep -Seconds 4
}

function Abrir-SitesEdgeLeve {
    param (
        # Aceita um dicionário (chave = Nome pro Log, valor = URL do site)
        [Parameter(Mandatory=$true)]
        [System.Collections.IDictionary]$Sites
    )

    if ($Sites.Count -eq 0) { return }

    Write-Host "Preparando o Edge leve..." -ForegroundColor Cyan -BackgroundColor Black

    # --- O ANTÍDOTO ANTI-CRASH NATIVO ---
    # Vamos direto no arquivo de preferências do seu perfil padrão do Edge
    $caminhoPrefs = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Preferences"
    
    if (Test-Path $caminhoPrefs) {
        Write-Host "  -> Aplicando vacina anti-crash no perfil do Edge..." -ForegroundColor Gray
        
        # Lê o arquivo inteiro como um blocão de texto
        $textoPrefs = Get-Content -Path $caminhoPrefs -Raw
        
        # Se ele achar a marcação de 'Crashed', ele troca para 'Normal'
        if ($textoPrefs -match '"exit_type":"Crashed"') {
            $textoPrefs = $textoPrefs -replace '"exit_type":"Crashed"', '"exit_type":"Normal"'
            
            # Salva de volta garantindo a codificação correta para não corromper o arquivo
            [System.IO.File]::WriteAllText($caminhoPrefs, $textoPrefs, [System.Text.Encoding]::UTF8)
        }
    }

    # --- MONTANDO O COMANDO ---
    $flagsEdge = "--new-window --disable-extensions"
    $urlsParaAbrir = ""

    foreach ($chave in $Sites.Keys) {
        $url = $Sites[$chave]
        # Adiciona a URL com aspas ao redor, separada por espaço
        $urlsParaAbrir += "`"$url`" "
        Write-Host "  -> Carregando: $chave" -ForegroundColor Cyan -BackgroundColor Black
    }

    # --- O TIRO FINAL ---
    $argumentosEdge = "$flagsEdge $urlsParaAbrir"
    Start-Process "msedge.exe" -ArgumentList $argumentosEdge

    Write-Host "Edge iniciado limpo e sem balões!" -ForegroundColor Green -BackgroundColor Black
}

Stop-Transcript