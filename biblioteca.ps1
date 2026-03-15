# ==========================================
# BIBLIOTECA DE FUNÇÕES - AUTOMAÇÃO MP
# ==========================================

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
        "chrome" = "Google Chrome"
        "msteams" = "Microsoft Teams"
        "mmc" = "Active Directory (MMC)"
        "Microsoft.ConfigurationManagement" = "SCCM"
        "saj" = "SAJMP"
        "olk" = "Novo Outlook"
        "FoxitPDFEditor" = "Foxit PDF"
        "Notepad" = "Bloco de Notas"
        "Time" = "Relógio do Windows"
    }

    foreach ($proc in $processos.Keys) {
        $nomeElegante = $processos[$proc]
        if (Get-Process -Name $proc -ErrorAction SilentlyContinue) {
            Write-Host "Processo $nomeElegante encontrado. Fechando..." -ForegroundColor Yellow -BackgroundColor Black
            Stop-Process -Name $proc -Force
        } else {
            Write-Host "Processo $nomeElegante não está aberto. Pulando..." -ForegroundColor DarkGray -BackgroundColor Black
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