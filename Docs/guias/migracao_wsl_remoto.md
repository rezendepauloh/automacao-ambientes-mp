# Guia de Migração da Bancada WSL (Streamlit + API REST) para Máquina Remota

Este guia descreve como migrar o ambiente WSL da Bancada (Streamlit na porta `8501` e API REST na porta `8510`) para uma máquina Windows remota, mantendo o `AmbientesHub` apontado para o novo host.

---

## 1. Preparação da Máquina Remota

1. **Habilitar o WSL e WinRM persistente (PowerShell Admin na máquina remota):**
   ```powershell
   Enable-PSRemoting -Force
   Set-Service -Name WinRM -StartupType Automatic
   Start-Service -Name WinRM
   Add-LocalGroupMember -SID "S-1-5-32-580" -Member "$env:USERDOMAIN\<usuario>"
   wsl --install
   Restart-Computer -Force
   ```
   _(O SID `S-1-5-32-580` corresponde ao grupo **Remote Management Users / Usuários de Gerenciamento Remoto**, permitindo que o `<usuario>` execute comandos WinRM a partir do `AmbientesHub` sem precisar de senha.)_
2. **Copiar o `.wslconfig` (opcional, para manter limites de memória/CPU):**
   ```powershell
   Copy-Item -Path "C:\Users\<usuario>\.wslconfig" -Destination "\\<HOST_REMOTO>\C$\Users\<usuario>\.wslconfig" -Force
   ```

---

## 2. Exportação e Importação da Distro WSL

1. **Na máquina de origem**, exporte a distro:
   ```powershell
   wsl --export Ubuntu-26.04 C:\temp\ubuntu-bancada.tar
   ```
2. **Envie o arquivo** para a máquina remota (usando conta administrativa com acesso ao `C$`):
   ```powershell
   Copy-Item -Path "C:\temp\ubuntu-bancada.tar" -Destination "\\<HOST_REMOTO>\C$\temp\ubuntu-bancada.tar" -Force
   ```
3. **Na máquina remota**, crie o diretório de destino, importe e defina como padrão:
   ```powershell
   New-Item -ItemType Directory -Path "C:\WSL\Ubuntu-26.04" -Force
   wsl --import Ubuntu-26.04 C:\WSL\Ubuntu-26.04 C:\temp\ubuntu-bancada.tar
   wsl --set-default Ubuntu-26.04
   ```
4. **Limpeza:** após validar com `wsl whoami`, remova `C:\temp\ubuntu-bancada.tar` em ambas as máquinas.

---

## 3. Exposição de Portas na Rede (`8501` e `8510`)

Se a API REST ou o Streamlit estiverem escutando apenas em `127.0.0.1` na máquina remota, libere as portas no Firewall do Windows e configure o `portproxy` (PowerShell Admin na máquina remota):

```powershell
# Liberar portas no Firewall
New-NetFirewallRule -DisplayName "Bancada Streamlit 8501" -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow
New-NetFirewallRule -DisplayName "Bancada API 8510" -Direction Inbound -Protocol TCP -LocalPort 8510 -Action Allow

# Encaminhar 0.0.0.0 -> 127.0.0.1 (persistente após reinício)
netsh interface portproxy add v4tov4 listenaddress=0.0.0.0 listenport=8501 connectaddress=127.0.0.1 connectport=8501
netsh interface portproxy add v4tov4 listenaddress=0.0.0.0 listenport=8510 connectaddress=127.0.0.1 connectport=8510
```

> **Nota:** Regras de firewall e `portproxy` são persistentes após reinício. Caso o WSL/Docker não inicie automaticamente no boot, configure uma tarefa no Agendador de Tarefas do Windows na máquina remota.

---

## 4. Inicialização Automática no Boot (Agendador de Tarefas)

Para que o WSL2 e os containers subam automaticamente no boot (mesmo sem ninguém logado):

1. Na máquina remota, abra `secpol.msc` → **Políticas Locais** → **Atribuição de Direitos de Usuário** → **Fazer logon como trabalho em lotes** (`SeBatchLogonRight`) e adicione `<DOMINIO>\<usuario>`.
2. Registre a tarefa agendada com `exec sleep infinity` para manter a VM do WSL2 ativa em segundo plano:

```powershell
$action   = New-ScheduledTaskAction -Execute "wsl.exe" -Argument "-d Ubuntu-26.04 -- bash -lc `"cd /home/<usuario>/projetos/bancada && docker compose up -d && exec sleep infinity`""
$trigger  = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName "BancadaStreamlitAutoStart" `
    -Action $action -Trigger $trigger -Settings $settings `
    -User "$env:USERDOMAIN\<usuario>" -Password "<senha>" -RunLevel Highest -Force
```

---

## 5. Configuração do `AmbientesHub`

No `.env` local (ou pela aba **Configurações & Credenciais** do `AmbientesHub`), aponte as URLs para o host remoto:

```env
URL_BANCADA_LOCAL=http://<HOST_REMOTO>:8501/
BANCADA_API_URL=http://<HOST_REMOTO>:8510
```

Quando `URL_BANCADA_LOCAL` aponta para um host remoto (diferente de `localhost`/`127.0.0.1`), o `BancadaDockerService` roteia automaticamente os botões **Iniciar**, **Parar**, **Reiniciar**, **Reconstruir** e **Conectar Logs** via WinRM (`Invoke-Command`).
