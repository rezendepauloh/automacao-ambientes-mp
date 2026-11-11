# 📖 Catálogo de Funções — biblioteca.ps1

Detalhamento técnico das assinaturas, parâmetros e funcionamento interno de cada função disponível no módulo.

---

## 1. `Limpar-Ambiente`

- **Objetivo:** Orquestrar o ciclo completo de saneamento e recuperação de recursos da estação.
- **Assinatura:** `Limpar-Ambiente` (sem parâmetros).
- **Etapas Executadas:**
  1. **Encerramento Transacional Office:** Conecta via COM ao Excel e Word em execução, salva documentos novos com carimbo temporal em `Downloads\` e documentos existentes in-place, antes de invocar `$app.Quit()`.
  2. **Varredura e Escudo de Processos:** Encerra navegadores e ferramentas pesadas listadas em `$processos`, ignorando processos com o PID atual, PID pai ou que contenham `"🛠️ Modo"` ou o título do terminal no título da janela.
  3. **Limpeza de Sessão do Edge:** Exclui o diretório `%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Sessions` para evitar restauração automática indesejada de dezenas de abas.
  4. **Modo Nuclear do Explorer:** Reinicia `explorer.exe` para liberar locks de arquivos e handles residuais do shell.
  5. **Supressão de Telemetria:** Interrompe `CompatTelRunner.exe` e envia sinal de parada ao serviço `DiagTrack`.

---

## 2. `Matar-Teams`

- **Objetivo:** Forçar a finalização de todas as variantes de executáveis do Microsoft Teams.
- **Assinatura:** `Matar-Teams` (sem parâmetros).
- **Alvos:** `"Teams"`, `"ms-teams"`, `"msteams"`.

---

## 3. `Abrir-PastasEmAbas`

- **Objetivo:** Abrir múltiplos diretórios organizados em abas dentro de uma única janela do Explorador de Arquivos do Windows.
- **Parâmetros:**
  - `-Pastas` (`[System.Collections.IDictionary]` - Obrigatório): Tabela com rótulo amigável na chave e caminho físico no valor.
- **Mecanismo:** Cria a primeira janela e utiliza automação via `WScript.Shell` enviando sequências de teclas (`Ctrl+T`, `Ctrl+L`, clipboard e `Enter`).

---

## 4. `Abrir-SitesEdgeLeve`

- **Objetivo:** Inicializar o navegador Microsoft Edge sem sobrecarga de extensões e prevenindo a mensagem de "Restaurar páginas após travamento".
- **Parâmetros:**
  - `-Sites` (`[System.Collections.IDictionary]` - Obrigatório): Tabela com rótulo na chave e URL no valor.
- **Mecanismo:** Modifica o JSON de preferências do perfil (`"exit_type":"Crashed"` -> `"exit_type":"Normal"`) e dispara `msedge.exe` com as flags `--new-window --disable-extensions`.

---

## 5. `Remover-CredenciaisMemoria`

- **Objetivo:** Evitar exposição de credenciais em memória RAM e reduzir vetores de inspeção em memória por ferramentas de monitoramento.
- **Parâmetros:**
  - `-Variaveis` (`[string[]]` - Opcional): Lista de nomes de variáveis a serem removidas do escopo pai (`-Scope 1`).
- **Mecanismo:** Executa `Remove-Variable` e força o acionamento de `[System.GC]::Collect()` e `[System.GC]::WaitForPendingFinalizers()`.

