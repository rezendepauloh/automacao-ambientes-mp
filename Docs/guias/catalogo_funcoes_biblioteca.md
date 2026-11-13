# 📖 Catálogo de Serviços Nativos em C# (AmbientesHub)

> Documentação técnica detalhada das classes de serviço que encapsulam e modernizam todas as operações do ecossistema `AmbientesHub`.

---

## 1. `CleanupService` (Saneamento & Escudo de Processos)

- **Objetivo:** Orquestrar o ciclo completo de saneamento, recuperação de memória e encerramento de processos com escudo de proteção.
- **Assinatura Principal:** `CleanEnvironment(CleanupOptions options, Action<string>? onLog = null)`
- **Mecanismos Internos:**
  1. **Encerramento Transacional do Office (COM):** Utiliza COM API via `Marshal.GetActiveObject` para Word e Excel, salva alterações em documentos abertos e libera os ponteiros com `Marshal.ReleaseComObject`.
  2. **Escudo de Proteção:** Ignora e imuniza contra encerramento o PID atual do aplicativo, o PID do processo pai e qualquer janela cujo título contenha `"🛠️ Modo"`.
  3. **Modo Nuclear do Explorer:** Finaliza a árvore de processos do `explorer.exe`, expurga travas residuais de arquivo e reinicia o shell nativo do Windows.
  4. **Supressão de Telemetria:** Interrompe `CompatTelRunner` e para o serviço `DiagTrack`.

---

## 2. `EnvironmentLauncherService` (Orquestrador de Modos)

- **Objetivo:** Realizar a transição segura para os diferentes modos operacionais de trabalho.
- **Assinatura Principal:** `LaunchEnvironment(string modeName, bool cleanFirst = true)`
- **Modos Suportados:**
  - **Chamados:** Dispara WhatsApp PWA, portais corporativos (OTRS, CitSmart) e Bancada Docker local.
  - **Contratos:** Inicializa ferramentas e portais de fiscalização contratual.
  - **Impressoras:** Abre utilitários de gestão de parque e bilhetagem de impressão.
  - **PXE:** Configura o ambiente para triagem de máquinas e bancada de formatação.
  - **Estudo:** Modo foco sem notificações e com supressão de distrações.

---

## 3. `BrowserLauncherService` (Orquestração de Navegadores)

- **Objetivo:** Abrir links e portais de forma segura, com validação contra Argument Injection e URL Schemas maliciosos.
- **Métodos:**
  - `OpenUrlsInFirefox(IEnumerable<string> urls)`: Abre abas em lote validando cada link via `InputValidator.IsSafeHttpUrl` e utilizando `ProcessStartInfo.ArgumentList`.
  - `OpenWhatsappPwa()`: Inicia o WhatsApp em janela de aplicativo PWA via Firefox com argumentos higienizados contra injeção de shell.
  - `ClearFirefoxSession()`: Expurga arquivos de sessão travados (`sessionstore.jsonlz4`).

---

## 4. `RemoteToolsService` (Acesso Remoto Rápido)

- **Objetivo:** Conectar instantaneamente a computadores remotos com validação rígida de parâmetros.
- **Métodos:**
  - `OpenRdp(string computerName)`: Dispara `mstsc.exe /v:<host>` usando `ArgumentList`.
  - `OpenSccmRemoteControl(string computerName, string? siteServer)`: Localiza e executa o `CmRcViewer.exe` corporativo.
  - `OpenRemoteExplorer(string computerName)`: Conecta e abre o compartilhamento administrativo `\\host\C$` no Windows Explorer.

---

## 5. `CredentialService` (Cofre Seguro DPAPI)

- **Objetivo:** Armazenar e recuperar credenciais criptografadas atreladas à sessão do operador atual.
- **Métodos:**
  - `SaveCredential(string xmlPath, string username, string plainPassword)`: Criptografa a senha com DPAPI (`DataProtectionScope.CurrentUser`), grava em Clixml compatível e aplica permissões de arquivo restritas (ACL 600).
  - `ReadPlaintextCredential(string xmlPath)`: Lê e valida o cofre contra ataques de XML DoS / XXE (limite de 64 KB e `DtdProcessing.Prohibit`), descriptografando a senha em memória.

---

## 6. `AppLoggerService` (Logs e Auditoria do Sistema)

- **Objetivo:** Gravar e recuperar logs estruturados de forma thread-safe com rotatividade automática e notificações em tempo real para a interface gráfica.
- **Propriedades e Eventos:**
  - `LogDirectoryPath`: Retorna o caminho do diretório de logs (`C:\AmbientesVirtuais\AmbientesHub\logs`).
  - `CurrentLogFilePath`: Retorna o caminho do arquivo ativo (`app.log`).
  - `OnLogLineWritten`: Evento disparado a cada linha gravada para sincronização em tempo real com a UI.
- **Métodos:**
  - `Log(LogLevel level, string message, string? caller = null)`: Grava com carimbo de tempo, nível (`INFO`, `WARNING`, `DANGER`, `SUCCESS`), sanitização de quebras de linha (`InputValidator.SanitizeLogText`) e rotação ao atingir 3 MB.
  - `ReadRecentLogs(int maxLines = 1000)`: Lê as últimas linhas do arquivo `app.log` usando `FileShare.ReadWrite` com tratamento seguro de concorrência.
  - Métodos utilitários: `Info()`, `Warning()`, `Danger()`, `Success()`.

---

## 7. `RegistryOptimizerService` (Otimizações de Registro & Rollback)

- **Objetivo:** Aplicar ajustes no Registro do Windows com segurança cirúrgica, backup prévio em formato `.reg` e rollback em 1 clique.
- **Métodos Principais:**
  - `GetDefaultTweaks()`: Retorna o catálogo de 12 otimizações curadas nas categorias: Responsividade & UI, Inicialização & Encerramento, Sistema & Background e Rede & Multimídia.
  - `ScanTweak(RegistryTweakItem tweak)` / `ScanAll()`: Inspeciona as chaves no registro em tempo real sem alterar dados, comparando os valores atuais contra os padrões e os valores recomendados.
  - `CreateBackup(IEnumerable<RegistryTweakItem> tweaks, string? customPath = null)`: Gera arquivo `.reg` (Windows Registry Editor Version 5.00) e snapshot JSON com permissões ACL 600 em `C:\AmbientesVirtuais\Backups`.
  - `ApplyTweak(RegistryTweakItem tweak)` / `ApplyBatch()`: Salva o estado anterior para rollback e aplica o valor otimizado, validando privilégios de Administrador em chaves `HKLM`.
  - `RollbackTweak(RegistryTweakItem tweak)` / `RollbackBatch()`: Reverte o valor ao estado original gravado no ponto de restauração ou ao padrão Windows.
