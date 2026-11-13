# 🔒 Arquitetura de Segurança & Hardening (SecOps)

> Documento canônico de engenharia de segurança do **AmbientesHub** e dos scripts legados em `Scripts/`.

---

## 1. Princípios de Projeto

1. **Fail-Closed por Padrão**: nenhuma rotina de rede, túnel WSL/Docker, exclusão em AD/SCCM, abertura de ferramenta remota ou escrita de cofre DPAPI é iniciada quando variáveis críticas estão ausentes, vazias ou fora da whitelist.
2. **Validação Centralizada por Whitelist (`InputValidator`)**: toda entrada externa (TextBoxes WPF, `.env`, retorno WMI remoto) é validada com expressões regulares compiladas, com `RegexTimeout` de 250 ms (anti-ReDoS) e verificação de cobertura integral (`m.Index == 0 && m.Length == value.Length`, impedindo bypass por `\n` final).
3. **Passagem Estruturada de Argumentos**: uso de `ProcessStartInfo.ArgumentList` em vez de interpolação manual de strings em `Arguments` para `mstsc.exe`, `CmRcViewer.exe`, `explorer.exe` e `firefox.exe`.
4. **Zero Segredos ou Dados Pessoais no Repositório**: nomes de usuários reais, caminhos pessoais de WSL/OneDrive, códigos de site e hostnames internos vivem exclusivamente em `AmbientesHub/.env` (ignorado pelo Git) e nos cofres DPAPI em `Credenciais/` (ignorados pelo Git).
5. **Defesa em Profundidade em Parsers, Downloads e Sistema de Arquivos**:
   - XML Clixml lido com `DtdProcessing.Prohibit`, `XmlResolver = null` e teto de 64 KB (anti-XXE / Billion Laughs).
   - Downloads MSI restritos a `https://`, com streaming e teto rígido de 300 MB em arquivo temporário com `Guid` único e `FileMode.CreateNew`.
   - Limpeza recursiva ignora `FileAttributes.ReparsePoint` para impedir *Junction/Symlink Traversal*.
   - ACL de `.env` e `Credenciais/*.xml` endurecida via `ConfigService.HardenFilePermissions` (herança desabilitada; acesso restrito ao usuário atual, `BUILTIN\Administrators` e `NT AUTHORITY\SYSTEM` — equivalente a `chmod 600`).

---

## 2. Controles Implementados por Camada

| Camada / Serviço | Controle de Segurança |
| :--- | :--- |
| [InputValidator.cs](../../AmbientesHub/Services/InputValidator.cs) | Whitelists (`IsValidHostName`, `IsValidSearchTerm`, `IsValidUserName`, `IsValidWslDistro`, `IsValidWslPath`, `IsValidSiteCode`, `IsSafeHttpUrl`, `IsSafeLocalExecutablePath`, `IsValidEnvKey`) + escapadores `EscapeLdap` (RFC 4515), `EscapeWql`, `EscapePowerShellSingleQuoted`, `SanitizeEnvValue`, `SanitizeLogText`. |
| [AdSccmService.cs](../../AmbientesHub/Services/AdSccmService.cs) | Exclusão (`CleanComputer`) exige hostname exato e consulta igualdade exata em LDAP (`cn={name}`), WQL (`NetbiosName = '{name}'`) e PowerShell CMSite (`-eq`), eliminando exclusão em massa por wildcard. Consultas escapam LDAP/WQL/PS. |
| [RemoteToolsService.cs](../../AmbientesHub/Services/RemoteToolsService.cs) | `RequireValidHost` + `ProcessStartInfo.ArgumentList` em RDP, CmRcViewer e Explorer; `CMRCVIEWER_PATH` validado com `IsSafeLocalExecutablePath`. |
| [RemoteManagementHelper.cs](../../AmbientesHub/Services/RemoteManagementHelper.cs) | `PingHost`, `CreateScope`, `EnsureAdminShareConnected` e `InvokeCommandOnTarget` validam o host e rejeitam `\r`/`\n`/`\0` antes de abrir DCOM/WMI/SMB. |
| [RemoteUserProfileService.cs](../../AmbientesHub/Services/RemoteUserProfileService.cs) | Whitelist de host e usuário antes de `net user /delete`; validação estrita `^[A-Za-z]:\\Users\\[A-Za-z0-9._\- ]{1,128}\z` antes do fallback `rmdir /S /Q`. |
| [RemoteMaintenanceService.cs](../../AmbientesHub/Services/RemoteMaintenanceService.cs) | `CleanDirectoryContentsSafe` detecta `FileAttributes.ReparsePoint` e remove apenas o elo sem atravessar junctions/symlinks. |
| [BancadaDockerService.cs](../../AmbientesHub/Services/BancadaDockerService.cs) | Defaults pessoais removidos; `ValidateTarget` exige `BANCADA_WSL_DISTRO` e `BANCADA_WSL_DIR` válidos no `.env` antes de chamar `wsl.exe`. |
| [BrowserLauncherService.cs](../../AmbientesHub/Services/BrowserLauncherService.cs) | `OpenUrlsInFirefox` aceita apenas `http://`/`https://` via `ArgumentList`; `OpenWhatsappPwa` valida `WHATSAPP_PWA_ARGS` e não registra argumentos de perfil no log. |
| [ConfigService.cs](../../AmbientesHub/Services/ConfigService.cs) | Valida chaves (`IsValidEnvKey`), remove CR/LF/NUL dos valores (`SanitizeEnvValue`) e aplica ACL restrita (`HardenFilePermissions`) ao carregar e salvar `.env`. |
| [CredentialService.cs](../../AmbientesHub/Services/CredentialService.cs) | Fail-closed para usuário/senha vazios, leitura XML anti-XXE com limite de 64 KB, DPAPI `CurrentUser` e ACL restrita no arquivo `.xml` gerado. |
| [DeviceAnalyzerService.cs](../../AmbientesHub/Services/DeviceAnalyzerService.cs) | Inventário somente-leitura (não altera WinRM/Firewall/TrustedHosts no host alvo) + `RedactSensitiveCommandLine` mascara senhas/tokens em linhas de comando coletadas. |
| [SoftwareProvisioningService.cs](../../AmbientesHub/Services/SoftwareProvisioningService.cs) | Validação de `SCCM_SITE_SERVER`/`SCCM_SITE_CODE`, arquivos temporários com `Guid` único, download MSI somente HTTPS com limite de 300 MB. |
| [AppLoggerService.cs](../../AmbientesHub/Services/AppLoggerService.cs) | `InputValidator.SanitizeLogText` neutraliza quebras de linha em todas as entradas de log (anti *Log Forging*). |
| `Scripts/*.ps1` | Dot-sourcing ancorado em `$PSScriptRoot` com trava fail-closed se `config.ps1` não existir; remoção de `runas.exe /savecred` em favor de `Import-Clixml` + `Start-Process -Credential`. |

---

## 3. Verificação Contínua na Suíte de Testes

O arquivo [SecurityAuditTests.cs](../../tests/Integration/SecurityAuditTests.cs) é executado a cada `dotnet test` e falha o build caso:
- Qualquer arquivo rastreado pelo Git (`git ls-files`) contenha chaves privadas, tokens de API, `ConvertTo-SecureString -AsPlainText` ou identificadores pessoais/organizacionais proibidos;
- Arquivos sensíveis (`.env`, `Credenciais/`, `temp/`, `config.ps1`, `login.ps1`, `*.log`) passem a ser rastreados;
- `.env.example` e `AmbientesHub/.env.example` divirjam entre si ou deixem de documentar qualquer chave consumida pelo `ConfigService`;
- `HardenFilePermissions` mantenha herança de ACL ou permita acesso a `Authenticated Users` / `BUILTIN\Users`;
- As travas *fail-closed* de validação de host, usuário, argumentos de navegador, XML/XXE ou credenciais deixem de rejeitar entradas inválidas.

