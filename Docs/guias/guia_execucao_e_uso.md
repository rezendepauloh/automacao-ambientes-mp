# 📖 Guia de Execução e Uso dos Serviços C# (AmbientesHub)

> Manual técnico de referência e uso dos serviços nativos em C# (.NET 8) que substituíram integralmente os scripts PowerShell legados (`biblioteca.ps1` e rotinas de ambiente).

---

## 1. Arquitetura de Serviços Nativos

Toda a lógica de automação, manipulação de janelas, limpeza transacional e orquestração de rede reside sob o namespace `AmbientesHub.Services`:

| Serviço C#                       | Responsabilidade Principal                                                                                                             |
| :------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------- |
| **`CleanupService`**             | Encerramento transacional do Office via COM, supressão de telemetria, expurgo de caches e reinício limpo do Explorer.                  |
| **`EnvironmentLauncherService`** | Orquestração dos Modos de Trabalho (Chamados, Contratos, Impressoras, PXE, Estudo).                                                    |
| **`BrowserLauncherService`**     | Abertura validada (HTTP/HTTPS) de URLs no Firefox/Edge e execução segura do WhatsApp PWA.                                              |
| **`CredentialService`**          | Gravação e leitura criptografada de credenciais com Windows DPAPI (`DataProtectionScope.CurrentUser`) com restrição de ACL NTFS (600). |
| **`RemoteToolsService`**         | Disparo sanitizado com `ArgumentList` de RDP (`mstsc.exe`), SCCM Remote Control (`CmRcViewer.exe`) e Explorer remoto.                  |
| **`AdSccmService`**              | Consultas e desprovisionamento seguro de computadores e usuários via LDAP nativo e WMI/CMSite.                                         |
| **`FolderLauncherService`**      | Abertura de múltiplos diretórios com validação de caminho.                                                                             |

---

## 2. Como Executar e Compilar o Projeto

O projeto é compilado e executado através do .NET 8 SDK:

### A. Executar em Modo de Desenvolvimento (Hot Reload / Debug)

```powershell
cd C:\AmbientesVirtuais\AmbientesHub
dotnet run
```

### B. Compilar e Publicar o Executável Autônomo (Portátil)

Gera um executável único com runtime embutido (Single-File), dispensando pré-instalação do .NET na máquina de destino:

```powershell
cd C:\AmbientesVirtuais\AmbientesHub
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o .\bin\Portatil
```

_Binário pronto em:_ `C:\AmbientesVirtuais\AmbientesHub\bin\Portatil\AmbientesHub.exe`

### C. Compilar Publicação Leve (Framework-Dependent)

Gera binário enxuto para parques corporativos que já distribuem o .NET Desktop Runtime 8 via SCCM/Intune:

```powershell
cd C:\AmbientesVirtuais\AmbientesHub
dotnet publish -c Release -r win-x64 --self-contained false -o .\bin\Publish
```

_Binário pronto em:_ `C:\AmbientesVirtuais\AmbientesHub\bin\Publish\AmbientesHub.exe`

---

## 3. Exemplos Práticos de Invocação dos Serviços em C#

### A. Saneamento Completo de Ambiente com Escudo de Proteção

```csharp
var cleanup = new CleanupService();
cleanup.CleanEnvironment(new CleanupOptions
{
    CloseOffice = true,                 // Transacional: Salva Word/Excel antes de fechar
    CloseBrowsers = true,               // Encerra Edge, Chrome e libera memória
    CloseFirefox = false,               // Preserva o Firefox onde o Hub ou PWA roda
    CloseTeams = true,                  // Derruba todas as instâncias do Teams
    CloseAdministrativeTools = true,
    CloseUtilities = true,
    StopTelemetry = true                // Finaliza CompatTelRunner e desativa DiagTrack
});
```

### B. Acionamento de Modo de Trabalho

```csharp
var launcher = new EnvironmentLauncherService();
// Aciona o Modo Chamados (limpa o ambiente previamente e orquestra navegadores/PWA)
launcher.LaunchEnvironment("Chamados", cleanFirst: true);
```

### C. Conexão Remota Sanitizada (Anti Argument Injection)

```csharp
var remote = new RemoteToolsService();
// Dispara RDP validando o hostname contra whitelists rígidas
remote.OpenRdp("ESTACAO-01");

// Dispara visualizador de controle remoto SCCM
remote.OpenSccmRemoteControl("ESTACAO-01");
```

### D. Armazenamento Seguro de Credenciais DPAPI

```csharp
// Salva credencial criptografada na sessão do usuário atual com ACL 600
CredentialService.SaveCredential(@"C:\AmbientesVirtuais\Credenciais\cred_admin.xml", "admin_local", "SenhaForte123!");

// Recupera a credencial de forma segura
var (success, msg, username, password) = CredentialService.ReadPlaintextCredential();
```
