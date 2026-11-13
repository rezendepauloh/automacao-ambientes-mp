# 📖 Guia de Operação do Ambientes Hub (WPF-UI / C#)

O **Ambientes Hub** é a interface desktop nativa Windows (Fluent Design / Mica) desenvolvida em **C# (.NET 8 com WPF-UI)** com o objetivo de centralizar a alternância de ambientes corporativos de trabalho, fornecer launcher com proteção contra injeção para conexões remotas (RDP e SCCM) e executar rotinas profundas de saneamento de memória, desprovisionamento no AD/SCCM e cofre de senhas DPAPI.

---

## 1. Como Compilar e Executar a Aplicação

Abra o terminal do PowerShell na pasta do projeto e execute:

```powershell
# Navegar até o diretório da aplicação
cd C:\AmbientesVirtuais\AmbientesHub

# Compilar e executar em modo de desenvolvimento
dotnet run
```

Também é possível disparar o binário executável gerado:

```powershell
Start-Process "C:\AmbientesVirtuais\AmbientesHub\bin\Debug\net8.0-windows\AmbientesHub.exe"
```

---

## 2. Como Gerar o Executável Pronto para Uso (`.exe`)

Dependendo do objetivo de distribuição no parque, você pode gerar a aplicação em dois formatos:

### Opção A: Executável Autocontido Portátil (Roda em Qualquer PC sem .NET Instalado)

Embute o runtime do .NET 8 e todas as dependências em um único executável portátil (`Single-File`). Não requer que o computador de destino tenha o SDK ou Runtime .NET instalados:

```powershell
cd C:\AmbientesVirtuais\AmbientesHub

# Gera o executável único autocontido
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o .\bin\Portatil
```

_Executável autônomo disponível em:_
`C:\AmbientesVirtuais\AmbientesHub\bin\Portatil\AmbientesHub.exe`

### Opção B: Framework-Dependent (Executável Leve)

Para máquinas corporativas que já possuem o _.NET Desktop Runtime 8.0_ instalado (geralmente distribuído via SCCM ou Windows Update):

```powershell
cd C:\AmbientesVirtuais\AmbientesHub

# Gera a publicação enxuta
dotnet publish -c Release -r win-x64 --self-contained false -o .\bin\Publish
```

_Executável leve disponível em:_
`C:\AmbientesVirtuais\AmbientesHub\bin\Publish\AmbientesHub.exe`

---

## 3. Módulos e Funcionalidades

### A. Aba "Ambientes"

- Apresenta cartões interativos para cada um dos ambientes corporativos:
  - **🛠️ Modo Chamados:** Abre os portais OTRS/CitSmart e pastas de triagem.
  - **🛠️ Modo Contratos:** Carrega documentações contratuais e fluxos.
  - **🛠️ Modo Impressoras:** Carrega atalhos e ferramentas de gestão de filas.
  - **🛠️ Modo PXE:** Ambiente preparado para máquinas de bancada e formatação.
  - **🛠️ Modo Estudo:** Ambiente focado sem distrações.

### B. Aba "Acesso Remoto"

- **Entrada única:** Permite digitar o nome de rede (ex: `DOMINIO-80600`) ou endereço IP.
- **Botão "SCCM Remoto":** Invoca diretamente o `CmRcViewer.exe` com a flag do site corporativo sem a necessidade de abrir o Console do MECM/SCCM.
- **Botão "RDP Padrão":** Dispara o `mstsc.exe` já configurado com a máquina alvo.

### C. Aba "Limpeza AD & SCCM"

- Realiza desprovisionamento transacional de estações de trabalho de forma 100% nativa em C# (`AdSccmService`).
- Remove contas de computador no Active Directory via LDAP nativo (`DeleteTree`) e no Configuration Manager via WMI com PacketPrivacy ou provedor CMSite PowerShell sanitizado.
- Exige validação estrita do hostname (anti-wildcard) e caixa de diálogo de confirmação crítica antes de qualquer exclusão.

### D. Aba "Performance & Limpeza"

- Dispara o saneamento de recursos via `CleanupService`:
  - Salva documentos abertos do Word e Excel de forma transacional via COM antes de encerrá-los, liberando ponteiros com `Marshal.ReleaseComObject`.
  - Encerra navegadores pesados (Chrome, Edge) e Teams liberando memória.
  - Imuniza o próprio Hub, seu processo pai e janelas rotuladas com `"🛠️ Modo"` através do escudo de processos.
  - Modo Nuclear do Explorer para zerar caches bloqueados de ícones e arquivos temporários.

### E. Aba "Otimizações de Registro (Tweaks & Rollback)" (Item 11)

- **Ajustes Cirúrgicos de Performance:** Lista curada de tweaks no Windows Registry focados em redução de latência (`MenuShowDelay`), eliminação do atraso de boot (`StartupDelayInMSec`), priorização de foreground em multitarefa (`SystemResponsiveness`), desativação de throttling de rede TCP (`NetworkThrottlingIndex`) e supressão de bloatware em segundo plano.
- **Backup Obrigatório Pré-Gravação:** Gera automaticamente arquivos `.reg` (compatíveis com o Editor de Registro do Windows) e snapshot JSON em `C:\AmbientesVirtuais\Backups\RegBackup_yyyyMMdd_HHmmss.reg` antes de qualquer alteração.
- **Restauração em 1 Clique (Rollback):** Permite reverter valores anteriores ou padrões da Microsoft individualmente ou em lote.
- **Segurança e Privilégios:** Separação visual clara entre chaves em espaço de usuário (`HKCU` - seguras e sem elevação) e chaves do sistema (`HKLM` - com badge indicando necessidade de execução como Administrador).

### F. Aba "Logs da Aplicação" (Item 12)

- **Observabilidade em Tempo Real:** Visualizador integrado dos logs gravados em `AmbientesHub\logs\app.log`.
- **Destaque Sintático Colorido:** Formatação rica no console WPF com identificação visual instantânea por cores para carimbos de data/hora, níveis (`INFO`, `WARNING`, `DANGER`, `SUCCESS`) e classes chamadoras.
- **Filtros Dinâmicos:** Filtro por nível e busca por texto em tempo real com rolagem automática.

### G. Aba "Configurações & Credenciais" (Item 13 - Último da Barra Lateral)

- **Centralização Unificada:** Módulo posicionado estrategicamente como o último item da navegação, consolidando todas as variáveis de ambiente (`.env`), seleção tipográfica, arranque com o Windows e cofres seguros DPAPI.

### H. Inicialização com Windows & Bandeja do Sistema

- **Iniciar com o Windows:** Alternador (ToggleSwitch) nas Configurações para habilitar ou desabilitar o arranque automático do aplicativo após login (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`).
- **Atalho no Menu Iniciar:** Alternador nas Configurações que cria e remove o atalho `Ambientes Hub.lnk` na pasta de Programas do Menu Iniciar (`AppData\Roaming\Microsoft\Windows\Start Menu\Programs`) com metadados e ícone corporativo.
- **Minimizar para a Bandeja (Tray Icon):** Mantém o aplicativo acessível próximo ao relógio. Suporta duplo clique para restaurar a janela e menu de contexto com opções de alternância entre Modo Comum e Modo Administrador (com badge indicativo no cabeçalho da janela).
- **Identidade Visual:** Exibe o ícone e a descrição oficiais no Gerenciador de Tarefas do Windows.
