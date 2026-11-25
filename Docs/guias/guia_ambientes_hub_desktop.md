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

### E. Aba "Backup & Restore (USMT)" (Item 10)

- **Migração Local ou Remota (`UsmtBackupService`):** Permite fazer backup (`scanstate.exe`) e restauração (`loadstate.exe`) de perfis de usuários tanto na máquina local (`.`) quanto em estações remotas via compartilhamento administrativo + WMI (`Win32_Process.Create`) com streaming do log em tempo real.
- **Seleção Individual & Segurança Total:** Lista os perfis presentes na máquina para marcação individual por checkbox (`MigApp.xml`, `MigDocs.xml`, `MigUser.xml`, compactação e criptografia opcional com `/encrypt` / `/key`), **sem nunca excluir nenhum perfil de usuário**.
- **Detecção & Download Automático do USMT:** Localiza `USMT_PATH` (ADK ou pasta local) e oferece botão de download/instalação automatizada do recurso `OptionId.UserStateMigrationTool` do Windows ADK oficial, preenchendo `USMT_PATH` automaticamente no `.env`.

### F. Aba "Bancada Docker (8501)" (Item 11)

- **Controle da Stack WSL/Docker (`BancadaDockerService`):** Inicia, para (`down`), reinicia e reconstrói (`build --no-cache`) containers com `PYTHONUNBUFFERED=1` e reconexão automática aos logs do `docker compose`, além de disparar notificações toast do Windows ao concluir tarefas em segundo plano.

### G. Aba "Mapa de Rota & Ping" (Item 19)

- **Traceroute Visual Nativo (`NetworkPathService` + `NetworkPathView`):** Traça a árvore L3 da estação local até o destino em tempo real usando `Canvas` WPF + curvas Bézier cúbicas (orientação Horizontal ou Vertical, Zoom e Pan).
- **Ping Contínuo & Telemetria:** Monitoramento contínuo com contagem de enviados/recebidos/perda, latência (último/mín/méd/máx) e gráfico de barras das últimas 24 amostras.
- **Enriquecimento Multicamada:**
  - **MAC Address:** Interface local (`NetworkInterface`), tabela ARP local (`SendARP` via `iphlpapi.dll` para Gateway/mesma sub-rede) e fallback SCCM.
  - **Usuário Conectado:** Sessão ativa via WMI (`Win32_ComputerSystem`) com fallback em segundo plano para AD/SCCM (`CurrentLogonUser`/`PrimaryUser`).
  - **Localização & Switches:** Mapeamento local de faixas CIDR (`subnet_map.json`), OU do Active Directory e integração opcional com **Zabbix API** (`ZABBIX_API_URL` e `ZABBIX_API_TOKEN` no `.env`).
- **Relatório HTML5 Responsivo:** Gera relatório adaptado para desktop e celular com datas no padrão brasileiro (`dd/MM/yyyy HH:mm:ss`) salvo no diretório de `CHAMADOS_REPORTS_DIR`.
- **Ações Rápidas, Portas TCP & Wake-on-LAN:** No painel inspetor lateral, permite abrir **SCCM**, **RDP** ou **`C$`**, testar em 1 segundo as portas **3389 (RDP)**, **445 (SMB)**, **135 (WMI)** e **5985 (WinRM)**, e enviar **Magic Packet WoL** quando houver MAC Address disponível.

### H. Aba "Monitor Zabbix" (Item 20)

- **Visão de Comarcas, Prédios & Switches (`ZabbixMonitorView`):** Lista todos os hosts monitorados no Zabbix (`host.get`) com KPIs (`Total`, `Online`, `Offline`, `Exibindo Agora`), paginação (25 itens por página) e filtros por cidade/prédio (`hostgroups`), status (`Todos`, `Offline`, `Online`) e busca textual.
- **Detalhes & Integração:** O botão **ℹ️ Detalhes** carrega inventário, alertas ativos e métricas recentes em tabela organizada; o botão **🌐 Traçar Rota** abre o equipamento direto no *Mapa de Rota & Ping*.

### I. Aba "Vizinhos de Rede" (Item 21)

- **Mapeamento de Sub-rede & Tabela ARP (`SubnetNeighborsService` + `SubnetNeighborsView`):** A partir de um hostname ou IP de referência, mapeia a sub-rede `/24` em paralelo (Ping + PTR + ARP local), tenta coletar a tabela ARP remota via `MSFT_NetNeighbor` (`root\StandardCimv2`) e opcionalmente complementa com máquinas da mesma OU no AD.
- **Árvore Bézier & Inspetor:** Exibe o nó raiz (sub-rede/gateway/switch Zabbix) conectado por curvas Bézier aos vizinhos encontrados; ao clicar em qualquer vizinho, carrega usuário conectado e MAC em background e libera botões de 1 clique (**SCCM**, **RDP**, **C$**, **Traçar Rota** e **Portas TCP**).

### J. Aba "Servidores & RDP" (Item 22)

- **Catálogo Criptografado & Compatível com RDM (`RdpCatalogService` + `RdpServersView`):** Importa e exporta JSON no formato do Devolutions Remote Desktop Manager (`Connections`) e persiste os servidores criptografados com **DPAPI (`CurrentUser`) + ACL 600** em `AmbientesHub\rdp_servers.dat`.
- **Editor Segura® PAM & Conexão Dual:** Lista servidores em cards paginados com busca e filtro (`Segura®` vs `Direto`), permite editar e aplicar o padrão Segura® (`usuario[dominio\adm@SERVIDOR]`), e conectar tanto **em abas embutidas** dentro do Hub (`EmbeddedRdpHost` via `mstscax.dll`) quanto pelo **`mstsc.exe` do Windows**.
- **Perfis de Qualidade, Resolução Dinâmica & Tela Cheia por Monitor:** Oferece 3 perfis (`⚖️ Balanceado 24-bit`, `🚀 Melhor Estabilidade 16-bit`, `✨ Melhor Visualização 32-bit`), adaptação dinâmica da resolução e DPI ao painel/monitor atual (`UpdateSessionDisplaySettings` + `SmartSizing`) e botão **`⛶ Tela Cheia`** que abre em resolução nativa no monitor onde a janela está posicionada.

### K. Aba "Print Server & Portas" (Item 23)

- **Automação de Portas PaperCut no Print Management (`PrinterMigrationService` + `PrinterMigrationView`):** Verifica a conectividade da impressora no novo IP (Ping ICMP + porta RAW `9100` + interface web + MAC ARP) e automatiza o fluxo no servidor de impressão:
  - Se a impressora já estiver com a porta `PAPERCUT_<NOVO_IP>`, interrompe a alteração e oferece envio de **Página de Teste** + abertura da **Fila de Impressão**.
  - Se estiver com IP antigo, cria a porta `PAPERCUT_<NOVO_IP>` (`PaperCut TCP/IP Port` na porta `9100`) caso ainda não exista e atualiza a impressora (`Set-Printer`).
  - Mantém estruturada a função `Provision-NewPrinterQueue` para criação de filas novas e o atalho direto para o **PaperCut Admin** (`URL_PAPERCUT`).
  - **Execução em ambientes com PAM / 2FA:** Quando a conta administrativa exige autenticação intermediada por proxy PAM com 2FA (ex.: formato `usuario[dominio\adm@SERVIDOR]`), o WinRM direto (`Invoke-Command`) não autentica; nesse caso, utilize **Copiar Script PowerShell**, abra o servidor na aba **Servidores & RDP** e cole no **PowerShell (Admin)** dentro da sessão.

### G. Aba "Otimizações de Registro (Tweaks & Rollback)" (Item 12)

- **Ajustes Cirúrgicos de Performance:** Lista curada de tweaks no Windows Registry focados em redução de latência (`MenuShowDelay`), eliminação do atraso de boot (`StartupDelayInMSec`), priorização de foreground em multitarefa (`SystemResponsiveness`), desativação de throttling de rede TCP (`NetworkThrottlingIndex`) e supressão de bloatware em segundo plano.
- **Backup Obrigatório Pré-Gravação:** Gera automaticamente arquivos `.reg` (compatíveis com o Editor de Registro do Windows) e snapshot JSON em `C:\AmbientesVirtuais\Backups\RegBackup_yyyyMMdd_HHmmss.reg` antes de qualquer alteração.
- **Restauração em 1 Clique (Rollback):** Permite reverter valores anteriores ou padrões da Microsoft individualmente ou em lote.
- **Segurança e Privilégios:** Separação visual clara entre chaves em espaço de usuário (`HKCU` - seguras e sem elevação) e chaves do sistema (`HKLM` - com badge indicando necessidade de execução como Administrador).

### H. Aba "Documentação" (Acima de Logs da Aplicação)

- **Leitor Markdown Integrado (`DocumentationService` + `DocumentationView`):** Indexa automaticamente todos os arquivos `.md` de `Docs/` agrupados por categoria (`Índice Geral`, `Guias Operacionais`, `Arquitetura & SecOps`, `Issues & Homologação`) com proteção estrita contra *path traversal* e *symlinks*.
- **Busca Instantânea & Navegação Hipertexto:** Filtra por título ou conteúdo completo dos manuais, suporta links clicáveis entre documentos `.md` com histórico (`◀` / `▶`), controle de zoom tipográfico (`A-` / `A+`) e botão **📋 Copiar** individual em cada bloco de código.

### I. Aba "Logs da Aplicação"

- **Observabilidade em Tempo Real:** Visualizador integrado dos logs gravados em `AmbientesHub\logs\app.log`.
- **Destaque Sintático Colorido:** Formatação rica no console WPF com identificação visual instantânea por cores para carimbos de data/hora, níveis (`INFO`, `WARNING`, `DANGER`, `SUCCESS`) e classes chamadoras.
- **Filtros Dinâmicos:** Filtro por nível e busca por texto em tempo real com rolagem automática.

### J. Aba "Configurações & Credenciais" (Último da Barra Lateral)

- **Centralização Unificada:** Módulo posicionado como o último item da navegação, consolidando todas as variáveis de ambiente (`.env`), caminho do USMT (`USMT_PATH`) com detecção/download, botões de preenchimento padrão por campo, seleção tipográfica, arranque com o Windows e cofres seguros DPAPI.

### K. Inicialização com Windows, Bandeja do Sistema & Notificações Toast

- **Iniciar com o Windows:** Alternador (ToggleSwitch) nas Configurações para habilitar ou desabilitar o arranque automático do aplicativo após login (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`).
- **Atalho no Menu Iniciar:** Alternador nas Configurações que cria e remove o atalho `Ambientes Hub.lnk` na pasta de Programas do Menu Iniciar (`AppData\Roaming\Microsoft\Windows\Start Menu\Programs`) com metadados e ícone corporativo.
- **Menu de Contexto Produtivo na Bandeja (Tray Icon):** Acesso rápido com submenus para alternar ambientes de trabalho, saneamento imediato de RAM, controle completo da Bancada Docker (Start/Stop/Rebuild/Browser/Status/Pasta), conexões remotas RDP/SCCM com leitura inteligente do Clipboard, acesso aos logs e alternância entre Modo Comum e Modo Administrador (com badge indicativo no cabeçalho da janela).
- **Notificações Toast do Windows:** Balões informativos ao concluir tarefas assíncronas demoradas (Backup/Restore USMT, Manutenção Remota, Expurgo de Perfis, Diagnóstico de Dispositivo, Instalação de Softwares, Desprovisionamento AD/SCCM, Otimizações de Registro e eventos da Bancada Docker).
