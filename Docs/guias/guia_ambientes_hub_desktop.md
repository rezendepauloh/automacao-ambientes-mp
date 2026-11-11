# 📖 Guia de Operação do Ambientes Hub (WinUI / C#)

O **Ambientes Hub** é a interface desktop moderna (Fluent Design Windows 11) desenvolvida em .NET 8 com o objetivo de centralizar a alternância de ambientes de trabalho, fornecer launcher leve para conexões remotas e permitir operações de limpeza e desprovisionamento sem a lentidão de consoles pesados.

---

## 1. Como Compilar e Executar a Aplicação

Abra o terminal do PowerShell na pasta do projeto e execute:

```powershell
# Navegar até o diretório da aplicação
cd C:\AmbientesVirtuais\AmbientesHub

# Compilar e executar em modo de desenvolvimento
dotnet run
```

---

## 2. Como Gerar o Executável Pronto para Uso (`.exe`)

Para gerar uma versão final e otimizada que pode ser colocada na Área de Trabalho ou fixada na Barra de Tarefas:

```powershell
cd C:\AmbientesVirtuais\AmbientesHub

# Gera a publicação autocontida para Windows x64
dotnet publish -c Release -r win-x64 --self-contained false -o .\bin\Publish
```

O executável estará pronto em:
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
- **Entrada única:** Permite digitar o nome de rede (ex: `MPE-80600`) ou endereço IP.
- **Botão "SCCM Remoto":** Invoca diretamente o `CmRcViewer.exe` com a flag do site corporativo sem a necessidade de abrir o Console do MECM/SCCM.
- **Botão "RDP Padrão":** Dispara o `mstsc.exe` já configurado com a máquina alvo.

### C. Aba "Limpeza AD & SCCM"
- Integra a rotina desprovisionadora de `D:\ArrumarPC\remover_computador_ad_sccm.ps1`.
- Exibe o log em tempo real dentro da própria janela do Hub, com confirmação de segurança antes de efetuar a exclusão.

### D. Aba "Performance & Limpeza"
- Dispara o saneamento do Windows: salva documentos do Office com segurança, fecha processos pesados sem derrubar o próprio Hub e limpa histórico de abas travadas do Edge.

