# 🚀 Automação de Ambientes de Trabalho com PowerShell e Selenium

Este repositório contém um conjunto de scripts modulares em PowerShell criados para automatizar a preparação do ambiente de trabalho e de estudos. Com um clique, o sistema encerra distrações, salva trabalhos pendentes, limpa resíduos de memória e abre todas as ferramentas, pastas e sites necessários (com login automatizado).

## ✨ Funcionalidades (Superpoderes)

- **🧹 Limpeza Nuclear (`biblioteca.ps1`)**: Salva automaticamente documentos abertos do Word/Excel, encerra processos desnecessários e reinicia o Windows Explorer para liberar memória.
- **🎯 Sniper Anti-Telemetria**: Identifica e neutraliza processos de telemetria do Windows (como o `CompatTelRunner`) para garantir 100% de performance da CPU e Disco durante o foco.
- **👻 Teclado Fantasma (Automação de UI)**: Agrupa múltiplas pastas do Windows Explorer em uma única janela utilizando abas dinâmicas via `WScript.Shell`.
- **🤖 Login Automático Invisível (Selenium WebDriver)**:
  - Abre o Microsoft Edge em um ambiente limpo.
  - Bypass de Preenchimento Automático: Injeta credenciais criptografadas de forma forçada (`Ctrl+A`), garantindo sucesso no login em sistemas complexos.
  - Verificação Inteligente de Sessão: Só realiza o login se o cookie de sessão estiver expirado.
- **📚 Perfis de Contexto**:
  - `chamados.ps1`: Prepara o ambiente operacional para atendimento (SIMP, OTRS/Central, Planilhas de controle, WhatsApp, Teams).
  - `estudo.ps1`: Prepara o ambiente blindado para foco total (Plataformas EAD, Youtube Music, Pomodoro/Relógio do Windows).
  - `contratos.ps1`: Carrega o escopo de trabalho focado na gestão e fiscalização de contratos.
  - `pxe.ps1`: Prepara o ambiente técnico para procedimentos de formatação e manutenção via rede (PXE).

## 🛠️ Pré-requisitos

1.  **Sistema Operacional**: Windows 10 ou 11.
2.  **Linguagem**: PowerShell 5.1 ou superior.
3.  **Dependência Externa**: Biblioteca `WebDriver.dll` do Selenium 4+.
4.  **Cofre de Senhas**: É necessário gerar um arquivo local `credenciais.xml` usando o `Export-Clixml` do PowerShell contendo o objeto de credencial de rede (Protegido pela API de Proteção de Dados do Windows - DPAPI).

## ⚠️ Aviso de Segurança Importante

**NUNCA comite o seu arquivo `credenciais.xml`**. O repositório já conta com um `.gitignore` configurado para impedir o envio acidental de arquivos XML. Mantenha seu cofre de senhas apenas na sua máquina local.

## 🚀 Como usar

1. Clone o repositório na sua máquina.
2. Certifique-se de que a `WebDriver.dll` está no caminho especificado nos scripts (ex: `C:\AmbientesVirtuais\`).
3. Crie atalhos na sua Área de Trabalho ou Menu Iniciar apontando para `chamados.ps1` ou `estudo.ps1`.
4. Clique e assista à mágica acontecer.
