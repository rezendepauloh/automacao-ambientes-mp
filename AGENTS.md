# AGENTS.md

Guia e diretrizes para agentes de IA e desenvolvedores que operam ou modificam este repositório de automação de ambientes no Windows.

---

## 1. Visão Geral do Projeto

Este repositório contém rotinas de automação voltadas para preparar, limpar e configurar ambientes de trabalho no Windows (incluindo rotinas de encerramento seguro do Microsoft Office, fechamento seletivo de processos, controle do Explorer, configuração do Edge e gerenciamento de memória).

### Componentes Atuais
- [biblioteca.ps1](file:///c:/AmbientesVirtuais/biblioteca.ps1): Módulo central de funções PowerShell.
- [Docs/](file:///c:/AmbientesVirtuais/Docs/README.md): Documentação técnica completa organizada nos moldes corporativos (guias operacionais e decisões arquiteturais).

---

## 2. Regras Operacionais Obrigatórias

- **Sem Execução Automática**: O agente nunca deve disparar comandos de terminal de forma autônoma. Todo comando sugerido deve ser fornecido em bloco de código formatado para execução manual pelo usuário.
- **Respostas Concisas**: Respostas diretas e sem introduções ou explicações redundantes.
- **Contexto e Ignore**: Respeitar sempre os arquivos de contexto (`.antigravity-context.md`) e ignore (`.antigravityignore`) na raiz.

---

## 3. Diretrizes de Manutenção dos Scripts

1. **Interoperabilidade Windows / COM**: Qualquer modificação que interaja com Word/Excel deve liberar objetos COM (`[System.Runtime.InteropServices.Marshal]::ReleaseComObject`) e tratar exceções silenciosamente quando apropriado.
2. **Proteção de Janelas / Processos**: Preservar a lógica do escudo de proteção baseada em PID pai (`$meuPaiPid`), PID atual (`$PID`) e títulos de janela protegidos (`"🛠️ Modo"`).
3. **Compatibilidade PowerShell**: Manter compatibilidade entre Windows PowerShell 5.1 e PowerShell 7+ (`pwsh`), atentando para comandos COM e manipulação de processos.

