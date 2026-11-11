# AGENTS.md

Guia e diretrizes para agentes de IA e desenvolvedores que operam ou modificam este repositório de automação de ambientes no Windows.

---

## 1. Visão Geral do Projeto

Este repositório contém o ecossistema de produtividade e saneamento de ambientes de trabalho no Windows, unificando a aplicação desktop nativa Fluent (**AmbientesHub**) e a biblioteca legado de scripts PowerShell.

### Componentes Atuais
- [AmbientesHub/](file:///c:/AmbientesVirtuais/AmbientesHub): Aplicação Desktop nativa C# (.NET 8 com WPF-UI) para troca de ambientes, acesso remoto e limpeza.
- [Scripts/](file:///c:/AmbientesVirtuais/Scripts): Pasta que concentra os scripts PowerShell legados (`biblioteca.ps1`, rotinas de login, etc.).
- [Docs/](file:///c:/AmbientesVirtuais/Docs/README.md): Documentação técnica completa e decisões arquiteturais.

---

## 2. Regras Operacionais Obrigatórias

- **Sem Execução Automática**: O agente nunca deve disparar comandos de terminal de forma autônoma. Todo comando sugerido deve ser fornecido em bloco de código formatado para execução manual pelo usuário.
- **Respostas Concisas**: Respostas diretas e sem introduções ou explicações redundantes.
- **Contexto e Ignore**: Respeitar sempre os arquivos de contexto (`.antigravity-context.md`) e ignore (`.antigravityignore`) na raiz.

---

## 3. Diretrizes de Manutenção

1. **Prioridade C# Nativo**: Qualquer nova funcionalidade ou automação deve ser preferencialmente implementada dentro do `AmbientesHub` em código C# nativo e autônomo.
2. **Interoperabilidade Windows / COM**: Qualquer modificação que interaja com Word/Excel deve liberar objetos COM (`Marshal.ReleaseComObject`) e tratar exceções silenciosamente quando apropriado.
3. **Escudo de Proteção**: Preservar a lógica do escudo baseada em PID atual, PID pai e títulos protegidos contendo `"🛠️ Modo"`.
