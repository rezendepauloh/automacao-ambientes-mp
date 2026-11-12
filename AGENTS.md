# AGENTS.md

Guia e diretrizes para agentes de IA e desenvolvedores que operam ou modificam este repositório de automação de ambientes no Windows.

---

## 1. Visão Geral do Projeto

Este repositório contém o ecossistema de produtividade e saneamento de ambientes de trabalho no Windows, unificando a aplicação desktop nativa Fluent (**AmbientesHub**) e a biblioteca legado de scripts PowerShell.

### Componentes Atuais

- [AmbientesHub/](/AmbientesHub): Aplicação Desktop nativa C# (.NET 8 com WPF-UI) para troca de ambientes, acesso remoto e limpeza.
- [tests/](/tests): Suíte unificada de testes automatizados em C# (xUnit e runner nativo).
- [Scripts/](/Scripts): Pasta que concentra os scripts PowerShell legados (`biblioteca.ps1`, rotinas de login, etc.).
- [Docs/](/Docs/README.md): Documentação técnica completa e decisões arquiteturais.

---

## 2. Regras Operacionais Obrigatórias

- **Sem Execução Automática**: O agente nunca deve disparar comandos de terminal de forma autônoma. Todo comando sugerido deve ser fornecido em bloco de código formatado para execução manual pelo usuário.
- **Respostas Concisas**: Respostas diretas e sem introduções ou explicações redundantes.
- **Contexto e Ignore**: Respeitar sempre os arquivos de contexto (`.antigravity-context.md`) e ignore (`.antigravityignore`) na raiz.
- **Cobertura de Testes Mandatória**: Toda nova funcionalidade, serviço, modelo ou refatoração deve incluir testes automatizados em [tests/](file:///c:/AmbientesVirtuais/tests), assegurando taxa de sucesso de 100% na suíte.
- **Privacidade e Zero Dados Pessoais Hardcoded**: É estritamente proibido inserir nomes de usuários reais, logins de rede corporativa, senhas ou dados identificáveis no código-fonte, arquivos XAML, testes, exemplos ou documentação. Nomes reais devem residir unicamente no arquivo `.env` local da máquina (não versionado), e os testes ou exemplos devem usar mocks neutros (ex: `usuario.sobrenome`, `DOMINIO\usuario.teste`).

---

## 3. Diretrizes de Manutenção

1. **Prioridade C# Nativo**: Qualquer nova funcionalidade ou automação deve ser preferencialmente implementada dentro do `AmbientesHub` em código C# nativo e autônomo.
2. **Interoperabilidade Windows / COM**: Qualquer modificação que interaja com Word/Excel deve liberar objetos COM (`Marshal.ReleaseComObject`) e tratar exceções silenciosamente quando apropriado.
3. **Escudo de Proteção**: Preservar a lógica do escudo baseada em PID atual, PID pai e títulos protegidos contendo `"🛠️ Modo"`.
4. **Criação Contínua de Testes**: Ao criar ou modificar serviços (`AdSccmService`, `RemoteToolsService`, `CleanupService`, etc.) ou modelos (`MachineDetails`, etc.), criar ou atualizar imediatamente os testes xUnit correspondentes em `tests/Unit/`, `tests/Services/` ou `tests/Integration/`.
5. **Sanitização de Placeholders e Mocks**: Ao criar interfaces visuais, tooltips ou testes automatizados, utilizar sempre identidades fictícias genéricas (`seu_usuario_admin`, `usuario.sobrenome`, `DOMINIO\usuario.primario`), garantindo conformidade com repositórios públicos no GitHub.
