# 📚 Documentação Técnica — Ambientes Virtuais & Automações Corporativas

> **Central unificada de documentação de engenharia, arquitetura, manuais operacionais e guias práticos do repositório de automação de ambientes de trabalho no Windows.**  
> Rotinas de limpeza de memória, encerramento transacional seguro do Office via COM, supressão de telemetria, orquestração de navegadores e manipulação de janelas.

---

## 🗂️ Estrutura e Mapa de Navegação da Documentação

```text
Docs/
├── README.md                                         # Este índice central unificado
│
├── 🏛️ arquitetura/                                   # Decisões fundamentais de engenharia e modelagem
│   ├── analise_arquitetural_powershell_vs_csharp.md  # Comparativo técnico detalhado de performance e manutenção
│   ├── ciclo_de_vida_e_seguranca_processos.md        # Lógica de escudos de proteção (PID/Título) e COM Office
│   ├── hub_winui3_arquitetura.md                     # Arquitetura do Hub Desktop WinUI 3 + C# (Remoto, Ambientes e AD)
│   └── seguranca_hardening.md                        # Arquitetura SecOps: whitelists, fail-closed, DPAPI e ACLs
│
└── 📖 guias/                                         # Manuais práticos de operação e scripts
    ├── guia_ambientes_hub_desktop.md                 # Manual completo do aplicativo desktop AmbientesHub
    ├── suite_testes_automatizados.md                 # Manual da suíte de testes automatizados xUnit e runner
    ├── guia_execucao_e_uso.md                        # Como carregar e executar as rotinas legadas (Scripts/biblioteca.ps1)
    ├── catalogo_funcoes_biblioteca.md                # Referência técnica de cada função da biblioteca.ps1
    └── preparacao_ambiente_dotnet.md                 # Instalação do .NET 8/9 SDK para o Hub WinUI 3
│
└── 🚨 issues/                                        # Pendências e homologações
    └── pendencias_testes_acesso_remoto_e_ambientes.md# Checklist de validação humana para Acesso Remoto e Modos
```

---

## 🚀 Acessos Rápidos Mais Utilizados

### 1. Operação Técnica Diária & Testes
- **Suíte de Testes Automatizados (xUnit):** [Guia da Suíte de Testes](guias/suite_testes_automatizados.md)
- **Manual do Ambientes Hub Desktop:** [Guia do Aplicativo](guias/guia_ambientes_hub_desktop.md)
- **Como executar e carregar as funções:** [Guia de Execução e Uso](guias/guia_execucao_e_uso.md)
- **Referência completa das funções:** [Catálogo de Funções da Biblioteca](guias/catalogo_funcoes_biblioteca.md)

### 2. Engenharia e Arquitetura de Software
- **Arquitetura de Segurança & Hardening (SecOps):** [Segurança e Hardening](arquitetura/seguranca_hardening.md)
- **PowerShell vs C# (.NET AOT) para performance:** [Análise Arquitetural](arquitetura/analise_arquitetural_powershell_vs_csharp.md)
- **Mecanismos de segurança e encerramento COM:** [Ciclo de Vida e Segurança de Processos](arquitetura/ciclo_de_vida_e_seguranca_processos.md)

