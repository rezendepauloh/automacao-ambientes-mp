# 📚 Documentação Técnica — Ambientes Virtuais & Automações MP

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
│   └── ciclo_de_vida_e_seguranca_processos.md        # Lógica de escudos de proteção (PID/Título) e COM Office
│
└── 📖 guias/                                         # Manuais práticos de operação e scripts
    ├── guia_execucao_e_uso.md                        # Como carregar e executar as rotinas no PowerShell
    └── catalogo_funcoes_biblioteca.md                # Referência técnica de cada função da biblioteca.ps1
```

---

## 🚀 Acessos Rápidos Mais Utilizados

### 1. Operação Técnica Diária
- **Como executar e carregar as funções:** [Guia de Execução e Uso](guias/guia_execucao_e_uso.md)
- **Referência completa das funções:** [Catálogo de Funções da Biblioteca](guias/catalogo_funcoes_biblioteca.md)

### 2. Engenharia e Arquitetura de Software
- **PowerShell vs C# (.NET AOT) para performance:** [Análise Arquitetural](arquitetura/analise_arquitetural_powershell_vs_csharp.md)
- **Mecanismos de segurança e encerramento COM:** [Ciclo de Vida e Segurança de Processos](arquitetura/ciclo_de_vida_e_seguranca_processos.md)

