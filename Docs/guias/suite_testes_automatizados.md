# 🧪 Guia da Suíte de Testes Automatizados — Ambientes Hub

> **Manual de execução, arquitetura e convenções da suíte de testes em C# (.NET 8 com xUnit).**

---

## 1. Visão Geral

A suíte de testes do **Ambientes Hub** foi desenvolvida para assegurar estabilidade, integridade de processos e segurança em todas as operações de automação de ambientes, consulta ao Active Directory/SCCM e manipulação de credenciais DPAPI.

O repositório adota a política de **100% de sucesso contínuo** (0 falhas, 0 erros).

---

## 2. Estrutura Modular dos Testes

Os testes estão organizados em três camadas dentro da pasta [`tests/`](file:///c:/AmbientesVirtuais/tests):

```text
tests/
├── AmbientesHub.Tests.csproj               # Projeto de testes .NET 8 com xUnit
├── run_all.ps1                             # Script runner interativo com saída formatada ANSI
│
├── Unit/                                   # Testes Unitários de Modelos e Serviços
│   ├── MachineDetailsTests.cs              # Lógica de fallback de usuário ativo, formatação de IP e MAC
│   ├── ConfigServiceTests.cs               # Carregamento de chaves do .env e valores padrão seguros
│   ├── AppLoggerServiceTests.cs            # Gravação de logs com rotação e thread-safety em concorrência
│   └── CredentialServiceTests.cs           # Criptografia e descriptografia de cofre DPAPI do usuário
│
├── Services/                               # Testes de Serviços e Utilitários
│   ├── RemoteToolsServiceTests.cs          # Localização do CmRcViewer, mstsc e validação de argumentos
│   └── ClipboardHelperTests.cs             # Validações seguras de transferência para a área de transferência
│
└── Integration/                            # Testes de Integração e Segurança
    ├── ProcessShieldTests.cs               # Validação do escudo de proteção (PID atual e títulos "🛠️ Modo")
    └── ScriptIntegrityTests.cs             # Integridade estrutural e codificação UTF-8 dos scripts PowerShell
```

---

## 3. Como Executar os Testes

### Opção 1: Runner Interativo PowerShell (Recomendado)
Apresenta um cabeçalho estilizado com medição de tempo e saída colorida:

```powershell
pwsh -File C:\AmbientesVirtuais\tests\run_all.ps1
```

### Opção 2: Linha de Comando Nativa .NET CLI
Executa a suíte de testes diretamente via SDK do .NET:

```powershell
dotnet test C:\AmbientesVirtuais\tests\AmbientesHub.Tests.csproj
```

### Opção 3: Executar com Saída Detalhada de Cada Teste
```powershell
dotnet test C:\AmbientesVirtuais\tests\AmbientesHub.Tests.csproj --logger "console;verbosity=detailed"
```

---

## 4. Diretrizes para Novos Testes

Ao criar novas implementações no código-fonte em [`AmbientesHub/`](file:///c:/AmbientesVirtuais/AmbientesHub):

1. **Novos Modelos ou Propriedades:** Devem ser cobertos em `tests/Unit/`.
2. **Novos Serviços:** Devem ter validações de exceções e caminhos felizes em `tests/Services/`.
3. **Regras de Encerramento e Escudo:** Qualquer alteração no fechamento de janelas ou proteção de PIDs deve ser validada em `tests/Integration/ProcessShieldTests.cs`.

