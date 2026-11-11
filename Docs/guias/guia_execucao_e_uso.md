# 📖 Guia de Execução e Uso

Guia passo a passo para carregar e executar as funções do módulo `biblioteca.ps1` no Windows PowerShell (5.1) ou PowerShell 7+ (`pwsh`).

---

## 1. Pré-Requisitos

- **Sistema Operacional:** Windows 10 ou Windows 11.
- **Permissão de Execução:** Política que permita carregar scripts locais (ex: `RemoteSigned` ou `Bypass`).
- **Privilégios:**
  - Modo padrão (usuário comum): Executa todas as funções normais (fechamento de processos do usuário, salvamento Office, abas do Edge/Explorer).
  - Modo Administrador (opcional): Necessário apenas para interromper serviços de sistema como o `DiagTrack` (Telemetria).

---

## 2. Como Carregar as Funções (Dot-Sourcing)

Abra o terminal do PowerShell na raiz do projeto (`C:\AmbientesVirtuais`) e carregue as funções na sessão atual:

```powershell
# Carrega todas as funções na memória do terminal atual
. .\biblioteca.ps1
```

---

## 3. Exemplos Práticos de Execução

### A. Limpeza Geral de Ambiente
Executa o ciclo completo: salva documentos do Office, fecha aplicações pesadas respeitando escudos de PID/Título, limpa cache de sessão do Edge, reinicia o Explorer e mata processos de telemetria.

```powershell
Limpar-Ambiente
```

---

### B. Fechamento Rápido do Teams
Derruba processos persistentes do Teams (legado, moderno e web):

```powershell
Matar-Teams
```

---

### C. Abertura Agrupada de Pastas em Abas do Explorer
Abre o Explorador de Arquivos e adiciona pastas como abas na mesma janela:

```powershell
$pastasDeTrabalho = [ordered]@{
    "Projetos"  = "C:\Projetos"
    "Downloads" = "$env:USERPROFILE\Downloads"
    "Documentos"= "$env:USERPROFILE\Documents"
}

Abrir-PastasEmAbas -Pastas $pastasDeTrabalho
```

---

### D. Abertura Otimizada do Microsoft Edge
Abre o Edge em modo leve (sem extensões) com a vacina anti-crash aplicada:

```powershell
$meusSites = [ordered]@{
    "Portal Corporativo" = "https://intranet.exemplo.gov.br"
    "Chamados"           = "https://chamados.exemplo.gov.br"
}

Abrir-SitesEdgeLeve -Sites $meusSites
```

---

### E. Limpeza de Credenciais da Memória
Remove variáveis que continham credenciais e força a coleta imediata pelo Garbage Collector:

```powershell
# Usando os nomes de variáveis padrão
Remover-CredenciaisMemoria

# Ou informando variáveis customizadas
Remover-CredenciaisMemoria -Variaveis @("tokenAcesso", "minhaSenha")
```

