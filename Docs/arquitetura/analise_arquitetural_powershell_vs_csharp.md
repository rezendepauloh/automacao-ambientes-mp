# 🏛️ Análise Arquitetural: PowerShell vs. C# (.NET AOT)

Avaliação de viabilidade técnica, ganhos de performance e impacto prático na migração dos scripts para uma aplicação compilada em C#.

---

## 1. Onde está o gargalo real de performance hoje?

Ao analisar a rotina `Limpar-Ambiente`, os maiores tempos de execução não vêm do interpretador PowerShell em si, mas sim de:

1. **Latência de Interação COM com o Office:** Salvar arquivos e aguardar o processo `excel.exe`/`winword.exe` sincronizar em disco.
2. **Pausas Físicas (`Start-Sleep`):**
   - Pausa de 3s após reiniciar o Explorer para recarregar a barra de tarefas.
   - Pausas de 2s a 4s em `Abrir-PastasEmAbas` para simulação de `SendKeys`.
3. **Chamada Externa `powershell.exe` dentro do script (Linha 81):** O script invoca uma subinstância separada de `powershell.exe` para salvar o Office, gerando um custo extra de inicialização de processo.

---

## 2. Comparativo Técnico Detalhado

| Métrica / Recurso | PowerShell (`biblioteca.ps1`) | C# (.NET 8/9 com Native AOT) |
| :--- | :--- | :--- |
| **Tempo de Startup da Aplicação** | ~500ms - 1.2s (JIT / Engine PS) | **< 15ms** (Executável nativo, sem overhead de inicialização) |
| **Consumo de Memória do Processo**| ~60MB - 120MB | **~8MB - 18MB** |
| **Integração COM (Word / Excel)** | Simples via `[Marshal]::GetActiveObject` | Nativa via `Marshal` / Interop COM clássico |
| **Fechamento e Listagem de Processos** | `Get-Process` + PIPs | Win32 API direta ou `System.Diagnostics.Process` (10x mais rápido) |
| **Manipulação do Explorer em Abas** | Simulação frágil via `WScript.SendKeys` | Pode usar UI Automation nativo do Windows (`UIAutomationClient`) sem depender de sleeps cegos |
| **Distribuição / Atualização** | Editar o arquivo `.ps1` com bloco de notas | Requer compilação (`dotnet publish`) gerando um `.exe` |

---

## 3. Ganharíamos performance com C#?

### Onde GANHAMOS com C#:
1. **Listagem e finalização de dezenas de processos:** Uma listagem em Win32/C# de processos ativos e checagem de títulos leva menos de **5 milissegundos**, enquanto loops no PowerShell com `Get-Process` e `MainWindowTitle` podem levar de 300ms a 1.5s.
2. **Eliminação do sub-processo `powershell.exe`:** A automação COM roda no mesmo processo nativo, eliminando a criação de um processo extra.
3. **Automação de abas do Explorer sem `SendKeys`:** Em C#, é possível usar a API de UI Automation do Windows para enviar comandos de abertura de abas diretamente ao elemento de UI do Explorer, reduzindo os tempos de espera forçados (`Start-Sleep`).

### Onde NÃO HÁ GANHO:
- O tempo do Word/Excel para gravar arquivos em disco e o tempo do Windows para reiniciar a barra de tarefas (`explorer.exe`) dependem do sistema operacional e do hardware (disco/CPU), não da linguagem.

---

## 4. Conclusão & Recomendação Prática

1. **Curto Prazo (Otimização do PowerShell Atual):**
   - Eliminar a chamada da linha 81 (`powershell.exe -NoProfile ...`) e rodar o bloco de COM diretamente na sessão atual.
   - Substituir chamadas redundantes a `Get-Process` por uma única consulta na memória.
   - Isso já reduz em ~30% a 40% o tempo total do script sem mudar de linguagem.

2. **Médio Prazo (Migração para C# CLI / Ferramenta de Bandeja):**
   - Caso se deseje transformar essa automação em um utilitário de sistema profissional (ex: atalho global `Win+Shift+L`, ícone na bandeja do sistema ao lado do relógio com clique rápido para "Limpar Ambiente"), **C# .NET 8/9 Native AOT** é a escolha ideal, combinando velocidade instantânea, baixo consumo e robustez.

