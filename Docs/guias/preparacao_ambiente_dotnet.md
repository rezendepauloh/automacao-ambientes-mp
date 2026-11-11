# 📖 Guia de Instalação e Preparação do Ambiente .NET & WinUI 3

Para compilar e executar o Hub Desktop em C# e WinUI 3 no Windows, precisamos do **.NET 8 ou .NET 9 SDK**.

---

## 1. Situação Atual Detectada

No terminal, o comando `dotnet --version` retornou:
> `No .NET SDKs were found.`

Isso indica que o runtime ou SDK do .NET ainda não está instalado ou não está no `PATH`.

---

## 2. Opções de Instalação Recomendadas

### Opção A: Via Windows Package Manager (`winget`) — Mais Rápido
Abra o PowerShell como Administrador e execute:

```powershell
# Instala o .NET 8 SDK oficial da Microsoft
winget install Microsoft.DotNet.SDK.8 --silent --accept-package-agreements --accept-source-agreements
```

*(Ou para .NET 9):*
```powershell
winget install Microsoft.DotNet.SDK.9 --silent --accept-package-agreements --accept-source-agreements
```

---

### Opção B: Download Oficial do Instalador (.exe)
1. Acesse: [https://dotnet.microsoft.com/download/dotnet/8.0](https://dotnet.microsoft.com/download/dotnet/8.0)
2. Baixe o instalador do **.NET SDK x64**.
3. Execute o instalador e avance até a conclusão.

---

## 3. Validação pós-instalação

Após instalar, abra uma **nova janela do terminal** (para carregar o novo PATH do sistema) e execute:

```powershell
dotnet --version
```

A saída deverá exibir a versão instalada (ex: `8.0.404` ou `9.0.100`).

