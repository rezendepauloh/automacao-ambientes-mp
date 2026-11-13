# 📖 Guia de Instalação e Preparação do Ambiente .NET

Para compilar, publicar e executar o Ambientes Hub e sua suíte de testes xUnit, é necessário o **.NET 8.0 SDK** (ou superior).

---

## 1. Pré-Requisitos Recomendados

- **Windows 10 (Build 19041+) ou Windows 11**.
- **.NET 8.0 SDK**: Necessário para desenvolvimento, compilação de binários e execução dos testes xUnit.

---

## 2. Opções de Instalação

### Opção A: Via Windows Package Manager (`winget`) — Recomendado
Abra o PowerShell como Administrador e execute:

```powershell
# Instala o .NET 8 SDK oficial da Microsoft
winget install Microsoft.DotNet.SDK.8 --silent --accept-package-agreements --accept-source-agreements
```

---

### Opção B: Download Oficial do Instalador (.exe)
1. Acesse: [https://dotnet.microsoft.com/download/dotnet/8.0](https://dotnet.microsoft.com/download/dotnet/8.0)
2. Baixe o instalador do **.NET SDK x64**.
3. Execute o instalador e avance até a conclusão.

---

## 3. Validação e Compilação Inicial

Após instalar, abra uma nova janela de terminal e execute:

```powershell
# 1. Validar versão do SDK instalada
dotnet --version

# 2. Restaurar dependências e compilar a aplicação
dotnet build C:\AmbientesVirtuais\AmbientesHub\AmbientesHub.csproj

# 3. Executar a suíte completa de testes unitários e de integração
pwsh -NoProfile -File C:\AmbientesVirtuais\tests\run_all.ps1
```

