# 🖨️ Guia Prático — Migração e Cadastro de Impressoras (Print Server + PaperCut)

> **Passo a passo simples para verificar, atualizar IP ou cadastrar uma nova fila de impressora no Windows Print Server usando a aba `Print Server & Portas` do AmbientesHub.**

---

## 1. Pré-requisitos Rápidos

1. Ter o **nome da fila/host** (ex.: `PRT-0001`), o **novo IP** (ex.: `10.0.0.50`) e o **servidor de impressão** (ex.: `srv-print01`).
2. Se o servidor usar acesso via **PAM / 2FA** (ex.: `usuario[dominio\adm@SERVIDOR]`), use o fluxo **Copiar Script PowerShell + aba Servidores & RDP** (não precisa de WinRM direto).

---

## 2. Passo a Passo no AmbientesHub

1. Abra a aba **Scripts & Migração → 🖨️ Print Server & Portas**.
2. Preencha:
   - **Servidor de Impressão:** ex.: `srv-print01` (ou `.` se estiver rodando direto no servidor).
   - **Nome do Host / Fila:** ex.: `PRT-0001`.
   - **Novo Endereço IPv4:** ex.: `10.0.0.50` (a porta `PAPERCUT_10.0.0.50` / `9100` é montada automaticamente).
   - **Setor / Localização** e **Modelo / Driver Padrão** (padrão: `HP Universal Printing PCL 6`).
3. Clique em **🔍 1. Testar IP da Impressora na Rede** para confirmar que a impressora responde a **Ping** e à porta **9100 (RAW)** no novo IP.
4. Clique em **📋 2. Copiar Script PowerShell**.
5. Clique em **🖥️ Abrir RDP** (ou abra o servidor na aba **Servidores & RDP**), entre com sua conta admin, abra o **PowerShell como Administrador** e cole o script (`Ctrl+V` + `Enter`).

---

## 3. O que o Script Faz Automaticamente no Servidor

* **Cenário A — Impressora já está no IP certo:**
  - Informa `[PASSO 4] TUDO CERTO!`, não altera nada, envia a **Página de Teste** (se marcado) e abre a janela da fila.
* **Cenário B — Impressora existe, mas está com IP antigo:**
  - Cria a porta `PAPERCUT_<NOVO_IP>` (se ainda não existir) via `Add-PrinterPort` e atualiza a impressora (`Set-Printer`).
* **Cenário C — Impressora ainda não existe no servidor:**
  - Cria a porta `PAPERCUT_<NOVO_IP>`, valida se o driver está instalado e cadastra a fila (`Add-Printer` com `Shared=$true`, `Published=$false`, `Location` e `Comment`), enviando a página de teste em seguida.

> **Dica:** Se o **Print Management** já estiver aberto no servidor, aperte **F5 (Refresh)** para ver a porta e a fila novas na lista.

---

## 4. Etapa Final — PaperCut Admin

1. Clique em **🌐 3. Abrir PaperCut Admin** no AmbientesHub (usa `URL_PAPERCUT` do `.env`).
2. Localize a impressora pelo nome (`PRT-0001`), atualize/confira o IP caso necessário e verifique se a página de teste foi contabilizada no histórico.
