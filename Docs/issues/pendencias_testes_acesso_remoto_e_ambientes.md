# 🚨 Pendências de Teste Manual e Validação Operacional

Este documento registra os itens desenvolvidos no **Ambientes Hub** que necessitam de homologação e testes práticos pelo operador técnico com máquinas reais e rede corporativa.

---

## 📋 Lista de Verificação (Checklist)

### 1. 🖥️ Módulo de Acesso Remoto

- [ ] **Teste SCCM Remote (`CmRcViewer.exe`):**
  - **Cenário:** Digitar o nome de uma estação ligada na rede (ex: `DOMINIO-00000`) e clicar em **SCCM Remoto (CmRcViewer)**.
  - **Resultado Esperado:** Disparo imediato do visualizador leve do SCCM conectando no host via `\\server.dominio`, sem abrir o Console MECM.
- [ ] **Teste RDP Padrão (`mstsc.exe`):**
  - **Cenário:** Digitar o IP ou nome da máquina e clicar em **RDP Padrão (mstsc)**.
  - **Resultado Esperado:** Abertura da janela de Conexão de Área de Trabalho Remota com o campo `/v:<host>` preenchido.
- [ ] **Acionamento via Teclado:**
  - **Cenário:** Digitar o host na caixa de texto e pressionar `Enter`.
  - **Resultado Esperado:** Disparo automático da conexão padrão do SCCM.

---

### 2. 🛠️ Validação dos Demais Ambientes de Trabalho

- [ ] **Modo Contratos:**
  - **Ação:** Clicar no card `🛠️ Modo Contratos`.
  - **Resultado Esperado:** Saneamento prévio das tarefas pesadas e abertura dos portais de contratos no Firefox.
- [ ] **Modo Impressoras:**
  - **Ação:** Clicar no card `🛠️ Modo Impressoras`.
  - **Resultado Esperado:** Fechamento prévio e abertura dos painéis de impressoras e filas.
- [ ] **Modo PXE:**
  - **Ação:** Clicar no card `🛠️ Modo PXE`.
  - **Resultado Esperado:** Preparação da bancada de suporte com abertura das ferramentas de formatação e reset.
- [ ] **Modo Estudo:**
  - **Ação:** Clicar no card `🛠️ Modo Estudo`.
  - **Resultado Esperado:** Abertura das ferramentas de estudo e leitura sem distrações.

---

### 3. 🗑️ Desprovisionamento e Exclusão Real (AD & SCCM)

- [ ] **Exclusão de Máquina de Teste:**
  - **Cenário:** Testar o botão **"Excluir do AD & SCCM"** em uma máquina real de laboratório/teste.
  - **Resultado Esperado:** Diálogo de confirmação `Sim/Não` e exclusão efetiva da conta no AD e do ResourceID no SCCM, liberando reformatacão imediata via PXE.
