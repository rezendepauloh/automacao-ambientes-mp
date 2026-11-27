# ☁️ Guia de Sincronização dos FAQs do SharePoint (Fonte Única)

> **Procedimento operacional para manter os tutoriais `.md` em `Tutoriais-FAQs\Manuais Markdown\` sempre atualizados a partir das páginas `SitePages` do SharePoint, sem duplicidade manual.**

---

## 1. Pré-requisitos no `.env`

Na aba **⚙️ Configurações & Credenciais** (ou diretamente em `AmbientesHub/.env`), configure as duas chaves abaixo:

- `SHAREPOINT_FAQS_DIR`: caminho local sincronizado pelo OneDrive (ex.: `C:\Users\usuario.teste\OneDrive - Organizacao\Tutoriais-FAQs`).
- `SHAREPOINT_FAQS_WEB_URL`: URL web da biblioteca no SharePoint (ex.: `https://empresa.sharepoint.com/sites/equipe/Shared%20Documents/Tutoriais-FAQs`).

> [!NOTE]
> Se `SHAREPOINT_FAQS_DIR` estiver vazio, o `AmbientesHub` opera em modo *fail-closed* exibindo apenas a documentação técnica de `Docs/`.

---

## 2. Como Sincronizar Novos FAQs ou Alterações

Sempre que um FAQ for criado ou editado no SharePoint:

1. No `AmbientesHub`, abra a aba **📚 Documentação** e clique em **`☁️ Sincronizar`** no canto superior direito.
2. Quando perguntado se deseja abrir o feed de `SitePages` no **Firefox**, clique em **Sim** (o Hub abre automaticamente a URL da REST API usando sua sessão já autenticada no Firefox).
3. No Firefox, o download do arquivo `items` (ou `items.xml`) acontece automaticamente na pasta `Downloads` (se o XML abrir na tela, pressione **`Ctrl + S`** para salvá-lo).
4. De volta ao `AmbientesHub`, selecione o arquivo baixado na janela de seleção (já abre em `Downloads`).
5. O `SharePointFaqSyncService` converte automaticamente todas as páginas para Markdown em `SHAREPOINT_FAQS_DIR\Manuais Markdown\`, atualiza o `README.md` e recarrega a lista na hora.
6. O **OneDrive** envia os `.md` atualizados de volta para o SharePoint automaticamente para toda a equipe.

---

## 3. Imagens de `SiteAssets` & Economia de Disco (OneDrive Sob Demanda)

- **Imagens em `Imagens FAQ/` e `Imagens FAQ/SiteAssets/`:** são renderizadas inline apenas quando você abre o tutorial específico (e embutidas em Base64 no botão **📦 Exportar HTML**).
- **Vídeos `.mp4` e PDFs:** aparecem como cards sob demanda — o OneDrive **não** baixa o vídeo ao abrir o tutorial, apenas ao clicar em **`▶ Assistir Aqui`** (player embutido) ou **`🪟 Externo`**.
- **Para liberar espaço em disco:** no Windows Explorer, clique com o botão direito nas pastas `Imagens FAQ` ou `Vídeos FAQ` → **Liberar espaço**.
