# Painel Giovanna Adv

Painel de gestão do escritório da Giovanna Comácio: clientes, processos, honorários e parcelas a receber, gastos por tipo/área/serviço, campanhas e documentos no Google Drive.

Roda como **Web App do Google Apps Script** na conta Google da Giovanna. Os dados ficam numa planilha no Drive dela, e os documentos em pastas por cliente e por processo.

## Como as atualizações chegam ao painel

Tudo em `src/` é enviado automaticamente para o projeto Apps Script quando muda na branch `main`, pelo workflow `.github/workflows/publicar.yml` (usa o [clasp](https://github.com/google/clasp), ferramenta oficial do Google).

- Use o link **`/dev`** do app (Implantar → Testar implantações): ele sempre roda o código mais recente, sem implantar nada.
- Na primeira abertura de uma versão nova, o próprio painel ajusta a planilha (colunas e listas novas). Dados existentes não são alterados.
- Para rodar o envio manualmente: aba **Actions** → *Publicar no Apps Script* → *Run workflow*.

## Configuração (uma vez só)

1. Na conta da Giovanna, ative a **API do Google Apps Script**: <https://script.google.com/home/usersettings>.
2. Em qualquer computador com Node.js, rode `npx @google/clasp@2.4.2 login` e entre com a **conta da Giovanna**. Isso cria o arquivo `.clasprc.json` na pasta do usuário (`C:\Users\<você>\.clasprc.json` no Windows, `~/.clasprc.json` no Mac/Linux).
3. No GitHub: **Settings → Secrets and variables → Actions → New repository secret**:
   - `CLASPRC_JSON`: o conteúdo inteiro do `.clasprc.json`.
   - `SCRIPT_ID`: o ID do script (Apps Script → Configurações do projeto → ID do script).
   - `DEPLOYMENT_ID` *(opcional)*: só se quiser que o link `/exec` também seja atualizado a cada envio (Implantar → Gerenciar implantações → ID da implantação).
4. Aba **Actions** → *Publicar no Apps Script* → *Run workflow*. Se ficar verde, está tudo ligado.

O `CLASPRC_JSON` dá acesso aos projetos Apps Script da conta da Giovanna. Mantenha o repositório **privado**. Para revogar: <https://myaccount.google.com/connections> → clasp → remover acesso.

## Estrutura

| Arquivo | Papel |
|---|---|
| `src/Codigo.gs` | Publicação do app, API chamada pelas telas, `setup()` e migração automática |
| `src/Repo.gs` | Camada de dados. Hoje é a planilha; numa migração futura (ex.: Cloudflare D1), só este arquivo muda |
| `src/Drive.gs` | Pastas e documentos de clientes e processos |
| `src/Index.html`, `src/Estilos.html`, `src/App.html` | Telas. Toda chamada ao servidor passa por `call()` no `App.html` |
| `src/appsscript.json` | Manifesto: fuso de Brasília, app web executando como a dona, acesso só dela |
