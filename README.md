# FlowTask

Gerenciador de tarefas (Google Apps Script + planilha "Gerenciador de Tarefas").

- `src/` — código do Apps Script (Code.js, Index.html, appsscript.json)
- Cada push na `main` publica automaticamente via GitHub Actions (`clasp push` + atualização da implantação, mesmo link).
- Requer o segredo `CLASPRC_JSON` (conteúdo do `~/.clasprc.json` gerado por `clasp login`).
