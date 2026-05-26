Passo a passo: receber respostas no Google Sheets via Apps Script

1) Criar uma Google Sheet nova
- Acesse https://sheets.google.com e crie uma nova planilha. Anote o ID na URL (parte entre `/d/` e `/edit`).

2) Abrir Apps Script
- Na planilha, menu Extensões → Apps Script.

3) Colar o código
- Substitua o conteúdo do editor pelo arquivo `google_apps_script_webapp.gs` presente nesta pasta.

4) Configurar Script Properties (TARGET_SHEET_ID)
- No editor Apps Script: Clique em `Project Settings` (ícone de engrenagem) → `Project properties` → `Script properties` → adicione `TARGET_SHEET_ID` com o ID da planilha criada.

5) Salvar e Deploy
- Clique em `Deploy` → `New deployment` → tipo: `Web app`.
- Em `Execute as` escolha `Me`.
- Em `Who has access` escolha `Anyone` ou `Anyone, even anonymous` (permite POSTs sem autenticação). Confirme.
- Copie a URL do Web App.

6) Atualizar seu HTML
- No arquivo `formulario_governanca_ti_v2.html`, preencha a constante `GOOGLE_APPS_SCRIPT_URL` com a URL copiada.

7) Teste
- Abra seu HTML no navegador local.
- Responda o questionário; ao finalizar, o frontend fará um POST à URL e a linha será adicionada na aba `Responses`.

8) Métricas no Sheets
- Na própria planilha, use fórmulas para calcular KPIs, por exemplo:
  - Total: =COUNTA(Responses!A:A)-1
  - Média de percentual: =AVERAGE(Responses!D2:D)
  - Taxa de aprovação (>=80%): =COUNTIF(Responses!D2:D, ">=80") / (COUNTA(Responses!A:A)-1)
  - Acerto por pergunta: será necessário parsear a coluna `answers` (JSON) ou gravar respostas em colunas separadas.

Observações de segurança e privacidade
- Se usar `Anyone, even anonymous`, qualquer pessoa com a URL pode enviar POSTs. Para apresentação isso costuma ser aceitável, mas mantenha a URL privada.
- Para produção/maiores volumes, considere autenticação e validação adicional.

Precisa que eu preencha a constante `GOOGLE_APPS_SCRIPT_URL` no HTML com a sua URL após você fazer o deploy, ou quer que eu gere também uma versão que grava cada resposta em colunas separadas (útil para calcular acerto por pergunta diretamente no Sheets)?
