# Configuração (uma vez)

São três peças: a credencial do Google para o GitHub escrever na planilha,
os segredos no GitHub e a atualização agendada no Power BI Service.

## 1. Credencial do Google: conta de serviço (recomendado)

A conta de serviço é um "usuário robô" do Google. Ela não expira e não depende
do seu login.

1. Entre no [Google Cloud Console](https://console.cloud.google.com/) com o
   e-mail que vai ser o responsável pela automação (pode ser outro, não o do
   `client_secret.json`). Aceite os termos se for o primeiro acesso. Não precisa
   de cartão de crédito: a Sheets API é gratuita.
2. No seletor de projeto (topo da tela) > *Novo projeto*. Nome: `automacao-dda`.
   Crie e confira que ele ficou selecionado no topo.
3. *APIs e serviços > Biblioteca* > procure **Google Sheets API** > *Ativar*.
4. *IAM e administrador > Contas de serviço > Criar conta de serviço*.
   Nome: `dda-github`. Não precisa dar papel nenhum.
5. Abra a conta criada > *Chaves > Adicionar chave > Criar nova chave > JSON*.
   Um arquivo `.json` é baixado. **Não envie esse arquivo para o repositório.**
6. Copie o e-mail da conta (algo como `dda-github@<projeto>.iam.gserviceaccount.com`).
7. Na planilha do Google que alimenta o painel, clique em *Compartilhar* e dê
   acesso de **Editor** a esse e-mail. Desmarque *Notificar pessoas* (o robô não
   tem caixa de entrada). O compartilhamento é feito com a conta dona da planilha,
   não importa qual e-mail criou a conta de serviço.

Se a sua organização bloquear a criação de chaves, use a alternativa do item 2b.

## 2. Segredos no GitHub

No repositório: *Settings > Secrets and variables > Actions > New repository secret*.

**2a.** Nome `GOOGLE_SERVICE_ACCOUNT_JSON`, valor: o conteúdo inteiro do `.json`
da conta de serviço (abra no Bloco de Notas, copie tudo e cole).

**2b. (alternativa)** Nome `GOOGLE_TOKEN_JSON`, valor: o conteúdo do
`token_sheets.json` que o seu script atual já gera. Atenção: se o app OAuth do
Google Cloud estiver em modo *Teste*, esse token expira em 7 dias. Para não
expirar, publique o app (*Tela de consentimento OAuth > Publicar app*).

Se um dia a planilha mudar, crie a *variable* (aba Variables) `SPREADSHEET_ID`.
Sem ela, o workflow usa o ID atual (`19AbG5og...`).

## 3. Primeira execução

*Actions > Atualizar base DDA > Run workflow*, marcando **Extrair mesmo sem
versão nova do painel**. Leva cerca de 10 minutos. Confira a planilha.

A partir daí, o workflow roda sozinho todo dia e só extrai quando o Tableau
tiver uma versão nova. Se uma execução falhar, o GitHub envia um e-mail e a
próxima tentativa acontece no horário seguinte.

## 4. Atualização agendada no Power BI Service

1. Em [app.powerbi.com](https://app.powerbi.com), abra o workspace do painel e
   o **modelo semântico** *Painel Monitoramento DDA*: *... > Configurações*.
2. *Credenciais da fonte de dados > Editar credenciais* da fonte Google
   Sheets: método **OAuth2**, entre com a conta Google que tem acesso à planilha.
3. *Atualizar > Atualização agendada*: ligue, fuso **Brasília**, frequência
   **Diária**, horários **11:00** e **16:00** (depois das duas checagens do GitHub).
4. Marque *Enviar notificações de falha* para você.

Com isso, não é mais preciso abrir o Power BI Desktop para atualizar.

## Opcional: ler direto do GitHub, como no projeto OCI

Os mesmos dados ficam em `dados/abas/*.csv` neste repositório. Se um dia a
planilha do Google atrapalhar, cada consulta do Power BI pode trocar a origem
para o CSV (credencial **Anônimo**), por exemplo:

```powerquery
let
    Fonte = Csv.Document(Web.Contents(
        "https://raw.githubusercontent.com/DGMOG-SES/monitoramento-dda/main/dados/abas/aba12.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Cabecalhos = Table.PromoteHeaders(Fonte, [PromoteAllScalars = true])
in
    Cabecalhos
```

Os passos seguintes de cada consulta continuam iguais, porque as colunas e os
valores são os mesmos da planilha.
