# Monitoramento das DDA (Pernambuco)

Extrai do painel [Monitoramento das DDA](https://public.tableau.com/app/profile/dda.brasil/viz/MonitoramentodasDDA/1-MonitoramentoBrasil2026)
(Tableau Public, Sivep-DDA/SVSA/MS) as tabelas de Pernambuco que alimentam o
painel do Power BI da DGMOG, e grava o resultado **na mesma planilha do Google**
que o Power BI já lê. O link da planilha não muda.

Tudo roda sozinho no GitHub Actions:

1. Duas vezes por dia (10h17 e 15h17 de Brasília) o workflow confere se o
   painel do Tableau foi republicado. Se não foi, para aí.
2. Se foi, `extrair_dda.py` baixa as 8 tabelas (crosstab) com filtro UF = PE e
   todos os anos, calcula `municipiosAlerta` e valida o resultado.
3. `enviar_sheets.py` substitui o conteúdo das abas na planilha.
4. Os CSVs vão para `dados/abas/` neste repositório (histórico no git).
5. O Power BI Service atualiza o painel pela atualização agendada do workspace.

Configuração inicial: [docs/configuracao.md](docs/configuracao.md).

## Abas

| Aba | View do Tableau | Conteúdo |
|---|---|---|
| aba12 | 12-Casos /SE (Mun) | casos por município × ano × SE |
| aba15 | 15-Casos Fx Et/SE e Ano (Mun) | por faixa etária |
| aba16 | 16-Casos Plan Trat /SE e Ano (Mun) | por plano de tratamento |
| aba17 | 17-US-DDA/SE (Mun) | unidades sentinela implantadas e que informaram |
| aba23 / aba24 | 23/24-Limites DC 2026/2025 (UF) | diagrama de controle, PE |
| aba25 / aba26 | 25/26-Limites DC 2026/2025 (Mun) | diagrama de controle por município |
| municipiosAlerta | derivada de aba25 + aba26 | formato longo |

## Rodar no computador

```bash
pip install -r requirements.txt
python extrair_dda.py --csv-dir dados/abas --sem-xlsx   # extrai
python enviar_sheets.py                                 # envia à planilha
python extrair_dda.py --diagnostico -v                  # se algo quebrar
```

## Como a extração funciona

O Tableau Public não tem API de dados oficial. O script usa a mesma conversa
VizQL do navegador: `startSession/viewing` → `bootstrapSession` → filtros →
`export-crosstab`. Dois detalhes que não estão em lugar nenhum documentado:

- a configuração da sessão vem de `POST .../startSession/viewing` (a página
  da view não traz mais o `tsConfigContainer` preenchido);
- o cabeçalho de resposta `global-session-header` (endereço do servidor que
  abriu a sessão) precisa ser reenviado em todo pedido; sem ele o balanceador
  manda o pedido a outro servidor e a resposta é 410.

Antes de publicar, `validar()` confere 185 municípios, 53 semanas e 5 linhas nas
abas de UF. Se algo vier errado, o workflow falha e a planilha não é tocada.
