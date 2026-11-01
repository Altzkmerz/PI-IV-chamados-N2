# Solução final — Predição e classificação de chamados N2

## O que entregar / usar

| Arquivo | Para que serve |
|---|---|
| `dashboard_solucao_final.html` | Interface de visualização da solução (abrir no navegador) |
| `Chamados_N2_com_predicao.xlsx` | Base pontuada para o Power BI (`Relatorio de Chamados.pbix`) |
| `gerar_solucao.js` | Script que gera scores (Naive Bayes + regressão logística) |
| `metricas.json` | Métricas do modelo para o relatório |

## Como atualizar

1. Coloque a nova exportação de atendimentos na pasta `PI IV`.
2. No terminal, em `Solucao_Final`: `node gerar_solucao.js`
3. No Power BI, aponte a fonte para `Chamados_N2_com_predicao.xlsx` e atualize.

## Como o modelo decide

- **Nível (N1/N2/N3/Misto):** Naive Bayes no título + categoria + prioridade.
- **Risco de ciclo longo:** regressão logística; ciclo > 10 dias = risco alto.
- **Ação sugerida:** devolver N1, escalar N3 ou priorizar no N2.

## Números da base de 26/10/2026

- 614 chamados · 27 abertos (4,4%)
- Acurácia do nível (holdout): 78,4%
- F1 de risco alto (holdout): 0,81
- 23 abertos com risco alto · 2 com perfil N1 · 9 com perfil N3
