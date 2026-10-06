# PMP — Plano Mestre de Produção

Gera a planilha `PMP_PSI.xlsx` a partir da carteira **FGI 8.5-01.02 - Carteira de Projetos**.

```bash
python3 pmp/1_extrair_carteira.py "FGI_8.5-01.02_-_Carteira_de_Projetos.xlsx"   # gera pmp_data.json
python3 pmp/2_gerar_pmp.py PMP_PSI.xlsx
```

Abas geradas: `LEIAME`, `PMP` (1 linha por painel, mês + % restante por etapa),
`CARGA x CAPACIDADE` (utilização e backlog por centro), `PMP MENSAL` (saída de painéis por projeto/mês),
`PARAMETROS` (headcount, dias úteis, eficiência) e `TEMPOS` (horas padrão Tipo+Complexidade × etapa).

Os scripts são só o *bootstrap*: depois de gerada, a planilha é mantida direto no Excel.

> **Não versionar a carteira nem o PMP gerado neste repositório** — ele é público e os arquivos têm dados de clientes.
