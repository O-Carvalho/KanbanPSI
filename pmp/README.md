# PMP — Plano Mestre de Produção

Gera a planilha do PMP a partir da carteira **FGI 8.5-01.02 - Carteira de Projetos**.

```bash
python3 pmp/gerar_pmp.py "FGI_8.5-01.02_-_Carteira_de_Projetos.xlsx" PMP_PSI.xlsx pmp/carteira.pq
```

Fluxo: carteira (tabela `PLANEJADO`) → aba `CARTEIRA` (Power Query `carteira.pq`, chave `OS&ETAPA`)
→ aba `PMP` (PROCV) → `CARGA x CAPACIDADE` e `PMP MENSAL`.

- `PMP`: 1 linha por painel. Mês e horas planejadas vêm da carteira; horas realizadas e "Concluído" são digitados.
- `CARGA x CAPACIDADE`: saldo (planejado − realizado) por centro × capacidade, backlog acumulado,
  conferência carteira × PMP e aderência ao tempo padrão.
- `PMP MENSAL`: painéis prontos para expedir por projeto e mês.

O script é só o *bootstrap*: depois de gerada, a planilha se atualiza pelo Power Query.

> **Não versionar a carteira nem o PMP gerado neste repositório** — ele é público e os arquivos têm dados de clientes.
