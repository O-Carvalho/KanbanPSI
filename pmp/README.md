# PMP — Plano Mestre de Produção

Gera a planilha do PMP a partir da carteira **FGI 8.5-01.02 - Carteira de Projetos**.

```bash
python3 pmp/gerar_pmp.py "FGI_8.5-01.02_-_Carteira_de_Projetos.xlsx" PMP_PSI.xlsx pmp/carteira.pq
```

Fluxo: carteira (tabela `PLANEJADO`) → aba `CARTEIRA` (tabela `PLANEJADO`, query do Power Query) e cadastro do SharePoint → aba `DB PROJETOS` (tabela `DB_PROJETOS`)
→ aba `PMP` (PROCV; etapas MEC → ELE → CDP → PLAT → NORM/INSP/EMB) → `CARGA x CAPACIDADE` e `PMP MENSAL`.

- `PMP`: 1 linha por painel. Dados do painel por PROCV na `DB_PROJETOS` (origem na `PLANEJADO`); mês e horas de cada etapa por MÍNIMOSES/SOMASES (OS + etapa + ativo); horas realizadas e "Concluído" são digitados.
- `CARGA x CAPACIDADE`: saldo (planejado − realizado) por centro × capacidade, backlog acumulado,
  conferência carteira × PMP e aderência ao tempo padrão.
- `PMP MENSAL`: painéis prontos para expedir por projeto e mês.

As fórmulas usam referência estruturada (`PLANEJADO[Horas / Mês]`, `DB_PROJETOS[[OS]:[GIGA]]`), então sobrevivem ao recarregamento das queries.

O script é só o *bootstrap*: depois de gerada, a planilha se atualiza pelo Power Query.

> **Não versionar a carteira nem o PMP gerado neste repositório** — ele é público e os arquivos têm dados de clientes.
