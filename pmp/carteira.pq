// Power Query (linguagem M) — consulta "CARTEIRA"
// Lê a tabela PLANEJADO da carteira FGI 8.5-01.02 e entrega 1 linha por OS × ETAPA, com a coluna CHAVE (OS&ETAPA) usada pelos PROCV do PMP.
// Troque o caminho abaixo pelo caminho da sua carteira (ou use Dados > Obter Dados > Do SharePoint, se ela estiver no SharePoint/OneDrive).
let
    Caminho = "C:\Users\SEU_USUARIO\...\FGI 8.5-01.02 - Carteira de Projetos.xlsx",
    Fonte = Excel.Workbook(File.Contents(Caminho), null, true),
    Planejado = Fonte{[Item="PLANEJADO", Kind="Table"]}[Data],
    Ativos = Table.SelectRows(Planejado, each [OS] <> null and [PROJETO ATIVO PLANEJAMENTO] = "SIM" and [ORIGEM] <> "CANCELADO"),
    Colunas = Table.SelectColumns(Ativos, {"OS", "CÓDIGO DO PROJETO", "SUBESTAÇÃO", "TAG", "TIPO PNL", "COMPLEXIDADE", "ORIGEM", "ETAPA", "Mês Planejado", "Horas / Mês", "GIGA"}),
    Tipos = Table.TransformColumnTypes(Colunas, {{"OS", Int64.Type}, {"Mês Planejado", type date}, {"Horas / Mês", type number}}),
    SemErros = Table.ReplaceErrorValues(Tipos, {{"Mês Planejado", null}, {"Horas / Mês", 0}}),
    ComChave = Table.AddColumn(SemErros, "CHAVE", each Text.From([OS]) & [ETAPA], type text),
    Ordem = Table.ReorderColumns(ComChave, {"CHAVE", "OS", "CÓDIGO DO PROJETO", "SUBESTAÇÃO", "TAG", "TIPO PNL", "COMPLEXIDADE", "ORIGEM", "ETAPA", "Mês Planejado", "Horas / Mês", "GIGA"})
in
    Ordem
