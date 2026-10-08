// Power Query (linguagem M) — consulta "CARTEIRA"
// Traz a tabela PLANEJADO da carteira FGI 8.5-01.02 exatamente como ela é (mesmas colunas, mesma ordem, todas as linhas).
// O PMP faz os filtros (ATIVO = SIM, ORIGEM <> CANCELADO) nas próprias fórmulas.
// Troque o caminho abaixo pelo caminho da sua carteira (ou use Dados > Obter Dados > Do SharePoint, se ela estiver no SharePoint/OneDrive).
let
    Caminho = "C:\Users\SEU_USUARIO\...\FGI 8.5-01.02 - Carteira de Projetos.xlsx",
    Fonte = Excel.Workbook(File.Contents(Caminho), null, true),
    Planejado = Fonte{[Item="PLANEJADO", Kind="Table"]}[Data],
    ComOS = Table.SelectRows(Planejado, each [OS] <> null),
    Tipos = Table.TransformColumnTypes(ComOS, {{"OS", Int64.Type}, {"Mês Planejado", type date}, {"Horas / Mês", type number}}),
    SemErros = Table.ReplaceErrorValues(Tipos, List.Transform(Table.ColumnNames(Tipos), each {_, null}))
in
    SemErros
