// =====================================================================
//  CONTROLE DE ESTOQUE — CONSULTAS POWER QUERY (M)
//  Como usar: no Power BI Desktop > Página Inicial > Transformar dados >
//  Nova Fonte > Consulta Nula. Abra o "Editor Avançado", apague tudo e
//  cole o bloco da consulta. Renomeie a consulta EXATAMENTE com o nome
//  indicado em "NOME DA CONSULTA". Crie na ordem abaixo.
//
//  Consultas carregadas no modelo (as demais: botão direito >
//  desmarcar "Habilitar carga"):
//      Estoque, Historico, Parametros
// =====================================================================


// ---------------------------------------------------------------------
// NOME DA CONSULTA: pPastaBase        (NÃO carregar)
// Caminho da pasta "ControleEstoque" — termine com a barra "\".
// Depois de criar, você pode transformá-la em Parâmetro:
// Página Inicial > Gerenciar Parâmetros > Novo (tipo Texto) e colar o valor.
// Se a pasta estiver no OneDrive/SharePoint sincronizado, use o caminho
// local sincronizado (ex.: C:\Users\SEU_USUARIO\OneDrive - EMPRESA\ControleEstoque\)
// ou veja a variante SharePoint no final deste arquivo.
// ---------------------------------------------------------------------
"C:\ControleEstoque\" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]


// ---------------------------------------------------------------------
// NOME DA CONSULTA: Funcoes           (NÃO carregar)
// Biblioteca de funções usadas pelas outras consultas.
// ---------------------------------------------------------------------
let
    Funcoes = [

        // Texto em MAIÚSCULAS, sem acentos, só letras/números, com espaço nas pontas.
        // Ex.: "Pote 320ml - AÇAÍ" -> " POTE 320ML ACAI "
        Normalizar = (t as any) as text =>
            let
                s        = if t = null then "" else Text.Upper(Text.From(t)),
                mapa     = {{"Á","A"},{"À","A"},{"Â","A"},{"Ã","A"},{"Ä","A"},{"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},
                            {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},{"Ó","O"},{"Ò","O"},{"Ô","O"},{"Õ","O"},{"Ö","O"},
                            {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},{"Ç","C"},{"Ñ","N"}},
                semAc    = List.Accumulate(mapa, s, (acc, p) => Text.Replace(acc, p{0}, p{1})),
                soAlfa   = Text.Combine(List.Transform(Text.ToList(semAc),
                               each if Text.Contains("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", _) then _ else " ")),
                compacto = Text.Combine(List.Select(Text.Split(soAlfa, " "), each _ <> ""), " ")
            in
                " " & compacto & " ",

        // Converte número vindo como número ou texto pt-BR ("1.000", "12,5"); erros/#N/A -> null
        Num = (v as any) as nullable number =>
            if v = null then null
            else if Value.Is(v, type number) then v
            else
                let t = Text.Trim(Text.From(v))
                in  if t = "" or Text.StartsWith(t, "#") then null
                    else try Number.From(t, "pt-BR") otherwise null,

        // Código sempre como texto ("7103")
        Codigo = (v as any) as nullable text =>
            if v = null then null
            else
                let t = Text.Trim(if Value.Is(v, type number) then Number.ToText(v, "0") else Text.From(v))
                in  if t = "" or Text.StartsWith(t, "#") then null else t,

        // Número de OS como texto. Trata "27.313" (texto ou decimal) -> "27313"
        OS = (v as any) as nullable text =>
            if v = null then null
            else if Value.Is(v, type number) then
                Number.ToText(if v <> Number.Round(v) and v < 1000 then Number.Round(v * 1000) else v, "0")
            else
                let t = Text.Remove(Text.Trim(Text.From(v)), {".", " "})
                in  if t = "" or Text.StartsWith(t, "#") then null else t,

        // Data a partir de data, número serial do Excel ou texto "dd/mm", "dd/mm/aa", "dd/mm/aaaa"
        Data = (v as any, anoPadrao as number) as nullable date =>
            if v = null then null
            else if Value.Is(v, type date) then v
            else if Value.Is(v, type datetime) then Date.From(v)
            else if Value.Is(v, type datetimezone) then Date.From(DateTimeZone.RemoveZone(v))
            else if Value.Is(v, type number) then (try Date.From(v) otherwise null)
            else
                let
                    t = Text.Trim(Text.From(v)),
                    p = Text.Split(Text.Replace(Text.Replace(t, "-", "/"), ".", "/"), "/")
                in
                    if List.Count(p) < 2 then null
                    else try #date(
                            if List.Count(p) >= 3 then
                                (let y = Number.From(p{2}) in if y < 100 then 2000 + y else y)
                            else anoPadrao,
                            Number.From(p{1}),
                            Number.From(p{0}))
                         otherwise null,

        // Data de referência do arquivo diário: procura "aaaa-mm-dd" no nome
        // (ex.: Estoque_2026-09-29.xlsx). Se não achar, usa a data de modificação.
        DataDoArquivo = (nome as text, modificado as any) as date =>
            let
                posicoes   = List.Numbers(0, List.Max({Text.Length(nome) - 9, 0})),
                candidatas = List.Transform(posicoes, (i) =>
                                try Date.FromText(Text.Middle(nome, i, 10), [Format = "yyyy-MM-dd"]) otherwise null),
                achou      = List.First(List.RemoveNulls(candidatas), null)
            in
                if achou <> null then achou else Date.From(modificado),

        // Hoje no fuso de Brasília (funciona igual no Desktop e no Serviço)
        Hoje = () as date => Date.From(DateTimeZone.SwitchZone(DateTimeZone.UtcNow(), -3)),

        // Lê um arquivo diário (xlsx/xlsm/csv) e devolve colunas padronizadas.
        // Encontra sozinho a linha de cabeçalho (a que contém "Cód Novo").
        LerPlanilha = (conteudo as binary, extensao as text) as table =>
            let
                ext  = Text.Lower(extensao),
                Bruto =
                    if ext = ".csv" then
                        let
                            txt   = Text.FromBinary(conteudo, 65001),
                            linha = List.First(Lines.FromText(txt), ""),
                            delim = if Text.Contains(linha, ";") then ";" else ","
                        in  Csv.Document(conteudo, [Delimiter = delim, Encoding = 65001, QuoteStyle = QuoteStyle.Csv])
                    else
                        let
                            wb    = Excel.Workbook(conteudo, null, true),
                            abas  = Table.SelectRows(wb, each [Kind] = "Sheet"),
                            pref  = Table.SelectRows(abas, each Text.Contains(Text.Lower([Item]), "estoque")),
                            aba   = if Table.RowCount(pref) > 0 then pref{0}[Data] else abas{0}[Data]
                        in  aba,
                SemErro  = Table.ReplaceErrorValues(Bruto, List.Transform(Table.ColumnNames(Bruto), each {_, null})),
                Linhas   = Table.ToRows(SemErro),
                PosCab   = List.PositionOf(
                              List.Transform(Linhas, (r) => List.Contains(List.Transform(r, each Text.Trim(Text.From(_ ?? ""))), "Cód Novo")),
                              true),
                SemTopo  = if PosCab > 0 then Table.Skip(SemErro, PosCab) else SemErro,
                Cab      = Table.PromoteHeaders(SemTopo, [PromoteAllScalars = true]),
                NomesOk  = Table.TransformColumnNames(Cab, Text.Trim),
                Mapa     = {
                              {"Cód Antigo", "CodAntigo"}, {"Cód Novo", "Codigo"}, {"Descrição produto", "Descricao"},
                              {"Média Mensal", "MediaMensal"}, {"Quantidade por caixa", "QtdPorCaixa"},
                              {"Estoque", "Estoque"}, {"Empenhado", "Empenhado"}, {"Disponível", "DisponivelPlanilha"},
                              {"OS", "OS_Planilha"}, {"Última Movimentação", "Mov_Planilha"},
                              {"Previsão", "Prev_Planilha"}, {"Observação", "Obs_Planilha"}
                           },
                Sel      = Table.SelectColumns(NomesOk, List.Transform(Mapa, each _{0}), MissingField.UseNull),
                Ren      = Table.RenameColumns(Sel, Mapa, MissingField.Ignore)
            in
                Ren
    ]
in
    Funcoes


// ---------------------------------------------------------------------
// NOME DA CONSULTA: stgControle       (NÃO carregar)
// Abre o arquivo de campos editáveis.
// ---------------------------------------------------------------------
let
    Fonte = Excel.Workbook(File.Contents(pPastaBase & "Controle_Estoque.xlsx"), true, true)
in
    Fonte


// ---------------------------------------------------------------------
// NOME DA CONSULTA: cfg               (NÃO carregar)
// Parâmetros como registro: cfg[LimiteCritico], cfg[MargemAtencao] ...
// ---------------------------------------------------------------------
let
    T       = stgControle{[Item = "tbParametros", Kind = "Table"]}[Data],
    Validos = Table.SelectRows(T, each [#"Parâmetro"] <> null and Text.Trim(Text.From([#"Parâmetro"])) <> ""),
    Lido    = Record.FromList(
                  List.Transform(Validos[Valor], each Funcoes[Num](_)),
                  List.Transform(Validos[#"Parâmetro"], each Text.Trim(Text.From(_)))),
    Padrao  = [LimiteCritico = 40, MargemAtencao = 15, DiasPorMes = 30, DiasAvisoPrazo = 5,
               DiasSemMovimentacao = 30, DiasHistorico = 180],
    // Valor da planilha prevalece; se faltar/estiver vazio, usa o padrão
    Final   = Record.FromList(
                  List.Transform(Record.FieldNames(Padrao),
                      (n) => let v = Record.FieldOrDefault(Lido, n, null) in if v = null then Record.Field(Padrao, n) else v),
                  Record.FieldNames(Padrao))
in
    Final


// ---------------------------------------------------------------------
// NOME DA CONSULTA: Parametros        (CARREGAR — usado nas medidas DAX)
// ---------------------------------------------------------------------
let
    Tabela = Table.FromRecords({cfg}),
    Tipos  = Table.TransformColumnTypes(Tabela,
                 List.Transform(Table.ColumnNames(Tabela), each {_, Int64.Type}))
in
    Tipos


// ---------------------------------------------------------------------
// NOME DA CONSULTA: stgDiario         (NÃO carregar)
// Junta TODOS os arquivos da pasta 01_Estoque_Diario (um por dia).
// O arquivo mais recente = estoque atual; os demais = histórico.
// Se houver 2 arquivos com a mesma data, vale o modificado por último.
// ---------------------------------------------------------------------
let
    Fonte     = Folder.Files(pPastaBase & "01_Estoque_Diario"),
    Validos   = Table.SelectRows(Fonte, each
                    List.Contains({".xlsx", ".xlsm", ".csv"}, Text.Lower([Extension]))
                    and not Text.StartsWith([Name], "~$")),
    ComData   = Table.AddColumn(Validos, "DataRef", each Funcoes[DataDoArquivo]([Name], [Date modified]), type date),
    Ordenado  = Table.Buffer(Table.Sort(ComData, {{"DataRef", Order.Descending}, {"Date modified", Order.Descending}})),
    UmPorDia  = Table.Distinct(Ordenado, {"DataRef"}),
    DataMax   = List.Max(UmPorDia[DataRef]),
    Janela    = Table.SelectRows(UmPorDia, each [DataRef] > Date.AddDays(DataMax, -cfg[DiasHistorico])),
    ComDados  = Table.AddColumn(Janela, "Dados", each Funcoes[LerPlanilha]([Content], [Extension])),
    SoPrecisa = Table.SelectColumns(ComDados, {"DataRef", "Name", "Dados"}),
    Expandido = Table.ExpandTableColumn(SoPrecisa, "Dados",
                    {"CodAntigo","Codigo","Descricao","MediaMensal","QtdPorCaixa","Estoque","Empenhado",
                     "DisponivelPlanilha","OS_Planilha","Mov_Planilha","Prev_Planilha","Obs_Planilha"}),
    Convertido = Table.TransformColumns(Expandido, {
                    {"Codigo",             Funcoes[Codigo], type nullable text},
                    {"CodAntigo",          Funcoes[Codigo], type nullable text},
                    {"Descricao",          each if _ = null then null else Text.Trim(Text.From(_)), type nullable text},
                    {"MediaMensal",        Funcoes[Num], type nullable number},
                    {"QtdPorCaixa",        Funcoes[Num], type nullable number},
                    {"Estoque",            Funcoes[Num], type nullable number},
                    {"Empenhado",          Funcoes[Num], type nullable number},
                    {"DisponivelPlanilha", Funcoes[Num], type nullable number},
                    {"OS_Planilha",        Funcoes[OS],  type nullable text},
                    {"Mov_Planilha",       each if _ = null or Text.Trim(Text.From(_)) = "" then null else Text.Trim(Text.From(_)), type nullable text},
                    {"Obs_Planilha",       each if _ = null or Text.Trim(Text.From(_)) = "" then null else Text.Trim(Text.From(_)), type nullable text}
                 }),
    ComPrazo  = Table.AddColumn(Convertido, "Prazo_Planilha", each Funcoes[Data]([Prev_Planilha], Date.Year([DataRef])), type nullable date),
    // Texto na coluna Previsão que não é data (ex.: "Liberando O.S") vira observação
    ComObs    = Table.ReplaceValue(ComPrazo, each [Obs_Planilha],
                    each if [Obs_Planilha] = null and [Prev_Planilha] <> null and [Prazo_Planilha] = null
                         then Text.Trim(Text.From([Prev_Planilha])) else [Obs_Planilha],
                    Replacer.ReplaceValue, {"Obs_Planilha"}),
    SemPrev   = Table.RemoveColumns(ComObs, {"Prev_Planilha"}),
    ComCodigo = Table.SelectRows(SemPrev, each [Codigo] <> null and [Codigo] <> "Cód Novo"),
    Unicos    = Table.Distinct(ComCodigo, {"DataRef", "Codigo"}),
    // Disponível = Estoque - Empenhado (se a planilha não tiver estoque, usa o Disponível dela)
    ComDisp   = Table.AddColumn(Unicos, "Disponivel", each
                    if [Estoque] = null and [Empenhado] = null then [DisponivelPlanilha]
                    else (([Estoque] ?? 0) - ([Empenhado] ?? 0)), type nullable number)
in
    ComDisp


// ---------------------------------------------------------------------
// NOME DA CONSULTA: Controle          (NÃO carregar)
// Campos editáveis (tabela tbControleOS) + regras da lista de movimentações.
// ---------------------------------------------------------------------
let
    T        = stgControle{[Item = "tbControleOS", Kind = "Table"]}[Data],
    Ren      = Table.RenameColumns(T, {
                   {"Código", "Codigo"}, {"O.S.", "OS_Controle"}, {"Última Movimentação", "Mov_Controle"},
                   {"Data Movimentação", "DataMovimentacao"}, {"Prazo", "Prazo_Controle"},
                   {"Franquia (ajuste manual)", "FranquiaManual"}, {"Categoria (ajuste manual)", "CategoriaManual"},
                   {"Consumo Mensal Manual", "ConsumoMensalManual"}, {"Observação", "Obs_Controle"}
               }, MissingField.Ignore),
    Txt      = (v) => if v = null or Text.Trim(Text.From(v)) = "" then null else Text.Trim(Text.From(v)),
    Conv     = Table.TransformColumns(Ren, {
                   {"Codigo",              Funcoes[Codigo], type nullable text},
                   {"OS_Controle",         Funcoes[OS],     type nullable text},
                   {"Mov_Controle",        Txt, type nullable text},
                   {"DataMovimentacao",    each Funcoes[Data](_, Date.Year(Funcoes[Hoje]())), type nullable date},
                   {"Prazo_Controle",      each Funcoes[Data](_, Date.Year(Funcoes[Hoje]())), type nullable date},
                   {"FranquiaManual",      Txt, type nullable text},
                   {"CategoriaManual",     Txt, type nullable text},
                   {"ConsumoMensalManual", Funcoes[Num], type nullable number},
                   {"Obs_Controle",        Txt, type nullable text}
               }),
    Sel      = Table.SelectColumns(Conv, {"Codigo","OS_Controle","Mov_Controle","DataMovimentacao","Prazo_Controle",
                                          "FranquiaManual","CategoriaManual","ConsumoMensalManual","Obs_Controle"},
                                   MissingField.UseNull),
    Validos  = Table.SelectRows(Sel, each [Codigo] <> null),
    Unicos   = Table.Distinct(Validos, {"Codigo"})
in
    Unicos


// ---------------------------------------------------------------------
// NOME DA CONSULTA: Historico         (CARREGAR)
// Uma linha por produto por dia (snapshot). Calcula dias de estoque,
// status, estoque anterior, entradas e saídas.
// ---------------------------------------------------------------------
let
    Lim      = cfg[LimiteCritico],
    Margem   = cfg[MargemAtencao],
    DiasMes  = cfg[DiasPorMes],

    Base     = Table.SelectColumns(stgDiario, {"DataRef","Codigo","MediaMensal","QtdPorCaixa","Estoque","Empenhado","Disponivel"}),
    ComCtrl  = Table.NestedJoin(Base, {"Codigo"}, Table.SelectColumns(Controle, {"Codigo","ConsumoMensalManual"}),
                                {"Codigo"}, "c", JoinKind.LeftOuter),
    ExpCtrl  = Table.ExpandTableColumn(ComCtrl, "c", {"ConsumoMensalManual"}),
    Consumo  = Table.AddColumn(ExpCtrl, "ConsumoMensal", each [ConsumoMensalManual] ?? [MediaMensal], type nullable number),
    Diario   = Table.AddColumn(Consumo, "ConsumoDiario", each
                   if [ConsumoMensal] = null or [ConsumoMensal] <= 0 then null else [ConsumoMensal] / DiasMes, type nullable number),
    Dias     = Table.AddColumn(Diario, "DiasEstoque", each
                   if [ConsumoDiario] = null or [Disponivel] = null then null
                   else Number.Round([Disponivel] / [ConsumoDiario], 0), type nullable number),
    Status   = Table.AddColumn(Dias, "StatusEstoque", each
                   if [Estoque] = null and [Disponivel] = null then "Sem dados"
                   else if [DiasEstoque] = null then (if ([Disponivel] ?? 0) < 0 then "Crítico" else "Sem consumo")
                   else if [DiasEstoque] < Lim then "Crítico"
                   else if [DiasEstoque] < Lim + Margem then "Atenção"
                   else "Seguro", type text),

    // Estoque/Status do snapshot anterior do MESMO produto
    Agrupado = Table.Group(Status, {"Codigo"}, {{"t", (g) =>
                   let
                       s     = Table.Buffer(Table.Sort(g, {{"DataRef", Order.Ascending}})),
                       est   = List.Buffer(s[Estoque]),
                       sts   = List.Buffer(s[StatusEstoque]),
                       dts   = List.Buffer(s[DataRef]),
                       idx   = Table.AddIndexColumn(s, "i", 0, 1, Int64.Type),
                       ant   = Table.AddColumn(idx, "EstoqueAnterior", each if [i] = 0 then null else est{[i] - 1}, type nullable number),
                       stAnt = Table.AddColumn(ant, "StatusAnterior", each if [i] = 0 then null else sts{[i] - 1}, type nullable text),
                       dtAnt = Table.AddColumn(stAnt, "DataAnterior", each if [i] = 0 then null else dts{[i] - 1}, type nullable date)
                   in
                       Table.RemoveColumns(dtAnt, {"i"}), type table}}),
    Expandir = Table.ExpandTableColumn(Agrupado, "t",
                   {"DataRef","MediaMensal","ConsumoMensalManual","QtdPorCaixa","Estoque","Empenhado","Disponivel",
                    "ConsumoMensal","ConsumoDiario","DiasEstoque","StatusEstoque","EstoqueAnterior","StatusAnterior","DataAnterior"}),
    Var      = Table.AddColumn(Expandir, "Variacao", each
                   if [EstoqueAnterior] = null or [Estoque] = null then null else [Estoque] - [EstoqueAnterior], type nullable number),
    Entradas = Table.AddColumn(Var, "Entradas", each if [Variacao] <> null and [Variacao] > 0 then [Variacao] else 0, type number),
    Saidas   = Table.AddColumn(Entradas, "Saidas", each if [Variacao] <> null and [Variacao] < 0 then -[Variacao] else 0, type number),
    Tipos    = Table.TransformColumnTypes(Saidas, {{"DataRef", type date}, {"Codigo", type text},
                   {"MediaMensal", type number}, {"ConsumoMensalManual", type number}, {"QtdPorCaixa", type number},
                   {"Estoque", type number}, {"Empenhado", type number}, {"Disponivel", type number},
                   {"ConsumoMensal", type number}, {"ConsumoDiario", type number}, {"DiasEstoque", Int64.Type},
                   {"EstoqueAnterior", type number}})
in
    Tipos


// ---------------------------------------------------------------------
// NOME DA CONSULTA: Estoque           (CARREGAR — tabela principal)
// Uma linha por produto, com a posição do arquivo mais recente +
// campos editáveis + franquia/categoria + regras de OS e movimentação.
// ---------------------------------------------------------------------
let
    Hoje      = Funcoes[Hoje](),
    DataAtual = List.Max(Historico[DataRef]),
    PrimeiraData = List.Min(Historico[DataRef]),
    TemAnterior  = List.Count(List.Distinct(Historico[DataRef])) > 1,

    // --- posição atual ---
    Atual     = Table.SelectRows(Historico, each [DataRef] = DataAtual),
    Cadastro  = Table.SelectColumns(Table.SelectRows(stgDiario, each [DataRef] = DataAtual),
                    {"Codigo","CodAntigo","Descricao","OS_Planilha","Mov_Planilha","Prazo_Planilha","Obs_Planilha"}),
    J1        = Table.NestedJoin(Atual, {"Codigo"}, Cadastro, {"Codigo"}, "cad", JoinKind.LeftOuter),
    E1        = Table.ExpandTableColumn(J1, "cad", {"CodAntigo","Descricao","OS_Planilha","Mov_Planilha","Prazo_Planilha","Obs_Planilha"}),

    // --- última alteração de estoque (a partir do histórico) ---
    Mudancas  = Table.SelectRows(Historico, each [EstoqueAnterior] <> null and [Estoque] <> [EstoqueAnterior]),
    UltMud    = Table.Group(Mudancas, {"Codigo"}, {{"DataUltimaAlteracao", each List.Max([DataRef]), type date}}),
    J2        = Table.NestedJoin(E1, {"Codigo"}, UltMud, {"Codigo"}, "um", JoinKind.LeftOuter),
    E2        = Table.ExpandTableColumn(J2, "um", {"DataUltimaAlteracao"}),

    // --- campos editáveis (Controle prevalece; vazio -> usa a planilha diária) ---
    J3        = Table.NestedJoin(E2, {"Codigo"}, Controle, {"Codigo"}, "ctl", JoinKind.LeftOuter),
    E3        = Table.ExpandTableColumn(J3, "ctl",
                    {"OS_Controle","Mov_Controle","DataMovimentacao","Prazo_Controle","FranquiaManual","CategoriaManual","Obs_Controle"}),
    OS        = Table.AddColumn(E3, "OS", each [OS_Controle] ?? [OS_Planilha], type nullable text),
    Mov       = Table.AddColumn(OS, "UltimaMovimentacao", each [Mov_Controle] ?? [Mov_Planilha], type nullable text),
    Prazo     = Table.AddColumn(Mov, "Prazo", each [Prazo_Controle] ?? [Prazo_Planilha], type nullable date),
    Obs       = Table.AddColumn(Prazo, "Observacao", each [Obs_Controle] ?? [Obs_Planilha], type nullable text),

    // --- franquia e categoria pela descrição (tabelas tbFranquias / tbCategorias) ---
    tbFr      = Table.Buffer(Table.Sort(Table.SelectRows(stgControle{[Item = "tbFranquias", Kind = "Table"]}[Data],
                    each [#"Palavra-chave"] <> null and Text.Upper(Text.From([Ativa] ?? "Sim")) <> "NÃO"
                         and Text.Upper(Text.From([Ativa] ?? "Sim")) <> "NAO"), {{"Prioridade", Order.Ascending}})),
    frChaves  = List.Buffer(List.Transform(tbFr[#"Palavra-chave"], Funcoes[Normalizar])),
    frNomes   = List.Buffer(tbFr[Franquia]),
    tbCat     = Table.Buffer(Table.Sort(Table.SelectRows(stgControle{[Item = "tbCategorias", Kind = "Table"]}[Data],
                    each [#"Palavra-chave"] <> null), {{"Prioridade", Order.Ascending}})),
    catChaves = List.Buffer(List.Transform(tbCat[#"Palavra-chave"], each Text.Trim(Funcoes[Normalizar](_)))),
    catNomes  = List.Buffer(tbCat[Categoria]),
    Achar     = (desc, chaves as list, nomes as list, padrao as text) =>
                    let n = Funcoes[Normalizar](desc),
                        p = List.PositionOf(List.Transform(chaves, each Text.Contains(n, _)), true)
                    in  if p = -1 then padrao else nomes{p},
    Fr        = Table.AddColumn(Obs, "Franquia", each [FranquiaManual] ?? Achar([Descricao], frChaves, frNomes, "Não identificada"), type text),
    Cat       = Table.AddColumn(Fr, "Categoria", each [CategoriaManual] ?? Achar([Descricao], catChaves, catNomes, "Outros"), type text),

    // --- regras de OS ---
    tbMov     = stgControle{[Item = "tbMovimentacoes", Kind = "Table"]}[Data],
    Encerram  = List.Buffer(List.Transform(
                    Table.SelectRows(tbMov, each Text.Upper(Text.From([#"Encerra OS"] ?? "")) = "SIM")[#"Movimentação"],
                    each Text.Upper(Text.Trim(_)))),
    OSAberta  = Table.AddColumn(Cat, "OSAberta", each
                    [OS] <> null and not List.Contains(Encerram, Text.Upper([UltimaMovimentacao] ?? "")), type logical),
    DiasPrazo = Table.AddColumn(OSAberta, "DiasParaPrazo", each
                    if [Prazo] = null then null else Duration.Days([Prazo] - Hoje), Int64.Type),
    StatusOS  = Table.AddColumn(DiasPrazo, "StatusOS", each
                    if [OS] = null then "Sem OS"
                    else if not [OSAberta] then "Concluída"
                    else if [Prazo] = null then "Sem prazo"
                    else if [DiasParaPrazo] < 0 then "Prazo vencido"
                    else if [DiasParaPrazo] <= cfg[DiasAvisoPrazo] then "Próximo do vencimento"
                    else "Dentro do prazo", type text),

    // --- sem movimentação ---
    UltAtiv   = Table.AddColumn(StatusOS, "UltimaAtividade", each List.Max({[DataMovimentacao], [DataUltimaAlteracao]}), type nullable date),
    DiasSem   = Table.AddColumn(UltAtiv, "DiasSemMovimentacao", each
                    if [UltimaAtividade] = null then null else Duration.Days(Hoje - [UltimaAtividade]), Int64.Type),
    SemMov    = Table.AddColumn(DiasSem, "SemMovimentacao", each
                    if Text.Upper([UltimaMovimentacao] ?? "") = "SEM MOVIMENTAÇÃO" then true
                    else if [UltimaAtividade] <> null then [DiasSemMovimentacao] > cfg[DiasSemMovimentacao]
                    else if Duration.Days(Hoje - PrimeiraData) >= cfg[DiasSemMovimentacao] then true
                    else [UltimaMovimentacao] = null, type logical),

    // --- alertas ---
    EmAlerta  = Table.AddColumn(SemMov, "EmAlerta", each List.Contains({"Crítico","Atenção"}, [StatusEstoque]), type logical),
    Novo      = Table.AddColumn(EmAlerta, "NovoAlerta", each
                    [EmAlerta] and (not TemAnterior or [DataAnterior] = null or not List.Contains({"Crítico","Atenção"}, [StatusAnterior] ?? "")),
                    type logical),
    Ordem     = Table.AddColumn(Novo, "OrdemStatus", each
                    List.PositionOf({"Crítico","Atenção","Seguro","Sem consumo","Sem dados"}, [StatusEstoque]) + 1, Int64.Type),

    // --- unidades (estoque está em caixas/pacotes na planilha) ---
    EstUn     = Table.AddColumn(Ordem, "EstoqueUnidades", each
                    if [Estoque] = null or [QtdPorCaixa] = null then null else [Estoque] * [QtdPorCaixa], type nullable number),
    DispUn    = Table.AddColumn(EstUn, "DisponivelUnidades", each
                    if [Disponivel] = null or [QtdPorCaixa] = null then null else [Disponivel] * [QtdPorCaixa], type nullable number),
    // Reposição sugerida (cx) para voltar à faixa segura
    Repos     = Table.AddColumn(DispUn, "ReposicaoSugerida", each
                    if [ConsumoDiario] = null then 0
                    else List.Max({0, Number.RoundUp((cfg[LimiteCritico] + cfg[MargemAtencao]) * [ConsumoDiario] - ([Disponivel] ?? 0), 0)}),
                    type number),
    // Campo único de pesquisa: código, descrição ou OS
    Busca     = Table.AddColumn(Repos, "Busca", each
                    Text.Combine(List.RemoveNulls({[Codigo], [Descricao], if [OS] = null then null else "OS " & [OS]}), " | "), type text),

    Final     = Table.SelectColumns(Busca, {
                    "DataRef","Codigo","CodAntigo","Descricao","Franquia","Categoria",
                    "MediaMensal","ConsumoMensalManual","ConsumoMensal","ConsumoDiario","QtdPorCaixa",
                    "Estoque","Empenhado","Disponivel","EstoqueUnidades","DisponivelUnidades",
                    "DiasEstoque","StatusEstoque","OrdemStatus","StatusAnterior","EmAlerta","NovoAlerta","ReposicaoSugerida",
                    "OS","OSAberta","UltimaMovimentacao","DataMovimentacao","Prazo","DiasParaPrazo","StatusOS",
                    "DataUltimaAlteracao","UltimaAtividade","DiasSemMovimentacao","SemMovimentacao","Observacao","Busca"})
in
    Final


// =====================================================================
//  VARIANTES (opcional)
// =====================================================================

// (A) Pasta diária no SharePoint (sem OneDrive sincronizado). Em stgDiario,
//     troque a linha "Fonte = Folder.Files(...)" por:
//
//     Fonte = Table.SelectRows(
//                 SharePoint.Files("https://SUAEMPRESA.sharepoint.com/sites/SEUSITE", [ApiVersion = 15]),
//                 each Text.Contains([Folder Path], "/ControleEstoque/01_Estoque_Diario/")),
//
//     e em stgControle:
//
//     Fonte = Excel.Workbook(
//                 SharePoint.Files("https://SUAEMPRESA.sharepoint.com/sites/SEUSITE", [ApiVersion = 15])
//                     {[Name = "Controle_Estoque.xlsx"]}[Content], true, true)

// (B) Campos editáveis numa Lista do SharePoint (edição pelo Power Apps dentro
//     do relatório — ver guia). Substitua a consulta "Controle" por:
//
// let
//     Fonte  = SharePoint.Tables("https://SUAEMPRESA.sharepoint.com/sites/SEUSITE", [Implementation = "2.0", ViewMode = "All"]),
//     Lista  = Fonte{[Title = "ControleOS"]}[Items],
//     Sel    = Table.SelectColumns(Lista, {"Codigo","OS","UltimaMovimentacao","DataMovimentacao","Prazo",
//                                          "FranquiaManual","CategoriaManual","ConsumoMensalManual","Observacao"}, MissingField.UseNull),
//     Ren    = Table.RenameColumns(Sel, {{"OS","OS_Controle"},{"UltimaMovimentacao","Mov_Controle"},
//                                        {"Prazo","Prazo_Controle"},{"Observacao","Obs_Controle"}}),
//     Conv   = Table.TransformColumns(Ren, {{"Codigo", Funcoes[Codigo]}, {"OS_Controle", Funcoes[OS]},
//                                           {"DataMovimentacao", each Funcoes[Data](_, 2026), type nullable date},
//                                           {"Prazo_Controle", each Funcoes[Data](_, 2026), type nullable date},
//                                           {"ConsumoMensalManual", Funcoes[Num], type nullable number}}),
//     Unicos = Table.Distinct(Table.SelectRows(Conv, each [Codigo] <> null), {"Codigo"})
// in
//     Unicos
//     (Obs.: coluna de escolha "UltimaMovimentacao" pode vir como registro — expanda o campo "Value".)
