"""Gera o projeto Power BI (PBIP) "Eventos de curta duração".

Uso: python3 powerbi/src/build_pbip.py
Saída: powerbi/projeto/ (Eventos.pbip, Eventos.SemanticModel/, Eventos.Report/)

Fontes editáveis nesta pasta:
  tema-azulterra.json   tema escuro do relatório
  estilo-html.css       estilo dos visuais HTML Content
  ../deneb/*.vl.json    gráficos Deneb (Vega-Lite): linha do tempo e calendário
"""
import json
import pathlib
import shutil

SRC = pathlib.Path(__file__).resolve().parent
ROOT = SRC.parent
OUT = ROOT / "projeto"
NOME = "Eventos"

SCHEMA = "https://developer.microsoft.com/json-schemas/fabric"
S_PBIP = SCHEMA + "/pbip/pbipProperties/1.0.0/schema.json"
S_PBIR = SCHEMA + "/item/report/definitionProperties/2.0.0/schema.json"
S_PBISM = SCHEMA + "/item/semanticModel/definitionProperties/1.0.0/schema.json"
S_VERSION = SCHEMA + "/item/report/definition/versionMetadata/1.0.0/schema.json"
S_REPORT = SCHEMA + "/item/report/definition/report/1.2.0/schema.json"
S_PAGES = SCHEMA + "/item/report/definition/pagesMetadata/1.0.0/schema.json"
S_PAGE = SCHEMA + "/item/report/definition/page/1.3.0/schema.json"
S_VISUAL = SCHEMA + "/item/report/definition/visualContainer/1.3.0/schema.json"

DENEB = "deneb7E15AEF80B9E4D4F8E12924291ECE89A"
HTML = "htmlContent443BE3AD55E043BF878BED274D3A6855"

# --------------------------------------------------------------------------
# Modelo semântico (TMSL / model.bim)
# --------------------------------------------------------------------------

def esc(expr):
    """Escapa texto para HTML dentro de DAX."""
    return 'SUBSTITUTE ( SUBSTITUTE ( SUBSTITUTE ( {0}, "&", "&amp;" ), "<", "&lt;" ), ">", "&gt;" )'.format(expr)


M_EVENTOS = [
    "let",
    "    // CaminhoExcel: caminho do arquivo (Excel > Arquivo > Informações > Copiar caminho) ou arquivo local.",
    "    Bruto0 = Text.Trim(CaminhoExcel),",
    "    Caminho = Text.BeforeDelimiter(Bruto0, \"?\"),",
    "    EhWeb = Text.StartsWith(Text.Lower(Caminho), \"http\"),",
    "    Partes = if EhWeb then Uri.Parts(Bruto0) else null,",
    "    Segmentos = if EhWeb then List.Select(Text.Split(Partes[Path], \"/\"), each _ <> \"\") else {},",
    "    // Site do SharePoint/OneDrive = servidor + /sites/<nome> (ou /teams/<nome>, /personal/<nome>)",
    "    PosSite = List.PositionOfAny(List.Transform(Segmentos, Text.Lower), {\"sites\", \"teams\", \"personal\"}),",
    "    Site = if EhWeb then Partes[Scheme] & \"://\" & Partes[Host] & (if PosSite >= 0 then \"/\" & Segmentos{PosSite} & \"/\" & Segmentos{PosSite + 1} else \"\") else null,",
    "    // Nome do arquivo: parâmetro file= (links do navegador) ou último trecho do caminho",
    "    NomeArquivo = if not EhWeb then null else if Record.HasFields(Partes[Query], \"file\") then Record.Field(Partes[Query], \"file\") else List.Last(Segmentos),",
    "    Arquivos = if EhWeb then SharePoint.Files(Site, [ApiVersion = 15]) else null,",
    "    Encontrados = if EhWeb then Table.SelectRows(Arquivos, each Text.Lower([Name]) = Text.Lower(NomeArquivo) or Text.Lower(Uri.EscapeDataString([Name])) = Text.Lower(NomeArquivo)) else null,",
    "    Conteudo = if EhWeb then",
    "            (if Table.IsEmpty(Encontrados) then error Error.Record(\"Arquivo não encontrado\", \"O arquivo '\" & NomeArquivo & \"' não foi encontrado no site \" & Site & \". Confira o parâmetro CaminhoExcel.\") else Encontrados{0}[Content])",
    "        else File.Contents(Caminho),",
    "    Pasta = Excel.Workbook(Conteudo, null, true),",
    "    // 1) Tabela do Excel chamada Eventos; 2) senão, primeira planilha, com o cabeçalho localizado nas 10 primeiras linhas",
    "    TemTabela = Table.MatchesAnyRows(Pasta, each [Kind] = \"Table\" and [Item] = \"Eventos\"),",
    "    Planilha = Table.SelectRows(Pasta, each [Kind] = \"Sheet\"){0}[Data],",
    "    PosCabecalho = List.PositionOf(List.Transform(Table.ToRows(Table.FirstN(Planilha, 10)), (linha) => List.Contains(List.Transform(linha, each if _ = null then \"\" else Text.Lower(Text.Trim(Text.From(_)))), \"evento\")), true),",
    "    DaPlanilha = Table.PromoteHeaders(Table.Skip(Planilha, if PosCabecalho < 0 then 0 else PosCabecalho), [PromoteAllScalars = true]),",
    "    Bruto = if TemTabela then Pasta{[Item = \"Eventos\", Kind = \"Table\"]}[Data] else DaPlanilha,",
    "    Colunas = Table.SelectColumns(Bruto, {\"Mês\", \"Data início\", \"Data fim\", \"Evento\", \"Status\", \"Transmissão LF\", \"Observação\"}, MissingField.UseNull),",
    "    Tipos = Table.TransformColumnTypes(Colunas, {{\"Mês\", type text}, {\"Data início\", type date}, {\"Data fim\", type date}, {\"Evento\", type text}, {\"Status\", type text}, {\"Transmissão LF\", type text}, {\"Observação\", type text}}, \"pt-BR\"),",
    "    Limpo = Table.TransformColumns(Tipos, {{\"Evento\", each if _ = null then null else Text.Trim(_), type text}, {\"Status\", each if _ = null then null else Text.Trim(_), type text}}),",
    "    Validos = Table.SelectRows(Limpo, each [Evento] <> null and [Evento] <> \"\" and [#\"Data início\"] <> null),",
    "    ComFim = Table.ReplaceValue(Validos, each [#\"Data fim\"], each if [#\"Data fim\"] = null then [#\"Data início\"] else [#\"Data fim\"], Replacer.ReplaceValue, {\"Data fim\"}),",
    "    DatasOk = Table.SelectRows(ComFim, each [#\"Data fim\"] >= [#\"Data início\"]),",
    "    Ordenado = Table.Sort(DatasOk, {{\"Data início\", Order.Ascending}, {\"Evento\", Order.Ascending}}),",
    "    ComID = Table.AddIndexColumn(Ordenado, \"ID\", 1, 1, Int64.Type)",
    "in",
    "    ComID",
]


def data_col(name, dtype, fmt=None, hidden=False, summarize="none"):
    c = {"name": name, "dataType": dtype, "sourceColumn": name, "summarizeBy": summarize}
    if fmt:
        c["formatString"] = fmt
    if hidden:
        c["isHidden"] = True
    return c


def calc_col(name, dtype, expr, fmt=None, hidden=False):
    c = {"type": "calculated", "name": name, "dataType": dtype, "isDataTypeInferred": False,
         "expression": expr, "summarizeBy": "none"}
    if fmt:
        c["formatString"] = fmt
    if hidden:
        c["isHidden"] = True
    return c


def ctab_col(name, dtype, fmt=None, hidden=False, sort_by=None):
    c = {"type": "calculatedTableColumn", "name": name, "dataType": dtype, "isNameInferred": True,
         "isDataTypeInferred": True, "sourceColumn": "[" + name + "]", "summarizeBy": "none"}
    if fmt:
        c["formatString"] = fmt
    if hidden:
        c["isHidden"] = True
    if sort_by:
        c["sortByColumn"] = sort_by
    return c


def measure(name, expr, fmt=None, hidden=False):
    m = {"name": name, "expression": expr}
    if fmt:
        m["formatString"] = fmt
    if hidden:
        m["isHidden"] = True
    return m


DATA = "dd/MM/yyyy"

eventos_cols = [
    data_col("Mês", "string", hidden=True),
    data_col("Data início", "dateTime", DATA),
    data_col("Data fim", "dateTime", DATA),
    data_col("Evento", "string"),
    data_col("Status", "string"),
    data_col("Transmissão LF", "string"),
    data_col("Observação", "string"),
    data_col("ID", "int64", "0", hidden=True),
    calc_col("StatusCat", "string",
             'VAR s = LOWER ( TRIM ( Eventos[Status] ) ) RETURN SWITCH ( TRUE (), s = "confirmado", "Confirmado", s = "a confirmar", "A confirmar", s = "cancelado", "Cancelado", "Outros" )'),
    calc_col("Cor", "string",
             'SWITCH ( Eventos[StatusCat], "Confirmado", "#7AE582", "A confirmar", "#F2C75C", "Cancelado", "#F2857E", "#B9A995" )', hidden=True),
    calc_col("InicioISO", "string", 'FORMAT ( Eventos[Data início], "yyyy-MM-dd" )', hidden=True),
    calc_col("FimISO", "string", 'FORMAT ( Eventos[Data fim], "yyyy-MM-dd" )', hidden=True),
    calc_col("Lane", "int64", "MOD ( Eventos[ID] - 1, 4 )", "0", hidden=True),
    calc_col("Dias", "int64", "INT ( Eventos[Data fim] - Eventos[Data início] ) + 1", "0"),
    calc_col("Datas", "string",
             'IF ( Eventos[Data início] = Eventos[Data fim], FORMAT ( Eventos[Data início], "dd/MM/yyyy" ), FORMAT ( Eventos[Data início], "dd/MM" ) & " a " & FORMAT ( Eventos[Data fim], "dd/MM/yyyy" ) )'),
    calc_col("Dia da semana", "string",
             'VAR a = FORMAT ( Eventos[Data início], "dddd" ) VAR b = FORMAT ( Eventos[Data fim], "dddd" ) RETURN UPPER ( LEFT ( a, 1 ) ) & MID ( a, 2, 30 ) & IF ( Eventos[Data início] <> Eventos[Data fim], " a " & b )'),
    calc_col("Mês de início", "string", 'FORMAT ( Eventos[Data início], "yyyy-MM" )', hidden=True),
]

MESES_EXPR = (
    "VAR mn = MIN ( Eventos[Data início] ) "
    "VAR mx = MAX ( Eventos[Data fim] ) "
    "VAR ini = DATE ( YEAR ( mn ), MONTH ( mn ), 1 ) "
    "RETURN SELECTCOLUMNS ( FILTER ( CALENDAR ( ini, mx ), DAY ( [Date] ) = 1 ), "
    "\"Início do mês\", [Date], "
    "\"Chave\", FORMAT ( [Date], \"yyyy-MM\" ), "
    "\"Mês\", VAR t = FORMAT ( [Date], \"mmmm\" ) RETURN UPPER ( LEFT ( t, 1 ) ) & MID ( t, 2, 30 ) & \" de \" & YEAR ( [Date] ), "
    "\"Fim do mês\", EOMONTH ( [Date], 0 ) )"
)

EVENTOMES_EXPR = (
    "SELECTCOLUMNS ( FILTER ( CROSSJOIN ( "
    "SELECTCOLUMNS ( Eventos, \"EvID\", Eventos[ID], \"EvIni\", Eventos[Data início], \"EvFim\", Eventos[Data fim] ), "
    "SELECTCOLUMNS ( Meses, \"MChave\", Meses[Chave], \"MIni\", Meses[Início do mês], \"MFim\", Meses[Fim do mês] ) ), "
    "[MIni] <= [EvFim] && [MFim] >= [EvIni] ), "
    "\"ID\", [EvID], \"Chave\", [MChave] )"
)

PERIODO_EXPR = 'DATATABLE ( "Meses", INTEGER, "Período", STRING, { { 1, "1 mês" }, { 2, "2 meses" }, { 3, "3 meses" } } )'

E_EVENTO = esc("Eventos[Evento]")
E_STATUS = esc("Eventos[Status]")
E_TRANS = esc('COALESCE ( Eventos[Transmissão LF], "—" )')
E_OBS = esc('COALESCE ( Eventos[Observação], "—" )')

HTML_DESTAQUES = f"""
VAR hoje = TODAY ()
VAR base = FILTER ( ALL ( Eventos ), Eventos[StatusCat] <> "Cancelado" && Eventos[Data fim] >= hoje )
VAR prox = MINX ( FILTER ( base, Eventos[Data início] > hoje ), Eventos[Data início] )
VAR sel = FILTER ( base, Eventos[Data início] <= hoje || Eventos[Data início] = prox )
RETURN
    IF (
        ISEMPTY ( sel ),
        "<div class='vazio'>Nenhum evento futuro na planilha.</div>",
        "<div class='next'>"
            & CONCATENATEX (
                sel,
                VAR ini = Eventos[Data início]
                VAR fim = Eventos[Data fim]
                VAR n = INT ( ini - hoje )
                VAR aviso =
                    IF ( ini = hoje && fim = hoje, "Acontece hoje",
                        IF ( ini <= hoje,
                            "Em andamento · " & IF ( fim = hoje, "termina hoje", "até " & FORMAT ( fim, "dd/MM" ) ),
                            "Próximo · " & IF ( n = 1, "amanhã", "faltam " & n & " dias" ) ) )
                RETURN
                    "<div class='hero' style='border-left-color:" & Eventos[Cor] & "'>"
                        & "<div class='top'><span class='eb'>" & aviso & "</span><span class='st'><i style='background:" & Eventos[Cor] & "'></i>" & {E_STATUS} & "</span></div>"
                        & "<div class='nm'>" & {E_EVENTO} & "</div>"
                        & "<div class='meta'><div><b>" & IF ( ini = fim, "Data", "Datas" ) & "</b>" & Eventos[Datas] & "</div>"
                        & "<div><b>Dia da semana</b>" & Eventos[Dia da semana] & "</div>"
                        & "<div><b>Transmissão LF</b>" & {E_TRANS} & "</div></div>"
                        & "<div class='obs'><b>Observação</b>" & {E_OBS} & "</div>"
                        & "</div>",
                "",
                Eventos[Data início], ASC,
                Eventos[Evento], ASC
            )
            & "</div>"
    )
""".strip()

HTML_CARDS = f"""
VAR hoje = TODAY ()
RETURN
    IF (
        ISEMPTY ( Eventos ),
        "<div class='vazio'>Nenhum evento com os filtros selecionados.</div>",
        "<div class='cards'>"
            & CONCATENATEX (
                Eventos,
                "<div class='card" & IF ( Eventos[Data fim] < hoje, " past" ) & "' style='border-top-color:" & Eventos[Cor] & "'>"
                    & IF ( NOT ISBLANK ( Eventos[Status] ), "<span class='badge'><i class='dot' style='background:" & Eventos[Cor] & "'></i>" & {E_STATUS} & "</span>" )
                    & "<h3>" & {E_EVENTO} & "</h3>"
                    & "<div class='when'><strong>" & Eventos[Datas] & "</strong> · " & Eventos[Dias] & IF ( Eventos[Dias] = 1, " dia", " dias" ) & "</div>"
                    & "<div class='kv'>"
                    & IF ( NOT ISBLANK ( Eventos[Transmissão LF] ), "<b>Transmissão LF</b><span>" & {esc("Eventos[Transmissão LF]")} & "</span>" )
                    & IF ( NOT ISBLANK ( Eventos[Observação] ), "<b>Observação</b><span>" & {esc("Eventos[Observação]")} & "</span>" )
                    & "</div></div>",
                "",
                Eventos[Data início], ASC,
                Eventos[Evento], ASC
            )
            & "</div>"
    )
""".strip()

HTML_LISTA = f"""
VAR hoje = TODAY ()
RETURN
    IF (
        ISEMPTY ( Eventos ),
        "<div class='vazio'>Nenhum evento com os filtros selecionados.</div>",
        "<table><thead><tr><th>Data</th><th>Evento</th><th>Status</th><th>Transmissão LF</th><th>Observação</th></tr></thead><tbody>"
            & CONCATENATEX (
                VALUES ( Eventos[Mês de início] ),
                VAR m = Eventos[Mês de início]
                VAR linhas = FILTER ( Eventos, Eventos[Mês de início] = m )
                VAR d = MINX ( linhas, Eventos[Data início] )
                VAR nome = FORMAT ( d, "mmmm" )
                VAR qtd = COUNTROWS ( linhas )
                RETURN
                    "<tr class='grp'><td colspan='5'>" & UPPER ( LEFT ( nome, 1 ) ) & MID ( nome, 2, 30 ) & " de " & YEAR ( d )
                        & "<span>" & qtd & IF ( qtd = 1, " evento", " eventos" ) & "</span></td></tr>"
                        & CONCATENATEX (
                            linhas,
                            "<tr" & IF ( Eventos[Data fim] < hoje, " class='past'" ) & ">"
                                & "<td class='dt'>" & Eventos[Datas] & "</td>"
                                & "<td class='nm2'>" & {E_EVENTO} & "</td>"
                                & "<td><i class='dot' style='background:" & Eventos[Cor] & "'></i>" & {E_STATUS} & "</td>"
                                & "<td>" & {esc("Eventos[Transmissão LF]")} & "</td>"
                                & "<td>" & {esc("Eventos[Observação]")} & "</td></tr>",
                            "",
                            Eventos[Data início], ASC,
                            Eventos[Evento], ASC
                        ),
                "",
                Eventos[Mês de início], ASC
            )
            & "</tbody></table>"
    )
""".strip()

HTML_LEGENDA = """
VAR c = CALCULATE ( COUNTROWS ( Eventos ), Eventos[StatusCat] = "Confirmado" ) + 0
VAR a = CALCULATE ( COUNTROWS ( Eventos ), Eventos[StatusCat] = "A confirmar" ) + 0
VAR x = CALCULATE ( COUNTROWS ( Eventos ), Eventos[StatusCat] = "Cancelado" ) + 0
VAR o = CALCULATE ( COUNTROWS ( Eventos ), Eventos[StatusCat] = "Outros" ) + 0
RETURN
    "<div class='leg'>"
        & "<div><i class='dot' style='background:#7AE582'></i>Confirmado<span>" & c & "</span></div>"
        & "<div><i class='dot' style='background:#F2C75C'></i>A confirmar<span>" & a & "</span></div>"
        & "<div><i class='dot' style='background:#F2857E'></i>Cancelado<span>" & x & "</span></div>"
        & IF ( o > 0, "<div><i class='dot' style='background:#B9A995'></i>Outros<span>" & o & "</span></div>" )
        & "</div>"
""".strip()

HTML_CABECALHO = '"<div class=\'hd\'><h1>Eventos de curta duração</h1><p>Eventos com oportunidades de venda de campanhas no digital</p></div>"'

TITULO = """
VAR n = COUNTROWS ( Eventos ) + 0
RETURN "<div class='hd'><h1 style='font-size:18px'>{titulo}</h1><p>" & {sub} & "</p></div>"
""".strip()

SUB_LINHA = (
    'IF ( [Mes Selecionado] <> "", '
    'SELECTEDVALUE ( Meses[Mês] ) & " (filtro de mês) · " & n & IF ( n = 1, " evento", " eventos" ), '
    'VAR fimJ = EDATE ( TODAY (), [Periodo Meses] ) - 1 '
    'VAR q = COUNTROWS ( FILTER ( Eventos, Eventos[Data início] <= fimJ && Eventos[Data fim] >= TODAY () ) ) + 0 '
    'RETURN "De " & FORMAT ( TODAY (), "dd/MM/yyyy" ) & " a " & FORMAT ( fimJ, "dd/MM/yyyy" ) & " · " & q & IF ( q = 1, " evento", " eventos" ) )'
)
SUB_CAL = 'IF ( [Mes Selecionado] <> "", SELECTEDVALUE ( Meses[Mês] ), "Mês atual · use o filtro Mês para trocar" )'
SUB_QTD = 'n & IF ( n = 1, " evento", " eventos" )'

medidas = [
    measure("Periodo Meses", "SELECTEDVALUE ( Periodo[Meses], 1 )", "0"),
    measure("Mes Selecionado",
            'IF ( ISFILTERED ( Meses[Mês] ) || ISFILTERED ( Meses[Chave] ), IF ( HASONEVALUE ( Meses[Chave] ), VALUES ( Meses[Chave] ), "" ), "" )'),
    measure("Qtd Eventos", "COUNTROWS ( Eventos ) + 0", "0"),
    measure("HTML Destaques", HTML_DESTAQUES),
    measure("HTML Cards", HTML_CARDS),
    measure("HTML Lista", HTML_LISTA),
    measure("HTML Legenda", HTML_LEGENDA),
    measure("HTML Cabecalho", HTML_CABECALHO),
    measure("HTML Titulo Linha", TITULO.replace("{titulo}", "Linha do tempo").replace("{sub}", SUB_LINHA)),
    measure("HTML Titulo Calendario", TITULO.replace("{titulo}", "Calendário").replace("{sub}", SUB_CAL)),
    measure("HTML Titulo Cards", TITULO.replace("{titulo}", "Cards").replace("{sub}", SUB_QTD)),
    measure("HTML Titulo Lista", TITULO.replace("{titulo}", "Lista").replace("{sub}", SUB_QTD + ' & " · agrupados por mês de início"')),
]

model = {
    "compatibilityLevel": 1567,
    "model": {
        "culture": "pt-BR",
        "sourceQueryCulture": "pt-BR",
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "expressions": [
            {
                "name": "CaminhoExcel",
                "kind": "m",
                "expression": "\"COLE_AQUI_O_CAMINHO_DO_EXCEL\" meta [IsParameterQuery=true, Type=\"Text\", IsParameterQueryRequired=true]",
                "annotations": [{"name": "PBI_ResultType", "value": "Text"}],
            }
        ],
        "tables": [
            {
                "name": "Eventos",
                "columns": eventos_cols,
                "partitions": [{"name": "Eventos", "mode": "import",
                                "source": {"type": "m", "expression": M_EVENTOS}}],
            },
            {
                "name": "Meses",
                "columns": [
                    ctab_col("Início do mês", "dateTime", DATA, hidden=True),
                    ctab_col("Chave", "string", hidden=True),
                    ctab_col("Mês", "string", sort_by="Chave"),
                    ctab_col("Fim do mês", "dateTime", DATA, hidden=True),
                ],
                "partitions": [{"name": "Meses", "mode": "import",
                                "source": {"type": "calculated", "expression": MESES_EXPR}}],
            },
            {
                "name": "EventoMes",
                "isHidden": True,
                "columns": [ctab_col("ID", "int64", "0"), ctab_col("Chave", "string")],
                "partitions": [{"name": "EventoMes", "mode": "import",
                                "source": {"type": "calculated", "expression": EVENTOMES_EXPR}}],
            },
            {
                "name": "Periodo",
                "columns": [ctab_col("Meses", "int64", "0", hidden=True),
                            ctab_col("Período", "string", sort_by="Meses")],
                "partitions": [{"name": "Periodo", "mode": "import",
                                "source": {"type": "calculated", "expression": PERIODO_EXPR}}],
            },
            {
                "name": "Medidas",
                "columns": [ctab_col("Coluna", "string", hidden=True)],
                "partitions": [{"name": "Medidas", "mode": "import",
                                "source": {"type": "calculated", "expression": 'ROW ( "Coluna", "" )'}}],
                "measures": medidas,
            },
        ],
        "relationships": [
            {"name": "EventoMes_Eventos", "fromTable": "EventoMes", "fromColumn": "ID",
             "toTable": "Eventos", "toColumn": "ID", "crossFilteringBehavior": "bothDirections"},
            {"name": "EventoMes_Meses", "fromTable": "EventoMes", "fromColumn": "Chave",
             "toTable": "Meses", "toColumn": "Chave"},
        ],
        "annotations": [
            {"name": "PBI_QueryOrder", "value": "[\"CaminhoExcel\",\"Eventos\"]"},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
            {"name": "PBIDesktopVersion", "value": "2.130"},
        ],
    },
}

# --------------------------------------------------------------------------
# Relatório (PBIR)
# --------------------------------------------------------------------------

def lit(s):
    return {"expr": {"Literal": {"Value": "'" + s.replace("'", "''") + "'"}}}


def lbool(b):
    return {"expr": {"Literal": {"Value": "true" if b else "false"}}}


def lnum(n):
    return {"expr": {"Literal": {"Value": f"{n}D"}}}


def color(hexv):
    return {"solid": {"color": lit(hexv)}}


def col(entity, prop):
    return {"field": {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}},
            "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}


def meas(entity, prop):
    return {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}},
            "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}


NO_TITLE = {
    "title": [{"properties": {"show": lbool(False)}}],
    "background": [{"properties": {"show": lbool(False)}}],
    "border": [{"properties": {"show": lbool(False)}}],
    "dropShadow": [{"properties": {"show": lbool(False)}}],
}


def container(name, x, y, w, h, z, visual):
    return {"$schema": S_VISUAL, "name": name,
            "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z},
            "visual": visual}


CSS = (SRC / "estilo-html.css").read_text(encoding="utf-8")


def html_visual(measure_name, bg=None):
    v = {"visualType": HTML,
         "query": {"queryState": {"content": {"projections": [meas("Medidas", measure_name)]}}},
         "objects": {
             "stylesheet": [{"properties": {"stylesheet": lit(CSS)}}],
             "contentFormatting": [{"properties": {"fontColour": color("#FFFAF3"), "fontSize": lnum(11)}}],
         },
         "visualContainerObjects": json.loads(json.dumps(NO_TITLE))}
    if bg:
        v["visualContainerObjects"]["background"] = [{"properties": {"show": lbool(True), "color": color(bg), "transparency": lnum(0)}}]
    return v


def slicer_visual(entity, prop, group, single=False):
    objs = {"data": [{"properties": {"mode": lit("Dropdown")}}],
            "header": [{"properties": {"show": lbool(True)}}]}
    if single:
        objs["selection"] = [{"properties": {"singleSelect": lbool(True)}}]
    return {"visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [col(entity, prop)]}}},
            "objects": objs,
            "visualContainerObjects": {"title": [{"properties": {"show": lbool(False)}}],
                                       "background": [{"properties": {"show": lbool(False)}}]},
            "syncGroup": {"groupName": group, "fieldChanges": True, "filterChanges": True}}


def deneb_visual(spec_file, fields):
    spec = json.loads((ROOT / "deneb" / spec_file).read_text(encoding="utf-8"))
    spec.pop("$schema", None)
    projections = [col("Eventos", f) if f not in ("Periodo Meses", "Mes Selecionado") else meas("Medidas", f)
                   for f in fields]
    return {"visualType": DENEB,
            "query": {"queryState": {"dataset": {"projections": projections}}},
            "objects": {"vega": [{"properties": {
                "jsonSpec": lit(json.dumps(spec, ensure_ascii=False, separators=(",", ":"))),
                "jsonConfig": lit("{}"),
                "provider": lit("vegaLite"),
                "enableTooltips": lbool(True),
                "isNewDialogOpen": lbool(False)}}]},
            "visualContainerObjects": {"title": [{"properties": {"show": lbool(False)}}],
                                       "background": [{"properties": {"show": lbool(True), "color": color("#152739"), "transparency": lnum(0)}}],
                                       "border": [{"properties": {"show": lbool(True), "color": color("#2B4561"), "radius": lnum(10)}}]}}


W, H = 1280, 900
MAIN_X, MAIN_W = 272, 992

PAGES = [
    ("linhaDoTempo", "Linha do tempo", "HTML Titulo Linha"),
    ("calendario", "Calendário", "HTML Titulo Calendario"),
    ("cards", "Cards", "HTML Titulo Cards"),
    ("lista", "Lista", "HTML Titulo Lista"),
]


def page_visuals(page_id):
    v = []
    z = 0

    def add(name, x, y, w, h, visual):
        nonlocal z
        z += 1000
        v.append(container(f"{page_id}_{name}", x, y, w, h, z, visual))

    add("logo", 16, 16, 48, 48, {
        "visualType": "image",
        "objects": {"general": [{"properties": {"imageUrl": {"expr": {"ResourcePackageItem": {
            "PackageName": "RegisteredResources", "PackageType": 1, "ItemName": "logo.png"}}}}}],
            "imageScaling": [{"properties": {"imageScalingType": lit("Fit")}}]},
        "visualContainerObjects": json.loads(json.dumps(NO_TITLE))})
    add("cabecalho", 72, 12, 560, 60, html_visual("HTML Cabecalho"))
    add("navegacao", 652, 20, 612, 44, {
        "visualType": "pageNavigator",
        "visualContainerObjects": json.loads(json.dumps(NO_TITLE))})
    add("fMes", 16, 92, 240, 60, slicer_visual("Meses", "Mês", "Mes"))
    add("fEvento", 16, 158, 240, 60, slicer_visual("Eventos", "Evento", "Evento"))
    add("fStatus", 16, 224, 240, 60, slicer_visual("Eventos", "Status", "Status"))
    add("fTransmissao", 16, 290, 240, 60, slicer_visual("Eventos", "Transmissão LF", "Transmissao"))
    y_leg = 360
    if page_id == "linhaDoTempo":
        add("fPeriodo", 16, 356, 240, 60, slicer_visual("Periodo", "Período", "Periodo", single=True))
        y_leg = 426
    add("legenda", 16, y_leg, 240, 140, html_visual("HTML Legenda"))
    add("destaques", MAIN_X, 84, MAIN_W, 170, html_visual("HTML Destaques"))
    titulo = dict(PAGES_BY_ID[page_id])
    add("titulo", MAIN_X, 262, MAIN_W, 54, html_visual(titulo["measure"]))
    if page_id == "linhaDoTempo":
        add("grafico", MAIN_X, 322, MAIN_W, 560, deneb_visual("linha-do-tempo.vl.json",
            ["Evento", "InicioISO", "FimISO", "Status", "Cor", "Lane", "ID", "Periodo Meses", "Mes Selecionado"]))
    elif page_id == "calendario":
        add("grafico", MAIN_X, 322, MAIN_W, 560, deneb_visual("calendario.vl.json",
            ["Evento", "InicioISO", "FimISO", "Status", "Cor", "ID", "Mes Selecionado"]))
    elif page_id == "cards":
        add("grafico", MAIN_X, 322, MAIN_W, 560, html_visual("HTML Cards"))
    else:
        add("grafico", MAIN_X, 322, MAIN_W, 560, html_visual("HTML Lista", bg="#152739"))
    return v


PAGES_BY_ID = {p[0]: {"name": p[1], "measure": p[2]} for p in PAGES}


def write(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(obj, (dict, list)):
        path.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    else:
        path.write_text(obj, encoding="utf-8")


def build():
    if OUT.exists():
        shutil.rmtree(OUT)
    sm = OUT / f"{NOME}.SemanticModel"
    rp = OUT / f"{NOME}.Report"

    write(OUT / f"{NOME}.pbip", {"$schema": S_PBIP, "version": "1.0",
                                 "artifacts": [{"report": {"path": f"{NOME}.Report"}}],
                                 "settings": {"enableAutoRecovery": True}})
    write(OUT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")

    write(sm / "definition.pbism", {"$schema": S_PBISM, "version": "1.0", "settings": {}})
    write(sm / "model.bim", model)

    write(rp / "definition.pbir", {"$schema": S_PBIR, "version": "4.0",
                                   "datasetReference": {"byPath": {"path": f"../{NOME}.SemanticModel"}}})
    d = rp / "definition"
    write(d / "version.json", {"$schema": S_VERSION, "version": "2.0.0"})
    write(d / "report.json", {
        "$schema": S_REPORT,
        "themeCollection": {"customTheme": {"name": "tema-azulterra.json", "reportVersionAtImport": "5.59",
                                            "type": "RegisteredResources"}},
        "layoutOptimization": "None",
        "objects": {"outspacePane": [{"properties": {"expanded": lbool(False), "visible": lbool(True)}}]},
        "publicCustomVisuals": [DENEB, HTML],
        "resourcePackages": [{"name": "RegisteredResources", "type": "RegisteredResources", "items": [
            {"name": "tema-azulterra.json", "path": "tema-azulterra.json", "type": "CustomTheme"},
            {"name": "logo.png", "path": "logo.png", "type": "Image"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True},
    })
    write(d / "pages" / "pages.json", {"$schema": S_PAGES, "pageOrder": [p[0] for p in PAGES],
                                       "activePageName": PAGES[0][0]})
    for pid, display, _ in PAGES:
        write(d / "pages" / pid / "page.json", {
            "$schema": S_PAGE, "name": pid, "displayName": display, "displayOption": "FitToWidth",
            "height": H, "width": W,
            "objects": {"background": [{"properties": {"color": color("#0E1B29"), "transparency": lnum(0)}}],
                        "outspace": [{"properties": {"color": color("#0E1B29"), "transparency": lnum(0)}}]},
        })
        for vis in page_visuals(pid):
            write(d / "pages" / pid / "visuals" / vis["name"] / "visual.json", vis)

    res = rp / "StaticResources" / "RegisteredResources"
    res.mkdir(parents=True, exist_ok=True)
    shutil.copy(SRC / "tema-azulterra.json", res / "tema-azulterra.json")
    shutil.copy(ROOT.parent / "eventos" / "logo.png", res / "logo.png")
    print("Projeto gerado em", OUT)


if __name__ == "__main__":
    build()
