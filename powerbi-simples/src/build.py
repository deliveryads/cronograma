"""Gera o projeto Power BI simples "Eventos de curta duração" (Cards + Lista).

Só visuais nativos (cartão de várias linhas, tabela e segmentações) e nenhuma
medida DAX: todas as colunas são calculadas no Power Query.

Uso: python3 powerbi-simples/src/build.py
Saída: powerbi-simples/projeto/
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

# --------------------------------------------------------------------------
# Power Query: lê o Excel (SharePoint/OneDrive ou arquivo local) e calcula
# todas as colunas de exibição. Nenhuma coluna calculada em DAX.
# --------------------------------------------------------------------------
M_EVENTOS = r'''let
    // CaminhoExcel: Excel > Arquivo > Informações > Copiar caminho (ou caminho local, ex.: C:\Eventos\Eventos.xlsx)
    Texto = Text.Trim(CaminhoExcel),
    Caminho = Text.BeforeDelimiter(Texto, "?"),
    EhWeb = Text.StartsWith(Text.Lower(Caminho), "http"),
    Partes = if EhWeb then Uri.Parts(Texto) else null,
    Segmentos = if EhWeb then List.Select(Text.Split(Partes[Path], "/"), each _ <> "") else {},
    PosSite = List.PositionOfAny(List.Transform(Segmentos, Text.Lower), {"sites", "teams", "personal"}),
    Site = if EhWeb then Partes[Scheme] & "://" & Partes[Host] & (if PosSite >= 0 then "/" & Segmentos{PosSite} & "/" & Segmentos{PosSite + 1} else "") else null,
    NomeArquivo = if not EhWeb then null else if Record.HasFields(Partes[Query], "file") then Record.Field(Partes[Query], "file") else List.Last(Segmentos),
    NomeDecod = if NomeArquivo = null then "" else Text.Lower(Text.Replace(NomeArquivo, "%20", " ")),
    BaseNome = Text.BeforeDelimiter(NomeDecod, ".", {0, RelativePosition.FromEnd}),
    // 1ª tentativa: baixar direto pelo caminho do arquivo (API do SharePoint)
    CaminhoServidor = if EhWeb then Text.Replace(Partes[Path], "%20", " ") else null,
    Direto = if EhWeb then try Binary.Buffer(Web.Contents(Site, [RelativePath = "_api/web/GetFileByServerRelativeUrl('" & Text.Replace(CaminhoServidor, "'", "''") & "')/$value"])) otherwise null else null,
    // 2ª tentativa: procurar o arquivo pelo nome entre os arquivos do site
    Arquivos = if EhWeb and Direto = null then Table.Buffer(Table.SelectRows(SharePoint.Files(Site, [ApiVersion = 15]), each Text.EndsWith(Text.Lower([Name] ?? ""), ".xlsx"))) else null,
    Exatos = if Arquivos = null then null else Table.SelectRows(Arquivos, each Text.Lower([Name] ?? "") = NomeDecod),
    Parecidos = if Arquivos = null then null else Table.Sort(Table.SelectRows(Arquivos, each Text.Contains(Text.Lower([Name] ?? ""), BaseNome)), {{"Date modified", Order.Descending}}),
    Escolhido = if Arquivos = null then null else if not Table.IsEmpty(Exatos) then Exatos else Parecidos,
    Disponiveis = if Arquivos = null then "" else Text.Combine(List.FirstN(List.Transform(Table.ToRecords(Arquivos), each [Name] & " (" & [Folder Path] & ")"), 15), "; "),
    Conteudo = if not EhWeb then File.Contents(Caminho)
        else if Direto <> null then Direto
        else if not Table.IsEmpty(Escolhido) then Escolhido{0}[Content]
        else error Error.Record("Arquivo não encontrado", "O arquivo '" & NomeArquivo & "' não foi encontrado no site " & Site & ". Arquivos .xlsx encontrados nesse site: " & (if Disponiveis = "" then "nenhum (verifique se você tem acesso ao site)" else Disponiveis)),
    Pasta = Excel.Workbook(Conteudo, null, true),
    // Tabela do Excel chamada Eventos; se não houver, primeira aba com o cabeçalho nas 10 primeiras linhas
    TemTabela = Table.MatchesAnyRows(Pasta, each [Kind] = "Table" and [Item] = "Eventos"),
    Planilha = Table.SelectRows(Pasta, each [Kind] = "Sheet"){0}[Data],
    PosCabecalho = List.PositionOf(List.Transform(Table.ToRows(Table.FirstN(Planilha, 10)), (linha) => List.Contains(List.Transform(linha, each if _ = null then "" else Text.Lower(Text.Trim(Text.From(_)))), "evento")), true),
    DaPlanilha = Table.PromoteHeaders(Table.Skip(Planilha, if PosCabecalho < 0 then 0 else PosCabecalho), [PromoteAllScalars = true]),
    Bruto = if TemTabela then Pasta{[Item = "Eventos", Kind = "Table"]}[Data] else DaPlanilha,
    Colunas = Table.SelectColumns(Bruto, {"Data início", "Data fim", "Evento", "Status", "Transmissão LF", "Observação"}, MissingField.UseNull),
    Tipos = Table.TransformColumnTypes(Colunas, {{"Data início", type date}, {"Data fim", type date}, {"Evento", type text}, {"Status", type text}, {"Transmissão LF", type text}, {"Observação", type text}}, "pt-BR"),
    // Textos: vazio vira null; nada de função de texto sobre null
    Limpo = Table.TransformColumns(Tipos, {
        {"Evento", each let t = Text.Trim(Text.From(_) ?? "") in if t = "" then null else t, type text},
        {"Status", each let t = Text.Trim(Text.From(_) ?? "") in if t = "" then null else t, type text},
        {"Transmissão LF", each let t = Text.Trim(Text.From(_) ?? "") in if t = "" then null else t, type text},
        {"Observação", each let t = Text.Trim(Text.From(_) ?? "") in if t = "" then null else t, type text}}),
    Validos = Table.SelectRows(Limpo, each [Evento] <> null and [#"Data início"] <> null),
    // Data fim vazia = evento de um dia
    ComFim = Table.AddColumn(Validos, "Fim", each if [#"Data fim"] = null then [#"Data início"] else [#"Data fim"], type date),
    SemFimOriginal = Table.RemoveColumns(ComFim, {"Data fim"}),
    Renomeado = Table.RenameColumns(SemFimOriginal, {{"Fim", "Data fim"}}),
    DatasOk = Table.SelectRows(Renomeado, each [#"Data fim"] >= [#"Data início"]),
    Hoje = Date.From(DateTime.LocalNow()),
    // Colunas de exibição
    ComDias = Table.AddColumn(DatasOk, "Dias", each Duration.Days([#"Data fim"] - [#"Data início"]) + 1, Int64.Type),
    DataTxt = (d as date, comAno as logical) as text =>
        Text.PadStart(Text.From(Date.Day(d)), 2, "0") & "/" & Text.PadStart(Text.From(Date.Month(d)), 2, "0") & (if comAno then "/" & Text.From(Date.Year(d)) else ""),
    ComPeriodo = Table.AddColumn(ComDias, "Período", each try
        (if [#"Data início"] = [#"Data fim"] then DataTxt([#"Data início"], true)
        else DataTxt([#"Data início"], false) & " a " & DataTxt([#"Data fim"], true)) otherwise null, type text),
    NomesDias = {"Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"},
    DiaTxt = (d as date) as text => NomesDias{Date.DayOfWeek(d, Day.Sunday)},
    ComDiaSemana = Table.AddColumn(ComPeriodo, "Dia da semana", each try
        (if [#"Data início"] = [#"Data fim"] then DiaTxt([#"Data início"])
        else DiaTxt([#"Data início"]) & " a " & Text.Lower(DiaTxt([#"Data fim"]))) otherwise null, type text),
    NomesMeses = {"Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"},
    ComMes = Table.AddColumn(ComDiaSemana, "Mês", each try
        (NomesMeses{Date.Month([#"Data início"]) - 1} & " de " & Text.From(Date.Year([#"Data início"]))) otherwise null, type text),
    ComOrdemMes = Table.AddColumn(ComMes, "Ordem do mês", each Date.Year([#"Data início"]) * 100 + Date.Month([#"Data início"]), Int64.Type),
    ComLegenda = Table.AddColumn(ComOrdemMes, "Status (legenda)", each
        let s = Text.Lower(Text.From([Status]) ?? "") in
            if s = "confirmado" then "🟢 Confirmado"
            else if s = "a confirmar" then "🟡 A confirmar"
            else if s = "cancelado" then "🔴 Cancelado"
            else "⚪ " & (Text.From([Status]) ?? "Sem status"), type text),
    ComSituacao = Table.AddColumn(ComLegenda, "Situação", each
        if [#"Data fim"] < Hoje then "Realizado"
        else if [#"Data início"] <= Hoje then "Em andamento"
        else "Próximo", type text),
    Ordenado = Table.Sort(ComSituacao, {{"Data início", Order.Ascending}, {"Evento", Order.Ascending}}),
    Final = Table.ReorderColumns(Ordenado, {"Evento", "Período", "Dia da semana", "Status (legenda)", "Transmissão LF", "Observação", "Situação", "Mês", "Data início", "Data fim", "Dias", "Status", "Ordem do mês"})
in
    Final'''


def col(name, dtype, fmt=None, hidden=False, sort_by=None):
    c = {"name": name, "dataType": dtype, "sourceColumn": name, "summarizeBy": "none"}
    if fmt:
        c["formatString"] = fmt
    if hidden:
        c["isHidden"] = True
    if sort_by:
        c["sortByColumn"] = sort_by
    return c


model = {
    "compatibilityLevel": 1567,
    "model": {
        "culture": "pt-BR",
        "sourceQueryCulture": "pt-BR",
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "expressions": [{
            "name": "CaminhoExcel",
            "kind": "m",
            "expression": "\"COLE_AQUI_O_CAMINHO_DO_EXCEL\" meta [IsParameterQuery=true, Type=\"Text\", IsParameterQueryRequired=true]",
            "annotations": [{"name": "PBI_ResultType", "value": "Text"}],
        }],
        "tables": [{
            "name": "Eventos",
            "columns": [
                col("Evento", "string"),
                col("Período", "string"),
                col("Dia da semana", "string"),
                col("Status (legenda)", "string"),
                col("Transmissão LF", "string"),
                col("Observação", "string"),
                col("Situação", "string"),
                col("Mês", "string", sort_by="Ordem do mês"),
                col("Data início", "dateTime", "dd/MM/yyyy"),
                col("Data fim", "dateTime", "dd/MM/yyyy"),
                col("Dias", "int64", "0"),
                col("Status", "string"),
                col("Ordem do mês", "int64", "0", hidden=True),
            ],
            "partitions": [{"name": "Eventos", "mode": "import",
                            "source": {"type": "m", "expression": M_EVENTOS.split("\n")}}],
        }],
        "annotations": [
            {"name": "PBI_QueryOrder", "value": "[\"CaminhoExcel\",\"Eventos\"]"},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
        ],
    },
}

# --------------------------------------------------------------------------
# Relatório: 2 páginas (Cards e Lista), só visuais nativos
# --------------------------------------------------------------------------


def lit(s):
    return {"expr": {"Literal": {"Value": "'" + s.replace("'", "''") + "'"}}}


def lbool(b):
    return {"expr": {"Literal": {"Value": "true" if b else "false"}}}


def lnum(n):
    return {"expr": {"Literal": {"Value": f"{n}D"}}}


def color(hexv):
    return {"solid": {"color": lit(hexv)}}


def field(prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": "Eventos"}}, "Property": prop}}


def proj(prop):
    return {"field": field(prop), "queryRef": f"Eventos.{prop}", "nativeQueryRef": prop}


def sort_by_start():
    return {"sort": [{"field": field("Data início"), "direction": "Ascending"}], "isDefaultSort": False}


def title_objs(text):
    return {
        "title": [{"properties": {"show": lbool(True), "text": lit(text), "fontColor": color("#FFFAF3"),
                                  "fontSize": lnum(14), "bold": lbool(True)}}],
        "background": [{"properties": {"show": lbool(True), "color": color("#152739"), "transparency": lnum(0)}}],
        "border": [{"properties": {"show": lbool(True), "color": color("#2B4561"), "radius": lnum(8)}}],
    }


def container(name, x, y, w, h, z, visual):
    return {"$schema": S_VISUAL, "name": name,
            "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z},
            "visual": visual}


def slicer(prop, group):
    return {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [proj(prop)]}}},
        "objects": {"data": [{"properties": {"mode": lit("Dropdown")}}]},
        "visualContainerObjects": {"title": [{"properties": {"show": lbool(False)}}],
                                   "background": [{"properties": {"show": lbool(False)}}]},
        "syncGroup": {"groupName": group, "fieldChanges": True, "filterChanges": True},
    }


CAMPOS_CARD = ["Evento", "Período", "Dia da semana", "Status (legenda)", "Transmissão LF", "Observação", "Data início"]
CAMPOS_LISTA = ["Data início", "Período", "Evento", "Dia da semana", "Status (legenda)", "Transmissão LF", "Observação", "Situação"]


def cards_visual():
    return {
        "visualType": "multiRowCard",
        "query": {"queryState": {"Values": {"projections": [proj(p) for p in CAMPOS_CARD]}},
                  "sortDefinition": sort_by_start()},
        "objects": {
            "dataLabels": [{"properties": {"color": color("#FFFAF3"), "fontSize": lnum(13)}}],
            "categoryLabels": [{"properties": {"show": lbool(True), "color": color("#CBB9A3"), "fontSize": lnum(9)}}],
            "card": [{"properties": {"barShow": lbool(True), "barColor": color("#F0A77C"),
                                     "outlineColor": color("#2B4561")}}],
        },
        "visualContainerObjects": title_objs("Eventos (cards)"),
    }


def lista_visual():
    return {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [proj(p) for p in CAMPOS_LISTA]}},
                  "sortDefinition": sort_by_start()},
        "objects": {
            "columnHeaders": [{"properties": {"fontColor": color("#FFFAF3"), "backColor": color("#1C3147"),
                                              "fontSize": lnum(11), "bold": lbool(True)}}],
            "values": [{"properties": {"fontColorPrimary": color("#FFFAF3"), "backColorPrimary": color("#152739"),
                                       "fontColorSecondary": color("#FFFAF3"), "backColorSecondary": color("#1A2E40"),
                                       "fontSize": lnum(11)}}],
            "grid": [{"properties": {"gridHorizontal": lbool(True), "gridHorizontalColor": color("#2B4561"),
                                     "rowPadding": lnum(6)}}],
        },
        "visualContainerObjects": title_objs("Eventos (lista)"),
    }


W, H = 1280, 720
FILTROS = [("Mês", "Mes"), ("Situação", "Situacao"), ("Status (legenda)", "Status"),
           ("Transmissão LF", "Transmissao"), ("Evento", "Evento")]
PAGES = [("cards", "Cards"), ("lista", "Lista")]


def page_visuals(pid):
    v = []
    z = 0
    for i, (prop, group) in enumerate(FILTROS):
        z += 1000
        v.append(container(f"{pid}_filtro{i + 1}", 16, 16 + i * 66, 220, 60, z, slicer(prop, group)))
    z += 1000
    main = cards_visual() if pid == "cards" else lista_visual()
    v.append(container(f"{pid}_principal", 252, 16, 1012, 688, z, main))
    return v


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
        "resourcePackages": [{"name": "RegisteredResources", "type": "RegisteredResources", "items": [
            {"name": "tema-azulterra.json", "path": "tema-azulterra.json", "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True},
    })
    write(d / "pages" / "pages.json", {"$schema": S_PAGES, "pageOrder": [p[0] for p in PAGES],
                                       "activePageName": PAGES[0][0]})
    for pid, display in PAGES:
        write(d / "pages" / pid / "page.json", {
            "$schema": S_PAGE, "name": pid, "displayName": display, "displayOption": "FitToPage",
            "height": H, "width": W,
            "objects": {"background": [{"properties": {"color": color("#0E1B29"), "transparency": lnum(0)}}],
                        "outspace": [{"properties": {"color": color("#0E1B29"), "transparency": lnum(0)}}]},
        })
        for vis in page_visuals(pid):
            write(d / "pages" / pid / "visuals" / vis["name"] / "visual.json", vis)
    res = rp / "StaticResources" / "RegisteredResources"
    res.mkdir(parents=True, exist_ok=True)
    shutil.copy(SRC / "tema-azulterra.json", res / "tema-azulterra.json")
    print("Projeto gerado em", OUT)


if __name__ == "__main__":
    build()
