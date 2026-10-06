/**
 * Eventos de curta duração — versão Google Apps Script (acesso restrito).
 *
 * Implantação como App da Web:
 *   Executar como: "Usuário que acessa o app da Web"
 *   Quem pode acessar: "Qualquer pessoa com uma Conta do Google"
 * Assim a planilha é lida com a conta de quem abre a página: só vê os dados
 * quem tiver a planilha compartilhada com o próprio e-mail.
 *
 * Passo a passo: apps-script/README.md
 */

var SHEET_ID = "1GJ_ShKOHqaLLfpgWJHIrXu44fPMYBaSSSrRiVEbFkYU";
var SHEET_NAME = ""; // vazio = primeira aba

var NOME_ALIASES = ["evento", "nome do evento", "nome", "titulo"];
var INICIO_ALIASES = ["data inicio", "data de inicio", "inicio", "data"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Eventos de curta duração")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function norm_(s) {
  return String(s == null ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

function hasAny_(labels, aliases) {
  return aliases.some(function (a) { return labels.indexOf(a) !== -1; });
}

/**
 * Devolve a tabela no mesmo formato da resposta do Google Sheets (gviz),
 * que a página já sabe interpretar. Datas viram "Date(ano,mês-1,dia)".
 */
function getEventos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (!sheet) throw new Error("Aba não encontrada: " + SHEET_NAME);

  var range = sheet.getDataRange();
  var values = range.getValues();
  var shown = range.getDisplayValues();
  var tz = ss.getSpreadsheetTimeZone();

  // Cabeçalho: primeira linha (entre as 10 primeiras) com Evento e Data início.
  var h = 0;
  for (var i = 0; i < Math.min(10, values.length); i++) {
    var labels = values[i].map(norm_);
    if (hasAny_(labels, NOME_ALIASES) && hasAny_(labels, INICIO_ALIASES)) { h = i; break; }
  }

  var cols = (values[h] || []).map(function (label) { return { label: String(label) }; });
  var rows = values.slice(h + 1).map(function (row, r) {
    return {
      c: row.map(function (v, c) {
        if (v === "" || v === null) return null;
        if (v instanceof Date) {
          var y = Number(Utilities.formatDate(v, tz, "yyyy"));
          var m = Number(Utilities.formatDate(v, tz, "M")) - 1;
          var d = Number(Utilities.formatDate(v, tz, "d"));
          return { v: "Date(" + y + "," + m + "," + d + ")" };
        }
        return { v: v, f: shown[h + 1 + r][c] };
      })
    };
  });

  return {
    headerRow: h + 1,
    table: { cols: cols, rows: rows },
    user: Session.getActiveUser().getEmail()
  };
}
