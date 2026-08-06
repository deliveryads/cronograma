/**
 * Backend do Cronograma de Responsabilidades — Google Apps Script Web App.
 *
 * Como implantar: veja APPS_SCRIPT.md na raiz do repositório.
 *
 * Este script espera 7 abas na planilha, todas com cabeçalho na primeira linha:
 *
 *   Squad              -> id | nome | roteiros
 *   Analistas          -> id | nome | squadId | funcao
 *   Férias             -> id | titulo | analistaId | squadId | dataInicio | dataFim | descricao
 *   Cursos             -> id | titulo | analistaId | squadId | dataInicio | dataFim | descricao
 *   Extraordinários    -> id | titulo | analistaId | squadId | dataInicio | dataFim | descricao
 *   Datas Importantes  -> id | titulo | analistaId | squadId | dataInicio | dataFim | descricao
 *   Ponto              -> id | titulo | analistaId | squadId | dataInicio | dataFim | descricao
 *
 * As 5 últimas abas compartilham exatamente as mesmas colunas — cada uma vira um
 * "tipo" de evento (ferias, curso, reuniao, data_importante, ponto) na página.
 */

var SHEET_TIPOS = {
  "Férias": "ferias",
  "Cursos": "curso",
  "Extraordinários": "reuniao",
  "Datas Importantes": "data_importante",
  "Ponto": "ponto"
};

var TIPO_TO_SHEET = {
  ferias: "Férias",
  curso: "Cursos",
  reuniao: "Extraordinários",
  data_importante: "Datas Importantes",
  ponto: "Ponto"
};

var EVENT_HEADERS = ["id", "titulo", "analistaId", "squadId", "dataInicio", "dataFim", "descricao"];

function readSheetAsObjects_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1)
    .filter(function (row) {
      return row.some(function (cell) { return cell !== "" && cell !== null; });
    })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (header, i) { obj[header] = row[i]; });
      return obj;
    });
}

function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return value ? String(value) : "";
}

function doGet(e) {
  var squads = readSheetAsObjects_("Squad").map(function (s) {
    return {
      id: String(s.id),
      nome: String(s.nome || ""),
      roteiros: s.roteiros
        ? String(s.roteiros).split(",").map(function (r) { return r.trim(); }).filter(Boolean)
        : []
    };
  });

  var analistas = readSheetAsObjects_("Analistas").map(function (a) {
    return {
      id: String(a.id),
      nome: String(a.nome || ""),
      squadId: String(a.squadId || ""),
      funcao: String(a.funcao || "")
    };
  });

  var eventos = [];
  Object.keys(SHEET_TIPOS).forEach(function (sheetName) {
    var tipo = SHEET_TIPOS[sheetName];
    readSheetAsObjects_(sheetName).forEach(function (row) {
      eventos.push({
        id: String(row.id),
        tipo: tipo,
        titulo: String(row.titulo || ""),
        analistaId: row.analistaId ? String(row.analistaId) : "",
        squadId: row.squadId ? String(row.squadId) : "",
        dataInicio: formatDate_(row.dataInicio),
        dataFim: formatDate_(row.dataFim),
        descricao: row.descricao ? String(row.descricao) : ""
      });
    });
  });

  return jsonResponse_({ squads: squads, analistas: analistas, eventos: eventos });
}

function sheetForTipo_(tipo) {
  var name = TIPO_TO_SHEET[tipo];
  if (!name) throw new Error("Tipo de evento inválido: " + tipo);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Aba não encontrada: "' + name + '"');
  return sheet;
}

function findRowById_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // número de linha (1-indexado)
  }
  return -1;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var tipo = payload.tipo;
    var evento = payload.evento || {};
    var sheet = sheetForTipo_(tipo);

    if (action === "create") {
      var id = Utilities.getUuid();
      var row = EVENT_HEADERS.map(function (h) { return h === "id" ? id : (evento[h] || ""); });
      sheet.appendRow(row);
      evento.id = id;
      evento.tipo = tipo;
      return jsonResponse_({ ok: true, evento: evento });
    }

    if (action === "update") {
      var rowNum = findRowById_(sheet, evento.id);
      if (rowNum === -1) return jsonResponse_({ ok: false, error: "Evento não encontrado: " + evento.id });
      var updateRow = EVENT_HEADERS.map(function (h) { return h === "id" ? evento.id : (evento[h] || ""); });
      sheet.getRange(rowNum, 1, 1, EVENT_HEADERS.length).setValues([updateRow]);
      evento.tipo = tipo;
      return jsonResponse_({ ok: true, evento: evento });
    }

    if (action === "delete") {
      var delRowNum = findRowById_(sheet, evento.id);
      if (delRowNum === -1) return jsonResponse_({ ok: false, error: "Evento não encontrado: " + evento.id });
      sheet.deleteRow(delRowNum);
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: "Ação inválida: " + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
