const SPREADSHEET_ID = "1mQpPKhT7xOFCOXHWlmkdCFbeFH1accwwzfGMDfLU5RQ";
const PLAYLIST_SHEET_NAME = "Playlist";
const PLAYLIST_HEADERS = ["Fecha y hora", "Canción"];

function doGet(e) {
  const parameters = e && e.parameter ? e.parameter : {};
  const action = String(parameters.action || "");
  if (action !== "songs" && action !== "submit") {
    return ContentService.createTextOutput("Servicio de playlist activo.");
  }

  const callback = String(parameters.callback || "");
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    throw new Error("Solicitud inválida.");
  }

  let result;
  if (action === "submit") {
    if (parameters.website) {
      result = { ok: false, error: "Solicitud rechazada." };
    } else {
      const song = normalizeSong_(parameters.song);
      if (!song) {
        result = { ok: false, error: "La canción debe tener entre 1 y 160 caracteres." };
      } else {
        saveSong_(song);
        result = { ok: true };
      }
    }
  } else {
    result = { songs: getSongs_() };
  }

  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(result) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function setupPlaylistSheet() {
  const sheet = getPlaylistSheet_();
  ensureHeaders_(sheet);
  sheet.getRange("A:A").setNumberFormat("dd/MM/yyyy HH:mm");
}

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.website) {
    return jsonResponse_({ ok: false, error: "Solicitud rechazada." });
  }

  const song = normalizeSong_(params.song);
  if (!song) {
    return jsonResponse_({ ok: false, error: "La canción debe tener entre 1 y 160 caracteres." });
  }

  saveSong_(song);
  return jsonResponse_({ ok: true });
}

function normalizeSong_(value) {
  const song = String(value || "").replace(/[\u0000-\u001f\u007f]/g, " ").trim();
  return song && song.length <= 160 ? song : "";
}

function saveSong_(song) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getPlaylistSheet_();
    ensureHeaders_(sheet);
    sheet.appendRow([new Date(), escapeFormula_(song)]);
  } finally {
    lock.releaseLock();
  }
}

function getSongs_() {
  const sheet = getPlaylistSheet_();
  ensureHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  return sheet.getRange(2, 1, lastRow - 1, PLAYLIST_HEADERS.length).getValues()
    .filter(function (row) {
      return String(row[1] || "").trim();
    })
    .map(function (row) {
      return {
        date: row[0] instanceof Date
          ? Utilities.formatDate(row[0], timeZone, "dd/MM/yyyy HH:mm")
          : String(row[0] || ""),
        song: String(row[1] || "").replace(/^'(?=[=+\-@])/, "")
      };
    });
}

function getPlaylistSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(PLAYLIST_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(PLAYLIST_SHEET_NAME);
  }

  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, PLAYLIST_HEADERS.length).setValues([PLAYLIST_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  const actualHeaders = sheet.getRange(1, 1, 1, PLAYLIST_HEADERS.length).getValues()[0];
  if (PLAYLIST_HEADERS.some(function (header, index) {
    return actualHeaders[index] !== header;
  })) {
    throw new Error('La primera fila debe contener: "' + PLAYLIST_HEADERS.join('", "') + '".');
  }
}

function escapeFormula_(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
