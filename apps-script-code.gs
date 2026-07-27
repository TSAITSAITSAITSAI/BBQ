const SPREADSHEET_ID = "1tTAVg4fn4Zf7-HA2ZBfkYUqoiQNMndNlRJ0wv0peqnQ";
const SHEET_NAME = "認領清單";
const HEADERS = ["id", "item", "owner", "amount", "note", "confirmed", "createdAt", "updatedAt"];

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const result = params.action === "list" ? listClaims() : { ok: true, rows: [] };
  return output(result, params.callback);
}

function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    const action = payload.action || "create";

    if (action === "update") {
      return output(updateClaim(payload));
    }

    if (action === "delete") {
      return output(deleteClaim(payload));
    }

    return output(createClaim(payload));
  } catch (error) {
    return output({ ok: false, message: error.message });
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return sheet;
  }

  migrateHeaders(sheet);
  return sheet;
}

function migrateHeaders(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);

  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((value) => String(value || "").trim());
  const isCurrent = HEADERS.every((header, index) => currentHeaders[index] === header);

  if (isCurrent) {
    return;
  }

  const headerIndex = {};
  currentHeaders.forEach((header, index) => {
    if (header) headerIndex[header] = index;
  });

  const oldRows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const nextRows = oldRows.map((row) => HEADERS.map((header) => {
    if (header === "note" && headerIndex[header] === undefined) return "";
    const index = headerIndex[header];
    return index === undefined ? "" : row[index];
  }));

  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  if (nextRows.length) {
    sheet.getRange(2, 1, nextRows.length, HEADERS.length).setValues(nextRows);
  }
}

function listClaims() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { ok: true, rows: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const rows = values
    .filter((row) => row[0] || row[1] || row[2])
    .map((row) => ({
      id: String(row[0] || ""),
      item: String(row[1] || ""),
      owner: String(row[2] || ""),
      amount: String(row[3] || ""),
      note: String(row[4] || ""),
      confirmed: String(row[5] || "否"),
      createdAt: String(row[6] || ""),
      updatedAt: String(row[7] || "")
    }));

  return { ok: true, rows };
}

function createClaim(payload) {
  const sheet = getSheet();
  const now = new Date();
  const item = clean(payload.item);
  const owner = clean(payload.owner);

  if (!item || !owner) {
    throw new Error("品項和認領人都要填。");
  }

  sheet.appendRow([
    Utilities.getUuid(),
    item,
    owner,
    clean(payload.amount),
    clean(payload.note),
    payload.confirmed === "是" ? "是" : "否",
    now,
    now
  ]);

  return { ok: true };
}

function updateClaim(payload) {
  const sheet = getSheet();
  const id = clean(payload.id);

  if (!id) {
    throw new Error("缺少資料 ID。");
  }

  const rowNumber = findRowNumberById(sheet, id);
  const oldRow = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
  const item = clean(payload.item);
  const owner = clean(payload.owner);

  if (!item || !owner) {
    throw new Error("品項和認領人都要填。");
  }

  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([[
    id,
    item,
    owner,
    clean(payload.amount),
    clean(payload.note),
    payload.confirmed === "是" ? "是" : "否",
    oldRow[6] || new Date(),
    new Date()
  ]]);

  return { ok: true };
}

function deleteClaim(payload) {
  const sheet = getSheet();
  const id = clean(payload.id);

  if (!id) {
    throw new Error("缺少資料 ID。");
  }

  sheet.deleteRow(findRowNumberById(sheet, id));
  return { ok: true };
}

function findRowNumberById(sheet, id) {
  const lastRow = sheet.getLastRow();
  const ids = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const index = ids.findIndex((value) => String(value) === id);

  if (index < 0) {
    throw new Error("找不到要處理的資料。");
  }

  return index + 2;
}

function output(data, callback) {
  const json = JSON.stringify(data);

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function clean(value) {
  return String(value || "").trim();
}
