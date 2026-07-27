const SPREADSHEET_ID = "1tTAVg4fn4Zf7-HA2ZBfkYUqoiQNMndNlRJ0wv0peqnQ";
const SHEET_NAME = "認領清單";
const HEADERS = ["id", "item", "owner", "amount", "confirmed", "createdAt", "updatedAt"];

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
  }

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = HEADERS.every((header, index) => firstRow[index] === header);

  if (!hasHeaders) {
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  return sheet;
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
      amount: row[3] === "" ? "" : Number(row[3]),
      confirmed: String(row[4] || "否"),
      createdAt: String(row[5] || ""),
      updatedAt: String(row[6] || "")
    }));

  return { ok: true, rows };
}

function createClaim(payload) {
  const sheet = getSheet();
  const now = new Date();
  const row = [
    Utilities.getUuid(),
    clean(payload.item),
    clean(payload.owner),
    Number(payload.amount || 0),
    payload.confirmed === "是" ? "是" : "否",
    now,
    now
  ];

  if (!row[1] || !row[2]) {
    throw new Error("品項和認領人都要填。");
  }

  sheet.appendRow(row);
  return { ok: true };
}

function updateClaim(payload) {
  const sheet = getSheet();
  const id = clean(payload.id);

  if (!id) {
    throw new Error("缺少資料 ID。");
  }

  const lastRow = sheet.getLastRow();
  const ids = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const index = ids.findIndex((value) => String(value) === id);

  if (index < 0) {
    throw new Error("找不到要更新的資料。");
  }

  const rowNumber = index + 2;
  const oldRow = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
  const nextRow = [
    id,
    clean(payload.item),
    clean(payload.owner),
    Number(payload.amount || 0),
    payload.confirmed === "是" ? "是" : "否",
    oldRow[5] || new Date(),
    new Date()
  ];

  if (!nextRow[1] || !nextRow[2]) {
    throw new Error("品項和認領人都要填。");
  }

  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([nextRow]);
  return { ok: true };
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

