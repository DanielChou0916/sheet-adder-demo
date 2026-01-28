function timed_(fn) {
  // time cost estimation
  const t0 = Date.now();
  const result = fn();
  const ms = Date.now() - t0;

  // 確保回傳物件都有 ms
  if (result && typeof result === "object") {
    result.ms = ms;
    return result;
  }
  return { ok: true, result, ms };
}


function doGet(e) {
  const p = e.parameter || {};
  const action = (p.action || "").trim();

  try {
  if (action === "getBounds") return jsonp_(p.callback, timed_(() => handleGetBounds_(p)));
  if (action === "addCols") return jsonp_(p.callback, timed_(() => handleAddCols_(p)));
  if (action === "getCell") return jsonp_(p.callback, timed_(() => handleGetCell_(p)));
  if (action === "setCell") return jsonp_(p.callback, timed_(() => handleSetCell_(p)));

    return jsonp_(p.callback, { ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonp_(p.callback, { ok: false, error: String(err && err.stack ? err.stack : err) });
  }
}

// ===== (3) 保留：你原本的 A+B->C 計算 =====
function handleGetBounds_(p) {
  const sheetId = required_(p, "sheetId");
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheets()[0];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const headers = (lastCol >= 1)
    ? sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    : [];

  return { ok: true, lastRow, lastCol, headers };
}

function handleAddCols_(p) {
  const sheetId = required_(p, "sheetId");

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0]; // demo：第一個 sheet

  // demo：fix A,B -> calculate C=A+B
  // const readA = sheet.getRange("A:A").getValues();
  // const readB = sheet.getRange("B:B").getValues();

  const lastRow = sheet.getLastRow();
  const aVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // A2:A(lastRow)
  const bVals = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B2:B(lastRow)
  if (lastRow < 2) return { ok: true, message: "No data rows" };

  sheet.getRange("C1").setValue("sum");

  const out = [];
  for (let i = 0; i < (lastRow - 1); i++) {
    const a = Number(aVals[i][0]) || 0;
    const b = Number(bVals[i][0]) || 0;
    out.push([a + b]);
  }

  sheet.getRange(2, 3, out.length, 1).setValues(out); // (row=2, col=3=C)
  return { ok: true, message: "Wrote " + out.length + " rows" };
}

// ===== (4) 新增：讀單一格，回傳 featureName/value/type =====
function handleGetCell_(p) {
  const sheetId = required_(p, "sheetId");
  const row = parseInt(required_(p, "row"), 10);
  const col = (required_(p, "col") || "").toUpperCase();

  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheets()[0];

  const colIdx = colToIndexAZ_(col);

  // feature_name 用第 1 列 header
  const featureName = String(sh.getRange(1, colIdx).getDisplayValue() || "");

  const cell = sh.getRange(row, colIdx);
  const v = cell.getValue();
  const display = cell.getDisplayValue();

  return {
    ok: true,
    row,
    col,
    featureName: featureName || "(no header)",
    value: display,
    type: classify_(v)
  };
}

// ===== (4) 新增：覆寫單一格 =====
function handleSetCell_(p) {
  const sheetId = required_(p, "sheetId");
  const row = parseInt(required_(p, "row"), 10);
  const col = (required_(p, "col") || "").toUpperCase();
  const raw = (p.value === undefined || p.value === null) ? "" : String(p.value);

  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheets()[0];
  const colIdx = colToIndexAZ_(col);

  const cell = sh.getRange(row, colIdx);

  // MVP：嘗試轉型，不行就當字串
  const parsed = smartParse_(raw);
  cell.setValue(parsed);

  const v2 = cell.getValue();
  return {
    ok: true,
    row,
    col,
    writtenValue: cell.getDisplayValue(),
    writtenType: classify_(v2)
  };
}

// ===== helpers =====
function jsonp_(callbackName, obj) {
  const cb = (callbackName && String(callbackName).trim()) ? String(callbackName).trim() : null;
  const payload = JSON.stringify(obj);
  const text = cb ? `${cb}(${payload});` : payload;

  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function required_(p, key) {
  const v = p[key];
  if (v === undefined || v === null || String(v).trim() === "") {
    throw new Error("Missing parameter: " + key);
  }
  return String(v).trim();
}

// A~Z only
function colToIndexAZ_(col) {
  if (!/^[A-Z]$/.test(col)) throw new Error("Invalid col (A-Z only): " + col);
  return (col.charCodeAt(0) - "A".charCodeAt(0)) + 1;
}

function classify_(v) {
  if (v === null || v === "" || typeof v === "undefined") return "blank";
  if (v instanceof Date) return "date";
  if (typeof v === "number") return "number";
  if (typeof v === "boolean") return "boolean";
  return "string";
}

function smartParse_(raw) {
  const s = String(raw);
  if (s.trim() === "") return "";

  if (/^(true|false)$/i.test(s.trim())) return s.trim().toLowerCase() === "true";
  if (/^[+-]?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(s.trim())) return Number(s.trim());
  if (/^\d{4}-\d{2}-\d{2}$/.test(s.trim())) {
    const d = new Date(s.trim() + "T00:00:00");
    if (!isNaN(d.getTime())) return d;
  }
  return s;
}
