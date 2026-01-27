
window.addEventListener("error", (e) => {
  const el = document.getElementById("msg");
  if (el) el.textContent = "JS Error: " + (e.message || e.error);
});
// ====== 改成你的 Apps Script Web App /exec URL ======
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxxfqp5bvpbyM95aWH_oTwkkaZrU2Rk-jN-FtPgBUJtKVOiyvQaM6LngfLwBpj5r2gW/exec";
let SHEET_BOUNDS = null; // {lastRow,lastCol,headers}

async function ensureBounds(sheetId) {
  if (SHEET_BOUNDS && SHEET_BOUNDS.sheetId === sheetId) return;

  const res = await jsonp({ action: "getBounds", sheetId });
  if (!res.ok) throw new Error(res.error || "getBounds failed");
  SHEET_BOUNDS = { sheetId, ...res };

  // 更新 row max
  const rowInput = document.getElementById("rowNum");
  rowInput.max = String(res.lastRow || 1);

  // 更新 col dropdown：A..lastCol，並顯示 header
  const sel = document.getElementById("colLetter");
  sel.innerHTML = "";
  const A = "A".charCodeAt(0);
  const n = Math.min(res.lastCol || 1, 26); // 你目前只支援 A~Z
  for (let i = 0; i < n; i++) {
    const col = String.fromCharCode(A + i);
    const h = (res.headers && res.headers[i]) ? res.headers[i] : "";
    const opt = document.createElement("option");
    opt.value = col;
    opt.textContent = h ? `${col} (${h})` : col;
    sel.appendChild(opt);
  }
}




function extractSheetId(input) {
  const s = (input || "").trim();
  if (!s.startsWith("http")) return s;
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s;
}

function setMsg(t) {
  document.getElementById("msg").textContent = t;
}

function setReadOut(t) {
  document.getElementById("readOut").textContent = t;
}

// JSONP helper
function jsonp(params) {
  return new Promise((resolve, reject) => {
    const cbName = "cb_" + Math.random().toString(36).slice(2);
    window[cbName] = (data) => {
      resolve(data);
      cleanup();
    };

    function cleanup() {
      try { delete window[cbName]; } catch (e) { window[cbName] = undefined; }
      if (script && script.parentNode) script.parentNode.removeChild(script);
    }

    const q = new URLSearchParams(params);
    q.set("callback", cbName);
    q.set("_", Date.now().toString());

    const script = document.createElement("script");
    script.src = `${SCRIPT_URL}?${q.toString()}`;
    script.onerror = () => {
      cleanup();
      reject(new Error("JSONP request failed (network or URL issue)"));
    };
    document.head.appendChild(script);
  });
}

// (3) A+B -> C
document.getElementById("runBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) {
    setMsg("Please paste a sheet URL/ID.");
    return;
  }

  try {
    // 先抓一次 bounds（限制 row/col 範圍）
    await ensureBounds(sheetId);

    setMsg("Running A+B=C ...");
    const res = await jsonp({ action: "addCols", sheetId });
    if (!res.ok) throw new Error(res.error || "unknown");

    setMsg(res.message || "Done.");

    // ✅ 重要：計算後清 bounds，下一次 Load/Save 會重新抓最新的 lastRow/lastCol
    SHEET_BOUNDS = null;
  } catch (e) {
    setMsg("Failed: " + String(e));
  }
});


// (4) Load cell
document.getElementById("loadBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  const row = parseInt(document.getElementById("rowNum").value, 10);
  const col = document.getElementById("colLetter").value;

  if (!sheetId) {
    setMsg("Please paste a sheet URL/ID.");
    return;
  }
  if (!row || row < 1) {
    setMsg("Row must be >= 1.");
    return;
  }
  if (!col) {
    setMsg("Please select a column.");
    return;
  }

  try {
    // ✅ 重要：Load 前確保 bounds 是最新的
    await ensureBounds(sheetId);

    // 可選：避免剛寫入後 sheet 顯示層延遲
    await new Promise(r => setTimeout(r, 300));

    setMsg("Loading cell...");
    const res = await jsonp({ action: "getCell", sheetId, row: String(row), col });
    if (!res.ok) throw new Error(res.error || "unknown");

    setReadOut(
      `data point ${res.row} has ${res.featureName} (${res.value})\n` +
      `data type: ${res.type}\n` +
      `cell: ${res.col}${res.row}`
    );
    setMsg("Loaded.");
  } catch (e) {
    setMsg("Load failed: " + String(e));
  }
});


// (4) Save cell
document.getElementById("saveBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  const row = parseInt(document.getElementById("rowNum").value, 10);
  const col = document.getElementById("colLetter").value;
  const value = document.getElementById("newValue").value;

  if (!sheetId) {
    setMsg("Please paste a sheet URL/ID.");
    return;
  }
  if (!row || row < 1) {
    setMsg("Row must be >= 1.");
    return;
  }
  if (!col) {
    setMsg("Please select a column.");
    return;
  }

  try {
    await ensureBounds(sheetId);

    setMsg("Saving...");
    const res = await jsonp({ action: "setCell", sheetId, row: String(row), col, value });
    if (!res.ok) throw new Error(res.error || "unknown");

    setMsg(`Saved: ${col}${row} = ${res.writtenValue} (type=${res.writtenType})`);

    // ✅ 寫入後，清 bounds（如果寫入造成 lastRow/lastCol 變動）
    SHEET_BOUNDS = null;

    // ✅ 讀回驗證：確保顯示的是最新值（你要的 rigor）
    const reread = await jsonp({ action: "getCell", sheetId, row: String(row), col });
    if (reread.ok) {
      setReadOut(
        `data point ${reread.row} has ${reread.featureName} (${reread.value})\n` +
        `data type: ${reread.type}\n` +
        `cell: ${reread.col}${reread.row}`
      );
    }
  } catch (e) {
    setMsg("Save failed: " + String(e));
  }
});

