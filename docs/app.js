// ====== 改成你的 Apps Script Web App /exec URL ======
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxxfqp5bvpbyM95aWH_oTwkkaZrU2Rk-jN-FtPgBUJtKVOiyvQaM6LngfLwBpj5r2gW/exec";

// A~Z 下拉
(function initColDropdown() {
  const sel = document.getElementById("colLetter");
  const A = "A".charCodeAt(0);
  for (let i = 0; i < 26; i++) {
    const c = String.fromCharCode(A + i);
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    sel.appendChild(opt);
  }
})();

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
  setMsg("Running A+B=C ...");
  try {
    const res = await jsonp({ action: "addCols", sheetId });
    if (!res.ok) throw new Error(res.error || "unknown");
    setMsg(res.message || "Done.");
  } catch (e) {
    setMsg("Failed: " + String(e));
  }
});

// (4) Load cell
document.getElementById("loadBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  const row = parseInt(document.getElementById("rowNum").value, 10);
  const col = document.getElementById("colLetter").value;

  if (!sheetId || !row || !col) {
    setMsg("Please fill sheetId, row, col.");
    return;
  }

  setMsg("Loading cell...");
  try {
    const res = await jsonp({ action: "getCell", sheetId, row: String(row), col });
    if (!res.ok) throw new Error(res.error || "unknown");

    // 你要的輸出格式：
    // 1) data point {row} has {feature_name} (its value).
    // 2) And data type
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

  if (!sheetId || !row || !col) {
    setMsg("Please fill sheetId, row, col.");
    return;
  }

  setMsg("Saving...");
  try {
    const res = await jsonp({ action: "setCell", sheetId, row: String(row), col, value });
    if (!res.ok) throw new Error(res.error || "unknown");

    setMsg(`Saved: ${col}${row} = ${res.writtenValue} (type=${res.writtenType})`);

    // 自動再讀一次刷新顯示
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
