window.addEventListener("error", (e) => {
  const el = document.getElementById("msg");
  if (el) el.textContent = "JS Error: " + (e.message || e.error);
});

// ====== 改成你的 Apps Script Web App /exec URL ======
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxxfqp5bvpbyM95aWH_oTwkkaZrU2Rk-jN-FtPgBUJtKVOiyvQaM6LngfLwBpj5r2gW/exec";

let SHEET_BOUNDS = null; // {sheetId,lastRow,lastCol,headers,ms}
let PLOT = null; // Chart.js instance

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

function setPlotInfo(t) {
  const el = document.getElementById("plotInfo");
  if (el) el.textContent = t;
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
    q.set("_", Date.now().toString()); // bust cache

    const script = document.createElement("script");
    script.src = `${SCRIPT_URL}?${q.toString()}`;
    script.onerror = () => {
      cleanup();
      reject(new Error("JSONP request failed (network or URL issue)"));
    };
    document.head.appendChild(script);
  });
}

// 初始化 dropdown（避免空的）
(function initDefaults() {
  const colSel = document.getElementById("colLetter");
  const plotSel = document.getElementById("plotCol");
  if (colSel && colSel.options.length === 0) colSel.add(new Option("A", "A"));
  if (plotSel && plotSel.options.length === 0) plotSel.add(new Option("A", "A"));
})();

async function ensureBounds(sheetId) {
  if (SHEET_BOUNDS && SHEET_BOUNDS.sheetId === sheetId) return SHEET_BOUNDS;

  const res = await jsonp({ action: "getBounds", sheetId });
  if (!res.ok) throw new Error(res.error || "getBounds failed");
  SHEET_BOUNDS = { sheetId, ...res };

  // row max
  const rowInput = document.getElementById("rowNum");
  if (rowInput) rowInput.max = String(res.lastRow || 1);

  // build options A..lastCol (<=26)
  const A = "A".charCodeAt(0);
  const n = Math.max(1, Math.min(Number(res.lastCol || 1), 26));

  const headers = res.headers || [];

  function buildOptions() {
    const frag = document.createDocumentFragment();
    for (let i = 0; i < n; i++) {
      const col = String.fromCharCode(A + i);
      const h = headers[i] ? String(headers[i]) : "";
      frag.appendChild(new Option(h ? `${col} (${h})` : col, col));
    }
    return frag;
  }

  // update colLetter
  const colSel = document.getElementById("colLetter");
  if (colSel) {
    const prev = colSel.value || "A";
    colSel.replaceChildren(buildOptions());
    colSel.value = prev;
    if (!colSel.value) colSel.value = "A";
  }

  // update plotCol
  const plotSel = document.getElementById("plotCol");
  if (plotSel) {
    const prev = plotSel.value || "A";
    plotSel.replaceChildren(buildOptions());
    plotSel.value = prev;
    if (!plotSel.value) plotSel.value = "A";
  }

  return SHEET_BOUNDS;
}

// ===== Buttons =====

// Load bounds
document.getElementById("loadSheetBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) { setMsg("Please paste a sheet URL/ID first."); return; }
  try {
    setMsg("Loading sheet bounds...");
    const b = await ensureBounds(sheetId);
    const lastColLetter = String.fromCharCode("A".charCodeAt(0) + Math.min(b.lastCol || 1, 26) - 1);
    setMsg(`Bounds loaded: rows=1..${b.lastRow}, cols=A..${lastColLetter} (backend ms=${b.ms ?? "N/A"})`);
  } catch (e) {
    setMsg("Failed to load bounds: " + String(e));
  }
});

// (3) A+B -> C
document.getElementById("runBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }

  try {
    setMsg("Running A+B=C ...");
    const res = await jsonp({ action: "addCols", sheetId });
    if (!res.ok) throw new Error(res.error || "unknown");
    setMsg(`${res.message || "Done."} (backend ms=${res.ms ?? "N/A"})`);
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

  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }
  if (!row || row < 1) { setMsg("Row must be >= 1."); return; }
  if (!col) { setMsg("Please select a column."); return; }

  try {
    await ensureBounds(sheetId);
    setMsg("Loading cell...");
    const res = await jsonp({ action: "getCell", sheetId, row: String(row), col });
    if (!res.ok) throw new Error(res.error || "unknown");

    setReadOut(
      `data point ${res.row} has ${res.featureName} (${res.value})\n` +
      `data type: ${res.type}\n` +
      `cell: ${res.col}${res.row}\n` +
      `backend ms: ${res.ms ?? "N/A"}`
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

  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }
  if (!row || row < 1) { setMsg("Row must be >= 1."); return; }
  if (!col) { setMsg("Please select a column."); return; }

  try {
    await ensureBounds(sheetId);
    setMsg("Saving...");
    const res = await jsonp({ action: "setCell", sheetId, row: String(row), col, value });
    if (!res.ok) throw new Error(res.error || "unknown");

    setMsg(`Saved: ${col}${row} = ${res.writtenValue} (type=${res.writtenType}) (backend ms=${res.ms ?? "N/A"})`);
    SHEET_BOUNDS = null;

    const reread = await jsonp({ action: "getCell", sheetId, row: String(row), col });
    if (reread.ok) {
      setReadOut(
        `data point ${reread.row} has ${reread.featureName} (${reread.value})\n` +
        `data type: ${reread.type}\n` +
        `cell: ${reread.col}${reread.row}\n` +
        `backend ms: ${reread.ms ?? "N/A"}`
      );
    }
  } catch (e) {
    setMsg("Save failed: " + String(e));
  }
});

// ===== Plot =====

function isMostlyNumeric(arr) {
  let num = 0, non = 0;
  for (const s of arr) {
    const t = String(s ?? "").trim();
    if (!t) continue;
    const v = Number(t);
    if (!Number.isNaN(v)) num++; else non++;
  }
  return num >= non;
}

function buildCategoryCounts(arr, topK = 30) {
  const map = new Map();
  for (const s of arr) {
    const t = String(s ?? "").trim();
    if (!t) continue;
    map.set(t, (map.get(t) || 0) + 1);
  }
  const items = [...map.entries()].sort((a,b) => b[1]-a[1]).slice(0, topK);
  return { labels: items.map(x=>x[0]), values: items.map(x=>x[1]) };
}

function plotBar(labels, values, title) {
  const canvas = document.getElementById("plotCanvas");
  const ctx = canvas.getContext("2d");
  if (PLOT) PLOT.destroy();

  PLOT = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: title, data: values }]
    },
    options: {
      responsive: true,
      plugins: { title: { display: true, text: title } }
    }
  });
}

document.getElementById("plotBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }

  try {
    await ensureBounds(sheetId);
    const col = document.getElementById("plotCol").value || "A";

    setMsg(`Fetching column ${col}...`);
    setPlotInfo("");

    const t0 = performance.now();
    const res = await jsonp({ action: "getColumn", sheetId, col });
    const t1 = performance.now();

    if (!res.ok) throw new Error(res.error || "getColumn failed");

    const vals = res.values || [];
    const header = res.header || col;

    if (vals.length === 0) {
      setMsg("No data rows to plot.");
      if (PLOT) PLOT.destroy();
      return;
    }

    // numeric vs categorical
    if (isMostlyNumeric(vals)) {
      const cap = 200; // 避免幾千根 bar 爆炸
      const sliced = vals.slice(0, cap);
      const numbers = sliced.map(v => {
        const x = Number(String(v).trim());
        return Number.isNaN(x) ? 0 : x;
      });
      const labels = numbers.map((_, i) => String(i + 2)); // row start at 2
      plotBar(labels, numbers, `${header} (first ${sliced.length} rows)`);
      setMsg(`Plotted numeric column ${col}.`);
    } else {
      const cc = buildCategoryCounts(vals, 30);
      plotBar(cc.labels, cc.values, `${header} (top categories)`);
      setMsg(`Plotted categorical column ${col}.`);
    }

    setPlotInfo(`backend ms=${res.ms ?? "N/A"}, fetch+render ms=${Math.round(t1 - t0)}`);
  } catch (e) {
    setMsg("Plot failed: " + String(e));
  }
});

// Save PNG
document.getElementById("savePngBtn").addEventListener("click", () => {
  const canvas = document.getElementById("plotCanvas");
  const url = canvas.toDataURL("image/png");
  const a = document.createElement("a");
  a.href = url;
  a.download = "plot.png";
  a.click();
});

// Save PDF
document.getElementById("savePdfBtn").addEventListener("click", () => {
  const canvas = document.getElementById("plotCanvas");
  const imgData = canvas.toDataURL("image/png");

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

  // fit image into page
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = 30;

  const imgW = pageW - margin * 2;
  const imgH = (canvas.height / canvas.width) * imgW;

  const y = (pageH - imgH) / 2;
  pdf.addImage(imgData, "PNG", margin, Math.max(margin, y), imgW, imgH);
  pdf.save("plot.pdf");
});
