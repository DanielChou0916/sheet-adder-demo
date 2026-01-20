function extractSheetId(input) {
  const s = (input || "").trim();
  // 若使用者直接貼 ID（沒有 http），就原樣回傳
  if (!s.startsWith("http")) return s;

  // 從 URL 抽出 /d/<ID>/
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s; // 抽不到就回傳原字串（之後會報錯）
}

document.getElementById("runBtn").addEventListener("click", async () => {
  const sheetInput = document.getElementById("sheetId").value;
  const sheetId = extractSheetId(sheetInput);

  const scriptUrl = document.getElementById("scriptUrl").value.trim();
  const msg = document.getElementById("msg");

  if (!sheetId || !scriptUrl) {
    msg.textContent = "請填 sheet 連結/ID 與 Apps Script Web App URL";
    return;
  }

  const payload = {
    sheetId,
    readCols: ["A", "B"],
    writeCol: "C",
    header: "sum"
  };

  msg.textContent = "已送出請求（寫入中）...";

  try {
    await fetch(scriptUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });

    msg.textContent = `請求已送出！(sheetId=${sheetId}) 回到 Google Sheet 看 C 欄是否出現 sum 與結果。`;
  } catch (e) {
    msg.textContent = "送出失敗：" + String(e);
  }
});
