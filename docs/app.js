const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxxfqp5bvpbyM95aWH_oTwkkaZrU2Rk-jN-FtPgBUJtKVOiyvQaM6LngfLwBpj5r2gW/exec";

function extractSheetId(input) {
  const s = (input || "").trim();
  if (!s.startsWith("http")) return s; // user pasted pure id
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s;
}

document.getElementById("runBtn").addEventListener("click", async () => {
  const sheetInput = document.getElementById("sheetId").value;
  const sheetId = extractSheetId(sheetInput);
  const msg = document.getElementById("msg");

  if (!sheetId) {
    msg.textContent = "請貼上 Google Sheet 連結或 sheetId";
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
    await fetch(SCRIPT_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });

    msg.textContent = `已送出！回到 Google Sheet 看 C 欄是否出現 sum 與結果。`;
  } catch (e) {
    msg.textContent = "送出失敗：" + String(e);
  }
});
