document.getElementById("runBtn").addEventListener("click", async () => {
  const sheetId = document.getElementById("sheetId").value.trim();
  const scriptUrl = document.getElementById("scriptUrl").value.trim();
  const msg = document.getElementById("msg");

  if (!sheetId || !scriptUrl) {
    msg.textContent = "請填 sheetId 與 scriptUrl";
    return;
  }

  // 送給 Apps Script 的資料
  const payload = {
    sheetId,
    // 這個 demo 固定：讀 A,B 寫到 C，並寫表頭 "sum"
    readCols: ["A", "B"],
    writeCol: "C",
    header: "sum"
  };

  msg.textContent = "已送出請求（寫入中）...";

  // ⚠️ 為了避開 CORS 麻煩，這裡用 no-cors：
  // 你將拿不到可讀的回應，但 Apps Script 仍會收到請求並寫回。
  try {
    await fetch(scriptUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });

    msg.textContent = "請求已送出！到 Google Sheet 看看 C 欄是否出現 sum 與計算結果。";
  } catch (e) {
    msg.textContent = "送出失敗：" + String(e);
  }
});
