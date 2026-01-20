//Attention Needed!!!!!!!! Replace the address of back end codes in "placeholder=" -->
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
    msg.textContent = "Please paste Google Sheet link";
    return;
  }

  const payload = {
    sheetId,
    readCols: ["A", "B"],
    writeCol: "C",
    header: "sum"
  };

  msg.textContent = "Request sent(writing)...";

  try {
    await fetch(SCRIPT_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });

    msg.textContent = `Please check the sheet again for results`;
  } catch (e) {
    msg.textContent = "Failed to send request" + String(e);
  }
});
