function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetId = data.sheetId;

    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0]; // demo：用第一張工作表

    // demo：固定 A,B → C
    const readA = sheet.getRange("A:A").getValues();
    const readB = sheet.getRange("B:B").getValues();

    // 找到最後一列（以 A 欄為準）
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return _ok("No data rows");

    // 寫表頭
    sheet.getRange("C1").setValue("sum");

    // 逐列計算（從第2列開始）
    const out = [];
    for (let r = 2; r <= lastRow; r++) {
      const a = Number(readA[r - 1][0]) || 0;
      const b = Number(readB[r - 1][0]) || 0;
      out.push([a + b]);
    }

    sheet.getRange(2, 3, out.length, 1).setValues(out); // (row=2, col=3=C)
    return _ok("Wrote " + out.length + " rows");
  } catch (err) {
    return _ok("Error: " + err);
  }
}

function _ok(text) {
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.TEXT);
}
