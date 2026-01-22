function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetId = data.sheetId;

    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0]; // demo

    // demo：fix A,B -> calculate C=A+B
    const readA = sheet.getRange("A:A").getValues();
    const readB = sheet.getRange("B:B").getValues();

    // Find the last row
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return _ok("No data rows");

    // Write column head
    sheet.getRange("C1").setValue("sum");

    // Calculation row by row (start from the second row)
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
