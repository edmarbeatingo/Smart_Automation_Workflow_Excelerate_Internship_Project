function consolidateLiveMirror() {

  const sources = [
    { id: "1-2PRBt08G8i1Bsk9-Ktb1mZLDKHVPLw7K3DTTsKF5e8", sheet: "MASTERFILE" },
    { id: "1XA2P3e_QZzO5Qkwsie9kxmHE9St-zVcTyvMsO89p2XA", sheet: "Email_Orders" },
    { id: "1CJrkLYMZbdFQ6i9VH040ZwcON44phgrneFMSkjAU6so", sheet: "ORDER_FORM (Responses)" }
  ];

  const targetSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Centralized_Order");

  if (!targetSheet) {
    Logger.log("Centralized_Order sheet not found.");
    return;
  }

  let masterData = [];

  sources.forEach(source => {

    const ss = SpreadsheetApp.openById(source.id);
    const sheet = ss.getSheetByName(source.sheet);

    if (!sheet) {
      Logger.log("Sheet not found: " + source.sheet);
      return;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    data.shift(); 

    data.forEach(row => {

      let completionRaw = row[15]; 

      if (!completionRaw) return;

      let completionDate = completionRaw instanceof Date
        ? completionRaw
        : new Date(completionRaw);

      if (isNaN(completionDate.getTime())) return;

      row[15] = completionDate;
      masterData.push(row);

    });

  });

  masterData.sort((a, b) => b[15] - a[15]);

  if (targetSheet.getLastRow() > 1) {
    targetSheet
      .getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn())
      .clearContent();
  }

  if (masterData.length > 0) {
    targetSheet.getRange(
      2,
      1,
      masterData.length,
      masterData[0].length
    ).setValues(masterData);
  }

  Logger.log("Live mirror sync completed successfully.");
}
