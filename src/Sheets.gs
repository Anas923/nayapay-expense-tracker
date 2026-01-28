/**
 * Writes transaction data to the active sheet.
 * @param {Object[]} dataList Array of parsed transaction objects.
 */
function updateSheetWithTransactions(dataList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Find the right sheet
  let sheet = ss.getSheets()[0];
  let sheets = ss.getSheets();
  for (let s of sheets) {
    if (s.getRange(1, 1).getValue().toString().trim() === "Date") {
      sheet = s;
      break;
    }
  }

  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date",
      "Amount",
      "Currency",
      "Merchant",
      "Category",
      "Type",
      "Reference ID",
      "Note",
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f3f3f3");
    // Force Reference ID column (G/7) to be Plain Text to preserve leading zeroes
    sheet.getRange(1, 7, sheet.getMaxRows(), 1).setNumberFormat("@");
  }

  // 2. Load existing IDs
  const lastRow = sheet.getLastRow();
  let idSet = new Set();

  if (lastRow > 1) {
    // Force the column to Plain Text before reading just to be safe
    sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat("@");

    const rawIds = sheet
      .getRange(2, 7, lastRow - 1, 1)
      .getDisplayValues()
      .flat();
    rawIds.forEach((id) => {
      const cleanId = String(id).trim();
      if (cleanId) idSet.add(cleanId);
    });
  }

  // 3. Process new data
  dataList.forEach((data) => {
    const currentId = String(data.referenceId).trim();

    if (currentId && !idSet.has(currentId)) {
      let autoCategory = "";
      const merchant = data.merchant.toLowerCase();
      if (merchant.includes("nayatel")) autoCategory = "Bills";
      if (merchant.includes("mepco")) autoCategory = "Bills";
      if (merchant.includes("remitly")) autoCategory = "Remittance";
      if (merchant.includes("super mart")) autoCategory = "Food";
      if (merchant.includes("shell") || merchant.includes("fuel"))
        autoCategory = "Travel";

      sheet.appendRow([
        data.date,
        data.amount,
        "PKR",
        data.merchant,
        autoCategory,
        data.type,
        currentId,
        data.note,
      ]);
      idSet.add(currentId);
      console.log(`Added new transaction: ${currentId} (${data.merchant})`);
    }
  });

  SpreadsheetApp.flush();
  setupCategoryValidation();
}
