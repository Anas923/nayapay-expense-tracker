/**
 * Initializes the Settings and Summary sheets if they don't exist.
 */
function initializeSmartSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Settings Sheet
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.appendRow(["Setting", "Value"]);
    settingsSheet
      .getRange(1, 1, 1, 2)
      .setFontWeight("bold")
      .setBackground("#f3f3f3");

    // Add Opening Balance setting
    settingsSheet.appendRow(["Opening Balance (Jan 1st)", 0]);
    settingsSheet.appendRow([]); // Divider

    // Add Categories header
    settingsSheet.appendRow(["Categories", ""]);
    settingsSheet.getRange(4, 1).setFontWeight("bold").setBackground("#f3f3f3");

    // Add default categories
    const defaults = [
      "Food",
      "Travel",
      "Bills",
      "Shopping",
      "Entertainment",
      "Remittance",
      "Investment",
      "Salary",
      "Transfer",
      "LoanByme",
      "LoanFromme",
    ];
    defaults.forEach((cat) => settingsSheet.appendRow([cat, ""]));
  } else {
    // If it exists, check if "Opening Balance" is there at row 2
    const firstCell = settingsSheet.getRange(2, 1).getValue().toString().trim();
    if (!firstCell.includes("Opening Balance")) {
      // Logic to upgrade existing sheet without losing categories
      settingsSheet.insertRowsBefore(1, 4); // Insert 4 rows at the very top
      settingsSheet
        .getRange(1, 1, 1, 2)
        .setValues([["Setting", "Value"]])
        .setFontWeight("bold")
        .setBackground("#f3f3f3");
      settingsSheet
        .getRange(2, 1, 1, 2)
        .setValues([["Opening Balance (Jan 1st)", 0]]);
      settingsSheet.getRange(3, 1).setValue(""); // Spacer
      settingsSheet
        .getRange(4, 1, 1, 2)
        .setValues([["Categories", ""]])
        .setFontWeight("bold")
        .setBackground("#f3f3f3");

      // Clean up the old headers that got pushed down
      // We assume the old list was just a single column of categories
      // But let's be safe and just let the user clean up any double headers
    }
  }

  // 2. Setup Summary Sheet
  let summarySheet = ss.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Summary");
  }
  setupSummaryFormulas(summarySheet);

  // 3. Create Charts
  createDashboardCharts(summarySheet);
}

/**
 * Helper to find where categories start in Settings sheet.
 */
function sheetRowForCategoryStart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) return 4;

  const values = settingsSheet.getRange("A1:A10").getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim() === "Categories") return i + 1;
  }
  return 4;
}

/**
 * Sets up the live formulas in the Summary sheet.
 */
function setupSummaryFormulas(sheet) {
  sheet.clear();
  sheet.appendRow([
    "Month",
    "Total Income",
    "Total Expense",
    "Monthly Savings",
    "Closing Balance",
  ]);
  sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");

  const year = new Date().getFullYear();
  const mainSheetName = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()[0]
    .getName();

  // Reference to Opening Balance in Settings (Cell B2)
  const openingBalanceRef = "'Settings'!B2";

  for (let m = 0; m < 12; m++) {
    const monthDate = new Date(year, m, 1);
    const row = m + 2;
    sheet
      .getRange(row, 1)
      .setValue(
        Utilities.formatDate(
          monthDate,
          Session.getScriptTimeZone(),
          "MMMM yyyy",
        ),
      );

    // Income Formula
    sheet
      .getRange(row, 2)
      .setFormula(
        `=SUMIFS('${mainSheetName}'!B:B, '${mainSheetName}'!F:F, "Income", '${mainSheetName}'!A:A, ">="&DATE(${year}, ${m + 1}, 1), '${mainSheetName}'!A:A, "<"&DATE(${year}, ${m + 2}, 1))`,
      );

    // Expense Formula
    sheet
      .getRange(row, 3)
      .setFormula(
        `=SUMIFS('${mainSheetName}'!B:B, '${mainSheetName}'!F:F, "Expense", '${mainSheetName}'!A:A, ">="&DATE(${year}, ${m + 1}, 1), '${mainSheetName}'!A:A, "<"&DATE(${year}, ${m + 2}, 1))`,
      );

    // Monthly Savings Formula
    sheet.getRange(row, 4).setFormula(`=B${row}-C${row}`);

    // Closing Balance Formula
    if (row === 2) {
      // First month: Opening Balance + Savings
      sheet.getRange(row, 5).setFormula(`=${openingBalanceRef} + D${row}`);
    } else {
      // Future months: Previous Month's Balance + Savings
      sheet.getRange(row, 5).setFormula(`=E${row - 1} + D${row}`);
    }
  }

  sheet.getRange("B2:E13").setNumberFormat("#,##0.00");

  // Add Category breakdown table for Pie Chart
  sheet.getRange("G1").setValue("Expense by Category").setFontWeight("bold");
  sheet
    .getRange("G2")
    .setFormula(
      `=QUERY('${mainSheetName}'!B:F, "SELECT E, SUM(B) WHERE F = 'Expense' AND B IS NOT NULL GROUP BY E LABEL SUM(B) ''", 1)`,
    );
}

/**
 * Creates or updates charts on the Summary sheet.
 */
function createDashboardCharts(sheet) {
  const charts = sheet.getCharts();
  charts.forEach((chart) => sheet.removeChart(chart));

  // 1. Trend Chart (Income vs Expense)
  const trendChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("A1:C13"))
    .setPosition(15, 1, 0, 0)
    .setOption("title", "Monthly Income vs Expense")
    .setOption("width", 600)
    .setOption("height", 400)
    .build();
  sheet.insertChart(trendChart);

  // 2. Category Pie Chart
  // Data: G2:H (where QUERY puts the results)
  const categoryChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("G2:H20"))
    .setPosition(15, 7, 0, 0)
    .setOption("title", "Expenses by Category")
    .setOption("is3D", true)
    .setOption("width", 600)
    .setOption("height", 400)
    .build();
  sheet.insertChart(categoryChart);
}

/**
 * Applies data validation (dropdown) to the Category column in the Transactions sheet.
 */
function setupCategoryValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const settingsSheet = ss.getSheetByName("Settings");

  if (!settingsSheet) return;

  const startRow = sheetRowForCategoryStart() + 1;
  const categoryRange = settingsSheet.getRange(startRow, 1, 100, 1);
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(categoryRange)
    .setAllowInvalid(true)
    .build();

  mainSheet
    .getRange(2, 5, mainSheet.getMaxRows() - 1, 1)
    .setDataValidation(validationRule);
}
