/**
 * Main function to sync Nayapay transactions from Gmail to Google Sheets.
 * @param {string|null} specificAfterDate Optional date to sync from.
 */
function syncNayapayTransactions(specificAfterDate = null) {
  console.log("Starting Nayapay sync...");

  // 1. Initialize structural sheets (Summary & Settings)
  initializeSmartSheets();

  // 2. Detect sync strategy
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const isSheetEmpty = mainSheet.getLastRow() <= 1;
  const isSpecificSync = specificAfterDate !== null;

  // 3. Fetch and parse emails
  const messages = fetchNayapayEmails(isSheetEmpty, specificAfterDate);
  console.log(`Fetched ${messages.length} messages from Gmail.`);

  const parsedData = [];

  messages.forEach((message) => {
    try {
      const data = parseNayapayEmail(message);
      if (data) {
        parsedData.push(data);
      }
    } catch (e) {
      console.error("Error parsing message: " + message.getSubject(), e);
    }
  });

  if (parsedData.length > 0) {
    parsedData.sort((a, b) => a.date - b.date);
    updateSheetWithTransactions(parsedData);
    console.log(`Successfully synced ${parsedData.length} transactions.`);
  } else {
    console.log("No new transactions found.");
  }

  // 4. Update last sync timestamp (Only if it's a regular sync, not a manual recalibration)
  if (!isSpecificSync) {
    PropertiesService.getScriptProperties().setProperty(
      "LAST_SYNC_TIME",
      new Date().getTime().toString(),
    );
  }
}

/**
 * Triggers a redo of the sync for a specific date or full year.
 */
function forceRecalibrate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Recalibrate Sync",
    "Enter a start date (YYYY/MM/DD) or leave empty for a full year sync:",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const inputDate = response.getResponseText().trim();
    if (inputDate === "") {
      // Full recalibration
      resetSyncTime();
      syncNayapayTransactions();
      ui.alert(
        "Recalibration Complete",
        "The sheet has been fully re-synced.",
        ui.ButtonSet.OK,
      );
    } else {
      // Date-specific recalibration
      // Basic regex check for YYYY/MM/DD
      if (/^\d{4}\/\d{2}\/\d{2}$/.test(inputDate)) {
        syncNayapayTransactions(inputDate);
        ui.alert(
          "Sync Complete",
          `Successfully re-synced from ${inputDate}.`,
          ui.ButtonSet.OK,
        );
      } else {
        ui.alert(
          "Invalid Date",
          "Please use the format YYYY/MM/DD (e.g. 2026/01/25)",
          ui.ButtonSet.OK,
        );
      }
    }
  }
}

/**
 * Creates a custom menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("NayaPay")
    .addItem("Sync Now", "syncNayapayTransactions")
    .addItem("Recalibrate (Select Date)", "forceRecalibrate")
    .addSeparator()
    .addItem("Setup Automation", "createTrigger")
    .addToUi();
}

/**
 * Creates a time-driven trigger to run the sync every minute.
 */
function createTrigger() {
  initializeSmartSheets();

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("syncNayapayTransactions")
    .timeBased()
    .everyMinutes(1)
    .create();
}
