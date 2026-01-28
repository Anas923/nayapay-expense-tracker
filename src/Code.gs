/**
 * Main function to sync Nayapay transactions from Gmail to Google Sheets.
 */
function syncNayapayTransactions() {
  console.log("Starting Nayapay sync...");

  // 1. Initialize structural sheets (Summary & Settings)
  initializeSmartSheets();

  // 2. Detect if we need a full history sync (is the transaction sheet empty?)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const isSheetEmpty = mainSheet.getLastRow() <= 1;

  // 3. Fetch and parse emails
  const messages = fetchNayapayEmails(isSheetEmpty);
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

  // 4. Update last sync timestamp
  PropertiesService.getScriptProperties().setProperty(
    "LAST_SYNC_TIME",
    new Date().getTime().toString(),
  );
}

/**
 * Triggers a full redo of the sync for the current year.
 * Useful if a transaction was missed or deleted.
 */
function forceRecalibrate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Recalibrate Sync",
    "This will re-scan all Nayapay emails from Jan 1st of this year. Existing transactions will not be duplicated. Proceed?",
    ui.ButtonSet.YES_NO,
  );

  if (response === ui.Button.YES) {
    resetSyncTime();
    syncNayapayTransactions();
    ui.alert(
      "Recalibration Complete",
      "The sheet has been fully re-synced.",
      ui.ButtonSet.OK,
    );
  }
}

/**
 * Creates a custom menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("NayaPay")
    .addItem("Sync Now", "syncNayapayTransactions")
    .addItem("Recalibrate (Full Year)", "forceRecalibrate")
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
