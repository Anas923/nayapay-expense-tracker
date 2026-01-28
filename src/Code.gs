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
  const isSheetEmpty = mainSheet.getLastRow() <= 1; // 0 or 1 (headers only)

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
    // Sort by date ascending to append in order
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
