/**
 * Fetches Nayapay transaction emails with a safety buffer.
 * @param {boolean} forceHistory If true, fetches all emails since the start of the current year.
 */
function fetchNayapayEmails(forceHistory = false) {
  const props = PropertiesService.getScriptProperties();
  const lastSync = props.getProperty("LAST_SYNC_TIME");

  // Base query
  let query = "from:service@nayapay.com";

  if (lastSync && !forceHistory) {
    // Optimization with Safety: Subtract 10 minutes (600 seconds)
    // to handle overlapping messages or Gmail indexing delays.
    const bufferedTime = Math.floor(lastSync / 1000) - 600;
    query += ` after:${bufferedTime}`;
  } else {
    // First run or empty sheet: fetch for the entire current year
    const year = new Date().getFullYear();
    query += ` after:${year}/01/01`;
    console.log(
      `Initial/Full sync requested. Searching since Jan 1st, ${year}.`,
    );
  }

  console.log(`Searching Gmail with query: ${query}`);
  const threads = GmailApp.search(query, 0, 100);
  const messages = [];

  threads.forEach((thread) => {
    messages.push(...thread.getMessages());
  });

  return messages;
}

/**
 * Resets the sync timestamp to force a full re-scan on next run.
 */
function resetSyncTime() {
  PropertiesService.getScriptProperties().deleteProperty("LAST_SYNC_TIME");
  console.log("Sync time reset successful.");
}
