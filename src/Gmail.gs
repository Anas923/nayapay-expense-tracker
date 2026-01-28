/**
 * Fetches Nayapay transaction emails with a safety buffer.
 * @param {boolean} forceHistory If true, fetches all emails since the start of the current year.
 * @param {string} specificAfterDate Optional date (YYYY/MM/DD) to fetch after.
 */
function fetchNayapayEmails(forceHistory = false, specificAfterDate = null) {
  const props = PropertiesService.getScriptProperties();
  const lastSync = props.getProperty("LAST_SYNC_TIME");

  // Base query
  let query = "from:service@nayapay.com";

  if (specificAfterDate) {
    // Priority 1: User specified a date
    query += ` after:${specificAfterDate}`;
    console.log(
      `Manual recalibration requested. Searching since ${specificAfterDate}.`,
    );
  } else if (lastSync && !forceHistory) {
    // Priority 2: Optimize with Safety buffer (normal sync)
    const bufferedTime = Math.floor(lastSync / 1000) - 600;
    query += ` after:${bufferedTime}`;
  } else {
    // Priority 3: Full year history (first run or forced full scan)
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
