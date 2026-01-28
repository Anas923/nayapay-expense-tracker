/**
 * Fetches Nayapay transaction emails.
 * @param {boolean} forceHistory If true, fetches all emails since the start of the current year.
 */
function fetchNayapayEmails(forceHistory = false) {
  const props = PropertiesService.getScriptProperties();
  const lastSync = props.getProperty("LAST_SYNC_TIME");

  // Base query
  let query = "from:service@nayapay.com";

  if (lastSync && !forceHistory) {
    // Optimization: Only fetch since last sync (Gmail search 'after' uses seconds)
    query += ` after:${Math.floor(lastSync / 1000)}`;
  } else {
    // First run or empty sheet: fetch for the entire current year
    const year = new Date().getFullYear();
    query += ` after:${year}/01/01`;
    console.log(
      `Initial/Full sync requested. Searching since Jan 1st, ${year}.`,
    );
  }

  console.log(`Searching Gmail with query: ${query}`);
  const threads = GmailApp.search(query, 0, 100); // 100 threads is plenty for a yearly history
  const messages = [];

  threads.forEach((thread) => {
    messages.push(...thread.getMessages());
  });

  return messages;
}
