/**
 * Parses Nayapay transaction emails into structured data.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The email message to parse.
 * @return {Object|null} Structured transaction data or null if not a transaction.
 */
function parseNayapayEmail(message) {
  const plainBody = message.getPlainBody();
  const htmlBody = message.getBody();
  const subject = message.getSubject();

  // 1. Ignore non-financial or precursor emails
  if (
    subject.match(/OTP|bill is due|got a new bill|failed international|failed/i)
  ) {
    return null;
  }

  // Clean HTML: Remove style and script tags before stripping other tags
  const cleanHtml = htmlBody
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]*>/g, " ");

  const searchContent = plainBody + "\n" + cleanHtml;

  const data = {
    date: message.getDate(),
    amount: 0,
    merchant: "Unknown",
    type: "Expense", // Default to Expense
    referenceId: "",
    note: "",
  };

  // 2. Extract Transaction ID
  // Pattern 1: Look for labels followed by alphanumeric ID
  let txIdMatch = searchContent.match(
    /(?:Transaction ID|Reference ID|RRN|Ref ID)\s*[:\-]?\s*([a-z0-9]{4,})/i,
  );

  // Pattern 2: Fallback to just finding a 24-char hex string
  if (!txIdMatch) {
    txIdMatch = searchContent.match(/\b([a-f0-9]{24})\b/i);
  }

  if (txIdMatch) {
    data.referenceId = txIdMatch[1];
  } else if (subject.match(/sent|got|spent|withdrawal|remittance/i)) {
    data.referenceId = "GMAIL_" + message.getId();
  } else {
    return null;
  }

  // 3. Extract Amount
  let amountMatch = searchContent.match(/Rs\.\s*([\d,]+\.?\d*)/i);
  if (!amountMatch) {
    amountMatch = subject.match(/Rs\.\s*([\d,]+\.?\d*)/i);
  }

  if (amountMatch) {
    data.amount = parseFloat(amountMatch[1].replace(/,/g, ""));
  }

  // 4. Determine Type and Merchant (Strictly Income or Expense)
  if (
    subject.includes("You sent") ||
    subject.includes("You spent") ||
    subject.includes("withdrawal")
  ) {
    data.type = "Expense";
    const merchantMatch = subject.match(/to (.*) 💸|at (.*) 💳|at (.*) 📱/i);
    data.merchant = merchantMatch
      ? (merchantMatch[1] || merchantMatch[2] || merchantMatch[3]).trim()
      : "Expenditure";

    if (subject.includes("withdrawal")) {
      data.merchant = "ATM Cash Withdrawal";
    }
  } else if (
    subject.includes("You got") ||
    subject.includes("Inward Foreign Remittance")
  ) {
    data.type = "Income";
    const merchantMatch = subject.match(/from (.*) 🎉|from (.*) to/i);
    data.merchant = merchantMatch
      ? (merchantMatch[1] || merchantMatch[2]).trim()
      : "Incoming Funds";

    if (data.merchant === "Incoming Funds") {
      const fromMatch = searchContent.match(/from (.*) to your NayaPay/i);
      if (fromMatch) data.merchant = fromMatch[1].trim();
    }

    if (subject.includes("Foreign Remittance")) {
      const sourceMatch = searchContent.match(/Source Bank\s*[:\-]?\s*(.*)/i);
      data.merchant = sourceMatch
        ? sourceMatch[1].trim().split(/\s{2,}/)[0]
        : "Foreign Remittance";
    }
  } else {
    // If we can't determine type but it's clearly a transaction, guess based on keywords
    data.type = subject.match(/received|got|inward/i) ? "Income" : "Expense";
    data.merchant = "Nayapay Transaction";
  }

  return data;
}
