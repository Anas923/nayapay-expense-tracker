# Nayapay Expense Tracker 🛡️📈

An automated financial tracking system that syncs **NayaPay** transaction emails from Gmail into **Google Sheets** in real-time.

![Automation Process](https://img.shields.io/badge/Maintained%3F-yes-green.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Platform: Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-blue.svg)

## 🚀 Key Features
- **⚡ Real-time Sync**: Automatically checks your Gmail for new transactions every minute.
- **📊 Interactive Dashboard**: Visualizes your monthly spending trends and category breakdowns with live charts.
- **🧠 Intelligent Classification**: Automatically detects Income/Expense types and suggests categories based on merchants.
- **💰 Balance Reconciliation**: Includes an "Opening Balance" feature to ensure your sheet matches your actual bank balance 1:1.
- **🔄 Deduplication**: Robust Reference ID tracking prevents duplicate entries, even if the script runs multiple times.

## 🛠️ Setup Instructions

### For Developers (clasp)
1. Clone the repository.
2. Run `clasp login` and `clasp push`.
3. Open the sheet and run `syncNayapayTransactions`.

### For General Users (Copy-Paste)
1. Create a new Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Copy the files from the `src/` directory into the editor.
4. Run `syncNayapayTransactions` and authorize the permissions.
5. Run `createTrigger` to enable 1-minute auto-sync.

## 📊 Sheet Structure
- **Transactions**: The raw log of all your spending and earnings.
- **Settings**: Manage your custom categories and set your Jan 1st Opening Balance.
- **Summary**: Live monthly breakdown with automated charts.

## 🤝 Contributing
Feel free to fork this project and submit pull requests for more "Auto-Category" keywords or advanced dashboard features!

## 📜 License
Published under the MIT License. Feel free to share and modify!

---
*Created to simplify financial tracking for NayaPay users.*
