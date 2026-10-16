# Unified Financial Automation Setup Guide - V4 (Enterprise Edition)

This comprehensive deployment guide provides the exact database structures, Apps Script automation engines, and system procedures to upgrade your spreadsheet into an **autonomous, self-categorizing, and self-allocating financial ecosystem**. 

This upgraded V4 system incorporates two newly engineered enterprise features:
1. **Dynamic Database-Driven Auto-Categorization (`Category_Rules_DB`)**: Removes hardcoded classification rules from code. You can now define, add, and adjust keyword-to-category rules directly inside your sheet interface.
2. **Double-Entry Monthly Savings Sweep Engine (`Savings_Goals`)**: Calculates your positive net surplus (Total Inflow - Total Outflow) at the close of any month, sweeps it across active savings goals based on custom percentages, programmatically updates balances, and appends double-entry auditable ledger rows to your ledger to maintain perfect accounting.

---

## 1. Setup of the New Supporting Sheets

When you run **`Generate Automated Sheet`** under V4, the compilation engine will programmatically configure and style two brand-new tabs in your workbook.

### A. The `Category_Rules_DB` Tab
This sheet serves as a **Single Source of Truth** for transaction mapping, decoupling your data categorizer from the Apps Script code.
* **Headers**: `Keyword` (Column A), `Category` (Column B)
* **Data Validation**: Column B is strictly bound to your categories in `Enum!A3:A11` via drop-down range validation.
* **Default Seeding**: Seeded with over 70 retail, food, travel, and utility keyword mappings.
* **Customization**: To add a new rule (e.g., mapping any description containing "Equinox" to "Personal"), simply write `equinox` in Column A and choose `Personal` in Column B. No code edits are required!

### B. The `Savings_Goals` Tab
This portal operates as your strategic financial progress tracker.
* **Headers**: 
  * `Goal Name` (Column A): Descriptive tag (e.g., "Emergency Fund", "Travel Account").
  * `Target Amount` (Column B): Numeric goal limit ($).
  * `Current Balance` (Column C): Incremented programmatically by sweeps or manual edits.
  * `Target Date` (Column D): Scheduled deadline (YYYY-MM-DD).
  * `Sweep Percentage` (Column E): Set percentage of your monthly net surplus allocated to this goal (e.g., `50.0%`).
  * `Progress Meter` (Column F): programmatically generated visual progress sparklines (`=SPARKLINE()`).
* **Relational Rules**: The sum of all active allocations in Column E must equal **100% or less**. The script will automatically intercept and block the sweep if your allocations exceed 100%.

---

## 2. Dynamic Database-Driven Auto-Categorization

The classification routine `autoCategorizeMerchant()` has been completely upgraded:
* **The Process**: Every time a transaction is recorded via your Gmail scraper, Chase CSV statement ingester, or manual entry HTML form, the script triggers.
* **Keyword Matching**: It queries your `Category_Rules_DB` sheet, matches the raw description against active keywords (case-insensitive), and returns the mapped category.
* **Resilient Fallback**: If the rules sheet is unavailable, or a keyword match isn't found, it gracefully falls back to a robust built-in hardcoded ruleset, and defaults to "Miscellaneous" only if no keyword matches exist.
* **Historical Cleanup**: Running **`Run AI Auto-Categorization`** from your custom toolbar menu scans your entire database and automatically backfills any "Miscellaneous" or blank cells with matching categories based on your live rule database.

---

## 3. The Double-Entry Monthly Savings Sweep Engine

To satisfy **Relational Database Normalization** and **Double-Entry Bookkeeping Principles**, we do not simply overwrite numbers when moving money to savings. Instead, we log a complete audit trail:

### How the Savings Sweep Engine Operates:
1. **Surplus Calculation**: Reads the dynamic Monthly Dashboard tab KPI cells (`A7` and `B7`) to determine your actual `Net Savings` (Net Income - Outflows) for your selected month.
2. **Goal Selection**: Queries the `Savings_Goals` tab to pull active goals, target criteria, and sweep percentages.
3. **Double-Entry Ingestion**: For each swept goal, the engine automatically appends two corresponding rows to your `Transactions_DB` ledger:
   * **Debit Row (Outflow)**: Logs the sweep amount as an `OUTFLOW` from your standard transactional account (e.g., `Chase Checking`) under category `Savings`.
   * **Credit Row (Inflow)**: Logs the sweep amount as an `INFLOW` to your savings target under account `Savings` and category `Savings`.
4. **Balance Aggregation**: Adds the swept amount directly to the goal's `Current Balance` cell in the `Savings_Goals` sheet, updating the visual progress bar instantly!
5. **Deduplication Safeguards**: Employs Google's `PropertiesService` state tracking. If you attempt to sweep the same month twice, a modal pop-up warning will alert you to prevent double-counting.

---

## 4. Full Script Integration Walkthrough

To apply these additions to your spreadsheet setup:

1. **Clear Legacy Code**: Open **Extensions > Apps Script** in your sheet. Overwrite all contents in `main.gs` with the code in your newly generated **`unified-financial-automation-v4.gs`** file.
2. **Verify HTML Sidebar File**: Ensure your sidebar file is named exactly `SidebarView` inside the Apps Script HTML tab list. 
 > `SidebarView.html`file was added in `unified-financial-system-setup.md` `step 5. *UI View File: SidebarView.html*`

3. **Deploy the Web App**: If you have not done so already, click **Deploy > New deployment**, select **Web app**, set execute as **Me**, set access to **Anyone**, and click **Deploy**. Copy this URL and paste it as the `const WEB_APP_URL` inside your **`ManualEntry.html`** file.
4. **Initialize Workbook**: Open your spreadsheet and click **`Financial Automation 🚀` > `Generate Automated Sheet`**.
   * The script will programmatically construct, style, and populate all eight operational tabs: `About`, `Enum`, `Transactions_DB`, `Budget_History_DB`, `The Monthly Dashboard Tab`, `Control`, `Category_Rules_DB`, and `Savings_Goals`.
   * Open the clickable URL popped open on your screen to view your completed workbook!

---

## 5. Testing the Completed System

### Test 1: Manual Transaction Ingestion & Rules-Based Mapping
1. Open your `ManualEntry.html` file in your browser, enter a transaction for `Target Store`, set amount to `45.00`, Category to `Miscellaneous` (or leave empty), and click **Submit**.
2. Open your `Transactions_DB` sheet—you will find the row appended. Because `target` matches our `Category_Rules_DB` list, the category will be automatically written as `Home` instead of "Miscellaneous"!

### Test 2: Execute Monthly Savings Sweep Closeout
1. Head to your `Transactions_DB` and insert a dummy income row: Date: `2026-08-25`, Merchant: `Payroll Paycheck`, Amount: `5000.00`, Flow: `INFLOW`, Category: `Miscellaneous` (which rules maps to Miscellaneous), Account: `Chase Checking`, Channel: `MANUAL`, Status: `SETTLED`.
2. Add a dummy expense row: Merchant: `Electric Utility`, Amount: `150.00`, Flow: `OUTFLOW`, Category: `Miscellaneous` (which rules dynamically categorize as `Bills & Utilities`!), Account: `Chase Checking`.
3. Open **`The Monthly Dashboard Tab`**, select **August 2026**. Your Total Income will show `$5,000.00`, Total Spending `$150.00`, and Net Savings of **`$4,850.00`**.
4. Click **`Financial Automation 🚀` > `Execute Monthly Savings Sweep 💰`**. 
5. Confirm the pop-up prompt. The script will calculate allocations, update balances on `Savings_Goals` instantly, and append the dual-debit-credit ledger rows to `Transactions_DB` to keep your system perfectly balanced under institutional accounting principles!

