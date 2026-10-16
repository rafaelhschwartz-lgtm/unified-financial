# Unified Financial Automation Setup Guide - V11 (Master Enterprise Edition) (Master Edition)

This comprehensive setup manual guides you through deploying, testing, and expanding your **Enterprise-Grade, Self-Expanding Personal Finance Ledger System**. The architecture complies strictly with the **FAST Modeling Standard** (Flexibility, Appropriateness, Structure, Transparency) and the **ICAEW Spreadsheet Design Principles**.

---

## 1. V11 Master Enhancements

### A. Auto-Adding Accounts to `Enum` (`ensureAccountInEnum`)
* **The Capability:** Whenever an account is detected from a CSV filename, CSV row header, Gmail purchase notification, Google Voice SMS forward, or manual entry (e.g. `Chase Checking 5678`, `Amex Gold 4321`), the script checks `Enum!B3:B` (`Account Sources`).
* **The Auto-Expansion:** If the account is not listed, the script automatically appends it to Column B on the `Enum` tab and updates data validation dropdowns across `Transactions_DB` in real-time. You never have to manually edit dropdown ranges when opening new bank accounts!

### B. Auto-Adding Categories to `Enum` (`ensureCategoryInEnum`)
* **The Capability:** When a new category is specified or inferred by auto-categorization rules, `ensureCategoryInEnum()` checks `Enum!A3:A` (`Budget Categories`).
* **The Auto-Expansion:** If missing, it appends the category to `Enum!A:A`, creates a default `$0.00` historical budget entry in `Budget_History_DB`, and refreshes category dropdown validations on `Transactions_DB` and `Category_Rules_DB`.

### C. Auto-Learning Categorization Rules (`Category_Rules_DB`)
* **The Capability:** When transactions arrive with unlisted merchants (e.g., `TRADER JOE'S #521`), the script extracts keywords (`trader joe`).
* **The Auto-Expansion:** If no matching rule exists in `Category_Rules_DB`, the engine infers the category (`Groceries`), cleans the title (`Trader Joe's`), and **automatically appends a new rule row `[keyword, Clean Name, Category]` directly into `Category_Rules_DB`**! Future transactions from that merchant will use this rule automatically without any code editing.

### D. Multi-Account Smart Detection Engine
* **The Capability:** Handles multiple bank accounts of the same type (e.g., two checking accounts or multiple credit cards).
* **The Detection:** The parser evaluates account numbers or card nicknames from:
  1. Filenames in `Chase_CSV_Uploads` (e.g., `checking_1234.csv`, `sapphire_9012.csv`).
  2. CSV statement columns (`Account Number`).
  3. Email bodies and SMS text snippets (`card ending in 5678`, `acct #9012`).
  4. Automatically maps the transaction to `Chase Checking 1234`, `Chase Sapphire 9012`, etc., and ensures the account exists in `Enum`.

### E. "➕ Quick Add Tools" Toolbar Menu
* **The Capability:** Added a dedicated sub-menu under **`Financial Automation 🚀` > `➕ Quick Add Tools`**:
  * **`➕ Quick Add Transaction`**: Modal prompt to log a date, merchant, amount, category, and account in seconds.
  * **`➕ Add New Budget Category`**: Prompts for a category and registers it in `Enum`.
  * **`➕ Add New Account Source`**: Prompts for an account name and registers it in `Enum`.
  * **`➕ Add Merchant Category Rule`**: Prompts for keyword, clean name, and target category to populate `Category_Rules_DB`.

---

## 2. Step-by-Step Deployment Instructions

### Step 1: Overwrite Master Script Code
1. Open your Google Sheet and navigate to **Extensions > Apps Script**.
2. Delete any legacy `.gs` files (like `Setup.gs` or old `main.gs` blocks) to prevent trigger collisions.
3. Open your primary script file (`main.gs`), delete its contents, and paste the complete code from **`unified-financial-automation-v10.gs`**.
4. Click the **Save (floppy disk icon)**.

### Step 2: Deploy HTML Sidebar Files
1. In the Apps Script left-hand sidebar, click **`+` > HTML**. Name the file exactly **`SidebarView`** (omit `.html`). Paste the code from Section 4 below and save.
2. Click **`+` > HTML**. Name the file exactly **`TransactionEditorView`** (omit `.html`). Paste the code from Section 5 below and save.

### Step 3: Initialize Workbook Tabs In-Place
1. Return to your Google Sheet browser tab and refresh the page.
2. Select **`Financial Automation 🚀` > `Initialize Sheets in This Workbook`**.
3. Confirm the prompt and grant authorization. All operational sheets (**`About`**, **`Enum`**, **`Transactions_DB`**, **`Budget_History_DB`**, **`The Monthly Dashboard Tab`**, **`Control`**, **`Category_Rules_DB`**, **`Savings_Goals`**) will build in-place with pre-linked formulas and dynamic charts!

---

## 3. Scheduled Background Automation Triggers

To run your ledger 100% hands-off, set up these three time-driven triggers in Apps Script (**Triggers / Alarm Clock icon**):

1. **`pullGmailTransactionAlerts`**:
   * Event source: `Time-driven` | Type: `Hour timer` | Interval: `Every hour`.
   * Automatically parses Chase purchase emails and Google Voice transaction text messages, auto-detects account digits, learns new rules, and logs deduplicated entries.
2. **`checkBudgetBreaches`**:
   * Event source: `Time-driven` | Type: `Day timer` | Time: `8 AM to 9 AM`.
   * Evaluates actual category spending against planned monthly boundaries. Sends an email alert if a breach occurs (with state-tracking to limit to 1 alert per category per month).
3. **`sendWeeklyBudgetSummary`**:
   * Event source: `Time-driven` | Type: `Weekly timer` | Day: `Every Monday 7 AM to 8 AM`.
   * Dispatches a mobile-responsive HTML briefing newsletter summarizing month-to-date KPIs, spending velocity, and category statuses.

---

## 4. SidebarView.html (Historical Budget Allocator)

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 15px; font-family: 'Arial', sans-serif; background-color: #fafafa; }
      .header-title { font-weight: bold; color: #1f4e78; margin-bottom: 20px; font-size: 1.2rem; display: flex; align-items: center; }
      .header-title i { margin-right: 8px; }
      .category-row { margin-bottom: 12px; background: white; padding: 10px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
      .category-label { font-weight: bold; color: #333; font-size: 0.95rem; }
      .save-btn { width: 100%; background-color: #1f4e78 !important; margin-top: 15px; }
      .loader { margin: 20px auto; text-align: center; }
      .period-indicator { background: #e8eff7; padding: 8px; border-radius: 4px; margin-bottom: 15px; font-size: 0.9rem; color: #1f4e78; font-weight: bold; text-align: center; }
    </style>
  </head>
  <body>
    <div class="header-title">
      <i class="material-icons">account_balance_wallet</i>
      <span>Budget Ceilings</span>
    </div>
    
    <div id="periodIndicator" class="period-indicator">Loading active period...</div>
    <div id="loading" class="loader">
      <div class="preloader-wrapper small active">
        <div class="spinner-layer spinner-blue-only">
          <div class="circle-clipper left"><div class="circle"></div></div>
          <div class="gap-patch"><div class="circle"></div></div>
          <div class="circle-clipper right"><div class="circle"></div></div>
        </div>
      </div>
    </div>
    
    <form id="budgetForm" style="display:none;">
      <div id="categoryContainer"></div>
      <button class="btn waves-effect waves-light save-btn" type="submit" name="action">
        Save Budget Targets
        <i class="material-icons right">save</i>
      </button>
    </form>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script>
      let activeYear = "";
      let activeMonth = "";

      document.addEventListener('DOMContentLoaded', function() {
        google.script.run.withSuccessHandler(initializePeriod).getDashboardPeriod();
      });

      function initializePeriod(period) {
        activeYear = period.year;
        activeMonth = period.month;
        document.getElementById("periodIndicator").innerText = "Period: " + activeMonth + " " + activeYear;
        google.script.run.withSuccessHandler(renderCategories).getCategoryBudgets(activeYear, activeMonth);
      }

      function renderCategories(data) {
        document.getElementById("loading").style.display = "none";
        document.getElementById("budgetForm").style.display = "block";
        const container = document.getElementById("categoryContainer");
        container.innerHTML = "";
        
        data.forEach(item => {
          const div = document.createElement("div");
          div.className = "category-row";
          div.innerHTML = `
            <div class="category-label">${item.category}</div>
            <div class="input-field inline" style="margin:0; width:100%;">
              <i class="material-icons prefix" style="font-size:1.3rem; margin-top:8px;">attach_money</i>
              <input type="number" step="0.01" min="0" class="validate budget-input" data-category="${item.category}" value="${item.amount.toFixed(2)}" style="height:2rem; margin:0;">
            </div>
          `;
          container.appendChild(div);
        });
      }

      document.getElementById("budgetForm").addEventListener("submit", function(e) {
        e.preventDefault();
        const inputs = document.querySelectorAll(".budget-input");
        const budgetData = [];
        
        inputs.forEach(input => {
          budgetData.push({
            category: input.getAttribute("data-category"),
            amount: parseFloat(input.value) || 0
          });
        });
        
        M.toast({html: 'Syncing to Database... ⏳', classes: 'blue'});
        google.script.run.withSuccessHandler(onSaveSuccess).saveCategoryBudgets(activeYear, activeMonth, budgetData);
      });

      function onSaveSuccess(response) {
        if (response === "SUCCESS") {
          M.toast({html: 'Budgets saved successfully! 🎉', classes: 'green'});
        } else {
          M.toast({html: 'Save failed ❌', classes: 'red'});
        }
      }
    </script>
  </body>
</html>
```

---

## 5. TransactionEditorView.html (Paginated Transaction Ledger Editor)

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 12px; font-family: 'Arial', sans-serif; background-color: #f9fbfd; font-size: 0.9rem; }
      .header-title { font-weight: bold; color: #1f4e78; margin-bottom: 12px; font-size: 1.1rem; display: flex; align-items: center; justify-content: space-between; }
      .card-panel { padding: 12px; margin-bottom: 12px; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
      .input-field { margin-top: 5px; margin-bottom: 5px; }
      .input-field input, .input-field select { height: 2.2rem; font-size: 0.85rem; }
      .btn-small { background-color: #1f4e78 !important; padding: 0 12px; height: 28px; line-height: 28px; font-size: 0.8rem; }
      .btn-danger { background-color: #db4437 !important; }
      .btn-secondary { background-color: #757575 !important; }
      .tx-item { padding: 8px; border-bottom: 1px solid #e0e0e0; cursor: pointer; transition: background 0.2s; }
      .tx-item:hover { background-color: #e8eff7; }
      .tx-merchant { font-weight: bold; color: #333; }
      .tx-meta { font-size: 0.75rem; color: #666; display: flex; justify-content: space-between; margin-top: 2px; }
      .tx-amount-out { color: #db4437; font-weight: bold; }
      .tx-amount-in { color: #2e7d32; font-weight: bold; }
      .pagination-controls { display: flex; align-items: center; justify-content: space-between; margin-top: 8px; font-size: 0.8rem; }
      .loader { text-align: center; margin: 15px 0; }
    </style>
  </head>
  <body>
    <div class="header-title">
      <span><i class="material-icons left" style="font-size:1.2rem;">edit_note</i>Ledger Editor</span>
      <button class="btn-small waves-effect waves-light" onclick="clearForm()">➕ New</button>
    </div>

    <!-- FORM PANEL -->
    <div class="card-panel">
      <input type="hidden" id="txRowIdx" value="">
      
      <div class="row" style="margin-bottom:0;">
        <div class="input-field col s6">
          <input type="date" id="txDate">
        </div>
        <div class="input-field col s6">
          <input type="number" id="txAmount" step="0.01" placeholder="Amount ($)">
        </div>
      </div>

      <div class="input-field" style="margin-top:0;">
        <input type="text" id="txMerchant" placeholder="Merchant / Payee Name">
      </div>

      <div class="row" style="margin-bottom:0;">
        <div class="input-field col s6">
          <select id="txFlow" class="browser-default">
            <option value="OUTFLOW">OUTFLOW</option>
            <option value="INFLOW">INFLOW</option>
            <option value="TRANSFER">TRANSFER</option>
          </select>
        </div>
        <div class="input-field col s6">
          <select id="txCategory" class="browser-default">
            <option value="">Category...</option>
          </select>
        </div>
      </div>

      <div class="input-field" style="margin-top:5px;">
        <select id="txAccount" class="browser-default">
          <option value="">Account Source...</option>
        </select>
      </div>

      <div style="display:flex; justify-content:space-between; margin-top:10px;">
        <button class="btn-small waves-effect waves-light" onclick="saveTransaction()">Save</button>
        <button class="btn-small btn-secondary waves-effect" onclick="clearForm()">Clear</button>
        <button class="btn-small btn-danger waves-effect" id="btnDelete" style="display:none;" onclick="deleteTransaction()">Delete</button>
      </div>
    </div>

    <!-- SEARCH & TRANSACTIONS LIST -->
    <div class="input-field" style="margin-bottom:8px;">
      <input type="text" id="searchInput" placeholder="🔍 Search payee, category, account..." onkeyup="debounceSearch()">
    </div>

    <div id="loader" class="loader" style="display:none;">
      <div class="preloader-wrapper micro active">
        <div class="spinner-layer spinner-blue-only">
          <div class="circle-clipper left"><div class="circle"></div></div>
        </div>
      </div>
    </div>

    <div id="txList"></div>

    <div class="pagination-controls">
      <button class="btn-small waves-effect btn-secondary" id="btnPrev" onclick="changePage(-1)">Prev</button>
      <span id="pageInfo">Page 1 of 1</span>
      <button class="btn-small waves-effect btn-secondary" id="btnNext" onclick="changePage(1)">Next</button>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script>
      let currentPage = 1;
      let totalPages = 1;
      let searchTimeout = null;

      document.addEventListener('DOMContentLoaded', function() {
        loadEnumDropdowns();
        loadTransactions(1);
        document.getElementById("txDate").value = new Date().toISOString().split('T')[0];
      });

      function loadEnumDropdowns() {
        google.script.run.withSuccessHandler(function(enums) {
          const catSelect = document.getElementById("txCategory");
          const acctSelect = document.getElementById("txAccount");
          
          catSelect.innerHTML = '<option value="">Category...</option>';
          enums.categories.forEach(c => {
            catSelect.innerHTML += `<option value="${c}">${c}</option>`;
          });
          
          acctSelect.innerHTML = '<option value="">Account Source...</option>';
          enums.accounts.forEach(a => {
            acctSelect.innerHTML += `<option value="${a}">${a}</option>`;
          });
        }).getEnumValues();
      }

      function loadTransactions(page) {
        document.getElementById("loader").style.display = "block";
        const query = document.getElementById("searchInput").value;
        google.script.run.withSuccessHandler(renderTransactions).getTransactionsPage(page, 8, query);
      }

      function renderTransactions(res) {
        document.getElementById("loader").style.display = "none";
        currentPage = res.currentPage;
        totalPages = res.totalPages;
        document.getElementById("pageInfo").innerText = `Page ${currentPage} of ${totalPages} (${res.totalRecords})`;
        
        document.getElementById("btnPrev").disabled = (currentPage <= 1);
        document.getElementById("btnNext").disabled = (currentPage >= totalPages);
        
        const container = document.getElementById("txList");
        container.innerHTML = "";
        
        if (!res.data || res.data.length === 0) {
          container.innerHTML = '<div style="text-align:center; color:#888; padding:15px;">No transactions found.</div>';
          return;
        }

        res.data.forEach(item => {
          const div = document.createElement("div");
          div.className = "tx-item";
          const amountClass = item.flow === "INFLOW" ? "tx-amount-in" : "tx-amount-out";
          const amountPrefix = item.flow === "INFLOW" ? "+" : "-";
          
          div.innerHTML = `
            <div class="tx-merchant">${item.merchantClean || item.merchantRaw}</div>
            <div class="tx-meta">
              <span>${item.date} &bull; ${item.category} (${item.accountSource})</span>
              <span class="${amountClass}">${amountPrefix}$${item.amount.toFixed(2)}</span>
            </div>
          `;
          
          div.onclick = function() {
            populateForm(item);
          };
          container.appendChild(div);
        });
      }

      function populateForm(item) {
        document.getElementById("txRowIdx").value = item.rowIdx;
        document.getElementById("txDate").value = item.date;
        document.getElementById("txAmount").value = item.amount;
        document.getElementById("txMerchant").value = item.merchantRaw;
        document.getElementById("txFlow").value = item.flow;
        document.getElementById("txCategory").value = item.category;
        document.getElementById("txAccount").value = item.accountSource;
        document.getElementById("btnDelete").style.display = "inline-block";
      }

      function clearForm() {
        document.getElementById("txRowIdx").value = "";
        document.getElementById("txDate").value = new Date().toISOString().split('T')[0];
        document.getElementById("txAmount").value = "";
        document.getElementById("txMerchant").value = "";
        document.getElementById("txFlow").value = "OUTFLOW";
        document.getElementById("txCategory").value = "";
        document.getElementById("txAccount").value = "";
        document.getElementById("btnDelete").style.display = "none";
      }

      function saveTransaction() {
        const rowIdx = document.getElementById("txRowIdx").value;
        const txData = {
          rowIdx: rowIdx,
          date: document.getElementById("txDate").value,
          amount: parseFloat(document.getElementById("txAmount").value) || 0,
          merchantRaw: document.getElementById("txMerchant").value,
          flow: document.getElementById("txFlow").value,
          category: document.getElementById("txCategory").value,
          accountSource: document.getElementById("txAccount").value,
          channel: "MANUAL",
          status: "SETTLED"
        };

        if (!txData.merchantRaw || txData.amount <= 0) {
          M.toast({html: 'Please enter merchant and amount > 0 ⚠️', classes: 'orange'});
          return;
        }

        M.toast({html: 'Saving transaction... ⏳', classes: 'blue'});
        google.script.run.withSuccessHandler(function(res) {
          M.toast({html: `Transaction ${res.toLowerCase()}! 🎉`, classes: 'green'});
          clearForm();
          loadEnumDropdowns();
          loadTransactions(currentPage);
        }).addOrUpdateTransaction(txData);
      }

      function deleteTransaction() {
        const rowIdx = document.getElementById("txRowIdx").value;
        if (!rowIdx) return;

        if (confirm("Are you sure you want to delete this transaction from your sheet?")) {
          M.toast({html: 'Deleting row... ⏳', classes: 'blue'});
          google.script.run.withSuccessHandler(function(res) {
            M.toast({html: 'Transaction deleted! 🗑️', classes: 'green'});
            clearForm();
            loadTransactions(currentPage);
          }).deleteTransactionRow(rowIdx);
        }
      }

      function changePage(delta) {
        const target = currentPage + delta;
        if (target >= 1 && target <= totalPages) {
          loadTransactions(target);
        }
      }

      function debounceSearch() {
        clearTimeout(searchTimeout);
        searchTimeout = setTimeout(function() {
          loadTransactions(1);
        }, 300);
      }
    </script>
  </body>
</html>
```


## 6. Isolated Trip & Project Expense Tracker Engine (`Trip_Budget`)

To prevent special high-ticket expenses (such as a vacation, wedding, conference, or home remodel) from distorting your regular monthly household spending metrics, V11 includes a dedicated **Trip & Project Expense Tracker Generator**:

### A. How to Create a New Trip/Project Budget
1. Click **`Financial Automation 🚀` > `➕ Create Trip / Project Budget`** on your toolbar.
2. Enter the title of your trip or project (e.g., `NYC Trip 2026`, `Israel Vacation`, `Kitchen Remodel`).
3. Enter your overall target planned budget limit (e.g., `$3,500.00`).
4. The script will programmatically create a clean, dedicated tab named **`Trip_[Your_Title]`** with:
   * **KPI Summary Header Cards**: Total Planned Budget, Total Actual Spend, Variance Remaining, and Status (`ON TRACK ✅` / `OVER BUDGET ❌`).
   * **Category Summary Table**: Automatically computes actual spending across trip categories (`Flights & Airfare`, `Lodging & Hotels`, `Meals & Dining`, `Events & Tickets`, `Sightseeing`, `Ground Transportation`, `Miscellaneous`) using dynamic `=SUMIF()` formulas.
   * **Itemized Expense Log**: An isolated ledger where you can record dates, descriptions, costs, trip categories (via dropdown), and payment sources.

---

## 6. Categorization Assistant Sidebar (`CategorizerView.html`)

To provide an ultra-convenient visual interface for sorting uncategorized transactions, checks, Zelle payments, and wire transfers, create a third HTML file in Apps Script named **`CategorizerView`** (omit `.html`) and paste this code:

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 12px; font-family: 'Arial', sans-serif; background-color: #f8f9fa; }
      .header-title { font-weight: bold; color: #1f4e78; margin-bottom: 12px; font-size: 1.1rem; display: flex; align-items: center; justify-content: space-between; }
      .queue-badge { background-color: #1f4e78; color: white; border-radius: 12px; padding: 2px 10px; font-size: 0.8rem; }
      .tx-card { background: white; border-radius: 8px; padding: 14px; box-shadow: 0 2px 5px rgba(0,0,0,0.08); margin-bottom: 15px; border-left: 5px solid #1f4e78; }
      .tx-amount { font-size: 1.2rem; font-weight: bold; }
      .outflow { color: #d32f2f; }
      .inflow { color: #2e7d32; }
      .raw-merchant { font-size: 0.85rem; color: #666; word-break: break-word; background: #f1f3f4; padding: 6px; border-radius: 4px; margin: 8px 0; font-family: monospace; }
      .btn-apply { width: 100%; background-color: #1f4e78 !important; margin-top: 10px; }
      .btn-skip { width: 100%; background-color: #757575 !important; margin-top: 6px; }
      .empty-state { text-align: center; padding: 30px 10px; }
      .empty-state i { font-size: 3.5rem; color: #2e7d32; margin-bottom: 10px; }
    </style>
  </head>
  <body>
    <div class="header-title">
      <span>Categorizer Assistant</span>
      <span id="queueBadge" class="queue-badge">0 left</span>
    </div>

    <div id="loader" style="text-align:center; margin-top:30px;">
      <div class="preloader-wrapper small active"><div class="spinner-layer spinner-blue-only"><div class="circle-clipper left"><div class="circle"></div></div></div></div>
      <p style="color:#666; font-size:0.85rem;">Scanning uncategorized transactions...</p>
    </div>

    <div id="emptyState" class="empty-state" style="display:none;">
      <i class="material-icons">check_circle</i>
      <h5>All Caught Up!</h5>
      <p style="color:#666;">Every transaction in your database is cleanly categorized.</p>
    </div>

    <div id="cardContainer" style="display:none;">
      <div class="tx-card">
        <div style="display:flex; justify-content:space-between; align-items:center;">
          <span id="txDate" style="color:#666;"></span>
          <span id="txAmount" class="tx-amount"></span>
        </div>
        
        <div style="margin-top:8px;">
          <span style="font-size:0.75rem; color:#888; font-weight:bold;">Raw Statement Text:</span>
          <div id="txRawMerchant" class="raw-merchant"></div>
        </div>

        <div class="input-field" style="margin-top:10px; margin-bottom:10px;">
          <input type="text" id="cleanNameInput">
          <label for="cleanNameInput" class="active">Standardized Clean Name</label>
        </div>

        <div style="margin-top:10px;">
          <label style="color:#333; font-weight:bold;">Assign Category:</label>
          <select id="categorySelect" class="browser-default" style="margin-top:4px; border:1px solid #ccc; border-radius:4px; padding:6px;"></select>
        </div>

        <div style="margin-top:12px; background:#e8f0fe; padding:8px; border-radius:4px;">
          <label>
            <input type="checkbox" id="saveRuleCheck" class="filled-in" checked="checked" />
            <span style="font-size:0.82rem; color:#1a73e8; font-weight:bold;">Save rule for keyword:</span>
          </label>
          <input type="text" id="keywordRuleInput" style="height:1.8rem; margin:4px 0 0 0; background:white; padding:2px 6px; border-radius:3px;">
        </div>

        <button class="btn waves-effect waves-light btn-apply" onclick="applyCategory()">Categorize & Save Rule</button>
        <button class="btn waves-effect waves-light btn-skip" onclick="skipTransaction()">Skip For Now</button>
      </div>
    </div>
  </body>
</html>
```

---

## 7. How Special Transaction Types are Handled (Checks, Zelle, Wires, ACH)

To resolve non-standard bank transaction strings, the V11 engine applies structured pattern recognition:

1. **Checks (`CHECK #1024` / `CHECK PAID`)**:
   - Automatically extracts Check Numbers (e.g. `Check #1024`).
   - Default Category: `Checks & Banking` (or custom rule from `Category_Rules_DB`).
2. **QuickPay / Zelle (`ZELLE TO JOHN DOE REF 9812`)**:
   - Extracts Payer / Payee names and cleans them (e.g. `Zelle: John Doe`).
   - Flow: Correctly assigns `INFLOW` or `OUTFLOW`.
   - Default Category: `Personal` or `Transfer` (or custom auto-learned rule).
3. **Wire Transfers (`WIRE TRANS OUT REF# 8812`)**:
   - Cleans to `Wire Transfer (8812)`.
   - Flow: Assigns `TRANSFER` or `OUTFLOW`.
4. **ACH / Direct Deposit (`ACH CREDIT PAYROLL`)**:
   - Cleans to `Direct Deposit / ACH`.
   - Category: Automatically maps to `Income` for inflow credits.
