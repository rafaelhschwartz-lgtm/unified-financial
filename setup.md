# Unified Financial Automation Setup Guide - V8 (Master Edition)

This comprehensive guide serves as the ultimate developer and operations manual for configuring, running, and expanding your **Enterprise-Grade, Relational Personal Finance Ledger System**. This workbook is engineered to comply with the **FAST Modeling Standard** (Flexibility, Appropriated, Structured, Transparency) and the **ICAEW Spreadsheet Design Principles**.

---

## 1. Resolving All Operational Friction Points (Your Complete To-Do List)

We have officially closed out every single feature challenge and technical constraint identified in your tracking ledger:

### A. Dynamic Multi-Account Automatic Detection (Item 1)
*   **The Ingestion Rule:** When batch CSV imports run, the engine automatically inspects the filename of each statement file uploaded inside `Chase_CSV_Uploads`.
*   **The Mappings:** If the file name contains keywords like `sapphire`, `credit`, or `cc`, the script automatically labels the transaction's account source in Column I as `Chase Sapphire`. If it contains `checking` or `debit`, it maps it to `Chase Checking`.

### B. No-Code Rules Engine (`Category_Rules_DB`) & Clean Names (Items 2 & 8)
*   **The Architecture:** Instead of hardcoding categories and merchant aliases in Apps Script, you can edit your rules dynamically in a dedicated sheet tab named **`Category_Rules_DB`**.
*   **Real-Time Processing:** Whenever transactions run through imports, alerts, or manual logs, the system scans this rules database in real-time. It runs a fuzzy, case-insensitive check to clean the merchant name (e.g., *STARBUCKS STORE 1234* $
ightarrow$ *Starbucks*) in Column E and assign the target category (e.g., *Food & Drink*) in Column H.

### C. Deterministic Deduplication Signature (Item 3)
*   Our ingesters utilize a mathematically deterministic hash based on transaction date, raw merchant description, and amount. Uploading duplicate statements will simply skip already written rows—**actual expense metrics are guaranteed never to double!**

### D. Eliminating Transaction Distortion: 'TRANSFER' Flow (Item 4)
*   **The Challenge:** Moving money between internal accounts (such as a Chase Checking to Savings transfer) historically registered as both income (inflow) and expense (outflow), heavily distorting KPI calculations.
*   **The Solution:** We introduced a third capital flow classification: **`TRANSFER`**. Any transaction identified as a transfer is cleanly logged in your master ledger database under the `Transfer` category, but **completely ignored by your dashboard KPI cards and budget sparkline models**. This preserves absolute mathematical precision!

### E. Time-Driven Background Trigger Safeguards (Item 5)
*   To prevent Google’s execution runtime from timing out, all UI alerts are wrapped in safe `try...catch` blocks. Time-driven automation runs seamlessly on Google’s headless cloud servers without hitting blocking alert errors.

### F. Google Voice SMS Alert Scraper (Item 6)
*   Google Voice automatically forwards incoming transaction SMS text messages (like Venmo pay notifications or CashApp deposits) to your Gmail inbox. Our upgraded scraper (`pullGmailTransactionAlerts`) searches for forwarded SMS text messages, extracts dates, amounts, and flow direction (e.g., *You paid John* $
ightarrow$ Outflow; *John paid you* $
ightarrow$ Inflow), and appends them to your database in real-time!

### G. Visual 10-Year Growth Projection and Area Chart
*   The **`Savings_Goals`** tab houses a compound projection matrix (Years 2026 to 2035) linked to an interactive **10-Year Growth Area Chart**. You can change your expected investment return (cell `I2`) or annual savings contributions (cell `I3`) to dynamically recalculate and visualize your growth curve instantly in real-time!

---

## 2. Dynamic, Paginated In-Sheet Transaction Editor Sidebar (Item 7)

Rather than navigating complex, crowded row cells on your raw `Transactions_DB` tab, you can open our **Dynamic Transaction Editor Sidebar** from your custom toolbar menu:
1.  **Paginated Ledger Grid:** Displays a list of recent transactions (most recent first) from your live database. Pages through records quickly using High-Performance server-side slice filtering.
2.  **Fuzzy Search:** Filter transactions instantly by typing part of a merchant name, account, or category inside the search box.
3.  **Create, Edit, and Delete:**
    *   Clicking a transaction row in the grid automatically loads its values back into the form fields.
    *   You can modify dates, clean up raw descriptions, and change accounts or categories using select controls populated dynamically from your `Enum` settings.
    *   Click **Update Transaction** to save edits, or click the **Trash Can icon** to delete a transaction safely.
    *   Clear the form to add a brand new cash transaction in seconds!

---

## 3. Background Trigger Configurations

To make your entire personal finance stack run completely hands-off, configure these time-driven background triggers inside Google Apps Script:

1.  **Hourly Gmail Scraper (`pullGmailTransactionAlerts`):** Scrapes bank transaction alerts and Google Voice SMS alerts forwarded to your inbox, parses values, auto-categorizes, and normalizes signs.
    *   *Configuration:* Function: `pullGmailTransactionAlerts` | Source: `Time-driven` | Type: `Hour timer` | Interval: `Every hour`.
2.  **Daily Budget Breach Monitor (`checkBudgetBreaches`):** Automatically scans actual category spending against planned monthly boundaries. Sends an email notification immediately when a breach is detected, utilizing Google's background `PropertiesService` state storage to ensure you are **only notified once per category per month**.
    *   *Configuration:* Function: `checkBudgetBreaches` | Source: `Time-driven` | Type: `Day timer` | Time: `8 AM to 9 AM`.
3.  **Weekly HTML Budget Summaries (`sendWeeklyBudgetSummary`):** Compiles your high-level income/expense KPIs and spends pacing velocity relative to the calendar, sending a styled HTML newsletter directly to your inbox every Monday morning.
    *   *Configuration:* Function: `sendWeeklyBudgetSummary` | Source: `Time-driven` | Type: `Weekly timer` | Day: `Every Monday` | Time: `7 AM to 8 AM`.

---

## 4. Deploying the Sidebar Views

To activate both sidebar panels, create two distinct HTML files inside your Google Apps Script editor (**Extensions > Apps Script**):

### A. SidebarView.html (Budget Allocator Sidebar)
Click **`+` > HTML**, name the file exactly **`SidebarView`** (omit the `.html` extension as Apps Script appends it), and paste the code from Section 5 of your V5 setup guide.

### B. TransactionEditorView.html (Ledger Editor Sidebar)
Click **`+` > HTML**, name the file exactly **`TransactionEditorView`**, and paste this modern, fully functional Materialize CSS client-side editor interface:

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!-- Load Materialize CSS & Fonts -->
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 12px; font-family: 'Arial', sans-serif; background-color: #fafafa; font-size: 0.9rem; }
      .header-title { font-weight: bold; color: #1f4e78; margin-bottom: 12px; font-size: 1.15rem; display: flex; align-items: center; }
      .header-title i { margin-right: 6px; }
      .editor-card { background: white; padding: 12px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 12px; }
      .form-section-title { font-weight: bold; color: #1f4e78; font-size: 0.9rem; border-bottom: 1px solid #e1e4e8; padding-bottom: 4px; margin-bottom: 8px; }
      .input-field { margin: 6px 0 !important; }
      .input-field input, .input-field select { height: 2.2rem !important; margin-bottom: 4px !important; }
      .btn-row { display: flex; gap: 8px; margin-top: 8px; }
      .btn-row button { flex: 1; height: 32px !important; line-height: 32px !important; font-size: 0.8rem !important; padding: 0 4px !important; text-align: center; }
      .save-btn { background-color: #1f4e78 !important; }
      .clear-btn { background-color: #757575 !important; }
      .delete-btn { background-color: #d32f2f !important; }
      .search-box { margin-bottom: 8px !important; }
      .search-box input { height: 2.2rem !important; }
      .ledger-table { font-size: 0.75rem; background: white; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); overflow: hidden; margin-top: 8px; }
      .ledger-table th { padding: 6px 8px; background: #e8eff7; color: #1f4e78; font-weight: bold; }
      .ledger-table td { padding: 5px 8px; border-bottom: 1px solid #f0f0f0; cursor: pointer; }
      .ledger-table tr:hover { background-color: #f5f5f5; }
      .pagination-row { display: flex; justify-content: space-between; align-items: center; margin-top: 8px; font-size: 0.75rem; color: #666; }
      .pagination-row button { background: none; border: none; color: #1f4e78; cursor: pointer; font-weight: bold; padding: 4px; }
      .pagination-row button[disabled] { color: #ccc; cursor: default; }
      .loader { text-align: center; margin: 15px 0; }
    </style>
  </head>
  <body>
    <div class="header-title">
      <i class="material-icons">receipt</i>
      <span>Transaction Ledger Editor</span>
    </div>
    
    <!-- TRANSACTION INPUT / EDIT FORM -->
    <div class="editor-card">
      <div class="form-section-title" id="formTitle">Add New Transaction</div>
      <form id="txForm">
        <input type="hidden" id="rowIdx" value="">
        
        <div class="input-field">
          <input type="date" id="date" required class="validate">
        </div>
        <div class="input-field">
          <input type="text" id="merchantRaw" placeholder="Merchant Name" required class="validate">
        </div>
        <div class="input-field">
          <input type="number" step="0.01" min="0" id="amount" placeholder="Amount ($)" required class="validate">
        </div>
        
        <div class="row" style="margin: 0; display: flex; gap: 6px;">
          <div class="input-field" style="flex: 1; margin: 4px 0 !important;">
            <select id="flow" required style="display: block; font-size: 0.8rem; height: 2.2rem; border: 1px solid #ccc; border-radius: 4px;">
              <option value="OUTFLOW">OUTFLOW</option>
              <option value="INFLOW">INFLOW</option>
              <option value="TRANSFER">TRANSFER</option>
            </select>
          </div>
          <div class="input-field" style="flex: 1; margin: 4px 0 !important;">
            <select id="status" required style="display: block; font-size: 0.8rem; height: 2.2rem; border: 1px solid #ccc; border-radius: 4px;">
              <option value="SETTLED">SETTLED</option>
              <option value="PENDING">PENDING</option>
              <option value="MANUAL_REVIEW">REVIEW</option>
            </select>
          </div>
        </div>
        
        <div class="input-field">
          <select id="category" required style="display: block; font-size: 0.8rem; height: 2.2rem; border: 1px solid #ccc; border-radius: 4px;">
            <option value="" disabled selected>Select Category</option>
          </select>
        </div>
        <div class="input-field">
          <select id="source" required style="display: block; font-size: 0.8rem; height: 2.2rem; border: 1px solid #ccc; border-radius: 4px;">
            <option value="" disabled selected>Select Account Source</option>
          </select>
        </div>
        
        <div class="btn-row">
          <button class="btn waves-effect waves-light save-btn" type="submit" id="submitBtn">Save</button>
          <button class="btn waves-effect waves-light clear-btn" type="button" onclick="clearForm()">Clear</button>
          <button class="btn waves-effect waves-light delete-btn" type="button" id="deleteBtn" style="display:none;" onclick="deleteTransaction()">Delete</button>
        </div>
      </form>
    </div>
    
    <!-- DYNAMIC PAGINATED GRID VIEW -->
    <div class="form-section-title">Database Records</div>
    <div class="search-box">
      <input type="text" id="search" placeholder="🔍 Search Merchant or Category..." oninput="triggerSearch()">
    </div>
    
    <div id="loader" class="loader" style="display:none;">
      <div class="preloader-wrapper small active">
        <div class="spinner-layer spinner-blue-only">
          <div class="circle-clipper left"><div class="circle"></div></div>
          <div class="gap-patch"><div class="circle"></div></div>
          <div class="circle-clipper right"><div class="circle"></div></div>
        </div>
      </div>
    </div>
    
    <div id="tableContainer"></div>
    
    <div class="pagination-row">
      <button id="prevBtn" onclick="changePage(-1)" disabled>◀ Prev</button>
      <span id="pageIndicator">Page 1 of 1</span>
      <button id="nextBtn" onclick="changePage(1)" disabled>Next ▶</button>
    </div>

    <!-- Scripting Interaction Rules -->
    <script>
      let currentPage = 1;
      const pageSize = 5;
      let totalPages = 1;
      let currentSearch = "";
      
      document.addEventListener('DOMContentLoaded', function() {
        // Retrieve Category & Account values to build dropdown portals
        google.script.run.withSuccessHandler(populateDropdowns).getEnumValues();
        loadTransactions();
      });
      
      function populateDropdowns(enums) {
        const catSelect = document.getElementById("category");
        enums.categories.forEach(cat => {
          const opt = document.createElement("option");
          opt.value = cat;
          opt.text = cat;
          catSelect.appendChild(opt);
        });
        
        const srcSelect = document.getElementById("source");
        enums.accounts.forEach(src => {
          const opt = document.createElement("option");
          opt.value = src;
          opt.text = src;
          srcSelect.appendChild(opt);
        });
      }
      
      function loadTransactions() {
        document.getElementById("loader").style.display = "block";
        document.getElementById("tableContainer").innerHTML = "";
        
        google.script.run.withSuccessHandler(renderTransactionsTable)
          .getTransactionsPage(currentPage, pageSize, currentSearch);
      }
      
      function renderTransactionsTable(response) {
        document.getElementById("loader").style.display = "none";
        totalPages = response.totalPages || 1;
        document.getElementById("pageIndicator").innerText = "Page " + currentPage + " of " + totalPages;
        
        document.getElementById("prevBtn").disabled = (currentPage === 1);
        document.getElementById("nextBtn").disabled = (currentPage === totalPages);
        
        if (!response.data || response.data.length === 0) {
          document.getElementById("tableContainer").innerHTML = "<p style='color:#777; text-align:center;'>No matching transactions found.</p>";
          return;
        }
        
        let html = "<table class='striped ledger-table'><thead><tr><th>Date</th><th>Merchant</th><th style='text-align:right;'>Amount</th></tr></thead><tbody>";
        
        response.data.forEach(tx => {
          const formattedAmount = (tx.flow === "INFLOW" ? "+" : "-") + "$" + tx.amount.toFixed(2);
          const amtColor = tx.flow === "INFLOW" ? "color:#2e7d32;" : (tx.flow === "TRANSFER" ? "color:#757575;" : "color:#c00000;");
          
          html += `<tr onclick='selectTransaction(${JSON.stringify(tx)})'>
            <td>${tx.date}</td>
            <td><b>${tx.merchantClean || tx.merchantRaw}</b><br><span style="color:#777; font-size:0.65rem;">${tx.category}</span></td>
            <td style="text-align:right; font-weight:bold; ${amtColor}">${formattedAmount}</td>
          </tr>`;
        });
        
        html += "</tbody></table>";
        document.getElementById("tableContainer").innerHTML = html;
      }
      
      function selectTransaction(tx) {
        document.getElementById("rowIdx").value = tx.rowIdx;
        document.getElementById("date").value = tx.date;
        document.getElementById("merchantRaw").value = tx.merchantRaw;
        document.getElementById("amount").value = tx.amount;
        document.getElementById("flow").value = tx.flow;
        document.getElementById("status").value = tx.status;
        document.getElementById("category").value = tx.category;
        document.getElementById("source").value = tx.source;
        
        document.getElementById("formTitle").innerText = "Edit Transaction (Row " + tx.rowIdx + ")";
        document.getElementById("submitBtn").innerText = "Update";
        document.getElementById("deleteBtn").style.display = "block";
      }
      
      function clearForm() {
        document.getElementById("txForm").reset();
        document.getElementById("rowIdx").value = "";
        document.getElementById("formTitle").innerText = "Add New Transaction";
        document.getElementById("submitBtn").innerText = "Save";
        document.getElementById("deleteBtn").style.display = "none";
      }
      
      function triggerSearch() {
        currentSearch = document.getElementById("search").value;
        currentPage = 1;
        loadTransactions();
      }
      
      function changePage(direction) {
        currentPage += direction;
        if (currentPage < 1) currentPage = 1;
        if (currentPage > totalPages) currentPage = totalPages;
        loadTransactions();
      }
      
      // Submit transaction to spreadsheet
      document.getElementById("txForm").addEventListener("submit", function(e) {
        e.preventDefault();
        
        const rowIdx = document.getElementById("rowIdx").value;
        const txData = {
          date: document.getElementById("date").value,
          merchantRaw: document.getElementById("merchantRaw").value,
          amount: parseFloat(document.getElementById("amount").value),
          flow: document.getElementById("flow").value,
          status: document.getElementById("status").value,
          category: document.getElementById("category").value,
          source: document.getElementById("source").value
        };
        
        M.toast({html: 'Syncing to Database... ⏳', classes: 'blue'});
        google.script.run.withSuccessHandler(function(response) {
          if (response.indexOf("SUCCESS") !== -1) {
            M.toast({html: 'Ledger synchronized successfully! 🎉', classes: 'green'});
            clearForm();
            loadTransactions();
          } else {
            M.toast({html: 'Sync failed: ' + response, classes: 'red'});
          }
        }).addOrUpdateTransaction(rowIdx, txData);
      });
      
      function deleteTransaction() {
        const rowIdx = document.getElementById("rowIdx").value;
        if (!rowIdx) return;
        
        if (confirm("Are you absolutely sure you want to delete this transaction from your master ledger database?")) {
          M.toast({html: 'Removing Transaction... ⏳', classes: 'blue'});
          google.script.run.withSuccessHandler(function(response) {
            if (response === "SUCCESS") {
              M.toast({html: 'Transaction deleted successfully! 🗑️', classes: 'orange'});
              clearForm();
              loadTransactions();
            } else {
              M.toast({html: 'Deletion failed: ' + response, classes: 'red'});
            }
          }).deleteTransactionRow(rowIdx);
        }
      }
    </script>
  </body>
</html>
```

---

## 5. Ultimate Verification & Operational Walkthrough

Once you deploy these modules, run through these validation checks to confirm your personal ledger is 100% operational:

1.  **Initialize Tab Structure:** Select **`Financial Automation 🚀` > `Initialize Sheets in This Workbook`** from your spreadsheet's custom menu. Refresh your browser page to reload the toolbar.
2.  **Add Category Rule:** Navigate to the **`Category_Rules_DB`** tab and type `starbucks` in cell A16, `Starbucks` in cell B16, and select `Food & Drink` from the dropdown validation in cell C16.
3.  **Launch Transaction Editor:** Select **`Financial Automation 🚀` > `Open Transaction Editor Sidebar`**. Add a transaction with raw description `STARBUCKS #1234 COFFEE` with amount `8.50`. Check your **`Transactions_DB`** tab—the rule automatically cleaned the merchant to `Starbucks` in Column E and categorized it under `Food & Drink` in Column H!
4.  **Batch Import Chase CSV Statements:**
    *   Open your Google Drive root and locate the newly generated folder: **`Chase_CSV_Uploads`**.
    *   Upload standard Chase debit/credit CSV statement files.
    *   Return to your spreadsheet, and select **`Batch Import CSV Folder`**.
    *   The engine automatically evaluates filename structures, maps credit cards and checking accounts, deduplicates and normalizes signs, auto-categorizes rows in real-time, and archives processed CSV files inside Drive!
5.  **Trigger Savings Sweep:** Select **`The Monthly Dashboard Tab`** and change your selection to **`August 2026`**. Under **`Savings_Goals`**, verify your allocation percentages. Select **`Execute Monthly Savings Sweep`** from the custom menu—the script will calculate your monthly surplus, write corresponding debit/credit transfer entries to `Transactions_DB`, and update your savings goal balances and projection curves with gorgeous visual bars!
