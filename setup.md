# Unified Financial Automation Setup Guide - V16 (User-Friendly Modal Edition)

This comprehensive guide serves as the ultimate manual for configuring, testing, and expanding your **Enterprise-Grade, Relational Personal Finance Ledger System**. This workbook is engineered to comply with the **FAST Modeling Standard** (Flexibility, Appropriateness, Structure, Transparency) and the **ICAEW Spreadsheet Design Principles**.

---

## 1. V16 User-Experience Upgrade: Rich HTML Modal Form Dialogs

In previous versions, adding a transaction, a new CD investment, a merchant category rule, or a trip budget relied on text prompts (`ui.prompt()`) where users had to type comma-separated strings (e.g. `Marcus, CD-8812, 10000, 2026-01-15, 12, 5.25, YES, 4.50`).

In **V16**, all comma-separated prompt dialogs have been replaced with **Polished, Responsive HTML Modal Dialogs**:

### A. Add Certificate of Deposit (CD) Modal (`promptAddCD`)
* **Interactive Controls**: Features date pickers (`<input type="date">`), styled currency inputs, APY percentage fields, and a dropdown for `Auto-Renew? (YES/NO)`.
* **Zero Syntax Errors**: Automatically formats dates, currency values, maturity calculations (`=EDATE()`), compounding yield formulas, and updates the CD Principal vs Accrued Yield chart seamlessly.

### B. Quick Add Transaction Modal (`quickAddTransactionPrompt`)
* **Dynamic Dropdowns**: Category and Payment Account Source dropdowns are dynamically populated directly from your active **`Enum`** tab!
* **Flow Direction Selector**: Easy dropdown selection for `OUTFLOW (Expense)`, `INFLOW (Income)`, or `TRANSFER (Internal)`.
* **Trip Tag Field**: Optional text field to immediately attach trip tags (e.g., `Israel_TLV_2026`).

### C. Merchant Category Rule Modal (`quickAddRulePrompt`)
* **Clean Interface**: Input boxes for Merchant Keyword, Clean Display Title, and a Category dropdown selector.
* **Retroactive Auto-Apply**: Saves the rule to **`Category_Rules_DB`** and immediately retroactively updates all matching uncategorized transactions in **`Transactions_DB`**.

### D. Create Trip / Project Expense Tracker Modal (`promptTripBudget`)
* **Structured Input**: Form fields for Trip Title (e.g., `Israel Vacation 2026`) and Total Planned Budget ($).
* **Instant Tab Generation**: Programmatically constructs the `Trip_[Title]` tab with `=SUMIFS()` budget tracking and isolated expense logs.

---
## 2. Step-by-Step Deployment Instructions

1. **Clear Legacy Backend Code**: Open **Extensions > Apps Script** in your Google Sheet. Delete any old `.gs` files (such as `main.gs` or `Setup.gs`).
2. **Deploy V12 Backend Script**: Create or open `main.gs`, paste the entire code from **`unified-financial-automation-v12.gs`**, and click **Save**.
3. **Deploy HTML Client Sidebars**:
   * Click **`+` > HTML**, name the file **`SidebarView`**, and paste the code from Section 3. Save.
   * Click **`+` > HTML**, name the file **`TransactionEditorView`**, and paste the code from Section 4. Save.
   * Click **`+` > HTML**, name the file **`CategorizerView`**, and paste the code from Section 5. Save.
4. **Initialize Workbook**: Return to your spreadsheet, refresh the page, and select **`Financial Automation 🚀` > `Initialize Sheets in This Workbook`**. Allow permissions when prompted.

---

## 3. Client Sidebar Code 1: `SidebarView.html` (Budget Allocator)

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
      .category-row { margin-bottom: 12px; background: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
      .category-label { font-weight: bold; color: #333; font-size: 0.95rem; margin-bottom: 6px; }
      .currency-field { display: flex; align-items: center; border-bottom: 1px solid #1f4e78; background: #f8f9fa; border-radius: 4px; padding: 2px 8px; }
      .currency-symbol { font-weight: bold; color: #1f4e78; font-size: 1.1rem; margin-right: 6px; }
      .currency-field input { border: none !important; box-shadow: none !important; margin: 0 !important; height: 2.2rem !important; font-size: 1rem !important; font-weight: bold; color: #002060; width: 100%; background: transparent !important; }
      .save-btn { width: 100%; background-color: #1f4e78 !important; margin-top: 20px; height: 42px; font-weight: bold; }
      .loader { margin: 20px auto; text-align: center; }
      .period-indicator { background: #e8eff7; padding: 10px; border-radius: 6px; margin-bottom: 15px; font-size: 0.95rem; color: #1f4e78; font-weight: bold; text-align: center; border: 1px solid #d9e1f2; }
    </style>
  </head>
  <body>
    <div class="header-title">
      <i class="material-icons">account_balance_wallet</i>
      <span>Budget History Ceilings</span>
    </div>
    
    <div id="periodIndicator" class="period-indicator">Loading period...</div>
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
      <button class="btn waves-effect waves-light save-btn" type="submit">
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
            <div class="currency-field">
              <span class="currency-symbol">$</span>
              <input type="number" step="0.01" min="0" class="budget-input" data-category="${item.category}" value="${parseFloat(item.amount || 0).toFixed(2)}">
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
        
        M.toast({html: 'Syncing Budgets... ⏳', classes: 'blue'});
        google.script.run.withSuccessHandler(onSaveSuccess).saveCategoryBudgets(activeYear, activeMonth, budgetData);
      });

      function onSaveSuccess(response) {
        if (response === "SUCCESS") {
          M.toast({html: 'Budgets Saved Successfully! 🎉', classes: 'green'});
        } else {
          M.toast({html: 'Save Failed ❌', classes: 'red'});
        }
      }
    </script>
  </body>
</html>
```

---

## 4. Client Sidebar Code 2: `TransactionEditorView.html` (Ledger Editor)

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 12px; font-family: 'Arial', sans-serif; background-color: #f8f9fa; }
      .editor-header { font-weight: bold; color: #1f4e78; margin-bottom: 12px; font-size: 1.1rem; display: flex; align-items: center; justify-content: space-between; }
      .editor-card { background: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 15px; }
      .input-field { margin-top: 5px; margin-bottom: 5px; }
      .btn-row { display: flex; gap: 8px; margin-top: 10px; }
      .search-box { margin-bottom: 10px; }
      .tx-item { background: white; padding: 10px; border-radius: 4px; margin-bottom: 8px; border-left: 4px solid #1f4e78; cursor: pointer; transition: all 0.2s; }
      .tx-item:hover { background: #e8eff7; }
      .tx-item.outflow { border-left-color: #db4437; }
      .tx-item.inflow { border-left-color: #385723; }
      .tx-item.transfer { border-left-color: #1f4e78; }
      .tx-merchant { font-weight: bold; font-size: 0.9rem; color: #333; }
      .tx-details { font-size: 0.75rem; color: #666; display: flex; justify-content: space-between; margin-top: 4px; }
      .pagination-controls { display: flex; justify-content: space-between; align-items: center; margin-top: 10px; font-size: 0.85rem; }
    </style>
  </head>
  <body>
    <div class="editor-header">
      <span><i class="material-icons left">edit_note</i>Ledger Editor</span>
      <button class="btn-small blue darken-3" onclick="resetForm()"><i class="material-icons left">add</i>New</button>
    </div>

    <div class="editor-card">
      <input type="hidden" id="txRowIdx" value="">
      <div class="row" style="margin-bottom:0;">
        <div class="input-field col s6">
          <input type="date" id="txDate">
        </div>
        <div class="input-field col s6">
          <input type="number" step="0.01" id="txAmount" placeholder="Amount ($)">
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
      <div class="input-field">
        <select id="txAccount" class="browser-default">
          <option value="">Account Source...</option>
        </select>
      </div>
      <div class="btn-row">
        <button class="btn btn-small blue darken-3 waves-effect" onclick="saveTx()" style="flex:1;">Save</button>
        <button class="btn btn-small red darken-2 waves-effect" id="btnDelete" onclick="deleteTx()" style="display:none;"><i class="material-icons">delete</i></button>
        <button class="btn btn-small grey waves-effect" onclick="resetForm()">Clear</button>
      </div>
    </div>

    <div class="search-box">
      <input type="text" id="searchInput" placeholder="Search merchant, category, account..." onkeyup="debounceSearch()">
    </div>

    <div id="txList">Loading transactions...</div>

    <div class="pagination-controls">
      <button class="btn-small grey lighten-2 black-text" id="btnPrev" onclick="changePage(-1)">Prev</button>
      <span id="pageInfo">Page 1</span>
      <button class="btn-small grey lighten-2 black-text" id="btnNext" onclick="changePage(1)">Next</button>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script>
      let currentPage = 1;
      let totalPages = 1;
      let searchTimeout = null;

      document.addEventListener('DOMContentLoaded', function() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('txDate').value = today;
        loadEnums();
        loadPage(1);
      });

      function loadEnums() {
        google.script.run.withSuccessHandler(function(enums) {
          const catSelect = document.getElementById('txCategory');
          catSelect.innerHTML = '<option value="">Category...</option>';
          enums.categories.forEach(c => {
            catSelect.innerHTML += `<option value="${c}">${c}</option>`;
          });

          const acctSelect = document.getElementById('txAccount');
          acctSelect.innerHTML = '<option value="">Account Source...</option>';
          enums.accounts.forEach(a => {
            acctSelect.innerHTML += `<option value="${a}">${a}</option>`;
          });
        }).getEnumValues();
      }

      function loadPage(page) {
        currentPage = page;
        const query = document.getElementById('searchInput').value;
        google.script.run.withSuccessHandler(renderPage).getTransactionsPage(currentPage, 10, query);
      }

      function renderPage(res) {
        totalPages = res.totalPages;
        document.getElementById('pageInfo').innerText = `Page ${res.currentPage} of ${totalPages} (${res.totalRecords})`;
        document.getElementById('btnPrev').disabled = (res.currentPage <= 1);
        document.getElementById('btnNext').disabled = (res.currentPage >= totalPages);

        const container = document.getElementById('txList');
        if (!res.data || res.data.length === 0) {
          container.innerHTML = '<div style="text-align:center; color:#999; padding:15px;">No transactions found.</div>';
          return;
        }

        let html = '';
        res.data.forEach(item => {
          const flowClass = item.flow.toLowerCase();
          const cleanName = item.merchantClean || item.merchantRaw || "Transaction";
          html += `
            <div class="tx-item ${flowClass}" onclick='selectTx(${JSON.stringify(item)})'>
              <div class="tx-merchant">${cleanName}</div>
              <div class="tx-details">
                <span>${item.date} • ${item.category}</span>
                <strong style="color:${flowClass === 'inflow' ? '#385723' : (flowClass === 'transfer' ? '#1f4e78' : '#db4437')}">
                  ${flowClass === 'inflow' ? '+' : '-'}$${item.amount.toFixed(2)}
                </strong>
              </div>
            </div>`;
        });
        container.innerHTML = html;
      }

      function selectTx(item) {
        document.getElementById('txRowIdx').value = item.rowIdx;
        document.getElementById('txDate').value = item.date;
        document.getElementById('txMerchant').value = item.merchantClean || item.merchantRaw;
        document.getElementById('txAmount').value = item.amount;
        document.getElementById('txFlow').value = item.flow;
        document.getElementById('txCategory').value = item.category;
        document.getElementById('txAccount').value = item.accountSource;
        document.getElementById('btnDelete').style.display = 'inline-block';
      }

      function resetForm() {
        document.getElementById('txRowIdx').value = '';
        document.getElementById('txDate').value = new Date().toISOString().split('T')[0];
        document.getElementById('txMerchant').value = '';
        document.getElementById('txAmount').value = '';
        document.getElementById('txFlow').value = 'OUTFLOW';
        document.getElementById('txCategory').value = '';
        document.getElementById('txAccount').value = '';
        document.getElementById('btnDelete').style.display = 'none';
      }

      function saveTx() {
        const txData = {
          rowIdx: document.getElementById('txRowIdx').value,
          date: document.getElementById('txDate').value,
          merchant: document.getElementById('txMerchant').value,
          amount: parseFloat(document.getElementById('txAmount').value) || 0,
          flow: document.getElementById('txFlow').value,
          category: document.getElementById('txCategory').value,
          accountSource: document.getElementById('txAccount').value
        };

        if (!txData.merchant || txData.amount <= 0) {
          M.toast({html: 'Please enter merchant and valid amount', classes: 'orange'});
          return;
        }

        M.toast({html: 'Saving transaction... ⏳', classes: 'blue'});
        google.script.run.withSuccessHandler(function(res) {
          M.toast({html: 'Transaction saved! 🎉', classes: 'green'});
          resetForm();
          loadPage(currentPage);
        }).addOrUpdateTransaction(txData);
      }

      function deleteTx() {
        const rowIdx = document.getElementById('txRowIdx').value;
        if (!rowIdx) return;

        if (confirm("Delete this transaction permanently?")) {
          M.toast({html: 'Deleting... ⏳', classes: 'blue'});
          google.script.run.withSuccessHandler(function(res) {
            M.toast({html: 'Transaction deleted! 🗑️', classes: 'green'});
            resetForm();
            loadPage(currentPage);
          }).deleteTransactionRow(rowIdx);
        }
      }

      function debounceSearch() {
        clearTimeout(searchTimeout);
        searchTimeout = setTimeout(function() {
          loadPage(1);
        }, 300);
      }

      function changePage(delta) {
        const newPage = currentPage + delta;
        if (newPage >= 1 && newPage <= totalPages) {
          loadPage(newPage);
        }
      }
    </script>
  </body>
</html>
```

---

## 5. Client Sidebar Code 3: `CategorizerView.html` (Categorization Assistant)

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 12px; font-family: 'Arial', sans-serif; background-color: #f8f9fa; }
      .card-header { font-weight: bold; color: #1f4e78; margin-bottom: 12px; font-size: 1.1rem; display: flex; align-items: center; justify-content: space-between; }
      .queue-card { background: white; padding: 15px; border-radius: 6px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); margin-bottom: 15px; border-top: 4px solid #1f4e78; }
      .merchant-title { font-weight: bold; font-size: 1.1rem; color: #333; margin-bottom: 4px; }
      .raw-snippet { font-size: 0.75rem; color: #888; margin-bottom: 10px; font-style: italic; word-break: break-all; }
      .amount-badge { font-size: 1.2rem; font-weight: bold; color: #db4437; margin-bottom: 12px; }
      .btn-categorize { width: 100%; background-color: #1f4e78 !important; margin-top: 10px; height: 40px; font-weight: bold; }
      .queue-counter { background: #e8eff7; color: #1f4e78; padding: 4px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: bold; }
    </style>
  </head>
  <body>
    <div class="card-header">
      <span><i class="material-icons left">psychology</i>Categorization Assistant</span>
      <span id="queueCounter" class="queue-counter">0 Left</span>
    </div>

    <div id="loader" style="text-align:center; padding:20px;">
      <div class="preloader-wrapper small active">
        <div class="spinner-layer spinner-blue-only">
          <div class="circle-clipper left"><div class="circle"></div></div>
          <div class="gap-patch"><div class="circle"></div></div>
          <div class="circle-clipper right"><div class="circle"></div></div>
        </div>
      </div>
      <p style="color:#666; font-size:0.85rem; margin-top:10px;">Scanning transactions...</p>
    </div>

    <div id="emptyState" style="display:none; text-align:center; padding:30px 10px; background:white; border-radius:6px; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
      <i class="material-icons green-text" style="font-size:3rem;">check_circle</i>
      <h6 style="font-weight:bold; color:#333; margin-top:10px;">All Transactions Categorized! 🎉</h6>
      <p style="color:#666; font-size:0.85rem;">Your ledger is clean and fully classified.</p>
    </div>

    <div id="queueContainer" style="display:none;">
      <div class="queue-card">
        <div id="cardDate" style="font-size:0.8rem; color:#666; font-weight:bold;">2026-09-09</div>
        <div id="cardMerchant" class="merchant-title">Starbucks</div>
        <div id="cardRaw" class="raw-snippet">Raw: SQ *STARBUCKS STORE 1234</div>
        <div id="cardAmount" class="amount-badge">$5.75</div>

        <div class="input-field" style="margin-top:0;">
          <select id="catSelect" class="browser-default">
            <option value="">Select Category...</option>
          </select>
        </div>

        <p style="margin-top:10px; margin-bottom:10px;">
          <label>
            <input type="checkbox" id="checkSaveRule" class="filled-in" checked="checked" />
            <span style="font-size:0.85rem; color:#333;">Save rule for keyword: <strong id="ruleKeyword">starbucks</strong></span>
          </label>
        </p>

        <button class="btn waves-effect waves-light btn-categorize" onclick="submitCategory()">
          Categorize & Apply Rule
          <i class="material-icons right">check</i>
        </button>
      </div>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script>
      let queueData = [];
      let categoriesList = [];
      let currentIndex = 0;

      document.addEventListener('DOMContentLoaded', function() {
        loadQueue();
      });

      function loadQueue() {
        google.script.run.withSuccessHandler(function(res) {
          queueData = res.queue || [];
          categoriesList = res.categories || [];
          
          const select = document.getElementById('catSelect');
          select.innerHTML = '<option value="">Select Category...</option>';
          categoriesList.forEach(c => {
            select.innerHTML += `<option value="${c}">${c}</option>`;
          });

          document.getElementById('loader').style.display = 'none';
          if (queueData.length === 0) {
            document.getElementById('emptyState').style.display = 'block';
            document.getElementById('queueContainer').style.display = 'none';
            document.getElementById('queueCounter').innerText = "0 Left";
          } else {
            document.getElementById('emptyState').style.display = 'none';
            document.getElementById('queueContainer').style.display = 'block';
            currentIndex = 0;
            renderCard();
          }
        }).getUncategorizedQueue();
      }

      function renderCard() {
        if (currentIndex >= queueData.length) {
          loadQueue();
          return;
        }

        const item = queueData[currentIndex];
        document.getElementById('queueCounter').innerText = `${queueData.length - currentIndex} Left`;
        document.getElementById('cardDate').innerText = item.date;
        document.getElementById('cardMerchant').innerText = item.cleanMerchant;
        document.getElementById('cardRaw').innerText = "Raw: " + item.rawMerchant;
        document.getElementById('cardAmount').innerText = "$" + item.amount.toFixed(2);
        
        const select = document.getElementById('catSelect');
        select.value = item.suggestedCategory || "";

        const keyword = item.cleanMerchant.toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
        document.getElementById('ruleKeyword').innerText = keyword || item.cleanMerchant.toLowerCase();
      }

      function submitCategory() {
        if (currentIndex >= queueData.length) return;

        const item = queueData[currentIndex];
        const category = document.getElementById('catSelect').value;
        const saveRule = document.getElementById('checkSaveRule').checked;
        const keyword = document.getElementById('ruleKeyword').innerText;

        if (!category) {
          M.toast({html: 'Please select a category', classes: 'orange'});
          return;
        }

        M.toast({html: 'Categorizing & learning rule... ⏳', classes: 'blue'});
        google.script.run.withSuccessHandler(function(res) {
          M.toast({html: 'Categorized successfully! 🎉', classes: 'green'});
          currentIndex++;
          renderCard();
        }).applyCategoryWithRule(item.rowIdx, category, keyword, saveRule);
      }
    </script>
  </body>
</html>
```


---

## 7. Certificate of Deposit (CD) Portfolio & Renewable Interest Engine (`CD_Tracker`)

To track fixed-income Certificates of Deposit across multiple banking institutions with varying renewal APYs, V14 introduces the **CD Portfolio Engine**:

### A. How Variable Rates & Auto-Renewals Work
1. **Initial Interest Period**: During the active term, compound yield accrues based on `Initial APY` and initial `Issue Date`.
2. **Auto-Renewal Rollover**: If `Auto-Renew?` is set to `YES`, when `TODAY()` passes the `Maturity Date`, the principal is automatically updated to include the accrued interest, and compound yield transitions to calculating accrued interest using `Renewal APY`.
3. **Dynamic Formula**:
   ```excel
   =IF(D3="", 0, IF(TODAY()<=F3, C3*(1 + G3*(MAX(0, TODAY()-D3)/365)), IF(H3="YES", (C3*(1+G3*(E3/12)))*(1 + I3*(MAX(0, TODAY()-F3)/365)), C3*(1+G3*(E3/12)))))
   ```

---

## 8. Routing Trip Transactions & Auto-Generated Travel Agendas (`Agenda_[TripName]`)

To cleanly isolate travel expenses from daily household metrics:
1. **Trip Tagging**: Use `Category_Rules_DB` or the **Trip / Project Expense Tracker** (`promptTripBudget`) to tag trip transactions.
2. **Auto-Generated Itinerary Tab**: Select **`Financial Automation 🚀` > `➕ Create Travel Itinerary & Agenda`** to build a structured `Agenda_[TripName]` sheet with sections for:
   * **Flight & Transportation Details** (Carrier, Confirmation Code, Flight #, Departure/Arrival).
   * **Accommodation & Lodging** (Hotel/Airbnb Name, Check-In/Check-Out, Nightly Rates).
   * **Daily Activity Schedule** (Time, Location, Reservation Details, Costs).

---

## 15. Adding & Tracking Multiple Certificates of Deposit (CDs)

Your automation suite provides **two seamless methods** to add and manage as many Certificate of Deposit investments as you need across different banking institutions:

### Method 1: Interactive Quick-Add Toolbar Tool (`➕ Add CD Investment`)
1. Open your Google Sheet and click **`Financial Automation 🚀` > `➕ Add CD Investment`**.
2. An interactive prompt window will open. Enter your CD details separated by commas:
   `Bank Name, Cert #, Principal ($), Issue Date (YYYY-MM-DD), Term (Months), Initial APY (%), Auto-Renew (YES/NO), Renewal APY (%)`
   * **Example:** `Marcus, CD-8812, 10000, 2026-01-15, 12, 5.25, YES, 4.50`
3. Click **OK**. The script automatically:
   * Inserts a new row in your `CD_Tracker` table before the total summary row.
   * Calculates the Maturity Date using `=EDATE(Issue_Date, Term)`.
   * Configures dynamic daily compounding and auto-renewal transition formulas.
   * Formats numbers ($ currency, YYYY-MM-DD dates, APY percentages).
   * Updates the portfolio total summary row and regenerates the **CD Portfolio Comparison Chart**!

---

### Method 2: Direct Sheet Entry on `CD_Tracker`
You can also type new CDs directly into empty rows on the **`CD_Tracker`** sheet tab:
1. Open the **`CD_Tracker`** tab.
2. In a blank row under your existing CDs (e.g. Row 6, Row 7, etc.), fill in:
   * **Col A (Bank / Institution):** e.g., `Discover Bank`
   * **Col B (Certificate #):** e.g., `CD-9041`
   * **Col C (Principal Deposit):** e.g., `15000.00`
   * **Col D (Issue Date):** e.g., `2026-01-10`
   * **Col E (Term in Months):** e.g., `18`
   * **Col F (Maturity Date):** Type `=EDATE(D6, E6)`
   * **Col G (Initial APY %):** e.g., `5.00%` (or `0.05`)
   * **Col H (Auto-Renew?):** Select `YES` or `NO` from the dropdown selector.
   * **Col I (Renewal APY %):** e.g., `4.25%` (or `0.0425`)
   * **Col J (Current Value):** Copy formula `=IF(D6="", 0, IF(TODAY()<=F6, C6*(1 + G6*(MAX(0, TODAY()-D6)/365)), IF(H6="YES", (C6*(1+G6*(E6/12)))*(1 + I6*(MAX(0, TODAY()-F6)/365)), C6*(1+G6*(E6/12)))))`
   * **Col K (Maturity Payout):** Copy formula `=C6*(1 + G6*(E6/12))`
   * **Col L (Status):** Copy formula `=IF(TODAY()<F6, "ACTIVE 🟢", IF(H6="YES", "RENEWED 🔄", "MATURED 🟡"))`
3. Update the summary totals row at the bottom `=SUM(C3:C7)` so all your CD deposits and accrued yield remain completely synchronized!
