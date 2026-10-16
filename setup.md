# Unified Financial Automation Setup Guide - V6 (Standardized Edition)

This comprehensive guide serves as the ultimate manual for configuring, testing, and expanding your **Enterprise-Grade, Relational Personal Finance Ledger System**. This workbook is engineered to comply with the **FAST Modeling Standard** (Flexibility, Appropriateness, Structure, Transparency) and the **ICAEW Spreadsheet Design Principles**.

---

---

## 1. Resolving Key Architecture Bugs

### A. The "Double Creation & Multiple URLs" Bug
**The Cause:** In Google Apps Script, all `.gs` files in the left-hand sidebar are compiled together inside a single V8 Javascript runtime. If you pasted the previous versions of the setup code into separate script files (e.g. `main.gs` and `Setup.gs` or `user.gs` or `code.gs`) without clearing your old scripts, **multiple declarations of `onOpen()` and `createNewSpreadsheet()` were active simultaneously**.
* Whichever trigger compiled last overrode the toolbar menu.
* Clicking "Generate Automated Sheet" ran both files' setup logic, creating **two completely separate Google Sheets files back-to-back with two different URL popups**.

**The Fix:** 
1. Delete all legacy script files in your Apps Script editor. You must have only **one** backend file (name it `main.gs`) and copy the code of **`unified-financial-automation-v6.gs`** into it.
2. We have updated your toolbar menu to replace `createNewSpreadsheet()` with the much smarter, highly requested option: **`Initialize Sheets in This Workbook`**.

---

### B. The "Missing Scripts inside Generated Sheet" Constraint
**The Cause:** When Google Apps Script runs `SpreadsheetApp.create("My Automated Spreadsheet")`, it spawns a *completely blank new file* on your Google Drive. Because Google Sheet bound scripts are sandboxed inside their parent file, programmatically spawned sheets **never inherit your Apps Script triggers, forms, HTML sidebars, or automated parsers**. They are simply blank templates.

**The Fix:** 
Instead of spawning a new sheet, run our new command: **`Initialize Sheets in This Workbook`** (`promptInPlaceInitialization`).
* This programmatically builds, resets, styles, and links all required database tabs (**`About`**, **`Enum`**, **`Transactions_DB`**, **`Budget_History_DB`**, **`The Monthly Dashboard Tab`**, **`Control`**, **`Category_Rules_DB`**, **`Savings_Goals`**) **directly inside the active workbook you are currently viewing!**
* Since the sheets are initialized in-place, all of your background Gmail scrapers, daily budget breach notifications, HTML manual entry APIs, and sidebar interfaces remain **fully connected, active, and scripted** with zero file transfers!

---

### C. The "Charts Setup Failed / setChartType" Error
**The Cause:** In Google Sheets' standard V8 runtime, `Charts.ChartType.DONUT` is an invalid enum value. It throws an exception, crashes the programmatic chart engine, and skips inserting both the Column Chart and the Donut Chart.
* **The Fix:** We have programmatically updated the category expense breakdown chart type to `Charts.ChartType.PIE` and configured the `.setOption('pieHole', 0.4)` parameter. This creates a beautiful, native donut chart that is fully compatible with Google Sheets' chart engine!

---

### D. Double Confirmation Pop-ups on Sheet Generation
**The Cause:** Previously, clicking "Generate Automated Sheet" triggered a basic synchronous blocking `ui.alert()` box *and then* immediately opened an HTML modal dialog, causing duplicate popup clicks.
* **The Fix:** We removed the redundant basic alert and consolidated the success notification into a single HTML modal that contains a direct clickable link to your new spreadsheet.

---

## 2. Advanced Multi-Year Savings Projection Model & Visual Chart

We have programmatically structured a powerful **Interactive Long-Term Savings Planner** inside your new **`Savings_Goals`** tab:

### A. Editable Parameter Inputs
You can dynamically forecast your compound wealth accumulation using the two editable parameter cells in columns H and I:
*   **Cell `I2` (Assumed Annual Growth Rate):** Defaults to `6.0%`. You can change this to model different investment returns (e.g., `4%` for conservative bonds, `8-10%` for standard stock portfolios).
*   **Cell `I3` (Annual Goals Savings Contribution):** Defaults to `$6,000.00` ($500/month). Adjust this to match your annual target savings velocity.

### B. 10-Year Dynamic Compound Interest Ledger
Starting in Column H, Row 5, the script designs a 10-year growth matrix (from Year 2026 through 2035) with the following relational formulas:
*   **Starting Balance (Year 1):** `=SUM(C3:C5)` - Automatically aggregates your live, current balances across all defined goals!
*   **Starting Balance (Subsequent Years):** Relinked to the previous year's ending balance (e.g., `=L6` for row 7).
*   **Annual Savings Addition:** Linked directly to your input cell `=$I$3`.
*   **Compound Interest Growth:** Calculates interest based on your editable rate cell `=(I6+J6)*$I$2`.
*   **Projected Ending Balance:** Integrates additions and compound interest `=SUM(I6:K6)`.

### C. Compound Interest Graph (Visual Area Chart)
To satisfy **ICAEW Principle 19: Visualizing performance**, a dynamic **10-Year Growth Area Chart** is programmatically constructed in cell `N1` on your `Savings_Goals` tab. 
* Because this chart is natively linked to the dynamic formula cells, changing your annual contribution in cell `I3` or your investment return rate in cell `I2` will **instantly update the area chart's compound curve in real-time** with absolute zero script latency!

---

## 3. Database-Driven AI Auto-Categorization ruleset (`Category_Rules_DB`)

To keep your categorization rules flexible and avoid editing backend script code, we have introduced a **No-Code Rules Engine**:

### A. How It Works
*   We created a dedicated database tab called **`Category_Rules_DB`**. Column A houses `Merchant Keyword Rule` (e.g., *starbucks*, *grubhub*, *wholefd*) and Column B houses `Budget Category Target` (e.g., *Food & Drink*, *Groceries*).
*   When a transaction is imported via your manual HTML form, Gmail scrapper, or Chase CSV folder, the script searches this sheet first. It runs a case-insensitive, fuzzy text check on the raw merchant descriptor against your keyword list. If a match is found, the transaction is **instantly sorted** into the correct category!
*   If no keyword rules match, the script automatically defaults to a built-in static dictionary, and then falls back to `Miscellaneous` to prevent crashes.

### B. How to Use It (Zero-Code Rules Maintenance)
To categorize a new recurring merchant, you do **not** need to touch Apps Script. Simply:
1.  Open the **`Category_Rules_DB`** sheet.
2.  Add a keyword rule (e.g., `spotify` in Column A, and choose `Entertainment` from the dropdown in Column B).
3.  Next time an import runs, it will categorize Spotify automatically!
4.  To retroactively clean your ledger, click **`Financial Automation 🚀` > `Run AI Auto-Categorization`** from your toolbar.

---

## 4. Setting Up Your Background Scheduled Automation Triggers

To make your system completely autonomous, set up these three scheduled background triggers inside Google Apps Script:

### A. Hourly Gmail Scraper Trigger (`pullGmailTransactionAlerts`)
*   **Objective:** Scrapes unread Chase purchase alerts, normalizes signs, deduplicates against Gmail Message IDs, categorizes them using your `Category_Rules_DB` tab, and appends them to your ledger.
*   **Trigger Config:** Function: `pullGmailTransactionAlerts` | Source: `Time-driven` | Type: `Hour timer` | Interval: `Every hour`.

### B. Daily Budget Breach Inspector Trigger (`checkBudgetBreaches`)
*   **Objective:** Scans actual spending vs planned budgets on your Monthly Dashboard. If a limit is breached, it dispatches an email alert.
*   **Spam Prevention Safeguard:** Uses `PropertiesService` background state-tracking. You will receive **only one email alert per category per month**, even if you record 10 more transactions in that category!
*   **Trigger Config:** Function: `checkBudgetBreaches` | Source: `Time-driven` | Type: `Day timer` | Time: `8 AM to 9 AM` (or daily morning window).

### C. Weekly HTML Budget Summary Digest Trigger (`sendWeeklyBudgetSummary`)
*   **Objective:** Dispatches a beautiful, responsive HTML briefing newsletter to your inbox every Monday morning. It outlines high-level income/expense KPIs, net savings, and category velocity statuses (**`⚠️ Overrun`**, **`⚡ Fast Velocity`**, or **`On Track`**) relative to how much of the calendar month has elapsed.
*   **Trigger Config:** Function: `sendWeeklyBudgetSummary` | Source: `Time-driven` | Type: `Weekly timer` | Day: `Every Monday` | Time: `7 AM to 8 AM`.

---

## 5. Client Sidebar Template File (`SidebarView.html`)

Create a separate **HTML** file in your Google Apps Script editor, name it exactly **`SidebarView`** (omit the `.html` extension as Apps Script appends it), and paste this Materialize CSS budget manager interface:

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!-- Load Google Material Icons & Materialize CSS -->
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

    <!-- Load Materialize JS -->
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

*(Note: We added a backend helper function `getDashboardPeriod()` in `unified-financial-automation-v6.gs` to feed the exact selected dropdown year and month directly to this HTML view upon launching!)*
