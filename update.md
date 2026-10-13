# Advanced Financial Automation Extensions: Weekly Summaries, Savings Tracking & Hands-Off CSV Ingestion

This guide outlines three advanced modules to elevate your personal financial workbook under the **FAST Modeling Standard** and **ICAEW Spreadsheet Principles**. 

---

## 1. Weekly Budget Summary Email Engine

To maintain high visual engagement and professional compliance, we establish a **Weekly HTML Digest Engine**. This background script compiles actual spending against monthly boundaries, computes progress ratios, and dispatches a responsive email newsletter to your inbox.

### Step 1: Add the Apps Script Code
Add this function to the bottom of your `main.gs` / `unified-financial-automation-v3.gs` file:

```javascript
/**
 * Compiles a weekly budget summary report and dispatches a rich HTML email.
 * Evaluates monthly spending velocity and category status.
 * Enforces FAST compliance by referencing normalized database totals.
 */
function sendWeeklyBudgetSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  const enumSheet = ss.getSheetByName("Enum");
  if (!dash || !enumSheet) {
    Logger.log("Dashboard or Enum sheet not found. Aborting summary.");
    return;
  }

  const year = dash.getRange("C4").getValue().toString().trim();
  const month = dash.getRange("D4").getValue().toString().trim();
  if (month === "All") {
    Logger.log("Skipping weekly summary for cumulative 'All' month selector.");
    return;
  }

  const userEmail = Session.getActiveUser().getEmail();
  const lastEnumRow = enumSheet.getLastRow();
  const categoryCount = lastEnumRow - 2;

  // Retrieve KPI Block values
  const totalIncome = parseFloat(dash.getRange("A7").getValue() || 0);
  const totalSpending = parseFloat(dash.getRange("B7").getValue() || 0);
  const netSavings = parseFloat(dash.getRange("C7").getValue() || 0);
  const savingsRate = (dash.getRange("D7").getValue() * 100).toFixed(1);

  // Retrieve budget items (Category, Planned, Actual, Diff)
  const budgetData = dash.getRange(10, 1, categoryCount, 4).getValues();

  // Determine current day status (e.g. Progress through the month)
  const today = new Date();
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const currentDay = today.getDate();
  const monthElapsedPercent = ((currentDay / daysInMonth) * 100).toFixed(0);

  // Construct HTML Newsletter
  let html = `
  <div style="font-family: Arial, sans-serif; background-color: #f4f6f9; padding: 20px; max-width: 600px; margin: auto; border: 1px solid #e1e4e8;">
    <!-- Header Block -->
    <div style="background-color: #1f4e78; color: white; padding: 25px; text-align: center; border-radius: 5px 5px 0 0;">
      <h2 style="margin: 0; font-size: 22px; letter-spacing: 1px;">WEEKLY FINANCIAL BRIEFING</h2>
      <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Period Status: ${month} ${year} (${monthElapsedPercent}% Elapsed)</p>
    </div>

    <!-- KPI Summary Block -->
    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Month-to-Date KPIs</h3>
      <table style="width: 100%; text-align: center; border-collapse: collapse;">
        <tr>
          <td style="padding: 10px; width: 25%;">
            <span style="font-size: 11px; color: #666; display: block; text-transform: uppercase;">Total Income</span>
            <span style="font-weight: bold; font-size: 16px; color: #385723;">$${totalIncome.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block; text-transform: uppercase;">Total Spending</span>
            <span style="font-weight: bold; font-size: 16px; color: #db4437;">$${totalSpending.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block; text-transform: uppercase;">Net Savings</span>
            <span style="font-weight: bold; font-size: 16px; color: ${netSavings >= 0 ? '#385723' : '#c00000'};">$${netSavings.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block; text-transform: uppercase;">Savings Rate</span>
            <span style="font-weight: bold; font-size: 16px; color: #1f4e78;">${savingsRate}%</span>
          </td>
        </tr>
      </table>
    </div>

    <!-- Category Variance Analysis -->
    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Budget Performance Breakdown</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <thead>
          <tr style="background-color: #f8f9fa; color: #333; text-align: left; font-weight: bold;">
            <th style="padding: 8px; border-bottom: 1px solid #e1e4e8;">Category</th>
            <th style="padding: 8px; border-bottom: 1px solid #e1e4e8; text-align: right;">Planned</th>
            <th style="padding: 8px; border-bottom: 1px solid #e1e4e8; text-align: right;">Actual</th>
            <th style="padding: 8px; border-bottom: 1px solid #e1e4e8; text-align: right;">Status</th>
          </tr>
        </thead>
        <tbody>`;

  budgetData.forEach(row => {
    const category = row[0];
    const planned = parseFloat(row[1] || 0);
    const actual = parseFloat(row[2] || 0);
    const progressPercent = planned > 0 ? (actual / planned) * 100 : 0;
    
    let statusText = "";
    let statusColor = "#4285f4"; // Blue standard

    if (planned === 0) {
      statusText = "Unplanned";
      statusColor = "#777777";
    } else if (actual > planned) {
      statusText = "⚠️ Overrun";
      statusColor = "#db4437";
    } else if (progressPercent > monthElapsedPercent) {
      statusText = "⚡ Fast Velocity";
      statusColor = "#ff9900";
    } else {
      statusText = "✅ On Track";
      statusColor = "#0f9d58";
    }

    html += `
      <tr style="border-bottom: 1px solid #f0f0f0;">
        <td style="padding: 8px; font-weight: bold; color: #333;">${category}</td>
        <td style="padding: 8px; text-align: right; color: #002060;">$${planned.toFixed(2)}</td>
        <td style="padding: 8px; text-align: right; color: #333;">$${actual.toFixed(2)}</td>
        <td style="padding: 8px; text-align: right; font-weight: bold; color: ${statusColor};">${statusText}</td>
      </tr>`;
  });

  html += `
        </tbody>
      </table>
    </div>

    <!-- Call to Action Block -->
    <div style="margin-top: 20px; text-align: center; color: #666; font-size: 12px; line-height: 1.5;">
      <p>This report compiles live rows from your <b>Transactions_DB</b> tab.</p>
      <p style="margin-top: 10px;"><i>Dispatched automatically by your Google Sheets Financial Automation System 🚀</i></p>
    </div>
  </div>`;

  MailApp.sendEmail({
    to: userEmail,
    subject: `📊 Weekly Financial Summary: ${month} ${year} Briefing`,
    htmlBody: html
  });

  Logger.log("Weekly budget summary email successfully dispatched.");
}
```

### Step 2: Configure the Time-Driven Trigger
To set this to run automatically every week without touching a single button:
1. In your Apps Script editor, click the **Triggers (alarm clock icon)** on the left sidebar.
2. Click **`+ Add Trigger`** in the bottom-right corner.
3. Select `sendWeeklyBudgetSummary` as the function to run.
4. Set the event source to **`Time-driven`**.
5. Set the type of trigger to **`Weekly timer`**.
6. Select your preferred day and time (e.g., **`Every Monday morning between 7 AM and 8 AM`**).
7. Save the trigger.

---

## 2. Visual Savings Goal Progress Tracker

To keep your personal goals cleanly separated from transactional balances, we establish a dedicated **`Savings_Goals`** tracking tab. This module provides a complete system to set goals, track saved sums, compute days remaining, and display dynamic progress meters.

### Step 1: Add the Goal Table Creation Code
Add this function to the bottom of your Apps Script editor to let your spreadsheet programmatically build your Savings tab:

```javascript
/**
 * Programmatically creates and structures a dedicated 'Savings_Goals' tab in your sheet.
 * Generates dynamic calculations and beautiful visual sparklines for visual pacing.
 * Enforces relational normalization and FAST standards.
 */
function setupSavingsGoalTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let goalsSheet = ss.getSheetByName("Savings_Goals");
  
  if (!goalsSheet) {
    goalsSheet = ss.insertSheet("Savings_Goals");
  } else {
    goalsSheet.clear();
  }
  
  // Style Master Title Block
  const titleRange = goalsSheet.getRange("A1:G1");
  titleRange.merge();
  titleRange.setValue("SAVINGS GOALS PROGRESS PORTAL");
  titleRange.setFontFamily("Arial")
            .setFontSize(14)
            .setFontWeight("bold")
            .setFontColor("#ffffff")
            .setBackground("#1f4e78")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
  goalsSheet.setRowHeight(1, 40);

  // Set Headers (Row 3)
  const headers = [["Goal Name", "Target Amount", "Current Balance", "Progress %", "Visual Bar Tracker", "Target Date", "Days Remaining"]];
  goalsSheet.getRange("A3:G3").setValues(headers).setFontWeight("bold");
  
  // Add some default dummy goals
  const initialGoals = [
    ["Emergency Fund (6 Months)", 10000.00, 2500.00, "=C4/B4", `=IF(D4>0, SPARKLINE(C4, {"charttype","bar";"max", B4;"color1", "#0f9d58"}), "")`, "2027-12-31", "=MAX(0, F4-TODAY())"],
    ["Travel Fund (Japan 2027)", 5000.00, 1200.00, "=C5/B5", `=IF(D5>0, SPARKLINE(C5, {"charttype","bar";"max", B5;"color1", "#4285f4"}), "")`, "2027-04-30", "=MAX(0, F5-TODAY())"],
    ["Crypto & Stock Portfolio", 15000.00, 7500.00, "=C6/B6", `=IF(D6>0, SPARKLINE(C6, {"charttype","bar";"max", B6;"color1", "#1f4e78"}), "")`, "2028-06-30", "=MAX(0, F6-TODAY())"]
  ];
  
  goalsSheet.getRange(4, 1, initialGoals.length, 7).setFormulasAndValues(initialGoals);
  
  // Apply formatting
  goalsSheet.getRange("B4:C100").setNumberFormat("$#,##0.00").setFontColor("#002060"); // Input colors
  goalsSheet.getRange("D4:D100").setNumberFormat("0.0%").setFontWeight("bold");
  goalsSheet.getRange("F4:F100").setNumberFormat("yyyy-mm-dd");
  goalsSheet.getRange("G4:G100").setNumberFormat("#,##0").setFontColor("#666666");

  // Summary Row (Row 7)
  goalsSheet.getRange("A7:G7").setValues([["Total Consolidated Savings", "=SUM(B4:B6)", "=SUM(C4:C6)", "=C7/B7", "", "", ""]]).setFontWeight("bold");
  goalsSheet.getRange("B7:C7").setNumberFormat("$#,##0.00");
  goalsSheet.getRange("D7").setNumberFormat("0.0%");
  goalsSheet.getRange("A7:G7").setBackground("#f2f2f2");

  // Style Header Cells
  goalsSheet.getRange("A3:G3").setFontColor("#000000").setBackground("#d9e1f2").setHorizontalAlignment("center");
  goalsSheet.setRowHeight(3, 25);
  
  goalsSheet.autoResizeColumns(1, 7);
  
  // Alert User
  const ui = SpreadsheetApp.getUi();
  ui.alert("Savings Goal Tracker Setup! 🎯", "A beautiful savings tracker sheet has been programmatically created and formatted under the 'Savings_Goals' tab. Check it out!", ui.ButtonSet.OK);
}
```

*To integrate this function into your custom spreadsheet menu, update your unified `onOpen()` function in Apps Script to include:*
```javascript
  ui.createMenu("Financial Automation 🚀")
    .addItem("Generate Automated Sheet", "createNewSpreadsheet")
    .addSeparator()
    .addItem("Batch Import CSV Folder", "importAllPendingCSVs")
    .addItem("Build Savings Goal Tracker", "setupSavingsGoalTracker") // Added item
    .addItem("Open Budget History Sidebar", "showBudgetSidebar")
    .addToUi();
```

---

## 3. Next Steps to 100% Fully Automate Chase CSV Imports

While your new `importAllPendingCSVs()` function handles multi-file ingestion, archiving, and deduplication perfectly, you still have to manually log into your bank, download the file, and click the script menu. 

To reach a **"Zero-Click" fully autonomous ingestion pipeline** for your Chase CSVs, execute these three tactical implementation steps:

### 1. Zero-Click File Delivery (Google Drive for Desktop Sync)
To eliminate the step of opening Google Drive in your web browser to upload statement files:
1. Download and install **Google Drive for Desktop** on your computer.
2. Once installed, Google Drive will appear as a virtual hard drive on your computer (typically as `G:`).
3. Navigate into your Drive folder and locate the **`Chase_CSV_Uploads`** folder generated by your Apps Script.
4. Pin or create a desktop shortcut to this folder.
5. **The Streamlined Workflow:** Whenever you download a statement from Chase, save it directly to this desktop shortcut. Drive for Desktop will silently sync it to the cloud folder in less than 2 seconds with zero browser clicks required.

### 2. Time-Driven Background Trigger Execution
To eliminate the requirement of manually clicking "Batch Import CSV Folder" in Sheets:
1. In the Apps Script sidebar, click the **Triggers icon** (alarm clock).
2. Click **`+ Add Trigger`**.
3. Select **`importAllPendingCSVs`** as the target function.
4. Set event source to **`Time-driven`**.
5. Set type of trigger to **`Day timer`** and choose **`Every night between midnight and 1 AM`** (or set an **`Hour timer`** if you upload statements frequently).
6. Save the trigger.

*When the trigger fires, the script scans the Drive folder silently in the background, appends the transactions to your sheet, deduplicates overlaps, and archives the processed statements. You will simply wake up to an updated dashboard.*

### 3. Automatic Ingestion Auto-Categorization Chain
Currently, newly imported transactions on `Transactions_DB` default to the `"Miscellaneous"` category. To automate the pipeline all the way through to your dashboard charts:
Update your `importAllPendingCSVs()` code in Apps Script so that right after parsing and appending rows, it automatically triggers the AI auto-categorization engine. Add this single function call at the very end of your file loop:
```javascript
// Automatically execute categorization rules over newly imported rows
autoCategorizeAllTransactions();
```
*This binds your CSV data parser directly to your merchant categorization rules, meaning new expenses are instantly mapped to categories (e.g. "Whole Foods" -> "Groceries") and reflected on your charts instantly with no manual categorization required.*
