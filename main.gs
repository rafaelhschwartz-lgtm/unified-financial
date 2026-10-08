/**
 *  ===========================================================================
 *  GOOGLE SHEETS FINANCIAL AUTOMATION SYSTEM (UNIFIED ENGINE)
 *  Standardized under the FAST Modeling Standard and ICAEW Spreadsheet Principles.
 *  
 *  This code consolidates all your automation channels, triggers, and layouts
 *  into a single, high-performance master script to prevent collisions.
 *  ===========================================================================
 */

// ===========================================================================
// 1. CONSOLIDATED TRIGGER SYSTEM (onOpen & onEdit)
// ===========================================================================

/**
 *  Consolidated onOpen trigger (Resolves prior triggers collision).
 *  Runs automatically when the spreadsheet is opened.
 *  Builds a unified custom menu for complete operations.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Financial Automation 🚀")
    .addItem("Generate Automated Sheet", "createNewSpreadsheet")
    .addSeparator()
    .addItem("Import Chase CSV File", "promptForChaseCSV")
    .addItem("Open Budget History Sidebar", "showBudgetSidebar")
    .addSeparator()
    .addItem("Run AI Auto-Categorization", "autoCategorizeUnassigned")
    .addItem("Check Budget Breaches", "checkBudgetBreaches")
    .addToUi();
}

/**
 *  Consolidated onEdit background trigger.
 *  Automates refreshing dashboard budgets when period selection changes (C4/D4).
 */
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // Watch Year selector (C4) and Month selector (D4) on the Dashboard
  if (sheetName === "The Monthly Dashboard Tab") {
    const row = range.getRow();
    const col = range.getColumn();
    if ((row === 4 && col === 3) || (row === 4 && col === 4)) {
      try {
        syncDashboardBudgets(sheet);
      } catch (err) {
        console.error("Failed to sync historical budgets: " + err.toString());
      }
    }
  }
}

// ===========================================================================
// 2. MAIN SHEET COMPILER (DYNAMIC RANGES)
// ===========================================================================

/**
 *  Creates a programmatically generated master spreadsheet containing the entire architecture.
 *  Includes safety wrappers to avoid headless execution failures.
 */
function createNewSpreadsheet() {
  const ss = SpreadsheetApp.create("My Automated Spreadsheet");
  
  // Build and link the entire architecture (Passing 'ss' directly)
  createWholeBook(ss);
  
  const url = ss.getUrl();
  Logger.log("Your new spreadsheet is ready at: " + url);
  
  // Safely trigger the pop-up/modal dialog
  try {
    const ui = SpreadsheetApp.getUi();
    const alertTitle = "Workbook Successfully Created! 🚀";
    const alertMessage = "Your automated spreadsheet is ready.\n\n" +
                         "Copy the URL below to open your new file:\n\n" + url + "\n\n" +
                         "💡 Don't forget to configure your conditional formatting on the 'Control' sheet!";
    ui.alert(alertTitle, alertMessage, ui.ButtonSet.OK);
    showUrlModal(url);
  } catch (e) {
    Logger.log("Notice: Pop-up modal could not be displayed because the script is running without an active browser UI context.");
    Logger.log("You can still access your compiled spreadsheet using the URL logged above.");
  }
}

/**
 *  Standardized Workbook Setup Script under FAST and ICAEW Principles.
 *  Dynamically sizes category listings and configures all inter-sheet relationships.
 */
function createWholeBook(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("No spreadsheet found. Ensure a spreadsheet is created or active.");
  }
  
  // Define required sheets (Includes the historical budget database)
  const sheetNames = ["About", "Enum", "Transactions_DB", "Budget_History_DB", "The Monthly Dashboard Tab", "Control"];
  const sheets = {};
  
  // Reset or create each sheet
  sheetNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clear();
      if (sheet.getFilter()) sheet.getFilter().remove();
      sheet.getDataRange().setDataValidation(null);
    }
    sheets[name] = sheet;
  });
  
  // Chronological arrangement by functional class
  sheets["About"].activate(); ss.moveActiveSheet(1);
  sheets["Enum"].activate(); ss.moveActiveSheet(2);
  sheets["Transactions_DB"].activate(); ss.moveActiveSheet(3);
  sheets["Budget_History_DB"].activate(); ss.moveActiveSheet(4);
  sheets["The Monthly Dashboard Tab"].activate(); ss.moveActiveSheet(5);
  sheets["Control"].activate(); ss.moveActiveSheet(6);
  
  // Helper styling functions
  function formatHeader(sheet, rangeString, title) {
    const range = sheet.getRange(rangeString);
    range.merge();
    range.setValue(title);
    range.setFontFamily("Arial")
         .setFontSize(14)
         .setFontWeight("bold")
         .setFontColor("#ffffff")
         .setBackground("#1f4e78") // Classic corporate deep blue
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
    sheet.setRowHeight(range.getRow(), 40);
  }
  
  function formatSubHeader(sheet, rangeString) {
    const range = sheet.getRange(rangeString);
    range.setFontFamily("Arial")
         .setFontSize(10)
         .setFontWeight("bold")
         .setFontColor("#000000")
         .setBackground("#d9e1f2") // Accent light blue
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
    sheet.setRowHeight(range.getRow(), 25);
  }
  
  // -------------------------------------------------------------------------
  // ABOUT TAB
  // -------------------------------------------------------------------------
  const about = sheets["About"];
  formatHeader(about, "A1:D1", "GOOGLE SHEETS FINANCIAL AUTOMATION SYSTEM");
  about.getRange("A3").setValue("System Overview").setFontWeight("bold").setFontSize(12);
  about.getRange("A4").setValue("This workbook is engineered under the FAST Financial Modeling Standard and the ICAEW Spreadsheet Principles. It decouples raw transactional ledgers (Transactions_DB & Budget_History_DB) from the dynamic monthly dashboard.");
  
  about.getRange("A6").setValue("Functional Color Code Guidelines").setFontWeight("bold").setFontSize(11);
  about.getRange("A7:B7").setValues([["Blue Text:", "Hardcoded manual input variables"]]);
  about.getRange("A7").setFontColor("#002060").setFontWeight("bold");
  about.getRange("A8:B8").setValues([["Black Text:", "Formulas and local calculation cells"]]);
  about.getRange("A8").setFontColor("#000000").setFontWeight("bold");
  about.getRange("A9:B9").setValues([["Green Text:", "Cross-sheet references and linkages"]]);
  about.getRange("A9").setFontColor("#385723").setFontWeight("bold");
  about.getRange("A10:B10").setValues([["Red Text:", "Warning/Error checking indicators"]]);
  about.getRange("A10").setFontColor("#c00000").setFontWeight("bold");
  
  about.getRange("A12").setValue("Operational Instructions").setFontWeight("bold").setFontSize(11);
  about.getRange("A13").setValue("1. Maintain standard Category and Account options on the 'Enum' tab.");
  about.getRange("A14").setValue("2. Log manual entries via the HTML web form, Gmail automation, or paste them directly to 'Transactions_DB' using positive magnitudes.");
  about.getRange("A15").setValue("3. Review high-level budget variances, KPIs, and actual spending patterns on 'The Monthly Dashboard Tab'.");
  about.getRange("A16").setValue("4. Click 'Financial Automation 🚀' -> 'Open Budget History Sidebar' to manage historical monthly budget ceilings.");
  about.getRange("A17").setValue("5. Check the 'Control' sheet's master indicator to ensure zero mathematical anomalies (Sign breaches, blank categories, or invalid flow types).");
  about.autoResizeColumns(1, 4);
  
  // -------------------------------------------------------------------------
  // ENUM TAB
  // -------------------------------------------------------------------------
  const enumSheet = sheets["Enum"];
  formatHeader(enumSheet, "A1:B1", "SYSTEM ENUMS / CATEGORIES & ACCOUNTS");
  enumSheet.getRange("A2:B2").setValues([["Budget Categories", "Account Sources"]]).setFontWeight("bold");
  
  const categories = [
    ["Food & Drink"], ["Groceries"], ["Bills & Utilities"], ["Home"], 
    ["Transportation"], ["Entertainment"], ["Personal"], ["Travel"], ["Miscellaneous"]
  ];
  const accounts = [
    ["Cash"], ["Chase Checking"], ["Chase Sapphire"], ["TD Bank"], ["Credit Card"], ["Savings"]
  ];
  enumSheet.getRange(3, 1, categories.length, 1).setValues(categories);
  enumSheet.getRange(3, 2, accounts.length, 1).setValues(accounts);
  enumSheet.autoResizeColumns(1, 2);
  
  // -------------------------------------------------------------------------
  // TRANSACTIONS_DB TAB
  // -------------------------------------------------------------------------
  const db = sheets["Transactions_DB"];
  const dbHeaders = [
    "Transaction_ID", "Timestamp", "Transaction_Date", "Merchant_Raw", 
    "Merchant_Clean", "Amount", "Capital_Flow", "Category", 
    "Account_Source", "Ingestion_Channel", "Sync_Status"
  ];
  db.getRange(1, 1, 1, dbHeaders.length).setValues([dbHeaders]);
  formatSubHeader(db, "A1:K1");
  db.setRowHeight(1, 28);
  
  // Set dropdown validation dynamically pointing to Enum sheet
  const categoryRange = enumSheet.getRange("A3:A11");
  const categoryValidation = SpreadsheetApp.newDataValidation().requireValueInRange(categoryRange).setAllowInvalid(false).build();
  db.getRange("H2:H1000").setDataValidation(categoryValidation);
  
  const accountRange = enumSheet.getRange("B3:B8");
  const accountValidation = SpreadsheetApp.newDataValidation().requireValueInRange(accountRange).setAllowInvalid(false).build();
  db.getRange("I2:I1000").setDataValidation(accountValidation);
  
  const flowValidation = SpreadsheetApp.newDataValidation().requireValueInList(["INFLOW", "OUTFLOW"]).setAllowInvalid(false).build();
  db.getRange("G2:G1000").setDataValidation(flowValidation);
  
  const channelValidation = SpreadsheetApp.newDataValidation().requireValueInList(["WEB_FORM", "CHASE_CSV", "GMAIL_ALERT", "PLAID_API", "MANUAL"]).setAllowInvalid(false).build();
  db.getRange("J2:J1000").setDataValidation(channelValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation().requireValueInList(["SETTLED", "PENDING", "MANUAL_REVIEW"]).setAllowInvalid(false).build();
  db.getRange("K2:K1000").setDataValidation(statusValidation);
  
  // Standard dummy row to initialize formulas
  db.getRange("A2:K2").setValues([[
    "INIT_SETUP_DUMMY", new Date(), new Date(), "INITIAL SYSTEM SETUP", "Initial Setup", 0.01, "INFLOW", "Miscellaneous", "Cash", "MANUAL", "SETTLED"
  ]]);
  db.getRange("F2:F1000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  db.getRange("C2:C1000").setNumberFormat("yyyy-mm-dd");
  db.autoResizeColumns(1, 11);
  
  // -------------------------------------------------------------------------
  // BUDGET_HISTORY_DB TAB (HISTORICAL DATA STORE)
  // -------------------------------------------------------------------------
  const budgetDB = sheets["Budget_History_DB"];
  formatHeader(budgetDB, "A1:D1", "HISTORICAL BUDGET RECORD DATABASE");
  const budgetHeaders = [["Year", "Month", "Category", "Budget_Amount"]];
  budgetDB.getRange("A2:D2").setValues(budgetHeaders).setFontWeight("bold");
  formatSubHeader(budgetDB, "A2:D2");
  budgetDB.setRowHeight(2, 25);
  
  // Populate starting dummy allocations for year 2026, Month August (match dashboard default)
  const initialAllocations = [];
  categories.forEach(cat => {
    initialAllocations.push([2026, "August", cat[0], 0.00]); // Starts at 0, updated by user in sidebar
  });
  budgetDB.getRange(3, 1, initialAllocations.length, 4).setValues(initialAllocations);
  budgetDB.getRange("D3:D1000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  budgetDB.autoResizeColumns(1, 4);
  
  // -------------------------------------------------------------------------
  // THE MONTHLY DASHBOARD TAB (DYNAMIC COCKPIT)
  // -------------------------------------------------------------------------
  const dash = sheets["The Monthly Dashboard Tab"];
  formatHeader(dash, "A1:G1", "INTERACTIVE MONTHLY FINANCIAL DASHBOARD");
  dash.getRange("A2").setFormula('=IF(Control!$A$2="✅ SYSTEM OK", "✅ SYSTEM OK", "❌ ERROR DETECTED")');
  dash.getRange("A2").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f2f2f2");
  
  // Period Dropdown Selectors
  dash.getRange("C3").setValue("Select Year:").setFontWeight("bold");
  const yearValidation = SpreadsheetApp.newDataValidation().requireValueInList(["2024", "2025", "2026", "2027", "2028"]).setAllowInvalid(false).build();
  dash.getRange("C4").setDataValidation(yearValidation).setValue("2026").setFontColor("#002060").setFontWeight("bold");
  
  dash.getRange("D3").setValue("Select Month:").setFontWeight("bold");
  const monthValidation = SpreadsheetApp.newDataValidation().requireValueInList([
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "All"
  ]).setAllowInvalid(false).build();
  dash.getRange("D4").setDataValidation(monthValidation).setValue("August").setFontColor("#002060").setFontWeight("bold");
  
  // KPI Summary Block
  const kpiHeaders = [["Total Income", "Total Spending", "Net Savings", "Savings Rate"]];
  dash.getRange("A6:D6").setValues(kpiHeaders).setFontWeight("bold").setHorizontalAlignment("center");
  formatSubHeader(dash, "A6:D6");
  
  dash.getRange("A7").setFormula(
    '=IF($D$4="All", SUMIFS(Transactions_DB!F:F, Transactions_DB!G:G, "INFLOW", Transactions_DB!C:C, ">="&DATE($C$4, 1, 1), Transactions_DB!C:C, "<="&DATE($C$4, 12, 31)), SUMIFS(Transactions_DB!F:F, Transactions_DB!G:G, "INFLOW", Transactions_DB!C:C, ">="&DATEVALUE($D$4&" 1, "&$C$4), Transactions_DB!C:C, "<="&EOMONTH(DATEVALUE($D$4&" 1, "&$C$4), 0)))'
  );
  dash.getRange("B7").setFormula(
    '=IF($D$4="All", SUMIFS(Transactions_DB!F:F, Transactions_DB!G:G, "OUTFLOW", Transactions_DB!C:C, ">="&DATE($C$4, 1, 1), Transactions_DB!C:C, "<="&DATE($C$4, 12, 31)), SUMIFS(Transactions_DB!F:F, Transactions_DB!G:G, "OUTFLOW", Transactions_DB!C:C, ">="&DATEVALUE($D$4&" 1, "&$C$4), Transactions_DB!C:C, "<="&EOMONTH(DATEVALUE($D$4&" 1, "&$C$4), 0)))'
  );
  dash.getRange("C7").setFormula('=A7-B7');
  dash.getRange("D7").setFormula('=IFERROR(C7/A7, 0)');
  dash.getRange("A7:C7").setNumberFormat("$#,##0.00").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#385723");
  dash.getRange("D7").setNumberFormat("0.0%").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");
  
  // Extended Formulas Integration (MoM & Uncategorized alerts)
  dash.getRange("F3").setValue("Uncategorized Alerts:").setFontWeight("bold");
  dash.getRange("F4").setFormula('=COUNTIFS(Transactions_DB!C:C, "<>", Transactions_DB!H:H, "")').setFontColor("#c00000").setFontWeight("bold").setHorizontalAlignment("center");
  
  dash.getRange("G3").setValue("Net Goal progress:").setFontWeight("bold");
  dash.getRange("G4").setFormula('=IFERROR(C7/1500, 0)').setNumberFormat("0.0%").setFontColor("#385723").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Budget & Progress Table (Enforcing dynamic category indexing)
  dash.getRange("A9:E9").setValues([["Category", "Planned Budget", "Actual Spending", "Difference", "Budget Progress"]]);
  formatSubHeader(dash, "A9:E9");
  
  // Programmatically resolve dynamic length of Enum categories
  const enumCategoriesCount = categories.length; // N=9 (adaptable if Enum updates)
  for (let r = 0; r < enumCategoriesCount; r++) {
    const rowIdx = 10 + r;
    // Link category name dynamically to Enum tab to maintain single source of truth
    dash.getRange(rowIdx, 1).setFormula(`=Enum!A${3 + r}`).setFontColor("#385723").setFontWeight("bold");
    
    // Set default planned budgets (These will be updated programmatically from Budget_History_DB)
    dash.getRange(rowIdx, 2).setValue(0).setFontColor("#002060");
    
    // Calculate actual outflows dynamically from Transactions_DB
    dash.getRange(rowIdx, 3).setFormula(
      `=IF($D$4="All", SUMIFS(Transactions_DB!F:F, Transactions_DB!H:H, A${rowIdx}, Transactions_DB!G:G, "OUTFLOW", Transactions_DB!C:C, ">="&DATE($C$4, 1, 1), Transactions_DB!C:C, "<="&DATE($C$4, 12, 31)), SUMIFS(Transactions_DB!F:F, Transactions_DB!H:H, A${rowIdx}, Transactions_DB!G:G, "OUTFLOW", Transactions_DB!C:C, ">="&DATEVALUE($D$4&" 1, "&$C$4), Transactions_DB!C:C, "<="&EOMONTH(DATEVALUE($D$4&" 1, "&$C$4), 0)))`
    );
    
    // Compute Difference remaining
    dash.getRange(rowIdx, 4).setFormula(`=B${rowIdx}-C${rowIdx}`);
    
    // Set Warning Sparkline Progress Bars
    dash.getRange(rowIdx, 5).setFormula(
      `=IF(B${rowIdx}>0, SPARKLINE(C${rowIdx}, {"charttype","bar";"max", B${rowIdx};"color1", IF(C${rowIdx}>B${rowIdx}, "#db4437", "#4285f4")}), "")`
    );
  }
  
  // Format numeric columns on dashboard category table
  dash.getRange(10, 2, enumCategoriesCount, 3).setNumberFormat("$#,##0.00");
  
  // Total summary row (Row index is dynamically 10 + N)
  const totalRowIdx = 10 + enumCategoriesCount; // 19
  dash.getRange(totalRowIdx, 1, 1, 4).setValues([[
    "Total Tracked Outflows", `=SUM(B10:B${totalRowIdx-1})`, `=SUM(C10:C${totalRowIdx-1})`, `=SUM(D10:D${totalRowIdx-1})`
  ]]).setFontWeight("bold");
  dash.getRange(totalRowIdx, 2, 1, 3).setNumberFormat("$#,##0.00");
  dash.getRange(totalRowIdx, 1, 1, 5).setBackground("#f2f2f2");
  
  // LEDGER HEADER & FEED (Rows 12+N onwards)
  const ledgerHeaderIdx = totalRowIdx + 2; // Row 21
  dash.getRange(ledgerHeaderIdx, 1, 1, 7).setValues([["Date", "Merchant Name", "Amount", "Flow Direction", "Category", "Account Source", "Channel"]]);
  formatSubHeader(dash, `A${ledgerHeaderIdx}:G${ledgerHeaderIdx}`);
  
  const ledgerFormulaIdx = ledgerHeaderIdx + 1; // Row 22
  dash.getRange(ledgerFormulaIdx, 1).setFormula(
    `=IF($D$4="All", FILTER({Transactions_DB!C2:C, Transactions_DB!D2:D, Transactions_DB!F2:F, Transactions_DB!G2:G, Transactions_DB!H2:H, Transactions_DB!I2:I, Transactions_DB!J2:J}, YEAR(Transactions_DB!C2:C) = $C$4), FILTER({Transactions_DB!C2:C, Transactions_DB!D2:D, Transactions_DB!F2:F, Transactions_DB!G2:G, Transactions_DB!H2:H, Transactions_DB!I2:I, Transactions_DB!J2:J}, YEAR(Transactions_DB!C2:C) = $C$4, MONTH(Transactions_DB!C2:C) = MONTH($D$4&1)))`
  );
  
  dash.getRange(ledgerFormulaIdx, 1, 100, 1).setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
  dash.getRange(ledgerFormulaIdx, 3, 100, 1).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  dash.getRange(ledgerFormulaIdx, 4, 100, 1).setHorizontalAlignment("center");
  
  dash.autoResizeColumns(1, 7);
  dash.setFrozenRows(5); // Freeze controller pane
  
  // Initial Sync from budget history table for default August 2026
  syncDashboardBudgets(dash);
  
  // =========================================================================
  // 5.5 PROGRAMMATIC DYNAMIC CHARTS (ICAEW Principle 19: Data Visualization)
  // =========================================================================
  try {
    // Clear any existing charts on the dashboard to prevent duplicates on regeneration
    const existingCharts = dash.getCharts();
    existingCharts.forEach(c => dash.removeChart(c));
    
    // Chart 1: Budget vs. Actual (Side-by-Side Column Chart)
    const budgetVsActualChart = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(9, 1, enumCategoriesCount + 1, 3)) // Category, Planned, Actual (includes headers in row 9)
      .setPosition(9, 9, 0, 0) // Positioned at Row 9, Column I (Col 9)
      .setOption('title', 'Monthly Planned Budget vs. Actual Spending')
      .setOption('legend', { position: 'top' })
      .setOption('width', 500)
      .setOption('height', 300)
      .setOption('colors', ['#4285f4', '#db4437']) // Blue for Planned, Red for Actual
      .setOption('hAxis', {
        title: 'Categories',
        textStyle: { fontSize: 10 }
      })
      .setOption('vAxis', {
        title: 'Amount ($)'
      })
      .build();
    dash.insertChart(budgetVsActualChart);
    
    // Chart 2: Category Expense Allocation (Donut Chart)
    const expenseBreakdownChart = dash.newChart()
      .setChartType(Charts.ChartType.DONUT)
      .addRange(dash.getRange(9, 1, enumCategoriesCount + 1, 1)) // Categories including header
      .addRange(dash.getRange(9, 3, enumCategoriesCount + 1, 1)) // Actual Spending including header
      .setPosition(9, 16, 0, 0) // Positioned at Row 9, Column P (Col 16)
      .setOption('title', 'Expense Allocation Breakdown')
      .setOption('legend', { position: 'right' })
      .setOption('width', 450)
      .setOption('height', 300)
      .setOption('pieHole', 0.4) // Makes it a donut chart
      .build();
    dash.insertChart(expenseBreakdownChart);
    
    Logger.log("Dynamic charts successfully generated on the Dashboard!");
  } catch (err) {
    Logger.log("Notice: Dynamic charts could not be created programmatically: " + err.toString());
  }
  
  // -------------------------------------------------------------------------
  // CONTROL TAB (Integrity Monitors)
  // -------------------------------------------------------------------------
  const control = sheets["Control"];
  formatHeader(control, "A1:D1", "WORKBOOK CONTROL CHECKS & BALANCES");
  control.getRange("A3:D3").setValues([["Check Name", "Formula / Logic Description", "Current Value", "Status"]]);
  formatSubHeader(control, "A3:D3");
  
  // Check 1: Sign Normalization
  control.getRange("A4").setValue("Sign Normalization Check");
  control.getRange("B4").setValue("Counts transactions with negative amounts in Column F (Should be 0)");
  control.getRange("C4").setFormula('=COUNTIF(Transactions_DB!F:F, "<0")');
  control.getRange("D4").setFormula('=IF(C4=0, "OK", "ERROR: Negative values found")');
  
  // Check 2: Blank Category
  control.getRange("A5").setValue("Blank Category Check");
  control.getRange("B5").setValue("Counts transactions that have dates but no category assigned (Should be 0)");
  control.getRange("C5").setFormula('=COUNTIFS(Transactions_DB!C2:C, "<>", Transactions_DB!H2:H, "")');
  control.getRange("D5").setFormula('=IF(C5=0, "OK", "ERROR: Uncategorized transactions")');
  
  // Check 3: Invalid Flow
  control.getRange("A6").setValue("Capital Flow Classification Check");
  control.getRange("B6").setValue("Counts rows missing standardized Capital_Flow definitions in Column G");
  control.getRange("C6").setFormula('=COUNTIFS(Transactions_DB!C2:C, "<>", Transactions_DB!G2:G, "<>INFLOW", Transactions_DB!G2:G, "<>OUTFLOW")');
  control.getRange("D6").setFormula('=IF(C6=0, "OK", "ERROR: Invalid flow class")');
  
  // Master alert formula
  control.getRange("A2").setFormula('=IF(COUNTIF(D4:D6, "ERROR*")=0, "✅ SYSTEM OK", "❌ ERROR DETECTED")');
  control.getRange("A2").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  control.autoResizeColumns(1, 4);
  
  // Apply visual conditional colors safely
  applyControlTabFormatting(ss);
  
  Logger.log("Enterprise personal finance workbook successfully compiled under FAST/ICAEW design parameters!");
}

/**
 *  Applies Light Red/Green conditional colors programmatically to Control!D4:D6.
 */
function applyControlTabFormatting(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  
  const controlSheet = ss.getSheetByName("Control");
  if (!controlSheet) return;
  
  const range = controlSheet.getRange("D4:D6");
  
  const ruleOK = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OK")
    .setBackground("#D9EAD3")
    .setRanges([range])
    .build();
    
  const ruleError = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith("ERROR")
    .setBackground("#F4CCCC")
    .setRanges([range])
    .build();
    
  const rules = controlSheet.getConditionalFormatRules();
  rules.push(ruleOK, ruleError);
  controlSheet.setConditionalFormatRules(rules);
  Logger.log("Successfully applied conditional formatting to Control!D4:D6.");
}

// ===========================================================================
// 3. BACKGROUND SYNCHRONIZER (Budget_History_DB -> Dashboard)
// ===========================================================================

/**
 *  Queries Budget_History_DB and synchronizes the values into Column B of Dashboard.
 */
function syncDashboardBudgets(dash) {
  const ss = dash.getParent();
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  const enumSheet = ss.getSheetByName("Enum");
  if (!budgetDB || !enumSheet) return;
  
  // Retrieve selected Period
  const year = parseInt(dash.getRange("C4").getValue(), 10);
  const month = dash.getRange("D4").getValue().toString().trim();
  
  // Retrieve active categories from Dashboard column A (Rows 10 down to A18)
  const lastEnumRow = enumSheet.getLastRow();
  if (lastEnumRow < 3) return;
  const categories = enumSheet.getRange(3, 1, lastEnumRow - 2, 1).getValues().flat().filter(String);
  const numCategories = categories.length;
  
  // Load entire budget history database into memory for processing speed
  const dbLastRow = budgetDB.getLastRow();
  let dbData = [];
  if (dbLastRow >= 3) {
    dbData = budgetDB.getRange(3, 1, dbLastRow - 2, 4).getValues();
  }
  
  // Map historical values to active categories for the selected period
  const updatedBudgets = [];
  categories.forEach(cat => {
    // Search history for matching Year, Month, Category
    const match = dbData.find(row => 
      parseInt(row[0], 10) === year && 
      row[1].toString().trim() === month && 
      row[2].toString().trim() === cat
    );
    updatedBudgets.push([match ? match[3] : 0.00]); // Default to 0.00 if no record exists
  });
  
  // Write the retrieved historical planned values into Column B
  dash.getRange(10, 2, numCategories, 1).setValues(updatedBudgets);
  Logger.log("Synced " + numCategories + " category budgets for Period: " + year + " " + month);
}

// ===========================================================================
// 4. SIDEBAR BUDGET CONTROLLER ENGINE
// ===========================================================================

/**
 *  Displays the custom sidebar to edit historical monthly budgets.
 */
function showBudgetSidebar() {
  const html = HtmlService.createTemplateFromFile("SidebarView");
  const htmlOutput = html.evaluate()
    .setTitle("Historical Budget Allocator")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 *  Sidebar Endpoint: Returns active categories and their currently loaded budgets for Year & Month.
 */
function getCategoryBudgets(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  
  const categories = enumSheet.getRange(3, 1, enumSheet.getLastRow() - 2, 1).getValues().flat().filter(String);
  const dbLastRow = budgetDB.getLastRow();
  let dbData = [];
  if (dbLastRow >= 3) {
    dbData = budgetDB.getRange(3, 1, dbLastRow - 2, 4).getValues();
  }
  
  // Map matching allocations
  return categories.map(cat => {
    const match = dbData.find(row => 
      parseInt(row[0], 10) === parseInt(year, 10) && 
      row[1].toString().trim() === month.toString().trim() && 
      row[2].toString().trim() === cat
    );
    return {
      category: cat,
      amount: match ? match[3] : 0.00
    };
  });
}

/**
 *  Sidebar Endpoint: IDEMPOTENT update/insert of budgets into Budget_History_DB.
 */
function saveCategoryBudgets(year, month, budgetData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const dbLastRow = budgetDB.getLastRow();
    let range;
    let dbData = [];
    if (dbLastRow >= 3) {
      range = budgetDB.getRange(3, 1, dbLastRow - 2, 4);
      dbData = range.getValues();
    }
    
    // Process each incoming allocation item
    budgetData.forEach(item => {
      const matchIndex = dbData.findIndex(row => 
        parseInt(row[0], 10) === parseInt(year, 10) && 
        row[1].toString().trim() === month.toString().trim() && 
        row[2].toString().trim() === item.category
      );
      
      if (matchIndex !== -1) {
        // IDEMPOTENT update: rewrite amount in the matching memory row index
        dbData[matchIndex][3] = parseFloat(item.amount);
      } else {
        // Insert new row if record didn't exist before
        dbData.push([parseInt(year, 10), month.toString().trim(), item.category, parseFloat(item.amount)]);
      }
    });
    
    // Reset and paste updated dataset
    if (dbLastRow >= 3) budgetDB.getRange(3, 1, dbLastRow - 2, 4).clearContent();
    budgetDB.getRange(3, 1, dbData.length, 4).setValues(dbData);
    
    // Instantly refresh the active Monthly Dashboard if Year and Month align!
    const dash = ss.getSheetByName("The Monthly Dashboard Tab");
    if (dash) {
      const dashYear = parseInt(dash.getRange("C4").getValue(), 10);
      const dashMonth = dash.getRange("D4").getValue().toString().trim();
      if (dashYear === parseInt(year, 10) && dashMonth === month.toString().trim()) {
        syncDashboardBudgets(dash);
      }
    }
    
    return "SUCCESS";
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// 5. EXISTING AUTOMATION CHANNELS (Consolidated & Untouched)
// ===========================================================================

/**
 *  Handles incoming POST requests from manual external HTML forms.
 */
function doPost(e) {
  const sheetName = 'Transactions_DB';
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Transactions_DB sheet not found.");
    
    const rowData = JSON.parse(e.postData.contents);
    const dateValue = new Date(rowData.Date);
    const timestamp = new Date();
    const transactionId = "MANUAL-" + timestamp.getTime();
    
    const cleanMerchant = rowData.Merchant ? rowData.Merchant.trim() : "Form Entry";
    const resolvedCategory = rowData.Category && rowData.Category !== "" ? rowData.Category : autoCategorizeMerchant(cleanMerchant);
    
    sheet.appendRow([
      transactionId,
      timestamp,
      dateValue,
      cleanMerchant,
      cleanMerchant, // Clean defaults to raw for forms
      Math.abs(parseFloat(rowData.Amount)),
      rowData.Flow,
      resolvedCategory,
      rowData.Source,
      "WEB_FORM",
      "SETTLED"
    ]);
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'id': transactionId })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'error': error.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 *  Automatically parses unread Chase transaction alert emails from Gmail.
 */
function pullGmailTransactionAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return;
  
  const lastRow = dbSheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }
  
  const searchQuery = 'from:Chase "transaction alert" is:unread';
  const threads = GmailApp.search(searchQuery, 0, 10);
  let newRowsCount = 0;
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const msgId = msg.getId();
      
      if (existingIDs.indexOf(msgId) !== -1) continue; // Deduplication
      
      const body = msg.getPlainBody();
      // Scrape Merchant, Date, Amount (Simplified placeholder example, extend regex to taste)
      const amountMatch = body.match(/\$(\d+\.\d{2})/);
      if (!amountMatch) continue;
      
      const amount = parseFloat(amountMatch[1]);
      const date = msg.getDate();
      
      // Attempt to extract merchant details from Chase email body
      let merchantRaw = "CHASE TRANSACTION ALERT";
      let merchantClean = "Chase Purchase";
      const merchantMatch = body.match(/at\s+([^,\n\r\.]+)/i) || body.match(/with\s+([^,\n\r\.]+)/i);
      if (merchantMatch) {
        merchantRaw = merchantMatch[0].trim();
        merchantClean = merchantMatch[1].trim();
      }
      
      const suggestedCategory = autoCategorizeMerchant(merchantClean);
      
      dbSheet.appendRow([
        msgId,
        new Date(),
        date,
        merchantRaw,
        merchantClean,
        amount,
        "OUTFLOW",
        suggestedCategory,
        "Chase Sapphire",
        "GMAIL_ALERT",
        "SETTLED"
      ]);
      msg.markRead();
      newRowsCount++;
    }
  }
  Logger.log("Completed Gmail parsing. Imported " + newRowsCount + " new transactions.");
}

/**
 *  Prompt user for a Drive File ID to run CSV import.
 */
function promptForChaseCSV() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Import Chase CSV",
    "Please enter the Google Drive File ID of your uploaded Chase CSV:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button("OK")) {
    const fileId = response.getResponseText().trim();
    if (fileId === "") {
      ui.alert("Error: File ID cannot be blank.");
      return;
    }
    try {
      const result = importChaseCSV(fileId);
      ui.alert("Import Status ✅", result, ui.ButtonSet.OK);
    } catch (err) {
      ui.alert("Import Failed ❌", err.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 *  Imports and normalizes Chase credit/debit statements.
 */
function importChaseCSV(fileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) throw new Error("Transactions_DB tab not found.");
  
  let file;
  if (fileId) {
    file = DriveApp.getFileById(fileId);
  } else {
    const files = DriveApp.getFilesByName("Chase.csv");
    if (files.hasNext()) {
      file = files.next();
    }
  }
  
  if (!file) return "No CSV statement detected in Drive.";
  
  const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  if (csvData.length < 2) return "Empty CSV file.";
  
  const lastRow = dbSheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }
  
  const csvHeaders = csvData[0].map(h => h.trim());
  const colDate = csvHeaders.indexOf("Transaction Date") !== -1 ? csvHeaders.indexOf("Transaction Date") : csvHeaders.indexOf("Posting Date");
  const colMerchant = csvHeaders.indexOf("Description");
  const colAmount = csvHeaders.indexOf("Amount");
  
  if (colDate === -1 || colMerchant === -1 || colAmount === -1) {
    throw new Error("Could not map Chase CSV headers. Ensure Date, Description, and Amount columns exist.");
  }
  
  let newRowsCount = 0;
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row[colDate]) continue;
    
    const rawAmountStr = row[colAmount].toString().replace(/'/g, "").trim();
    const rawAmount = parseFloat(rawAmountStr);
    if (isNaN(rawAmount)) continue;
    
    // Deterministic deduplication signature
    const sigId = "CHASE-CSV-" + row[colDate] + "-" + row[colMerchant].substring(0, 10).replace(/\s/g, "") + "-" + Math.abs(rawAmount).toFixed(2);
    if (existingIDs.indexOf(sigId) !== -1) continue;
    
    const amountMagnitude = Math.abs(rawAmount);
    const flow = rawAmount < 0 ? "OUTFLOW" : "INFLOW";
    
    const cleanMerchant = row[colMerchant].trim();
    const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
    
    dbSheet.appendRow([
      sigId,
      new Date(),
      new Date(row[colDate]),
      row[colMerchant],
      cleanMerchant, // Cleaned default
      amountMagnitude,
      flow,
      suggestedCategory, // Dynamic Auto-Categorized Category
      "Chase Checking",
      "CHASE_CSV",
      "SETTLED"
    ]);
    newRowsCount++;
  }
  return "Successfully imported " + newRowsCount + " transaction rows.";
}

// ===========================================================================
// 6. INTERNAL MODAL WINDOW BUILDERS (Clickable URL & UI helper)
// ===========================================================================

function showUrlModal(url) {
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
      <h3 style="color: #1f4e78;">Your Automated Sheet is Compiled!</h3>
      <p>The entire ledger schema and dynamic dashboards have been written cleanly.</p>
      <div style="margin: 30px 0;">
        <a href="${url}" target="_blank" style="background-color: #1f4e78; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px;">Open New Spreadsheet ➔</a>
      </div>
      <p style="color: #666; font-size: 11px;">If the link does not open, copy the URL below:<br>${url}</p>
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Compilation Success!");
}


// ===========================================================================
// 4.5 AI AUTO-CATEGORIZER & EMAIL BREACH ALERTS ENGINES
// ===========================================================================

/**
 * Dynamic rules-based auto-categorizer.
 * Scrapes raw merchant names and assigns appropriate categories.
 */
function autoCategorizeMerchant(merchantRaw) {
  if (!merchantRaw) return "Miscellaneous";
  const merchant = merchantRaw.toLowerCase();
  
  // Rules-based keyword mapping
  if (merchant.indexOf("starbucks") !== -1 || merchant.indexOf("mcdonald") !== -1 || 
      merchant.indexOf("dunkin") !== -1 || merchant.indexOf("restaurant") !== -1 || 
      merchant.indexOf("cafe") !== -1 || merchant.indexOf("grubhub") !== -1 || 
      merchant.indexOf("ubereats") !== -1 || merchant.indexOf("doorndash") !== -1 ||
      merchant.indexOf("pizza") !== -1 || merchant.indexOf("deli") !== -1) {
    return "Food & Drink";
  }
  if (merchant.indexOf("kroger") !== -1 || merchant.indexOf("walmart") !== -1 || 
      merchant.indexOf("wholefd") !== -1 || merchant.indexOf("safeway") !== -1 || 
      merchant.indexOf("supermarket") !== -1 || merchant.indexOf("trader joe") !== -1 ||
      merchant.indexOf("aldi") !== -1 || merchant.indexOf("costco") !== -1) {
    return "Groceries";
  }
  if (merchant.indexOf("netflix") !== -1 || merchant.indexOf("spotify") !== -1 || 
      merchant.indexOf("hulu") !== -1 || merchant.indexOf("youtube") !== -1 || 
      merchant.indexOf("disney") !== -1 || merchant.indexOf("steam") !== -1 ||
      merchant.indexOf("cinema") !== -1 || merchant.indexOf("ticketmaster") !== -1) {
    return "Entertainment";
  }
  if (merchant.indexOf("uber") !== -1 || merchant.indexOf("lyft") !== -1 || 
      merchant.indexOf("chevron") !== -1 || merchant.indexOf("shell") !== -1 || 
      merchant.indexOf("bp") !== -1 || merchant.indexOf("subway") !== -1 || 
      merchant.indexOf("transit") !== -1 || merchant.indexOf("gas station") !== -1 ||
      merchant.indexOf("parking") !== -1 || merchant.indexOf("toll") !== -1) {
    return "Transportation";
  }
  if (merchant.indexOf("comcast") !== -1 || merchant.indexOf("coned") !== -1 || 
      merchant.indexOf("electric") !== -1 || merchant.indexOf("att") !== -1 || 
      merchant.indexOf("verizon") !== -1 || merchant.indexOf("water") !== -1 ||
      merchant.indexOf("utility") !== -1 || merchant.indexOf("insurance") !== -1 ||
      merchant.indexOf("power") !== -1 || merchant.indexOf("gas company") !== -1) {
    return "Bills & Utilities";
  }
  if (merchant.indexOf("amazon") !== -1 || merchant.indexOf("target") !== -1 || 
      merchant.indexOf("home depot") !== -1 || merchant.indexOf("ikea") !== -1 ||
      merchant.indexOf("lowes") !== -1 || merchant.indexOf("bed bath") !== -1 ||
      merchant.indexOf("furniture") !== -1 || merchant.indexOf("hardware") !== -1) {
    return "Home";
  }
  if (merchant.indexOf("delta") !== -1 || merchant.indexOf("airbnb") !== -1 || 
      merchant.indexOf("hotel") !== -1 || merchant.indexOf("expedia") !== -1 || 
      merchant.indexOf("flight") !== -1 || merchant.indexOf("united airlines") !== -1 ||
      merchant.indexOf("travel") !== -1 || merchant.indexOf("booking") !== -1) {
    return "Travel";
  }
  if (merchant.indexOf("nike") !== -1 || merchant.indexOf("macys") !== -1 || 
      merchant.indexOf("ulta") !== -1 || merchant.indexOf("barber") !== -1 ||
      merchant.indexOf("sephora") !== -1 || merchant.indexOf("salon") !== -1 ||
      merchant.indexOf("clothing") !== -1 || merchant.indexOf("pharmacy") !== -1 ||
      merchant.indexOf("cvs") !== -1 || merchant.indexOf("walgreens") !== -1) {
    return "Personal";
  }
  
  return "Miscellaneous"; // Default fallback
}

/**
 * Loops through the entire Transactions_DB and categorizes rows missing a category.
 * Also cleans "Miscellaneous" defaults if keywords are found.
 */
function autoCategorizeUnassigned() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) {
    SpreadsheetApp.getUi().alert("Error ❌", "Transactions_DB sheet not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return;
  
  const range = dbSheet.getRange(2, 1, lastRow - 1, 11);
  const data = range.getValues();
  let updatedCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const rowIdx = 2 + i;
    const category = data[i][7]; // Column H
    const rawMerchant = data[i][3]; // Column D
    const cleanMerchant = data[i][4]; // Column E
    
    // Categorize if cell is blank, missing, or defaulted to Miscellaneous
    if (!category || category === "" || category === "Miscellaneous") {
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant || rawMerchant);
      if (suggestedCategory !== "Miscellaneous") {
        dbSheet.getRange(rowIdx, 8).setValue(suggestedCategory); // Set Column H (Category)
        updatedCount++;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert("Auto-Categorization Complete 🤖", "Successfully updated " + updatedCount + " transactions based on rules-based merchant patterns.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Evaluates monthly spending on the active dashboard.
 * Sends email notifications if actual spending breaches planned targets.
 * Employs standard State tracking to prevent email spam.
 */
function checkBudgetBreaches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  const enumSheet = ss.getSheetByName("Enum");
  if (!dash || !enumSheet) {
    Logger.log("Dashboard or Enum sheet not found.");
    return;
  }
  
  const year = dash.getRange("C4").getValue().toString().trim();
  const month = dash.getRange("D4").getValue().toString().trim();
  
  if (month === "All") {
    Logger.log("Skipping budget breach checks for cumulative 'All' selection.");
    return;
  }
  
  const lastEnumRow = enumSheet.getLastRow();
  if (lastEnumRow < 3) return;
  const enumCategoriesCount = lastEnumRow - 2;
  
  // Read category budget data block from Dashboard A10:D(10+N-1)
  const dataRange = dash.getRange(10, 1, enumCategoriesCount, 4);
  const data = dataRange.getValues();
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const userEmail = Session.getActiveUser().getEmail();
  let breaches = [];
  
  for (let i = 0; i < data.length; i++) {
    const category = data[i][0];
    const planned = parseFloat(data[i][1]);
    const actual = parseFloat(data[i][2]);
    const diff = parseFloat(data[i][3]); // Planned - Actual
    
    if (planned > 0 && actual > planned) {
      const breachKey = `breach_${year}_${month}_${category}`;
      const alreadySent = scriptProperties.getProperty(breachKey);
      
      if (!alreadySent) {
        breaches.push({
          category: category,
          planned: planned,
          actual: actual,
          overrun: Math.abs(diff)
        });
        // Flag as sent to avoid duplicate emails
        scriptProperties.setProperty(breachKey, "true");
      }
    }
  }
  
  if (breaches.length > 0) {
    let emailBody = `<h3>⚠️ BUDGET BREACH ALERT - ${month} ${year}</h3>`;
    emailBody += `<p>The following spending limits have been breached on your Interactive Financial Dashboard:</p>`;
    emailBody += `<table border="1" cellpadding="5" style="border-collapse: collapse; font-family: Arial, sans-serif;">`;
    emailBody += `<tr style="background-color: #1f4e78; color: white;"><th>Category</th><th>Planned Budget</th><th>Actual Spent</th><th>Overrun Amount</th></tr>`;
    
    breaches.forEach(b => {
      emailBody += `<tr>`;
      emailBody += `<td><b>${b.category}</b></td>`;
      emailBody += `<td style="color: blue;">$${b.planned.toFixed(2)}</td>`;
      emailBody += `<td style="color: red;">$${b.actual.toFixed(2)}</td>`;
      emailBody += `<td style="color: #c00000; font-weight: bold;">$${b.overrun.toFixed(2)}</td>`;
      emailBody += `</tr>`;
    });
    
    emailBody += `</table>`;
    emailBody += `<p>Please review your transactions on <b>'The Monthly Dashboard Tab'</b> and adjust your spending or re-allocate your historical budgets using the Budget Allocation Sidebar.</p>`;
    emailBody += `<br><p><i>Sent automatically by your Google Sheets Financial Automation System 🚀</i></p>`;
    
    MailApp.sendEmail({
      to: userEmail,
      subject: `⚠️ Budget Breach Notification: ${breaches.length} Category Alert(s)`,
      htmlBody: emailBody
    });
    
    Logger.log(`Dispatched budget breach emails for ${breaches.length} categories.`);
    try {
      SpreadsheetApp.getUi().alert("Email Alerts Dispatched! ✉️", "Dispatched " + breaches.length + " breach warning(s) to " + userEmail, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  } else {
    Logger.log("No new budget breaches detected.");
    try {
      SpreadsheetApp.getUi().alert("Budget Check Complete ✅", "All actual spending remains within planned monthly boundaries. No alerts sent.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }
}

