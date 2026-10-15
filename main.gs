/**
 *  ===========================================================================
 *  GOOGLE SHEETS FINANCIAL AUTOMATION SYSTEM (UNIFIED MASTER ENGINE - V7)
 *  Standardized under the FAST Modeling Standard and ICAEW Spreadsheet Principles.
 *  
 *  This file unifies:
 *    1. Consolidated trigger system (onOpen, onEdit).
 *    2. In-place workbook tab initialization (About, Enum, Transactions_DB,
 *       Budget_History_DB, Dashboard, Control, Category_Rules_DB, Savings_Goals).
 *    3. Programmatic Dynamic Visualizations (Side-by-Side and Pie/Donut Charts).
 *    4. Gmail alerts scraper (Chase Alerts) AND Google Voice SMS transaction alerts scraper.
 *    5. Multi-account automatic detection for folder-based CSV statement ingesters.
 *    6. Database-driven No-Code Rules auto-categorizer (Category_Rules_DB).
 *    7. 'TRANSFER' flow type configuration to ignore internal transfer distortion.
 *    8. Interactive double-entry Monthly Savings Sweep (Savings_Goals -> Transactions_DB).
 *    9. 10-Year growth projection matrix and live Area Chart.
 *   10. Dynamic, paginated In-Sheet Transaction Editor Sidebar.
 *  ===========================================================================
 */

// ===========================================================================
// 1. CONSOLIDATED TRIGGER SYSTEM (onOpen & onEdit)
// ===========================================================================

/**
 * Consolidated onOpen trigger. Runs automatically when the spreadsheet is opened.
 * Builds a unified, non-conflicting master menu for operations.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Financial Automation 🚀")
    .addItem("Initialize Sheets in This Workbook", "promptInPlaceInitialization")
    .addSeparator()
    .addItem("Batch Import CSV Folder", "importAllPendingCSVs")
    .addItem("Run AI Auto-Categorization", "autoCategorizeUnassigned")
    .addItem("Execute Monthly Savings Sweep", "promptSavingsSweep")
    .addSeparator()
    .addItem("Open Budget History Sidebar", "showBudgetSidebar")
    .addItem("Open Transaction Editor Sidebar", "showTransactionEditorSidebar")
    .addSeparator()
    .addItem("Check Budget Breaches", "checkBudgetBreaches")
    .addItem("Dispatch Weekly Summary Email", "sendWeeklyBudgetSummary")
    .addToUi();
}

/**
 * Consolidated onEdit background trigger.
 * Monitors dashboard dropdown edits to trigger zero-latency budget sync.
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
// 2. MAIN SHEET COMPILER (DYNAMIC RANGES & VISUALIZATIONS)
// ===========================================================================

/**
 * Prompt user before performing in-place sheet initialization in the active spreadsheet.
 */
function promptInPlaceInitialization() {
  const ui = SpreadsheetApp.getUi();
  const title = "Initialize Spreadsheet? ⚠️";
  const msg = "This will initialize all required operational tabs ('About', 'Enum', " +
               "'Transactions_DB', 'Budget_History_DB', 'The Monthly Dashboard Tab', 'Control', " +
               "'Category_Rules_DB', 'Savings_Goals') directly inside THIS spreadsheet workbook.\n\n" +
               "Any existing sheets with these names will be cleared. Do you want to proceed?";
  const response = ui.alert(title, msg, ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log(`[Init Log] Initializing database and dashboard sheets in-place.`);
    createWholeBook(ss);
    ui.alert("Success! 🎉", "Your financial sheets, dashboards, databases, and projection models have been structured successfully.", ui.ButtonSet.OK);
  }
}

/**
 * Programmatic sheet generator. Creates a new sheet.
 */
function createNewSpreadsheet() {
  const ss = SpreadsheetApp.create("My Automated Spreadsheet");
  console.log(`[Spreadsheet Creation] Generating new spreadsheet: "My Automated Spreadsheet"`);
  
  createWholeBook(ss);
  
  const url = ss.getUrl();
  Logger.log("Your new spreadsheet is ready at: " + url);
  
  try {
    const ui = SpreadsheetApp.getUi();
    const alertTitle = "Workbook Successfully Created! 🚀";
    const alertMessage = "Your automated spreadsheet is ready.\n\n" +
                         "Copy the URL below to open your new file:\n\n" + url + "\n\n" +
                         "💡 Tip: Run 'Initialize Sheets in This Workbook' inside your current file to keep all bound scripts connected!";
    ui.alert(alertTitle, alertMessage, ui.ButtonSet.OK);
    showUrlModal(url);
  } catch (e) {
    Logger.log("Notice: Pop-up modal could not be displayed because the script is running without an active browser UI context.");
    Logger.log("You can still access your compiled spreadsheet using the URL logged above: " + url);
  }
}

/**
 * Standardized Workbook Setup Script under FAST and ICAEW Principles.
 * Dynamically sizes category listings and configures all inter-sheet relationships.
 */
function createWholeBook(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("No spreadsheet found. Ensure a spreadsheet is created or active.");
  }
  
  console.log(`[Whole Book Compiler] Constructing sheet structures for Spreadsheet: ${ss.getId()}`);
  
  // Define required sheets
  const sheetNames = ["About", "Enum", "Transactions_DB", "Budget_History_DB", "The Monthly Dashboard Tab", "Control", "Category_Rules_DB", "Savings_Goals"];
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
  sheets["Category_Rules_DB"].activate(); ss.moveActiveSheet(7);
  sheets["Savings_Goals"].activate(); ss.moveActiveSheet(8);
  
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
  about.getRange("A14").setValue("2. Log manual entries via the HTML web form, sidebar panel, Gmail automation, or paste them directly using positive magnitudes.");
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
    ["Transportation"], ["Entertainment"], ["Personal"], ["Travel"], ["Miscellaneous"], ["Savings"], ["Transfer"]
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
    "Transaction_ID", "Timestamp", "Transaction_Date", "Merchant_Raw", "Merchant_Clean", "Amount", "Capital_Flow", "Category", "Account_Source", "Ingestion_Channel", "Sync_Status"
  ];
  db.getRange(1, 1, 1, dbHeaders.length).setValues([dbHeaders]);
  formatSubHeader(db, "A1:K1");
  db.setRowHeight(1, 28);
  
  // Set dropdown validation dynamically pointing to Enum sheet
  const categoryEnumRange = enumSheet.getRange(3, 1, categories.length, 1);
  const categoryValidation = SpreadsheetApp.newDataValidation().requireValueInRange(categoryEnumRange).setAllowInvalid(false).build();
  db.getRange("H2:H5000").setDataValidation(categoryValidation);
  
  const accountEnumRange = enumSheet.getRange(3, 2, accounts.length, 1);
  const accountValidation = SpreadsheetApp.newDataValidation().requireValueInRange(accountEnumRange).setAllowInvalid(false).build();
  db.getRange("I2:I5000").setDataValidation(accountValidation);
  
  const flowValidation = SpreadsheetApp.newDataValidation().requireValueInList(["INFLOW", "OUTFLOW", "TRANSFER"]).setAllowInvalid(false).build();
  db.getRange("G2:G5000").setDataValidation(flowValidation);
  
  const channelValidation = SpreadsheetApp.newDataValidation().requireValueInList(["WEB_FORM", "CHASE_CSV", "GMAIL_ALERT", "GOOGLE_VOICE_SMS", "SIDEBAR_EDITOR", "MANUAL"]).setAllowInvalid(false).build();
  db.getRange("J2:J5000").setDataValidation(channelValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation().requireValueInList(["SETTLED", "PENDING", "MANUAL_REVIEW"]).setAllowInvalid(false).build();
  db.getRange("K2:K5000").setDataValidation(statusValidation);
  
  // Standard dummy row to initialize formulas
  db.getRange("A2:K2").setValues([["INIT_SETUP_DUMMY", new Date(), new Date(), "INITIAL SYSTEM SETUP", "Initial Setup", 0.01, "INFLOW", "Miscellaneous", "Cash", "MANUAL", "SETTLED"]]);
  db.getRange("F2:F5000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  db.getRange("C2:C5000").setNumberFormat("yyyy-mm-dd");
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
  budgetDB.getRange("D3:D5000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  budgetDB.autoResizeColumns(1, 4);
  
  // -------------------------------------------------------------------------
  // CATEGORY_RULES_DB (Rules mapping raw merchant descriptions to clean name and category)
  // -------------------------------------------------------------------------
  const rulesDb = sheets["Category_Rules_DB"];
  formatHeader(rulesDb, "A1:C1", "NO-CODE TRANSACTION CLASSIFICATION RULES");
  rulesDb.getRange("A2:C2").setValues([["Merchant Keyword Rule (Raw)", "Standardized Clean Name", "Mapped Category"]]).setFontWeight("bold");
  formatSubHeader(rulesDb, "A2:C2");
  rulesDb.setRowHeight(2, 25);
  
  const startingRules = [
    ["starbucks", "Starbucks", "Food & Drink"],
    ["mcdonald", "McDonalds", "Food & Drink"],
    ["kroger", "Kroger", "Groceries"],
    ["wholefd", "Whole Foods", "Groceries"],
    ["netflix", "Netflix", "Entertainment"],
    ["spotify", "Spotify", "Entertainment"],
    ["comcast", "Comcast", "Bills & Utilities"],
    ["coned", "ConEd Utilities", "Bills & Utilities"],
    ["amazon", "Amazon Purchase", "Home"],
    ["target", "Target Purchase", "Personal"],
    ["uber", "Uber Transit", "Transportation"],
    ["lyft", "Lyft Transit", "Transportation"],
    ["chase transfer", "Chase Savings Sweep", "Transfer"]
  ];
  rulesDb.getRange(3, 1, startingRules.length, 3).setValues(startingRules);
  rulesDb.getRange("C3:C1000").setDataValidation(categoryValidation);
  rulesDb.getRange("A3:A1000").setFontColor("#002060"); // Blue user input color
  rulesDb.getRange("B3:B1000").setFontColor("#002060");
  rulesDb.autoResizeColumns(1, 3);
  
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
  
  // Programmatically resolve dynamic length of Enum categories (excluding Transfer)
  const enumCategoriesCount = categories.length - 1; // Exclude Transfer row
  for (let r = 0; r < enumCategoriesCount; r++) {
    const rowIdx = 10 + r;
    // Link category name dynamically to Enum tab to maintain single source of truth
    dash.getRange(rowIdx, 1).setFormula(`=Enum!A${3 + r}`).setFontColor("#385723").setFontWeight("bold");
    
    // Set default planned budgets (These will be updated programmatically from Budget_History_DB)
    dash.getRange(rowIdx, 2).setValue(0).setFontColor("#002060");
    
    // Calculate actual outflows dynamically from Transactions_DB (Summing only standard OUTFLOWS, ignoring transfers)
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
  const totalRowIdx = 10 + enumCategoriesCount; // 20
  dash.getRange(totalRowIdx, 1, 1, 4).setValues([["Total Tracked Outflows", `=SUM(B10:B${totalRowIdx-1})`, `=SUM(C10:C${totalRowIdx-1})`, `=SUM(D10:D${totalRowIdx-1})`]]).setFontWeight("bold");
  dash.getRange(totalRowIdx, 2, 1, 3).setNumberFormat("$#,##0.00");
  dash.getRange(totalRowIdx, 1, 1, 5).setBackground("#f2f2f2");
  
  // LEDGER HEADER & FEED (Rows 12+N onwards)
  const ledgerHeaderIdx = totalRowIdx + 2; // Row 22
  dash.getRange(ledgerHeaderIdx, 1, 1, 7).setValues([["Date", "Merchant Name", "Amount", "Flow Direction", "Category", "Account Source", "Channel"]]);
  formatSubHeader(dash, `A${ledgerHeaderIdx}:G${ledgerHeaderIdx}`);
  
  const ledgerFormulaIdx = ledgerHeaderIdx + 1; // Row 23
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
  
  // Dynamic charts generation
  try {
    const existingCharts = dash.getCharts();
    existingCharts.forEach(c => dash.removeChart(c));
    
    // Chart 1: Budget vs. Actual (Side-by-Side Column Chart)
    const budgetVsActualChart = dash.newChart()
      .setChartType(Sheets.ChartType.COLUMN)
      .addRange(dash.getRange(9, 1, enumCategoriesCount + 1, 3))
      .setPosition(9, 9, 0, 0)
      .setOption('title', 'Monthly Planned Budget vs. Actual Spending')
      .setOption('legend', { position: 'top' })
      .setOption('width', 500)
      .setOption('height', 300)
      .setOption('colors', ['#4285f4', '#db4437']) // Blue for Planned, Red for Actual
      .setOption('hAxis', { title: 'Categories', textStyle: { fontSize: 10 } })
      .setOption('vAxis', { title: 'Amount ($)' })
      .build();
    dash.insertChart(budgetVsActualChart);
    
    // Chart 2: Category Expense Allocation (Donut Chart using native PIE + pieHole)
    const expenseBreakdownChart = dash.newChart()
      .setChartType(Sheets.ChartType.PIE)
      .addRange(dash.getRange(9, 1, enumCategoriesCount + 1, 1))
      .addRange(dash.getRange(9, 3, enumCategoriesCount + 1, 1))
      .setPosition(9, 16, 0, 0)
      .setOption('title', 'Expense Allocation Breakdown')
      .setOption('legend', { position: 'right' })
      .setOption('width', 450)
      .setOption('height', 300)
      .setOption('pieHole', 0.4)
      .build();
    dash.insertChart(expenseBreakdownChart);
    console.log("[Whole Book Compiler] Programmable dashboard column and pie charts successfully generated.");
  } catch (err) {
    console.error("[Whole Book Compiler Error] Charts setup failed: " + err.toString());
  }
  
  // -------------------------------------------------------------------------
  // SAVINGS GOALS TAB
  // -------------------------------------------------------------------------
  const goalsSheet = sheets["Savings_Goals"];
  formatHeader(goalsSheet, "A1:F1", "SAVINGS GOALS PROGRESS PORTAL");
  goalsSheet.getRange("A2:F2").setValues([["Goal Name", "Target Amount", "Current Balance", "Target Date", "Sweep Percentage", "Progress Meter"]]).setFontWeight("bold");
  formatSubHeader(goalsSheet, "A2:F2");
  goalsSheet.setRowHeight(2, 25);
  
  const defaultGoals = [
    ["Emergency Fund", 10000.00, 1500.00, "2027-12-31", 0.50, ""],
    ["Travel Fund", 5000.00, 250.00, "2027-06-30", 0.30, ""],
    ["Investment Portfolio", 20000.00, 0.00, "2028-12-31", 0.20, ""]
  ];
  goalsSheet.getRange(3, 1, defaultGoals.length, 6).setValues(defaultGoals);
  
  // Apply formulas for Sparkline progress bars in column F
  for (let g = 0; g < defaultGoals.length; g++) {
    const rIdx = 3 + g;
    goalsSheet.getRange(rIdx, 6).setFormula(`=IF(B${rIdx}>0, SPARKLINE(C${rIdx}, {"charttype","bar";"max", B${rIdx};"color1", "#2e7d32"}), "")`);
  }
  
  goalsSheet.getRange("B3:C100").setNumberFormat("$#,##0.00");
  goalsSheet.getRange("D3:D100").setNumberFormat("yyyy-mm-dd");
  goalsSheet.getRange("E3:E100").setNumberFormat("0.0%");
  goalsSheet.getRange("A3:E100").setFontColor("#002060"); // User editable
  goalsSheet.autoResizeColumns(1, 6);
  
  // Setup 10-Year projection ledger below savings goals matrix
  goalsSheet.getRange("H1:L1").setValues([["LONG-TERM WEALTH GROWTH RETIREMENT FORECAST", "", "", "", ""]]).setFontWeight("bold").merge();
  goalsSheet.getRange("H1:L1").setBackground("#1f4e78").setFontColor("#ffffff").setHorizontalAlignment("center");
  goalsSheet.setRowHeight(1, 30);
  
  goalsSheet.getRange("H2").setValue("Expected Annual Investment Return (I2):").setFontWeight("bold");
  goalsSheet.getRange("I2").setValue(0.06).setNumberFormat("0.0%").setFontColor("#002060").setFontWeight("bold");
  goalsSheet.getRange("H3").setValue("Model Annual Savings Contribution (I3):").setFontWeight("bold");
  goalsSheet.getRange("I3").setValue(6000.00).setNumberFormat("$#,##0.00").setFontColor("#002060").setFontWeight("bold");
  
  goalsSheet.getRange("H5:L5").setValues([["Year", "Starting Net Worth", "New Contributions", "Compound Interest", "Projected Ending Balance"]]).setFontWeight("bold");
  formatSubHeader(goalsSheet, "H5:L5");
  
  for (let y = 0; g = 10, y < 10; y++) {
    const rowIdx = 6 + y;
    const yearNum = 2026 + y;
    goalsSheet.getRange(rowIdx, 8).setValue(yearNum);
    if (y === 0) {
      goalsSheet.getRange(rowIdx, 9).setFormula("=SUM(C3:C5)"); // Aggregates active savings balances
    } else {
      goalsSheet.getRange(rowIdx, 9).setFormula(`=L${rowIdx - 1}`);
    }
    goalsSheet.getRange(rowIdx, 10).setFormula("=$I$3"); // Contribution
    goalsSheet.getRange(rowIdx, 11).setFormula(`=(I${rowIdx}+J${rowIdx})*$I$2`); // Dynamic rate growth
    goalsSheet.getRange(rowIdx, 12).setFormula(`=SUM(I${rowIdx}:K${rowIdx})`); // Dynamic ending total
  }
  goalsSheet.getRange("I6:L15").setNumberFormat("$#,##0.00");
  goalsSheet.getRange("H6:H15").setNumberFormat("0000").setHorizontalAlignment("center");
  
  // Insert Multi-Year Savings Projection Graph (Area Chart)
  try {
    const projectionChart = goalsSheet.newChart()
      .setChartType(Sheets.ChartType.AREA)
      .addRange(goalsSheet.getRange("H5:H15")) // Year Range
      .addRange(goalsSheet.getRange("L5:L15")) // Projected Ending Balance
      .setPosition(1, 14, 0, 0) // Positioned nicely at Column N (Col 14), Row 1
      .setOption('title', 'Multi-Year Long-Term Savings Projection (10-Year Growth)')
      .setOption('legend', { position: 'none' })
      .setOption('width', 520)
      .setOption('height', 320)
      .setOption('colors', ['#2e7d32']) // Green for savings growth
      .setOption('hAxis', { title: 'Year', format: '0000' })
      .setOption('vAxis', { title: 'Projected Net Worth ($)' })
      .build();
    goalsSheet.insertChart(projectionChart);
    console.log("[Whole Book Compiler] Programmable Area projection chart successfully generated.");
  } catch (err) {
    console.error("[Whole Book Compiler Error] Projection chart setup failed: " + err.toString());
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
  control.getRange("C6").setFormula('=COUNTIFS(Transactions_DB!C2:C, "<>", Transactions_DB!G2:G, "<>INFLOW", Transactions_DB!G2:G, "<>OUTFLOW", Transactions_DB!G2:G, "<>TRANSFER")');
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
  
  // Retrieve active categories from Dashboard column A
  const lastEnumRow = enumSheet.getLastRow();
  if (lastEnumRow < 3) return;
  const categories = enumSheet.getRange(3, 1, lastEnumRow - 3, 1).getValues().flat().filter(String); // Exclude Transfer (last row)
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

function getDashboardPeriod() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  if (!dash) return { year: "2026", month: "August" };
  return {
    year: dash.getRange("C4").getValue().toString().trim(),
    month: dash.getRange("D4").getValue().toString().trim()
  };
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
  
  const categories = enumSheet.getRange(3, 1, enumSheet.getLastRow() - 3, 1).getValues().flat().filter(String); // Excludes Transfer row
  const dbLastRow = budgetDB.getLastRow();
  let dbData = [];
  if (dbLastRow >= 3) {
    dbData = budgetDB.getRange(3, 1, dbLastRow - 2, 4).getValues();
  }
  
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
    
    budgetData.forEach(item => {
      const matchIndex = dbData.findIndex(row => 
        parseInt(row[0], 10) === parseInt(year, 10) && 
        row[1].toString().trim() === month.toString().trim() && 
        row[2].toString().trim() === item.category
      );
      
      if (matchIndex !== -1) {
        dbData[matchIndex][3] = parseFloat(item.amount);
      } else {
        dbData.push([parseInt(year, 10), month.toString().trim(), item.category, parseFloat(item.amount)]);
      }
    });
    
    if (dbLastRow >= 3) budgetDB.getRange(3, 1, dbLastRow - 2, 4).clearContent();
    budgetDB.getRange(3, 1, dbData.length, 4).setValues(dbData);
    
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
    
    const rawMerchant = rowData.Merchant;
    const cleanMerchant = autoCleanMerchantName(rawMerchant);
    const resolvedCategory = rowData.Category && rowData.Category !== "" ? rowData.Category : autoCategorizeMerchant(cleanMerchant);
    
    sheet.appendRow([
      transactionId,
      timestamp,
      dateValue,
      rawMerchant,
      cleanMerchant,
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
 *  Automatically parses unread Chase transaction alert emails from Gmail
 *  AND Google Voice SMS messages forwarded to Gmail.
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
  
  // Ingest Channel 1: Traditional Gmail Chase Alerts
  const chaseQuery = 'from:Chase "transaction alert" is:unread';
  const chaseThreads = GmailApp.search(chaseQuery, 0, 10);
  let chaseCount = 0;
  
  for (let i = 0; i < chaseThreads.length; i++) {
    const messages = chaseThreads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const msgId = msg.getId();
      
      if (existingIDs.indexOf(msgId) !== -1) continue;
      
      const body = msg.getPlainBody();
      const amountMatch = body.match(/\$(\d+\.\d{2})/);
      if (!amountMatch) continue;
      
      const amount = parseFloat(amountMatch[1]);
      const date = msg.getDate();
      
      // Attempt to clean merchant from text snippet
      let merchantRaw = "CHASE TRANSACTION ALERT";
      const merchMatch = body.match(/at\s+([^\n\r,.]+?)\s+on/i) || body.match(/charge\s+of\s+.*?\s+at\s+([^\n\r,.]+)/i);
      if (merchMatch) {
        merchantRaw = merchMatch[1].trim();
      }
      
      const cleanMerchant = autoCleanMerchantName(merchantRaw);
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
      
      dbSheet.appendRow([
        msgId,
        new Date(),
        date,
        merchantRaw,
        cleanMerchant,
        amount,
        "OUTFLOW",
        suggestedCategory,
        "Chase Sapphire",
        "GMAIL_ALERT",
        "SETTLED"
      ]);
      msg.markRead();
      existingIDs.push(msgId);
      chaseCount++;
    }
  }
  
  // Ingest Channel 2: Google Voice SMS forward notifications (Venmo, CashApp, etc.)
  const gVoiceQuery = 'from:voice-noreply@google.com "SMS" is:unread';
  const voiceThreads = GmailApp.search(gVoiceQuery, 0, 15);
  let gvoiceCount = 0;
  
  for (let i = 0; i < voiceThreads.length; i++) {
    const messages = voiceThreads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const msgId = msg.getId();
      
      if (existingIDs.indexOf(msgId) !== -1) continue;
      
      const body = msg.getPlainBody();
      
      // Regex parsing for amount (e.g., $15.50 venmo or purchase)
      const amountMatch = body.match(/\$(\d+\.\d{2})/);
      if (!amountMatch) continue;
      
      const amount = parseFloat(amountMatch[1]);
      const date = msg.getDate();
      
      // Extract merchant or sender names
      let merchantRaw = "Google Voice Transaction";
      let flow = "OUTFLOW";
      let category = "Miscellaneous";
      
      if (body.toLowerCase().indexOf("venmo") !== -1) {
        merchantRaw = "Venmo Transaction";
        // Parse flow: "You paid John" vs "John paid you"
        if (body.toLowerCase().match(/paid\s+you/i) || body.toLowerCase().match(/received/i)) {
          flow = "INFLOW";
          category = "Income";
        }
      } else if (body.toLowerCase().indexOf("cash app") !== -1) {
        merchantRaw = "CashApp Transaction";
        if (body.toLowerCase().indexOf("sent you") !== -1) {
          flow = "INFLOW";
          category = "Income";
        }
      }
      
      // Regex to try parsing "paid [Name] $XX.XX for [Notes]"
      const venmoMatch = body.match(/paid\s+([^\$]+?)\s+\$/i);
      if (venmoMatch) {
        merchantRaw = "Venmo to " + venmoMatch[1].trim();
      }
      
      const cleanMerchant = autoCleanMerchantName(merchantRaw);
      const suggestedCategory = category !== "Income" ? autoCategorizeMerchant(cleanMerchant) : "Miscellaneous";
      
      dbSheet.appendRow([
        msgId,
        new Date(),
        date,
        merchantRaw,
        cleanMerchant,
        amount,
        flow,
        suggestedCategory,
        "Cash",
        "GOOGLE_VOICE_SMS",
        "SETTLED"
      ]);
      msg.markRead();
      existingIDs.push(msgId);
      gvoiceCount++;
    }
  }
  
  console.log(`[Gmail/Google Voice Parsing Engine] Ingest cycle complete. Imported ${chaseCount} Chase Alerts, ${gvoiceCount} Google Voice SMS Alerts.`);
}

/**
 *  Imports and normalizes Chase credit/debit statements.
 *  Uses folder name checks or filename structures to dynamically assign 'Account_Source'
 */
function importChaseCSV(fileId, filename) {
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
  
  // Rule 1: Dynamic Account Detection based on file context
  let detectedAccount = "Chase Checking"; // Default
  const checkString = (filename || file.getName() || "").toLowerCase();
  if (checkString.indexOf("sapphire") !== -1 || checkString.indexOf("credit") !== -1 || checkString.indexOf("cc") !== -1) {
    detectedAccount = "Chase Sapphire";
  } else if (checkString.indexOf("checking") !== -1 || checkString.indexOf("debit") !== -1) {
    detectedAccount = "Chase Checking";
  }
  
  let newRowsCount = 0;
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row[colDate]) continue;
    
    const rawAmountStr = row[colAmount].toString().replace(/'/g, "").trim();
    const rawAmount = parseFloat(rawAmountStr);
    if (isNaN(rawAmount)) continue;
    
    const merchantRaw = row[colMerchant];
    const cleanMerchant = autoCleanMerchantName(merchantRaw);
    
    // Deterministic deduplication signature
    const sigId = "CHASE-CSV-" + row[colDate] + "-" + merchantRaw.substring(0, 10).replace(/\s/g, "") + "-" + Math.abs(rawAmount).toFixed(2);
    if (existingIDs.indexOf(sigId) !== -1) continue;
    
    const amountMagnitude = Math.abs(rawAmount);
    
    // Rule 4: Handle "TRANSFER" flow types to prevent double counting
    let flow = rawAmount < 0 ? "OUTFLOW" : "INFLOW";
    let suggestedCategory = autoCategorizeMerchant(cleanMerchant);
    
    if (cleanMerchant.toLowerCase().indexOf("transfer") !== -1 || cleanMerchant.toLowerCase().indexOf("payment thank you") !== -1) {
      flow = "TRANSFER";
      suggestedCategory = "Transfer";
    }
    
    dbSheet.appendRow([
      sigId,
      new Date(),
      new Date(row[colDate]),
      merchantRaw,
      cleanMerchant,
      amountMagnitude,
      flow,
      suggestedCategory,
      detectedAccount,
      "CHASE_CSV",
      "SETTLED"
    ]);
    newRowsCount++;
  }
  return "Successfully imported " + newRowsCount + " transaction rows into: " + detectedAccount;
}

/**
 * Automatically scans a dedicated Google Drive folder for pending CSV statement files.
 * Batch imports, deduplicates, runs auto-categorizations, and archives processed files.
 */
function importAllPendingCSVs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) {
    ui.alert("Configuration Error ⚠️", "Transactions_DB sheet not found.", ui.ButtonSet.OK);
    return;
  }

  // 1. Locate or create the target landing folder in Google Drive
  let uploadFolder;
  const folders = DriveApp.getFoldersByName("Chase_CSV_Uploads");
  if (folders.hasNext()) {
    uploadFolder = folders.next();
  } else {
    uploadFolder = DriveApp.createFolder("Chase_CSV_Uploads");
    uploadFolder.createFolder("Archive");
    
    ui.alert(
      "Folder Setup Created! 📂",
      "I have automatically created a folder named 'Chase_CSV_Uploads' in your Google Drive root directory.\n\n" +
      "To batch import transactions:\n" +
      "1. Drag and drop all your bank CSV files directly into 'Chase_CSV_Uploads'.\n" +
      "2. Return here and click 'Batch Import CSV Folder' again to ingest all statements at once!",
      ui.ButtonSet.OK
    );
    return;
  }

  let archiveFolder;
  const archiveFolders = uploadFolder.getFoldersByName("Archive");
  if (archiveFolders.hasNext()) {
    archiveFolder = archiveFolders.next();
  } else {
    archiveFolder = uploadFolder.createFolder("Archive");
  }

  const files = uploadFolder.getFiles();
  let filesProcessed = 0;
  let totalNewRows = 0;

  const lastRow = dbSheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "text/csv" && !file.getName().toLowerCase().endsWith(".csv")) {
      continue;
    }

    try {
      const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      if (csvData.length < 2) continue;

      const csvHeaders = csvData[0].map(h => h.trim());
      const colDate = csvHeaders.indexOf("Transaction Date") !== -1 ? csvHeaders.indexOf("Transaction Date") : csvHeaders.indexOf("Posting Date");
      const colMerchant = csvHeaders.indexOf("Description");
      const colAmount = csvHeaders.indexOf("Amount");

      if (colDate === -1 || colMerchant === -1 || colAmount === -1) continue;

      // Dynamic account detection
      let detectedAccount = "Chase Checking";
      const filename = file.getName().toLowerCase();
      if (filename.indexOf("sapphire") !== -1 || filename.indexOf("credit") !== -1 || filename.indexOf("cc") !== -1) {
        detectedAccount = "Chase Sapphire";
      } else if (filename.indexOf("checking") !== -1 || filename.indexOf("debit") !== -1) {
        detectedAccount = "Chase Checking";
      }

      let fileRowsAdded = 0;
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        if (!row[colDate]) continue;

        const rawAmountStr = row[colAmount].toString().replace(/'/g, "").trim();
        const rawAmount = parseFloat(rawAmountStr);
        if (isNaN(rawAmount)) continue;

        const merchantRaw = row[colMerchant];
        const cleanMerchant = autoCleanMerchantName(merchantRaw);

        const sigId = "CHASE-CSV-" + row[colDate] + "-" + merchantRaw.substring(0, 10).replace(/\s/g, "") + "-" + Math.abs(rawAmount).toFixed(2);
        if (existingIDs.indexOf(sigId) !== -1) continue;

        const amountMagnitude = Math.abs(rawAmount);
        
        let flow = rawAmount < 0 ? "OUTFLOW" : "INFLOW";
        let suggestedCategory = autoCategorizeMerchant(cleanMerchant);
        
        if (cleanMerchant.toLowerCase().indexOf("transfer") !== -1 || cleanMerchant.toLowerCase().indexOf("payment thank you") !== -1) {
          flow = "TRANSFER";
          suggestedCategory = "Transfer";
        }

        dbSheet.appendRow([
          sigId,
          new Date(),
          new Date(row[colDate]),
          merchantRaw,
          cleanMerchant,
          amountMagnitude,
          flow,
          suggestedCategory,
          detectedAccount,
          "CHASE_CSV",
          "SETTLED"
        ]);

        existingIDs.push(sigId);
        fileRowsAdded++;
      }

      file.moveTo(archiveFolder);
      filesProcessed++;
      totalNewRows += fileRowsAdded;

    } catch (err) {
      console.error("Critical failure during batch run on file: " + file.getName() + " - " + err.toString());
    }
  }

  if (filesProcessed === 0) {
    ui.alert("Process Status ℹ️", "No pending .csv statement files were found inside Google Drive folder 'Chase_CSV_Uploads'.", ui.ButtonSet.OK);
  } else {
    ui.alert(
      "Batch Import Complete! 🎉",
      "Successfully processed " + filesProcessed + " CSV statements.\n\n" +
      "Added " + totalNewRows + " new transactions to your Transactions_DB tab.",
      ui.ButtonSet.OK
    );
  }
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
// 7. ADVANCED EMAIL ALERTS, DIGESTS & RULES ENGINES
// ===========================================================================

/**
 * Dynamic rules-based auto-categorizer.
 * Programmatically loads rules from Category_Rules_DB tab first to enforce no-code standards.
 * Falls back to built-in rules if table is empty or missing.
 */
function autoCategorizeMerchant(merchantRaw) {
  if (!merchantRaw) return "Miscellaneous";
  const merchant = merchantRaw.toLowerCase().trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  
  if (rulesSheet) {
    const lastRow = rulesSheet.getLastRow();
    if (lastRow >= 3) {
      const rulesData = rulesSheet.getRange(3, 1, lastRow - 2, 3).getValues();
      for (let i = 0; i < rulesData.length; i++) {
        const keyword = rulesData[i][0].toString().toLowerCase().trim();
        const categoryVal = rulesData[i][2].toString().trim();
        if (keyword !== "" && merchant.indexOf(keyword) !== -1) {
          return categoryVal;
        }
      }
    }
  }
  
  // Built-in hardcoded rules fallback
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
 * Standardizes merchant descriptions to clean names based on Category_Rules_DB tab entries.
 * For example: 'STARBUCKS STORE 1234' becomes 'Starbucks'
 */
function autoCleanMerchantName(merchantRaw) {
  if (!merchantRaw) return "Unknown Merchant";
  const merchant = merchantRaw.toLowerCase().trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  
  if (rulesSheet) {
    const lastRow = rulesSheet.getLastRow();
    if (lastRow >= 3) {
      const rulesData = rulesSheet.getRange(3, 1, lastRow - 2, 2).getValues();
      for (let i = 0; i < rulesData.length; i++) {
        const keyword = rulesData[i][0].toString().toLowerCase().trim();
        const cleanName = rulesData[i][1].toString().trim();
        if (keyword !== "" && merchant.indexOf(keyword) !== -1) {
          return cleanName;
        }
      }
    }
  }
  
  // Basic Regex cleansing to strip extra merchant ID strings
  let cleaned = merchantRaw.replace(/[0-9]+/g, "").replace(/\s+/g, " ").trim();
  // Capitalize first letters
  return cleaned.replace(/\b\w/g, c => c.toUpperCase());
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
  
  try {
    SpreadsheetApp.getUi().alert("Auto-Categorization Complete 🤖", "Successfully updated " + updatedCount + " transactions based on database rules-based merchant patterns.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.log("Trace: Auto-categorizer completed run. Updated rows: " + updatedCount);
  }
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
    Logger.log("Skipping budget breach checks for cumulative 'All' month selector.");
    return;
  }
  
  const lastEnumRow = enumSheet.getLastRow();
  if (lastEnumRow < 3) return;
  const enumCategoriesCount = lastEnumRow - 3; // Exclude Transfer
  
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
  const categoryCount = lastEnumRow - 3; // Excluding Transfer

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

  // Programmatically build HTML rows for categories
  let budgetRowsHtml = "";
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
      statusText = "⚡ Fast Pacing";
      statusColor = "#ff9900";
    } else {
      statusText = "On Track";
      statusColor = "#0f9d58";
    }

    budgetRowsHtml += `
      <tr style="border-bottom: 1px solid #f0f0f0;">
        <td style="padding: 10px 12px; font-weight: bold; color: #333333;">${category}</td>
        <td style="padding: 10px 12px; text-align: right; color: #002060; font-family: monospace;">$${planned.toFixed(2)}</td>
        <td style="padding: 10px 12px; text-align: right; color: #333333; font-family: monospace;">$${actual.toFixed(2)}</td>
        <td style="padding: 10px 12px; text-align: right; font-weight: bold; color: ${statusColor};">${statusText}</td>
      </tr>`;
  });

  // Construct HTML Newsletter Layout
  const htmlBody = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Weekly Financial Briefing</title>
</head>
<body style="margin: 0; padding: 0; background-color: #f4f6f9; font-family: Arial, sans-serif;">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f4f6f9; padding: 20px 10px;">
    <tr>
      <td align="center">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; border: 1px solid #e1e4e8; border-collapse: collapse; border-radius: 6px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
          <tr>
            <td style="background-color: #1f4e78; padding: 30px 20px; text-align: center;">
              <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: bold; letter-spacing: 1px; text-transform: uppercase;">Weekly Financial Briefing</h1>
              <p style="margin: 8px 0 0 0; color: #d9e1f2; font-size: 14px; font-weight: normal; letter-spacing: 0.5px;">
                Period Status: <strong>${month} ${year}</strong> &nbsp;|&nbsp; <strong>${monthElapsedPercent}%</strong> of Month Elapsed
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding: 25px 20px 15px 20px;">
              <h2 style="margin: 0 0 15px 0; color: #1f4e78; font-size: 16px; font-weight: bold; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px; text-transform: uppercase;">Month-To-Date KPIs</h2>
              <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td width="50%" style="padding: 0 10px 15px 0; box-sizing: border-box;">
                    <div style="background-color: #f8f9fa; border: 1px solid #e1e4e8; border-radius: 4px; padding: 15px; text-align: center;">
                      <span style="font-size: 11px; color: #6c757d; font-weight: bold; text-transform: uppercase; display: block; margin-bottom: 5px;">Total Income</span>
                      <strong style="font-size: 20px; color: #385723; display: block;">$${totalIncome.toFixed(2)}</strong>
                    </div>
                  </td>
                  <td width="50%" style="padding: 0 0 15px 10px; box-sizing: border-box;">
                    <div style="background-color: #f8f9fa; border: 1px solid #e1e4e8; border-radius: 4px; padding: 15px; text-align: center;">
                      <span style="font-size: 11px; color: #6c757d; font-weight: bold; text-transform: uppercase; display: block; margin-bottom: 5px;">Total Spending</span>
                      <strong style="font-size: 20px; color: #db4437; display: block;">$${totalSpending.toFixed(2)}</strong>
                    </div>
                  </td>
                </tr>
                <tr>
                  <td width="50%" style="padding: 5px 10px 0 0; box-sizing: border-box;">
                    <div style="background-color: #f8f9fa; border: 1px solid #e1e4e8; border-radius: 4px; padding: 15px; text-align: center;">
                      <span style="font-size: 11px; color: #6c757d; font-weight: bold; text-transform: uppercase; display: block; margin-bottom: 5px;">Net Savings</span>
                      <strong style="font-size: 20px; color: ${netSavings >= 0 ? "#385723" : "#db4437"}; display: block;">$${netSavings.toFixed(2)}</strong>
                    </div>
                  </td>
                  <td width="50%" style="padding: 5px 0 0 10px; box-sizing: border-box;">
                    <div style="background-color: #f8f9fa; border: 1px solid #e1e4e8; border-radius: 4px; padding: 15px; text-align: center;">
                      <span style="font-size: 11px; color: #6c757d; font-weight: bold; text-transform: uppercase; display: block; margin-bottom: 5px;">Savings Rate</span>
                      <strong style="font-size: 20px; color: #1f4e78; display: block;">${savingsRate}%</strong>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding: 15px 20px 25px 20px;">
              <h2 style="margin: 0 0 15px 0; color: #1f4e78; font-size: 16px; font-weight: bold; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px; text-transform: uppercase;">Budget Pacing Performance</h2>
              <table border="0" cellpadding="8" cellspacing="0" width="100%" style="border-collapse: collapse; font-size: 13px; text-align: left;">
                <thead>
                  <tr style="background-color: #1f4e78; color: #ffffff;">
                    <th style="border-radius: 4px 0 0 0; padding-left: 12px; font-weight: bold;">Category</th>
                    <th style="text-align: right; font-weight: bold;">Planned</th>
                    <th style="text-align: right; font-weight: bold;">Actual</th>
                    <th style="text-align: right; border-radius: 0 4px 0 0; padding-right: 12px; font-weight: bold;">Status</th>
                  </tr>
                </thead>
                <tbody>
                  ${budgetRowsHtml}
                </tbody>
              </table>
            </td>
          </tr>
          <tr>
            <td style="background-color: #f8f9fa; border-top: 1px solid #e1e4e8; padding: 25px 20px; text-align: center;">
              <p style="margin: 0 0 15px 0; font-size: 13px; color: #555555; line-height: 1.4;">
                This digest was programmatically compiled from your centralized transactional database. Open your dynamic interactive cockpit to modify historical thresholds, adjust allocations, or log cash adjustments.
              </p>
              <table border="0" cellpadding="0" cellspacing="0" style="margin: auto;">
                <tr>
                  <td align="center" style="border-radius: 4px; background-color: #1f4e78;">
                    <a href="${ss.getUrl()}" target="_blank" style="padding: 10px 20px; border: 1px solid #1f4e78; border-radius: 4px; font-family: Helvetica, Arial, sans-serif; font-size: 14px; font-weight: bold; color: #ffffff; text-decoration: none; display: inline-block;">Open Interactive Dashboard ➔</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="background-color: #1f4e78; padding: 15px 20px; text-align: center; border-radius: 0 0 6px 6px;">
              <span style="font-size: 11px; color: #a6b9d0; display: block;">
                Sent automatically by your <strong>Google Sheets Financial Automation System 🚀</strong>
              </span>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  MailApp.sendEmail({
    to: userEmail,
    subject: `📊 Weekly Financial Summary: ${month} ${year} Briefing`,
    htmlBody: htmlBody
  });

  Logger.log("Weekly budget summary email successfully dispatched.");
}

// ===========================================================================
// 8. DOUBLE-ENTRY MONTHLY SAVINGS SWEEP SYSTEM
// ===========================================================================

function promptSavingsSweep() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  if (!dash) {
    ui.alert("Configuration Error ⚠️", "Dashboard sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  const year = dash.getRange("C4").getValue().toString().trim();
  const month = dash.getRange("D4").getValue().toString().trim();
  
  if (month === "All") {
    ui.alert("Invalid Selection ⚠️", "Please select a specific calendar month (not 'All') on your Dashboard before executing the savings sweep.", ui.ButtonSet.OK);
    return;
  }
  
  const netSavings = parseFloat(dash.getRange("C7").getValue() || 0);
  if (netSavings <= 0) {
    ui.alert(
      "No Sweepable Surplus ℹ️",
      "Your Net Savings for " + month + " " + year + " is " + formatCurrency(netSavings) + ".\n" +
      "A sweep requires a positive surplus magnitude.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  const confirmMsg = "Do you want to execute your Monthly Savings Sweep for " + month + " " + year + "?\n\n" +
                     "Computed Surplus: " + formatCurrency(netSavings) + "\n\n" +
                     "This will close out the month and sweep your savings velocity across your goals based on your active allocations.";
  
  const response = ui.alert("Execute Savings Sweep? 💰", confirmMsg, ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    try {
      const summaryResult = executeSavingsSweep(ss, year, month, netSavings);
      ui.alert("Sweep Executed Successfully! 🎉", summaryResult, ui.ButtonSet.OK);
    } catch (err) {
      ui.alert("Sweep Refused ❌", err.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * Runs a transactional double-entry ledger allocation sweep of net income into Savings_Goals.
 */
function executeSavingsSweep(ss, year, month, netSavings) {
  const goalsSheet = ss.getSheetByName("Savings_Goals");
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!goalsSheet || !dbSheet) {
    throw new Error("Required database sheets ('Savings_Goals' or 'Transactions_DB') are missing.");
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const sweepKey = "sweep_run_" + year + "_" + month;
  const alreadyRun = scriptProperties.getProperty(sweepKey);
  
  if (alreadyRun === "true") {
    const forceConfirm = SpreadsheetApp.getUi().alert(
      "Duplicate Action Triggered ⚠️",
      "A savings sweep has already been logged for " + month + " " + year + ".\n\n" +
      "Rerunning the sweep will double-allocate and could distort balances.\n\n" +
      "Do you want to override and execute anyway?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (forceConfirm !== SpreadsheetApp.getUi().Button.YES) {
      throw new Error("Action cancelled by user to protect database integrity.");
    }
  }
  
  // Read allocation profiles from Savings_Goals sheet
  const goalsLastRow = goalsSheet.getLastRow();
  let data = [];
  if (goalsLastRow >= 3) {
    data = goalsSheet.getRange(3, 1, goalsLastRow - 2, 5).getValues();
  }
  
  if (data.length === 0) {
    throw new Error("No goals found on 'Savings_Goals' tab.");
  }
  
  // Verify allocation percentages total 100%
  let totalPercentage = 0;
  for (let i = 0; i < data.length; i++) {
    totalPercentage += parseFloat(data[i][4]) || 0;
  }
  
  if (Math.abs(totalPercentage - 1.0) > 0.001) {
    throw new Error("Total sweep allocation percentages must equal exactly 100% (" + (totalPercentage * 100).toFixed(0) + "% computed). Please correct your allocations.");
  }
  
  // Resolve month indices for chronological timestamping
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthIdx = monthNames.indexOf(month);
  const sweepDate = new Date(parseInt(year, 10), monthIdx + 1, 0); // Last day of that billing month
  
  let logSummary = "Allocated " + formatCurrency(netSavings) + " as follows:";
  let entriesLogged = 0;
  
  for (let i = 0; i < data.length; i++) {
    const goalName = data[i][0];
    const currentBal = parseFloat(data[i][2]) || 0;
    const sweepPct = parseFloat(data[i][4]) || 0;
    
    if (sweepPct <= 0) continue;
    
    const allocationAmount = netSavings * sweepPct;
    
    // Programmatic update of balances in goals tab
    const rowIdx = 3 + i;
    goalsSheet.getRange(rowIdx, 3).setValue(currentBal + allocationAmount);
    
    // DOUBLE-ENTRY BOOKKEEPING TRANSACTION CORRESPONDENCE
    const timestamp = new Date();
    const batchId = "SWEEP-" + year + "-" + month.substring(0,3).toUpperCase() + "-" + timestamp.getTime();
    
    // Debit Funding Account (OUTFLOW)
    dbSheet.appendRow([
      batchId + "-DEBIT-" + i,
      timestamp,
      sweepDate,
      "REALLOCATE SURPLUS TO " + goalName.toUpperCase(),
      "Sweep to " + goalName,
      allocationAmount,
      "OUTFLOW",
      "Savings",
      "Chase Checking",
      "WEB_FORM",
      "SETTLED"
    ]);
    
    // Credit Savings Destination (INFLOW)
    dbSheet.appendRow([
      batchId + "-CREDIT-" + i,
      timestamp,
      sweepDate,
      "RECEIVE ALLOCATION FOR " + goalName.toUpperCase(),
      "Sweep to " + goalName,
      allocationAmount,
      "INFLOW",
      "Savings",
      "Savings",
      "WEB_FORM",
      "SETTLED"
    ]);
    
    logSummary += "\n• " + goalName + ": " + formatCurrency(allocationAmount) + " (" + (sweepPct * 100).toFixed(0) + "%)";
    entriesLogged += 2;
  }
  
  if (entriesLogged > 0) {
    scriptProperties.setProperty(sweepKey, "true");
    return logSummary;
  } else {
    return "No active allocations were computed. Ensure your goals have non-zero sweep percentages.";
  }
}

function formatCurrency(amount) {
  return "$" + parseFloat(amount).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

// ===========================================================================
// 9. HIGH-PERFORMANCE DYNAMIC TRANSACTION EDITOR SIDEBAR
// ===========================================================================

function showTransactionEditorSidebar() {
  const html = HtmlService.createTemplateFromFile("TransactionEditorView");
  const htmlOutput = html.evaluate()
    .setTitle("Transaction Editor Ledger")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Sidebar Endpoint: Returns a subset of transactions from Transactions_DB to feed the pagination table.
 * Includes optional fuzzy search queries.
 */
function getTransactionsPage(pageNum, pageSize, searchQuery) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return { data: [], totalPages: 0 };
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return { data: [], totalPages: 0 };
  
  const rawValues = dbSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const query = (searchQuery || "").toLowerCase().trim();
  
  // Sort chronologically descending, with transaction row indexing preserved
  const mapped = rawValues.map((row, index) => {
    return {
      rowIdx: index + 2,
      txId: row[0],
      timestamp: row[1],
      date: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
      merchantRaw: row[3],
      merchantClean: row[4],
      amount: parseFloat(row[5]) || 0,
      flow: row[6],
      category: row[7],
      source: row[8],
      channel: row[9],
      status: row[10]
    };
  }).reverse(); // Most recent first
  
  const filtered = mapped.filter(tx => {
    if (query === "") return true;
    return (tx.merchantRaw || "").toLowerCase().indexOf(query) !== -1 ||
           (tx.merchantClean || "").toLowerCase().indexOf(query) !== -1 ||
           (tx.category || "").toLowerCase().indexOf(query) !== -1 ||
           (tx.source || "").toLowerCase().indexOf(query) !== -1;
  });
  
  const totalItems = filtered.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  
  const startIdx = (pageNum - 1) * pageSize;
  const pageData = filtered.slice(startIdx, startIdx + pageSize);
  
  return {
    data: pageData,
    totalPages: totalPages,
    totalItems: totalItems
  };
}

/**
 * Sidebar Endpoint: Safely fetches category settings to build select controls.
 */
function getEnumValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  if (!enumSheet) return { categories: [], accounts: [] };
  
  const lastRow = enumSheet.getLastRow();
  if (lastRow < 3) return { categories: [], accounts: [] };
  
  const cats = enumSheet.getRange(3, 1, lastRow - 2, 1).getValues().flat().filter(String);
  const accs = enumSheet.getRange(3, 2, enumSheet.getRange("B3:B").getLastRow() - 2, 1).getValues().flat().filter(String);
  
  return {
    categories: cats,
    accounts: accs
  };
}

/**
 * Sidebar Endpoint: Commits transaction creations or edits back to the sheet cleanly.
 */
function addOrUpdateTransaction(rowIdx, txData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return "ERROR: Transactions_DB tab missing.";
  
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const txDate = new Date(txData.date);
    const amount = Math.abs(parseFloat(txData.amount)) || 0.00;
    const flow = txData.flow;
    const category = txData.category;
    const source = txData.source;
    const merchantRaw = txData.merchantRaw;
    const cleanMerchant = autoCleanMerchantName(merchantRaw);
    const channel = txData.channel || "SIDEBAR_EDITOR";
    const status = txData.status || "SETTLED";
    
    if (rowIdx && parseInt(rowIdx, 10) > 1) {
      // IDEMPOTENT Edit existing row
      const targetRow = parseInt(rowIdx, 10);
      const txId = dbSheet.getRange(targetRow, 1).getValue() || ("MANUAL-EDIT-" + new Date().getTime());
      
      dbSheet.getRange(targetRow, 1, 1, 11).setValues([[
        txId,
        new Date(), // Update timestamp of edit
        txDate,
        merchantRaw,
        cleanMerchant,
        amount,
        flow,
        category,
        source,
        channel,
        status
      ]]);
      console.log(`[Transaction Editor Sidebar] Successfully updated transaction at Row ${targetRow}.`);
      return "SUCCESS: Transaction updated.";
    } else {
      // Append a fresh transaction row
      const txId = "MANUAL-" + new Date().getTime();
      dbSheet.appendRow([
        txId,
        new Date(),
        txDate,
        merchantRaw,
        cleanMerchant,
        amount,
        flow,
        category,
        source,
        channel,
        status
      ]);
      console.log("[Transaction Editor Sidebar] Successfully appended new manual transaction.");
      return "SUCCESS: Transaction added.";
    }
  } catch (err) {
    return "ERROR: " + err.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sidebar Endpoint: Deletes a transaction row safely.
 */
function deleteTransactionRow(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return "ERROR: Transactions_DB tab missing.";
  
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const targetRow = parseInt(rowIdx, 10);
    if (targetRow > 1) {
      dbSheet.deleteRow(targetRow);
      console.log(`[Transaction Editor Sidebar] Successfully deleted transaction at Row ${targetRow}.`);
      return "SUCCESS";
    }
    return "ERROR: Invalid row selected.";
  } catch (err) {
    return "ERROR: " + err.toString();
  } finally {
    lock.releaseLock();
  }
}
