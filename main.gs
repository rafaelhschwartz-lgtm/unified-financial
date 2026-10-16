/**
 *  ===========================================================================
 *  GOOGLE SHEETS FINANCIAL AUTOMATION SYSTEM (UNIFIED MASTER ENGINE - V5)
 *  Standardized under the FAST Modeling Standard and ICAEW Spreadsheet Principles.
 *  
 *  This master script consolidates all automation channels, triggers, layouts,
 *  automated categorizations, double-entry sweeps, email alerts, HTML digest
 *  summaries, and dynamic visualizations into a single high-performance system.
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
        console.log(`[onEdit Log] Period selection altered on Dashboard. Refreshing budget history...`);
        syncDashboardBudgets(sheet);
      } catch (err) {
        console.error("[onEdit Error] Failed to sync historical budgets: " + err.toString());
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
  
  const sheetNames = ["About", "Enum", "Transactions_DB", "Budget_History_DB", "The Monthly Dashboard Tab", "Control", "Category_Rules_DB", "Savings_Goals"];
  const sheets = {};
  
  // Reset or create each sheet
  sheetNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      console.log(`[Whole Book Compiler] Inserted missing tab: ${name}`);
    } else {
      sheet.clear();
      if (sheet.getFilter()) sheet.getFilter().remove();
      sheet.getDataRange().setDataValidation(null);
      console.log(`[Whole Book Compiler] Cleared and prepared existing tab: ${name}`);
    }
    sheets[name] = sheet;
  });
  
  // Arrange sheet ordering chronologically by functional class
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
  
  const channelValidation = SpreadsheetApp.newDataValidation().requireValueInList(["WEB_FORM", "CHASE_CSV", "GMAIL_ALERT", "PLAID_API", "MANUAL", "SAVINGS_SWEEP"]).setAllowInvalid(false).build();
  db.getRange("J2:J1000").setDataValidation(channelValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation().requireValueInList(["SETTLED", "PENDING", "MANUAL_REVIEW"]).setAllowInvalid(false).build();
  db.getRange("K2:K1000").setDataValidation(statusValidation);
  
  // Standard dummy row to initialize formulas
  db.getRange("A2:K2").setValues([["INIT_SETUP_DUMMY", new Date(), new Date(), "INITIAL SYSTEM SETUP", "Initial Setup", 0.01, "INFLOW", "Miscellaneous", "Cash", "MANUAL", "SETTLED"]]);
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
  // CATEGORY RULES DATABASE TAB (Category_Rules_DB)
  // -------------------------------------------------------------------------
  const rulesDb = sheets["Category_Rules_DB"];
  formatHeader(rulesDb, "A1:B1", "AI RULES-BASED CATEGORIZATION LOOKUP DATABASE");
  rulesDb.getRange("A2:B2").setValues([["Merchant Keyword Rule", "Budget Category Target"]]).setFontWeight("bold");
  formatSubHeader(rulesDb, "A2:B2");
  rulesDb.setRowHeight(2, 25);
  
  const startingRules = [
    ["starbucks", "Food & Drink"],
    ["mcdonald", "Food & Drink"],
    ["dunkin", "Food & Drink"],
    ["grubhub", "Food & Drink"],
    ["ubereats", "Food & Drink"],
    ["doorndash", "Food & Drink"],
    ["kroger", "Groceries"],
    ["wholefd", "Groceries"],
    ["safeway", "Groceries"],
    ["trader joe", "Groceries"],
    ["aldi", "Groceries"],
    ["costco", "Groceries"],
    ["netflix", "Entertainment"],
    ["spotify", "Entertainment"],
    ["hulu", "Entertainment"],
    ["uber", "Transportation"],
    ["lyft", "Transportation"],
    ["chevron", "Transportation"],
    ["shell", "Transportation"],
    ["comcast", "Bills & Utilities"],
    ["coned", "Bills & Utilities"],
    ["verizon", "Bills & Utilities"],
    ["amazon", "Home"],
    ["target", "Home"],
    ["delta", "Travel"],
    ["airbnb", "Travel"],
    ["nike", "Personal"],
    ["cvs", "Personal"]
  ];
  rulesDb.getRange(3, 1, startingRules.length, 2).setValues(startingRules);
  const categoryEnumRange = enumSheet.getRange("A3:A11");
  const validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(categoryEnumRange).setAllowInvalid(false).build();
  rulesDb.getRange("B3:B1000").setDataValidation(validationRule);
  rulesDb.getRange("A3:A1000").setFontColor("#002060"); // Blue for user inputs
  rulesDb.autoResizeColumns(1, 2);

  // -------------------------------------------------------------------------
  // SAVINGS GOALS TAB & MULTI-YEAR SAVINGS PROJECTION CHART
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
  
  // -- Programmatic Multi-Year Savings Projection Model --
  // Setup editable assumption parameters in columns H & I
  goalsSheet.getRange("H1:I1").setValues([["Projection Parameter", "User Allocation Value"]]).setFontWeight("bold");
  formatSubHeader(goalsSheet, "H1:I1");
  goalsSheet.getRange("H2:I2").setValues([["Assumed Annual Growth Rate", 0.06]]);
  goalsSheet.getRange("H3:I3").setValues([["Annual Goals Savings Contribution", 6000.00]]);
  goalsSheet.getRange("I2").setNumberFormat("0.0%").setFontColor("#002060");
  goalsSheet.getRange("I3").setNumberFormat("$#,##0.00").setFontColor("#002060");
  
  // Construct 10-year projection ledger table
  goalsSheet.getRange("H5:L5").setValues([["Year", "Starting Balance", "Annual Savings Contribution", "Compound Interest", "Projected Ending Balance"]]).setFontWeight("bold");
  formatSubHeader(goalsSheet, "H5:L5");
  
  // Year 1 (2026) dynamic calculations
  goalsSheet.getRange("H6").setValue(2026);
  goalsSheet.getRange("I6").setFormula("=SUM(C3:C5)"); // Starts from sum of current goal balances
  goalsSheet.getRange("J6").setFormula("=$I$3");
  goalsSheet.getRange("K6").setFormula("=(I6+J6)*$I$2");
  goalsSheet.getRange("L6").setFormula("=SUM(I6:K6)");
  
  // Years 2 to 10 calculations loop
  for (let y = 1; y < 10; y++) {
    const rowIdx = 6 + y;
    goalsSheet.getRange(rowIdx, 1).setValue(2026 + y); // Year
    goalsSheet.getRange(rowIdx, 2).setFormula(`=L${rowIdx - 1}`); // Starts from previous Year's ending
    goalsSheet.getRange(rowIdx, 3).setFormula("=$I$3"); // Dynamic addition reference
    goalsSheet.getRange(rowIdx, 4).setFormula(`=(I${rowIdx}+J${rowIdx})*$I$2`); // Dynamic rate growth
    goalsSheet.getRange(rowIdx, 5).setFormula(`=SUM(I${rowIdx}:K${rowIdx})`); // Dynamic ending total
  }
  goalsSheet.getRange("I6:L15").setNumberFormat("$#,##0.00");
  goalsSheet.getRange("H6:H15").setNumberFormat("0000").setHorizontalAlignment("center");
  
  // Insert Multi-Year Savings Projection Graph (Area Chart)
  try {
    const projectionChart = goalsSheet.newChart()
      .setChartType(Charts.ChartType.AREA)
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
    console.log(`[Whole Book Compiler] Programmatically compiled 10-year Area Chart on Savings_Goals.`);
  } catch (err) {
    console.error("[Whole Book Compiler Error] Projection graph creation failed: " + err.toString());
  }

  goalsSheet.autoResizeColumns(1, 12);

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
  const enumCategoriesCount = categories.length; // N=9
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
  
  // Dynamic charts generation
  try {
    const existingCharts = dash.getCharts();
    existingCharts.forEach(c => dash.removeChart(c));
    
    // Chart 1: Budget vs. Actual (Side-by-Side Column Chart)
    const budgetVsActualChart = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
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
    
    // Chart 2: Category Expense Allocation (Donut Chart)
    const expenseBreakdownChart = dash.newChart()
      .setChartType(Charts.ChartType.PIE)
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
  
  console.log("[Whole Book Compiler] Finished compiler run. Workbook structure is fully operational.");
}

/**
 * Applies Light Red/Green conditional colors programmatically to Control!D4:D6.
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
  console.log("[Format Engine] Applied conditional formatting rules successfully to Control!D4:D6.");
}

// ===========================================================================
// 3. BACKGROUND SYNCHRONIZER (Budget_History_DB -> Dashboard)
// ===========================================================================

/**
 * Queries Budget_History_DB and synchronizes the values into Column B of Dashboard.
 */
function syncDashboardBudgets(dash) {
  const ss = dash.getParent();
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  const enumSheet = ss.getSheetByName("Enum");
  if (!budgetDB || !enumSheet) return;
  
  // Retrieve selected Period
  const year = parseInt(dash.getRange("C4").getValue(), 10);
  const month = dash.getRange("D4").getValue().toString().trim();
  
  // Retrieve active categories from Enum sheet
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
  console.log("[Budget Sync Engine] Synchronized " + numCategories + " category budgets for Period: " + year + " " + month);
}

// ===========================================================================
// 4. SIDEBAR BUDGET CONTROLLER ENGINE
// ===========================================================================

/**
 * Displays the custom sidebar to edit historical monthly budgets.
 */
function showBudgetSidebar() {
  const html = HtmlService.createTemplateFromFile("SidebarView");
  const htmlOutput = html.evaluate()
    .setTitle("Historical Budget Allocator")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  console.log("[UI Service] Rendered Budget History Allocation Sidebar to view.");
}

/**
 * Sidebar Endpoint: Returns active categories and their currently loaded budgets for Year & Month.
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
 * Sidebar Endpoint: IDEMPOTENT update/insert of budgets into Budget_History_DB.
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
        dbData[matchIndex][3] = parseFloat(item.amount);
      } else {
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
    console.log(`[Budget History Admin] Logged ${budgetData.length} budget updates into history DB.`);
    return "SUCCESS";
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// 5. EXISTING AUTOMATION CHANNELS (Consolidated & Cleaned)
// ===========================================================================

/**
 * Handles incoming POST requests from manual external HTML forms.
 * Normalizes values, runs auto-categorization automatically, and logs cleanly.
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
    
    const cleanMerchant = rowData.Merchant.toString().trim();
    const resolvedCategory = rowData.Category && rowData.Category !== "" ? rowData.Category : autoCategorizeMerchant(cleanMerchant);
    
    sheet.appendRow([
      transactionId,
      timestamp,
      dateValue,
      rowData.Merchant,
      cleanMerchant,
      Math.abs(parseFloat(rowData.Amount)),
      rowData.Flow,
      resolvedCategory,
      rowData.Source,
      "WEB_FORM",
      "SETTLED"
    ]);
    
    console.log(`[POST Ingestion Log] Ingested manual transaction: ${transactionId} - Merchant: ${cleanMerchant} - Amount: $${rowData.Amount}`);
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'id': transactionId })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("[POST Ingestion Error] API insertion crashed: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'error': error.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Automatically parses unread Chase transaction alert emails from Gmail.
 * Runs keyword auto-categorization rules before appending rows.
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
  
  console.log(`[Gmail Scraper Engine] Searching unread transaction emails query: "${searchQuery}"`);
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const msgId = msg.getId();
      
      if (existingIDs.indexOf(msgId) !== -1) continue; // Deduplication check
      
      const body = msg.getPlainBody();
      const amountMatch = body.match(/\$(\d+\.\d{2})/);
      if (!amountMatch) continue;
      
      const amount = parseFloat(amountMatch[1]);
      const date = msg.getDate();
      
      // Parse merchant if regex pattern allows (fallback default description)
      const merchantClean = "Chase Credit Card Purchase";
      const suggestedCategory = autoCategorizeMerchant(merchantClean);
      
      dbSheet.appendRow([
        msgId,
        new Date(),
        date,
        "CHASE TRANSACTION ALERT",
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
      console.log(`[Gmail Scraper Engine] Ingested email message alert ${msgId} for $${amount}`);
    }
  }
  console.log("[Gmail Scraper Engine] Parsing run completed. Imported " + newRowsCount + " new alerts.");
}

/**
 * Standardized single Chase CSV prompts.
 */
function promptForChaseCSV() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Import Chase CSV Statement",
    "Please enter the Google Drive File ID of your Chase statement .csv file:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
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
 * Imports and normalizes Chase credit/debit statements from standard Drive file.
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
    
    const sigId = "CHASE-CSV-" + row[colDate] + "-" + row[colMerchant].substring(0, 10).replace(/\s/g, "") + "-" + Math.abs(rawAmount).toFixed(2);
    if (existingIDs.indexOf(sigId) !== -1) continue;
    
    const cleanMerchant = row[colMerchant].toString().trim();
    const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
    
    dbSheet.appendRow([
      sigId,
      new Date(),
      new Date(row[colDate]),
      row[colMerchant],
      cleanMerchant,
      Math.abs(rawAmount),
      rawAmount < 0 ? "OUTFLOW" : "INFLOW",
      suggestedCategory,
      "Chase Checking",
      "CHASE_CSV",
      "SETTLED"
    ]);
    newRowsCount++;
  }
  console.log(`[Chase Statement Ingestion] Upload finished. Ingested ${newRowsCount} transactions.`);
  return "Successfully imported " + newRowsCount + " transaction rows.";
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

  console.log(`[Batch CSV Ingestion] Checking folder "Chase_CSV_Uploads" for statement files...`);

  // Sequential file processing loop
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

      if (colDate === -1 || colMerchant === -1 || colAmount === -1) {
        console.warn(`[Batch CSV Ingestion] Skipping file ${file.getName()} because expected headers are missing.`);
        continue;
      }

      let fileRowsAdded = 0;
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        if (!row[colDate]) continue;

        const rawAmountStr = row[colAmount].toString().replace(/'/g, "").trim();
        const rawAmount = parseFloat(rawAmountStr);
        if (isNaN(rawAmount)) continue;

        const sigId = "CHASE-CSV-" + row[colDate] + "-" + row[colMerchant].substring(0, 10).replace(/\s/g, "") + "-" + Math.abs(rawAmount).toFixed(2);
        if (existingIDs.indexOf(sigId) !== -1) continue;

        const cleanMerchant = row[colMerchant].toString().trim();
        const suggestedCategory = autoCategorizeMerchant(cleanMerchant);

        dbSheet.appendRow([
          sigId,
          new Date(),
          new Date(row[colDate]),
          row[colMerchant],
          cleanMerchant,
          Math.abs(rawAmount),
          rawAmount < 0 ? "OUTFLOW" : "INFLOW",
          suggestedCategory,
          "Chase Checking",
          "CHASE_CSV",
          "SETTLED"
        ]);

        existingIDs.push(sigId);
        fileRowsAdded++;
      }

      // Safely archive the processed statement
      file.moveTo(archiveFolder);
      filesProcessed++;
      totalNewRows += fileRowsAdded;
      console.log(`[Batch CSV Ingestion] Successfully imported statement: ${file.getName()} (${fileRowsAdded} rows added).`);

    } catch (err) {
      console.error("[Batch CSV Ingestion Error] Critical crash processing: " + file.getName() + " -> " + err.toString());
    }
  }

  if (filesProcessed === 0) {
    ui.alert("Process Status ℹ️", "No pending .csv statement files were found inside 'Chase_CSV_Uploads' folder.", ui.ButtonSet.OK);
  } else {
    ui.alert(
      "Batch Import Complete! 🎉",
      "Successfully processed " + filesProcessed + " CSV statements.\n\n" +
      "Added " + totalNewRows + " new transactions to your Transactions_DB tab.\n\n" +
      "Processed files have been safely archived to prevent future double-counting.",
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
 * Scrapes merchant names, checks database rules on sheet 'Category_Rules_DB' first,
 * and falls back to standard dictionary mapping if not found.
 */
function autoCategorizeMerchant(merchantRaw) {
  if (!merchantRaw) return "Miscellaneous";
  const merchant = merchantRaw.toLowerCase().trim();
  
  // 1. Try querying real-time database rules sheet Category_Rules_DB first
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName("Category_Rules_DB");
    if (rulesSheet) {
      const lastRow = rulesSheet.getLastRow();
      if (lastRow >= 3) {
        const rulesData = rulesSheet.getRange(3, 1, lastRow - 2, 2).getValues();
        for (let r = 0; r < rulesData.length; r++) {
          const keyword = rulesData[r][0].toString().toLowerCase().trim();
          const targetCat = rulesData[r][1].toString().trim();
          if (keyword !== "" && merchant.indexOf(keyword) !== -1) {
            console.log(`[Categorization Match] Found database rule: "${keyword}" -> "${targetCat}" for merchant "${merchantRaw}"`);
            return targetCat;
          }
        }
      }
    }
  } catch (dbErr) {
    console.warn("[Categorization Warning] Database rules search failed, falling back to static rules. Details: " + dbErr.toString());
  }

  // 2. Built-in Fallback Rules Dictionary (Guarantees zero-failure code)
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
 * Also cleans "Miscellaneous" defaults if database keywords are found.
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
  
  console.log(`[Auto-Categorization] Processing ${data.length} transactions in the database...`);
  
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
  
  console.log(`[Auto-Categorization] Run finished. Categorized ${updatedCount} transactions.`);
  SpreadsheetApp.getUi().alert("Auto-Categorization Complete 🤖", "Successfully updated " + updatedCount + " transactions based on database-driven category matching.", SpreadsheetApp.getUi().ButtonSet.OK);
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
  
  const dataRange = dash.getRange(10, 1, enumCategoriesCount, 4);
  const data = dataRange.getValues();
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const userEmail = Session.getActiveUser().getEmail();
  let breaches = [];
  
  console.log(`[Budget Breaches] Analyzing actual spending against planned thresholds for Period: ${month} ${year}`);
  
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
    
    console.log(`[Budget Breaches] Dispatched budget breach emails for ${breaches.length} categories.`);
    try {
      SpreadsheetApp.getUi().alert("Email Alerts Dispatched! ✉️", "Dispatched " + breaches.length + " breach warning(s) to " + userEmail, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  } else {
    console.log("[Budget Breaches] Checked limits. No new budget breaches detected.");
    try {
      SpreadsheetApp.getUi().alert("Budget Check Complete ✅", "All actual spending remains within planned monthly boundaries. No alerts sent.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }
}

/**
 * Compiles a weekly budget summary report and dispatches a rich, structured HTML email.
 * Evaluates monthly spending velocity and category status.
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

  // Retrieve budget items (Category, Planned, Actual)
  const budgetData = dash.getRange(10, 1, categoryCount, 3).getValues();

  // Determine current progress through the month
  const today = new Date();
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const currentDay = today.getDate();
  const monthElapsedPercent = ((currentDay / daysInMonth) * 100).toFixed(0);

  console.log(`[Weekly Briefing Email] Compiling financial performance breakdown...`);

  // Construct HTML Newsletter
  let html = `
  <div style="font-family: Arial, sans-serif; background-color: #f4f6f9; padding: 20px; max-width: 600px; margin: auto; border: 1px solid #e1e4e8;">
    <div style="background-color: #1f4e78; color: white; padding: 25px; text-align: center; border-radius: 5px 5px 0 0;">
      <h2 style="margin: 0; font-size: 22px; letter-spacing: 1px;">WEEKLY FINANCIAL BRIEFING</h2>
      <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Period Status: ${month} ${year} (${monthElapsedPercent}% Elapsed)</p>
    </div>

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

    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Budget Performance Velocity</h3>
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
    let statusColor = "#4285f4";

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

  console.log(`[Weekly Briefing Email] Weekly budget summary successfully dispatched to ${userEmail}`);
  try {
    SpreadsheetApp.getUi().alert("Briefing Email Sent! 📬", "Successfully compiled and dispatched HTML briefing newsletter to " + userEmail, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// ===========================================================================
// 8. DOUBLE-ENTRY MONTHLY INCOME SAVINGS SWEEP SYSTEM
// ===========================================================================

/**
 * Triggers interactive confirmation to execute monthly savings sweep.
 */
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
      "No Surplus Found 📉",
      "Your calculated Net Savings for " + month + " " + year + " is " + formatCurrency(netSavings) + ".\n" +
      "A sweep requires a positive surplus magnitude.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  const response = ui.alert(
    "Confirm Savings Sweep 💰",
    "Your Net Savings surplus for " + month + " " + year + " is " + formatCurrency(netSavings) + ".\n\n" +
    "This will close out the month and sweep your savings velocity across your goals based on your active allocations.\n\n" +
    "Do you want to proceed with executing the double-entry transactions sweep?",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      const result = executeSavingsSweep(ss, year, month, netSavings);
      ui.alert("Sweep Success! 🎉", result, ui.ButtonSet.OK);
    } catch (err) {
      ui.alert("Sweep Action Crashed ❌", err.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * Executes a relational savings sweep by pulling allocation percentages from Savings_Goals,
 * calculating split amounts, and appending double-entry offsetting transaction pairs inside Transactions_DB.
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
  
  // Safety prompt warning to prevent duplicate allocation sweeps
  if (alreadyRun) {
    const ui = SpreadsheetApp.getUi();
    const safetyCheck = ui.alert(
      "Duplicate Sweep Warning ⚠️",
      "Our background State database flags that a Monthly Savings Sweep was ALREADY executed for " + month + " " + year + ".\n\n" +
      "Rerunning the sweep will double-allocate and could distort balances. Are you absolutely sure you want to proceed?",
      ui.ButtonSet.YES_NO
    );
    if (safetyCheck !== ui.Button.YES) {
      throw new Error("Action cancelled by user.");
    }
  }
  
  const lastGoalRow = goalsSheet.getLastRow();
  if (lastGoalRow < 3) {
    throw new Error("No savings goals defined in 'Savings_Goals' tab.");
  }
  
  const data = goalsSheet.getRange(3, 1, lastGoalRow - 2, 5).getValues();
  let totalPercentage = 0;
  data.forEach(row => totalPercentage += parseFloat(row[4] || 0));
  
  if (Math.abs(totalPercentage - 1.0) > 0.001) {
    throw new Error("Total sweep allocation percentages exceed 100% (" + (totalPercentage * 100).toFixed(0) + "%). Please correct your allocations on Savings_Goals.");
  }
  
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthIdx = months.indexOf(month);
  const sweepDate = new Date(parseInt(year, 10), monthIdx + 1, 0); // Last day of that billing month
  
  let sweepRowsCount = 0;
  let logSummary = "Allocations executed successfully:";
  
  for (let i = 0; i < data.length; i++) {
    const goalName = data[i][0];
    const sweepPct = parseFloat(data[i][4]) || 0;
    
    if (sweepPct <= 0) continue;
    
    const allocationAmount = netSavings * sweepPct;
    const baseId = "SWEEP-" + year + "-" + month.substring(0,3).toUpperCase() + "-" + goalName.toUpperCase().replace(/\s/g, "");
    
    // Offset Transaction 1: DEBIT/OUTFLOW from Chase Checking Account
    const debitId = baseId + "-DEBIT";
    dbSheet.appendRow([
      debitId,
      new Date(),
      sweepDate,
      "SAVINGS SWEEP ALLOCATION - " + goalName.toUpperCase(),
      "Sweep To " + goalName,
      allocationAmount,
      "OUTFLOW",
      "Savings",
      "Chase Checking",
      "SAVINGS_SWEEP",
      "SETTLED"
    ]);
    
    // Offset Transaction 2: CREDIT/INFLOW to Savings Account
    const creditId = baseId + "-CREDIT";
    dbSheet.appendRow([
      creditId,
      new Date(),
      sweepDate,
      "SAVINGS SWEEP ALLOCATION - " + goalName.toUpperCase(),
      "Sweep To " + goalName,
      allocationAmount,
      "INFLOW",
      "Savings",
      "Savings",
      "SAVINGS_SWEEP",
      "SETTLED"
    ]);
    
    // Update the Current Balance on Goals sheet directly in-place
    const goalRowCell = goalsSheet.getRange(3 + i, 3);
    const prevBalance = parseFloat(goalRowCell.getValue() || 0);
    goalRowCell.setValue(prevBalance + allocationAmount);
    
    sweepRowsCount += 2;
    logSummary += "\n• " + goalName + ": " + formatCurrency(allocationAmount) + " (" + (sweepPct * 100).toFixed(0) + "%)";
  }
  
  if (sweepRowsCount > 0) {
    scriptProperties.setProperty(sweepKey, "true");
    console.log(`[Sweep Service] Successfully ran monthly surplus sweep for ${month} ${year}. Injected ${sweepRowsCount} double-entry rows.`);
    return logSummary;
  } else {
    return "No active allocations were computed. Ensure your goals have non-zero sweep percentages.";
  }
}

function formatCurrency(amount) {
  return "$" + parseFloat(amount).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}
