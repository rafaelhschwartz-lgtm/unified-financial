

// ===========================================================================
// GLOBAL SHEET FORMATTING HELPERS
// ===========================================================================

/**
 * Global helper: Formats a merged title header banner on a sheet.
 */
function formatHeader(sheet, rangeString, title) {
  if (!sheet) return;
  const range = sheet.getRange(rangeString);
  range.merge();
  range.setValue(title);
  range.setFontFamily("Arial")
       .setFontSize(14)
       .setFontWeight("bold")
       .setFontColor("#ffffff")
       .setBackground("#1f4e78")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
  sheet.setRowHeight(range.getRow(), 40);
}

/**
 * Global helper: Formats a sub-header row on a sheet.
 */
function formatSubHeader(sheet, rangeString) {
  if (!sheet) return;
  const range = sheet.getRange(rangeString);
  range.setFontFamily("Arial")
       .setFontSize(10)
       .setFontWeight("bold")
       .setFontColor("#000000")
       .setBackground("#d9e1f2")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
  sheet.setRowHeight(range.getRow(), 25);
}

/**
 *  ===========================================================================
 *  GOOGLE SHEETS FINANCIAL AUTOMATION SYSTEM (UNIFIED MASTER ENGINE - V12)
 *  Standardized under the FAST Modeling Standard and ICAEW Spreadsheet Principles.
 *  
 *  V12 Master System Features:
 *    1. Consolidated trigger system (onOpen, onEdit).
 *    2. In-place workbook tab initialization (About, Enum, Transactions_DB,
 *       Budget_History_DB, Dashboard, Control, Category_Rules_DB, Savings_Goals).
 *    3. Programmatic Dynamic Visualizations (Side-by-Side Column and Pie/Donut Charts).
 *    4. Gmail purchase alerts scraper AND Google Voice SMS transaction alerts scraper.
 *    5. Multi-account automatic detection for folder-based CSV statement ingesters.
 *    6. Database-driven No-Code Rules auto-categorizer (Category_Rules_DB).
 *    7. Automatic Category & Account auto-expansion in Enum with ZERO validation errors!
 *    8. Multi-pattern advanced merchant string cleansing & title-casing.
 *    9. 'TRANSFER' flow type configuration to ignore internal transfer distortion.
 *   10. Interactive double-entry Monthly Savings Sweep (Savings_Goals -> Transactions_DB).
 *   11. 10-Year growth projection matrix and live Area Chart.
 *   12. Dynamic, paginated In-Sheet Transaction Editor Sidebar.
 *   13. Visual Categorization Assistant Sidebar with 1-click batch rule learning.
 *   14. Isolated Trip & Project Expense Tracker Engine.
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
    .addItem("Open Categorization Assistant 🤖", "showCategorizerSidebar")
    .addItem("Open Transaction Editor Sidebar", "showTransactionEditorSidebar")
    .addItem("Open Budget History Sidebar", "showBudgetSidebar")
    .addSeparator()
    .addItem("Batch Import CSV Folder", "importAllPendingCSVs")
    .addItem("Run AI Auto-Categorization", "autoCategorizeUnassigned")
    .addItem("Execute Monthly Savings Sweep", "promptSavingsSweep")
    .addSeparator()
    .addSubMenu(ui.createMenu("➕ Quick Add & Tools")
      .addItem("➕ Quick Add Transaction", "quickAddTransactionPrompt")
      .addItem("➕ Create Trip / Project Budget", "promptTripBudget")
      .addItem("➕ Add New Budget Category", "quickAddCategoryPrompt")
      .addItem("➕ Add New Account Source", "quickAddAccountPrompt")
      .addItem("➕ Add Merchant Category Rule", "quickAddRulePrompt"))
    .addSeparator()
    .addItem("Check Budget Breaches", "checkBudgetBreaches")
    .addItem("Dispatch Weekly Summary Email", "sendWeeklyBudgetSummary")
    .addItem("➕ Add CD Investment", "promptAddCD")
    .addSeparator()
    .addItem("➕ Create Travel Itinerary & Agenda", "promptTravelItinerary")
    .addItem("📈 Build CD Portfolio Growth Tracker", "setupCDTracker")
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
// 1.5 DYNAMIC ENUM & DATA VALIDATION AUTO-EXPANSION ENGINE
// ===========================================================================

/**
 * Refreshes dropdown data validations dynamically across all sheets to prevent
 * "Invalid: Input must fall within specified range" errors.
 */
function refreshEnumValidations(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  const dbSheet = ss.getSheetByName("Transactions_DB");
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  if (!enumSheet) return;

  const lastCatRow = Math.max(3, enumSheet.getLastRow());
  const catRange = enumSheet.getRange("A3:A" + lastCatRow);
  const catValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(catRange)
    .setAllowInvalid(false)
    .build();

  const lastAcctRow = Math.max(3, enumSheet.getLastRow());
  const acctRange = enumSheet.getRange("B3:B" + lastAcctRow);
  const acctValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(acctRange)
    .setAllowInvalid(false)
    .build();

  if (dbSheet) {
    dbSheet.getRange("H2:H5000").setDataValidation(catValidation);
    dbSheet.getRange("I2:I5000").setDataValidation(acctValidation);

    const flowValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["INFLOW", "OUTFLOW", "TRANSFER"])
      .setAllowInvalid(false)
      .build();
    dbSheet.getRange("G2:G5000").setDataValidation(flowValidation);

    const channelValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["WEB_FORM", "CHASE_CSV", "GMAIL_ALERT", "GOOGLE_VOICE_SMS", "PLAID_API", "MANUAL", "SAVINGS_SWEEP"])
      .setAllowInvalid(false)
      .build();
    dbSheet.getRange("J2:J5000").setDataValidation(channelValidation);

    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["SETTLED", "PENDING", "MANUAL_REVIEW"])
      .setAllowInvalid(false)
      .build();
    dbSheet.getRange("K2:K5000").setDataValidation(statusValidation);
  }

  if (rulesSheet) {
    rulesSheet.getRange("C3:C5000").setDataValidation(catValidation);
  }
  console.log("[Validation Engine] Successfully refreshed dynamic dropdown validations across all sheets.");
}

/**
 * Checks if a category exists in Enum!A:A. If not, automatically appends it
 * and refreshes dropdown data validations across the workbook.
 */
function ensureCategoryInEnum(categoryName, ss) {
  if (!categoryName) return;
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  if (!enumSheet) return;

  const catName = categoryName.toString().trim();
  if (!catName) return;

  const lastRow = enumSheet.getLastRow();
  let existingCats = [];
  if (lastRow >= 3) {
    existingCats = enumSheet.getRange(3, 1, lastRow - 2, 1).getValues().flat().map(c => c.toString().trim().toLowerCase());
  }

  if (existingCats.indexOf(catName.toLowerCase()) === -1) {
    enumSheet.appendRow([catName, ""]);
    console.log(`[Enum Engine] Automatically registered new category in Enum!A:A: "${catName}"`);

    if (budgetDB) {
      const today = new Date();
      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
      const currentMonth = months[today.getMonth()];
      const currentYear = today.getFullYear();
      budgetDB.appendRow([currentYear, currentMonth, catName, 0.00]);
    }
    refreshEnumValidations(ss);
  }
}

/**
 * Checks if an account source exists in Enum!B:B. If not, automatically appends it
 * and refreshes dropdown data validations across the workbook.
 */
function ensureAccountInEnum(accountName, ss) {
  if (!accountName) return;
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  if (!enumSheet) return;

  const acctName = accountName.toString().trim();
  if (!acctName) return;

  const lastRow = enumSheet.getLastRow();
  let existingAccts = [];
  if (lastRow >= 3) {
    existingAccts = enumSheet.getRange(3, 2, lastRow - 2, 1).getValues().flat().map(a => a.toString().trim().toLowerCase());
  }

  if (existingAccts.indexOf(acctName.toLowerCase()) === -1) {
    const nextRow = Math.max(3, lastRow + 1);
    enumSheet.getRange(nextRow, 2).setValue(acctName);
    console.log(`[Enum Engine] Automatically registered new account source in Enum!B:B: "${acctName}"`);
    refreshEnumValidations(ss);
  }
}

/**
 * Smart Multi-Account Detector for CSV files, Bank Emails, and Google Voice Messages.
 */
function detectAccountSource(filename, csvHeader, csvRow) {
  const fileLower = (filename || "").toLowerCase();
  
  if (fileLower.indexOf("1234") !== -1 || fileLower.indexOf("primary") !== -1) return "Chase Checking 1234";
  if (fileLower.indexOf("5678") !== -1 || fileLower.indexOf("joint") !== -1) return "Chase Checking 5678";
  if (fileLower.indexOf("9012") !== -1 || fileLower.indexOf("sapphire") !== -1) return "Chase Sapphire 9012";
  if (fileLower.indexOf("3456") !== -1 || fileLower.indexOf("freedom") !== -1) return "Chase Freedom 3456";
  if (fileLower.indexOf("td") !== -1) return "TD Bank";

  if (csvHeader && csvRow) {
    const colAccount = csvHeader.indexOf("Account") !== -1 ? csvHeader.indexOf("Account") : csvHeader.indexOf("Account Number");
    if (colAccount !== -1 && csvRow[colAccount]) {
      const acctStr = csvRow[colAccount].toString().trim();
      if (acctStr.endsWith("1234")) return "Chase Checking 1234";
      if (acctStr.endsWith("5678")) return "Chase Checking 5678";
      if (acctStr.endsWith("9012")) return "Chase Sapphire 9012";
      if (acctStr.endsWith("3456")) return "Chase Freedom 3456";
    }
  }

  return "Chase Checking 1234"; // Default fallback
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
 */
function createWholeBook(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("No spreadsheet found. Ensure a spreadsheet is created or active.");
  }
  
  console.log(`[Whole Book Compiler] Constructing sheet structures for Spreadsheet: ${ss.getId()}`);
  
  const sheetNames = ["About", "Enum", "Transactions_DB", "Budget_History_DB", "The Monthly Dashboard Tab", "Control", "Category_Rules_DB", "Savings_Goals"];
  const sheets = {};
  
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
  
  sheets["About"].activate(); ss.moveActiveSheet(1);
  sheets["Enum"].activate(); ss.moveActiveSheet(2);
  sheets["Transactions_DB"].activate(); ss.moveActiveSheet(3);
  sheets["Budget_History_DB"].activate(); ss.moveActiveSheet(4);
  sheets["The Monthly Dashboard Tab"].activate(); ss.moveActiveSheet(5);
  sheets["Control"].activate(); ss.moveActiveSheet(6);
  sheets["Category_Rules_DB"].activate(); ss.moveActiveSheet(7);
  sheets["Savings_Goals"].activate(); ss.moveActiveSheet(8);
  
  
  
  
  
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
  about.getRange("A17").setValue("5. Check the 'Control' sheet's master indicator to ensure zero mathematical anomalies.");
  about.autoResizeColumns(1, 4);
  
  // -------------------------------------------------------------------------
  // ENUM TAB (EXPANDED MASTER LISTS)
  // -------------------------------------------------------------------------
  const enumSheet = sheets["Enum"];
  formatHeader(enumSheet, "A1:B1", "SYSTEM ENUMS / CATEGORIES & ACCOUNTS");
  enumSheet.getRange("A2:B2").setValues([["Budget Categories", "Account Sources"]]).setFontWeight("bold");
  
  const categories = [
    ["Food & Drink"], ["Groceries"], ["Bills & Utilities"], ["Home"], 
    ["Transportation"], ["Entertainment"], ["Personal"], ["Travel"], 
    ["Transfer"], ["Income"], ["Checks & Banking"], ["Savings & Investments"],
    ["Health & Medical"], ["Education"], ["Miscellaneous"]
  ];
  const accounts = [
    ["Cash"], ["Chase Checking 1234"], ["Chase Checking 5678"], ["Chase Sapphire 9012"], ["TD Bank"], ["Credit Card"], ["Savings"]
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
  
  // Initial dummy row
  db.getRange("A2:K2").setValues([["INIT_SETUP_DUMMY", new Date(), new Date(), "INITIAL SYSTEM SETUP", "Initial Setup", 0.01, "INFLOW", "Miscellaneous", "Cash", "MANUAL", "SETTLED"]]);
  db.getRange("F2:F1000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  db.getRange("C2:C1000").setNumberFormat("yyyy-mm-dd");
  db.autoResizeColumns(1, 11);
  
  // -------------------------------------------------------------------------
  // BUDGET_HISTORY_DB TAB
  // -------------------------------------------------------------------------
  const budgetDB = sheets["Budget_History_DB"];
  formatHeader(budgetDB, "A1:D1", "HISTORICAL BUDGET RECORD DATABASE");
  const budgetHeaders = [["Year", "Month", "Category", "Budget_Amount"]];
  budgetDB.getRange("A2:D2").setValues(budgetHeaders).setFontWeight("bold");
  formatSubHeader(budgetDB, "A2:D2");
  budgetDB.setRowHeight(2, 25);
  
  const initialAllocations = [];
  categories.forEach(cat => {
    initialAllocations.push([2026, "August", cat[0], 0.00]);
  });
  budgetDB.getRange(3, 1, initialAllocations.length, 4).setValues(initialAllocations);
  budgetDB.getRange("D3:D1000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  budgetDB.autoResizeColumns(1, 4);
  
  // -------------------------------------------------------------------------
  // CATEGORY RULES DATABASE TAB (Category_Rules_DB)
  // -------------------------------------------------------------------------
  const rulesDb = sheets["Category_Rules_DB"];
  formatHeader(rulesDb, "A1:C1", "AI RULES-BASED CATEGORIZATION LOOKUP DATABASE");
  rulesDb.getRange("A2:C2").setValues([["Merchant Keyword Rule", "Clean Merchant Name Target", "Budget Category Target"]]).setFontWeight("bold");
  formatSubHeader(rulesDb, "A2:C2");
  rulesDb.setRowHeight(2, 25);
  
  const startingRules = [
    ["starbucks", "Starbucks", "Food & Drink"],
    ["mcdonald", "McDonald's", "Food & Drink"],
    ["kroger", "Kroger", "Groceries"],
    ["walmart", "Walmart", "Groceries"],
    ["trader joe", "Trader Joe's", "Groceries"],
    ["landaus supermarket", "Landaus Supermarket", "Groceries"],
    ["home depot", "The Home Depot", "Home"],
    ["comcast", "Comcast", "Bills & Utilities"],
    ["verizon", "Verizon", "Bills & Utilities"],
    ["uber", "Uber", "Transportation"],
    ["chevron", "Chevron", "Transportation"],
    ["exxon", "Exxon", "Transportation"],
    ["netflix", "Netflix", "Entertainment"],
    ["zelle", "Zelle Transfer", "Transfer"],
    ["wire", "Wire Transfer", "Transfer"],
    ["check", "Check Payment", "Checks & Banking"]
  ];
  rulesDb.getRange(3, 1, startingRules.length, 3).setValues(startingRules);
  rulesDb.getRange("A3:C1000").setFontColor("#002060");
  rulesDb.autoResizeColumns(1, 3);
  
  // Apply dynamic validations across sheets
  refreshEnumValidations(ss);

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
  
  for (let g = 0; g < defaultGoals.length; g++) {
    const rIdx = 3 + g;
    goalsSheet.getRange(rIdx, 6).setFormula(`=IF(B${rIdx}>0, SPARKLINE(C${rIdx}, {"charttype","bar";"max", B${rIdx};"color1", "#2e7d32"}), "")`);
  }
  
  goalsSheet.getRange("B3:C100").setNumberFormat("$#,##0.00");
  goalsSheet.getRange("D3:D100").setNumberFormat("yyyy-mm-dd");
  goalsSheet.getRange("E3:E100").setNumberFormat("0.0%");
  goalsSheet.getRange("A3:E100").setFontColor("#002060");
  
  goalsSheet.getRange("H1:I1").setValues([["Projection Parameter", "User Allocation Value"]]).setFontWeight("bold");
  formatSubHeader(goalsSheet, "H1:I1");
  goalsSheet.getRange("H2:I2").setValues([["Assumed Annual Growth Rate", 0.06]]);
  goalsSheet.getRange("H3:I3").setValues([["Annual Goals Savings Contribution", 6000.00]]);
  goalsSheet.getRange("I2").setNumberFormat("0.0%").setFontColor("#002060");
  goalsSheet.getRange("I3").setNumberFormat("$#,##0.00").setFontColor("#002060");
  
  goalsSheet.getRange("H5:L5").setValues([["Year", "Starting Balance", "Annual Savings Contribution", "Compound Interest", "Projected Ending Balance"]]).setFontWeight("bold");
  formatSubHeader(goalsSheet, "H5:L5");
  
  goalsSheet.getRange("H6").setValue(2026);
  goalsSheet.getRange("I6").setFormula("=SUM(C3:C5)");
  goalsSheet.getRange("J6").setFormula("=$I$3");
  goalsSheet.getRange("K6").setFormula("=(I6+J6)*$I$2");
  goalsSheet.getRange("L6").setFormula("=SUM(I6:K6)");
  
  for (let y = 1; y < 10; y++) {
    const rowIdx = 6 + y;
    goalsSheet.getRange(rowIdx, 8).setValue(2026 + y);
    goalsSheet.getRange(rowIdx, 9).setFormula(`=L${rowIdx - 1}`);
    goalsSheet.getRange(rowIdx, 10).setFormula("=$I$3");
    goalsSheet.getRange(rowIdx, 11).setFormula(`=(I${rowIdx}+J${rowIdx})*$I$2`);
    goalsSheet.getRange(rowIdx, 12).setFormula(`=SUM(I${rowIdx}:K${rowIdx})`);
  }
  goalsSheet.getRange("I6:L15").setNumberFormat("$#,##0.00");
  goalsSheet.getRange("H6:H15").setNumberFormat("0000").setHorizontalAlignment("center");
  
  try {
    const projectionChart = goalsSheet.newChart()
      .setChartType(Charts.ChartType.AREA)
      .addRange(goalsSheet.getRange("H5:H15"))
      .addRange(goalsSheet.getRange("L5:L15"))
      .setPosition(1, 14, 0, 0)
      .setOption('title', 'Multi-Year Long-Term Savings Projection (10-Year Growth)')
      .setOption('legend', { position: 'none' })
      .setOption('width', 520)
      .setOption('height', 320)
      .setOption('colors', ['#2e7d32'])
      .setOption('hAxis', { title: 'Year', format: '0000' })
      .setOption('vAxis', { title: 'Projected Net Worth ($)' })
      .build();
    goalsSheet.insertChart(projectionChart);
    console.log("[Whole Book Compiler] Programmable Area projection chart successfully generated.");
  } catch (err) {
    console.error("[Whole Book Compiler Error] Projection chart setup failed: " + err.toString());
  }
  goalsSheet.autoResizeColumns(1, 12);
  
  // -------------------------------------------------------------------------
  // THE MONTHLY DASHBOARD TAB
  // -------------------------------------------------------------------------
  const dash = sheets["The Monthly Dashboard Tab"];
  formatHeader(dash, "A1:G1", "INTERACTIVE MONTHLY FINANCIAL DASHBOARD");
  dash.getRange("A2").setFormula('=IF(Control!$A$2="✅ SYSTEM OK", "✅ SYSTEM OK", "❌ ERROR DETECTED")');
  dash.getRange("A2").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f2f2f2");
  
  dash.getRange("C3").setValue("Select Year:").setFontWeight("bold");
  const yearValidation = SpreadsheetApp.newDataValidation().requireValueInList(["2024", "2025", "2026", "2027", "2028"]).setAllowInvalid(false).build();
  dash.getRange("C4").setDataValidation(yearValidation).setValue("2026").setFontColor("#002060").setFontWeight("bold");
  
  dash.getRange("D3").setValue("Select Month:").setFontWeight("bold");
  const monthValidation = SpreadsheetApp.newDataValidation().requireValueInList([
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "All"
  ]).setAllowInvalid(false).build();
  dash.getRange("D4").setDataValidation(monthValidation).setValue("August").setFontColor("#002060").setFontWeight("bold");
  
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
  
  dash.getRange("F3").setValue("Uncategorized Alerts:").setFontWeight("bold");
  dash.getRange("F4").setFormula('=COUNTIFS(Transactions_DB!C:C, "<>", Transactions_DB!H:H, "")').setFontColor("#c00000").setFontWeight("bold").setHorizontalAlignment("center");
  
  dash.getRange("G3").setValue("Net Goal progress:").setFontWeight("bold");
  dash.getRange("G4").setFormula('=IFERROR(C7/1500, 0)').setNumberFormat("0.0%").setFontColor("#385723").setFontWeight("bold").setHorizontalAlignment("center");
  
  dash.getRange("A9:E9").setValues([["Category", "Planned Budget", "Actual Spending", "Difference", "Budget Progress"]]);
  formatSubHeader(dash, "A9:E9");
  
  const enumCategoriesCount = categories.length;
  for (let r = 0; r < enumCategoriesCount; r++) {
    const rowIdx = 10 + r;
    dash.getRange(rowIdx, 1).setFormula(`=Enum!A${3 + r}`).setFontColor("#385723").setFontWeight("bold");
    dash.getRange(rowIdx, 2).setValue(0).setFontColor("#002060");
    dash.getRange(rowIdx, 3).setFormula(
      `=IF($D$4="All", SUMIFS(Transactions_DB!F:F, Transactions_DB!H:H, A${rowIdx}, Transactions_DB!G:G, "OUTFLOW", Transactions_DB!C:C, ">="&DATE($C$4, 1, 1), Transactions_DB!C:C, "<="&DATE($C$4, 12, 31)), SUMIFS(Transactions_DB!F:F, Transactions_DB!H:H, A${rowIdx}, Transactions_DB!G:G, "OUTFLOW", Transactions_DB!C:C, ">="&DATEVALUE($D$4&" 1, "&$C$4), Transactions_DB!C:C, "<="&EOMONTH(DATEVALUE($D$4&" 1, "&$C$4), 0)))`
    );
    dash.getRange(rowIdx, 4).setFormula(`=B${rowIdx}-C${rowIdx}`);
    dash.getRange(rowIdx, 5).setFormula(
      `=IF(B${rowIdx}>0, SPARKLINE(C${rowIdx}, {"charttype","bar";"max", B${rowIdx};"color1", IF(C${rowIdx}>B${rowIdx}, "#db4437", "#4285f4")}), "")`
    );
  }
  
  dash.getRange(10, 2, enumCategoriesCount, 3).setNumberFormat("$#,##0.00");
  const totalRowIdx = 10 + enumCategoriesCount;
  dash.getRange(totalRowIdx, 1, 1, 4).setValues([[
    "Total Tracked Outflows", `=SUM(B10:B${totalRowIdx-1})`, `=SUM(C10:C${totalRowIdx-1})`, `=SUM(D10:D${totalRowIdx-1})`
  ]]).setFontWeight("bold");
  dash.getRange(totalRowIdx, 2, 1, 3).setNumberFormat("$#,##0.00");
  dash.getRange(totalRowIdx, 1, 1, 5).setBackground("#f2f2f2");
  
  const ledgerHeaderIdx = totalRowIdx + 2;
  dash.getRange(ledgerHeaderIdx, 1, 1, 7).setValues([["Date", "Merchant Name", "Amount", "Flow Direction", "Category", "Account Source", "Channel"]]);
  formatSubHeader(dash, `A${ledgerHeaderIdx}:G${ledgerHeaderIdx}`);
  
  const ledgerFormulaIdx = ledgerHeaderIdx + 1;
  dash.getRange(ledgerFormulaIdx, 1).setFormula(
    `=IF($D$4="All", FILTER({Transactions_DB!C2:C, Transactions_DB!D2:D, Transactions_DB!F2:F, Transactions_DB!G2:G, Transactions_DB!H2:H, Transactions_DB!I2:I, Transactions_DB!J2:J}, YEAR(Transactions_DB!C2:C) = $C$4), FILTER({Transactions_DB!C2:C, Transactions_DB!D2:D, Transactions_DB!F2:F, Transactions_DB!G2:G, Transactions_DB!H2:H, Transactions_DB!I2:I, Transactions_DB!J2:J}, YEAR(Transactions_DB!C2:C) = $C$4, MONTH(Transactions_DB!C2:C) = MONTH($D$4&1)))`
  );
  
  dash.getRange(ledgerFormulaIdx, 1, 100, 1).setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
  dash.getRange(ledgerFormulaIdx, 3, 100, 1).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  dash.getRange(ledgerFormulaIdx, 4, 100, 1).setHorizontalAlignment("center");
  dash.autoResizeColumns(1, 7);
  dash.setFrozenRows(5);
  
  syncDashboardBudgets(dash);
  
  try {
    const existingCharts = dash.getCharts();
    existingCharts.forEach(c => dash.removeChart(c));
    
    const budgetVsActualChart = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(9, 1, enumCategoriesCount + 1, 3))
      .setPosition(9, 9, 0, 0)
      .setOption('title', 'Monthly Planned Budget vs. Actual Spending')
      .setOption('legend', { position: 'top' })
      .setOption('width', 500)
      .setOption('height', 300)
      .setOption('colors', ['#4285f4', '#db4437'])
      .setOption('hAxis', { title: 'Categories', textStyle: { fontSize: 10 } })
      .setOption('vAxis', { title: 'Amount ($)' })
      .build();
    dash.insertChart(budgetVsActualChart);
    
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
  // CONTROL TAB
  // -------------------------------------------------------------------------
  const control = sheets["Control"];
  formatHeader(control, "A1:D1", "WORKBOOK CONTROL CHECKS & BALANCES");
  control.getRange("A3:D3").setValues([["Check Name", "Formula / Logic Description", "Current Value", "Status"]]);
  formatSubHeader(control, "A3:D3");
  
  control.getRange("A4").setValue("Sign Normalization Check");
  control.getRange("B4").setValue("Counts transactions with negative amounts in Column F (Should be 0)");
  control.getRange("C4").setFormula('=COUNTIF(Transactions_DB!F:F, "<0")');
  control.getRange("D4").setFormula('=IF(C4=0, "OK", "ERROR: Negative values found")');
  
  control.getRange("A5").setValue("Blank Category Check");
  control.getRange("B5").setValue("Counts transactions that have dates but no category assigned (Should be 0)");
  control.getRange("C5").setFormula('=COUNTIFS(Transactions_DB!C2:C, "<>", Transactions_DB!H2:H, "")');
  control.getRange("D5").setFormula('=IF(C5=0, "OK", "ERROR: Uncategorized transactions")');
  
  control.getRange("A6").setValue("Capital Flow Classification Check");
  control.getRange("B6").setValue("Counts rows missing standardized Capital_Flow definitions in Column G");
  control.getRange("C6").setFormula('=COUNTIFS(Transactions_DB!C2:C, "<>", Transactions_DB!G2:G, "<>INFLOW", Transactions_DB!G2:G, "<>OUTFLOW", Transactions_DB!G2:G, "<>TRANSFER")');
  control.getRange("D6").setFormula('=IF(C6=0, "OK", "ERROR: Invalid flow class")');
  
  control.getRange("A2").setFormula('=IF(COUNTIF(D4:D6, "ERROR*")=0, "✅ SYSTEM OK", "❌ ERROR DETECTED")');
  control.getRange("A2").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  control.autoResizeColumns(1, 4);
  
  applyControlTabFormatting(ss);
  console.log("[Whole Book Compiler] Finished compiler run. Workbook structure is fully operational.");
}

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
// 3. TOOLBAR QUICK-ADD PROMPTS
// ===========================================================================



function quickAddCategoryPrompt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.prompt("➕ Add New Budget Category", "Enter new budget category name:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() == ui.Button.OK) {
    const catName = resp.getResponseText().trim();
    if (catName) {
      ensureCategoryInEnum(catName, ss);
      ui.alert("Category Registered! ✅", `Category "${catName}" added to Enum!A:A and History database.`, ui.ButtonSet.OK);
    }
  }
}

function quickAddAccountPrompt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.prompt("➕ Add New Account Source", "Enter new account name (e.g., Chase Checking 5678):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() == ui.Button.OK) {
    const acctName = resp.getResponseText().trim();
    if (acctName) {
      ensureAccountInEnum(acctName, ss);
      ui.alert("Account Source Registered! ✅", `Account "${acctName}" added to Enum!B:B and dropdown rules.`, ui.ButtonSet.OK);
    }
  }
}



// ===========================================================================
// 4. BACKGROUND SYNCHRONIZER & SIDEBARS
// ===========================================================================

function syncDashboardBudgets(dash) {
  const ss = dash.getParent();
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  const enumSheet = ss.getSheetByName("Enum");
  if (!budgetDB || !enumSheet) return;

  const year = parseInt(dash.getRange("C4").getValue(), 10);
  const month = dash.getRange("D4").getValue().toString().trim();

  const lastEnumRow = enumSheet.getLastRow();
  if (lastEnumRow < 3) return;
  const categories = enumSheet.getRange(3, 1, lastEnumRow - 2, 1).getValues().flat().filter(String);
  const numCategories = categories.length;

  const dbLastRow = budgetDB.getLastRow();
  let dbData = [];
  if (dbLastRow >= 3) {
    dbData = budgetDB.getRange(3, 1, dbLastRow - 2, 4).getValues();
  }

  const updatedBudgets = [];
  categories.forEach(cat => {
    const match = dbData.find(row => 
      parseInt(row[0], 10) === year && 
      row[1].toString().trim() === month && 
      row[2].toString().trim() === cat
    );
    updatedBudgets.push([match ? match[3] : 0.00]);
  });

  dash.getRange(10, 2, numCategories, 1).setValues(updatedBudgets);
  console.log("[Budget Sync Engine] Synchronized " + numCategories + " category budgets for Period: " + year + " " + month);
}

function getDashboardPeriod() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  if (!dash) return { year: "2026", month: "August" };
  return {
    year: dash.getRange("C4").getValue().toString().trim() || "2026",
    month: dash.getRange("D4").getValue().toString().trim() || "August"
  };
}

function showBudgetSidebar() {
  const html = HtmlService.createTemplateFromFile("SidebarView");
  const htmlOutput = html.evaluate()
    .setTitle("Historical Budget Allocator")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  console.log("[UI Service] Rendered Budget History Allocation Sidebar.");
}

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
  
  return categories.map(cat => {
    const match = dbData.find(row => 
      parseInt(row[0], 10) === parseInt(year, 10) && 
      row[1].toString().trim() === month.toString().trim() && 
      row[2].toString().trim() === cat
    );
    return {
      category: cat,
      amount: match ? parseFloat(match[3]) : 0.00
    };
  });
}

function saveCategoryBudgets(year, month, budgetData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetDB = ss.getSheetByName("Budget_History_DB");
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const dbLastRow = budgetDB.getLastRow();
    let dbData = [];
    if (dbLastRow >= 3) {
      dbData = budgetDB.getRange(3, 1, dbLastRow - 2, 4).getValues();
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
// 5. INGESTION CHANNELS & AUTOMATED PARSERS
// ===========================================================================

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
    
    const rawMerchant = rowData.Merchant || "Web Form Entry";
    const cleanMerchant = autoCleanMerchantName(rawMerchant);
    const category = rowData.Category || autoCategorizeMerchant(cleanMerchant);
    const accountSource = rowData.Source || "Cash";

    ensureCategoryInEnum(category, ss);
    ensureAccountInEnum(accountSource, ss);

    sheet.appendRow([
      transactionId,
      timestamp,
      dateValue,
      rawMerchant,
      cleanMerchant,
      Math.abs(parseFloat(rowData.Amount)),
      rowData.Flow || "OUTFLOW",
      category,
      accountSource,
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

function pullGmailTransactionAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return;
  
  const lastRow = dbSheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }
  
  // Channel 1: Traditional Chase Transaction Email Alerts
  const chaseQuery = 'from:Chase "transaction alert" is:unread';
  const chaseThreads = GmailApp.search(chaseQuery, 0, 10);
  let newRowsCount = 0;
  
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
      
      let merchantRaw = "CHASE TRANSACTION ALERT";
      const merchMatch = body.match(/at\s+([^\n\r,.]+?)\s+on/i) || body.match(/charge\s+of\s+.*?\s+at\s+([^\n\r,.]+)/i);
      if (merchMatch) merchantRaw = merchMatch[1].trim();
      
      const acctMatch = body.match(/(?:ending in|account|card)\s*#?\s*(\d{4})/i);
      let targetAccount = "Chase Checking 1234";
      if (acctMatch && acctMatch[1]) {
        targetAccount = "Chase Card " + acctMatch[1];
      }

      const cleanMerchant = autoCleanMerchantName(merchantRaw);
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant);

      ensureCategoryInEnum(suggestedCategory, ss);
      ensureAccountInEnum(targetAccount, ss);

      dbSheet.appendRow([
        msgId, new Date(), date, merchantRaw, cleanMerchant, amount, "OUTFLOW",
        suggestedCategory, targetAccount, "GMAIL_ALERT", "SETTLED"
      ]);
      msg.markRead();
      existingIDs.push(msgId);
      newRowsCount++;
    }
  }

  // Channel 2: Google Voice Forwarded Transaction SMS Texts
  const voiceQuery = 'from:voice-noreply@google.com "SMS" is:unread';
  const voiceThreads = GmailApp.search(voiceQuery, 0, 10);
  
  for (let i = 0; i < voiceThreads.length; i++) {
    const messages = voiceThreads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const msgId = "GV-SMS-" + msg.getId();
      if (existingIDs.indexOf(msgId) !== -1) continue;

      const body = msg.getPlainBody();
      const amountMatch = body.match(/\$(\d+(?:\.\d{2})?)/);
      if (!amountMatch) continue;

      const amount = parseFloat(amountMatch[1]);
      const date = msg.getDate();

      let flow = "OUTFLOW";
      if (/paid you|sent you|received|deposit|credit/i.test(body)) {
        flow = "INFLOW";
      }

      let rawMerchant = "Google Voice Transaction";
      const merchMatch = body.match(/at\s+([^\n\r,.]+)/i) || body.match(/from\s+([^\n\r,.]+)/i) || body.match(/to\s+([^\n\r,.]+)/i);
      if (merchMatch) rawMerchant = merchMatch[1].trim();

      const acctMatch = body.match(/(?:ending in|account|card)\s*#?\s*(\d{4})/i);
      let targetAccount = "Chase Checking 1234";
      if (acctMatch && acctMatch[1]) {
        targetAccount = "Chase Card " + acctMatch[1];
      }

      const cleanMerchant = autoCleanMerchantName(rawMerchant);
      const category = autoCategorizeMerchant(cleanMerchant);

      ensureCategoryInEnum(category, ss);
      ensureAccountInEnum(targetAccount, ss);

      dbSheet.appendRow([
        msgId, new Date(), date, rawMerchant, cleanMerchant, amount, flow,
        category, targetAccount, "GOOGLE_VOICE_SMS", "SETTLED"
      ]);
      msg.markRead();
      existingIDs.push(msgId);
      newRowsCount++;
    }
  }
  console.log(`[Alert Parser] Processed emails and SMS. Imported ${newRowsCount} transactions.`);
}

function promptForChaseCSV() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Import Chase CSV",
    "Please enter the Google Drive File ID of your uploaded Chase CSV:",
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

function importChaseCSV(fileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) throw new Error("Transactions_DB tab not found.");
  
  let file;
  if (fileId) {
    file = DriveApp.getFileById(fileId);
  } else {
    const files = DriveApp.getFilesByName("Chase.csv");
    if (files.hasNext()) file = files.next();
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
  
  const detectedAccount = detectAccountSource(file.getName(), csvHeaders, csvData[1]);
  ensureAccountInEnum(detectedAccount, ss);

  let newRowsCount = 0;
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
    
    const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
    ensureCategoryInEnum(suggestedCategory, ss);

    dbSheet.appendRow([
      sigId, new Date(), new Date(row[colDate]), merchantRaw, cleanMerchant,
      Math.abs(rawAmount), rawAmount < 0 ? "OUTFLOW" : "INFLOW", suggestedCategory,
      detectedAccount, "CHASE_CSV", "SETTLED"
    ]);
    newRowsCount++;
  }
  console.log(`[Chase Ingestion] Import finished. Ingested ${newRowsCount} rows.`);
  return "Successfully imported " + newRowsCount + " transaction rows.";
}

function importAllPendingCSVs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) {
    ui.alert("Configuration Error ⚠️", "Transactions_DB sheet not found.", ui.ButtonSet.OK);
    return;
  }

  let uploadFolder;
  const folders = DriveApp.getFoldersByName("Chase_CSV_Uploads");
  if (folders.hasNext()) {
    uploadFolder = folders.next();
  } else {
    uploadFolder = DriveApp.createFolder("Chase_CSV_Uploads");
    uploadFolder.createFolder("Archive");
    ui.alert(
      "Folder Setup Created! 📂",
      "Created 'Chase_CSV_Uploads' folder in Google Drive.\n\n" +
      "1. Drag and drop bank CSV files into 'Chase_CSV_Uploads'.\n" +
      "2. Click 'Batch Import CSV Folder' again to ingest all statements at once!",
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
    if (file.getMimeType() !== "text/csv" && !file.getName().toLowerCase().endsWith(".csv")) continue;

    try {
      const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      if (csvData.length < 2) continue;

      const csvHeaders = csvData[0].map(h => h.trim());
      const colDate = csvHeaders.indexOf("Transaction Date") !== -1 ? csvHeaders.indexOf("Transaction Date") : csvHeaders.indexOf("Posting Date");
      const colMerchant = csvHeaders.indexOf("Description");
      const colAmount = csvHeaders.indexOf("Amount");

      if (colDate === -1 || colMerchant === -1 || colAmount === -1) continue;

      const accountSource = detectAccountSource(file.getName(), csvHeaders, csvData[1]);
      ensureAccountInEnum(accountSource, ss);

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

        const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
        ensureCategoryInEnum(suggestedCategory, ss);

        dbSheet.appendRow([
          sigId, new Date(), new Date(row[colDate]), merchantRaw, cleanMerchant,
          Math.abs(rawAmount), rawAmount < 0 ? "OUTFLOW" : "INFLOW", suggestedCategory,
          accountSource, "CHASE_CSV", "SETTLED"
        ]);

        existingIDs.push(sigId);
        fileRowsAdded++;
      }

      file.moveTo(archiveFolder);
      filesProcessed++;
      totalNewRows += fileRowsAdded;
    } catch (err) {
      console.error("Failure processing file " + file.getName() + ": " + err.toString());
    }
  }

  if (filesProcessed === 0) {
    ui.alert("Process Status ℹ️", "No pending .csv statement files found inside 'Chase_CSV_Uploads'.", ui.ButtonSet.OK);
  } else {
    ui.alert("Batch Import Complete! 🎉", `Processed ${filesProcessed} CSV files.\nAdded ${totalNewRows} new transaction rows.`, ui.ButtonSet.OK);
  }
}

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
// 6. MULTI-PATTERN ADVANCED MERCHANT CLEANSING & AUTO-LEARNING CATEGORIZER
// ===========================================================================

function autoCleanMerchantName(merchantRaw) {
  if (!merchantRaw) return "Unknown Merchant";
  const raw = merchantRaw.toString().trim();
  const merchantLower = raw.toLowerCase();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  
  if (rulesSheet && rulesSheet.getLastRow() >= 3) {
    const rulesData = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 2).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      const keyword = rulesData[i][0].toString().toLowerCase().trim();
      const cleanName = rulesData[i][1].toString().trim();
      if (keyword !== "" && merchantLower.indexOf(keyword) !== -1 && cleanName !== "") {
        return cleanName;
      }
    }
  }

  // Specialized Banking & Transfer Pattern Recognition
  if (/zelle\s+payment\s+(from|to)\s+(.*)/i.test(raw)) {
    const m = raw.match(/zelle\s+payment\s+(from|to)\s+(.*)/i);
    if (m && m[2]) {
      let name = m[2].replace(/\s+bacz.*$/i, '').replace(/[\/\#\-].*$/, '').trim();
      return "Zelle: " + name.replace(/\b\w/g, c => c.toUpperCase());
    }
  }
  if (/zelle\s+from\s+(.*)/i.test(raw)) {
    const m = raw.match(/zelle\s+from\s+(.*)/i);
    if (m && m[1]) return "Zelle: " + m[1].replace(/[\/\#\-].*$/, '').trim().replace(/\b\w/g, c => c.toUpperCase());
  }
  if (/check\s*#?\s*(\d+)/i.test(raw) || /chk\s*#?\s*(\d+)/i.test(raw)) {
    const m = raw.match(/(check|chk)\s*#?\s*(\d+)/i);
    if (m && m[2]) return "Check #" + m[2];
    return "Check Payment";
  }
  if (/wire\s*(trans|transfer|incoming|out)?/i.test(raw)) {
    const m = raw.match(/(ref\#?\s*\d+|ref\s*\w+)/i);
    if (m) return "Wire Transfer (" + m[1] + ")";
    return "Wire Transfer";
  }
  if (/irs\s*usataxpymt|nys\s*dtf|ny\s*state\s*nysttaxrfd/i.test(raw)) {
    if (/nys|ny\s*state/i.test(raw)) return "NYS Tax Payment";
    return "IRS Tax Payment";
  }
  if (/atm\s*withdrawal|atm\s*check\s*deposit/i.test(raw)) {
    if (/deposit/i.test(raw)) return "ATM Check Deposit";
    return "ATM Withdrawal";
  }

  // Strip Banking Noise Prefixes
  let cleaned = raw
    .replace(/^pos\s+(debit|purchase|withdrawal|card\s+purchase)\s+/i, '')
    .replace(/^orig\s+co\s+name:\s*/i, '')
    .replace(/^ppd\s+id:\s*/i, '')
    .replace(/sec:ppd\s*orig\s*id:.*$/i, '')
    .replace(/^ach\s+co_entry\s+descr:\s*/i, '')
    .replace(/webpayment\s+web\s+id:.*$/i, '')
    .replace(/remote\s+online\s+deposit.*$/i, 'Remote Online Deposit')
    .replace(/deposit\s+id\s+number.*$/i, 'Deposit')
    .replace(/^(sq\s*\*|tst\s\*)/i, '');

  // Strip Location Noise, Store Numbers, Hashes, Slashes, Dashes
  cleaned = cleaned
    .replace(/#.*$/g, '')
    .replace(/\s+--\s+.*$/g, '')
    .replace(/\s*[\/\\]\s*.*$/g, '')
    .replace(/\s*-\s*(ny|sc|ca|nj|fl|tx|monroe|monsey|spring\s+valley|columbia|brooklyn).*$/i, '')
    .replace(/\s+(monroe|monsey|spring\s+valley|columbia|brooklyn|new\s+york|ny|sc|ca|nj|fl)\s*$/i, '')
    .replace(/[0-9]{3,}/g, '')
    .replace(/[-_\/*#()]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!cleaned || cleaned.length < 2) {
    cleaned = raw.replace(/[0-9]+/g, "").replace(/[-_\/*#()]+/g, " ").replace(/\s+/g, " ").trim();
  }

  return cleaned.replace(/\b\w/g, c => c.toUpperCase());
}

function autoCategorizeMerchant(merchantRaw) {
  if (!merchantRaw) return "Miscellaneous";
  const merchant = merchantRaw.toString().toLowerCase().trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  
  if (rulesSheet && rulesSheet.getLastRow() >= 3) {
    const rulesData = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 3).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      const keyword = rulesData[i][0].toString().toLowerCase().trim();
      const categoryVal = rulesData[i][2].toString().trim();
      if (keyword !== "" && merchant.indexOf(keyword) !== -1 && categoryVal !== "") {
        return categoryVal;
      }
    }
  }
  
  // Keyword pattern inference engine
  let suggestedCategory = "Miscellaneous";
  if (/starbucks|mcdonald|dunkin|restaurant|cafe|grubhub|ubereats|doordash|pizza|deli|bakery/i.test(merchant)) {
    suggestedCategory = "Food & Drink";
  } else if (/kroger|walmart|wholefd|safeway|supermarket|trader joe|aldi|costco|landaus/i.test(merchant)) {
    suggestedCategory = "Groceries";
  } else if (/netflix|spotify|hulu|youtube|disney|steam|cinema|theater/i.test(merchant)) {
    suggestedCategory = "Entertainment";
  } else if (/uber|lyft|chevron|shell|bp|subway|transit|gas station|parking|toll/i.test(merchant)) {
    suggestedCategory = "Transportation";
  } else if (/comcast|electric|att|verizon|water|utility|insurance|power|tax/i.test(merchant)) {
    suggestedCategory = "Bills & Utilities";
  } else if (/amazon|target|home depot|ikea|lowes|furniture/i.test(merchant)) {
    suggestedCategory = "Home";
  } else if (/zelle|wire|transfer|payment to/i.test(merchant)) {
    suggestedCategory = "Transfer";
  } else if (/check|chk/i.test(merchant)) {
    suggestedCategory = "Checks & Banking";
  } else if (/payroll|direct dep|salary|refund/i.test(merchant)) {
    suggestedCategory = "Income";
  }

  // Auto-Learn new rule directly into Category_Rules_DB if rulesSheet exists
  if (rulesSheet && merchant.length > 2) {
    const cleanName = autoCleanMerchantName(merchantRaw);
    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
      try {
        const keyword = cleanName.toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
        if (keyword.length > 2) {
          rulesSheet.appendRow([keyword, cleanName, suggestedCategory]);
          console.log(`[Auto-Rule Engine] Auto-learned rule: "${keyword}" -> ${cleanName} (${suggestedCategory})`);
        }
      } catch (e) {
        console.error("Auto-rule append failed: " + e.toString());
      } finally {
        lock.releaseLock();
      }
    }
  }

  ensureCategoryInEnum(suggestedCategory, ss);
  return suggestedCategory;
}

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
    const category = data[i][7];
    const rawMerchant = data[i][3];
    const cleanMerchant = data[i][4] || autoCleanMerchantName(rawMerchant);
    
    if (!category || category === "" || category === "Miscellaneous") {
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
      if (suggestedCategory !== "Miscellaneous") {
        dbSheet.getRange(rowIdx, 5).setValue(cleanMerchant);
        dbSheet.getRange(rowIdx, 8).setValue(suggestedCategory);
        ensureCategoryInEnum(suggestedCategory, ss);
        updatedCount++;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert("Auto-Categorization Complete 🤖", `Successfully updated ${updatedCount} transactions and cleaned merchant titles.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===========================================================================
// 7. CATEGORIZATION ASSISTANT SIDEBAR ENGINE
// ===========================================================================

function showCategorizerSidebar() {
  const html = HtmlService.createTemplateFromFile("CategorizerView");
  const htmlOutput = html.evaluate()
    .setTitle("Categorization Assistant")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getUncategorizedQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return { queue: [], categories: [] };

  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return { queue: [], categories: [] };

  const rawValues = dbSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const queue = [];

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];
    const cat = row[7] ? row[7].toString().trim() : "";
    if (!cat || cat === "" || cat === "Miscellaneous") {
      queue.push({
        rowIdx: i + 2,
        date: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        rawMerchant: row[3],
        cleanMerchant: row[4] || autoCleanMerchantName(row[3]),
        amount: parseFloat(row[5]) || 0,
        suggestedCategory: autoCategorizeMerchant(row[4] || row[3])
      });
    }
  }

  const enums = getEnumValues();
  return {
    queue: queue.slice(0, 50),
    categories: enums.categories
  };
}

function applyCategoryWithRule(rowIdx, category, keywordRule, saveRule) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");

  if (!dbSheet) return "ERROR: Sheet not found";

  ensureCategoryInEnum(category, ss);
  dbSheet.getRange(rowIdx, 8).setValue(category);

  if (saveRule && keywordRule && rulesSheet) {
    const rawVal = dbSheet.getRange(rowIdx, 4).getValue().toString();
    const cleanVal = dbSheet.getRange(rowIdx, 5).getValue().toString() || autoCleanMerchantName(rawVal);
    rulesSheet.appendRow([keywordRule.toLowerCase().trim(), cleanVal, category]);

    const lastRow = dbSheet.getLastRow();
    if (lastRow > 1) {
      const merchants = dbSheet.getRange(2, 4, lastRow - 1, 1).getValues();
      const categories = dbSheet.getRange(2, 8, lastRow - 1, 1).getValues();
      for (let i = 0; i < merchants.length; i++) {
        const mStr = merchants[i][0].toString().toLowerCase();
        const cStr = categories[i][0] ? categories[i][0].toString().trim() : "";
        if (mStr.indexOf(keywordRule.toLowerCase().trim()) !== -1 && (!cStr || cStr === "Miscellaneous")) {
          dbSheet.getRange(i + 2, 8).setValue(category);
        }
      }
    }
  }
  return "SUCCESS";
}

// ===========================================================================
// 8. HIGH-PERFORMANCE DYNAMIC TRANSACTION EDITOR SIDEBAR
// ===========================================================================

function showTransactionEditorSidebar() {
  const html = HtmlService.createTemplateFromFile("TransactionEditorView");
  const htmlOutput = html.evaluate()
    .setTitle("Transaction Editor Ledger")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getTransactionsPage(pageNum, pageSize, searchQuery) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return { data: [], totalPages: 0 };
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return { data: [], totalPages: 0 };
  
  const rawValues = dbSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const query = (searchQuery || "").toLowerCase().trim();
  
  const mapped = rawValues.map((row, index) => {
    return {
      rowIdx: index + 2,
      txId: row[0] || "",
      timestamp: row[1] || "",
      date: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
      merchantRaw: row[3] || "",
      merchantClean: row[4] || "",
      amount: parseFloat(row[5]) || 0,
      flow: row[6] || "OUTFLOW",
      category: row[7] || "",
      accountSource: row[8] || "",
      channel: row[9] || "",
      status: row[10] || ""
    };
  }).reverse();
  
  const filtered = mapped.filter(item => {
    if (!query) return true;
    const catStr = (item.category || "").toString().toLowerCase();
    const merchCleanStr = (item.merchantClean || "").toString().toLowerCase();
    const merchRawStr = (item.merchantRaw || "").toString().toLowerCase();
    const acctStr = (item.accountSource || "").toString().toLowerCase();
    const txIdStr = (item.txId || "").toString().toLowerCase();
    const amtStr = (item.amount || 0).toString();
    const flowStr = (item.flow || "").toString().toLowerCase();

    return (
      merchRawStr.indexOf(query) !== -1 ||
      merchCleanStr.indexOf(query) !== -1 ||
      catStr.indexOf(query) !== -1 ||
      acctStr.indexOf(query) !== -1 ||
      txIdStr.indexOf(query) !== -1 ||
      amtStr.indexOf(query) !== -1 ||
      flowStr.indexOf(query) !== -1
    );
  });
  
  const size = pageSize || 10;
  const page = pageNum || 1;
  const totalPages = Math.ceil(filtered.length / size) || 1;
  const startIndex = (page - 1) * size;
  const pageData = filtered.slice(startIndex, startIndex + size);
  
  return {
    data: pageData,
    totalPages: totalPages,
    currentPage: page,
    totalRecords: filtered.length
  };
}

function getEnumValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  if (!enumSheet) return { categories: [], accounts: [] };
  
  const lastRow = enumSheet.getLastRow();
  if (lastRow < 3) return { categories: [], accounts: [] };
  
  const categories = enumSheet.getRange(3, 1, lastRow - 2, 1).getValues().flat().filter(String);
  const accounts = enumSheet.getRange(3, 2, lastRow - 2, 1).getValues().flat().filter(String);
  
  return { categories: categories, accounts: accounts };
}

function addOrUpdateTransaction(txData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return "ERROR: Sheet not found";
  
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const rawMerchant = txData.merchant || "Manual Entry";
    const cleanMerchant = autoCleanMerchantName(rawMerchant);
    const catVal = txData.category || autoCategorizeMerchant(cleanMerchant);
    const acctVal = txData.accountSource || "Cash";

    ensureCategoryInEnum(catVal, ss);
    ensureAccountInEnum(acctVal, ss);

    if (txData.rowIdx && parseInt(txData.rowIdx, 10) > 1) {
      const rowNum = parseInt(txData.rowIdx, 10);
      dbSheet.getRange(rowNum, 3).setValue(new Date(txData.date));
      dbSheet.getRange(rowNum, 4).setValue(rawMerchant);
      dbSheet.getRange(rowNum, 5).setValue(cleanMerchant);
      dbSheet.getRange(rowNum, 6).setValue(Math.abs(parseFloat(txData.amount)));
      dbSheet.getRange(rowNum, 7).setValue(txData.flow || "OUTFLOW");
      dbSheet.getRange(rowNum, 8).setValue(catVal);
      dbSheet.getRange(rowNum, 9).setValue(acctVal);
      return "SUCCESS_UPDATE";
    } else {
      const newTxId = "MANUAL-" + new Date().getTime();
      dbSheet.appendRow([
        newTxId, new Date(), new Date(txData.date), rawMerchant, cleanMerchant,
        Math.abs(parseFloat(txData.amount)), txData.flow || "OUTFLOW", catVal,
        acctVal, "MANUAL", "SETTLED"
      ]);
      return "SUCCESS_ADD";
    }
  } finally {
    lock.releaseLock();
  }
}

function deleteTransactionRow(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return "ERROR: Sheet not found";
  
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const rowNum = parseInt(rowIdx, 10);
    if (rowNum > 1 && rowNum <= dbSheet.getLastRow()) {
      dbSheet.deleteRow(rowNum);
      return "SUCCESS_DELETE";
    }
    return "ERROR: Invalid Row";
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// 9. ISOLATED TRIP & PROJECT EXPENSE TRACKER ENGINE
// ===========================================================================



function createTravelBudgetTab(tripTitle, totalBudget) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanTitle = (tripTitle || "Vacation_Trip").replace(/[^a-zA-Z0-9 ]/g, "").trim();
  const sheetName = "Trip_" + cleanTitle.replace(/\s+/g, "_");
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  

  

  formatHeader(sheet, "A1:F1", `${cleanTitle.toUpperCase()} ISOLATED EXPENSE TRACKER`);
  
  const kpiHeaders = [["Total Planned Budget", "Total Actual Spend", "Variance Remaining", "Trip Budget Status"]];
  sheet.getRange("A3:D3").setValues(kpiHeaders);
  formatSubHeader(sheet, "A3:D3");
  
  const budget = totalBudget || 3000.00;
  sheet.getRange("A4").setValue(budget).setNumberFormat("$#,##0.00").setFontWeight("bold").setFontColor("#002060");
  sheet.getRange("B4").setFormula("=SUM(C18:C100)").setNumberFormat("$#,##0.00").setFontWeight("bold").setFontColor("#db4437");
  sheet.getRange("C4").setFormula("=A4-B4").setNumberFormat("$#,##0.00").setFontWeight("bold").setFontColor("#385723");
  sheet.getRange("D4").setFormula('=IF(C4>=0, "ON TRACK ✅", "OVER BUDGET ❌")').setFontWeight("bold").setHorizontalAlignment("center");
  
  const catHeaders = [["Category", "Planned Budget", "Actual Spend", "Variance"]];
  sheet.getRange("A7:D7").setValues(catHeaders);
  formatSubHeader(sheet, "A7:D7");
  
  const categories = [
    ["Flights & Airfare", budget * 0.30, "=SUMIF(D18:D100, A9, C18:C100)", "=B9-C9"],
    ["Lodging & Hotels", budget * 0.35, "=SUMIF(D18:D100, A10, C18:C100)", "=B10-C10"],
    ["Meals & Dining", budget * 0.15, "=SUMIF(D18:D100, A11, C18:C100)", "=B11-C11"],
    ["Events & Tickets", budget * 0.10, "=SUMIF(D18:D100, A12, C18:C100)", "=B12-C12"],
    ["Ground Transportation", budget * 0.05, "=SUMIF(D18:D100, A13, C18:C100)", "=B13-C13"],
    ["Miscellaneous", budget * 0.05, "=SUMIF(D18:D100, A14, C18:C100)", "=B14-C14"]
  ];
  sheet.getRange(9, 1, categories.length, 4).setValues(categories);
  sheet.getRange("B9:D14").setNumberFormat("$#,##0.00");
  sheet.getRange("A9:A14").setFontWeight("bold");
  
  const logHeaders = [["Date", "Item Description", "Cost ($)", "Category", "Payment Account Source", "Notes / Confirmation"]];
  sheet.getRange("A17:F17").setValues(logHeaders);
  formatSubHeader(sheet, "A17:F17");
  
  const sampleExpenses = [
    ["2026-09-05", "Roundtrip Flight Tickets", 850.00, "Flights & Airfare", "Chase Sapphire 9012", "Confirmation #ABC123"],
    ["2026-09-05", "Hotel Stay Deposit", 450.00, "Lodging & Hotels", "Chase Sapphire 9012", "3 Nights Reservation"]
  ];
  sheet.getRange(18, 1, sampleExpenses.length, 6).setValues(sampleExpenses);
  
  const tripCats = ["Flights & Airfare", "Lodging & Hotels", "Meals & Dining", "Events & Tickets", "Ground Transportation", "Miscellaneous"];
  const catRule = SpreadsheetApp.newDataValidation().requireValueInList(tripCats).setAllowInvalid(false).build();
  sheet.getRange("D18:D100").setDataValidation(catRule);
  
  sheet.getRange("C18:C100").setNumberFormat("$#,##0.00").setFontColor("#002060");
  sheet.getRange("A18:A100").setNumberFormat("yyyy-mm-dd");
  sheet.autoResizeColumns(1, 6);
  
  console.log(`[Trip Engine] Compiled isolated trip tracker: ${sheetName}`);
}

// ===========================================================================
// 10. BACKGROUND AUTOMATED EMAIL & ALERT ENGINES
// ===========================================================================

function checkBudgetBreaches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  const enumSheet = ss.getSheetByName("Enum");
  if (!dash || !enumSheet) return;
  
  const year = dash.getRange("C4").getValue().toString().trim();
  const month = dash.getRange("D4").getValue().toString().trim();
  
  if (month === "All") return;
  
  const lastEnumRow = enumSheet.getLastRow();
  if (lastEnumRow < 3) return;
  const enumCategoriesCount = lastEnumRow - 2;
  
  const data = dash.getRange(10, 1, enumCategoriesCount, 4).getValues();
  const scriptProperties = PropertiesService.getScriptProperties();
  const userEmail = Session.getActiveUser().getEmail();
  let breaches = [];
  
  for (let i = 0; i < data.length; i++) {
    const category = data[i][0];
    const planned = parseFloat(data[i][1]);
    const actual = parseFloat(data[i][2]);
    const diff = parseFloat(data[i][3]);
    
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
    
    console.log(`[Breach Alert Engine] Dispatched budget breach email to ${userEmail} for ${breaches.length} categories.`);
  }
}

function sendWeeklyBudgetSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  const enumSheet = ss.getSheetByName("Enum");
  if (!dash || !enumSheet) return;

  const year = dash.getRange("C4").getValue().toString().trim();
  const month = dash.getRange("D4").getValue().toString().trim();
  if (month === "All") return;

  const userEmail = Session.getActiveUser().getEmail();
  const lastEnumRow = enumSheet.getLastRow();
  const categoryCount = lastEnumRow - 2;

  const totalIncome = parseFloat(dash.getRange("A7").getValue() || 0);
  const totalSpending = parseFloat(dash.getRange("B7").getValue() || 0);
  const netSavings = parseFloat(dash.getRange("C7").getValue() || 0);
  const savingsRate = (dash.getRange("D7").getValue() * 100).toFixed(1);

  const budgetData = dash.getRange(10, 1, categoryCount, 4).getValues();

  const today = new Date();
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const currentDay = today.getDate();
  const monthElapsedPercent = ((currentDay / daysInMonth) * 100).toFixed(0);

  let html = `
  <div style="font-family: Arial, sans-serif; background-color: #f4f6f9; padding: 20px; max-width: 600px; margin: auto; border: 1px solid #e1e4e8;">
    <div style="background-color: #1f4e78; color: white; padding: 25px; text-align: center; border-radius: 5px 5px 0 0;">
      <h2 style="margin: 0; font-size: 22px; letter-spacing: 1px;">WEEKLY FINANCIAL BRIEFING</h2>
      <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Period Status: ${month} ${year} (${monthElapsedPercent}% Elapsed)</p>
    </div>

    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px;">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Month-to-Date KPIs</h3>
      <table style="width: 100%; text-align: center; border-collapse: collapse;">
        <tr>
          <td style="padding: 10px; width: 25%;">
            <span style="font-size: 11px; color: #666; display: block;">TOTAL INCOME</span>
            <span style="font-weight: bold; font-size: 16px; color: #385723;">$${totalIncome.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block;">TOTAL SPENDING</span>
            <span style="font-weight: bold; font-size: 16px; color: #db4437;">$${totalSpending.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block;">NET SAVINGS</span>
            <span style="font-weight: bold; font-size: 16px; color: ${netSavings >= 0 ? '#385723' : '#c00000'};">$${netSavings.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block;">SAVINGS RATE</span>
            <span style="font-weight: bold; font-size: 16px; color: #1f4e78;">${savingsRate}%</span>
          </td>
        </tr>
      </table>
    </div>

    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px;">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Budget Performance Breakdown</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <thead>
          <tr style="background-color: #f8f9fa; color: #333; text-align: left;">
            <th style="padding: 8px;">Category</th>
            <th style="padding: 8px; text-align: right;">Planned</th>
            <th style="padding: 8px; text-align: right;">Actual</th>
            <th style="padding: 8px; text-align: right;">Status</th>
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

    <div style="margin-top: 20px; text-align: center; color: #666; font-size: 12px;">
      <p>Dispatched automatically by your <strong>Google Sheets Financial Automation System 🚀</strong></p>
    </div>
  </div>`;

  MailApp.sendEmail({
    to: userEmail,
    subject: `📊 Weekly Financial Summary: ${month} ${year} Briefing`,
    htmlBody: html
  });

  console.log("[Weekly Summary Engine] Dispatched HTML briefing email.");
}

function promptSavingsSweep() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  
  if (!dash) {
    ui.alert("Error ❌", "The Monthly Dashboard Tab not found.", ui.ButtonSet.OK);
    return;
  }

  const year = dash.getRange("C4").getValue().toString().trim();
  const month = dash.getRange("D4").getValue().toString().trim();
  
  if (month === "All") {
    ui.alert("Invalid Selection ⚠️", "Please select a specific calendar month on your Dashboard before executing the savings sweep.", ui.ButtonSet.OK);
    return;
  }

  const netSavings = parseFloat(dash.getRange("C7").getValue() || 0);
  if (netSavings <= 0) {
    ui.alert("No Net Surplus ℹ️", `Net Savings for ${month} ${year} is $${netSavings.toFixed(2)}. A sweep requires a positive surplus.`, ui.ButtonSet.OK);
    return;
  }

  const title = `Execute Monthly Savings Sweep (${month} ${year})? 🚀`;
  const msg = `Net Savings Surplus: $${netSavings.toFixed(2)}\n\n` +
               `This will allocate your net surplus across your Savings Goals based on active target percentages on the 'Savings_Goals' tab.\n\n` +
               `Do you want to proceed?`;

  const resp = ui.alert(title, msg, ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) {
    try {
      const summary = executeSavingsSweep(ss, year, month, netSavings);
      ui.alert("Savings Sweep Complete! 🎉", summary, ui.ButtonSet.OK);
    } catch (err) {
      ui.alert("Sweep Execution Failed ❌", err.toString(), ui.ButtonSet.OK);
    }
  }
}

function executeSavingsSweep(ss, year, month, netSavings) {
  const goalsSheet = ss.getSheetByName("Savings_Goals");
  const dbSheet = ss.getSheetByName("Transactions_DB");
  
  if (!goalsSheet || !dbSheet) {
    throw new Error("Required database sheets ('Savings_Goals' or 'Transactions_DB') are missing.");
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const sweepKey = "sweep_run_" + year + "_" + month;
  const alreadyRun = scriptProperties.getProperty(sweepKey);

  if (alreadyRun) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert("Duplicate Sweep Warning ⚠️", `A savings sweep for ${month} ${year} was already executed previously.\nDo you want to force re-run?`, ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) {
      return "Sweep canceled by user.";
    }
  }

  const lastGoalRow = goalsSheet.getLastRow();
  if (lastGoalRow < 3) throw new Error("No goals found on 'Savings_Goals' tab.");

  const data = goalsSheet.getRange(3, 1, lastGoalRow - 2, 6).getValues();
  let totalPercentage = 0;
  
  for (let i = 0; i < data.length; i++) {
    totalPercentage += parseFloat(data[i][4]) || 0;
  }

  if (totalPercentage > 1.01) {
    throw new Error(`Total sweep allocation percentages exceed 100% (${(totalPercentage * 100).toFixed(0)}%). Please correct allocations on Savings_Goals.`);
  }

  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthIdx = months.indexOf(month);
  const sweepDate = new Date(parseInt(year, 10), monthIdx + 1, 0);

  let logSummary = `Successfully swept $${netSavings.toFixed(2)} net surplus for ${month} ${year}:\n`;
  let rowsAdded = 0;

  for (let i = 0; i < data.length; i++) {
    const goalName = data[i][0];
    const currentBal = parseFloat(data[i][2]) || 0;
    const sweepPct = parseFloat(data[i][4]) || 0;

    if (sweepPct <= 0) continue;

    const allocationAmount = netSavings * sweepPct;
    const newBal = currentBal + allocationAmount;
    goalsSheet.getRange(3 + i, 3).setValue(newBal);

    const timestamp = new Date();
    const txIdOut = "SWEEP-OUT-" + timestamp.getTime() + "-" + i;
    dbSheet.appendRow([
      txIdOut, timestamp, sweepDate, `Savings Sweep Allocation: ${goalName}`, `Savings Sweep: ${goalName}`,
      allocationAmount, "OUTFLOW", "Savings & Investments", "Chase Checking 1234", "SAVINGS_SWEEP", "SETTLED"
    ]);

    const txIdIn = "SWEEP-IN-" + timestamp.getTime() + "-" + i;
    dbSheet.appendRow([
      txIdIn, timestamp, sweepDate, `Savings Sweep Received: ${goalName}`, `Savings Sweep: ${goalName}`,
      allocationAmount, "INFLOW", "Savings & Investments", "Savings", "SAVINGS_SWEEP", "SETTLED"
    ]);

    logSummary += `\n• ${goalName}: $${allocationAmount.toFixed(2)} (${(sweepPct * 100).toFixed(0)}%)`;
    rowsAdded++;
  }

  if (rowsAdded > 0) {
    scriptProperties.setProperty(sweepKey, "true");
    return logSummary;
  } else {
    return "No active allocations computed. Ensure goals have non-zero sweep percentages.";
  }
}

function formatCurrency(amount) {
  return "$" + parseFloat(amount).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}


// ===========================================================================
// 12. CERTIFICATE OF DEPOSIT (CD) PORTFOLIO & VARIABLE RENEWAL ENGINE
// ===========================================================================

/**
 * Menu Trigger: Prompts user and constructs/resets the CD_Tracker sheet tab.
 */
function setupCDTracker() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cdSheet = ss.getSheetByName("CD_Tracker");
  
  if (!cdSheet) {
    cdSheet = ss.insertSheet("CD_Tracker");
  } else {
    const res = ui.alert("Reset CD Tracker? ⚠️", "Do you want to re-initialize the CD_Tracker tab? Existing data will be reformatted.", ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    cdSheet.clear();
  }

  // Title Block
  formatHeader(cdSheet, "A1:L1", "CERTIFICATE OF DEPOSIT (CD) PORTFOLIO & RENEWAL TRACKER");
  
  // Table Headers
  const headers = [["Bank / Institution", "Certificate #", "Principal Deposit", "Issue Date", "Term (Months)", "Maturity Date", "Initial APY", "Auto-Renew?", "Renewal APY", "Current Value", "Maturity Payout", "Status"]];
  cdSheet.getRange("A2:L2").setValues(headers).setFontWeight("bold");
  formatSubHeader(cdSheet, "A2:L2");
  cdSheet.setRowHeight(2, 25);

  // Sample CD Portfolio rows (demonstrating variable rates and renewal terms)
  const sampleCDs = [
    ["Marcus by Goldman Sachs", "CD-8812", 10000.00, "2025-10-15", 12, "", 0.0525, "YES", 0.0450, "", "", ""],
    ["Discover Bank", "CD-9041", 15000.00, "2026-01-10", 18, "", 0.0500, "YES", 0.0425, "", "", ""],
    ["Capital One", "CD-3321", 5000.00, "2026-05-01", 6, "", 0.0475, "NO", 0.0000, "", "", ""]
  ];

  cdSheet.getRange(3, 1, sampleCDs.length, 12).setValues(sampleCDs);

  // Dynamic Formulas for each CD row
  for (let i = 0; i < sampleCDs.length; i++) {
    const r = 3 + i;
    // Maturity Date: =EDATE(Issue_Date, Term)
    cdSheet.getRange(r, 6).setFormula(`=EDATE(D${r}, E${r})`);
    
    // Current Value: Accounts for compounding elapsed time & auto-renewal transition if TODAY() > Maturity Date
    cdSheet.getRange(r, 10).setFormula(
      `=IF(D${r}="", 0, IF(TODAY()<=F${r}, C${r}*(1 + G${r}*(MAX(0, TODAY()-D${r})/365)), IF(H${r}="YES", (C${r}*(1+G${r}*(E${r}/12)))*(1 + I${r}*(MAX(0, TODAY()-F${r})/365)), C${r}*(1+G${r}*(E${r}/12)))))`
    );

    // Total Maturity Payout
    cdSheet.getRange(r, 11).setFormula(`=C${r}*(1 + G${r}*(E${r}/12))`);

    // Status Indicator
    cdSheet.getRange(r, 12).setFormula(
      `=IF(TODAY()<F${r}, "ACTIVE 🟢", IF(H${r}="YES", "RENEWED 🔄", "MATURED 🟡"))`
    );
  }

  // Formats
  cdSheet.getRange("C3:C100").setNumberFormat("$#,##0.00");
  cdSheet.getRange("J3:K100").setNumberFormat("$#,##0.00");
  cdSheet.getRange("D3:D100").setNumberFormat("yyyy-mm-dd");
  cdSheet.getRange("F3:F100").setNumberFormat("yyyy-mm-dd");
  cdSheet.getRange("G3:G100").setNumberFormat("0.00%");
  cdSheet.getRange("I3:I100").setNumberFormat("0.00%");

  // Auto-Renew Data Validation Dropdown
  const renewValidation = SpreadsheetApp.newDataValidation().requireValueInList(["YES", "NO"]).setAllowInvalid(false).build();
  cdSheet.getRange("H3:H100").setDataValidation(renewValidation);

  // Portfolio Totals Summary Row
  const totalRow = 3 + sampleCDs.length + 1;
  cdSheet.getRange(totalRow, 1).setValue("Total CD Portfolio Value:").setFontWeight("bold");
  cdSheet.getRange(totalRow, 3).setFormula(`=SUM(C3:C${totalRow-2})`).setFontWeight("bold").setNumberFormat("$#,##0.00");
  cdSheet.getRange(totalRow, 10).setFormula(`=SUM(J3:J${totalRow-2})`).setFontWeight("bold").setNumberFormat("$#,##0.00");
  cdSheet.getRange(totalRow, 11).setFormula(`=SUM(K3:K${totalRow-2})`).setFontWeight("bold").setNumberFormat("$#,##0.00");
  cdSheet.getRange(totalRow, 1, 1, 12).setBackground("#f2f2f2");

  cdSheet.autoResizeColumns(1, 12);

  // Programmatic CD Portfolio Bar Chart
  try {
    const existingCharts = cdSheet.getCharts();
    existingCharts.forEach(c => cdSheet.removeChart(c));

    const cdChart = cdSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(cdSheet.getRange(2, 1, sampleCDs.length + 1, 1)) // Bank name
      .addRange(cdSheet.getRange(2, 3, sampleCDs.length + 1, 1)) // Principal
      .addRange(cdSheet.getRange(2, 10, sampleCDs.length + 1, 1)) // Current Value
      .setPosition(totalRow + 2, 1, 0, 0)
      .setOption('title', 'CD Principal vs Current Value Accrual')
      .setOption('legend', { position: 'top' })
      .setOption('width', 600)
      .setOption('height', 320)
      .setOption('colors', ['#1f4e78', '#2e7d32'])
      .build();
    cdSheet.insertChart(cdChart);
  } catch (err) {
    console.error("CD Chart build error: " + err.toString());
  }

  ui.alert("CD Portfolio Tracker Online! 📈", "The CD_Tracker tab has been generated with dynamic compounding formulas, renewal rates, and portfolio charts.", ui.ButtonSet.OK);
}

// ===========================================================================
// 13. AUTOMATED TRAVEL PLANNER & ITINERARY GENERATOR
// ===========================================================================

/**
 * Menu Trigger: Prompts user for a trip name and builds a structured Travel Itinerary/Agenda tab.
 */
function promptTravelItinerary() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Create Travel Itinerary ✈️", "Please enter the Trip Name (e.g., Israel_TLV_2026 or NYC_Weekend):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    const tripTitle = resp.getResponseText().trim().replace(/[^a-zA-Z0-9_]/g, "_");
    if (!tripTitle) {
      ui.alert("Invalid Name", "Please provide a valid trip title.", ui.ButtonSet.OK);
      return;
    }
    createTravelAgendaTab(tripTitle);
  }
}

/**
 * Constructs a comprehensive Travel Itinerary & Agenda tab based on travel planner templates.
 */
function createTravelAgendaTab(tripTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Agenda_" + tripTitle;
  let agendaSheet = ss.getSheetByName(sheetName);

  if (!agendaSheet) {
    agendaSheet = ss.insertSheet(sheetName);
  } else {
    agendaSheet.clear();
  }

  // Header Banner
  formatHeader(agendaSheet, "A1:G1", "TRAVEL ITINERARY & AGENDA: " + tripTitle.replace(/_/g, " ").toUpperCase());

  // 1. FLIGHT & TRANSPORTATION DETAILS BLOCK
  agendaSheet.getRange("A3:G3").setValues([["Flight / Transportation Details", "", "", "", "", "", ""]]);
  formatSubHeader(agendaSheet, "A3:G3");
  const flightHeaders = [["Leg / Direction", "Carrier", "Flight #", "Departure Date & Time", "Arrival Date & Time", "Confirmation Code", "Seat / Class"]];
  agendaSheet.getRange("A4:G4").setValues(flightHeaders).setFontWeight("bold");

  const sampleFlights = [
    ["Outbound", "Delta Airlines", "DL-198", "2026-09-20 18:30", "2026-09-21 11:45", "DL-CONF-8812", "14B / Economy"],
    ["Return", "Delta Airlines", "DL-199", "2026-09-30 14:15", "2026-09-30 20:00", "DL-CONF-8812", "15A / Economy"]
  ];
  agendaSheet.getRange(5, 1, sampleFlights.length, 7).setValues(sampleFlights);

  // 2. ACCOMMODATION BLOCK
  const hotelRow = 8;
  agendaSheet.getRange(`A${hotelRow}:G${hotelRow}`).setValues([["Accommodation & Lodging", "", "", "", "", "", ""]]);
  formatSubHeader(agendaSheet, `A${hotelRow}:G${hotelRow}`);
  const hotelHeaders = [["Hotel / Airbnb Name", "Address / Location", "Check-In Date", "Check-Out Date", "Confirmation #", "Contact Phone", "Nightly Rate"]];
  agendaSheet.getRange(hotelRow + 1, 1, 1, 7).setValues(hotelHeaders).setFontWeight("bold");

  const sampleHotels = [
    ["Tel Aviv Beach Hotel", "Hayarkon St 121, Tel Aviv", "2026-09-21", "2026-09-30", "HOTEL-TLV-991", "+972 3-555-0192", 220.00]
  ];
  agendaSheet.getRange(hotelRow + 2, 1, sampleHotels.length, 7).setValues(sampleHotels);
  agendaSheet.getRange(`G${hotelRow + 2}`).setNumberFormat("$#,##0.00");

  // 3. DAILY SCHEDULE & ITINERARY MATRIX
  const scheduleRow = 12;
  agendaSheet.getRange(`A${scheduleRow}:G${scheduleRow}`).setValues([["Daily Agenda & Activity Schedule", "", "", "", "", "", ""]]);
  formatSubHeader(agendaSheet, `A${scheduleRow}:G${scheduleRow}`);
  const scheduleHeaders = [["Date", "Time", "Activity / Event Description", "Location / Venue", "Reservation / Tickets", "Est. Cost ($)", "Status"]];
  agendaSheet.getRange(scheduleRow + 1, 1, 1, 7).setValues(scheduleHeaders).setFontWeight("bold");

  const sampleActivities = [
    ["2026-09-21", "14:00", "Hotel Check-in & Rest", "Tel Aviv Beach Hotel", "Confirmed", 0.00, "COMPLETED"],
    ["2026-09-22", "10:00", "Old Jaffa Walking Tour", "Clock Tower, Jaffa", "Tour Group A", 45.00, "PLANNED"],
    ["2026-09-23", "19:00", "Group Dinner @ Schwartz TLV", "Rothschild Blvd", "Res #781", 120.00, "PLANNED"]
  ];
  agendaSheet.getRange(scheduleRow + 2, 1, sampleActivities.length, 7).setValues(sampleActivities);
  agendaSheet.getRange(`F${scheduleRow + 2}:F${scheduleRow + 2 + sampleActivities.length}`).setNumberFormat("$#,##0.00");

  agendaSheet.autoResizeColumns(1, 7);

  const ui = SpreadsheetApp.getUi();
  ui.alert("Itinerary Created! ✈️", "Created " + sheetName + " tab with pre-built Flight, Lodging, and Daily Activity schedules.", ui.ButtonSet.OK);
}


/**
 * Quick-Add Tool: Prompts user to enter a new Certificate of Deposit (CD) investment,
 * appends the row to 'CD_Tracker', and automatically configures compounding formulas, 
 * maturity dates, renewal rates, and summary totals.
 */



// ===========================================================================
// 14. USER-FRIENDLY HTML MODAL DIALOG INPUT ENGINES (V16 UPGRADE)
// ===========================================================================

/**
 * Menu Action: Opens a user-friendly HTML form modal to add a Certificate of Deposit (CD).
 */
function promptAddCD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cdSheet = ss.getSheetByName("CD_Tracker");
  if (!cdSheet) {
    SpreadsheetApp.getUi().alert("CD_Tracker Tab Missing ⚠️", "Please run '📈 Build CD Portfolio Growth Tracker' from the menu first to initialize the CD tab.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
        <style>
          body { padding: 15px; font-family: Arial, sans-serif; background-color: #fafafa; }
          .modal-title { font-size: 1.2rem; font-weight: bold; color: #1f4e78; margin-bottom: 15px; text-align: center; }
          .input-field { margin-bottom: 10px; }
          .input-field label { font-size: 0.85rem; color: #555; }
          .btn-submit { background-color: #1f4e78 !important; width: 100%; margin-top: 15px; }
          .btn-cancel { background-color: #757575 !important; width: 100%; margin-top: 5px; }
          .row { margin-bottom: 5px; }
        </style>
      </head>
      <body>
        <div class="modal-title">📈 Add Certificate of Deposit (CD)</div>
        <form id="cdForm">
          <div class="row">
            <div class="input-field col s6">
              <input id="bank" type="text" placeholder="e.g. Marcus by Goldman Sachs" required>
              <label class="active" for="bank">Bank / Institution</label>
            </div>
            <div class="input-field col s6">
              <input id="certNum" type="text" placeholder="e.g. CD-8812" required>
              <label class="active" for="certNum">Certificate #</label>
            </div>
          </div>
          <div class="row">
            <div class="input-field col s6">
              <input id="principal" type="number" step="0.01" placeholder="10000.00" required>
              <label class="active" for="principal">Principal Deposit ($)</label>
            </div>
            <div class="input-field col s6">
              <input id="issueDate" type="date" required>
              <label class="active" for="issueDate">Issue Date</label>
            </div>
          </div>
          <div class="row">
            <div class="input-field col s4">
              <input id="term" type="number" value="12" min="1" required>
              <label class="active" for="term">Term (Months)</label>
            </div>
            <div class="input-field col s4">
              <input id="initialApy" type="number" step="0.01" placeholder="5.25" required>
              <label class="active" for="initialApy">Initial APY (%)</label>
            </div>
            <div class="input-field col s4">
              <select id="autoRenew" class="browser-default" style="margin-top: 5px; height: 2.2rem; border-radius: 4px;">
                <option value="YES" selected>YES (Auto-Renew)</option>
                <option value="NO">NO (Payout)</option>
              </select>
              <label class="active" style="top: -12px;">Auto-Renew?</label>
            </div>
          </div>
          <div class="row">
            <div class="input-field col s12">
              <input id="renewalApy" type="number" step="0.01" placeholder="4.50" value="4.50">
              <label class="active" for="renewalApy">Renewal APY (%) (if auto-renewed)</label>
            </div>
          </div>
          <button class="btn waves-effect waves-light btn-submit" type="button" onclick="submitCD()">
            Save CD Investment 🚀
          </button>
        </form>

        <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
        <script>
          document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('issueDate').valueAsDate = new Date();
          });

          function submitCD() {
            const bank = document.getElementById('bank').value.trim();
            const certNum = document.getElementById('certNum').value.trim();
            const principal = parseFloat(document.getElementById('principal').value) || 0;
            const issueDate = document.getElementById('issueDate').value;
            const term = parseInt(document.getElementById('term').value, 10) || 12;
            const initialApy = (parseFloat(document.getElementById('initialApy').value) || 0) / 100;
            const autoRenew = document.getElementById('autoRenew').value;
            const renewalApy = (parseFloat(document.getElementById('renewalApy').value) || 0) / 100;

            if (!bank || !certNum || principal <= 0 || !issueDate) {
              M.toast({html: 'Please fill in Bank, Cert #, Principal, and Date!', classes: 'red'});
              return;
            }

            M.toast({html: 'Saving CD Investment... ⏳', classes: 'blue'});
            google.script.run
              .withSuccessHandler(function(res) {
                M.toast({html: res, classes: 'green'});
                setTimeout(function() { google.script.host.close(); }, 1200);
              })
              .withFailureHandler(function(err) {
                M.toast({html: 'Error: ' + err.message, classes: 'red'});
              })
              .addCDFromForm({
                bank: bank,
                certNum: certNum,
                principal: principal,
                issueDate: issueDate,
                term: term,
                initialApy: initialApy,
                autoRenew: autoRenew,
                renewalApy: renewalApy
              });
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(530);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Add Certificate of Deposit (CD)");
}

/**
 * Backend Handler: Receives clean JSON object from the CD Modal form and appends it to CD_Tracker.
 */
function addCDFromForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cdSheet = ss.getSheetByName("CD_Tracker");
  if (!cdSheet) throw new Error("CD_Tracker sheet not found.");

  const lastRow = cdSheet.getLastRow();
  let targetRow = 3;
  
  if (lastRow >= 3) {
    const banks = cdSheet.getRange(3, 1, lastRow - 2, 1).getValues();
    let foundTotal = false;
    for (let i = 0; i < banks.length; i++) {
      if (banks[i][0].toString().toLowerCase().includes("total cd portfolio")) {
        targetRow = 3 + i;
        foundTotal = true;
        break;
      }
    }
    if (!foundTotal) targetRow = lastRow + 1;
  }

  cdSheet.insertRowBefore(targetRow);
  
  cdSheet.getRange(targetRow, 1).setValue(data.bank);
  cdSheet.getRange(targetRow, 2).setValue(data.certNum);
  cdSheet.getRange(targetRow, 3).setValue(data.principal).setNumberFormat("$#,##0.00");
  cdSheet.getRange(targetRow, 4).setValue(new Date(data.issueDate)).setNumberFormat("yyyy-mm-dd");
  cdSheet.getRange(targetRow, 5).setValue(data.term).setNumberFormat("#,##0");
  
  cdSheet.getRange(targetRow, 6).setFormula(`=EDATE(D${targetRow}, E${targetRow})`).setNumberFormat("yyyy-mm-dd");
  cdSheet.getRange(targetRow, 7).setValue(data.initialApy).setNumberFormat("0.00%");
  cdSheet.getRange(targetRow, 8).setValue(data.autoRenew);
  cdSheet.getRange(targetRow, 9).setValue(data.renewalApy).setNumberFormat("0.00%");
  
  cdSheet.getRange(targetRow, 10).setFormula(
    `=IF(D${targetRow}="", 0, IF(TODAY()<=F${targetRow}, C${targetRow}*(1 + G${targetRow}*(MAX(0, TODAY()-D${targetRow})/365)), IF(H${targetRow}="YES", (C${targetRow}*(1+G${targetRow}*(E${targetRow}/12)))*(1 + I${targetRow}*(MAX(0, TODAY()-F${targetRow})/365)), C${targetRow}*(1+G${targetRow}*(E${targetRow}/12)))))`
  ).setNumberFormat("$#,##0.00");

  cdSheet.getRange(targetRow, 11).setFormula(`=C${targetRow}*(1 + G${targetRow}*(E${targetRow}/12))`).setNumberFormat("$#,##0.00");
  
  cdSheet.getRange(targetRow, 12).setFormula(
    `=IF(TODAY()<F${targetRow}, "ACTIVE 🟢", IF(H${targetRow}="YES", "RENEWED 🔄", "MATURED 🟡"))`
  );

  const renewValidation = SpreadsheetApp.newDataValidation().requireValueInList(["YES", "NO"]).setAllowInvalid(false).build();
  cdSheet.getRange(targetRow, 8).setDataValidation(renewValidation);

  const newLastRow = cdSheet.getLastRow();
  let totRow = 0;
  for (let r = 3; r <= newLastRow; r++) {
    const val = cdSheet.getRange(r, 1).getValue().toString();
    if (val.toLowerCase().includes("total cd portfolio")) {
      totRow = r;
      break;
    }
  }

  if (totRow > 0) {
    cdSheet.getRange(totRow, 3).setFormula(`=SUM(C3:C${totRow-1})`);
    cdSheet.getRange(totRow, 10).setFormula(`=SUM(J3:J${totRow-1})`);
    cdSheet.getRange(totRow, 11).setFormula(`=SUM(K3:K${totRow-1})`);
  }

  cdSheet.autoResizeColumns(1, 12);
  return `Successfully added ${data.bank} (${data.certNum}) for $${data.principal.toFixed(2)}! 🎉`;
}

/**
 * Menu Action: Opens a user-friendly HTML modal to quick-add a transaction.
 */
function quickAddTransactionPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enums = getEnumValues();
  
  let catOptions = enums.categories.map(c => `<option value="${c}">${c}</option>`).join('');
  let acctOptions = enums.accounts.map(a => `<option value="${a}">${a}</option>`).join('');

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
        <style>
          body { padding: 15px; font-family: Arial, sans-serif; background-color: #fafafa; }
          .modal-title { font-size: 1.2rem; font-weight: bold; color: #1f4e78; margin-bottom: 15px; text-align: center; }
          .btn-submit { background-color: #1f4e78 !important; width: 100%; margin-top: 15px; }
          .row { margin-bottom: 5px; }
          select { display: block; height: 2.2rem; border-radius: 4px; border: 1px solid #ccc; }
        </style>
      </head>
      <body>
        <div class="modal-title">💳 Quick Add Transaction</div>
        <form id="txForm">
          <div class="row">
            <div class="input-field col s6">
              <input id="txDate" type="date" required>
              <label class="active">Date</label>
            </div>
            <div class="input-field col s6">
              <input id="amount" type="number" step="0.01" placeholder="0.00" required>
              <label class="active">Amount ($)</label>
            </div>
          </div>
          <div class="row">
            <div class="input-field col s12">
              <input id="merchant" type="text" placeholder="e.g. Starbucks, Target, Delta" required>
              <label class="active">Merchant Name / Description</label>
            </div>
          </div>
          <div class="row">
            <div class="col s4">
              <label style="font-size: 0.8rem; color: #555;">Capital Flow</label>
              <select id="flow">
                <option value="OUTFLOW" selected>OUTFLOW (Expense)</option>
                <option value="INFLOW">INFLOW (Income)</option>
                <option value="TRANSFER">TRANSFER (Internal)</option>
              </select>
            </div>
            <div class="col s4">
              <label style="font-size: 0.8rem; color: #555;">Category</label>
              <select id="category">
                <option value="" disabled selected>Select Category...</option>
                ${catOptions}
              </select>
            </div>
            <div class="col s4">
              <label style="font-size: 0.8rem; color: #555;">Account Source</label>
              <select id="account">
                <option value="" disabled selected>Select Account...</option>
                ${acctOptions}
              </select>
            </div>
          </div>
          <div class="row" style="margin-top: 10px;">
            <div class="input-field col s12">
              <input id="tripTag" type="text" placeholder="e.g. Israel_TLV_2026 (Optional)">
              <label class="active">Trip / Project Tag (Optional)</label>
            </div>
          </div>
          <button class="btn waves-effect waves-light btn-submit" type="button" onclick="submitTx()">
            Submit Transaction 🚀
          </button>
        </form>

        <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
        <script>
          document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('txDate').valueAsDate = new Date();
          });

          function submitTx() {
            const date = document.getElementById('txDate').value;
            const merchant = document.getElementById('merchant').value.trim();
            const amount = parseFloat(document.getElementById('amount').value) || 0;
            const flow = document.getElementById('flow').value;
            const category = document.getElementById('category').value;
            const account = document.getElementById('account').value;
            const tripTag = document.getElementById('tripTag').value.trim();

            if (!date || !merchant || amount <= 0) {
              M.toast({html: 'Please fill in Date, Merchant, and Amount!', classes: 'red'});
              return;
            }

            M.toast({html: 'Saving Transaction... ⏳', classes: 'blue'});
            google.script.run
              .withSuccessHandler(function(res) {
                M.toast({html: res, classes: 'green'});
                setTimeout(function() { google.script.host.close(); }, 1200);
              })
              .withFailureHandler(function(err) {
                M.toast({html: 'Error: ' + err.message, classes: 'red'});
              })
              .quickAddTxFromForm({
                date: date,
                merchant: merchant,
                amount: amount,
                flow: flow,
                category: category,
                account: account,
                tripTag: tripTag
              });
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(520)
    .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Quick Add Transaction");
}

/**
 * Backend Handler: Receives JSON from Quick Add Tx modal and appends to Transactions_DB.
 */
function quickAddTxFromForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) throw new Error("Transactions_DB sheet not found.");

  const rawMerchant = data.merchant || "Manual Expense";
  const cleanMerchant = autoCleanMerchantName(rawMerchant);
  const catVal = data.category || autoCategorizeMerchant(cleanMerchant);
  const acctVal = data.account || "Cash";

  ensureCategoryInEnum(catVal, ss);
  ensureAccountInEnum(acctVal, ss);

  const txId = "MANUAL-" + new Date().getTime();
  dbSheet.appendRow([
    txId,
    new Date(),
    new Date(data.date),
    rawMerchant,
    cleanMerchant,
    Math.abs(parseFloat(data.amount)),
    data.flow || "OUTFLOW",
    catVal,
    acctVal,
    "MANUAL",
    "SETTLED",
    data.tripTag || ""
  ]);

  return `Logged $${Math.abs(parseFloat(data.amount)).toFixed(2)} at ${cleanMerchant}! 🎉`;
}

/**
 * Menu Action: Opens a user-friendly HTML modal to add a Merchant Category Rule.
 */
function quickAddRulePrompt() {
  const enums = getEnumValues();
  let catOptions = enums.categories.map(c => `<option value="${c}">${c}</option>`).join('');

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
        <style>
          body { padding: 15px; font-family: Arial, sans-serif; background-color: #fafafa; }
          .modal-title { font-size: 1.2rem; font-weight: bold; color: #1f4e78; margin-bottom: 15px; text-align: center; }
          .btn-submit { background-color: #1f4e78 !important; width: 100%; margin-top: 15px; }
          select { display: block; height: 2.2rem; border-radius: 4px; border: 1px solid #ccc; margin-top: 5px; }
        </style>
      </head>
      <body>
        <div class="modal-title">🏷️ Add Merchant Category Rule</div>
        <form id="ruleForm">
          <div class="input-field">
            <input id="keyword" type="text" placeholder="e.g. starbucks, delta, trader joe" required>
            <label class="active">Merchant Keyword (Lowercase)</label>
          </div>
          <div class="input-field">
            <input id="cleanName" type="text" placeholder="e.g. Starbucks, Delta Airlines" required>
            <label class="active">Clean Display Name</label>
          </div>
          <div style="margin-top: 10px;">
            <label style="font-size: 0.8rem; color: #555;">Target Budget Category</label>
            <select id="category">
              <option value="" disabled selected>Select Category...</option>
              ${catOptions}
            </select>
          </div>
          <button class="btn waves-effect waves-light btn-submit" type="button" onclick="submitRule()">
            Save Rule & Apply to Ledger 🚀
          </button>
        </form>

        <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
        <script>
          function submitRule() {
            const keyword = document.getElementById('keyword').value.toLowerCase().trim();
            const cleanName = document.getElementById('cleanName').value.trim();
            const category = document.getElementById('category').value;

            if (!keyword || !cleanName || !category) {
              M.toast({html: 'Please fill in Keyword, Clean Name, and Category!', classes: 'red'});
              return;
            }

            M.toast({html: 'Saving Category Rule... ⏳', classes: 'blue'});
            google.script.run
              .withSuccessHandler(function(res) {
                M.toast({html: res, classes: 'green'});
                setTimeout(function() { google.script.host.close(); }, 1200);
              })
              .withFailureHandler(function(err) {
                M.toast({html: 'Error: ' + err.message, classes: 'red'});
              })
              .addRuleFromForm({ keyword: keyword, cleanName: cleanName, category: category });
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Add Merchant Category Rule");
}

/**
 * Backend Handler: Saves rule to Category_Rules_DB and retroactively updates matching uncategorized rows.
 */
function addRuleFromForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!rulesSheet) throw new Error("Category_Rules_DB sheet not found.");

  ensureCategoryInEnum(data.category, ss);
  rulesSheet.appendRow([data.keyword.toLowerCase().trim(), data.cleanName, data.category]);

  let updatedRows = 0;
  if (dbSheet) {
    const lastRow = dbSheet.getLastRow();
    if (lastRow > 1) {
      const rawData = dbSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (let i = 0; i < rawData.length; i++) {
        const mStr = rawData[i][3].toString().toLowerCase();
        const cStr = rawData[i][7] ? rawData[i][7].toString().trim() : "";
        if (mStr.indexOf(data.keyword.toLowerCase().trim()) !== -1 && (!cStr || cStr === "Miscellaneous")) {
          dbSheet.getRange(i + 2, 5).setValue(data.cleanName);
          dbSheet.getRange(i + 2, 8).setValue(data.category);
          updatedRows++;
        }
      }
    }
  }

  return `Saved rule for '${data.cleanName}'! Updated ${updatedRows} existing rows. 🎉`;
}

/**
 * Menu Action: Opens a user-friendly HTML modal to create a Trip / Project Tracker.
 */
function promptTripBudget() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
        <style>
          body { padding: 15px; font-family: Arial, sans-serif; background-color: #fafafa; }
          .modal-title { font-size: 1.2rem; font-weight: bold; color: #1f4e78; margin-bottom: 15px; text-align: center; }
          .btn-submit { background-color: #1f4e78 !important; width: 100%; margin-top: 15px; }
        </style>
      </head>
      <body>
        <div class="modal-title">✈️ Create Trip / Project Tracker</div>
        <form id="tripForm">
          <div class="input-field">
            <input id="tripTitle" type="text" placeholder="e.g. Israel Vacation 2026, NYC Weekend" required>
            <label class="active">Trip or Project Title</label>
          </div>
          <div class="input-field">
            <input id="totalBudget" type="number" step="100" placeholder="3500.00" required>
            <label class="active">Total Planned Budget ($)</label>
          </div>
          <button class="btn waves-effect waves-light btn-submit" type="button" onclick="submitTrip()">
            Create Trip Tracker Sheet 🚀
          </button>
        </form>

        <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
        <script>
          function submitTrip() {
            const title = document.getElementById('tripTitle').value.trim();
            const budget = parseFloat(document.getElementById('totalBudget').value) || 0;

            if (!title || budget <= 0) {
              M.toast({html: 'Please fill in Title and Budget!', classes: 'red'});
              return;
            }

            M.toast({html: 'Building Trip Tracker... ⏳', classes: 'blue'});
            google.script.run
              .withSuccessHandler(function(res) {
                M.toast({html: res, classes: 'green'});
                setTimeout(function() { google.script.host.close(); }, 1200);
              })
              .withFailureHandler(function(err) {
                M.toast({html: 'Error: ' + err.message, classes: 'red'});
              })
              .createTripFromForm({ title: title, budget: budget });
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Create Trip / Project Tracker");
}

/**
 * Backend Handler: Creates the isolated Trip tab from the Trip Modal form.
 */
function createTripFromForm(data) {
  createTravelBudgetTab(data.title, data.budget);
  return `Created 'Trip_${data.title.replace(/[^a-zA-Z0-9]/g, "_")}' tab with $${data.budget.toFixed(2)} budget! ✈️`;
}
