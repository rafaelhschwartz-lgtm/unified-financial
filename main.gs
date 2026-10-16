/**
 *  ===========================================================================
 *  GOOGLE SHEETS FINANCIAL AUTOMATION SYSTEM (UNIFIED MASTER ENGINE - V11)
 *  Standardized under the FAST Modeling Standard and ICAEW Spreadsheet Principles.
 *  
 *  V9 Master Enhancements:
 *    1. Auto-Adding Accounts to Enum: Automatically appends new accounts (e.g. Chase Checking 1234)
 *       to Enum!B:B and updates dropdown validations dynamically.
 *    2. Auto-Adding Categories to Enum: Automatically appends new categories to Enum!A:A and
 *       populates budget historical structures.
 *    3. Auto-Learning Categorization Rules: Automatically learns unlisted merchants and appends
 *       new keyword rules to Category_Rules_DB.
 *    4. Multi-Account Smart Detection: Detects account numbers & types from filenames, CSV headers,
 *       email alerts, and Google Voice SMS forwards.
 *    5. "Quick Add" Toolbar Tools: UI modal prompts to quickly add Transactions, Categories, Accounts,
 *       and Merchant Rules without touching raw sheet cells.
 *    6. Consolidated Trigger System (onOpen, onEdit).
 *    7. Dynamic 10-Year Growth Projection Matrix & Area Chart.
 *    8. Responsive HTML Weekly Summaries & Single-Alert Budget Breach Monitors.
 *    9. Paginated In-Sheet Transaction Editor Sidebar.
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
    .addSubMenu(ui.createMenu("➕ Quick Add Tools")
      .addItem("➕ Quick Add Transaction", "quickAddTransactionPrompt")
      .addItem("➕ Add New Budget Category", "quickAddCategoryPrompt")
      .addItem("➕ Add New Account Source", "quickAddAccountPrompt")
      .addItem("➕ Add Merchant Category Rule", "quickAddRulePrompt"))
    .addSeparator()
    .addItem("Open Budget History Sidebar", "showBudgetSidebar")
    .addItem("Open Transaction Editor Sidebar", "showTransactionEditorSidebar")
    .addItem("Open Categorization Assistant 🤖", "showCategorizerSidebar")
    .addSeparator()
    .addItem("Check Budget Breaches", "checkBudgetBreaches")
    .addItem("Dispatch Weekly Summary Email", "sendWeeklyBudgetSummary")
    .addSeparator()
    .addItem("➕ Create Trip / Project Budget", "promptTripBudget")
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
// 2. DYNAMIC ENUM ENGINE & AUTO-VALIDATION HELPERS
// ===========================================================================

/**
 * Ensures an Account Source exists in Enum!B3:B.
 * If missing, automatically appends it and updates database validation rules.
 */
function ensureAccountInEnum(accountName) {
  if (!accountName) return "Cash";
  const name = accountName.toString().trim();
  if (!name) return "Cash";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  if (!enumSheet) return name;
  
  const lastRow = enumSheet.getLastRow();
  let existingAccounts = [];
  if (lastRow >= 3) {
    existingAccounts = enumSheet.getRange(3, 2, lastRow - 2, 1).getValues().flat().map(a => a.toString().trim().toLowerCase());
  }
  
  if (existingAccounts.indexOf(name.toLowerCase()) === -1) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
      try {
        let bRow = 3;
        while (enumSheet.getRange(bRow, 2).getValue() !== "") {
          bRow++;
        }
        enumSheet.getRange(bRow, 2).setValue(name).setFontColor("#002060");
        console.log(`[Enum Engine] Auto-added new Account Source to Enum tab: "${name}" at B${bRow}`);
        
        refreshEnumValidations(ss);
      } catch (e) {
        console.error("Error auto-adding account to Enum: " + e.toString());
      } finally {
        lock.releaseLock();
      }
    }
  }
  return name;
}

/**
 * Ensures a Budget Category exists in Enum!A3:A.
 * If missing, appends it, adds default historical allocations, and refreshes validations.
 */
function ensureCategoryInEnum(categoryName) {
  if (!categoryName) return "Miscellaneous";
  const cat = categoryName.toString().trim();
  if (!cat || cat === "Miscellaneous") return "Miscellaneous";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enumSheet = ss.getSheetByName("Enum");
  if (!enumSheet) return cat;
  
  const lastRow = enumSheet.getLastRow();
  let existingCategories = [];
  if (lastRow >= 3) {
    existingCategories = enumSheet.getRange(3, 1, lastRow - 2, 1).getValues().flat().map(c => c.toString().trim().toLowerCase());
  }
  
  if (existingCategories.indexOf(cat.toLowerCase()) === -1) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
      try {
        let aRow = 3;
        while (enumSheet.getRange(aRow, 1).getValue() !== "") {
          aRow++;
        }
        enumSheet.getRange(aRow, 1).setValue(cat).setFontColor("#385723").setFontWeight("bold");
        console.log(`[Enum Engine] Auto-added new Budget Category to Enum tab: "${cat}" at A${aRow}`);
        
        const budgetDB = ss.getSheetByName("Budget_History_DB");
        if (budgetDB) {
          const today = new Date();
          const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
          const currMonth = months[today.getMonth()];
          const currYear = today.getFullYear();
          budgetDB.appendRow([currYear, currMonth, cat, 0.00]);
        }
        
        refreshEnumValidations(ss);
      } catch (e) {
        console.error("Error auto-adding category to Enum: " + e.toString());
      } finally {
        lock.releaseLock();
      }
    }
  }
  return cat;
}

/**
 * Refreshes dropdown data validation rules on Transactions_DB and Category_Rules_DB
 * to point dynamically to the full height of Enum!A3:A and Enum!B3:B.
 */
function refreshEnumValidations(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  
  const enumSheet = ss.getSheetByName("Enum");
  const dbSheet = ss.getSheetByName("Transactions_DB");
  const rulesDb = ss.getSheetByName("Category_Rules_DB");
  if (!enumSheet) return;
  
  const lastCatRow = Math.max(3, enumSheet.getLastRow());
  const categoryRange = enumSheet.getRange(3, 1, Math.max(1, lastCatRow - 2), 1);
  const categoryValidation = SpreadsheetApp.newDataValidation().requireValueInRange(categoryRange).setAllowInvalid(true).build();
  
  let bRow = 3;
  while (enumSheet.getRange(bRow, 2).getValue() !== "") {
    bRow++;
  }
  const lastAcctRow = Math.max(3, bRow - 1);
  const accountRange = enumSheet.getRange(3, 2, Math.max(1, lastAcctRow - 2), 1);
  const accountValidation = SpreadsheetApp.newDataValidation().requireValueInRange(accountRange).setAllowInvalid(true).build();
  
  if (dbSheet) {
    dbSheet.getRange("H2:H1000").setDataValidation(categoryValidation);
    dbSheet.getRange("I2:I1000").setDataValidation(accountValidation);
  }
  
  if (rulesDb) {
    rulesDb.getRange("B3:B1000").setDataValidation(categoryValidation);
  }
}

/**
 * Dynamically detects specific accounts (including digits or card types like 'Chase Checking 1234')
 * from filenames, CSV row data, or email/SMS body text.
 */
function detectAccountSource(filename, csvHeader, csvRow, textBody) {
  let detected = "Chase Checking";
  const fn = (filename || "").toLowerCase();
  const text = (textBody || "").toLowerCase();
  
  const digitMatch = fn.match(/(\d{4})/) || text.match(/(?:ending in|account|card|acct)\s*#?\s*(\d{4})/i) || text.match(/#(\d{4})/);
  const acctDigits = digitMatch ? digitMatch[1] : null;
  
  if (fn.includes("sapphire") || text.includes("sapphire")) {
    detected = acctDigits ? `Chase Sapphire ${acctDigits}` : "Chase Sapphire";
  } else if (fn.includes("freedom") || text.includes("freedom")) {
    detected = acctDigits ? `Chase Freedom ${acctDigits}` : "Chase Freedom";
  } else if (fn.includes("checking") || text.includes("checking") || fn.includes("debit")) {
    detected = acctDigits ? `Chase Checking ${acctDigits}` : "Chase Checking";
  } else if (fn.includes("savings") || text.includes("savings")) {
    detected = acctDigits ? `Chase Savings ${acctDigits}` : "Chase Savings";
  } else if (fn.includes("td") || text.includes("td bank")) {
    detected = acctDigits ? `TD Bank ${acctDigits}` : "TD Bank";
  } else if (fn.includes("amex") || text.includes("american express")) {
    detected = acctDigits ? `Amex Gold ${acctDigits}` : "Amex Gold";
  } else if (acctDigits) {
    detected = `Chase Account ${acctDigits}`;
  }
  
  if (csvHeader && csvRow) {
    const colAcct = csvHeader.indexOf("Account") !== -1 ? csvHeader.indexOf("Account") : csvHeader.indexOf("Account Number");
    if (colAcct !== -1 && csvRow[colAcct]) {
      const val = csvRow[colAcct].toString().trim();
      const match = val.match(/(\d{4})$/);
      if (match) {
        detected = `Chase Account ${match[1]}`;
      }
    }
  }
  
  return ensureAccountInEnum(detected);
}

// ===========================================================================
// 3. MAIN SHEET COMPILER (DYNAMIC RANGES & VISUALIZATIONS)
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
  
  function formatHeader(sheet, rangeString, title) {
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
  
  function formatSubHeader(sheet, rangeString) {
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
  
  // ABOUT TAB
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
  
  // ENUM TAB
  const enumSheet = sheets["Enum"];
  formatHeader(enumSheet, "A1:B1", "SYSTEM ENUMS / CATEGORIES & ACCOUNTS");
  enumSheet.getRange("A2:B2").setValues([["Budget Categories", "Account Sources"]]).setFontWeight("bold");
  
  const categories = [
    ["Food & Drink"], ["Groceries"], ["Bills & Utilities"], ["Home"], 
    ["Transportation"], ["Entertainment"], ["Personal"], ["Travel"], ["Miscellaneous"], ["Transfer"]
  ];
  const accounts = [
    ["Cash"], ["Chase Checking 1234"], ["Chase Checking 5678"], ["Chase Sapphire 9012"], ["TD Bank"], ["Credit Card"], ["Savings"]
  ];
  enumSheet.getRange(3, 1, categories.length, 1).setValues(categories);
  enumSheet.getRange(3, 2, accounts.length, 1).setValues(accounts);
  enumSheet.autoResizeColumns(1, 2);
  
  // TRANSACTIONS_DB TAB
  const db = sheets["Transactions_DB"];
  const dbHeaders = [
    "Transaction_ID", "Timestamp", "Transaction_Date", "Merchant_Raw", 
    "Merchant_Clean", "Amount", "Capital_Flow", "Category", 
    "Account_Source", "Ingestion_Channel", "Sync_Status"
  ];
  db.getRange(1, 1, 1, dbHeaders.length).setValues([dbHeaders]);
  formatSubHeader(db, "A1:K1");
  db.setRowHeight(1, 28);
  
  const categoryRange = enumSheet.getRange("A3:A20");
  const categoryValidation = SpreadsheetApp.newDataValidation().requireValueInRange(categoryRange).setAllowInvalid(true).build();
  db.getRange("H2:H1000").setDataValidation(categoryValidation);
  
  const accountRange = enumSheet.getRange("B3:B20");
  const accountValidation = SpreadsheetApp.newDataValidation().requireValueInRange(accountRange).setAllowInvalid(true).build();
  db.getRange("I2:I1000").setDataValidation(accountValidation);
  
  const flowValidation = SpreadsheetApp.newDataValidation().requireValueInList(["INFLOW", "OUTFLOW", "TRANSFER"]).setAllowInvalid(false).build();
  db.getRange("G2:G1000").setDataValidation(flowValidation);
  
  const channelValidation = SpreadsheetApp.newDataValidation().requireValueInList(["WEB_FORM", "CHASE_CSV", "GMAIL_ALERT", "GOOGLE_VOICE", "PLAID_API", "MANUAL", "SAVINGS_SWEEP"]).setAllowInvalid(false).build();
  db.getRange("J2:J1000").setDataValidation(channelValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation().requireValueInList(["SETTLED", "PENDING", "MANUAL_REVIEW"]).setAllowInvalid(false).build();
  db.getRange("K2:K1000").setDataValidation(statusValidation);
  
  db.getRange("A2:K2").setValues([[
    "INIT_SETUP_DUMMY", new Date(), new Date(), "INITIAL SYSTEM SETUP", "Initial Setup", 0.01, "INFLOW", "Miscellaneous", "Cash", "MANUAL", "SETTLED"
  ]]);
  db.getRange("F2:F1000").setNumberFormat("$#,##0.00").setFontColor("#002060");
  db.getRange("C2:C1000").setNumberFormat("yyyy-mm-dd");
  db.autoResizeColumns(1, 11);
  
  // BUDGET_HISTORY_DB TAB
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
  
  // CATEGORY RULES DATABASE TAB
  const rulesDb = sheets["Category_Rules_DB"];
  formatHeader(rulesDb, "A1:C1", "AI RULES-BASED CATEGORIZATION LOOKUP DATABASE");
  rulesDb.getRange("A2:C2").setValues([["Merchant Keyword Rule", "Clean Merchant Name", "Budget Category Target"]]).setFontWeight("bold");
  formatSubHeader(rulesDb, "A2:C2");
  rulesDb.setRowHeight(2, 25);
  
  const startingRules = [
    ["starbucks", "Starbucks", "Food & Drink"],
    ["mcdonald", "McDonald's", "Food & Drink"],
    ["kroger", "Kroger", "Groceries"],
    ["walmart", "Walmart", "Groceries"],
    ["wholefd", "Whole Foods", "Groceries"],
    ["trader joe", "Trader Joe's", "Groceries"],
    ["netflix", "Netflix", "Entertainment"],
    ["spotify", "Spotify", "Entertainment"],
    ["uber", "Uber", "Transportation"],
    ["lyft", "Lyft", "Transportation"],
    ["chevron", "Chevron", "Transportation"],
    ["shell", "Shell", "Transportation"],
    ["comcast", "Comcast", "Bills & Utilities"],
    ["verizon", "Verizon", "Bills & Utilities"],
    ["amazon", "Amazon", "Home"],
    ["target", "Target", "Home"],
    ["transfer", "Internal Transfer", "Transfer"]
  ];
  rulesDb.getRange(3, 1, startingRules.length, 3).setValues(startingRules);
  rulesDb.getRange("C3:C1000").setDataValidation(categoryValidation);
  rulesDb.getRange("A3:B1000").setFontColor("#002060");
  rulesDb.autoResizeColumns(1, 3);
  
  // SAVINGS GOALS TAB & MULTI-YEAR PROJECTION MODEL
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
  
  // THE MONTHLY DASHBOARD TAB
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
  
  // CONTROL TAB
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
  const ruleOK = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("OK").setBackground("#D9EAD3").setRanges([range]).build();
  const ruleError = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith("ERROR").setBackground("#F4CCCC").setRanges([range]).build();
  const rules = controlSheet.getConditionalFormatRules();
  rules.push(ruleOK, ruleError);
  controlSheet.setConditionalFormatRules(rules);
  console.log("[Format Engine] Applied conditional formatting rules successfully to Control!D4:D6.");
}

// ===========================================================================
// 4. QUICK ADD TOOLBAR PROMPTS
// ===========================================================================

function quickAddTransactionPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  const mRes = ui.prompt("➕ Quick Add Transaction", "Enter Merchant Name (e.g. Starbucks, Kroger):", ui.ButtonSet.OK_CANCEL);
  if (mRes.getSelectedButton() !== ui.Button.OK || !mRes.getResponseText().trim()) return;
  const merchant = mRes.getResponseText().trim();
  
  const aRes = ui.prompt("➕ Quick Add Transaction", `Enter Amount for "${merchant}" (e.g. 15.50):`, ui.ButtonSet.OK_CANCEL);
  if (aRes.getSelectedButton() !== ui.Button.OK || !aRes.getResponseText().trim()) return;
  const amount = Math.abs(parseFloat(aRes.getResponseText().trim()) || 0);
  
  const cRes = ui.prompt("➕ Quick Add Transaction", `Enter Category (or leave blank to auto-categorize):`, ui.ButtonSet.OK_CANCEL);
  let category = cRes.getSelectedButton() === ui.Button.OK ? cRes.getResponseText().trim() : "";
  if (!category) {
    category = autoCategorizeMerchant(merchant);
  } else {
    ensureCategoryInEnum(category);
  }
  
  const acRes = ui.prompt("➕ Quick Add Transaction", `Enter Account Source (e.g. Chase Checking 1234, Cash):`, ui.ButtonSet.OK_CANCEL);
  let account = acRes.getSelectedButton() === ui.Button.OK ? acRes.getResponseText().trim() : "Cash";
  if (!account) account = "Cash";
  ensureAccountInEnum(account);
  
  addOrUpdateTransaction({
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    merchantRaw: merchant,
    merchantClean: autoCleanMerchantName(merchant),
    amount: amount,
    flow: "OUTFLOW",
    category: category,
    accountSource: account,
    channel: "MANUAL",
    status: "SETTLED"
  });
  
  ui.alert("Success! 🎉", `Transaction logged successfully:

Merchant: ${merchant}
Amount: $${amount.toFixed(2)}
Category: ${category}
Account: ${account}`, ui.ButtonSet.OK);
}

function quickAddCategoryPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("➕ Add New Budget Category", "Enter New Category Name (e.g. Subscriptions, Pet Care):", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim()) {
    const cat = res.getResponseText().trim();
    ensureCategoryInEnum(cat);
    ui.alert("Category Added! 🎉", `"${cat}" has been added to your Enum sheet and dropdown validations.`, ui.ButtonSet.OK);
  }
}

function quickAddAccountPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("➕ Add New Account Source", "Enter New Account Name (e.g. Chase Checking 5678, Amex Gold):", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim()) {
    const acct = res.getResponseText().trim();
    ensureAccountInEnum(acct);
    ui.alert("Account Added! 🎉", `"${acct}" has been added to your Enum sheet and dropdown validations.`, ui.ButtonSet.OK);
  }
}

function quickAddRulePrompt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  if (!rulesSheet) {
    ui.alert("Error ❌", "Category_Rules_DB sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  const kwRes = ui.prompt("➕ Add Merchant Rule", "Enter Merchant Keyword Rule (e.g. spotify, trader joe):", ui.ButtonSet.OK_CANCEL);
  if (kwRes.getSelectedButton() !== ui.Button.OK || !kwRes.getResponseText().trim()) return;
  const kw = kwRes.getResponseText().trim().toLowerCase();
  
  const clRes = ui.prompt("➕ Add Merchant Rule", `Enter Clean Merchant Name for "${kw}":`, ui.ButtonSet.OK_CANCEL);
  const cleanName = clRes.getSelectedButton() === ui.Button.OK ? clRes.getResponseText().trim() : kw;
  
  const catRes = ui.prompt("➕ Add Merchant Rule", `Enter Target Category for "${kw}" (e.g. Entertainment):`, ui.ButtonSet.OK_CANCEL);
  if (catRes.getSelectedButton() !== ui.Button.OK || !catRes.getResponseText().trim()) return;
  const cat = catRes.getResponseText().trim();
  
  ensureCategoryInEnum(cat);
  rulesSheet.appendRow([kw, cleanName, cat]);
  ui.alert("Rule Saved! 🤖", `Rule created:
Keyword: "${kw}"
Clean Name: "${cleanName}"
Category: "${cat}"`, ui.ButtonSet.OK);
}

// ===========================================================================
// 5. BACKGROUND SYNCHRONIZER & SIDEBAR ENGINES
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
  console.log(`[Budget Sync Engine] Synchronized ${numCategories} category budgets for Period: ${year} ${month}`);
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

function showBudgetSidebar() {
  const html = HtmlService.createTemplateFromFile("SidebarView");
  const htmlOutput = html.evaluate()
    .setTitle("Historical Budget Allocator")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  console.log("[UI Service] Rendered Budget History Allocation Sidebar to view.");
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
      amount: match ? match[3] : 0.00
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
// 6. EXISTING AUTOMATION CHANNELS (Consolidated & Auto-Learning)
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
    
    const rawMerchant = rowData.Merchant || "Manual Entry";
    const cleanMerchant = autoCleanMerchantName(rawMerchant);
    const suggestedCategory = rowData.Category && rowData.Category !== "" ? rowData.Category : autoCategorizeMerchant(cleanMerchant);
    const accountSource = ensureAccountInEnum(rowData.Source || "Cash");
    ensureCategoryInEnum(suggestedCategory);
    
    sheet.appendRow([
      transactionId,
      timestamp,
      dateValue,
      rawMerchant,
      cleanMerchant,
      Math.abs(parseFloat(rowData.Amount)),
      rowData.Flow || "OUTFLOW",
      suggestedCategory,
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
  
  // Channel 1: Gmail Chase Alerts
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
      
      let merchantRaw = "CHASE TRANSACTION ALERT";
      const merchMatch = body.match(/at\s+([^\n\r,.]+?)\s+on/i) || body.match(/charge\s+of\s+.*?\s+at\s+([^\n\r,.]+)/i);
      if (merchMatch) {
        merchantRaw = merchMatch[1].trim();
      }
      
      const cleanMerchant = autoCleanMerchantName(merchantRaw);
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
      const accountSource = detectAccountSource("Chase Alert", null, null, body);
      
      dbSheet.appendRow([
        msgId,
        new Date(),
        date,
        merchantRaw,
        cleanMerchant,
        amount,
        "OUTFLOW",
        suggestedCategory,
        accountSource,
        "GMAIL_ALERT",
        "SETTLED"
      ]);
      msg.markRead();
      chaseCount++;
      existingIDs.push(msgId);
    }
  }
  
  // Channel 2: Google Voice SMS Alerts
  const gvoiceQuery = 'from:voice-noreply@google.com "transaction" is:unread';
  const gvoiceThreads = GmailApp.search(gvoiceQuery, 0, 10);
  let gvoiceCount = 0;
  
  for (let i = 0; i < gvoiceThreads.length; i++) {
    const messages = gvoiceThreads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const msgId = msg.getId();
      
      if (existingIDs.indexOf(msgId) !== -1) continue;
      
      const body = msg.getPlainBody();
      const amountMatch = body.match(/\$(\d+\.\d{2})/);
      if (!amountMatch) continue;
      
      const amount = parseFloat(amountMatch[1]);
      const date = msg.getDate();
      
      let flow = "OUTFLOW";
      if (/received|paid you|credited|deposit/i.test(body)) {
        flow = "INFLOW";
      }
      
      let merchantRaw = "GOOGLE VOICE SMS ALERT";
      const merchMatch = body.match(/at\s+([^\n\r,.]+)/i) || body.match(/from\s+([^\n\r,.]+)/i);
      if (merchMatch) {
        merchantRaw = merchMatch[1].trim();
      }
      
      const cleanMerchant = autoCleanMerchantName(merchantRaw);
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
      const accountSource = detectAccountSource("Google Voice", null, null, body);
      
      dbSheet.appendRow([
        msgId,
        new Date(),
        date,
        merchantRaw,
        cleanMerchant,
        amount,
        flow,
        suggestedCategory,
        accountSource,
        "GOOGLE_VOICE",
        "SETTLED"
      ]);
      msg.markRead();
      gvoiceCount++;
      existingIDs.push(msgId);
    }
  }
  
  console.log(`[Gmail Scraper Engine] Ingested ${chaseCount} Chase Alerts, ${gvoiceCount} Google Voice SMS Alerts.`);
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
    throw new Error("Could not map Chase CSV headers automatically.");
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
    
    const merchantRaw = row[colMerchant];
    const cleanMerchant = autoCleanMerchantName(merchantRaw);
    const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
    const accountSource = detectAccountSource(file.getName(), csvHeaders, row, null);
    
    dbSheet.appendRow([
      sigId,
      new Date(),
      new Date(row[colDate]),
      merchantRaw,
      cleanMerchant,
      Math.abs(rawAmount),
      rawAmount < 0 ? "OUTFLOW" : "INFLOW",
      suggestedCategory,
      accountSource,
      "CHASE_CSV",
      "SETTLED"
    ]);
    newRowsCount++;
  }
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
      "Created 'Chase_CSV_Uploads' folder in Google Drive.

" +
      "1. Drag and drop bank CSV files into 'Chase_CSV_Uploads'.
" +
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

      let fileRowsAdded = 0;
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        if (!row[colDate]) continue;

        const rawAmountStr = row[colAmount].toString().replace(/'/g, "").trim();
        const rawAmount = parseFloat(rawAmountStr);
        if (isNaN(rawAmount)) continue;

        const sigId = "CHASE-CSV-" + row[colDate] + "-" + row[colMerchant].substring(0, 10).replace(/\s/g, "") + "-" + Math.abs(rawAmount).toFixed(2);
        if (existingIDs.indexOf(sigId) !== -1) continue;

        const merchantRaw = row[colMerchant];
        const cleanMerchant = autoCleanMerchantName(merchantRaw);
        const suggestedCategory = autoCategorizeMerchant(cleanMerchant);
        const accountSource = detectAccountSource(file.getName(), csvHeaders, row, null);

        dbSheet.appendRow([
          sigId,
          new Date(),
          new Date(row[colDate]),
          merchantRaw,
          cleanMerchant,
          Math.abs(rawAmount),
          rawAmount < 0 ? "OUTFLOW" : "INFLOW",
          suggestedCategory,
          accountSource,
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
      console.error(`Error processing file ${file.getName()}: ${err.toString()}`);
    }
  }

  if (filesProcessed === 0) {
    ui.alert("Process Status ℹ️", "No pending .csv files found inside 'Chase_CSV_Uploads'.", ui.ButtonSet.OK);
  } else {
    ui.alert(
      "Batch Import Complete! 🎉",
      `Successfully processed ${filesProcessed} CSV statements.
Added ${totalNewRows} new transactions.
Processed files archived.`,
      ui.ButtonSet.OK
    );
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
// 7. AUTO-CATEGORIZATION & AUTO-LEARNING ENGINE
// ===========================================================================

function autoCategorizeMerchant(merchantRaw) {
  if (!merchantRaw) return "Miscellaneous";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  const cleanRaw = merchantRaw.toString().trim();
  
  const keyword = cleanRaw.toLowerCase()
                          .replace(/[0-9]+/g, '')
                          .replace(/#\w+/g, '')
                          .replace(/\s+/g, ' ')
                          .trim();

  if (rulesSheet && rulesSheet.getLastRow() >= 3) {
    const rulesData = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 3).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      const ruleKw = rulesData[i][0].toString().toLowerCase().trim();
      const ruleCat = rulesData[i][2].toString().trim();
      if (ruleKw !== "" && keyword.indexOf(ruleKw) !== -1) {
        ensureCategoryInEnum(ruleCat);
        return ruleCat;
      }
    }
  }

  let suggestedCat = "Miscellaneous";
  if (/starbucks|mcdonald|dunkin|restaurant|cafe|grubhub|ubereats|doordash|pizza|deli|bakery|burger|taco/i.test(keyword)) {
    suggestedCat = "Food & Drink";
  } else if (/kroger|walmart|wholefd|safeway|supermarket|trader joe|aldi|costco|grocer|food market/i.test(keyword)) {
    suggestedCat = "Groceries";
  } else if (/netflix|spotify|hulu|youtube|disney|steam|cinema|theatre|ticketmaster/i.test(keyword)) {
    suggestedCat = "Entertainment";
  } else if (/uber|lyft|chevron|shell|bp|subway|transit|gas station|parking|toll|exxon|mobil/i.test(keyword)) {
    suggestedCat = "Transportation";
  } else if (/comcast|electric|att|verizon|water|utility|insurance|power|energy|mobile/i.test(keyword)) {
    suggestedCat = "Bills & Utilities";
  } else if (/amazon|target|home depot|ikea|lowes|furniture|hardware/i.test(keyword)) {
    suggestedCat = "Home";
  } else if (/delta|airbnb|hotel|expedia|flight|united airlines|travel|booking/i.test(keyword)) {
    suggestedCat = "Travel";
  } else if (/nike|macys|ulta|barber|sephora|salon|clothing|pharmacy|cvs|walgreens/i.test(keyword)) {
    suggestedCat = "Personal";
  } else if (/transfer|payment|zelle|venmo|credit card payment/i.test(keyword)) {
    suggestedCat = "Transfer";
  }

  const cleanTitle = keyword.split(' ')
                            .map(w => w.charAt(0).toUpperCase() + w.slice(1))
                            .join(' ');

  if (rulesSheet && keyword.length >= 3) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
      try {
        rulesSheet.appendRow([keyword, cleanTitle || cleanRaw, suggestedCat]);
        console.log(`[Auto-Rules Engine] Auto-added learned merchant rule to Category_Rules_DB: "${keyword}" -> "${suggestedCat}"`);
      } catch (err) {
        console.error("Failed to auto-append rule to Category_Rules_DB: " + err.toString());
      } finally {
        lock.releaseLock();
      }
    }
  }

  ensureCategoryInEnum(suggestedCat);
  return suggestedCat;
}

function autoCleanMerchantName(merchantRaw) {
  if (!merchantRaw) return "Unknown Merchant";
  const merchant = merchantRaw.toLowerCase().trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  
  if (rulesSheet && rulesSheet.getLastRow() >= 3) {
    const rulesData = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 2).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      const keyword = rulesData[i][0].toString().toLowerCase().trim();
      const cleanName = rulesData[i][1].toString().trim();
      if (keyword !== "" && merchant.indexOf(keyword) !== -1) {
        return cleanName;
      }
    }
  }
  
  let cleaned = merchantRaw.replace(/[0-9]+/g, "").replace(/\s+/g, " ").trim();
  return cleaned.replace(/\b\w/g, c => c.toUpperCase());
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
    const cleanMerchant = data[i][4];
    
    if (!category || category === "" || category === "Miscellaneous") {
      const suggestedCategory = autoCategorizeMerchant(cleanMerchant || rawMerchant);
      if (suggestedCategory !== "Miscellaneous") {
        dbSheet.getRange(rowIdx, 8).setValue(suggestedCategory);
        updatedCount++;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert("Auto-Categorization Complete 🤖", "Successfully updated " + updatedCount + " transactions based on rules-based merchant patterns.", SpreadsheetApp.getUi().ButtonSet.OK);
}

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
  
  const dataRange = dash.getRange(10, 1, enumCategoriesCount, 4);
  const data = dataRange.getValues();
  
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
    emailBody += `<p>Spending limits breached on your Interactive Financial Dashboard:</p>`;
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
    emailBody += `<p>Please review your transactions on <b>'The Monthly Dashboard Tab'</b>.</p>`;
    emailBody += `<br><p><i>Sent automatically by your Google Sheets Financial Automation System 🚀</i></p>`;
    
    MailApp.sendEmail({
      to: userEmail,
      subject: `⚠️ Budget Breach Notification: ${breaches.length} Category Alert(s)`,
      htmlBody: emailBody
    });
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
      <h2 style="margin: 0; font-size: 22px;">WEEKLY FINANCIAL BRIEFING</h2>
      <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Period: ${month} ${year} (${monthElapsedPercent}% Elapsed)</p>
    </div>

    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px;">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Month-to-Date KPIs</h3>
      <table style="width: 100%; text-align: center; border-collapse: collapse;">
        <tr>
          <td style="padding: 10px; width: 25%;">
            <span style="font-size: 11px; color: #666; display: block;">INCOME</span>
            <span style="font-weight: bold; font-size: 16px; color: #385723;">$${totalIncome.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block;">SPENDING</span>
            <span style="font-weight: bold; font-size: 16px; color: #db4437;">$${totalSpending.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block;">SAVINGS</span>
            <span style="font-weight: bold; font-size: 16px; color: ${netSavings >= 0 ? '#385723' : '#c00000'};">$${netSavings.toFixed(2)}</span>
          </td>
          <td style="padding: 10px; width: 25%; border-left: 1px solid #e1e4e8;">
            <span style="font-size: 11px; color: #666; display: block;">RATE</span>
            <span style="font-weight: bold; font-size: 16px; color: #1f4e78;">${savingsRate}%</span>
          </td>
        </tr>
      </table>
    </div>

    <div style="background-color: white; padding: 20px; margin-top: 15px; border-radius: 5px;">
      <h3 style="color: #1f4e78; margin-top: 0; border-bottom: 2px solid #d9e1f2; padding-bottom: 8px;">Budget Performance</h3>
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
    
    let statusText = "✅ On Track";
    let statusColor = "#0f9d58";

    if (planned === 0) {
      statusText = "Unplanned";
      statusColor = "#777777";
    } else if (actual > planned) {
      statusText = "⚠️ Overrun";
      statusColor = "#db4437";
    } else if (progressPercent > monthElapsedPercent) {
      statusText = "⚡ Fast Velocity";
      statusColor = "#ff9900";
    }

    html += `
      <tr style="border-bottom: 1px solid #f0f0f0;">
        <td style="padding: 8px; font-weight: bold;">${category}</td>
        <td style="padding: 8px; text-align: right;">$${planned.toFixed(2)}</td>
        <td style="padding: 8px; text-align: right;">$${actual.toFixed(2)}</td>
        <td style="padding: 8px; text-align: right; font-weight: bold; color: ${statusColor};">${statusText}</td>
      </tr>`;
  });

  html += `
        </tbody>
      </table>
    </div>

    <div style="margin-top: 20px; text-align: center; color: #666; font-size: 12px;">
      <p style="margin-top: 10px;"><i>Dispatched automatically by your <strong>Google Sheets Financial Automation System 🚀</strong></i></p>
    </div>
  </div>`;

  MailApp.sendEmail({
    to: userEmail,
    subject: `📊 Weekly Financial Summary: ${month} ${year} Briefing`,
    htmlBody: html
  });
}

function promptSavingsSweep() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("The Monthly Dashboard Tab");
  if (!dash) {
    ui.alert("Error ❌", "Dashboard sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  const month = dash.getRange("D4").getValue().toString().trim();
  const year = dash.getRange("C4").getValue().toString().trim();
  
  if (month === "All") {
    ui.alert("Invalid Selection ⚠️", "Please select a specific month before executing the savings sweep.", ui.ButtonSet.OK);
    return;
  }
  
  const netSavings = parseFloat(dash.getRange("C7").getValue() || 0);
  if (netSavings <= 0) {
    ui.alert("Zero Surplus ⚠️", `Monthly Net Savings for ${month} ${year} is $${netSavings.toFixed(2)}. Sweep requires a positive surplus.`, ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.alert(
    "Execute Monthly Savings Sweep? 💰",
    `Monthly Net Surplus for ${month} ${year} is $${netSavings.toFixed(2)}.

Execute relational double-entry sweep to allocated Savings Goals?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      const summary = executeSavingsSweep(ss, year, month, netSavings);
      ui.alert("Sweep Complete! 🎉", summary, ui.ButtonSet.OK);
    } catch (err) {
      ui.alert("Sweep Failed ❌", err.toString(), ui.ButtonSet.OK);
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
    const confirmReRun = ui.alert(
      "Sweep Previously Executed ⚠️",
      `A savings sweep for ${month} ${year} was already run previously.
Do you want to re-run it?`,
      ui.ButtonSet.YES_NO
    );
    if (confirmReRun !== ui.Button.YES) {
      return "Sweep cancelled by user.";
    }
  }
  
  const lastGoalRow = goalsSheet.getLastRow();
  if (lastGoalRow < 3) throw new Error("No goals found on Savings_Goals tab.");
  
  const data = goalsSheet.getRange(3, 1, lastGoalRow - 2, 5).getValues();
  let totalPercentage = 0;
  for (let i = 0; i < data.length; i++) {
    totalPercentage += parseFloat(data[i][4]) || 0;
  }
  
  if (totalPercentage > 1.001) {
    throw new Error("Total sweep allocation percentages exceed 100%. Please correct allocations on Savings_Goals tab.");
  }
  
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthIdx = months.indexOf(month);
  const sweepDate = new Date(parseInt(year, 10), monthIdx + 1, 0);
  
  let logSummary = `Successfully swept $${netSavings.toFixed(2)} surplus for ${month} ${year}:`;
  let entriesAdded = 0;
  
  for (let i = 0; i < data.length; i++) {
    const goalName = data[i][0];
    const sweepPct = parseFloat(data[i][4]) || 0;
    
    if (sweepPct <= 0) continue;
    
    const allocationAmount = netSavings * sweepPct;
    const currentBal = parseFloat(data[i][2]) || 0;
    goalsSheet.getRange(3 + i, 3).setValue(currentBal + allocationAmount);
    
    const timeId = new Date().getTime() + i;
    
    // Double-entry 1: Debit Outflow
    dbSheet.appendRow([
      "SWEEP-OUT-" + year + "-" + month + "-" + i,
      new Date(),
      sweepDate,
      "Monthly Savings Sweep: " + goalName,
      "Savings Sweep (" + goalName + ")",
      allocationAmount,
      "OUTFLOW",
      "Savings",
      "Chase Checking 1234",
      "SAVINGS_SWEEP",
      "SETTLED"
    ]);
    
    // Double-entry 2: Credit Inflow
    dbSheet.appendRow([
      "SWEEP-IN-" + year + "-" + month + "-" + i,
      new Date(),
      sweepDate,
      "Goal Allocation: " + goalName,
      "Goal Deposit (" + goalName + ")",
      allocationAmount,
      "INFLOW",
      "Savings",
      "Savings",
      "SAVINGS_SWEEP",
      "SETTLED"
    ]);
    
    entriesAdded += 2;
    logSummary += `
• ${goalName}: $${allocationAmount.toFixed(2)} (${(sweepPct * 100).toFixed(0)}%)`;
  }
  
  scriptProperties.setProperty(sweepKey, "true");
  return logSummary;
}

function formatCurrency(amount) {
  return "$" + parseFloat(amount).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

// ===========================================================================
// 8. HIGH-PERFORMANCE DYNAMIC TRANSACTION EDITOR SIDEBAR
// ===========================================================================


// ===========================================================================
// 7.5 CHECKS, ZELLE, WIRE & SPECIAL TRANSACTION CLASSIFICATION ENGINE
// ===========================================================================

/**
 * Enhanced Auto-Cleanser for Checks, Zelle/QuickPay, Wire Transfers, and ACH.
 */
function autoCleanMerchantName(merchantRaw) {
  if (!merchantRaw) return "Unknown Merchant";
  const raw = merchantRaw.toString().trim();
  const lower = raw.toLowerCase();

  // 1. Check Rule in Category_Rules_DB
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");
  if (rulesSheet && rulesSheet.getLastRow() >= 3) {
    const rulesData = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 2).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      const keyword = rulesData[i][0].toString().toLowerCase().trim();
      const cleanName = rulesData[i][1].toString().trim();
      if (keyword !== "" && lower.indexOf(keyword) !== -1) {
        return cleanName;
      }
    }
  }

  // 2. Pattern Matching for Special Transaction Types
  // A. Check Detection
  const checkMatch = raw.match(/check\s*#?\s*(\d+)/i) || raw.match(/chk\s*#?\s*(\d+)/i);
  if (checkMatch) {
    return "Check #" + checkMatch[1];
  }
  if (/^check|^chk|check paid/i.test(lower)) {
    return "Check Payment";
  }

  // B. QuickPay / Zelle Detection
  const zelleMatch = raw.match(/(?:zelle|quickpay|qp)\s+(?:to|from|with)?\s*([a-z0-9\s\.\-']+)/i);
  if (zelleMatch) {
    let person = zelleMatch[1].replace(/ref#?\w+/i, '').replace(/std\s*add/i, '').trim();
    if (person) {
      person = person.replace(/\w/g, c => c.toUpperCase());
      return "Zelle: " + person;
    }
    return "Zelle Transfer";
  }

  // C. Wire Transfer Detection
  const wireMatch = raw.match(/wire\s*(?:transfer|out|in|trans)?\s*(?:ref#?\s*(\w+))?/i);
  if (wireMatch) {
    const ref = wireMatch[1] ? " (" + wireMatch[1] + ")" : "";
    return "Wire Transfer" + ref;
  }

  // D. ACH / Direct Deposit
  if (/direct\s*dep|payroll|ach\s*credit/i.test(lower)) {
    return "Direct Deposit / ACH";
  }
  if (/ach\s*debit|auto\s*pmt|online\s*pmt/i.test(lower)) {
    return "ACH Electronic Payment";
  }

  // 3. Fallback General Cleansing
  let cleaned = raw.replace(/[0-9]{5,}/g, "").replace(/\s+/g, " ").trim();
  return cleaned.replace(/\w/g, c => c.toUpperCase());
}

/**
 * Enhanced Auto-Categorizer with support for Checks, Zelle, Wires, and ACH.
 */
function autoCategorizeMerchant(merchantRaw) {
  if (!merchantRaw) return "Miscellaneous";
  const raw = merchantRaw.toString().trim();
  const lower = raw.toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");

  // 1. Evaluate Category_Rules_DB First (Single Source of Truth)
  if (rulesSheet && rulesSheet.getLastRow() >= 3) {
    const rulesData = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 3).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      const keyword = rulesData[i][0].toString().toLowerCase().trim();
      const categoryVal = rulesData[i][2].toString().trim();
      if (keyword !== "" && lower.indexOf(keyword) !== -1) {
        return categoryVal;
      }
    }
  }

  // 2. Special Transaction Default Routing Rules
  if (/check|chk/i.test(lower)) {
    return "Checks & Banking";
  }
  if (/zelle|quickpay|qp/i.test(lower)) {
    return "Personal";
  }
  if (/wire\s*transfer|wire/i.test(lower)) {
    return "Transfer";
  }
  if (/direct\s*dep|payroll/i.test(lower)) {
    return "Income";
  }

  // 3. Standard Keyword Fallback Rules
  if (/starbucks|mcdonald|dunkin|restaurant|cafe|grubhub|ubereats|doordash|pizza|deli/i.test(lower)) {
    return "Food & Drink";
  }
  if (/kroger|walmart|wholefd|safeway|supermarket|trader joe|aldi|costco/i.test(lower)) {
    return "Groceries";
  }
  if (/netflix|spotify|hulu|youtube|disney|steam|cinema/i.test(lower)) {
    return "Entertainment";
  }
  if (/uber|lyft|chevron|shell|bp|subway|transit|gas station|parking|toll/i.test(lower)) {
    return "Transportation";
  }
  if (/comcast|electric|att|verizon|water|utility|insurance|power/i.test(lower)) {
    return "Bills & Utilities";
  }
  if (/amazon|target|home depot|ikea|lowes/i.test(lower)) {
    return "Home";
  }

  return "Miscellaneous";
}

// ===========================================================================
// 9.5 CATEGORIZATION ASSISTANT SIDEBAR ENGINE
// ===========================================================================

function showCategorizerSidebar() {
  const html = HtmlService.createTemplateFromFile("CategorizerView");
  const htmlOutput = html.evaluate()
    .setTitle("Categorization Assistant 🤖")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Sidebar Endpoint: Returns all uncategorized or Miscellaneous transactions.
 */
function getUncategorizedQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) return [];

  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return [];

  const rawValues = dbSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const queue = [];

  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];
    const category = (row[7] || "").toString().trim();
    const rawMerchant = (row[3] || "").toString().trim();
    const cleanMerchant = (row[4] || "").toString().trim();
    const amount = parseFloat(row[5] || 0);

    if (!category || category === "" || category === "Miscellaneous" || category === "Uncategorized") {
      queue.push({
        rowIdx: i + 2,
        txId: row[0],
        date: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        rawMerchant: rawMerchant,
        cleanMerchant: cleanMerchant || autoCleanMerchantName(rawMerchant),
        amount: amount,
        flow: row[6] || "OUTFLOW",
        category: category || "Miscellaneous",
        source: row[8] || ""
      });
    }
  }

  return queue;
}

/**
 * Sidebar Endpoint: Applies category to a transaction and optionally saves rule to Category_Rules_DB.
 */
function applyCategoryWithRule(rowIdx, selectedCategory, keywordRule, cleanNameRule) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  const rulesSheet = ss.getSheetByName("Category_Rules_DB");

  if (!dbSheet) throw new Error("Transactions_DB not found.");

  // Ensure category is registered in Enum
  ensureCategoryInEnum(selectedCategory);

  // 1. Update the single targeted transaction row
  dbSheet.getRange(rowIdx, 8).setValue(selectedCategory);
  if (cleanNameRule) {
    dbSheet.getRange(rowIdx, 5).setValue(cleanNameRule);
  }

  // 2. If user opted to save a rule, record in Category_Rules_DB and batch update existing matches
  let batchUpdatedCount = 1;
  if (keywordRule && keywordRule.trim() !== "") {
    const kw = keywordRule.toLowerCase().trim();
    
    if (rulesSheet) {
      // Append to Category_Rules_DB if not already present
      let ruleExists = false;
      if (rulesSheet.getLastRow() >= 3) {
        const existingKW = rulesSheet.getRange(3, 1, rulesSheet.getLastRow() - 2, 1).getValues().flat().map(k => k.toString().toLowerCase().trim());
        if (existingKW.includes(kw)) ruleExists = true;
      }
      if (!ruleExists) {
        rulesSheet.appendRow([kw, cleanNameRule || keywordRule, selectedCategory]);
      }
    }

    // Batch update all other matching Miscellaneous / Uncategorized rows in Transactions_DB!
    const lastRow = dbSheet.getLastRow();
    if (lastRow >= 2) {
      const dbData = dbSheet.getRange(2, 1, lastRow - 1, 11).getValues();
      for (let r = 0; r < dbData.length; r++) {
        const currRowIdx = r + 2;
        if (currRowIdx === rowIdx) continue;

        const currCategory = (dbData[r][7] || "").toString().trim();
        const currRaw = (dbData[r][3] || "").toString().toLowerCase();

        if ((!currCategory || currCategory === "Miscellaneous" || currCategory === "Uncategorized") && currRaw.indexOf(kw) !== -1) {
          dbSheet.getRange(currRowIdx, 8).setValue(selectedCategory);
          if (cleanNameRule) dbSheet.getRange(currRowIdx, 5).setValue(cleanNameRule);
          batchUpdatedCount++;
        }
      }
    }
  }

  return { success: true, updatedCount: batchUpdatedCount };
}


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
      txId: row[0],
      timestamp: row[1],
      date: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
      merchantRaw: row[3],
      merchantClean: row[4],
      amount: parseFloat(row[5]) || 0,
      flow: row[6],
      category: row[7],
      accountSource: row[8],
      channel: row[9],
      status: row[10]
    };
  }).reverse();
  
  const filtered = mapped.filter(item => {
    if (!query) return true;
    return (
      item.merchantRaw.toString().toLowerCase().includes(query) ||
      item.merchantClean.toString().toLowerCase().includes(query) ||
      item.category.toString().toLowerCase().includes(query) ||
      item.accountSource.toString().toLowerCase().includes(query) ||
      item.txId.toString().toLowerCase().includes(query)
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
  
  return {
    categories: categories,
    accounts: accounts
  };
}

function addOrUpdateTransaction(txData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) throw new Error("Transactions_DB sheet not found.");
  
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const rawMerchant = txData.merchantRaw || "Manual Entry";
    const cleanMerchant = txData.merchantClean || autoCleanMerchantName(rawMerchant);
    const category = txData.category ? ensureCategoryInEnum(txData.category) : autoCategorizeMerchant(cleanMerchant);
    const accountSource = ensureAccountInEnum(txData.accountSource || "Cash");
    const dateVal = txData.date ? new Date(txData.date) : new Date();
    const amountVal = Math.abs(parseFloat(txData.amount) || 0);
    const flowVal = txData.flow || "OUTFLOW";
    const channelVal = txData.channel || "MANUAL";
    const statusVal = txData.status || "SETTLED";
    
    if (txData.rowIdx && parseInt(txData.rowIdx, 10) >= 2) {
      const rowIdx = parseInt(txData.rowIdx, 10);
      dbSheet.getRange(rowIdx, 3).setValue(dateVal);
      dbSheet.getRange(rowIdx, 4).setValue(rawMerchant);
      dbSheet.getRange(rowIdx, 5).setValue(cleanMerchant);
      dbSheet.getRange(rowIdx, 6).setValue(amountVal);
      dbSheet.getRange(rowIdx, 7).setValue(flowVal);
      dbSheet.getRange(rowIdx, 8).setValue(category);
      dbSheet.getRange(rowIdx, 9).setValue(accountSource);
      dbSheet.getRange(rowIdx, 10).setValue(channelVal);
      dbSheet.getRange(rowIdx, 11).setValue(statusVal);
      return "UPDATED";
    } else {
      const txId = "MANUAL-" + new Date().getTime();
      dbSheet.appendRow([
        txId,
        new Date(),
        dateVal,
        rawMerchant,
        cleanMerchant,
        amountVal,
        flowVal,
        category,
        accountSource,
        channelVal,
        statusVal
      ]);
      return "CREATED";
    }
  } finally {
    lock.releaseLock();
  }
}

function deleteTransactionRow(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Transactions_DB");
  if (!dbSheet) throw new Error("Transactions_DB sheet not found.");
  
  const row = parseInt(rowIdx, 10);
  if (isNaN(row) || row < 2) throw new Error("Invalid row index for deletion.");
  
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    dbSheet.deleteRow(row);
    return "DELETED";
  } finally {
    lock.releaseLock();
  }
}


// ===========================================================================
// 10. ISOLATED TRIP & PROJECT EXPENSE TRACKER ENGINE
// ===========================================================================

/**
 * Prompt user to dynamically generate a dedicated Travel/Project Budget Tab.
 */
function promptTripBudget() {
  const ui = SpreadsheetApp.getUi();
  const titleResp = ui.prompt("Create Trip / Project Budget", "Enter a title for this trip or project (e.g. NYC Trip 2026, Israel Vacation, Kitchen Remodel):", ui.ButtonSet.OK_CANCEL);
  if (titleResp.getSelectedButton() !== ui.Button.OK) return;
  
  const tripTitle = titleResp.getResponseText().trim();
  if (!tripTitle) {
    ui.alert("Error ❌", "Trip/Project title cannot be blank.", ui.ButtonSet.OK);
    return;
  }
  
  const budgetResp = ui.prompt("Set Total Target Budget", "Enter the total overall planned budget amount ($):", ui.ButtonSet.OK_CANCEL);
  let totalBudget = 3000.00;
  if (budgetResp.getSelectedButton() === ui.Button.OK) {
    const val = parseFloat(budgetResp.getResponseText().replace(/[^0-9.]/g, ''));
    if (!isNaN(val) && val > 0) totalBudget = val;
  }
  
  createTravelBudgetTab(tripTitle, totalBudget);
  ui.alert("Trip Budget Created! ✈️", `Successfully generated isolated budget sheet '${tripTitle}' with a $${totalBudget.toFixed(2)} target limit.`, ui.ButtonSet.OK);
}

/**
 * Programmatically builds and formats a dedicated 'Trip_Budget' sheet.
 * Follows FAST Standard & ICAEW Spreadsheet Principles.
 */
function createTravelBudgetTab(tripTitle, totalBudget) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanSlug = tripTitle ? tripTitle.replace(/[^a-zA-Z0-9]/g, "_") : "Trip_Budget";
  const sheetName = "Trip_" + cleanSlug;
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    if (sheet.getFilter()) sheet.getFilter().remove();
  }
  
  // 1. Title Header
  const titleText = (tripTitle || "SPECIAL TRIP / EVENT BUDGET").toUpperCase() + " TRACKER";
  sheet.getRange("A1:F1").merge().setValue(titleText)
       .setFontFamily("Arial").setFontSize(14).setFontWeight("bold")
       .setFontColor("#FFFFFF").setBackground("#1F4E78")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);
  
  // 2. KPI Summary Cards
  const kpiHeaders = [["Total Planned Budget", "Total Actual Spend", "Variance Remaining", "Budget Status"]];
  sheet.getRange("A3:D3").setValues(kpiHeaders).setFontWeight("bold").setBackground("#D9E1F2").setFontColor("#000000").setHorizontalAlignment("center");
  
  sheet.getRange("A4").setValue(totalBudget || 3000.00).setNumberFormat("$#,##0.00").setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#002060");
  sheet.getRange("B4").setFormula("=SUM(C9:C15)").setNumberFormat("$#,##0.00").setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#385723");
  sheet.getRange("C4").setFormula("=A4-B4").setNumberFormat("$#,##0.00").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D4").setFormula('=IF(C4>=0, "ON TRACK ✅", "OVER BUDGET ❌")').setFontWeight("bold").setHorizontalAlignment("center");
  
  // 3. Category Summary Table
  const catHeaders = [["Trip Category", "Planned Budget", "Actual Spend", "Variance"]];
  sheet.getRange("A7:D7").setValues(catHeaders).setFontWeight("bold").setBackground("#1F4E78").setFontColor("#FFFFFF");
  
  const categories = [
    ["Flights & Airfare", (totalBudget * 0.25), "=SUMIF(D18:D100, A9, C18:C100)", "=B9-C9"],
    ["Lodging & Hotels", (totalBudget * 0.35), "=SUMIF(D18:D100, A10, C18:C100)", "=B10-C10"],
    ["Events & Tickets", (totalBudget * 0.10), "=SUMIF(D18:D100, A11, C18:C100)", "=B11-C11"],
    ["Meals & Dining", (totalBudget * 0.15), "=SUMIF(D18:D100, A12, C18:C100)", "=B12-C12"],
    ["Sightseeing & Activities", (totalBudget * 0.08), "=SUMIF(D18:D100, A13, C18:C100)", "=B13-C13"],
    ["Ground Transportation", (totalBudget * 0.05), "=SUMIF(D18:D100, A14, C18:C100)", "=B14-C14"],
    ["Miscellaneous", (totalBudget * 0.02), "=SUMIF(D18:D100, A15, C18:C100)", "=B15-C15"]
  ];
  sheet.getRange(9, 1, categories.length, 4).setValues(categories);
  sheet.getRange("B9:D15").setNumberFormat("$#,##0.00");
  sheet.getRange("A9:A15").setFontWeight("bold");
  
  // Total summary row for categories
  sheet.getRange(16, 1, 1, 4).setValues([["Total Trip Category Budget", "=SUM(B9:B15)", "=SUM(C9:C15)", "=SUM(D9:D15)"]]).setFontWeight("bold").setBackground("#F2F2F2");
  sheet.getRange("B16:D16").setNumberFormat("$#,##0.00");
  
  // 4. Itemized Expense Log
  const logHeaders = [["Date", "Item Description", "Cost ($)", "Trip Category", "Payment Source", "Notes / Confirmation"]];
  sheet.getRange("A17:F17").setValues(logHeaders).setFontWeight("bold").setBackground("#1F4E78").setFontColor("#FFFFFF");
  sheet.setRowHeight(17, 25);
  
  // Validation for Category column in log
  const catRange = sheet.getRange("A9:A15");
  const catVal = SpreadsheetApp.newDataValidation().requireValueInRange(catRange).setAllowInvalid(false).build();
  sheet.getRange("D18:D100").setDataValidation(catVal);
  
  // Validation for Account Source in log from Enum sheet
  const enumSheet = ss.getSheetByName("Enum");
  if (enumSheet && enumSheet.getLastRow() >= 3) {
    const acctRange = enumSheet.getRange(3, 2, enumSheet.getLastRow() - 2, 1);
    const acctVal = SpreadsheetApp.newDataValidation().requireValueInRange(acctRange).setAllowInvalid(true).build();
    sheet.getRange("E18:E100").setDataValidation(acctVal);
  }
  
  sheet.getRange("C18:C100").setNumberFormat("$#,##0.00").setFontColor("#002060");
  sheet.getRange("A18:A100").setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
  sheet.getRange("A18:F100").setFontFamily("Arial").setFontSize(10);
  
  // Dummy starting expense sample
  sheet.getRange("A18:F18").setValues([[
    new Date(), "Sample Roundtrip Flight Reservation", (totalBudget * 0.20), "Flights & Airfare", "Chase Sapphire 9012", "Conf: Sample Flight #101"
  ]]);
  
  sheet.autoResizeColumns(1, 6);
  console.log(`[Trip Budget Engine] Created trip budget tab '${sheetName}' successfully.`);
}
