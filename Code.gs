// Simple accounting system for keeping track of 4H expenses.
// 2025/02/01
// v 0.1
// Sunny Chow

// ---
// Section: Utility Functions
// ---
function _enumerateMonths(startingDate, endingDate) {
  // Fix endingDate (make it span all the way to the last ms of that date)
  endingDate.setTime(endingDate.getTime() + 86399999);

  if (startingDate > endingDate) {
    throw new Error("Start Date must be before End Date");
  }

  let results = [];
  let tempStartDate = new Date(startingDate);

  while (tempStartDate < endingDate) {
    let tempEndDate = new Date(tempStartDate.getFullYear(), tempStartDate.getMonth() + 1, 1);
    tempEndDate.setTime(tempEndDate.getTime() - 1); // 1 ms before.

    if (tempEndDate > endingDate) {
      tempEndDate = endingDate;
    }
    results.push([
      tempStartDate, tempEndDate
    ]);

    tempStartDate = new Date(tempEndDate.getFullYear(), tempEndDate.getMonth(), tempEndDate.getDate() + 1);
  }
  return results;
}

function _getBalances(beginningBalance, entries, fnDelta) {
  // Fetch data from the active sheet
  var balances = [];
  var balanceIdx = 0;

  // for each entry create an updated balance element
  entries.forEach(entry => {
    var prev = 0;

    if (balanceIdx == 0) {
      prev = beginningBalance;
    }
    else {
      prev = balances[balanceIdx - 1];
    }

    balances.push(prev + fnDelta(entry));
    balanceIdx++;
  });

  return balances;
}

function _getBalancesPerLedgerEntry(beginningBalance, entries) {
  return _getBalances(beginningBalance, entries, item => { return item.income - item.expense});
}

function _getBalancesPerBudgetPlanningEntry(beginningBalance, entries) {
  return _getBalances(beginningBalance, entries, item => { return item.income.budgeted - item.expense.budgeted});
}

function _createMonthlyLedgers(beginningBalance, startingDate, endingDate, entries) {
  var monthlyLedgers = [];
  var lastBalance = beginningBalance;

  const monthRanges = _enumerateMonths(startingDate, endingDate);

  monthRanges.forEach(monthTuple => {
    const tempLedger = new Ledger(monthTuple[0], monthTuple[1], lastBalance, entries, (entry) => { return entry.dateReconciled});
    lastBalance = tempLedger.endingBalance;
    monthlyLedgers.push(tempLedger);
  });

  return monthlyLedgers;
}

function _convertToSparse(data, cols) {
  var sparse = [];
  for (var i = 0; i < data.length; ++i) {
    var temp = [];
    for (var j = 0; j < data[i].length; ++j) {
      temp.push(data[i][j])
      for (var skip = 1; skip < cols[j]; ++skip) {
        temp.push("");
      }
    }
    sparse[i] = temp;
  }
  return sparse;
}

// ---
// Section: Utility Classes
// ---
class BudgetEntry {
  constructor(budgeted = 0) {
    this.actual = 0;
    this.budgeted = budgeted;
  }

  accumulate(val) {
     this.actual += val;
  }
}

class CombinedBudgetEntry {
  constructor(subAccount, category, budgetedIncome, budgetedExpense) {
    this.subAccount = subAccount;
    this.category = category;
    this.income = new BudgetEntry(budgetedIncome);
    this.expense = new BudgetEntry(budgetedExpense);
  }

  accumulate(entry) {
     this.income.accumulate(entry.income);
     this.expense.accumulate(entry.expense);
  }
}

class LedgerEntry {
  constructor(dateRecorded, dateReconciled, number, subject, category, subAccount, description, income, expense) {
    this.dateRecorded = dateRecorded;
    this.dateReconciled = dateReconciled;
    this.number = number;
    this.subject = subject;
    this.subAccount = subAccount;
    this.category  = category;
    this.description = description;
    this.income = income;
    this.expense = expense;
  }

  label() {
    let txt = this.subAccount != "" ? this.subAccount + ": " : "";
    txt += this.category;
    txt += this.description != "" ? ", " + this.description : "";
    txt += this.subject != "" ? ", " + this.subject : "";
    return txt
  }
}

class ClubInformation {
  constructor(county, name, address, ein, treasurerName, treasurerPhone, treasurerEmail, bankName, bankAccountNo, yearStart, yearEnd, bankBalance) {
    this.county = county;
    this.name = name;
    this.address = address;
    this.ein = ein;
    this.treasurerName = treasurerName;
    this.treasurerPhone = treasurerPhone;
    this.treasurerEmail = treasurerEmail;
    this.bankName = bankName;
    this.bankAccountNo = bankAccountNo;
    this.startingDate = yearStart;
    this.endingDate = yearEnd;
    this.bankBalance = bankBalance;
  }
}


class Ledger {
  constructor(startingDate, endingDate, beginningBalance, entries, fnDate = (a) => {return a.dateRecorded;} ) {
    this.startingDate = startingDate;
    this.endingDate = endingDate;
    this.beginningBalance = beginningBalance;
    this.endingBalance = this.beginningBalance;
    this.entries = [];
    entries.forEach(entry => {
      const date = fnDate(entry);
      let startingDate = this.startingDate;
      let endingDate = this.endingDate;
      if (date == null || isNaN(date) || date < this.startingDate || date > this.endingDate) {
        return;
      }

      this.addEntry(entry);
    });
  }

  addEntry(entry) {
    this.entries.push(entry);
    this.endingBalance -= entry.expense;
    this.endingBalance += entry.income;
  }

  getIncomeEntries() {
    let incomeEntries = [];
    this.entries.forEach (entry => {
      if (entry.income > 0.0) {
        incomeEntries.push(entry);
      }
    });
    return incomeEntries;
  }

  getExpenseEntries() {
    let expenseEntries = [];
    this.entries.forEach (entry => {
      if (entry.expense > 0.0) {
        expenseEntries.push(entry);
      }
    });
    return expenseEntries;
  }

  getTotalIncome() {
    let entries = this.getIncomeEntries();

    let sum = entries.length > 0 ? entries.reduce((accumulator, item) =>
      { return accumulator + item.income; }, 0) : 0.0;

    return sum;
  }

  getTotalExpense() {
    let entries = this.getExpenseEntries();

    let sum = entries.length > 0 ? entries.reduce((accumulator, item) =>
      { return accumulator + item.expense; }, 0) : 0.0;

    return sum;
  }
}

class BudgetMap {
  constructor(budgetEntries) {
    let lookup = {};
    budgetEntries.forEach( entry => {
        if (!lookup[entry.subAccount]) {
          lookup[entry.subAccount] = {}
        }
        lookup[entry.subAccount][entry.category] = entry;
    });
    this.lut = lookup;
  }

  addEntry(entry) {
    let subAccountLookup = entry.subAccount;
    let categoryLookup = entry.category;

    // Not found, redirect into it's own Uncategorized bin
    if (!this.lut[subAccountLookup] ||
       !this.lut[subAccountLookup][categoryLookup]) {
      categoryLookup = categoryLookup + " - NEW";

      if (!this.lut[subAccountLookup]) {
         this.lut[subAccountLookup] = {};
      }

      if (!this.lut[subAccountLookup][categoryLookup]) {
        this.lut[subAccountLookup][categoryLookup]
          = new CombinedBudgetEntry(subAccountLookup, categoryLookup, 0, 0);
      }
    }

    this.lut[subAccountLookup][categoryLookup].accumulate(entry);
  }

  _addCategoriesFromSubAccount(data, acct, fnEntrySelect) {
    let sortedCategories = Object.keys(this.lut[acct]).sort();
    sortedCategories.forEach( cat => {
        let entry = fnEntrySelect(this.lut[acct][cat]);

        // Skip entries where no money was budgeted or spent.
        if (entry.actual === 0.0 && entry.budgeted === 0.0 ) { 
          return;
        }

        let name = acct === "" ? this.lut[acct][cat].category :  acct + " - " + this.lut[acct][cat].category;
        data.push([name, entry.budgeted, entry.actual]);
    });
  }

  _getAccumulations(fnEntrySelect) {
    let data = []

    this._addCategoriesFromSubAccount(data, "", fnEntrySelect);

    // Add a Projects sub-heading.
    data.push(["PROJECTS (SUB_ACCOUNTS)", "", ""]);

    // Add sub-projects
    let sortedSubaccounts = Object.keys(this.lut).sort();

    sortedSubaccounts.forEach( acct => {
      if (acct === "") { return };

      this._addCategoriesFromSubAccount(data, acct, fnEntrySelect);
    });
    return data;
  }

  getIncomeEntries() {
    return this._getAccumulations(entry => { return entry.income; });
  }

  getExpenseEntries() {
    return this._getAccumulations(entry => { return entry.expense; });
  }

  getBudgetedIncome() {
    return this.reduce((accum, entry) => { return accum + entry.income.budgeted; });
  }

  getBudgetedExpense() {
     return this.reduce((accum, entry) => { return accum + entry.expense.budgeted; }); 
  }

  reduce(fnPerEntry) {
    let accum = 0;
    Object.values(this.lut).forEach( cats => {
      let entries = Object.values(cats);
      entries.forEach(entry => {
          accum = fnPerEntry(accum, entry);
      })
    });
    return accum;
  }
}

// ---
// Section: Sheet Functions
// ---
class BasicSheet {
  constructor(spreadsheet, name) {
    this.name = name;
    this.sheet = spreadsheet.getSheetByName(name);
    if (!this.sheet) {
      this.sheet = spreadsheet.insertSheet(name);
    }

    this.curRow = 1;
    this.dateRule = SpreadsheetApp.newDataValidation().requireDate().build();
    this.currencyRule = SpreadsheetApp.newDataValidation().requireNumberBetween(-1e9, 1e9).build();
    this.emailRule = SpreadsheetApp.newDataValidation().requireTextIsEmail().build();
  }

  isEmpty() {
    var data = this.sheet.getDataRange().getValues();

    return (data.length === 1 && data[0].length === 1 && data[0][0] === "");
  }

  clear() {
    this.sheet.clear();
  }

	_toFmt(typeStr) {
		var fmt = "@";
    if (typeStr === "date") {
      fmt = "MM/dd/yyyy";
    }
    if (typeStr == "email") {
      fmt = "%";
    }
    if (typeStr === "int") {
      fmt = "0";
    }
    if (typeStr === "currency") {
      fmt = "$#,##0.00"
    }
		return fmt;
	}

	_toRule(typeStr) {
		var rule = null;
    if (typeStr === "date") {
      rule = this.dateRule;
    }
    if (typeStr == "email") {
      rule = this.emailRule;
    }
    if (typeStr === "currency") {
      rule = this.currencyRule;
    }
		return rule;
	}
}

class FormattedSheet extends BasicSheet {
  constructor(spreadsheet, name, numCols, width) {
    super(spreadsheet, name);
    this.nCols = numCols;
    this.width = width;
    const colWidth = this.width / this.nCols;
    this.sheet.setColumnWidths(1, numCols, colWidth);
  }

  heading(text) {
    this.sheet.getRange(this.curRow, 1, 1, 1).setValues([[text]]).setBackground("#008000").setFontColor("#FFFFFF");
    this.sheet.getRange(this.curRow, 1, 1, this.nCols).mergeAcross();
    this.curRow++;
  }

  paragraph(text) {
    this.sheet.getRange(this.curRow, 1, 1, 1).setValues([[text]]);
     this.sheet.getRange(this.curRow, 1, 1, this.nCols).mergeAcross()
      .setHorizontalAlignment("center").setWrap(false);
    this.curRow++;
  }

  labelText(label, text, col) {
    if (col + 1 > this.nCols) {
      throw Error("Label / Text pair at position " + col + " will not fit!");
    }

    this.sheet.getRange(this.curRow, col, 1, 2).setValues([[label, text]]);
    this.sheet.getRange(this.curRow, col, 1, 1).setFontWeight("bold").setHorizontalAlignment("right").setWrap(false)
    this.sheet.getRange(this.curRow, col+1, 1, 1).setHorizontalAlignment("left").setWrap(false);
  }

  newline() {
    this.curRow++;
  }

  startTable(header, colWidths, formats = null) {
    if (header.length != colWidths.length) {
      throw Error("colWidths and header must be arrays of the same length.");
    }

    const colsUsed = colWidths.reduce((accum, value) => accum + value, 0);

    if (colsUsed > this.nCols) {
      throw Error("the summation of colWidths must be less than " + this.nCols);
    }

    // Store table state for endTable
    this.activeTable = {
      colWidths: colWidths,
      colsUsed: colsUsed,
      tableRowStart: this.curRow,
      formats: formats,
      dataRowCount: 0
    };

    // header
    let sparseHeader = _convertToSparse([header], colWidths);
    this.sheet.getRange(this.curRow, 1, 1, colsUsed).setValues(sparseHeader).setBackground("#008000").setFontColor("#FFFFFF");
    this.curRow++;
  }

  addTableRow(data) {
    if (!this.activeTable) {
      throw Error("Must call startTable() before adding rows");
    }
    if (data.length != this.activeTable.colWidths.length) {
      throw Error("data length must match colWidths length");
    }

    let sparseData = _convertToSparse([data], this.activeTable.colWidths);
    this.sheet.getRange(this.curRow, 1, 1, this.activeTable.colsUsed).setValues(sparseData);
    this.curRow++;
    this.activeTable.dataRowCount++;
  }

  endTable(footer) {
    if (!this.activeTable) {
      throw Error("Must call startTable() before endTable()");
    }
    if (footer.length != this.activeTable.colWidths.length) {
      throw Error("footer and header must be arrays of the same length.");
    }

    // footer
    let sparseFooter = _convertToSparse([footer], this.activeTable.colWidths);
    this.sheet.getRange(this.curRow, 1, 1, this.activeTable.colsUsed).setValues(sparseFooter).setFontWeight('bold');
    this.curRow++;

    const totalTableRows = this.activeTable.dataRowCount + 1; // +1 for footer

    // Merge cells according to colWidths
    let colIdx = 1;
    this.activeTable.colWidths.forEach(width => {
      this.sheet.getRange(this.activeTable.tableRowStart, colIdx, totalTableRows + 1, width).mergeAcross();
      colIdx += width;
    });

    // Add borders
    this.sheet.getRange(this.activeTable.tableRowStart, 1, totalTableRows + 1, this.activeTable.colsUsed)
      .setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);

    // Apply formats (for data rows only, not header or footer)
    if (this.activeTable.formats != null) {
      if (this.activeTable.formats.length != this.activeTable.colWidths.length) {
        throw Error("formats and colWidths must be arrays of the same length");
      }
      colIdx = 1;
      for (let i = 0; i < this.activeTable.formats.length; ++i) {
        let colFormat = this._toFmt(this.activeTable.formats[i])
        this.sheet.getRange(this.activeTable.tableRowStart + 1, colIdx, this.activeTable.dataRowCount, 1).setNumberFormat(colFormat);
        colIdx += this.activeTable.colWidths[i];
      }
    }

    // Clear table state
    this.activeTable = null;
  }

  addTable(header, data, footer, colWidths, formats = null) {
    if (header.length != data[0].length) {
      throw Error("header and data lengths must match");
    }
    if (footer.length != header.length) {
      throw Error("footer and header must be arrays of the same length.");
    }

    // Use the new startTable/endTable methods
    this.startTable(header, colWidths, formats);

    data.forEach(row => {
      this.addTableRow(row);
    });

    this.endTable(footer);
  }
}

class UserInputSheet extends BasicSheet {
  constructor(spreadsheet, name) {
    super(spreadsheet, name);
  }

  // Used to create a table for user entry.
  setTable(header) {
	  var headerTitles = [];
	  header.forEach(item => {
		  headerTitles.push(item[0]);
	  });

		var nR = 1;
	  this.sheet.getRange(nR, 1, 1, headerTitles.length).setValues([headerTitles])
		  .setFontWeight("bold").setBackground("#008000").setFontColor("#FFFFFF"); 

	  this.sheet.autoResizeColumns(1, headerTitles.length); 

		nR++;

		var idx = 1;
		var affectRows = 255;
		// Set validation rules
		header.forEach(item => {
			const fmt = this._toFmt(item[1]);
			const rule = this._toRule(item[1]);
			this.sheet.getRange(nR, idx, affectRows, 1).setDataValidation(rule).setNumberFormat(fmt);
			idx++;
		});
  }

  retrieveTableEntries(header) {
    // FIXME: hardcoded idxs...
    const actualHeader = this.sheet.getRange(1, 1, 1, header.length).getValues();

    for (var i = 0; i < header.length; ++i) {
      if (header[i][0] !== actualHeader[0][i]) {
         throw Error("Header for table has changed.  Expected: " + header[i][0] + " Actual: " + actualHeader[0][i]);
      }
    }

    // FIXME: hardcode header idx
    // (-1 to get rid of header)
    return this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1, header.length).getValues();
  }

  updateColumn(colIdx, content) {
    var contentRowOrder = [];
    content.forEach(item => {
        contentRowOrder.push([item]);
    });

    this.sheet.getRange(2, colIdx, content.length, 1).setValues(contentRowOrder)
      .setBackground("#DDDDDD");
  }

  addForm(form) {
    form.forEach(formRowContent => {

      this.sheet.getRange(this.curRow, 1, 1, 1).setValues([[formRowContent[0]]])
        .setFontWeight("bold").setBackground("#008000").setFontColor("#FFFFFF");

      var fmt = this._toFmt(formRowContent[1]);
      var rule = this._toRule(formRowContent[1]);

      if (rule != null) {
        this.sheet.getRange(this.curRow, 2, 1, 1).setDataValidation(rule).setNumberFormat(fmt);
      }

      // hack
      this.curRow++;
    });

	  this.sheet.autoResizeColumn(1);
  }

  retrieveForm(form) {
    var formResponses = [];
    var rowIdx = 1;
    form.forEach(row => {
      var values = this.sheet.getRange(rowIdx, 1, 1, 2).getValues();
      if (values[0][0] !== row[0]) {
        throw Error("On row " + rowNo + ", Expected: " + row[0] + ", Found:" + values[0][0] );
      }

      if (row[1] == "date") {
        formResponses.push(new Date(values[0][1]));
      }
      else {
        formResponses.push(values[0][1]);
      }
      rowIdx++;
    });
    return formResponses;
  }
}

// ---
// Section: Compiled Tables
// ---
class AnnualFinancialReportSheet extends FormattedSheet {
  constructor(spreadsheet, clubInfo, monthlyLedgers, totalLedger) {
    super(spreadsheet, "Form 6.3 - Annual Financial Report", 4, 800);
    this.clubInfo = clubInfo
    this.monthlyLedgers = monthlyLedgers;
    this.totalLedger = totalLedger;
    this.clear();
  }

  publish() {
    this.heading("Form 6.3 ANNUAL FINANCIAL REPORT");
    this.paragraph(
      this.clubInfo.startingDate.toLocaleDateString('en-US', {month: 'long', day: 'numeric', year: 'numeric'}) + 
      " to " + 
      this.clubInfo.endingDate.toLocaleDateString('en-US', {month: 'long', day: 'numeric', year: 'numeric'}));
    this.labelText("County", this.clubInfo.county, 1);
    this.labelText("Treasurer Name", this.clubInfo.treasurerName, 3);
    this.newline();
    this.labelText("Club Name", this.clubInfo.name, 1);
    this.labelText("Treasurer Phone", this.clubInfo.treasurerPhone, 3);
    this.newline();
    this.labelText("EIN", this.clubInfo.ein, 1);
    this.labelText("Treasurer Email", this.clubInfo.treasurerEmail, 3);
    this.newline();
    this.labelText("Bank Account", "Checking", 1);
    this.newline();
    this.labelText("Bank Name", this.clubInfo.bankName, 1);
    this.labelText("Last 4 Digits of Account No", this.clubInfo.bankAccountNo, 3);
    this.newline();
    this.newline();
    this.labelText("Bank Balance at the end of the previous year", this.clubInfo.bankBalance, 3);
    this.newline();
    this.newline();

    const data = this.monthlyLedgers.map(item => { return [
        item.startingDate.toLocaleDateString('en-US', {month: 'long'}),
        item.getTotalIncome(),
        item.getTotalExpense(),
        item.endingBalance
      ];
    });

    this.addTable(["MONTH", "TOTAL INCOME", "TOTAL EXPENSES", "BALANCE"],
                  data,
                  ["Total", this.totalLedger.getTotalIncome(), this.totalLedger.getTotalExpense(), this.totalLedger.endingBalance],
                  [1, 1, 1, 1],
                  ["string", "currency", "currency", "currency"]);
  }
}

class MonthlyReportSheet extends FormattedSheet {
  constructor(spreadsheet, clubInfo, monthlyLedgers) {
    super(spreadsheet, "Form 6.1 - Monthly Ledger", 4, 800);
    this.clubInfo = clubInfo;
    this.monthlyLedgers = monthlyLedgers;
    this.clear();
  }

  publish() {
    this.monthlyLedgers.forEach(ledger => {
      this.heading("Form 6.1 - 4-H Monthly Meeting Report Form");
      this.labelText("Club Name:", this.clubInfo.county, 1);
      this.newline();
      this.labelText("Location:", this.clubInfo.address, 1);
      this.newline();
      this.labelText("Month:", ledger.startingDate.toLocaleDateString('en-US', {month: 'long', year: 'numeric'}), 1);
      this.newline();
      this.labelText("Total Opening Balance $:", ledger.beginningBalance, 3);
      this.newline();
      this.newline();

      let emptyEntry = [["-", ""]];
      // Add Income Entries
      {
        let entries = ledger.getIncomeEntries();
        let data = entries.length > 0 ? entries.map(entry => {
          return [entry.label(), entry.income];
        }) : emptyEntry;
        this.addTable(
          ["Income (SOURCE, USE, PURPOSE)", "Amount"],
          data,
          ["Total Income", ledger.getTotalIncome()],
          [ 3, 1],
          ["string", "currency"]
        );
      }
      this.newline();

      // Add Expense Entries
      {
        let entries = ledger.getExpenseEntries();
        let data = entries.length > 0 ? entries.map(entry => {
          return [entry.label(), entry.expense];
        }) : emptyEntry;
        this.addTable(
          ["Expense (DESCRIBE)", "Amount"],
          data,
          ["Total Expenses", ledger.getTotalExpense()],
          [ 3, 1],
          ["string", "currency"]
        );
      }
      this.labelText("Closing Balance $:", ledger.endingBalance, 3);
      this.newline();
      this.newline();
    });
  }
}

class BudgetReportSheet extends FormattedSheet {
  constructor(spreadsheet, clubInfo, budgetMap, monthlyLedgers, totalLedger) {
    super(spreadsheet, "Form 8.4 - Budget", 5, 800);
    this.clubInfo = clubInfo;
    this.monthlyLedgers = monthlyLedgers;
    this.totalLedger = totalLedger;
    this.budgetMap = budgetMap;
    this.clear();
  }

  publish() {
    this.heading("4-H CLUB BUDGET");
    this.labelText("Club Name", this.clubInfo.name, 1);
    this.newline();
    this.paragraph(
      this.clubInfo.startingDate.toLocaleDateString('en-US', {month: 'long', day: 'numeric', year: 'numeric'}) + 
      " to " +
      this.clubInfo.endingDate.toLocaleDateString('en-US', {month: 'long', day: 'numeric', year: 'numeric'}));
    this.labelText("Total Opening Balance: $", this.totalLedger.beginningBalance, 4);
    this.newline();

    this.addTable(
          ["Estimated Income (SOURCE, USE, PURPOSE)", "BUDGETED", "ACTUAL"],
          this.budgetMap.getIncomeEntries(),
          ["Total Income: $", this.budgetMap.getBudgetedIncome(), this.totalLedger.getTotalIncome()],
          [ 3, 1, 1],
          [ "string", "currency", "currency"]);

    this.newline();

    this.addTable(
          ["Estimated Expenses (DESCRIBE)", "BUDGETED", "ACTUAL"],
          this.budgetMap.getExpenseEntries(),
          ["Total Expenses: $", this.budgetMap.getBudgetedExpense(), this.totalLedger.getTotalExpense()],
          [ 3, 1, 1],
          ["string", "currency", "currency"]);

    this.newline();

    this.labelText("Beginning Balance", this.totalLedger.beginningBalance, 4);
    this.newline();
    this.labelText("Income + ", this.totalLedger.getTotalIncome(), 4);
    this.newline();
    this.labelText("Expenses - ", this.totalLedger.getTotalExpense(), 4);
    this.newline();
    this.labelText("Total Closing Balance", this.totalLedger.endingBalance, 4);
    this.newline();
    this.newline();
    this.labelText("We certify that this budget was approved by the club meeting on (date)", "[Fill In]", 4);
    this.newline();
    this.newline();
    this.labelText("Club President" , "[Print Name]", 2);
    this.labelText("Signature", "[Sign]", 4);
    this.newline();
    this.labelText("Treasurer", "[Print Name]", 2);
    this.labelText("Signature", "[Sign]", 4);
    this.newline();
    this.labelText("Club Leader", "[Print Name]", 2);
    this.labelText("Signature", "[Sign]", 4);
    this.newline();
    this.labelText("County Director or designee*", "[Print Name]", 2);
    this.labelText("Signature", "[Sign]", 4);
  }
};

// ---
// Section: Tables & Forms for User Input
// ---
class LedgerSheet extends UserInputSheet {
   constructor(spreadsheet) {
    super(spreadsheet, "Ledger");
    this.entries = [];

	this.header = [
		["Date", "date"],
		["Date Reconciled", "date"],
    ["Check No/Receipt No", "string"],
    ["To/From", "string"],
		["Purpose/Category", "string"],
    ["Sub-Account", "string"],
		["Description", "string"],
		["Income ($)", "currency"],
		["Expense ($)", "currency"],
		["Balance ($)", "currency"]];
  }

  populate() {
		if (!this.isEmpty()) {
			return;
		}

		this.setTable(this.header);
  }

  retrieve() {
    const results = this.retrieveTableEntries(this.header);
    results.forEach(row => {
      this.entries.push(new LedgerEntry(
         row[0] = new Date(row[0]),  // Date
         row[1] = new Date(row[1]),  // Date Reconciled
         row[2],                     // Check No / Receipt No
         row[3],                     // To / From
         row[4],                     // Purpose / Category
         row[5],                     // Sub-Account
         row[6],                     // Description
         row[7] === "" ? 0 : row[7], // Income
         row[8] === "" ? 0 : row[8], // Expense
      ));
    });
    return this.entries;
  }

   updateBalances(balances) {
     this.updateColumn(this.header.length, balances);
   }
}

class BudgetPlanningSheet extends UserInputSheet {
	constructor(spreadsheet) {
		super(spreadsheet, "Budget Planning");
		this.header = [
			["Sub Account", "string"],
			["Category", "string"],
			["Budgeted Income ($)", "currency"],
			["Budgeted Expense ($)", "currency"],
			["Balance ($)", "currency"],
		];
	}

  populate() {
		if (!this.isEmpty()) {
			return;
		}

		this.setTable(this.header);
  }

  retrieve() {
    const results = this.retrieveTableEntries(this.header);
    var entries = [];
    results.forEach(row => {
      entries.push(new CombinedBudgetEntry(
        row[0],
        row[1],
        row[2] === "" ? 0.0 : row[2],
        row[3] === "" ? 0.0 : row[3]));
    });
    return entries;
  }

  updateBalances(balances) {
    this.updateColumn(this.header.length, balances);
  }
}

class ClubInformationSheet extends UserInputSheet {
  constructor(spreadsheet) {
    super(spreadsheet, "Club Information");
    this.contents = [
      ["County", "string"],
      ["Club Name", "string"],
			["Club Address", "string"],
      ["Club EIN", "string"],
      ["Treasurer Name", "string"],
      ["Treasurer Phone", "string"],
      ["Treasurer Email", "email"],
      ["Bank Name", "string"],
      ["Bank Account No (last 4 digits)", "string"],
      ["Year Start", "date"],
      ["Year End", "date"],
      ["Bank Balance (end of previous year)", "currency"]
    ];
  }

  populate() {
    if (!this.isEmpty()) {
      return;
    }

    this.addForm(this.contents);
  }

  retrieve() {
    var inputs = [];
    inputs = this.retrieveForm(this.contents);
    return new ClubInformation(
      inputs[0],
      inputs[1],
      inputs[2],
      inputs[3],
      inputs[4],
      inputs[5],
      inputs[6],
      inputs[7],
      inputs[8],
      inputs[9],
      inputs[10],
      inputs[11],
    );
  }
}

// ---
// Section: Test Functions
// ---
function test_retrieve_form() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let cis = new ClubInformationSheet(spreadsheet);
  let items = cis.retrieve();
}

function test_sparse() {
  let data = [["item 1",  1]];
  let cols =  [       3,  2];
  let sparse = _convertToSparse(data, cols);
  console.log(sparse);

  data = [["item 2",  2]];
  cols =  [       1,  4];
  sparse = _convertToSparse(data, cols);
  console.log(sparse);
}

function test_formatSheet(){
  let sheet = new FormattedSheet(
    SpreadsheetApp.getActiveSpreadsheet(),
    "Test Format",
    4,
    800);

  sheet.clear();

  sheet.heading("Heading");
  sheet.labelText("Label 1", "Text 1", 1);
  sheet.labelText("Label 2", "Text 2", 3);
  sheet.newline();
  sheet.labelText("Label 3", "Text 3", 1);
  sheet.labelText("Label 4", "Text 4", 3);
  sheet.newline();
  sheet.newline();

  sheet.addTable(["Label", "Value1", "Value2"],
                 [["foo", 1, 2], ["bar", 3, 4]],
                 ["Total", 5, 6],
                 [     2,      1,       1]);

  // Test new startTable/endTable methods
  sheet.newline();
  sheet.heading("Testing startTable/endTable");
  sheet.startTable(["Item", "Quantity", "Price"], [2, 1, 1], ["string", "int", "currency"]);
  sheet.addTableRow(["Apples", 5, 2.50]);
  sheet.addTableRow(["Bananas", 3, 1.25]);
  sheet.endTable(["Total", 8, 3.75]);
}

function test_BudgetMap() {
  let budgetEntries = [
    new CombinedBudgetEntry("", "Tools", 0, 100),
    new CombinedBudgetEntry("", "Stationary", 0, 200),
    new CombinedBudgetEntry("Trees", "Tools", 0, 300),
  ];

  let budgetMap = new BudgetMap(budgetEntries);

  budgetMap.addEntry(new LedgerEntry(
    new Date("2025/01/04"),
    new Date("2025/01/07"),
    "",
    "Orchard Supply",
    "Tools",
    "",
    "Pruning tools",
    0,
    50
  ));

  budgetMap.addEntry(new LedgerEntry(
    new Date("2025/01/08"),
    new Date("2025/01/10"),
    "",
    "Staples",
    "Stationary",
    "",
    "Paper",
    0,
    10
  ));

  budgetMap.addEntry(new LedgerEntry(
    new Date("2025/02/08"),
    new Date("2025/02/10"),
    "",
    "Home Depot",
    "Tools",
    "",
    "Lawn Mower",
    0,
    500
  ));

  budgetMap.addEntry(new LedgerEntry(
    new Date("2025/02/08"),
    new Date("2025/02/10"),
    "",
    "Baskin Robbins",
    "Fundraiser",
    "",
    "Party",
    500,
    0
  ));

  budgetMap.addEntry(new LedgerEntry(
    new Date("2025/03/02"),
    new Date("2025/03/10"),
    "",
    "Safeway",
    "Party",
    "",
    "Food & Drinks",
    0,
    100
  ));

  budgetMap.addEntry(new LedgerEntry(
    new Date("2025/03/04"),
    new Date("2025/03/13"),
    "",
    "Lowes",
    "Tools",
    "Trees",
    "Pruners",
    0,
    75
  ));


  //  [ 'Uncategorized', 0, 500 ],
  //  [ 'PROJECTS (SUB_ACCOUNTS)', '', '' ],
  let incomeEntries = budgetMap.getIncomeEntries();
  console.log(incomeEntries);

  //[ [ 'Tools', 100, 550 ],
  //  [ 'Stationary', 200, 10 ],
  //  [ 'Uncategorized', 0, 100 ],
  //  [ 'PROJECTS (SUB_ACCOUNTS)', '', '' ],
  //  [ 'Trees - Tools', 300, 75 ] ]
  let expenseEntries = budgetMap.getExpenseEntries();
  console.log(expenseEntries);

  let totalBudgetExpense = budgetMap.getBudgetedExpense();
  console.log(totalBudgetExpense);
}

// ---
// Section: User Functions
// ---
function create_input_tables() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let cis = new ClubInformationSheet(spreadsheet);
  cis.populate();

  let ledger = new LedgerSheet(spreadsheet);
  ledger.populate();

  let budgetPlanning = new BudgetPlanningSheet(spreadsheet);
  budgetPlanning.populate();
}

function publish_4HForms() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let clubInformationSheet = new ClubInformationSheet(spreadsheet);
  const clubInfo = clubInformationSheet.retrieve();

  let ledgerSheet = new LedgerSheet(spreadsheet);
  const entries = ledgerSheet.retrieve();

  const totalLedger = new Ledger(clubInfo.startingDate, clubInfo.endingDate, clubInfo.bankBalance, entries, (entry) => { return entry.dateReconciled});

  let monthlyLedgers = _createMonthlyLedgers(
    clubInfo.bankBalance,
    clubInfo.startingDate,
    clubInfo.endingDate,
    entries);

  // Annual Financial Report
  {
    const financialReport = new AnnualFinancialReportSheet(spreadsheet, clubInfo, monthlyLedgers, totalLedger);
    financialReport.publish();
  }

  // Monthly Ledgers
  {
    const monthlyReport = new MonthlyReportSheet(spreadsheet, clubInfo, monthlyLedgers);
    monthlyReport.publish();
  }

  // Budget Report
  {
    let budgetSheet = new BudgetPlanningSheet(spreadsheet);
    const budgetEntries = budgetSheet.retrieve();

    let budgetMap = new BudgetMap(budgetEntries);
    entries.forEach( entry => { budgetMap.addEntry(entry); });

    const budgetReport = new BudgetReportSheet(spreadsheet, clubInfo, budgetMap, monthlyLedgers, totalLedger);
    budgetReport.publish();
  }
}

function update_balances() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let clubInformationSheet = new ClubInformationSheet(spreadsheet);
  clubInfo = clubInformationSheet.retrieve();

  {
    let ledgerSheet = new LedgerSheet(spreadsheet);
    const entries = ledgerSheet.retrieve();
    const balances = _getBalancesPerLedgerEntry(clubInfo.bankBalance, entries);
    ledgerSheet.updateBalances(balances);
  }

  {
    let budgetPlanningSheet = new BudgetPlanningSheet(spreadsheet);
    const entries = budgetPlanningSheet.retrieve();
    const balances = _getBalancesPerBudgetPlanningEntry(clubInfo.bankBalance, entries);
    budgetPlanningSheet.updateBalances(balances);
  }
}

function run_tests() {
  test_formatSheet();
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or SlidesApp or FormApp.
  menu.addItem('Start', 'create_input_tables');
  menu.addItem('Update Balances', 'update_balances');
  menu.addItem('Generate 4H Forms', 'publish_4HForms');
  menu.addItem('Run tests', 'run_tests');
  menu.addToUi();
}
