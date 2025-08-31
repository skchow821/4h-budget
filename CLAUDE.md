# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script-based accounting system for tracking 4H club expenses and budgets. The entire codebase consists of a single JavaScript file (`Code.gs`) that integrates with Google Sheets.

## Architecture

The system is built around several key classes organized into logical sections:

### Core Data Models
- `BudgetEntry` - Individual budget line items
- `CombinedBudgetEntry` - Budget entries with project and category
- `LedgerEntry` - Individual expense/income transactions  
- `ClubInformation` - Club metadata and bank balance

### Processing Classes  
- `Ledger` - Manages collections of ledger entries with balance calculations
- `BudgetMap` - Maps budget categories to actual expenses and tracks variances

### Sheet Classes (Google Sheets Integration)
- `BasicSheet` - Base class for Google Sheets operations
- `FormattedSheet` - Adds formatting capabilities 
- `UserInputSheet` - Handles user data entry sheets
- Report sheets: `AnnualFinancialReportSheet`, `MonthlyReportSheet`, `BudgetReportSheet`
- Input sheets: `LedgerSheet`, `BudgetPlanningSheet`

### Key Functions
- `_enumerateMonths()` - Breaks date ranges into monthly periods
- `_getBalances()` - Calculates running balances for entries
- `_createMonthlyLedgers()` - Groups transactions by month
- `_convertToSparse()` - Formats data for Google Sheets

## Development

### Testing
The file includes several test functions prefixed with `test_`:
- `test_retrieve_form()` - Tests form data retrieval
- `test_sparse()` - Tests sparse data conversion  
- `test_formatSheet()` - Tests sheet formatting
- `test_BudgetMap()` - Tests budget mapping functionality

To run tests, execute the individual test functions in the Google Apps Script environment.

### Google Apps Script Integration
This code is designed to run in Google Apps Script and uses:
- `SpreadsheetApp` for Google Sheets integration
- `Logger` for debugging output
- Google Sheets API for data manipulation

The main entry point appears to be at the bottom of the file where balance calculations are performed for both ledger and budget planning sheets.
