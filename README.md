# ACCT Excel Template

This is a custom Microsoft Excel template powered by VBA that automates formatting and layout for accounting tasks. It is designed specifically for:

- **Accounting students** following structured formats
- **Instructors** preparing formatted examples
- **Anyone** who wants to save time on repetitive accounting layouts

The template is tailored for the University of Calgary courses:

- **ACCT217:** Kimmel, Financial Accounting, 9e Canadian
- **ACCT323:** Weygandt, Managerial Accounting, 6e Canadian

---

## Features

- Automated formatting for common accounting tables and layouts
- Pre-built structures for journal entries, ledgers, financial statements, and more
- Consistent, professional appearance for all accounting documents
- Easy-to-use VBA macros for one-click formatting

---

## VBA Modules Overview
The template's functionality is organized into several VBA modules, each handling specific accounting automation tasks:

### Core Formatting Modules

- **[`FormatTable.bas`](Modules/FormatTable.bas)**: Automatically formats selected data as professional accounting tables with proper headers, borders, and total row formatting. Detects header rows and applies bold formatting with underlines to non-numeric cells, plus special formatting for total/net rows.

- **[`FormatLayouts.bas`](Modules/FormatLayouts.bas)**: Provides standardized layouts for common accounting documents:
  - Journal entries with Date, Account, Description, Debit, Credit columns
  - General ledger format with running balance
  - T-account layouts with proper dividers
  - Financial statement formatting with merged headers

### IFRS-Specific Modules

- **[`IFRSLibrary.bas`](Modules/IFRSLibrary.bas)**: Contains a comprehensive library of IFRS-compliant account names and line items. Creates dropdown lists for consistent account naming across financial statements, covering assets, liabilities, equity, income statement items, and managerial accounting terms.

- **[`IFRSLayouts.bas`](Modules/IFRSLayouts.bas)**: Specialized formatting functions for IFRS-compliant financial statements, including statement bodies, calculation tables, and proper totaling conventions.

### Utility Modules

- **[`Bootstrap.bas`](Modules/Bootstrap.bas)**: Template initialization and setup functions. Applies default formatting (Times New Roman, 11pt, accounting number format) and manages the IFRS dropdown functionality.

- **[`Indent.bas`](Modules/Indent.bas)**: Provides indentation controls for account hierarchies and statement formatting. Includes functions to increase (`IndentPlus4`) and decrease (`IndentMinus4`) indentation levels.

- **[`Globals.bas`](Modules/Globals.bas)**: Defines global constants for consistent formatting across all modules, including font settings, number formats, and date formats.

### Keyboard Shortcuts

The template includes keyboard shortcuts for quick access to formatting functions:
- **Ctrl+Shift+B**: Format Table
- **Ctrl+Shift+J**: Format Journal
- **Ctrl+Shift+T**: Format T-Account  
- **Ctrl+Shift+L**: Format Ledger
- **Ctrl+Shift+S**: Format Statement
- **Ctrl+Shift+I**: Apply IFRS Dropdown
- **Ctrl+Shift+U**: Remove Data Validation

All modules include error handling and are designed to work with selected ranges in Excel worksheets.

---

## How to Use

### 1. Open the Template

1. Download the `ACCT.xltm` or `ACCT 2.0.xltm` file from this repository.
2. Double-click the file to open it in Microsoft Excel. This will create a new workbook based on the template.

### 2. Enable Macros

When prompted, enable macros to allow the VBA automation to run. Macros are required for all automated formatting features.

### 3. Use the Automated Tools

The template provides several VBA-powered tools accessible via custom buttons or the "Developer" tab:

- **Format Table:** Automatically formats selected data as a professional accounting table.
- **Format Layouts:** Apply standard layouts for journal entries, ledgers, and statements.
- **Indent/Outdent:** Quickly adjust indentation for account names or line items.
- **Bootstrap:** Reset or initialize a new sheet with default settings.

Refer to the in-template instructions or button tooltips for details on each feature.

### 4. Customize as Needed

You can modify the template to suit your specific assignment or teaching needs. The VBA code is organized in modules for easy editing (see the `Modules/` folder for source code).

---

## Requirements

- Microsoft Excel (Windows or Mac) with macro support (usually .xlsm or .xltm files)
- Macros must be enabled for automation to work

---

## Support

For questions, suggestions, or bug reports, please open an issue in this repository.

---

## License

See the `LICENSE` file for details.

