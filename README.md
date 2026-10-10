# Financial Reconciliation Automation

A Windows Tkinter GUI for reconciling general ledger and subledger transactions using Python. Designed as an educational demonstration of accounting controls and exception reporting, with synthetic sample data.

## Features
- Upload CSV or Excel (.xlsx) for the GL and subledger
- Match on transaction ID and account
- Detect missing transactions, amount/date differences, and duplicate keys
- Display counts and exceptions in a desktop interface
- Export exception details and summary to Excel
- Uses Decimal for currency; never automatically posts or clears exceptions

## Getting started
1. Install Python 3.10+ on Windows.
2. Run `pip install -r requirements.txt` in a terminal.
3. Double-click `financial_reconciliation.pyw` (or use `pythonw financial_reconciliation.pyw`).
4. Load the sample CSV files from `samples/`, then click **Run Reconciliation**.
5. Click **Export Report** to create an Excel workbook.

## Input schema
Each input must have these headers: `transaction_id,account,date,amount`. Dates should be supplied consistently as `YYYY-MM-DD` text. Amounts are signed monetary values, in the same currency and sign convention in both files. Matching is on `(transaction_id, account)`. Duplicated keys are flagged for manual investigation rather than silently netted.

## Limitations and controls
This demonstration is not a production accounting system. It does not validate posting periods, exchange rates, journal balancing, approval evidence, or authoritative transaction identity. It does not access company systems or transmit data. Reconciliation output should be independently reviewed before use in financial reporting.
