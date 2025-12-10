# QuickBooks Invoice Connector

A command-line tool for synchronising invoices between Excel and QuickBooks Desktop.
The tool compares records, detects conflicts, identifies missing invoices, and writes a detailed JSON report.

------------------------------------------------------------
SETUP PROJECT
------------------------------------------------------------

Clone the repository and install dependencies using Poetry:

    poetry install

All development happens inside the src/ directory.

------------------------------------------------------------
QUALITY CHECK
------------------------------------------------------------

Install pre-commit hooks (only once):

    poetry run pre-commit install

Run all pre-commit checks manually:

    poetry run pre-commit run --all-file

Auto-fix formatting with Ruff:

    poetry run ruff check --fix

------------------------------------------------------------
EXECUTION (DEVELOPMENT MODE)
------------------------------------------------------------

Run the CLI using Poetry:

    poetry run python -m src.cli --excel company_data.xlsx

Optional arguments:

    --company "C:\Path\To\Company.QBW"
    --output report.json

Example:

    poetry run python -m src.cli --excel company_data.xlsx --output invoice_sync_report.json

------------------------------------------------------------
BUILD WINDOWS .EXE
------------------------------------------------------------

Use PyInstaller inside the Poetry environment:

    poetry run pyinstaller --onefile --name qb-invoice-sync --hidden-import win32timezone --hidden-import win32com.client src/cli.py

The final executable will be located in:

    dist/qb-invoice-sync.exe

------------------------------------------------------------
RUN THE EXE
------------------------------------------------------------

From the dist/ folder:

    .\qb-invoice-sync.exe --excel "..\company_data.xlsx"

With custom output:

    .\qb-invoice-sync.exe --excel "..\company_data.xlsx" --output "invoice_sync_report.json"

------------------------------------------------------------
EXAMPLE JSON OUTPUT
------------------------------------------------------------
```
{
  "status": "success",
  "generated_at": "2025-12-08T19:03:14+00:00",
  "same_invoice": 1,
  "added_invoices": [
    {
      "record_id": "5000",
      "customer": "ExcelOnly",
      "invoice_number": "EX-500",
      "invoice_date": "2023-01-15",
      "invoice_amount": 5040.0
    }
  ],
  "conflicts": [
    {
      "record_id": "7656",
      "reason": "data_mismatch",
      "excel_customer": "NBT Group LTD",
      "excel_invoice_number": "23-0401",
      "excel_invoice_date": "2023-09-13",
      "excel_invoice_amount": 13100.0,
      "qb_customer": "NBT Group LTD",
      "qb_invoice_number": "23-0401",
      "qb_invoice_date": "2023-09-13",
      "qb_invoice_amount": 0.0
    },
    {
      "record_id": "8000",
      "reason": "missing_in_excel",
      "excel_customer": null,
      "excel_invoice_number": null,
      "excel_invoice_date": null,
      "excel_invoice_amount": null,
      "qb_customer": "QBOnlyCustomer",
      "qb_invoice_number": "QB-8000",
      "qb_invoice_date": "2024-02-20",
      "qb_invoice_amount": 0.0
    }
  ],
  "error": null
}
```
