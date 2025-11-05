# stdlib only CLI
from __future__ import annotations
import argparse
from pathlib import Path
from src.runner import run_invoice_sync

def main():
    p = argparse.ArgumentParser(
        prog="qb-invoice-sync",
        description="Sync company Excel invoices with QuickBooks and write a JSON report."
    )
    p.add_argument("--excel", required=True, help="Path to company .xlsx file")
    p.add_argument("--sheet", default="account credit vendor",
                   help="Worksheet name to use (default: 'account credit vendor')")
    p.add_argument("--company", default="", help="QuickBooks company file path (leave blank to use open file)")
    p.add_argument("--report", default="invoice_sync_report.json", help="Output JSON report path")
    p.add_argument("--dry-run", action="store_true",
                   help="Compare only: do NOT add Excel-only invoices to QuickBooks")
    args = p.parse_args()

    report_path = run_invoice_sync(
        excel_path=args.excel,
        sheet_name=args.sheet,
        company_file=(args.company or None),
        report_path=args.report,
        dry_run=args.dry_run,
    )
    print(f"Report: {Path(report_path).resolve()}")

if __name__ == "__main__":
    main()
