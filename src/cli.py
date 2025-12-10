from __future__ import annotations
import argparse
from pathlib import Path
from src.runner import run_invoice_sync


def main():
    parser = argparse.ArgumentParser(
        prog="qb-invoice-sync",
        description="Sync Excel invoices with QuickBooks and generate a JSON report.",
    )

    parser.add_argument(
        "--excel",
        required=True,
        help="Path to Excel invoice workbook",
    )

    parser.add_argument(
        "--company",
        default="",
        help="QuickBooks company file (leave blank to use currently open file)",
    )

    parser.add_argument(
        "--output",
        default="invoice_sync_report.json",
        help="Output JSON report path",
    )

    args = parser.parse_args()

    report_path = run_invoice_sync(
        excel_path=args.excel,
        company_file=(args.company or None),
        output_path=args.output,
    )

    print(f"Report: {Path(report_path).resolve()}")


if __name__ == "__main__":
    main()
