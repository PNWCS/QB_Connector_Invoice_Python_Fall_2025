"""Main orchestration for QuickBooks Invoice Sync (Full Auto-Sync)."""

from __future__ import annotations
from pathlib import Path
import json
from datetime import datetime, timezone
from typing import Dict, Any, List

from openpyxl import load_workbook
from src.excel_reader import extract_invoices
from src.qb_gateway import fetch_invoices, add_invoice
from src.comparer import compare_invoices
from src.models import Invoice, Conflict, ComparisonReport


def _iso_ts() -> str:
    """Return UTC timestamp in ISO format."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def update_excel_with_qb_only(excel_path: str, qb_only_invoices: List[Invoice], sheet_name: str = "account credit vendor"):
    """Append QB-only invoices to Excel so Excel stays synced."""
    if not qb_only_invoices:
        return

    print(f"🧩 Updating Excel with {len(qb_only_invoices)} QuickBooks-only records...")

    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Worksheet '{sheet_name}' not found.")
    sheet = wb[sheet_name]

    for inv in qb_only_invoices:
        sheet.append([
            "",  # Parent ID
            inv.record_id,
            inv.invoice_number,
            inv.invoice_date.isoformat() if inv.invoice_date else "",
            "", "", "", "",
            inv.invoice_amount,
            inv.customer,
            "", "", "", "",
            inv.invoice_amount,
            "",
        ])

    wb.save(excel_path)
    wb.close()
    print(f"✅ Excel updated successfully with {len(qb_only_invoices)} new QB rows.")


def run_invoice_sync(
    excel_path: str,
    sheet_name: str = "account credit vendor",
    company_file: str | None = None,
    report_path: str | None = None,
    dry_run: bool = False,
) -> Path:
    """
    Full bi-directional synchronization:
      - Excel → QB (new or updated)
      - QB → Excel (missing in Excel)
      - Generates JSON summary
    """
    report_file = Path(report_path or "invoice_sync_report.json")

    payload: Dict[str, Any] = {
        "status": "success",
        "generated_at": _iso_ts(),
        "added_invoices": [],
        "updated_invoices": [],
        "conflicts": [],
        "error": None,
    }

    try:
        # Step 1️⃣ Read Excel
        excel_invoices = extract_invoices(excel_path, sheet_name=sheet_name)
        print(f"✅ Loaded {len(excel_invoices)} invoices from Excel")

        # Step 2️⃣ Read QuickBooks
        try:
            qb_invoices = fetch_invoices(company_file)
            print(f"✅ Fetched {len(qb_invoices)} invoices from QuickBooks")
        except Exception as e:
            print(f"⚠️ QuickBooks fetch failed: {e}")
            qb_invoices = []

        # Step 3️⃣ Compare
        comp: ComparisonReport = compare_invoices(excel_invoices, qb_invoices)
        print(
            f"🔍 Compare — Same: {len(comp.same_same)}, "
            f"Excel-only: {len(comp.excel_only)}, QB-only: {len(comp.qb_only)}, "
            f"Conflicts: {len(comp.conflicts)}"
        )

        added: List[Invoice] = []
        updated: List[Invoice] = []

        # Step 4️⃣ Add Excel-only invoices to QuickBooks
        for inv in comp.excel_only:
            if not inv.invoice_number or not inv.customer or not inv.invoice_amount:
                print(f"⚠️ Skipping incomplete row: {inv}")
                continue
            if dry_run:
                print(f"🧪 Dry-run: would add invoice {inv.invoice_number}")
                continue
            try:
                print(f"📤 Adding new invoice {inv.invoice_number} for {inv.customer}...")
                added_inv = add_invoice(company_file, inv)
                added.append(added_inv)
                print(f"✅ Added invoice {inv.invoice_number}")
            except Exception as e:
                print(f"⚠️ Failed to add invoice {inv.invoice_number}: {e}")

        # Step 5️⃣ Update changed invoices automatically
        for conflict in comp.conflicts:
            excel_inv = conflict.excel_data
            qb_inv = conflict.qb_data
            if not excel_inv or not qb_inv:
                continue
            if dry_run:
                print(f"🧪 Dry-run: would update invoice {excel_inv.invoice_number}")
                continue
            try:
                print(f"🔁 Updating invoice {excel_inv.invoice_number} (changes detected)...")
                add_invoice(company_file, excel_inv)
                updated.append(excel_inv)
                print(f"✅ Updated invoice {excel_inv.invoice_number}")
            except Exception as e:
                print(f"⚠️ Failed to update invoice {excel_inv.invoice_number}: {e}")

        # Step 6️⃣ Write QB-only invoices back to Excel
        try:
            update_excel_with_qb_only(excel_path, comp.qb_only, sheet_name)
        except Exception as e:
            print(f"⚠️ Failed to update Excel: {e}")

        # Step 7️⃣ Prepare JSON report
        payload["added_invoices"] = [
            {"record_id": i.record_id, "name": i.invoice_number, "source": i.source}
            for i in added
        ]
        payload["updated_invoices"] = [
            {"record_id": i.record_id, "name": i.invoice_number, "source": i.source}
            for i in updated
        ]
        payload["conflicts"] = [
            {
                "record_id": c.record_id,
                "excel_name": c.excel_data.invoice_number if c.excel_data else None,
                "qb_name": c.qb_data.invoice_number if c.qb_data else None,
                "reason": c.reason,
            }
            for c in comp.conflicts
        ]

    except Exception as e:
        payload["status"] = "error"
        payload["error"] = str(e)
        print(f"❌ Sync failed: {e}")

    # Step 8️⃣ Write report
    with open(report_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
    print(f"📝 Report written to {report_file}")
    return report_file


if __name__ == "__main__":
    run_invoice_sync("company_data.xlsx")
