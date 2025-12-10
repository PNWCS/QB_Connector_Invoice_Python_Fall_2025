"""
Production runner for Excel ↔ QuickBooks invoice comparison.

Produces JSON in required structure:
{
    "status": "success",
    "generated_at": "...",
    "added_invoices": [...],
    "conflicts": [...],
    "same_invoice": X,
    "error": null
}
"""

from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Any

from src.excel_reader import extract_invoices
from src.qb_gateway import fetch_invoices, add_invoices_batch
from src.comparer import compare_invoices
from src.models import Invoice, Conflict


def ts() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# ----------------------------
# Base Invoice Serializer
# ----------------------------
def serialize(inv: Invoice) -> Dict[str, Any]:
    return {
        "record_id": inv.record_id,
        "customer": inv.customer,
        "invoice_number": inv.invoice_number,
        "invoice_date": inv.invoice_date.isoformat(),
        "invoice_amount": inv.invoice_amount,
    }


# ----------------------------
# Conflict Serializers
# ----------------------------
def serialize_flat_conflict(c: Conflict) -> Dict[str, Any]:
    e = c.excel_data
    q = c.qb_data

    return {
        "record_id": c.record_id,
        "reason": c.reason,
        "excel_customer": e.customer if e else None,
        "excel_invoice_number": e.invoice_number if e else None,
        "excel_invoice_date": e.invoice_date.isoformat() if e else None,
        "excel_invoice_amount": e.invoice_amount if e else None,
        "qb_customer": q.customer if q else None,
        "qb_invoice_number": q.invoice_number if q else None,
        "qb_invoice_date": q.invoice_date.isoformat() if q else None,
        "qb_invoice_amount": q.invoice_amount if q else None,
    }


def serialize_missing_in_excel(inv: Invoice) -> Dict[str, Any]:
    return {
        "record_id": inv.record_id,
        "reason": "missing_in_excel",
        "excel_customer": None,
        "excel_invoice_number": None,
        "excel_invoice_date": None,
        "excel_invoice_amount": None,
        "qb_customer": inv.customer,
        "qb_invoice_number": inv.invoice_number,
        "qb_invoice_date": inv.invoice_date.isoformat(),
        "qb_invoice_amount": inv.invoice_amount,
    }


# ----------------------------
# MAIN PROCESS
# ----------------------------
def _execute_sync(excel_path: Path, output_path: Path) -> Path:
    conflicts: List[Dict[str, Any]] = []  # mypy required annotation
    added: List[Invoice] = []  # mypy required annotation

    payload: Dict[str, Any] = {
        "status": "success",
        "generated_at": ts(),
        "same_invoice": 0,
        "added_invoices": [],
        "conflicts": conflicts,
        "error": None,
    }

    try:
        excel_invoices = extract_invoices(excel_path)
        qb_invoices = fetch_invoices()

        report = compare_invoices(excel_invoices, qb_invoices)

        # Add Excel-only invoices
        if report.excel_only:
            add_invoices_batch(report.excel_only)
            added = report.excel_only

        # Data mismatch conflicts
        conflicts.extend(serialize_flat_conflict(c) for c in report.conflicts)

        # QB-only → missing_in_excel conflicts
        conflicts.extend(serialize_missing_in_excel(i) for i in report.qb_only)

        # Fill JSON
        payload["added_invoices"] = [serialize(i) for i in added]
        payload["same_invoice"] = len(report.same_invoice)

    except Exception as exc:
        payload["status"] = "error"
        payload["error"] = str(exc)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=4)

    return output_path


def run_invoice_sync(
    excel_path: str,
    company_file=None,  # unused but kept for compatibility
    output_path: str = "invoice_sync_report.json",
) -> Path:
    return _execute_sync(Path(excel_path), Path(output_path))


if __name__ == "__main__":
    run_invoice_sync("company_data.xlsx")
    print("invoice_sync_report.json written.")
