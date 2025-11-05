"""Comparison helpers for invoices.

This module compares invoice data from Excel and QuickBooks, identifying
which invoices are identical, missing in one source, or have mismatched data.
"""

from __future__ import annotations
from typing import Dict, List
from src.models import Invoice, Conflict, ComparisonReport


def compare_invoices(
    excel_invoices: List[Invoice],
    qb_invoices: List[Invoice],
) -> ComparisonReport:
    """Compare Excel vs QuickBooks invoices and detect discrepancies.

    This function reconciles two datasets — one from Excel, one from QuickBooks —
    by comparing their record identifiers and key fields such as invoice number,
    amount, and date.

    The result categorizes invoices as:
      - same_same: identical in both sources
      - excel_only: exists only in Excel
      - qb_only: exists only in QuickBooks
      - conflicts: same ID, but differing data values

    **Input Parameters:**
    :param excel_invoices: List of Invoice objects read from Excel
    :param qb_invoices: List of Invoice objects read from QuickBooks

    **Return Value:**
    :return: ComparisonReport containing categorized comparison results.
    """

    report = ComparisonReport()

    # Create fast lookup maps by record_id (Excel Child ID ↔ QB Memo)
    excel_by_id: Dict[str, Invoice] = {inv.record_id: inv for inv in excel_invoices}
    qb_by_id: Dict[str, Invoice] = {inv.record_id: inv for inv in qb_invoices}

    # Combine all record IDs from both sources
    all_ids = set(excel_by_id.keys()) | set(qb_by_id.keys())

    for rid in all_ids:
        e_inv = excel_by_id.get(rid)
        q_inv = qb_by_id.get(rid)

        # Case 1: Exists in both
        if e_inv and q_inv:
            same_number = e_inv.invoice_number == q_inv.invoice_number
            same_date = e_inv.invoice_date == q_inv.invoice_date
            same_amount = abs(e_inv.invoice_amount - q_inv.invoice_amount) < 0.01

            if same_number and same_date and same_amount:
                report.same_same.append(e_inv)
            else:
                report.conflicts.append(
                    Conflict(
                        record_id=rid,
                        excel_data=e_inv,
                        qb_data=q_inv,
                        reason="data_mismatch",
                    )
                )

        # Case 2: Exists only in Excel
        elif e_inv and not q_inv:
            report.excel_only.append(e_inv)

        # Case 3: Exists only in QuickBooks
        elif q_inv and not e_inv:
            report.qb_only.append(q_inv)

    return report


__all__ = ["compare_invoices"]


if __name__ == "__main__":
    from datetime import date

    # Manual test data
    excel_data = [
        Invoice("A1", "Test6", "INV001", date(2024, 1, 1), 100.0, "excel"),
        Invoice("A2", "Test7", "INV002", date(2024, 1, 2), 200.0, "excel"),
    ]

    qb_data = [
        Invoice("A1", "Test8", "INV001", date(2024, 1, 1), 100.0, "quickbooks"),
        Invoice("A3", "Test9", "INV003", date(2024, 1, 3), 300.0, "quickbooks"),
    ]

    report = compare_invoices(excel_data, qb_data)

    print(
        f"Same: {len(report.same_same)} | "
        f"Excel-only: {len(report.excel_only)} | "
        f"QB-only: {len(report.qb_only)} | "
        f"Conflicts: {len(report.conflicts)}"
    )
