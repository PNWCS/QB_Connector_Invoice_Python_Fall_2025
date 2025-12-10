"""Comparison helpers for invoices."""

from __future__ import annotations
from typing import Dict, List
from src.models import Invoice, Conflict, ComparisonReport


def compare_invoices(
    excel_invoices: List[Invoice],
    qb_invoices: List[Invoice],
) -> ComparisonReport:
    report = ComparisonReport()

    excel_by_id: Dict[str, Invoice] = {inv.record_id: inv for inv in excel_invoices}
    qb_by_id: Dict[str, Invoice] = {inv.record_id: inv for inv in qb_invoices}

    all_ids = set(excel_by_id.keys()) | set(qb_by_id.keys())

    for rid in all_ids:
        excel_inv = excel_by_id.get(rid)
        qb_inv = qb_by_id.get(rid)

        if excel_inv and qb_inv:
            num_match = excel_inv.invoice_number == qb_inv.invoice_number
            date_match = excel_inv.invoice_date == qb_inv.invoice_date
            amt_match = abs(excel_inv.invoice_amount - qb_inv.invoice_amount) < 0.01

            if num_match and date_match and amt_match:
                report.same_invoice.append(excel_inv)
            else:
                report.conflicts.append(
                    Conflict(
                        record_id=rid,
                        excel_data=excel_inv,
                        qb_data=qb_inv,
                        reason="data_mismatch",
                    )
                )

        elif excel_inv and not qb_inv:
            report.excel_only.append(excel_inv)

        elif qb_inv and not excel_inv:
            report.qb_only.append(qb_inv)

    return report


__all__ = ["compare_invoices"]
