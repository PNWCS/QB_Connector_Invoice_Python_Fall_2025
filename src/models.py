"""Domain models for invoice synchronisation between Excel and QuickBooks."""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Literal, Optional
from datetime import date

SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["data_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class Invoice:
    """Represents a single invoice from Excel or QuickBooks."""

    record_id: str
    customer: str
    invoice_number: str
    invoice_date: date
    invoice_amount: float
    source: SourceLiteral

    def __str__(self) -> str:
        return (
            f"Invoice(id={self.record_id}, customer={self.customer}, "
            f"number={self.invoice_number}, date={self.invoice_date}, "
            f"amount={self.invoice_amount}, source={self.source})"
        )


@dataclass(slots=True)
class Conflict:
    """Describes an ID present in both sources but with mismatching data."""

    record_id: str
    excel_data: Optional[Invoice]
    qb_data: Optional[Invoice]
    reason: ConflictReason


@dataclass(slots=True)
class ComparisonReport:
    """
    Container for comparison results:
      - same_invoice: matching invoices in Excel and QB
      - excel_only: invoices only in Excel
      - qb_only: invoices only in QuickBooks
      - conflicts: invoices with same ID but mismatching fields
    """

    same_invoice: list[Invoice] = field(default_factory=list)
    excel_only: list[Invoice] = field(default_factory=list)
    qb_only: list[Invoice] = field(default_factory=list)
    conflicts: list[Conflict] = field(default_factory=list)


__all__ = [
    "Invoice",
    "Conflict",
    "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]
