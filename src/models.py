"""Domain models for invoice synchronisation between Excel and QuickBooks."""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Literal, Optional
from datetime import date

# Define literal types for clarity
SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["data_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class Invoice:
    """Represents a single invoice from either Excel or QuickBooks."""

    record_id: str           # Excel "Child ID" ↔ QuickBooks "Memo"
    customer: str            # Customer name
    invoice_number: str      # Invoice reference number
    invoice_date: date       # Invoice date
    invoice_amount: float    # Invoice total amount
    source: SourceLiteral    # Origin of this record ("excel" or "quickbooks")

    def __str__(self) -> str:
        """Readable string representation of this invoice."""
        return (
            f"Invoice(id={self.record_id}, customer={self.customer}, "
            f"number={self.invoice_number}, date={self.invoice_date}, "
            f"amount={self.invoice_amount}, source={self.source})"
        )


@dataclass(slots=True)
class Conflict:
    """Describes a mismatch or missing invoice between Excel and QuickBooks."""

    record_id: str
    excel_data: Optional[Invoice]
    qb_data: Optional[Invoice]
    reason: ConflictReason


@dataclass(slots=True)
class ComparisonReport:
    """Groups comparison results between Excel and QuickBooks."""

    same_same: list[Invoice] = field(default_factory=list)
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
