"""
QuickBooks COM gateway for invoices.

Auto-detects a working Item and uses it for invoice creation.
Never depends on 'Services' or 'Sales' existing.
"""

from __future__ import annotations
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List, Tuple
from datetime import date, datetime

try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # type: ignore

from src.models import Invoice

APP_NAME = "Quickbooks Connector"


# -------------------------
# XML Escape Helper
# -------------------------
def _escape_xml(value: str) -> str:
    """Safe XML escaping (ElementTree has no escape())."""
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _require_win32com() -> None:
    if win32com is None:
        raise RuntimeError(
            "pywin32 is required to communicate with QuickBooks Desktop."
        )


@contextmanager
def _qb_session() -> Iterator[Tuple[object, object]]:
    """Open/close a QuickBooks RPC session."""
    _require_win32com()
    session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")  # type: ignore[attr-defined]
    session.OpenConnection2("", APP_NAME, 1)
    ticket = session.BeginSession("", 0)
    try:
        yield session, ticket
    finally:
        try:
            session.EndSession(ticket)
        finally:
            session.CloseConnection()


def _parse_response(raw_xml: str) -> ET.Element:
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing statusCode.")

    code = int(response.get("statusCode", "0"))
    msg = response.get("statusMessage", "") or ""

    if code not in (0, 1):
        raise RuntimeError(f"QuickBooks error {code}: {msg}")

    return root


def _send_qbxml(xml: str) -> ET.Element:
    with _qb_session() as (session, ticket):
        raw = session.ProcessRequest(ticket, xml)  # type: ignore[attr-defined]
    return _parse_response(raw)


# ---------------------------------------------------------
# Fetch Invoices
# ---------------------------------------------------------
def fetch_invoices() -> List[Invoice]:
    qbxml = """<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
 <QBXMLMsgsRq onError="stopOnError">
  <InvoiceQueryRq>
    <IncludeLineItems>false</IncludeLineItems>
  </InvoiceQueryRq>
 </QBXMLMsgsRq>
</QBXML>"""

    root = _send_qbxml(qbxml)

    invoices: List[Invoice] = []
    seen: set[tuple[str, str]] = set()

    for inv in root.findall(".//InvoiceRet"):
        record_id = inv.findtext("Memo") or inv.findtext("RefNumber") or ""
        customer = (inv.findtext("CustomerRef/FullName") or "").strip()
        number = (inv.findtext("RefNumber") or "").strip()
        date_str = inv.findtext("TxnDate")
        amount_str = inv.findtext("Amount") or inv.findtext("Subtotal") or "0"

        if not record_id or not number or not customer:
            continue

        key = (customer.lower(), number)
        if key in seen:
            continue
        seen.add(key)

        try:
            d = date.fromisoformat(date_str) if date_str else date.today()
        except Exception:
            d = date.today()

        try:
            amt = float(amount_str)
        except Exception:
            amt = 0.0

        invoices.append(
            Invoice(
                record_id=record_id,
                customer=customer,
                invoice_number=number,
                invoice_date=d,
                invoice_amount=amt,
                source="quickbooks",
            )
        )

    print(f"Total unique invoices fetched: {len(invoices)}")
    return invoices


# ---------------------------------------------------------
# Item Selection / Creation
# ---------------------------------------------------------
def _get_first_item_name() -> str:
    """Return name of first available item; create AutoItem if none exist."""
    qbxml = """<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
 <QBXMLMsgsRq onError="stopOnError">
  <ItemQueryRq/>
 </QBXMLMsgsRq>
</QBXML>"""

    root = _send_qbxml(qbxml)

    items = (
        root.findall(".//ItemInventoryRet")
        + root.findall(".//ItemNonInventoryRet")
        + root.findall(".//ItemServiceRet")
    )

    if items:
        name = items[0].findtext("Name") or "AutoItem"
        print(f"Using existing item: {name}")
        return name

    # Create fallback item
    create_xml = """<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
 <QBXMLMsgsRq onError="stopOnError">
  <ItemNonInventoryAddRq>
   <ItemNonInventoryAdd>
    <Name>AutoItem</Name>
    <IncomeAccountRef><FullName>Income</FullName></IncomeAccountRef>
   </ItemNonInventoryAdd>
  </ItemNonInventoryAddRq>
 </QBXMLMsgsRq>
</QBXML>"""

    try:
        _send_qbxml(create_xml)
        print("Created AutoItem in QuickBooks.")
    except Exception as exc:
        print(f"Failed to create AutoItem: {exc}")

    return "AutoItem"


# ---------------------------------------------------------
# Customer Creation
# ---------------------------------------------------------
def add_customer(name: str) -> None:
    """Create a customer in QuickBooks if missing (best-effort)."""
    if not name:
        return

    name_xml = _escape_xml(name)

    xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
 <QBXMLMsgsRq onError="continueOnError">
  <CustomerAddRq>
   <CustomerAdd>
    <Name>{name_xml}</Name>
   </CustomerAdd>
  </CustomerAddRq>
 </QBXMLMsgsRq>
</QBXML>"""

    try:
        _send_qbxml(xml)
        print(f"Customer '{name}' verified or created.")
    except Exception as exc:
        print(f"Customer '{name}' creation issue: {exc}")


# ---------------------------------------------------------
# Batch Invoice Add
# ---------------------------------------------------------
def add_invoices_batch(invoices: List[Invoice]) -> None:
    if not invoices:
        print("No invoices to add.")
        return

    item_name = _get_first_item_name()

    existing = fetch_invoices()
    existing_keys = {(i.customer.lower(), i.invoice_number) for i in existing}

    new = [
        inv
        for inv in invoices
        if (inv.customer.lower(), inv.invoice_number) not in existing_keys
    ]

    if not new:
        print("No new invoices to add.")
        return

    # Ensure customers exist
    for cust in sorted({inv.customer for inv in new}):
        add_customer(cust)

    requests: List[str] = []

    for inv in new:
        d = (
            inv.invoice_date.strftime("%Y-%m-%d")
            if isinstance(inv.invoice_date, (date, datetime))
            else str(inv.invoice_date)
        )

        cust = _escape_xml(inv.customer)
        refnum = _escape_xml(inv.invoice_number)
        memo = _escape_xml(inv.record_id)
        item = _escape_xml(item_name)

        requests.append(
            f"""
    <InvoiceAddRq>
      <InvoiceAdd>
        <CustomerRef><FullName>{cust}</FullName></CustomerRef>
        <TxnDate>{d}</TxnDate>
        <RefNumber>{refnum}</RefNumber>
        <Memo>{memo}</Memo>
        <InvoiceLineAdd>
          <ItemRef><FullName>{item}</FullName></ItemRef>
          <Quantity>1</Quantity>
          <Amount>{inv.invoice_amount:.2f}</Amount>
        </InvoiceLineAdd>
      </InvoiceAdd>
    </InvoiceAddRq>
"""
        )

    full_xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
 <QBXMLMsgsRq onError="continueOnError">
  {"".join(requests)}
 </QBXMLMsgsRq>
</QBXML>"""

    try:
        _send_qbxml(full_xml)
        print(f"Successfully added {len(new)} invoices.")
    except Exception as exc:
        print(f"Error adding invoices: {exc}")
