"""QuickBooks COM gateway for invoices.

This module communicates with QuickBooks Desktop via the QBXML Request Processor
COM interface (pywin32). It allows querying invoices and, optionally, creating
customers and items when needed.
"""

from __future__ import annotations
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List
from datetime import date, datetime

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

from src.models import Invoice

APP_NAME = "QuickBooks Invoice Connector"



def _require_win32com() -> None:
    """Ensure the pywin32 library is available."""
    if win32com is None:
        raise RuntimeError("pywin32 is required to communicate with QuickBooks Desktop.")


@contextmanager
def _qb_session() -> Iterator[tuple[object, object]]:
    """Context manager for safely opening and closing a QuickBooks session."""
    _require_win32com()
    session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
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
    """Parse and validate a QuickBooks QBXML response."""
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing status information.")
    status_code = int(response.get("statusCode", "0"))
    status_message = response.get("statusMessage", "")
    if status_code not in (0, 1):  # 0=OK, 1=No results
        raise RuntimeError(f"QuickBooks error {status_code}: {status_message}")
    return root


def _send_qbxml(qbxml: str) -> ET.Element:
    """Send QBXML to QuickBooks and return parsed XML root."""
    with _qb_session() as (session, ticket):
        print(f"Sending QBXML:\n{qbxml}")
        raw_response = session.ProcessRequest(ticket, qbxml)
        print(f"Received QBXML:\n{raw_response}")
    return _parse_response(raw_response)



def fetch_invoices(company_file: str | None = None) -> List[Invoice]:
    """Fetch all invoices from QuickBooks (deduplicated by invoice number + customer)."""
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <InvoiceQueryRq>\n"
        "      <IncludeLineItems>false</IncludeLineItems>\n"
        "    </InvoiceQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    root = _send_qbxml(qbxml)
    invoices: List[Invoice] = []
    seen_keys = set()  # <-- Track unique (customer, invoice_number)

    for inv_ret in root.findall(".//InvoiceRet"):
        record_id = inv_ret.findtext("Memo") or inv_ret.findtext("RefNumber")
        customer = (inv_ret.findtext("CustomerRef/FullName") or "").strip()
        invoice_number = (inv_ret.findtext("RefNumber") or "").strip()
        date_str = inv_ret.findtext("TxnDate")
        invoice_date = date.fromisoformat(date_str) if date_str else date.today()
        amount_str = inv_ret.findtext("Subtotal") or inv_ret.findtext("Amount") or "0"

        try:
            invoice_amount = float(amount_str)
        except ValueError:
            invoice_amount = 0.0

        if not record_id or not invoice_number or not customer:
            continue

        key = (customer.lower(), invoice_number)
        if key in seen_keys:
            continue  # Skip duplicates
        seen_keys.add(key)

        invoices.append(
            Invoice(
                record_id=record_id,
                customer=customer,
                invoice_number=invoice_number,
                invoice_date=invoice_date,
                invoice_amount=invoice_amount,
                source="quickbooks",
            )
        )

    print(f"Total unique invoices fetched: {len(invoices)}")
    return invoices




def add_customer(customer_name: str) -> None:
    """Create a customer in QuickBooks if missing."""
    if not customer_name:
        raise ValueError("Customer name is required to create a customer.")
    qbxml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CustomerAddRq>
      <CustomerAdd>
        <Name>{customer_name}</Name>
      </CustomerAdd>
    </CustomerAddRq>
  </QBXMLMsgsRq>
</QBXML>"""
    _send_qbxml(qbxml)
    print(f" Created customer '{customer_name}' in QuickBooks.")


def add_item_service(item_name: str):
    """Create a service item in QuickBooks if it doesn't exist."""
    qbxml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <ItemServiceAddRq>
      <ItemServiceAdd>
        <Name>{item_name}</Name>
        <SalesOrPurchase>
          <Desc>Auto-created by Invoice Sync</Desc>
          <AccountRef><FullName>Sales</FullName></AccountRef>
        </SalesOrPurchase>
      </ItemServiceAdd>
    </ItemServiceAddRq>
  </QBXMLMsgsRq>
</QBXML>"""
    _send_qbxml(qbxml)
    print(f"Created missing service item '{item_name}' in QuickBooks.")


if __name__ == "__main__":  
    try:
        invoices = fetch_invoices()
        for inv in invoices:
            print(inv)
        print(f" Total invoices fetched: {len(invoices)}")
    except Exception as e:
        print(f" Error fetching invoices: {e}")
