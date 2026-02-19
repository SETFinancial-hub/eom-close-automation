"""
Parser for DPV LLC / Bank of America monthly invoices.
Extracts interest expense and principal balance from PDF.
"""
import pdfplumber
import re
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass


@dataclass
class DPVStatementData:
    """Parsed data from DPV/BofA statement."""
    invoice_date: str = ""
    invoice_number: str = ""
    customer_name: str = "DPV LLC"
    total_due: Decimal = Decimal("0")
    principal_balance: Decimal = Decimal("0")
    invoice_due_date: str = ""
    # Accrual details
    accrual_entries: list = None  # List of daily accrual entries
    interest_payment: Decimal = Decimal("0")  # Mid-month interest payment

    def __post_init__(self):
        if self.accrual_entries is None:
            self.accrual_entries = []


def parse_dpv_statement(filepath: str) -> DPVStatementData:
    """Parse DPV LLC Bank of America PDF invoice."""
    data = DPVStatementData()

    with pdfplumber.open(filepath) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # Extract invoice date
    m = re.search(r"InvoiceDate:\s*([\d\-A-Za-z]+)", full_text)
    if m:
        data.invoice_date = m.group(1).strip()

    # Extract invoice number
    m = re.search(r"Invoice\s+([A-Z0-9]+)", full_text)
    if m:
        data.invoice_number = m.group(1).strip()

    # Extract total due
    m = re.search(r"TotalDue:\s*USD\s*([\d,]+\.?\d*)", full_text)
    if m:
        data.total_due = _to_decimal(Decimal(m.group(1).replace(",", "")))

    # Extract due date
    m = re.search(r"InvoiceDueDate:\s*([\d\-A-Za-z]+)", full_text)
    if m:
        data.invoice_due_date = m.group(1).strip()

    # Extract principal balance (the accrued balance amount)
    # Look for the recurring USD amount in accrual lines
    balances = re.findall(r"USD\s*([\d,]+\.?\d+)\s+USD", full_text)
    if balances:
        # The most common balance is the principal
        from collections import Counter
        balance_counts = Counter(balances)
        most_common = balance_counts.most_common(1)[0][0]
        data.principal_balance = _to_decimal(Decimal(most_common.replace(",", "")))

    # Extract interest payment (mid-month)
    m = re.search(r"InterestPayment\s+([\d,]+\.?\d*)", full_text)
    if m:
        data.interest_payment = _to_decimal(Decimal(m.group(1).replace(",", "")))

    return data


def _to_decimal(value) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
