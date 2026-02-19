"""
Parser for PAT LLC Series 47 (Pier) monthly statements.
Extracts key facility data from PDF.
"""
import pdfplumber
import re
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass


@dataclass
class PierStatementData:
    """Parsed data from Pier facility statement."""
    borrower: str = ""
    reporting_period: str = ""
    interest_rate_base: Decimal = Decimal("0")  # 13.50%
    sofr_floor: Decimal = Decimal("0")          # 4.00%
    beginning_balance: Decimal = Decimal("0")
    advances: Decimal = Decimal("0")
    principal_payments: Decimal = Decimal("0")
    ending_balance: Decimal = Decimal("0")
    accrued_interest_due: Decimal = Decimal("0")
    origination_fee_outstanding: Decimal = Decimal("0")
    default_reserve_balance_due: Decimal = Decimal("0")
    total_amount_due: Decimal = Decimal("0")
    daily_accrual_rate: Decimal = Decimal("0")  # daily interest amount


def parse_pier_statement(filepath: str) -> PierStatementData:
    """Parse Pier facility PDF statement."""
    data = PierStatementData()

    with pdfplumber.open(filepath) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # Extract borrower
    m = re.search(r"Borrower:\s*(.+?)(?:\n|Reporting)", full_text)
    if m:
        data.borrower = m.group(1).strip()

    # Extract reporting period
    m = re.search(r"period:\s*(.+?)(?:\n|Days)", full_text)
    if m:
        data.reporting_period = m.group(1).strip()

    # Extract interest rate
    m = re.search(r"Interest:\s*([\d.]+)%\s*\+\s*SOFR\s*\(([\d.]+)%\s*floor\)", full_text)
    if m:
        data.interest_rate_base = Decimal(m.group(1))
        data.sofr_floor = Decimal(m.group(2))

    # Extract key financial figures
    patterns = {
        "beginning_balance": r"Beginning Balance\s+\$?([\d,]+\.?\d*)",
        "ending_balance": r"Ending Balance\s+\$?([\d,]+\.?\d*)",
        "accrued_interest_due": r"Accrued Interest Due\s+\$?([\d,]+\.?\d*)",
        "origination_fee_outstanding": r"Origination Fee Outstanding\s+\$?([\d,]+\.?\d*)",
        "default_reserve_balance_due": r"Default Reserve Balance Due\s+\$?([\d,]+\.?\d*)",
        "total_amount_due": r"Total Amount Due\s+\$?([\d,]+\.?\d*)",
    }

    for field_name, pattern in patterns.items():
        m = re.search(pattern, full_text)
        if m:
            value = m.group(1).replace(",", "")
            setattr(data, field_name, _to_decimal(Decimal(value)))

    # Check for advances/payments (may show as "-")
    m = re.search(r"Advances in Period\s+[\$]?([\d,]+\.?\d*|-)", full_text)
    if m and m.group(1) != "-":
        data.advances = _to_decimal(Decimal(m.group(1).replace(",", "")))

    m = re.search(r"Principal Payments in Period\s+[\$]?([\d,]+\.?\d*|-)", full_text)
    if m and m.group(1) != "-":
        data.principal_payments = _to_decimal(Decimal(m.group(1).replace(",", "")))

    # Extract daily accrual from the daily schedule
    m = re.search(r"\d{1,2}/\d{1,2}/\d{4}\s+[\d.]+%\s+\$([\d,]+\.?\d*)", full_text)
    if m:
        data.daily_accrual_rate = _to_decimal(Decimal(m.group(1).replace(",", "")))

    return data


def _to_decimal(value) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
