"""
Parser for Metacorp charge-off sale closing statements.
Data is extracted from the PNG image — values are hardcoded per month
or can be manually entered. Future: OCR integration.
"""
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass
import json
from pathlib import Path


@dataclass
class MetacorpSaleData:
    """Data from Metacorp charge-off sale closing statement."""
    agreement_date: str = ""
    purchaser: str = "Metacorp, LLC"
    seller: str = "SET Financial Corporation"
    num_accounts: int = 0
    total_current_balance: Decimal = Decimal("0")
    purchase_price_pct: Decimal = Decimal("0")
    transfer_amount: Decimal = Decimal("0")
    file_creation_date: str = ""
    closing_date: str = ""


def create_metacorp_sale(
    num_accounts: int,
    total_balance: Decimal,
    purchase_pct: Decimal,
    transfer_amount: Decimal,
    agreement_date: str = "",
    closing_date: str = "",
) -> MetacorpSaleData:
    """Create Metacorp sale data from manual inputs or OCR."""
    return MetacorpSaleData(
        agreement_date=agreement_date,
        num_accounts=num_accounts,
        total_current_balance=total_balance,
        purchase_price_pct=purchase_pct,
        transfer_amount=transfer_amount,
        closing_date=closing_date,
    )


def load_from_json(json_path: str) -> MetacorpSaleData:
    """Load Metacorp sale data from a JSON file."""
    with open(json_path, 'r') as f:
        data = json.load(f)
    return MetacorpSaleData(
        agreement_date=data.get("agreement_date", ""),
        num_accounts=int(data["num_accounts"]),
        total_current_balance=Decimal(data["total_balance"]),
        purchase_price_pct=Decimal(data["purchase_pct"]),
        transfer_amount=Decimal(data["transfer_amount"]),
        file_creation_date=data.get("file_creation_date", ""),
        closing_date=data.get("closing_date", ""),
    )


# January 2026 data (from PNG closing statement)
JANUARY_2026_SALE = MetacorpSaleData(
    agreement_date="2/06/2026",
    num_accounts=101,
    total_current_balance=Decimal("192157.63"),
    purchase_price_pct=Decimal("5.10"),
    transfer_amount=Decimal("9800.04"),
    file_creation_date="2/06/2026",
    closing_date="2/06/2026",
)
