"""
Parser for Collection Register Excel files.
Reads Sheet2 (raw transaction data) and produces summary metrics.
"""
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class CollectionSummary:
    """Aggregated collection metrics for the month."""
    total_collected: Decimal = Decimal("0")
    principal: Decimal = Decimal("0")
    interest_collected: Decimal = Decimal("0")
    interest_rebate: Decimal = Decimal("0")
    late_fees: Decimal = Decimal("0")
    nsf_fees: Decimal = Decimal("0")
    insurance_rebate: Decimal = Decimal("0")
    balance_renewed: Decimal = Decimal("0")
    recovery: Decimal = Decimal("0")
    amount_to_refund: Decimal = Decimal("0")
    allotment_fee: Decimal = Decimal("0")
    cash_received: Decimal = Decimal("0")
    transaction_count: int = 0
    # By branch
    by_branch: dict = field(default_factory=dict)
    # By portfolio
    by_portfolio: dict = field(default_factory=dict)


def parse_collection_register(filepath: str) -> tuple[pd.DataFrame, CollectionSummary]:
    """
    Parse Collection Register Excel file.
    Returns (raw_dataframe, summary).
    """
    # Read Sheet2 (raw data)
    df = pd.read_excel(filepath, sheet_name="Sheet2", engine="openpyxl")

    # Ensure numeric columns
    numeric_cols = [
        "TotalCollected", "Principal", "InterestCollected", "InterestRebate",
        "LateFees", "NSFFees", "InsuranceRebate", "BalanceRenewed",
        "Recovery", "AmountToRefund", "AllotmentFee", "CashReceived"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Build summary
    summary = CollectionSummary(
        total_collected=_to_decimal(df["TotalCollected"].sum()),
        principal=_to_decimal(df["Principal"].sum()),
        interest_collected=_to_decimal(df["InterestCollected"].sum()),
        interest_rebate=_to_decimal(df["InterestRebate"].sum()),
        late_fees=_to_decimal(df["LateFees"].sum()),
        nsf_fees=_to_decimal(df["NSFFees"].sum()),
        insurance_rebate=_to_decimal(df["InsuranceRebate"].sum()),
        balance_renewed=_to_decimal(df["BalanceRenewed"].sum()),
        recovery=_to_decimal(df["Recovery"].sum()),
        amount_to_refund=_to_decimal(df["AmountToRefund"].sum()),
        allotment_fee=_to_decimal(df["AllotmentFee"].sum()),
        cash_received=_to_decimal(df["CashReceived"].sum()),
        transaction_count=len(df),
    )

    # By branch
    for branch_id, group in df.groupby("BranchID"):
        summary.by_branch[int(branch_id)] = {
            "total_collected": _to_decimal(group["TotalCollected"].sum()),
            "principal": _to_decimal(group["Principal"].sum()),
            "interest_collected": _to_decimal(group["InterestCollected"].sum()),
            "late_fees": _to_decimal(group["LateFees"].sum()),
            "nsf_fees": _to_decimal(group["NSFFees"].sum()),
            "recovery": _to_decimal(group["Recovery"].sum()),
            "cash_received": _to_decimal(group["CashReceived"].sum()),
            "count": len(group),
        }

    # By portfolio
    for port_id, group in df.groupby("PortfolioID"):
        summary.by_portfolio[int(port_id)] = {
            "total_collected": _to_decimal(group["TotalCollected"].sum()),
            "principal": _to_decimal(group["Principal"].sum()),
            "count": len(group),
        }

    return df, summary


def _to_decimal(value) -> Decimal:
    """Convert float to Decimal with 2 decimal places."""
    return Decimal(str(round(float(value), 2))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
