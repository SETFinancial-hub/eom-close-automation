"""
Parser for Charge Off Excel files.
Reads Sheet2 (raw charge-off data) and produces summary metrics.
"""
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field


@dataclass
class ChargeOffSummary:
    """Aggregated charge-off metrics for the month."""
    note_amount: Decimal = Decimal("0")
    charge_off_amount: Decimal = Decimal("0")       # Net (P&L impact)
    pc_interest_rebate: Decimal = Decimal("0")       # Unearned reversed
    total_charge_off_amt: Decimal = Decimal("0")     # Gross (balance removed)
    prior_month_balance: Decimal = Decimal("0")
    account_count: int = 0
    # By branch
    by_branch: dict = field(default_factory=dict)
    # By portfolio
    by_portfolio: dict = field(default_factory=dict)


def parse_charge_offs(filepath: str) -> tuple[pd.DataFrame, ChargeOffSummary]:
    """
    Parse Charge Off Excel file.
    Returns (raw_dataframe, summary).
    """
    df = pd.read_excel(filepath, sheet_name="Sheet2", engine="openpyxl")

    numeric_cols = [
        "NoteAmount", "PriorMonthBalance", "ChargeOffAmount",
        "PCInterestRebate", "TotalChargeOffAmt"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    summary = ChargeOffSummary(
        note_amount=_to_decimal(df["NoteAmount"].sum()),
        charge_off_amount=_to_decimal(df["ChargeOffAmount"].sum()),
        pc_interest_rebate=_to_decimal(df["PCInterestRebate"].sum()),
        total_charge_off_amt=_to_decimal(df["TotalChargeOffAmt"].sum()),
        prior_month_balance=_to_decimal(df["PriorMonthBalance"].sum()),
        account_count=len(df),
    )

    # Validation: net + unearned should = total
    expected_total = summary.charge_off_amount + summary.pc_interest_rebate
    if abs(expected_total - summary.total_charge_off_amt) > Decimal("1.00"):
        summary._validation_warning = (
            f"Net ({summary.charge_off_amount}) + Unearned ({summary.pc_interest_rebate}) "
            f"= {expected_total} vs Total ({summary.total_charge_off_amt})"
        )

    # By branch
    for branch_id, group in df.groupby("BranchID"):
        summary.by_branch[int(branch_id)] = {
            "charge_off_amount": _to_decimal(group["ChargeOffAmount"].sum()),
            "pc_interest_rebate": _to_decimal(group["PCInterestRebate"].sum()),
            "total_charge_off_amt": _to_decimal(group["TotalChargeOffAmt"].sum()),
            "count": len(group),
        }

    # By portfolio
    for port_id, group in df.groupby("PortfolioID"):
        summary.by_portfolio[int(port_id)] = {
            "charge_off_amount": _to_decimal(group["ChargeOffAmount"].sum()),
            "total_charge_off_amt": _to_decimal(group["TotalChargeOffAmt"].sum()),
            "count": len(group),
        }

    return df, summary


def _to_decimal(value) -> Decimal:
    return Decimal(str(round(float(value), 2))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
