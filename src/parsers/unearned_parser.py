"""
Parser for Unearned Register Excel files.
Reads Sheet2 (loan-level unearned income schedule) and produces summary metrics.
This is the most critical file — it drives finance income recognition and insurance earnings.
"""
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field


@dataclass
class UnearnedSummary:
    """Aggregated unearned income balances at month-end."""
    # Interest unearned
    unearned_new_interest: Decimal = Decimal("0")
    unearned_existing_interest: Decimal = Decimal("0")
    total_unearned_interest: Decimal = Decimal("0")

    # Finance fees unearned
    unearned_new_finance_fees: Decimal = Decimal("0")
    unearned_existing_finance_fees: Decimal = Decimal("0")
    total_unearned_finance_fees: Decimal = Decimal("0")

    # Insurance unearned (by type)
    unearned_credit_life: Decimal = Decimal("0")
    unearned_disability: Decimal = Decimal("0")
    unearned_iui: Decimal = Decimal("0")
    unearned_property: Decimal = Decimal("0")
    unearned_vsi: Decimal = Decimal("0")
    total_unearned_insurance: Decimal = Decimal("0")

    # Portfolio metrics
    total_current_balance: Decimal = Decimal("0")
    original_finance_charge: Decimal = Decimal("0")
    interest_collected_month: Decimal = Decimal("0")
    loan_count: int = 0

    # By branch
    by_branch: dict = field(default_factory=dict)
    # By portfolio
    by_portfolio: dict = field(default_factory=dict)


def parse_unearned_register(filepath: str) -> tuple[pd.DataFrame, UnearnedSummary]:
    """
    Parse Unearned Register Excel file.
    Returns (raw_dataframe, summary).
    """
    df = pd.read_excel(filepath, sheet_name="Sheet2", engine="openpyxl")

    numeric_cols = [
        "CurrentBalance", "OriginalFinanceCharge",
        "InterestCollectedMonth", "InterestCollectedToDate",
        "UnearnedNewInterest", "UnearnedExistingInterest",
        "UnearnedNewFinanceFees", "UnearnedExistingFinanceFees",
        "UnearnedNewCreditLife", "UnearnedExistingCreditLife",
        "UnearnedNewDisability", "UnearnedExistingDisability",
        "UnearnedNewIUI", "UnearnedExistingIUI",
        "UnearnedNewProperty", "UnearnedExistingProperty",
        "UnearnedNewVSI", "UnearnedExistingVSI",
        "OriginalCreditLife", "OriginalDisability", "OriginalIUI",
        "OriginalProperty", "OriginalVSI", "OriginalFinanceFees",
        "AmountFinanced",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Compute totals
    unearned_new_int = _to_decimal(df["UnearnedNewInterest"].sum())
    unearned_exist_int = _to_decimal(df["UnearnedExistingInterest"].sum())

    unearned_new_ff = _to_decimal(df["UnearnedNewFinanceFees"].sum())
    unearned_exist_ff = _to_decimal(df["UnearnedExistingFinanceFees"].sum())

    # Insurance unearned (new + existing for each type)
    ue_life = _to_decimal(df["UnearnedNewCreditLife"].sum() + df["UnearnedExistingCreditLife"].sum())
    ue_disability = _to_decimal(df["UnearnedNewDisability"].sum() + df["UnearnedExistingDisability"].sum())
    ue_iui = _to_decimal(df["UnearnedNewIUI"].sum() + df["UnearnedExistingIUI"].sum())
    ue_property = _to_decimal(df["UnearnedNewProperty"].sum() + df["UnearnedExistingProperty"].sum())
    ue_vsi = _to_decimal(df["UnearnedNewVSI"].sum() + df["UnearnedExistingVSI"].sum())

    summary = UnearnedSummary(
        unearned_new_interest=unearned_new_int,
        unearned_existing_interest=unearned_exist_int,
        total_unearned_interest=unearned_new_int + unearned_exist_int,
        unearned_new_finance_fees=unearned_new_ff,
        unearned_existing_finance_fees=unearned_exist_ff,
        total_unearned_finance_fees=unearned_new_ff + unearned_exist_ff,
        unearned_credit_life=ue_life,
        unearned_disability=ue_disability,
        unearned_iui=ue_iui,
        unearned_property=ue_property,
        unearned_vsi=ue_vsi,
        total_unearned_insurance=ue_life + ue_disability + ue_iui + ue_property + ue_vsi,
        total_current_balance=_to_decimal(df["CurrentBalance"].sum()),
        original_finance_charge=_to_decimal(df["OriginalFinanceCharge"].sum()),
        interest_collected_month=_to_decimal(df["InterestCollectedMonth"].sum()),
        loan_count=len(df),
    )

    # By branch
    for branch_id, group in df.groupby("BranchId"):
        branch_int = int(branch_id)
        ue_int = _to_decimal(group["UnearnedNewInterest"].sum() + group["UnearnedExistingInterest"].sum())
        summary.by_branch[branch_int] = {
            "total_unearned_interest": ue_int,
            "current_balance": _to_decimal(group["CurrentBalance"].sum()),
            "count": len(group),
        }

    # By portfolio
    for port_id, group in df.groupby("PortfolioId"):
        port_int = int(port_id)
        ue_int = _to_decimal(group["UnearnedNewInterest"].sum() + group["UnearnedExistingInterest"].sum())
        summary.by_portfolio[port_int] = {
            "total_unearned_interest": ue_int,
            "current_balance": _to_decimal(group["CurrentBalance"].sum()),
            "count": len(group),
        }

    return df, summary


def _to_decimal(value) -> Decimal:
    return Decimal(str(round(float(value), 2))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
