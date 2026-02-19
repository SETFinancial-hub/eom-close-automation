"""
Parser for Loan Register Excel files.
Reads Sheet2 (raw origination data) and produces summary metrics.
"""
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field


@dataclass
class LoanRegisterSummary:
    """Aggregated loan origination metrics for the month."""
    note_amount: Decimal = Decimal("0")
    finance_charge: Decimal = Decimal("0")
    cash_to_borrower: Decimal = Decimal("0")
    credit_life_premium: Decimal = Decimal("0")
    ah_premium: Decimal = Decimal("0")
    apr_fees: Decimal = Decimal("0")
    balance_renewed: Decimal = Decimal("0")
    loan_count: int = 0
    # By branch
    by_branch: dict = field(default_factory=dict)
    # By portfolio
    by_portfolio: dict = field(default_factory=dict)
    # By loan type
    by_loan_type: dict = field(default_factory=dict)


def parse_loan_register(filepath: str) -> tuple[pd.DataFrame, LoanRegisterSummary]:
    """
    Parse Loan Register Excel file.
    Returns (raw_dataframe, summary).
    """
    df = pd.read_excel(filepath, sheet_name="Sheet2", engine="openpyxl")

    numeric_cols = [
        "NoteAmount", "FinanceCharge", "CashToBorrower",
        "OriginalCreditLifePremium", "OriginalAndHPremium",
        "APRFees", "BalanceRenewed", "PandIPaymentAmount",
        "OriginalAcquisitionAmount"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    summary = LoanRegisterSummary(
        note_amount=_to_decimal(df["NoteAmount"].sum()),
        finance_charge=_to_decimal(df["FinanceCharge"].sum()),
        cash_to_borrower=_to_decimal(df["CashToBorrower"].sum()),
        credit_life_premium=_to_decimal(df["OriginalCreditLifePremium"].sum()),
        ah_premium=_to_decimal(df["OriginalAndHPremium"].sum()),
        apr_fees=_to_decimal(df["APRFees"].sum()),
        balance_renewed=_to_decimal(df["BalanceRenewed"].sum()),
        loan_count=len(df),
    )

    # Check for duplicate loan numbers
    dupes = df[df.duplicated(subset=["LoanNumber"], keep=False)]
    if len(dupes) > 0:
        summary._duplicate_loans = dupes["LoanNumber"].unique().tolist()

    # By branch
    for branch_id, group in df.groupby("BranchID"):
        summary.by_branch[int(branch_id)] = {
            "note_amount": _to_decimal(group["NoteAmount"].sum()),
            "cash_to_borrower": _to_decimal(group["CashToBorrower"].sum()),
            "finance_charge": _to_decimal(group["FinanceCharge"].sum()),
            "count": len(group),
        }

    # By portfolio
    for port_id, group in df.groupby("PortfolioID"):
        summary.by_portfolio[int(port_id)] = {
            "note_amount": _to_decimal(group["NoteAmount"].sum()),
            "count": len(group),
        }

    # By loan type
    for lt, group in df.groupby("LoanType"):
        summary.by_loan_type[lt] = {
            "note_amount": _to_decimal(group["NoteAmount"].sum()),
            "finance_charge": _to_decimal(group["FinanceCharge"].sum()),
            "count": len(group),
        }

    return df, summary


def _to_decimal(value) -> Decimal:
    return Decimal(str(round(float(value), 2))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
