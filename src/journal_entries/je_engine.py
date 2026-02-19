"""
Journal Entry Engine for SET Financial Corporation EOM Close.
Generates all 10 journal entry types from parsed source data.
"""
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field
from datetime import date
from typing import Optional
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from config.accounts import *


@dataclass
class JournalEntryLine:
    """Single line of a journal entry."""
    account_code: str
    account_name: str
    debit: Decimal = Decimal("0")
    credit: Decimal = Decimal("0")
    memo: str = ""


@dataclass
class JournalEntry:
    """Complete journal entry with header and lines."""
    je_number: str           # e.g., "JE-1"
    description: str
    je_date: date
    lines: list = field(default_factory=list)
    source_file: str = ""
    notes: str = ""
    review_required: bool = False

    @property
    def total_debits(self) -> Decimal:
        return sum(line.debit for line in self.lines)

    @property
    def total_credits(self) -> Decimal:
        return sum(line.credit for line in self.lines)

    @property
    def is_balanced(self) -> bool:
        return abs(self.total_debits - self.total_credits) < Decimal("0.02")

    def add_line(self, account_code: str, debit: Decimal = Decimal("0"),
                 credit: Decimal = Decimal("0"), memo: str = ""):
        name = ACCOUNT_NAMES.get(account_code, account_code)
        self.lines.append(JournalEntryLine(
            account_code=account_code,
            account_name=name,
            debit=_d(debit),
            credit=_d(credit),
            memo=memo,
        ))


def generate_je1_finance_income(
    prior_month_unearned_interest: Decimal,
    current_month_unearned_interest: Decimal,
    je_date: date,
) -> JournalEntry:
    """
    JE-1: Finance Income Recognition (Earned Interest).
    Earned = Prior month unearned - Current month unearned.
    (Positive = income earned as unearned balance decreases)
    Note: New loan originations add to unearned (handled in JE-3),
    so the net change here reflects the earning process.
    """
    earned_interest = prior_month_unearned_interest - current_month_unearned_interest

    je = JournalEntry(
        je_number="JE-1",
        description="Finance Income Recognition - Earned Interest",
        je_date=je_date,
        source_file="Unearned Register",
    )

    if earned_interest > 0:
        je.add_line(UNEARNED_PRECOMPUTED_INTEREST, debit=earned_interest,
                    memo="Decrease in unearned interest (earning)")
        je.add_line(FINANCE_INCOME, credit=earned_interest,
                    memo="Earned interest revenue")
    elif earned_interest < 0:
        # Unearned increased (more new loans than earned) — unusual but possible
        je.add_line(FINANCE_INCOME, debit=abs(earned_interest),
                    memo="Adjustment - unearned increase exceeded earning")
        je.add_line(UNEARNED_PRECOMPUTED_INTEREST, credit=abs(earned_interest),
                    memo="Increase in unearned interest")
        je.review_required = True
        je.notes = "Unearned interest increased net — verify new originations vs earnings."

    return je


def generate_je2_insurance_earnings(
    prior_unearned_insurance: dict,
    current_unearned_insurance: dict,
    je_date: date,
) -> JournalEntry:
    """
    JE-2: Insurance Premium Earnings.
    Each insurance type tracked separately.
    prior/current dicts have keys: CreditLife, Disability, IUI, Property, VSI
    """
    je = JournalEntry(
        je_number="JE-2",
        description="Insurance Premium Earnings",
        je_date=je_date,
        source_file="Unearned Register",
    )

    for ins_type, mapping in INSURANCE_MAPPING.items():
        prior = prior_unearned_insurance.get(ins_type, Decimal("0"))
        current = current_unearned_insurance.get(ins_type, Decimal("0"))
        earned = prior - current

        if earned != Decimal("0") and abs(earned) >= Decimal("0.01"):
            if earned > 0:
                je.add_line(mapping["unearned"], debit=earned,
                            memo=f"Earned {ins_type} premium")
                je.add_line(mapping["earned"], credit=earned,
                            memo=f"{ins_type} commission earned")
            else:
                je.add_line(mapping["earned"], debit=abs(earned),
                            memo=f"Reverse {ins_type} over-earned")
                je.add_line(mapping["unearned"], credit=abs(earned),
                            memo=f"Increase unearned {ins_type}")

    return je


def generate_je3_originations(
    note_amount: Decimal,
    finance_charge: Decimal,
    credit_life_premium: Decimal,
    ah_premium: Decimal,
    cash_to_borrower: Decimal,
    balance_renewed: Decimal,
    je_date: date,
) -> JournalEntry:
    """
    JE-3: Loan Originations.
    DR: Loans Receivable Gross (NoteAmount)
    CR: Unearned Pre-Computed Interest (FinanceCharge)
    CR: Unearned Insurance (premiums)
    CR: Bank/Funding (CashToBorrower)
    CR: Loans Receivable Gross (BalanceRenewed — for renewals)
    """
    je = JournalEntry(
        je_number="JE-3",
        description="Loan Originations",
        je_date=je_date,
        source_file="Loan Register",
    )

    # Debit gross receivable for full note amount
    if note_amount > 0:
        je.add_line(LOANS_RECEIVABLE_GROSS, debit=note_amount,
                    memo="New loan originations - note amount")

    # Credit unearned interest (finance charge on PC loans)
    if finance_charge > 0:
        je.add_line(UNEARNED_PRECOMPUTED_INTEREST, credit=finance_charge,
                    memo="Unearned interest on new loans")

    # Credit unearned insurance
    if credit_life_premium > 0:
        je.add_line(UNEARNED_LIFE_INS, credit=credit_life_premium,
                    memo="Unearned life insurance on new loans")
    if ah_premium > 0:
        je.add_line(UNEARNED_AH_INS, credit=ah_premium,
                    memo="Unearned A&H insurance on new loans")

    # Credit bank for net cash disbursed
    # Note: For NLS simple interest loans, NoteAmount = CashToBorrower (no finance charge)
    if cash_to_borrower > 0:
        je.add_line("FUNDING_BANK", credit=cash_to_borrower,
                    memo="Cash disbursed to borrowers")

    # Credit receivable for renewed balances (nets out the renewal portion)
    if balance_renewed > 0:
        je.add_line(LOANS_RECEIVABLE_GROSS, credit=balance_renewed,
                    memo="Balance renewed on renewal loans (reduces net new receivable)")

    return je


def generate_je4_collections(
    cash_received: Decimal,
    principal: Decimal,
    interest_collected: Decimal,
    interest_rebate: Decimal,
    balance_renewed: Decimal,
    late_fees: Decimal,
    nsf_fees: Decimal,
    amount_to_refund: Decimal,
    insurance_rebate: Decimal,
    recovery: Decimal,
    je_date: date,
) -> JournalEntry:
    """
    JE-4: Collections / Payments.
    DR: Bank (CashReceived — actual cash deposited)
    DR: Loans Receivable Gross (BalanceRenewed — renewals, if any)
    CR: Loans Receivable Gross (residual — principal + interest portion reducing receivable)
    CR: Delinquent/NSF Fees (LateFees + NSFFees)
    CR: Refunds (AmountToRefund — negative credit / debit)
    CR: Insurance Rebates
    CR: Customer Recoveries

    Note: Interest collected reduces the gross receivable here. The corresponding
    earned interest recognition (DR Unearned, CR Finance Income) is in JE-1.
    Together, these keep the net loans receivable correct.
    """
    je = JournalEntry(
        je_number="JE-4",
        description="Collections / Payments Received",
        je_date=je_date,
        source_file="Collection Register",
    )

    # DEBITS: Cash in + renewal offset
    if cash_received > 0:
        je.add_line("PAYMENT_BANK", debit=cash_received,
                    memo="Cash collections received")
    if balance_renewed > 0:
        je.add_line(LOANS_RECEIVABLE_GROSS, debit=balance_renewed,
                    memo="Renewal balances (old loan paid by new loan)")

    # CREDITS: Loan balance reduction + fee income
    # The receivable credit = cash - fees - recovery + refunds (the loan-applied portion)
    fees_and_other = late_fees + nsf_fees + insurance_rebate + recovery - amount_to_refund
    receivable_credit = cash_received + balance_renewed - fees_and_other

    if receivable_credit > 0:
        je.add_line(LOANS_RECEIVABLE_GROSS, credit=receivable_credit,
                    memo=f"Loan balance reduction (principal ${principal} + interest ${interest_collected})")

    if late_fees > 0:
        je.add_line(DELINQUENT_NSF_FEES, credit=late_fees,
                    memo="Late fees collected")
    if nsf_fees > 0:
        je.add_line(DELINQUENT_NSF_FEES, credit=nsf_fees,
                    memo="NSF fees collected")
    if amount_to_refund > 0:
        je.add_line(REFUNDS, debit=amount_to_refund,
                    memo="Customer refunds issued")
    if insurance_rebate > 0:
        je.add_line(EARNED_INS_REBATES, credit=insurance_rebate,
                    memo="Insurance premium rebates")
    if recovery > 0:
        je.add_line(CUSTOMER_RECOVERIES, credit=recovery,
                    memo="Collections on charged-off accounts")

    return je


def generate_je5_charge_offs(
    total_charge_off_amt: Decimal,
    charge_off_amount: Decimal,
    pc_interest_rebate: Decimal,
    je_date: date,
) -> JournalEntry:
    """
    JE-5: Charge-Offs.
    When a loan charges off, the full balance (TotalChargeOffAmt) is removed
    from gross receivable. The net loss (ChargeOffAmount) hits P&L. The
    unearned portion (PCInterestRebate) is reversed since it was never earned.

    DR: 610500 Bad Debt Writeoffs (ChargeOffAmount — net P&L expense)
    DR: 110010 Unearned Interest (PCInterestRebate — reverse unearned)
    CR: 110200 Accumulated Charge Offs (TotalChargeOffAmt — increase contra)

    Balance check: ChargeOffAmount + PCInterestRebate = TotalChargeOffAmt
    """
    je = JournalEntry(
        je_number="JE-5",
        description="Monthly Charge-Offs",
        je_date=je_date,
        source_file="Charge Off Register",
    )

    # Debits: expense + unearned reversal
    je.add_line(BAD_DEBT_WRITEOFFS, debit=charge_off_amount,
                memo="Net charge-off expense (P&L impact)")
    if pc_interest_rebate > 0:
        je.add_line(UNEARNED_PRECOMPUTED_INTEREST, debit=pc_interest_rebate,
                    memo="Reverse unearned interest on charged-off loans")

    # Credit: accumulated charge-offs (contra to gross receivable)
    je.add_line(ACCUMULATED_CHARGE_OFFS, credit=total_charge_off_amt,
                memo="Record charge-offs in contra account")

    return je


def generate_je6_bad_debt_sale(
    transfer_amount: Decimal,
    je_date: date,
    num_accounts: int = 0,
    total_balance: Decimal = Decimal("0"),
) -> JournalEntry:
    """
    JE-6: Bad Debt Sales (Metacorp).
    DR: Bank (transfer amount received)
    CR: Sale of Bad Debt (revenue)
    """
    je = JournalEntry(
        je_number="JE-6",
        description=f"Bad Debt Sale to Metacorp ({num_accounts} accounts, ${total_balance} balance)",
        je_date=je_date,
        source_file="Metacorp Closing Statement",
    )

    je.add_line("OPERATING_BANK", debit=transfer_amount,
                memo=f"Metacorp sale proceeds ({num_accounts} accts)")
    je.add_line(SALE_OF_BAD_DEBT, credit=transfer_amount,
                memo="Revenue from bad debt sale")

    return je


def generate_je7_pier_interest(
    accrued_interest: Decimal,
    je_date: date,
    ending_balance: Decimal = Decimal("0"),
) -> JournalEntry:
    """
    JE-7: Pier Interest Accrual.
    DR: Interest Expense
    CR: Accrued Expenses
    """
    je = JournalEntry(
        je_number="JE-7",
        description=f"Pier Facility Interest Accrual (balance: ${ending_balance})",
        je_date=je_date,
        source_file="PAT LLC Series 47 Statement",
    )

    je.add_line(INTEREST_EXPENSE, debit=accrued_interest,
                memo="Pier facility monthly interest")
    je.add_line(ACCRUED_EXPENSES, credit=accrued_interest,
                memo="Accrued Pier interest payable")

    return je


def generate_je8_dpv_interest(
    interest_amount: Decimal,
    je_date: date,
    principal_balance: Decimal = Decimal("0"),
) -> JournalEntry:
    """
    JE-8: DPV LLC Interest.
    DR: DPV Interest (31200 - equity/LOC account)
    CR: Bank (auto-debited from BofA DDA)
    """
    je = JournalEntry(
        je_number="JE-8",
        description=f"DPV LLC Interest (BofA LOC, balance: ${principal_balance})",
        je_date=je_date,
        source_file="DPV LLC / BofA Invoice",
    )

    je.add_line(DPV_INTEREST, debit=interest_amount,
                memo="DPV LLC BofA interest (SOFR-based)")
    je.add_line("DPV_BANK", credit=interest_amount,
                memo="Auto-debited from BofA DDA")

    return je


def generate_je9_allowance(
    adjustment_amount: Decimal,
    je_date: date,
    portfolio_balance: Decimal = Decimal("0"),
    target_allowance_pct: Decimal = Decimal("0"),
) -> JournalEntry:
    """
    JE-9: Allowance for Credit Losses.
    DR: Allowance Adjustment (610502)
    CR: Allowance for Credit Losses (110002)
    Note: Flagged for human review — judgment-based entry.
    """
    je = JournalEntry(
        je_number="JE-9",
        description="Allowance for Credit Losses Adjustment",
        je_date=je_date,
        source_file="Management Estimate",
        review_required=True,
        notes=f"Portfolio balance: ${portfolio_balance}. "
              f"Target allowance: {target_allowance_pct}%. "
              f"This is a judgment-based entry requiring management review.",
    )

    if adjustment_amount > 0:
        # Increase allowance (provision expense)
        je.add_line(ALLOWANCE_ADJUSTMENT, debit=adjustment_amount,
                    memo="Provision for credit losses")
        je.add_line(ALLOWANCE_CREDIT_LOSSES, credit=adjustment_amount,
                    memo="Increase allowance balance")
    elif adjustment_amount < 0:
        # Decrease allowance (release)
        je.add_line(ALLOWANCE_CREDIT_LOSSES, debit=abs(adjustment_amount),
                    memo="Release excess allowance")
        je.add_line(ALLOWANCE_ADJUSTMENT, credit=abs(adjustment_amount),
                    memo="Allowance release (benefit)")

    return je


def generate_je10_recoveries(
    recovery_amount: Decimal,
    je_date: date,
) -> JournalEntry:
    """
    JE-10: Recoveries on charged-off accounts.
    DR: Bank
    CR: Customer Recoveries (420100)
    """
    je = JournalEntry(
        je_number="JE-10",
        description="Recoveries on Charged-Off Accounts",
        je_date=je_date,
        source_file="Collection Register (Recovery column)",
    )

    if recovery_amount > 0:
        je.add_line("PAYMENT_BANK", debit=recovery_amount,
                    memo="Cash recovered on charged-off accounts")
        je.add_line(CUSTOMER_RECOVERIES, credit=recovery_amount,
                    memo="Recovery revenue")

    return je


def _d(value) -> Decimal:
    """Ensure Decimal with 2 decimal places."""
    if isinstance(value, Decimal):
        return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return Decimal(str(round(float(value), 2))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
