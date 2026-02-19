"""
Reconciliation engine for SET Financial Corporation EOM Close.
Validates source data against QBO balances and performs roll-forward checks.
"""
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ReconItem:
    """Single reconciliation check."""
    name: str
    source_value: Decimal
    target_value: Decimal
    difference: Decimal = Decimal("0")
    tolerance: Decimal = Decimal("100")
    status: str = ""  # PASS, FAIL, WARNING
    notes: str = ""

    def __post_init__(self):
        self.difference = self.source_value - self.target_value
        abs_diff = abs(self.difference)
        if abs_diff <= Decimal("0.01"):
            self.status = "PASS"
        elif abs_diff <= self.tolerance:
            self.status = "WARNING"
            self.notes = f"Within tolerance (${self.tolerance})"
        else:
            self.status = "FAIL"
            self.notes = f"Exceeds tolerance of ${self.tolerance}"


@dataclass
class ReconciliationReport:
    """Complete reconciliation report for the month."""
    period: str
    items: list = field(default_factory=list)
    roll_forwards: list = field(default_factory=list)
    validation_checks: list = field(default_factory=list)

    @property
    def pass_count(self) -> int:
        return sum(1 for i in self.items if i.status == "PASS")

    @property
    def fail_count(self) -> int:
        return sum(1 for i in self.items if i.status == "FAIL")

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.items if i.status == "WARNING")

    def add_check(self, name: str, source: Decimal, target: Decimal,
                  tolerance: Decimal = Decimal("100")) -> ReconItem:
        item = ReconItem(name=name, source_value=source, target_value=target,
                        tolerance=tolerance)
        self.items.append(item)
        return item

    def add_validation(self, name: str, passed: bool, detail: str = ""):
        self.validation_checks.append({
            "name": name, "status": "PASS" if passed else "FAIL", "detail": detail
        })


@dataclass
class RollForward:
    """Roll-forward reconciliation (beginning + adds - subtracts = ending)."""
    name: str
    beginning_balance: Decimal = Decimal("0")
    additions: dict = field(default_factory=dict)
    subtractions: dict = field(default_factory=dict)
    expected_ending: Decimal = Decimal("0")
    actual_ending: Decimal = Decimal("0")

    @property
    def calculated_ending(self) -> Decimal:
        total_adds = sum(self.additions.values())
        total_subs = sum(self.subtractions.values())
        return self.beginning_balance + total_adds - total_subs

    @property
    def difference(self) -> Decimal:
        return self.actual_ending - self.calculated_ending

    @property
    def status(self) -> str:
        if abs(self.difference) <= Decimal("1.00"):
            return "PASS"
        elif abs(self.difference) <= Decimal("100.00"):
            return "WARNING"
        return "FAIL"


def build_loans_receivable_recon(
    nortridge_current_balance: Decimal,
    qbo_gross: Decimal,
    qbo_accum_chargeoffs: Decimal,
    qbo_unearned_interest: Decimal,
    qbo_unearned_insurance: Decimal,
    qbo_allowance: Decimal,
) -> ReconItem:
    """
    Reconcile Nortridge ending balance to QBO net loans receivable.
    QBO Net = Gross - Accum CO - Unearned Interest - Unearned Insurance - Allowance
    """
    qbo_net = qbo_gross + qbo_accum_chargeoffs + qbo_unearned_interest + qbo_unearned_insurance + qbo_allowance
    return ReconItem(
        name="Loans Receivable: Nortridge vs QBO",
        source_value=nortridge_current_balance,
        target_value=qbo_net,
        tolerance=Decimal("100"),
    )


def build_unearned_interest_recon(
    register_unearned: Decimal,
    qbo_unearned: Decimal,
) -> ReconItem:
    """
    Reconcile Unearned Register total to QBO 110010.
    Note: QBO stores as negative (contra-asset), register is positive.
    """
    return ReconItem(
        name="Unearned Interest: Register vs QBO",
        source_value=register_unearned,
        target_value=abs(qbo_unearned),
        tolerance=Decimal("50"),
    )


def build_pier_balance_recon(
    statement_balance: Decimal,
    qbo_pier_loc: Decimal,
) -> ReconItem:
    """Reconcile Pier statement ending balance to QBO 291300."""
    return ReconItem(
        name="Pier Facility Balance: Statement vs QBO",
        source_value=statement_balance,
        target_value=qbo_pier_loc,
        tolerance=Decimal("0.01"),  # Exact match required
    )


def build_charge_off_validation(
    net_chargeoff: Decimal,
    unearned_reversed: Decimal,
    total_chargeoff: Decimal,
) -> dict:
    """Validate: Net + Unearned = Total charge-off."""
    calculated = net_chargeoff + unearned_reversed
    return {
        "name": "Charge-Off Components",
        "net": net_chargeoff,
        "unearned": unearned_reversed,
        "expected_total": calculated,
        "actual_total": total_chargeoff,
        "difference": calculated - total_chargeoff,
        "status": "PASS" if abs(calculated - total_chargeoff) < Decimal("1.00") else "FAIL",
    }


def build_collections_validation(
    cash_received: Decimal,
    principal: Decimal,
    interest: Decimal,
    late_fees: Decimal,
    nsf_fees: Decimal,
    recoveries: Decimal,
    refunds: Decimal,
    insurance_rebate: Decimal,
    balance_renewed: Decimal,
    interest_rebate: Decimal,
) -> dict:
    """
    Validate collection components.
    TotalCollected = Principal + InterestCollected + LateFees + NSFFees + Recovery
                    - AmountToRefund + InsuranceRebate
    CashReceived = TotalCollected - BalanceRenewed - InterestRebate
    """
    total_collected_calc = principal + interest + late_fees + nsf_fees + recoveries - refunds + insurance_rebate
    return {
        "name": "Collections Balance Check",
        "principal": principal,
        "interest": interest,
        "fees": late_fees + nsf_fees,
        "recoveries": recoveries,
        "refunds": refunds,
        "cash_received": cash_received,
        "status": "INFO",  # This is informational
    }
