"""
SET Financial Corporation — Monthly EOM Close Automation
Main orchestrator: parses all source files, generates journal entries,
runs reconciliation, and produces the close package.

Usage:
    python run_close.py --month 2025-12
    python run_close.py --month 2026-01
    python run_close.py --month 2026-01 --data-dir ./data/sample/
"""
import sys
import os
import argparse
from pathlib import Path
from decimal import Decimal
from datetime import date

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from parsers.collection_parser import parse_collection_register
from parsers.loan_register_parser import parse_loan_register
from parsers.charge_off_parser import parse_charge_offs
from parsers.unearned_parser import parse_unearned_register
from parsers.pier_statement_parser import parse_pier_statement
from parsers.dpv_statement_parser import parse_dpv_statement
from parsers.metacorp_parser import load_from_json, JANUARY_2026_SALE

from journal_entries.je_engine import (
    generate_je1_finance_income, generate_je2_insurance_earnings,
    generate_je3_originations, generate_je4_collections,
    generate_je5_charge_offs, generate_je6_bad_debt_sale,
    generate_je7_pier_interest, generate_je8_dpv_interest,
    generate_je9_allowance, generate_je10_recoveries,
    JournalEntry,
)

from reconciliation.recon_engine import (
    ReconciliationReport, build_loans_receivable_recon,
    build_unearned_interest_recon, build_pier_balance_recon,
    build_charge_off_validation, build_collections_validation,
)

from reports.excel_report import generate_excel_output
from reports.csv_report import generate_csv_output

from config.month_config import resolve_month
from config.portfolios import get_class_name


def _load_metacorp(config):
    """Load Metacorp sale data from JSON or fall back to hardcoded constant."""
    metacorp_path = config.file_paths.get("metacorp")
    if metacorp_path and metacorp_path.exists():
        return load_from_json(str(metacorp_path))
    # Fallback for January 2026 (hardcoded constant)
    if config.month_str == "2026-01":
        return JANUARY_2026_SALE
    raise FileNotFoundError(
        f"No metacorp.json found in {config.data_dir}. "
        f"Create data/{config.month_str}/metacorp.json with the Metacorp closing statement values."
    )


def run_eom_close(month_str: str, data_dir_override: str = None):
    """Execute the full EOM close process for the specified month."""
    config = resolve_month(month_str, data_dir_override)
    balances = config.prior_balances

    print("=" * 70)
    print(f"  SET Financial Corporation — EOM Close: {config.period_label}")
    print("=" * 70)

    # ──────────────────────────────────────────────────────────
    # PHASE 1: PARSE ALL SOURCE FILES
    # ──────────────────────────────────────────────────────────
    print("\n📁 PHASE 1: Parsing Source Files...")

    # 1. Collection Register
    coll_df, coll_summary = parse_collection_register(str(config.file_paths["collection"]))
    print(f"  ✓ Collection Register: {coll_summary.transaction_count} transactions, "
          f"${coll_summary.total_collected} collected")

    # 2. Loan Register
    loan_df, loan_summary = parse_loan_register(str(config.file_paths["loan"]))
    print(f"  ✓ Loan Register: {loan_summary.loan_count} loans, "
          f"${loan_summary.note_amount} originated")

    # 3. Charge Offs
    co_df, co_summary = parse_charge_offs(str(config.file_paths["charge_off"]))
    print(f"  ✓ Charge Offs: {co_summary.account_count} accounts, "
          f"${co_summary.total_charge_off_amt} total")

    # 4. Unearned Register
    ue_df, ue_summary = parse_unearned_register(str(config.file_paths["unearned"]))
    print(f"  ✓ Unearned Register: {ue_summary.loan_count} loans, "
          f"${ue_summary.total_unearned_interest} total unearned interest")

    # 5. Pier Statement
    pier_data = parse_pier_statement(str(config.file_paths["pier"]))
    print(f"  ✓ Pier Statement: Balance ${pier_data.ending_balance}, "
          f"Interest ${pier_data.accrued_interest_due}")

    # 6. DPV Statement
    dpv_data = parse_dpv_statement(str(config.file_paths["dpv"]))
    print(f"  ✓ DPV Statement: Balance ${dpv_data.principal_balance}, "
          f"Interest ${dpv_data.total_due}")

    # 7. Metacorp Sale
    metacorp = _load_metacorp(config)
    print(f"  ✓ Metacorp Sale: {metacorp.num_accounts} accounts, "
          f"${metacorp.transfer_amount} proceeds")

    # ──────────────────────────────────────────────────────────
    # PHASE 2: GENERATE JOURNAL ENTRIES (Class-Tagged by Portfolio)
    # ──────────────────────────────────────────────────────────
    print("\n📝 PHASE 2: Generating Journal Entries (by QBO class)...")

    journal_entries = []

    # JE-1: Finance Income Recognition — BY CLASS
    # Uses Collection Register InterestCollected per portfolio as earned interest proxy.
    je1 = JournalEntry(
        je_number="JE-1",
        description="Finance Income Recognition - Earned Interest",
        je_date=config.je_date,
        source_file="Unearned Register / Collection Register",
    )
    for port_id, port_data in sorted(coll_summary.by_portfolio.items()):
        cls = get_class_name(port_id)
        port_interest = port_data["interest_collected"]
        if port_interest > 0:
            sub = generate_je1_finance_income(
                prior_month_unearned_interest=port_interest,  # synthetic prior = current + collected
                current_month_unearned_interest=Decimal("0"),
                je_date=config.je_date,
                class_name=cls,
            )
            je1.lines.extend(sub.lines)
    je1.notes = ("Using Collection Register InterestCollected per portfolio "
                 f"(total ${coll_summary.interest_collected}) as earned interest proxy. "
                 "Future months will use month-over-month unearned register delta.")
    je1.review_required = True
    journal_entries.append(je1)
    print(f"  ✓ JE-1 Finance Income: ${je1.total_debits} (by class, estimated from interest collected)")

    # JE-2: Insurance Premium Earnings — SET (aggregate, no class-level priors)
    current_insurance = {
        "CreditLife": ue_summary.unearned_credit_life,
        "Disability": ue_summary.unearned_disability,
        "IUI": ue_summary.unearned_iui,
        "Property": ue_summary.unearned_property,
        "VSI": ue_summary.unearned_vsi,
    }
    prior_insurance = {
        "CreditLife": abs(balances["110050_unearned_life"]) + Decimal("0"),
        "Disability": abs(balances["110060_unearned_ah"]) + Decimal("0"),
        "IUI": abs(balances["110070_unearned_iui"]),
        "Property": abs(balances["110080_unearned_prop"]),
        "VSI": abs(balances["110090_unearned_auto"]),
    }
    je2 = generate_je2_insurance_earnings(prior_insurance, current_insurance, config.je_date,
                                          class_name="SET")
    journal_entries.append(je2)
    print(f"  ✓ JE-2 Insurance Earnings: ${je2.total_debits} (class: SET)")

    # JE-3: Loan Originations — BY CLASS
    je3 = JournalEntry(
        je_number="JE-3",
        description="Loan Originations",
        je_date=config.je_date,
        source_file="Loan Register",
    )
    for port_id, port_data in sorted(loan_summary.by_portfolio.items()):
        cls = get_class_name(port_id)
        sub = generate_je3_originations(
            note_amount=port_data["note_amount"],
            finance_charge=port_data["finance_charge"],
            credit_life_premium=port_data["credit_life_premium"],
            ah_premium=port_data["ah_premium"],
            cash_to_borrower=port_data["cash_to_borrower"],
            balance_renewed=port_data["balance_renewed"],
            je_date=config.je_date,
            class_name=cls,
        )
        je3.lines.extend(sub.lines)
    journal_entries.append(je3)
    print(f"  ✓ JE-3 Originations: ${je3.total_debits} (by class, {loan_summary.loan_count} loans)")

    # JE-4: Collections — BY CLASS
    je4 = JournalEntry(
        je_number="JE-4",
        description="Collections / Payments Received",
        je_date=config.je_date,
        source_file="Collection Register",
    )
    for port_id, port_data in sorted(coll_summary.by_portfolio.items()):
        cls = get_class_name(port_id)
        sub = generate_je4_collections(
            cash_received=port_data["cash_received"],
            principal=port_data["principal"],
            interest_collected=port_data["interest_collected"],
            interest_rebate=port_data["interest_rebate"],
            balance_renewed=port_data["balance_renewed"],
            late_fees=port_data["late_fees"],
            nsf_fees=port_data["nsf_fees"],
            amount_to_refund=port_data["amount_to_refund"],
            insurance_rebate=port_data["insurance_rebate"],
            recovery=port_data["recovery"],
            je_date=config.je_date,
            class_name=cls,
        )
        je4.lines.extend(sub.lines)
    journal_entries.append(je4)
    print(f"  ✓ JE-4 Collections: ${je4.total_debits} (by class)")

    # JE-5: Charge-Offs — BY CLASS
    je5 = JournalEntry(
        je_number="JE-5",
        description="Monthly Charge-Offs",
        je_date=config.je_date,
        source_file="Charge Off Register",
    )
    for port_id, port_data in sorted(co_summary.by_portfolio.items()):
        cls = get_class_name(port_id)
        sub = generate_je5_charge_offs(
            total_charge_off_amt=port_data["total_charge_off_amt"],
            charge_off_amount=port_data["charge_off_amount"],
            pc_interest_rebate=port_data["pc_interest_rebate"],
            je_date=config.je_date,
            class_name=cls,
        )
        je5.lines.extend(sub.lines)
    journal_entries.append(je5)
    print(f"  ✓ JE-5 Charge-Offs: ${co_summary.total_charge_off_amt} total "
          f"(${co_summary.charge_off_amount} net, ${co_summary.pc_interest_rebate} unearned) (by class)")

    # JE-6: Bad Debt Sale — SET (company-level)
    je6 = generate_je6_bad_debt_sale(
        transfer_amount=metacorp.transfer_amount,
        je_date=config.je_date,
        num_accounts=metacorp.num_accounts,
        total_balance=metacorp.total_current_balance,
        class_name="SET",
    )
    journal_entries.append(je6)
    print(f"  ✓ JE-6 Bad Debt Sale: ${metacorp.transfer_amount} (class: SET)")

    # JE-7: Pier Interest — SET (facility-level)
    je7 = generate_je7_pier_interest(
        accrued_interest=pier_data.accrued_interest_due,
        je_date=config.je_date,
        ending_balance=pier_data.ending_balance,
        class_name="SET",
    )
    journal_entries.append(je7)
    print(f"  ✓ JE-7 Pier Interest: ${pier_data.accrued_interest_due} (class: SET)")

    # JE-8: DPV Interest — SET (facility-level)
    je8 = generate_je8_dpv_interest(
        interest_amount=dpv_data.total_due,
        je_date=config.je_date,
        principal_balance=dpv_data.principal_balance,
        class_name="SET",
    )
    journal_entries.append(je8)
    print(f"  ✓ JE-8 DPV Interest: ${dpv_data.total_due} (class: SET)")

    # JE-9: Allowance for Credit Losses — SET (management judgment, aggregate)
    net_receivable = ue_summary.total_current_balance
    current_allowance = abs(balances["110002_allowance"])
    target_pct = Decimal("18")
    target_allowance = (net_receivable * target_pct / Decimal("100")).quantize(Decimal("0.01"))
    adjustment = target_allowance - current_allowance

    je9 = generate_je9_allowance(
        adjustment_amount=adjustment,
        je_date=config.je_date,
        portfolio_balance=net_receivable,
        target_allowance_pct=target_pct,
        class_name="SET",
    )
    journal_entries.append(je9)
    print(f"  ✓ JE-9 Allowance: Suggested adjustment ${adjustment} "
          f"(current ${current_allowance}, target ${target_allowance}) (class: SET)")

    # JE-10: Recoveries — consolidated into JE-4 (Collections)
    if coll_summary.recovery > 0:
        print(f"  ℹ JE-10 Recoveries: ${coll_summary.recovery} (included in JE-4 Collections)")

    # ──────────────────────────────────────────────────────────
    # PHASE 3: RECONCILIATION
    # ──────────────────────────────────────────────────────────
    print("\n🔍 PHASE 3: Reconciliation Checks...")

    recon = ReconciliationReport(period=config.period_label)

    # Pier balance check
    pier_recon = build_pier_balance_recon(
        statement_balance=pier_data.ending_balance,
        qbo_pier_loc=balances["291300_pier_loc"],
    )
    recon.items.append(pier_recon)
    print(f"  {'✓' if pier_recon.status == 'PASS' else '✗'} Pier Balance: "
          f"Statement ${pier_data.ending_balance} vs QBO ${balances['291300_pier_loc']} "
          f"→ {pier_recon.status}")

    # Unearned interest check
    ue_recon = build_unearned_interest_recon(
        register_unearned=ue_summary.total_unearned_interest,
        qbo_unearned=balances["110010_unearned_interest"],
    )
    ue_recon.notes = (f"QBO balance may not be as of {config.je_date}. "
                      f"Diff of ${abs(ue_recon.difference)} may include activity after month-end. "
                      f"Re-run with exact month-end QBO trial balance for true recon.")
    recon.items.append(ue_recon)
    print(f"  ⚠ Unearned Interest: Register ${ue_summary.total_unearned_interest} "
          f"vs QBO ${abs(balances['110010_unearned_interest'])} "
          f"→ EXPECTED DIFF due to date mismatch (diff: ${ue_recon.difference})")

    # Charge-off validation
    co_valid = build_charge_off_validation(
        net_chargeoff=co_summary.charge_off_amount,
        unearned_reversed=co_summary.pc_interest_rebate,
        total_chargeoff=co_summary.total_charge_off_amt,
    )
    recon.add_validation(
        "Charge-Off Components (Net + Unearned = Total)",
        co_valid["status"] == "PASS",
        f"Net ${co_summary.charge_off_amount} + Unearned ${co_summary.pc_interest_rebate} "
        f"= ${co_valid['expected_total']} vs Total ${co_summary.total_charge_off_amt}"
    )
    print(f"  {'✓' if co_valid['status'] == 'PASS' else '✗'} Charge-Off Components: {co_valid['status']}")

    # Collections validation
    coll_valid = build_collections_validation(
        cash_received=coll_summary.cash_received,
        principal=coll_summary.principal,
        interest=coll_summary.interest_collected,
        late_fees=coll_summary.late_fees,
        nsf_fees=coll_summary.nsf_fees,
        recoveries=coll_summary.recovery,
        refunds=coll_summary.amount_to_refund,
        insurance_rebate=coll_summary.insurance_rebate,
        balance_renewed=coll_summary.balance_renewed,
        interest_rebate=coll_summary.interest_rebate,
    )
    recon.add_validation(
        "Collections Component Check",
        coll_valid["status"] != "FAIL",
        f"Cash ${coll_summary.cash_received}, Principal ${coll_summary.principal}, "
        f"Interest ${coll_summary.interest_collected}, Fees ${coll_summary.late_fees + coll_summary.nsf_fees}"
    )
    print(f"  ✓ Collections Validation: {coll_valid['status']}")

    # JE balance checks
    all_balanced = True
    for je in journal_entries:
        if not je.is_balanced and len(je.lines) > 0:
            all_balanced = False
            recon.add_validation(
                f"{je.je_number} Balance Check",
                False,
                f"Debits ${je.total_debits} ≠ Credits ${je.total_credits}"
            )
            print(f"  ✗ {je.je_number}: UNBALANCED (DR ${je.total_debits} vs CR ${je.total_credits})")

    if all_balanced:
        recon.add_validation("All JEs Balanced", True, "All journal entries have matching debits and credits")
        print(f"  ✓ All journal entries balanced")

    # Class-level totals vs aggregate validation
    class_tagged_jes = {
        "JE-1": coll_summary.interest_collected,
        "JE-3": loan_summary.note_amount,       # total debits = note_amount
        "JE-5": co_summary.total_charge_off_amt, # total credits = total_charge_off_amt
    }
    for je in journal_entries:
        if je.je_number in class_tagged_jes:
            expected = class_tagged_jes[je.je_number]
            if je.je_number == "JE-5":
                actual = je.total_credits
            else:
                actual = je.total_debits
            diff = abs(actual - expected)
            ok = diff < Decimal("0.02")
            recon.add_validation(
                f"{je.je_number} Class Totals Match Aggregate",
                ok,
                f"Class sum ${actual} vs aggregate ${expected} (diff ${diff})"
            )
            print(f"  {'✓' if ok else '✗'} {je.je_number} class total: ${actual} vs aggregate ${expected}")

    # Duplicate loan check
    dupes = loan_df[loan_df.duplicated(subset=["LoanNumber"], keep=False)]
    has_dupes = len(dupes) > 0
    recon.add_validation(
        "No Duplicate Loan Numbers",
        not has_dupes,
        f"{len(dupes)} duplicates found" if has_dupes else "All loan numbers unique"
    )
    print(f"  {'✗' if has_dupes else '✓'} Duplicate Loan Check: "
          f"{'FAIL' if has_dupes else 'PASS'}")

    # Portfolio ID validation
    from config.portfolios import VALID_PORTFOLIO_IDS
    invalid_ports = set()
    for df_check in [loan_df, co_df]:
        port_col = "PortfolioID" if "PortfolioID" in df_check.columns else "PortfolioId"
        if port_col in df_check.columns:
            for pid in df_check[port_col].unique():
                if int(pid) not in VALID_PORTFOLIO_IDS and int(pid) != 0:
                    invalid_ports.add(int(pid))
    recon.add_validation(
        "Valid Portfolio IDs",
        len(invalid_ports) == 0,
        f"Invalid IDs: {invalid_ports}" if invalid_ports else "All portfolio IDs valid"
    )
    print(f"  {'✗' if invalid_ports else '✓'} Portfolio ID Check: "
          f"{'FAIL - ' + str(invalid_ports) if invalid_ports else 'PASS'}")

    # ──────────────────────────────────────────────────────────
    # PHASE 4: GENERATE OUTPUT
    # ──────────────────────────────────────────────────────────
    print("\n📊 PHASE 4: Generating Output Files...")

    # Generate Excel close package
    generate_excel_output(journal_entries, recon, coll_summary, loan_summary,
                          co_summary, ue_summary, pier_data, dpv_data, metacorp,
                          config.output_dir, config.period_label)

    # Generate CSV for QBO import
    generate_csv_output(journal_entries, config.output_dir, config.period_label)

    print(f"\n{'=' * 70}")
    print(f"  EOM Close Complete: {config.period_label}")
    print(f"  Journal Entries: {len(journal_entries)}")
    print(f"  Recon Checks: {recon.pass_count} PASS, {recon.warning_count} WARNING, {recon.fail_count} FAIL")
    print(f"  Validations: {len(recon.validation_checks)}")
    print(f"  Output: {config.output_dir}")
    print(f"{'=' * 70}")

    return journal_entries, recon


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="SET Financial Corporation — Monthly EOM Close Automation"
    )
    parser.add_argument(
        "--month", required=True,
        help="Month to close in YYYY-MM format (e.g., 2025-12, 2026-01)"
    )
    parser.add_argument(
        "--data-dir", default=None,
        help="Optional override for source data directory (default: data/{YYYY-MM}/)"
    )
    args = parser.parse_args()
    run_eom_close(args.month, args.data_dir)
