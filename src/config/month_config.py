"""
Month configuration for parameterized EOM close runs.
Resolves file paths, prior-period balances, and metacorp data for any month.
"""
from dataclasses import dataclass, field
from decimal import Decimal
from datetime import date
from pathlib import Path
from typing import Optional
import calendar
import glob


PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


@dataclass
class MonthConfig:
    """All configuration needed to run an EOM close for a given month."""
    month_str: str          # "2025-12"
    period_label: str       # "December 2025"
    je_date: date           # last day of month
    data_dir: Path
    output_dir: Path
    prior_balances: dict    # QBO balances as of prior month-end
    file_paths: dict        # resolved source file paths


# ──────────────────────────────────────────────────────────────
# Prior-period QBO balances, keyed by the month they apply TO.
# These are the ending balances from the PRIOR month's close,
# used as opening balances for the current month.
# ──────────────────────────────────────────────────────────────
KNOWN_BALANCES = {
    # Balances as of 11/30/2025 (used for December 2025 close)
    # These represent the QBO state BEFORE December activity.
    # Since December is the accountant's last finalized month,
    # we derive these from the December tie-out ending balances
    # by backing out December JEs. For now, we use the same
    # balances that were hardcoded for the January close — the
    # QBO snapshot as of ~2/16/2026 which effectively captures
    # the December close state.
    "2025-12": {
        "110001_gross": Decimal("35243907.50"),
        "110200_accum_co": Decimal("-33313354.46"),
        "110002_allowance": Decimal("-328194.02"),
        "110010_unearned_interest": Decimal("-55038.50"),
        "110050_unearned_life": Decimal("-1.28"),
        "110060_unearned_ah": Decimal("-0.71"),
        "110070_unearned_iui": Decimal("0"),
        "110080_unearned_prop": Decimal("0"),
        "110090_unearned_auto": Decimal("0"),
        "total_unearned_insurance": Decimal("-1.99"),
        "291300_pier_loc": Decimal("1591414.81"),
        "290060_accrued_expenses": Decimal("95349.58"),
    },
    # Balances as of 12/31/2025 (used for January 2026 close)
    # Same values — these ARE the December ending balances.
    "2026-01": {
        "110001_gross": Decimal("35243907.50"),
        "110200_accum_co": Decimal("-33313354.46"),
        "110002_allowance": Decimal("-328194.02"),
        "110010_unearned_interest": Decimal("-55038.50"),
        "110050_unearned_life": Decimal("-1.28"),
        "110060_unearned_ah": Decimal("-0.71"),
        "110070_unearned_iui": Decimal("0"),
        "110080_unearned_prop": Decimal("0"),
        "110090_unearned_auto": Decimal("0"),
        "total_unearned_insurance": Decimal("-1.99"),
        "291300_pier_loc": Decimal("1591414.81"),
        "290060_accrued_expenses": Decimal("95349.58"),
    },
}


def _find_file(data_dir: Path, patterns: list[str], label: str) -> Optional[Path]:
    """Find a file matching any of the given glob patterns."""
    for pattern in patterns:
        matches = list(data_dir.glob(pattern))
        if matches:
            return matches[0]
    return None


def _resolve_files(data_dir: Path, year: int, month: int) -> dict:
    """Resolve source file paths for a given month using pattern matching."""
    month_name = calendar.month_name[month]
    month_abbr = calendar.month_abbr[month]
    yr2 = str(year)[-2:]

    files = {}

    # Collection Register
    files["collection"] = _find_file(data_dir, [
        f"Collection Register_{month_name}{year}.xlsx",
        f"Collection Register_{month_abbr}{year}.xlsx",
        f"*Collection*Register*{year}*.xlsx",
        f"*Collection*Register*{month_name}*.xlsx",
    ], "Collection Register")

    # Loan Register
    files["loan"] = _find_file(data_dir, [
        f"Loan Register_{month_name}{year}.xlsx",
        f"Loan Register_{month_abbr}{year}.xlsx",
        f"*Loan*Register*{year}*.xlsx",
        f"*Loan*Register*{month_name}*.xlsx",
    ], "Loan Register")

    # Charge Offs
    files["charge_off"] = _find_file(data_dir, [
        f"Charge Offs_{month_name}{year}.xlsx",
        f"Charge Offs_{month_abbr}{year}.xlsx",
        f"*Charge*Off*{year}*.xlsx",
        f"*Charge*Off*{month_name}*.xlsx",
    ], "Charge Offs")

    # Unearned Register
    files["unearned"] = _find_file(data_dir, [
        f"Unearned_{month_name}{year}.xlsx",
        f"Unearned_{month_abbr}{year}.xlsx",
        f"*Unearned*{year}*.xlsx",
        f"*Unearned*{month_name}*.xlsx",
    ], "Unearned Register")

    # Pier Statement (PAT LLC Series 47)
    files["pier"] = _find_file(data_dir, [
        f"PAT LLC Series 47 - SET - {month_name} {year}.pdf",
        f"PAT*LLC*{month_name}*{year}*.pdf",
        f"*PAT*LLC*{month_name}*.pdf",
        f"*PAT*LLC*{month_abbr}*.pdf",
        f"*Pier*{month_name}*.pdf",
    ], "Pier Statement")

    # DPV Statement — naming is inconsistent (e.g., "2026 - DPV LLC - JAN.pdf"
    # vs "2025 - DPV LLC - Dec Invoice.pdf")
    files["dpv"] = _find_file(data_dir, [
        f"{year} - DPV LLC - {month_abbr.upper()}.pdf",
        f"{year} - DPV LLC - {month_name}.pdf",
        f"{year} - DPV LLC - {month_abbr} Invoice.pdf",
        f"{year} - DPV LLC - {month_abbr.capitalize()} Invoice.pdf",
        f"*DPV*LLC*{month_abbr}*.pdf",
        f"*DPV*{year}*.pdf",
        f"*dpv*{month_abbr}*.pdf",
    ], "DPV Statement")

    # Metacorp JSON
    files["metacorp"] = _find_file(data_dir, [
        "metacorp.json",
    ], "Metacorp Data")

    return files


def resolve_month(month_str: str, data_dir_override: Optional[str] = None) -> MonthConfig:
    """
    Resolve a month string like "2025-12" into a full MonthConfig.

    Args:
        month_str: "YYYY-MM" format
        data_dir_override: Optional override for the data directory path
    """
    year, month = int(month_str[:4]), int(month_str[5:7])
    month_name = calendar.month_name[month]
    last_day = calendar.monthrange(year, month)[1]

    # Data directory
    if data_dir_override:
        data_dir = Path(data_dir_override)
    else:
        data_dir = PROJECT_ROOT / "data" / month_str

    # Output directory
    output_dir = PROJECT_ROOT / "output" / month_str
    output_dir.mkdir(parents=True, exist_ok=True)

    # Prior balances
    if month_str not in KNOWN_BALANCES:
        raise ValueError(
            f"No prior-period balances found for {month_str}. "
            f"Add them to KNOWN_BALANCES in month_config.py. "
            f"Available months: {list(KNOWN_BALANCES.keys())}"
        )
    prior_balances = KNOWN_BALANCES[month_str]

    # Resolve files
    file_paths = _resolve_files(data_dir, year, month)

    # Validate required files exist
    required = ["collection", "loan", "charge_off", "unearned", "pier", "dpv"]
    missing = [k for k in required if file_paths.get(k) is None]
    if missing:
        available = list(data_dir.glob("*")) if data_dir.exists() else []
        avail_str = "\n  ".join(str(f.name) for f in available) if available else "(directory empty or missing)"
        raise FileNotFoundError(
            f"Missing required files for {month_str} in {data_dir}:\n"
            f"  Missing: {missing}\n"
            f"  Available files:\n  {avail_str}"
        )

    return MonthConfig(
        month_str=month_str,
        period_label=f"{month_name} {year}",
        je_date=date(year, month, last_day),
        data_dir=data_dir,
        output_dir=output_dir,
        prior_balances=prior_balances,
        file_paths=file_paths,
    )
