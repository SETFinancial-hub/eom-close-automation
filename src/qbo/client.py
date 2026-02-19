"""
QBO API Client
--------------
Wraps Intuit QuickBooks Online REST API v3 calls.
All report endpoints return pandas DataFrames for easy merging
with Nortridge portfolio data.

Error handling:
- HTTP errors: logged with intuit_tid for Intuit support troubleshooting
- Auth errors: delegated to QBOAuth (auto-refresh / TokenExpiredError)
- Validation errors: 400 responses parsed and logged with detail
"""
import os
import logging
from typing import Optional
from pathlib import Path

import requests
import pandas as pd

from .auth import QBOAuth

QBO_BASE = "https://quickbooks.api.intuit.com/v3/company"
SANDBOX_BASE = "https://sandbox-quickbooks.api.intuit.com/v3/company"

# ------------------------------------------------------------------
# Logging — errors written to logs/qbo.log for Intuit troubleshooting
# ------------------------------------------------------------------
_LOG_DIR = Path(__file__).parent.parent.parent / "logs"
_LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(_LOG_DIR / "qbo.log"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("qbo.client")


class QBOClient:
    """QuickBooks Online API client for financial report retrieval."""

    def __init__(self, auth: Optional[QBOAuth] = None):
        self.auth = auth or QBOAuth()
        self._env = os.environ.get("QBO_ENVIRONMENT", "sandbox").lower()
        self._base = SANDBOX_BASE if self._env == "sandbox" else QBO_BASE

    def _get(self, path: str, params: Optional[dict] = None) -> dict:
        url = f"{self._base}/{self.auth.realm_id}/{path}"
        resp = requests.get(
            url,
            headers={
                "Authorization": f"Bearer {self.auth.access_token}",
                "Accept": "application/json",
            },
            params=params or {},
            timeout=30,
        )

        # Capture intuit_tid from response headers for Intuit support
        intuit_tid = resp.headers.get("intuit_tid", "n/a")

        if not resp.ok:
            # Parse validation/syntax errors from QBO error body
            try:
                err_body = resp.json()
                err_detail = err_body.get("Fault", {}).get("Error", [{}])
                err_msg = "; ".join(
                    f"{e.get('code','?')}: {e.get('Message','?')} — {e.get('Detail','')}"
                    for e in err_detail
                ) if err_detail else str(err_body)
            except Exception:
                err_msg = resp.text[:500]

            log.error(
                "QBO API error | status=%s | intuit_tid=%s | url=%s | error=%s",
                resp.status_code, intuit_tid, url, err_msg,
            )
            resp.raise_for_status()

        log.info("QBO API call | status=%s | intuit_tid=%s | url=%s",
                 resp.status_code, intuit_tid, url)
        return resp.json()

    # ------------------------------------------------------------------
    # Financial Reports
    # ------------------------------------------------------------------

    def profit_and_loss(self, start_date: str, end_date: str) -> pd.DataFrame:
        """
        Pull Profit & Loss report for a date range.

        Args:
            start_date: YYYY-MM-DD
            end_date: YYYY-MM-DD

        Returns:
            DataFrame with columns: account, amount
        """
        data = self._get("reports/ProfitAndLoss", {
            "start_date": start_date,
            "end_date": end_date,
            "accounting_method": "Accrual",
        })
        return self._parse_report(data, "ProfitAndLoss")

    def balance_sheet(self, as_of_date: str) -> pd.DataFrame:
        """
        Pull Balance Sheet as of a specific date.

        Args:
            as_of_date: YYYY-MM-DD

        Returns:
            DataFrame with columns: account, amount
        """
        data = self._get("reports/BalanceSheet", {
            "start_date": as_of_date,
            "end_date": as_of_date,
            "accounting_method": "Accrual",
        })
        return self._parse_report(data, "BalanceSheet")

    def cash_flow(self, start_date: str, end_date: str) -> pd.DataFrame:
        """Pull Cash Flow Statement for a date range."""
        data = self._get("reports/CashFlow", {
            "start_date": start_date,
            "end_date": end_date,
        })
        return self._parse_report(data, "CashFlow")

    def general_ledger(self, start_date: str, end_date: str) -> pd.DataFrame:
        """Pull General Ledger detail for a date range."""
        data = self._get("reports/GeneralLedger", {
            "start_date": start_date,
            "end_date": end_date,
            "accounting_method": "Accrual",
        })
        return self._parse_report(data, "GeneralLedger")

    def accounts_receivable_aging(self) -> pd.DataFrame:
        """Pull AR Aging Summary."""
        data = self._get("reports/AgedReceivables")
        return self._parse_report(data, "AgedReceivables")

    # ------------------------------------------------------------------
    # Chart of Accounts
    # ------------------------------------------------------------------

    def chart_of_accounts(self) -> pd.DataFrame:
        """Return full Chart of Accounts."""
        data = self._get("query", {"query": "SELECT * FROM Account MAXRESULTS 1000"})
        rows = data.get("QueryResponse", {}).get("Account", [])
        return pd.DataFrame([
            {
                "id": a.get("Id"),
                "name": a.get("Name"),
                "fully_qualified_name": a.get("FullyQualifiedName"),
                "account_type": a.get("AccountType"),
                "account_sub_type": a.get("AccountSubType"),
                "active": a.get("Active"),
                "current_balance": a.get("CurrentBalance", 0.0),
            }
            for a in rows
        ])

    # ------------------------------------------------------------------
    # Report parser
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_report(data: dict, report_type: str) -> pd.DataFrame:
        """
        Flatten QBO report JSON into a DataFrame.
        QBO reports use a nested Row/Rows/ColData structure.
        """
        rows_out = []
        report = data.get("Rows", {})

        def _walk(rows, indent=0):
            for row in rows.get("Row", []):
                row_type = row.get("type", "")
                col_data = row.get("ColData", [])
                if col_data:
                    values = [c.get("value", "") for c in col_data]
                    rows_out.append({
                        "row_type": row_type,
                        "account": values[0] if len(values) > 0 else "",
                        "amount": _to_float(values[1] if len(values) > 1 else ""),
                        "indent": indent,
                    })
                if "Rows" in row:
                    _walk(row["Rows"], indent + 1)

        _walk(report)
        df = pd.DataFrame(rows_out)
        df.attrs["report_type"] = report_type
        return df


def _to_float(val: str) -> float:
    try:
        return float(str(val).replace(",", "").strip() or 0)
    except (ValueError, TypeError):
        return 0.0
