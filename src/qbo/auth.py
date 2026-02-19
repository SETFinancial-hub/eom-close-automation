"""
QBO OAuth 2.0 Authentication Handler
--------------------------------------
Handles the OAuth 2.0 flow for QuickBooks Online.
For internal use: SET Financial Corporation only.

Flow:
  1. Generate authorization URL → user approves in browser
  2. Exchange authorization code for tokens
  3. Store/refresh tokens via environment or tokens.json (gitignored)
"""
import os
import json
import time
import webbrowser
from urllib.parse import urlencode, urlparse, parse_qs
from pathlib import Path

import requests
from requests.auth import HTTPBasicAuth

# Intuit OAuth 2.0 endpoints
AUTH_ENDPOINT = "https://appcenter.intuit.com/connect/oauth2"
TOKEN_ENDPOINT = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
REVOKE_ENDPOINT = "https://developer.api.intuit.com/v2/oauth2/tokens/revoke"

# QBO scopes needed for financial reporting
SCOPES = [
    "com.intuit.quickbooks.accounting",
]

TOKEN_FILE = Path(__file__).parent.parent.parent / "tokens.json"  # gitignored


class QBOAuth:
    """Manages OAuth 2.0 tokens for QuickBooks Online."""

    def __init__(self):
        self.client_id = os.environ["QBO_CLIENT_ID"]
        self.client_secret = os.environ["QBO_CLIENT_SECRET"]
        self.redirect_uri = os.environ["QBO_REDIRECT_URI"]
        self.realm_id = os.environ.get("QBO_REALM_ID", "")
        self._tokens: dict = {}
        self._load_tokens()

    # ------------------------------------------------------------------
    # Token persistence
    # ------------------------------------------------------------------

    def _load_tokens(self):
        """Load tokens from file or environment."""
        if TOKEN_FILE.exists():
            with open(TOKEN_FILE) as f:
                self._tokens = json.load(f)
        else:
            self._tokens = {
                "access_token": os.environ.get("QBO_ACCESS_TOKEN", ""),
                "refresh_token": os.environ.get("QBO_REFRESH_TOKEN", ""),
                "expiry": float(os.environ.get("QBO_TOKEN_EXPIRY", "0")),
                "realm_id": self.realm_id,
            }

    def _save_tokens(self):
        """Persist tokens to gitignored file."""
        TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_FILE, "w") as f:
            json.dump(self._tokens, f, indent=2)

    # ------------------------------------------------------------------
    # Authorization flow
    # ------------------------------------------------------------------

    def get_authorization_url(self) -> str:
        """Return URL for user to authorize the app."""
        params = {
            "client_id": self.client_id,
            "response_type": "code",
            "scope": " ".join(SCOPES),
            "redirect_uri": self.redirect_uri,
            "state": "setfinancial_eom",
        }
        return f"{AUTH_ENDPOINT}?{urlencode(params)}"

    def exchange_code(self, authorization_code: str, realm_id: str):
        """Exchange authorization code for access + refresh tokens."""
        resp = requests.post(
            TOKEN_ENDPOINT,
            auth=HTTPBasicAuth(self.client_id, self.client_secret),
            headers={"Accept": "application/json"},
            data={
                "grant_type": "authorization_code",
                "code": authorization_code,
                "redirect_uri": self.redirect_uri,
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        self._tokens = {
            "access_token": data["access_token"],
            "refresh_token": data["refresh_token"],
            "expiry": time.time() + data.get("expires_in", 3600) - 60,
            "realm_id": realm_id,
        }
        self._save_tokens()
        print(f"[QBOAuth] Tokens obtained for realm {realm_id}")

    # ------------------------------------------------------------------
    # Token refresh
    # ------------------------------------------------------------------

    def _refresh(self):
        """Refresh access token using refresh token."""
        resp = requests.post(
            TOKEN_ENDPOINT,
            auth=HTTPBasicAuth(self.client_id, self.client_secret),
            headers={"Accept": "application/json"},
            data={
                "grant_type": "refresh_token",
                "refresh_token": self._tokens["refresh_token"],
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        self._tokens["access_token"] = data["access_token"]
        self._tokens["refresh_token"] = data.get("refresh_token", self._tokens["refresh_token"])
        self._tokens["expiry"] = time.time() + data.get("expires_in", 3600) - 60
        self._save_tokens()

    @property
    def access_token(self) -> str:
        """Return a valid access token, refreshing if needed."""
        if not self._tokens.get("access_token"):
            raise RuntimeError(
                "No QBO access token. Run: python src/qbo/setup_oauth.py"
            )
        if time.time() >= self._tokens.get("expiry", 0):
            self._refresh()
        return self._tokens["access_token"]

    @property
    def realm_id(self) -> str:
        return self._tokens.get("realm_id", os.environ.get("QBO_REALM_ID", ""))

    @realm_id.setter
    def realm_id(self, value: str):
        self._tokens["realm_id"] = value

    def revoke(self):
        """Revoke tokens (disconnect app)."""
        token = self._tokens.get("refresh_token") or self._tokens.get("access_token")
        if not token:
            print("[QBOAuth] No tokens to revoke.")
            return
        resp = requests.post(
            REVOKE_ENDPOINT,
            auth=HTTPBasicAuth(self.client_id, self.client_secret),
            headers={"Accept": "application/json"},
            json={"token": token},
            timeout=30,
        )
        resp.raise_for_status()
        self._tokens = {}
        if TOKEN_FILE.exists():
            TOKEN_FILE.unlink()
        print("[QBOAuth] Tokens revoked and deleted.")
