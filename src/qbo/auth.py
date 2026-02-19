"""
QBO OAuth 2.0 Authentication Handler
--------------------------------------
Handles the OAuth 2.0 flow for QuickBooks Online.
For internal use: SET Financial Corporation only.

Flow:
  1. Generate authorization URL → user approves in browser
  2. Exchange authorization code for tokens
  3. Store/refresh tokens via environment or tokens.json (gitignored)

Endpoints are loaded from Intuit's discovery document (not hardcoded)
per Intuit's security requirements.
"""
import os
import json
import time
import secrets
from urllib.parse import urlencode
from pathlib import Path

import requests
from requests.auth import HTTPBasicAuth

# Intuit discovery document — provides latest OAuth 2.0 endpoints
DISCOVERY_DOC_URL = "https://developer.intuit.com/.well-known/openid_configuration"

# QBO scopes needed for financial reporting
SCOPES = [
    "com.intuit.quickbooks.accounting",
]

TOKEN_FILE = Path(__file__).parent.parent.parent / "tokens.json"  # gitignored

# Errors that mean the refresh token is invalid — user must re-authorize
_REAUTH_ERRORS = {"invalid_grant", "AuthenticationFailed", "token_expired"}


def _load_discovery() -> dict:
    """
    Fetch Intuit's OpenID discovery document to get current OAuth endpoints.
    Cached in memory for the process lifetime.
    """
    if not hasattr(_load_discovery, "_cache"):
        resp = requests.get(DISCOVERY_DOC_URL, timeout=10)
        resp.raise_for_status()
        _load_discovery._cache = resp.json()
    return _load_discovery._cache


class TokenExpiredError(Exception):
    """Refresh token has expired or been revoked — user must re-authorize."""


class QBOAuth:
    """Manages OAuth 2.0 tokens for QuickBooks Online."""

    def __init__(self):
        self.client_id = os.environ["QBO_CLIENT_ID"]
        self.client_secret = os.environ["QBO_CLIENT_SECRET"]
        self.redirect_uri = os.environ["QBO_REDIRECT_URI"]
        self._realm_id = os.environ.get("QBO_REALM_ID", "")
        self._tokens: dict = {}
        self._csrf_state: str = ""
        self._load_tokens()

    # ------------------------------------------------------------------
    # Endpoint discovery
    # ------------------------------------------------------------------

    @property
    def _auth_endpoint(self) -> str:
        return _load_discovery().get(
            "authorization_endpoint",
            "https://appcenter.intuit.com/connect/oauth2",
        )

    @property
    def _token_endpoint(self) -> str:
        return _load_discovery().get(
            "token_endpoint",
            "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer",
        )

    @property
    def _revoke_endpoint(self) -> str:
        return _load_discovery().get(
            "revocation_endpoint",
            "https://developer.api.intuit.com/v2/oauth2/tokens/revoke",
        )

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
                "realm_id": self._realm_id,
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
        """
        Return URL for user to authorize the app.
        Generates a cryptographically random CSRF state token.
        """
        self._csrf_state = secrets.token_urlsafe(16)
        params = {
            "client_id": self.client_id,
            "response_type": "code",
            "scope": " ".join(SCOPES),
            "redirect_uri": self.redirect_uri,
            "state": self._csrf_state,
        }
        return f"{self._auth_endpoint}?{urlencode(params)}"

    def validate_csrf_state(self, returned_state: str):
        """Validate state parameter to prevent CSRF attacks."""
        if not self._csrf_state:
            return  # state not tracked in this session (e.g. after restart)
        if returned_state != self._csrf_state:
            raise ValueError(
                f"CSRF state mismatch. Expected {self._csrf_state!r}, got {returned_state!r}. "
                "Do not reuse authorization URLs."
            )

    def exchange_code(self, authorization_code: str, realm_id: str, state: str = ""):
        """Exchange authorization code for access + refresh tokens."""
        if state:
            self.validate_csrf_state(state)

        resp = requests.post(
            self._token_endpoint,
            auth=HTTPBasicAuth(self.client_id, self.client_secret),
            headers={"Accept": "application/json"},
            data={
                "grant_type": "authorization_code",
                "code": authorization_code,
                "redirect_uri": self.redirect_uri,
            },
            timeout=30,
        )
        self._handle_token_response(resp)
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
    # Token refresh with error handling
    # ------------------------------------------------------------------

    def _refresh(self):
        """
        Refresh access token using refresh token.
        Raises TokenExpiredError if the refresh token is invalid/expired,
        prompting the user to re-authorize via setup_oauth.py.
        """
        resp = requests.post(
            self._token_endpoint,
            auth=HTTPBasicAuth(self.client_id, self.client_secret),
            headers={"Accept": "application/json"},
            data={
                "grant_type": "refresh_token",
                "refresh_token": self._tokens["refresh_token"],
            },
            timeout=30,
        )
        self._handle_token_response(resp)
        data = resp.json()
        self._tokens["access_token"] = data["access_token"]
        self._tokens["refresh_token"] = data.get(
            "refresh_token", self._tokens["refresh_token"]
        )
        self._tokens["expiry"] = time.time() + data.get("expires_in", 3600) - 60
        self._save_tokens()

    def _handle_token_response(self, resp: requests.Response):
        """
        Inspect token endpoint responses for specific OAuth errors
        and raise typed exceptions for proper handling upstream.
        """
        if resp.ok:
            return
        try:
            body = resp.json()
        except Exception:
            resp.raise_for_status()
            return

        error_code = body.get("error", "") or body.get("code", "")

        if error_code in _REAUTH_ERRORS:
            # Clear stored tokens — user must go through OAuth flow again
            self._tokens = {}
            if TOKEN_FILE.exists():
                TOKEN_FILE.unlink()
            raise TokenExpiredError(
                "QuickBooks refresh token has expired or been revoked. "
                "Re-authorize by running: python src/qbo/setup_oauth.py"
            )

        resp.raise_for_status()

    @property
    def access_token(self) -> str:
        """Return a valid access token, refreshing automatically if expired."""
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
        """Revoke tokens and clear local storage (disconnect app)."""
        token = self._tokens.get("refresh_token") or self._tokens.get("access_token")
        if not token:
            print("[QBOAuth] No tokens to revoke.")
            return
        resp = requests.post(
            self._revoke_endpoint,
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
