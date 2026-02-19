"""
QBO OAuth Setup Script
-----------------------
Run this ONCE to complete the OAuth 2.0 flow and store tokens.

Usage:
    python src/qbo/setup_oauth.py

Steps:
  1. Opens your browser to the Intuit authorization page
  2. You approve access in your QuickBooks account
  3. Intuit redirects to the redirect_uri with a code parameter
  4. Paste the full redirect URL back here
  5. Tokens are saved to tokens.json (gitignored)
"""
import os
import sys
import webbrowser
from urllib.parse import urlparse, parse_qs
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from dotenv import load_dotenv
load_dotenv()

from src.qbo.auth import QBOAuth


def main():
    print("=" * 60)
    print("  SET Financial — QuickBooks OAuth Setup")
    print("=" * 60)

    auth = QBOAuth()

    url = auth.get_authorization_url()
    print(f"\n[1] Opening browser for authorization...\n")
    print(f"    URL: {url}\n")

    try:
        webbrowser.open(url)
    except Exception:
        print("    (Could not auto-open browser — please copy the URL above)")

    print("[2] After approving access in QuickBooks, you will be redirected.")
    print("    The redirect URL will look like:")
    print("    https://setfinancial-hub.github.io/eom-close-automation/disconnect?code=...&realmId=...\n")

    redirect_url = input("    Paste the full redirect URL here: ").strip()

    parsed = urlparse(redirect_url)
    params = parse_qs(parsed.query)

    code = params.get("code", [None])[0]
    realm_id = params.get("realmId", [None])[0]
    state = params.get("state", [None])[0]

    if not code or not realm_id:
        print("\n[ERROR] Could not parse 'code' or 'realmId' from the URL.")
        print("        Make sure you pasted the full redirect URL.")
        sys.exit(1)

    print(f"\n[3] Validating CSRF state and exchanging authorization code for tokens...")
    auth.exchange_code(code, realm_id, state=state or "")

    print(f"\n[SUCCESS] Tokens saved. Realm ID: {realm_id}")
    print(f"          You can now run the MCP server: python src/qbo/mcp_server.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
