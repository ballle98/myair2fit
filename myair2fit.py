"""
myair2fit - Import myAir CPAP sleep data into Fitbit.

Reads SLEEP_RECORD.csv from a myAir export ZIP and POSTs each sleep session
to the Fitbit Sleep API (POST /1.2/user/-/sleep.json).
"""

import argparse
import base64
import csv
import hashlib
import http.server
import json
import os
import secrets
import sys
import tempfile
import threading
import time
import urllib.parse
import webbrowser
import zipfile
from datetime import datetime, date

import requests
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DEFAULT_START_TIME = "22:30"
FITBIT_AUTH_URI = "https://www.fitbit.com/oauth2/authorize"
FITBIT_TOKEN_URI = "https://api.fitbit.com/oauth2/token"
FITBIT_SLEEP_URL = "https://api.fitbit.com/1.2/user/-/sleep.json"
REDIRECT_PORT = 8080
REDIRECT_URI = f"http://localhost:{REDIRECT_PORT}/callback"
TOKEN_FILE = ".fitbit_tokens.json"
SCOPES = "sleep"

# ---------------------------------------------------------------------------
# OAuth 2.0 helpers (Authorization Code Grant with PKCE)
# ---------------------------------------------------------------------------

def _generate_pkce():
    """Return (code_verifier, code_challenge) for PKCE."""
    verifier = secrets.token_urlsafe(64)[:128]
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return verifier, challenge


class _CallbackHandler(http.server.BaseHTTPRequestHandler):
    """HTTP handler that captures the authorization code from the redirect."""

    auth_code = None
    error = None

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        if "code" in params:
            _CallbackHandler.auth_code = params["code"][0]
            self._respond("Authorization successful! You can close this tab.")
        else:
            _CallbackHandler.error = params.get("error", ["unknown"])[0]
            self._respond(f"Authorization failed: {_CallbackHandler.error}")

    def _respond(self, message):
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(
            f"<html><body><h2>{message}</h2></body></html>".encode()
        )

    def log_message(self, format, *args):
        pass  # suppress request logs


def _authorize(client_id: str) -> dict:
    """Run the full OAuth2 PKCE flow and return token dict."""
    verifier, challenge = _generate_pkce()
    state = secrets.token_urlsafe(16)

    auth_params = urllib.parse.urlencode({
        "response_type": "code",
        "client_id": client_id,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
        "state": state,
    })
    auth_url = f"{FITBIT_AUTH_URI}?{auth_params}"

    # Start local server to receive the callback
    server = http.server.HTTPServer(("localhost", REDIRECT_PORT), _CallbackHandler)
    thread = threading.Thread(target=server.handle_request, daemon=True)
    thread.start()

    print(f"Opening browser for Fitbit authorization...")
    print(f"If the browser doesn't open, visit:\n{auth_url}\n")
    webbrowser.open(auth_url)

    # Wait for callback (up to 120 seconds)
    thread.join(timeout=120)
    server.server_close()

    if _CallbackHandler.error:
        print(f"Authorization error: {_CallbackHandler.error}", file=sys.stderr)
        sys.exit(1)
    if not _CallbackHandler.auth_code:
        print("Timed out waiting for authorization.", file=sys.stderr)
        sys.exit(1)

    # Exchange code for tokens
    resp = requests.post(FITBIT_TOKEN_URI, data={
        "grant_type": "authorization_code",
        "code": _CallbackHandler.auth_code,
        "redirect_uri": REDIRECT_URI,
        "client_id": client_id,
        "code_verifier": verifier,
    }, headers={"Content-Type": "application/x-www-form-urlencoded"})

    if resp.status_code != 200:
        print(f"Token exchange failed ({resp.status_code}): {resp.text}", file=sys.stderr)
        sys.exit(1)

    tokens = resp.json()
    tokens["obtained_at"] = time.time()
    return tokens


def _refresh_tokens(client_id: str, refresh_token: str) -> dict:
    """Use a refresh token to get new access + refresh tokens."""
    resp = requests.post(FITBIT_TOKEN_URI, data={
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
    }, headers={"Content-Type": "application/x-www-form-urlencoded"})

    if resp.status_code != 200:
        print(f"Token refresh failed ({resp.status_code}): {resp.text}", file=sys.stderr)
        return None

    tokens = resp.json()
    tokens["obtained_at"] = time.time()
    return tokens


def _save_tokens(tokens: dict):
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def _load_tokens() -> dict | None:
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            return json.load(f)
    return None


def get_access_token(client_id: str) -> str:
    """Return a valid access token, authorizing or refreshing as needed."""
    tokens = _load_tokens()

    if tokens:
        # Check if token is expired (with 60s buffer)
        expires_in = tokens.get("expires_in", 28800)
        obtained_at = tokens.get("obtained_at", 0)
        if time.time() < obtained_at + expires_in - 60:
            return tokens["access_token"]

        # Try refresh
        print("Access token expired, refreshing...")
        new_tokens = _refresh_tokens(client_id, tokens["refresh_token"])
        if new_tokens:
            _save_tokens(new_tokens)
            return new_tokens["access_token"]
        print("Refresh failed, re-authorizing...")

    # Full authorization flow
    tokens = _authorize(client_id)
    _save_tokens(tokens)
    return tokens["access_token"]


# ---------------------------------------------------------------------------
# CSV parsing
# ---------------------------------------------------------------------------

def load_sleep_records(csv_path: str, start_date: date = None, end_date: date = None) -> list[dict]:
    """Parse SLEEP_RECORD.csv and return list of {date, usage_hours} dicts."""
    records = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            session_date_str = row.get("SESSION_DATE", "").strip()
            usage_hours_str = row.get("USAGE_HOURS", "").strip()

            if not session_date_str or not usage_hours_str:
                continue

            try:
                session_date = date.fromisoformat(session_date_str)
                usage_hours = float(usage_hours_str)
            except (ValueError, TypeError):
                continue

            if usage_hours <= 0:
                continue
            if start_date and session_date < start_date:
                continue
            if end_date and session_date > end_date:
                continue

            records.append({
                "date": session_date,
                "usage_hours": usage_hours,
            })

    # Sort chronologically
    records.sort(key=lambda r: r["date"])
    return records


# ---------------------------------------------------------------------------
# Fitbit API
# ---------------------------------------------------------------------------

def post_sleep(access_token: str, sleep_date: date, usage_hours: float,
               start_time: str = DEFAULT_START_TIME) -> dict:
    """POST a single sleep log to Fitbit. Returns response dict."""
    duration_ms = int(usage_hours * 3_600_000)

    resp = requests.post(
        FITBIT_SLEEP_URL,
        headers={"Authorization": f"Bearer {access_token}"},
        data={
            "date": sleep_date.isoformat(),
            "startTime": start_time,
            "duration": duration_ms,
        },
    )
    return {"status": resp.status_code, "body": resp.json() if resp.content else {}}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def find_sleep_csv(source_path: str) -> str:
    """Given a ZIP file or directory, return path to SLEEP_RECORD.csv."""
    if zipfile.is_zipfile(source_path):
        tmp_dir = tempfile.mkdtemp(prefix="myair2fit_")
        with zipfile.ZipFile(source_path, "r") as zf:
            zf.extractall(tmp_dir)
        # Search for SLEEP_RECORD.csv in extracted contents
        for root, _dirs, files in os.walk(tmp_dir):
            for fname in files:
                if fname.upper() == "SLEEP_RECORD.CSV":
                    return os.path.join(root, fname)
        print(f"SLEEP_RECORD.csv not found in ZIP archive.", file=sys.stderr)
        sys.exit(1)

    if os.path.isdir(source_path):
        candidate = os.path.join(source_path, "SLEEP_RECORD.csv")
        if os.path.isfile(candidate):
            return candidate
        print(f"SLEEP_RECORD.csv not found in {source_path}", file=sys.stderr)
        sys.exit(1)

    if os.path.isfile(source_path) and source_path.upper().endswith(".CSV"):
        return source_path

    print(f"Cannot process: {source_path}", file=sys.stderr)
    sys.exit(1)


def main():
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Import myAir CPAP sleep data into Fitbit."
    )
    parser.add_argument("source", help="Path to myAir export ZIP, extracted directory, or SLEEP_RECORD.csv")
    parser.add_argument("--start-date", type=date.fromisoformat,
                        help="Only import records on or after this date (yyyy-MM-dd)")
    parser.add_argument("--end-date", type=date.fromisoformat,
                        help="Only import records on or before this date (yyyy-MM-dd)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview records without posting to Fitbit")
    parser.add_argument("--start-time", default=DEFAULT_START_TIME,
                        help=f"Sleep start time in HH:mm (default: {DEFAULT_START_TIME})")
    args = parser.parse_args()

    client_id = os.getenv("FITBIT_CLIENT_ID")
    if not client_id and not args.dry_run:
        print("FITBIT_CLIENT_ID not set. Copy .env.example to .env and add your Client ID.",
              file=sys.stderr)
        sys.exit(1)

    csv_path = find_sleep_csv(args.source)
    records = load_sleep_records(csv_path, args.start_date, args.end_date)

    if not records:
        print("No sleep records found matching the criteria.")
        return

    print(f"Found {len(records)} sleep record(s) to import.\n")

    if args.dry_run:
        print(f"{'Date':<14} {'Hours':>6}  {'Duration (ms)':>14}  Start")
        print("-" * 52)
        for rec in records:
            duration_ms = int(rec["usage_hours"] * 3_600_000)
            print(f"{rec['date'].isoformat():<14} {rec['usage_hours']:>6.2f}  {duration_ms:>14,}  {args.start_time}")
        print(f"\nDry run complete. Use without --dry-run to POST to Fitbit.")
        return

    access_token = get_access_token(client_id)

    success = 0
    errors = 0
    for i, rec in enumerate(records, 1):
        date_str = rec["date"].isoformat()
        hours = rec["usage_hours"]
        print(f"[{i}/{len(records)}] {date_str}  {hours:.2f}h ... ", end="", flush=True)

        result = post_sleep(access_token, rec["date"], hours, args.start_time)

        if result["status"] in (200, 201):
            print("OK")
            success += 1
        elif result["status"] == 401:
            # Token may have expired mid-run, try refresh
            print("token expired, refreshing... ", end="", flush=True)
            access_token = get_access_token(client_id)
            result = post_sleep(access_token, rec["date"], hours, args.start_time)
            if result["status"] in (200, 201):
                print("OK")
                success += 1
            else:
                print(f"FAILED ({result['status']}): {result['body']}")
                errors += 1
        else:
            print(f"FAILED ({result['status']}): {result['body']}")
            errors += 1

    print(f"\nDone. {success} imported, {errors} failed.")


if __name__ == "__main__":
    main()
