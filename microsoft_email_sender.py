"""
Bulk Email Sender with Certificate & Status Report Attachments
==============================================================
Reads from the dataset Excel file, validates emails, derives certificate
filenames, attaches both files, sends in batches, and logs all results.

Uses Microsoft Graph API with OAuth 2.0 Device Code Flow (delegated,
Mail.Send permission — no admin consent required).

Single master log
-----------------
Every run appends to ONE file: <output>/master_results.txt (plus an Excel
view, master_results.xlsx). On startup the script reads the master file and
skips anyone already marked "sent", so no --resume argument is needed.

If the master file exists, it is only appended to (never rewritten or migrated
into). Only when NO master file is found, and old results_*.txt files exist in
the output folder, are their "sent" rows merged into a new master (one-time).
The Excel copy is likewise appended to if it exists, and built once if not.

Setup
-----
1. Register an app in Azure Active Directory (Entra ID).
2. Under "Authentication", add a Mobile/Desktop platform and enable
   "Allow public client flows" (needed for device code).
3. Under "API permissions", add:
      Microsoft Graph → Delegated → Mail.Send
   (No admin consent needed for delegated Mail.Send.)
4. Note your Tenant ID and Client ID.

First run
---------
The script will print a URL and a short code, e.g.:
    To sign in, visit https://microsoft.com/devicelogin and enter code ABCD1234
Open the URL in any browser, sign in as the sending mailbox owner, and approve.
The token (including refresh token) is cached to ~/.graph_token_cache.json and
reused on subsequent runs — you won't be prompted again unless the token expires
(refresh tokens last 90 days for M365 accounts by default).

Usage:
    python send_emails_graph.py \
        --dataset testing_email_dataset_060320.xlsx \
        --attachments ./attachments \
        --tenant-id  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" \
        --client-id  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" \
        --sender "sender@yourdomain.com" \
        --batch-size 10 \
        --batch-delay 60 \
        --send-delay 3 \
        --session-limit 100
"""

import argparse
import base64
import json
import logging
import re
import sys
import time
from datetime import datetime
from functools import wraps
from pathlib import Path

import dns.resolver
import pandas as pd
import requests

# ── Logging ──────────────────────────────────────────────────────────────────

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"email_run_{timestamp}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ── CC address — receives every email ────────────────────────────────────────
CC_RECIPIENT = "ffpclusterdevelopment@gmail.com"

# ── Graph API endpoints ───────────────────────────────────────────────────────
GRAPH_DEVICE_CODE_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"
GRAPH_TOKEN_URL       = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
GRAPH_SEND_URL        = "https://graph.microsoft.com/v1.0/me/sendMail"
GRAPH_SCOPE           = "https://graph.microsoft.com/Mail.Send offline_access"
TOKEN_CACHE_PATH      = Path.home() / ".graph_token_cache.json"

# ── Master log name ───────────────────────────────────────────────────────────
MASTER_LOG_NAME = "master_results"

# ── Retry decorator ───────────────────────────────────────────────────────────

def retry(max_attempts: int = 3, delay: float = 2.0, backoff: float = 2.0,
          exceptions: tuple = (Exception,)):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            current_delay = delay
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts:
                        log.error("%s failed after %d attempts. Error: %s",
                                  func.__name__, max_attempts, e)
                        raise
                    log.warning("%s attempt %d failed: %s. Retrying in %.2fs...",
                                func.__name__, attempt, e, current_delay)
                    time.sleep(current_delay)
                    current_delay *= backoff
        return wrapper
    return decorator

# ── OAuth 2.0 Device Code token manager ──────────────────────────────────────

class GraphTokenManager:
    """
    Delegated auth via Device Code Flow.

    - On first use (or if the cache is missing/expired) it prints a URL + code
      and blocks until the user authenticates in a browser.
    - Caches access + refresh tokens to TOKEN_CACHE_PATH.
    - Silently refreshes with the refresh token on subsequent runs.
    """

    def __init__(self, tenant_id: str, client_id: str):
        self._tenant_id  = tenant_id
        self._client_id  = client_id
        self._token: str | None = None
        self._refresh_token: str | None = None
        self._expires_at: float = 0.0
        self._load_cache()

    # ── public ────────────────────────────────────────────────────────────────

    def get_token(self) -> str:
        if self._token and time.time() < self._expires_at - 60:
            return self._token
        if self._refresh_token:
            try:
                self._refresh()
                return self._token  # type: ignore[return-value]
            except Exception as e:
                log.warning("Refresh token failed (%s) — re-authenticating via device code.", e)
        self._device_code_flow()
        return self._token  # type: ignore[return-value]

    # ── private ───────────────────────────────────────────────────────────────

    def _token_url(self) -> str:
        return GRAPH_TOKEN_URL.format(tenant_id=self._tenant_id)

    def _device_code_flow(self):
        """Initiate device code flow; block until user authenticates."""
        # Step 1 — request device + user codes
        dc_resp = requests.post(
            GRAPH_DEVICE_CODE_URL.format(tenant_id=self._tenant_id),
            data={"client_id": self._client_id, "scope": GRAPH_SCOPE},
            timeout=30,
        )
        dc_resp.raise_for_status()
        dc = dc_resp.json()

        # Step 2 — show instructions to user
        print("\n" + "=" * 60)
        print("ACTION REQUIRED — Microsoft authentication")
        print(f"  1. Open:  {dc['verification_uri']}")
        print(f"  2. Enter: {dc['user_code']}")
        print("=" * 60 + "\n")
        log.info("Waiting for device code authentication (expires in %ss)...",
                 dc.get("expires_in", "?"))

        # Step 3 — poll until granted or expired
        interval = int(dc.get("interval", 5))
        deadline = time.time() + int(dc.get("expires_in", 900))

        while time.time() < deadline:
            time.sleep(interval)
            poll = requests.post(
                self._token_url(),
                data={
                    "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
                    "client_id":   self._client_id,
                    "device_code": dc["device_code"],
                },
                timeout=30,
            )
            payload = poll.json()
            error = payload.get("error")

            if error == "authorization_pending":
                continue
            elif error == "slow_down":
                interval += 5
                continue
            elif error:
                raise RuntimeError(f"Device code auth failed: {error} — "
                                   f"{payload.get('error_description', '')}")

            self._store_tokens(payload)
            log.info("Device code authentication successful.")
            return

        raise TimeoutError("Device code flow timed out — user did not authenticate in time.")

    def _refresh(self):
        resp = requests.post(
            self._token_url(),
            data={
                "grant_type":    "refresh_token",
                "client_id":     self._client_id,
                "refresh_token": self._refresh_token,
                "scope":         GRAPH_SCOPE,
            },
            timeout=30,
        )
        resp.raise_for_status()
        self._store_tokens(resp.json())
        log.info("Graph API token refreshed silently.")

    def _store_tokens(self, payload: dict):
        self._token         = payload["access_token"]
        self._refresh_token = payload.get("refresh_token", self._refresh_token)
        self._expires_at    = time.time() + int(payload.get("expires_in", 3600))
        self._save_cache()

    def _save_cache(self):
        try:
            TOKEN_CACHE_PATH.write_text(json.dumps({
                "access_token":  self._token,
                "refresh_token": self._refresh_token,
                "expires_at":    self._expires_at,
            }))
        except Exception as e:
            log.warning("Could not save token cache: %s", e)

    def _load_cache(self):
        try:
            data = json.loads(TOKEN_CACHE_PATH.read_text())
            self._token         = data.get("access_token")
            self._refresh_token = data.get("refresh_token")
            self._expires_at    = float(data.get("expires_at", 0))
            log.info("Loaded cached token from %s.", TOKEN_CACHE_PATH)
        except FileNotFoundError:
            pass  # first run — normal
        except Exception as e:
            log.warning("Could not load token cache (%s) — will re-authenticate.", e)

# ── Master log: load already-sent emails ──────────────────────────────────────

def load_sent_emails(master_path: Path) -> set[str]:
    """Return the set of emails marked 'sent' in the master results file."""
    already_sent: set[str] = set()
    if not master_path.exists():
        log.info("No master log yet (%s) — nobody will be skipped.", master_path)
        return already_sent
    try:
        df = pd.read_csv(master_path, sep="\t", dtype=str).fillna("")
        sent_rows = df[df["result"].str.strip().str.lower() == "sent"]
        already_sent = {e.strip().lower() for e in sent_rows["email"] if e.strip()}
        log.info("Master log: %d unique already-sent emails loaded from %s "
                 "— these will be skipped.", len(already_sent), master_path.name)
    except Exception as e:
        # Fail loudly: continuing with an empty set could cause duplicate sends.
        log.error("Could not read master log %s: %s", master_path, e)
        raise
    return already_sent

# ── Email validation ──────────────────────────────────────────────────────────

EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

TRUSTED_DOMAINS = {
    "gmail.com", "googlemail.com",
    "yahoo.com", "yahoo.co.uk", "yahoo.com.ng",
    "outlook.com", "hotmail.com", "hotmail.co.uk", "live.com",
    "icloud.com", "me.com", "mac.com",
    "protonmail.com", "proton.me",
    "zoho.com",
}


def validate_email_format(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email.strip()))


def check_mx_record(domain: str) -> bool:
    try:
        records = dns.resolver.resolve(domain, "MX", lifetime=10)
        return len(records) > 0
    except Exception:
        return False


def validate_email(email: str) -> tuple[bool, str]:
    email = email.strip().lower()
    if not validate_email_format(email):
        return False, "invalid_format"
    domain = email.split("@")[1]
    if domain in TRUSTED_DOMAINS:
        return True, "ok"
    if not check_mx_record(domain):
        return False, "no_mx_record"
    return True, "ok"

# ── File name derivation ──────────────────────────────────────────────────────

def derive_certificate_name(status_report_name: str) -> str:
    return re.sub(
        r"Status\s+Report\s+CAC",
        "CERTIFICATE",
        status_report_name,
        flags=re.IGNORECASE,
    )

# ── Build Graph API JSON payload ──────────────────────────────────────────────

def _encode_attachment(filepath: Path) -> dict:
    """Return a Graph API fileAttachment object for the given file."""
    with open(filepath, "rb") as f:
        content = base64.b64encode(f.read()).decode("utf-8")
    return {
        "@odata.type":  "#microsoft.graph.fileAttachment",
        "name":         filepath.name,
        "contentType":  "application/pdf",
        "contentBytes": content,
    }


def build_message(
    sender: str,
    recipient: str,
    first_name: str,
    surname: str,
    business_name: str,
    status_report_path: Path | None,
    certificate_path: Path | None,
) -> dict:
    """Return the Graph API /sendMail request body as a Python dict."""

    body_text = f"""\
Dear {first_name.title()} {surname.title()},

Congratulations on the successful registration of {business_name.title()} with the Corporate Affairs Commission (CAC) under the MSME Competitiveness and Industrial Productivity Program (MCIPP) Implemented by Ethnocentrique and funded by the Abia State Government.

Please find attached your:
  1. Status Report
  2. Certificate of Registration

Kindly keep these documents safe, as they serve as official proof of your business registration.

Should you have any questions or require further assistance, do not hesitate to reach out.

Best regards,

Adedapo Adegboyega-Conde
Team Lead,
MCIPP
"""

    attachments = []
    for path in [status_report_path, certificate_path]:
        if path and path.exists():
            attachments.append(_encode_attachment(path))
        elif path:
            log.warning("Attachment not found, skipping: %s", path)

    message: dict = {
        "subject": f"Your CAC Registration Documents - {business_name}",
        "body": {
            "contentType": "Text",
            "content":     body_text,
        },
        "toRecipients": [
            {"emailAddress": {"address": recipient}},
        ],
        "ccRecipients": [
            {"emailAddress": {"address": CC_RECIPIENT}},
        ],
    }

    if attachments:
        message["attachments"] = attachments

    return {"message": message, "saveToSentItems": "true"}

# ── Result tracking (single master file) ──────────────────────────────────────

class ResultTracker:
    COLUMNS = [
        "timestamp", "run_id", "row_index", "surname", "firstname",
        "email", "business_name", "status_report_file",
        "certificate_file", "result", "reason",
    ]

    def __init__(self, output_dir: Path):
        output_dir.mkdir(exist_ok=True)
        self.output_dir = output_dir
        self.log_path   = output_dir / f"{MASTER_LOG_NAME}.txt"
        self.xlsx_path  = output_dir / f"{MASTER_LOG_NAME}.xlsx"
        self.records: list[dict] = []   # this run only (for the summary)

        if self.log_path.exists():
            # Master detected -> never touch or rewrite it; just append later.
            log.info("Master log found: %s — appending to it (no migration).",
                     self.log_path.resolve())
            if self.log_path.stat().st_size == 0:
                # Exists but completely empty: add the header only.
                with open(self.log_path, "w", encoding="utf-8") as f:
                    f.write("\t".join(self.COLUMNS) + "\n")
        else:
            # No master detected -> this is the only time we migrate/create.
            log.info("No master log at %s — creating it.", self.log_path.resolve())
            if not self._migrate_old_results():
                with open(self.log_path, "w", encoding="utf-8") as f:
                    f.write("\t".join(self.COLUMNS) + "\n")

    def _migrate_old_results(self) -> bool:
        """
        One-time: fold 'sent' rows from legacy results_*.txt files into a new
        master file. Returns True if a master was created.
        """
        old_files = sorted(self.output_dir.glob("results_*.txt"))
        if not old_files:
            return False

        frames = []
        for f in old_files:
            try:
                frames.append(pd.read_csv(f, sep="\t", dtype=str).fillna(""))
            except Exception as e:
                log.warning("Migration: could not read %s (%s) — skipped.", f.name, e)
        if not frames:
            return False

        df = pd.concat(frames, ignore_index=True)
        df = df[df["result"].str.strip().str.lower() == "sent"].copy()
        df["_e"] = df["email"].str.strip().str.lower()
        df = df.drop_duplicates("_e").drop(columns="_e")
        df["run_id"] = "migrated"
        df = df.reindex(columns=self.COLUMNS).fillna("")
        df.to_csv(self.log_path, sep="\t", index=False, encoding="utf-8")
        log.info("Migration: created master log with %d sent records from %d old file(s).",
                 len(df), len(old_files))
        return True

    def record(self, **kwargs):
        kwargs.setdefault("timestamp", datetime.now().isoformat(timespec="seconds"))
        kwargs["run_id"] = timestamp
        self.records.append(kwargs)
        # Strip tabs/newlines so error messages can't corrupt the TSV rows
        row = "\t".join(
            re.sub(r"[\t\r\n]+", " ", str(kwargs.get(c, ""))) for c in self.COLUMNS
        )
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(row + "\n")

    def save_excel(self):
        """
        Excel view of the master.
        - If the .xlsx already exists: append only this run's rows to it.
        - If it doesn't exist: build it once from the full master .txt.
        """
        try:
            if self.xlsx_path.exists():
                if not self.records:
                    log.info("No new rows this run — Excel master left untouched.")
                    return
                from openpyxl import load_workbook
                wb = load_workbook(self.xlsx_path)
                ws = wb.active
                for r in self.records:
                    ws.append([str(r.get(c, "")) for c in self.COLUMNS])
                wb.save(self.xlsx_path)
                log.info("Appended %d rows to Excel master -> %s",
                         len(self.records), self.xlsx_path)
            else:
                df = pd.read_csv(self.log_path, sep="\t", dtype=str).fillna("")
                df.to_excel(self.xlsx_path, index=False)
                log.info("Excel master created -> %s (%d rows)", self.xlsx_path, len(df))
        except PermissionError:
            log.warning("Close %s in Excel. The .txt master is up to date; "
                        "this run's rows were NOT added to the Excel copy.",
                        self.xlsx_path)

    def summary(self):
        counts: dict[str, int] = {}
        for r in self.records:
            counts[r["result"]] = counts.get(r["result"], 0) + 1
        return counts

# ── Core send logic ───────────────────────────────────────────────────────────

@retry(max_attempts=3, delay=2.0, backoff=2.0, exceptions=(Exception,))
def send_one(token_mgr: GraphTokenManager, payload: dict):
    """POST the sendMail payload via Graph API. Raises on non-2xx responses."""
    headers = {
        "Authorization": f"Bearer {token_mgr.get_token()}",
        "Content-Type":  "application/json",
    }
    resp = requests.post(GRAPH_SEND_URL, headers=headers, json=payload, timeout=60)
    if resp.status_code == 202:
        return  # Graph returns 202 Accepted on success — no body
    raise RuntimeError(
        f"Graph API returned HTTP {resp.status_code}: {resp.text[:400]}"
    )


def run(args):
    attachments_dir = Path(args.attachments)
    output_dir      = Path(args.output)
    tracker         = ResultTracker(output_dir)

    already_sent  = load_sent_emails(tracker.log_path)
    session_limit = args.session_limit
    session_sent  = 0

    log.info("Loading dataset: %s", args.dataset)
    df = pd.read_excel(args.dataset, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df = df.fillna("")
    log.info("  %d rows loaded.", len(df))

    if session_limit:
        log.info("  Session limit: %d sends — will stop after reaching it.", session_limit)

    token_mgr = GraphTokenManager(args.tenant_id, args.client_id)
    # Eagerly authenticate before iterating rows
    token_mgr.get_token()

    sent = failed = invalid = skipped = resumed = 0
    session_limit_hit = False

    for batch_start in range(0, len(df), args.batch_size):

        if session_limit_hit:
            break

        batch = df.iloc[batch_start : batch_start + args.batch_size]
        log.info(
            "--- Batch %d: rows %d-%d ---",
            batch_start // args.batch_size + 1,
            batch_start + 1,
            batch_start + len(batch),
        )

        for idx, row in batch.iterrows():

            if session_limit is not None and session_sent >= session_limit:
                log.info("Session limit of %d reached (%d sent this run). Stopping.",
                         session_limit, session_sent)
                session_limit_hit = True
                break

            row_num   = idx + 2
            email_raw = row.get("EMAIL", "").strip()
            firstname = row.get("FIRSTNAME", "").strip()
            surname   = row.get("SURNAME", "").strip()
            biz_name  = row.get("Business Name", "").strip()
            src_file  = row.get("Source File", "").strip()

            if not email_raw:
                log.warning("Row %d: empty email — skipping", row_num)
                tracker.record(
                    row_index=row_num, surname=surname, firstname=firstname,
                    email="", business_name=biz_name,
                    status_report_file=src_file, certificate_file="",
                    result="skipped", reason="empty_email",
                )
                skipped += 1
                continue

            # Already sent in a previous run (per master log) — skip silently,
            # without writing another row, so the master doesn't keep growing.
            if email_raw.lower() in already_sent:
                log.info("Row %d: %s already sent — skipping", row_num, email_raw)
                resumed += 1
                continue

            ok, reason = validate_email(email_raw)
            if not ok:
                log.warning("Row %d: %s -> INVALID (%s)", row_num, email_raw, reason)
                tracker.record(
                    row_index=row_num, surname=surname, firstname=firstname,
                    email=email_raw, business_name=biz_name,
                    status_report_file=src_file, certificate_file="",
                    result="invalid", reason=reason,
                )
                invalid += 1
                continue

            status_path = attachments_dir / src_file if src_file else None
            cert_name   = derive_certificate_name(src_file) if src_file else ""
            cert_path   = attachments_dir / cert_name if cert_name else None

            missing = []
            if status_path and not status_path.exists():
                missing.append(f"status_report:{src_file}")
            if cert_path and not cert_path.exists():
                missing.append(f"certificate:{cert_name}")
            if missing and not args.allow_missing_attachments:
                log.warning("Row %d: missing file(s): %s — skipping", row_num, missing)
                tracker.record(
                    row_index=row_num, surname=surname, firstname=firstname,
                    email=email_raw, business_name=biz_name,
                    status_report_file=src_file, certificate_file=cert_name,
                    result="failed", reason=f"missing_attachments:{','.join(missing)}",
                )
                failed += 1
                continue

            payload = build_message(
                sender=args.sender,
                recipient=email_raw,
                first_name=firstname,
                surname=surname,
                business_name=biz_name,
                status_report_path=status_path,
                certificate_path=cert_path,
            )

            try:
                send_one(token_mgr, payload)
                log.info("Row %d: SENT -> %s (%s)", row_num, email_raw, biz_name)
                tracker.record(
                    row_index=row_num, surname=surname, firstname=firstname,
                    email=email_raw, business_name=biz_name,
                    status_report_file=src_file, certificate_file=cert_name,
                    result="sent", reason="ok",
                )
                sent += 1
                session_sent += 1
                # Also guard against the same email appearing twice in the dataset
                already_sent.add(email_raw.lower())
            except Exception as e:
                log.error("Row %d: send error — %s", row_num, e)
                tracker.record(
                    row_index=row_num, surname=surname, firstname=firstname,
                    email=email_raw, business_name=biz_name,
                    status_report_file=src_file, certificate_file=cert_name,
                    result="failed", reason=str(e),
                )
                failed += 1

            time.sleep(args.send_delay)

        if not session_limit_hit and batch_start + args.batch_size < len(df):
            log.info("Batch complete. Pausing %ds ...", args.batch_delay)
            time.sleep(args.batch_delay)

    tracker.save_excel()

    log.info("=" * 60)
    log.info(
        "Run complete — sent:%d  failed:%d  invalid:%d  "
        "skipped:%d  already-sent(skipped):%d",
        sent, failed, invalid, skipped, resumed,
    )
    if session_limit:
        log.info("Session limit: %d/%d used this run", session_sent, session_limit)
    log.info("Master log   -> %s", tracker.log_path)
    log.info("Master Excel -> %s", tracker.xlsx_path)

# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        description="Bulk CAC certificate email sender — Microsoft Graph API edition"
    )
    p.add_argument("--dataset",     required=True, help="Path to input Excel file")
    p.add_argument("--attachments", default="attachments", help="Dir containing PDFs")
    p.add_argument("--output",      default="output",
                   help="Dir for the master results log (default: output)")

    # ── Graph / Azure AD auth ─────────────────────────────────────────────────
    p.add_argument("--tenant-id",     required=True,
                   help="Azure AD tenant (directory) ID")
    p.add_argument("--client-id",     required=True,
                   help="App registration client (application) ID")
    p.add_argument("--sender",        required=True,
                   help="Email shown in the From: field (must match the authenticated account)")

    # ── Sending behaviour ─────────────────────────────────────────────────────
    p.add_argument("--batch-size",  type=int, default=10,
                   help="Emails per batch (default: 10)")
    p.add_argument("--batch-delay", type=int, default=60,
                   help="Seconds to pause between batches (default: 60)")
    p.add_argument("--send-delay",  type=int, default=3,
                   help="Seconds to pause between individual emails (default: 3)")
    p.add_argument("--session-limit", type=int, default=None, metavar="N",
                   help="Stop after sending N emails this run")
    p.add_argument("--allow-missing-attachments", action="store_true",
                   help="Send even if one or both PDF attachments are missing")
    return p.parse_args()


if __name__ == "__main__":
    run(parse_args())