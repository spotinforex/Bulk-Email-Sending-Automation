"""
Bulk Email Sender with Certificate & Status Report Attachments
==============================================================
Reads from the dataset Excel file, validates emails, derives certificate
filenames, attaches both files, sends in batches, and logs all results.

Usage:
    python send_emails.py --dataset testing_email_dataset_060320.xlsx \
                          --attachments ./attachments \
                          --sender "sender@example.com" \
                          --smtp-host smtp.gmail.com \
                          --smtp-port 587 \
                          --smtp-user "sender@example.com" \
                          --smtp-pass "your_app_password" \
                          --batch-size 10 \
                          --batch-delay 60 \
                          --send-delay 3 \
                          --resume output/results_20260509_143022.txt output/results_20260510_090000.txt \
                          --session-limit 100
"""

import argparse
import asyncio
import logging
import re
import smtplib
import sys
import time
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from functools import wraps
from pathlib import Path

import dns.resolver
import pandas as pd

# ── Logging ──────────────────────────────────────────────────────────────────

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"email_run_{timestamp}.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ── Retry decorator ───────────────────────────────────────────────────────────

def retry(max_attempts: int = 3, delay: float = 2.0, backoff: float = 2.0, exceptions: tuple = (Exception,)):
    """
    Retry decorator for both sync and async functions.
    Args:
        max_attempts: Maximum number of attempts
        delay: Initial delay between retries in seconds
        backoff: Multiplier applied to delay after each retry (exponential backoff)
        exceptions: Tuple of exceptions to catch and retry on
    """
    def decorator(func):
        @wraps(func)
        async def async_wrapper(*args, **kwargs):
            current_delay = delay
            for attempt in range(1, max_attempts + 1):
                try:
                    return await func(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts:
                        logging.error(
                            "%s failed after %d attempts. Error: %s",
                            func.__name__, max_attempts, e
                        )
                        raise
                    logging.warning(
                        "%s attempt %d failed: %s. Retrying in %.2fs...",
                        func.__name__, attempt, e, current_delay
                    )
                    await asyncio.sleep(current_delay)
                    current_delay *= backoff

        @wraps(func)
        def sync_wrapper(*args, **kwargs):
            current_delay = delay
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts:
                        logging.error(
                            "%s failed after %d attempts. Error: %s",
                            func.__name__, max_attempts, e
                        )
                        raise
                    logging.warning(
                        "%s attempt %d failed: %s. Retrying in %.2fs...",
                        func.__name__, attempt, e, current_delay
                    )
                    time.sleep(current_delay)
                    current_delay *= backoff

        if asyncio.iscoroutinefunction(func):
            return async_wrapper
        return sync_wrapper
    return decorator

# ── Resume: load already-sent emails from one or more log files ───────────────

def load_sent_emails(log_files: list[str]) -> set[str]:
    """
    Read one or more previous result .txt log files and return the set of
    email addresses that were already successfully sent.
    Handles missing files gracefully and reports per-file counts.
    """
    already_sent: set[str] = set()

    if not log_files:
        return already_sent

    for log_file in log_files:
        path = Path(log_file)
        if not path.exists():
            log.warning(f"Resume file not found, skipping: {log_file}")
            continue

        count_before = len(already_sent)
        try:
            df = pd.read_csv(path, sep="\t", dtype=str).fillna("")
            sent_rows = df[df["result"].str.strip().str.lower() == "sent"]
            emails = {e.strip().lower() for e in sent_rows["email"] if e.strip()}
            already_sent.update(emails)
            count_added = len(already_sent) - count_before
            log.info(f"Resume — {path.name}: {count_added} sent emails loaded "
                     f"({len(sent_rows)} sent rows found)")
        except Exception as e:
            log.error(f"Could not read resume file {log_file}: {e}")

    log.info(f"Resume total: {len(already_sent)} unique already-sent emails across "
             f"{len(log_files)} file(s) — these will be skipped.")
    return already_sent

# ── Email validation ──────────────────────────────────────────────────────────

EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
)


def validate_email_format(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email.strip()))


def check_mx_record(domain: str) -> bool:
    """Return True if the domain has at least one MX record."""
    try:
        records = dns.resolver.resolve(domain, "MX", lifetime=10)
        return len(records) > 0
    except Exception:
        return False


def validate_email(email: str) -> tuple[bool, str]:
    """Full validation: format then MX check. Returns (ok, reason)."""
    email = email.strip().lower()
    if not validate_email_format(email):
        return False, "invalid_format"
    domain = email.split("@")[1]
    if not check_mx_record(domain):
        return False, "no_mx_record"
    return True, "ok"


# ── File name derivation ──────────────────────────────────────────────────────

def derive_certificate_name(status_report_name: str) -> str:
    """
    Replace 'Status Report CAC' (case-insensitive) with 'CERTIFICATE'
    in the filename, keeping the rest (business name + extension) intact.

    Example:
        "Status Report CAC - UGOPRINCE VOGUE FASHION.pdf"
        -> "CERTIFICATE - UGOPRINCE VOGUE FASHION.pdf"
    """
    return re.sub(
        r"Status\s+Report\s+CAC",
        "CERTIFICATE",
        status_report_name,
        flags=re.IGNORECASE,
    )


# ── Email composition ─────────────────────────────────────────────────────────

def build_email(
    sender: str,
    recipient: str,
    first_name: str,
    surname: str,
    business_name: str,
    status_report_path: Path,
    certificate_path: Path,
) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Cc"] = "Success.Isaac@ethnocentrique.com"
    msg["Subject"] = f"Your CAC Registration Documents - {business_name}"

    body = f"""\
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
    msg.attach(MIMEText(body, "plain"))

    for filepath in [status_report_path, certificate_path]:
        if filepath and filepath.exists():
            with open(filepath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{filepath.name}"',
            )
            msg.attach(part)
        else:
            log.warning(f"Attachment not found, skipping: {filepath}")

    return msg


# ── Result tracking ───────────────────────────────────────────────────────────

class ResultTracker:
    COLUMNS = [
        "timestamp", "row_index", "surname", "firstname",
        "email", "business_name", "status_report_file",
        "certificate_file", "result", "reason",
    ]

    def __init__(self, output_dir: Path):
        output_dir.mkdir(exist_ok=True)
        self.log_path  = output_dir / f"results_{timestamp}.txt"
        self.xlsx_path = output_dir / f"results_{timestamp}.xlsx"
        self.records: list[dict] = []

        # Write text header
        with open(self.log_path, "w") as f:
            f.write("\t".join(self.COLUMNS) + "\n")

    def record(self, **kwargs):
        kwargs.setdefault("timestamp", datetime.now().isoformat(timespec="seconds"))
        self.records.append(kwargs)

        # Append to text log immediately (crash-safe)
        with open(self.log_path, "a") as f:
            row = "\t".join(str(kwargs.get(c, "")) for c in self.COLUMNS)
            f.write(row + "\n")

    def save_excel(self):
        df = pd.DataFrame(self.records, columns=self.COLUMNS)
        df.to_excel(self.xlsx_path, index=False)
        log.info(f"Excel results saved -> {self.xlsx_path}")

    def summary(self):
        counts = {}
        for r in self.records:
            counts[r["result"]] = counts.get(r["result"], 0) + 1
        return counts


# ── SMTP connection (auto-switches based on port) ────────────────────────────

def connect_smtp(host: str, port: int, user: str, password: str) -> smtplib.SMTP:
    """
    Connect and authenticate to an SMTP server, choosing the right
    transport automatically based on port:

      465  → SMTP_SSL (implicit TLS from the start)
      587  → SMTP + STARTTLS (upgrade plain connection to TLS)
      2525 → SMTP + STARTTLS (common ISP-friendly fallback, same as 587)
      25   → SMTP plain (no TLS — last resort, most providers reject this)
      other→ attempt STARTTLS, fall back to plain if unsupported
    """
    log.info(f"Connecting to SMTP {host}:{port} ...")

    if port == 465:
        # Implicit SSL — wrap the socket before any SMTP handshake
        smtp = smtplib.SMTP_SSL(host, port, timeout=30)
        smtp.ehlo()
        log.info("Transport: SMTP_SSL (port 465)")
    elif port in (587, 2525):
        # Explicit TLS via STARTTLS
        smtp = smtplib.SMTP(host, port, timeout=30)
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        log.info(f"Transport: STARTTLS (port {port})")
    elif port == 25:
        # Plain — no encryption (most cloud providers block inbound on 25)
        smtp = smtplib.SMTP(host, port, timeout=30)
        smtp.ehlo()
        log.warning("Transport: plain SMTP (port 25) — no encryption")
    else:
        # Unknown port — try STARTTLS, fall back to plain
        smtp = smtplib.SMTP(host, port, timeout=30)
        smtp.ehlo()
        try:
            smtp.starttls()
            smtp.ehlo()
            log.info(f"Transport: STARTTLS (port {port})")
        except smtplib.SMTPNotSupportedError:
            log.warning(f"Transport: plain SMTP (port {port}) — STARTTLS not supported")

    smtp.login(user, password)
    log.info("SMTP authenticated.")
    return smtp


# ── Core send logic ───────────────────────────────────────────────────────────

@retry(max_attempts=3, delay=2.0, backoff=2.0, exceptions=(Exception,))
def send_one(smtp: smtplib.SMTP, sender: str, msg: MIMEMultipart, recipient: str):
    smtp.sendmail(sender, recipient, msg.as_string())


def run(args):
    attachments_dir = Path(args.attachments)
    output_dir      = Path(args.output)
    tracker         = ResultTracker(output_dir)

    # ── Load already-sent emails from previous log files ──────────────────────
    already_sent = load_sent_emails(args.resume or [])

    # ── Session limit ─────────────────────────────────────────────────────────
    session_limit = args.session_limit   # None means unlimited
    session_sent  = 0                    # successful sends THIS run only

    # Load dataset
    log.info(f"Loading dataset: {args.dataset}")
    df = pd.read_excel(args.dataset, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df = df.fillna("")
    log.info(f"  {len(df)} rows loaded.")

    if session_limit:
        log.info(f"  Session limit: {session_limit} sends — will stop after reaching it.")

    # ── SMTP connection (auto-switches mode based on port) ───────────────────
    smtp = connect_smtp(args.smtp_host, args.smtp_port, args.smtp_user, args.smtp_pass)

    sent = failed = invalid = skipped = resumed = 0
    session_limit_hit = False

    try:
        for batch_start in range(0, len(df), args.batch_size):

            if session_limit_hit:
                break

            batch = df.iloc[batch_start : batch_start + args.batch_size]
            log.info(
                f"--- Batch {batch_start // args.batch_size + 1}: "
                f"rows {batch_start + 1}-{batch_start + len(batch)} ---"
            )

            for idx, row in batch.iterrows():

                # ── Session limit check ───────────────────────────────────────
                if session_limit is not None and session_sent >= session_limit:
                    log.info(
                        f"Session limit of {session_limit} reached "
                        f"({session_sent} sent this run). Stopping."
                    )
                    session_limit_hit = True
                    break

                row_num   = idx + 2          # 1-indexed + header
                email_raw = row.get("EMAIL", "").strip()
                firstname = row.get("FIRSTNAME", "").strip()
                surname   = row.get("SURNAME", "").strip()
                biz_name  = row.get("Business Name", "").strip()
                src_file  = row.get("Source File", "").strip()

                if not email_raw:
                    log.warning(f"Row {row_num}: empty email — skipping")
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email="", business_name=biz_name,
                        status_report_file=src_file, certificate_file="",
                        result="skipped", reason="empty_email",
                    )
                    skipped += 1
                    continue

                # ── Resume check ──────────────────────────────────────────────
                if email_raw.lower() in already_sent:
                    log.info(f"Row {row_num}: {email_raw} already sent in a previous run — skipping")
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email=email_raw, business_name=biz_name,
                        status_report_file=src_file, certificate_file="",
                        result="resumed", reason="already_sent_previous_run",
                    )
                    resumed += 1
                    continue

                # Validate
                ok, reason = validate_email(email_raw)
                if not ok:
                    log.warning(f"Row {row_num}: {email_raw} -> INVALID ({reason})")
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email=email_raw, business_name=biz_name,
                        status_report_file=src_file, certificate_file="",
                        result="invalid", reason=reason,
                    )
                    invalid += 1
                    continue

                # Derive attachment paths
                status_path = attachments_dir / src_file if src_file else None
                cert_name   = derive_certificate_name(src_file) if src_file else ""
                cert_path   = attachments_dir / cert_name if cert_name else None

                missing = []
                if status_path and not status_path.exists():
                    missing.append(f"status_report: {src_file}")
                if cert_path and not cert_path.exists():
                    missing.append(f"certificate: {cert_name}")
                if missing and not args.allow_missing_attachments:
                    log.warning(
                        f"Row {row_num}: missing file(s): {missing} — skipping"
                    )
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email=email_raw, business_name=biz_name,
                        status_report_file=src_file, certificate_file=cert_name,
                        result="failed", reason=f"missing_attachments:{','.join(missing)}",
                    )
                    failed += 1
                    continue

                # Build & send
                msg = build_email(
                    sender=args.sender,
                    recipient=email_raw,
                    first_name=firstname,
                    surname=surname,
                    business_name=biz_name,
                    status_report_path=status_path,
                    certificate_path=cert_path,
                )

                try:
                    send_one(smtp, args.sender, msg, email_raw)
                    log.info(f"Row {row_num}: SENT -> {email_raw} ({biz_name})")
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email=email_raw, business_name=biz_name,
                        status_report_file=src_file, certificate_file=cert_name,
                        result="sent", reason="ok",
                    )
                    sent += 1
                    session_sent += 1
                except smtplib.SMTPRecipientsRefused as e:
                    log.error(f"Row {row_num}: recipient refused — {e}")
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email=email_raw, business_name=biz_name,
                        status_report_file=src_file, certificate_file=cert_name,
                        result="failed", reason=f"smtp_refused:{e}",
                    )
                    failed += 1
                except Exception as e:
                    log.error(f"Row {row_num}: send error — {e}")
                    tracker.record(
                        row_index=row_num, surname=surname, firstname=firstname,
                        email=email_raw, business_name=biz_name,
                        status_report_file=src_file, certificate_file=cert_name,
                        result="failed", reason=str(e),
                    )
                    failed += 1

                time.sleep(args.send_delay)

            # Inter-batch pause (skip after last batch or if limit hit)
            if not session_limit_hit and batch_start + args.batch_size < len(df):
                log.info(f"Batch complete. Pausing {args.batch_delay}s ...")
                time.sleep(args.batch_delay)

    finally:
        smtp.quit()
        tracker.save_excel()

    log.info("=" * 60)
    log.info(
        f"Run complete — "
        f"sent:{sent}  failed:{failed}  invalid:{invalid}  "
        f"skipped:{skipped}  resumed(already-sent):{resumed}"
    )
    if session_limit:
        log.info(f"Session limit: {session_sent}/{session_limit} used this run")
    log.info(f"Text log  -> {tracker.log_path}")
    log.info(f"Excel log -> {tracker.xlsx_path}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(description="Bulk CAC certificate email sender")
    p.add_argument("--dataset",           required=True,  help="Path to input Excel file")
    p.add_argument("--attachments",       default="attachments", help="Dir containing PDFs")
    p.add_argument("--output",            default="output",      help="Dir for result logs")
    p.add_argument("--sender",            required=True,  help="From address")
    p.add_argument("--smtp-host",         required=True)
    p.add_argument("--smtp-port",         type=int, default=587)
    p.add_argument("--smtp-user",         required=True)
    p.add_argument("--smtp-pass",         required=True)
    p.add_argument("--batch-size",        type=int, default=10,  help="Emails per batch")
    p.add_argument("--batch-delay",       type=int, default=60,  help="Seconds between batches")
    p.add_argument("--send-delay",        type=int, default=3,   help="Seconds between emails")
    p.add_argument("--resume",            nargs="+", metavar="LOG_FILE",
                   help="One or more previous result .txt files to load already-sent emails from")
    p.add_argument("--session-limit",     type=int, default=None, metavar="N",
                   help="Stop after sending N emails this run (does not count resumed/skipped)")
    p.add_argument("--allow-missing-attachments", action="store_true",
                   help="Send email even if one or both attachments are missing")
    return p.parse_args()


if __name__ == "__main__":
    run(parse_args())
