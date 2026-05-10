"""
Dry Run — Preview & Validate Without Sending
=============================================
Run this before send_emails.py to catch issues early.

Usage:
    python dry_run.py --dataset testing_email_dataset_060320.xlsx \
                      --attachments ./attachments
"""

import argparse
import re
import sys
from pathlib import Path

import dns.resolver
import pandas as pd

EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")


def validate_email_format(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email.strip()))


def check_mx_record(domain: str) -> bool:
    try:
        records = dns.resolver.resolve(domain, "MX", lifetime=10)
        return len(records) > 0
    except Exception:
        return False


def derive_certificate_name(status_report_name: str) -> str:
    return re.sub(
        r"Status\s+Report\s+CAC",
        "CERTIFICATE",
        status_report_name,
        flags=re.IGNORECASE,
    )


def run(args):
    attachments_dir = Path(args.attachments)
    df = pd.read_excel(args.dataset, dtype=str).fillna("")
    df.columns = [c.strip() for c in df.columns]

    print(f"\n{'='*70}")
    print(f"  DRY RUN — {args.dataset}  ({len(df)} rows)")
    print(f"  Attachments dir: {attachments_dir.resolve()}")
    print(f"{'='*70}\n")

    issues = []
    rows_ok = 0

    for idx, row in df.iterrows():
        row_num  = idx + 2
        email    = row.get("EMAIL", "").strip()
        name     = f"{row.get('FIRSTNAME','')} {row.get('SURNAME','')}".strip()
        biz      = row.get("Business Name", "").strip()
        src_file = row.get("Source File", "").strip()

        row_issues = []

        # Email checks
        if not email:
            row_issues.append("empty email")
        elif not validate_email_format(email):
            row_issues.append(f"bad format: {email}")
        else:
            domain = email.split("@")[1]
            if not check_mx_record(domain):
                row_issues.append(f"no MX record for {domain}")

        # Attachment checks
        if src_file:
            sp = attachments_dir / src_file
            if not sp.exists():
                row_issues.append(f"status report missing: {src_file}")
            cert_name = derive_certificate_name(src_file)
            cp = attachments_dir / cert_name
            if not cp.exists():
                row_issues.append(f"certificate missing: {cert_name}")
        else:
            row_issues.append("no Source File value")

        if row_issues:
            issues.append((row_num, name, biz, row_issues))
            for issue in row_issues:
                print(f"  [ROW {row_num:>4}] {name:<30} | {issue}")
        else:
            cert_name = derive_certificate_name(src_file)
            print(
                f"  [ROW {row_num:>4}] ✓  {name:<28}  "
                f"{email}  →  {cert_name}"
            )
            rows_ok += 1

    print(f"\n{'='*70}")
    print(f"  SUMMARY: {rows_ok} OK  |  {len(issues)} with issues")
    print(f"{'='*70}\n")

    if issues:
        sys.exit(1)


def parse_args():
    p = argparse.ArgumentParser(description="Dry-run validation before sending")
    p.add_argument("--dataset",     required=True)
    p.add_argument("--attachments", default="attachments")
    return p.parse_args()


if __name__ == "__main__":
    run(parse_args())
