# CAC Certificate Bulk Email Sender

Reads a dataset of registrants, validates every email address, attaches the correct
**Status Report** and **Certificate** PDFs to each email, sends in rate-limited batches,
and logs every outcome to both a plain-text file and an Excel workbook.

---

## Files

| File | Purpose |
|------|---------|
| `send_emails.py` | Main sender — run this to actually send |
| `dry_run.py` | Validation preview — run this first |

---
## Attachment naming convention

The dataset's **Source File** column holds the status report filename, e.g.:

```
Status Report CAC - UGOPRINCE VOGUE FASHION.pdf
```

The script automatically derives the matching certificate name by replacing
`"Status Report CAC"` → `"CERTIFICATE"`:

```
CERTIFICATE - UGOPRINCE VOGUE FASHION.pdf
```

Both files must exist in the `--attachments` directory.

---

## Step 1 — Dry run (validation only, no emails sent)

```bash
python dry_run.py \
  --dataset  testing_email_dataset_060320.xlsx \
  --attachments ./attachments
```

Checks:
- Email format (regex)
- Domain MX record (DNS)
- Status report file exists
- Certificate file exists (derived name)

Fix any issues flagged before proceeding.

---

## Step 2 — Send

```bash
python send_emails.py \
  --dataset     testing_email_dataset_060320.xlsx \
  --attachments ./attachments \
  --sender      "noreply@yourorg.com" \
  --smtp-host   smtp.gmail.com \
  --smtp-port   587 \
  --smtp-user   "noreply@yourorg.com" \
  --smtp-pass   "your_app_password" \
  --batch-size  10 \
  --batch-delay 60 \
  --send-delay  3
```

### All flags

| Flag | Default | Description |
|------|---------|-------------|
| `--dataset` | *(required)* | Path to Excel dataset |
| `--attachments` | `attachments` | Directory containing PDF files |
| `--output` | `output` | Directory where logs are saved |
| `--sender` | *(required)* | From address shown to recipients |
| `--smtp-host` | *(required)* | SMTP server hostname |
| `--smtp-port` | `587` | SMTP port (587=STARTTLS, 465=SSL, 25=plain, 2525) |
| `--smtp-user` | *(required)* | SMTP login username |
| `--smtp-pass` | *(required)* | SMTP login password / app password |
| `--batch-size` | `10` | Emails per batch before a long pause |
| `--batch-delay` | `60` | Seconds to pause between batches |
| `--send-delay` | `3` | Seconds between individual emails |
| `--resume` | *(none)* | One or more previous result `.txt` files to skip already-sent emails |
| `--session-limit` | *(none)* | Stop after sending N emails this run |
| `--allow-missing-attachments` | off | Send email even if PDF(s) not found |

---

## Output / logs

After a run, two files appear in `./output/`:

```
output/
  results_20260509_143022.txt    ← tab-separated, written line-by-line (crash-safe)
  results_20260509_143022.xlsx   ← Excel workbook, written at end of run
logs/
  email_run_20260509_143022.log  ← full console log with timestamps
```

### Result values

| result | meaning |
|--------|---------|
| `sent` | Email delivered to SMTP server |
| `failed` | SMTP error or missing attachment |
| `invalid` | Bad email format or no MX record |
| `skipped` | Empty email cell |
| `resumed` | Skipped — already sent in a previous run |

---

## Resuming after a crash

The text log is written line-by-line so it survives a crash mid-run. On the next
run, pass the previous result file(s) via `--resume` and the script will
automatically skip any email already marked `sent`:

```bash
python send_emails.py \
  ... \
  --resume output/results_20260509_143022.txt
```

You can pass **multiple log files** at once to consolidate skips across several
previous runs:

```bash
python send_emails.py \
  ... \
  --resume output/results_20260509_143022.txt \
           output/results_20260510_090000.txt \
           output/results_20260511_080000.txt
```

At startup the script reports a per-file breakdown and a combined total:

```
Resume — results_20260509_143022.txt: 47 sent emails loaded
Resume — results_20260510_090000.txt: 31 sent emails loaded
Resume total: 78 unique already-sent emails across 3 file(s) — these will be skipped.
```

Resumed rows are logged as `result=resumed` in the new output file so you retain
a full audit trail without double-counting.

---

## Capping a run with `--session-limit`

To send only a fixed number of emails per run (e.g. to stay within daily SMTP
limits), use `--session-limit`:

```bash
python send_emails.py \
  ... \
  --session-limit 100
```

The counter only increments on a **successful send** — resumed, skipped, invalid,
and failed rows do not count toward the limit. Once the limit is hit the script
stops cleanly, flushes logs, and closes the SMTP connection. The final log line
shows how many were used:

```
Session limit: 100/100 used this run
```

Combine with `--resume` to run in daily chunks across multiple sessions without
ever double-sending:

```bash
# Day 1 — sends up to 100
python send_emails.py ... --session-limit 100

# Day 2 — skips yesterday's 100, sends the next batch
python send_emails.py ... --session-limit 100 --resume output/results_20260509_143022.txt
```
---

## Gmail-specific notes

1. Enable **2-Step Verification** on your Google account.
2. Generate an **App Password** (Google Account → Security → App Passwords).
3. Use the 16-character app password as `--smtp-pass`.
4. Gmail's free tier allows ~500 emails/day; G Suite/Workspace allows ~2,000/day.
   Tune `--batch-size`, `--batch-delay`, and `--session-limit` accordingly.


