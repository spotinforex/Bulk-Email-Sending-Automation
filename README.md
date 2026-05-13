# CAC Certificate Bulk Email Sender

Reads a dataset of registrants, validates every email address, attaches the correct
**Status Report** and **Certificate** PDFs to each email, sends in rate-limited batches,
and logs every outcome to both a plain-text file and an Excel workbook.

Two sender scripts are available depending on your email provider:

| Script | Provider | Auth method |
|--------|----------|-------------|
| `gmail_send_emails.py` | Gmail / Google Workspace | App Password (SMTP) |
| `microsoft_email_sender.py` | Outlook / Microsoft 365 | Microsoft Graph API (OAuth2) |

---

## Files

| File | Purpose |
|------|---------|
| `gmail_send_emails.py` | Sender for Gmail accounts |
| `microsoft_email_sender.py` | Sender for Microsoft/Outlook accounts |
| `dry_run.py` | Validation preview — run this first regardless of provider |

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

Run this before either sender script to catch issues early.

```bash
python dry_run.py \
  --dataset     testing_email_dataset_060320.xlsx \
  --attachments ./attachments
```

Checks performed:
- Email format (regex)
- Domain MX record (DNS)
- Status report file exists
- Certificate file exists (derived name)

Fix any issues flagged before proceeding.

---

## Step 2A — Send via Gmail (`gmail_send_emails.py`)

### Gmail setup (one time)

1. Enable **2-Step Verification** on the Google account.
2. Go to Google Account → Security → **App Passwords**.
3. Generate a password for "Mail" — you'll get a 16-character string.
4. Use that string as `--smtp-pass`.

> Gmail free tier: ~500 emails/day. Google Workspace: ~2,000/day.
> Tune `--batch-size`, `--batch-delay`, and `--session-limit` accordingly.

### Run command

```bash
python gmail_send_emails.py \
  --dataset     testing_email_dataset_060320.xlsx \
  --attachments ./attachments \
  --sender      "you@gmail.com" \
  --smtp-host   smtp.gmail.com \
  --smtp-port   587 \
  --smtp-user   "you@gmail.com" \
  --smtp-pass   "your_16char_app_password" \
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
| `--smtp-port` | `587` | SMTP port (587=STARTTLS, 465=SSL, 2525=fallback) |
| `--smtp-user` | *(required)* | SMTP login username |
| `--smtp-pass` | *(required)* | App password |
| `--batch-size` | `10` | Emails per batch before a long pause |
| `--batch-delay` | `60` | Seconds to pause between batches |
| `--send-delay` | `3` | Seconds between individual emails |
| `--resume` | *(none)* | One or more previous result `.txt` files to skip already-sent emails |
| `--session-limit` | *(none)* | Stop after sending N emails this run |
| `--allow-missing-attachments` | off | Send email even if PDF(s) not found |

---

## Step 2B — Send via Microsoft (`microsoft_email_sender.py`)

### Azure app setup (one time)

1. Go to [portal.azure.com](https://portal.azure.com) and sign in.
2. Search **App registrations** → New registration → give it a name → Register.
3. Note your **Application (client) ID** and **Directory (tenant) ID**.
4. Go to **API permissions** → Add a permission → Microsoft Graph →
   **Delegated** → `Mail.Send` → Add.
5. Go to **Authentication** → Add a redirect URI:
   ```
   https://login.microsoftonline.com/common/oauth2/nativeclient
   ```
   Then scroll to **Advanced settings** → set **Allow public client flows** to **Yes** → Save.

### First run — device code login

On first run the script prints a URL and a short code:

```
ACTION REQUIRED — Microsoft authentication
  1. Open:  https://microsoft.com/devicelogin
  2. Enter: ABCD1234
```

Open the URL in any browser, sign in with the sender's Microsoft account, and approve.
The token is cached to `~/.graph_token_cache.json` and reused on subsequent runs
— you won't be prompted again unless the token expires (refresh tokens last 90 days).

### Run command

```bash
python microsoft_email_sender.py \
  --dataset     testing_email_dataset_060320.xlsx \
  --attachments ./attachments \
  --sender      "you@outlook.com" \
  --tenant-id   "your-tenant-id" \
  --client-id   "your-client-id" \
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
| `--sender` | *(required)* | Microsoft sender email address |
| `--tenant-id` | *(required)* | Azure AD directory (tenant) ID |
| `--client-id` | *(required)* | Azure app registration client ID |
| `--batch-size` | `10` | Emails per batch before a long pause |
| `--batch-delay` | `60` | Seconds to pause between batches |
| `--send-delay` | `3` | Seconds between individual emails |
| `--resume` | *(none)* | One or more previous result `.txt` files to skip already-sent emails |
| `--session-limit` | *(none)* | Stop after sending N emails this run |
| `--allow-missing-attachments` | off | Send email even if PDF(s) not found |

---

## Output / logs

Both scripts write to the same output structure:

```
output/
  results_20260509_143022.txt    <- tab-separated, written line-by-line (crash-safe)
  results_20260509_143022.xlsx   <- Excel workbook, written at end of run
logs/
  email_run_20260509_143022.log  <- full console log with timestamps
```

### Result values

| result | meaning |
|--------|---------|
| `sent` | Email delivered successfully |
| `failed` | API/SMTP error or missing attachment |
| `invalid` | Bad email format or no MX record |
| `skipped` | Empty email cell |
| `resumed` | Skipped — already sent in a previous run |

---

## Resuming after a crash

The text log is written line-by-line so it survives a crash mid-run. On the next
run, pass the previous result file(s) via `--resume` and the script will
automatically skip any email already marked `sent`:

```bash
python microsoft_email_sender.py \
  ... \
  --resume output/results_20260509_143022.txt
```

You can pass **multiple log files** at once to consolidate skips across several
previous runs:

```bash
python microsoft_email_sender.py \
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

Resumed rows are logged as `result=resumed` so you retain a full audit trail
without double-counting.

---

## Capping a run with `--session-limit`

To send only a fixed number of emails per run, use `--session-limit`:

```bash
python microsoft_email_sender.py \
  ... \
  --session-limit 100
```

The counter only increments on a **successful send** — resumed, skipped, invalid,
and failed rows do not count. Combine with `--resume` to run in daily chunks:

```bash
# Day 1 — sends up to 100
python microsoft_email_sender.py ... --session-limit 100

# Day 2 — skips yesterday's 100, sends the next batch
python microsoft_email_sender.py ... --session-limit 100 --resume output/results_20260509_143022.txt
```
