# Recruiter Gmail Outreach Agent

A lightweight, high-performance automated Python agent to send individualized outreach emails with resume attachments to recruiters using the Gmail API with OAuth 2.0.

---

## Project Structure

```text
Recruiter Gmail Outreach Agent/
│
├── main.py                     # Main execution script & concurrent worker engine
├── requirements.txt            # Python package dependencies
├── README.md                   # Setup and execution guide
├── .gitignore                  # Ignores credentials, tokens, .env, .venv, PDFs, data, templates
├── .env.example                # Sample environment configuration
│
├── config/
│   ├── README.md               # Detailed OAuth credential setup instructions
│   ├── credentials.json.example# Sample client credentials structure
│   ├── credentials.json        # Your actual OAuth client credentials (git-ignored)
│   └── token.json              # Auto-generated OAuth access token (git-ignored)
│
├── data/
│   ├── recruiter_contacts.csv  # Temporary input queue for recruiter emails
│   └── recruiter_email_backup.xlsx # Reusable public recruiter directory
│
├── templates/
│   └── template.html           # HTML email body template
│
├── attachments/
│   ├── README.md               # Instructions for attaching your resume
│   └── resume.pdf              # Your resume PDF to attach (git-ignored)
│
├── tests/
│   └── test_agent.py           # Comprehensive 30-test suite (non-destructive)
│
├── logs/                       # Persistent outreach execution logs (git-ignored)
│   └── outreach.log
│
└── .venv/                      # Python virtual environment (git-ignored)
```

---

## Features

- **Concurrent Multi-Worker Engine**: Sends emails concurrently across 3 thread-safe workers with independent Gmail API transports.
- **Configurable Email Subject**: Subject is configured exclusively via `.env` (`EMAIL_SUBJECT`) as the single source of truth.
- **Two User-Facing Delivery Statuses**:
  - `SENT`: Confirmed acceptance by the Gmail API.
  - `UNDELIVERED (reason)`: Confirmed delivery failure with extracted bounce reason (HTML entities unescaped).
- **Duplicate Prevention with Configurable Cooldown**:
  - Pre-flight history check against `logs/outreach.log`.
  - Configurable cooldown window (`DUPLICATE_COOLDOWN_DAYS`, default: 10 days).
  - Skips addresses sent within the cooldown window (calculated from original submission timestamp).
  - Confirmed `UNDELIVERED` addresses are permanently excluded and flagged for manual review (`UNDELIVERED REVIEW`) regardless of age.
  - `FAILED_TO_SUBMIT` addresses remain immediately eligible for retry.
- **CSV Auto-Clear & Crash Recovery**:
  - `data/recruiter_contacts.csv` functions as a temporary input queue.
  - Never clears at startup.
  - Automatically clears processed recipient rows upon completion while preserving the header.
  - Retains unprocessed or uncertain rows if interrupted or crashed.
- **Excel Public Directory**:
  - Automatically appends successfully submitted emails to `data/recruiter_email_backup.xlsx` with case-insensitive deduplication.
  - Automatically removes confirmed permanent invalid addresses (mailbox not found, user unknown, address not found).
  - Retains addresses with temporary failures (mailbox full, greylisting, delayed).
- **Minimal, Clean Terminal Dashboard**:
  - Clean startup with non-fatal dependency warnings suppressed.
  - Compact single-line startup header.
  - Real-time live dashboard with progress bar, active worker panels, and an aligned 8-counter grid.
  - Restructured final execution summary with conditional detail sections and a compact footer.
- **Persistent Logging**: Writes durable one-line-per-email records with immediate disk flushing (`fsync`) to prevent duplicate sending across crashes.
- **DRY_RUN Mode**: Full simulation of workflow, dashboard, and CSV auto-clearing without sending real emails.

---

## Quickstart Guide

### 1. Clone Repository
```bash
git clone https://github.com/chaudharyo7/Recruiter-Gmail-Outreach-Agent.git
cd Recruiter-Gmail-Outreach-Agent
```

### 2. Create Virtual Environment
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Requirements
```bash
pip install -r requirements.txt
```

### 4. Configure Gmail API and OAuth Credentials
1. Go to the [Google Cloud Console](https://console.cloud.google.com/) and create a project.
2. Enable the **Gmail API** in **APIs & Services > Library**.
3. In **OAuth consent screen**:
   - Set User Type to **External**.
   - Add scopes: `https://www.googleapis.com/auth/gmail.send` and `https://www.googleapis.com/auth/gmail.readonly`.
   - Add your Gmail address under **Test Users**.
4. In **Credentials**, create an **OAuth Client ID** for **Desktop App**.
5. Download the JSON, rename it to `credentials.json`, and place it in the project root or `config/`:
   ```text
   credentials.json
   ```
   *(See `config/README.md` and `config/credentials.json.example` for details).*

### 5. Prepare Recruiter CSV Queue
Add target recruiter emails to `data/recruiter_contacts.csv`:
```csv
email
recruiter1@company.com
recruiter2@company.com
```

### 6. Add Resume Attachment
Copy your resume PDF into `attachments/` and name it `resume.pdf` (or set `RESUME_FILE` in `.env`):
```text
attachments/resume.pdf
```

### 7. Centralized Configuration (`.env`)

Copy `.env.example` to `.env` and configure your settings:
```bash
cp .env.example .env
```

| Section | Setting | Default | Description |
| :--- | :--- | :--- | :--- |
| **Email** | `SENDER_EMAIL` | *(empty)* | Optional sender email address in `From` header. |
| | `EMAIL_SUBJECT` | *(empty)* | **Required**. Single source of truth for email subject line. |
| | `EMAIL_TEMPLATE_FILE`| `templates/template.html` | Path to HTML email body template. |
| | `RESUME_FILE` | `attachments/resume.pdf` | Path to resume PDF attachment. |
| | `RECRUITER_CSV_FILE` | `data/recruiter_contacts.csv` | Path to temporary input CSV queue. |
| **Workers** | `NUM_WORKERS` | `3` | Parallel worker threads in `ThreadPoolExecutor`. |
| | `MAX_RETRIES` | `3` | Retries with exponential backoff on transient API errors (429, 500, 502, 503, 504). |
| | `RETRY_BACKOFF_SECONDS`| `2` | Initial retry backoff delay in seconds. |
| | `GMAIL_API_REQUEST_DELAY`| `0` | Rate-limiting delay in seconds between requests per worker. |
| **Excel Directory**| `EXCEL_DIRECTORY_ENABLED`| `true` | When `true`, manages `data/recruiter_email_backup.xlsx`. |
| | `EXCEL_DIRECTORY_FILE`| `data/recruiter_email_backup.xlsx` | Path to reusable recruiter backup directory. |
| **Duplicate Prevention**| `DUPLICATE_CHECK_ENABLED`| `true` | When `true`, enforces time-based sending cooldown checks against log history. |
| | `DUPLICATE_COOLDOWN_DAYS`| `10` | Duration of sending cooldown window in days (positive integer). |
| **Bounce Detection**| `BOUNCE_CHECK_ENABLED` | `true` | When `true`, scans inbox for delivery failure notifications and updates records. |
| | `BOUNCE_CHECK_DELAY_SECONDS`| `5` | Delay in seconds before checking for immediate bounces. |
| **Logging** | `LOGGING_ENABLED` | `true` | When `true`, writes durable status records to `outreach.log`. |
| | `LOG_FILE` | `logs/outreach.log` | Path to persistent outreach log file. |
| **Dashboard** | `DASHBOARD_ENABLED` | `true` | When `true`, renders live terminal dashboard and formatted summary. |
| | `DASHBOARD_REFRESH_RATE`| `4` | Dashboard refresh rate in Hz. |
| **OAuth** | `GMAIL_CREDENTIALS_FILE`| `credentials.json` | Path to client credentials JSON. |
| | `GMAIL_TOKEN_FILE` | `token.json` | Path to saved OAuth token JSON. |
| **Execution** | `DRY_RUN` | `false` | When `true`, simulates execution and auto-clearing **without sending real emails**. |

### 8. Run Verification Tests
Verify all functionality, concurrency, Excel deduplication, cooldown timers, bounce handling, and CSV auto-clearing:
```bash
python -m unittest discover tests
```

### 9. Run Outreach Agent
Execute the agent:
```bash
python main.py
```

- On first run, a browser tab opens to authorize your Google Account with `gmail.send` and `gmail.readonly` scopes.
- `token.json` is saved locally for future runs.
- The agent displays a live progress dashboard and outputs the final execution summary:

```text
Recruiter Gmail Outreach Agent  •  Workers: 3  •  Duplicate Check: 10d cooldown  •  Mode: LIVE

Final Execution Summary
――――――――――――――――――――――――――――――――――――――――――――――――――――
TOTAL INPUT:          10
NEW RECIPIENTS:       7
SENT:                 7
UNDELIVERED:          0
COOLDOWN SKIPPED (10-DAY COOLDOWN): 2
UNDELIVERED REVIEW:   1
INVALID:              0
FAILED TO SUBMIT:     0

Cooldown Skipped (10-Day Cooldown)
  • recruiter@example.com: Previous Submission: 2026-11-15 10:00:00 | Eligible Again: 2026-11-25 10:00:00

Undelivered Review (Prior Bounces Flagged for Review)
  • invalid.mailbox@domain.com: Status: UNDELIVERED (550 5.1.1 Address not found)

――――――――――――――――――――――――――――――――――――――――――――――――――――
Elapsed: 00:04  •  Rate: 105.0 msg/min  •  Log: logs/outreach.log  •  Directory: data/recruiter_email_backup.xlsx
――――――――――――――――――――――――――――――――――――――――――――――――――――
```
