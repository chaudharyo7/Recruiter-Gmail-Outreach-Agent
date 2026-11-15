# Recruiter Gmail Outreach Agent

A lightweight, automated Python agent to send individualized outreach emails with resume attachments to recruiters using the Gmail API with OAuth 2.0.

---

## Project Structure

```text
Recruiter Gmail Outreach Agent/
│
├── main.py                     # Main execution script
├── requirements.txt            # Python package dependencies
├── README.md                   # Setup and execution guide
├── .gitignore                  # Ignores credentials, tokens, .env, .venv, and PDFs
├── .env.example                # Sample environment configuration
│
├── config/
│   ├── README.md               # Detailed OAuth credential setup instructions
│   ├── credentials.json.example# Sample client credentials structure
│   ├── credentials.json        # Your actual OAuth client credentials (git-ignored)
│   └── token.json              # Auto-generated OAuth access token (git-ignored)
│
├── data/
│   └── recruiter_contacts.csv  # Recruiter email list (pre-populated with dummy sample rows)
│
├── templates/
│   └── template.html           # Predefined HTML email template with contact placeholders
│
├── attachments/
│   ├── README.md               # Instructions for attaching your resume
│   └── resume.pdf              # Your resume PDF to attach (git-ignored)
│
├── tests/
│   └── test_agent.py           # Non-destructive test suite
│
├── logs/                       # Persistent outreach execution logs (git-ignored)
│   └── outreach.log
│
└── .venv/                      # Python virtual environment (git-ignored)
```

---

## Features

- **CSV Ingestion**: Reads recipient emails from `data/recruiter_contacts.csv`.
- **Validation**: Verifies email address format using regex and skips invalid addresses.
- **Gmail API with OAuth 2.0**: Secure authentication using `config/credentials.json` without storing raw passwords.
- **HTML Email Template**: Sends formatted emails using `templates/template.html`.
- **Automatic Attachment**: Attaches your PDF resume from `attachments/` to each email.
- **Individual Delivery**: Emails are sent individually (direct `To`, no CC/BCC leak).
- **Error Handling**: Gracefully handles Gmail API HTTP errors and individual sending failures without crashing the batch.
- **Accurate Status Reporting**: Distinguishes `SUBMITTED`, `BOUNCED`, `UNKNOWN`, and `FAILED_TO_SUBMIT`. Never assumes API acceptance is delivery.
- **Bounce Detection**: Detects and matches delivery failure notifications from Mail Delivery Subsystem / Mailer-Daemon / postmaster.
- **Persistent Logging**: Automatically records timestamps, recipient addresses, subjects, message IDs, and status to `logs/outreach.log`.
- **Summary Report**: Prints real-time status and a categorized breakdown of attempted, submitted, bounced, failed to submit, invalid, and unknown counts.

---

## Quickstart Guide

Follow these steps to configure and run the outreach agent:

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
5. Download the JSON, rename it to `credentials.json`, and place it in the `config/` directory:
   ```text
   config/credentials.json
   ```
   *(See `config/README.md` and `config/credentials.json.example` for details).*

### 5. Prepare Recruiter CSV
Open `data/recruiter_contacts.csv` and replace the dummy example emails with your target recruiter addresses:
```csv
email
recruiter1@company.com
recruiter2@company.com
```

### 6. Add Personal Resume
Copy your PDF resume into the `attachments/` directory and name it `resume.pdf`:
```text
attachments/resume.pdf
```
*(If you prefer a custom filename, place it in `attachments/` and set `RESUME_FILE` in `.env`).*

### 7. Centralized Configuration (`.env`)

All configurable settings are managed centrally through a single `.env` file. Copy `.env.example` to `.env`:
```bash
cp .env.example .env
```

| Section | Setting | Default | Description |
| :--- | :--- | :--- | :--- |
| **Email** | `SENDER_EMAIL` | *(empty)* | Optional sender email address to include in the `From` header. |
| | `EMAIL_SUBJECT` | *(empty)* | Custom email subject. If left empty, extracts subject from the HTML template comment (`<!-- Subject: ... -->`). |
| | `EMAIL_TEMPLATE_FILE`| `templates/template.html` | Path to HTML email template. |
| | `RESUME_FILE` | `attachments/resume.pdf` | Path to your resume PDF attachment. |
| | `RECRUITER_CSV_FILE` | `data/recruiter_contacts.csv` | Path to input CSV containing recruiter emails. |
| **Workers** | `NUM_WORKERS` | `3` | Number of parallel worker threads in `ThreadPoolExecutor`. |
| | `MAX_RETRIES` | `3` | Max retries with exponential backoff on transient Gmail API errors (429, 500, 502, 503, 504). |
| | `RETRY_BACKOFF_SECONDS`| `2` | Initial retry backoff delay in seconds for transient API errors. |
| | `GMAIL_API_REQUEST_DELAY`| `0` | Delay in seconds between API requests per worker to pace rate limits. |
| **Excel Directory**| `EXCEL_DIRECTORY_ENABLED`| `true` | When `true`, automatically registers all successfully `SUBMITTED` emails into the public directory. |
| | `EXCEL_DIRECTORY_FILE`| `data/recruiter_email_backup.xlsx` | Path to the reusable public recruiter Excel directory. |
| **Bounce Detection**| `BOUNCE_CHECK_ENABLED` | `true` | When `true`, scans inbox for bounce notifications and updates log records. |
| | `BOUNCE_CHECK_DELAY_SECONDS`| `5` | Delay in seconds before checking for immediate bounces. |
| **Logging** | `LOGGING_ENABLED` | `true` | When `true`, writes persistent one-line status records to `outreach.log`. |
| | `LOG_FILE` | `logs/outreach.log` | Path to persistent outreach log file. |
| **Dashboard** | `DASHBOARD_ENABLED` | `true` | When `true`, displays a real-time live terminal dashboard with worker states and progress. |
| | `DASHBOARD_REFRESH_RATE`| `4` | Dashboard refresh rate in Hz (screen renders per second). |
| **OAuth** | `GMAIL_CREDENTIALS_FILE`| `credentials.json` | Path to Google OAuth client credentials JSON (falls back to `config/credentials.json`). |
| | `GMAIL_TOKEN_FILE` | `token.json` | Path to saved OAuth token JSON (falls back to `config/token.json`). |
| **Execution** | `DRY_RUN` | `false` | When `true`, simulates end-to-end execution, worker states, and dashboard progress **without sending real emails**. |

### 8. Customize Email Template
`templates/template.html` manages your email body and default subject:
- Define your email subject in the HTML comment at the very top:
  ```html
  <!-- Subject: Application for DevOps / Cloud Engineer Opportunities -->
  ```
- Edit the HTML body and update signature placeholders with your contact info and links.
Changing the subject comment automatically changes the subject of outgoing emails without modifying any code (unless overridden by `EMAIL_SUBJECT` in `.env`).

### 9. Run Verification Tests (Optional)
Run the unit test suite to verify configuration loading, CSV parsing, 3-worker concurrency, Excel directory, and DRY_RUN mode without sending emails:
```bash
python -m unittest discover tests
```

### 10. Run the Agent
Execute the outreach script:
```bash
python main.py
```

- On first run, a browser tab will open asking you to authorize your Google Account with the required `gmail.send` and `gmail.readonly` scopes.
- Once authorized, `config/token.json` is generated automatically for future runs.
- The agent sends each email individually, checks for bounce notifications, and prints a compact summary:

```text
=== Recruiter Outreach Agent ===

[1/2] recruiter@example.com - SUBMITTED
[2/2] hr@example.com - INVALID

Summary:
Submitted: 1
Bounced: 0
Failed: 0
Invalid: 1
Unknown: 0

Logs: logs/outreach.log
```
