# Recruiter Gmail Outreach Agent

A lightweight, automated Python agent to send individualized outreach emails with resume attachments to recruiters using Gmail API OAuth 2.0.

---

## Project Structure

```text
Recruiter Gmail Outreach Agent/
│
├── main.py                     # Main execution script
├── requirements.txt            # Python package dependencies
├── README.md                   # Setup and execution guide
├── .gitignore                  # Ignores credentials, tokens, .env, .venv, PDFs, private CSVs
├── .env.example                # Sample environment configuration
│
├── config/
│   ├── .gitkeep
│   ├── credentials.json        # OAuth 2.0 Client credentials (git-ignored)
│   └── token.json              # Generated user access/refresh token (git-ignored)
│
├── data/
│   ├── .gitkeep
│   ├── recruiters_sample.csv   # Sample template CSV with dummy addresses
│   └── recruiters.csv          # Your target recruiter email list (git-ignored)
│
├── templates/
│   └── template.html           # Predefined HTML email template
│
├── attachments/
│   ├── .gitkeep
│   └── resume.pdf              # Your resume attached to each email (git-ignored)
│
├── tests/
│   └── test_agent.py           # Non-destructive test suite
│
└── .venv/                      # Python virtual environment
```

---

## Features

- **CSV Ingestion**: Reads recipient emails from `data/recruiters.csv` (or `data/recruiters_sample.csv`).
- **Validation**: Verifies email address format using regex and skips invalid addresses.
- **Gmail API with OAuth 2.0**: Secure authentication using `config/credentials.json` without storing raw passwords.
- **HTML Email Template**: Sends formatted emails using `templates/template.html`.
- **Automatic Attachment**: Attaches your PDF resume from `attachments/` to each email.
- **Individual Delivery**: Emails are sent individually (direct `To`, no CC/BCC leak).
- **Error Handling**: Gracefully handles Gmail API HTTP errors and individual sending failures without crashing the batch.
- **Summary Report**: Prints real-time status and a final breakdown of sent, failed, and skipped addresses with corresponding reasons.

---

## 1. Setup & Installation

1. **Virtual Environment**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

2. **Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

---

## 2. Configuration & Preparation

1. **OAuth Credentials**:
   Download your OAuth 2.0 Desktop Client JSON from Google Cloud Console and place it at:
   ```text
   config/credentials.json
   ```

2. **Recruiter List**:
   Add your target recruiter emails to `data/recruiters.csv`:
   ```csv
   email
   recruiter1@company.com
   recruiter2@company.com
   ```

3. **Resume**:
   Place your PDF resume into the `attachments/` directory (e.g. `attachments/resume.pdf` or set `RESUME_FILE` in `.env`).

4. **Email Template**:
   Customize the email body text and signature placeholders in:
   ```text
   templates/template.html
   ```

---

## 3. Testing (Without Sending Emails)

Run the unit test suite to verify PDF validity, credentials format, regex validation, CSV parsing, and MIME email construction with attachment:

```bash
python tests/test_agent.py
```
or
```bash
python -m unittest discover tests
```

---

## 4. Execution

To start the agent and send emails, run from the project root:

```bash
python main.py
```

### What to Expect on First Run:
1. The script will open a browser window asking you to log into your Google Account.
2. Select your account and grant permission to send emails.
   *(If prompted with "Google hasn't verified this app", click **Advanced > Go to Recruiter Outreach Agent (unsafe)** — this is standard for personal test apps).*
3. Once approved, the script automatically saves `config/token.json` so you will not need to re-authenticate on future runs.
4. The agent processes each recruiter email sequentially, validates its format, attaches the resume, sends the email individually, and prints progress.
5. A summary report is displayed in the terminal upon completion.
