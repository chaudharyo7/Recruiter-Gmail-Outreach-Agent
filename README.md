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
│   └── recruiters.csv          # Recruiter email list (pre-populated with dummy sample rows)
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
└── .venv/                      # Python virtual environment (git-ignored)
```

---

## Features

- **CSV Ingestion**: Reads recipient emails from `data/recruiters.csv`.
- **Validation**: Verifies email address format using regex and skips invalid addresses.
- **Gmail API with OAuth 2.0**: Secure authentication using `config/credentials.json` without storing raw passwords.
- **HTML Email Template**: Sends formatted emails using `templates/template.html`.
- **Automatic Attachment**: Attaches your PDF resume from `attachments/` to each email.
- **Individual Delivery**: Emails are sent individually (direct `To`, no CC/BCC leak).
- **Error Handling**: Gracefully handles Gmail API HTTP errors and individual sending failures without crashing the batch.
- **Summary Report**: Prints real-time status and a final breakdown of sent, failed, and skipped addresses with corresponding reasons.

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
   - Add scope: `https://www.googleapis.com/auth/gmail.send`.
   - Add your Gmail address under **Test Users**.
4. In **Credentials**, create an **OAuth Client ID** for **Desktop App**.
5. Download the JSON, rename it to `credentials.json`, and place it in the `config/` directory:
   ```text
   config/credentials.json
   ```
   *(See `config/README.md` and `config/credentials.json.example` for details).*

### 5. Prepare Recruiter CSV
Open `data/recruiters.csv` and replace the dummy example emails with your target recruiter addresses:
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

### 7. Customize Email Template
Edit `templates/template.html` to adjust the body text and update the signature placeholders with your name and links.

### 8. Run Verification Tests (Optional)
Run the unit test suite to verify configuration, CSV parsing, email generation, and PDF attachment without sending emails:
```bash
python tests/test_agent.py
```
or
```bash
python -m unittest discover tests
```

### 9. Run the Agent
Execute the outreach script:
```bash
python main.py
```

- On first run, a browser tab will open asking you to authorize your Google Account with the `gmail.send` scope.
- Once authorized, `config/token.json` is generated automatically for future runs.
- The agent will send each email individually and print a summary report with sent, failed, and skipped counts.
