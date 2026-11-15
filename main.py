import os
import re
import csv
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import List, Tuple

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    Request = None
    Credentials = None
    InstalledAppFlow = None
    build = None
    HttpError = Exception

# Configuration
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
DEFAULT_CSV = os.getenv(
    'CSV_FILE',
    os.path.join('data', 'recruiters.csv')
    if os.path.exists(os.path.join('data', 'recruiters.csv'))
    else os.path.join('data', 'recruiters_sample.csv')
)
DEFAULT_TEMPLATE = os.getenv('TEMPLATE_FILE', os.path.join('templates', 'template.html'))
DEFAULT_RESUME = os.getenv(
    'RESUME_FILE',
    os.path.join('attachments', 'Yamandeep_Chaudhary_Resume.pdf')
    if os.path.exists(os.path.join('attachments', 'Yamandeep_Chaudhary_Resume.pdf'))
    else os.path.join('attachments', 'resume.pdf')
)
DEFAULT_SUBJECT = os.getenv('EMAIL_SUBJECT', 'Application & Resume Submission')
CREDENTIALS_FILE = os.getenv('CREDENTIALS_FILE', os.path.join('config', 'credentials.json'))
TOKEN_FILE = os.getenv('TOKEN_FILE', os.path.join('config', 'token.json'))

EMAIL_REGEX = re.compile(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')


def is_valid_email(email: str) -> bool:
    """Validate format of an email address."""
    return bool(EMAIL_REGEX.match(email.strip()))


def get_gmail_service():
    """Authenticate and return the Gmail API service instance."""
    if None in (Request, Credentials, InstalledAppFlow, build):
        raise ImportError(
            "Missing Google client libraries. Run 'pip install -r requirements.txt' first."
        )

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"'{CREDENTIALS_FILE}' not found. Please download OAuth client credentials "
                    "from Google Cloud Console and place them in the project root."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for subsequent runs
        token_dir = os.path.dirname(TOKEN_FILE)
        if token_dir:
            os.makedirs(token_dir, exist_ok=True)
        with open(TOKEN_FILE, 'w') as token_file:
            token_file.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


def read_recruiter_emails(csv_path: str) -> Tuple[List[str], List[str]]:
    """
    Read recruiter emails from a CSV file.
    Validates format and returns (valid_emails, invalid_emails).
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    valid_emails = []
    invalid_emails = []

    with open(csv_path, mode='r', encoding='utf-8') as f:
        reader = csv.reader(f)
        rows = list(reader)

        if not rows:
            return valid_emails, invalid_emails

        # Detect if header exists (e.g. contains 'email' or 'mail')
        first_row = [col.strip().lower() for col in rows[0]]
        email_col_idx = 0
        start_row_idx = 0

        if any('email' in col or 'mail' in col for col in first_row):
            start_row_idx = 1
            for idx, col in enumerate(first_row):
                if 'email' in col or 'mail' in col:
                    email_col_idx = idx
                    break

        for row in rows[start_row_idx:]:
            if not row or len(row) <= email_col_idx:
                continue
            email = row[email_col_idx].strip()
            if not email:
                continue

            if is_valid_email(email):
                valid_emails.append(email)
            else:
                invalid_emails.append(email)

    return valid_emails, invalid_emails


def create_message(to_email: str, subject: str, html_content: str, resume_path: str) -> dict:
    """Create a MIME message with HTML body and PDF attachment, encoded for Gmail API."""
    message = MIMEMultipart()
    message['to'] = to_email
    message['subject'] = subject

    # Attach HTML body
    body_part = MIMEText(html_content, 'html', 'utf-8')
    message.attach(body_part)

    # Attach resume.pdf
    with open(resume_path, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='pdf')
        filename = os.path.basename(resume_path)
        attachment.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(attachment)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw_message}


def main():
    print("=== Recruiter Gmail Outreach Agent ===")

    # 1. Check prerequisite files
    if not os.path.exists(DEFAULT_TEMPLATE):
        print(f"Error: Email template '{DEFAULT_TEMPLATE}' not found.")
        return

    if not os.path.exists(DEFAULT_RESUME):
        print(f"Error: Resume attachment '{DEFAULT_RESUME}' not found.")
        return

    with open(DEFAULT_TEMPLATE, 'r', encoding='utf-8') as f:
        html_template = f.read()

    # 2. Read and validate emails from CSV
    try:
        valid_emails, invalid_emails = read_recruiter_emails(DEFAULT_CSV)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return

    print(f"Total valid recipients found: {len(valid_emails)}")
    if invalid_emails:
        print(f"Skipped {len(invalid_emails)} invalid email address(es):")
        for inv in invalid_emails:
            print(f"  - [SKIPPED] {inv}")

    if not valid_emails:
        print("No valid email addresses to send. Exiting.")
        return

    # 3. Authenticate with Gmail API
    print("\nAuthenticating with Gmail API...")
    try:
        service = get_gmail_service()
    except Exception as e:
        print(f"Authentication failed: {e}")
        return

    # 4. Send emails individually
    sent_emails = []
    failed_emails = []

    print("\nSending emails...")
    for idx, recipient in enumerate(valid_emails, start=1):
        print(f"[{idx}/{len(valid_emails)}] Sending to: {recipient}...", end=" ", flush=True)
        try:
            message_body = create_message(
                to_email=recipient,
                subject=DEFAULT_SUBJECT,
                html_content=html_template,
                resume_path=DEFAULT_RESUME
            )
            service.users().messages().send(userId='me', body=message_body).execute()
            sent_emails.append(recipient)
            print("SUCCESS")
        except HttpError as error:
            failed_emails.append((recipient, f"Gmail API HTTP {error.resp.status}: {error.reason}"))
            print(f"FAILED (API error: {error.reason})")
        except Exception as error:
            failed_emails.append((recipient, str(error)))
            print(f"FAILED (Error: {error})")

    # 5. Display Summary
    print("\n" + "=" * 45)
    print("               SUMMARY REPORT                ")
    print("=" * 45)
    print(f"Total Attempted : {len(valid_emails)}")
    print(f"Total Sent      : {len(sent_emails)}")
    print(f"Total Failed    : {len(failed_emails)}")
    print(f"Total Invalid   : {len(invalid_emails)}")
    print("-" * 45)

    if sent_emails:
        print("\nSent Addresses:")
        for addr in sent_emails:
            print(f"  ✓ {addr}")

    if failed_emails:
        print("\nFailed Addresses:")
        for addr, err in failed_emails:
            print(f"  ✗ {addr} (Reason: {err})")

    if invalid_emails:
        print("\nSkipped Invalid Addresses:")
        for addr in invalid_emails:
            print(f"  - {addr}")
    print("=" * 45)


if __name__ == '__main__':
    main()
