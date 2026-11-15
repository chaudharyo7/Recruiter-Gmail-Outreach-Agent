import os
import re
import csv
import time
import html
import base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import List, Tuple, Dict, Any, Optional

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
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly'
]
DEFAULT_CSV = os.getenv('CSV_FILE', os.path.join('data', 'recruiters.csv'))
DEFAULT_TEMPLATE = os.getenv('TEMPLATE_FILE', os.path.join('templates', 'template.html'))
DEFAULT_RESUME = os.getenv(
    'RESUME_FILE',
    os.path.join('attachments', 'Yamandeep_Chaudhary_Resume.pdf')
    if os.path.exists(os.path.join('attachments', 'Yamandeep_Chaudhary_Resume.pdf'))
    else os.path.join('attachments', 'resume.pdf')
)
CREDENTIALS_FILE = os.getenv('CREDENTIALS_FILE', os.path.join('config', 'credentials.json'))
TOKEN_FILE = os.getenv('TOKEN_FILE', os.path.join('config', 'token.json'))
LOG_DIR = 'logs'
LOG_FILE = os.getenv('LOG_FILE', os.path.join(LOG_DIR, 'outreach.log'))
BOUNCE_CHECK_DELAY_SECONDS = int(os.getenv('BOUNCE_CHECK_DELAY_SECONDS', '5'))

EMAIL_REGEX = re.compile(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')
SUBJECT_REGEX = re.compile(r'<!--\s*Subject:\s*(.+?)\s*-->', re.IGNORECASE)
BOUNCE_SENDER_KEYWORDS = ['mailer-daemon', 'mail delivery subsystem', 'postmaster']


def sanitize_log_field(text: Any) -> str:
    """Sanitize field value: unescape HTML entities, strip and replace newlines/carriage returns with single spaces."""
    if text is None:
        return "None"
    s = html.unescape(str(text)).replace('\r', ' ').replace('\n', ' ').strip()
    return re.sub(r'\s+', ' ', s) if s else "None"


def format_email_log_line(
    timestamp: str,
    recipient: str,
    status: str,
    subject: str = "N/A",
    message_id: str = "N/A",
    submitted_time: str = "N/A",
    bounce_time: str = "N/A",
    error: str = "None"
) -> str:
    """Format an email record into strictly one physical line:
    [TIMESTAMP] | [RECIPIENT] | [STATUS] | Subject: [SUBJECT] | ID: [MESSAGE_ID] | Submitted: [TIME] | Bounce: [TIME/N/A] | Error: [ERROR/NONE]
    """
    s_ts = sanitize_log_field(timestamp)
    s_rec = sanitize_log_field(recipient)
    s_st = sanitize_log_field(status)
    s_sub = sanitize_log_field(subject) if subject and str(subject).strip() != "None" else "N/A"
    s_mid = sanitize_log_field(message_id) if message_id and str(message_id).strip() != "None" else "N/A"
    s_sub_time = sanitize_log_field(submitted_time) if submitted_time and str(submitted_time).strip() != "None" else "N/A"
    s_bt = sanitize_log_field(bounce_time) if bounce_time and str(bounce_time).strip() != "None" else "N/A"
    s_err = sanitize_log_field(error) if error and str(error).strip() != "None" else "None"

    return (
        f"{s_ts} | {s_rec} | {s_st} | Subject: {s_sub} | ID: {s_mid} | "
        f"Submitted: {s_sub_time} | Bounce: {s_bt} | Error: {s_err}"
    )


class OutreachLogger:
    """Persistent logger maintaining exactly one physical line per email in outreach.log."""

    def __init__(self, log_path: Optional[str] = None):
        self.log_path = log_path or LOG_FILE
        log_dir = os.path.dirname(self.log_path)
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)

    def log_event(self, level: str, message: str) -> None:
        """No-op: outreach.log must strictly contain only email records (no events, auth logs, or debug lines)."""
        pass

    def log_email_record(self, record: Dict[str, Any]) -> None:
        """Append exactly one physical log line for an email."""
        ts = record.get('timestamp') or datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        sub_time = record.get('submitted_time')
        if not sub_time or sub_time == "None":
            if record.get('status') in ('SUBMITTED', 'UNKNOWN'):
                sub_time = ts.split(' ')[1] if ' ' in ts else ts
            else:
                sub_time = "N/A"

        line = format_email_log_line(
            timestamp=ts,
            recipient=record['recipient'],
            status=record['status'],
            subject=record.get('subject') or "N/A",
            message_id=record.get('message_id') or "N/A",
            submitted_time=sub_time,
            bounce_time=record.get('bounce_time') or "N/A",
            error=record.get('error') or "None"
        )
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(line + "\n")

    def update_email_record(
        self,
        recipient: str,
        new_status: str,
        error: str = "None",
        bounce_time: Optional[str] = None,
        message_id: Optional[str] = None,
        subject: Optional[str] = None,
        remark: Optional[str] = None
    ) -> bool:
        """
        Update an existing email record line in outreach.log in-place.
        Preserves original submission timestamp, subject, and message_id.
        Ensures exactly one line per email across status changes.
        """
        if not os.path.exists(self.log_path):
            return False

        with open(self.log_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        target_email = recipient.lower().strip()
        updated = False

        # Search in reverse to update the most recent record for this recipient
        for idx in range(len(lines) - 1, -1, -1):
            line = lines[idx]
            parts = [p.strip() for p in line.split('|')]
            if len(parts) >= 3:
                line_recipient = parts[1].lower().strip()
                if line_recipient == target_email:
                    orig_timestamp = parts[0]

                    # Extract Subject
                    orig_subj = subject
                    if not orig_subj:
                        for p in parts[3:]:
                            if p.startswith('Subject:'):
                                orig_subj = re.sub(r'^Subject:\s*', '', p).strip()
                                break
                    if not orig_subj:
                        orig_subj = "N/A"

                    # Extract Message ID
                    orig_mid = message_id
                    if not orig_mid or orig_mid == "N/A":
                        for p in parts[3:]:
                            if p.startswith('ID:') or p.startswith('Message-ID:'):
                                extracted_id = re.sub(r'^(?:ID|Message-ID):\s*', '', p).strip()
                                if extracted_id and extracted_id != "N/A":
                                    orig_mid = extracted_id
                                    break
                    if not orig_mid:
                        orig_mid = "N/A"

                    # Extract Submitted time
                    orig_sub_time = "N/A"
                    for p in parts[3:]:
                        if p.startswith('Submitted:'):
                            orig_sub_time = re.sub(r'^Submitted:\s*', '', p).strip()
                            break
                    if orig_sub_time == "N/A" and ('SUBMITTED' in parts[2] or new_status in ('SUBMITTED', 'UNKNOWN', 'BOUNCED')):
                        orig_sub_time = orig_timestamp.split(' ')[1] if ' ' in orig_timestamp else orig_timestamp

                    # Determine bounce_time
                    cur_bounce = "N/A"
                    if bounce_time:
                        cur_bounce = bounce_time
                    elif new_status == 'BOUNCED':
                        cur_bounce = datetime.now().strftime('%H:%M:%S')
                    else:
                        for p in parts[3:]:
                            if p.startswith('Bounce:') or p.startswith('Bounce-Time:'):
                                cur_bounce = re.sub(r'^(?:Bounce|Bounce-Time):\s*', '', p).strip()
                                break

                    eff_error = error if error and error != "None" else (remark if remark and remark != "None" else "None")

                    updated_line = format_email_log_line(
                        timestamp=orig_timestamp,
                        recipient=parts[1],
                        status=new_status,
                        subject=orig_subj,
                        message_id=orig_mid,
                        submitted_time=orig_sub_time,
                        bounce_time=cur_bounce,
                        error=eff_error
                    )
                    lines[idx] = updated_line + "\n"
                    updated = True
                    break

        if updated:
            with open(self.log_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)

        return updated

    def get_unconfirmed_records(self) -> List[Dict[str, Any]]:
        """Retrieve previously logged records with SUBMITTED or UNKNOWN status for delayed bounce checks."""
        if not os.path.exists(self.log_path):
            return []

        unconfirmed = []
        with open(self.log_path, 'r', encoding='utf-8') as f:
            for line in f:
                parts = [p.strip() for p in line.split('|')]
                if len(parts) >= 3:
                    status = parts[2]
                    bounce_str = "N/A"
                    for p in parts[3:]:
                        if p.startswith('Bounce:') or p.startswith('Bounce-Time:'):
                            bounce_str = re.sub(r'^(?:Bounce|Bounce-Time):\s*', '', p).strip()
                            break

                    if status in ('SUBMITTED', 'UNKNOWN') and bounce_str == 'N/A':
                        mid = "N/A"
                        subj = "N/A"
                        sub_time = "N/A"
                        for p in parts[3:]:
                            if p.startswith('ID:') or p.startswith('Message-ID:'):
                                mid = re.sub(r'^(?:ID|Message-ID):\s*', '', p).strip()
                            elif p.startswith('Subject:'):
                                subj = re.sub(r'^Subject:\s*', '', p).strip()
                            elif p.startswith('Submitted:'):
                                sub_time = re.sub(r'^Submitted:\s*', '', p).strip()

                        unconfirmed.append({
                            'timestamp': parts[0],
                            'recipient': parts[1],
                            'status': status,
                            'subject': subj,
                            'message_id': mid,
                            'submitted_time': sub_time,
                            'bounce_time': 'N/A'
                        })
        return unconfirmed


def parse_template(template_path: str) -> Tuple[str, str]:
    """
    Read HTML template file, extract the email subject from the HTML comment,
    and return (subject, html_body).
    Raises ValueError if subject is missing or empty.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Email template not found: {template_path}")

    with open(template_path, 'r', encoding='utf-8') as f:
        content = f.read()

    match = SUBJECT_REGEX.search(content)
    if not match:
        raise ValueError(
            f"Email subject not found in '{template_path}'. "
            "Please define it using: <!-- Subject: Your Subject Here -->"
        )

    subject = match.group(1).strip()
    if not subject:
        raise ValueError(
            f"Email subject in '{template_path}' cannot be empty. "
            "Please provide a valid subject: <!-- Subject: Your Subject Here -->"
        )

    html_body = (content[:match.start()] + content[match.end():]).lstrip()
    return subject, html_body


def is_valid_email(email_str: str) -> bool:
    """Validate format of an email address."""
    return bool(EMAIL_REGEX.match(email_str.strip()))


def get_gmail_service():
    """Authenticate and return Gmail API service instance."""
    if None in (Request, Credentials, InstalledAppFlow, build):
        raise ImportError(
            "Missing Google client libraries. Run 'pip install -r requirements.txt' first."
        )

    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            import json
            with open(TOKEN_FILE, 'r', encoding='utf-8') as tf:
                token_data = json.load(tf)
            stored_scopes = set(token_data.get('scopes', []))
            # Verify stored token actually grants every required scope in SCOPES
            if set(SCOPES).issubset(stored_scopes):
                creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
            else:
                creds = None
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"'{CREDENTIALS_FILE}' not found. Please download OAuth client credentials "
                    "from Google Cloud Console and place them in config/credentials.json."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        token_dir = os.path.dirname(TOKEN_FILE)
        if token_dir:
            os.makedirs(token_dir, exist_ok=True)
        with open(TOKEN_FILE, 'w') as token_file:
            token_file.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


def read_recruiter_emails(csv_path: str) -> Tuple[List[str], List[str]]:
    """
    Read recruiter emails from CSV file.
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
            email_val = row[email_col_idx].strip()
            if not email_val:
                continue

            if is_valid_email(email_val):
                valid_emails.append(email_val)
            else:
                invalid_emails.append(email_val)

    return valid_emails, invalid_emails


def create_message(to_email: str, subject: str, html_content: str, resume_path: str) -> dict:
    """Create a MIME message with HTML body and PDF attachment, encoded for Gmail API."""
    message = MIMEMultipart()
    message['to'] = to_email
    message['subject'] = subject

    # Attach HTML body
    body_part = MIMEText(html_content, 'html', 'utf-8')
    message.attach(body_part)

    # Attach resume PDF
    with open(resume_path, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='pdf')
        filename = os.path.basename(resume_path)
        attachment.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(attachment)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw_message}


def extract_message_text(payload: dict) -> str:
    """Recursively extract text and decoded bodies from a Gmail message payload."""
    text_parts = []

    # Check headers
    headers = payload.get('headers', [])
    for h in headers:
        name = h.get('name', '').lower()
        if name in ('from', 'to', 'subject', 'message-id', 'in-reply-to', 'references'):
            text_parts.append(h.get('value', ''))

    # Check body
    body_data = payload.get('body', {}).get('data')
    if body_data:
        try:
            decoded = base64.urlsafe_b64decode(body_data.encode('utf-8')).decode('utf-8', errors='ignore')
            text_parts.append(decoded)
        except Exception:
            pass

    # Check nested parts recursively
    parts = payload.get('parts', [])
    for part in parts:
        text_parts.append(extract_message_text(part))

    return " ".join(text_parts)


def parse_bounce_notification(msg_full: dict) -> Optional[Dict[str, Any]]:
    """Inspect a Gmail message and parse it if it is a genuine delivery failure notification."""
    snippet = msg_full.get('snippet', '')
    payload = msg_full.get('payload', {})
    text_content = extract_message_text(payload)
    full_text = f"{snippet} {text_content}".strip()
    full_text_lower = full_text.lower()

    headers = {h.get('name', '').lower(): h.get('value', '') for h in payload.get('headers', [])}
    sender = headers.get('from', '').lower()
    subject = headers.get('subject', '').lower()

    is_bounce_sender = any(k in sender for k in BOUNCE_SENDER_KEYWORDS)
    is_bounce_subject = any(
        k in subject for k in [
            'delivery status notification',
            'undelivered mail',
            'mail delivery failed',
            'failure notice',
            'returned to sender'
        ]
    )

    if not (is_bounce_sender or is_bounce_subject):
        return None

    # Extract SMTP error code e.g. 550 5.1.1 or 554
    error_code = "None"
    code_match = re.search(r'\b(5\d{2}(?:\s+\d\.\d\.\d)?)\b', full_text)
    if code_match:
        error_code = f"SMTP {code_match.group(1).strip()}"

    # Extract bounce reason
    reason = "Recipient mailbox rejected message"
    reason_patterns = [
        r'(address couldn\'t be found[^\r\n<>]*)',
        r'((?:recipient|the email account|user|mailbox)[^\r\n<>]*(?:does not exist|not found|unavailable|rejected)[^\r\n<>]*)',
        r'(?:because\s+|diagnostic-code:\s*|reason:\s*)([^\r\n<>]+)',
        r'(?:550|554)\s+(?:[\d.]+\s+)?([^\r\n<>]+)',
    ]
    for pattern in reason_patterns:
        match = re.search(pattern, full_text_lower)
        if match:
            extracted = match.group(1).strip()
            extracted = re.split(r'\.\s+', extracted)[0].strip().rstrip('.')
            if extracted:
                reason = extracted
                break
    else:
        if snippet:
            reason = snippet.split('.')[0].strip()

    return {
        'id': msg_full.get('id'),
        'sender': sender,
        'subject': subject,
        'date': headers.get('date', ''),
        'snippet': snippet,
        'full_text_lower': full_text_lower,
        'reason': reason,
        'error_code': error_code
    }


def detect_bounces(
    service,
    submitted_records: List[Dict[str, Any]],
    logger: Optional[OutreachLogger] = None,
    delay_seconds: int = 0,
    cached_bounces: Optional[List[Dict[str, Any]]] = None
) -> List[Dict[str, Any]]:
    """
    Check Gmail inbox for delivery failure notifications from mailer-daemon,
    mail delivery subsystem, or postmaster.
    Matches notifications to submitted records, updates status to BOUNCED or UNKNOWN,
    and updates outreach.log in-place without duplicate records.
    """
    if not submitted_records:
        return submitted_records

    if cached_bounces is not None and len(cached_bounces) > 0:
        bounces = cached_bounces
    else:
        if delay_seconds > 0:
            time.sleep(delay_seconds)

        query = 'from:(mailer-daemon OR "Mail Delivery Subsystem" OR postmaster)'
        try:
            resp = service.users().messages().list(userId='me', q=query, maxResults=50).execute()
            messages = resp.get('messages', [])
        except HttpError as e:
            if e.resp.status == 403:
                error_detail = "Insufficient OAuth permissions (gmail.readonly scope required)"
            else:
                error_detail = f"Gmail API HTTP {e.resp.status}: {e.reason}"
            for rec in submitted_records:
                if rec.get('status') == 'SUBMITTED':
                    rec['status'] = 'UNKNOWN'
                    rec['remark'] = f"Bounce check failed ({error_detail}); delivery unconfirmed"
                    rec['error'] = "None"
                    if logger:
                        logger.update_email_record(
                            recipient=rec['recipient'],
                            new_status='UNKNOWN',
                            remark=rec['remark'],
                            error=rec['error'],
                            bounce_time='N/A',
                            message_id=rec.get('message_id'),
                            subject=rec.get('subject')
                        )
            return submitted_records
        except Exception as e:
            error_detail = str(e)
            for rec in submitted_records:
                if rec.get('status') == 'SUBMITTED':
                    rec['status'] = 'UNKNOWN'
                    rec['remark'] = f"Bounce check unavailable ({error_detail}); delivery unconfirmed"
                    rec['error'] = "None"
                    if logger:
                        logger.update_email_record(
                            recipient=rec['recipient'],
                            new_status='UNKNOWN',
                            remark=rec['remark'],
                            error=rec['error'],
                            bounce_time='N/A',
                            message_id=rec.get('message_id'),
                            subject=rec.get('subject')
                        )
            return submitted_records

        # Retrieve and parse bounce notifications
        bounces = []
        for msg_meta in messages:
            msg_id = msg_meta.get('id')
            try:
                msg_full = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
                parsed = parse_bounce_notification(msg_full)
                if parsed:
                    bounces.append(parsed)
            except Exception:
                pass

        if cached_bounces is not None:
            cached_bounces.extend(bounces)

    bounce_time = datetime.now().strftime('%H:%M:%S')

    # Match bounces against submitted records
    for rec in submitted_records:
        rec_email = rec['recipient'].lower().strip()
        rec_mid = str(rec.get('message_id', '')).strip()
        matched_bounce = None

        for b in bounces:
            # Match original recipient address or Message-ID in bounce notification
            if rec_email in b['full_text_lower'] or (rec_mid != 'N/A' and rec_mid != '' and rec_mid.lower() in b['full_text_lower']):
                matched_bounce = b
                break

        if matched_bounce:
            rec['status'] = 'BOUNCED'
            rec['bounce_time'] = bounce_time
            rec['error'] = matched_bounce['error_code'] if matched_bounce['error_code'] != "None" else matched_bounce['reason']
            rec['remark'] = matched_bounce['reason']
            if logger:
                logger.update_email_record(
                    recipient=rec['recipient'],
                    new_status='BOUNCED',
                    remark=rec['remark'],
                    error=rec['error'],
                    bounce_time=rec['bounce_time'],
                    message_id=rec.get('message_id'),
                    subject=rec.get('subject')
                )
        else:
            # If no bounce found, ensure status is UNKNOWN (delivery remains unconfirmed)
            if rec['status'] == 'SUBMITTED':
                rec['status'] = 'UNKNOWN'
                rec['remark'] = "No bounce detected; delivery unconfirmed"
                rec['error'] = "None"
                rec['bounce_time'] = "N/A"
                if logger:
                    logger.update_email_record(
                        recipient=rec['recipient'],
                        new_status='UNKNOWN',
                        remark=rec['remark'],
                        error=rec['error'],
                        bounce_time=rec['bounce_time'],
                        message_id=rec.get('message_id'),
                        subject=rec.get('subject')
                    )

    return submitted_records


def main():
    logger = OutreachLogger()

    print("=== Recruiter Outreach Agent ===\n")

    # 1. Check prerequisite files & parse template
    if not os.path.exists(DEFAULT_RESUME):
        err_msg = f"Resume attachment '{DEFAULT_RESUME}' not found."
        print(f"Error: {err_msg}")
        return

    try:
        subject, html_template = parse_template(DEFAULT_TEMPLATE)
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return

    # 2. Read and validate emails from CSV
    try:
        valid_emails, invalid_emails = read_recruiter_emails(DEFAULT_CSV)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return

    total_recipients = len(invalid_emails) + len(valid_emails)
    if total_recipients == 0:
        print("No email addresses found in CSV.")
        return

    # 3. Authenticate with Gmail API
    service = None
    try:
        service = get_gmail_service()
    except Exception as e:
        print(f"Authentication failed: {e}")
        return

    curr_index = 1
    all_records: List[Dict[str, Any]] = []
    submitted_records: List[Dict[str, Any]] = []

    # Process invalid emails
    for inv in invalid_emails:
        print(f"[{curr_index}/{total_recipients}] {inv} - INVALID")
        curr_index += 1
        now = datetime.now()
        rec = {
            'timestamp': now.strftime('%Y-%m-%d %H:%M:%S'),
            'recipient': inv,
            'status': 'INVALID',
            'subject': subject,
            'message_id': 'N/A',
            'submitted_time': 'N/A',
            'bounce_time': 'N/A',
            'remark': 'Invalid email format; skipped',
            'error': 'None'
        }
        all_records.append(rec)
        logger.log_email_record(rec)

    # Submit valid emails
    for recipient in valid_emails:
        now = datetime.now()
        rec = {
            'timestamp': now.strftime('%Y-%m-%d %H:%M:%S'),
            'recipient': recipient,
            'subject': subject,
            'status': None,
            'message_id': 'N/A',
            'submitted_time': now.strftime('%H:%M:%S'),
            'bounce_time': 'N/A',
            'remark': 'None',
            'error': 'None'
        }
        try:
            message_body = create_message(
                to_email=recipient,
                subject=subject,
                html_content=html_template,
                resume_path=DEFAULT_RESUME
            )
            resp = service.users().messages().send(userId='me', body=message_body).execute()
            msg_id = resp.get('id', 'N/A')
            rec['message_id'] = msg_id
            rec['status'] = 'SUBMITTED'
            rec['remark'] = 'Gmail API accepted email; delivery unconfirmed'
            rec['error'] = 'None'
            submitted_records.append(rec)
            all_records.append(rec)
            logger.log_email_record(rec)
            print(f"[{curr_index}/{total_recipients}] {recipient} - SUBMITTED")
        except HttpError as error:
            err_text = f"HTTP {error.resp.status}: {error.reason}"
            rec['status'] = 'FAILED_TO_SUBMIT'
            rec['submitted_time'] = 'N/A'
            rec['remark'] = 'Gmail API rejected sending request'
            rec['error'] = err_text
            all_records.append(rec)
            logger.log_email_record(rec)
            print(f"[{curr_index}/{total_recipients}] {recipient} - FAILED_TO_SUBMIT")
        except Exception as error:
            err_text = str(error)
            rec['status'] = 'FAILED_TO_SUBMIT'
            rec['submitted_time'] = 'N/A'
            rec['remark'] = 'Sending failed'
            rec['error'] = err_text
            all_records.append(rec)
            logger.log_email_record(rec)
            print(f"[{curr_index}/{total_recipients}] {recipient} - FAILED_TO_SUBMIT")
        curr_index += 1

    # 4. Check for Bounces (Current run + Delayed bounces from previous unconfirmed emails)
    if service:
        cached_bounces: List[Dict[str, Any]] = []
        # Check current batch
        if submitted_records:
            detect_bounces(
                service,
                submitted_records,
                logger=logger,
                delay_seconds=BOUNCE_CHECK_DELAY_SECONDS,
                cached_bounces=cached_bounces
            )

        # Check previously unconfirmed emails from logs (delayed bounces)
        try:
            prev_unconfirmed = logger.get_unconfirmed_records()
            curr_emails = {r['recipient'].lower() for r in submitted_records}
            prev_to_check = [r for r in prev_unconfirmed if r['recipient'].lower() not in curr_emails]
            if prev_to_check:
                detect_bounces(
                    service,
                    prev_to_check,
                    logger=logger,
                    delay_seconds=0,
                    cached_bounces=cached_bounces if cached_bounces else None
                )
        except Exception:
            pass

    # 5. Calculate Summary Counts
    submitted_count = len(submitted_records)
    bounced_count = sum(1 for r in submitted_records if r.get('status') == 'BOUNCED')
    failed_count = sum(1 for r in all_records if r.get('status') == 'FAILED_TO_SUBMIT')
    invalid_count = len(invalid_emails)
    unknown_count = sum(1 for r in submitted_records if r.get('status') == 'UNKNOWN')

    # 6. Display Minimal Terminal Output
    print("\nSummary:")
    print(f"Submitted: {submitted_count}")
    print(f"Bounced: {bounced_count}")
    print(f"Failed: {failed_count}")
    print(f"Invalid: {invalid_count}")
    print(f"Unknown: {unknown_count}\n")
    print(f"Logs: {LOG_FILE}")


if __name__ == '__main__':
    main()
