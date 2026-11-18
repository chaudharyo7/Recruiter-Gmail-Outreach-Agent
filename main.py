import os
import re
import csv
import sys
import time
import html
import queue
import base64
import threading
import concurrent.futures
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import List, Tuple, Dict, Any, Optional, Set

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    from dotenv import load_dotenv, dotenv_values
    load_dotenv()
except ImportError:
    dotenv_values = lambda *args, **kwargs: {}

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

# Centralized Configuration Loader and Parser Helpers
def parse_bool(val: Any, default: bool = False) -> bool:
    """Parse a boolean value from environment string representation."""
    if val is None:
        return default
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    if s in ('true', '1', 'yes', 'y', 'on', 'enabled'):
        return True
    if s in ('false', '0', 'no', 'n', 'off', 'disabled', ''):
        return False
    return default


def parse_int(val: Any, default: int, min_val: Optional[int] = None) -> int:
    """Parse an integer value with optional lower bound constraint."""
    try:
        res = int(str(val).strip())
        if min_val is not None and res < min_val:
            return default
        return res
    except (ValueError, TypeError):
        return default


def parse_float(val: Any, default: float, min_val: Optional[float] = None) -> float:
    """Parse a float value with optional lower bound constraint."""
    try:
        res = float(str(val).strip())
        if min_val is not None and res < min_val:
            return default
        return res
    except (ValueError, TypeError):
        return default


class Config:
    """Centralized configuration loader and validator for Recruiter Outreach Agent.
    Manages all settings from .env, falling back to sensible defaults.
    """

    def __init__(self, env_path: Optional[str] = None):
        env_dict = {}
        if env_path and os.path.exists(env_path):
            env_dict = dotenv_values(env_path)
        elif os.path.exists('.env'):
            env_dict = dotenv_values('.env')

        def get_val(key: str) -> Optional[str]:
            if env_path is not None:
                if key in env_dict and env_dict[key] is not None:
                    return env_dict[key]
                return os.getenv(key)
            if key in os.environ and os.environ[key] is not None:
                return os.environ[key]
            if key in env_dict and env_dict[key] is not None:
                return env_dict[key]
            return None

        # EMAIL CONFIGURATION
        self.sender_email: Optional[str] = get_val('SENDER_EMAIL') or None
        self.email_subject: Optional[str] = get_val('EMAIL_SUBJECT') or None

        raw_template = get_val('EMAIL_TEMPLATE_FILE') or get_val('TEMPLATE_FILE')
        self.email_template_file: str = self._resolve_path(
            raw_template,
            defaults=['templates/template.html']
        )

        raw_resume = get_val('RESUME_FILE')
        self.resume_file: str = self._resolve_path(
            raw_resume,
            defaults=[
                'attachments/Yamandeep_Chaudhary_Resume.pdf',
                'attachments/resume.pdf'
            ]
        )

        raw_csv = get_val('RECRUITER_CSV_FILE') or get_val('CSV_FILE')
        self.recruiter_csv_file: str = self._resolve_path(
            raw_csv,
            defaults=['data/recruiter_contacts.csv']
        )

        # WORKERS & PERFORMANCE
        self.num_workers: int = parse_int(get_val('NUM_WORKERS'), default=3, min_val=1)
        self.max_retries: int = parse_int(get_val('MAX_RETRIES'), default=3, min_val=0)
        self.retry_backoff_seconds: float = parse_float(get_val('RETRY_BACKOFF_SECONDS'), default=2.0, min_val=0.0)
        self.gmail_api_request_delay: float = parse_float(get_val('GMAIL_API_REQUEST_DELAY'), default=0.0, min_val=0.0)

        # EXCEL DIRECTORY
        self.excel_directory_enabled: bool = parse_bool(get_val('EXCEL_DIRECTORY_ENABLED'), default=True)
        self.excel_directory_file: str = get_val('EXCEL_DIRECTORY_FILE') or os.path.join('data', 'recruiter_email_backup.xlsx')

        # DUPLICATE PREVENTION
        self.duplicate_check_enabled: bool = parse_bool(get_val('DUPLICATE_CHECK_ENABLED'), default=True)
        raw_cooldown = get_val('DUPLICATE_COOLDOWN_DAYS')
        if raw_cooldown is not None and str(raw_cooldown).strip() != '':
            try:
                self.duplicate_cooldown_days: int = int(str(raw_cooldown).strip())
            except (ValueError, TypeError):
                self.duplicate_cooldown_days = 0
        else:
            self.duplicate_cooldown_days = 10

        # BOUNCE DETECTION
        self.bounce_check_enabled: bool = parse_bool(get_val('BOUNCE_CHECK_ENABLED'), default=True)
        self.bounce_check_delay_seconds: int = parse_int(get_val('BOUNCE_CHECK_DELAY_SECONDS'), default=5, min_val=0)

        # LOGGING
        self.logging_enabled: bool = parse_bool(get_val('LOGGING_ENABLED'), default=True)
        self.log_file: str = get_val('LOG_FILE') or os.path.join('logs', 'outreach.log')

        # DASHBOARD
        self.dashboard_enabled: bool = parse_bool(get_val('DASHBOARD_ENABLED'), default=True)
        self.dashboard_refresh_rate: float = parse_float(get_val('DASHBOARD_REFRESH_RATE'), default=4.0, min_val=0.1)

        # GMAIL API / OAUTH
        raw_creds = get_val('GMAIL_CREDENTIALS_FILE') or get_val('CREDENTIALS_FILE')
        self.gmail_credentials_file: str = self._resolve_path(
            raw_creds,
            defaults=['credentials.json', 'config/credentials.json']
        )

        raw_token = get_val('GMAIL_TOKEN_FILE') or get_val('TOKEN_FILE')
        self.gmail_token_file: str = self._resolve_path(
            raw_token,
            defaults=['token.json', 'config/token.json']
        )

        # EXECUTION
        self.dry_run: bool = parse_bool(get_val('DRY_RUN'), default=False)

    @staticmethod
    def _resolve_path(path: Optional[str], defaults: List[str]) -> str:
        if path:
            if os.path.exists(path):
                return path
            # Check candidate defaults if specified path does not exist
            for candidate in defaults:
                if os.path.exists(candidate):
                    return candidate
            return path
        for candidate in defaults:
            if os.path.exists(candidate):
                return candidate
        return defaults[0]

    def validate(self) -> None:
        """Validate paths and configuration parameters."""
        if self.num_workers < 1:
            raise ValueError(f"NUM_WORKERS must be an integer >= 1, got {self.num_workers}")
        if self.max_retries < 0:
            raise ValueError(f"MAX_RETRIES must be an integer >= 0, got {self.max_retries}")
        if self.retry_backoff_seconds < 0:
            raise ValueError(f"RETRY_BACKOFF_SECONDS must be >= 0, got {self.retry_backoff_seconds}")
        if self.gmail_api_request_delay < 0:
            raise ValueError(f"GMAIL_API_REQUEST_DELAY must be >= 0, got {self.gmail_api_request_delay}")
        if self.dashboard_refresh_rate <= 0:
            raise ValueError(f"DASHBOARD_REFRESH_RATE must be > 0, got {self.dashboard_refresh_rate}")
        if self.bounce_check_delay_seconds < 0:
            raise ValueError(f"BOUNCE_CHECK_DELAY_SECONDS must be >= 0, got {self.bounce_check_delay_seconds}")
        if self.duplicate_cooldown_days < 1:
            raise ValueError(f"DUPLICATE_COOLDOWN_DAYS must be a positive integer >= 1, got {self.duplicate_cooldown_days}")


_CONFIG = Config()


def get_config(env_path: Optional[str] = None) -> Config:
    """Return the global configuration instance, or a fresh one if env_path is provided."""
    global _CONFIG
    if env_path is not None:
        return Config(env_path=env_path)
    return _CONFIG


# Configuration constants maintained for module-level compatibility
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly'
]
DEFAULT_CSV = _CONFIG.recruiter_csv_file
DEFAULT_TEMPLATE = _CONFIG.email_template_file
DEFAULT_RESUME = _CONFIG.resume_file
CREDENTIALS_FILE = _CONFIG.gmail_credentials_file
TOKEN_FILE = _CONFIG.gmail_token_file
LOG_DIR = os.path.dirname(_CONFIG.log_file) or 'logs'
LOG_FILE = _CONFIG.log_file
BOUNCE_CHECK_DELAY_SECONDS = _CONFIG.bounce_check_delay_seconds
DEFAULT_EXCEL_DIRECTORY = _CONFIG.excel_directory_file
NUM_WORKERS = _CONFIG.num_workers
DUPLICATE_CHECK_ENABLED = _CONFIG.duplicate_check_enabled
DUPLICATE_COOLDOWN_DAYS = _CONFIG.duplicate_cooldown_days

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
        self._lock = threading.Lock()
        log_dir = os.path.dirname(self.log_path)
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)

    def log_event(self, level: str, message: str) -> None:
        """No-op: outreach.log must strictly contain only email records (no events, auth logs, or debug lines)."""
        pass

    def log_email_record(self, record: Dict[str, Any]) -> None:
        """Append exactly one physical log line for an email."""
        self.log_email_records([record])

    def log_email_records(self, records: List[Dict[str, Any]]) -> None:
        """Batch append multiple email records in a single disk I/O write."""
        if not records:
            return
        lines = []
        for record in records:
            ts = record.get('timestamp') or datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            sub_time = record.get('submitted_time')
            if not sub_time or sub_time == "None":
                st = record.get('status')
                if st in ('SENT', 'SUBMITTED', 'UNKNOWN') or (st and str(st).startswith('UNDELIVERED')):
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
            lines.append(line + "\n")

        with self._lock:
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.writelines(lines)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except OSError:
                    pass

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
        """Update an existing email record line in outreach.log in-place."""
        eff_err = error if error and error != "None" else (remark if remark and remark != "None" else "None")
        updates = {
            recipient.lower().strip(): {
                'new_status': new_status,
                'error': eff_err,
                'bounce_time': bounce_time,
                'message_id': message_id,
                'subject': subject
            }
        }
        return self.batch_update_email_records(updates) > 0

    def batch_update_email_records(self, updates_by_email: Dict[str, Dict[str, Any]]) -> int:
        """
        Batch update multiple email records in a single read-write file pass.
        Preserves original submission timestamp, subject, and message_id.
        Key is lowercase recipient email address.
        """
        if not os.path.exists(self.log_path) or not updates_by_email:
            return 0

        with self._lock:
            with open(self.log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            pending_updates = {k.lower().strip(): v for k, v in updates_by_email.items()}
            updated_count = 0
            now_time = datetime.now().strftime('%H:%M:%S')

            # Scan in reverse to update the most recent record for each recipient
            for idx in range(len(lines) - 1, -1, -1):
                line = lines[idx]
                parts = [p.strip() for p in line.split('|')]
                if len(parts) >= 3:
                    line_recipient = parts[1].lower().strip()
                    if line_recipient in pending_updates:
                        upd = pending_updates[line_recipient]
                        orig_timestamp = parts[0]

                        # Extract Subject
                        orig_subj = upd.get('subject')
                        if not orig_subj or orig_subj == "N/A":
                            for p in parts[3:]:
                                if p.startswith('Subject:'):
                                    orig_subj = re.sub(r'^Subject:\s*', '', p).strip()
                                    break
                        if not orig_subj:
                            orig_subj = "N/A"

                        # Extract Message ID
                        orig_mid = upd.get('message_id')
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
                        new_status = upd.get('new_status', parts[2])
                        is_sub = (
                            'SENT' in parts[2]
                            or 'SUBMITTED' in parts[2]
                            or new_status in ('SENT', 'SUBMITTED', 'BOUNCED')
                            or str(new_status).startswith('UNDELIVERED')
                        )
                        if orig_sub_time == "N/A" and is_sub:
                            orig_sub_time = orig_timestamp.split(' ')[1] if ' ' in orig_timestamp else orig_timestamp

                        # Determine bounce_time
                        cur_bounce = "N/A"
                        if upd.get('bounce_time'):
                            cur_bounce = upd['bounce_time']
                        elif str(new_status).startswith('UNDELIVERED') or new_status == 'BOUNCED':
                            cur_bounce = now_time
                        else:
                            for p in parts[3:]:
                                if p.startswith('Bounce:') or p.startswith('Bounce-Time:'):
                                    cur_bounce = re.sub(r'^(?:Bounce|Bounce-Time):\s*', '', p).strip()
                                    break

                        eff_error = upd.get('error', 'None')
                        if not eff_error or eff_error == "None":
                            rem = upd.get('remark')
                            if rem and rem != "None":
                                eff_error = rem
                            else:
                                eff_error = "None"

                        lines[idx] = format_email_log_line(
                            timestamp=orig_timestamp,
                            recipient=parts[1],
                            status=new_status,
                            subject=orig_subj,
                            message_id=orig_mid,
                            submitted_time=orig_sub_time,
                            bounce_time=cur_bounce,
                            error=eff_error
                        ) + "\n"
                        updated_count += 1
                        del pending_updates[line_recipient]
                        if not pending_updates:
                            break

            if updated_count > 0:
                with open(self.log_path, 'w', encoding='utf-8') as f:
                    f.writelines(lines)
                    f.flush()
                    try:
                        os.fsync(f.fileno())
                    except OSError:
                        pass

            return updated_count

    def get_unconfirmed_records(self) -> List[Dict[str, Any]]:
        """Retrieve previously logged records with SENT status (and legacy SUBMITTED/UNKNOWN) for delayed bounce checks."""
        if not os.path.exists(self.log_path):
            return []

        unconfirmed = []
        with self._lock:
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

                        is_active_sent = (
                            status in ('SENT', 'SUBMITTED', 'UNKNOWN')
                            and not status.startswith('UNDELIVERED')
                            and status != 'BOUNCED'
                            and bounce_str == 'N/A'
                        )
                        if is_active_sent:
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
                                'bounce_time': 'N/A',
                                'remark': 'Sent; awaiting delayed bounce check',
                                'error': 'None'
                            })
        return unconfirmed


def parse_log_timestamp(ts_str: str) -> Optional[datetime]:
    """Parse timestamp string from outreach.log into datetime object."""
    if not ts_str:
        return None
    cleaned = ts_str.strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(cleaned)
    except Exception:
        return None


def parse_log_line(line: str) -> Optional[Dict[str, Any]]:
    """Parse a single physical outreach.log line into structured fields."""
    parts = [p.strip() for p in line.split('|')]
    if len(parts) < 3:
        return None

    orig_timestamp = parts[0]
    recipient = parts[1]
    status = parts[2]

    subject = "N/A"
    message_id = "N/A"
    submitted_time = "N/A"
    bounce_time = "N/A"
    error = "None"

    for p in parts[3:]:
        if p.startswith('Subject:'):
            subject = re.sub(r'^Subject:\s*', '', p).strip()
        elif p.startswith('ID:') or p.startswith('Message-ID:'):
            message_id = re.sub(r'^(?:ID|Message-ID):\s*', '', p).strip()
        elif p.startswith('Submitted:'):
            submitted_time = re.sub(r'^Submitted:\s*', '', p).strip()
        elif p.startswith('Bounce:') or p.startswith('Bounce-Time:'):
            bounce_time = re.sub(r'^(?:Bounce|Bounce-Time):\s*', '', p).strip()
        elif p.startswith('Error:'):
            error = re.sub(r'^Error:\s*', '', p).strip()

    return {
        'timestamp': orig_timestamp,
        'recipient': recipient,
        'status': status,
        'subject': subject,
        'message_id': message_id,
        'submitted_time': submitted_time,
        'bounce_time': bounce_time,
        'error': error
    }


def load_outreach_history(log_path: str) -> Dict[str, Dict[str, Any]]:
    """Load historical recipient outcomes from outreach.log.
    Key is lowercase recipient address.
    Tracks permanent/temporary bounces, last successful submission timestamp,
    and last status.
    Uses original submission timestamp (parts[0]) for cooldown calculations.
    """
    if not os.path.exists(log_path):
        return {}

    history: Dict[str, Dict[str, Any]] = {}
    with open(log_path, 'r', encoding='utf-8') as f:
        for line in f:
            entry = parse_log_line(line)
            if not entry:
                continue

            email_key = entry['recipient'].lower().strip()
            if not email_key:
                continue

            status = entry['status']
            ts_dt = parse_log_timestamp(entry['timestamp'])
            is_undelivered = status.startswith('UNDELIVERED') or status == 'BOUNCED'
            is_successful_submission = (
                status in ('SENT', 'SUBMITTED', 'UNKNOWN')
                and not is_undelivered
                and entry['bounce_time'] == 'N/A'
            )

            if email_key not in history:
                history[email_key] = {
                    'recipient': entry['recipient'],
                    'is_undelivered': is_undelivered,
                    'undelivered_status': status if is_undelivered else None,
                    'last_submitted_dt': ts_dt if is_successful_submission else None,
                    'last_status': status,
                    'last_error': entry['error'],
                    'entries': [entry]
                }
            else:
                h = history[email_key]
                h['entries'].append(entry)
                h['last_status'] = status
                h['last_error'] = entry['error']
                if is_undelivered:
                    h['is_undelivered'] = True
                    h['undelivered_status'] = status
                if is_successful_submission and ts_dt:
                    if h['last_submitted_dt'] is None or ts_dt > h['last_submitted_dt']:
                        h['last_submitted_dt'] = ts_dt

    return history


def evaluate_recipient_eligibility(
    recipient: str,
    history_map: Dict[str, Dict[str, Any]],
    cooldown_days: int = 10,
    cooldown_enabled: bool = True,
    current_time: Optional[datetime] = None
) -> Dict[str, Any]:
    """Evaluate whether an input recipient is eligible for outreach or should be skipped.
    Rules:
    - INVALID email format: skip until corrected.
    - UNDELIVERED in history: skip and flag for manual review, regardless of age
      (preserved even when cooldown is disabled).
    - SENT within cooldown period: skip (show previous submission & eligible dates).
    - SENT exactly cooldown_days ago or earlier: eligible to send.
    - FAILED_TO_SUBMIT: eligible for retry.
    - New recipient: eligible to send.
    """
    clean_email = recipient.strip()
    if not is_valid_email(clean_email):
        return {
            'action': 'INVALID',
            'reason': 'Invalid email address format; skipped until corrected',
            'recipient': clean_email
        }

    email_key = clean_email.lower()
    hist = history_map.get(email_key)

    if hist:
        # 1. UNDELIVERED in history: skip and flag for manual review, regardless of age
        if hist.get('is_undelivered'):
            return {
                'action': 'UNDELIVERED_REVIEW',
                'reason': f"Confirmed bounce: {hist.get('undelivered_status', 'UNDELIVERED')}",
                'status': hist.get('undelivered_status', 'UNDELIVERED'),
                'recipient': clean_email
            }

        # 2. Check cooldown for previously submitted emails
        last_dt = hist.get('last_submitted_dt')
        if last_dt is not None:
            if cooldown_enabled:
                now = current_time or datetime.now()
                cooldown_delta = timedelta(days=cooldown_days)
                elapsed = now - last_dt
                if elapsed < cooldown_delta:
                    # SENT within the last cooldown_days days: SKIP
                    eligible_dt = last_dt + cooldown_delta
                    return {
                        'action': 'SKIP_COOLDOWN',
                        'reason': f"Sent within {cooldown_days}-day cooldown",
                        'submitted_date': last_dt.strftime('%Y-%m-%d %H:%M:%S'),
                        'eligible_date': eligible_dt.strftime('%Y-%m-%d %H:%M:%S'),
                        'last_submitted_dt': last_dt,
                        'eligible_dt': eligible_dt,
                        'recipient': clean_email
                    }
                else:
                    # SENT exactly cooldown_days ago or earlier: eligible to send!
                    return {
                        'action': 'SEND',
                        'reason': f"Sent >= {cooldown_days} days ago (cooldown expired)",
                        'last_submitted_dt': last_dt,
                        'recipient': clean_email
                    }
            else:
                # Cooldown checking disabled: eligible to send
                return {
                    'action': 'SEND',
                    'reason': "Duplicate cooldown check disabled; eligible to send",
                    'last_submitted_dt': last_dt,
                    'recipient': clean_email
                }

        # 3. FAILED_TO_SUBMIT: eligible for retry
        if hist.get('last_status') == 'FAILED_TO_SUBMIT':
            return {
                'action': 'SEND',
                'reason': 'Previous submission failed; eligible for retry',
                'recipient': clean_email
            }

    # 4. New recipient not in history: eligible to send
    return {
        'action': 'SEND',
        'reason': 'New recipient; eligible to send',
        'recipient': clean_email
    }


def safe_clear_csv_queue(
    csv_path: str,
    safely_processed_emails: Set[str]
) -> int:
    """Safely update data/recruiter_contacts.csv temporary input queue.
    Removes rows whose recipient emails have confirmed, safely recorded outcomes
    (including skipped duplicates, undelivered review, invalid, and completed sends).
    Preserves the header and any unprocessed rows or rows with uncertain outcomes.
    If all input recipients were safely processed, leaves only the header.
    Never clears at startup.
    Thread-safe and crash-safe (via atomic temp-file replace and fsync).
    Returns count of removed rows.
    """
    if not os.path.exists(csv_path) or not safely_processed_emails:
        return 0

    lower_safe = {e.lower().strip() for e in safely_processed_emails if e and isinstance(e, str) and e.strip()}
    if not lower_safe:
        return 0

    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = list(csv.reader(f))
    except Exception:
        return 0

    if not reader:
        return 0

    first_row = reader[0]
    first_row_clean = [col.strip().lower() for col in first_row]
    email_col_idx = 0
    has_header = False

    if any('email' in col or 'mail' in col or 'recipient' in col for col in first_row_clean):
        has_header = True
        for idx, col in enumerate(first_row_clean):
            if 'email' in col or 'mail' in col or 'recipient' in col:
                email_col_idx = idx
                break

    header_row = first_row if has_header else None
    data_rows = reader[1:] if has_header else reader

    remaining_data_rows = []
    removed_count = 0

    for row in data_rows:
        if not row:
            continue
        row_email = row[email_col_idx].strip().lower() if len(row) > email_col_idx else ""
        if row_email and row_email in lower_safe:
            removed_count += 1
        else:
            remaining_data_rows.append(row)

    # Atomic write to temporary file then replace
    temp_csv = csv_path + ".tmp"
    try:
        with open(temp_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if header_row is not None:
                writer.writerow(header_row)
            for r in remaining_data_rows:
                writer.writerow(r)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass

        os.replace(temp_csv, csv_path)
    except Exception:
        if os.path.exists(temp_csv):
            try:
                os.remove(temp_csv)
            except OSError:
                pass
        return 0

    return removed_count


class RecruiterDirectory:
    """Thread-safe public recruiter email directory saved to data/recruiter_email_backup.xlsx.
    Maintains a single 'Email' column with case-insensitively deduplicated addresses.
    Preserves existing entries between runs.
    """

    def __init__(self, excel_path: str = DEFAULT_EXCEL_DIRECTORY):
        self.excel_path = excel_path
        self._lock = threading.Lock()
        self._emails_lower: Set[str] = set()
        self._init_file()

    def _init_file(self) -> None:
        if openpyxl is None:
            return
        with self._lock:
            parent = os.path.dirname(os.path.abspath(self.excel_path))
            if parent:
                os.makedirs(parent, exist_ok=True)
            if os.path.exists(self.excel_path):
                wb = None
                try:
                    wb = openpyxl.load_workbook(self.excel_path)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                        val = row[0]
                        if val and isinstance(val, str) and val.strip():
                            self._emails_lower.add(val.strip().lower())
                    return
                except Exception:
                    pass
                finally:
                    if wb:
                        wb.close()
            wb = openpyxl.Workbook()
            try:
                ws = wb.active
                ws.title = "Recruiters"
                ws.cell(row=1, column=1, value="Email")
                wb.save(self.excel_path)
            finally:
                wb.close()

    def add_email(self, email_str: str) -> bool:
        """Add a submitted email to the Excel directory if not already present (case-insensitive).
        Thread-safe. Returns True if added, False if duplicate or skipped.
        """
        if openpyxl is None:
            return False
        clean = email_str.strip()
        lower = clean.lower()
        if not lower:
            return False

        with self._lock:
            if lower in self._emails_lower:
                return False

            wb = None
            try:
                if os.path.exists(self.excel_path):
                    wb = openpyxl.load_workbook(self.excel_path)
                    ws = wb.active
                else:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Recruiters"
                    ws.cell(row=1, column=1, value="Email")

                if ws.cell(row=1, column=1).value != "Email":
                    ws.cell(row=1, column=1, value="Email")

                next_row = ws.max_row + 1
                ws.cell(row=next_row, column=1, value=clean)
                wb.save(self.excel_path)
                self._emails_lower.add(lower)
                return True
            except Exception:
                return False
            finally:
                if wb:
                    wb.close()

    def remove_email(self, email_str: str) -> bool:
        """Remove an email from the Excel directory if present (case-insensitive).
        Thread-safe. Returns True if removed, False otherwise.
        """
        return self.remove_emails([email_str]) > 0

    def remove_emails(self, email_list: List[str]) -> int:
        """Remove a collection of emails from the Excel directory (case-insensitive).
        Thread-safe. Preserves single 'Email' column and all other valid addresses.
        Returns count of removed emails.
        """
        if openpyxl is None or not os.path.exists(self.excel_path):
            return 0
        to_remove = {e.strip().lower() for e in email_list if e and isinstance(e, str) and e.strip()}
        if not to_remove:
            return 0

        with self._lock:
            matching = to_remove.intersection(self._emails_lower)
            if not matching:
                return 0

            wb = None
            try:
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active
                remaining = []
                for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                    val = row[0]
                    if val and isinstance(val, str) and val.strip():
                        if val.strip().lower() not in matching:
                            remaining.append(val.strip())

                new_ws = wb.create_sheet(title="_temp_recruiters", index=0)
                new_ws.cell(row=1, column=1, value="Email")
                for idx, email_val in enumerate(remaining, start=2):
                    new_ws.cell(row=idx, column=1, value=email_val)

                for sheet_name in list(wb.sheetnames):
                    if sheet_name != "_temp_recruiters":
                        wb.remove(wb[sheet_name])
                new_ws.title = "Recruiters"
                wb.save(self.excel_path)
                self._emails_lower.difference_update(matching)
                return len(matching)
            except Exception:
                return 0
            finally:
                if wb:
                    wb.close()

    def get_all_emails(self) -> List[str]:
        """Return all emails currently stored in the directory."""
        if openpyxl is None or not os.path.exists(self.excel_path):
            return []
        with self._lock:
            wb = None
            try:
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active
                emails = []
                for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                    val = row[0]
                    if val and isinstance(val, str) and val.strip():
                        emails.append(val.strip())
                return emails
            except Exception:
                return []
            finally:
                if wb:
                    wb.close()


class LiveDashboard:
    """Thread-safe live terminal dashboard tracking workers, progress, and counters."""

    def __init__(
        self,
        total_recipients: int,
        is_live: Optional[bool] = None,
        num_workers: int = 3,
        refresh_rate: float = 4.0,
        cooldown_days: int = 10,
        cooldown_enabled: bool = True
    ):
        self.total_recipients = max(0, total_recipients)
        self.is_live = sys.stdout.isatty() if is_live is None else is_live
        self.num_workers = max(1, num_workers)
        self.min_render_interval = 1.0 / max(0.1, refresh_rate)
        self.cooldown_days = max(1, cooldown_days)
        self.cooldown_enabled = cooldown_enabled
        self._lock = threading.Lock()
        self.start_time = time.time()
        self.completed = 0
        self.last_render_time = 0.0
        self.lines_rendered = 0

        # Worker panels configured for num_workers
        self.workers: Dict[int, Dict[str, str]] = {
            i: {'state': 'IDLE', 'recipient': '-'}
            for i in range(1, self.num_workers + 1)
        }

        # Counters: required display metrics
        self.counters: Dict[str, int] = {
            'TOTAL_INPUT': self.total_recipients,
            'NEW_RECIPIENTS': 0,
            'SKIPPED_COOLDOWN': 0,
            'UNDELIVERED_REVIEW': 0,
            'INVALID': 0,
            'FAILED_TO_SUBMIT': 0,
            'SENT': 0,
            'UNDELIVERED': 0
        }
        self.skipped_cooldown_recipients: List[Dict[str, Any]] = []
        self.undelivered_review_recipients: List[Dict[str, Any]] = []
        self.undelivered_recipients: List[Dict[str, Any]] = []
        self.execution_errors: List[Dict[str, Any]] = []

    def update_worker(self, worker_id: int, state: str, recipient: str) -> None:
        with self._lock:
            if worker_id in self.workers:
                self.workers[worker_id]['state'] = state
                self.workers[worker_id]['recipient'] = recipient if recipient else '-'
            self._render_locked()

    def increment_counter(self, counter_name: str, delta: int = 1) -> None:
        with self._lock:
            if counter_name in self.counters:
                self.counters[counter_name] += delta
            else:
                self.counters[counter_name] = delta
            self._render_locked()

    def set_counter(self, counter_name: str, value: int) -> None:
        with self._lock:
            self.counters[counter_name] = value
            self._render_locked()

    def set_skipped_cooldown_recipients(self, records: List[Dict[str, Any]]) -> None:
        with self._lock:
            self.skipped_cooldown_recipients = list(records)

    def set_undelivered_review_recipients(self, records: List[Dict[str, Any]]) -> None:
        with self._lock:
            self.undelivered_review_recipients = list(records)

    def set_undelivered_recipients(self, records: List[Dict[str, Any]]) -> None:
        with self._lock:
            self.undelivered_recipients = list(records)

    def add_execution_error(self, recipient: str, error_msg: str) -> None:
        with self._lock:
            self.execution_errors.append({'recipient': recipient, 'error': error_msg})

    def set_execution_errors(self, errors: List[Dict[str, Any]]) -> None:
        with self._lock:
            self.execution_errors = list(errors)

    def record_completed(self) -> None:
        with self._lock:
            self.completed += 1
            self._render_locked()

    def _render_locked(self, force: bool = False) -> None:
        now = time.time()
        if not force and (now - self.last_render_time) < self.min_render_interval:
            return
        self.last_render_time = now

        if not self.is_live:
            return

        elapsed = max(0.001, now - self.start_time)
        rate = (self.completed / elapsed) * 60.0
        pct = (self.completed / self.total_recipients * 100.0) if self.total_recipients > 0 else 100.0
        mins, secs = divmod(int(elapsed), 60)
        elapsed_str = f"{mins:02d}:{secs:02d}"

        bar_len = 25
        filled = int(bar_len * (self.completed / self.total_recipients)) if self.total_recipients > 0 else bar_len
        bar = "█" * filled + "░" * (bar_len - filled)

        cd_desc = f"{self.cooldown_days}-DAY COOLDOWN" if self.cooldown_enabled else "DISABLED"
        lines = [
            "======================= Recruiter Outreach Live Dashboard =======================",
            f"Config: Workers={self.num_workers} | Duplicate Check={cd_desc}",
            f"Progress: [{bar}] {self.completed}/{self.total_recipients} ({pct:.1f}%) | Time: {elapsed_str} | Rate: {rate:.1f} msg/min",
            "--------------------------------- Worker Panels ---------------------------------"
        ]
        for w_id in range(1, self.num_workers + 1):
            w = self.workers.get(w_id, {'state': 'IDLE', 'recipient': '-'})
            lines.append(f" [Worker {w_id}] State: {w['state']:<10} | Current: {w['recipient']}")

        lines.extend([
            "------------------------------------ Counters -----------------------------------",
            f" TOTAL INPUT: {self.total_recipients} | NEW RECIPIENTS: {self.counters.get('NEW_RECIPIENTS', 0)} | SENT: {self.counters.get('SENT', 0)}",
            f" SKIPPED ({self.cooldown_days}-DAY COOLDOWN): {self.counters.get('SKIPPED_COOLDOWN', 0)} | UNDELIVERED REVIEW: {self.counters.get('UNDELIVERED_REVIEW', 0)}",
            f" INVALID: {self.counters.get('INVALID', 0)} | FAILED TO SUBMIT: {self.counters.get('FAILED_TO_SUBMIT', 0)} | UNDELIVERED: {self.counters.get('UNDELIVERED', 0)}",
            "================================================================================="
        ])

        if self.lines_rendered > 0:
            sys.stdout.write(f"\033[{self.lines_rendered}A")
        for line in lines:
            sys.stdout.write(f"\033[2K\r{line}\n")
        sys.stdout.flush()
        self.lines_rendered = len(lines)

    def render_final_summary(
        self,
        log_path: str = LOG_FILE,
        excel_path: str = DEFAULT_EXCEL_DIRECTORY,
        undelivered_recipients: Optional[List[Dict[str, Any]]] = None,
        execution_errors: Optional[List[Dict[str, Any]]] = None
    ) -> None:
        with self._lock:
            self._render_locked(force=True)
            elapsed = max(0.001, time.time() - self.start_time)
            rate = (self.completed / elapsed) * 60.0
            pct = (self.completed / self.total_recipients * 100.0) if self.total_recipients > 0 else 100.0
            mins, secs = divmod(int(elapsed), 60)
            elapsed_str = f"{mins:02d}:{secs:02d}"

            sent_cnt = self.counters.get('SENT', 0)
            undelivered_cnt = self.counters.get('UNDELIVERED', 0)
            skipped_cnt = self.counters.get('SKIPPED_COOLDOWN', 0)
            review_cnt = self.counters.get('UNDELIVERED_REVIEW', 0)
            invalid_cnt = self.counters.get('INVALID', 0)
            failed_cnt = self.counters.get('FAILED_TO_SUBMIT', 0)
            new_cnt = self.counters.get('NEW_RECIPIENTS', 0)

            cd_str = f"ENABLED ({self.cooldown_days}-Day Cooldown)" if self.cooldown_enabled else "DISABLED"

            summary = [
                "",
                "=========================== Final Execution Summary ===========================",
                f"Duplicate Check:      {cd_str}",
                f"Total Input:          {self.total_recipients}",
                f"New Recipients:       {new_cnt}",
                f"Processed:            {self.completed}",
                f"Progress:             {pct:.1f}%",
                f"Elapsed Time:         {elapsed_str}",
                f"Processing Rate:      {rate:.1f} emails/min",
                "",
                "--- Status Counters ---",
                f"TOTAL INPUT:          {self.total_recipients}",
                f"NEW RECIPIENTS:       {new_cnt}",
                f"SKIPPED ({self.cooldown_days}-DAY COOLDOWN): {skipped_cnt}",
                f"UNDELIVERED REVIEW:   {review_cnt}",
                f"INVALID:              {invalid_cnt}",
                f"FAILED TO SUBMIT:     {failed_cnt}",
                f"SENT:                 {sent_cnt}",
                f"UNDELIVERED:          {undelivered_cnt}"
            ]

            if self.skipped_cooldown_recipients:
                summary.append("")
                summary.append(f"--- Skipped Recipients ({self.cooldown_days}-Day Cooldown) ---")
                for s in self.skipped_cooldown_recipients:
                    summary.append(f"  - {s['recipient']}: Previous Submission: {s.get('submitted_date', 'N/A')} | Eligible Again: {s.get('eligible_date', 'N/A')}")

            if self.undelivered_review_recipients:
                summary.append("")
                summary.append("--- Undelivered Review (Flagged for Manual Review) ---")
                for u in self.undelivered_review_recipients:
                    summary.append(f"  - {u['recipient']}: Status: {u.get('status', 'UNDELIVERED')}")

            undelivered = undelivered_recipients if undelivered_recipients is not None else self.undelivered_recipients
            if undelivered:
                summary.append("")
                summary.append("--- Undelivered Recipients ---")
                for u in undelivered:
                    summary.append(f"Recipient: {u['recipient']}")
                    summary.append(f"Status: {u['status']}")

            errs = execution_errors if execution_errors is not None else self.execution_errors
            if errs:
                summary.append("")
                summary.append("--- Execution Error Details ---")
                for e in errs:
                    summary.append(f"  - {e['recipient']}: {e.get('error', 'Error')}")

            summary.extend([
                "",
                f"Logs:                 {log_path}",
                f"Public Directory:     {excel_path}",
                "===============================================================================",
                ""
            ])
            print("\n".join(summary))


def load_template(template_path: str) -> str:
    """Read HTML template file and return HTML content.
    Email subject is configured exclusively via .env (EMAIL_SUBJECT), not in the template.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Email template not found: {template_path}")

    with open(template_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Strip legacy subject comment if present so it doesn't appear in the HTML body
    cleaned = SUBJECT_REGEX.sub('', content).lstrip()
    return cleaned


def parse_template(template_path: str) -> Tuple[str, str]:
    """Compatibility helper: reads HTML template file and returns (subject, html_body).
    Email subject is retrieved from .env EMAIL_SUBJECT as the single source of truth.
    """
    html_body = load_template(template_path)
    cfg = get_config()
    subject = (cfg.email_subject or '').strip()
    return subject, html_body


def is_valid_email(email_str: str) -> bool:
    """Validate format of an email address."""
    return bool(EMAIL_REGEX.match(email_str.strip()))


def get_gmail_service(
    credentials_file: Optional[str] = None,
    token_file: Optional[str] = None
):
    """Authenticate and return Gmail API service instance."""
    if None in (Request, Credentials, InstalledAppFlow, build):
        raise ImportError(
            "Missing Google client libraries. Run 'pip install -r requirements.txt' first."
        )

    cfg = get_config()
    creds_path = credentials_file or cfg.gmail_credentials_file
    tok_path = token_file or cfg.gmail_token_file

    # Fallback to config/ if not found in root
    if not os.path.exists(creds_path) and os.path.exists(os.path.join('config', 'credentials.json')):
        creds_path = os.path.join('config', 'credentials.json')

    if not os.path.exists(tok_path) and os.path.exists(os.path.join('config', 'token.json')):
        tok_path = os.path.join('config', 'token.json')

    creds = None
    if os.path.exists(tok_path):
        try:
            import json
            with open(tok_path, 'r', encoding='utf-8') as tf:
                token_data = json.load(tf)
            stored_scopes = set(token_data.get('scopes', []))
            # Verify stored token actually grants every required scope in SCOPES
            if set(SCOPES).issubset(stored_scopes):
                creds = Credentials.from_authorized_user_file(tok_path, SCOPES)
            else:
                creds = None
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                raise FileNotFoundError(
                    f"'{creds_path}' not found. Please download OAuth client credentials "
                    "from Google Cloud Console and place them in config/credentials.json."
                )
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)

        token_dir = os.path.dirname(tok_path)
        if token_dir:
            os.makedirs(token_dir, exist_ok=True)
        with open(tok_path, 'w') as token_file:
            token_file.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    service._agent_credentials = creds
    _THREAD_LOCAL.service = service
    return service


_THREAD_LOCAL = threading.local()


def get_thread_safe_service(base_service):
    """Return a thread-isolated Gmail API service client.
    Each thread gets its own Resource instance and independent httplib2.Http transport,
    preventing C-level SSL heap corruption in libmalloc on macOS ARM64.
    If base_service is a mock, None, or lacks credentials, returns base_service directly.
    """
    if base_service is None:
        return None

    # Transparent mock compatibility for unit tests
    if (
        hasattr(base_service, '_mock_return_value')
        or hasattr(base_service, 'mock_calls')
        or hasattr(base_service, 'assert_called_with')
    ):
        return base_service

    # Check credentials
    creds = getattr(base_service, '_agent_credentials', None)
    if not creds:
        creds = getattr(base_service, '_credentials', None)

    if not creds or build is None:
        return base_service

    thread_service = getattr(_THREAD_LOCAL, 'service', None)
    if thread_service is None:
        try:
            thread_service = build('gmail', 'v1', credentials=creds)
            thread_service._agent_credentials = creds
            _THREAD_LOCAL.service = thread_service
        except Exception:
            return base_service

    return thread_service


def read_raw_recruiter_emails(csv_path: str) -> List[str]:
    """Read all raw recipient entries from CSV file in row order, preserving order and duplicates."""
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    emails: List[str] = []
    with open(csv_path, mode='r', encoding='utf-8') as f:
        reader = csv.reader(f)
        try:
            first_row = next(reader)
        except StopIteration:
            return emails

        first_row_clean = [col.strip().lower() for col in first_row]
        email_col_idx = 0
        has_header = False

        if any('email' in col or 'mail' in col or 'recipient' in col for col in first_row_clean):
            has_header = True
            for idx, col in enumerate(first_row_clean):
                if 'email' in col or 'mail' in col or 'recipient' in col:
                    email_col_idx = idx
                    break

        # Process first row if it was data rather than header
        if not has_header and first_row:
            val = first_row[0].strip()
            if val:
                emails.append(val)

        for row in reader:
            if not row or len(row) <= email_col_idx:
                continue
            val = row[email_col_idx].strip()
            if val:
                emails.append(val)

    return emails


def read_recruiter_emails(csv_path: str) -> Tuple[List[str], List[str]]:
    """
    Read recruiter emails from CSV file in a single streaming pass.
    Validates format, eliminates duplicates, and returns (valid_emails, invalid_emails).
    """
    raw_list = read_raw_recruiter_emails(csv_path)
    valid_emails = []
    invalid_emails = []
    seen = set()

    for email_val in raw_list:
        norm = email_val.lower()
        if norm in seen:
            continue
        seen.add(norm)
        if is_valid_email(email_val):
            valid_emails.append(email_val)
        else:
            invalid_emails.append(email_val)

    return valid_emails, invalid_emails


_RESUME_CACHE: Dict[str, Tuple[str, bytes]] = {}
_RESUME_CACHE_LOCK = threading.Lock()


def get_cached_resume(resume_path: str) -> Tuple[str, bytes]:
    """Retrieve filename and bytes of resume, caching in-memory to prevent repeated disk reads."""
    with _RESUME_CACHE_LOCK:
        if resume_path in _RESUME_CACHE:
            return _RESUME_CACHE[resume_path]
    with open(resume_path, 'rb') as f:
        data = f.read()
    filename = os.path.basename(resume_path)
    with _RESUME_CACHE_LOCK:
        _RESUME_CACHE[resume_path] = (filename, data)
    return filename, data


def create_message(
    to_email: str,
    subject: str,
    html_content: str,
    resume_path: str,
    resume_data: Optional[Tuple[str, bytes]] = None,
    from_email: Optional[str] = None
) -> dict:
    """Create a MIME message with HTML body and PDF attachment, encoded for Gmail API.
    Reuses in-memory cached resume binary data to eliminate redundant disk reads.
    """
    message = MIMEMultipart()
    message['to'] = to_email
    if from_email:
        message['from'] = from_email
    message['subject'] = subject

    # Attach HTML body
    body_part = MIMEText(html_content, 'html', 'utf-8')
    message.attach(body_part)

    # Attach resume PDF using cached bytes
    if resume_data is not None:
        filename, data = resume_data
    else:
        filename, data = get_cached_resume(resume_path)

    attachment = MIMEApplication(data, _subtype='pdf')
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


def execute_with_retry(
    api_call_fn,
    max_retries: int = 3,
    initial_delay: float = 0.4,
    backoff_factor: float = 2.0
) -> Any:
    """Execute an API call with exponential backoff for transient errors.
    Retries on HTTP 429 (rate limit), 500, 502, 503, 504, and transient network errors.
    Never retries permanent failures (HTTP 400, 401, 403, 404).
    """
    delay = initial_delay
    last_err = None

    for attempt in range(max_retries + 1):
        try:
            return api_call_fn()
        except HttpError as error:
            status_code = getattr(getattr(error, 'resp', None), 'status', None)
            if status_code in (429, 500, 502, 503, 504) and attempt < max_retries:
                time.sleep(delay)
                delay *= backoff_factor
                last_err = error
                continue
            raise
        except (ConnectionError, TimeoutError, OSError) as error:
            if attempt < max_retries:
                time.sleep(delay)
                delay *= backoff_factor
                last_err = error
                continue
            raise

    if last_err:
        raise last_err


TEMPORARY_BOUNCE_PATTERNS = [
    r'\bmailbox\s+full\b',
    r'\bover\s+quota\b',
    r'\bquota\s+exceeded\b',
    r'\bstorage\s+quota\b',
    r'\bexceeded\s+storage\b',
    r'\btemporar(?:y|ily)\b',
    r'\b(?:delivery|message|mail)?\s*temporar(?:y|ily)\s+delayed?\b',
    r'\b(?:delivery|message|mail)\s+delayed?\b',
    r'\bdelayed?\s+(?:delivery|mail|message)\b',
    r'\bgreylist(?:ed|ing)?\b',
    r'\btry\s+again\b',
    r'\bretry\s+later\b',
    r'\brate\s+limit\b',
    r'\btoo\s+many\s+connections\b',
    r'\bserver\s+busy\b',
    r'\b4\.\d\.\d\b',
    r'\b45[0-2]\b',
    r'\b552\b',
    r'\b5\.2\.2\b'
]

PERMANENT_INVALID_RECIPIENT_PATTERNS = [
    r'\bdoes\s+not\s+exist\b',
    r'\baddress\s+(?:couldn\'?t\s+be\s+found|could\s+not\s+be\s+found|not\s+found)\b',
    r'\binvalid\s+recipient\b',
    r'\brecipient\s+invalid\b',
    r'\brecipient\s+(?:not\s+found|rejected|unknown)\b',
    r'\buser\s+(?:unknown|not\s+found|does\s+not\s+exist)\b',
    r'\bno\s+such\s+(?:mailbox|user|recipient|address|person)\b',
    r'\bmailbox\s+(?:does\s+not\s+exist|not\s+found|unavailable)\b',
    r'\baccount\s+(?:does\s+not\s+exist|disabled|not\s+found)\b',
    r'\b5\.1\.1\b',
    r'\b5\.1\.2\b',
    r'\b550[- ]5\.1\.1\b',
]


def is_permanent_invalid_recipient_bounce(bounce_info: Any) -> bool:
    """Determine whether a bounce notification indicates a confirmed permanent invalid recipient
    (e.g., mailbox does not exist, address not found, user unknown, no such mailbox).
    Returns False for temporary failures (e.g. mailbox full, server errors, greylisting, delayed).
    """
    if isinstance(bounce_info, dict):
        text_to_check = " ".join([
            str(bounce_info.get('reason', '')),
            str(bounce_info.get('snippet', '')),
            str(bounce_info.get('full_text_lower', '')),
            str(bounce_info.get('error_code', ''))
        ]).lower()
    else:
        text_to_check = str(bounce_info).lower()

    # 1. Check for temporary failure indicators first
    for pat in TEMPORARY_BOUNCE_PATTERNS:
        if re.search(pat, text_to_check):
            return False

    # 2. Check for permanent invalid recipient indicators
    for pat in PERMANENT_INVALID_RECIPIENT_PATTERNS:
        if re.search(pat, text_to_check):
            return True

    return False


def detect_bounces(
    service,
    submitted_records: List[Dict[str, Any]],
    logger: Optional[OutreachLogger] = None,
    directory: Optional[RecruiterDirectory] = None,
    delay_seconds: int = 0,
    cached_bounces: Optional[List[Dict[str, Any]]] = None,
    config: Optional[Config] = None
) -> List[Dict[str, Any]]:
    """
    Check Gmail inbox for delivery failure notifications from mailer-daemon,
    mail delivery subsystem, or postmaster.
    Uses concurrent message retrieval and batch logging for high performance.
    Matches notifications to submitted records, updates status to UNDELIVERED (reason),
    removes confirmed permanent invalid recipients from the Excel directory,
    and updates outreach.log in-place without duplicate records.
    """
    if not submitted_records:
        return submitted_records

    cfg = config or get_config()

    if cached_bounces is not None and len(cached_bounces) > 0:
        bounces = cached_bounces
    else:
        if delay_seconds > 0:
            time.sleep(delay_seconds)

        query = 'from:(mailer-daemon OR "Mail Delivery Subsystem" OR postmaster)'
        svc = get_thread_safe_service(service)
        try:
            resp = execute_with_retry(
                lambda: svc.users().messages().list(userId='me', q=query, maxResults=50).execute()
            )
            messages = resp.get('messages', [])
        except Exception:
            # Never downgrade or mark UNDELIVERED on bounce-check errors or missing permissions
            return submitted_records

        # Concurrently retrieve and parse bounce notifications
        bounces = []
        if messages:
            def fetch_single_bounce(msg_meta):
                msg_id = msg_meta.get('id')
                try:
                    thread_svc = get_thread_safe_service(service)
                    msg_full = execute_with_retry(
                        lambda: thread_svc.users().messages().get(userId='me', id=msg_id, format='full').execute()
                    )
                    return parse_bounce_notification(msg_full)
                except Exception:
                    return None

            max_fetch_workers = min(cfg.num_workers, len(messages))
            if max_fetch_workers > 1:
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_fetch_workers) as executor:
                    parsed_results = list(executor.map(fetch_single_bounce, messages))
                bounces = [p for p in parsed_results if p is not None]
            else:
                for msg_meta in messages:
                    res = fetch_single_bounce(msg_meta)
                    if res:
                        bounces.append(res)

        if cached_bounces is not None:
            cached_bounces.extend(bounces)

    bounce_time = datetime.now().strftime('%H:%M:%S')
    pending_updates = {}
    permanent_removals = []

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
            reason = matched_bounce['reason'].strip()
            new_status = reason if reason.startswith('UNDELIVERED') else f"UNDELIVERED ({reason})"
            rec['status'] = new_status
            rec['bounce_time'] = bounce_time
            rec['remark'] = reason
            rec['error'] = matched_bounce['error_code'] if matched_bounce['error_code'] != "None" else reason
            pending_updates[rec['recipient']] = {
                'new_status': new_status,
                'error': rec['error'],
                'remark': rec['remark'],
                'bounce_time': rec['bounce_time'],
                'message_id': rec.get('message_id'),
                'subject': rec.get('subject')
            }
            if directory and is_permanent_invalid_recipient_bounce(matched_bounce):
                permanent_removals.append(rec['recipient'])
        else:
            # Unmatched records remain SENT (never changed to UNKNOWN)
            if rec.get('status') in ('SUBMITTED', None):
                rec['status'] = 'SENT'

    # Batch update log file in a single file pass
    if logger and pending_updates:
        logger.batch_update_email_records(pending_updates)

    # Remove permanent invalid recipient addresses from Excel directory
    if directory and permanent_removals:
        directory.remove_emails(permanent_removals)

    return submitted_records


def process_recipients_concurrently(
    recipients: List[str],
    service,
    subject: str,
    html_template: str,
    resume_path: str,
    logger: Optional[OutreachLogger] = None,
    directory: Optional[RecruiterDirectory] = None,
    dashboard: Optional[LiveDashboard] = None,
    num_workers: Optional[int] = None,
    api_call_delay: Optional[float] = None,
    config: Optional[Config] = None,
    safely_processed_set: Optional[Set[str]] = None
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Process recipients across concurrent workers independently.
    Handles deduplication within run, validation, MIME creation, thread-safe sending,
    Excel directory registration for SUBMITTED emails, live dashboard state updates,
    and DRY_RUN mode.
    Returns (all_records, submitted_records).
    """
    cfg = config or get_config()
    workers_count = num_workers if num_workers is not None else cfg.num_workers
    delay = api_call_delay if api_call_delay is not None else cfg.gmail_api_request_delay
    retries = cfg.max_retries
    retry_delay = cfg.retry_backoff_seconds
    is_dry_run = cfg.dry_run
    from_email = cfg.sender_email
    logging_active = cfg.logging_enabled and logger is not None
    excel_active = cfg.excel_directory_enabled and directory is not None

    all_records: List[Dict[str, Any]] = []
    submitted_records: List[Dict[str, Any]] = []

    seen_lock = threading.Lock()
    seen_in_run: Set[str] = set()

    api_lock = threading.Lock()
    log_lock = threading.Lock()
    records_lock = threading.Lock()

    worker_slots: queue.Queue = queue.Queue()
    for i in range(1, workers_count + 1):
        worker_slots.put(i)

    resume_data = get_cached_resume(resume_path)

    def worker_task(recipient_raw: str) -> Optional[Dict[str, Any]]:
        slot = worker_slots.get()
        norm_email = recipient_raw.strip()
        lower_email = norm_email.lower()

        try:
            # 1. State: VALIDATING
            if dashboard:
                dashboard.update_worker(slot, state='VALIDATING', recipient=norm_email)

            # In-run duplicate check
            with seen_lock:
                if lower_email in seen_in_run:
                    is_duplicate = True
                else:
                    seen_in_run.add(lower_email)
                    is_duplicate = False

            if is_duplicate:
                if safely_processed_set is not None:
                    safely_processed_set.add(lower_email)
                if dashboard:
                    dashboard.increment_counter('SKIPPED')
                    dashboard.record_completed()
                    dashboard.update_worker(slot, state='IDLE', recipient='-')
                return None

            # Email format validation
            if not is_valid_email(norm_email):
                now = datetime.now()
                rec = {
                    'timestamp': now.strftime('%Y-%m-%d %H:%M:%S'),
                    'recipient': norm_email,
                    'status': 'INVALID',
                    'subject': subject,
                    'message_id': 'N/A',
                    'submitted_time': 'N/A',
                    'bounce_time': 'N/A',
                    'error': 'Invalid email address format; skipped'
                }
                if logging_active:
                    with log_lock:
                        logger.log_email_record(rec)
                if safely_processed_set is not None:
                    safely_processed_set.add(lower_email)
                with records_lock:
                    all_records.append(rec)
                if dashboard:
                    dashboard.increment_counter('INVALID')
                    dashboard.add_execution_error(norm_email, 'Invalid email address format; skipped')
                    dashboard.update_worker(slot, state='FAILED', recipient=norm_email)
                    dashboard.record_completed()
                    dashboard.update_worker(slot, state='IDLE', recipient='-')
                return rec

            # 2. State: SENDING
            if dashboard:
                dashboard.update_worker(slot, state='SENDING', recipient=norm_email)

            now = datetime.now()
            rec = {
                'timestamp': now.strftime('%Y-%m-%d %H:%M:%S'),
                'recipient': norm_email,
                'subject': subject,
                'status': None,
                'message_id': 'N/A',
                'submitted_time': now.strftime('%H:%M:%S'),
                'bounce_time': 'N/A',
                'error': 'None'
            }

            if is_dry_run:
                # DRY RUN: simulate submission safely without sending real email
                if delay > 0:
                    time.sleep(delay)
                msg_id = f"dry_run_{abs(hash(norm_email)) % 1000000:06d}"
                rec['message_id'] = msg_id
                rec['status'] = 'SENT'

                if excel_active:
                    directory.add_email(norm_email)

                if logging_active:
                    with log_lock:
                        logger.log_email_record(rec)

                if safely_processed_set is not None:
                    safely_processed_set.add(lower_email)

                with records_lock:
                    all_records.append(rec)
                    submitted_records.append(rec)

                if dashboard:
                    dashboard.increment_counter('SENT')
                    dashboard.update_worker(slot, state='SENT', recipient=norm_email)
                    dashboard.record_completed()
                    dashboard.update_worker(slot, state='IDLE', recipient='-')

                return rec

            try:
                message_body = create_message(
                    to_email=norm_email,
                    subject=subject,
                    html_content=html_template,
                    resume_path=resume_path,
                    resume_data=resume_data,
                    from_email=from_email
                )

                thread_svc = get_thread_safe_service(service)

                def api_call():
                    with api_lock:
                        if delay > 0:
                            time.sleep(delay)
                    return thread_svc.users().messages().send(userId='me', body=message_body).execute()

                resp = execute_with_retry(api_call, max_retries=retries, initial_delay=retry_delay)
                msg_id = resp.get('id', 'N/A')
                rec['message_id'] = msg_id
                rec['status'] = 'SENT'

                # Add to public Excel directory if enabled
                if excel_active:
                    directory.add_email(norm_email)

                if logging_active:
                    with log_lock:
                        logger.log_email_record(rec)

                if safely_processed_set is not None:
                    safely_processed_set.add(lower_email)

                with records_lock:
                    all_records.append(rec)
                    submitted_records.append(rec)

                if dashboard:
                    dashboard.increment_counter('SENT')
                    dashboard.update_worker(slot, state='SENT', recipient=norm_email)
                    dashboard.record_completed()
                    dashboard.update_worker(slot, state='IDLE', recipient='-')

                return rec

            except HttpError as error:
                err_text = f"HTTP {error.resp.status}: {error.reason}"
                rec['status'] = 'FAILED_TO_SUBMIT'
                rec['submitted_time'] = 'N/A'
                rec['error'] = err_text

                if logging_active:
                    with log_lock:
                        logger.log_email_record(rec)

                if safely_processed_set is not None:
                    safely_processed_set.add(lower_email)

                with records_lock:
                    all_records.append(rec)

                if dashboard:
                    dashboard.increment_counter('FAILED_TO_SUBMIT')
                    dashboard.add_execution_error(norm_email, err_text)
                    dashboard.update_worker(slot, state='FAILED', recipient=norm_email)
                    dashboard.record_completed()
                    dashboard.update_worker(slot, state='IDLE', recipient='-')

                return rec

            except Exception as error:
                rec['status'] = 'FAILED_TO_SUBMIT'
                rec['submitted_time'] = 'N/A'
                rec['error'] = str(error)

                if logging_active:
                    with log_lock:
                        logger.log_email_record(rec)

                if safely_processed_set is not None:
                    safely_processed_set.add(lower_email)

                with records_lock:
                    all_records.append(rec)

                if dashboard:
                    dashboard.increment_counter('FAILED_TO_SUBMIT')
                    dashboard.add_execution_error(norm_email, str(error))
                    dashboard.update_worker(slot, state='FAILED', recipient=norm_email)
                    dashboard.record_completed()
                    dashboard.update_worker(slot, state='IDLE', recipient='-')

                return rec

        finally:
            worker_slots.put(slot)

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers_count) as executor:
        futures = [executor.submit(worker_task, r) for r in recipients]
        for f in concurrent.futures.as_completed(futures):
            f.result()

    return all_records, submitted_records


def main():
    config = get_config()
    try:
        config.validate()
    except Exception as e:
        print(f"Configuration Error: {e}")
        return

    logger = OutreachLogger(config.log_file) if config.logging_enabled else None
    directory = RecruiterDirectory(config.excel_directory_file) if config.excel_directory_enabled else None

    # Startup status banner reflecting active settings
    cd_status = f"ENABLED ({config.duplicate_cooldown_days}-Day Cooldown)" if config.duplicate_check_enabled else "DISABLED"
    print("===============================================================================")
    print(f"[*] Starting Recruiter Gmail Outreach Agent")
    print(f"[*] Config: Workers={config.num_workers} | Duplicate Check={cd_status} | Dry Run={config.dry_run}")
    print("===============================================================================")

    # 1. Check prerequisite files & parse template
    if not os.path.exists(config.resume_file):
        print(f"Error: Resume attachment '{config.resume_file}' not found.")
        return

    if not os.path.exists(config.email_template_file):
        print(f"Error: Email template '{config.email_template_file}' not found.")
        return

    subject = (config.email_subject or '').strip()
    if not subject:
        print("Error: EMAIL_SUBJECT must be set in .env. Email subject is configured exclusively via .env.")
        return

    try:
        html_template = load_template(config.email_template_file)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return

    # 2. Read raw recruiter emails from CSV (temporary input queue; never cleared at startup)
    try:
        raw_recipients = read_raw_recruiter_emails(config.recruiter_csv_file)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return

    total_recipients = len(raw_recipients)
    if total_recipients == 0:
        print("No email addresses found in CSV.")
        return

    # 3. Pre-flight Duplicate & Eligibility check against outreach.log before worker assignment
    history_map = load_outreach_history(config.log_file) if (config.logging_enabled and os.path.exists(config.log_file)) else {}

    eligible_recipients: List[str] = []
    skipped_cooldown_records: List[Dict[str, Any]] = []
    undelivered_review_records: List[Dict[str, Any]] = []
    invalid_records: List[Dict[str, Any]] = []
    seen_in_input: Set[str] = set()
    safely_processed: Set[str] = set()

    for r_raw in raw_recipients:
        r_clean = r_raw.strip()
        r_lower = r_clean.lower()
        if not r_clean:
            continue

        # Prevent queuing duplicates within the input queue in the same execution
        if r_lower in seen_in_input:
            safely_processed.add(r_lower)
            continue
        seen_in_input.add(r_lower)

        eval_res = evaluate_recipient_eligibility(
            recipient=r_clean,
            history_map=history_map,
            cooldown_days=config.duplicate_cooldown_days,
            cooldown_enabled=config.duplicate_check_enabled
        )

        action = eval_res['action']
        if action == 'SKIP_COOLDOWN':
            skipped_cooldown_records.append({
                'recipient': r_clean,
                'submitted_date': eval_res.get('submitted_date', 'N/A'),
                'eligible_date': eval_res.get('eligible_date', 'N/A'),
                'reason': eval_res.get('reason', '')
            })
            safely_processed.add(r_lower)
        elif action == 'UNDELIVERED_REVIEW':
            undelivered_review_records.append({
                'recipient': r_clean,
                'status': eval_res.get('status', 'UNDELIVERED'),
                'reason': eval_res.get('reason', '')
            })
            safely_processed.add(r_lower)
        elif action == 'INVALID':
            inv_rec = {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'recipient': r_clean,
                'status': 'INVALID',
                'subject': subject,
                'message_id': 'N/A',
                'submitted_time': 'N/A',
                'bounce_time': 'N/A',
                'error': 'Invalid email address format; skipped'
            }
            if logger and config.logging_enabled:
                logger.log_email_record(inv_rec)
            invalid_records.append(inv_rec)
            safely_processed.add(r_lower)
        else: # SEND
            eligible_recipients.append(r_clean)

    # 4. Authenticate with Gmail API (skip if dry_run or no eligible recipients)
    service = None
    if eligible_recipients and not config.dry_run:
        try:
            service = get_gmail_service(
                credentials_file=config.gmail_credentials_file,
                token_file=config.gmail_token_file
            )
        except Exception as e:
            print(f"Authentication failed: {e}")
            return
    elif config.dry_run:
        print("Note: Running in DRY_RUN mode. Real emails will not be sent.")

    # 5. Initialize dashboard if enabled
    dashboard = None
    if config.dashboard_enabled:
        dashboard = LiveDashboard(
            total_recipients=total_recipients,
            num_workers=config.num_workers,
            refresh_rate=config.dashboard_refresh_rate,
            cooldown_days=config.duplicate_cooldown_days,
            cooldown_enabled=config.duplicate_check_enabled
        )
        dashboard.set_counter('TOTAL_INPUT', total_recipients)
        dashboard.set_counter('NEW_RECIPIENTS', len(eligible_recipients))
        dashboard.set_counter('SKIPPED_COOLDOWN', len(skipped_cooldown_records))
        dashboard.set_counter('UNDELIVERED_REVIEW', len(undelivered_review_records))
        dashboard.set_counter('INVALID', len(invalid_records))
        dashboard.set_skipped_cooldown_recipients(skipped_cooldown_records)
        dashboard.set_undelivered_review_recipients(undelivered_review_records)
        for inv in invalid_records:
            dashboard.add_execution_error(inv['recipient'], inv['error'])

    all_records: List[Dict[str, Any]] = list(invalid_records)
    submitted_records: List[Dict[str, Any]] = []

    try:
        # 6. Process eligible recipients concurrently across 3 workers
        if eligible_recipients:
            worker_all, worker_sub = process_recipients_concurrently(
                recipients=eligible_recipients,
                service=service,
                subject=subject,
                html_template=html_template,
                resume_path=config.resume_file,
                logger=logger,
                directory=directory,
                dashboard=dashboard,
                config=config,
                safely_processed_set=safely_processed
            )
            all_records.extend(worker_all)
            submitted_records.extend(worker_sub)

        # 7. Check for Bounces (Current run + Delayed bounces from previous unconfirmed emails)
        if config.bounce_check_enabled and service and not config.dry_run:
            cached_bounces: List[Dict[str, Any]] = []
            if submitted_records:
                detect_bounces(
                    service,
                    submitted_records,
                    logger=logger,
                    directory=directory,
                    delay_seconds=config.bounce_check_delay_seconds,
                    cached_bounces=cached_bounces,
                    config=config
                )

            if logger:
                try:
                    prev_unconfirmed = logger.get_unconfirmed_records()
                    curr_emails = {r['recipient'].lower() for r in submitted_records}
                    prev_to_check = [r for r in prev_unconfirmed if r['recipient'].lower() not in curr_emails]
                    if prev_to_check:
                        detect_bounces(
                            service,
                            prev_to_check,
                            logger=logger,
                            directory=directory,
                            delay_seconds=0,
                            cached_bounces=cached_bounces if cached_bounces else None,
                            config=config
                        )
                except Exception:
                    pass

    finally:
        # 8. Safe CSV auto-clear: clear safely processed recipient rows while preserving header and unprocessed rows
        if safely_processed and os.path.exists(config.recruiter_csv_file):
            safe_clear_csv_queue(config.recruiter_csv_file, safely_processed)

    # 9. Final Summary
    undelivered_records = [
        r for r in submitted_records
        if str(r.get('status', '')).startswith('UNDELIVERED') or r.get('status') == 'BOUNCED'
    ]
    sent_count = max(0, len(submitted_records) - len(undelivered_records))
    undelivered_count = len(undelivered_records)
    failed_submit_count = sum(1 for r in all_records if r.get('status') == 'FAILED_TO_SUBMIT')

    execution_errors = [
        r for r in all_records
        if r.get('status') in ('FAILED_TO_SUBMIT', 'INVALID')
    ]

    if dashboard:
        dashboard.set_counter('SENT', sent_count)
        dashboard.set_counter('UNDELIVERED', undelivered_count)
        dashboard.set_counter('FAILED_TO_SUBMIT', failed_submit_count)
        dashboard.set_undelivered_recipients(undelivered_records)
        dashboard.set_execution_errors(execution_errors)
        dashboard.render_final_summary(
            log_path=config.log_file if config.logging_enabled else "Disabled",
            excel_path=config.excel_directory_file if config.excel_directory_enabled else "Disabled"
        )
    else:
        print("\n=== Final Execution Summary ===")
        print(f"Duplicate Check:      {cd_status}")
        print(f"Total Input:          {total_recipients}")
        print(f"New Recipients:       {len(eligible_recipients)}")
        print(f"SKIPPED ({config.duplicate_cooldown_days}-DAY COOLDOWN):   {len(skipped_cooldown_records)}")
        print(f"UNDELIVERED REVIEW:   {len(undelivered_review_records)}")
        print(f"INVALID:              {len(invalid_records)}")
        print(f"FAILED TO SUBMIT:     {failed_submit_count}")
        print(f"SENT:                 {sent_count}")
        print(f"UNDELIVERED:          {undelivered_count}")

        if skipped_cooldown_records:
            print(f"\n--- Skipped Recipients ({config.duplicate_cooldown_days}-Day Cooldown) ---")
            for scr in skipped_cooldown_records:
                print(f"  - {scr['recipient']}: Previous Submission: {scr['submitted_date']} | Eligible Again: {scr['eligible_date']}")

        if undelivered_review_records:
            print("\n--- Undelivered Review (Flagged for Manual Review) ---")
            for urr in undelivered_review_records:
                print(f"  - {urr['recipient']}: Status: {urr['status']}")

        if undelivered_records:
            print("\n--- Undelivered Recipients ---")
            for ur in undelivered_records:
                print(f"Recipient: {ur['recipient']}")
                print(f"Status: {ur['status']}")

        if execution_errors:
            print("\n--- Execution Error Details ---")
            for ee in execution_errors:
                print(f"  - {ee['recipient']}: {ee.get('error', 'Error')}")

        if config.logging_enabled:
            print(f"\nLogs:                 {config.log_file}")
        if config.excel_directory_enabled:
            print(f"Public Directory:     {config.excel_directory_file}")


if __name__ == '__main__':
    main()
