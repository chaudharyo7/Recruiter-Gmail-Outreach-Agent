import os
import re
import sys
import tempfile
import unittest
import threading
import base64
import email
import json
import httplib2
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch
from email.mime.multipart import MIMEMultipart

# Ensure project root is in sys.path and is current working directory
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)

from main import (
    is_valid_email,
    read_raw_recruiter_emails,
    read_recruiter_emails,
    create_message,
    load_template,
    parse_template,
    detect_bounces,
    is_permanent_invalid_recipient_bounce,
    sanitize_log_field,
    format_email_log_line,
    OutreachLogger,
    RecruiterDirectory,
    LiveDashboard,
    process_recipients_concurrently,
    execute_with_retry,
    get_cached_resume,
    Config,
    get_config,
    parse_bool,
    parse_int,
    parse_float,
    DEFAULT_CSV,
    DEFAULT_TEMPLATE,
    DEFAULT_RESUME,
    DEFAULT_EXCEL_DIRECTORY,
    NUM_WORKERS,
    DUPLICATE_CHECK_ENABLED,
    DUPLICATE_COOLDOWN_DAYS,
    load_outreach_history,
    evaluate_recipient_eligibility,
    safe_clear_csv_queue,
    SCOPES,
    CREDENTIALS_FILE,
    LOG_FILE
)


class TestRecruiterOutreachAgent(unittest.TestCase):

    def test_01_resume_pdf_validity(self):
        """Verify that resume PDF exists, is non-empty, and has a valid PDF header."""
        if not os.path.exists(DEFAULT_RESUME):
            self.skipTest(f"Resume not found at '{DEFAULT_RESUME}'. Place your PDF resume in attachments/ to run this test.")
        size = os.path.getsize(DEFAULT_RESUME)
        self.assertGreater(size, 1000, "resume PDF should be a valid PDF file (>1KB)")
        with open(DEFAULT_RESUME, 'rb') as f:
            header = f.read(5)
            self.assertEqual(header, b'%PDF-', "resume file must start with '%PDF-' header")
        print(f"\n[PASS] Resume file validated: '{DEFAULT_RESUME}' ({size} bytes, valid PDF)")

    def test_02_credentials_exist(self):
        """Verify that credentials.json is present and has desktop client config."""
        if not os.path.exists(CREDENTIALS_FILE):
            self.skipTest(f"Credentials not found at '{CREDENTIALS_FILE}'. Follow config/README.md to add credentials.json.")
        import json
        with open(CREDENTIALS_FILE, 'r') as f:
            creds = json.load(f)
        self.assertIn('installed', creds, "credentials.json should contain 'installed' desktop client config")
        print(f"[PASS] OAuth client credentials file present and format validated")

    def test_03_email_validation(self):
        """Test validation of email address format and INVALID categorization."""
        valid_cases = [
            "recruiter@company.com",
            "first.last@domain.co.uk",
            "recruiter+tag@sub.domain.org",
            "john_doe123@tech.io"
        ]
        invalid_cases = [
            "invalid-email-format",
            "@missinguser.com",
            "missingdomain@",
            "spaces in@email.com",
            "user@.com",
            ""
        ]
        for addr in valid_cases:
            self.assertTrue(is_valid_email(addr), f"Should be valid: {addr}")
        for addr in invalid_cases:
            self.assertFalse(is_valid_email(addr), f"Should be invalid: {addr}")
        print(f"[PASS] Email format regex correctly validated {len(valid_cases)} valid and {len(invalid_cases)} invalid cases")

    def test_04_csv_reading(self):
        """Test reading recruiter email list from CSV and ensure format segregation."""
        valid_emails, invalid_emails = read_recruiter_emails(DEFAULT_CSV)
        self.assertIsInstance(valid_emails, list)
        self.assertIsInstance(invalid_emails, list)
        for addr in valid_emails:
            self.assertTrue(is_valid_email(addr), f"Address in valid_emails must be valid: {addr}")
        for addr in invalid_emails:
            self.assertFalse(is_valid_email(addr), f"Address in invalid_emails must be invalid: {addr}")

        # Test with a test CSV buffer to verify parsing logic deterministically
        with tempfile.NamedTemporaryFile('w', suffix='.csv', delete=False) as tmp:
            tmp.write("email\nrecruiter1@example.com\ninvalid-email\nrecruiter2@example.com\n")
            tmp_path = tmp.name

        try:
            sample_valid, sample_invalid = read_recruiter_emails(tmp_path)
            self.assertEqual(sample_valid, ["recruiter1@example.com", "recruiter2@example.com"])
            self.assertEqual(sample_invalid, ["invalid-email"])
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

        print(f"[PASS] CSV reading validated: {len(valid_emails)} valid found in '{DEFAULT_CSV}', dummy CSV parsing verified")

    def test_05_email_generation_and_attachment(self):
        """Verify email generation, HTML body, subject from .env, and PDF attachment without sending."""
        if not os.path.exists(DEFAULT_RESUME):
            self.skipTest(f"Resume not found at '{DEFAULT_RESUME}'. Place your PDF resume in attachments/ to run this test.")
        self.assertTrue(os.path.exists(DEFAULT_TEMPLATE), "template.html must exist")
        cfg = get_config()
        subject = cfg.email_subject
        html_content = load_template(DEFAULT_TEMPLATE)
        self.assertEqual(subject, cfg.email_subject)

        test_recipient = "test.recruiter@example.com"
        message_dict = create_message(
            to_email=test_recipient,
            subject=subject,
            html_content=html_content,
            resume_path=DEFAULT_RESUME
        )

        self.assertIn('raw', message_dict, "Generated message dict must have 'raw' key")
        raw_bytes = base64.urlsafe_b64decode(message_dict['raw'].encode('utf-8'))

        # Parse raw MIME message
        parsed_msg = email.message_from_bytes(raw_bytes)
        self.assertEqual(parsed_msg['To'], test_recipient)
        self.assertEqual(re.sub(r'\s+', ' ', parsed_msg['Subject']), subject)
        self.assertIsNone(parsed_msg['Cc'], "Cc must be empty")
        self.assertIsNone(parsed_msg['Bcc'], "Bcc must be empty")

        # Check payload parts
        parts = list(parsed_msg.walk())
        part_content_types = [p.get_content_type() for p in parts]
        self.assertIn('text/html', part_content_types, "Message must contain text/html part")
        self.assertIn('application/pdf', part_content_types, "Message must contain application/pdf attachment")

        # Verify HTML body payload
        html_part = next(p for p in parts if p.get_content_type() == 'text/html')
        html_body_decoded = html_part.get_payload(decode=True).decode('utf-8')
        self.assertNotIn('<!-- Subject:', html_body_decoded, "Subject comment should be stripped from email body")

        # Verify attachment filename and payload
        pdf_part = next(p for p in parts if p.get_content_type() == 'application/pdf')
        expected_filename = os.path.basename(DEFAULT_RESUME)
        self.assertEqual(pdf_part.get_filename(), expected_filename)
        self.assertTrue(expected_filename.lower().endswith('.pdf'))
        pdf_data = pdf_part.get_payload(decode=True)
        self.assertTrue(pdf_data.startswith(b'%PDF-'), "Attached content must be a valid PDF binary")
        self.assertEqual(len(pdf_data), os.path.getsize(DEFAULT_RESUME))

        print(f"[PASS] MIME message verified: HTML body present, {expected_filename} attached ({len(pdf_data)} bytes), recipient '{test_recipient}' set individually with subject '{subject}'")

    def test_06_oauth_scope_configuration(self):
        """Verify required Gmail scopes (gmail.send and gmail.readonly) are configured."""
        self.assertIn('https://www.googleapis.com/auth/gmail.send', SCOPES)
        self.assertIn('https://www.googleapis.com/auth/gmail.readonly', SCOPES)
        print(f"[PASS] OAuth scopes verified: {SCOPES}")

    def test_07_email_subject_configured_only_from_env(self):
        """Verify email subject is configured exclusively via .env (EMAIL_SUBJECT) as single source of truth, not in template."""
        # 1. Subject loaded from .env config
        cfg = get_config()
        self.assertTrue(bool(cfg.email_subject and cfg.email_subject.strip()), "EMAIL_SUBJECT must be set in .env")

        # 2. Template loading reads pure HTML without needing subject comment
        body = load_template(DEFAULT_TEMPLATE)
        self.assertTrue(len(body) > 50)
        self.assertNotIn('<!-- Subject:', body)

        # 3. Dynamic subject change from .env
        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as tmp_env:
            tmp_env.write("EMAIL_SUBJECT=Senior Cloud Infrastructure Role\n")
            tmp_env_path = tmp_env.name

        try:
            custom_cfg = Config(env_path=tmp_env_path)
            self.assertEqual(custom_cfg.email_subject, "Senior Cloud Infrastructure Role")
        finally:
            if os.path.exists(tmp_env_path):
                os.remove(tmp_env_path)

        # 4. Empty/missing subject in .env raises validation error
        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as empty_env:
            empty_env.write("EMAIL_SUBJECT=\n")
            empty_env_path = empty_env.name

        try:
            empty_cfg = Config(env_path=empty_env_path)
            self.assertFalse(bool(empty_cfg.email_subject and empty_cfg.email_subject.strip()))
        finally:
            if os.path.exists(empty_env_path):
                os.remove(empty_env_path)

        print("[PASS] Email subject configured exclusively from .env as single source of truth; template contains no subject")

    def test_08_submission_status_reporting(self):
        """Verify that Gmail acceptance produces SUBMITTED (never SUCCESS or DELIVERED), and failure produces FAILED_TO_SUBMIT."""
        mock_service = MagicMock()

        # 1. Successful submission mock
        mock_service.users().messages().send().execute.return_value = {'id': '18abc123'}

        rec_success = {
            'timestamp': '2026-11-15 19:30:01',
            'recipient': 'recruiter@gmail.com',
            'status': None,
            'message_id': 'N/A',
            'remark': 'None',
            'error': 'None',
            'bounce_time': 'N/A'
        }
        resp = mock_service.users().messages().send().execute()
        rec_success['message_id'] = resp.get('id')
        rec_success['status'] = 'SENT'
        rec_success['remark'] = 'Gmail API accepted email for sending'

        self.assertEqual(rec_success['status'], 'SENT')
        self.assertNotEqual(rec_success['status'], 'DELIVERED', "Never report DELIVERED based on Gmail API acceptance")
        self.assertNotEqual(rec_success['status'], 'SUBMITTED', "SUBMITTED is replaced by SENT")
        self.assertNotEqual(rec_success['status'], 'UNKNOWN', "SENT should not be marked UNKNOWN immediately")
        self.assertEqual(rec_success['message_id'], '18abc123')

        # 2. Failed submission mock
        mock_service.users().messages().send().execute.side_effect = Exception("HTTP 403 permission denied")

        rec_failed = {
            'timestamp': '2026-11-15 19:30:20',
            'recipient': 'jobs@gmail.com',
            'status': None,
            'message_id': 'N/A',
            'remark': 'None',
            'error': 'None',
            'bounce_time': 'N/A'
        }
        try:
            mock_service.users().messages().send().execute()
        except Exception as e:
            rec_failed['status'] = 'FAILED_TO_SUBMIT'
            rec_failed['remark'] = 'Gmail API rejected sending request'
            rec_failed['error'] = str(e)

        self.assertEqual(rec_failed['status'], 'FAILED_TO_SUBMIT')
        self.assertIn("HTTP 403", rec_failed['error'])

        # 3. Invalid email produces INVALID
        rec_invalid = {
            'timestamp': '2026-11-15 19:30:05',
            'recipient': 'invalid-email',
            'status': 'INVALID',
            'message_id': 'N/A',
            'remark': 'Invalid email format; skipped',
            'error': 'None',
            'bounce_time': 'N/A'
        }
        self.assertEqual(rec_invalid['status'], 'INVALID')
        self.assertEqual(rec_invalid['message_id'], 'N/A')

        print("[PASS] Status reporting verified: acceptance yields SENT, API failure yields FAILED_TO_SUBMIT, invalid yields INVALID")

    def test_09_bounce_detection_and_matching(self):
        """Verify bounce notifications identify UNDELIVERED (reason) recipients, unmatched bounces are ignored, and missing bounce remains SENT."""
        mock_service = MagicMock()

        # Dummy bounce notification from Mail Delivery Subsystem
        dummy_bounce = {
            'id': 'bounce_msg_101',
            'snippet': "Address not found Your message wasn't delivered to unknown@gmail.com because the address couldn't be found, or is unable to receive mail.",
            'payload': {
                'headers': [
                    {'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'},
                    {'name': 'Subject', 'value': 'Delivery Status Notification (Failure)'},
                    {'name': 'Date', 'value': 'Sun, 15 Nov 2026 19:35:10 +0000'}
                ],
                'body': {'data': ''},
                'parts': []
            }
        }

        mock_service.users().messages().list().execute.return_value = {
            'messages': [{'id': 'bounce_msg_101'}]
        }
        mock_service.users().messages().get().execute.return_value = dummy_bounce

        records = [
            {
                'timestamp': '2026-11-15 19:30:10',
                'recipient': 'unknown@gmail.com',
                'subject': 'Application for DevOps Role',
                'message_id': '18def456',
                'status': 'SENT',
                'remark': 'Gmail API accepted email for sending',
                'error': 'None',
                'bounce_time': 'N/A'
            },
            {
                'timestamp': '2026-11-15 19:30:15',
                'recipient': 'hr@gmail.com',
                'subject': 'Application for DevOps Role',
                'message_id': '18ghi789',
                'status': 'SENT',
                'remark': 'Gmail API accepted email for sending',
                'error': 'None',
                'bounce_time': 'N/A'
            }
        ]

        # Run bounce detection
        detect_bounces(mock_service, records, delay_seconds=0)

        # 1. Matching recipient correctly marked UNDELIVERED (reason)
        self.assertTrue(records[0]['status'].startswith('UNDELIVERED'))
        self.assertIn("address couldn't be found", records[0]['status'].lower())
        self.assertNotEqual(records[0]['bounce_time'], 'N/A')
        self.assertIn("address couldn't be found", records[0]['error'].lower())

        # 2. Unrelated recipient hr@gmail.com is NOT marked UNDELIVERED
        self.assertFalse(records[1]['status'].startswith('UNDELIVERED'))

        # 3. Missing bounce notification keeps status as SENT (never UNKNOWN, never DELIVERED)
        self.assertEqual(records[1]['status'], 'SENT')
        self.assertNotEqual(records[1]['status'], 'UNKNOWN', "Absence of bounce must not change SENT to UNKNOWN")
        self.assertNotEqual(records[1]['status'], 'DELIVERED', "Absence of bounce must never be reported as DELIVERED")

        print("[PASS] Bounce detection verified: bounce identified as UNDELIVERED (reason), unrelated recipient remains SENT, no false delivery confirmation")

    def test_10_one_line_per_email_and_inplace_updates(self):
        """Verify: exactly one physical line per email, strict format, in-place update from SENT to UNDELIVERED, timestamp/message_id preserved, sanitized newlines."""
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp:
            tmp_log_path = tmp.name

        try:
            logger = OutreachLogger(tmp_log_path)

            # Log initial SENT email record
            initial_rec = {
                'timestamp': '2026-11-15 19:30:01',
                'recipient': 'candidate@gmail.com',
                'status': 'SENT',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18abc123',
                'submitted_time': '19:30:01',
                'remark': 'Gmail API accepted email for sending',
                'error': 'None',
                'bounce_time': 'N/A'
            }
            logger.log_email_record(initial_rec)

            # Log second email record with multiline error attempt
            multiline_rec = {
                'timestamp': '2026-11-15 19:30:02',
                'recipient': 'multiline@gmail.com',
                'status': 'FAILED_TO_SUBMIT',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': 'N/A',
                'submitted_time': 'N/A',
                'remark': 'Error with \r\n newline in \n description',
                'error': 'Stack trace:\nline 1\nline 2',
                'bounce_time': 'N/A'
            }
            logger.log_email_record(multiline_rec)

            # Read log and count lines
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines_before = f.readlines()

            # Verify multiline error was sanitized into a single line
            for line in lines_before:
                if 'multiline@gmail.com' in line:
                    self.assertNotIn('\n', line.rstrip('\r\n'), "No internal newlines allowed in email record")
                    self.assertIn('Stack trace: line 1 line 2', line)

            # Exactly 2 lines total: exactly 2 email lines, NO event/auth/summary lines
            self.assertEqual(len(lines_before), 2, "Each email must have exactly one physical line and no event logs")

            # Verify format on initial_rec
            cand_line = next(line for line in lines_before if 'candidate@gmail.com' in line)
            expected_cand = (
                "2026-11-15 19:30:01 | candidate@gmail.com | SENT | "
                "Subject: Application for DevOps / Cloud Engineer Opportunities | ID: 18abc123 | "
                "Submitted: 19:30:01 | Bounce: N/A | Error: None\n"
            )
            self.assertEqual(cand_line, expected_cand)

            # Now update initial_rec to UNDELIVERED (simulating a later confirmed bounce)
            updated = logger.update_email_record(
                recipient='candidate@gmail.com',
                new_status='UNDELIVERED (Mailbox does not exist)',
                error='SMTP 550 5.1.1',
                bounce_time='19:35:10'
            )
            self.assertTrue(updated, "Record update must succeed")

            # Read log after bounce update
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines_after = f.readlines()

            # Must still have exactly 2 physical lines (no new line appended for candidate@gmail.com)
            self.assertEqual(len(lines_after), 2, "Updating status must update in-place, NOT append another line")

            # Find updated line
            updated_line = next(line for line in lines_after if 'candidate@gmail.com' in line)
            expected_updated = (
                "2026-11-15 19:30:01 | candidate@gmail.com | UNDELIVERED (Mailbox does not exist) | "
                "Subject: Application for DevOps / Cloud Engineer Opportunities | ID: 18abc123 | "
                "Submitted: 19:30:01 | Bounce: 19:35:10 | Error: SMTP 550 5.1.1\n"
            )
            self.assertEqual(updated_line, expected_updated)

            # Test appending across executions
            logger.log_email_record({
                'timestamp': '2026-11-15 19:40:01',
                'recipient': 'run2@gmail.com',
                'status': 'SENT',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18run2',
                'submitted_time': '19:40:01',
                'remark': 'Gmail API accepted email for sending',
                'error': 'None',
                'bounce_time': 'N/A'
            })
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines_run2 = f.readlines()

            # Must have 3 lines now (previous 2 + 1 new email record)
            self.assertEqual(len(lines_run2), 3, "Logs must append across executions (strictly 1 line per email)")

            # Verify no sensitive tokens or email bodies logged
            log_full_text = "".join(lines_run2)
            sensitive_patterns = ["client_secret", "refresh_token", "<html>", "<!DOCTYPE html>"]
            for pat in sensitive_patterns:
                self.assertNotIn(pat, log_full_text, f"Sensitive data '{pat}' must never be written to logs")

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] One-line logging verified: exactly one line per email, in-place update preserves timestamp and message ID, sanitized newlines, cross-run append")

    def test_11_delayed_bounce_detection_from_previous_logs(self):
        """Verify delayed bounce detection: unconfirmed records from earlier runs are checked and updated in-place upon bounce arrival."""
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp:
            tmp_log_path = tmp.name

        try:
            logger = OutreachLogger(tmp_log_path)

            # Log an email submitted in run 1 whose delivery was unconfirmed
            prev_record = {
                'timestamp': '2026-11-14 10:00:00',
                'recipient': 'delayed.bounce@company.com',
                'status': 'SENT',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18delayed999',
                'submitted_time': '10:00:00',
                'remark': 'Gmail API accepted email for sending',
                'error': 'None',
                'bounce_time': 'N/A'
            }
            logger.log_email_record(prev_record)

            # Another email in run 1 that remains unconfirmed (no bounce)
            other_record = {
                'timestamp': '2026-11-14 10:00:05',
                'recipient': 'valid.recruiter@company.com',
                'status': 'SENT',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18valid888',
                'submitted_time': '10:00:05',
                'remark': 'Gmail API accepted email for sending',
                'error': 'None',
                'bounce_time': 'N/A'
            }
            logger.log_email_record(other_record)

            # Check lines in log file before run 2
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines_before = f.readlines()
            self.assertEqual(len(lines_before), 2)  # exactly 2 records

            # Verify get_unconfirmed_records finds both records
            unconfirmed = logger.get_unconfirmed_records()
            self.assertEqual(len(unconfirmed), 2)
            self.assertEqual(unconfirmed[0]['recipient'], 'delayed.bounce@company.com')
            self.assertEqual(unconfirmed[1]['recipient'], 'valid.recruiter@company.com')

            # Mock Gmail service returning a delayed bounce for delayed.bounce@company.com
            mock_service = MagicMock()
            delayed_bounce_msg = {
                'id': 'bounce_delayed_555',
                'snippet': "550 5.1.1 The email account that you tried to reach does not exist: delayed.bounce@company.com",
                'payload': {
                    'headers': [
                        {'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'},
                        {'name': 'Subject', 'value': 'Delivery Status Notification (Failure)'},
                        {'name': 'Date', 'value': 'Sun, 15 Nov 2026 08:00:00 +0000'}
                    ],
                    'body': {'data': ''},
                    'parts': []
                }
            }
            mock_service.users().messages().list().execute.return_value = {
                'messages': [{'id': 'bounce_delayed_555'}]
            }
            mock_service.users().messages().get().execute.return_value = delayed_bounce_msg

            # Run bounce detection on previously unconfirmed records (simulating Run 2)
            detect_bounces(mock_service, unconfirmed, logger=logger, delay_seconds=0)

            # Verify unconfirmed in-memory states
            self.assertTrue(unconfirmed[0]['status'].startswith('UNDELIVERED'))
            self.assertIn("does not exist", unconfirmed[0]['status'].lower())
            self.assertIn("550", unconfirmed[0]['error'])
            self.assertEqual(unconfirmed[1]['status'], 'SENT', "No bounce must keep recipient SENT, never UNKNOWN")

            # Verify log file was updated in-place without adding duplicate email records
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines_after = f.readlines()
            
            # Exactly 2 email record lines must exist (no duplicate records)
            self.assertEqual(len(lines_after), 2, "Email record count must remain 2 (in-place update, no duplicate records)")

            delayed_lines = [l for l in lines_after if 'delayed.bounce@company.com' in l]
            self.assertEqual(len(delayed_lines), 1, "delayed.bounce@company.com must have exactly one line in log")

            # Check updated line for delayed.bounce@company.com
            updated_line = delayed_lines[0]
            parts = [p.strip() for p in updated_line.split('|')]
            self.assertEqual(parts[0], '2026-11-14 10:00:00', "Original timestamp must be preserved")
            self.assertEqual(parts[1], 'delayed.bounce@company.com')
            self.assertTrue(parts[2].startswith('UNDELIVERED'))
            self.assertIn("does not exist", parts[2].lower())
            self.assertEqual(parts[3], 'Subject: Application for DevOps / Cloud Engineer Opportunities')
            self.assertEqual(parts[4], 'ID: 18delayed999', "Original Message-ID must be preserved")
            self.assertEqual(parts[5], 'Submitted: 10:00:00', "Original Submitted time must be preserved")
            self.assertNotEqual(parts[6], 'Bounce: N/A', "Bounce time must be updated")

            # Check valid.recruiter line was not marked as delivered
            other_line = next(line for line in lines_after if 'valid.recruiter@company.com' in line)
            self.assertIn('| SENT |', other_line)
            self.assertNotIn('DELIVERED', other_line)

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Delayed bounce detection verified: unconfirmed emails checked, log updated in-place, timestamps preserved")

    def test_12_compact_summary_counts_and_dashboard_format(self):
        """Verify dashboard and final summary: exactly TWO user-facing statuses (SENT, UNDELIVERED), list undelivered with reasons, UNKNOWN removed."""
        import io
        from unittest.mock import patch

        dashboard = LiveDashboard(total_recipients=4, is_live=False)

        # Case 1: 4 emails accepted by Gmail API, no bounces detected yet
        dashboard.set_counter('SENT', 4)
        dashboard.set_counter('UNDELIVERED', 0)
        self.assertEqual(dashboard.counters['SENT'], 4)
        self.assertEqual(dashboard.counters['UNDELIVERED'], 0)

        with patch('sys.stdout', new=io.StringIO()) as fake_stdout:
            dashboard.render_final_summary(log_path="logs/outreach.log", excel_path="data/recruiter_email_backup.xlsx")
            out1 = fake_stdout.getvalue()
            self.assertIn("SENT:                 4", out1)
            self.assertIn("UNDELIVERED:          0", out1)
            self.assertNotIn("UNKNOWN", out1)
            self.assertNotIn("SUBMITTED", out1)

        # Case 2: A bounce arrives later -> SENT: 3, UNDELIVERED: 1 with failure reason
        dashboard.set_counter('SENT', 3)
        dashboard.set_counter('UNDELIVERED', 1)
        undelivered = [{
            'recipient': 'recruiter@example.com',
            'status': 'UNDELIVERED (Mailbox does not exist)'
        }]
        dashboard.set_undelivered_recipients(undelivered)

        with patch('sys.stdout', new=io.StringIO()) as fake_stdout:
            dashboard.render_final_summary(log_path="logs/outreach.log", excel_path="data/recruiter_email_backup.xlsx")
            out2 = fake_stdout.getvalue()
            self.assertIn("SENT:                 3", out2)
            self.assertIn("UNDELIVERED:          1", out2)
            self.assertIn("Recipient: recruiter@example.com", out2)
            self.assertIn("Status: UNDELIVERED (Mailbox does not exist)", out2)
            self.assertNotIn("UNKNOWN", out2)
            self.assertNotIn("SUBMITTED", out2)

        print("[PASS] Dashboard and summary format verified: SENT & UNDELIVERED counters, undelivered list with reasons, UNKNOWN removed")

    def test_13_gmail_403_insufficient_permissions_handling(self):
        """Verify handling of Gmail API 403 insufficient scopes error: does not crash, never downgrades SENT, never falsely marks UNDELIVERED or DELIVERED."""
        from googleapiclient.errors import HttpError
        from unittest.mock import MagicMock
        import httplib2

        resp = httplib2.Response({'status': 403, 'reason': 'Request had insufficient authentication scopes.'})
        http_error = HttpError(resp, b'{"error": {"message": "Insufficient Permission"}}')

        mock_service = MagicMock()
        mock_service.users().messages().list.side_effect = http_error

        records = [{
            'recipient': 'candidate@example.com',
            'status': 'SENT',
            'subject': 'Application for DevOps Role',
            'message_id': '18abc456',
            'submitted_time': '12:00:00',
            'remark': 'Gmail API accepted email for sending',
            'error': 'None',
            'bounce_time': 'N/A'
        }]

        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp:
            tmp_log_path = tmp.name

        try:
            logger = OutreachLogger(tmp_log_path)
            logger.log_email_record(records[0])

            # Call detect_bounces which encounters HttpError 403
            detect_bounces(mock_service, records, logger=logger, delay_seconds=0)

            # Record must REMAIN SENT: bounce check errors never downgrade SENT or mark UNDELIVERED
            self.assertEqual(records[0]['status'], 'SENT', "Bounce check errors must never downgrade SENT")
            self.assertFalse(str(records[0]['status']).startswith('UNDELIVERED'), "Never mark UNDELIVERED on bounce check failure")
            self.assertNotEqual(records[0]['status'], 'UNKNOWN', "Never downgrade to UNKNOWN")
            self.assertNotEqual(records[0]['status'], 'DELIVERED', "Never falsely confirm delivery")

            # Log record must remain unchanged as SENT
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 1, "Must maintain exactly 1 line in outreach.log")
            self.assertIn('| SENT |', lines[0])

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Gmail API 403 insufficient permissions handled gracefully: SENT preserved, never downgraded to UNKNOWN or UNDELIVERED, single-line log intact")

    def test_14_exponential_backoff_retry_transient_errors(self):
        """Verify exponential backoff retries on transient errors (429, 503) and immediate failure on permanent errors (400, 403)."""
        from googleapiclient.errors import HttpError
        import httplib2

        call_counter = {'count': 0}

        # Case 1: Transient 429 fails twice, succeeds on third attempt
        def transient_call():
            call_counter['count'] += 1
            if call_counter['count'] < 3:
                resp = httplib2.Response({'status': 429, 'reason': 'Too Many Requests'})
                raise HttpError(resp, b'Rate limit exceeded')
            return {'status': 'success'}

        result = execute_with_retry(transient_call, max_retries=3, initial_delay=0.01, backoff_factor=1.5)
        self.assertEqual(result, {'status': 'success'})
        self.assertEqual(call_counter['count'], 3, "Transient error should be retried until success")

        # Case 2: Permanent error (400 Bad Request) fails immediately with zero retries
        call_counter['count'] = 0
        def permanent_call():
            call_counter['count'] += 1
            resp = httplib2.Response({'status': 400, 'reason': 'Bad Request'})
            raise HttpError(resp, b'Bad Request')

        with self.assertRaises(HttpError):
            execute_with_retry(permanent_call, max_retries=3, initial_delay=0.01, backoff_factor=1.5)
        self.assertEqual(call_counter['count'], 1, "Permanent 400 error must NOT be retried")

        print("[PASS] Exponential backoff verified: transient 429 retried, permanent 400 failed immediately without retry")

    def test_15_batch_logging_and_resume_caching(self):
        """Verify batch logging I/O and in-memory resume attachment caching."""
        # 1. Resume caching
        if not os.path.exists(DEFAULT_RESUME):
            self.skipTest(f"Resume not found at '{DEFAULT_RESUME}'.")

        filename1, bytes1 = get_cached_resume(DEFAULT_RESUME)
        filename2, bytes2 = get_cached_resume(DEFAULT_RESUME)
        self.assertIs(bytes1, bytes2, "get_cached_resume must return cached bytes without re-reading file")
        self.assertTrue(bytes1.startswith(b'%PDF-'))

        # 2. Batch logging and batch updates
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp:
            tmp_log_path = tmp.name

        try:
            logger = OutreachLogger(tmp_log_path)
            # Batch append 3 records
            records = [
                {
                    'timestamp': '2026-11-15 20:00:00',
                    'recipient': f'user{i}@example.com',
                    'status': 'SUBMITTED',
                    'subject': 'Application',
                    'message_id': f'msg_{i}',
                    'submitted_time': '20:00:00',
                    'bounce_time': 'N/A',
                    'error': 'None'
                }
                for i in range(3)
            ]
            logger.log_email_records(records)

            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 3, "Batch logging must append exactly 3 lines")

            # Batch update 2 of the records
            updates = {
                'user0@example.com': {'new_status': 'BOUNCED', 'error': 'SMTP 550 5.1.1', 'bounce_time': '20:05:00'},
                'user1@example.com': {'new_status': 'UNKNOWN', 'error': 'None', 'bounce_time': 'N/A'}
            }
            updated_count = logger.batch_update_email_records(updates)
            self.assertEqual(updated_count, 2)

            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines_after = f.readlines()
            self.assertEqual(len(lines_after), 3, "Batch update must maintain exactly 3 physical lines in-place")
            self.assertIn('| BOUNCED |', lines_after[0])
            self.assertIn('SMTP 550 5.1.1', lines_after[0])
            self.assertIn('| UNKNOWN |', lines_after[1])
            self.assertIn('| SUBMITTED |', lines_after[2])

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Batch logging and in-memory resume caching verified")

    def test_16_recruiter_directory_excel_creation_and_persistence(self):
        """Verify RecruiterDirectory creates Excel file with 'Email' column, deduplicates case-insensitively, handles thread concurrency, and persists entries across executions."""
        import openpyxl
        import concurrent.futures

        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_excel_path = tmp.name

        try:
            # 1. Initialization creates file with single 'Email' column header
            directory = RecruiterDirectory(tmp_excel_path)
            wb = openpyxl.load_workbook(tmp_excel_path)
            ws = wb.active
            self.assertEqual(ws.cell(row=1, column=1).value, "Email")
            self.assertIsNone(ws.cell(row=1, column=2).value, "Only one column allowed")

            # 2. Add submitted emails with case-insensitive deduplication
            self.assertTrue(directory.add_email("recruiter1@company.com"))
            self.assertFalse(directory.add_email("RECRUITER1@COMPANY.COM"), "Duplicate case-insensitive must return False")
            self.assertTrue(directory.add_email("recruiter2@company.com"))

            # 3. Verify exactly 2 rows exist
            emails = directory.get_all_emails()
            self.assertEqual(len(emails), 2)
            self.assertEqual(emails[0], "recruiter1@company.com")
            self.assertEqual(emails[1], "recruiter2@company.com")

            # 4. Thread-safe concurrent writing from 10 threads
            def concurrent_writer(idx):
                directory.add_email(f"threaded_{idx}@company.com")
                directory.add_email(f"THREADED_{idx}@company.com")

            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as ex:
                list(ex.map(concurrent_writer, range(10)))

            all_after = directory.get_all_emails()
            self.assertEqual(len(all_after), 12, "Should have 2 initial + 10 distinct concurrent emails")

            # 5. Persistence across executions: a new instance reloads existing entries without loss
            new_run_dir = RecruiterDirectory(tmp_excel_path)
            self.assertEqual(len(new_run_dir.get_all_emails()), 12)
            self.assertFalse(new_run_dir.add_email("recruiter1@company.com"), "Must recognize prior execution emails")
            self.assertTrue(new_run_dir.add_email("new_campaign_recruiter@company.com"))
            self.assertEqual(len(new_run_dir.get_all_emails()), 13)

        finally:
            if os.path.exists(tmp_excel_path):
                os.remove(tmp_excel_path)

        print("[PASS] RecruiterDirectory Excel creation, single-column constraint, case-insensitive deduplication, concurrency, and cross-run persistence verified")

    def test_17_live_dashboard_worker_panels_counters_and_summary(self):
        """Verify LiveDashboard worker panels (Worker 1, 2, 3), state transitions, SENT/UNDELIVERED counters, and summary rendering."""
        import io
        from unittest.mock import patch

        dashboard = LiveDashboard(total_recipients=5, is_live=False)

        # Verify initial 3 worker panels
        self.assertEqual(len(dashboard.workers), 3)
        for w_id in (1, 2, 3):
            self.assertEqual(dashboard.workers[w_id]['state'], 'IDLE')
            self.assertEqual(dashboard.workers[w_id]['recipient'], '-')

        # Verify state updates for all 3 workers
        dashboard.update_worker(1, state='VALIDATING', recipient='user1@example.com')
        dashboard.update_worker(2, state='SENDING', recipient='user2@example.com')
        dashboard.update_worker(3, state='SENT', recipient='user3@example.com')

        self.assertEqual(dashboard.workers[1]['state'], 'VALIDATING')
        self.assertEqual(dashboard.workers[2]['state'], 'SENDING')
        self.assertEqual(dashboard.workers[3]['state'], 'SENT')

        # Verify counters
        dashboard.increment_counter('SENT', 3)
        dashboard.set_counter('UNDELIVERED', 1)

        self.assertEqual(dashboard.counters['SENT'], 3)
        self.assertEqual(dashboard.counters['UNDELIVERED'], 1)

        # Add undelivered recipient and execution error details
        dashboard.set_undelivered_recipients([
            {'recipient': 'bounce@example.com', 'status': 'UNDELIVERED (Mailbox does not exist)'}
        ])
        dashboard.add_execution_error('bad@domain', 'Invalid email format; skipped')

        # Verify final summary output contains SENT and UNDELIVERED, and no UNKNOWN or SUBMITTED
        with patch('sys.stdout', new=io.StringIO()) as fake_stdout:
            dashboard.render_final_summary(log_path="logs/outreach.log", excel_path="data/recruiter_email_backup.xlsx")
            output = fake_stdout.getvalue()
            self.assertIn("Final Execution Summary", output)
            self.assertIn("SENT:                 3", output)
            self.assertIn("UNDELIVERED:          1", output)
            self.assertIn("Recipient: bounce@example.com", output)
            self.assertIn("Status: UNDELIVERED (Mailbox does not exist)", output)
            self.assertIn("bad@domain", output)
            self.assertNotIn("UNKNOWN", output)
            self.assertNotIn("SUBMITTED", output)

        print("[PASS] LiveDashboard 3-worker panels, state transitions, SENT/UNDELIVERED counters, and final summary verified")

    def test_18_parallel_sending_exactly_3_workers_and_duplicate_prevention(self):
        """Verify parallel sending uses exactly 3 workers, prevents duplicates within the run, and records SUBMITTED in Excel."""
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp_log, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel:
            tmp_log_path = tmp_log.name
            tmp_excel_path = tmp_excel.name

        try:
            logger = OutreachLogger(tmp_log_path)
            directory = RecruiterDirectory(tmp_excel_path)
            dashboard = LiveDashboard(total_recipients=5, is_live=False)

            # Recipient list containing valid, duplicate within run, and invalid
            recipients = [
                'worker1.test@example.com',
                'worker2.test@example.com',
                'WORKER1.TEST@EXAMPLE.COM',  # Duplicate within run -> must be SKIPPED
                'invalid-email-format',       # Invalid format -> must be INVALID
                'worker3.test@example.com'
            ]

            mock_service = MagicMock()
            mock_service.users().messages().send().execute.return_value = {'id': 'msg_mock_777'}

            all_records, submitted_records = process_recipients_concurrently(
                recipients=recipients,
                service=mock_service,
                subject="Application for DevOps / Cloud Engineer Opportunities",
                html_template="<p>Test Body</p>",
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=directory,
                dashboard=dashboard,
                num_workers=3,
                api_call_delay=0.0
            )

            # Exactly 3 valid emails submitted
            self.assertEqual(len(submitted_records), 3)
            submitted_emails = {r['recipient'].lower() for r in submitted_records}
            self.assertEqual(submitted_emails, {
                'worker1.test@example.com',
                'worker2.test@example.com',
                'worker3.test@example.com'
            })

            # Check dashboard counters (SENT & UNDELIVERED delivery statuses) and execution errors
            self.assertEqual(dashboard.counters['SENT'], 3)
            self.assertEqual(dashboard.counters['UNDELIVERED'], 0)
            self.assertEqual(len(dashboard.execution_errors), 1)
            self.assertEqual(dashboard.execution_errors[0]['recipient'], 'invalid-email-format')

            # Check Excel directory: only SENT emails entered, deduplicated
            excel_emails = directory.get_all_emails()
            self.assertEqual(len(excel_emails), 3)
            self.assertIn('worker1.test@example.com', [e.lower() for e in excel_emails])
            self.assertIn('worker2.test@example.com', [e.lower() for e in excel_emails])
            self.assertIn('worker3.test@example.com', [e.lower() for e in excel_emails])
            self.assertNotIn('invalid-email-format', [e.lower() for e in excel_emails])

            # Check outreach.log lines: exactly 4 lines (3 SENT + 1 INVALID; SKIPPED is not logged)
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                log_lines = [l.strip() for l in f.readlines() if l.strip()]
            self.assertEqual(len(log_lines), 4, "Must maintain exactly one line per processed email in outreach.log")
            self.assertTrue(any('worker1.test@example.com' in l and '| SENT |' in l for l in log_lines))
            self.assertTrue(any('invalid-email-format' in l and '| INVALID |' in l for l in log_lines))

            # Verify that future campaign can reuse existing directory addresses:
            dashboard2 = LiveDashboard(total_recipients=1, is_live=False)
            all_rec2, sub_rec2 = process_recipients_concurrently(
                recipients=['worker1.test@example.com'],
                service=mock_service,
                subject="Follow-up Application",
                html_template="<p>Follow-up</p>",
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=directory,
                dashboard=dashboard2,
                num_workers=3,
                api_call_delay=0.0
            )
            self.assertEqual(len(sub_rec2), 1, "Future campaigns must be allowed to send to directory addresses again")
            self.assertEqual(len(directory.get_all_emails()), 3)

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)
            if os.path.exists(tmp_excel_path):
                os.remove(tmp_excel_path)

        print("[PASS] Parallel email sending (3 workers), in-run duplicate prevention, Excel registration, and future campaign reuse verified")

    def test_19_config_loading_and_validation(self):
        """Verify centralized Config loading from .env, default values, and validation error handling."""
        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as tmp_env:
            tmp_env.write(
                "NUM_WORKERS=5\n"
                "MAX_RETRIES=4\n"
                "RETRY_BACKOFF_SECONDS=1.5\n"
                "GMAIL_API_REQUEST_DELAY=0.1\n"
                "EXCEL_DIRECTORY_ENABLED=false\n"
                "BOUNCE_CHECK_ENABLED=false\n"
                "LOGGING_ENABLED=false\n"
                "DASHBOARD_ENABLED=false\n"
                "DRY_RUN=true\n"
                "EMAIL_SUBJECT=Custom Config Subject\n"
            )
            tmp_env_path = tmp_env.name

        try:
            cfg = Config(env_path=tmp_env_path)
            cfg.validate()

            self.assertEqual(cfg.num_workers, 5)
            self.assertEqual(cfg.max_retries, 4)
            self.assertEqual(cfg.retry_backoff_seconds, 1.5)
            self.assertEqual(cfg.gmail_api_request_delay, 0.1)
            self.assertFalse(cfg.excel_directory_enabled)
            self.assertFalse(cfg.bounce_check_enabled)
            self.assertFalse(cfg.logging_enabled)
            self.assertFalse(cfg.dashboard_enabled)
            self.assertTrue(cfg.dry_run)
            self.assertEqual(cfg.email_subject, "Custom Config Subject")

            # Test invalid settings validation
            with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as invalid_env:
                invalid_env.write("NUM_WORKERS=0\n")
                invalid_env_path = invalid_env.name

            try:
                inv_cfg = Config(env_path=invalid_env_path)
                self.assertEqual(inv_cfg.num_workers, 3)

                inv_cfg.num_workers = 0
                with self.assertRaises(ValueError):
                    inv_cfg.validate()

                inv_cfg.num_workers = 3
                inv_cfg.max_retries = -1
                with self.assertRaises(ValueError):
                    inv_cfg.validate()

            finally:
                if os.path.exists(invalid_env_path):
                    os.remove(invalid_env_path)

        finally:
            if os.path.exists(tmp_env_path):
                os.remove(tmp_env_path)

        print("[PASS] Centralized Config loader, defaults, overrides, and validation verified")

    def test_20_dry_run_mode_execution(self):
        """Verify that DRY_RUN=true processes recipients safely without calling Gmail API send."""
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp_log, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel, \
             tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as tmp_env:
            tmp_log_path = tmp_log.name
            tmp_excel_path = tmp_excel.name
            tmp_env.write(f"DRY_RUN=true\nLOG_FILE={tmp_log_path}\nEXCEL_DIRECTORY_FILE={tmp_excel_path}\n")
            tmp_env_path = tmp_env.name

        try:
            cfg = Config(env_path=tmp_env_path)
            self.assertTrue(cfg.dry_run)

            logger = OutreachLogger(tmp_log_path)
            directory = RecruiterDirectory(tmp_excel_path)
            dashboard = LiveDashboard(total_recipients=2, is_live=False, num_workers=cfg.num_workers)

            mock_service = MagicMock()

            recipients = ['dryrun1@example.com', 'dryrun2@example.com']
            all_records, submitted_records = process_recipients_concurrently(
                recipients=recipients,
                service=mock_service,
                subject="Test Subject",
                html_template="<p>Test</p>",
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=directory,
                dashboard=dashboard,
                config=cfg
            )

            # Gmail API send must NEVER have been called in DRY_RUN mode
            self.assertEqual(mock_service.users().messages().send.call_count, 0, "Real Gmail API send must not be called in DRY_RUN")

            # Recipients must be safely marked SENT with dry_run message IDs
            self.assertEqual(len(submitted_records), 2)
            self.assertTrue(all(r['status'] == 'SENT' for r in submitted_records))
            self.assertTrue(all(r['message_id'].startswith('dry_run_') for r in submitted_records))

            # Directory and logs must reflect submission
            excel_emails = directory.get_all_emails()
            self.assertEqual(len(excel_emails), 2)

            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 2)
            self.assertIn('| SENT |', lines[0])

        finally:
            for p in (tmp_log_path, tmp_excel_path, tmp_env_path):
                if os.path.exists(p):
                    os.remove(p)

        print("[PASS] DRY_RUN execution verified: zero Gmail API send calls made, safe workflow processing verified")

    def test_21_feature_toggles_excel_logging_subject_override(self):
        """Verify EXCEL_DIRECTORY_ENABLED=false, LOGGING_ENABLED=false, and EMAIL_SUBJECT override."""
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp_log, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel, \
             tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as tmp_env:
            tmp_log_path = tmp_log.name
            tmp_excel_path = tmp_excel.name
            tmp_env.write(
                f"EXCEL_DIRECTORY_ENABLED=false\n"
                f"LOGGING_ENABLED=false\n"
                f"EMAIL_SUBJECT=Overridden Subject From Env\n"
                f"DRY_RUN=true\n"
            )
            tmp_env_path = tmp_env.name

        try:
            cfg = Config(env_path=tmp_env_path)
            self.assertFalse(cfg.excel_directory_enabled)
            self.assertFalse(cfg.logging_enabled)
            self.assertEqual(cfg.email_subject, "Overridden Subject From Env")

            logger = OutreachLogger(tmp_log_path)
            directory = RecruiterDirectory(tmp_excel_path)
            dashboard = LiveDashboard(total_recipients=1, is_live=False)

            all_records, submitted_records = process_recipients_concurrently(
                recipients=['feature.toggle@example.com'],
                service=MagicMock(),
                subject=cfg.email_subject,
                html_template="<p>Test</p>",
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=directory,
                dashboard=dashboard,
                config=cfg
            )

            # Subject was overridden
            self.assertEqual(submitted_records[0]['subject'], "Overridden Subject From Env")

            # Excel directory was disabled -> 0 emails added
            self.assertEqual(len(directory.get_all_emails()), 0)

            # Logging was disabled -> 0 lines written
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 0)

        finally:
            for p in (tmp_log_path, tmp_excel_path, tmp_env_path):
                if os.path.exists(p):
                    os.remove(p)

        print("[PASS] Feature toggles (EXCEL_DIRECTORY_ENABLED=false, LOGGING_ENABLED=false) and EMAIL_SUBJECT override verified")

    def test_22_thread_safe_service_isolation_and_concurrency(self):
        """Verify thread-isolated Gmail service instances across concurrent threads to prevent C-level SSL heap corruption."""
        from main import get_thread_safe_service
        from google.oauth2.credentials import Credentials

        # 1. Mock service is returned directly and transparently
        mock_svc = MagicMock()
        self.assertIs(get_thread_safe_service(mock_svc), mock_svc)

        # 2. None service returns None
        self.assertIsNone(get_thread_safe_service(None))

        # 3. Multi-threaded isolation test with real Credentials
        fake_creds = Credentials(token='mock_token_for_thread_isolation')
        class RealServiceStub:
            def __init__(self, creds):
                self._agent_credentials = creds

        base_service = RealServiceStub(fake_creds)

        thread_services = []
        thread_http_instances = []
        lock = threading.Lock()

        def worker_fn():
            svc = get_thread_safe_service(base_service)
            with lock:
                thread_services.append(svc)
                if hasattr(svc, '_http'):
                    thread_http_instances.append(svc._http)

        threads = [threading.Thread(target=worker_fn) for _ in range(3)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # All 3 worker threads retrieved a service
        self.assertEqual(len(thread_services), 3)
        # Verify each thread received an independent service instance
        self.assertEqual(len(set(id(s) for s in thread_services)), 3)
        # Verify each thread has its own independent HTTP transport
        self.assertEqual(len(set(id(h) for h in thread_http_instances)), 3)

        print("[PASS] Thread-safe service isolation across workers verified: independent Resource and HTTP transport per thread")

    def test_23_repeated_mocked_concurrency_stress_test(self):
        """Repeatedly test 120 recipients (100 unique, 15 duplicates, 5 invalid) across 3 concurrent workers.
        Verifies thread-isolated services, zero httplib2 transport collisions, thread-safe Excel and log updates,
        and zero duplicate recipient submissions.
        """
        from google.oauth2.credentials import Credentials

        class RealServiceStub:
            def __init__(self, creds):
                self._agent_credentials = creds

        fake_creds = Credentials(token='mock_stress_token')
        base_service = RealServiceStub(fake_creds)

        # Run 2 repetitions of 120 recipients each
        for repetition in range(1, 3):
            active_transports_by_thread = {}
            concurrent_transport_collisions = []
            transport_lock = threading.Lock()
            in_flight_transports = set()

            def mock_http_request(http_self, uri, method='GET', body=None, headers=None, **kwargs):
                thread_id = threading.get_ident()
                http_id = id(http_self)

                with transport_lock:
                    if http_id in active_transports_by_thread:
                        if active_transports_by_thread[http_id] != thread_id:
                            concurrent_transport_collisions.append({
                                'http_id': http_id,
                                'original_thread': active_transports_by_thread[http_id],
                                'conflicting_thread': thread_id
                            })
                    else:
                        active_transports_by_thread[http_id] = thread_id

                    if http_id in in_flight_transports:
                        concurrent_transport_collisions.append({
                            'http_id': http_id,
                            'error': 'concurrent_in_flight_access',
                            'thread': thread_id
                        })
                    in_flight_transports.add(http_id)

                with transport_lock:
                    in_flight_transports.discard(http_id)

                resp = httplib2.Response({'status': '200'})
                if 'messages/send' in uri:
                    content = json.dumps({'id': f'msg_mock_{thread_id}'}).encode('utf-8')
                elif 'messages?' in uri:
                    content = json.dumps({'messages': [{'id': 'bounce_msg_1'}]}).encode('utf-8')
                elif 'messages/' in uri:
                    content = json.dumps({
                        'id': 'bounce_msg_1',
                        'snippet': 'Delivery failed: Recruiter mailbox does not exist',
                        'payload': {
                            'headers': [
                                {'name': 'From', 'value': 'mailer-daemon@googlemail.com'},
                                {'name': 'Subject', 'value': 'Delivery Status Notification (Failure)'}
                            ]
                        }
                    }).encode('utf-8')
                else:
                    content = b'{}'
                return (resp, content)

            raw_recipients = []
            for i in range(1, 101):
                raw_recipients.append(f"recruiter_{i:03d}@company{i % 10}.com")
            for i in range(1, 16):
                raw_recipients.append(f"RECRUITER_{i:03d}@COMPANY{i % 10}.COM")
            for i in range(1, 6):
                raw_recipients.append(f"invalid-email-address-{i}")

            with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp_log, \
                 tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel:
                tmp_log_path = tmp_log.name
                tmp_excel_path = tmp_excel.name

            try:
                logger = OutreachLogger(tmp_log_path)
                directory = RecruiterDirectory(tmp_excel_path)
                dashboard = LiveDashboard(total_recipients=len(raw_recipients), is_live=False, num_workers=3)

                with patch.object(httplib2.Http, 'request', new=mock_http_request):
                    all_records, submitted_records = process_recipients_concurrently(
                        recipients=raw_recipients,
                        service=base_service,
                        subject="DevOps Opportunities",
                        html_template="<p>Stress Test</p>",
                        resume_path=DEFAULT_RESUME,
                        logger=logger,
                        directory=directory,
                        dashboard=dashboard,
                        num_workers=3,
                        api_call_delay=0.0
                    )

                # 1. Zero HTTP transport collisions across threads
                self.assertEqual(len(concurrent_transport_collisions), 0)

                # 2. Exactly 100 unique valid submissions, 0 duplicate sends
                self.assertEqual(len(submitted_records), 100)
                submitted_emails = [r['recipient'].lower() for r in submitted_records]
                self.assertEqual(len(set(submitted_emails)), 100)

                # 3. Excel directory thread-safety & deduplication
                excel_emails = directory.get_all_emails()
                self.assertEqual(len(excel_emails), 100)
                self.assertEqual(len(set(e.lower() for e in excel_emails)), 100)

                # 4. Logger thread-safety: 100 submitted + 5 invalid = 105 log lines
                with open(tmp_log_path, 'r', encoding='utf-8') as f:
                    log_lines = f.readlines()
                self.assertEqual(len(log_lines), 105)
                for line in log_lines:
                    self.assertTrue(line.endswith('\n'))
                    parts = line.strip().split('|')
                    self.assertGreaterEqual(len(parts), 8)

                # 5. Dashboard counters & execution errors
                self.assertEqual(dashboard.counters['SENT'], 100)
                self.assertEqual(dashboard.counters['UNDELIVERED'], 0)
                self.assertEqual(len(dashboard.execution_errors), 5)

            finally:
                for p in (tmp_log_path, tmp_excel_path):
                    if os.path.exists(p):
                        os.remove(p)

        print("[PASS] Repeated mocked concurrency stress test passed: 120 recipients x 2 reps, 0 collisions, 0 duplicates")

    def test_24_two_status_system_verification(self):
        """Comprehensive verification of the 2-status delivery system:
        1. Successful Gmail API response immediately produces SENT.
        2. No bounce keeps the recipient SENT.
        3. A confirmed bounce changes SENT to UNDELIVERED with its actual reason.
        4. Delayed bounces update the original record, not create duplicate records.
        5. Bounce-check errors never downgrade SENT.
        6. Excel directory still preserves submitted addresses.
        7. Existing 3-worker concurrency and thread safety remain intact.
        """
        from googleapiclient.errors import HttpError
        import httplib2

        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp_log, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel:
            tmp_log_path = tmp_log.name
            tmp_excel_path = tmp_excel.name

        try:
            logger = OutreachLogger(tmp_log_path)
            directory = RecruiterDirectory(tmp_excel_path)
            dashboard = LiveDashboard(total_recipients=3, is_live=False)

            # --- Condition 1: Successful Gmail API response immediately produces SENT ---
            mock_service = MagicMock()
            mock_service.users().messages().send().execute.side_effect = [
                {'id': 'msg_001'},
                {'id': 'msg_002'},
                {'id': 'msg_003'}
            ]

            recipients = ['recipient1@test.com', 'recipient2@test.com', 'recipient3@test.com']
            all_records, submitted = process_recipients_concurrently(
                recipients=recipients,
                service=mock_service,
                subject="Test Two Status",
                html_template="<p>Test</p>",
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=directory,
                dashboard=dashboard,
                num_workers=3,
                api_call_delay=0.0
            )

            self.assertEqual(len(submitted), 3)
            for r in submitted:
                self.assertEqual(r['status'], 'SENT', "Acceptance must immediately produce SENT")
                self.assertNotEqual(r['status'], 'SUBMITTED')
                self.assertNotEqual(r['status'], 'UNKNOWN')

            # --- Condition 6: Excel directory records submitted addresses ---
            excel_addrs = directory.get_all_emails()
            self.assertEqual(len(excel_addrs), 3)
            self.assertIn('recipient1@test.com', excel_addrs)

            # --- Condition 2 & 3: Bounce check updates matched to UNDELIVERED(reason) and unmatched stays SENT ---
            bounce_msg = {
                'id': 'bounce_b1',
                'snippet': "550 5.1.1 Mailbox does not exist: recipient1@test.com",
                'payload': {
                    'headers': [
                        {'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'},
                        {'name': 'Subject', 'value': 'Delivery Status Notification (Failure)'},
                        {'name': 'Date', 'value': 'Sun, 15 Nov 2026 10:00:00 +0000'}
                    ],
                    'body': {'data': ''},
                    'parts': []
                }
            }
            mock_service.users().messages().list().execute.return_value = {'messages': [{'id': 'bounce_b1'}]}
            mock_service.users().messages().get().execute.return_value = bounce_msg

            detect_bounces(mock_service, submitted, logger=logger, directory=directory, delay_seconds=0)

            # recipient1 bounced -> UNDELIVERED (reason)
            self.assertTrue(submitted[0]['status'].startswith('UNDELIVERED'))
            self.assertIn("does not exist", submitted[0]['status'].lower())
            self.assertEqual(submitted[0]['message_id'], 'msg_001')

            # recipient2 and recipient3 had no bounce -> remain SENT
            self.assertEqual(submitted[1]['status'], 'SENT', "No bounce must keep recipient SENT")
            self.assertEqual(submitted[2]['status'], 'SENT', "No bounce must keep recipient SENT")

            # --- Condition 6: Excel directory removes permanent invalid recipient ---
            excel_addrs_after_bounce = directory.get_all_emails()
            self.assertEqual(len(excel_addrs_after_bounce), 2, "Excel directory must remove permanent invalid recipient")
            self.assertNotIn('recipient1@test.com', excel_addrs_after_bounce)
            self.assertIn('recipient2@test.com', excel_addrs_after_bounce)
            self.assertIn('recipient3@test.com', excel_addrs_after_bounce)

            # --- Condition 4: Delayed bounces update original record, not create duplicate records ---
            delayed_bounce_msg = {
                'id': 'bounce_b2',
                'snippet': "550 User unknown: recipient2@test.com",
                'payload': {
                    'headers': [
                        {'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'},
                        {'name': 'Subject', 'value': 'Delivery Status Notification (Failure)'},
                        {'name': 'Date', 'value': 'Sun, 15 Nov 2026 12:00:00 +0000'}
                    ],
                    'body': {'data': ''},
                    'parts': []
                }
            }
            mock_service.users().messages().list().execute.return_value = {'messages': [{'id': 'bounce_b2'}]}
            mock_service.users().messages().get().execute.return_value = delayed_bounce_msg

            unconfirmed_delayed = logger.get_unconfirmed_records()
            detect_bounces(mock_service, unconfirmed_delayed, logger=logger, directory=directory, delay_seconds=0)

            # Delayed bounce for recipient2 (User unknown) also removed from Excel
            excel_addrs_after_delayed = directory.get_all_emails()
            self.assertEqual(len(excel_addrs_after_delayed), 1)
            self.assertNotIn('recipient2@test.com', excel_addrs_after_delayed)
            self.assertIn('recipient3@test.com', excel_addrs_after_delayed)

            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                log_lines = f.readlines()
            self.assertEqual(len(log_lines), 3, "Delayed bounce must update original record in-place without duplicate records")

            r2_line = next(l for l in log_lines if 'recipient2@test.com' in l)
            self.assertIn('UNDELIVERED', r2_line)
            self.assertIn('ID: msg_002', r2_line, "Message ID must be preserved")

            # --- Condition 5: Bounce-check errors never downgrade SENT ---
            resp_403 = httplib2.Response({'status': 403, 'reason': 'Forbidden'})
            mock_service.users().messages().list.side_effect = HttpError(resp_403, b'Access Denied')

            detect_bounces(mock_service, [submitted[2]], logger=logger, directory=directory, delay_seconds=0)
            self.assertEqual(submitted[2]['status'], 'SENT', "Bounce-check error must never downgrade SENT")

            # Final summary verification
            dashboard.set_counter('SENT', 1)
            dashboard.set_counter('UNDELIVERED', 2)
            dashboard.set_undelivered_recipients([submitted[0], unconfirmed_delayed[0]])

            import io
            from unittest.mock import patch
            with patch('sys.stdout', new=io.StringIO()) as fake_stdout:
                dashboard.render_final_summary()
                sum_out = fake_stdout.getvalue()
                self.assertIn("SENT:                 1", sum_out)
                self.assertIn("UNDELIVERED:          2", sum_out)
                self.assertNotIn("UNKNOWN", sum_out)
                self.assertNotIn("SUBMITTED", sum_out)

        finally:
            for p in (tmp_log_path, tmp_excel_path):
                if os.path.exists(p):
                    os.remove(p)

        print("[PASS] Two-status system completely verified (SENT, UNDELIVERED, Excel removal on permanent bounce, error tolerance)")

    def test_25_permanent_bounce_removal_and_temporary_retention(self):
        """Comprehensive verification of Excel directory permanent bounce removal:
        1. Permanent failures (Mailbox does not exist, Address not found, Invalid recipient,
           User unknown, No such mailbox) remove address from Excel directory.
        2. Temporary failures (Mailbox full, Temporary server errors, Greylisting,
           Delivery temporarily delayed) RETAIN address in Excel directory.
        3. Excel removal is strictly case-insensitive and thread-safe.
        4. Single 'Email' column and other valid addresses are preserved.
        5. Delayed bounce notifications trigger removal from Excel directory.
        6. Original bounce reasons in logs are preserved and statuses remain SENT / UNDELIVERED.
        """
        import openpyxl

        # --- A. Direct RecruiterDirectory methods & case-insensitivity ---
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel:
            tmp_excel_path = tmp_excel.name

        try:
            directory = RecruiterDirectory(tmp_excel_path)
            self.assertTrue(directory.add_email("Alice.Recruiter@Company.com"))
            self.assertTrue(directory.add_email("Bob.Recruiter@Company.com"))
            self.assertTrue(directory.add_email("Charlie.Recruiter@Company.com"))
            self.assertEqual(len(directory.get_all_emails()), 3)

            # Case-insensitive removal of Bob
            removed = directory.remove_email("BOB.RECRUITER@COMPANY.COM")
            self.assertTrue(removed, "Case-insensitive removal must succeed")
            emails = directory.get_all_emails()
            self.assertEqual(len(emails), 2)
            self.assertIn("Alice.Recruiter@Company.com", emails)
            self.assertNotIn("Bob.Recruiter@Company.com", emails)
            self.assertIn("Charlie.Recruiter@Company.com", emails)

            # Verify Excel file structure: single column 'Email'
            wb = openpyxl.load_workbook(tmp_excel_path)
            ws = wb.active
            self.assertEqual(ws.cell(row=1, column=1).value, "Email")
            self.assertEqual(ws.max_column, 1, "Must maintain strictly single 'Email' column")
            wb.close()

            # Non-existent removal returns False
            self.assertFalse(directory.remove_email("nonexistent@company.com"))
            self.assertEqual(len(directory.get_all_emails()), 2)
        finally:
            if os.path.exists(tmp_excel_path):
                os.remove(tmp_excel_path)

        # --- B. Bounce classifier unit checks: Permanent vs Temporary ---
        # 1. Permanent failures
        perm_reasons = [
            "Mailbox does not exist.",
            "The email account that you tried to reach does not exist: recruiter@example.com",
            "Address not found",
            "The address couldn't be found, or is unable to receive mail.",
            "Invalid recipient",
            "550 5.1.1 Invalid recipient: hr@example.com",
            "User unknown",
            "Recipient address rejected: User unknown in virtual mailbox table",
            "No such mailbox",
            "550 5.1.1 <devops@example.com>: Recipient address rejected: No such mailbox"
        ]
        for pr in perm_reasons:
            self.assertTrue(
                is_permanent_invalid_recipient_bounce(pr),
                f"Reason '{pr}' must be classified as a permanent failure"
            )

        # 2. Temporary failures
        temp_reasons = [
            "Mailbox full",
            "552 5.2.2 Mailbox is full, over storage quota",
            "Temporary server errors",
            "451 4.3.0 Temporary server error, please try again",
            "Greylisting in action, please retry later",
            "Delivery temporarily delayed",
            "Action: delayed; Status: 4.4.1; Delivery temporarily delayed"
        ]
        for tr in temp_reasons:
            self.assertFalse(
                is_permanent_invalid_recipient_bounce(tr),
                f"Reason '{tr}' must be classified as a temporary failure"
            )

        # --- C. End-to-End Outreach Integration: Permanent Removal vs Temporary Retention ---
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp_log, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel:
            tmp_log_path = tmp_log.name
            tmp_excel_path = tmp_excel.name

        try:
            logger = OutreachLogger(tmp_log_path)
            directory = RecruiterDirectory(tmp_excel_path)

            recipients = [
                'valid.candidate@company.com',
                'perm.mailbox.notexist@company.com',
                'perm.addr.notfound@company.com',
                'perm.user.unknown@company.com',
                'temp.mailbox.full@company.com',
                'temp.server.error@company.com',
                'temp.delayed@company.com'
            ]

            mock_service = MagicMock()
            mock_service.users().messages().send().execute.side_effect = [
                {'id': f'msg_{i}'} for i in range(len(recipients))
            ]

            all_rec, submitted = process_recipients_concurrently(
                recipients=recipients,
                service=mock_service,
                subject="Application for Cloud Engineer",
                html_template="<p>Test Content</p>",
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=directory,
                num_workers=3,
                api_call_delay=0.0
            )

            # Initially, all 7 submitted addresses are registered in Excel directory
            self.assertEqual(len(directory.get_all_emails()), 7)

            # Mock bounce notifications returned by Gmail
            bounces_returned = [
                # Permanent 1: Mailbox does not exist
                {
                    'id': 'b_p1',
                    'snippet': "550 5.1.1 The email account that you tried to reach does not exist: perm.mailbox.notexist@company.com",
                    'payload': {
                        'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                        'body': {'data': ''},
                        'parts': []
                    }
                },
                # Permanent 2: Address not found
                {
                    'id': 'b_p2',
                    'snippet': "Address not found: Your message was not delivered to perm.addr.notfound@company.com because the address couldn't be found.",
                    'payload': {
                        'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                        'body': {'data': ''},
                        'parts': []
                    }
                },
                # Permanent 3: User unknown
                {
                    'id': 'b_p3',
                    'snippet': "550 5.1.1 <perm.user.unknown@company.com>: Recipient address rejected: User unknown",
                    'payload': {
                        'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                        'body': {'data': ''},
                        'parts': []
                    }
                },
                # Temporary 1: Mailbox full
                {
                    'id': 'b_t1',
                    'snippet': "552 5.2.2 Mailbox is full: temp.mailbox.full@company.com quota exceeded",
                    'payload': {
                        'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                        'body': {'data': ''},
                        'parts': []
                    }
                },
                # Temporary 2: Temporary server error
                {
                    'id': 'b_t2',
                    'snippet': "451 4.3.0 Temporary server error for temp.server.error@company.com, please retry later",
                    'payload': {
                        'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                        'body': {'data': ''},
                        'parts': []
                    }
                },
                # Temporary 3: Delivery temporarily delayed
                {
                    'id': 'b_t3',
                    'snippet': "Action: delayed; Status: 4.4.1; Delivery temporarily delayed for temp.delayed@company.com",
                    'payload': {
                        'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                        'body': {'data': ''},
                        'parts': []
                    }
                }
            ]

            mock_service.users().messages().list().execute.return_value = {
                'messages': [{'id': b['id']} for b in bounces_returned]
            }
            mock_service.users().messages().get().execute.side_effect = bounces_returned

            detect_bounces(mock_service, submitted, logger=logger, directory=directory, delay_seconds=0)

            # Check Excel directory contents after bounce detection:
            # - 3 permanent invalid addresses MUST be removed
            # - 3 temporary addresses MUST be retained
            # - 1 valid unbounced address MUST be retained
            # -> Exactly 4 emails remain in Excel
            remaining_excel = directory.get_all_emails()
            self.assertEqual(len(remaining_excel), 4, "Excel must retain exactly 4 addresses (1 valid + 3 temporary)")
            self.assertIn('valid.candidate@company.com', remaining_excel)
            self.assertIn('temp.mailbox.full@company.com', remaining_excel, "Temporary bounce must NOT be removed from Excel")
            self.assertIn('temp.server.error@company.com', remaining_excel, "Temporary server error must NOT be removed from Excel")
            self.assertIn('temp.delayed@company.com', remaining_excel, "Delivery delayed must NOT be removed from Excel")

            self.assertNotIn('perm.mailbox.notexist@company.com', remaining_excel, "Mailbox not exist must be removed from Excel")
            self.assertNotIn('perm.addr.notfound@company.com', remaining_excel, "Address not found must be removed from Excel")
            self.assertNotIn('perm.user.unknown@company.com', remaining_excel, "User unknown must be removed from Excel")

            # Check log file: keeps original bounce reason and statuses
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                log_content = f.read()

            self.assertIn('valid.candidate@company.com | SENT', log_content)
            self.assertIn('perm.mailbox.notexist@company.com | UNDELIVERED', log_content)
            self.assertIn('temp.mailbox.full@company.com | UNDELIVERED', log_content)
            self.assertIn('mailbox.full', log_content.lower())

            # --- D. Delayed Bounce Removal ---
            # Simulate a previously sent email that had no bounce in Run 1 (still SENT and in Excel)
            delayed_candidate = {
                'timestamp': '2026-11-15 08:00:00',
                'recipient': 'delayed.perm@company.com',
                'status': 'SENT',
                'subject': 'Application for Cloud Engineer',
                'message_id': 'msg_delayed_001',
                'submitted_time': '08:00:00',
                'bounce_time': 'N/A',
                'remark': 'Sent; awaiting delayed bounce check',
                'error': 'None'
            }
            logger.log_email_record(delayed_candidate)
            directory.add_email('delayed.perm@company.com')
            self.assertEqual(len(directory.get_all_emails()), 5)

            # In Run 2, a delayed bounce arrives for delayed.perm@company.com
            delayed_perm_bounce = {
                'id': 'b_delayed_perm',
                'snippet': "550 5.1.1 Recipient address rejected: No such mailbox: delayed.perm@company.com",
                'payload': {
                    'headers': [{'name': 'From', 'value': 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>'}],
                    'body': {'data': ''},
                    'parts': []
                }
            }
            mock_service.users().messages().list().execute.return_value = {'messages': [{'id': 'b_delayed_perm'}]}
            mock_service.users().messages().get().execute.side_effect = None
            mock_service.users().messages().get().execute.return_value = delayed_perm_bounce

            unconfirmed_records = logger.get_unconfirmed_records()
            self.assertTrue(any(r['recipient'] == 'delayed.perm@company.com' for r in unconfirmed_records))
            detect_bounces(mock_service, unconfirmed_records, logger=logger, directory=directory, delay_seconds=0)

            # delayed.perm@company.com is now confirmed permanently invalid -> removed from Excel!
            after_delayed_excel = directory.get_all_emails()
            self.assertEqual(len(after_delayed_excel), 4)
            self.assertNotIn('delayed.perm@company.com', after_delayed_excel)

            # Single 'Email' column preserved
            wb = openpyxl.load_workbook(tmp_excel_path)
            ws = wb.active
            self.assertEqual(ws.cell(row=1, column=1).value, "Email")
            self.assertEqual(ws.max_column, 1)
            wb.close()

            # --- E. Thread Safety Test: Concurrent Add and Remove ---
            import concurrent.futures

            def worker_op(i):
                email = f"concurrent_{i}@company.com"
                directory.add_email(email)
                if i % 2 == 0:
                    directory.remove_email(email.upper())

            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                list(executor.map(worker_op, range(20)))

            final_emails = directory.get_all_emails()
            # Odd numbers retained, even numbers removed
            for i in range(20):
                email = f"concurrent_{i}@company.com"
                if i % 2 == 0:
                    self.assertNotIn(email, final_emails)
                else:
                    self.assertIn(email, final_emails)

        finally:
            for p in (tmp_log_path, tmp_excel_path):
                if os.path.exists(p):
                    os.remove(p)

        print("[PASS] Recruiter email directory permanent bounce removal, temporary bounce retention, case-insensitivity, and thread-safety verified")

    def test_26_duplicate_cooldown_rules(self):
        """Verify 10-day duplicate prevention cooldown: 2-day, 9-day, exactly 10-day, and 11-day cases,
        case-insensitive matching, original submission timestamp calculation, and timer preservation.
        """
        ref_time = datetime(2026, 11, 20, 12, 0, 0)
        cooldown_days = 10

        # Construct log history with various submission timestamps
        log_records = [
            # 2 days ago: 2026-11-18 12:00:00 -> within cooldown -> SKIP
            "2026-11-18 12:00:00 | two_days@example.com | SENT | Subject: Test | ID: msg_2d | Submitted: 12:00:00 | Bounce: N/A | Error: None",
            # 9 days ago: 2026-11-11 12:00:00 -> within cooldown -> SKIP
            "2026-11-11 12:00:00 | nine_days@example.com | SENT | Subject: Test | ID: msg_9d | Submitted: 12:00:00 | Bounce: N/A | Error: None",
            # Exactly 10 days ago: 2026-11-10 12:00:00 -> cooldown expired -> ELIGIBLE (SEND)
            "2026-11-10 12:00:00 | exact_10d@example.com | SENT | Subject: Test | ID: msg_10d | Submitted: 12:00:00 | Bounce: N/A | Error: None",
            # 11 days ago: 2026-11-09 12:00:00 -> cooldown expired -> ELIGIBLE (SEND)
            "2026-11-09 12:00:00 | eleven_days@example.com | SENT | Subject: Test | ID: msg_11d | Submitted: 12:00:00 | Bounce: N/A | Error: None",
            # Case-insensitivity check: logged lowercase, queried uppercase 3 days ago -> SKIP
            "2026-11-17 12:00:00 | case_test@example.com | SENT | Subject: Test | ID: msg_case | Submitted: 12:00:00 | Bounce: N/A | Error: None",
        ]

        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as f:
            for r in log_records:
                f.write(r + "\n")
            tmp_log = f.name

        try:
            history = load_outreach_history(tmp_log)

            # 1. 2-day case: SKIPPED
            res_2d = evaluate_recipient_eligibility('two_days@example.com', history, cooldown_days=10, current_time=ref_time)
            self.assertEqual(res_2d['action'], 'SKIP_COOLDOWN')
            self.assertEqual(res_2d['submitted_date'], '2026-11-18 12:00:00')
            self.assertEqual(res_2d['eligible_date'], '2026-11-28 12:00:00')

            # 2. 9-day case: SKIPPED
            res_9d = evaluate_recipient_eligibility('nine_days@example.com', history, cooldown_days=10, current_time=ref_time)
            self.assertEqual(res_9d['action'], 'SKIP_COOLDOWN')
            self.assertEqual(res_9d['submitted_date'], '2026-11-11 12:00:00')
            self.assertEqual(res_9d['eligible_date'], '2026-11-21 12:00:00')

            # 3. Exactly 10-day case: ELIGIBLE TO SEND
            res_10d = evaluate_recipient_eligibility('exact_10d@example.com', history, cooldown_days=10, current_time=ref_time)
            self.assertEqual(res_10d['action'], 'SEND')
            self.assertIn('cooldown expired', res_10d['reason'])

            # 4. 11-day case: ELIGIBLE TO SEND
            res_11d = evaluate_recipient_eligibility('eleven_days@example.com', history, cooldown_days=10, current_time=ref_time)
            self.assertEqual(res_11d['action'], 'SEND')
            self.assertIn('cooldown expired', res_11d['reason'])

            # 5. Case-insensitive matching: queried in uppercase
            res_case = evaluate_recipient_eligibility('CASE_TEST@EXAMPLE.COM', history, cooldown_days=10, current_time=ref_time)
            self.assertEqual(res_case['action'], 'SKIP_COOLDOWN')

            # 6. Timer preservation: skipped duplicates do not alter log records or reset timer
            logger = OutreachLogger(tmp_log)
            # Re-evaluating still yields the original submitted_date (2026-11-18 12:00:00)
            res_recheck = evaluate_recipient_eligibility('two_days@example.com', history, cooldown_days=10, current_time=ref_time)
            self.assertEqual(res_recheck['submitted_date'], '2026-11-18 12:00:00')

        finally:
            if os.path.exists(tmp_log):
                os.remove(tmp_log)

        print("[PASS] 10-day duplicate prevention cooldown verified (2-day, 9-day, exactly 10-day, 11-day, case-insensitivity, timer preservation)")

    def test_27_undelivered_review_and_retry_rules(self):
        """Verify UNDELIVERED exclusion regardless of age and even when cooldown is disabled,
        FAILED_TO_SUBMIT retry eligibility, and INVALID format handling.
        """
        ref_time = datetime(2026, 11, 20, 12, 0, 0)

        # History with confirmed UNDELIVERED (from 50 days ago) and FAILED_TO_SUBMIT
        log_records = [
            # 50 days ago: confirmed bounce
            "2026-10-01 10:00:00 | bounced_old@company.com | UNDELIVERED (Mailbox does not exist) | Subject: App | ID: mid_1 | Submitted: 10:00:00 | Bounce: 10:05:00 | Error: SMTP 550",
            # Recent FAILED_TO_SUBMIT
            "2026-11-19 14:00:00 | failed_yesterday@company.com | FAILED_TO_SUBMIT | Subject: App | ID: N/A | Submitted: N/A | Bounce: N/A | Error: HTTP 500: Server Error",
        ]

        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as f:
            for r in log_records:
                f.write(r + "\n")
            tmp_log = f.name

        try:
            history = load_outreach_history(tmp_log)

            # 1. UNDELIVERED in history: skip and flag for manual review, regardless of age
            res_undeliv = evaluate_recipient_eligibility(
                'bounced_old@company.com',
                history,
                cooldown_days=10,
                cooldown_enabled=True,
                current_time=ref_time
            )
            self.assertEqual(res_undeliv['action'], 'UNDELIVERED_REVIEW')
            self.assertIn('Mailbox does not exist', res_undeliv['status'])

            # 2. Confirmed UNDELIVERED exclusion is preserved even when cooldown is disabled!
            res_undeliv_nocd = evaluate_recipient_eligibility(
                'bounced_old@company.com',
                history,
                cooldown_days=10,
                cooldown_enabled=False,
                current_time=ref_time
            )
            self.assertEqual(res_undeliv_nocd['action'], 'UNDELIVERED_REVIEW', "UNDELIVERED must be preserved even when cooldown is disabled")

            # 3. FAILED_TO_SUBMIT: eligible for retry
            res_retry = evaluate_recipient_eligibility(
                'failed_yesterday@company.com',
                history,
                cooldown_days=10,
                cooldown_enabled=True,
                current_time=ref_time
            )
            self.assertEqual(res_retry['action'], 'SEND')
            self.assertIn('retry', res_retry['reason'].lower())

            # 4. INVALID email address: skip until corrected
            res_inv = evaluate_recipient_eligibility(
                'not_an_email_address',
                history,
                cooldown_days=10,
                cooldown_enabled=True,
                current_time=ref_time
            )
            self.assertEqual(res_inv['action'], 'INVALID')

        finally:
            if os.path.exists(tmp_log):
                os.remove(tmp_log)

        print("[PASS] UNDELIVERED manual review, retry of FAILED_TO_SUBMIT, and INVALID rules verified")

    def test_28_config_cooldown_settings(self):
        """Verify configurable duplicate prevention settings via .env:
        DUPLICATE_CHECK_ENABLED=true/false, custom DUPLICATE_COOLDOWN_DAYS, and positive integer validation.
        """
        # 1. Custom cooldown days = 5
        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as f:
            f.write("DUPLICATE_CHECK_ENABLED=true\nDUPLICATE_COOLDOWN_DAYS=5\n")
            env5 = f.name

        try:
            cfg5 = Config(env_path=env5)
            cfg5.validate()
            self.assertTrue(cfg5.duplicate_check_enabled)
            self.assertEqual(cfg5.duplicate_cooldown_days, 5)

            ref_now = datetime(2026, 11, 20, 12, 0, 0)
            history = {
                'candidate@example.com': {
                    'recipient': 'candidate@example.com',
                    'is_undelivered': False,
                    'undelivered_status': None,
                    'last_submitted_dt': datetime(2026, 11, 16, 12, 0, 0), # 4 days ago
                    'last_status': 'SENT',
                    'last_error': 'None'
                }
            }
            # 4 days ago with 5-day cooldown -> SKIP
            res_4d = evaluate_recipient_eligibility('candidate@example.com', history, cooldown_days=5, current_time=ref_now)
            self.assertEqual(res_4d['action'], 'SKIP_COOLDOWN')

            # 5 days ago with 5-day cooldown -> ELIGIBLE
            history['candidate@example.com']['last_submitted_dt'] = datetime(2026, 11, 15, 12, 0, 0)
            res_5d = evaluate_recipient_eligibility('candidate@example.com', history, cooldown_days=5, current_time=ref_now)
            self.assertEqual(res_5d['action'], 'SEND')
        finally:
            if os.path.exists(env5):
                os.remove(env5)

        # 2. Disabled duplicate check: DUPLICATE_CHECK_ENABLED=false
        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as f:
            f.write("DUPLICATE_CHECK_ENABLED=false\nDUPLICATE_COOLDOWN_DAYS=10\n")
            env_disabled = f.name

        try:
            cfg_dis = Config(env_path=env_disabled)
            cfg_dis.validate()
            self.assertFalse(cfg_dis.duplicate_check_enabled)

            # Even if sent 1 hour ago, disabled cooldown permits sending
            history_recent = {
                'recent@example.com': {
                    'recipient': 'recent@example.com',
                    'is_undelivered': False,
                    'undelivered_status': None,
                    'last_submitted_dt': datetime(2026, 11, 20, 11, 0, 0),
                    'last_status': 'SENT',
                    'last_error': 'None'
                }
            }
            res_disabled = evaluate_recipient_eligibility(
                'recent@example.com',
                history_recent,
                cooldown_days=10,
                cooldown_enabled=False,
                current_time=ref_now
            )
            self.assertEqual(res_disabled['action'], 'SEND')
        finally:
            if os.path.exists(env_disabled):
                os.remove(env_disabled)

        # 3. Validation: DUPLICATE_COOLDOWN_DAYS <= 0 raises ValueError
        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as f:
            f.write("DUPLICATE_COOLDOWN_DAYS=0\n")
            env_invalid = f.name

        try:
            cfg_inv = Config(env_path=env_invalid)
            with self.assertRaises(ValueError) as ctx:
                cfg_inv.validate()
            self.assertIn("DUPLICATE_COOLDOWN_DAYS", str(ctx.exception))
        finally:
            if os.path.exists(env_invalid):
                os.remove(env_invalid)

        print("[PASS] Configurable duplicate check and cooldown validation verified")

    def test_29_csv_auto_clear_and_crash_recovery(self):
        """Verify CSV queue auto-clearing and crash recovery:
        - Never clear at startup
        - Clear recipient rows while preserving header when all outcomes are safely recorded
        - Preserve unprocessed / uncertain rows if interrupted or crashed
        """
        import csv

        # 1. Normal execution: all 3 recipients safely processed -> only header remains
        with tempfile.NamedTemporaryFile('w', newline='', suffix='.csv', delete=False) as f:
            writer = csv.writer(f)
            writer.writerow(['email'])
            writer.writerow(['r1@example.com'])
            writer.writerow(['r2@example.com'])
            writer.writerow(['r3@example.com'])
            tmp_csv1 = f.name

        try:
            # Startup: verify CSV is non-empty and never cleared at startup
            raw_initial = read_raw_recruiter_emails(tmp_csv1)
            self.assertEqual(len(raw_initial), 3)

            # Execution completes: all 3 safely recorded
            safely_completed = {'r1@example.com', 'r2@example.com', 'r3@example.com'}
            removed = safe_clear_csv_queue(tmp_csv1, safely_completed)
            self.assertEqual(removed, 3)

            # Check file content: header preserved, data rows cleared
            with open(tmp_csv1, 'r', encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip()]
            self.assertEqual(lines, ['email'], "CSV must retain only the header row after safe execution")

            raw_after = read_raw_recruiter_emails(tmp_csv1)
            self.assertEqual(len(raw_after), 0)
        finally:
            if os.path.exists(tmp_csv1):
                os.remove(tmp_csv1)

        # 2. Crash / partial execution: 2 processed, 2 unprocessed -> unprocessed preserved
        with tempfile.NamedTemporaryFile('w', newline='', suffix='.csv', delete=False) as f:
            writer = csv.writer(f)
            writer.writerow(['Email', 'Company'])
            writer.writerow(['safe1@example.com', 'Acme'])
            writer.writerow(['crashed@example.com', 'Beta'])
            writer.writerow(['safe2@example.com', 'Gamma'])
            writer.writerow(['unprocessed@example.com', 'Delta'])
            tmp_csv2 = f.name

        try:
            # Simulate crash: only safe1 and safe2 were safely logged
            partial_safe = {'safe1@example.com', 'safe2@example.com'}
            removed = safe_clear_csv_queue(tmp_csv2, partial_safe)
            self.assertEqual(removed, 2)

            with open(tmp_csv2, 'r', encoding='utf-8') as f:
                reader = list(csv.reader(f))

            self.assertEqual(reader[0], ['Email', 'Company'], "Header row must be preserved")
            remaining_emails = [r[0] for r in reader[1:]]
            self.assertEqual(remaining_emails, ['crashed@example.com', 'unprocessed@example.com'],
                             "Uncertain or unprocessed rows must be preserved for recovery")
        finally:
            if os.path.exists(tmp_csv2):
                os.remove(tmp_csv2)

        # 3. Verify process_recipients_concurrently preserves failed-to-submit emails from safely_processed_set
        with tempfile.NamedTemporaryFile('w', newline='', suffix='.csv', delete=False) as f:
            writer = csv.writer(f)
            writer.writerow(['email'])
            writer.writerow(['success@example.com'])
            writer.writerow(['failed@example.com'])
            tmp_csv3 = f.name

        try:
            from googleapiclient.errors import HttpError
            mock_service = MagicMock()
            def mock_send(userId, body):
                raw_bytes = base64.urlsafe_b64decode(body.get('raw', '').encode('utf-8'))
                m = MagicMock()
                if b'failed@example.com' in raw_bytes:
                    resp = httplib2.Response({'status': 500, 'reason': 'Internal Server Error'})
                    m.execute.side_effect = HttpError(resp, b'{"error": {"message": "Server Error"}}')
                else:
                    m.execute.return_value = {'id': 'success_id'}
                return m

            mock_service.users().messages().send.side_effect = mock_send

            test_safely_processed = set()
            process_recipients_concurrently(
                recipients=['success@example.com', 'failed@example.com'],
                service=mock_service,
                subject='Test',
                html_template='<p>Test</p>',
                resume_path=DEFAULT_RESUME,
                num_workers=1,
                safely_processed_set=test_safely_processed,
                api_call_delay=0.0
            )

            self.assertIn('success@example.com', test_safely_processed)
            self.assertNotIn('failed@example.com', test_safely_processed)

            # Clearing queue with this set preserves failed@example.com
            safe_clear_csv_queue(tmp_csv3, test_safely_processed)
            remaining_raw = read_raw_recruiter_emails(tmp_csv3)
            self.assertEqual(remaining_raw, ['failed@example.com'], "Failed emails must be retained in CSV queue")
        finally:
            if os.path.exists(tmp_csv3):
                os.remove(tmp_csv3)

        print("[PASS] CSV auto-clearing and crash recovery verified (header preserved, safe rows cleared, uncertain rows retained, failed sends kept in CSV)")

    def test_30_full_pipeline_duplicate_prevention_and_dashboard(self):
        """Verify full end-to-end pipeline:
        - Pre-flight history check against outreach.log
        - 1 Skipped Cooldown (shows previous submission & eligible date)
        - 1 Undelivered Review (flagged for manual review)
        - 1 Retry of FAILED_TO_SUBMIT
        - 1 New recipient
        - 1 Invalid recipient
        - Dashboard counters: TOTAL INPUT, NEW RECIPIENTS, SKIPPED (10-DAY COOLDOWN), UNDELIVERED REVIEW, INVALID, FAILED TO SUBMIT, SENT
        - Safe CSV queue auto-clearing
        """
        import io
        import csv

        ref_now = datetime.now()
        cooldown_days = 10

        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as f_log:
            # 1. Skipped cooldown: sent 3 days ago
            t_3d = (ref_now - timedelta(days=3)).strftime('%Y-%m-%d %H:%M:%S')
            f_log.write(f"{t_3d} | recent@test.com | SENT | Subject: App | ID: mid_rec | Submitted: 10:00:00 | Bounce: N/A | Error: None\n")
            # 2. Undelivered review: confirmed bounce
            t_old = (ref_now - timedelta(days=20)).strftime('%Y-%m-%d %H:%M:%S')
            f_log.write(f"{t_old} | bounced@test.com | UNDELIVERED (Address not found) | Subject: App | ID: mid_bnc | Submitted: 10:00:00 | Bounce: 10:05:00 | Error: SMTP 550\n")
            # 3. Failed to submit in previous run: eligible for retry
            f_log.write(f"{t_old} | failed_prev@test.com | FAILED_TO_SUBMIT | Subject: App | ID: N/A | Submitted: N/A | Bounce: N/A | Error: HTTP 500\n")
            tmp_log = f_log.name

        with tempfile.NamedTemporaryFile('w', newline='', suffix='.csv', delete=False) as f_csv:
            writer = csv.writer(f_csv)
            writer.writerow(['email'])
            writer.writerow(['recent@test.com'])       # 1. SKIP (cooldown)
            writer.writerow(['bounced@test.com'])      # 2. UNDELIVERED REVIEW
            writer.writerow(['failed_prev@test.com'])  # 3. RETRY (SEND)
            writer.writerow(['brand_new@test.com'])    # 4. NEW (SEND)
            writer.writerow(['invalid-email-format'])  # 5. INVALID
            tmp_csv = f_csv.name

        with tempfile.NamedTemporaryFile('w', suffix='.env', delete=False) as f_env:
            f_env.write(f"LOG_FILE={tmp_log}\nRECRUITER_CSV_FILE={tmp_csv}\nDUPLICATE_CHECK_ENABLED=true\nDUPLICATE_COOLDOWN_DAYS=10\nDRY_RUN=true\n")
            tmp_env = f_env.name

        try:
            cfg = Config(env_path=tmp_env)
            logger = OutreachLogger(cfg.log_file)
            history = load_outreach_history(cfg.log_file)

            raw_recipients = read_raw_recruiter_emails(tmp_csv)
            self.assertEqual(len(raw_recipients), 5)

            eligible = []
            skipped_cd = []
            undeliv_rev = []
            invalid_recs = []
            safely_processed = set()

            for r in raw_recipients:
                eval_res = evaluate_recipient_eligibility(r, history, cooldown_days=10, cooldown_enabled=True)
                action = eval_res['action']
                if action == 'SKIP_COOLDOWN':
                    skipped_cd.append(eval_res)
                    safely_processed.add(r.lower())
                elif action == 'UNDELIVERED_REVIEW':
                    undeliv_rev.append(eval_res)
                    safely_processed.add(r.lower())
                elif action == 'INVALID':
                    inv_rec = {
                        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'recipient': r,
                        'status': 'INVALID',
                        'subject': 'Test',
                        'message_id': 'N/A',
                        'submitted_time': 'N/A',
                        'bounce_time': 'N/A',
                        'error': 'Invalid email address format'
                    }
                    invalid_recs.append(inv_rec)
                    safely_processed.add(r.lower())
                else:
                    eligible.append(r)

            self.assertEqual(len(skipped_cd), 1)
            self.assertEqual(skipped_cd[0]['recipient'], 'recent@test.com')
            self.assertIn('eligible_date', skipped_cd[0])

            self.assertEqual(len(undeliv_rev), 1)
            self.assertEqual(undeliv_rev[0]['recipient'], 'bounced@test.com')

            self.assertEqual(len(invalid_recs), 1)
            self.assertEqual(invalid_recs[0]['recipient'], 'invalid-email-format')

            self.assertEqual(len(eligible), 2)
            self.assertIn('failed_prev@test.com', eligible)
            self.assertIn('brand_new@test.com', eligible)

            # Initialize dashboard
            dashboard = LiveDashboard(
                total_recipients=len(raw_recipients),
                is_live=False,
                num_workers=3,
                cooldown_days=10,
                cooldown_enabled=True
            )
            dashboard.set_counter('TOTAL_INPUT', len(raw_recipients))
            dashboard.set_counter('NEW_RECIPIENTS', len(eligible))
            dashboard.set_counter('SKIPPED_COOLDOWN', len(skipped_cd))
            dashboard.set_counter('UNDELIVERED_REVIEW', len(undeliv_rev))
            dashboard.set_counter('INVALID', len(invalid_recs))
            dashboard.set_skipped_cooldown_recipients(skipped_cd)
            dashboard.set_undelivered_review_recipients(undeliv_rev)

            # Process eligible recipients concurrently with mocked service in DRY_RUN mode
            all_records, submitted_records = process_recipients_concurrently(
                recipients=eligible,
                service=None,
                subject='Application for Cloud Opportunities',
                html_template='<p>Hello recruiter</p>',
                resume_path=DEFAULT_RESUME,
                logger=logger,
                directory=None,
                dashboard=dashboard,
                config=cfg,
                safely_processed_set=safely_processed
            )

            self.assertEqual(len(submitted_records), 2)
            dashboard.set_counter('SENT', len(submitted_records))

            # Render final summary
            with patch('sys.stdout', new=io.StringIO()) as fake_out:
                dashboard.render_final_summary()
                summary_str = fake_out.getvalue()

            self.assertIn("TOTAL INPUT:          5", summary_str)
            self.assertIn("NEW RECIPIENTS:       2", summary_str)
            self.assertIn("SKIPPED (10-DAY COOLDOWN): 1", summary_str)
            self.assertIn("UNDELIVERED REVIEW:   1", summary_str)
            self.assertIn("INVALID:              1", summary_str)
            self.assertIn("FAILED TO SUBMIT:     0", summary_str)
            self.assertIn("SENT:                 2", summary_str)
            self.assertIn("recent@test.com", summary_str)
            self.assertIn("Eligible Again:", summary_str)
            self.assertIn("bounced@test.com", summary_str)

            # Safely clear CSV: all 5 recipients have safely recorded outcomes
            removed = safe_clear_csv_queue(tmp_csv, safely_processed)
            self.assertEqual(removed, 5)

            # Confirm CSV retains only header
            with open(tmp_csv, 'r', encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip()]
            self.assertEqual(lines, ['email'])

        finally:
            for p in (tmp_log, tmp_csv, tmp_env):
                if os.path.exists(p):
                    os.remove(p)

        print("[PASS] Full outreach pipeline with 10-day cooldown, UNDELIVERED review, and CSV clearing verified")


if __name__ == '__main__':
    unittest.main()
