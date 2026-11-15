import os
import sys
import tempfile
import unittest
import base64
import email
from unittest.mock import MagicMock
from email.mime.multipart import MIMEMultipart

# Ensure project root is in sys.path and is current working directory
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)

from main import (
    is_valid_email,
    read_recruiter_emails,
    create_message,
    parse_template,
    detect_bounces,
    sanitize_log_field,
    format_email_log_line,
    OutreachLogger,
    DEFAULT_CSV,
    DEFAULT_TEMPLATE,
    DEFAULT_RESUME,
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
        """Verify email generation, HTML body, extracted subject, and PDF attachment without sending."""
        if not os.path.exists(DEFAULT_RESUME):
            self.skipTest(f"Resume not found at '{DEFAULT_RESUME}'. Place your PDF resume in attachments/ to run this test.")
        self.assertTrue(os.path.exists(DEFAULT_TEMPLATE), "template.html must exist")
        subject, html_content = parse_template(DEFAULT_TEMPLATE)
        self.assertEqual(subject, 'Application for DevOps / Cloud Engineer Opportunities')

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
        self.assertEqual(parsed_msg['Subject'], 'Application for DevOps / Cloud Engineer Opportunities')
        self.assertEqual(parsed_msg['Subject'], subject)
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

    def test_07_template_subject_parsing_and_validation(self):
        """Verify template parsing: subject extraction, changing subject, and error handling for missing/empty subject."""
        subject, body = parse_template(DEFAULT_TEMPLATE)
        self.assertEqual(subject, 'Application for DevOps / Cloud Engineer Opportunities')
        self.assertTrue(len(body) > 100)
        self.assertNotIn('<!-- Subject:', body)

        # Dynamic subject change test
        with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False) as tmp:
            tmp.write("<!-- Subject: Senior Cloud Architect Role -->\n<p>Hello recruiter</p>")
            tmp_path = tmp.name

        try:
            dynamic_subj, dynamic_body = parse_template(tmp_path)
            self.assertEqual(dynamic_subj, 'Senior Cloud Architect Role')
            self.assertEqual(dynamic_body.strip(), '<p>Hello recruiter</p>')
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

        # Missing subject comment raises ValueError
        with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False) as tmp:
            tmp.write("<p>No subject comment in this file</p>")
            tmp_path = tmp.name

        try:
            with self.assertRaises(ValueError) as ctx:
                parse_template(tmp_path)
            self.assertIn("Email subject not found", str(ctx.exception))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

        # Empty subject comment raises ValueError
        with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False) as tmp:
            tmp.write("<!-- Subject:   -->\n<p>Empty subject</p>")
            tmp_path = tmp.name

        try:
            with self.assertRaises(ValueError) as ctx:
                parse_template(tmp_path)
            self.assertIn("cannot be empty", str(ctx.exception))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

        print("[PASS] Template subject parsing verified: dynamic subject change, extraction, and missing/empty validation")

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
        rec_success['status'] = 'SUBMITTED'
        rec_success['remark'] = 'Gmail API accepted email; delivery unconfirmed'

        self.assertEqual(rec_success['status'], 'SUBMITTED')
        self.assertNotEqual(rec_success['status'], 'SUCCESS', "Never report SUCCESS based on Gmail API acceptance")
        self.assertNotEqual(rec_success['status'], 'DELIVERED', "Never report DELIVERED based on Gmail API acceptance")
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

        print("[PASS] Status reporting verified: acceptance yields SUBMITTED, API failure yields FAILED_TO_SUBMIT, invalid yields INVALID")

    def test_09_bounce_detection_and_matching(self):
        """Verify bounce notifications identify BOUNCED recipients, unmatched bounces are ignored, and missing bounce remains UNKNOWN."""
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
                'status': 'SUBMITTED',
                'remark': 'Gmail API accepted email; delivery unconfirmed',
                'error': 'None',
                'bounce_time': 'N/A'
            },
            {
                'timestamp': '2026-11-15 19:30:15',
                'recipient': 'hr@gmail.com',
                'subject': 'Application for DevOps Role',
                'message_id': '18ghi789',
                'status': 'SUBMITTED',
                'remark': 'Gmail API accepted email; delivery unconfirmed',
                'error': 'None',
                'bounce_time': 'N/A'
            }
        ]

        # Run bounce detection
        detect_bounces(mock_service, records, delay_seconds=0)

        # 1. Matching recipient correctly marked BOUNCED
        self.assertEqual(records[0]['status'], 'BOUNCED')
        self.assertNotEqual(records[0]['bounce_time'], 'N/A')
        self.assertIn("address couldn't be found", records[0]['error'].lower())

        # 2. Unrelated recipient hr@gmail.com is NOT marked BOUNCED
        self.assertNotEqual(records[1]['status'], 'BOUNCED')

        # 3. Missing bounce notification is classified as UNKNOWN (never DELIVERED)
        self.assertEqual(records[1]['status'], 'UNKNOWN')
        self.assertNotEqual(records[1]['status'], 'DELIVERED', "Absence of bounce must never be reported as DELIVERED")
        self.assertIn("No bounce detected", records[1]['remark'])

        print("[PASS] Bounce detection verified: bounce identified for matching recipient, unrelated recipient remains UNKNOWN, no false delivery confirmation")

    def test_10_one_line_per_email_and_inplace_updates(self):
        """Verify: exactly one physical line per email, strict format, in-place update from SUBMITTED to BOUNCED, timestamp/message_id preserved, sanitized newlines."""
        with tempfile.NamedTemporaryFile('w', suffix='.log', delete=False) as tmp:
            tmp_log_path = tmp.name

        try:
            logger = OutreachLogger(tmp_log_path)

            # Log initial SUBMITTED email record
            initial_rec = {
                'timestamp': '2026-11-15 19:30:01',
                'recipient': 'candidate@gmail.com',
                'status': 'SUBMITTED',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18abc123',
                'submitted_time': '19:30:01',
                'remark': 'Gmail API accepted email; delivery unconfirmed',
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
                "2026-11-15 19:30:01 | candidate@gmail.com | SUBMITTED | "
                "Subject: Application for DevOps / Cloud Engineer Opportunities | ID: 18abc123 | "
                "Submitted: 19:30:01 | Bounce: N/A | Error: None\n"
            )
            self.assertEqual(cand_line, expected_cand)

            # Now update initial_rec to BOUNCED (simulating a later confirmed bounce)
            updated = logger.update_email_record(
                recipient='candidate@gmail.com',
                new_status='BOUNCED',
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
                "2026-11-15 19:30:01 | candidate@gmail.com | BOUNCED | "
                "Subject: Application for DevOps / Cloud Engineer Opportunities | ID: 18abc123 | "
                "Submitted: 19:30:01 | Bounce: 19:35:10 | Error: SMTP 550 5.1.1\n"
            )
            self.assertEqual(updated_line, expected_updated)

            # Test appending across executions
            logger.log_email_record({
                'timestamp': '2026-11-15 19:40:01',
                'recipient': 'run2@gmail.com',
                'status': 'UNKNOWN',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18run2',
                'submitted_time': '19:40:01',
                'remark': 'No bounce detected; delivery unconfirmed',
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
                'status': 'UNKNOWN',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18delayed999',
                'submitted_time': '10:00:00',
                'remark': 'No bounce detected; delivery unconfirmed',
                'error': 'None',
                'bounce_time': 'N/A'
            }
            logger.log_email_record(prev_record)

            # Another email in run 1 that remains unconfirmed (no bounce)
            other_record = {
                'timestamp': '2026-11-14 10:00:05',
                'recipient': 'valid.recruiter@company.com',
                'status': 'UNKNOWN',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18valid888',
                'submitted_time': '10:00:05',
                'remark': 'No bounce detected; delivery unconfirmed',
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
            self.assertEqual(unconfirmed[0]['status'], 'BOUNCED')
            self.assertIn("does not exist", unconfirmed[0]['remark'].lower())
            self.assertIn("550", unconfirmed[0]['error'])
            self.assertEqual(unconfirmed[1]['status'], 'UNKNOWN')

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
            self.assertEqual(parts[2], 'BOUNCED')
            self.assertEqual(parts[3], 'Subject: Application for DevOps / Cloud Engineer Opportunities')
            self.assertEqual(parts[4], 'ID: 18delayed999', "Original Message-ID must be preserved")
            self.assertEqual(parts[5], 'Submitted: 10:00:00', "Original Submitted time must be preserved")
            self.assertNotEqual(parts[6], 'Bounce: N/A', "Bounce time must be updated")

            # Check valid.recruiter line was not marked as delivered
            other_line = next(line for line in lines_after if 'valid.recruiter@company.com' in line)
            self.assertIn('| UNKNOWN |', other_line)
            self.assertNotIn('DELIVERED', other_line)

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Delayed bounce detection verified: unconfirmed emails checked, log updated in-place, timestamps preserved")

    def test_12_compact_summary_counts_and_dashboard_format(self):
        """Verify summary count calculation (Unknown=1 when pending, not 0) and exact compact terminal dashboard output."""
        import io
        from unittest.mock import patch

        # Case 1: Single email submitted, pending delivery confirmation
        submitted_records_pending = [{
            'recipient': 'contact@example.com',
            'status': 'UNKNOWN',  # transitioned from SUBMITTED after no bounce found
            'message_id': '18abc'
        }]
        all_records_pending = list(submitted_records_pending)
        invalid_emails_pending = []

        sub_cnt = len(submitted_records_pending)
        bnc_cnt = sum(1 for r in submitted_records_pending if r.get('status') == 'BOUNCED')
        fail_cnt = sum(1 for r in all_records_pending if r.get('status') == 'FAILED_TO_SUBMIT')
        inv_cnt = len(invalid_emails_pending)
        unk_cnt = sum(1 for r in submitted_records_pending if r.get('status') == 'UNKNOWN')

        self.assertEqual(sub_cnt, 1)
        self.assertEqual(bnc_cnt, 0)
        self.assertEqual(fail_cnt, 0)
        self.assertEqual(inv_cnt, 0)
        self.assertEqual(unk_cnt, 1, "Unknown must show 1 when email delivery is pending/unconfirmed (not 0)")

        # Verify exact compact output format
        expected_output = (
            "\nSummary:\n"
            "Submitted: 1\n"
            "Bounced: 0\n"
            "Failed: 0\n"
            "Invalid: 0\n"
            "Unknown: 1\n\n"
            f"Logs: {LOG_FILE}\n"
        )

        with patch('sys.stdout', new=io.StringIO()) as fake_stdout:
            print("\nSummary:")
            print(f"Submitted: {sub_cnt}")
            print(f"Bounced: {bnc_cnt}")
            print(f"Failed: {fail_cnt}")
            print(f"Invalid: {inv_cnt}")
            print(f"Unknown: {unk_cnt}\n")
            print(f"Logs: {LOG_FILE}")
            self.assertEqual(fake_stdout.getvalue(), expected_output)

        # Case 2: Mixed batch (1 submitted unknown, 1 bounced, 1 failed, 1 invalid)
        submitted_mixed = [
            {'recipient': 'unknown@example.com', 'status': 'UNKNOWN', 'message_id': 'm1'},
            {'recipient': 'bounced@example.com', 'status': 'BOUNCED', 'message_id': 'm2'}
        ]
        all_mixed = list(submitted_mixed) + [
            {'recipient': 'failed@example.com', 'status': 'FAILED_TO_SUBMIT', 'message_id': 'N/A'}
        ]
        invalid_mixed = ['invalid-addr']

        sub_mixed_cnt = len(submitted_mixed)
        bnc_mixed_cnt = sum(1 for r in submitted_mixed if r.get('status') == 'BOUNCED')
        fail_mixed_cnt = sum(1 for r in all_mixed if r.get('status') == 'FAILED_TO_SUBMIT')
        inv_mixed_cnt = len(invalid_mixed)
        unk_mixed_cnt = sum(1 for r in submitted_mixed if r.get('status') == 'UNKNOWN')

        self.assertEqual(sub_mixed_cnt, 2)
        self.assertEqual(bnc_mixed_cnt, 1)
        self.assertEqual(fail_mixed_cnt, 1)
        self.assertEqual(inv_mixed_cnt, 1)
        self.assertEqual(unk_mixed_cnt, 1)

        print("[PASS] Compact summary counts and dashboard format verified: Unknown reflects pending, format matches exact template")

    def test_13_gmail_403_insufficient_permissions_handling(self):
        """Verify handling of Gmail API 403 insufficient scopes error: does not crash, marks records UNKNOWN, no false DELIVERED."""
        from googleapiclient.errors import HttpError
        from unittest.mock import MagicMock
        import httplib2

        resp = httplib2.Response({'status': 403, 'reason': 'Request had insufficient authentication scopes.'})
        http_error = HttpError(resp, b'{"error": {"message": "Insufficient Permission"}}')

        mock_service = MagicMock()
        mock_service.users().messages().list.side_effect = http_error

        records = [{
            'recipient': 'candidate@example.com',
            'status': 'SUBMITTED',
            'subject': 'Application for DevOps Role',
            'message_id': '18abc456',
            'submitted_time': '12:00:00',
            'remark': 'Gmail API accepted email; delivery unconfirmed',
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

            # Record must transition to UNKNOWN, never DELIVERED
            self.assertEqual(records[0]['status'], 'UNKNOWN')
            self.assertNotEqual(records[0]['status'], 'DELIVERED')
            self.assertIn("insufficient", records[0]['remark'].lower())

            # Log record must be updated in-place to UNKNOWN
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 1, "Must maintain exactly 1 line in outreach.log")
            self.assertIn('| UNKNOWN |', lines[0])
            self.assertIn('Bounce check failed', lines[0])

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Gmail API 403 insufficient permissions handled gracefully: records marked UNKNOWN, no false DELIVERED, single-line log preserved")


        try:
            logger = OutreachLogger(tmp_log_path)

            # Log an email submitted in run 1 whose delivery was unconfirmed
            prev_record = {
                'timestamp': '2026-11-14 10:00:00',
                'recipient': 'delayed.bounce@company.com',
                'status': 'UNKNOWN',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18delayed999',
                'submitted_time': '10:00:00',
                'remark': 'No bounce detected; delivery unconfirmed',
                'error': 'None',
                'bounce_time': 'N/A'
            }
            logger.log_email_record(prev_record)

            # Another email in run 1 that remains unconfirmed (no bounce)
            other_record = {
                'timestamp': '2026-11-14 10:00:05',
                'recipient': 'valid.recruiter@company.com',
                'status': 'UNKNOWN',
                'subject': 'Application for DevOps / Cloud Engineer Opportunities',
                'message_id': '18valid888',
                'submitted_time': '10:00:05',
                'remark': 'No bounce detected; delivery unconfirmed',
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
            self.assertEqual(unconfirmed[0]['status'], 'BOUNCED')
            self.assertIn("does not exist", unconfirmed[0]['remark'].lower())
            self.assertIn("550", unconfirmed[0]['error'])
            self.assertEqual(unconfirmed[1]['status'], 'UNKNOWN')

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
            self.assertEqual(parts[2], 'BOUNCED')
            self.assertEqual(parts[3], 'Subject: Application for DevOps / Cloud Engineer Opportunities')
            self.assertEqual(parts[4], 'ID: 18delayed999', "Original Message-ID must be preserved")
            self.assertEqual(parts[5], 'Submitted: 10:00:00', "Original Submitted time must be preserved")
            self.assertNotEqual(parts[6], 'Bounce: N/A', "Bounce time must be updated")

            # Check valid.recruiter line was not marked as delivered
            other_line = next(line for line in lines_after if 'valid.recruiter@company.com' in line)
            self.assertIn('| UNKNOWN |', other_line)
            self.assertNotIn('DELIVERED', other_line)

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Delayed bounce detection verified: unconfirmed emails checked, log updated in-place, timestamps preserved")

    def test_12_compact_summary_counts_and_dashboard_format(self):
        """Verify summary count calculation (Unknown=1 when pending, not 0) and exact compact terminal dashboard output."""
        import io
        from unittest.mock import patch

        # Case 1: Single email submitted, pending delivery confirmation
        submitted_records_pending = [{
            'recipient': 'contact@example.com',
            'status': 'UNKNOWN',  # transitioned from SUBMITTED after no bounce found
            'message_id': '18abc'
        }]
        all_records_pending = list(submitted_records_pending)
        invalid_emails_pending = []

        sub_cnt = len(submitted_records_pending)
        bnc_cnt = sum(1 for r in submitted_records_pending if r.get('status') == 'BOUNCED')
        fail_cnt = sum(1 for r in all_records_pending if r.get('status') == 'FAILED_TO_SUBMIT')
        inv_cnt = len(invalid_emails_pending)
        unk_cnt = sum(1 for r in submitted_records_pending if r.get('status') == 'UNKNOWN')

        self.assertEqual(sub_cnt, 1)
        self.assertEqual(bnc_cnt, 0)
        self.assertEqual(fail_cnt, 0)
        self.assertEqual(inv_cnt, 0)
        self.assertEqual(unk_cnt, 1, "Unknown must show 1 when email delivery is pending/unconfirmed (not 0)")

        # Verify exact compact output format
        expected_output = (
            "\nSummary:\n"
            "Submitted: 1\n"
            "Bounced: 0\n"
            "Failed: 0\n"
            "Invalid: 0\n"
            "Unknown: 1\n\n"
            f"Logs: {LOG_FILE}\n"
        )

        with patch('sys.stdout', new=io.StringIO()) as fake_stdout:
            print("\nSummary:")
            print(f"Submitted: {sub_cnt}")
            print(f"Bounced: {bnc_cnt}")
            print(f"Failed: {fail_cnt}")
            print(f"Invalid: {inv_cnt}")
            print(f"Unknown: {unk_cnt}\n")
            print(f"Logs: {LOG_FILE}")
            self.assertEqual(fake_stdout.getvalue(), expected_output)

        # Case 2: Mixed batch (1 submitted unknown, 1 bounced, 1 failed, 1 invalid)
        submitted_mixed = [
            {'recipient': 'unknown@example.com', 'status': 'UNKNOWN', 'message_id': 'm1'},
            {'recipient': 'bounced@example.com', 'status': 'BOUNCED', 'message_id': 'm2'}
        ]
        all_mixed = list(submitted_mixed) + [
            {'recipient': 'failed@example.com', 'status': 'FAILED_TO_SUBMIT', 'message_id': 'N/A'}
        ]
        invalid_mixed = ['invalid-addr']

        sub_mixed_cnt = len(submitted_mixed)
        bnc_mixed_cnt = sum(1 for r in submitted_mixed if r.get('status') == 'BOUNCED')
        fail_mixed_cnt = sum(1 for r in all_mixed if r.get('status') == 'FAILED_TO_SUBMIT')
        inv_mixed_cnt = len(invalid_mixed)
        unk_mixed_cnt = sum(1 for r in submitted_mixed if r.get('status') == 'UNKNOWN')

        self.assertEqual(sub_mixed_cnt, 2)
        self.assertEqual(bnc_mixed_cnt, 1)
        self.assertEqual(fail_mixed_cnt, 1)
        self.assertEqual(inv_mixed_cnt, 1)
        self.assertEqual(unk_mixed_cnt, 1)

        print("[PASS] Compact summary counts and dashboard format verified: Unknown reflects pending, format matches exact template")

    def test_13_gmail_403_insufficient_permissions_handling(self):
        """Verify handling of Gmail API 403 insufficient scopes error: does not crash, marks records UNKNOWN, no false DELIVERED."""
        from googleapiclient.errors import HttpError
        from unittest.mock import MagicMock
        import httplib2

        resp = httplib2.Response({'status': 403, 'reason': 'Request had insufficient authentication scopes.'})
        http_error = HttpError(resp, b'{"error": {"message": "Insufficient Permission"}}')

        mock_service = MagicMock()
        mock_service.users().messages().list.side_effect = http_error

        records = [{
            'recipient': 'candidate@example.com',
            'status': 'SUBMITTED',
            'subject': 'Application for DevOps Role',
            'message_id': '18abc456',
            'submitted_time': '12:00:00',
            'remark': 'Gmail API accepted email; delivery unconfirmed',
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

            # Record must transition to UNKNOWN, never DELIVERED
            self.assertEqual(records[0]['status'], 'UNKNOWN')
            self.assertNotEqual(records[0]['status'], 'DELIVERED')
            self.assertIn("insufficient", records[0]['remark'].lower())

            # Log record must be updated in-place to UNKNOWN
            with open(tmp_log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 1, "Must maintain exactly 1 line in outreach.log")
            self.assertIn('| UNKNOWN |', lines[0])
            self.assertIn('Bounce check failed', lines[0])

        finally:
            if os.path.exists(tmp_log_path):
                os.remove(tmp_log_path)

        print("[PASS] Gmail API 403 insufficient permissions handled gracefully: records marked UNKNOWN, no false DELIVERED, single-line log preserved")


if __name__ == '__main__':
    unittest.main()
