import os
import sys
import unittest
import base64
import email
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
    DEFAULT_CSV,
    DEFAULT_TEMPLATE,
    DEFAULT_RESUME,
    DEFAULT_SUBJECT,
    SCOPES,
    CREDENTIALS_FILE
)


class TestRecruiterOutreachAgent(unittest.TestCase):

    def test_01_resume_pdf_validity(self):
        """Verify that resume.pdf exists, is non-empty, and has a valid PDF header."""
        self.assertTrue(os.path.exists(DEFAULT_RESUME), "resume.pdf must exist")
        size = os.path.getsize(DEFAULT_RESUME)
        self.assertGreater(size, 1000, "resume.pdf should be a real PDF file (>1KB)")
        with open(DEFAULT_RESUME, 'rb') as f:
            header = f.read(5)
            self.assertEqual(header, b'%PDF-', "resume.pdf must start with '%PDF-' header")
        print(f"\n[PASS] Resume file validated: '{DEFAULT_RESUME}' ({size} bytes, valid PDF)")

    def test_02_credentials_exist(self):
        """Verify that credentials.json is present and has desktop client config."""
        self.assertTrue(os.path.exists(CREDENTIALS_FILE), f"{CREDENTIALS_FILE} must be present")
        import json
        with open(CREDENTIALS_FILE, 'r') as f:
            creds = json.load(f)
        self.assertIn('installed', creds, "credentials.json should contain 'installed' desktop client config")
        print(f"[PASS] OAuth client credentials file present and format validated")

    def test_03_email_validation(self):
        """Test validation of email address format."""
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

        # Also test with a test CSV buffer to verify parsing logic deterministically
        import tempfile
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
        """Verify email generation, HTML body, and PDF attachment without sending."""
        self.assertTrue(os.path.exists(DEFAULT_TEMPLATE), "template.html must exist")
        with open(DEFAULT_TEMPLATE, 'r', encoding='utf-8') as f:
            html_content = f.read()

        test_recipient = "test.recruiter@example.com"
        message_dict = create_message(
            to_email=test_recipient,
            subject=DEFAULT_SUBJECT,
            html_content=html_content,
            resume_path=DEFAULT_RESUME
        )

        self.assertIn('raw', message_dict, "Generated message dict must have 'raw' key")
        raw_bytes = base64.urlsafe_b64decode(message_dict['raw'].encode('utf-8'))

        # Parse raw MIME message
        parsed_msg = email.message_from_bytes(raw_bytes)
        self.assertEqual(parsed_msg['To'], test_recipient)
        self.assertEqual(parsed_msg['Subject'], DEFAULT_SUBJECT)
        self.assertIsNone(parsed_msg['Cc'], "Cc must be empty")
        self.assertIsNone(parsed_msg['Bcc'], "Bcc must be empty")

        # Check payload parts
        parts = list(parsed_msg.walk())
        part_content_types = [p.get_content_type() for p in parts]
        self.assertIn('text/html', part_content_types, "Message must contain text/html part")
        self.assertIn('application/pdf', part_content_types, "Message must contain application/pdf attachment")

        # Verify attachment filename and payload
        pdf_part = next(p for p in parts if p.get_content_type() == 'application/pdf')
        expected_filename = os.path.basename(DEFAULT_RESUME)
        self.assertEqual(pdf_part.get_filename(), expected_filename)
        self.assertTrue(expected_filename.lower().endswith('.pdf'))
        pdf_data = pdf_part.get_payload(decode=True)
        self.assertTrue(pdf_data.startswith(b'%PDF-'), "Attached content must be a valid PDF binary")
        self.assertEqual(len(pdf_data), os.path.getsize(DEFAULT_RESUME))

        print(f"[PASS] MIME message verified: HTML body present, {expected_filename} attached ({len(pdf_data)} bytes), recipient '{test_recipient}' set individually")

    def test_06_oauth_scope_configuration(self):
        """Verify required Gmail send scope is configured in the agent."""
        self.assertIn('https://www.googleapis.com/auth/gmail.send', SCOPES)
        print(f"[PASS] OAuth scope verified: {SCOPES}")


if __name__ == '__main__':
    unittest.main()

