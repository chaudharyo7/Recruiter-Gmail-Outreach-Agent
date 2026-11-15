# Resume Attachment Directory

Place your resume PDF file in this directory.

---

## Instructions

1. **Copy your resume here**:
   Copy your PDF resume into this folder and name it:
   ```text
   resume.pdf
   ```
   *(e.g., `attachments/resume.pdf`)*

2. **Custom Filename**:
   If you want to use a custom filename (e.g., `John_Doe_Resume.pdf`):
   - Place `John_Doe_Resume.pdf` inside `attachments/`.
   - Update `RESUME_FILE` in your `.env` file (or set the environment variable):
     ```bash
     RESUME_FILE=attachments/John_Doe_Resume.pdf
     ```

> **Security Note:** All `*.pdf` files in this directory are excluded by `.gitignore` to ensure personal resumes are never committed to version control.

