# Google Cloud OAuth 2.0 Credentials Setup

This directory holds the OAuth 2.0 Client credentials required to authenticate with the Gmail API.

---

## How to Get `credentials.json`

1. **Create a Google Cloud Project**:
   - Visit the [Google Cloud Console](https://console.cloud.google.com/).
   - Click the project dropdown at the top and select **New Project**.
   - Name your project (e.g., `Recruiter-Outreach-Agent`) and click **Create**.

2. **Enable Gmail API**:
   - In the navigation menu, go to **APIs & Services > Library**.
   - Search for **Gmail API**, click it, and click **Enable**.

3. **Configure OAuth Consent Screen**:
   - Navigate to **APIs & Services > OAuth consent screen**.
   - Choose **External** for User Type and click **Create**.
   - Enter an App Name (e.g., `Recruiter Outreach Agent`) and your email for support and developer contacts.
   - Under **Scopes**, click **Add or Remove Scopes**, search for `https://www.googleapis.com/auth/gmail.send`, select it, and click **Update**.
   - Under **Test Users**, add the personal Gmail address you will use to send emails.
   - Click **Save and Continue** to finish.

4. **Create OAuth Client ID**:
   - Go to **APIs & Services > Credentials**.
   - Click **+ Create Credentials** at the top and choose **OAuth client ID**.
   - Under **Application type**, choose **Desktop app**.
   - Set a name (e.g., `Outreach Desktop Client`) and click **Create**.
   - In the popup, click **Download JSON**.

5. **Place Credentials in this Directory**:
   - Rename the downloaded JSON file to `credentials.json`.
   - Place it in this folder:
     ```text
     config/credentials.json
     ```

> **Security Note:** Both `credentials.json` and the auto-generated `token.json` are listed in `.gitignore` and will never be committed to version control. Refer to `credentials.json.example` for the expected JSON structure.
