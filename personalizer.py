#!/usr/bin/env python3
"""
Email Personalizer
==================
Reads a draft email template and an Excel spreadsheet of recipients, then
creates one personalised Gmail draft per recipient.

Template placeholders
---------------------
{dear}    → replaced with the correct salutation + name (see below)

Salutation rules
----------------
Mr  → "Dear Mr {first_name}"
Ms  → "Dear Ms {first_name}"
Dr  → "Dear Dr {last_name}"

Subject
-------
The subject line is taken from the template (first line starting with
"Subject:") and is identical for every draft.

Usage
-----
1. Fill in recipients.xlsx with columns: title, first_name, last_name, email
2. Edit template.txt with your draft.  Use {dear} where the salutation goes.
3. Authenticate with Gmail (first run opens a browser):
       python personalizer.py
4. Check your Gmail Drafts folder.

Requirements
------------
    pip install openpyxl google-auth-oauthlib google-auth-httplib2 google-api-python-client
"""

import base64
import os
import re
import sys
from email.mime.text import MIMEText

# ── Gmail API imports ────────────────────────────────────────────────────────
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    GMAIL_AVAILABLE = True
except ImportError:
    GMAIL_AVAILABLE = False

# ── Gmail API imports ────────────────────────────────────────────────────────
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    GMAIL_AVAILABLE = True
except ImportError:
    GMAIL_AVAILABLE = False

# ── Excel import ─────────────────────────────────────────────────────────────
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ── Constants ────────────────────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]
CREDENTIALS_FILE = "credentials.json"   # OAuth client secret downloaded from GCP
TOKEN_FILE = "token.json"               # auto-generated after first login
TEMPLATE_FILE = "template.txt"
RECIPIENTS_FILE = "recipients.xlsx"
OUTPUT_DIR = "drafts_preview"           # used when Gmail is not configured


# ── Helpers ──────────────────────────────────────────────────────────────────

def build_salutation(title: str, first_name: str, last_name: str) -> str:
    """Return the correct salutation string for the {dear} placeholder."""
    title = title.strip()
    if title.lower() == "dr":
        return f"Dr {last_name.strip()}"
    return f"{title} {first_name.strip()}"


def parse_template(template_path: str) -> tuple[str, str]:
    """
    Read the template file and return (subject, body_without_subject_line).
    The subject is extracted from the first line that starts with 'Subject:'.
    """
    with open(template_path, "r", encoding="utf-8") as fh:
        content = fh.read()

    subject = ""
    body_lines = []
    subject_found = False

    for line in content.splitlines(keepends=True):
        if not subject_found and re.match(r"(?i)^subject\s*:", line):
            subject = re.sub(r"(?i)^subject\s*:\s*", "", line).strip()
            subject_found = True
        else:
            body_lines.append(line)

    body = "".join(body_lines).lstrip("\n")
    return subject, body


def personalise(body: str, title: str, first_name: str, last_name: str) -> str:
    """Replace {dear} in the body with the correct salutation."""
    salutation = build_salutation(title, first_name, last_name)
    return body.replace("{dear}", salutation)


def load_recipients(xlsx_path: str) -> list[dict]:
    """
    Load recipients from an Excel workbook (.xlsx).
    Reads the first sheet.  Required columns (case-insensitive header row):
        title, first_name, last_name, email
    Skips rows where the email cell is empty.
    """
    if not OPENPYXL_AVAILABLE:
        sys.exit("openpyxl is not installed. Run: pip install openpyxl")

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit(f"{xlsx_path} appears to be empty.")

    # First row = headers (normalise to lowercase, strip whitespace)
    headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]

    required = {"title", "first_name", "last_name", "email"}
    missing = required - set(headers)
    if missing:
        sys.exit(f"Excel sheet is missing columns: {missing}\n"
                 f"Found headers: {headers}")

    recipients = []
    for row in rows[1:]:
        record = dict(zip(headers, (str(v).strip() if v is not None else "" for v in row)))
        if record.get("email"):
            recipients.append(record)

    wb.close()
    return recipients


# ── Gmail draft creation ─────────────────────────────────────────────────────

def get_gmail_service():
    """Authenticate and return an authorised Gmail API service object."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def create_gmail_draft(service, to_email: str, subject: str, body: str) -> str:
    """Create a Gmail draft and return the draft ID."""
    message = MIMEText(body, "plain")
    message["to"] = to_email
    message["subject"] = subject

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    draft = service.users().drafts().create(
        userId="me",
        body={"message": {"raw": raw}}
    ).execute()
    return draft["id"]


# ── Local preview fallback ───────────────────────────────────────────────────

def save_preview(to_email: str, subject: str, body: str, index: int):
    """Save a personalised draft as a .txt file for review (no Gmail needed)."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_name = re.sub(r"[^\w@.\-]", "_", to_email)
    path = os.path.join(OUTPUT_DIR, f"{index:03d}_{safe_name}.txt")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(f"To: {to_email}\n")
        fh.write(f"Subject: {subject}\n")
        fh.write("-" * 60 + "\n")
        fh.write(body)
    return path


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    # 1. Parse template
    if not os.path.exists(TEMPLATE_FILE):
        sys.exit(f"Template file not found: {TEMPLATE_FILE}")
    subject, body_template = parse_template(TEMPLATE_FILE)
    if not subject:
        sys.exit("Template must include a 'Subject:' line.")
    print(f"Subject (same for all): {subject}\n")

    # 2. Load recipients
    if not os.path.exists(RECIPIENTS_FILE):
        sys.exit(f"Recipients file not found: {RECIPIENTS_FILE}")
    recipients = load_recipients(RECIPIENTS_FILE)
    print(f"Loaded {len(recipients)} recipient(s) from {RECIPIENTS_FILE}\n")

    # 3. Decide mode: Gmail or local preview
    use_gmail = GMAIL_AVAILABLE and os.path.exists(CREDENTIALS_FILE)
    if not use_gmail:
        if not GMAIL_AVAILABLE:
            print("Gmail libraries not installed — saving previews to ./drafts_preview/")
            print("Install with: pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client\n")
        else:
            print(f"'{CREDENTIALS_FILE}' not found — saving previews to ./{OUTPUT_DIR}/")
            print("To send to Gmail: download credentials.json from Google Cloud Console.\n")

    service = get_gmail_service() if use_gmail else None

    # 4. Process each recipient
    for i, row in enumerate(recipients, start=1):
        title      = row["title"].strip()
        first_name = row["first_name"].strip()
        last_name  = row["last_name"].strip()
        email      = row["email"].strip()

        personalised_body = personalise(body_template, title, first_name, last_name)
        salutation = build_salutation(title, first_name, last_name)

        if use_gmail:
            draft_id = create_gmail_draft(service, email, subject, personalised_body)
            print(f"[{i}] Draft created → To: {email}  |  Dear {salutation}  |  ID: {draft_id}")
        else:
            path = save_preview(email, subject, personalised_body, i)
            print(f"[{i}] Preview saved  → To: {email}  |  Dear {salutation}  |  File: {path}")

    print(f"\nDone — {len(recipients)} draft(s) created.")


if __name__ == "__main__":
    main()
