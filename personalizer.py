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

Authentication — three modes (tried in order)
---------------------------------------------
1. Gmail API   credentials.json present → uses GCP OAuth (full API)
2. App Password  .env present with GMAIL_ADDRESS + GMAIL_APP_PASSWORD
                 → connects via IMAP, appends to Drafts folder
                 No GCP project needed. Requires Gmail 2-Step Verification.
3. Preview only  neither configured → saves .txt files to drafts_preview/

Usage
-----
1. Fill in recipients.xlsx with columns: title, first_name, last_name, email
2. Edit template.txt with your draft.  Use {dear} where the salutation goes.
3. Add .env (App Password) or credentials.json (GCP) — see README.
4. Run:  python personalizer.py

Requirements
------------
    pip install openpyxl python-dotenv
    (+ google-auth-oauthlib google-auth-httplib2 google-api-python-client
       only if using GCP / credentials.json)
"""

import base64
import email.utils
import imaplib
import os
import re
import sys
from email.mime.text import MIMEText

# ── Optional: Gmail API ───────────────────────────────────────────────────────
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    GMAIL_API_AVAILABLE = True
except ImportError:
    GMAIL_API_AVAILABLE = False

# ── Optional: .env file support ───────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False

# ── Optional: Excel support ───────────────────────────────────────────────────
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ── Constants ─────────────────────────────────────────────────────────────────
SCOPES          = ["https://www.googleapis.com/auth/gmail.compose"]
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE       = "token.json"
TEMPLATE_FILE    = "template.txt"
RECIPIENTS_FILE  = "recipients.xlsx"
OUTPUT_DIR       = "drafts_preview"


# ── Helpers ───────────────────────────────────────────────────────────────────

def build_salutation(title: str, first_name: str, last_name: str) -> str:
    title = title.strip()
    if title.lower() == "dr":
        return f"Dr {last_name.strip()}"
    return f"{title} {first_name.strip()}"


def parse_template(template_path: str) -> tuple[str, str]:
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

    return subject, "".join(body_lines).lstrip("\n")


def personalise(body: str, title: str, first_name: str, last_name: str) -> str:
    return body.replace("{dear}", build_salutation(title, first_name, last_name))


def load_recipients(xlsx_path: str) -> list[dict]:
    if not OPENPYXL_AVAILABLE:
        sys.exit("openpyxl not installed. Run: pip install openpyxl")

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        sys.exit(f"{xlsx_path} is empty.")

    headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    missing = {"title", "first_name", "last_name", "email"} - set(headers)
    if missing:
        sys.exit(f"Excel sheet is missing columns: {missing}\nFound: {headers}")

    recipients = []
    for row in rows[1:]:
        record = dict(zip(headers, (str(v).strip() if v is not None else "" for v in row)))
        if record.get("email"):
            recipients.append(record)
    return recipients


# ── Mode 1: Gmail API (credentials.json) ─────────────────────────────────────

def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as fh:
            fh.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


def create_draft_via_api(service, to_email: str, subject: str, body: str) -> str:
    msg = MIMEText(body, "plain")
    msg["to"] = to_email
    msg["subject"] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    draft = service.users().drafts().create(
        userId="me", body={"message": {"raw": raw}}
    ).execute()
    return draft["id"]


# ── Mode 2: App Password via IMAP ─────────────────────────────────────────────

def get_imap_connection(gmail_address: str, app_password: str) -> imaplib.IMAP4_SSL:
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(gmail_address, app_password)
    return conn


def create_draft_via_imap(conn: imaplib.IMAP4_SSL, to_email: str,
                           subject: str, body: str, from_email: str):
    msg = MIMEText(body, "plain", "utf-8")
    msg["To"] = to_email
    msg["From"] = from_email
    msg["Subject"] = subject
    msg["Date"] = email.utils.formatdate()

    # Append the message to the [Gmail]/Drafts folder
    conn.append(
        "[Gmail]/Drafts",
        "\\Draft",
        imaplib.Time2Internaldate(None),
        msg.as_bytes()
    )


# ── Mode 3: Local preview ─────────────────────────────────────────────────────

def save_preview(to_email: str, subject: str, body: str, index: int) -> str:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe = re.sub(r"[^\w@.\-]", "_", to_email)
    path = os.path.join(OUTPUT_DIR, f"{index:03d}_{safe}.txt")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(f"To: {to_email}\nSubject: {subject}\n{'-'*60}\n{body}")
    return path


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # 1. Template
    if not os.path.exists(TEMPLATE_FILE):
        sys.exit(f"Template file not found: {TEMPLATE_FILE}")
    subject, body_template = parse_template(TEMPLATE_FILE)
    if not subject:
        sys.exit("Template must include a 'Subject:' line.")
    print(f"Subject (same for all): {subject}\n")

    # 2. Recipients
    if not os.path.exists(RECIPIENTS_FILE):
        sys.exit(f"Recipients file not found: {RECIPIENTS_FILE}")
    recipients = load_recipients(RECIPIENTS_FILE)
    print(f"Loaded {len(recipients)} recipient(s)\n")

    # 3. Pick authentication mode
    use_api      = GMAIL_API_AVAILABLE and os.path.exists(CREDENTIALS_FILE)
    gmail_addr   = os.environ.get("GMAIL_ADDRESS", "").strip()
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "").strip()
    use_imap     = bool(gmail_addr and app_password)

    if use_api:
        print("Mode: Gmail API (credentials.json)\n")
        service = get_gmail_service()
        imap_conn = None
    elif use_imap:
        print(f"Mode: App Password / IMAP ({gmail_addr})\n")
        service = None
        imap_conn = get_imap_connection(gmail_addr, app_password)
    else:
        print("Mode: Local preview only (no Gmail credentials found)")
        print("  → To use App Password: create a .env file with:")
        print("      GMAIL_ADDRESS=you@gmail.com")
        print("      GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx\n")
        service = None
        imap_conn = None

    # 4. Process each recipient
    for i, row in enumerate(recipients, start=1):
        title      = row["title"].strip()
        first_name = row["first_name"].strip()
        last_name  = row["last_name"].strip()
        to_email   = row["email"].strip()

        body = personalise(body_template, title, first_name, last_name)
        salutation = build_salutation(title, first_name, last_name)

        if use_api:
            draft_id = create_draft_via_api(service, to_email, subject, body)
            print(f"[{i}] Draft created (API)  → {to_email}  |  Dear {salutation}  |  id:{draft_id}")
        elif use_imap:
            create_draft_via_imap(imap_conn, to_email, subject, body, gmail_addr)
            print(f"[{i}] Draft created (IMAP) → {to_email}  |  Dear {salutation}")
        else:
            path = save_preview(to_email, subject, body, i)
            print(f"[{i}] Preview saved        → {to_email}  |  Dear {salutation}  |  {path}")

    if imap_conn:
        imap_conn.logout()

    print(f"\nDone — {len(recipients)} draft(s) created.")


if __name__ == "__main__":
    main()
