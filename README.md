# Email Personalizer

Reads a draft email template and a CSV of recipients, then creates one personalised Gmail draft per recipient.

## Quick start

```bash
pip install -r requirements.txt
python personalizer.py
```

---

## 1. Set up your recipients (`recipients.csv`)

| Column | Description |
|--------|-------------|
| `title` | `Mr`, `Ms`, or `Dr` |
| `first_name` | Recipient's first name |
| `last_name` | Recipient's last name |
| `email` | Recipient's email address |

```csv
title,first_name,last_name,email
Mr,James,Smith,james.smith@example.com
Ms,Sarah,Johnson,sarah.johnson@example.com
Dr,Emily,Brown,emily.brown@example.com
```

---

## 2. Write your draft email (`template.txt`)

- The **first line** must be `Subject: ...` — this subject is used for **every** draft unchanged.
- Use `{dear}` exactly where you want the salutation to appear.

```
Subject: Annual Conference Invitation 2026

Dear {dear},

Your email body here...
```

### Salutation rules

| Title | Result |
|-------|--------|
| `Mr`  | `Dear Mr {first_name}` |
| `Ms`  | `Dear Ms {first_name}` |
| `Dr`  | `Dear Dr {last_name}`  |

---

## 3. Connect to Gmail (optional)

Without Gmail credentials the tool saves previews to `drafts_preview/` as `.txt` files so you can review them first.

To push directly to your Gmail Drafts folder:

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project → enable the **Gmail API**.
3. Create an **OAuth 2.0 Desktop** client → download `credentials.json`.
4. Place `credentials.json` in this folder.
5. Run `python personalizer.py` — a browser window will ask you to authorise once.

The token is saved to `token.json` for future runs (no browser needed again).

---

## File overview

```
.
├── personalizer.py      # main script
├── template.txt         # your draft email template
├── recipients.csv       # list of recipients
├── requirements.txt     # Python dependencies
├── credentials.json     # (you provide) Gmail OAuth client secret
├── token.json           # (auto-generated) saved Gmail token
└── drafts_preview/      # (auto-generated) local previews when Gmail not configured
```
