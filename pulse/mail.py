"""Microsoft Graph sendMail, the same shape CIP uses for platform mail.

Credentials stay in the environment. When they are missing, schedules still save
and the caller is told that email was not sent.
"""
import base64
import json
import os
import re

import requests

from .bc_client import DATA_DIR

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PREFS = DATA_DIR / "notification_mail.json"
_KEYS = ("PULSE_MAIL_TENANT_ID", "PULSE_MAIL_CLIENT_ID", "PULSE_MAIL_CLIENT_SECRET", "PULSE_MAIL_SENDER")


def configured():
    return all((os.environ.get(k) or "").strip() for k in _KEYS)


def clean_recipients(addresses):
    out, seen = [], set()
    for raw in addresses or []:
        addr = str(raw or "").strip()
        key = addr.lower()
        if EMAIL_RE.match(addr) and key not in seen:
            seen.add(key)
            out.append(addr)
        if len(out) >= 20:
            break
    return out


def load_prefs():
    try:
        data = json.loads(PREFS.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    data.setdefault("email", "")
    data.setdefault("refresh", False)
    data.setdefault("reports", False)
    return data


def save_prefs(data):
    emails = clean_recipients([data.get("email")])
    PREFS.parent.mkdir(parents=True, exist_ok=True)
    PREFS.write_text(json.dumps({
        "email": emails[0] if emails else "",
        "refresh": bool(data.get("refresh")),
        "reports": bool(data.get("reports")),
    }, indent=2), encoding="utf-8")
    return load_prefs()


def _token():
    tenant = os.environ["PULSE_MAIL_TENANT_ID"].strip()
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={"client_id": os.environ["PULSE_MAIL_CLIENT_ID"].strip(),
              "client_secret": os.environ["PULSE_MAIL_CLIENT_SECRET"].strip(),
              "scope": "https://graph.microsoft.com/.default",
              "grant_type": "client_credentials"},
        timeout=30)
    if resp.status_code >= 400:
        return None, f"mail token failed ({resp.status_code})"
    return resp.json().get("access_token"), None


def send(to, subject, html, attachment=None):
    """attachment is (filename, bytes, content_type) or None. Returns 'sent' or a short error."""
    recipients = clean_recipients(to)
    if not recipients:
        return "no recipients"
    if not configured():
        return "mail sender is not configured"
    token, err = _token()
    if err:
        return err
    sender = os.environ["PULSE_MAIL_SENDER"].strip()
    message = {
        "subject": subject[:240],
        "body": {"contentType": "HTML", "content": html},
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in recipients],
    }
    if attachment:
        name, content, ctype = attachment
        message["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": ctype,
            "contentBytes": base64.b64encode(content).decode("ascii"),
        }]
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": message, "saveToSentItems": True},
        timeout=60)
    if resp.status_code >= 400:
        return f"mail failed ({resp.status_code})"
    return "sent"


def notice(kind, subject, html):
    prefs = load_prefs()
    if not prefs.get(kind) or not prefs.get("email"):
        return "off"
    return send([prefs["email"]], subject, html)
