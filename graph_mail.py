"""Mail versturen via Microsoft Graph (app-registratie met Application-recht Mail.Send)."""
import base64
import json
import os
import urllib.parse
import urllib.request


def _token():
    tenant = os.environ["GRAPH_TENANT_ID"]
    data = urllib.parse.urlencode({
        "client_id": os.environ["GRAPH_CLIENT_ID"],
        "client_secret": os.environ["GRAPH_CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    req = urllib.request.Request(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", data=data)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())["access_token"]


def split_addresses(value):
    return [a.strip() for a in (value or "").replace(",", ";").split(";") if a.strip()]


def send(to, subject, body, attachments=()):
    """to: lijst adressen. attachments: [(bestandsnaam, bytes)]."""
    if not to:
        raise RuntimeError("Geen ontvangers opgegeven.")
    sender = os.environ["MAIL_FROM"]
    msg = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        "attachments": [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "contentBytes": base64.b64encode(content).decode(),
        } for name, content in attachments],
    }
    req = urllib.request.Request(
        f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(sender)}/sendMail",
        data=json.dumps({"message": msg, "saveToSentItems": True}).encode(),
        headers={"Authorization": f"Bearer {_token()}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        if r.status != 202:
            raise RuntimeError(f"Graph sendMail gaf status {r.status}")
