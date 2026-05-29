"""
Reads drive-change alert emails from Outlook via Microsoft Graph API.
Uses the same Azure app registration as the main app — ensure Mail.Read
application permission is granted and admin-consented in Azure portal.
"""

import os
import requests
import msal
from datetime import date as _date, datetime, timedelta, timezone

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"), override=False)

GRAPH_API = "https://graph.microsoft.com/v1.0"

DC_USER_EMAIL     = os.getenv("DC_OUTLOOK_USER_EMAIL", "")
DC_SUBJECT_FILTER = os.getenv("DC_EMAIL_SUBJECT_FILTER", "Important: Drive Changes Issue")

_IST = timezone(timedelta(hours=5, minutes=30))


def _get_token() -> str:
    client_id     = os.getenv("AZURE_CLIENT_ID", "")
    client_secret = os.getenv("AZURE_CLIENT_SECRET", "")
    tenant_id     = os.getenv("AZURE_TENANT_ID", "")
    if not (client_id and client_secret and tenant_id):
        raise EnvironmentError(
            "Azure credentials missing. Set AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID in .env"
        )
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise ConnectionError(
            f"Azure token acquisition failed: {result.get('error_description', 'unknown error')}"
        )
    return result["access_token"]


def fetch_drive_change_emails(date_str: str | None = None) -> list[dict]:
    """
    Fetch drive-change alert emails for a specific IST date (YYYY-MM-DD).
    Defaults to today in IST when date_str is omitted.
    Returns list of raw email dicts sorted newest-first.
    Raises EnvironmentError if DC_OUTLOOK_USER_EMAIL is not configured.
    """
    if not DC_USER_EMAIL:
        raise EnvironmentError(
            "DC_OUTLOOK_USER_EMAIL is not set in .env — cannot read drive-change emails."
        )

    # Resolve the IST date
    if date_str:
        ist_date = datetime.fromisoformat(date_str).date()
    else:
        ist_date = datetime.now(_IST).date()

    # Convert IST midnight → midnight-next-day to UTC bounds
    day_start = datetime(ist_date.year, ist_date.month, ist_date.day, tzinfo=_IST)
    day_end   = day_start + timedelta(days=1)
    since = day_start.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    until = day_end.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    token   = _get_token()
    headers = {"Authorization": f"Bearer {token}"}

    # NOTE: $orderby omitted intentionally — Graph rejects it with contains() filter.
    filter_q = (
        f"contains(subject, '{DC_SUBJECT_FILTER}') "
        f"and receivedDateTime ge {since} "
        f"and receivedDateTime lt {until}"
    )
    url = (
        f"{GRAPH_API}/users/{DC_USER_EMAIL}/messages"
        f"?$filter={filter_q}"
        f"&$select=id,subject,body,receivedDateTime,isRead"
        f"&$top=500"
    )

    emails = []
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        for msg in data.get("value", []):
            emails.append({
                "id":           msg["id"],
                "subject":      msg.get("subject", ""),
                "received_at":  msg.get("receivedDateTime", ""),
                "is_read":      msg.get("isRead", False),
                "body":         msg["body"]["content"],
                "content_type": msg["body"]["contentType"],
            })
        url = data.get("@odata.nextLink")

    emails.sort(key=lambda e: e["received_at"], reverse=True)
    return emails
