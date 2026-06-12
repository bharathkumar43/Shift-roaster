"""
Parse drive-change alert email bodies into structured alert records.
"""
import re
from datetime import datetime, timezone, timedelta
from html.parser import HTMLParser

IST_OFFSET = timedelta(hours=5, minutes=30)

WINDOWS = {
    1: {"label": "1:00 AM – 1:00 PM IST",  "start_h": 1,  "end_h": 13},
    2: {"label": "1:00 PM – 1:00 AM IST",  "start_h": 13, "end_h": 25},
}


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self):
        return " ".join(self._parts)


def _strip_html(html_body: str) -> str:
    s = _HTMLStripper()
    s.feed(html_body)
    return s.get_text()


def _ist_hour(received_at: str) -> int:
    try:
        dt = datetime.fromisoformat(received_at.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return 0
    return dt.astimezone(timezone(IST_OFFSET)).hour


def classify_window(received_at: str) -> int:
    h = _ist_hour(received_at)
    return 1 if 1 <= h < 13 else 2


def _find(patterns: list[str], text: str) -> str:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def _normalize_status_type(status: str) -> str:
    """Return the first word of the status string, uppercased (e.g. 'Conflict details' → 'CONFLICT')."""
    if not status:
        return ""
    return status.strip().split()[0].upper()


def _extract_fields(text: str) -> dict:
    workspace_id = _find([
        r"workspace[\s_-]*id[\s:=]+([^\s\n,;<]+)",
        r"workspace[\s:=]+([A-Za-z0-9_\-]+)",
        r"\bwsid[\s:=]+([^\s\n,;<]+)",
    ], text)

    server_name = _find([
        r"server[\s_-]*name[\s:=]+(\S+)",
        r"source[\s_-]*server[\s:=]+(\S+)",
        r"server[\s:=]+(\S+)",
    ], text)

    client_name = _find([
        r"client[\s_-]*name[\s:=]+(\S+)",
        r"customer[\s_-]*name[\s:=]+(\S+)",
        r"project[\s_-]*name[\s:=]+(\S+)",
        r"client[\s:=]+(\S+)",
        r"project[\s:=]+(\S+)",
        r"customer[\s:=]+(\S+)",
        r"account[\s:=]+(\S+)",
        r"tenant[\s:=]+(\S+)",
    ], text)

    status = _find([
        r"status[\s:=]+([^\n,;<]+)",
        r"issue[\s:=]+([^\n,;<]+)",
        r"error[\s:=]+([^\n,;<]+)",
        r"reason[\s:=]+([^\n,;<]+)",
    ], text)

    return {
        "workspace_id": workspace_id,
        "server_name":  server_name,
        "client_name":  client_name,
        "status":       status,
    }


def parse_email(email: dict) -> dict:
    body = email.get("body", "")
    content_type = email.get("content_type", "text")
    text = _strip_html(body) if "html" in content_type.lower() else body

    fields = _extract_fields(text)
    window = classify_window(email.get("received_at", ""))

    return {
        "id":           email["id"],
        "received_at":  email.get("received_at", ""),
        "subject":      email.get("subject", ""),
        "window":       window,
        "workspace_id": fields["workspace_id"],
        "server_name":  fields["server_name"],
        "client_name":  fields["client_name"],
        "status":       fields["status"],
        "status_type":  _normalize_status_type(fields["status"]),
    }


def parse_all(emails: list[dict]) -> list[dict]:
    return [parse_email(e) for e in emails]
