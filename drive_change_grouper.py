"""
Group parsed drive-change alerts by IST time window and server/project domain.
Workspace IDs are deduplicated within each project group.
"""
from collections import defaultdict
from drive_change_parser import WINDOWS


def _project_name_from_server(server: str) -> str:
    """'vaticahealth.cloudfuze.com' → 'VATICAHEALTH'"""
    if not server:
        return ""
    name = server.lower()
    for suffix in [".cloudfuze.com", ".com", ".net", ".org", ".io", ".co.uk"]:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
            break
    return name.split(".")[0].upper()


def group_alerts(parsed_alerts: list[dict]) -> dict:
    """
    Groups by server_name (project domain). Returns:
    {
        "parse_failures": int,
        1: {
            "label":  "12:00 AM – 8:00 AM IST",
            "groups": [
                {
                    "server_name":    str,
                    "project_name":   str,   # VATICAHEALTH
                    "client_name":    str,   # first seen user
                    "primary_status": str,   # CONFLICT
                    "status_counts":  {str: int},
                    "workspace_ids":  [str, ...],
                    "count":          int,   # total emails
                    "alert_count":    int,   # unique workspace IDs
                }
            ],
            "total": int,
        },
        2: { ... },
        3: { ... },
    }
    """
    buckets: dict[int, dict[str, dict]] = {w: {} for w in WINDOWS}
    parse_failures = 0

    for alert in parsed_alerts:
        ws_id = (alert.get("workspace_id") or "").strip()
        if not ws_id:
            parse_failures += 1

        w = alert.get("window", 1)
        server = (alert.get("server_name") or "").strip()
        key = server or "(unknown)"

        if key not in buckets[w]:
            buckets[w][key] = {
                "server_name":   server,
                "project_name":  _project_name_from_server(server) or key.upper(),
                "client_name":   "",
                "workspace_ids": [],
                "_ws_seen":      set(),
                "status_counts": defaultdict(int),
                "count":         0,
            }

        g = buckets[w][key]
        g["count"] += 1

        st = (alert.get("status_type") or "").strip()
        if st:
            g["status_counts"][st] += 1

        if not g["client_name"]:
            g["client_name"] = (alert.get("client_name") or "").strip()

        if ws_id and ws_id not in g["_ws_seen"]:
            g["_ws_seen"].add(ws_id)
            g["workspace_ids"].append(ws_id)

    result: dict = {"parse_failures": parse_failures}

    for w_num, w_info in WINDOWS.items():
        groups_sorted = sorted(
            buckets[w_num].values(),
            key=lambda x: -len(x["workspace_ids"]),
        )
        groups = []
        for data in groups_sorted:
            sc = dict(data["status_counts"])
            primary = max(sc.items(), key=lambda x: x[1])[0] if sc else ""
            groups.append({
                "server_name":    data["server_name"],
                "project_name":   data["project_name"],
                "client_name":    data["client_name"],
                "workspace_ids":  data["workspace_ids"],
                "alert_count":    len(data["workspace_ids"]),
                "count":          data["count"],
                "status_counts":  sc,
                "primary_status": primary,
            })

        result[w_num] = {
            "label":  w_info["label"],
            "groups": groups,
            "total":  sum(g["alert_count"] for g in groups),
        }

    return result
