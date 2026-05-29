"""
Group parsed drive-change alerts by IST time window and server/project domain.
Workspace IDs are deduplicated within each project group.
Also provides a combined per-project view across all windows.
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
    Groups by server_name per window. Returns per-window breakdown AND a
    combined per-project list (one entry per project across all windows).
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

    # Per-window result (used for window-level totals in stats bar)
    result: dict = {"parse_failures": parse_failures}
    for w_num, w_info in WINDOWS.items():
        groups_sorted = sorted(buckets[w_num].values(), key=lambda x: -len(x["workspace_ids"]))
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

    # Combined per-project view: merge workspace IDs across all windows
    combined: dict[str, dict] = {}
    for w_num in WINDOWS:
        for proj in result[w_num]["groups"]:
            key = proj["server_name"] or proj["project_name"]
            if key not in combined:
                combined[key] = {
                    "server_name":     proj["server_name"],
                    "project_name":    proj["project_name"],
                    "client_name":     proj["client_name"],
                    "workspace_ids":   [],
                    "_ws_seen":        set(),
                    "status_counts":   defaultdict(int),
                    "count":           0,
                    "window_counts":   {1: 0, 2: 0, 3: 0},
                }
            c = combined[key]
            for ws in proj["workspace_ids"]:
                if ws not in c["_ws_seen"]:
                    c["_ws_seen"].add(ws)
                    c["workspace_ids"].append(ws)
            for st, n in proj["status_counts"].items():
                c["status_counts"][st] += n
            c["count"] += proj["count"]
            c["window_counts"][w_num] += proj["alert_count"]
            if not c["client_name"] and proj["client_name"]:
                c["client_name"] = proj["client_name"]

    projects_out = []
    for c in sorted(combined.values(), key=lambda x: -len(x["workspace_ids"])):
        sc = dict(c["status_counts"])
        primary = max(sc.items(), key=lambda x: x[1])[0] if sc else ""
        projects_out.append({
            "server_name":    c["server_name"],
            "project_name":   c["project_name"],
            "client_name":    c["client_name"],
            "workspace_ids":  c["workspace_ids"],
            "alert_count":    len(c["workspace_ids"]),
            "count":          c["count"],
            "status_counts":  sc,
            "primary_status": primary,
            "window_counts":  c["window_counts"],
        })

    result["projects"] = projects_out
    return result
