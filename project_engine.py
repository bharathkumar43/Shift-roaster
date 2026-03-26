import calendar
from datetime import date
from collections import defaultdict

from roster_engine import DAY_NAMES


def generate_project_coverage(projects, employees, shift_assignments, year, month):
    """
    Generate daily project coverage for the month.

    When the primary owner is on week-off, a takeover is assigned from
    the same shift AND same product type, picking the person with the
    fewest projects (least-projects-first).

    Args:
        projects: list of project dicts with 'name', 'product_type', 'employee_id', 'employee_name'
        employees: list of employee dicts with 'id', 'name', 'content_types', 'working_days'
        shift_assignments: dict mapping employee_name -> shift_num
        year, month: target period

    Returns:
        coverage: list of dicts, one per day, each containing:
            - date, day_name, day_abbr, day_num
            - projects: list of {project_name, product_type, owner, handler, is_takeover}
        warnings: list of warning strings
    """
    emp_lookup = {e["name"]: e for e in employees}
    emp_id_lookup = {e["id"]: e for e in employees}

    projects_by_owner = defaultdict(list)
    for p in projects:
        projects_by_owner[p["employee_name"]].append(p)

    num_days = calendar.monthrange(year, month)[1]
    coverage = []
    warnings = []

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]

        day_info = {
            "date": d.strftime("%b %d"),
            "day_name": day_name,
            "day_abbr": d.strftime("%a"),
            "day_num": day,
            "projects": []
        }

        takeover_counts = defaultdict(int)

        for proj in projects:
            owner_name = proj["employee_name"]
            owner = emp_lookup.get(owner_name)

            if not owner:
                continue

            owner_working = day_name in owner["working_days"]

            if owner_working:
                day_info["projects"].append({
                    "project_name": proj["name"],
                    "product_type": proj["product_type"],
                    "owner": owner_name,
                    "handler": owner_name,
                    "is_takeover": False,
                })
            else:
                handler = _find_takeover(
                    proj, owner_name, employees, shift_assignments,
                    day_name, takeover_counts, projects_by_owner
                )
                if handler:
                    takeover_counts[handler] += 1
                    day_info["projects"].append({
                        "project_name": proj["name"],
                        "product_type": proj["product_type"],
                        "owner": owner_name,
                        "handler": handler,
                        "is_takeover": True,
                    })
                else:
                    day_info["projects"].append({
                        "project_name": proj["name"],
                        "product_type": proj["product_type"],
                        "owner": owner_name,
                        "handler": None,
                        "is_takeover": True,
                    })
                    warnings.append(
                        f"{d.strftime('%b %d')} ({day_name}): No secondary available for "
                        f"project '{proj['name']}' (owner {owner_name} is off)"
                    )

        coverage.append(day_info)

    return coverage, warnings


def _find_takeover(project, owner_name, employees, shift_assignments,
                   day_name, takeover_counts, projects_by_owner):
    """
    Find the best takeover person for a project on a given day.

    Criteria:
      1. Must be on the same shift as the owner
      2. Must share the project's product_type in their content_types
      3. Must be working on this day
      4. Must not be the owner themselves
      5. Among eligible, pick the one with the fewest total projects
         (own projects + already-assigned takeovers today)
    """
    owner_shift = shift_assignments.get(owner_name)
    if owner_shift is None:
        return None

    product_type = project["product_type"]
    candidates = []

    for emp in employees:
        if emp["name"] == owner_name:
            continue
        if shift_assignments.get(emp["name"]) != owner_shift:
            continue
        if product_type not in emp["content_types"]:
            continue
        if day_name not in emp["working_days"]:
            continue

        own_project_count = len(projects_by_owner.get(emp["name"], []))
        today_takeover_count = takeover_counts.get(emp["name"], 0)
        total_load = own_project_count + today_takeover_count

        candidates.append((emp["name"], total_load))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[1])
    return candidates[0][0]
