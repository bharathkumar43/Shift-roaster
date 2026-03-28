import calendar
from datetime import date
from collections import defaultdict

from roster_engine import DAY_NAMES, SHIFTS


def generate_project_coverage(projects, employees, shift_assignments, year, month):
    """
    Generate daily project coverage for the month across ALL shifts.

    For each project:
      - The owner handles it in their shift
      - For the other 2 shifts, the system auto-assigns the best-fit employee
        (same product type, working that day, fewest projects)
      - When owner or assigned handler is off, a secondary is picked

    Returns:
        coverage: list of dicts per day, each with:
            projects: list of {project_name, product_type, owner, shifts: {1: {...}, 2: {...}, 3: {...}}}
        warnings: list of warning strings
    """
    emp_lookup = {e["name"]: e for e in employees}

    projects_by_owner = defaultdict(list)
    for p in projects:
        projects_by_owner[p["employee_name"]].append(p)

    emps_by_shift = defaultdict(list)
    for emp in employees:
        shift = shift_assignments.get(emp["name"])
        if shift:
            emps_by_shift[shift].append(emp)

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

        assignment_counts = defaultdict(int)

        for proj in projects:
            owner_name = proj["employee_name"]
            owner = emp_lookup.get(owner_name)
            if not owner:
                continue

            owner_shift = shift_assignments.get(owner_name)
            product_type = proj["product_type"]

            shift_handlers = {}

            for shift_num in [1, 2, 3]:
                if shift_num == owner_shift:
                    owner_working = day_name in owner["working_days"]
                    if owner_working:
                        shift_handlers[shift_num] = {
                            "handler": owner_name,
                            "is_secondary": False,
                            "is_owner_shift": True,
                        }
                    else:
                        secondary = _find_handler(
                            product_type, owner_name, shift_num,
                            employees, shift_assignments, day_name,
                            assignment_counts, projects_by_owner
                        )
                        if secondary:
                            assignment_counts[secondary] += 1
                            shift_handlers[shift_num] = {
                                "handler": secondary,
                                "is_secondary": True,
                                "is_owner_shift": True,
                            }
                        else:
                            shift_handlers[shift_num] = {
                                "handler": None,
                                "is_secondary": True,
                                "is_owner_shift": True,
                            }
                            warnings.append(
                                f"{d.strftime('%b %d')} ({day_name}): No handler in "
                                f"{SHIFTS[shift_num]['name']} for project '{proj['name']}' "
                                f"(owner {owner_name} is off)"
                            )
                else:
                    handler = _find_handler(
                        product_type, None, shift_num,
                        employees, shift_assignments, day_name,
                        assignment_counts, projects_by_owner
                    )
                    if handler:
                        assignment_counts[handler] += 1
                        shift_handlers[shift_num] = {
                            "handler": handler,
                            "is_secondary": False,
                            "is_owner_shift": False,
                        }
                    else:
                        shift_handlers[shift_num] = {
                            "handler": None,
                            "is_secondary": False,
                            "is_owner_shift": False,
                        }

            day_info["projects"].append({
                "project_name": proj["name"],
                "product_type": product_type,
                "owner": owner_name,
                "owner_shift": owner_shift,
                "shifts": shift_handlers,
            })

        coverage.append(day_info)

    return coverage, warnings


def _find_handler(product_type, exclude_name, shift_num,
                  employees, shift_assignments, day_name,
                  assignment_counts, projects_by_owner):
    """
    Find the best handler for a project in a given shift on a given day.

    Picks the employee with the fewest total assignments (own projects +
    already-assigned coverage today) who matches the product type and is working.
    """
    candidates = []

    for emp in employees:
        if emp["name"] == exclude_name:
            continue
        if shift_assignments.get(emp["name"]) != shift_num:
            continue
        if product_type not in emp["content_types"]:
            continue
        if day_name not in emp["working_days"]:
            continue

        own_count = len(projects_by_owner.get(emp["name"], []))
        assigned_count = assignment_counts.get(emp["name"], 0)
        total = own_count + assigned_count

        candidates.append((emp["name"], total))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[1])
    return candidates[0][0]
