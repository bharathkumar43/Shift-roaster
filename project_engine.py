import calendar
from datetime import date
from collections import defaultdict

from roster_engine import DAY_NAMES, SHIFTS


def generate_project_coverage(projects, employees, shift_assignments, year, month):
    """
    Generate daily project coverage for the month across ALL shifts.

    For each project, one fixed person is assigned per shift for the whole month.
    On their off days, the best available backup from the same shift takes over.
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

    fixed_assignments = _assign_fixed_handlers(
        projects, employees, shift_assignments, projects_by_owner
    )

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

        for proj in projects:
            owner_name = proj["employee_name"]
            owner = emp_lookup.get(owner_name)
            if not owner:
                continue

            owner_shift = shift_assignments.get(owner_name)
            product_type = proj["product_type"]
            proj_key = (proj["name"], product_type)

            shift_handlers = {}

            for shift_num in [1, 2, 3]:
                fixed_person = fixed_assignments.get(proj_key, {}).get(shift_num)

                if shift_num == owner_shift:
                    owner_working = day_name in owner["working_days"]
                    if owner_working:
                        shift_handlers[shift_num] = {
                            "handler": owner_name,
                            "is_secondary": False,
                            "is_owner_shift": True,
                        }
                    else:
                        backup = _find_backup(
                            product_type, owner_name, shift_num,
                            employees, shift_assignments, day_name
                        )
                        if backup:
                            shift_handlers[shift_num] = {
                                "handler": backup,
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
                                f"({owner_name} is off)"
                            )
                else:
                    if fixed_person:
                        fixed_emp = emp_lookup.get(fixed_person)
                        if fixed_emp and day_name in fixed_emp["working_days"]:
                            shift_handlers[shift_num] = {
                                "handler": fixed_person,
                                "is_secondary": False,
                                "is_owner_shift": False,
                            }
                        else:
                            backup = _find_backup(
                                product_type, fixed_person, shift_num,
                                employees, shift_assignments, day_name
                            )
                            shift_handlers[shift_num] = {
                                "handler": backup,
                                "is_secondary": True,
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


def _assign_fixed_handlers(projects, employees, shift_assignments, projects_by_owner):
    """
    For each project + shift combination (excluding the owner's shift),
    pick one fixed person for the whole month based on fewest existing projects.
    """
    emp_lookup = {e["name"]: e for e in employees}
    fixed = {}
    shift_load = defaultdict(int)

    for proj in projects:
        owner_name = proj["employee_name"]
        owner = emp_lookup.get(owner_name)
        if not owner:
            continue

        owner_shift = shift_assignments.get(owner_name)
        product_type = proj["product_type"]
        proj_key = (proj["name"], product_type)

        if proj_key not in fixed:
            fixed[proj_key] = {}

        for shift_num in [1, 2, 3]:
            if shift_num == owner_shift:
                fixed[proj_key][shift_num] = owner_name
                continue

            candidates = []
            for emp in employees:
                if emp["name"] == owner_name:
                    continue
                if shift_assignments.get(emp["name"]) != shift_num:
                    continue
                if product_type not in emp["content_types"]:
                    continue

                own_count = len(projects_by_owner.get(emp["name"], []))
                load = shift_load.get(emp["name"], 0)
                total = own_count + load
                candidates.append((emp["name"], total))

            if candidates:
                candidates.sort(key=lambda x: x[1])
                chosen = candidates[0][0]
                fixed[proj_key][shift_num] = chosen
                shift_load[chosen] += 1
            else:
                fixed[proj_key][shift_num] = None

    return fixed


def _find_backup(product_type, exclude_name, shift_num,
                 employees, shift_assignments, day_name):
    """
    Find a backup handler for a day when the fixed person is off.
    Picks anyone in the same shift with the matching product type who is working.
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

        candidates.append(emp["name"])

    return candidates[0] if candidates else None
