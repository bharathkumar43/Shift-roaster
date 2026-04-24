import calendar
from collections import defaultdict
from datetime import date

from roster_engine import (
    DAY_NAMES,
    SHIFTS,
    prepare_employees_for_roster_month,
    is_emp_scheduled_work_day,
)


def _emp_sort_key(emp_lookup, name):
    """Stable tiebreak: employee id (not alphabetical name)."""
    e = emp_lookup.get(name) or {}
    return (e.get("id") if e is not None else 0) or 0


def _pick_min_coverage(candidates, coverage_load, shift_num, product_type, emp_lookup):
    """Choose engineer with lowest coverage count for (shift, product_type); tiebreak by id."""
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda n: (
            coverage_load[(n, shift_num, product_type)],
            _emp_sort_key(emp_lookup, n),
        ),
    )


def generate_project_coverage(projects, employees, shift_assignments, year, month, leave_dates=None):
    """
    Generate daily project coverage for the month across ALL shifts.

    For each project, one fixed person is assigned per shift for the whole month.
    On their off days or leave days, the best available backup takes over.

    leave_dates: dict {employee_name: [date_str, ...]} of approved leaves
    """
    if leave_dates is None:
        leave_dates = {}

    employees, _ = prepare_employees_for_roster_month(employees, year, month)
    emp_lookup = {e["name"]: e for e in employees}

    seen_proj_keys = set()
    unique_projects = []
    for p in projects:
        if p["employee_name"] not in emp_lookup:
            continue
        key = (p["name"], p["product_type"])
        if key in seen_proj_keys:
            continue
        seen_proj_keys.add(key)
        unique_projects.append(p)

    projects_by_owner = defaultdict(list)
    for p in unique_projects:
        projects_by_owner[p["employee_name"]].append(p)

    emps_by_shift = defaultdict(list)
    for emp in employees:
        shift = shift_assignments.get(emp["name"])
        if shift:
            emps_by_shift[shift].append(emp)

    fixed_assignments = _assign_fixed_handlers(
        unique_projects, employees, shift_assignments, projects_by_owner, projects
    )

    num_days = calendar.monthrange(year, month)[1]
    coverage = []
    warnings = []
    # Balance backup picks across the month (per shift × product type)
    backup_load = defaultdict(int)

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]
        date_str = d.strftime("%Y-%m-%d")

        day_info = {
            "date": d.strftime("%b %d"),
            "day_name": day_name,
            "day_abbr": d.strftime("%a"),
            "day_num": day,
            "projects": []
        }

        for proj in unique_projects:
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
                    owner_off_weekly = not is_emp_scheduled_work_day(owner, d)
                    owner_on_leave = date_str in leave_dates.get(owner_name, [])
                    owner_available = not owner_off_weekly and not owner_on_leave

                    if owner_available:
                        shift_handlers[shift_num] = {
                            "handler": owner_name,
                            "is_secondary": False,
                            "is_owner_shift": True,
                        }
                    else:
                        backup = _find_backup(
                            product_type, owner_name, shift_num,
                            employees, shift_assignments, day_name,
                            date_str, leave_dates,
                            backup_load=backup_load,
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
                        fixed_off_weekly = (
                            not is_emp_scheduled_work_day(fixed_emp, d) if fixed_emp else True
                        )
                        fixed_on_leave = date_str in leave_dates.get(fixed_person, [])
                        fixed_available = fixed_emp and not fixed_off_weekly and not fixed_on_leave

                        if fixed_available:
                            shift_handlers[shift_num] = {
                                "handler": fixed_person,
                                "is_secondary": False,
                                "is_owner_shift": False,
                            }
                        else:
                            backup = _find_backup(
                                product_type, fixed_person, shift_num,
                                employees, shift_assignments, day_name,
                                date_str, leave_dates,
                                backup_load=backup_load,
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


def _assign_fixed_handlers(projects, employees, shift_assignments, projects_by_owner, all_projects=None):
    """
    Assign a fixed handler per shift for each project.

    Priority:
      1. Engineers who already have this project assigned in that shift (from DB)
      2. Among eligible engineers in that shift, pick whoever has the fewest
         cross-shift coverage slots so far for (shift, product_type) — then
         tiebreak by employee id (not name alphabet).

    Eligibility for a non-owner shift: same shift, content_types includes
    product_type. (Working days and leaves are applied per day in the
    coverage loop via backup.)
    """
    emp_lookup = {e["name"]: e for e in employees}
    fixed = {}

    if all_projects is None:
        all_projects = projects

    proj_assigned = defaultdict(list)
    for p in all_projects:
        if p["employee_name"] in emp_lookup:
            proj_assigned[p["name"].lower()].append(p["employee_name"])

    # Running count of fixed coverage slots per engineer per (shift, product_type)
    coverage_load = defaultdict(int)

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

            assigned_in_shift = [
                name for name in proj_assigned.get(proj["name"].lower(), [])
                if name != owner_name
                and shift_assignments.get(name) == shift_num
                and product_type in emp_lookup[name]["content_types"]
            ]
            if assigned_in_shift:
                chosen = _pick_min_coverage(
                    assigned_in_shift, coverage_load, shift_num, product_type, emp_lookup
                )
                fixed[proj_key][shift_num] = chosen
                coverage_load[(chosen, shift_num, product_type)] += 1
                continue

            candidates = []
            for emp in employees:
                if emp["name"] == owner_name:
                    continue
                if shift_assignments.get(emp["name"]) != shift_num:
                    continue
                if product_type not in emp["content_types"]:
                    continue
                candidates.append(emp["name"])

            if candidates:
                chosen = _pick_min_coverage(
                    candidates, coverage_load, shift_num, product_type, emp_lookup
                )
                fixed[proj_key][shift_num] = chosen
                coverage_load[(chosen, shift_num, product_type)] += 1
            else:
                fixed[proj_key][shift_num] = None

    return fixed


def _find_backup(product_type, exclude_name, shift_num,
                 employees, shift_assignments, day_name,
                 date_str="", leave_dates=None,
                 backup_load=None):
    """
    Find a backup handler for a day when the fixed person is off or on leave.

    Eligibility: same shift, content_types, working that weekday, not on leave.
    Among eligible engineers, prefer lowest backup_load for (shift, product_type)
    this month; tiebreak by employee id. Increments backup_load when set.
    """
    if leave_dates is None:
        leave_dates = {}
    if backup_load is None:
        backup_load = defaultdict(int)

    emp_lookup = {e["name"]: e for e in employees}
    candidates = []

    for emp in employees:
        if emp["name"] == exclude_name:
            continue
        if shift_assignments.get(emp["name"]) != shift_num:
            continue
        if product_type not in emp["content_types"]:
            continue
        if not is_emp_scheduled_work_day(emp, date.fromisoformat(date_str)):
            continue
        if date_str in leave_dates.get(emp["name"], []):
            continue

        candidates.append(emp["name"])

    if not candidates:
        return None

    chosen = _pick_min_coverage(
        candidates, backup_load, shift_num, product_type, emp_lookup
    )
    backup_load[(chosen, shift_num, product_type)] += 1
    return chosen
