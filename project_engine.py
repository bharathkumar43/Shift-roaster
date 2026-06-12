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


def _pick_min_coverage(candidates, coverage_load, shift_num, manager_name, emp_lookup):
    """Choose engineer with lowest backup coverage count; tiebreak by id."""
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda n: (
            coverage_load[(n, shift_num, manager_name)],
            _emp_sort_key(emp_lookup, n),
        ),
    )


def _greedy_pick(candidates, assignment_count, emp_lookup):
    """Pick the candidate with the fewest total project assignments globally (tiebreak by id)."""
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda n: (assignment_count[n], _emp_sort_key(emp_lookup, n)),
    )


def _resolve_manager(proj, emp_lookup):
    """
    Return the coverage-manager name for a project.

    - Project owned by a top-level manager (manager_name=='') → use owner name
    - Project owned by a sub-manager (Sriram has manager_name='Raghu') → use parent manager
    - Project still owned by an engineer (legacy) → use engineer's manager_name
    """
    owner = emp_lookup.get(proj["employee_name"]) or {}
    if owner.get("emp_role") == "manager":
        parent = (owner.get("manager_name") or "").strip()
        return parent if parent else owner.get("name", "")
    return (owner.get("manager_name") or "").strip()


def _resolve_division(manager_name, emp_lookup):
    mgr = emp_lookup.get(manager_name) or {}
    return mgr.get("division", "")


def _candidates_for(shift_num, manager_name, division, employees, shift_assignments):
    """
    Engineers eligible for a (project, shift) slot:
      primary  — emp has manager_name containing manager_name
      backup   — emp has backup_division matching the project's division
    Only employees with a shift assignment are included.
    """
    result = []
    for emp in employees:
        if shift_assignments.get(emp["name"]) != shift_num:
            continue
        emp_managers = [m.strip() for m in (emp.get("manager_name") or "").split(",") if m.strip()]
        if manager_name in emp_managers:
            result.append(emp["name"])
        elif division and emp.get("backup_division") == division:
            result.append(emp["name"])
    return result


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

    fixed_assignments, _ = _assign_fixed_handlers(
        unique_projects, employees, shift_assignments
    )

    num_days = calendar.monthrange(year, month)[1]
    coverage = []
    warnings = []

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]
        date_str = d.strftime("%Y-%m-%d")

        daily_backup_load = defaultdict(int)

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
            proj_manager = _resolve_manager(proj, emp_lookup)

            shift_handlers = {}

            for shift_num in [2, 3]:
                fixed_person = fixed_assignments.get(proj_key, {}).get(shift_num)

                if fixed_person:
                    fixed_emp = emp_lookup.get(fixed_person)
                    fixed_off = not is_emp_scheduled_work_day(fixed_emp, d) if fixed_emp else True
                    fixed_on_leave = date_str in leave_dates.get(fixed_person, [])

                    if fixed_emp and not fixed_off and not fixed_on_leave:
                        shift_handlers[shift_num] = {
                            "handler": fixed_person,
                            "is_secondary": False,
                            "is_owner_shift": (shift_num == owner_shift),
                        }
                    else:
                        backup = _find_backup(
                            proj_manager, fixed_person, shift_num,
                            employees, shift_assignments, day_name,
                            date_str, leave_dates,
                            backup_load=daily_backup_load,
                            emp_lookup=emp_lookup,
                        )
                        if backup:
                            shift_handlers[shift_num] = {
                                "handler": backup,
                                "is_secondary": True,
                                "is_owner_shift": (shift_num == owner_shift),
                            }
                        else:
                            shift_handlers[shift_num] = {
                                "handler": None,
                                "is_secondary": True,
                                "is_owner_shift": (shift_num == owner_shift),
                            }
                            warnings.append(
                                f"{d.strftime('%b %d')} ({day_name}): No handler in "
                                f"{SHIFTS[shift_num]['name']} for project '{proj['name']}' "
                                f"(assigned engineer is off)"
                            )
                else:
                    shift_handlers[shift_num] = {
                        "handler": None,
                        "is_secondary": False,
                        "is_owner_shift": (shift_num == owner_shift),
                    }

            day_info["projects"].append({
                "project_name": proj["name"],
                "product_type": product_type,
                "owner": owner_name,
                "owner_shift": owner_shift,
                "proj_manager": proj_manager,
                "shifts": shift_handlers,
            })

        coverage.append(day_info)

    return coverage, warnings


def _assign_fixed_handlers(projects, employees, shift_assignments):
    """
    Assign one fixed handler per (project, shift).

    Pure greedy: for each (project, shift), pick the least-loaded eligible engineer
    from the manager's team. Product-type filtering removed — team membership is the
    only eligibility criterion.

    assignment_count is global (across all shifts and projects) so the distribution
    stays within ≤1 between any two engineers in the same pool.
    """
    emp_lookup = {e["name"]: e for e in employees}
    fixed = {}
    assignment_count = defaultdict(int)

    for proj in projects:
        if not emp_lookup.get(proj["employee_name"]):
            continue
        proj_key = (proj["name"], proj["product_type"])
        if proj_key not in fixed:
            fixed[proj_key] = {}
        proj_manager = _resolve_manager(proj, emp_lookup)
        division = _resolve_division(proj_manager, emp_lookup)

        for shift_num in [2, 3]:
            pool = _candidates_for(shift_num, proj_manager, division, employees, shift_assignments)
            chosen = _greedy_pick(pool, assignment_count, emp_lookup) if pool else None
            fixed[proj_key][shift_num] = chosen
            if chosen:
                assignment_count[chosen] += 1

    return fixed, {}


def _find_backup(manager_name, exclude_name, shift_num,
                 employees, shift_assignments, day_name,
                 date_str="", leave_dates=None,
                 backup_load=None, emp_lookup=None):
    """
    Find a backup handler for a day when the fixed person is off or on leave.

    Eligibility: same shift, manager's team (or backup_division match), working that
    weekday, not on leave. Excludes the absent fixed person.
    """
    if leave_dates is None:
        leave_dates = {}
    if backup_load is None:
        backup_load = defaultdict(int)
    if emp_lookup is None:
        emp_lookup = {e["name"]: e for e in employees}

    division = _resolve_division(manager_name, emp_lookup)
    candidates = []

    for emp in employees:
        if emp["name"] == exclude_name:
            continue
        if shift_assignments.get(emp["name"]) != shift_num:
            continue
        emp_managers = [m.strip() for m in (emp.get("manager_name") or "").split(",") if m.strip()]
        in_team = (manager_name in emp_managers) or (division and emp.get("backup_division") == division)
        if not in_team:
            continue
        if not is_emp_scheduled_work_day(emp, date.fromisoformat(date_str)):
            continue
        if date_str in leave_dates.get(emp["name"], []):
            continue
        candidates.append(emp["name"])

    if not candidates:
        return None

    chosen = _pick_min_coverage(candidates, backup_load, shift_num, manager_name, emp_lookup)
    backup_load[(chosen, shift_num, manager_name)] += 1
    return chosen
