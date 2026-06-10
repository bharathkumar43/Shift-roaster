import calendar
from collections import defaultdict
from datetime import date

from roster_engine import (
    DAY_NAMES,
    SHIFTS,
    prepare_employees_for_roster_month,
    is_emp_scheduled_work_day,
)
import database as _db


def _emp_sort_key(emp_lookup, name):
    """Stable tiebreak: employee id (not alphabetical name)."""
    e = emp_lookup.get(name) or {}
    return (e.get("id") if e is not None else 0) or 0


def _pick_min_coverage(candidates, coverage_load, shift_num, product_type, emp_lookup):
    """Choose engineer with lowest coverage count for (shift, product_type); tiebreak by id.
    Used only for day-level backup selection within a month."""
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda n: (
            coverage_load[(n, shift_num, product_type)],
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

    saved_handlers = _db.get_project_shift_handlers()
    fixed_assignments, new_handlers = _assign_fixed_handlers(
        unique_projects, employees, shift_assignments, saved_handlers
    )
    if new_handlers:
        _db.save_project_shift_handlers(new_handlers)

    num_days = calendar.monthrange(year, month)[1]
    coverage = []
    warnings = []

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]
        date_str = d.strftime("%Y-%m-%d")

        # Reset backup load each day so returning engineers are not overloaded
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

            shift_handlers = {}

            for shift_num in [1, 2, 3]:
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
                            product_type, fixed_person, shift_num,
                            employees, shift_assignments, day_name,
                            date_str, leave_dates,
                            backup_load=daily_backup_load,
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
                "shifts": shift_handlers,
            })

        coverage.append(day_info)

    return coverage, warnings


def _assign_fixed_handlers(projects, employees, shift_assignments, saved_handlers=None):
    """
    Assign one fixed handler per (project, shift) with ≤ 1 count difference between engineers.

    Rule 1 — Owner pinning: in the owner's own shift the project is always assigned to
    the owner. This counts towards their global load.

    Rule 2 — Equal distribution for non-owner shifts: unassigned slots are grouped by
    (shift, product_type) and filled with the globally least-loaded eligible engineer,
    guaranteeing ≤ 1 difference within each eligible pool.

    Rule 3 — Stability: saved handlers (from project_shift_handlers DB) are reused
    as-is when still valid, preventing reshuffling when a new project is added.

    Two-pass approach:
      Pass 1 — lock in all determined assignments (owner-pinned + valid saved handlers)
               and seed the global assignment counter with their counts.
      Pass 2 — group remaining unassigned slots by (shift, product_type), then greedy-
               fill each group so the least-loaded eligible engineer is always chosen,
               achieving at most 1 project difference between engineers in the same pool.
    """
    if saved_handlers is None:
        saved_handlers = {}

    emp_lookup = {e["name"]: e for e in employees}
    fixed = {}
    new_handlers = {}
    # Global count per engineer name — single-shift employees so this equals per-shift count,
    # but tracking globally ensures cross-type fairness (e.g. CON+MSG engineer vs EML+MSG).
    assignment_count = defaultdict(int)

    def _is_valid_saved(handler_name, shift_num, product_type):
        emp = emp_lookup.get(handler_name)
        return (emp and shift_assignments.get(handler_name) == shift_num
                and product_type in emp.get("content_types", []))

    # ── Pass 1: lock determined assignments and seed the load counter ──────────
    for proj in projects:
        owner_name = proj["employee_name"]
        if not emp_lookup.get(owner_name):
            continue
        owner_shift = shift_assignments.get(owner_name)
        product_type = proj["product_type"]
        proj_key = (proj["name"], product_type)
        if proj_key not in fixed:
            fixed[proj_key] = {}

        for shift_num in [1, 2, 3]:
            if shift_num == owner_shift:
                fixed[proj_key][shift_num] = owner_name
                assignment_count[owner_name] += 1
            else:
                saved_pair = saved_handlers.get((proj["name"], product_type, shift_num))
                if saved_pair and _is_valid_saved(saved_pair[0], shift_num, product_type):
                    fixed[proj_key][shift_num] = saved_pair[0]
                    assignment_count[saved_pair[0]] += 1

    # ── Pass 2: group unassigned slots by (shift, type) then greedy-fill ───────
    def _candidates(shift_num, product_type):
        return [
            emp["name"] for emp in employees
            if shift_assignments.get(emp["name"]) == shift_num
            and product_type in emp.get("content_types", [])
        ]

    # Collect all unassigned (project, shift) pairs grouped by (shift, product_type)
    unassigned = defaultdict(list)
    for proj in projects:
        if not emp_lookup.get(proj["employee_name"]):
            continue
        product_type = proj["product_type"]
        proj_key = (proj["name"], product_type)
        for shift_num in [1, 2, 3]:
            if shift_num not in fixed.get(proj_key, {}):
                unassigned[(shift_num, product_type)].append(proj_key)

    # Assign each group: always pick the globally least-loaded eligible engineer
    for (shift_num, product_type), proj_keys in unassigned.items():
        pool = _candidates(shift_num, product_type)
        for pk in proj_keys:
            if pk not in fixed:
                fixed[pk] = {}
            chosen = _greedy_pick(pool, assignment_count, emp_lookup) if pool else None
            fixed[pk][shift_num] = chosen
            if chosen:
                assignment_count[chosen] += 1
                new_handlers[(pk[0], pk[1], shift_num)] = (chosen, None)

    return fixed, new_handlers


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
        if product_type not in emp.get("content_types", []):
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
