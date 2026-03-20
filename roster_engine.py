import calendar
from datetime import date
from math import ceil
from collections import defaultdict

SHIFTS = {
    1: {"name": "Shift 1", "time": "6:00 AM - 2:00 PM IST", "strength": "lean"},
    2: {"name": "Shift 2", "time": "1:00 PM - 10:00 PM IST", "strength": "strong"},
    3: {"name": "Shift 3", "time": "9:00 PM - 6:00 AM IST", "strength": "strong"},
}

CONTENT_TYPES = ["Content", "Email", "Message"]

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def assign_shifts(employees, night_shift_counts=None):
    """
    Assign each employee to a shift (1, 2, or 3).

    Distribution target: ~20% Shift 1, ~40% Shift 2, ~40% Shift 3.
    Night shift (3) rotates monthly via round-robin: employees with the
    fewest past night-shift months are picked first for Shift 3.

    Args:
        employees: list of employee dicts with 'name', 'content_types' (list), 'working_days'
        night_shift_counts: dict mapping employee_id -> number of months on night shift.
                           If None, no rotation preference is applied.
    """
    n = len(employees)
    if n == 0:
        return {}

    n_shift3 = max(1, round(n * 0.4))
    n_shift2 = max(1, round(n * 0.4))
    n_shift1 = max(1, n - n_shift3 - n_shift2)

    if n_shift1 + n_shift2 + n_shift3 != n:
        diff = n - (n_shift1 + n_shift2 + n_shift3)
        n_shift2 += diff

    if night_shift_counts:
        sorted_emps = sorted(
            employees,
            key=lambda e: night_shift_counts.get(e.get("id", 0), 0)
        )
    else:
        sorted_emps = list(employees)

    assignments = {}
    idx = 0

    for _ in range(n_shift3):
        assignments[sorted_emps[idx]["name"]] = 3
        idx += 1
    for _ in range(n_shift2):
        assignments[sorted_emps[idx]["name"]] = 2
        idx += 1
    for _ in range(n_shift1):
        assignments[sorted_emps[idx]["name"]] = 1
        idx += 1

    _ensure_content_type_coverage(assignments, employees)

    return assignments


def _ensure_content_type_coverage(assignments, employees):
    """Try to ensure each shift has at least one employee covering each content type."""
    emp_lookup = {e["name"]: e for e in employees}

    for shift_num in [1, 2, 3]:
        shift_names = [n for n, s in assignments.items() if s == shift_num]
        types_covered = set()
        for name in shift_names:
            for ct in emp_lookup[name]["content_types"]:
                types_covered.add(ct)

        for ct in CONTENT_TYPES:
            if ct not in types_covered:
                for other_shift in [1, 2, 3]:
                    if other_shift == shift_num:
                        continue
                    other_names = [n for n, s in assignments.items() if s == other_shift]
                    if len(other_names) <= 1:
                        continue
                    for candidate in other_names:
                        if ct in emp_lookup[candidate]["content_types"]:
                            assignments[candidate] = shift_num
                            break
                    if any(ct in emp_lookup[n]["content_types"]
                           for n in assignments if assignments[n] == shift_num):
                        break


def generate_roster(employees, year, month, night_shift_counts=None):
    """
    Generate a full monthly roster.

    Returns:
        roster: dict mapping date -> {shift_num: [employee_names]}
        warnings: list of warning strings for coverage gaps
        shift_assignments: dict mapping employee_name -> shift_num
    """
    shift_assignments = assign_shifts(employees, night_shift_counts)
    emp_lookup = {emp["name"]: emp for emp in employees}

    num_days = calendar.monthrange(year, month)[1]
    roster = {}
    warnings = []

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]

        daily = {1: [], 2: [], 3: []}

        for emp in employees:
            if day_name in emp["working_days"]:
                shift = shift_assignments[emp["name"]]
                daily[shift].append(emp["name"])

        roster[d] = daily

        for shift_num in [1, 2, 3]:
            shift_employees = daily[shift_num]
            if len(shift_employees) == 0:
                warnings.append(
                    f"{d.strftime('%b %d')} ({day_name}): {SHIFTS[shift_num]['name']} has NO coverage"
                )
                continue

            types_covered = set()
            for name in shift_employees:
                for ct in emp_lookup[name]["content_types"]:
                    types_covered.add(ct)

            for ct in CONTENT_TYPES:
                if ct not in types_covered:
                    warnings.append(
                        f"{d.strftime('%b %d')} ({day_name}): {SHIFTS[shift_num]['name']} "
                        f"missing '{ct}' coverage"
                    )

    return roster, warnings, shift_assignments


def get_roster_summary(employees, shift_assignments):
    """Return a summary of shift distribution for display."""
    summary = {s: {ct: [] for ct in CONTENT_TYPES} for s in [1, 2, 3]}

    for emp in employees:
        shift = shift_assignments[emp["name"]]
        for ct in emp["content_types"]:
            summary[shift][ct].append(emp["name"])

    return summary
