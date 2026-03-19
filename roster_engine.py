import calendar
from datetime import date
from math import ceil, floor
from collections import defaultdict

SHIFTS = {
    1: {"name": "Shift 1", "time": "6:00 AM - 2:00 PM IST", "strength": "lean"},
    2: {"name": "Shift 2", "time": "1:00 PM - 10:00 PM IST", "strength": "strong"},
    3: {"name": "Shift 3", "time": "9:00 PM - 6:00 AM IST", "strength": "strong"},
}

CONTENT_TYPES = ["Content & Email", "Message"]

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def assign_shifts(employees):
    """
    Assign each employee to a shift (1, 2, or 3).

    Distribution target per content type:
      - Shift 1: ~20%
      - Shift 2: ~40%
      - Shift 3: ~40%

    Each shift must have at least 1 person from each content type.
    """
    by_type = defaultdict(list)
    for emp in employees:
        by_type[emp["content_type"]].append(emp)

    assignments = {}

    for ctype, members in by_type.items():
        n = len(members)
        if n == 0:
            continue

        n_shift1 = max(1, round(n * 0.2))
        remaining = n - n_shift1
        n_shift2 = max(1, ceil(remaining / 2))
        n_shift3 = max(1, remaining - n_shift2)

        if n_shift1 + n_shift2 + n_shift3 != n:
            diff = n - (n_shift1 + n_shift2 + n_shift3)
            n_shift2 += diff

        idx = 0
        for i in range(n_shift1):
            assignments[members[idx]["name"]] = 1
            idx += 1
        for i in range(n_shift2):
            assignments[members[idx]["name"]] = 2
            idx += 1
        for i in range(n_shift3):
            assignments[members[idx]["name"]] = 3
            idx += 1

    return assignments


def generate_roster(employees, year, month):
    """
    Generate a full monthly roster.

    Returns:
        roster: dict mapping date -> {shift_num: [employee_names]}
        warnings: list of warning strings for coverage gaps
        shift_assignments: dict mapping employee_name -> shift_num
    """
    shift_assignments = assign_shifts(employees)

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
                types_covered.add(emp_lookup[name]["content_type"])

            for ct in CONTENT_TYPES:
                if ct not in types_covered:
                    warnings.append(
                        f"{d.strftime('%b %d')} ({day_name}): {SHIFTS[shift_num]['name']} "
                        f"missing '{ct}' coverage"
                    )

    return roster, warnings, shift_assignments


def get_roster_summary(employees, shift_assignments):
    """Return a summary of shift distribution for display."""
    summary = {1: {"Content & Email": [], "Message": []},
               2: {"Content & Email": [], "Message": []},
               3: {"Content & Email": [], "Message": []}}

    for emp in employees:
        shift = shift_assignments[emp["name"]]
        summary[shift][emp["content_type"]].append(emp["name"])

    return summary
