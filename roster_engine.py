import calendar
from datetime import date
from math import ceil
from collections import defaultdict
from itertools import combinations

SHIFTS = {
    1: {"name": "Shift 1", "time_ist": "6:00 AM - 2:00 PM IST", "time_est": "7:30 PM - 3:30 AM EST", "strength": "lean"},
    2: {"name": "Shift 2", "time_ist": "1:00 PM - 10:00 PM IST", "time_est": "2:30 AM - 11:30 AM EST", "strength": "strong"},
    3: {"name": "Shift 3", "time_ist": "9:00 PM - 6:00 AM IST", "time_est": "10:30 AM - 7:30 PM EST", "strength": "strong"},
}

CONTENT_TYPES = ["Content", "Email", "Message"]

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def assign_shifts(employees, night_shift_counts=None):
    """
    Assign each employee to a shift (1, 2, or 3).

    Strategy:
      1. Calculate target sizes (~20% Shift 1, ~40% Shift 2, ~40% Shift 3)
      2. Sort employees by night shift history for round-robin rotation
      3. Assign employees to shifts while trying to stagger off-days
         so every shift has at least 1 person working every day of the week
      4. Ensure content type coverage across shifts
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

    targets = {3: n_shift3, 2: n_shift2, 1: n_shift1}

    if night_shift_counts:
        sorted_emps = sorted(
            employees,
            key=lambda e: night_shift_counts.get(e.get("id", 0), 0)
        )
    else:
        sorted_emps = list(employees)

    assignments = _assign_with_staggered_offs(sorted_emps, targets)

    _ensure_content_type_coverage(assignments, employees)
    _fix_daily_gaps(assignments, employees)

    return assignments


def _get_off_days(emp):
    """Return the set of days this employee is NOT working."""
    return set(DAY_NAMES) - set(emp["working_days"])


def _assign_with_staggered_offs(sorted_emps, targets):
    """
    Assign employees to shifts such that off-days within each shift
    are staggered (minimizing days where all shift members are off).
    """
    shift_members = {1: [], 2: [], 3: []}
    assigned = set()

    for shift_num in [3, 2, 1]:
        target = targets[shift_num]
        available = [e for e in sorted_emps if e["name"] not in assigned]

        if not available:
            break

        if target == 1:
            shift_members[shift_num].append(available[0])
            assigned.add(available[0]["name"])
            continue

        best_combo = _find_best_stagger_combo(available, target)
        for emp in best_combo:
            shift_members[shift_num].append(emp)
            assigned.add(emp["name"])

    for emp in sorted_emps:
        if emp["name"] not in assigned:
            for shift_num in [2, 3, 1]:
                if len(shift_members[shift_num]) < targets[shift_num]:
                    shift_members[shift_num].append(emp)
                    assigned.add(emp["name"])
                    break

    assignments = {}
    for shift_num, members in shift_members.items():
        for emp in members:
            assignments[emp["name"]] = shift_num

    return assignments


def _find_best_stagger_combo(available, target):
    """
    Among available employees, find the combination of `target` people
    whose off-days overlap the least (best stagger).
    """
    if len(available) <= target:
        return available[:target]

    best_score = -1
    best_combo = available[:target]

    candidates = available[:min(len(available), 10)]

    for combo in combinations(candidates, min(target, len(candidates))):
        score = _stagger_score(combo)
        if score > best_score:
            best_score = score
            best_combo = list(combo)

    return best_combo


def _stagger_score(group):
    """
    Score a group of employees by how well their off-days are staggered.
    Higher = better (more days of the week covered by at least 1 person).
    """
    coverage = 0
    for day in DAY_NAMES:
        working = sum(1 for emp in group if day in emp["working_days"])
        if working > 0:
            coverage += 1
    return coverage


def _fix_daily_gaps(assignments, employees):
    """
    After initial assignment, check each day of the week.
    If a shift has zero working employees on a given day,
    try to swap someone from another shift to fill the gap.
    """
    emp_lookup = {e["name"]: e for e in employees}

    for day in DAY_NAMES:
        for shift_num in [1, 2, 3]:
            shift_names = [n for n, s in assignments.items() if s == shift_num]
            working_today = [n for n in shift_names if day in emp_lookup[n]["working_days"]]

            if len(working_today) > 0:
                continue

            for donor_shift in [1, 2, 3]:
                if donor_shift == shift_num:
                    continue
                donor_names = [n for n, s in assignments.items() if s == donor_shift]
                if len(donor_names) <= 1:
                    continue

                donors_working_today = [
                    n for n in donor_names if day in emp_lookup[n]["working_days"]
                ]
                if len(donors_working_today) <= 1:
                    continue

                candidate = donors_working_today[-1]

                would_break = False
                remaining_donors = [n for n in donor_names if n != candidate]
                for check_day in DAY_NAMES:
                    still_covered = any(
                        check_day in emp_lookup[n]["working_days"]
                        for n in remaining_donors
                    )
                    if not still_covered:
                        would_break = True
                        break

                if not would_break:
                    assignments[candidate] = shift_num
                    break


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

    off_day_warnings = _check_off_day_overlaps(shift_assignments, employees)
    warnings.extend(off_day_warnings)

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


def _check_off_day_overlaps(assignments, employees):
    """
    Check if employees in the same shift share off-days that leave
    the shift uncovered on certain days.
    """
    emp_lookup = {e["name"]: e for e in employees}
    warnings = []

    for shift_num in [1, 2, 3]:
        shift_names = [n for n, s in assignments.items() if s == shift_num]
        if len(shift_names) <= 1:
            if len(shift_names) == 1:
                off = _get_off_days(emp_lookup[shift_names[0]])
                if off:
                    warnings.append(
                        f"WARNING: {SHIFTS[shift_num]['name']} has only 1 person "
                        f"({shift_names[0]}), no coverage on their off-days: "
                        f"{', '.join(sorted(off))}"
                    )
            continue

        for day in DAY_NAMES:
            working = [n for n in shift_names if day in emp_lookup[n]["working_days"]]
            if len(working) == 0:
                off_people = ", ".join(shift_names)
                warnings.append(
                    f"OVERLAP: All {SHIFTS[shift_num]['name']} members "
                    f"({off_people}) are off on {day}. "
                    f"Consider staggering off-days or adding more employees."
                )

    return warnings


def get_roster_summary(employees, shift_assignments):
    """Return a summary of shift distribution for display."""
    summary = {s: {ct: [] for ct in CONTENT_TYPES} for s in [1, 2, 3]}

    for emp in employees:
        shift = shift_assignments[emp["name"]]
        for ct in emp["content_types"]:
            summary[shift][ct].append(emp["name"])

    return summary
