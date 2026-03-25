import calendar
from datetime import date
from math import ceil
from collections import defaultdict

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
      1. Guarantee: place at least 1 employee of each product type in each shift
      2. Distribute remaining employees per product type in 20/40/40 ratio
      3. Night shift rotation via round-robin
      4. Fix daily gaps from off-day overlaps
    """
    n = len(employees)
    if n == 0:
        return {}

    if night_shift_counts:
        sorted_emps = sorted(
            employees,
            key=lambda e: night_shift_counts.get(e.get("id", 0), 0)
        )
    else:
        sorted_emps = list(employees)

    total = len(sorted_emps)
    target_s1 = max(1, round(total * 0.2))
    target_s3 = max(1, round(total * 0.4))
    target_s2 = max(1, total - target_s1 - target_s3)
    if target_s1 + target_s2 + target_s3 != total:
        target_s2 += total - (target_s1 + target_s2 + target_s3)
    targets = {1: target_s1, 2: target_s2, 3: target_s3}

    assignments = {}
    assigned = set()

    _guarantee_type_coverage(sorted_emps, assignments, assigned, night_shift_counts)
    _distribute_remaining(sorted_emps, assignments, assigned, targets, night_shift_counts)
    _fix_daily_gaps(assignments, sorted_emps)

    return assignments


def _guarantee_type_coverage(sorted_emps, assignments, assigned, night_shift_counts):
    """
    For each product type, ensure at least 1 employee is placed in each shift.
    Prioritizes employees who handle fewer product types (specialists first),
    so multi-type employees remain available for flexible placement later.
    """
    emp_lookup = {e["name"]: e for e in sorted_emps}

    for ct in CONTENT_TYPES:
        candidates = [e for e in sorted_emps if ct in e["content_types"]]
        if not candidates:
            continue

        for shift_num in [1, 2, 3]:
            already_covered = any(
                assignments.get(e["name"]) == shift_num and ct in emp_lookup[e["name"]]["content_types"]
                for e in sorted_emps if e["name"] in assigned
            )
            if already_covered:
                continue

            unassigned = [e for e in candidates if e["name"] not in assigned]

            unassigned.sort(key=lambda e: len(e["content_types"]))

            placed = False
            for emp in unassigned:
                assignments[emp["name"]] = shift_num
                assigned.add(emp["name"])
                placed = True
                break

            if not placed:
                for emp in candidates:
                    if emp["name"] in assigned and assignments[emp["name"]] != shift_num:
                        continue
                    if emp["name"] not in assigned:
                        assignments[emp["name"]] = shift_num
                        assigned.add(emp["name"])
                        break


def _distribute_remaining(sorted_emps, assignments, assigned, targets, night_shift_counts):
    """
    Distribute unassigned employees across shifts respecting 20/40/40 targets
    and per-product-type balance.
    """
    shift_counts = {1: 0, 2: 0, 3: 0}
    for s in assignments.values():
        shift_counts[s] += 1

    remaining = [e for e in sorted_emps if e["name"] not in assigned]

    type_shift_counts = defaultdict(lambda: {1: 0, 2: 0, 3: 0})
    emp_lookup = {e["name"]: e for e in sorted_emps}
    for name, shift in assignments.items():
        for ct in emp_lookup[name]["content_types"]:
            type_shift_counts[ct][shift] += 1

    type_totals = defaultdict(int)
    for emp in sorted_emps:
        for ct in emp["content_types"]:
            type_totals[ct] += 1

    type_targets = {}
    for ct in CONTENT_TYPES:
        t = type_totals[ct]
        if t == 0:
            continue
        s1 = max(1, round(t * 0.2))
        s3 = max(1, round(t * 0.4))
        s2 = max(1, t - s1 - s3)
        if s1 + s2 + s3 != t:
            s2 += t - (s1 + s2 + s3)
        type_targets[ct] = {1: s1, 2: s2, 3: s3}

    if night_shift_counts:
        remaining.sort(key=lambda e: night_shift_counts.get(e.get("id", 0), 0))

    for emp in remaining:
        best_shift = _pick_best_shift(emp, shift_counts, targets, type_shift_counts, type_targets)
        assignments[emp["name"]] = best_shift
        assigned.add(emp["name"])
        shift_counts[best_shift] += 1
        for ct in emp["content_types"]:
            type_shift_counts[ct][best_shift] += 1


def _pick_best_shift(emp, shift_counts, targets, type_shift_counts, type_targets):
    """
    Pick the best shift for an employee considering:
      1. Which shifts still need more people (overall targets)
      2. Which shifts need this employee's product type(s) most
    """
    scores = {}
    for s in [1, 2, 3]:
        if shift_counts[s] >= targets[s]:
            scores[s] = -100
            continue

        score = targets[s] - shift_counts[s]

        for ct in emp["content_types"]:
            if ct in type_targets:
                needed = type_targets[ct][s] - type_shift_counts[ct][s]
                score += needed * 2

        scores[s] = score

    return max(scores, key=scores.get)


def _get_off_days(emp):
    return set(DAY_NAMES) - set(emp["working_days"])


def _fix_daily_gaps(assignments, employees):
    """
    Final pass: for each day, if a shift has zero people working,
    try to move someone from a shift with 2+ people that day.
    """
    emp_lookup = {e["name"]: e for e in employees}

    for day in DAY_NAMES:
        for shift_num in [1, 2, 3]:
            shift_names = [n for n, s in assignments.items() if s == shift_num]
            working_today = [n for n in shift_names if day in emp_lookup[n]["working_days"]]

            if len(working_today) > 0:
                continue

            for donor_shift in [2, 3, 1]:
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
                remaining_donors = [n for n in donor_names if n != candidate]
                would_break = any(
                    not any(d in emp_lookup[n]["working_days"] for n in remaining_donors)
                    for d in DAY_NAMES
                )

                if not would_break:
                    assignments[candidate] = shift_num
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

    type_dist_warnings = _check_type_distribution(shift_assignments, employees)
    warnings.extend(type_dist_warnings)

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


def _check_type_distribution(assignments, employees):
    emp_lookup = {e["name"]: e for e in employees}
    warnings = []

    for shift_num in [1, 2, 3]:
        shift_names = [n for n, s in assignments.items() if s == shift_num]
        type_counts = defaultdict(int)
        for name in shift_names:
            for ct in emp_lookup[name]["content_types"]:
                type_counts[ct] += 1

        for ct in CONTENT_TYPES:
            if type_counts[ct] == 0:
                has_any = any(ct in emp_lookup[n]["content_types"] for n in assignments)
                if has_any:
                    warnings.append(
                        f"DISTRIBUTION: {SHIFTS[shift_num]['name']} has no "
                        f"'{ct}' employees. Consider adding more '{ct}' workers."
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
