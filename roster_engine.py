import calendar
import json
import os
from collections import defaultdict
from datetime import date, timedelta
from math import ceil

SHIFTS = {
    1: {"name": "Shift 1", "time_ist": "6:00 AM - 2:00 PM IST", "time_est": "7:30 PM - 3:30 AM EST", "strength": "lean"},
    2: {"name": "Shift 2", "time_ist": "1:00 PM - 10:00 PM IST", "time_est": "2:30 AM - 11:30 AM EST", "strength": "strong"},
    3: {"name": "Shift 3", "time_ist": "9:00 PM - 6:00 AM IST", "time_est": "10:30 AM - 7:30 PM EST", "strength": "strong"},
}

CONTENT_TYPES = ["Content", "Email", "Message"]

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

DEFAULT_FIVE_DAY_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# Mandated shift headcount ratio (Shift 1 : Shift 2 : Shift 3) for a full team of 21.
SHIFT_MANDATE_WEIGHTS = (6, 7, 8)

# Comma-separated engineer names always kept on Shift 1 (Edit Shifts and generation honor this).
# Optional: set PINNED_SHIFT_1_NAMES in .env, e.g. PINNED_SHIFT_1_NAMES=Alice,Bob
# Default is empty so shift changes in the UI are not overridden for anyone.
_PINNED_SHIFT_1_RAW = os.getenv("PINNED_SHIFT_1_NAMES", "")
PINNED_SHIFT_1_LOWER = frozenset(
    x.strip().lower() for x in _PINNED_SHIFT_1_RAW.split(",") if x.strip()
)


def is_pinned_shift_1(emp_name):
    """Engineers fixed to Shift 1 (not moved off by gap-fix; merged into predefined shifts)."""
    return (emp_name or "").strip().lower() in PINNED_SHIFT_1_LOWER


def weekoffs_are_consecutive(working_days):
    """
    True if working_days lists exactly five distinct weekdays and the two days off
    are adjacent on the Monday-Sunday cycle (e.g. Sat-Sun, Fri-Sat, Sun-Mon).
    """
    ordered = [d for d in DAY_NAMES if d in working_days]
    if len(ordered) != 5:
        return False
    ws = set(ordered)
    off = [d for d in DAY_NAMES if d not in ws]
    if len(off) != 2:
        return False
    i, j = DAY_NAMES.index(off[0]), DAY_NAMES.index(off[1])
    a, b = min(i, j), max(i, j)
    return (b - a == 1) or (a == 0 and b == 6)


def month_key(year, month):
    return f"{year}-{month:02d}"


def _latest_snapshot_month_before(md, year, month):
    """
    Largest calendar month (y, m) strictly before (year, month) that has a key in md.
    Keys are expected as 'YYYY-MM' from snapshot_monthly_working_pattern.
    """
    best = None
    for key in md:
        if not isinstance(key, str):
            continue
        parts = key.split("-")
        if len(parts) != 2:
            continue
        try:
            y, m = int(parts[0]), int(parts[1])
        except ValueError:
            continue
        if m < 1 or m > 12:
            continue
        if (y < year) or (y == year and m < month):
            if best is None or (y, m) > best:
                best = (y, m)
    return best


def _months_from_to(ay, am, by, bm):
    """Number of forward calendar-month steps from month A to month B (same month -> 0)."""
    return (by - ay) * 12 + (bm - am)


def _normalize_monthly_working_dict(md):
    """Canonical 'YYYY-MM' keys so lookups match snapshot_monthly_working_pattern output."""
    if not isinstance(md, dict):
        return {}
    out = {}
    for key, val in md.items():
        if not isinstance(key, str):
            continue
        parts = key.strip().split("-")
        if len(parts) != 2:
            continue
        try:
            y, m = int(parts[0]), int(parts[1])
        except ValueError:
            continue
        if m < 1 or m > 12:
            continue
        ck = month_key(y, m)
        if isinstance(val, list):
            out[ck] = list(val)
        elif val is not None:
            out[ck] = val
    return out


def _parse_monthly_dict(emp):
    md = emp.get("monthly_working_days") or {}
    if isinstance(md, str):
        s = md.strip()
        if not s:
            return {}
        try:
            md = json.loads(s)
        except json.JSONDecodeError:
            return {}
    if not isinstance(md, dict):
        return {}
    return _normalize_monthly_working_dict(md)


def _infer_consecutive_off_block_start(working_five):
    """
    For a valid Mon-Sun week with five working days, return the DAY_NAMES index of the
    first of the two consecutive off days (off = that day and the next calendar day, wrapping).
    """
    ordered = coerce_to_five_day_pattern(working_five)
    ws = set(ordered)
    if len(ws) != 5:
        return None
    off = [d for d in DAY_NAMES if d not in ws]
    if len(off) != 2:
        return None
    i, j = DAY_NAMES.index(off[0]), DAY_NAMES.index(off[1])
    if (i + 1) % 7 == j:
        return i
    if (j + 1) % 7 == i:
        return j
    return None


def _working_days_from_off_block_start(start_idx):
    """Five working weekdays when off on start_idx and (start_idx+1) mod 7."""
    off_i = {start_idx % 7, (start_idx + 1) % 7}
    return [DAY_NAMES[i] for i in range(7) if i not in off_i]


def rotate_week_offs_forward(prev_working_days):
    """
    Move the two consecutive off-days forward by two weekdays on the cycle
    (e.g. Mon-Tue off -> Wed-Thu off), keeping five working days.
    """
    prev = coerce_to_five_day_pattern(prev_working_days)
    start = _infer_consecutive_off_block_start(prev)
    if start is None:
        return list(DEFAULT_FIVE_DAY_WEEK)
    new_start = (start + 2) % 7
    return _working_days_from_off_block_start(new_start)


def compute_mandated_shift_targets(n):
    """
    Shift sizes summing to n with ratio 6:7:8 (exact 6,7,8 when n==21).
    Uses largest-remainder fair apportionment for other team sizes.
    """
    if n <= 0:
        return {1: 0, 2: 0, 3: 0}
    w1, w2, w3 = SHIFT_MANDATE_WEIGHTS
    W = w1 + w2 + w3
    parts = [n * w1 / W, n * w2 / W, n * w3 / W]
    floors = [int(p) for p in parts]
    rem = n - sum(floors)
    order = sorted(range(3), key=lambda i: parts[i] - floors[i], reverse=True)
    for k in range(rem):
        floors[order[k]] += 1
    return {1: floors[0], 2: floors[1], 3: floors[2]}


def pattern_for_calendar_month(emp, year, month):
    """
    Working weekdays for that calendar month:
    - Stored monthly snapshot for YYYY-MM when present (written on roster Save, or
      on Generate draft so the next month can rotate without an extra Save).
    - Else: start from the latest snapshot strictly before this month, then apply
      rotate_week_offs_forward once for each calendar month from that anchor up to
      this month. So if only April exists, May rotates once from April, June twice.
    - If there is no snapshot before this month, use employee profile working_days
      (or Mon-Fri default).
    """
    md = _parse_monthly_dict(emp)
    k = month_key(year, month)
    if k in md:
        return md[k]
    anchor = _latest_snapshot_month_before(md, year, month)
    if anchor is None:
        return emp.get("working_days") or list(DEFAULT_FIVE_DAY_WEEK)
    ay, am = anchor
    anchor_key = month_key(ay, am)
    raw_anchor = md.get(anchor_key)
    if not raw_anchor:
        return emp.get("working_days") or list(DEFAULT_FIVE_DAY_WEEK)
    pat = coerce_to_five_day_pattern(raw_anchor)
    n = _months_from_to(ay, am, year, month)
    if n <= 0:
        return pat
    for _ in range(n):
        pat = rotate_week_offs_forward(pat)
    return pat


def normalize_five_day_pattern(working_days):
    """
    Return exactly five weekday names in calendar order.
    Raises ValueError if the list is not exactly five distinct working days.
    """
    if not working_days:
        return list(DEFAULT_FIVE_DAY_WEEK)
    ordered = [d for d in DAY_NAMES if d in working_days]
    if len(ordered) != 5:
        raise ValueError(
            f"Each employee must have exactly 5 working days (got {len(ordered)}: {working_days!r}). "
            "Pick two weekly off days."
        )
    if not weekoffs_are_consecutive(ordered):
        raise ValueError(
            "The two weekly off-days must be next to each other (e.g. Saturday-Sunday, "
            "Friday-Saturday, or Monday-Tuesday). They cannot be split across the week "
            f"(working days were: {ordered!r})."
        )
    return ordered


def coerce_to_five_day_pattern(working_days):
    """
    Same calendar-order five-day week as normalize, but never raises.
    Used for roster generation when monthly snapshots or legacy rows are incomplete.
    The two off-days must be consecutive; otherwise Mon-Fri is used.
    - 0 days: Mon-Fri default
    - 1-4: listed days plus next weekdays in calendar order until five (then validate)
    - 5: use as listed only if consecutive off-days; else Mon-Fri
    - 6+: first five in Mon-Sun order among those listed (then validate)
    """
    if not working_days:
        return list(DEFAULT_FIVE_DAY_WEEK)
    ordered = [d for d in DAY_NAMES if d in working_days]
    if len(ordered) >= 5:
        candidate = ordered[:5]
    else:
        for d in DAY_NAMES:
            if d not in ordered:
                ordered.append(d)
            if len(ordered) == 5:
                break
        candidate = ordered[:5]
    if weekoffs_are_consecutive(candidate):
        return candidate
    return list(DEFAULT_FIVE_DAY_WEEK)


def is_emp_scheduled_work_day(emp, d):
    """True if employee works roster shift on date d (prepared employee dict)."""
    dn = DAY_NAMES[d.weekday()]
    if dn not in emp["working_days"]:
        return False
    return d not in emp.get("_forced_off_dates", frozenset())


def _work_dates_in_iso_week_for_roster_month(emp, week_dates, roster_year, roster_month):
    """
    Dates in week_dates where the employee would work using the roster month's pattern
    only (same five weekdays for every calendar date in the ISO week).

    Using per-date calendar months for each day in the week could count more than five
    workdays when week-offs change on the 1st, which then triggered transition forced
    OFFs and produced long runs (e.g. four consecutive days off). The roster grid is
    only for one calendar month, so boundary weeks are evaluated with that month's
    pattern throughout.
    """
    pat = coerce_to_five_day_pattern(
        pattern_for_calendar_month(emp, roster_year, roster_month)
    )
    return [
        dt
        for dt in week_dates
        if DAY_NAMES[dt.weekday()] in pat
    ]


# Alias: older call sites (or stale bytecode) referenced this name.
_work_dates_in_iso_week = _work_dates_in_iso_week_for_roster_month


def _iter_boundary_weeks(year, month):
    """ISO weeks (Mon–Sun) that intersect this month and also spill into an adjacent month."""
    month_start = date(year, month, 1)
    num_days = calendar.monthrange(year, month)[1]
    month_end = date(year, month, num_days)
    seen = set()
    for dom in range(1, num_days + 1):
        d0 = date(year, month, dom)
        mon = d0 - timedelta(days=d0.weekday())
        if mon in seen:
            continue
        week = [mon + timedelta(days=i) for i in range(7)]
        seen.add(mon)
        in_month = [dt for dt in week if month_start <= dt <= month_end]
        spill = [dt for dt in week if dt < month_start or dt > month_end]
        if in_month and spill:
            yield tuple(week)


def compute_transition_forced_offs(employees, year, month):
    """
    Legacy hook: when week-offs were mixed per calendar date inside boundary ISO weeks,
    this could force extra OFF days inside the roster month. That stacked on top of the
    month's two weekly off-days and produced long OFF streaks (e.g. four days).

    Work dates in each boundary week are now counted using **this roster month's**
    pattern only, so a valid five-day week never exceeds five workdays in a Mon–Sun
    week and this pass no longer adds transition forced offs in normal operation.
    Kept for rare edge cases and warning continuity if logic changes later.
    """
    forced = defaultdict(set)
    warnings = []
    month_start = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    for week in _iter_boundary_weeks(year, month):
        for emp in employees:
            work_dates = _work_dates_in_iso_week(emp, week, year, month)
            if len(work_dates) <= 5:
                continue
            excess = len(work_dates) - 5
            in_roster = sorted(dt for dt in work_dates if month_start <= dt <= month_end)
            if len(in_roster) >= excess:
                removed = in_roster[-excess:]
                for dt in removed:
                    forced[emp["name"]].add(dt)
                warnings.append(
                    f"TRANSITION: {emp['name']} - capped ISO week at 5 work days: OFF on "
                    f"{', '.join(d.strftime('%b %d, %Y') for d in removed)} "
                    f"(week {week[0].strftime('%Y-%m-%d')}-{week[6].strftime('%Y-%m-%d')})."
                )
            else:
                for dt in in_roster:
                    forced[emp["name"]].add(dt)
                unresolved = excess - len(in_roster)
                warnings.append(
                    f"TRANSITION: {emp['name']} still has {unresolved} extra workday(s) in week "
                    f"{week[0].strftime('%Y-%m-%d')}-{week[6].strftime('%Y-%m-%d')} outside this roster month - "
                    "update monthly week-off history (save prior month roster or edit monthly patterns)."
                )

    return forced, warnings


def prepare_employees_for_roster_month(employees, year, month):
    """
    Build roster-ready employee dicts: working_days = normalized pattern for this month,
    _forced_off_dates = transition compensatory offs.
    """
    forced, transition_warnings = compute_transition_forced_offs(employees, year, month)
    prep = []
    prep_errors = []
    for e in employees:
        c = dict(e)
        raw = pattern_for_calendar_month(e, year, month)
        listed = [d for d in DAY_NAMES if d in (raw or [])]
        pat = coerce_to_five_day_pattern(raw)
        if len(listed) != 5:
            prep_errors.append(
                f"PATTERN: {e.get('name', '?')} - {len(listed)} working day(s) in month "
                f"{month_key(year, month)} {listed!r}; using {pat!r} for the roster. "
                "Edit the employee to set exactly five weekdays (two consecutive off-days)."
            )
        elif not weekoffs_are_consecutive(listed):
            prep_errors.append(
                f"PATTERN: {e.get('name', '?')} - week-offs must be two adjacent days "
                f"(e.g. Sat-Sun). Month {month_key(year, month)} had {listed!r}. "
                f"Using {pat!r} for the roster until you fix the profile."
            )
        c["working_days"] = pat
        c["_forced_off_dates"] = frozenset(forced.get(e["name"], set()))
        prep.append(c)
    return prep, transition_warnings + prep_errors


def assign_shifts(
    employees,
    night_shift_counts=None,
    prev_month_night_ids=None,
    fixed_assignments=None,
    *,
    relax_fixed_caps=False,
):
    """
    Assign each employee to a shift (1, 2, or 3).

    fixed_assignments: name -> shift for people already locked (e.g. Shift 1 pins).
    Mandated targets use len(employees) (6/7/8 for 21); remaining slots are filled
    after fixed placements.

    relax_fixed_caps: If True, do not raise when fixed_assignments already exceed
    per-shift targets (used for fully manual Edit Shifts saves).

    Strategy:
      1. Guarantee: place at least 1 employee of each product type in each shift
         without exceeding mandated per-shift caps
      2. Distribute remaining to mandated shift sizes (ratio 6:7:8 for 21 people)
      3. Night shift rotation via round-robin + avoid shift 3 if on night last month
      4. Fix daily gaps from off-day overlaps
    """
    n = len(employees)
    if n == 0:
        return {}

    prev_night = prev_month_night_ids or frozenset()
    fixed_assignments = dict(fixed_assignments or {})

    if night_shift_counts:
        sorted_emps = sorted(
            employees,
            key=lambda e: (
                night_shift_counts.get(e.get("id", 0), 0),
                1 if e.get("id") in prev_night else 0,
            ),
        )
    else:
        sorted_emps = sorted(
            employees,
            key=lambda e: (1 if e.get("id") in prev_night else 0,),
        )

    total = len(sorted_emps)
    targets = compute_mandated_shift_targets(total)

    assignments = dict(fixed_assignments)
    assigned = set(assignments.keys())
    shift_counts = {1: 0, 2: 0, 3: 0}
    for s in assignments.values():
        shift_counts[s] += 1

    for s in (1, 2, 3):
        if shift_counts[s] > targets[s] and not relax_fixed_caps:
            raise ValueError(
                f"Fixed shift assignments exceed mandated cap for shift {s}: "
                f"{shift_counts[s]} > {targets[s]} (targets {targets})."
            )

    _guarantee_type_coverage(
        sorted_emps, assignments, assigned, night_shift_counts, prev_night, targets, shift_counts
    )
    _distribute_remaining(sorted_emps, assignments, assigned, targets, night_shift_counts, prev_night)
    _fix_daily_gaps(assignments, sorted_emps)

    return assignments


def generate_roster_from_manual_assignments(
    employees,
    year,
    month,
    locked_shifts,
    prev_month_night_ids=None,
):
    """
    Build the monthly roster grid from a complete engineer name -> shift map.

    Skips mandated 6/7/8 caps (no ValueError for "too many on shift N"). Used when saving
    arbitrary Edit Shifts layouts. Coverage / type / off-day warnings are still produced.
    """
    locked_shifts = dict(locked_shifts or {})
    employees, prep_warnings = prepare_employees_for_roster_month(employees, year, month)
    shift_assignments = {}
    for e in employees:
        nm = e["name"]
        if nm not in locked_shifts:
            raise ValueError(f"Missing shift assignment for engineer: {nm}")
        shift_assignments[nm] = int(locked_shifts[nm])

    prev_night = frozenset(prev_month_night_ids or ())
    emp_lookup = {emp["name"]: emp for emp in employees}
    num_days = calendar.monthrange(year, month)[1]
    roster = {}
    warnings = list(prep_warnings)

    off_day_warnings = _check_off_day_overlaps(shift_assignments, employees)
    warnings.extend(off_day_warnings)

    type_dist_warnings = _check_type_distribution(shift_assignments, employees)
    warnings.extend(type_dist_warnings)

    if prev_night:
        emp_by_id = {e.get("id"): e for e in employees}
        for eid in prev_night:
            emp = emp_by_id.get(eid)
            if not emp:
                continue
            if shift_assignments.get(emp["name"]) == 3:
                warnings.append(
                    f"NIGHT ROTATION: {emp['name']} was on night shift (Shift 3) last month but "
                    "is still assigned to Shift 3 this month (coverage constraints). Consider manual adjustment."
                )

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]

        daily = {1: [], 2: [], 3: []}

        for emp in employees:
            if is_emp_scheduled_work_day(emp, d):
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


def _guarantee_type_coverage(
    sorted_emps, assignments, assigned, night_shift_counts, prev_month_night_ids, targets, shift_counts
):
    """
    For each product type, ensure at least 1 employee is placed in each shift.
    Does not assign past mandated targets[shift] for the full roster.
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

            def _cov_key(e):
                on_night_last = 1 if e.get("id") in prev_month_night_ids and shift_num == 3 else 0
                return (on_night_last, len(e["content_types"]))

            unassigned.sort(key=_cov_key)

            placed = False
            for emp in unassigned:
                if shift_counts[shift_num] >= targets[shift_num]:
                    continue
                assignments[emp["name"]] = shift_num
                assigned.add(emp["name"])
                shift_counts[shift_num] += 1
                placed = True
                break

            if not placed:
                cands = list(candidates)
                cands.sort(key=_cov_key)
                for emp in cands:
                    if emp["name"] in assigned and assignments[emp["name"]] != shift_num:
                        continue
                    if emp["name"] not in assigned:
                        if shift_counts[shift_num] >= targets[shift_num]:
                            continue
                        assignments[emp["name"]] = shift_num
                        assigned.add(emp["name"])
                        shift_counts[shift_num] += 1
                        break


def _distribute_remaining(sorted_emps, assignments, assigned, targets, night_shift_counts, prev_month_night_ids):
    """
    Distribute unassigned employees across shifts respecting mandated targets
    (6:7:8 ratio) and per-product-type balance.
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
        remaining.sort(
            key=lambda e: (
                night_shift_counts.get(e.get("id", 0), 0),
                1 if e.get("id") in prev_month_night_ids else 0,
            )
        )
    else:
        remaining.sort(key=lambda e: (1 if e.get("id") in prev_month_night_ids else 0,))

    for emp in remaining:
        best_shift = _pick_best_shift(
            emp, shift_counts, targets, type_shift_counts, type_targets, prev_month_night_ids
        )
        assignments[emp["name"]] = best_shift
        assigned.add(emp["name"])
        shift_counts[best_shift] += 1
        for ct in emp["content_types"]:
            type_shift_counts[ct][best_shift] += 1


def _pick_best_shift(emp, shift_counts, targets, type_shift_counts, type_targets, prev_month_night_ids):
    """
    Pick the best shift for an employee considering:
      1. Which shifts still need more people (overall targets)
      2. Which shifts need this employee's product type(s) most
      3. Penalize shift 3 if they were on night shift last month (no back-to-back months)
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

        if s == 3 and emp.get("id") in prev_month_night_ids:
            score -= 250

        scores[s] = score

    best = max(scores, key=scores.get)
    if scores[best] == -100:
        eligible = [s for s in (1, 2, 3) if shift_counts[s] < targets[s]]
        if eligible:
            return min(eligible, key=lambda s: shift_counts[s])
        return 1
    return best


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

                moved = False
                for candidate in reversed(donors_working_today):
                    if is_pinned_shift_1(candidate) and assignments.get(candidate) == 1 and shift_num != 1:
                        continue
                    remaining_donors = [n for n in donor_names if n != candidate]
                    would_break = any(
                        not any(d in emp_lookup[n]["working_days"] for n in remaining_donors)
                        for d in DAY_NAMES
                    )
                    if not would_break:
                        assignments[candidate] = shift_num
                        moved = True
                        break
                if moved:
                    break


def generate_roster(
    employees,
    year,
    month,
    night_shift_counts=None,
    predefined_shifts=None,
    prev_month_night_ids=None,
    *,
    relax_fixed_caps=False,
):
    """
    Generate a full monthly roster.

    predefined_shifts: optional fixed placements (e.g. Shift 1 pins). Everyone else
    is assigned so mandated shift sizes (6/7/8 for 21 engineers) are met.

    relax_fixed_caps: passed through to assign_shifts when fixed placements exceed targets.

    Employees are prepared with a normalized 5-day pattern for (year, month) and
    compensatory OFF days on ISO weeks that cross month boundaries when week-offs
    differ between months (effective from the 1st).

    prev_month_night_ids: employee ids who were on shift 3 last calendar month;
    they are deprioritized for shift 3 this month.

    Returns:
        roster: dict mapping date -> {shift_num: [employee_names]}
        warnings: list of warning strings for coverage gaps
        shift_assignments: dict mapping employee_name -> shift_num
    """
    employees, prep_warnings = prepare_employees_for_roster_month(employees, year, month)

    prev_night = frozenset(prev_month_night_ids or ())
    roster_size = len(employees)
    mandate_msgs = list(prep_warnings)
    if roster_size != 21:
        t = compute_mandated_shift_targets(roster_size)
        mandate_msgs.append(
            f"SHIFT MANDATE: Ideal split is 6 / 7 / 8 (21 engineers). "
            f"With {roster_size} engineer(s) using fair apportionment: "
            f"Shift 1={t[1]}, Shift 2={t[2]}, Shift 3={t[3]}."
        )

    shift_assignments = assign_shifts(
        employees,
        night_shift_counts,
        prev_night,
        fixed_assignments=dict(predefined_shifts or {}),
        relax_fixed_caps=relax_fixed_caps,
    )
    emp_lookup = {emp["name"]: emp for emp in employees}

    num_days = calendar.monthrange(year, month)[1]
    roster = {}
    warnings = list(mandate_msgs)

    off_day_warnings = _check_off_day_overlaps(shift_assignments, employees)
    warnings.extend(off_day_warnings)

    type_dist_warnings = _check_type_distribution(shift_assignments, employees)
    warnings.extend(type_dist_warnings)

    if prev_night:
        emp_by_id = {e.get("id"): e for e in employees}
        for eid in prev_night:
            emp = emp_by_id.get(eid)
            if not emp:
                continue
            if shift_assignments.get(emp["name"]) == 3:
                warnings.append(
                    f"NIGHT ROTATION: {emp['name']} was on night shift (Shift 3) last month but "
                    "is still assigned to Shift 3 this month (coverage constraints). Consider manual adjustment."
                )

    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_name = DAY_NAMES[weekday]

        daily = {1: [], 2: [], 3: []}

        for emp in employees:
            if is_emp_scheduled_work_day(emp, d):
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
