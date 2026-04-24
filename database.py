import json
import os

import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from werkzeug.security import generate_password_hash

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

DB_CONFIG = {
    "host": os.getenv("PG_HOST", "localhost"),
    "port": os.getenv("PG_PORT", "5432"),
    "user": os.getenv("PG_USER", "postgres"),
    "password": os.getenv("PG_PASSWORD", "postgres"),
    "dbname": os.getenv("PG_DATABASE", "roster_db"),
}


def get_db():
    conn = psycopg2.connect(**DB_CONFIG)
    return conn


def _fetchall(cur):
    """Convert cursor results to list of dicts."""
    if cur.description is None:
        return []
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def _fetchone(cur):
    """Convert single cursor result to dict or None."""
    if cur.description is None:
        return None
    row = cur.fetchone()
    if row is None:
        return None
    cols = [d[0] for d in cur.description]
    return dict(zip(cols, row))


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            content_types TEXT NOT NULL,
            working_days TEXT NOT NULL,
            emp_role TEXT NOT NULL DEFAULT 'engineer'
        )
    """)
    try:
        cur.execute("ALTER TABLE employees ADD COLUMN emp_role TEXT NOT NULL DEFAULT 'engineer'")
        conn.commit()
    except psycopg2.errors.DuplicateColumn:
        conn.rollback()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            product_type TEXT NOT NULL,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS rotation_history (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            shift_assigned INTEGER NOT NULL,
            UNIQUE(employee_id, year, month)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS saved_rosters (
            id SERIAL PRIMARY KEY,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            generated_at TIMESTAMP NOT NULL DEFAULT NOW(),
            excel_blob BYTEA NOT NULL,
            UNIQUE(year, month)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS saved_roster_data (
            id SERIAL PRIMARY KEY,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            roster_json TEXT NOT NULL,
            saved_at TIMESTAMP NOT NULL DEFAULT NOW(),
            UNIQUE(year, month)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL DEFAULT '',
            role TEXT NOT NULL DEFAULT 'user',
            employee_id INTEGER REFERENCES employees(id) ON DELETE SET NULL,
            created_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    try:
        cur.execute("ALTER TABLE users ADD COLUMN employee_id INTEGER REFERENCES employees(id) ON DELETE SET NULL")
        conn.commit()
    except psycopg2.errors.DuplicateColumn:
        conn.rollback()
    try:
        cur.execute(
            "ALTER TABLE employees ADD COLUMN monthly_working_days TEXT NOT NULL DEFAULT '{}'"
        )
        conn.commit()
    except psycopg2.errors.DuplicateColumn:
        conn.rollback()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS leaves (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            leave_date DATE NOT NULL,
            leave_type TEXT NOT NULL DEFAULT 'planned',
            reason TEXT DEFAULT '',
            approved_by TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT NOW(),
            UNIQUE(employee_id, leave_date)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS leave_requests (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            requested_by INTEGER REFERENCES users(id),
            leave_date DATE NOT NULL,
            leave_type TEXT NOT NULL DEFAULT 'planned',
            reason TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'pending',
            reviewed_by TEXT DEFAULT '',
            reviewed_at TIMESTAMP,
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS leave_balances (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            year INTEGER NOT NULL,
            total_allowed INTEGER DEFAULT 24,
            planned_used INTEGER DEFAULT 0,
            sick_used INTEGER DEFAULT 0,
            emergency_used INTEGER DEFAULT 0,
            UNIQUE(employee_id, year)
        )
    """)
    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        admin_user = os.getenv("ADMIN_USER", "admin")
        admin_pass = os.getenv("ADMIN_PASSWORD", "admin123")
        cur.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (%s, %s, %s, %s)",
            (admin_user, generate_password_hash(admin_pass), "Administrator", "admin")
        )
    conn.commit()
    cur.close()
    conn.close()


# ── Employee CRUD ────────────────────────────────────────

def add_employee(name, content_types, working_days, emp_role="engineer"):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO employees (name, content_types, working_days, emp_role) VALUES (%s, %s, %s, %s) RETURNING id",
            (name, json.dumps(content_types), json.dumps(working_days), emp_role)
        )
        emp_id = cur.fetchone()[0]
        conn.commit()
        return emp_id
    except psycopg2.IntegrityError:
        conn.rollback()
        return None
    finally:
        cur.close()
        conn.close()


def get_all_employees():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM employees ORDER BY id")
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return [_row_to_employee(r) for r in rows]


def get_employee_by_id(emp_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM employees WHERE id = %s", (emp_id,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return _row_to_employee(row) if row else None


def update_employee(emp_id, content_types, working_days, name=None):
    conn = get_db()
    cur = conn.cursor()
    if name:
        cur.execute(
            "UPDATE employees SET name = %s, content_types = %s, working_days = %s WHERE id = %s",
            (name, json.dumps(content_types), json.dumps(working_days), emp_id)
        )
    else:
        cur.execute(
            "UPDATE employees SET content_types = %s, working_days = %s WHERE id = %s",
            (json.dumps(content_types), json.dumps(working_days), emp_id)
        )
    conn.commit()
    cur.close()
    conn.close()


def snapshot_monthly_working_pattern(emp_id, year, month, working_days_list):
    """
    Store the canonical 5-day pattern used for a roster month (YYYY-MM key).
    Used after saving a roster so the next month can resolve week-1 transitions.
    """
    key = f"{year}-{month:02d}"
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT monthly_working_days FROM employees WHERE id = %s", (emp_id,))
    row = cur.fetchone()
    md = {}
    if row and row[0]:
        try:
            md = json.loads(row[0]) if isinstance(row[0], str) else (row[0] or {})
        except json.JSONDecodeError:
            md = {}
    if not isinstance(md, dict):
        md = {}
    md[key] = list(working_days_list)
    cur.execute(
        "UPDATE employees SET monthly_working_days = %s WHERE id = %s",
        (json.dumps(md), emp_id),
    )
    conn.commit()
    cur.close()
    conn.close()


def clear_monthly_working_snapshot_for_month(year, month):
    """
    Drop the YYYY-MM entry from every employee's monthly_working_days so the roster
    engine recomputes that month from earlier snapshots + rotation (avoids stale May
    data that matched April blocking forward week-off shifts).
    """
    keys = {f"{year}-{month:02d}", f"{year}-{month}"}
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, monthly_working_days FROM employees")
    rows = cur.fetchall()
    for emp_id, raw in rows:
        md = {}
        if raw:
            try:
                md = json.loads(raw) if isinstance(raw, str) else (raw or {})
            except json.JSONDecodeError:
                md = {}
        if not isinstance(md, dict):
            md = {}
        changed = False
        for k in list(md.keys()):
            if k in keys:
                del md[k]
                changed = True
        if changed:
            cur.execute(
                "UPDATE employees SET monthly_working_days = %s WHERE id = %s",
                (json.dumps(md), emp_id),
            )
    conn.commit()
    cur.close()
    conn.close()


def remove_employee(emp_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM employees WHERE id = %s", (emp_id,))
    conn.commit()
    cur.close()
    conn.close()


def clear_all_employees():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM employees")
    conn.commit()
    cur.close()
    conn.close()


def _row_to_employee(row):
    ct = row["content_types"]
    wd = row["working_days"]
    mwd = row.get("monthly_working_days", "{}")
    if mwd is None:
        mwd = "{}"
    if isinstance(mwd, str):
        try:
            mwd = json.loads(mwd) if mwd else {}
        except json.JSONDecodeError:
            mwd = {}
    return {
        "id": row["id"],
        "name": row["name"],
        "content_types": json.loads(ct) if isinstance(ct, str) else ct,
        "working_days": json.loads(wd) if isinstance(wd, str) else wd,
        "monthly_working_days": mwd if isinstance(mwd, dict) else {},
        "emp_role": row.get("emp_role", "engineer") if isinstance(row, dict) else "engineer",
    }


def get_employees_by_role(role):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM employees WHERE emp_role = %s ORDER BY id", (role,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return [_row_to_employee(r) for r in rows]


# ── Project CRUD ─────────────────────────────────────────

def add_project(name, product_type, employee_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO projects (name, product_type, employee_id) VALUES (%s, %s, %s)",
        (name, product_type, employee_id)
    )
    conn.commit()
    cur.close()
    conn.close()


def get_projects_for_employee(employee_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM projects WHERE employee_id = %s ORDER BY id", (employee_id,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_all_projects():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.id, p.name, p.product_type, p.employee_id, e.name AS employee_name
        FROM projects p
        JOIN employees e ON p.employee_id = e.id
        ORDER BY p.id
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def remove_project(project_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM projects WHERE id = %s", (project_id,))
    conn.commit()
    cur.close()
    conn.close()


def clear_projects_for_employee(employee_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM projects WHERE employee_id = %s", (employee_id,))
    conn.commit()
    cur.close()
    conn.close()


# ── Rotation History ─────────────────────────────────────

def save_rotation(employee_id, year, month, shift_assigned):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO rotation_history (employee_id, year, month, shift_assigned)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (employee_id, year, month)
        DO UPDATE SET shift_assigned = EXCLUDED.shift_assigned
    """, (employee_id, year, month, shift_assigned))
    conn.commit()
    cur.close()
    conn.close()


def save_all_rotations(assignments, employees, year, month):
    conn = get_db()
    cur = conn.cursor()
    name_to_id = {e["name"]: e["id"] for e in employees}
    for emp_name, shift in assignments.items():
        emp_id = name_to_id.get(emp_name)
        if emp_id:
            cur.execute("""
                INSERT INTO rotation_history (employee_id, year, month, shift_assigned)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (employee_id, year, month)
                DO UPDATE SET shift_assigned = EXCLUDED.shift_assigned
            """, (emp_id, year, month, shift))
    conn.commit()
    cur.close()
    conn.close()


def get_rotation_history():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT rh.employee_id, e.name AS employee_name, rh.year, rh.month, rh.shift_assigned
        FROM rotation_history rh
        JOIN employees e ON rh.employee_id = e.id
        ORDER BY rh.year DESC, rh.month DESC
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_night_shift_counts():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT employee_id, COUNT(*) as cnt
        FROM rotation_history
        WHERE shift_assigned = 3
        GROUP BY employee_id
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return {r["employee_id"]: r["cnt"] for r in rows}


def get_last_night_shift_month(employee_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT year, month FROM rotation_history
        WHERE employee_id = %s AND shift_assigned = 3
        ORDER BY year DESC, month DESC
        LIMIT 1
    """, (employee_id,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return (row["year"], row["month"]) if row else None


def get_shift_assignments_for_month(year, month):
    """Return {emp_name: shift_num} for all saved assignments in a month."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT e.name, rh.shift_assigned
        FROM rotation_history rh
        JOIN employees e ON rh.employee_id = e.id
        WHERE rh.year = %s AND rh.month = %s
    """, (year, month))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return {r["name"]: r["shift_assigned"] for r in rows}


def get_employee_ids_on_shift(year, month, shift_num):
    """Return set of employee ids assigned to shift_num for that calendar month."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT employee_id FROM rotation_history
        WHERE year = %s AND month = %s AND shift_assigned = %s
        """,
        (year, month, shift_num),
    )
    rows = {r["employee_id"] for r in _fetchall(cur)}
    cur.close()
    conn.close()
    return rows


def clear_shifts_for_month(year, month):
    """Delete all rotation_history rows for a given month (revert to auto)."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM rotation_history WHERE year = %s AND month = %s", (year, month))
    conn.commit()
    cur.close()
    conn.close()


def get_rotation_for_employee(employee_id, year, month):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT shift_assigned FROM rotation_history
        WHERE employee_id = %s AND year = %s AND month = %s
    """, (employee_id, year, month))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row["shift_assigned"] if row else None


# ── Saved Rosters ────────────────────────────────────────

def save_roster_excel(year, month, excel_bytes):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO saved_rosters (year, month, generated_at, excel_blob)
        VALUES (%s, %s, NOW(), %s)
        ON CONFLICT (year, month)
        DO UPDATE SET generated_at = NOW(), excel_blob = EXCLUDED.excel_blob
    """, (year, month, psycopg2.Binary(excel_bytes)))
    conn.commit()
    cur.close()
    conn.close()


def get_saved_roster(year, month):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT excel_blob FROM saved_rosters WHERE year = %s AND month = %s",
        (year, month)
    )
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return bytes(row["excel_blob"]) if row else None


def list_saved_rosters():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT year, month, generated_at
        FROM saved_rosters
        ORDER BY year DESC, month DESC
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


# ── Saved Roster Data (finalized rosters) ────────────────

def save_roster_data(year, month, roster_json_str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO saved_roster_data (year, month, roster_json, saved_at)
        VALUES (%s, %s, %s, NOW())
        ON CONFLICT (year, month)
        DO UPDATE SET roster_json = EXCLUDED.roster_json, saved_at = NOW()
    """, (year, month, roster_json_str))
    conn.commit()
    cur.close()
    conn.close()


def get_saved_roster_data(year, month):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT roster_json, saved_at FROM saved_roster_data WHERE year = %s AND month = %s",
        (year, month)
    )
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def list_finalized_rosters():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT year, month, saved_at
        FROM saved_roster_data
        ORDER BY year DESC, month DESC
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def clear_all_saved_rosters_and_rotation():
    """
    Remove every saved roster (Excel + JSON), all rotation_history rows,
    and reset per-month week-off snapshots so scheduling starts clean from profiles.
    """
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM saved_roster_data")
        cur.execute("DELETE FROM saved_rosters")
        cur.execute("DELETE FROM rotation_history")
        cur.execute("UPDATE employees SET monthly_working_days = %s", ("{}",))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


# ── Search ───────────────────────────────────────────────

def search_employees(query):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM employees WHERE name ILIKE %s ORDER BY name",
        (f"%{query}%",)
    )
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return [_row_to_employee(r) for r in rows]


def search_projects(query):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.id, p.name, p.product_type, p.employee_id, e.name AS employee_name
        FROM projects p
        JOIN employees e ON p.employee_id = e.id
        WHERE p.name ILIKE %s
        ORDER BY p.name
    """, (f"%{query}%",))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


# ── Users ─────────────────────────────────────────────────

def add_user(username, password_hash, full_name="", role="user", employee_id=None):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash, full_name, role, employee_id) VALUES (%s, %s, %s, %s, %s) RETURNING id",
            (username, password_hash, full_name, role, employee_id)
        )
        uid = cur.fetchone()[0]
        conn.commit()
        return uid
    except psycopg2.IntegrityError:
        conn.rollback()
        return None
    finally:
        cur.close()
        conn.close()


def auto_link_user(full_name):
    """
    Search employees by name (case-insensitive exact match).
    Returns the employee dict if exactly 1 match, otherwise None.
    """
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM employees WHERE LOWER(name) = LOWER(%s)",
        (full_name.strip(),)
    )
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    if len(rows) == 1:
        return _row_to_employee(rows[0])
    return None


def link_user_to_employee(user_id, employee_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET employee_id = %s WHERE id = %s", (employee_id, user_id))
    conn.commit()
    cur.close()
    conn.close()


def get_linked_employee(user_id):
    """Get the employee record linked to a user, or None."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT e.* FROM employees e
        JOIN users u ON u.employee_id = e.id
        WHERE u.id = %s
    """, (user_id,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return _row_to_employee(row) if row else None


def get_user_by_username(username):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE username = %s", (username,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def get_user_by_id(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE id = %s", (user_id,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def update_user_password(user_id, password_hash):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET password_hash = %s WHERE id = %s", (password_hash, user_id))
    conn.commit()
    cur.close()
    conn.close()


def update_user_profile(user_id, full_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET full_name = %s WHERE id = %s", (full_name, user_id))
    conn.commit()
    cur.close()
    conn.close()


def get_all_users():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, username, full_name, role, created_at FROM users ORDER BY id")
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


# ── Leaves ───────────────────────────────────────────────

def add_leave(employee_id, leave_date, leave_type, reason="", approved_by=""):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO leaves (employee_id, leave_date, leave_type, reason, approved_by)
            VALUES (%s, %s, %s, %s, %s) RETURNING id
        """, (employee_id, leave_date, leave_type, reason, approved_by))
        lid = cur.fetchone()[0]
        conn.commit()
        return lid
    except psycopg2.IntegrityError:
        conn.rollback()
        return None
    finally:
        cur.close()
        conn.close()


def cancel_leave(leave_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT employee_id, leave_date, leave_type FROM leaves WHERE id = %s", (leave_id,))
    row = _fetchone(cur)
    if row:
        cur.execute("DELETE FROM leaves WHERE id = %s", (leave_id,))
        conn.commit()
    cur.close()
    conn.close()
    return row


def get_leaves_for_month(year, month):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT l.id, l.employee_id, e.name AS employee_name, l.leave_date,
               l.leave_type, l.reason, l.approved_by, l.created_at
        FROM leaves l
        JOIN employees e ON l.employee_id = e.id
        WHERE EXTRACT(YEAR FROM l.leave_date) = %s
          AND EXTRACT(MONTH FROM l.leave_date) = %s
        ORDER BY l.leave_date, e.name
    """, (year, month))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_leaves_for_employee(employee_id, year=None):
    conn = get_db()
    cur = conn.cursor()
    if year:
        cur.execute("""
            SELECT * FROM leaves
            WHERE employee_id = %s AND EXTRACT(YEAR FROM leave_date) = %s
            ORDER BY leave_date
        """, (employee_id, year))
    else:
        cur.execute("SELECT * FROM leaves WHERE employee_id = %s ORDER BY leave_date DESC", (employee_id,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_leave_dates_map(year, month):
    """Return {employee_name: [date_str, ...]} for all leaves in a month."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT e.name, l.leave_date
        FROM leaves l
        JOIN employees e ON l.employee_id = e.id
        WHERE EXTRACT(YEAR FROM l.leave_date) = %s
          AND EXTRACT(MONTH FROM l.leave_date) = %s
        ORDER BY l.leave_date
    """, (year, month))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    result = {}
    for r in rows:
        name = r["name"]
        ld = r["leave_date"]
        date_str = ld.strftime("%Y-%m-%d") if hasattr(ld, "strftime") else str(ld)
        result.setdefault(name, []).append(date_str)
    return result


# ── Leave Requests ───────────────────────────────────────

def add_leave_request(employee_id, requested_by, leave_date, leave_type, reason=""):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO leave_requests (employee_id, requested_by, leave_date, leave_type, reason)
        VALUES (%s, %s, %s, %s, %s) RETURNING id
    """, (employee_id, requested_by, leave_date, leave_type, reason))
    rid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return rid


def get_pending_requests():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT lr.id, lr.employee_id, e.name AS employee_name, lr.leave_date,
               lr.leave_type, lr.reason, lr.status, lr.created_at,
               u.username AS requested_by_user
        FROM leave_requests lr
        JOIN employees e ON lr.employee_id = e.id
        LEFT JOIN users u ON lr.requested_by = u.id
        WHERE lr.status = 'pending'
        ORDER BY lr.leave_date
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_leave_requests_for_employee(employee_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT lr.*, e.name AS employee_name
        FROM leave_requests lr
        JOIN employees e ON lr.employee_id = e.id
        WHERE lr.employee_id = %s
        ORDER BY lr.created_at DESC
    """, (employee_id,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def approve_leave_request(request_id, reviewed_by):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE leave_requests SET status = 'approved', reviewed_by = %s, reviewed_at = NOW()
        WHERE id = %s RETURNING employee_id, leave_date, leave_type, reason
    """, (reviewed_by, request_id))
    row = cur.fetchone()
    conn.commit()
    cur.close()
    conn.close()
    return row


def reject_leave_request(request_id, reviewed_by):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE leave_requests SET status = 'rejected', reviewed_by = %s, reviewed_at = NOW()
        WHERE id = %s
    """, (reviewed_by, request_id))
    conn.commit()
    cur.close()
    conn.close()


def get_leave_request_by_id(request_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT lr.*, e.name AS employee_name
        FROM leave_requests lr
        JOIN employees e ON lr.employee_id = e.id
        WHERE lr.id = %s
    """, (request_id,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


# ── Leave Balances ───────────────────────────────────────

def get_or_create_balance(employee_id, year):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM leave_balances WHERE employee_id = %s AND year = %s",
        (employee_id, year)
    )
    row = _fetchone(cur)
    if not row:
        cur.execute("""
            INSERT INTO leave_balances (employee_id, year) VALUES (%s, %s) RETURNING *
        """, (employee_id, year))
        row = _fetchone(cur)
        conn.commit()
    cur.close()
    conn.close()
    return row


def increment_leave_used(employee_id, year, leave_type):
    conn = get_db()
    cur = conn.cursor()
    col = {"planned": "planned_used", "sick": "sick_used", "emergency": "emergency_used"}.get(leave_type)
    if not col:
        cur.close()
        conn.close()
        return
    cur.execute(f"""
        INSERT INTO leave_balances (employee_id, year, {col})
        VALUES (%s, %s, 1)
        ON CONFLICT (employee_id, year)
        DO UPDATE SET {col} = leave_balances.{col} + 1
    """, (employee_id, year))
    conn.commit()
    cur.close()
    conn.close()


def decrement_leave_used(employee_id, year, leave_type):
    conn = get_db()
    cur = conn.cursor()
    col = {"planned": "planned_used", "sick": "sick_used", "emergency": "emergency_used"}.get(leave_type)
    if not col:
        cur.close()
        conn.close()
        return
    cur.execute(f"""
        UPDATE leave_balances SET {col} = GREATEST({col} - 1, 0)
        WHERE employee_id = %s AND year = %s
    """, (employee_id, year))
    conn.commit()
    cur.close()
    conn.close()


def get_all_balances_for_year(year):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT lb.*, e.name AS employee_name
        FROM leave_balances lb
        JOIN employees e ON lb.employee_id = e.id
        WHERE lb.year = %s
        ORDER BY e.name
    """, (year,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_shift_strength(year, month, shift_assignments):
    """Return per-day, per-shift headcount accounting for weekly offs and leaves."""
    from datetime import date as dt_date

    from roster_engine import prepare_employees_for_roster_month, is_emp_scheduled_work_day

    leave_map = get_leave_dates_map(year, month)
    employees = get_employees_by_role("engineer")
    employees, _ = prepare_employees_for_roster_month(employees, year, month)

    import calendar as cal

    num_days = cal.monthrange(year, month)[1]

    strength = {}
    for day in range(1, num_days + 1):
        d = dt_date(year, month, day)
        date_str = d.strftime("%Y-%m-%d")
        daily = {1: 0, 2: 0, 3: 0}
        for emp in employees:
            shift = shift_assignments.get(emp["name"])
            if not shift:
                continue
            if not is_emp_scheduled_work_day(emp, d):
                continue
            if date_str in leave_map.get(emp["name"], []):
                continue
            daily[shift] += 1
        strength[day] = daily
    return strength
