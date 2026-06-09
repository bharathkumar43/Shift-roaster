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
    try:
        cur.execute(
            "ALTER TABLE employees ADD COLUMN learning_content_types TEXT NOT NULL DEFAULT '{}'"
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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS delta_events (
            id SERIAL PRIMARY KEY,
            project_name TEXT NOT NULL,
            product_type TEXT NOT NULL,
            start_date DATE NOT NULL,
            end_date DATE NOT NULL,
            manager_name TEXT DEFAULT '',
            start_time TEXT DEFAULT '',
            end_time TEXT DEFAULT '',
            created_by TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.commit()  # commit CREATE TABLE before ALTER TABLE migrations to avoid rollback undoing it
    for _col, _def in [
        ("manager_name", "TEXT DEFAULT ''"),
        ("start_time",   "TEXT DEFAULT ''"),
        ("end_time",     "TEXT DEFAULT ''"),
    ]:
        try:
            cur.execute(f"ALTER TABLE delta_events ADD COLUMN {_col} {_def}")
            conn.commit()
        except (psycopg2.errors.DuplicateColumn, psycopg2.errors.UndefinedTable):
            conn.rollback()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS delta_assignments (
            id SERIAL PRIMARY KEY,
            delta_event_id INTEGER NOT NULL REFERENCES delta_events(id) ON DELETE CASCADE,
            assignment_date DATE NOT NULL,
            shift_num INTEGER NOT NULL,
            engineer_name TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS project_lifecycle (
            id SERIAL PRIMARY KEY,
            project_name TEXT NOT NULL,
            product_type TEXT NOT NULL,
            migration_status TEXT NOT NULL DEFAULT 'final_validation',
            final_validation_date DATE,
            decommission_date DATE,
            UNIQUE(project_name, product_type)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS shift_config (
            id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
            shift1_target INTEGER NOT NULL DEFAULT 6,
            shift2_target INTEGER NOT NULL DEFAULT 6,
            shift3_target INTEGER NOT NULL DEFAULT 7
        )
    """)
    conn.commit()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS project_shift_handlers (
            project_name TEXT NOT NULL,
            product_type TEXT NOT NULL,
            shift_num INTEGER NOT NULL,
            handler_name TEXT NOT NULL,
            backup_handler_name TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (project_name, product_type, shift_num)
        )
    """)
    conn.commit()  # must commit CREATE TABLE before ALTER TABLE; rollback would undo the CREATE
    try:
        cur.execute(
            "ALTER TABLE project_shift_handlers ADD COLUMN backup_handler_name TEXT NOT NULL DEFAULT ''"
        )
        conn.commit()
    except (psycopg2.errors.DuplicateColumn, psycopg2.errors.UndefinedTable):
        conn.rollback()
    try:
        cur.execute("ALTER TABLE delta_events ADD COLUMN delta_status TEXT NOT NULL DEFAULT 'active'")
        conn.commit()
    except (psycopg2.errors.DuplicateColumn, psycopg2.errors.UndefinedTable):
        conn.rollback()
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
    sh_init_tables()


# ── Shift Configuration ──────────────────────────────────

def get_shift_config():
    """Return {1: n, 2: n, 3: n} target headcounts, or the default (6,6,7) if not set."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT shift1_target, shift2_target, shift3_target FROM shift_config WHERE id = 1")
    row = _fetchone(cur)
    cur.close()
    conn.close()
    if row:
        return {1: row["shift1_target"], 2: row["shift2_target"], 3: row["shift3_target"]}
    return {1: 6, 2: 6, 3: 7}


def save_shift_config(s1, s2, s3):
    """Persist admin-configured shift target headcounts."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO shift_config (id, shift1_target, shift2_target, shift3_target)
        VALUES (1, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            shift1_target = EXCLUDED.shift1_target,
            shift2_target = EXCLUDED.shift2_target,
            shift3_target = EXCLUDED.shift3_target
    """, (s1, s2, s3))
    conn.commit()
    cur.close()
    conn.close()


# ── Project Shift Handlers ───────────────────────────────

def get_project_shift_handlers():
    """Return {(project_name, product_type, shift_num): (handler_name, backup_handler_name|None)}."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT project_name, product_type, shift_num, handler_name, backup_handler_name "
        "FROM project_shift_handlers"
    )
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return {
        (r["project_name"], r["product_type"], r["shift_num"]): (
            r["handler_name"],
            r["backup_handler_name"] or None,
        )
        for r in rows
    }


def clear_project_shift_handlers():
    """Delete all saved handler assignments so they are recomputed fresh."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM project_shift_handlers")
    conn.commit()
    cur.close()
    conn.close()


def save_project_shift_handlers(handlers):
    """Upsert {(project_name, product_type, shift_num): (primary, backup|None)}."""
    if not handlers:
        return
    conn = get_db()
    cur = conn.cursor()
    for (proj_name, prod_type, shift_num), handler_info in handlers.items():
        if isinstance(handler_info, tuple):
            primary, backup = handler_info
        else:
            primary, backup = handler_info, None
        cur.execute("""
            INSERT INTO project_shift_handlers
                (project_name, product_type, shift_num, handler_name, backup_handler_name)
            VALUES (%s, %s, %s, %s, %s)
            ON CONFLICT (project_name, product_type, shift_num)
            DO UPDATE SET handler_name = EXCLUDED.handler_name,
                          backup_handler_name = EXCLUDED.backup_handler_name
        """, (proj_name, prod_type, shift_num, primary or '', backup or ''))
    conn.commit()
    cur.close()
    conn.close()


# ── Employee CRUD ────────────────────────────────────────

def add_employee(name, content_types, working_days, emp_role="engineer", learning_content_types=None):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO employees (name, content_types, working_days, emp_role, learning_content_types) VALUES (%s, %s, %s, %s, %s) RETURNING id",
            (name, json.dumps(content_types), json.dumps(working_days), emp_role, json.dumps({}))
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


def update_employee(emp_id, content_types, working_days, name=None, learning_content_types=None):
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
    lct = row.get("learning_content_types", "{}")
    if mwd is None:
        mwd = "{}"
    if isinstance(mwd, str):
        try:
            mwd = json.loads(mwd) if mwd else {}
        except json.JSONDecodeError:
            mwd = {}
    if lct is None:
        lct = "{}"
    if isinstance(lct, str):
        try:
            lct = json.loads(lct) if lct else {}
        except json.JSONDecodeError:
            lct = {}
    return {
        "id": row["id"],
        "name": row["name"],
        "content_types": json.loads(ct) if isinstance(ct, str) else ct,
        "working_days": json.loads(wd) if isinstance(wd, str) else wd,
        "monthly_working_days": mwd if isinstance(mwd, dict) else {},
        "learning_content_types": lct if isinstance(lct, dict) else {},
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


def get_user_by_username_ci(username):
    """Case-insensitive username lookup — used during Microsoft login migration."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(%s)", (username,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def update_user_username(user_id, new_username):
    """Rename a user's username — called once per user during Microsoft login migration."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET username = %s WHERE id = %s", (new_username, user_id))
    conn.commit()
    cur.close()
    conn.close()


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


# ── Delta Calendar CRUD ──────────────────────────────────

def add_delta_event(project_name, product_type, start_date, end_date, created_by="", manager_name="", start_time="", end_time=""):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO delta_events (project_name, product_type, start_date, end_date, created_by, manager_name, start_time, end_time) "
        "VALUES (%s, %s, %s, %s, %s, %s, %s, %s) RETURNING id",
        (project_name, product_type, start_date, end_date, created_by, manager_name, start_time, end_time)
    )
    delta_id = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return delta_id


def get_all_delta_events():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM delta_events ORDER BY start_date DESC, id DESC")
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_delta_events_by_product_type(product_type):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM delta_events WHERE product_type = %s ORDER BY start_date, id",
        (product_type,)
    )
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def save_delta_assignments(delta_event_id, assignments):
    """assignments: list of {date, shift_num, engineer_name}"""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM delta_assignments WHERE delta_event_id = %s", (delta_event_id,))
    for a in assignments:
        cur.execute(
            "INSERT INTO delta_assignments (delta_event_id, assignment_date, shift_num, engineer_name) "
            "VALUES (%s, %s, %s, %s)",
            (delta_event_id, a["date"], a["shift_num"], a["engineer_name"])
        )
    conn.commit()
    cur.close()
    conn.close()


def get_delta_assignments(delta_event_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM delta_assignments WHERE delta_event_id = %s ORDER BY assignment_date, shift_num",
        (delta_event_id,)
    )
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def delete_delta_event(delta_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM delta_events WHERE id = %s", (delta_id,))
    conn.commit()
    cur.close()
    conn.close()


def update_delta_status(delta_id, status):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE delta_events SET delta_status = %s WHERE id = %s", (status, delta_id))
    conn.commit()
    cur.close()
    conn.close()


def update_delta_events_status_by_project(project_name, product_type, status):
    """Update delta_status on all delta_events for a given project/product_type."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE delta_events SET delta_status = %s WHERE project_name = %s AND product_type = %s",
        (status, project_name, product_type)
    )
    conn.commit()
    cur.close()
    conn.close()


def upsert_project_lifecycle(project_name, product_type, migration_status, final_validation_date=None, decommission_date=None):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO project_lifecycle (project_name, product_type, migration_status, final_validation_date, decommission_date)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (project_name, product_type) DO UPDATE
            SET migration_status = EXCLUDED.migration_status,
                final_validation_date = COALESCE(EXCLUDED.final_validation_date, project_lifecycle.final_validation_date),
                decommission_date = COALESCE(EXCLUDED.decommission_date, project_lifecycle.decommission_date)
    """, (project_name, product_type, migration_status, final_validation_date, decommission_date))
    conn.commit()
    cur.close()
    conn.close()


def get_project_lifecycle_by_status(status):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM project_lifecycle WHERE migration_status = %s ORDER BY project_name, product_type",
        (status,)
    )
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def get_excluded_projects():
    """Return set of (project_name, product_type) that are final_validation or decommissioned."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT project_name, product_type FROM project_lifecycle WHERE migration_status IN ('final_validation', 'decommissioned')"
    )
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return {(r[0], r[1]) for r in rows}


def delete_project_lifecycle(project_name, product_type):
    """Remove a project from project_lifecycle so it becomes active again."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "DELETE FROM project_lifecycle WHERE project_name = %s AND product_type = %s",
        (project_name, product_type),
    )
    conn.commit()
    cur.close()
    conn.close()


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


# ── Shift Handover ──────────────────────────────────────────────────────────────

def sh_init_tables():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_handovers (
            id SERIAL PRIMARY KEY,
            handover_date DATE NOT NULL,
            shift_num INTEGER NOT NULL CHECK (shift_num IN (1,2,3)),
            project_name TEXT NOT NULL,
            submitted_by_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
            submitted_by_name TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT 'draft',
            lead_notes TEXT DEFAULT '',
            engineer_acknowledged BOOLEAN NOT NULL DEFAULT FALSE,
            engineer_acknowledged_by TEXT DEFAULT '',
            engineer_acknowledged_at TIMESTAMP,
            manager_acknowledged BOOLEAN NOT NULL DEFAULT FALSE,
            manager_acknowledged_by TEXT DEFAULT '',
            manager_acknowledged_at TIMESTAMP,
            created_at TIMESTAMP NOT NULL DEFAULT NOW(),
            updated_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    conn.commit()

    # Add UNIQUE constraint if missing
    try:
        cur.execute("""
            ALTER TABLE sh_handovers
            ADD CONSTRAINT sh_handovers_unique_slot
            UNIQUE (handover_date, shift_num, project_name)
        """)
        conn.commit()
    except Exception:
        conn.rollback()

    # Migrate old columns / add new columns to sh_handovers
    for _col, _def in [
        ("lead_notes",                 "TEXT DEFAULT ''"),
        ("engineer_acknowledged",      "BOOLEAN NOT NULL DEFAULT FALSE"),
        ("engineer_acknowledged_by",   "TEXT DEFAULT ''"),
        ("engineer_acknowledged_at",   "TIMESTAMP"),
        ("manager_acknowledged",       "BOOLEAN NOT NULL DEFAULT FALSE"),
        ("manager_acknowledged_by",    "TEXT DEFAULT ''"),
        ("manager_acknowledged_at",    "TIMESTAMP"),
        ("updated_at",                 "TIMESTAMP NOT NULL DEFAULT NOW()"),
    ]:
        try:
            cur.execute(f"ALTER TABLE sh_handovers ADD COLUMN {_col} {_def}")
            conn.commit()
        except Exception:
            conn.rollback()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_client_entries (
            id SERIAL PRIMARY KEY,
            handover_id INTEGER NOT NULL REFERENCES sh_handovers(id) ON DELETE CASCADE,
            client_name TEXT NOT NULL,
            tickets TEXT DEFAULT '',
            entry_status TEXT NOT NULL DEFAULT 'NA',
            engineer_worked TEXT DEFAULT '',
            issues TEXT DEFAULT '',
            engineer_notes TEXT DEFAULT '',
            manager_notes TEXT DEFAULT '',
            next_shift_engineer TEXT DEFAULT '',
            migration_report_sent BOOLEAN NOT NULL DEFAULT FALSE,
            drive_changes_alerts BOOLEAN NOT NULL DEFAULT FALSE,
            row_tint TEXT DEFAULT NULL,
            filled_by_name TEXT DEFAULT '',
            created_at TIMESTAMP NOT NULL DEFAULT NOW(),
            updated_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    conn.commit()

    # Add new columns to sh_client_entries if old schema
    for _col, _def in [
        ("tickets",               "TEXT DEFAULT ''"),
        ("engineer_worked",       "TEXT DEFAULT ''"),
        ("issues",                "TEXT DEFAULT ''"),
        ("engineer_notes",        "TEXT DEFAULT ''"),
        ("manager_notes",         "TEXT DEFAULT ''"),
        ("next_shift_engineer",   "TEXT DEFAULT ''"),
        ("migration_report_sent", "BOOLEAN NOT NULL DEFAULT FALSE"),
        ("drive_changes_alerts",  "BOOLEAN NOT NULL DEFAULT FALSE"),
        ("row_tint",              "TEXT DEFAULT NULL"),
        ("filled_by_name",        "TEXT DEFAULT ''"),
        ("updated_at",            "TIMESTAMP NOT NULL DEFAULT NOW()"),
    ]:
        try:
            cur.execute(f"ALTER TABLE sh_client_entries ADD COLUMN {_col} {_def}")
            conn.commit()
        except Exception:
            conn.rollback()

    # Add UNIQUE on (handover_id, client_name) if missing
    try:
        cur.execute("""
            ALTER TABLE sh_client_entries
            ADD CONSTRAINT sh_client_entries_unique
            UNIQUE (handover_id, client_name)
        """)
        conn.commit()
    except Exception:
        conn.rollback()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_mom (
            id SERIAL PRIMARY KEY,
            handover_id INTEGER NOT NULL REFERENCES sh_handovers(id) ON DELETE CASCADE,
            client_name TEXT NOT NULL DEFAULT '',
            file_name TEXT NOT NULL,
            file_data BYTEA NOT NULL,
            notes TEXT DEFAULT '',
            uploaded_by_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
            uploaded_by_name TEXT NOT NULL DEFAULT '',
            uploaded_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    conn.commit()

    for _col, _def in [
        ("client_name", "TEXT NOT NULL DEFAULT ''"),
        ("notes",       "TEXT DEFAULT ''"),
    ]:
        try:
            cur.execute(f"ALTER TABLE sh_mom ADD COLUMN {_col} {_def}")
            conn.commit()
        except Exception:
            conn.rollback()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_daily_notes (
            id SERIAL PRIMARY KEY,
            note_date DATE NOT NULL UNIQUE,
            duty_manager TEXT DEFAULT '',
            week_label TEXT DEFAULT '',
            key_issues TEXT DEFAULT '',
            actions_for_tomorrow TEXT DEFAULT '',
            updated_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    conn.commit()

    # sh_projects: shift handover project/department names
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_projects (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            sort_order INTEGER NOT NULL DEFAULT 0,
            created_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    conn.commit()

    # Seed default projects if the table is empty
    cur.execute("SELECT COUNT(*) FROM sh_projects")
    if cur.fetchone()[0] == 0:
        for i, pname in enumerate(["Content", "Email", "Messaging"], start=1):
            cur.execute(
                "INSERT INTO sh_projects (name, sort_order) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                (pname, i)
            )
        conn.commit()

    # sh_clients: managed client list per project
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_clients (
            id SERIAL PRIMARY KEY,
            project_name TEXT NOT NULL,
            client_name TEXT NOT NULL,
            is_active BOOLEAN NOT NULL DEFAULT TRUE,
            created_at TIMESTAMP NOT NULL DEFAULT NOW(),
            CONSTRAINT sh_clients_unique UNIQUE (project_name, client_name)
        )
    """)
    conn.commit()

    # sh_user_shifts: which shifts each user is assigned to
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sh_user_shifts (
            user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
            shift_1 BOOLEAN NOT NULL DEFAULT FALSE,
            shift_2 BOOLEAN NOT NULL DEFAULT FALSE,
            shift_3 BOOLEAN NOT NULL DEFAULT FALSE
        )
    """)
    conn.commit()

    # Add email and is_active columns to users table if missing
    for _col, _def in [
        ("email",     "TEXT DEFAULT ''"),
        ("is_active", "BOOLEAN NOT NULL DEFAULT TRUE"),
    ]:
        try:
            cur.execute(f"ALTER TABLE users ADD COLUMN {_col} {_def}")
            conn.commit()
        except Exception:
            conn.rollback()

    cur.close()
    conn.close()


def sh_get_project_names():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT name FROM sh_projects ORDER BY sort_order, name")
    rows = [r[0] for r in cur.fetchall()]
    cur.close()
    conn.close()
    return rows


def sh_add_project(name):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO sh_projects (name, sort_order)
            VALUES (%s, (SELECT COALESCE(MAX(sort_order),0)+1 FROM sh_projects))
            ON CONFLICT (name) DO NOTHING
        """, (name.strip(),))
        conn.commit()
    except Exception:
        conn.rollback()
    cur.close()
    conn.close()


def sh_delete_project(name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM sh_projects WHERE name = %s", (name,))
    conn.commit()
    cur.close()
    conn.close()


def sh_get_project_clients(project_name):
    """Return all distinct client names ever entered for a project, most recent first."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT ce.client_name
        FROM sh_client_entries ce
        JOIN sh_handovers h ON ce.handover_id = h.id
        WHERE h.project_name = %s
        ORDER BY ce.client_name
    """, (project_name,))
    rows = [r[0] for r in cur.fetchall()]
    cur.close()
    conn.close()
    return rows


def sh_get_or_create_handover(handover_date, shift_num, project_name, user_id, user_name):
    """Get existing draft/submitted handover or create a new draft."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM sh_handovers
        WHERE handover_date = %s AND shift_num = %s AND project_name = %s
    """, (handover_date, shift_num, project_name))
    row = _fetchone(cur)
    if not row:
        cur.execute("""
            INSERT INTO sh_handovers
                (handover_date, shift_num, project_name, submitted_by_user_id,
                 submitted_by_name, status)
            VALUES (%s, %s, %s, %s, %s, 'draft')
            ON CONFLICT (handover_date, shift_num, project_name) DO NOTHING
            RETURNING id
        """, (handover_date, shift_num, project_name, user_id, user_name))
        r = cur.fetchone()
        conn.commit()
        if r:
            handover_id = r[0]
        else:
            cur.execute("""
                SELECT id FROM sh_handovers
                WHERE handover_date = %s AND shift_num = %s AND project_name = %s
            """, (handover_date, shift_num, project_name))
            handover_id = cur.fetchone()[0]
        cur.execute("SELECT * FROM sh_handovers WHERE id = %s", (handover_id,))
        row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def sh_save_entries(handover_id, entries, user_name, lead_notes=None):
    """Upsert client entries for a handover. Only updates fields the user sent."""
    conn = get_db()
    cur = conn.cursor()
    if lead_notes is not None:
        cur.execute(
            "UPDATE sh_handovers SET lead_notes = %s, updated_at = NOW() WHERE id = %s",
            (lead_notes, handover_id)
        )
    for entry in entries:
        cn = (entry.get("client_name") or "").strip()
        if not cn:
            continue
        cur.execute("""
            INSERT INTO sh_client_entries
                (handover_id, client_name, tickets, entry_status,
                 engineer_worked, issues, engineer_notes, manager_notes,
                 next_shift_engineer, migration_report_sent, drive_changes_alerts,
                 row_tint, filled_by_name)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT (handover_id, client_name)
            DO UPDATE SET
                tickets               = EXCLUDED.tickets,
                entry_status          = EXCLUDED.entry_status,
                engineer_worked       = EXCLUDED.engineer_worked,
                issues                = EXCLUDED.issues,
                engineer_notes        = EXCLUDED.engineer_notes,
                manager_notes         = EXCLUDED.manager_notes,
                next_shift_engineer   = EXCLUDED.next_shift_engineer,
                migration_report_sent = EXCLUDED.migration_report_sent,
                drive_changes_alerts  = EXCLUDED.drive_changes_alerts,
                row_tint              = EXCLUDED.row_tint,
                filled_by_name        = EXCLUDED.filled_by_name,
                updated_at            = NOW()
        """, (
            handover_id,
            cn,
            entry.get("tickets", ""),
            entry.get("entry_status", "NA"),
            entry.get("engineer_worked", ""),
            entry.get("issues", ""),
            entry.get("engineer_notes", ""),
            entry.get("manager_notes", ""),
            entry.get("next_shift_engineer", ""),
            entry.get("migration_report_sent", False),
            entry.get("drive_changes_alerts", False),
            entry.get("row_tint"),
            user_name,
        ))
    conn.commit()
    cur.close()
    conn.close()


def sh_submit_handover(handover_id, user_id, user_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE sh_handovers
        SET status = 'submitted',
            submitted_by_user_id = %s,
            submitted_by_name = %s,
            engineer_acknowledged = FALSE,
            manager_acknowledged = FALSE,
            updated_at = NOW()
        WHERE id = %s
    """, (user_id, user_name, handover_id))
    conn.commit()
    cur.close()
    conn.close()


def sh_engineer_ack(handover_id, user_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE sh_handovers
        SET engineer_acknowledged = TRUE,
            engineer_acknowledged_by = %s,
            engineer_acknowledged_at = NOW(),
            updated_at = NOW()
        WHERE id = %s
    """, (user_name, handover_id))
    conn.commit()
    cur.close()
    conn.close()


def sh_manager_ack(handover_id, user_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE sh_handovers
        SET manager_acknowledged = TRUE,
            manager_acknowledged_by = %s,
            manager_acknowledged_at = NOW(),
            updated_at = NOW()
        WHERE id = %s
    """, (user_name, handover_id))
    conn.commit()
    cur.close()
    conn.close()


def sh_get_handover(handover_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM sh_handovers WHERE id = %s", (handover_id,))
    handover = _fetchone(cur)
    if not handover:
        cur.close()
        conn.close()
        return None
    cur.execute("""
        SELECT * FROM sh_client_entries WHERE handover_id = %s ORDER BY client_name
    """, (handover_id,))
    handover["client_entries"] = _fetchall(cur)
    cur.execute("""
        SELECT id, client_name, file_name, notes, uploaded_by_name, uploaded_at,
               octet_length(file_data) AS file_size
        FROM sh_mom WHERE handover_id = %s ORDER BY uploaded_at
    """, (handover_id,))
    handover["moms"] = _fetchall(cur)
    cur.close()
    conn.close()
    return handover


def sh_get_handover_by_slot(handover_date, shift_num, project_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM sh_handovers
        WHERE handover_date = %s AND shift_num = %s AND project_name = %s
    """, (handover_date, shift_num, project_name))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def sh_get_handovers(project_name=None, shift_num=None, status=None,
                     date_from=None, date_to=None, limit=20, offset=0):
    conn = get_db()
    cur = conn.cursor()
    wheres, params = [], []
    if project_name:
        wheres.append("project_name = %s"); params.append(project_name)
    if shift_num:
        wheres.append("shift_num = %s"); params.append(int(shift_num))
    if status:
        wheres.append("status = %s"); params.append(status)
    if date_from:
        wheres.append("handover_date >= %s"); params.append(date_from)
    if date_to:
        wheres.append("handover_date <= %s"); params.append(date_to)
    where_clause = ("WHERE " + " AND ".join(wheres)) if wheres else ""
    cur.execute(f"""
        SELECT * FROM sh_handovers {where_clause}
        ORDER BY handover_date DESC, shift_num, id DESC
        LIMIT %s OFFSET %s
    """, params + [limit, offset])
    rows = _fetchall(cur)
    cur.execute(f"SELECT COUNT(*) FROM sh_handovers {where_clause}", params)
    total = cur.fetchone()[0]
    cur.close()
    conn.close()
    return rows, total


def sh_upload_mom(handover_id, client_name, file_name, file_data, notes, user_id, user_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO sh_mom
            (handover_id, client_name, file_name, file_data, notes,
             uploaded_by_user_id, uploaded_by_name)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
    """, (handover_id, client_name, file_name,
          psycopg2.Binary(file_data), notes, user_id, user_name))
    conn.commit()
    cur.close()
    conn.close()


def sh_delete_mom(mom_id, user_id, is_admin):
    conn = get_db()
    cur = conn.cursor()
    if is_admin:
        cur.execute("DELETE FROM sh_mom WHERE id = %s", (mom_id,))
    else:
        cur.execute(
            "DELETE FROM sh_mom WHERE id = %s AND uploaded_by_user_id = %s",
            (mom_id, user_id)
        )
    conn.commit()
    cur.close()
    conn.close()


def sh_get_mom_file(mom_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT file_name, file_data FROM sh_mom WHERE id = %s", (mom_id,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    if not row:
        return None, None
    return row[0], bytes(row[1])


def sh_daily_dashboard(dashboard_date):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT h.*,
               (SELECT COUNT(*) FROM sh_client_entries ce WHERE ce.handover_id = h.id) AS entry_count
        FROM sh_handovers h
        WHERE h.handover_date = %s
        ORDER BY h.project_name, h.shift_num
    """, (dashboard_date,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    projects = sh_get_project_names()
    dashboard = {p: {1: None, 2: None, 3: None} for p in projects}
    for row in rows:
        proj, shift = row["project_name"], row["shift_num"]
        if proj in dashboard:
            dashboard[proj][shift] = row
    return dashboard


def sh_get_daily_notes(note_date):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM sh_daily_notes WHERE note_date = %s", (note_date,))
    row = _fetchone(cur)
    cur.close()
    conn.close()
    return row


def sh_upsert_daily_notes(note_date, duty_manager, week_label, key_issues, actions_for_tomorrow):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO sh_daily_notes
            (note_date, duty_manager, week_label, key_issues, actions_for_tomorrow)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (note_date) DO UPDATE SET
            duty_manager         = EXCLUDED.duty_manager,
            week_label           = EXCLUDED.week_label,
            key_issues           = EXCLUDED.key_issues,
            actions_for_tomorrow = EXCLUDED.actions_for_tomorrow,
            updated_at           = NOW()
    """, (note_date, duty_manager, week_label, key_issues, actions_for_tomorrow))
    conn.commit()
    cur.close()
    conn.close()


def sh_compliance_data(target_date):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT h.id, h.project_name, h.shift_num, h.status,
               h.engineer_acknowledged, h.manager_acknowledged,
               COUNT(ce.id) AS total_entries,
               SUM(CASE WHEN ce.engineer_notes <> '' THEN 1 ELSE 0 END) AS filled_entries
        FROM sh_handovers h
        LEFT JOIN sh_client_entries ce ON ce.handover_id = h.id
        WHERE h.handover_date = %s
        GROUP BY h.id, h.project_name, h.shift_num, h.status,
                 h.engineer_acknowledged, h.manager_acknowledged
        ORDER BY h.project_name, h.shift_num
    """, (target_date,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_dashboard_stats(dashboard_date):
    """Return total ticket count, open issues, resolved counts for the day."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT
            COUNT(ce.id) AS total_entries,
            SUM(CASE WHEN ce.entry_status NOT IN ('Completed') THEN 1 ELSE 0 END) AS open_issues,
            SUM(CASE WHEN ce.entry_status = 'Completed' THEN 1 ELSE 0 END) AS resolved
        FROM sh_client_entries ce
        JOIN sh_handovers h ON ce.handover_id = h.id
        WHERE h.handover_date = %s
    """, (dashboard_date,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return {
        "total": row[0] or 0,
        "open":  row[1] or 0,
        "resolved": row[2] or 0,
    }


# ── sh_clients management ─────────────────────────────────────────────────────

def sh_get_clients(project_name=None):
    conn = get_db()
    cur = conn.cursor()
    if project_name:
        cur.execute("""
            SELECT * FROM sh_clients WHERE project_name = %s ORDER BY client_name
        """, (project_name,))
    else:
        cur.execute("SELECT * FROM sh_clients ORDER BY project_name, client_name")
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_add_client(project_name, client_name):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO sh_clients (project_name, client_name)
            VALUES (%s, %s)
            ON CONFLICT (project_name, client_name) DO UPDATE SET is_active = TRUE
        """, (project_name, client_name.strip()))
        conn.commit()
    except Exception:
        conn.rollback()
    cur.close()
    conn.close()


def sh_toggle_client(client_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE sh_clients SET is_active = NOT is_active WHERE id = %s
        RETURNING is_active
    """, (client_id,))
    row = cur.fetchone()
    conn.commit()
    cur.close()
    conn.close()
    return row[0] if row else None


def sh_delete_client(client_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM sh_clients WHERE id = %s", (client_id,))
    conn.commit()
    cur.close()
    conn.close()


def sh_get_engineers():
    """Return list of user display names for dropdowns."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT COALESCE(NULLIF(full_name,''), username) AS name
        FROM users
        WHERE COALESCE(is_active, TRUE) = TRUE
        ORDER BY name
    """)
    rows = [r[0] for r in cur.fetchall()]
    cur.close()
    conn.close()
    return rows


_SH_PROJECT_TO_PRODUCT_TYPE = {
    "messaging": "Message",
    "message":   "Message",
    "email":     "Email",
    "content":   "Content",
}


def _sh_build_coverage_for_dates(year_month_pairs):
    """Run generate_project_coverage once per unique (year, month) and return merged day list."""
    from project_engine import generate_project_coverage

    projects  = get_all_projects()
    employees = get_employees_by_role("engineer")
    all_days  = {}  # (year, month, day_num) -> day_info

    for year, month in set(year_month_pairs):
        shift_map = get_shift_assignments_for_month(year, month)
        if not shift_map:
            continue
        coverage, _ = generate_project_coverage(projects, employees, shift_map, year, month)
        for day_info in coverage:
            all_days[(year, month, day_info["day_num"])] = day_info

    return all_days


def sh_get_client_shift_handlers(date_str, shift_num, project_name):
    """
    Return two dicts keyed by lowercase client/project name:
      curr_map  : {name_lower: engineer}  — handler in shift_num on date_str
      next_map  : {name_lower: engineer}  — handler in the next shift
      next_label: human-readable label for the next shift

    Engineers are looked up per individual project from generate_project_coverage so
    different clients in the same form get their own specific engineer.
    """
    from datetime import date as _date, timedelta

    d = _date.fromisoformat(date_str)
    if shift_num < 3:
        next_shift = shift_num + 1
        next_date  = d
    else:
        next_shift = 1
        next_date  = d + timedelta(days=1)

    shift_names = {1: "Morning", 2: "Afternoon", 3: "Night"}
    next_label  = f"{shift_names.get(next_shift, f'Shift {next_shift}')} · {next_date.strftime('%b %d')}"

    pt_filter = _SH_PROJECT_TO_PRODUCT_TYPE.get(project_name.lower(), project_name)

    pairs = [(d.year, d.month)]
    if (next_date.year, next_date.month) != (d.year, d.month):
        pairs.append((next_date.year, next_date.month))

    all_days = _sh_build_coverage_for_dates(pairs)

    def _build_map(lookup_date, s_num):
        key      = (lookup_date.year, lookup_date.month, lookup_date.day)
        day_info = all_days.get(key)
        if not day_info:
            return {}
        result = {}
        for proj in day_info["projects"]:
            if proj["product_type"] != pt_filter:
                continue
            handler = proj["shifts"].get(s_num, {}).get("handler")
            if handler:
                result[proj["project_name"].lower().strip()] = handler
        return result

    curr_map = _build_map(d, shift_num)
    next_map = _build_map(next_date, next_shift)

    return curr_map, next_map, next_label


def sh_get_engineers_on_shift(date_str, shift_num, project_name):
    """Unique engineer names handling project_name in shift_num on date_str (ordered by first seen)."""
    curr_map, _, _ = sh_get_client_shift_handlers(date_str, shift_num, project_name)
    seen, result = set(), []
    for eng in curr_map.values():
        if eng not in seen:
            seen.add(eng)
            result.append(eng)
    return result


def sh_get_next_shift_engineers(date_str, current_shift_num, project_name=""):
    """Return (unique_engineers_list, label) for the shift after current_shift_num."""
    _, next_map, label = sh_get_client_shift_handlers(date_str, current_shift_num, project_name)
    seen, result = set(), []
    for eng in next_map.values():
        if eng not in seen:
            seen.add(eng)
            result.append(eng)
    return result, label


def sh_get_active_clients(project_name):
    """Return active clients from sh_clients for a project, fall back to historical."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT client_name FROM sh_clients
        WHERE project_name = %s AND is_active = TRUE
        ORDER BY client_name
    """, (project_name,))
    rows = [r[0] for r in cur.fetchall()]
    if not rows:
        # Fall back to historically-used client names
        cur.execute("""
            SELECT DISTINCT ce.client_name
            FROM sh_client_entries ce
            JOIN sh_handovers h ON ce.handover_id = h.id
            WHERE h.project_name = %s
            ORDER BY ce.client_name
        """, (project_name,))
        rows = [r[0] for r in cur.fetchall()]
    cur.close()
    conn.close()
    return rows


# ── sh_users management ───────────────────────────────────────────────────────

def sh_get_all_users_with_shifts():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT ON (LOWER(COALESCE(NULLIF(u.email, ''), u.username)))
               u.id, u.username, u.full_name,
               COALESCE(u.email, '') AS email,
               u.role,
               COALESCE(u.is_active, TRUE) AS is_active,
               COALESCE(s.shift_1, FALSE) AS shift_1,
               COALESCE(s.shift_2, FALSE) AS shift_2,
               COALESCE(s.shift_3, FALSE) AS shift_3
        FROM users u
        LEFT JOIN sh_user_shifts s ON s.user_id = u.id
        ORDER BY LOWER(COALESCE(NULLIF(u.email, ''), u.username)), u.id
    """)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_update_user_role(user_id, new_role):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET role = %s WHERE id = %s", (new_role, user_id))
    conn.commit()
    cur.close()
    conn.close()


def sh_update_user_shifts(user_id, shift_1, shift_2, shift_3):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO sh_user_shifts (user_id, shift_1, shift_2, shift_3)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (user_id) DO UPDATE SET
            shift_1 = EXCLUDED.shift_1,
            shift_2 = EXCLUDED.shift_2,
            shift_3 = EXCLUDED.shift_3
    """, (user_id, shift_1, shift_2, shift_3))
    conn.commit()
    cur.close()
    conn.close()


def sh_toggle_user_status(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE users SET is_active = NOT COALESCE(is_active, TRUE)
        WHERE id = %s RETURNING is_active
    """, (user_id,))
    row = cur.fetchone()
    conn.commit()
    cur.close()
    conn.close()
    return row[0] if row else None


def sh_delete_user(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM sh_user_shifts WHERE user_id = %s", (user_id,))
    cur.execute("DELETE FROM users WHERE id = %s", (user_id,))
    conn.commit()
    cur.close()
    conn.close()


# ── Reports ───────────────────────────────────────────────────────────────────

def sh_report_by_date(report_date):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT h.handover_date, h.shift_num, h.project_name,
               h.submitted_by_name, h.status,
               ce.client_name, ce.entry_status, ce.engineer_worked,
               ce.issues, ce.engineer_notes, ce.tickets, ce.next_shift_engineer
        FROM sh_handovers h
        LEFT JOIN sh_client_entries ce ON ce.handover_id = h.id
        WHERE h.handover_date = %s
        ORDER BY h.project_name, h.shift_num, ce.client_name
    """, (report_date,))
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_report_by_employee(name_filter, date_from=None, date_to=None):
    conn = get_db()
    cur = conn.cursor()
    wheres = []
    params = []
    if name_filter:
        wheres.append("(h.submitted_by_name ILIKE %s OR ce.filled_by_name ILIKE %s)")
        params += [f"%{name_filter}%", f"%{name_filter}%"]
    if date_from:
        wheres.append("h.handover_date >= %s"); params.append(date_from)
    if date_to:
        wheres.append("h.handover_date <= %s"); params.append(date_to)
    where_clause = ("WHERE " + " AND ".join(wheres)) if wheres else ""
    cur.execute(f"""
        SELECT h.handover_date, h.shift_num, h.project_name,
               h.submitted_by_name, h.status,
               ce.client_name, ce.entry_status, ce.filled_by_name,
               ce.engineer_notes, ce.tickets
        FROM sh_handovers h
        LEFT JOIN sh_client_entries ce ON ce.handover_id = h.id
        {where_clause}
        ORDER BY h.handover_date DESC, h.project_name, h.shift_num
    """, params)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_report_by_client(client_filter, date_from=None, date_to=None):
    conn = get_db()
    cur = conn.cursor()
    wheres = []
    params = []
    if client_filter:
        wheres.append("ce.client_name ILIKE %s"); params.append(f"%{client_filter}%")
    if date_from:
        wheres.append("h.handover_date >= %s"); params.append(date_from)
    if date_to:
        wheres.append("h.handover_date <= %s"); params.append(date_to)
    where_clause = ("WHERE " + " AND ".join(wheres)) if wheres else ""
    cur.execute(f"""
        SELECT h.handover_date, h.shift_num, h.project_name,
               ce.client_name, ce.entry_status, ce.engineer_worked,
               ce.issues, ce.engineer_notes, ce.tickets
        FROM sh_handovers h
        JOIN sh_client_entries ce ON ce.handover_id = h.id
        {where_clause}
        ORDER BY h.handover_date DESC, ce.client_name, h.shift_num
    """, params)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_get_history_for_excel(project_name=None, shift_num=None, status=None,
                              date_from=None, date_to=None):
    conn = get_db()
    cur = conn.cursor()
    wheres, params = [], []
    if project_name:
        wheres.append("h.project_name = %s"); params.append(project_name)
    if shift_num:
        wheres.append("h.shift_num = %s"); params.append(int(shift_num))
    if status:
        wheres.append("h.status = %s"); params.append(status)
    if date_from:
        wheres.append("h.handover_date >= %s"); params.append(date_from)
    if date_to:
        wheres.append("h.handover_date <= %s"); params.append(date_to)
    where_clause = ("WHERE " + " AND ".join(wheres)) if wheres else ""
    cur.execute(f"""
        SELECT h.handover_date, h.shift_num, h.project_name,
               h.submitted_by_name, h.status, h.lead_notes,
               h.engineer_acknowledged, h.manager_acknowledged,
               ce.client_name, ce.tickets, ce.entry_status,
               ce.engineer_worked, ce.issues, ce.engineer_notes,
               ce.next_shift_engineer
        FROM sh_handovers h
        LEFT JOIN sh_client_entries ce ON ce.handover_id = h.id
        {where_clause}
        ORDER BY h.handover_date DESC, h.project_name, h.shift_num, ce.client_name
    """, params)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


def sh_get_client_history(project_name, client_name, date_from=None, date_to=None):
    """Return all handover entries for a specific client+project, ordered newest first."""
    conn = get_db()
    cur = conn.cursor()
    wheres = ["h.project_name = %s", "LOWER(ce.client_name) = LOWER(%s)"]
    params = [project_name, client_name]
    if date_from:
        wheres.append("h.handover_date >= %s"); params.append(date_from)
    if date_to:
        wheres.append("h.handover_date <= %s"); params.append(date_to)
    where_clause = "WHERE " + " AND ".join(wheres)
    cur.execute(f"""
        SELECT
            h.id            AS handover_id,
            h.handover_date,
            h.shift_num,
            h.project_name,
            h.submitted_by_name,
            h.status,
            ce.client_name,
            ce.tickets,
            ce.entry_status,
            ce.engineer_worked,
            ce.issues,
            ce.engineer_notes,
            ce.manager_notes,
            ce.next_shift_engineer,
            ce.migration_report_sent,
            ce.drive_changes_alerts,
            ce.row_tint,
            ce.filled_by_name
        FROM sh_handovers h
        JOIN sh_client_entries ce ON ce.handover_id = h.id
        {where_clause}
        ORDER BY h.handover_date DESC, h.shift_num
    """, params)
    rows = _fetchall(cur)
    cur.close()
    conn.close()
    return rows


# ── Engineer activity for tracking page ─────────────────────────────────────

def sh_engineer_activity(target_date):
    conn = get_db()
    cur = conn.cursor()

    # Who filled entries today and how many
    cur.execute("""
        SELECT ce.filled_by_name AS engineer,
               COUNT(ce.id) AS entries_filled,
               MAX(ce.updated_at) AS last_activity,
               ARRAY_AGG(DISTINCT h.shift_num ORDER BY h.shift_num) AS shifts_covered
        FROM sh_client_entries ce
        JOIN sh_handovers h ON ce.handover_id = h.id
        WHERE h.handover_date = %s
          AND ce.filled_by_name IS NOT NULL
          AND ce.filled_by_name <> ''
        GROUP BY ce.filled_by_name
        ORDER BY entries_filled DESC
    """, (target_date,))
    filled = _fetchall(cur)

    # All system users
    cur.execute("""
        SELECT id, COALESCE(full_name, username) AS name, role
        FROM users
        WHERE COALESCE(is_active, TRUE) = TRUE
        ORDER BY name
    """)
    all_users = _fetchall(cur)

    filled_names = {r["engineer"] for r in filled}

    # Shift-by-shift breakdown
    cur.execute("""
        SELECT h.shift_num, h.project_name, h.status,
               COUNT(ce.id) AS total_entries,
               SUM(CASE WHEN ce.engineer_notes <> '' OR ce.tickets <> '' THEN 1 ELSE 0 END) AS filled_entries,
               ARRAY_AGG(DISTINCT ce.filled_by_name) FILTER (WHERE ce.filled_by_name <> '') AS engineers
        FROM sh_handovers h
        LEFT JOIN sh_client_entries ce ON ce.handover_id = h.id
        WHERE h.handover_date = %s
        GROUP BY h.shift_num, h.project_name, h.status
        ORDER BY h.shift_num, h.project_name
    """, (target_date,))
    shift_breakdown = _fetchall(cur)

    cur.close()
    conn.close()

    not_filled = [u for u in all_users if u["name"] not in filled_names]

    return {
        "filled":         filled,
        "not_filled":     not_filled,
        "all_users":      all_users,
        "shift_breakdown": shift_breakdown,
    }
