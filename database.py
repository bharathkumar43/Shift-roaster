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
            working_days TEXT NOT NULL
        )
    """)
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
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL DEFAULT '',
            role TEXT NOT NULL DEFAULT 'user',
            created_at TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    # Seed default admin if no users exist
    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (%s, %s, %s, %s)",
            ("admin", generate_password_hash("admin123"), "Administrator", "admin")
        )
    conn.commit()
    cur.close()
    conn.close()


# ── Employee CRUD ────────────────────────────────────────

def add_employee(name, content_types, working_days):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO employees (name, content_types, working_days) VALUES (%s, %s, %s) RETURNING id",
            (name, json.dumps(content_types), json.dumps(working_days))
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


def update_employee(emp_id, content_types, working_days):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE employees SET content_types = %s, working_days = %s WHERE id = %s",
        (json.dumps(content_types), json.dumps(working_days), emp_id)
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
    return {
        "id": row["id"],
        "name": row["name"],
        "content_types": json.loads(ct) if isinstance(ct, str) else ct,
        "working_days": json.loads(wd) if isinstance(wd, str) else wd,
    }


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

def add_user(username, password_hash, full_name="", role="user"):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (%s, %s, %s, %s) RETURNING id",
            (username, password_hash, full_name, role)
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
