import sqlite3
import json
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "roster.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            content_types TEXT NOT NULL,
            working_days TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            product_type TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS rotation_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            shift_assigned INTEGER NOT NULL,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
            UNIQUE(employee_id, year, month)
        );

        CREATE TABLE IF NOT EXISTS saved_rosters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            generated_at TEXT NOT NULL,
            excel_blob BLOB NOT NULL,
            UNIQUE(year, month)
        );
    """)
    conn.commit()
    conn.close()


# ── Employee CRUD ────────────────────────────────────────

def add_employee(name, content_types, working_days):
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO employees (name, content_types, working_days) VALUES (?, ?, ?)",
            (name, json.dumps(content_types), json.dumps(working_days))
        )
        conn.commit()
        return conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    except sqlite3.IntegrityError:
        return None
    finally:
        conn.close()


def get_all_employees():
    conn = get_db()
    rows = conn.execute("SELECT * FROM employees ORDER BY id").fetchall()
    conn.close()
    return [_row_to_employee(r) for r in rows]


def get_employee_by_id(emp_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM employees WHERE id = ?", (emp_id,)).fetchone()
    conn.close()
    return _row_to_employee(row) if row else None


def remove_employee(emp_id):
    conn = get_db()
    conn.execute("DELETE FROM employees WHERE id = ?", (emp_id,))
    conn.commit()
    conn.close()


def clear_all_employees():
    conn = get_db()
    conn.execute("DELETE FROM employees")
    conn.commit()
    conn.close()


def _row_to_employee(row):
    return {
        "id": row["id"],
        "name": row["name"],
        "content_types": json.loads(row["content_types"]),
        "working_days": json.loads(row["working_days"]),
    }


# ── Project CRUD ─────────────────────────────────────────

def add_project(name, product_type, employee_id):
    conn = get_db()
    conn.execute(
        "INSERT INTO projects (name, product_type, employee_id) VALUES (?, ?, ?)",
        (name, product_type, employee_id)
    )
    conn.commit()
    conn.close()


def get_projects_for_employee(employee_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM projects WHERE employee_id = ? ORDER BY id", (employee_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_projects():
    conn = get_db()
    rows = conn.execute("""
        SELECT p.id, p.name, p.product_type, p.employee_id, e.name AS employee_name
        FROM projects p
        JOIN employees e ON p.employee_id = e.id
        ORDER BY p.id
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def remove_project(project_id):
    conn = get_db()
    conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
    conn.commit()
    conn.close()


# ── Rotation History ─────────────────────────────────────

def save_rotation(employee_id, year, month, shift_assigned):
    conn = get_db()
    conn.execute("""
        INSERT OR REPLACE INTO rotation_history (employee_id, year, month, shift_assigned)
        VALUES (?, ?, ?, ?)
    """, (employee_id, year, month, shift_assigned))
    conn.commit()
    conn.close()


def save_all_rotations(assignments, employees, year, month):
    conn = get_db()
    name_to_id = {e["name"]: e["id"] for e in employees}
    for emp_name, shift in assignments.items():
        emp_id = name_to_id.get(emp_name)
        if emp_id:
            conn.execute("""
                INSERT OR REPLACE INTO rotation_history (employee_id, year, month, shift_assigned)
                VALUES (?, ?, ?, ?)
            """, (emp_id, year, month, shift))
    conn.commit()
    conn.close()


def get_rotation_history():
    conn = get_db()
    rows = conn.execute("""
        SELECT rh.employee_id, e.name AS employee_name, rh.year, rh.month, rh.shift_assigned
        FROM rotation_history rh
        JOIN employees e ON rh.employee_id = e.id
        ORDER BY rh.year DESC, rh.month DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_night_shift_counts():
    conn = get_db()
    rows = conn.execute("""
        SELECT employee_id, COUNT(*) as cnt
        FROM rotation_history
        WHERE shift_assigned = 3
        GROUP BY employee_id
    """).fetchall()
    conn.close()
    return {r["employee_id"]: r["cnt"] for r in rows}


def get_last_night_shift_month(employee_id):
    conn = get_db()
    row = conn.execute("""
        SELECT year, month FROM rotation_history
        WHERE employee_id = ? AND shift_assigned = 3
        ORDER BY year DESC, month DESC
        LIMIT 1
    """, (employee_id,)).fetchone()
    conn.close()
    return (row["year"], row["month"]) if row else None


def get_rotation_for_employee(employee_id, year, month):
    conn = get_db()
    row = conn.execute("""
        SELECT shift_assigned FROM rotation_history
        WHERE employee_id = ? AND year = ? AND month = ?
    """, (employee_id, year, month)).fetchone()
    conn.close()
    return row["shift_assigned"] if row else None


# ── Saved Rosters ────────────────────────────────────────

def save_roster_excel(year, month, excel_bytes):
    conn = get_db()
    conn.execute("""
        INSERT OR REPLACE INTO saved_rosters (year, month, generated_at, excel_blob)
        VALUES (?, ?, datetime('now'), ?)
    """, (year, month, excel_bytes))
    conn.commit()
    conn.close()


def get_saved_roster(year, month):
    conn = get_db()
    row = conn.execute(
        "SELECT excel_blob FROM saved_rosters WHERE year = ? AND month = ?",
        (year, month)
    ).fetchone()
    conn.close()
    return bytes(row["excel_blob"]) if row else None


def list_saved_rosters():
    conn = get_db()
    rows = conn.execute("""
        SELECT year, month, generated_at
        FROM saved_rosters
        ORDER BY year DESC, month DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── Search ───────────────────────────────────────────────

def search_employees(query):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM employees WHERE name LIKE ? ORDER BY name",
        (f"%{query}%",)
    ).fetchall()
    conn.close()
    return [_row_to_employee(r) for r in rows]


def search_projects(query):
    conn = get_db()
    rows = conn.execute("""
        SELECT p.id, p.name, p.product_type, p.employee_id, e.name AS employee_name
        FROM projects p
        JOIN employees e ON p.employee_id = e.id
        WHERE p.name LIKE ?
        ORDER BY p.name
    """, (f"%{query}%",)).fetchall()
    conn.close()
    return [dict(r) for r in rows]
