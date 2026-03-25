"""
One-time migration script: SQLite (roster.db) -> PostgreSQL

Usage:
  1. Fill in your PostgreSQL credentials in .env
  2. Create the PostgreSQL database (e.g., CREATE DATABASE roster_db)
  3. Run: python migrate_data.py
"""

import sqlite3
import json
import os

import psycopg2
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

SQLITE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "roster.db")

PG_CONFIG = {
    "host": os.getenv("PG_HOST", "localhost"),
    "port": os.getenv("PG_PORT", "5432"),
    "user": os.getenv("PG_USER", "postgres"),
    "password": os.getenv("PG_PASSWORD", "postgres"),
    "dbname": os.getenv("PG_DATABASE", "roster_db"),
}


def migrate():
    if not os.path.exists(SQLITE_PATH):
        print(f"SQLite database not found at: {SQLITE_PATH}")
        print("Nothing to migrate.")
        return

    sqlite_conn = sqlite3.connect(SQLITE_PATH)
    sqlite_conn.row_factory = sqlite3.Row

    pg_conn = psycopg2.connect(**PG_CONFIG)
    pg_cur = pg_conn.cursor()

    print("Connected to both databases.\n")

    # Create tables in PostgreSQL
    pg_cur.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            content_types TEXT NOT NULL,
            working_days TEXT NOT NULL
        )
    """)
    pg_cur.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            product_type TEXT NOT NULL,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE
        )
    """)
    pg_cur.execute("""
        CREATE TABLE IF NOT EXISTS rotation_history (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            shift_assigned INTEGER NOT NULL,
            UNIQUE(employee_id, year, month)
        )
    """)
    pg_cur.execute("""
        CREATE TABLE IF NOT EXISTS saved_rosters (
            id SERIAL PRIMARY KEY,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            generated_at TIMESTAMP NOT NULL DEFAULT NOW(),
            excel_blob BYTEA NOT NULL,
            UNIQUE(year, month)
        )
    """)
    pg_conn.commit()
    print("PostgreSQL tables created.\n")

    # Migrate employees
    rows = sqlite_conn.execute("SELECT * FROM employees ORDER BY id").fetchall()
    emp_id_map = {}
    emp_count = 0
    for row in rows:
        old_id = row["id"]
        try:
            pg_cur.execute(
                "INSERT INTO employees (name, content_types, working_days) VALUES (%s, %s, %s) RETURNING id",
                (row["name"], row["content_types"], row["working_days"])
            )
            new_id = pg_cur.fetchone()[0]
            emp_id_map[old_id] = new_id
            emp_count += 1
        except psycopg2.IntegrityError:
            pg_conn.rollback()
            pg_cur.execute("SELECT id FROM employees WHERE name = %s", (row["name"],))
            existing = pg_cur.fetchone()
            if existing:
                emp_id_map[old_id] = existing[0]
            print(f"  Skipped employee '{row['name']}' (already exists)")
    pg_conn.commit()
    print(f"Employees migrated: {emp_count}")

    # Migrate projects
    rows = sqlite_conn.execute("SELECT * FROM projects ORDER BY id").fetchall()
    proj_count = 0
    for row in rows:
        new_emp_id = emp_id_map.get(row["employee_id"])
        if new_emp_id is None:
            print(f"  Skipped project '{row['name']}' (employee_id {row['employee_id']} not found)")
            continue
        try:
            pg_cur.execute(
                "INSERT INTO projects (name, product_type, employee_id) VALUES (%s, %s, %s)",
                (row["name"], row["product_type"], new_emp_id)
            )
            proj_count += 1
        except psycopg2.IntegrityError:
            pg_conn.rollback()
            print(f"  Skipped project '{row['name']}' (duplicate)")
    pg_conn.commit()
    print(f"Projects migrated: {proj_count}")

    # Migrate rotation history
    rows = sqlite_conn.execute("SELECT * FROM rotation_history ORDER BY id").fetchall()
    rot_count = 0
    for row in rows:
        new_emp_id = emp_id_map.get(row["employee_id"])
        if new_emp_id is None:
            continue
        try:
            pg_cur.execute("""
                INSERT INTO rotation_history (employee_id, year, month, shift_assigned)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (employee_id, year, month) DO NOTHING
            """, (new_emp_id, row["year"], row["month"], row["shift_assigned"]))
            rot_count += 1
        except psycopg2.IntegrityError:
            pg_conn.rollback()
    pg_conn.commit()
    print(f"Rotation history migrated: {rot_count}")

    # Migrate saved rosters
    rows = sqlite_conn.execute("SELECT * FROM saved_rosters ORDER BY id").fetchall()
    roster_count = 0
    for row in rows:
        try:
            pg_cur.execute("""
                INSERT INTO saved_rosters (year, month, generated_at, excel_blob)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (year, month) DO NOTHING
            """, (row["year"], row["month"], row["generated_at"],
                  psycopg2.Binary(row["excel_blob"])))
            roster_count += 1
        except psycopg2.IntegrityError:
            pg_conn.rollback()
    pg_conn.commit()
    print(f"Saved rosters migrated: {roster_count}")

    # Reset sequences to match the highest IDs
    for table in ["employees", "projects", "rotation_history", "saved_rosters"]:
        pg_cur.execute(f"SELECT MAX(id) FROM {table}")
        max_id = pg_cur.fetchone()[0]
        if max_id:
            pg_cur.execute(f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), %s)", (max_id,))
    pg_conn.commit()

    pg_cur.close()
    pg_conn.close()
    sqlite_conn.close()

    print(f"\nMigration complete!")
    print(f"  Employees: {emp_count}")
    print(f"  Projects: {proj_count}")
    print(f"  Rotation history: {rot_count}")
    print(f"  Saved rosters: {roster_count}")
    print(f"\nYour SQLite file (roster.db) has NOT been deleted.")


if __name__ == "__main__":
    migrate()
