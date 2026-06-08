"""
migrate_sh_data.py
==================
Step 1 — Auto-loads the Prisma dump into roster_db (finds psql.exe automatically).
Step 2 — Migrates data from Prisma tables into Flask sh_* tables in the same DB.

Run once:
    python migrate_sh_data.py
"""

import os
import sys
import subprocess
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

# ── Config ────────────────────────────────────────────────────────────────────

DUMP_PATH = r"C:\Users\BharathTummaganti\OneDrive - CloudFuze, Inc\Documents\Shift_handover\shift_handover_export.sql"

DB_CONFIG = {
    "host":     os.getenv("PG_HOST",     "localhost"),
    "port":     os.getenv("PG_PORT",     "5432"),
    "user":     os.getenv("PG_USER",     "postgres"),
    "password": os.getenv("PG_PASSWORD", "postgres"),
    "dbname":   os.getenv("PG_DATABASE", "roster_db"),
}

# Common PostgreSQL bin locations on Windows
PSQL_CANDIDATES = [
    r"C:\Program Files\PostgreSQL\18\bin\psql.exe",
    r"C:\Program Files\PostgreSQL\17\bin\psql.exe",
    r"C:\Program Files\PostgreSQL\16\bin\psql.exe",
    r"C:\Program Files\PostgreSQL\15\bin\psql.exe",
    r"C:\Program Files\PostgreSQL\14\bin\psql.exe",
    "psql",  # already in PATH
]

# ── Data mappings ─────────────────────────────────────────────────────────────

ENTRY_STATUS_MAP = {
    "COMPLETE":    "Completed",
    "IN_PROGRESS": "Active",
    "PENDING":     "Pending",
    "DELTA":       "Pending",
    "NA":          "N/A",
    "GOOD":        "Good",
    "BAD":         "Bad",
}

ROLE_MAP = {
    "ADMIN":    "admin",
    "LEAD":     "admin",
    "ENGINEER": "engineer",
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def get_conn():
    return psycopg2.connect(**DB_CONFIG, cursor_factory=psycopg2.extras.RealDictCursor)

def log(msg):
    print(f"  {msg}")

def find_psql():
    for p in PSQL_CANDIDATES:
        if os.path.isabs(p):
            if os.path.exists(p):
                return p
        else:
            return p   # rely on PATH
    return None

def prisma_tables_exist(conn):
    cur = conn.cursor()
    try:
        cur.execute('SELECT 1 FROM public."ShiftHandover" LIMIT 1')
        return True
    except Exception:
        conn.rollback()
        return False
    finally:
        cur.close()

# ── Step 1 — Load dump ────────────────────────────────────────────────────────

def load_dump():
    print("\n[0] Loading dump into roster_db...")

    if not os.path.exists(DUMP_PATH):
        print(f"  ERROR: Dump file not found: {DUMP_PATH}")
        sys.exit(1)

    psql = find_psql()
    if not psql:
        print("  ERROR: psql.exe not found.")
        sys.exit(1)

    print(f"  Using psql: {psql}")
    print(f"  Dump file: {DUMP_PATH}")

    env = os.environ.copy()
    env["PGPASSWORD"] = DB_CONFIG["password"]

    result = subprocess.run(
        [
            psql,
            "-h", DB_CONFIG["host"],
            "-p", str(DB_CONFIG["port"]),
            "-U", DB_CONFIG["user"],
            "-d", DB_CONFIG["dbname"],
            "-f", DUMP_PATH,
            "--set=ON_ERROR_STOP=0",   # continue on non-fatal errors
        ],
        env=env,
        capture_output=True,
        text=True,
    )

    # Show only real ERRORs (ignore "already exists" from re-runs)
    errors = [
        ln for ln in result.stderr.splitlines()
        if "ERROR" in ln and "already exists" not in ln and "duplicate" not in ln.lower()
    ]
    if errors:
        print("  Warnings during import (non-fatal):")
        for e in errors[:15]:
            print(f"    {e}")

    print(f"  Dump loaded (psql exit code: {result.returncode}).")

# ── Step 2 — Migrate ──────────────────────────────────────────────────────────

def migrate_projects(conn):
    print("\n[1] Projects...")
    cur = conn.cursor()
    cur.execute('SELECT name FROM public."Project" ORDER BY name')
    projects = cur.fetchall()
    for i, p in enumerate(projects, 1):
        cur.execute(
            "INSERT INTO sh_projects (name, sort_order) VALUES (%s,%s) ON CONFLICT (name) DO NOTHING",
            (p["name"], i),
        )
        log(p["name"])
    conn.commit()
    log(f"Done — {len(projects)} projects.")


def migrate_clients(conn):
    print("\n[2] Clients...")
    cur = conn.cursor()
    cur.execute("""
        SELECT c.name, c.active, p.name AS project_name
        FROM public."Client" c
        JOIN public."Project" p ON p.id = c."projectId"
        ORDER BY p.name, c."sortOrder"
    """)
    clients = cur.fetchall()
    for c in clients:
        cur.execute(
            """
            INSERT INTO sh_clients (project_name, client_name, is_active)
            VALUES (%s,%s,%s)
            ON CONFLICT (project_name, client_name) DO UPDATE SET is_active = EXCLUDED.is_active
            """,
            (c["project_name"], c["name"], c["active"]),
        )
    conn.commit()
    log(f"Done — {len(clients)} clients.")


def migrate_users(conn):
    print("\n[3] Users...")
    cur = conn.cursor()
    cur.execute('SELECT id, name, email, role, active FROM public."User" ORDER BY "createdAt"')
    users = cur.fetchall()

    user_name_map = {}
    inserted = skipped = 0
    for u in users:
        user_name_map[u["id"]] = u["name"] or ""
        email = (u["email"] or "").lower().strip()
        if not email:
            skipped += 1
            continue
        cur.execute("SELECT id FROM users WHERE username = %s OR email = %s", (email, email))
        existing = cur.fetchone()
        if existing:
            log(f"SKIP (exists): {u['name']} <{email}>")
            skipped += 1
            continue
        role = ROLE_MAP.get(u["role"], "engineer")
        cur.execute(
            "INSERT INTO users (username, full_name, email, password_hash, role, is_active) "
            "VALUES (%s,%s,%s,%s,%s,%s)",
            (email, u["name"] or "", email, "", role, u["active"]),
        )
        log(f"Inserted: {u['name']} <{email}> [{role}]")
        inserted += 1
    conn.commit()
    log(f"Done — {inserted} inserted, {skipped} skipped.")
    return user_name_map


def migrate_handovers(conn, user_name_map):
    print("\n[4] Shift handovers...")
    cur = conn.cursor()
    cur.execute("""
        SELECT h.id, h.date, h."shiftNumber", h.status,
               h."leadId", h."submittedById", h."leadNotes",
               h."engineerAcknowledged", h."engineerAcknowledgedAt", h."engineerAcknowledgedById",
               h."managerAcknowledged",  h."managerAcknowledgedAt",  h."managerAcknowledgedById",
               p.name AS project_name
        FROM public."ShiftHandover" h
        JOIN public."Project" p ON p.id = h."projectId"
        ORDER BY h.date, h."shiftNumber"
    """)
    handovers = cur.fetchall()

    handover_id_map = {}
    inserted = skipped = 0
    for h in handovers:
        status        = h["status"].lower()
        lead_name     = user_name_map.get(h["leadId"], "")
        sub_name      = user_name_map.get(h["submittedById"], lead_name)
        by_name       = sub_name if status == "submitted" else lead_name

        cur.execute(
            "SELECT id FROM sh_handovers WHERE handover_date=%s AND shift_num=%s AND project_name=%s",
            (h["date"], h["shiftNumber"], h["project_name"]),
        )
        existing = cur.fetchone()
        if existing:
            handover_id_map[h["id"]] = existing["id"]
            log(f"SKIP (exists): {h['project_name']} Shift {h['shiftNumber']} {h['date']}")
            skipped += 1
            continue

        cur.execute(
            """
            INSERT INTO sh_handovers
                (handover_date, shift_num, project_name, submitted_by_name, status, lead_notes,
                 engineer_acknowledged, engineer_acknowledged_by, engineer_acknowledged_at,
                 manager_acknowledged,  manager_acknowledged_by,  manager_acknowledged_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id
            """,
            (
                h["date"], h["shiftNumber"], h["project_name"], by_name,
                status, h["leadNotes"] or "",
                h["engineerAcknowledged"],
                user_name_map.get(h["engineerAcknowledgedById"], ""),
                h["engineerAcknowledgedAt"],
                h["managerAcknowledged"],
                user_name_map.get(h["managerAcknowledgedById"], ""),
                h["managerAcknowledgedAt"],
            ),
        )
        new_id = cur.fetchone()["id"]
        handover_id_map[h["id"]] = new_id
        log(f"Inserted: {h['project_name']} Shift {h['shiftNumber']} {h['date']} [{status}] by '{by_name}'")
        inserted += 1

    conn.commit()
    log(f"Done — {inserted} inserted, {skipped} skipped.")
    return handover_id_map


def migrate_entries(conn, handover_id_map, user_name_map):
    print("\n[5] Client entries...")
    cur = conn.cursor()
    cur.execute("""
        SELECT ce."shiftHandoverId", ce.tickets, ce.status,
               ce."engineerWorked", ce.issues, ce.updates, ce."handoverNotes",
               ce."managerNotes", ce."filledById", ce."rowTint",
               ce."migrationReportSent", ce."driveChangesAlerts",
               cl.name AS client_name
        FROM public."ClientEntry" ce
        JOIN public."Client" cl ON cl.id = ce."clientId"
        ORDER BY ce."createdAt"
    """)
    entries = cur.fetchall()

    inserted = missing = 0
    for e in entries:
        new_hid = handover_id_map.get(e["shiftHandoverId"])
        if not new_hid:
            missing += 1
            continue

        status_key   = str(e["status"]) if e["status"] else "NA"
        entry_status = ENTRY_STATUS_MAP.get(status_key, "N/A")
        eng_notes    = "\n".join(filter(None, [e["updates"] or "", e["handoverNotes"] or ""])).strip()

        cur.execute(
            """
            INSERT INTO sh_client_entries
                (handover_id, client_name, tickets, entry_status,
                 engineer_worked, issues, engineer_notes, manager_notes,
                 next_shift_engineer, migration_report_sent, drive_changes_alerts,
                 row_tint, filled_by_name)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT (handover_id, client_name) DO UPDATE SET
                tickets               = EXCLUDED.tickets,
                entry_status          = EXCLUDED.entry_status,
                engineer_worked       = EXCLUDED.engineer_worked,
                issues                = EXCLUDED.issues,
                engineer_notes        = EXCLUDED.engineer_notes,
                manager_notes         = EXCLUDED.manager_notes,
                migration_report_sent = EXCLUDED.migration_report_sent,
                drive_changes_alerts  = EXCLUDED.drive_changes_alerts,
                row_tint              = EXCLUDED.row_tint,
                filled_by_name        = EXCLUDED.filled_by_name,
                updated_at            = NOW()
            """,
            (
                new_hid, e["client_name"] or "",
                e["tickets"] or "", entry_status,
                e["engineerWorked"] or "", e["issues"] or "",
                eng_notes, e["managerNotes"] or "", "",
                bool(e["migrationReportSent"]), bool(e["driveChangesAlerts"]),
                e["rowTint"],
                user_name_map.get(e["filledById"], ""),
            ),
        )
        inserted += 1

    conn.commit()
    log(f"Done — {inserted} entries migrated, {missing} skipped (no matching handover).")


def migrate_daily_notes(conn):
    print("\n[6] Daily notes...")
    cur = conn.cursor()
    cur.execute('SELECT date, "dutyManager", week, "keyIssues", "actionsForTomorrow" FROM public."DailyDashboard"')
    notes = cur.fetchall()
    if not notes:
        log("Empty — skipping.")
        return
    for n in notes:
        cur.execute(
            """
            INSERT INTO sh_daily_notes (note_date, duty_manager, week_label, key_issues, actions_for_tomorrow)
            VALUES (%s,%s,%s,%s,%s)
            ON CONFLICT (note_date) DO UPDATE SET
                duty_manager=EXCLUDED.duty_manager, week_label=EXCLUDED.week_label,
                key_issues=EXCLUDED.key_issues, actions_for_tomorrow=EXCLUDED.actions_for_tomorrow
            """,
            (n["date"], n["dutyManager"] or "", n["week"] or "",
             n["keyIssues"] or "", n["actionsForTomorrow"] or ""),
        )
    conn.commit()
    log(f"Done — {len(notes)} records.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("Shift Handover Migration")
    print(f"  DB: {DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['dbname']}")
    print("=" * 60)

    # Step 1 — load dump if Prisma tables aren't there yet
    try:
        conn = get_conn()
    except Exception as e:
        print(f"\nERROR connecting to DB: {e}")
        sys.exit(1)

    if not prisma_tables_exist(conn):
        load_dump()
        conn.close()
        # Re-connect after dump load
        try:
            conn = get_conn()
        except Exception as e:
            print(f"\nERROR reconnecting: {e}")
            sys.exit(1)
        if not prisma_tables_exist(conn):
            print("\nERROR: Prisma tables still not found after dump load.")
            sys.exit(1)
    else:
        print('\nPrisma tables already present — skipping dump load.')

    # Step 2 — migrate
    migrate_projects(conn)
    migrate_clients(conn)
    user_name_map   = migrate_users(conn)
    handover_id_map = migrate_handovers(conn, user_name_map)
    migrate_entries(conn, handover_id_map, user_name_map)
    migrate_daily_notes(conn)

    conn.close()
    print("\n" + "=" * 60)
    print("Migration complete. Prisma tables are still in the DB.")
    print("To clean them up after verifying, run:")
    print('  python migrate_sh_data.py --cleanup')
    print("=" * 60)


def cleanup():
    """Drop the Prisma tables and types from roster_db."""
    print("\nCleaning up Prisma tables...")
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        DROP TABLE IF EXISTS
            public."ClientEntry", public."ShiftHandover", public."Client",
            public."Project", public."User", public."DailyDashboard", public."MOM",
            public."MigrationProject", public."MigrationItem", public."MigrationIssue",
            public."MigrationTask", public."MigrationTypeOption",
            public."MigrationProjectTicket", public."BatchRun",
            public."ProjectComment", public._prisma_migrations CASCADE
    """)
    cur.execute("""
        DROP TYPE IF EXISTS
            public."EntryStatus", public."HandoverStatus", public."IssueTicketStatus",
            public."MigrationCombination", public."MigrationItemStatus", public."MigrationPhase",
            public."MigrationProjectStatus", public."ProductType", public."Role",
            public."RowTint", public."TaskAssignedTo", public."TaskStatus" CASCADE
    """)
    conn.commit()
    conn.close()
    print("Done — Prisma tables removed.")


if __name__ == "__main__":
    if "--cleanup" in sys.argv:
        cleanup()
    else:
        main()
