"""
One-time script: import all clients from the old Prisma shift_handover DB
into roster_db's sh_clients table.

Run on the server:
    python migrate_clients.py

Requires the OLD_PG_* env vars pointing at the old Prisma database.
The roster_db connection comes from the normal .env (PG_* vars).
"""
import os
import psycopg2
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

# ── Old Prisma DB connection (shift_handover) ─────────────────────────────────
OLD_CONFIG = {
    "host":     os.getenv("OLD_PG_HOST", "localhost"),
    "port":     os.getenv("OLD_PG_PORT", "5433"),          # adjust if different
    "user":     os.getenv("OLD_PG_USER", "postgres"),
    "password": os.getenv("OLD_PG_PASSWORD", "postgres"),
    "dbname":   os.getenv("OLD_PG_DATABASE", "shift_handover"),
}

# ── New roster_db connection ───────────────────────────────────────────────────
NEW_CONFIG = {
    "host":     os.getenv("PG_HOST", "localhost"),
    "port":     os.getenv("PG_PORT", "5432"),
    "user":     os.getenv("PG_USER", "postgres"),
    "password": os.getenv("PG_PASSWORD", "postgres"),
    "dbname":   os.getenv("PG_DATABASE", "roster_db"),
}

PROJECT_TYPE_MAP = {
    "CONTENT": "Content",
    "EMAIL":   "Email",
    "MESSAGE": "Messaging",
}

def main():
    print("Connecting to old Prisma DB ...")
    old_conn = psycopg2.connect(**OLD_CONFIG)
    old_cur  = old_conn.cursor()

    old_cur.execute("""
        SELECT c.name, c."productType", p.name AS project
        FROM "Client" c
        JOIN "Project" p ON p.id = c."projectId"
        WHERE c."productType" IS NOT NULL
        ORDER BY p.name, c.name
    """)
    old_clients = old_cur.fetchall()
    old_cur.close()
    old_conn.close()
    print(f"Found {len(old_clients)} clients in old DB.")

    print("Connecting to roster_db ...")
    new_conn = psycopg2.connect(**NEW_CONFIG)
    new_cur  = new_conn.cursor()

    inserted = 0
    skipped  = 0
    unknown  = 0

    for (client_name, product_type, project_name) in old_clients:
        mapped_project = PROJECT_TYPE_MAP.get(product_type)
        if not mapped_project:
            print(f"  SKIP (unknown productType={product_type}): {client_name}")
            unknown += 1
            continue

        new_cur.execute("""
            INSERT INTO sh_clients (project_name, client_name)
            VALUES (%s, %s)
            ON CONFLICT (project_name, client_name) DO NOTHING
        """, (mapped_project, client_name.strip()))

        if new_cur.rowcount:
            inserted += 1
        else:
            skipped += 1

    new_conn.commit()
    new_cur.close()
    new_conn.close()

    print(f"\nDone.")
    print(f"  Inserted : {inserted}")
    print(f"  Already existed (skipped) : {skipped}")
    print(f"  Unknown type (skipped)    : {unknown}")


if __name__ == "__main__":
    main()
