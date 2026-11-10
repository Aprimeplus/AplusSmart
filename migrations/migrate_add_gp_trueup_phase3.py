import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Phase 3 GP True-Up (projects columns + reserve_release_queue table)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS final_gp_pct NUMERIC")
            print("- Added 'final_gp_pct' column to projects")
            cursor.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS final_sales_amount NUMERIC")
            print("- Added 'final_sales_amount' column to projects")
            cursor.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS final_cost_amount NUMERIC")
            print("- Added 'final_cost_amount' column to projects")
            cursor.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS needs_director_approval BOOLEAN DEFAULT FALSE")
            print("- Added 'needs_director_approval' column to projects")
            cursor.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS close_note TEXT")
            print("- Added 'close_note' column to projects")

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS reserve_release_queue (
                    id SERIAL PRIMARY KEY,
                    commission_id INTEGER REFERENCES commissions(id),
                    project_id INTEGER REFERENCES projects(id),
                    sale_key TEXT,
                    so_number TEXT,
                    reserve_amount NUMERIC,
                    release_ratio NUMERIC,
                    release_amount NUMERIC,
                    forfeited_amount NUMERIC,
                    project_gp_pct NUMERIC,
                    status TEXT DEFAULT 'Pending',
                    applied_payout_id INTEGER,
                    decided_by TEXT,
                    decided_at TIMESTAMP,
                    created_at TIMESTAMP DEFAULT NOW()
                )
            """)
            print("- Created 'reserve_release_queue' table")
            cursor.execute(
                "CREATE INDEX IF NOT EXISTS idx_reserve_release_queue_sale_key "
                "ON reserve_release_queue(sale_key, status)")
            print("- Added index on reserve_release_queue(sale_key, status)")
        conn.commit()
        print("Migration completed successfully!")
    except Exception as e:
        conn.rollback()
        print(f"Migration failed: {e}")
    finally:
        app.release_connection(conn)
        app.destroy()

if __name__ == "__main__":
    migrate_db()
