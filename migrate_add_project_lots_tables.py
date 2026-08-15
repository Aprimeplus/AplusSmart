import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Multi-Lot Project tables (projects, project_lots)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS projects (
                    id SERIAL PRIMARY KEY,
                    project_code TEXT UNIQUE NOT NULL,
                    project_name TEXT NOT NULL,
                    customer_name TEXT,
                    total_project_value NUMERIC DEFAULT 0,
                    status TEXT DEFAULT 'Open',
                    created_by TEXT,
                    created_at TIMESTAMP DEFAULT NOW(),
                    closed_at TIMESTAMP,
                    closed_by TEXT
                )
            """)
            print("- Created 'projects' table")

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS project_lots (
                    id SERIAL PRIMARY KEY,
                    project_id INTEGER REFERENCES projects(id),
                    lot_number INTEGER,
                    lot_name TEXT,
                    so_number TEXT UNIQUE,
                    lot_value NUMERIC DEFAULT 0,
                    delivered_flag BOOLEAN DEFAULT FALSE,
                    delivery_note_signed_flag BOOLEAN DEFAULT FALSE,
                    delivery_date DATE,
                    invoice_recorded_flag BOOLEAN DEFAULT FALSE,
                    payment_collected_flag BOOLEAN DEFAULT FALSE,
                    payment_collected_date DATE,
                    gp_estimate_pct NUMERIC,
                    kpi_qualified_flag BOOLEAN DEFAULT FALSE,
                    status TEXT DEFAULT 'Draft',
                    lock_reason TEXT,
                    lock_unlocked_by TEXT,
                    lock_unlocked_at TIMESTAMP,
                    created_by TEXT,
                    created_at TIMESTAMP DEFAULT NOW()
                )
            """)
            print("- Created 'project_lots' table")

            cursor.execute(
                "CREATE INDEX IF NOT EXISTS idx_project_lots_project_id ON project_lots(project_id)")
            print("- Added index on project_lots(project_id)")
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
