import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Adding defer_type and defer_decision columns to commissions table...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS defer_type TEXT DEFAULT NULL")
            print("- Added 'defer_type' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS defer_decision TEXT DEFAULT NULL")
            print("- Added 'defer_decision' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS defer_decision_reason TEXT DEFAULT NULL")
            print("- Added 'defer_decision_reason' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS defer_requested_by TEXT DEFAULT NULL")
            print("- Added 'defer_requested_by' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS defer_approved_by TEXT DEFAULT NULL")
            print("- Added 'defer_approved_by' column")
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
