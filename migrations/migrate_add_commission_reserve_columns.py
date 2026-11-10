import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Adding Commission Reserve columns to commissions table...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS commission_now_amount NUMERIC")
            print("- Added 'commission_now_amount' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS commission_reserve_amount NUMERIC")
            print("- Added 'commission_reserve_amount' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS reserve_status TEXT DEFAULT NULL")
            print("- Added 'reserve_status' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS reserve_payout_id INTEGER")
            print("- Added 'reserve_payout_id' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS reserve_decided_at TIMESTAMP")
            print("- Added 'reserve_decided_at' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS reserve_decided_by TEXT")
            print("- Added 'reserve_decided_by' column")
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
