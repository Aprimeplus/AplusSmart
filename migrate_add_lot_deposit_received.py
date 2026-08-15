import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Add deposit_received_flag/date to project_lots...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                "ALTER TABLE project_lots ADD COLUMN IF NOT EXISTS deposit_received_flag BOOLEAN DEFAULT FALSE")
            print("- Added 'deposit_received_flag' column to project_lots")
            cursor.execute(
                "ALTER TABLE project_lots ADD COLUMN IF NOT EXISTS deposit_received_date DATE")
            print("- Added 'deposit_received_date' column to project_lots")
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
