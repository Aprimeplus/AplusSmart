import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Add deposit_pct column to projects...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS deposit_pct NUMERIC")
            print("- Added 'deposit_pct' column to projects")
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
