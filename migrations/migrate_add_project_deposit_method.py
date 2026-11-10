import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: projects.deposit_method (spread / last_lot)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                ALTER TABLE projects
                ADD COLUMN IF NOT EXISTS deposit_method TEXT NOT NULL DEFAULT 'spread'
            """)
            print("- Added 'deposit_method' column to projects (default 'spread')")
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
