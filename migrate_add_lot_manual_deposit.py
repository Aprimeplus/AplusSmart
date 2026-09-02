import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: project_lots.manual_deposit_amount (deposit_method='spread' override)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                ALTER TABLE project_lots
                ADD COLUMN IF NOT EXISTS manual_deposit_amount NUMERIC
            """)
            print("- Added 'manual_deposit_amount' column to project_lots (NULL = ใช้ยอดคำนวณอัตโนมัติ)")
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
