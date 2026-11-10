import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: projects.planned_lot_count...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                ALTER TABLE projects
                ADD COLUMN IF NOT EXISTS planned_lot_count INTEGER
            """)
            print("- Added 'planned_lot_count' column to projects "
                  "(NULL = โครงการเก่าที่ยังไม่ได้ระบุ ใช้เพดานตามมูลค่าแทน)")
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
