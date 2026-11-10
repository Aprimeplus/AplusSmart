import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: project_lots.product_type (เก็บไว้เป็นข้อมูลอ้างอิง ไม่ใช้คำนวณ)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                ALTER TABLE project_lots
                ADD COLUMN IF NOT EXISTS product_type TEXT
            """)
            print("- Added 'product_type' column to project_lots")
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
