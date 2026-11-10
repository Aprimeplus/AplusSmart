import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Adding missing columns (subject, payment amounts) to commissions table...")
    
    app = AppContainer()
    conn = app.get_connection()
    
    try:
        with conn.cursor() as cursor:
            # 1. เพิ่มคอลัมน์ 'subject'
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS subject TEXT DEFAULT '-'")
            print("- Added 'subject' column")

            # 2. เพิ่มคอลัมน์ 'payment1_amount'
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS payment1_amount REAL DEFAULT 0.0")
            print("- Added 'payment1_amount' column")

            # 3. เพิ่มคอลัมน์ 'payment2_amount'
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS payment2_amount REAL DEFAULT 0.0")
            print("- Added 'payment2_amount' column")
            
            # (Optional) ถ้าต้องการเพิ่ม payment1_date/method ด้วย (ถ้ายังไม่มี)
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS payment1_date TEXT")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS payment1_method TEXT")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS payment2_date TEXT")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS payment2_method TEXT")

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