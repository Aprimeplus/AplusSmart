"""
Migration: เพิ่ม defer_source_month และ defer_source_year
เพื่อเก็บเดือน/ปีต้นทางของ SO ที่ถูกเลื่อน ใช้แสดงในรายงาน Export
"""
import psycopg2
import os

DB_URL = os.environ.get("DATABASE_URL", "postgresql://postgres:password@localhost:5432/aplus_com_test")

def run():
    conn = psycopg2.connect(DB_URL)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                ALTER TABLE commissions
                    ADD COLUMN IF NOT EXISTS defer_source_month INT,
                    ADD COLUMN IF NOT EXISTS defer_source_year  INT;
            """)
        conn.commit()
        print("✅ เพิ่ม defer_source_month, defer_source_year เรียบร้อยแล้ว")
    except Exception as e:
        conn.rollback()
        print(f"❌ Error: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    run()
