"""
Migration: เพิ่มคอลัมน์ wh_zone และ wh_coordinates ใน suppliers table
รัน: python migrate_add_wh_location.py
"""
import psycopg2

DB_PARAMS = {
    "host":     "Server-APrime",
    "dbname":   "aplus_com_test",
    "user":     "app_user",
    "password": "cailfornia123",
    "port":     5432,
}

try:
    conn = psycopg2.connect(**DB_PARAMS)
    with conn.cursor() as cur:
        cur.execute("""
            ALTER TABLE suppliers
            ADD COLUMN IF NOT EXISTS wh_zone        TEXT DEFAULT '',
            ADD COLUMN IF NOT EXISTS wh_coordinates TEXT DEFAULT '';
        """)
    conn.commit()
    print("✅ เพิ่มคอลัมน์ wh_zone และ wh_coordinates เรียบร้อยแล้ว")
except Exception as e:
    print(f"❌ Error: {e}")
finally:
    conn.close()
