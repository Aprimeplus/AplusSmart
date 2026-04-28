"""
migrate_add_zoning_columns.py
─────────────────────────────
เพิ่ม 3 columns ใน table suppliers สำหรับ Zoning Phase 1:
  - dispatch_zone     TEXT   โซนที่ตั้ง / จุดส่งสินค้า
  - service_area      TEXT   ขอบเขตการส่ง (National / Regional / Local)
  - logistics_assets  TEXT   ประเภทรถที่มี (comma-separated เช่น "6 ล้อ,10 ล้อ")

รันครั้งเดียว ปลอดภัย: ใช้ ADD COLUMN IF NOT EXISTS
"""

import psycopg2

DB_CFG = dict(
    host="Server-Aprime",
    dbname="aplus_com_test",
    user="app_user",
    password="cailfornia123",
)

MIGRATIONS = [
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS dispatch_zone    TEXT DEFAULT ''",
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS service_area     TEXT DEFAULT 'National'",
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS logistics_assets TEXT DEFAULT ''",
]

def run():
    conn = psycopg2.connect(**DB_CFG)
    try:
        with conn.cursor() as cur:
            for sql in MIGRATIONS:
                print(f"▶ {sql}")
                cur.execute(sql)
            conn.commit()
        print("\n✅ Migration สำเร็จ — เพิ่ม 3 columns (dispatch_zone, service_area, logistics_assets) แล้ว")
    except Exception as e:
        conn.rollback()
        print(f"❌ Error: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    run()