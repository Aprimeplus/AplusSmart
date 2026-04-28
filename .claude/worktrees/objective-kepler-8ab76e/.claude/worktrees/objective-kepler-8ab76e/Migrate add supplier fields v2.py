"""
migrate_add_supplier_fields_v2.py
──────────────────────────────────
เพิ่ม 3 columns ใน table suppliers:
  - business_type     TEXT   ประเภทธุรกิจ (โรงงาน / ยี่ปั๊วใหญ่ / ร้านโลคอล)
  - standard_focus    TEXT   มาตรฐานสินค้า (เกรดราชการ / เกรดทั่วไป)
  - credit_term_label TEXT   เครดิต (สด / เครดิต 2D / 3D / 7D ฯลฯ)

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
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS business_type     TEXT DEFAULT ''",
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS standard_focus    TEXT DEFAULT ''",
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS credit_term_label TEXT DEFAULT 'สด'",
]

def run():
    conn = psycopg2.connect(**DB_CFG)
    try:
        with conn.cursor() as cur:
            for sql in MIGRATIONS:
                print(f"▶ {sql}")
                cur.execute(sql)
            conn.commit()
        print("\n✅ Migration สำเร็จ — เพิ่ม 3 columns (business_type, standard_focus, credit_term_label) แล้ว")
    except Exception as e:
        conn.rollback()
        print(f"❌ Error: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    run()