"""
migrate_super_supplier_scoring.py
──────────────────────────────────
เพิ่ม fields ที่จำเป็นสำหรับระบบ Super Supplier Scoring (5 เกณฑ์ × 20%):
  suppliers table:
    - service_score  INT DEFAULT 0   (การบริการ — กรอกเอง)
    - quality_score  INT DEFAULT 0   (คุณภาพ — กรอกเอง)
  purchase_orders table:
    - expected_delivery_date  DATE   (วันส่งของที่ Supplier นัด)
    - actual_received_date    DATE   (วันที่รับของจริง)

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
    # ── Suppliers: เพิ่ม 2 คะแนนกรอกเอง ────────────────────────────────────
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS service_score INT DEFAULT 0",
    "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS quality_score INT DEFAULT 0",

    # ── Purchase Orders: เพิ่ม 2 วันสำหรับคำนวณ SLA ──────────────────────
    "ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS expected_delivery_date DATE",
    "ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS actual_received_date   DATE",
]

def run():
    conn = psycopg2.connect(**DB_CFG)
    try:
        with conn.cursor() as cur:
            for sql in MIGRATIONS:
                print(f"▶ {sql}")
                cur.execute(sql)
            conn.commit()
        print("\n✅ Migration สำเร็จ")
        print("   suppliers       ← service_score, quality_score")
        print("   purchase_orders ← expected_delivery_date, actual_received_date")
    except Exception as e:
        conn.rollback()
        print(f"❌ Error: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    run()
