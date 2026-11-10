"""
migrate_quality_events.py
--------------------------
สร้าง supplier_quality_events table และ initialize quality_score = 100
สำหรับ Supplier ที่ยังไม่มีคะแนน (ระบบใหม่เริ่ม 20/20 = 100 คะแนน)

รันครั้งเดียว ปลอดภัย: ใช้ CREATE TABLE IF NOT EXISTS
"""

import psycopg2

DB_CFG = dict(
    host="Server-Aprime",
    dbname="aplus_com_test",
    user="app_user",
    password="cailfornia123",
)

MIGRATIONS = [
    # สร้าง quality events table
    """
    CREATE TABLE IF NOT EXISTS supplier_quality_events (
        id              SERIAL PRIMARY KEY,
        supplier_name   TEXT    NOT NULL,
        event_type      TEXT    NOT NULL,
        delta           INT     NOT NULL,
        reason          TEXT,
        recovery_label  TEXT,
        resolved        BOOLEAN DEFAULT FALSE,
        parent_event_id INT,
        created_by      TEXT,
        created_at      TIMESTAMP DEFAULT NOW()
    )
    """,
    # initialize quality_score = 100 สำหรับ Supplier ที่ยังเป็น 0 (หรือ NULL)
    """
    UPDATE suppliers
    SET quality_score = 100
    WHERE quality_score IS NULL OR quality_score = 0
    """,
]

def run():
    conn = psycopg2.connect(**DB_CFG)
    try:
        with conn.cursor() as cur:
            for sql in MIGRATIONS:
                label = sql.strip().split('\n')[0][:80]
                print(">> " + label)
                cur.execute(sql)
            conn.commit()
        print("\nMigration OK")
        print("   supplier_quality_events table created (if not exists)")
        print("   quality_score initialized to 100 for all suppliers with 0")
    except Exception as e:
        conn.rollback()
        print("ERROR: " + str(e))
    finally:
        conn.close()

if __name__ == "__main__":
    run()
