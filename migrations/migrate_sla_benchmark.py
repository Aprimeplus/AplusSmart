"""
migrate_sla_benchmark.py
สร้างตาราง sla_benchmark และ migrate schema ให้รองรับการจับเวลาจาก row แรก
"""
import psycopg2

DB_CFG = dict(host="Server-APrime", dbname="aplus_com_test",
              user="app_user", password="cailfornia123")

SQL_CREATE = """
CREATE TABLE IF NOT EXISTS sla_benchmark (
    id           SERIAL PRIMARY KEY,
    so_number    TEXT,
    user_key     TEXT NOT NULL,
    row_temp_key UUID,
    started_at   TIMESTAMP NOT NULL DEFAULT NOW(),
    copied_at    TIMESTAMP,
    duration_min INTEGER
);

CREATE INDEX IF NOT EXISTS idx_sla_so   ON sla_benchmark(so_number);
CREATE INDEX IF NOT EXISTS idx_sla_user ON sla_benchmark(user_key);
CREATE INDEX IF NOT EXISTS idx_sla_uuid ON sla_benchmark(row_temp_key);
"""

# Migration สำหรับตารางที่มีอยู่แล้ว
SQL_MIGRATE = """
-- ทำให้ so_number เป็น nullable (รองรับ pending rows ที่ยังไม่มี SO)
ALTER TABLE sla_benchmark ALTER COLUMN so_number DROP NOT NULL;

-- เพิ่มคอลัมน์ row_temp_key ถ้ายังไม่มี
ALTER TABLE sla_benchmark ADD COLUMN IF NOT EXISTS row_temp_key UUID;

-- เพิ่ม index
CREATE INDEX IF NOT EXISTS idx_sla_uuid ON sla_benchmark(row_temp_key);

-- เปลี่ยน duration_min จาก GENERATED ALWAYS เป็น regular column
-- (เพื่อให้บันทึก business hours แทน wall-clock)
ALTER TABLE sla_benchmark ALTER COLUMN duration_min DROP EXPRESSION;
"""

if __name__ == "__main__":
    conn = psycopg2.connect(**DB_CFG)
    cur = conn.cursor()

    # ตรวจว่าตารางมีอยู่แล้วหรือเปล่า
    cur.execute("""
        SELECT EXISTS (
            SELECT 1 FROM information_schema.tables
            WHERE table_name = 'sla_benchmark'
        )
    """)
    exists = cur.fetchone()[0]

    if not exists:
        print("สร้างตาราง sla_benchmark ใหม่...")
        cur.execute(SQL_CREATE)
    else:
        print("ตาราง sla_benchmark มีอยู่แล้ว — migrate schema...")
        # แยก statement ที่อาจ fail ออกมา เพื่อ skip ได้ถ้า already done
        safe_stmts = [
            "ALTER TABLE sla_benchmark ALTER COLUMN so_number DROP NOT NULL",
            "ALTER TABLE sla_benchmark ADD COLUMN IF NOT EXISTS row_temp_key UUID",
            "CREATE INDEX IF NOT EXISTS idx_sla_uuid ON sla_benchmark(row_temp_key)",
        ]
        for stmt in safe_stmts:
            try:
                cur.execute(stmt)
                conn.commit()
            except Exception as e:
                conn.rollback()
                print(f"  skip: {e}")

        # DROP EXPRESSION — skip ถ้าไม่ใช่ generated column อยู่แล้ว
        try:
            cur.execute("ALTER TABLE sla_benchmark ALTER COLUMN duration_min DROP EXPRESSION")
            conn.commit()
            print("  duration_min: DROP EXPRESSION ✓")
        except Exception:
            conn.rollback()
            print("  duration_min: เป็น regular column อยู่แล้ว — skip")

    conn.commit()
    conn.close()
    print("✅ sla_benchmark schema พร้อมใช้งาน")
