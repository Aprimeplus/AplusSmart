"""สร้างตารางประวัติ cost_benchmarks_history (สำเนาข้อมูลก่อนถูกบันทึกทับ) — รันครั้งเดียว ปลอดภัยถ้ารันซ้ำ"""
from main_app import AppContainer

app = AppContainer()
conn = app.get_connection()
try:
    with conn.cursor() as cur:
        cur.execute("CREATE TABLE IF NOT EXISTS cost_benchmarks_history (LIKE cost_benchmarks EXCLUDING ALL)")
        cur.execute("ALTER TABLE cost_benchmarks_history ADD COLUMN IF NOT EXISTS snapshot_at TIMESTAMP DEFAULT now()")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_cbh_scope ON cost_benchmarks_history "
                    "(created_by, benchmark_year, benchmark_month, snapshot_at)")
    conn.commit()
    print("ok")
finally:
    app.release_connection(conn)
