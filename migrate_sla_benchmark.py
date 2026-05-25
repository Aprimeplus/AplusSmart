"""
migrate_sla_benchmark.py
สร้างตาราง sla_benchmark สำหรับเก็บ SLA ของจัดซื้อ
"""
import psycopg2

DB_CFG = dict(host="Server-APrime", dbname="aplus_com_test",
              user="app_user", password="cailfornia123")

SQL = """
CREATE TABLE IF NOT EXISTS sla_benchmark (
    id           SERIAL PRIMARY KEY,
    so_number    TEXT NOT NULL,
    user_key     TEXT NOT NULL,
    started_at   TIMESTAMP NOT NULL DEFAULT NOW(),
    copied_at    TIMESTAMP,
    duration_min INTEGER GENERATED ALWAYS AS (
        CASE WHEN copied_at IS NOT NULL
             THEN EXTRACT(EPOCH FROM (copied_at - started_at))::INTEGER / 60
             ELSE NULL
        END
    ) STORED
);

CREATE INDEX IF NOT EXISTS idx_sla_so ON sla_benchmark(so_number);
CREATE INDEX IF NOT EXISTS idx_sla_user ON sla_benchmark(user_key);
"""

if __name__ == "__main__":
    conn = psycopg2.connect(**DB_CFG)
    cur = conn.cursor()
    cur.execute(SQL)
    conn.commit()
    conn.close()
    print("✅ sla_benchmark table created")
