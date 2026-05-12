# migrate_markup_tiers.py
# รัน 1 ครั้งเพื่อสร้างตาราง markup_tiers ใน DB
# python migrate_markup_tiers.py

import psycopg2

_DB_CFG = dict(host="192.168.1.60", dbname="aplus_com_test",
               user="app_user", password="cailfornia123")

SQL = """
-- ===== Markup Tiers Table =====
CREATE TABLE IF NOT EXISTS markup_tiers (
    id              SERIAL PRIMARY KEY,
    main_category   TEXT NOT NULL,
    sub_category    TEXT,
    product_type    TEXT NOT NULL,
    product_champ   TEXT,
    cost_per_kg     NUMERIC(12,4),
    tier_name       TEXT NOT NULL,      -- 'Guest','T1','T2','T3','T4','T5'
    tier_order      INTEGER NOT NULL,   -- 0=Guest 1=T1 ... 5=T5
    markup_pct      NUMERIC(8,4),       -- e.g. 0.11 = 11%
    amount_range    TEXT,               -- ข้อความอธิบายช่วงยอด เช่น "น้อยกว่า 25,000 บ."
    qty_range       TEXT,               -- ข้อความอธิบายช่วงจำนวน เช่น "ไม่ถึง 1 ตัน"
    status_note     TEXT,               -- สถานะจาก Excel เช่น "รอ observe ราคาตลาด"
    is_active       BOOLEAN DEFAULT TRUE,
    excel_updated_date  DATE,
    excel_updated_by    TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_markup_main_cat  ON markup_tiers(main_category);
CREATE INDEX IF NOT EXISTS idx_markup_prod_type ON markup_tiers(product_type);
CREATE INDEX IF NOT EXISTS idx_markup_tier_name ON markup_tiers(tier_name);

-- Auto-update updated_at
CREATE OR REPLACE FUNCTION fn_markup_tiers_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS trg_markup_tiers_updated_at ON markup_tiers;
CREATE TRIGGER trg_markup_tiers_updated_at
    BEFORE UPDATE ON markup_tiers
    FOR EACH ROW EXECUTE FUNCTION fn_markup_tiers_updated_at();
"""

if __name__ == "__main__":
    print("Connecting to DB...")
    conn = psycopg2.connect(**_DB_CFG)
    conn.autocommit = True
    cur = conn.cursor()
    print("Running migration...")
    cur.execute(SQL)
    print("✅ markup_tiers table created (or already exists).")
    cur.close()
    conn.close()
