# migrate_markup_tiers.py
# รัน 1 ครั้งเพื่อสร้างตาราง markup_tiers ใน DB
# python migrate_markup_tiers.py

import psycopg2

_DB_CFG = dict(host="Server-APrime", dbname="aplus_com_test",
               user="app_user", password="cailfornia123")

SQL = """
-- ===== Markup Tiers Table =====
CREATE TABLE IF NOT EXISTS markup_tiers (
    id              SERIAL PRIMARY KEY,
    main_category   TEXT NOT NULL DEFAULT '',
    sub_category    TEXT,
    product_type    TEXT NOT NULL DEFAULT '',
    product_champ   TEXT,
    cost_per_kg     NUMERIC(12,4),
    tier_name       TEXT NOT NULL DEFAULT '',
    tier_order      INTEGER NOT NULL DEFAULT 0,
    markup_pct      NUMERIC(8,4),
    amount_range    TEXT,
    qty_range       TEXT,
    status_note     TEXT,
    is_active       BOOLEAN DEFAULT TRUE,
    excel_updated_date  DATE,
    excel_updated_by    TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- เพิ่ม column ที่อาจยังไม่มี (กรณีตารางมีอยู่แล้ว)
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS main_category      TEXT NOT NULL DEFAULT '';
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS sub_category       TEXT;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS product_type       TEXT NOT NULL DEFAULT '';
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS product_champ      TEXT;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS cost_per_kg        NUMERIC(12,4);
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS tier_name          TEXT NOT NULL DEFAULT '';
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS tier_order         INTEGER NOT NULL DEFAULT 0;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS markup_pct         NUMERIC(8,4);
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS amount_range       TEXT;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS qty_range          TEXT;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS status_note        TEXT;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS is_active          BOOLEAN DEFAULT TRUE;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS excel_updated_date DATE;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS excel_updated_by   TEXT;
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS created_at         TIMESTAMPTZ DEFAULT NOW();
ALTER TABLE markup_tiers ADD COLUMN IF NOT EXISTS updated_at         TIMESTAMPTZ DEFAULT NOW();

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
