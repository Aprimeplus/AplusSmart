import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: markup_guide_tiers (Mark up Guide - จัดซื้อ)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            # ตาราง Markup Guide ต่อ SKU — แต่ละแถวคือ 1 Tier (T1-T5) ของ 1 สินค้า
            # เงื่อนไขราคา+น้ำหนักรวม ต้องเข้าเกณฑ์ "ทั้งคู่พร้อมกัน" ถึงจะเข้า Tier นั้น (ตามที่ยืนยันกับ PM)
            # ไม่บังคับต้องมีครบ 5 Tier ต่อ SKU — สินค้าไหนตั้งไม่ครบก็ปล่อยว่างได้ (แค่ไม่มีแถวของ Tier นั้น)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS markup_guide_tiers (
                    id              SERIAL PRIMARY KEY,
                    product_code    TEXT NOT NULL REFERENCES products(product_code)
                                        ON UPDATE CASCADE ON DELETE CASCADE,
                    tier            SMALLINT NOT NULL CHECK (tier BETWEEN 1 AND 5),
                    markup_percent  NUMERIC(6,2) NOT NULL,
                    price_min       NUMERIC(14,2),   -- NULL = ไม่จำกัดขั้นต่ำ
                    price_max       NUMERIC(14,2),   -- NULL = ไม่จำกัดขั้นสูง
                    weight_min      NUMERIC(14,3),   -- น้ำหนักรวม (กก.) ขั้นต่ำ, NULL = ไม่จำกัด
                    weight_max      NUMERIC(14,3),   -- น้ำหนักรวม (กก.) ขั้นสูง, NULL = ไม่จำกัด
                    updated_at      TIMESTAMP NOT NULL DEFAULT NOW(),
                    updated_by      VARCHAR(50),      -- sale_key ของคนแก้ล่าสุด
                    UNIQUE (product_code, tier)
                )
            """)
            print("- Created table 'markup_guide_tiers'")

            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_markup_guide_tiers_product_code
                ON markup_guide_tiers (product_code)
            """)
            print("- Created index on product_code")

        conn.commit()
        print("Migration completed successfully!")
    except Exception as e:
        conn.rollback()
        print(f"Migration failed: {e}")
    finally:
        app.release_connection(conn)
        app.destroy()

if __name__ == "__main__":
    migrate_db()
