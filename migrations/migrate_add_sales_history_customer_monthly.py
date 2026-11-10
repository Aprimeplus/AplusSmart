import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: sales_history_customer_monthly (ยอดขายรายลูกค้าย้อนหลัง สำหรับ Customer Monitoring / Dormant Pool)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            # เก็บยอดขายสะสมต่อลูกค้าต่อเดือน จากไฟล์ "SO Analysis Report" ย้อนหลัง (ปี 2567/2568)
            # sale_key_raw เก็บรหัสพนักงานตามไฟล์เดิม (เช่น TG001, ID001) ไม่ได้ map กับ sale_key
            # ของระบบปัจจุบัน (PIYAWAN, ILADA ฯลฯ) ตามที่ตกลงกับ PM — ใช้แสดงผลอย่างเดียว ไม่ join กับ sales_users
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS sales_history_customer_monthly (
                    id              SERIAL PRIMARY KEY,
                    customer_code   TEXT NOT NULL,
                    customer_name   TEXT,
                    sale_key_raw    TEXT,
                    year            SMALLINT NOT NULL,   -- ค.ศ.
                    month           SMALLINT NOT NULL CHECK (month BETWEEN 1 AND 12),
                    total_amount    NUMERIC(14,2) NOT NULL DEFAULT 0,
                    source_sheet    TEXT,                -- เช่น 'SO All24', 'SO All25' (ที่มาของข้อมูล เผื่อสอบย้อนกลับ)
                    imported_at     TIMESTAMP NOT NULL DEFAULT NOW(),
                    UNIQUE (customer_code, sale_key_raw, year, month)
                )
            """)
            print("- Created table 'sales_history_customer_monthly'")

            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_sales_history_customer_monthly_customer
                ON sales_history_customer_monthly (customer_code)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_sales_history_customer_monthly_year_month
                ON sales_history_customer_monthly (year, month)
            """)
            print("- Created indexes")

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
