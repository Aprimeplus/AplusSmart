"""สร้างตาราง supplier_categories — รายการหมวดที่อนุมัติแล้วสำหรับดรอปดาวน์ "หมวดสินค้า" ของ Supplier

ดรอปดาวน์เดิมดึงเฉพาะหมวดที่ Supplier ใช้อยู่จริง หมวดใหม่ที่ยังไม่มีซัพใช้จึงไม่ขึ้น
ตารางนี้ทำให้เพิ่มหมวดใหม่ได้ที่ฐานข้อมูลเลย (ไม่ต้อง build setup ใหม่) — ดู tools/add_supplier_category.py

รัน (ปลอดภัยถ้ารันซ้ำ):  python -m migrations.migrate_add_supplier_categories
"""
from main_app import AppContainer

SEED = ["บริการ (บก)"]

app = AppContainer()
conn = app.get_connection()
try:
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS supplier_categories (
                category_name TEXT PRIMARY KEY,
                note          TEXT,
                created_at    TIMESTAMP DEFAULT now()
            )
        """)
        for name in SEED:
            cur.execute("INSERT INTO supplier_categories (category_name, note) VALUES (%s, %s) "
                        "ON CONFLICT DO NOTHING", (name, "อนุมัติโดยผู้จัดการฝ่ายขาย"))
    conn.commit()
    print("ok")
finally:
    app.release_connection(conn)
