"""เพิ่ม / ลบ หมวดสินค้าของ Supplier ในดรอปดาวน์ (ตาราง supplier_categories) — ไม่ต้อง build โปรแกรมใหม่

ใช้หลังผู้จัดการฝ่ายขายอนุมัติหมวดใหม่ตามขั้นตอน SOP "ขอเพิ่มหมวดใหม่" (ผู้ใช้เปิดโปรแกรมใหม่ถึงจะเห็น)

    python -m tools.add_supplier_category "บริการ (บก)"            # เพิ่ม
    python -m tools.add_supplier_category "บริการ (บก)" --remove    # ลบจากรายการ (ไม่แตะ Supplier ที่ใช้หมวดนี้อยู่)
    python -m tools.add_supplier_category --list                     # ดูรายการหมวดที่เพิ่มด้วยวิธีนี้

หมายเหตุ: หมวดที่มี Supplier ใช้อยู่แล้วจะขึ้นในดรอปดาวน์เองอยู่แล้ว ตารางนี้ใช้เฉพาะหมวดที่ "อนุมัติแล้วแต่ยังไม่มีซัพใช้"
"""
import argparse
import sys

from main_app import AppContainer


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("name", nargs="?", help="ชื่อหมวด เช่น 'บริการ (บก)'")
    ap.add_argument("--remove", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    if not a.list and not (a.name and a.name.strip()):
        ap.error("ต้องระบุชื่อหมวด หรือใช้ --list")

    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("CREATE TABLE IF NOT EXISTS supplier_categories ("
                        "category_name TEXT PRIMARY KEY, note TEXT, created_at TIMESTAMP DEFAULT now())")
            if a.list:
                cur.execute("SELECT category_name, created_at FROM supplier_categories ORDER BY category_name")
                for r in cur.fetchall():
                    print(r[0], "|", r[1])
            elif a.remove:
                cur.execute("DELETE FROM supplier_categories WHERE category_name = %s", (a.name.strip(),))
                print("ลบแล้ว" if cur.rowcount else "ไม่พบหมวดนี้ในรายการ")
            else:
                cur.execute("INSERT INTO supplier_categories (category_name, note) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                            (a.name.strip(), "เพิ่มด้วย tools/add_supplier_category.py"))
                print("เพิ่มแล้ว" if cur.rowcount else "มีหมวดนี้อยู่แล้ว")
        conn.commit()
    finally:
        app.release_connection(conn)


if __name__ == "__main__":
    main()
