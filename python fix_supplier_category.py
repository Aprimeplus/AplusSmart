# fix_supplier_category.py
# อัปเดต category ใน suppliers โดยดูจากข้อมูลใน cost_benchmarks
# รันครั้งเดียว — ปลอดภัย เพราะ UPDATE เฉพาะแถวที่ category ยังว่างอยู่

import psycopg2
from tkinter import messagebox, Tk

DB = dict(host="Server-Aprime", dbname="aplus_com_test",
          user="app_user", password="cailfornia123")

def fix():
    conn = None
    try:
        conn = psycopg2.connect(**DB)
        with conn.cursor() as cur:

            # 1. ดึง column ชื่อภาษาไทยจาก cost_benchmarks
            cur.execute("""
                SELECT column_name 
                FROM information_schema.columns 
                WHERE table_name = 'cost_benchmarks'
                ORDER BY ordinal_position
            """)
            cols = [r[0] for r in cur.fetchall()]
            print("Columns in cost_benchmarks:")
            for i, c in enumerate(cols):
                print(f"  {i}: {c}")

            # หา index ของ column "ชื่อ Supplier" และ "หมวด"
            supplier_col = next((c for c in cols if 'Supplier' in c and c != 'Supplier2' 
                                  and 'ID' not in c and 'Sup' not in c[:3]), None)
            category_col = next((c for c in cols if c in ('หมวด',) or 
                                  (len(c) < 6 and 'หมวด' in c)), None)

            print(f"\nSupplier column: {repr(supplier_col)}")
            print(f"Category column: {repr(category_col)}")

            if not supplier_col or not category_col:
                print("ERROR: ไม่พบ column ที่ต้องการ")
                messagebox.showerror("Error", 
                    f"ไม่พบ column:\n"
                    f"Supplier: {repr(supplier_col)}\n"
                    f"Category: {repr(category_col)}\n\n"
                    f"Columns ที่มี:\n" + "\n".join(cols[:20]))
                return

            # 2. ดึง mapping: supplier_name → category ที่ใช้บ่อยที่สุด
            cur.execute(f"""
                SELECT sup_name, top_cat, cnt
                FROM (
                    SELECT 
                        "{supplier_col}" AS sup_name,
                        "{category_col}" AS top_cat,
                        COUNT(*) AS cnt,
                        ROW_NUMBER() OVER (
                            PARTITION BY "{supplier_col}" 
                            ORDER BY COUNT(*) DESC
                        ) AS rn
                    FROM cost_benchmarks
                    WHERE "{supplier_col}" IS NOT NULL 
                      AND "{supplier_col}" != ''
                      AND "{category_col}" IS NOT NULL 
                      AND "{category_col}" != ''
                    GROUP BY "{supplier_col}", "{category_col}"
                ) ranked
                WHERE rn = 1
                ORDER BY cnt DESC
            """)
            mappings = cur.fetchall()
            print(f"\nพบ mapping {len(mappings)} รายการ:")
            for name, cat, cnt in mappings[:20]:
                print(f"  {name[:30]:<30} → {cat}  ({cnt} ครั้ง)")

            if not mappings:
                messagebox.showinfo("แจ้งเตือน", 
                    "ไม่พบข้อมูลใน cost_benchmarks\n"
                    "กรุณาตรวจสอบว่ามีข้อมูลในตารางเทียบราคาก่อน")
                return

            # 3. UPDATE suppliers ที่ยัง category ว่าง
            updated = 0
            skipped = 0
            for sup_name, category, cnt in mappings:
                cur.execute("""
                    UPDATE suppliers 
                    SET category = %s
                    WHERE supplier_name = %s
                      AND (category IS NULL OR category = '' OR category = 'Tier 2')
                """, (category, sup_name))
                rows = cur.rowcount
                if rows > 0:
                    updated += rows
                    print(f"  ✓ Updated: {sup_name[:30]} → {category}")
                else:
                    skipped += 1

            conn.commit()
            print(f"\nเสร็จสิ้น! Updated: {updated} แถว, Skipped: {skipped} แถว")

            # 4. แสดงสรุป category ที่ได้
            cur.execute("""
                SELECT category, COUNT(*) as cnt 
                FROM suppliers 
                WHERE category IS NOT NULL AND category != ''
                GROUP BY category 
                ORDER BY cnt DESC
            """)
            summary = cur.fetchall()
            summary_text = "\n".join([f"  {cat}: {cnt} ร้าน" for cat, cnt in summary])
            print(f"\nสรุป category:\n{summary_text}")

            messagebox.showinfo(
                "อัปเดตสำเร็จ",
                f"อัปเดต category จากประวัติการเทียบราคา\n\n"
                f"อัปเดต: {updated} รายการ\n"
                f"ข้าม (มีข้อมูลอยู่แล้ว): {skipped} รายการ\n\n"
                f"สรุป category:\n{summary_text}"
            )

    except Exception as e:
        if conn:
            conn.rollback()
        print(f"ERROR: {e}")
        import traceback; traceback.print_exc()
        messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาด:\n{e}")
    finally:
        if conn:
            conn.close()


if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno(
        "ยืนยัน",
        "จะอัปเดต category ของ Supplier\n"
        "โดยดูจากประวัติการเทียบราคา (cost_benchmarks)\n\n"
        "เฉพาะ Supplier ที่ยัง category ว่างอยู่เท่านั้น\n"
        "ดำเนินการต่อหรือไม่?"
    ):
        fix()
    root.destroy()