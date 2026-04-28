# migrate_add_new_fees.py
import psycopg2
from tkinter import messagebox, Tk

def add_new_fee_columns():
    conn = None
    try:
        conn = psycopg2.connect(
            host="192.168.1.60", dbname="aplus_com_test",
            user="app_user", password="cailfornia123"
        )
        with conn.cursor() as cur:
            print("กำลังเพิ่มคอลัมน์ใหม่ในตาราง commissions...")
            # เพิ่มคอลัมน์สำหรับ ค่าบริการอื่นๆ และ ตัวเลือก VAT ของยอดขายหลัก
            cur.execute("""
                ALTER TABLE commissions
                ADD COLUMN IF NOT EXISTS sales_service_vat_option TEXT,
                ADD COLUMN IF NOT EXISTS other_service_fee REAL,
                ADD COLUMN IF NOT EXISTS other_service_fee_vat_option TEXT;
            """)
            # รวมค่าตัดและเจาะเป็นคอลัมน์เดียว
            cur.execute("""
                ALTER TABLE commissions
                ADD COLUMN IF NOT EXISTS cutting_drilling_fee REAL,
                ADD COLUMN IF NOT EXISTS cutting_drilling_fee_vat_option TEXT;
            """)
            conn.commit()
        messagebox.showinfo("สำเร็จ", "เพิ่มคอลัมน์สำหรับค่าบริการใหม่เรียบร้อยแล้ว")
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถอัปเดตตารางได้: {e}")
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    root = Tk(); root.withdraw()
    if messagebox.askyesno("ยืนยัน", "คุณต้องการอัปเดตตาราง commissions เพื่อรองรับค่าบริการใหม่หรือไม่?"):
        add_new_fee_columns()
    root.destroy()