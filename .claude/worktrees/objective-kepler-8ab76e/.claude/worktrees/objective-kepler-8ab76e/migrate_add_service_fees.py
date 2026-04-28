# migrate_add_service_fees.py
import psycopg2
from tkinter import messagebox, Tk

def add_service_fee_columns():
    """
    สคริปต์สำหรับเพิ่มคอลัมน์ค่าบริการตัดและเจาะ
    ในตาราง commissions
    """
    conn = None
    try:
        # --- กรุณาตรวจสอบข้อมูลการเชื่อมต่อให้ตรงกับของคุณ ---
        conn = psycopg2.connect(
            host="192.168.1.60",
            dbname="aplus_com_test",
            user="app_user",
            password="cailfornia123"
        )
        print("เชื่อมต่อฐานข้อมูลสำเร็จ")
        
        with conn.cursor() as cur:
            print("กำลังเพิ่มคอลัมน์สำหรับค่าบริการตัดและเจาะ...")
            
            # เพิ่มคอลัมน์ 4 ตัว
            cur.execute("""
                ALTER TABLE commissions
                ADD COLUMN IF NOT EXISTS cutting_fee REAL,
                ADD COLUMN IF NOT EXISTS cutting_fee_vat_option TEXT,
                ADD COLUMN IF NOT EXISTS drilling_fee REAL,
                ADD COLUMN IF NOT EXISTS drilling_fee_vat_option TEXT;
            """)
            
            conn.commit()
            print("เพิ่มคอลัมน์เรียบร้อยแล้ว")
            
        messagebox.showinfo("สำเร็จ", "อัปเดตโครงสร้างตาราง 'commissions' สำเร็จ")

    except psycopg2.Error as e:
        print(f"Database error: {e}")
        messagebox.showerror("Database Error", f"ไม่สามารถอัปเดตตารางได้: {e}")
        if conn:
            conn.rollback()
    finally:
        if conn:
            conn.close()
            print("ปิดการเชื่อมต่อฐานข้อมูล")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("ยืนยันการอัปเดตฐานข้อมูล", "คุณต้องการเพิ่มคอลัมน์สำหรับค่าบริการตัด/เจาะหรือไม่?"):
        add_service_fee_columns()
    root.destroy()