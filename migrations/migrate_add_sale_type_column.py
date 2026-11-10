# migrate_add_sale_type_column.py
import psycopg2
from tkinter import messagebox, Tk

def add_sale_type_column():
    """
    สคริปต์สำหรับเพิ่มคอลัมน์ sale_type ในตาราง sales_users
    ให้รันสคริปต์นี้เพียงครั้งเดียวเท่านั้น
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
            print("กำลังตรวจสอบและเพิ่มคอลัมน์ 'sale_type'...")
            
            # ใช้ ALTER TABLE เพื่อเพิ่มคอลัมน์ถ้ายังไม่มีอยู่
            cur.execute("""
                ALTER TABLE sales_users
                ADD COLUMN IF NOT EXISTS sale_type TEXT;
            """)
            
            conn.commit()
            print("ตรวจสอบและเพิ่มคอลัมน์ 'sale_type' เรียบร้อยแล้ว")
            
        messagebox.showinfo("สำเร็จ", "อัปเดตโครงสร้างตาราง 'sales_users' เรียบร้อยแล้ว\nเพิ่มคอลัมน์ sale_type สำเร็จ")

    except psycopg2.Error as e:
        print(f"Database error: {e}")
        messagebox.showerror("Database Error", f"ไม่สามารถอัปเดตตารางได้: {e}")
        if conn:
            conn.rollback()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        messagebox.showerror("เกิดข้อผิดพลาด", f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")
    finally:
        if conn:
            conn.close()
            print("ปิดการเชื่อมต่อฐานข้อมูล")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("ยืนยันการอัปเดตฐานข้อมูล", "คุณต้องการเพิ่มคอลัมน์สำหรับประเภทฝ่ายขาย (sale_type) หรือไม่?\n(โปรดสำรองฐานข้อมูลก่อนดำเนินการ)"):
        add_sale_type_column()
    root.destroy()