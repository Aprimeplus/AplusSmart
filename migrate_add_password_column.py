# migrate_add_password_column.py
import psycopg2
from tkinter import messagebox, Tk

def add_password_hash_column():
    """
    สคริปต์สำหรับเพิ่มคอลัมน์ password_hash ในตาราง sales_users
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
            print("กำลังตรวจสอบและเพิ่มคอลัมน์ 'password_hash'...")
            
            # ใช้ ALTER TABLE เพื่อเพิ่มคอลัมน์ถ้ายังไม่มีอยู่
            # IF NOT EXISTS เป็น синтаксис ของ PostgreSQL ที่ช่วยให้รันซ้ำได้โดยไม่เกิด Error
            cur.execute("""
                ALTER TABLE sales_users
                ADD COLUMN IF NOT EXISTS password_hash TEXT;
            """)
            
            conn.commit()
            print("ตรวจสอบและเพิ่มคอลัมน์ 'password_hash' เรียบร้อยแล้ว")
            
            # (Optional) ตั้งค่า default password สำหรับ user ที่ยังไม่มี
            # หมายเหตุ: รหัสผ่านนี้ไม่ปลอดภัย ควรให้ HR ตั้งใหม่ให้ทุกคน
            # print("กำลังตั้งรหัสผ่านเริ่มต้นสำหรับผู้ใช้ที่ยังไม่มี...")
            # default_password = "password123"
            # hashed_password = bcrypt.hashpw(default_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            # cur.execute("UPDATE sales_users SET password_hash = %s WHERE password_hash IS NULL", (hashed_password,))
            # conn.commit()
            # print("ตั้งรหัสผ่านเริ่มต้นสำเร็จ")

        messagebox.showinfo("สำเร็จ", "อัปเดตโครงสร้างตาราง 'sales_users' เรียบร้อยแล้ว\nเพิ่มคอลัมน์ password_hash สำเร็จ")

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
    if messagebox.askyesno("ยืนยันการอัปเดตฐานข้อมูล", "คุณต้องการเพิ่มคอลัมน์สำหรับเก็บรหัสผ่านหรือไม่?\n(โปรดสำรองฐานข้อมูลก่อนดำเนินการ)"):
        add_password_hash_column()
    root.destroy()
