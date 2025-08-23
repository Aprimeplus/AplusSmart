# reset_hr_password.py (ฉบับแก้ไข - สร้าง User ใหม่หากไม่พบ)

import psycopg2
import bcrypt
from tkinter import messagebox, Tk

def reset_hr_password():
    """
    ทำการเชื่อมต่อฐานข้อมูลและอัปเดตรหัสผ่านของ user 'hr-007'
    หากไม่พบ จะทำการสร้าง User ขึ้นมาใหม่
    """
    # --- ข้อมูลการเชื่อมต่อ (กรุณาตรวจสอบให้ถูกต้อง) ---
    db_params = {
        "host": "192.168.1.60",
        "dbname": "aplus_com_test",
        "user": "app_user",
        "password": "cailfornia123"
    }
    
    user_to_reset = 'hr'
    new_password = '007'
    
    conn = None
    try:
        print("Connecting to database...")
        conn = psycopg2.connect(**db_params)
        print("Connection successful.")
        
        with conn.cursor() as cur:
            # 1. เข้ารหัสผ่านใหม่ด้วย bcrypt
            print(f"Hashing new password for user '{user_to_reset}'...")
            hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
            
            # 2. ลองอัปเดต password_hash ในฐานข้อมูลก่อน
            print(f"Attempting to update password hash in the database...")
            cur.execute(
                "UPDATE sales_users SET password_hash = %s WHERE sale_key = %s",
                (hashed_password.decode('utf-8'), user_to_reset)
            )
            
            # ตรวจสอบว่ามีการอัปเดตเกิดขึ้นจริงหรือไม่
            if cur.rowcount == 0:
                # ถ้าไม่เจอ user ให้สร้างใหม่
                print(f"User '{user_to_reset}' not found, creating new user...")
                insert_sql = """
                    INSERT INTO sales_users (sale_key, sale_name, password_hash, role, status)
                    VALUES (%s, %s, %s, %s, %s)
                """
                # กำหนดค่าสำหรับ User ใหม่
                new_user_data = (
                    user_to_reset,
                    'HR Admin', # ชื่อที่แสดง
                    hashed_password.decode('utf-8'),
                    'HR', # Role
                    'Active' # Status
                )
                cur.execute(insert_sql, new_user_data)
                conn.commit()
                message = f"สร้างผู้ใช้งาน '{user_to_reset}' และตั้งรหัสผ่านเรียบร้อยแล้ว"
                print(f"SUCCESS: {message}")
                messagebox.showinfo("สำเร็จ", message)

            else:
                # ถ้าอัปเดตสำเร็จ
                conn.commit()
                message = f"รีเซ็ตรหัสผ่านสำหรับผู้ใช้ '{user_to_reset}' เป็น '{new_password}' เรียบร้อยแล้ว"
                print(f"SUCCESS: {message}")
                messagebox.showinfo("สำเร็จ", message)

    except Exception as e:
        if conn:
            conn.rollback()
        print(f"An error occurred: {e}")
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถรีเซ็ตรหัสผ่านได้: {e}")
    finally:
        if conn:
            conn.close()
            print("Database connection closed.")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    
    reset_hr_password()