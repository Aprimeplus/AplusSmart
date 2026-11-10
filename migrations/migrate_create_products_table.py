# migrate_create_products_table.py
import psycopg2
from tkinter import messagebox, Tk

def create_products_table():
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
            print("กำลังสร้างตาราง 'products'...")

            # สร้างตารางใหม่สำหรับเก็บข้อมูลสินค้าหลัก
            cur.execute("""
                CREATE TABLE IF NOT EXISTS products (
                    id SERIAL PRIMARY KEY,
                    product_code TEXT UNIQUE NOT NULL,
                    product_name TEXT,
                    warehouse TEXT,
                    default_cost REAL,
                    default_weight REAL
                );
            """)

            conn.commit()
            print("สร้างตาราง 'products' เรียบร้อยแล้ว")

        messagebox.showinfo("สำเร็จ", "สร้างตาราง 'products' สำหรับเก็บข้อมูลสินค้าหลักเรียบร้อยแล้ว")

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถสร้างตารางได้: {e}")
        if conn:
            conn.rollback()
    finally:
        if conn:
            conn.close()
            print("ปิดการเชื่อมต่อฐานข้อมูล")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("ยืนยัน", "คุณต้องการสร้างตารางสำหรับข้อมูลสินค้าหลัก (Product Master) หรือไม่?"):
        create_products_table()
    root.destroy()