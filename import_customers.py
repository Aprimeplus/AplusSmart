# import_customers.py (ฉบับแก้ไข รองรับ PostgreSQL เวอร์ชันเก่า)
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import psycopg2
import chardet

# --- กรุณาตรวจสอบข้อมูลการเชื่อมต่อให้ตรงกับของคุณ ---
DB_CONFIG = {
    "host": "192.168.1.60",
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123"
}
# ----------------------------------------------------

def detect_encoding(file_path):
    """ตรวจหา Encoding ของไฟล์."""
    with open(file_path, 'rb') as f:
        raw_data = f.read(100000)
    return chardet.detect(raw_data)['encoding']

def import_customers_compatible():
    """
    สคริปต์สำหรับ Import หรือ Update ข้อมูลลูกค้า (รองรับ DB เก่า)
    โดยใช้วิธี UPDATE ก่อน ถ้าไม่สำเร็จค่อย INSERT
    """
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ลูกค้า (Excel หรือ CSV)",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )

    if not file_path:
        print("ไม่ได้เลือกไฟล์. ยกเลิกการทำงาน")
        return

    try:
        if file_path.endswith('.csv'):
            encoding = detect_encoding(file_path)
            df = pd.read_csv(file_path, encoding=encoding, dtype=str)
        else:
            df = pd.read_excel(file_path, dtype=str)
        
        df.columns = df.columns.str.strip().str.lower()
        df.rename(columns={'รหัสลูกค้า': 'customer_code', 'ชื่อลูกค้า': 'customer_name'}, inplace=True)
        
        if not {'customer_code', 'customer_name'}.issubset(df.columns):
            messagebox.showerror("คอลัมน์ไม่ถูกต้อง", "ไฟล์ของคุณต้องมีคอลัมน์ 'รหัสลูกค้า' และ 'ชื่อลูกค้า'")
            return
            
        df = df.where(pd.notnull(df), None)
        df.dropna(subset=['customer_code'], inplace=True)

        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลลูกค้าที่มีรหัสในไฟล์")
            return

        print(f"พบข้อมูลลูกค้า {len(df)} รายการในไฟล์. เริ่มการ Import...")

    except Exception as e:
        messagebox.showerror("ผิดพลาดในการอ่านไฟล์", f"ไม่สามารถอ่านไฟล์ได้: {e}")
        return

    conn = None
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        
        inserted_count = 0
        updated_count = 0

        # วนลูปเพื่อจัดการข้อมูลทีละแถว
        for index, row in df.iterrows():
            code = row['customer_code']
            name = row['customer_name']

            # 1. ลอง UPDATE ก่อน
            cur.execute(
                "UPDATE customers SET customer_name = %s WHERE customer_code = %s",
                (name, code)
            )
            
            # 2. ตรวจสอบว่ามีการ UPDATE เกิดขึ้นหรือไม่
            if cur.rowcount == 0:
                # ถ้าไม่มีการ UPDATE (rowcount=0) แสดงว่าไม่มีรหัสนี้อยู่ ให้ INSERT แทน
                cur.execute(
                    "INSERT INTO customers (customer_code, customer_name, credit_term) VALUES (%s, %s, %s)",
                    (code, name, 'เงินสด') # กำหนด credit_term เริ่มต้น
                )
                inserted_count += 1
            else:
                updated_count += 1
        
        conn.commit()

        summary_message = (
            f"Import ข้อมูลเสร็จสมบูรณ์!\n\n"
            f"เพิ่มลูกค้าใหม่: {inserted_count} รายการ\n"
            f"อัปเดตข้อมูลลูกค้าเก่า: {updated_count} รายการ"
        )
        print(summary_message)
        messagebox.showinfo("สำเร็จ", summary_message)

    except Exception as e:
        if conn:
            conn.rollback()
        print(f"เกิดข้อผิดพลาด: {e}")
        messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการ Import ข้อมูล: {e}")
    finally:
        if conn:
            cur.close()
            conn.close()
            print("ปิดการเชื่อมต่อฐานข้อมูล")

if __name__ == "__main__":
    import_customers_compatible()