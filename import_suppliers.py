# import_suppliers.py (ฉบับแก้ไข - แปลงข้อมูล Credit Term)

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import psycopg2
import psycopg2.extras
import traceback

def run_supplier_import():
    db_params = {
        "host": "192.168.1.60", 
        "dbname": "aplus_com_test",
        "user": "app_user", 
        "password": "cailfornia123"
    }

    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ข้อมูลซัพพลายเออร์ (Excel/CSV)",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )
    if not file_path:
        print("ยกเลิกการเลือกไฟล์")
        return

    try:
        # อ่านข้อมูลและบังคับให้เป็น String เพื่อความแน่นอน
        df = pd.read_excel(file_path, dtype=str) if file_path.endswith('.xlsx') else pd.read_csv(file_path, dtype=str)
        
        df.columns = [str(col).strip().lower() for col in df.columns]

        column_map = {
            'รหัส': 'supplier_code', 'code': 'supplier_code',
            'ชื่อซัพ': 'supplier_name', 'name': 'supplier_name',
            'เครดิต': 'credit_term', 'credit': 'credit_term'
        }
        df.rename(columns=lambda c: column_map.get(c, c), inplace=True)

        if not {'supplier_code', 'supplier_name'}.issubset(df.columns):
            messagebox.showerror("คอลัมน์ไม่ถูกต้อง", "ไฟล์ของคุณต้องมีคอลัมน์ 'รหัส' และ 'ชื่อซัพ' เป็นอย่างน้อย")
            return
            
        df = df.where(pd.notnull(df), None)
        df.dropna(subset=['supplier_code'], inplace=True)

        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลซัพพลายเออร์ที่มีรหัสในไฟล์")
            return

    except Exception as e:
        messagebox.showerror("ผิดพลาดในการอ่านไฟล์", f"ไม่สามารถอ่านไฟล์ได้: {e}")
        return

    conn = None
    try:
        conn = psycopg2.connect(**db_params)
        cur = conn.cursor()

        # --- START: EDIT - ลบข้อมูลเก่าทั้งหมดก่อน ---
        if messagebox.askyesno("ยืนยันการล้างข้อมูล", "คุณต้องการลบข้อมูลซัพพลายเออร์เก่าทั้งหมดก่อนนำเข้าข้อมูลใหม่ใช่หรือไม่?", icon='warning'):
            print("กำลังลบข้อมูลซัพพลายเออร์เก่า...")
            cur.execute("TRUNCATE TABLE suppliers RESTART IDENTITY CASCADE;")
            print("ลบข้อมูลเก่าเรียบร้อยแล้ว")
        # --- END: EDIT ---

        upsert_sql = """
        INSERT INTO suppliers (supplier_code, supplier_name, credit_term)
        VALUES (%s, %s, %s)
        ON CONFLICT (supplier_code) DO UPDATE SET
            supplier_name = EXCLUDED.supplier_name,
            credit_term = EXCLUDED.credit_term;
        """

        data_to_upsert = []
        
        # --- START: EDIT - เพิ่ม Logic การแปลงค่า ---
        credit_term_map = {
            '0': 'เงินสด',
            '7': 'Cr 7',
            '15': 'Cr 15',
            '30': 'Cr 30'
        }
        # --- END: EDIT ---

        for _, row in df.iterrows():
            # --- START: EDIT - แปลงค่า credit_term ---
            raw_credit = str(row.get('credit_term', '0')).strip()
            # ใช้ค่าที่แปลงแล้ว หรือใช้ค่าเดิมถ้าหาไม่เจอ
            display_term = credit_term_map.get(raw_credit, raw_credit)
            # --- END: EDIT ---

            data_to_upsert.append((
                row.get('supplier_code'),
                row.get('supplier_name'),
                display_term # <-- ใช้ค่าที่แปลงแล้ว
            ))

        psycopg2.extras.execute_batch(cur, upsert_sql, data_to_upsert)
        conn.commit()
        messagebox.showinfo("สำเร็จ", f"นำเข้า/อัปเดตข้อมูลซัพพลายเออร์จำนวน {len(df)} รายการเรียบร้อยแล้ว")

    except Exception as e:
        if conn: conn.rollback()
        messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดระหว่างการนำเข้าข้อมูล:\n{traceback.format_exc()}")
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    run_supplier_import()