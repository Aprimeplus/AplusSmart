# import_product_master.py (ฉบับปรับปรุง)
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import psycopg2
import psycopg2.extras
import traceback

def run_product_import():
    db_params = {
        "host": "192.168.1.60", "dbname": "aplus_com_test",
        "user": "app_user", "password": "cailfornia123"
    }

    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ข้อมูลสินค้า (Excel/CSV)",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )
    if not file_path: return

    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถอ่านไฟล์ได้: {e}")
        return

    # --- START: ปรับปรุงการจัดการคอลัมน์ ---
    # ทำความสะอาดชื่อคอลัมน์ในไฟล์ Excel/CSV (ลบช่องว่าง, ทำให้เป็นตัวเล็ก)
    df.columns = [str(col).strip().lower() for col in df.columns]

    # ขยาย Map ให้รองรับชื่อที่อาจแตกต่างกัน และเพิ่ม cost/weight
    column_map = {
    'รหัสสินค้า': 'product_code', 'product code': 'product_code',
    'ชื่อสินค้า': 'product_name', 'product name': 'product_name',
    'คลัง': 'warehouse', 'คลังสินค้า': 'warehouse',
    'ต้นทุน': 'default_cost', 'cost': 'default_cost',
    'น้ำหนัก': 'default_weight', 'weight': 'default_weight',
    'ราคา': 'default_price', 'price': 'default_price' 
}
    
    # แปลงชื่อคอลัมน์จากภาษาไทยเป็นอังกฤษ
    df.rename(columns=lambda c: column_map.get(c, c), inplace=True)
    
    # --- END: ปรับปรุงการจัดการคอลัมน์ ---

    conn = None
    try:
        conn = psycopg2.connect(**db_params)
        cur = conn.cursor()

        # --- START: ปรับปรุง SQL UPSERT ---
        upsert_sql = """
        INSERT INTO products (product_code, product_name, warehouse, default_cost, default_weight, default_price)
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (product_code) DO UPDATE SET
           product_name = EXCLUDED.product_name,
           warehouse = EXCLUDED.warehouse,
           default_cost = EXCLUDED.default_cost,
           default_weight = EXCLUDED.default_weight,
           default_price = EXCLUDED.default_price;
        """
        # --- END: ปรับปรุง SQL UPSERT ---

        data_to_upsert = []
        for _, row in df.iterrows():
            # แปลงค่าว่างให้เป็น None เพื่อให้ฐานข้อมูลเข้าใจ
            cost = row.get('default_cost')
            weight = row.get('default_weight')
            price = row.get('default_price')
           
            data_to_upsert.append((
        row.get('product_code'),
        row.get('product_name'),
        row.get('warehouse'),
        float(cost) if pd.notna(cost) else None,
        float(weight) if pd.notna(weight) else None,
        float(price) if pd.notna(price) else None # <-- เพิ่มบรรทัดนี้
    ))

        psycopg2.extras.execute_batch(cur, upsert_sql, data_to_upsert)
        conn.commit()
        messagebox.showinfo("สำเร็จ", f"นำเข้า/อัปเดตข้อมูลสินค้าจำนวน {len(df)} รายการเรียบร้อยแล้ว")

    except Exception as e:
        if conn: conn.rollback()
        messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดระหว่างการนำเข้าข้อมูล: {traceback.format_exc()}")
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    run_product_import()