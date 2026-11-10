import pandas as pd
from sqlalchemy import create_engine
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import urllib.parse

# ==========================================
# ⚙️ ตั้งค่าการเชื่อมต่อฐานข้อมูล (แก้ไขตรงนี้)
# ==========================================
DB_HOST = "Server-APrime"
DB_NAME = "aplus_com_test"  # ชื่อฐานข้อมูลของคุณ (เช็คจาก main_app.py หรือ pgAdmin)
DB_USER = "app_user"        # ชื่อผู้ใช้ (ปกติคือ postgres)
DB_PASS = "cailfornia123"           # ⚠️ ใส่รหัสผ่านฐานข้อมูลของคุณที่นี่
DB_PORT = "5432"
# ==========================================

def export_suppliers():
    root = tk.Tk()
    root.withdraw()  # ซ่อนหน้าต่างหลักของ Tkinter

    try:
        # 1. สร้างการเชื่อมต่อกับฐานข้อมูล
        # เข้ารหัสรหัสผ่านเพื่อรองรับอักขระพิเศษ
        encoded_password = urllib.parse.quote_plus(DB_PASS)
        db_url = f"postgresql://{DB_USER}:{encoded_password}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
        engine = create_engine(db_url)

        print("กำลังเชื่อมต่อฐานข้อมูล...")

        # 2. เขียนคำสั่ง SQL เพื่อดึงเฉพาะ Supplier Code และ Name
        # หมายเหตุ: ถ้าชื่อตารางในฐานข้อมูลไม่ใช่ 'suppliers' ให้แก้ตรง FROM ...
        query = """
            SELECT 
                supplier_code, 
                supplier_name 
            FROM suppliers 
            ORDER BY supplier_code ASC
        """

        # 3. ดึงข้อมูลมาใส่ Pandas DataFrame
        df = pd.read_sql(query, engine)

        if df.empty:
            messagebox.showwarning("ไม่พบข้อมูล", "ไม่มีข้อมูล Supplier ในฐานข้อมูล")
            return

        # เปลี่ยนชื่อหัวตารางให้สวยงาม (Optional)
        df.rename(columns={
            'supplier_code': 'รหัสผู้ขาย (Code)',
            'supplier_name': 'ชื่อผู้ขาย (Name)'
        }, inplace=True)

        print(f"ดึงข้อมูลสำเร็จ: {len(df)} รายการ")

        # 4. เปิดหน้าต่างให้เลือกที่บันทึกไฟล์
        default_filename = f"Supplier_List_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=default_filename,
            title="บันทึกไฟล์รายชื่อ Supplier"
        )

        if file_path:
            # 5. บันทึกเป็น Excel
            df.to_excel(file_path, index=False)
            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลเรียบร้อยแล้วที่:\n{file_path}")
            print("Export เสร็จสิ้น")
        else:
            print("ยกเลิกการบันทึก")

    except Exception as e:
        error_msg = f"เกิดข้อผิดพลาด:\n{str(e)}"
        print(error_msg)
        messagebox.showerror("Error", error_msg)

if __name__ == "__main__":
    export_suppliers()