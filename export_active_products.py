import pandas as pd
import psycopg2
import psycopg2.extras
from tkinter import filedialog, messagebox
from datetime import datetime
import os

def export_active_products_to_excel(db_connection):
    """
    Export ข้อมูลสินค้าเฉพาะรายการที่มีประวัติการใช้งาน (เคยถูกเปิด PO) 
    ออกเป็นไฟล์ Excel พร้อมคอลัมน์ System_ID สำหรับการ Import กลับ
    """
    try:
        # SQL Query: เลือกสินค้าจากตาราง products 
        # โดยมีเงื่อนไขว่า รหัสสินค้า (product_code) ต้องมีอยู่ในตาราง purchase_order_items
        query = """
            SELECT DISTINCT
                p.id,
                p.product_code,
                p.product_name,
                p.warehouse,
                p.last_unit_price,
                p.last_weight_per_unit
            FROM products p
            INNER JOIN purchase_order_items poi ON p.product_code = poi.product_code
            ORDER BY p.product_code
        """

        # อ่านข้อมูลด้วย Pandas
        df = pd.read_sql_query(query, db_connection)

        if df.empty:
            messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบสินค้าที่มีประวัติการใช้งานในระบบ")
            return

        # เปลี่ยนชื่อคอลัมน์ให้ตรงกับฟอร์ม Import ของคุณ (สำคัญมาก: System_ID)
        df.rename(columns={
            'id': 'System_ID',  # คีย์สำคัญสำหรับการ Update
            'product_code': 'รหัสสินค้า',
            'product_name': 'ชื่อสินค้า',
            'warehouse': 'คลัง',
            'last_unit_price': 'ราคาล่าสุด',
            'last_weight_per_unit': 'น้ำหนักล่าสุด'
        }, inplace=True)

        # ตั้งชื่อไฟล์ Default
        filename = f"active_products_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        
        # เปิดหน้าต่าง Save File
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="บันทึกไฟล์สินค้าที่มีการเคลื่อนไหว",
            initialfile=filename
        )

        if save_path:
            # บันทึกเป็น Excel
            df.to_excel(save_path, index=False)
            messagebox.showinfo(
                "สำเร็จ", 
                f"Export ข้อมูลสินค้าที่ใช้งานแล้วจำนวน {len(df)} รายการ\nเรียบร้อยแล้วที่:\n{save_path}\n\n"
                "Tip: คุณสามารถแก้ไข รหัส/ชื่อสินค้า ในไฟล์นี้\nแล้วนำกลับไป Import เพื่ออัปเดตข้อมูลได้ทันที"
            )

    except Exception as e:
        messagebox.showerror("Export Error", f"เกิดข้อผิดพลาดในการ Export: {e}")
        print(f"Error: {e}")

# ==============================================================================
# วิธีการนำไปทดสอบ (Test Run)
# ==============================================================================
if __name__ == "__main__":
    # ส่วนนี้สำหรับทดสอบรันไฟล์นี้เดี่ยวๆ (ต้องแก้ค่า DB Config ให้ตรงกับเครื่องคุณ)
    try:
        # จำลองการเชื่อมต่อ Database (ใส่ค่าให้ตรงกับของคุณ)
        conn = psycopg2.connect(
            dbname="aplus_com_test",
            user="app_user",
            password="cailfornia123",
            host="Server-APrime",
            port="5432"
        )
        
        # เรียกใช้งานฟังก์ชัน
        export_active_products_to_excel(conn)
        
        conn.close()
    except Exception as e:
        print("ไม่สามารถเชื่อมต่อฐานข้อมูลเพื่อทดสอบได้ (หากนำไปใช้จริงในแอป ให้ข้ามส่วนนี้ไป):", e)