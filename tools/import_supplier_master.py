import pandas as pd
from sqlalchemy import create_engine, text
import tkinter as tk
from tkinter import filedialog, messagebox
import urllib.parse

# ==========================================
# ⚙️ ตั้งค่าการเชื่อมต่อฐานข้อมูล (ต้องตรงกับไฟล์ Export)
# ==========================================
DB_HOST = "Server-APrime"
DB_NAME = "aplus_com_test"
DB_USER = "app_user"
DB_PASS = "cailfornia123"         
DB_PORT = "5432"
# ==========================================

def import_suppliers():
    root = tk.Tk()
    root.withdraw()

    try:
        # 1. เลือกไฟล์ Excel
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel ข้อมูล Supplier เพื่อนำเข้า",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if not file_path:
            return

        print(f"กำลังอ่านไฟล์: {file_path}")
        
        # 2. อ่านข้อมูลจาก Excel
        df = pd.read_excel(file_path)

        # 3. ตรวจสอบชื่อหัวคอลัมน์ (ต้องตรงกับตอน Export)
        # แมปชื่อไทยกลับเป็นชื่อภาษาอังกฤษใน DB
        column_map = {
            'รหัสผู้ขาย (Code)': 'supplier_code',
            'ชื่อผู้ขาย (Name)': 'supplier_name'
        }

        # ตรวจสอบว่ามีคอลัมน์ครบไหม
        if not set(column_map.keys()).issubset(df.columns):
            missing = set(column_map.keys()) - set(df.columns)
            messagebox.showerror("Format ผิดพลาด", f"ไม่พบคอลัมน์: {missing}\nกรุณาใช้ไฟล์ที่ Export ออกมาจากระบบ")
            return

        # เปลี่ยนชื่อคอลัมน์ให้ตรงกับ Database
        df.rename(columns=column_map, inplace=True)

        # ตัดช่องว่างหัวท้ายออก (กันความผิดพลาด)
        df['supplier_code'] = df['supplier_code'].astype(str).str.strip()
        df['supplier_name'] = df['supplier_name'].astype(str).str.strip()

        # 4. เชื่อมต่อฐานข้อมูล
        encoded_password = urllib.parse.quote_plus(DB_PASS)
        db_url = f"postgresql://{DB_USER}:{encoded_password}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
        engine = create_engine(db_url)
        connection = engine.connect()
        transaction = connection.begin()

        print("กำลังประมวลผลข้อมูล...")
        
        updated_count = 0
        inserted_count = 0

        try:
            for index, row in df.iterrows():
                code = row['supplier_code']
                name = row['supplier_name']

                # ข้ามแถวที่ไม่มี Code
                if not code or code.lower() == 'nan':
                    continue

                # 4.1 เช็คว่ามีข้อมูลอยู่แล้วหรือไม่
                check_query = text("SELECT COUNT(*) FROM suppliers WHERE supplier_code = :code")
                result = connection.execute(check_query, {"code": code}).scalar()

                if result > 0:
                    # 4.2 ถ้ามี -> อัปเดต (Update)
                    update_query = text("UPDATE suppliers SET supplier_name = :name WHERE supplier_code = :code")
                    connection.execute(update_query, {"name": name, "code": code})
                    updated_count += 1
                else:
                    # 4.3 ถ้าไม่มี -> เพิ่มใหม่ (Insert)
                    insert_query = text("INSERT INTO suppliers (supplier_code, supplier_name) VALUES (:code, :name)")
                    connection.execute(insert_query, {"code": code, "name": name})
                    inserted_count += 1

            transaction.commit()
            msg = (f"นำเข้าข้อมูลสำเร็จ!\n\n"
                   f"✅ เพิ่มใหม่: {inserted_count} รายการ\n"
                   f"🔄 อัปเดต: {updated_count} รายการ")
            print(msg)
            messagebox.showinfo("สำเร็จ", msg)

        except Exception as e:
            transaction.rollback()
            raise e
        finally:
            connection.close()

    except Exception as e:
        error_msg = f"เกิดข้อผิดพลาด:\n{str(e)}"
        print(error_msg)
        messagebox.showerror("Error", error_msg)

if __name__ == "__main__":
    import_suppliers()