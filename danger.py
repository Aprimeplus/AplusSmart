# danger.py (ฉบับปรับปรุงล่าสุด)

import psycopg2
from tkinter import messagebox, Tk
import bcrypt

def clear_transactional_data():
    """
    สคริปต์สำหรับล้างข้อมูลทำรายการ (Transactional Data) ในแอปพลิเคชัน
    โดยจะไม่ลบข้อมูล Master Data เช่น Users, Customers, Suppliers, Products
    """
    
    # --- ข้อมูลการเชื่อมต่อ PostgreSQL ---
    PG_HOST = "192.168.1.60" 
    PG_DBNAME = "aplus_com_test"
    PG_USER = "app_user"
    PG_PASSWORD = "cailfornia123"
    # ------------------------------------

    # --- START: ปรับปรุงลิสต์ตารางที่ต้องการล้าง ---
    # จัดลำดับใหม่ให้ลบข้อมูลย่อย (Child) ก่อนข้อมูลหลัก (Parent) และเพิ่มตาราง Archive
    tables_to_truncate = [
        # --- ข้อมูลย่อยที่อ้างอิงถึงตารางอื่น ---
        'purchase_order_items',      # รายการสินค้าใน PO (อ้างอิงถึง purchase_orders)
        'purchase_order_payments',   # การชำระเงิน PO (อ้างอิงถึง purchase_orders)
        'notifications',             # การแจ้งเตือน (อาจอ้างอิงถึง commissions หรือ po)

        # --- ตารางประวัติ (Log Tables) ---
        'comparison_logs',           # ประวัติการเปรียบเทียบข้อมูล
        'commission_payout_logs',    # ประวัติการจ่ายค่าคอม
        'audit_log',                 # ประวัติการใช้งาน

        # --- ตารางข้อมูลหลักของการทำรายการ ---
        'purchase_orders',           # ใบสั่งซื้อ
        'commissions',               # ข้อมูล Sales Order / Commissions

        # --- (ใหม่) ตารางข้อมูลที่ถูกจัดเก็บถาวร ---
        'commissions_archive',       # <<<< เพิ่ม: ข้อมูล SO ที่ถูก Archive
        'purchase_orders_archive',   # <<<< เพิ่ม: ข้อมูล PO ที่ถูก Archive
        'comparison_logs_archive',   # <<<< เพิ่ม: ประวัติการเทียบข้อมูลที่ถูก Archive
        'commission_payout_logs_archive' # <<<< เพิ่ม: ประวัติการจ่ายเงินที่ถูก Archive
    ]
    # --- END: สิ้นสุดการปรับปรุง ---

    conn = None
    try:
        print("กำลังเชื่อมต่อฐานข้อมูล PostgreSQL...")
        conn = psycopg2.connect(host=PG_HOST, dbname=PG_DBNAME, user=PG_USER, password=PG_PASSWORD)
        cur = conn.cursor()
        print("เชื่อมต่อสำเร็จ")

        # --- ล้างข้อมูลในตารางที่ระบุด้วย TRUNCATE ---
        for table in tables_to_truncate:
            try:
                print(f"กำลังล้างข้อมูลในตาราง: {table}...")
                # CASCADE: ลบข้อมูลที่เกี่ยวข้องในตารางอื่นด้วย (ถ้ามีการตั้ง FK ON DELETE CASCADE)
                # RESTART IDENTITY: รีเซ็ตลำดับ auto-increment ของ primary key
                sql = f"TRUNCATE TABLE {table} RESTART IDENTITY CASCADE;"
                cur.execute(sql)
                conn.commit() # Commit ทีละตารางเพื่อความปลอดภัย
                print(f"   -> ล้างข้อมูลตาราง {table} เรียบร้อยแล้ว")
            except psycopg2.errors.UndefinedTable:
                print(f"   -> ข้ามตาราง {table} เนื่องจากไม่พบในฐานข้อมูล")
                conn.rollback() # Rollback transaction ของตารางที่ error
            except Exception as e:
                print(f"เกิดข้อผิดพลาดขณะล้างตาราง {table}: {e}")
                conn.rollback()
                raise e

        print("\nกระบวนการล้างข้อมูลทำรายการเสร็จสิ้น!")
        messagebox.showinfo("สำเร็จ", "ล้างข้อมูลการทำรายการทั้งหมดในระบบเรียบร้อยแล้ว!\n"
                                      "ข้อมูลผู้ใช้, ลูกค้า, ซัพพลายเออร์, และสินค้ายังคงอยู่")

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถล้างข้อมูลทำรายการได้: {e}")
    finally:
        if conn:
            conn.close()
            print("ปิดการเชื่อมต่อฐานข้อมูล")

if __name__ == "__main__":
    root = Tk()
    root.withdraw() # ซ่อนหน้าต่าง Tkinter หลักที่ว่างเปล่า

    warning_message = (
        "คุณแน่ใจจริงๆ หรือที่จะรีเซ็ตข้อมูลการทำรายการทั้งหมดในฐานข้อมูล?\n\n"
        "การกระทำนี้จะลบข้อมูลต่อไปนี้ทั้งหมด:\n"
        "- ข้อมูลการขาย (Sales Orders / Commissions) ทั้งหมด\n"
        "- ข้อมูลการซื้อ (Purchase Orders และรายละเอียด) ทั้งหมด\n"
        "- ประวัติการใช้งาน (Audit Logs) และการแจ้งเตือน ทั้งหมด\n"
        "- ประวัติการเปรียบเทียบข้อมูล และประวัติการจ่ายเงิน ทั้งหมด\n"
        "- ข้อมูลเก่าที่เคยถูก Archive ไว้ทั้งหมด\n\n" # <--- เพิ่มคำอธิบาย
        "ข้อมูลผู้ใช้งาน, ลูกค้า, ซัพพลายเออร์, และสินค้าจะ *ไม่ถูกลบ* และยังคงอยู่\n\n"
        "!!!!!! การกระทำนี้ไม่สามารถย้อนกลับได้ !!!!!! "
    )
    
    if messagebox.askyesno("ยืนยันการล้างข้อมูลทำรายการ", warning_message, icon='warning'):
        clear_transactional_data()
    
    root.destroy()