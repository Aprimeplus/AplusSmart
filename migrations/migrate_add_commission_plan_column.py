# migrate_add_commission_plan_column.py
import psycopg2
from tkinter import messagebox, Tk

def add_commission_plan_column():
    conn = None
    try:
        conn = psycopg2.connect(
            host="192.168.1.60", dbname="aplus_com_test",
            user="app_user", password="cailfornia123"
        )
        with conn.cursor() as cur:
            print("กำลังเพิ่มคอลัมน์ 'commission_plan' ในตาราง sales_users...")
            cur.execute("""
                ALTER TABLE sales_users
                ADD COLUMN IF NOT EXISTS commission_plan TEXT;
            """)
            conn.commit()
        messagebox.showinfo("สำเร็จ", "เพิ่มคอลัมน์ commission_plan ในตาราง sales_users เรียบร้อยแล้ว")
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถอัปเดตตารางได้: {e}")
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    root = Tk(); root.withdraw()
    if messagebox.askyesno("ยืนยัน", "คุณต้องการเพิ่มคอลัมน์สำหรับแผนค่าคอมมิชชั่นหรือไม่?"):
        add_commission_plan_column()
    root.destroy()