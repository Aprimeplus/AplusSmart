# migrate_add_final_commission_column.py
import psycopg2
from tkinter import messagebox, Tk

def add_final_commission_column():
    conn = None
    try:
        conn = psycopg2.connect(
            host="192.168.1.60", dbname="aplus_com_test",
            user="app_user", password="cailfornia123"
        )
        with conn.cursor() as cur:
            print("กำลังเพิ่มคอลัมน์ 'final_commission' ในตาราง commissions...")
            cur.execute("""
                ALTER TABLE commissions
                ADD COLUMN IF NOT EXISTS final_commission REAL;
            """)
            conn.commit()
        messagebox.showinfo("สำเร็จ", "เพิ่มคอลัมน์ final_commission ในตาราง commissions เรียบร้อยแล้ว")
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถอัปเดตตารางได้: {e}")
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    root = Tk(); root.withdraw()
    if messagebox.askyesno("ยืนยัน", "คุณต้องการเพิ่มคอลัมน์สำหรับยอดค่าคอมมิชชั่นสุดท้ายหรือไม่?"):
        add_final_commission_column()
    root.destroy()