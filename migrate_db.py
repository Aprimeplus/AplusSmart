# migrate_db.py (ฉบับปรับปรุง)
import sqlite3
import psycopg2
import pandas as pd
from tkinter import messagebox, Tk
import traceback

def migrate_data_robust():
    SQLITE_DB_PATH = 'commission_history.db'
    
    # --- กรุณาแก้ไขข้อมูลการเชื่อมต่อ PostgreSQL ของคุณที่นี่ ---
    PG_HOST = "192.168.1.60" 
    PG_DBNAME = "aplus_com_test"
    PG_USER = "app_user"
    PG_PASSWORD = "cailfornia123"
    # ----------------------------------------------------

    pg_conn = None
    sqlite_conn = None
    try:
        print("Connecting to SQLite...")
        sqlite_conn = sqlite3.connect(SQLITE_DB_PATH)
        print("Connecting to PostgreSQL...")
        pg_conn = psycopg2.connect(host=PG_HOST, dbname=PG_DBNAME, user=PG_USER, password=PG_PASSWORD)
        print("Connections successful.")
    except Exception as e:
        messagebox.showerror("Connection Error", f"ไม่สามารถเชื่อมต่อฐานข้อมูลได้: {e}")
        return

    # <<< START OF CHANGE: ปรับลำดับตาราง >>>
    # ย้ายตารางแม่ (parent table) มาไว้ก่อนตารางลูก (child tables)
    table_names = ['sales_users', 'audit_log', 'commissions', 'purchase_orders']
    # <<< END OF CHANGE >>>
    
    pg_cursor = pg_conn.cursor()

    for table in table_names:
        try:
            print(f"\n=============================================")
            print(f"Reading data from SQLite table: {table}...")
            # อ่านข้อมูลจาก SQLite
            df = pd.read_sql_query(f"SELECT * FROM {table}", sqlite_conn)
            
            # แปลงชื่อคอลัมน์เป็นตัวพิมพ์เล็กทั้งหมด
            df.columns = [x.lower() for x in df.columns]
            
            if df.empty:
                print(f"Table {table} is empty in SQLite. Skipping.")
                continue

            # ล้างข้อมูลในตาราง PostgreSQL ก่อน Import
            print(f"Clearing PostgreSQL table: {table} before import...")
            pg_cursor.execute(f"TRUNCATE TABLE {table} RESTART IDENTITY CASCADE;")

            print(f"Preparing to write {len(df)} rows to PostgreSQL table: {table}...")
            
            # <<< START OF CHANGE: เปลี่ยนมา INSERT ทีละแถวเพื่อจัดการ Error ได้ดีขึ้น >>>
            cols = ", ".join([f'"{c}"' for c in df.columns]) # ใส่ double quote รอบชื่อคอลัมน์
            placeholders = ", ".join(["%s"] * len(df.columns))
            sql_insert = f"INSERT INTO {table} ({cols}) VALUES ({placeholders})"
            
            success_count = 0
            fail_count = 0

            # วน Loop เพื่อเพิ่มข้อมูลทีละแถว
            for index, row in df.iterrows():
                # แปลงข้อมูลในแถวเป็น tuple
                # จัดการกับค่า NaN ของ pandas ให้เป็น None ซึ่ง SQL เข้าใจได้
                data_tuple = tuple(None if pd.isna(x) else x for x in row.values)
                try:
                    pg_cursor.execute(sql_insert, data_tuple)
                    success_count += 1
                except Exception as row_error:
                    fail_count += 1
                    print(f"  - FAILED to insert row {index} into {table}. Error: {row_error}")
                    print(f"    Problematic data: {data_tuple}")
                    pg_conn.rollback() # Rollback เฉพาะแถวที่มีปัญหา
            
            pg_conn.commit() # Commit ข้อมูลทั้งหมดของตารางนี้เมื่อ Loop จบ
            # <<< END OF CHANGE >>>

            print(f"Finished migrating table {table}. Success: {success_count}, Failed: {fail_count}")

            # อัปเดต ID Sequence (ส่วนนี้ยังคงเดิมและถูกต้อง)
            if 'id' in df.columns:
                print(f"Updating sequence for table: {table}...")
                update_seq_sql = f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), coalesce(max(id), 1), max(id) IS NOT null) FROM {table};"
                pg_cursor.execute(update_seq_sql)
                pg_conn.commit()
                print(f"Sequence for {table} updated successfully.")
                
        except Exception as e:
            print(f"Could not migrate table {table}. Error: {e}")
            traceback.print_exc() # พิมพ์รายละเอียด error ทั้งหมด
            pg_conn.rollback()

    print("\nData migration process finished.")
    if sqlite_conn:
        sqlite_conn.close()
    if pg_conn:
        pg_cursor.close()
        pg_conn.close()
    messagebox.showinfo("สำเร็จ", "การย้ายข้อมูลเสร็จสิ้น!")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("Confirm Migration", "คุณต้องการเริ่มกระบวนการย้ายข้อมูลจาก SQLite ไปยัง PostgreSQL หรือไม่?\n(โปรดสำรองไฟล์ commission_history.db ก่อนดำเนินการ)"):
        migrate_data_robust()
    root.destroy()