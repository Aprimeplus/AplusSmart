# revert_so_status.py (เวอร์ชันปรับปรุง)
import psycopg2
import psycopg2.extras

# --- ตั้งค่าการเชื่อมต่อฐานข้อมูล (เหมือนใน main_app.py) ---
DB_PARAMS = {
    "host": "192.168.1.60",
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123"
}

def revert_so_status_tool():
    """
    เครื่องมือสำหรับย้อนสถานะ SO ที่ 'HR Verified' แล้ว
    - แสดงรายการ SO ทั้งหมดให้เลือก
    - สามารถย้อนสถานะทีละรายการหรือทั้งหมดได้
    """
    conn = None
    print("=" * 60)
    print("   Tool for Reverting 'HR Verified' SO Status (Improved)")
    print("=" * 60)

    try:
        conn = psycopg2.connect(**DB_PARAMS)
        while True:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. ดึงข้อมูล SO ทั้งหมดที่สามารถย้อนสถานะได้
                cursor.execute("""
                    SELECT so_number, timestamp, final_sales_amount 
                    FROM commissions 
                    WHERE status = 'HR Verified' AND is_active = 1 
                    ORDER BY timestamp DESC
                """)
                revertible_sos = cursor.fetchall()

                if not revertible_sos:
                    print("\n🎉 ไม่พบ SO ที่มีสถานะ 'HR Verified' ในระบบแล้ว")
                    break

                # 2. แสดงผลเป็นเมนูให้ผู้ใช้เลือก
                print("\nSO ทั้งหมดที่สามารถย้อนสถานะได้:")
                for i, so in enumerate(revertible_sos):
                    print(f"  [{i+1}] {so['so_number']} (Verified on: {so['timestamp']}, Sales: {so['final_sales_amount']:,.2f})")
                
                print("\nพิมพ์ 'exit' เพื่อออกจากโปรแกรม")
                
                # 3. รับ Input จากผู้ใช้
                choice = input("> กรุณาเลือกเบอร์ SO ที่ต้องการย้อนสถานะ (เช่น 1), พิมพ์ 'all' เพื่อย้อนสถานะทั้งหมด, หรือ 'exit' เพื่อออก: ").strip().lower()

                if choice == 'exit':
                    break
                
                so_to_revert = []

                if choice == 'all':
                    confirm = input("!! คุณแน่ใจหรือไม่ที่จะย้อนสถานะ SO ทั้งหมดที่แสดงในรายการ? พิมพ์ 'YES' เพื่อยืนยัน: ").strip()
                    if confirm == 'YES':
                        so_to_revert = [so['so_number'] for so in revertible_sos]
                    else:
                        print("ยกเลิกการดำเนินการ")
                        continue
                elif choice.isdigit() and 1 <= int(choice) <= len(revertible_sos):
                    so_to_revert.append(revertible_sos[int(choice) - 1]['so_number'])
                else:
                    print("⚠️ ตัวเลือกไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง")
                    continue
                
                # 4. ทำการอัปเดตฐานข้อมูล
                if so_to_revert:
                    sql_update = """
                        UPDATE commissions
                        SET 
                            status = 'Forwarded_To_HR', 
                            approver_sale_manager_key = NULL,
                            approval_date_sale_manager = NULL,
                            final_sales_amount = NULL,
                            final_cost_amount = NULL,
                            final_gp = NULL,
                            final_margin = NULL
                        WHERE 
                            so_number = ANY(%s) -- ใช้ ANY เพื่อรองรับการอัปเดตหลายรายการ
                            AND status = 'HR Verified'
                            AND is_active = 1;
                    """
                    cursor.execute(sql_update, (so_to_revert,))
                    conn.commit()
                    print(f"✅ สำเร็จ! ย้อนสถานะ SO จำนวน {cursor.rowcount} รายการเรียบร้อยแล้ว")

    except Exception as e:
        if conn:
            conn.rollback()
        print(f"❌ เกิดข้อผิดพลาด: {e}")
    finally:
        if conn:
            conn.close()
        print("ปิดการเชื่อมต่อฐานข้อมูลแล้ว")


if __name__ == "__main__":
    revert_so_status_tool()