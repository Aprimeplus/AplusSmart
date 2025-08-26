# revert_so_status.py (เวอร์ชันปรับปรุง)
import psycopg2
import psycopg2.extras
import json

# --- ตั้งค่าการเชื่อมต่อฐานข้อมูล (เหมือนใน main_app.py) ---
DB_PARAMS = {
    "host": "192.168.1.60",
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123"
}

def revert_so_status_tool():
    """
    เครื่องมือสำหรับย้อนสถานะ SO ที่ 'HR Verified' หรือ 'Paid' แล้ว
    - แสดงรายการ SO ที่เกี่ยวข้องทั้งหมดให้เลือก
    - สามารถย้อนสถานะทีละรายการหรือทั้งหมดได้
    - มีคำเตือนพิเศษสำหรับการย้อนสถานะ SO ที่จ่ายเงินไปแล้ว
    - บันทึกการทำงานลงใน Audit Log
    """
    conn = None
    print("=" * 70)
    print("      Tool for Reverting 'HR Verified' / 'Paid' SO Status")
    print("=" * 70)

    try:
        conn = psycopg2.connect(**DB_PARAMS)
        while True:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. ดึงข้อมูล SO ทั้งหมดที่สามารถย้อนสถานะได้ (HR Verified และ Paid)
                cursor.execute("""
                    SELECT 
                        c.id, c.so_number, c.status, c.timestamp, 
                        c.final_sales_amount, c.payout_id, u.sale_name
                    FROM commissions c
                    LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                    WHERE c.status IN ('HR Verified', 'Paid') AND c.is_active = 1 
                    ORDER BY c.status, c.timestamp DESC
                """)
                revertible_sos = cursor.fetchall()

                if not revertible_sos:
                    print("\n🎉 ไม่พบ SO ที่มีสถานะ 'HR Verified' หรือ 'Paid' ในระบบแล้ว")
                    break

                # 2. แสดงผลเป็นเมนูให้ผู้ใช้เลือก
                print("\nSO ทั้งหมดที่สามารถย้อนสถานะได้:")
                for i, so in enumerate(revertible_sos):
                    status_info = f"Status: {so['status']}"
                    if so['status'] == 'Paid':
                        status_info += f" (Payout ID: {so['payout_id'] or 'N/A'})"
                    
                    print(f"  [{i+1}] {so['so_number']} (Sale: {so['sale_name'] or 'N/A'}, {status_info})")
                
                print("\nพิมพ์ 'exit' เพื่อออกจากโปรแกรม")
                
                # 3. รับ Input จากผู้ใช้
                choice = input("> กรุณาเลือกเบอร์ SO ที่ต้องการย้อนสถานะ (เช่น 1), พิมพ์ 'all' เพื่อย้อนทั้งหมด, หรือ 'exit' เพื่อออก: ").strip().lower()

                if choice == 'exit':
                    break
                
                selected_records = []

                if choice == 'all':
                    selected_records = revertible_sos
                elif choice.isdigit() and 1 <= int(choice) <= len(revertible_sos):
                    selected_records.append(revertible_sos[int(choice) - 1])
                else:
                    print("⚠️  ตัวเลือกไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง")
                    continue
                
                if not selected_records:
                    continue

                # ตรวจสอบว่ามี SO ที่สถานะ 'Paid' อยู่ในรายการที่เลือกหรือไม่
                has_paid_so = any(rec['status'] == 'Paid' for rec in selected_records)
                
                print("-" * 30)
                print("รายการที่เลือก:")
                for rec in selected_records:
                    print(f"  - {rec['so_number']} (Status: {rec['status']})")
                
                if has_paid_so:
                    print("\n" + "!"*70)
                    print("!! คำเตือน: คุณได้เลือก SO ที่ถูกจ่ายเงินไปแล้ว (สถานะ 'Paid') !!")
                    print("การย้อนสถานะจะทำให้ข้อมูลประวัติการจ่ายเงิน (Payout Log) ไม่ถูกต้อง")
                    print("กรุณาตรวจสอบและยืนยันกับฝ่ายที่เกี่ยวข้องหลังดำเนินการเสร็จสิ้น")
                    print("!"*70)

                confirm = input("คุณแน่ใจหรือไม่ที่จะย้อนสถานะรายการที่เลือก? พิมพ์ 'YES' เพื่อยืนยัน: ").strip()

                if confirm != 'YES':
                    print("ยกเลิกการดำเนินการ")
                    continue
                    
                # 4. ทำการอัปเดตฐานข้อมูล
                so_to_revert = [rec['so_number'] for rec in selected_records]
                
                sql_update = """
                    UPDATE commissions
                    SET 
                        status = 'Forwarded_To_HR', 
                        approver_sale_manager_key = NULL,
                        approval_date_sale_manager = NULL,
                        final_sales_amount = NULL,
                        final_cost_amount = NULL,
                        final_gp = NULL,
                        final_margin = NULL,
                        payout_id = NULL -- << ล้าง Payout ID ที่เคยผูกอยู่
                    WHERE 
                        so_number = ANY(%s)
                        AND status IN ('HR Verified', 'Paid')
                        AND is_active = 1;
                """
                cursor.execute(sql_update, (so_to_revert,))
                updated_rows = cursor.rowcount
                conn.commit()
                print(f"✅ สำเร็จ! ย้อนสถานะ SO จำนวน {updated_rows} รายการเรียบร้อยแล้ว")

                # 5. บันทึก Audit Log
                with conn.cursor() as log_cursor:
                    for record in selected_records:
                        log_details = {
                            'reverted_so': record['so_number'],
                            'original_status': record['status'],
                            'original_payout_id': record['payout_id']
                        }
                        log_cursor.execute("""
                            INSERT INTO audit_log (action, table_name, record_id, user_info, summary_json)
                            VALUES (%s, %s, %s, %s, %s)
                        """, (
                            'Revert SO Status', 
                            'commissions', 
                            record['id'], 
                            'revert_so_status_tool.py', 
                            json.dumps(log_details)
                        ))
                    conn.commit()
                print("📝 การย้อนสถานะถูกบันทึกใน Audit Log เรียบร้อยแล้ว")


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