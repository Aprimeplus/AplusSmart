import psycopg2
import pandas as pd
from datetime import datetime
import traceback  # <--- เพิ่มบรรทัดนี้ครับ

# --- 1. ตั้งค่าการเชื่อมต่อฐานข้อมูล (แก้ไขให้ตรงกับเครื่องคุณ) ---
DB_CONFIG = {
    "dbname": "aplus_com_test",  # ชื่อ Database ของคุณ
    "user": "app_user",         # User
    "password": "cailfornia123",     # Password
    "host": "192.168.1.51",
    "port": "5432"
}

def test_commission_query():
    conn = None
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        print("\n" + "="*50)
        print("🧪 เริ่มการทดสอบ Logic การดึงข้อมูลคอมมิชชั่น")
        print("="*50)

        # --- 2. จำลองข้อมูล (Mock Data) ---
        # สร้าง Sale คนสมมติ
        test_sale_key = "TEST_SALE_001"
        cursor.execute("DELETE FROM commissions WHERE sale_key = %s", (test_sale_key,)) # ล้างของเก่าก่อน
        cursor.execute("DELETE FROM sales_users WHERE sale_key = %s", (test_sale_key,))
        
        cursor.execute("""
            INSERT INTO sales_users (sale_key, sale_name, role, status, commission_plan)
            VALUES (%s, 'Mr. Test Logic', 'Sale', 'Active', 'Plan A')
        """, (test_sale_key,))

        # สร้าง SO จำลอง 3 ใบ (สถานะ HR Verified, ยังไม่จ่าย)
        # ใบที่ 1: เดือน 10 (ตุลาคม) - ค้างจ่าย
        # ใบที่ 2: เดือน 11 (พฤศจิกายน) - เดือนปัจจุบัน
        # ใบที่ 3: เดือน 12 (ธันวาคม) - อนาคต
        
        mock_sos = [
            ('SO-TEST-OCT', 10, 2024, 10000), # เดือน 10
            ('SO-TEST-NOV', 11, 2024, 20000), # เดือน 11
            ('SO-TEST-DEC', 12, 2024, 30000)  # เดือน 12
        ]
        
        print(f"\n📝 สร้างข้อมูลจำลองสำหรับ Sale: {test_sale_key}")
        for so_num, month, year, amount in mock_sos:
            cursor.execute("""
                INSERT INTO commissions (
                    so_number, sale_key, commission_month, commission_year, 
                    sales_service_amount, final_sales_amount, final_cost_amount, 
                    status, is_active, payout_id
                ) VALUES (%s, %s, %s, %s, %s, %s, 0, 'HR Verified', 1, NULL)
            """, (so_num, test_sale_key, month, year, amount, amount))
            print(f"   - สร้าง {so_num}: เดือน {month}/{year} (สถานะ: HR Verified, รอจ่าย)")

        conn.commit()

        # --- 3. รัน Query ที่เราแก้ไข (จำลองการทำงานของ HR) ---
        # สมมติว่า HR เลือกคำนวณรอบ "พฤศจิกายน 2024" (เดือน 11)
        target_month = 11
        target_year = 2024
        
        print(f"\n🔍 HR เลือกคำนวณรอบ: {target_month}/{target_year}")
        print(f"   (คาดหวัง: ต้องเจอ SO เดือน 10 และ 11 แต่ต้องไม่เจอ 12)")

        # นี่คือ Query ที่อยู่ใน hr_screen.py ของคุณ
        query = """
            SELECT so_number, commission_month, commission_year 
            FROM commissions c
            WHERE c.sale_key = %s 
                AND c.status = 'HR Verified' 
                AND c.payout_id IS NULL
                AND c.is_active = 1
                AND (
                    (c.commission_year < %s) 
                    OR 
                    (c.commission_year = %s AND c.commission_month <= %s)
                )
        """
        
        # Params: (sale_key, year, year, month)
        params = (test_sale_key, target_year, target_year, target_month)
        
        df = pd.read_sql_query(query, conn, params=params)

        # --- 4. ตรวจสอบผลลัพธ์ ---
        print("\n📊 ผลลัพธ์ที่ได้จาก Query:")
        if df.empty:
            print("   ❌ ไม่พบข้อมูล (ผิดปกติ)")
        else:
            print(df.to_string(index=False))
            
            found_sos = df['so_number'].tolist()
            
            # Check Condition
            has_oct = 'SO-TEST-OCT' in found_sos
            has_nov = 'SO-TEST-NOV' in found_sos
            has_dec = 'SO-TEST-DEC' in found_sos
            
            print("\n✅ การตรวจสอบความถูกต้อง:")
            print(f"   1. เก็บตกเดือนเก่า (ต.ค.) ได้หรือไม่?  -> {'ผ่าน ✅' if has_oct else 'ไม่ผ่าน ❌'}")
            print(f"   2. ดึงเดือนปัจจุบัน (พ.ย.) ได้หรือไม่? -> {'ผ่าน ✅' if has_nov else 'ไม่ผ่าน ❌'}")
            print(f"   3. กันเดือนอนาคต (ธ.ค.) ออกหรือไม่?   -> {'ผ่าน ✅' if not has_dec else 'ไม่ผ่าน ❌ (หลุดมา)'}")
            
            if has_oct and has_nov and not has_dec:
                print("\n🎉 สรุป: ระบบทำงานถูกต้อง 100%!")
            else:
                print("\n⚠️ สรุป: ยังมีจุดผิดพลาด กรุณาเช็ค Logic อีกครั้ง")

        # --- 5. ล้างข้อมูลทดสอบ ---
        print("\n🧹 กำลังล้างข้อมูลทดสอบ...")
        cursor.execute("DELETE FROM commissions WHERE sale_key = %s", (test_sale_key,))
        cursor.execute("DELETE FROM sales_users WHERE sale_key = %s", (test_sale_key,))
        conn.commit()
        print("   ล้างข้อมูลเรียบร้อย")

    except Exception as e:
        print(f"\n❌ เกิดข้อผิดพลาด: {e}")
        traceback.print_exc()
    finally:
        if conn: conn.close()

if __name__ == "__main__":
    test_commission_query()