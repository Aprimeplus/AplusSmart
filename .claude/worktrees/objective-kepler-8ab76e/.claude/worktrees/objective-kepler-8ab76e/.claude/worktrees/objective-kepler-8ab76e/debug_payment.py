import pandas as pd
from main_app import AppContainer

def check_payment_data():
    # ใส่เลข SO ที่คุณกำลังทดสอบ (แก้ตรงนี้)
    target_so = "SO6901AM051" 
    
    print(f"--- กำลังตรวจสอบข้อมูลของ {target_so} ---")
    app = AppContainer()
    conn = app.get_connection()
    
    try:
        # ดึงค่า payment1_amount และ payment2_amount ออกมาดูตรงๆ
        df = pd.read_sql(
            f"SELECT id, so_number, payment1_amount, payment2_amount, total_payment_amount FROM commissions WHERE so_number = '{target_so}'", 
            app.pg_engine
        )
        
        if not df.empty:
            print("ผลลัพธ์ที่อยู่ในฐานข้อมูลตอนนี้:")
            print(df.to_string(index=False))
            
            p1 = df.iloc[0]['payment1_amount']
            
            if p1 is None or p1 == 0:
                print("\n🔴 ปัญหาเจอแล้ว: ยอด payment1_amount เป็น 0 หรือ ว่าง")
                print("   -> วิธีแก้: ต้องเข้าหน้าแก้ไข (Edit Window) แล้วกรอกยอดเงินลงไปใหม่ครับ")
            else:
                print(f"\n✅ ข้อมูลมีค่า: {p1}")
                print("   -> ถ้า PDF ยังไม่ขึ้น แสดงว่าไฟล์ po_document_generator.py ยังไม่อัปเดต")
        else:
            print("❌ ไม่พบ SO เบอร์นี้ในระบบ")
            
    except Exception as e:
        print(f"Error: {e}")
    finally:
        app.release_connection(conn)

if __name__ == "__main__":
    check_payment_data()