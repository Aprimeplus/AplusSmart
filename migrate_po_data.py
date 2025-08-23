# migrate_po_data.py (ฉบับแก้ไขล่าสุด)
import psycopg2
import psycopg2.extras
import json
from tkinter import messagebox, Tk
from datetime import datetime

def parse_date(date_str, format_str="%d-%b-%Y"):
    """แปลงค่า String ของวันที่จากฟอร์มให้เป็น Date Object ที่ถูกต้อง"""
    thai_month_map = {
        "ม.ค.": "Jan", "ก.พ.": "Feb", "มี.ค.": "Mar", "เม.ย.": "Apr", "พ.ค.": "May", "มิ.ย.": "Jun",
        "ก.ค.": "Jul", "ส.ค.": "Aug", "ก.ย.": "Sep", "ต.ค.": "Oct", "พ.ย.": "Nov", "ธ.ค.": "Dec"
    }
    try:
        if not date_str or not isinstance(date_str, str): return None
        parts = date_str.split('-')
        if len(parts) != 3: return None
        day, month_th, year_be = parts
        month_en = thai_month_map.get(month_th)
        if not month_en: return None
        year_ad = int(year_be) - 543
        return datetime.strptime(f"{day}-{month_en}-{year_ad}", format_str).date()
    except (ValueError, TypeError, KeyError, AttributeError):
        return None

def convert_to_float(value_str):
    """แปลง String ที่อาจมี comma ให้เป็น float อย่างปลอดภัย"""
    if value_str is None or value_str == '': return 0.0
    try:
        return float(str(value_str).replace(",", ""))
    except (ValueError, TypeError):
        return 0.0

def migrate_po_json_data():
    conn = None
    try:
        conn = psycopg2.connect(
            host="192.168.1.60", dbname="aplus_com_test",
            user="app_user", password="cailfornia123"
        )
        print("เชื่อมต่อฐานข้อมูลสำเร็จ")
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        
        cur.execute("SELECT id, form_data_json FROM purchase_orders WHERE form_data_json IS NOT NULL AND form_data_json != ''")
        all_pos = cur.fetchall()
        
        print(f"พบ {len(all_pos)} รายการ PO ที่ต้องย้ายข้อมูล...")

        for po_row in all_pos:
            po_id = po_row['id']
            form_data_json = po_row['form_data_json']
            try:
                print(f"\nกำลังประมวลผล PO ID: {po_id}")
                data = json.loads(form_data_json)
                
                summary = data.get('summary', {})
                shipping = data.get('shipping', {})

                update_sql = """
                    UPDATE purchase_orders SET
                        credit_term = %(credit_term)s, total_cost = %(total_cost)s, total_weight = %(total_weight)s, 
                        wht_3_percent_checked = %(wht_3_percent_checked)s, wht_3_percent_amount = %(wht_3_percent_amount)s,
                        vat_7_percent_checked = %(vat_7_percent_checked)s, vat_7_percent_amount = %(vat_7_percent_amount)s,
                        grand_total = %(grand_total)s, hired_truck_cost = %(hired_truck_cost)s, delivery_date = %(delivery_date)s,
                        shipping_company_name = %(shipping_company_name)s, truck_name = %(truck_name)s, 
                        actual_payment_amount = %(actual_payment_amount)s, actual_payment_date = %(actual_payment_date)s, 
                        supplier_account_number = %(supplier_account_number)s, shipping_type = %(shipping_type)s, 
                        shipping_notes = %(shipping_notes)s
                    WHERE id = %(po_id)s
                """
                update_params = {
                    "credit_term": data.get('credit_term'), "total_cost": convert_to_float(summary.get('total_cost')),
                    "total_weight": convert_to_float(summary.get('total_weight')), "wht_3_percent_checked": summary.get('wht_3_percent_checked', False),
                    "wht_3_percent_amount": convert_to_float(summary.get('wht_3_percent_amount')), "vat_7_percent_checked": summary.get('vat_7_percent_checked', False),
                    "vat_7_percent_amount": convert_to_float(summary.get('vat_7_percent_amount')), "grand_total": convert_to_float(summary.get('grand_total')),
                    "hired_truck_cost": convert_to_float(shipping.get('hired_truck_cost')), "delivery_date": parse_date(shipping.get('delivery_date')),
                    "shipping_company_name": shipping.get('company_name'), "truck_name": shipping.get('truck_name'),
                    "actual_payment_amount": convert_to_float(shipping.get('actual_payment_amount')), "actual_payment_date": parse_date(shipping.get('actual_payment_date')),
                    "supplier_account_number": shipping.get('supplier_account_number'), "shipping_type": shipping.get('type'),
                    "shipping_notes": shipping.get('shipping_notes'), "po_id": po_id
                }
                cur.execute(update_sql, update_params)
                print(f"  - อัปเดตตารางหลักสำเร็จ")

                # <<< START OF CHANGE: เพิ่มการลบข้อมูลเก่าก่อน Insert ใหม่ >>>
                cur.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = %s", (po_id,))
                cur.execute("DELETE FROM purchase_order_payments WHERE purchase_order_id = %s", (po_id,))
                # <<< END OF CHANGE >>>

                items = data.get('items', [])
                if items:
                    for item in items:
                        item_sql = """
                            INSERT INTO purchase_order_items 
                            (purchase_order_id, product_name, status, quantity, weight_per_unit, unit_price, total_weight, total_price)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        """
                        item_params = (po_id, item.get('product_name'), item.get('status'), convert_to_float(item.get('quantity')), 
                                       convert_to_float(item.get('weight_per_unit')), convert_to_float(item.get('unit_price')), 
                                       convert_to_float(item.get('total_weight')), convert_to_float(item.get('total')))
                        cur.execute(item_sql, item_params)
                    print(f"  - เพิ่ม {len(items)} รายการสินค้าสำเร็จ")
                
                payments = data.get('payments', [])
                if payments:
                    for payment in payments:
                        payment_sql = """
                            INSERT INTO purchase_order_payments
                            (purchase_order_id, payment_type, amount, payment_date) VALUES (%s, %s, %s, %s)
                        """
                        payment_params = (po_id, payment.get('type'), convert_to_float(payment.get('amount')), parse_date(payment.get('date')))
                        cur.execute(payment_sql, payment_params)
                    print(f"  - เพิ่ม {len(payments)} รายการชำระเงินสำเร็จ")

                conn.commit()
            except (json.JSONDecodeError, KeyError) as e:
                print(f"!!! ผิดพลาดในการประมวลผล PO ID: {po_id}, ข้อมูล JSON อาจไม่สมบูรณ์. Error: {e}")
                conn.rollback()
            except Exception as e:
                print(f"!!! เกิดข้อผิดพลาดร้ายแรงกับ PO ID: {po_id}. Error: {e}")
                conn.rollback()

        print("\nกระบวนการย้ายข้อมูลเสร็จสิ้น!")
        messagebox.showinfo("สำเร็จ", "การย้ายข้อมูล PO จาก JSON ไปยังตารางใหม่เสร็จสิ้น!")
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถเชื่อมต่อหรือทำงานกับฐานข้อมูลได้: {e}")
    finally:
        if conn:
            conn.close()
            print("ปิดการเชื่อมต่อฐานข้อมูล")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("ยืนยันการย้ายข้อมูล", "คุณต้องการเริ่มย้ายข้อมูล PO จาก JSON หรือไม่?\n**กรุณาสำรองฐานข้อมูลก่อนดำเนินการ!**"):
        migrate_po_json_data()
    root.destroy()