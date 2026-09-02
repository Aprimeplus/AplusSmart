# export_utils.py (ฉบับแก้ไขล่าสุด: เพิ่มฟังก์ชัน export รายละเอียดค่าคอม)

import tkinter as tk
from tkinter import messagebox, filedialog
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import traceback
import psycopg2.extras
from customtkinter import CTkToplevel, CTkFrame, CTkLabel, CTkButton
from custom_widgets import DateSelector

# --- START: เพิ่มคลาสสำหรับหน้าต่างเลือกช่วงเวลา ---
class DateRangeDialog(CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("เลือกช่วงเวลาสำหรับ Export")
        self.geometry("480x320")

        self.start_date = None
        self.end_date = None

        # --- Quick selection buttons ---
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=15)
        CTkButton(button_frame, text="เดือนนี้", command=self.set_this_month).pack(side="left", padx=5)
        CTkButton(button_frame, text="เดือนที่แล้ว", command=self.set_last_month).pack(side="left", padx=5)
        CTkButton(button_frame, text="ปีนี้", command=self.set_this_year).pack(side="left", padx=5)

        # --- Manual date selectors ---
        date_frame = CTkFrame(self)
        date_frame.pack(pady=10, padx=20, fill="x")
        date_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(date_frame, text="วันที่เริ่มต้น:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.start_date_selector = DateSelector(date_frame)
        self.start_date_selector.grid(row=0, column=1, padx=10, pady=10)

        CTkLabel(date_frame, text="วันที่สิ้นสุด:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.end_date_selector = DateSelector(date_frame)
        self.end_date_selector.grid(row=1, column=1, padx=10, pady=10)
        self.end_date_selector._set_to_today() # Set end date to today by default

        # --- Confirm/Cancel buttons ---
        confirm_frame = CTkFrame(self, fg_color="transparent")
        confirm_frame.pack(pady=20)
        CTkButton(confirm_frame, text="ตกลง", command=self.on_ok, width=120, font=("", 14, "bold")).pack(side="left", padx=10)
        CTkButton(confirm_frame, text="ยกเลิก", command=self.on_cancel, fg_color="gray").pack(side="left", padx=10)

        self.transient(master)
        self.grab_set()

    def set_this_month(self):
        today = date.today()
        start = today.replace(day=1)
        self.start_date_selector.set_date(start)
        self.end_date_selector.set_date(today)

    def set_last_month(self):
        today = date.today()
        first_day_of_this_month = today.replace(day=1)
        last_month = first_day_of_this_month - timedelta(days=1)
        start = last_month.replace(day=1)
        self.start_date_selector.set_date(start)
        self.end_date_selector.set_date(last_month)

    def set_this_year(self):
        today = date.today()
        start = today.replace(month=1, day=1)
        self.start_date_selector.set_date(start)
        self.end_date_selector.set_date(today)

    def on_ok(self):
        self.start_date = self.start_date_selector.get_date()
        self.end_date = self.end_date_selector.get_date()

        if not self.start_date:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกวันที่เริ่มต้น", parent=self)
            return
        
        if not self.end_date:
            self.end_date = date.today() # ถ้าไม่เลือกวันที่สิ้นสุด ให้เป็นวันนี้

        if self.start_date > self.end_date:
            messagebox.showwarning("วันที่ผิดพลาด", "วันที่เริ่มต้นต้องมาก่อนวันที่สิ้นสุด", parent=self)
            return

        self.destroy()

    def on_cancel(self):
        self.start_date = None
        self.end_date = None
        self.destroy()
# --- END: เพิ่มคลาสสำหรับหน้าต่างเลือกช่วงเวลา ---


def export_approved_pos_to_excel(parent_window, pg_engine):
    """
    (เวอร์ชันแก้ไข) Export ข้อมูล PO ที่มีสถานะ 'Approved' ทั้งหมด
    พร้อมข้อมูล SO ที่เชื่อมกัน ออกเป็นไฟล์ Excel
    """
    dialog = DateRangeDialog(parent_window)
    parent_window.wait_window(dialog)

    start_date = dialog.start_date
    end_date = dialog.end_date

    if not start_date or not end_date:
        print("Export canceled by user.")
        return

    try:
        # --- [แก้ไข] ใช้ Query ที่ JOIN ตารางสมบูรณ์แล้ว ---
        query = """
        SELECT
            po.*,
            c.bill_date,
            c.customer_name,
            c.customer_type,
            c.sales_service_amount,
            c.sale_key,
            su.sale_name
        FROM
            purchase_orders po
        LEFT JOIN
            commissions c ON po.so_number = c.so_number AND c.is_active = 1
        LEFT JOIN
            sales_users su ON c.sale_key = su.sale_key
        WHERE
            po.status = 'Approved'
            AND po.timestamp::date BETWEEN %s AND %s
        ORDER BY
            po.timestamp DESC;
        """

        df = pd.read_sql_query(query, pg_engine, params=(start_date, end_date))

        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูล PO ที่อนุมัติแล้วในช่วงเวลาที่เลือก", parent=parent_window)
            return

        # (แนะนำ) ดึง Header Map ตัวหลักมาจาก AppContainer เพื่อให้เป็นมาตรฐานเดียวกัน
        header_map = {}
        if hasattr(parent_window, 'app_container') and hasattr(parent_window.app_container, 'HEADER_MAP'):
             header_map = parent_window.app_container.HEADER_MAP
        
        # แปลงชื่อคอลัมน์ทั้งหมดเท่าที่มีใน Map
        df.rename(columns=lambda c: header_map.get(c, c), inplace=True)

        default_filename = f"approved_po_export_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="บันทึกไฟล์ PO ที่อนุมัติแล้ว",
            initialfile=default_filename,
            parent=parent_window
        )

        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลทั้งหมดเรียบร้อยแล้วที่:\n{save_path}", parent=parent_window)

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=parent_window)
        traceback.print_exc()


# --- START: ฟังก์ชันใหม่สำหรับ Export รายละเอียดค่าคอม ---

def export_commission_details_to_excel(parent_window, pg_engine, sale_key, sale_name=""):
    """
    Export ข้อมูลรายละเอียดค่าคอมมิชชั่น (สันนิษฐานว่าสถานะ 'Paid') 
    สำหรับพนักงานขายที่เลือก (sale_key) ตามช่วงวันที่
    """
    # 1. เปิดหน้าต่างเลือกช่วงวันที่
    dialog = DateRangeDialog(parent_window)
    parent_window.wait_window(dialog)

    start_date = dialog.start_date
    end_date = dialog.end_date

    if not start_date or not end_date:
        print("Export canceled by user.")
        return

    try:
        # 2. สร้าง Query เพื่อดึงข้อมูล
        # (เราจะดึงเฉพาะรายการ 'Paid' ที่อยู่ในช่วงวันที่ที่เลือก และสำหรับ sale_key นี้)
        query = """
        SELECT
            c.so_number,
            c.bill_date,
            c.customer_name,
            c.customer_type,
            c.sales_service_amount,
            c.final_sales_amount,
            c.final_cost_amount,
            c.final_gp,
            c.final_margin,
            c.final_commission,
            c.commission_plan,
            c.status,
            su.sale_name
        FROM
            commissions c
        LEFT JOIN
            sales_users su ON c.sale_key = su.sale_key
        WHERE
            c.sale_key = %s
            AND c.bill_date::date BETWEEN %s AND %s
            AND c.is_active = 1
            AND c.status = 'Paid'  -- ดึงเฉพาะรายการที่จ่ายแล้ว
        ORDER BY
            c.bill_date DESC;
        """
        
        # 3. ดึงข้อมูลด้วย Pandas
        df = pd.read_sql_query(query, pg_engine, params=(sale_key, start_date, end_date))

        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", 
                                   f"ไม่พบข้อมูลค่าคอมมิชชั่นที่ 'Paid' ของ {sale_name}\n"
                                   f"ในช่วงวันที่ {start_date.strftime('%d/%m/%Y')} ถึง {end_date.strftime('%d/%m/%Y')}", 
                                   parent=parent_window)
            return

        # 4. แปลงชื่อคอลัมน์ (ใช้ HEADER_MAP จาก AppContainer ถ้ามี)
        header_map = {}
        if hasattr(parent_window, 'app_container') and hasattr(parent_window.app_container, 'HEADER_MAP'):
             header_map = parent_window.app_container.HEADER_MAP
        
        df.rename(columns=lambda c: header_map.get(c, c), inplace=True)
        
        # 5. ถามที่บันทึกไฟล์
        safe_sale_name = sale_name.replace(" ", "_") if sale_name else sale_key
        default_filename = (f"commission_paid_{safe_sale_name}_"
                            f"{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx")
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="บันทึกรายละเอียดค่าคอมมิชชั่นที่จ่ายแล้ว",
            initialfile=default_filename,
            parent=parent_window
        )

        # 6. บันทึกไฟล์ Excel
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลสำเร็จ!\nบันทึกที่: {save_path}", parent=parent_window)

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=parent_window)
        traceback.print_exc()

# --- END: ฟังก์ชันใหม่ ---

import json
import psycopg2.extras # <--- อาจจะต้องเพิ่ม import นี้ ถ้ายังไม่มี

def export_payout_so_list_to_excel(parent_window, app_container, payout_id):
    """
    (แก้ไขใหม่) Export รายการ SO โดยคำนวณ Margin และ Status 
    ให้ตรงกับหน้าจอ PayoutDetailWindow 100%
    """
    conn = None
    try:
        conn = app_container.get_connection()
        
        # 1. ดึงข้อมูล Log
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            cursor.execute("SELECT so_ids_json, sale_key FROM commission_payout_logs WHERE id = %s", (payout_id,))
            log_data = cursor.fetchone()

        if not log_data:
            messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูล Payout Log ID: {payout_id}", parent=parent_window)
            return
            
        sale_key = log_data['sale_key']
        try:
            so_id_list = json.loads(log_data['so_ids_json'])
            so_ids_in_log = tuple(so_id_list)
        except:
            so_ids_in_log = ()

        if not so_ids_in_log:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบรายการ SO ในรอบการจ่ายนี้", parent=parent_window)
            return

        # 2. [แก้ไข Query] ดึงตัวแปรที่ต้องใช้คำนวณมาให้ครบ
        placeholders = ', '.join(['%s'] * len(so_ids_in_log))
        query = f"""
            SELECT 
                so_number, 
                sales_service_amount,   -- ยอดขายสินค้า (ตัวตั้ง)
                final_sales_amount,     -- ยอดขายสุทธิ
                final_cost_amount,      -- ต้นทุน
                cost_multiplier,        -- ตัวคูณต้นทุน
                difference_amount       -- ส่วนต่าง
            FROM commissions
            WHERE id IN ({placeholders})
        """
        df = pd.read_sql_query(query, app_container.pg_engine, params=so_ids_in_log)
        
        if df.empty:
            messagebox.showwarning("เตือน", "ไม่พบข้อมูล SO ในฐานข้อมูล", parent=parent_window)
            return

        # 3. [เพิ่ม Logic การคำนวณ] แบบเดียวกับ PayoutDetailWindow เป๊ะๆ
        
        # เตรียมตัวแปร (กันค่าว่าง)
        sales_product_only = pd.to_numeric(df['sales_service_amount'], errors='coerce').fillna(0)
        final_cost = pd.to_numeric(df['final_cost_amount'], errors='coerce').fillna(0)
        multiplier = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
        diff_amt = pd.to_numeric(df['difference_amount'], errors='coerce').fillna(0)
        
        # คำนวณกำไร (Profit) สูตร: (ยอดขายสินค้า - (ต้นทุน * ตัวคูณ)) + ส่วนต่าง
        profit = (sales_product_only - (final_cost * multiplier)) + diff_amt
        # ต้นทุนรวมที่ใช้เทียบ (หลังคูณตัวคูณแล้ว) — ใช้เป็นตัวหารของ Mark up %
        total_cost_used = final_cost * multiplier

        # คำนวณ Gross Margin % — [ตาม PM] เปลี่ยนชื่อจาก "Margin" เป็น "Gross Margin" (สูตรเดิมไม่เปลี่ยน)
        # สูตร: (Profit / ยอดขายสินค้า) * 100
        df['calculated_margin'] = (profit / sales_product_only.replace(0, np.nan)) * 100
        df['calculated_margin'] = df['calculated_margin'].fillna(0.0)

        # [ตาม PM] เพิ่ม Mark up % = (Profit / ต้นทุนรวม) * 100 — คนละสูตรกับ Gross Margin (ตัวหารเป็นต้นทุน ไม่ใช่ยอดขาย)
        df['calculated_markup'] = (profit / total_cost_used.replace(0, np.nan)) * 100
        df['calculated_markup'] = df['calculated_markup'].fillna(0.0)

        # กำหนดสถานะ (Status)
        df['status'] = df['calculated_margin'].apply(lambda x: 'Normal' if x >= 10.0 else 'Below Tier')

        # 4. จัดเตรียมคอลัมน์สำหรับ Export
        header_map = {}
        if hasattr(app_container, 'HEADER_MAP'):
             header_map = app_container.HEADER_MAP

        # Rename ให้สวยงาม
        df.rename(columns={
            'so_number': header_map.get('so_number', 'SO Number'),
            'sales_service_amount': 'ยอดขายสินค้า (Base)',
            'final_sales_amount': 'ยอดขายรวมสุทธิ (Final)',
            'calculated_margin': 'Gross Margin ที่คำนวณ (%)',
            'calculated_markup': 'Mark up ที่คำนวณ (%)',
            'status': 'สถานะ (Status)'
        }, inplace=True)

        # เลือกเฉพาะคอลัมน์ที่จำเป็น
        export_cols = [
            header_map.get('so_number', 'SO Number'),
            'สถานะ (Status)',
            'ยอดขายสินค้า (Base)',
            'Gross Margin ที่คำนวณ (%)',
            'Mark up ที่คำนวณ (%)',
            'ยอดขายรวมสุทธิ (Final)'
        ]
        
        # กรองเอาเฉพาะที่มีอยู่จริงใน df
        final_cols = [c for c in export_cols if c in df.columns]
        df_export = df[final_cols]

        # 5. ถามที่บันทึกไฟล์
        default_filename = f"Payout_{payout_id}_{sale_key}_SOs_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="บันทึกรายการ SO ที่จ่ายแล้ว",
            initialfile=default_filename,
            parent=parent_window
        )

        # 6. บันทึก
        if save_path:
            df_export.to_excel(save_path, index=False)
            messagebox.showinfo("สำเร็จ", f"Export รายการ SO สำเร็จ! (ตรงกับหน้าจอ)\nบันทึกที่: {save_path}", parent=parent_window)

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=parent_window)
        traceback.print_exc()
    finally:
        if conn:
            app_container.release_connection(conn)