# export_utils.py (ฉบับแก้ไขล่าสุด: เพิ่มการกรองข้อมูลตามวันที่)

import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import traceback

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