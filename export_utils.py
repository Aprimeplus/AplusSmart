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
    Export ข้อมูล Purchase Orders (PO) ที่มีสถานะ 'Approved' ออกเป็นไฟล์ Excel
    โดยมีการกรองข้อมูลตามช่วงเวลา
    """
    # --- START: เรียกใช้หน้าต่างเลือกเวลา ---
    dialog = DateRangeDialog(parent_window)
    parent_window.wait_window(dialog) # รอจนกว่าหน้าต่างเลือกเวลาจะปิด

    start_date = dialog.start_date
    end_date = dialog.end_date

    if not start_date or not end_date:
        print("Export canceled by user.")
        return # ผู้ใช้กดยกเลิก
    # --- END: เรียกใช้หน้าต่างเลือกเวลา ---

    try:
        # --- START: แก้ไข Query เพื่อเพิ่มเงื่อนไขวันที่ ---
        query = """
            SELECT 
                so.so_number,
                po.po_number,
                po.po_mode,
                po.rr_number,
                po.timestamp AS po_date,
                po.supplier_name,
                po.grand_total AS po_total_payable,
                po.user_key AS po_creator_key,
                so.bill_date AS so_date,
                so.customer_name,
                u.sale_name,
                (COALESCE(so.sales_service_amount, 0) + COALESCE(so.product_vat_7, 0)) AS so_grand_total
            FROM 
                purchase_orders po
            LEFT JOIN 
                commissions so ON po.so_number = so.so_number
            LEFT JOIN 
                sales_users u ON so.sale_key = u.sale_key
            WHERE 
                po.status = 'Approved'
                AND po.timestamp BETWEEN %s AND %s
            ORDER BY
                po.timestamp DESC;
        """
        # --- END: แก้ไข Query เพื่อเพิ่มเงื่อนไขวันที่ ---
        
        df = pd.read_sql_query(query, pg_engine, params=(start_date, end_date))
        
        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูล PO ที่อนุมัติแล้วในช่วงเวลาที่เลือก", parent=parent_window)
            return

        header_map = {
            'so_number': 'เลขที่ SO',
            'po_number': 'เลขที่ PO',
            'po_mode': 'ประเภท PO',
            'rr_number': 'เลขที่ RR', 
            'po_date': 'วันที่สร้าง PO',
            'supplier_name': 'ชื่อซัพพลายเออร์', 
            'po_total_payable': 'ยอดชำระ PO (บาท)',
            'po_creator_key': 'รหัสผู้สร้าง PO', 
            'so_date': 'วันที่เปิด SO', 
            'customer_name': 'ชื่อลูกค้า',
            'sale_name': 'พนักงานขาย', 
            'so_grand_total': 'ยอดขาย SO (บาท)'
        }
        df.rename(columns=header_map, inplace=True)
        
        ordered_columns = [
            'เลขที่ SO', 'เลขที่ PO', 'ประเภท PO', 'เลขที่ RR',
            'วันที่สร้าง PO', 'ชื่อซัพพลายเออร์', 'ยอดชำระ PO (บาท)',
            'รหัสผู้สร้าง PO', 'วันที่เปิด SO', 'ชื่อลูกค้า',
            'พนักงานขาย', 'ยอดขาย SO (บาท)'
        ]
        df = df[ordered_columns]
        
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
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้วที่:\n{save_path}", parent=parent_window)

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=parent_window)
        traceback.print_exc()