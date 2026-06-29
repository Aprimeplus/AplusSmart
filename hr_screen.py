# hr_screen.py (ฉบับแก้ไขและปรับปรุง)

import tkinter as tk
from tkinter import ttk
from customtkinter import (CTkFrame, CTkLabel, CTkEntry, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkTabview, filedialog,
                           CTkInputDialog, CTkOptionMenu, CTkCheckBox, CTkTextbox, CTkComboBox, CTkRadioButton, CTkToplevel)
from tkinter import messagebox, TclError
import pandas as pd
import psycopg2
import psycopg2.errors
from psycopg2.extras import DictCursor, execute_values
import numpy as np
from datetime import datetime, timedelta
import calendar
import chardet
import json
import bcrypt
import traceback
import os
import shutil
from tkinter import font as tkfont
from outstanding_dashboard_tab import OutstandingDashboardTab
# --- START: แก้ไขการ Import และลงทะเบียนฟอนต์ ---
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.font_manager import fontManager
from export_utils import export_commission_details_to_excel
import matplotlib
from cancellation_dialog import CancellationReasonDialog
from hr_windows import HRCoverSheetDialog



try:
    # ใช้ os.path.join เพื่อให้ทำงานได้ทุกระบบปฏิบัติการ
    font_path = os.path.join('resources', 'THSarabunNew.ttf')
    if os.path.exists(font_path):
        fontManager.addfont(font_path)
        # ตั้งค่าให้ Matplotlib รู้จักและใช้ฟอนต์นี้เป็นหลัก
        matplotlib.rc('font', family='TH Sarabun New')
    else:
        print("Warning: ไม่พบฟอนต์ THSarabunNew.ttf ในโฟลเดอร์ resources")
except Exception as e:
    print(f"Error registering font for matplotlib: {e}")
# --- END ---

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import FuncFormatter, MaxNLocator

from sqlalchemy import create_engine

# (import ส่วนที่เหลือของโปรแกรม)
from hr_windows import HRVerificationWindow, PayoutDetailWindow, PayoutCalculationViewer, SOPopupWindow, CalculationDetailViewer
from history_windows import PurchaseDetailWindow, CancelledHistoryWindow, STATUS_THAI_MAP
from custom_widgets import NumericEntry, DateSelector
import utils
import business_logic
# --- DIALOG CLASSES ---

# วางไว้ต่อจาก import หรือกลุ่ม Class Dialog อื่นๆ (เช่น ManualEntryDialog)

class DateRangeSelectionDialog(CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("กำหนดช่วงเวลา")
        self.geometry("350x250")
        self.grab_set()
        self.transient(master)
        
        self.start_date = None
        self.end_date = None
        self.confirmed = False

        # ใช้ DateSelector ที่มีอยู่ใน custom_widgets
        CTkLabel(self, text="วันที่เริ่มต้น:", font=master.label_font).pack(pady=(20, 5))
        self.start_picker = DateSelector(self)
        self.start_picker.pack(pady=5)

        CTkLabel(self, text="วันที่สิ้นสุด:", font=master.label_font).pack(pady=(10, 5))
        self.end_picker = DateSelector(self)
        self.end_picker.pack(pady=5)
        
        # ตั้งค่าเริ่มต้นเป็นวันนี้
        self.start_picker.set_date(datetime.now())
        self.end_picker.set_date(datetime.now())

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)
        CTkButton(btn_frame, text="ตกลง", command=self._on_confirm, fg_color="#16A34A").pack(side="left", padx=10)
        CTkButton(btn_frame, text="ยกเลิก", command=self.destroy, fg_color="gray").pack(side="left", padx=10)
        
        utils.center_window(self)

    def _on_confirm(self):
        self.start_date = self.start_picker.get_date()
        self.end_date = self.end_picker.get_date()
        
        if self.start_date > self.end_date:
            messagebox.showerror("ผิดพลาด", "วันที่เริ่มต้นต้องมาก่อนวันที่สิ้นสุด", parent=self)
            return
            
        self.confirmed = True
        self.destroy()

class SalesFilterDialog(CTkToplevel):
    def __init__(self, master, sales_list, current_selection, on_confirm):
        super().__init__(master)
        self.title("กรองพนักงานขาย")
        self.geometry("350x500")
        self.sales_list = sales_list  # [('SALE01', 'Name'), ...]
        self.on_confirm = on_confirm
        self.checkbox_vars = {}
        
        # แปลง current_selection (list) เป็น set เพื่อค้นหาเร็ว
        selected_set = set(current_selection) if current_selection else set([s[0] for s in sales_list])

        # --- Header & Buttons ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.pack(fill="x", padx=10, pady=10)
        
        CTkButton(top_frame, text="เลือกทั้งหมด", width=100, 
                  command=self.select_all, fg_color="#3B82F6").pack(side="left", padx=5)
        CTkButton(top_frame, text="ล้างทั้งหมด", width=100, 
                  command=self.deselect_all, fg_color="gray").pack(side="left", padx=5)

        # --- Checkbox List ---
        self.scroll_frame = CTkScrollableFrame(self, label_text="รายชื่อพนักงาน")
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

        for sale_key, sale_name in sales_list:
            var = tk.IntVar(value=1 if sale_key in selected_set else 0)
            self.checkbox_vars[sale_key] = var
            text_label = f"{sale_key} : {sale_name}"
            cb = CTkCheckBox(self.scroll_frame, text=text_label, variable=var)
            cb.pack(anchor="w", padx=10, pady=5)

        # --- Confirm Button ---
        CTkButton(self, text="ตกลง", command=self._confirm_selection, 
                  fg_color="#16A34A", font=("Arial", 16, "bold")).pack(fill="x", padx=20, pady=20)
        
        # Center Window
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        self.transient(master)
        self.grab_set()

    def select_all(self):
        for var in self.checkbox_vars.values(): var.set(1)

    def deselect_all(self):
        for var in self.checkbox_vars.values(): var.set(0)

    def _confirm_selection(self):
        selected_keys = [k for k, v in self.checkbox_vars.items() if v.get() == 1]
        self.on_confirm(selected_keys)
        self.destroy()

class ManualEntryDialog(CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("เพิ่มข้อมูลด้วยมือ")
        self.geometry("400x200")
        self.grab_set()
        self.transient(master)
        self.result = None

        main_frame = CTkFrame(self, fg_color="transparent")
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)
        main_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(main_frame, text="เลขที่ SO:").grid(row=0, column=0, sticky="w", pady=5)
        self.so_entry = CTkEntry(main_frame)
        self.so_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        CTkLabel(main_frame, text="ยอดขาย:").grid(row=1, column=0, sticky="w", pady=5)
        self.sales_entry = NumericEntry(main_frame)
        self.sales_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        CTkLabel(main_frame, text="ต้นทุน:").grid(row=2, column=0, sticky="w", pady=5)
        self.cost_entry = NumericEntry(main_frame)
        self.cost_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=10)
        CTkButton(button_frame, text="เพิ่มรายการ", command=self._on_add).pack(side="left", padx=10)
        CTkButton(button_frame, text="ยกเลิก", command=self.destroy, fg_color="gray").pack(side="left", padx=10)
        
        self.so_entry.focus_set()

    def _on_add(self):
        so = self.so_entry.get().strip()
        sales = self.sales_entry.get().strip()
        cost = self.cost_entry.get().strip()

        if not so or not sales or not cost:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลให้ครบทุกช่อง", parent=self)
            return
        
        try:
            float(sales.replace(",", ""))
            float(cost.replace(",", ""))
        except ValueError:
            messagebox.showerror("ข้อมูลผิดพลาด", "ยอดขายและต้นทุนต้องเป็นตัวเลข", parent=self)
            return

        self.result = {'so_number': so, 'sales_uploaded': sales, 'cost_uploaded': cost}
        self.destroy()

class ComparisonConfigDialog(CTkToplevel):
    def __init__(self, master, sales_keys):
        super().__init__(master)
        self.title("ตั้งค่าการเปรียบเทียบข้อมูล")
        self.geometry("500x550") # ปรับความสูงให้เหมาะสม
        self.grab_set()
        self.transient(master)

        self.result = None
        self.imported_df = None
        self.manual_df = pd.DataFrame(columns=['so_number', 'sales_uploaded', 'cost_uploaded'])

        self.grid_columnconfigure(0, weight=1)
        ### 4. ปรับ grid_rowconfigure ให้ถูกต้อง ###
        self.grid_rowconfigure(3, weight=1) # ให้ส่วนแสดงผล manual ขยายได้

        ### 1. แก้ไขลำดับการจัดวาง (grid row) และเลขลำดับ Label ให้ถูกต้อง ###
        
        # --- Section 1: เลือกรอบค่าคอม ---
        period_frame = CTkFrame(self)
        period_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        CTkLabel(period_frame, text="1. เลือกรอบค่าคอม:", font=master.label_font).pack(side="left", padx=10)

        current_time = datetime.now()
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.month_var = tk.StringVar(value=self.thai_months[current_time.month - 1])
        CTkOptionMenu(period_frame, variable=self.month_var, values=self.thai_months, command=self._check_run_button_state).pack(side="left", padx=5)

        year_options = [str(y + 543) for y in range(current_time.year - 2, current_time.year + 2)]
        self.year_var = tk.StringVar(value=str(current_time.year + 543))
        CTkOptionMenu(period_frame, variable=self.year_var, values=year_options, command=self._check_run_button_state).pack(side="left", padx=5)

        # --- Section 2: เลือกพนักงานขาย ---
        sales_frame = CTkFrame(self)
        sales_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        CTkLabel(sales_frame, text="2. เลือกพนักงานขาย:", font=master.label_font).pack(side="left", padx=10)

        ### 2. ลบโค้ดที่ซ้ำซ้อนและขัดแย้งกันออก เหลือแค่ส่วนนี้ ###
        self.placeholder = "กรุณาเลือกพนักงานขาย..."
        self.selected_sale = tk.StringVar(value=self.placeholder)
        self.sale_dropdown = CTkOptionMenu(sales_frame, variable=self.selected_sale, values=[self.placeholder] + sales_keys, command=self._check_run_button_state)
        self.sale_dropdown.pack(side="left", padx=10, pady=10)

        # --- Section 3: เลือกแหล่งข้อมูล ---
        source_frame = CTkFrame(self)
        source_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        source_frame.grid_columnconfigure(1, weight=1)
        CTkLabel(source_frame, text="3. เลือกแหล่งข้อมูล (อย่างน้อย 1 อย่าง):", font=master.label_font).grid(row=0, column=0, columnspan=2, sticky="w", padx=10)
        
        self.import_button = CTkButton(source_frame, text="นำเข้าไฟล์ Excel/CSV", command=self._on_import_file)
        self.import_button.grid(row=1, column=0, padx=10, pady=10)
        self.file_label = CTkLabel(source_frame, text="ยังไม่ได้เลือกไฟล์", text_color="gray", anchor="w")
        self.file_label.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.manual_button = CTkButton(source_frame, text="เพิ่มข้อมูลด้วยมือ...", command=self._on_add_manual)
        self.manual_button.grid(row=2, column=0, padx=10, pady=10)

        # --- Section 4: แสดงผล Manual Entry ---
        manual_display_frame = CTkScrollableFrame(self, label_text="รายการที่คีย์ด้วยมือ")
        manual_display_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        self.manual_display_label = CTkLabel(manual_display_frame, text="ไม่มีข้อมูล")
        self.manual_display_label.pack(pady=10)

        # --- Section 5: ปุ่ม Action ---
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=4, column=0, pady=20)
        self.run_button = CTkButton(button_frame, text="เริ่มการเปรียบเทียบ", command=self._on_run_comparison, state="disabled")
        self.run_button.pack(side="left", padx=10)
        CTkButton(button_frame, text="ยกเลิก", command=self.destroy, fg_color="gray").pack(side="left", padx=10)

        utils.center_window(self)

    def _check_run_button_state(self, *args):
        has_data = (self.imported_df is not None) or (not self.manual_df.empty)
        sale_selected = self.selected_sale.get() != self.placeholder
        
        # --- เพิ่มเงื่อนไขตรวจสอบเดือน/ปี ---
        period_selected = self.month_var.get() and self.year_var.get()
        
        # ปุ่มจะทำงานได้ก็ต่อเมื่อมีข้อมูล, เลือกเซลส์แล้ว และเลือกรอบเวลาแล้ว
        if has_data and sale_selected and period_selected:
            self.run_button.configure(state="normal")
        else:
            self.run_button.configure(state="disabled")

    def _on_import_file(self):
        file_path = filedialog.askopenfilename(title="เลือกไฟล์ Excel/CSV", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not file_path:
            return
        
        SO_ALIASES = {'so number', 'so_number', 'so no.', 'เลขที่ so', 'อ้างถึง'}

        def _find_header_row(raw_df):
            """สแกนหา row ที่มี SO column — คืน index ของ header row (หรือ None)"""
            for i, row in raw_df.iterrows():
                vals = {str(v).lower().strip() for v in row.values if pd.notna(v)}
                if vals & SO_ALIASES:
                    return i
            return None

        try:
            if file_path.endswith('.csv'):
                with open(file_path, 'rb') as f: result = chardet.detect(f.read())
                df = pd.read_csv(file_path, encoding=result['encoding'])
            else:
                raw = pd.read_excel(file_path, header=None)
                header_row = _find_header_row(raw)
                if header_row is None:
                    # ลองอ่านแบบปกติ (header=0) ถ้าหา SO row ไม่เจอ
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_excel(file_path, header=header_row)

            self.imported_df = df
            self.file_label.configure(text=os.path.basename(file_path), text_color="green")
            self._check_run_button_state()
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถอ่านไฟล์ได้: {e}", parent=self)
            self.imported_df = None
            self.file_label.configure(text="การนำเข้าล้มเหลว", text_color="red")
            self._check_run_button_state()

    def _on_add_manual(self):
        dialog = ManualEntryDialog(self)
        self.wait_window(dialog)

        # ย้ายทุกอย่างที่ใช้ new_row เข้ามาในนี้
        if dialog.result:
            new_entry = {
                'so_number': dialog.result['so_number'],
                'sales_uploaded': float(dialog.result['sales_uploaded'].replace(",", "")),
                'cost_uploaded': float(dialog.result['cost_uploaded'].replace(",", ""))
            }
            new_row = pd.DataFrame([new_entry])
            
            # <<< START: ย้ายโค้ดส่วนนี้เข้ามาข้างใน >>>
            # แก้ไขปัญหานี้และ Future Warning ไปพร้อมกัน
            if self.manual_df.empty:
                self.manual_df = new_row
            else:
                self.manual_df = pd.concat([self.manual_df, new_row], ignore_index=True)
            
            # ซ่อน Label "ไม่มีข้อมูล" ถ้ามีข้อมูลแถวแรก
            if len(self.manual_df) == 1:
                self.manual_display_label.pack_forget()
            
            # แสดงข้อมูลที่เพิ่มเข้ามา
            entry_text = f"SO: {new_entry['so_number']}, Sales: {new_entry['sales_uploaded']:,.2f}, Cost: {new_entry['cost_uploaded']:,.2f}"
            CTkLabel(self.manual_display_label.master, text=entry_text).pack(anchor="w", padx=10)
            
            # เปิดใช้งานปุ่ม "เริ่มการเปรียบเทียบ"
            self._check_run_button_state()

    def _on_run_comparison(self):
        thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        selected_month_num = thai_month_map.get(self.month_var.get())
        selected_year_ad = int(self.year_var.get()) - 543

        self.result = {
            "salesperson": self.selected_sale.get(),
            "month": selected_month_num, # <-- เพิ่ม
            "year": selected_year_ad,      # <-- เพิ่ม
            "imported_df": self.imported_df,
            "manual_df": self.manual_df
        }
        self.destroy()

class AnnualArchiveDialog(CTkToplevel):
    def __init__(self, master, current_year):
        super().__init__(master)
        self.title("บันทึกประจำปี")
        self.geometry("450x300")
        self.grab_set()
        self.transient(master)

        self.archive_mode = tk.StringVar(value="annual")
        self.selected_month = tk.StringVar(value="")
        self.selected_year = tk.StringVar(value=str(current_year - 1))

        thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(thai_months)}
        year_list = [str(y) for y in range(current_year - 5, current_year + 1)]

        CTkLabel(self, text="เลือกโหมดการบันทึกประจำปี:", font=CTkFont(size=16, weight="bold")).pack(pady=10)

        mode_frame = CTkFrame(self, fg_color="transparent")
        mode_frame.pack(pady=5)
        CTkRadioButton(mode_frame, text="บันทึกทั้งปี (ไฟล์รวม)", variable=self.archive_mode, value="annual", command=self._toggle_month_selector).pack(anchor="w", pady=2)
        CTkRadioButton(mode_frame, text="บันทึกรายเดือน (เลือกเดือน)", variable=self.archive_mode, value="monthly", command=self._toggle_month_selector).pack(anchor="w", pady=2)
        # --- START: เพิ่มตัวเลือกใหม่ ---
        CTkRadioButton(mode_frame, text="บันทึกทั้งปี (แยกไฟล์รายเดือน)", variable=self.archive_mode, value="annual_by_month", command=self._toggle_month_selector).pack(anchor="w", pady=2)

        year_frame = CTkFrame(self, fg_color="transparent")
        year_frame.pack(pady=5)
        CTkLabel(year_frame, text="ปีที่ต้องการบันทึก:").pack(side="left", padx=5)
        self.year_menu = CTkOptionMenu(year_frame, variable=self.selected_year, values=year_list)
        self.year_menu.pack(side="left", padx=5)

        self.month_frame = CTkFrame(self, fg_color="transparent")
        CTkLabel(self.month_frame, text="เดือนที่ต้องการบันทึก:").pack(side="left", padx=5)
        self.month_menu = CTkOptionMenu(self.month_frame, variable=self.selected_month, values=thai_months)
        self.month_menu.pack(side="left", padx=5)
        
        self.selected_month.set(thai_months[datetime.now().month - 1])

        self._toggle_month_selector()

        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=20)
        CTkButton(button_frame, text="ตกลง", command=self._on_confirm).pack(side="left", padx=10)
        CTkButton(button_frame, text="ยกเลิก", command=self._on_cancel, fg_color="gray").pack(side="left", padx=10)

        self.result = None

    def _toggle_month_selector(self):
    # จะแสดง Dropdown ก็ต่อเมื่อเลือกโหมด "รายเดือน" เท่านั้น
        if self.archive_mode.get() == "monthly":
            self.month_frame.pack(pady=5)
        else:
            self.month_frame.pack_forget()

    def _on_confirm(self):
        mode = self.archive_mode.get()
        year = self.selected_year.get()
        month_num = None
        if mode == "monthly":
            month_name = self.selected_month.get()
            if not month_name:
                messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกเดือนที่ต้องการบันทึก", parent=self)
                return
            month_num = self.thai_month_map.get(month_name)
            if not month_num:
                messagebox.showerror("ผิดพลาด", "ไม่สามารถระบุเดือนที่เลือกได้", parent=self)
                return
        self.result = {"mode": mode, "year": int(year), "month": month_num}
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()

class PayoutConfirmDialog(CTkToplevel):
    """
    Custom confirmation dialog สำหรับยืนยันการจ่ายคอมมิชชั่น (Paid)
    แสดงรายละเอียด SO งวดปัจจุบัน + SO ตกหล่นจากงวดก่อน พร้อมยอดเงิน
    """
    THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม",
                   "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม",
                   "พฤศจิกายน", "ธันวาคม"]

    def __init__(self, master, period_text, val_gross, val_wht, val_net,
                 comm_df, selected_year, selected_month):
        super().__init__(master)
        self.result = False
        self.title("ยืนยันการจ่ายคอมมิชชั่น")
        self.resizable(False, True)
        self.grab_set()
        self.transient(master)

        # ---- แยก SO งวดปัจจุบัน vs ตกหล่น ----
        self._current_rows = []
        self._old_rows = []
        if comm_df is not None and not comm_df.empty:
            df = comm_df.copy()
            df['commission_year']  = pd.to_numeric(df['commission_year'],  errors='coerce').fillna(0).astype(int)
            df['commission_month'] = pd.to_numeric(df['commission_month'], errors='coerce').fillna(0).astype(int)
            if 'final_sales_amount' not in df.columns:
                df['final_sales_amount'] = 0.0
            df['final_sales_amount'] = pd.to_numeric(df['final_sales_amount'], errors='coerce').fillna(0.0)

            cur_mask = (df['commission_year'] == selected_year) & (df['commission_month'] == selected_month)
            for _, r in df[cur_mask].iterrows():
                self._current_rows.append({'so': str(r['so_number']), 'amount': float(r['final_sales_amount'])})
            for _, r in df[~cur_mask].iterrows():
                mo = int(r['commission_month'])
                yr = int(r['commission_year'])
                mo_str = self.THAI_MONTHS[mo - 1] if 1 <= mo <= 12 else f"เดือน {mo}"
                yr_be  = yr + 543
                self._old_rows.append({'so': str(r['so_number']), 'amount': float(r['final_sales_amount']),
                                       'period': f"{mo_str} {yr_be}"})

        # ---- สร้าง UI ----
        W = 560
        self._build_ui(period_text, val_gross, val_wht, val_net)
        self.update_idletasks()
        H = min(self.winfo_reqheight() + 20, 680)
        self._center(master, W, H)
        self.geometry(f"{W}x{H}")

    # ------------------------------------------------------------------
    def _build_ui(self, period_text, val_gross, val_wht, val_net):
        # ── HEADER ──────────────────────────────────────────────────────
        hdr = CTkFrame(self, fg_color="#1E3A5F", corner_radius=0)
        hdr.pack(fill="x")
        CTkLabel(hdr, text="💰  ยืนยันการจ่ายคอมมิชชั่น",
                 font=CTkFont(size=16, weight="bold"),
                 text_color="white").pack(pady=(12, 2))
        CTkLabel(hdr, text=f"งวด: {period_text}",
                 font=CTkFont(size=13),
                 text_color="#A8D0F0").pack(pady=(0, 12))

        body = CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # ── สรุปยอดเงิน ─────────────────────────────────────────────────
        fin_box = CTkFrame(body, fg_color="#F0F4FF", corner_radius=10)
        fin_box.pack(fill="x", pady=(0, 10))
        fin_box.grid_columnconfigure(1, weight=1)

        def fin_row(parent, row_idx, label, value_text, color="#1F1F1F", bold=False):
            wt = "bold" if bold else "normal"
            CTkLabel(parent, text=label,  font=CTkFont(size=13, weight=wt),
                     text_color=color).grid(row=row_idx, column=0, sticky="w", padx=14, pady=3)
            CTkLabel(parent, text=value_text, font=CTkFont(size=13, weight=wt),
                     text_color=color).grid(row=row_idx, column=1, sticky="e", padx=14, pady=3)

        fin_row(fin_box, 0, "ยอดคอมมิชชั่นขั้นต้น (Gross)", f"{val_gross:,.2f} บาท")
        fin_row(fin_box, 1, "ภาษีหัก ณ ที่จ่าย (3%)",        f"-{val_wht:,.2f} บาท", color="#DC2626")

        sep = CTkFrame(fin_box, height=1, fg_color="#CBD5E1")
        sep.grid(row=2, column=0, columnspan=2, sticky="ew", padx=14, pady=4)

        fin_row(fin_box, 3, "✅  ยอดโอนสุทธิ (Net)",
                f"{val_net:,.2f} บาท", color="#16A34A", bold=True)

        # ── รายการ SO ───────────────────────────────────────────────────
        total = len(self._current_rows) + len(self._old_rows)
        CTkLabel(body,
                 text=f"📦  รายการ SO ทั้งหมด ({total} รายการ)",
                 font=CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(4, 4))

        so_scroll = CTkScrollableFrame(body, height=230, fg_color="#F8F9FA", corner_radius=8)
        so_scroll.pack(fill="x")

        # งวดปัจจุบัน
        if self._current_rows:
            CTkLabel(so_scroll,
                     text=f"✅  งวดปัจจุบัน ({len(self._current_rows)} รายการ)",
                     font=CTkFont(size=12, weight="bold"),
                     text_color="#16A34A").pack(anchor="w", padx=10, pady=(6, 2))
            for row in self._current_rows:
                CTkLabel(so_scroll,
                         text=f"   • {row['so']}   ยอดขาย: {row['amount']:,.0f} บาท",
                         font=CTkFont(size=12),
                         text_color="#374151").pack(anchor="w", padx=10, pady=1)

        # ตกหล่น
        if self._old_rows:
            CTkLabel(so_scroll,
                     text=f"⚠️  ตกหล่นจากงวดก่อน ({len(self._old_rows)} รายการ)",
                     font=CTkFont(size=12, weight="bold"),
                     text_color="#B45309").pack(anchor="w", padx=10, pady=(8, 2))
            for row in self._old_rows:
                CTkLabel(so_scroll,
                         text=f"   • {row['so']}   [{row['period']}]   ยอดขาย: {row['amount']:,.0f} บาท",
                         font=CTkFont(size=12),
                         text_color="#92400E").pack(anchor="w", padx=10, pady=1)

        if not self._current_rows and not self._old_rows:
            CTkLabel(so_scroll, text="(ไม่มีข้อมูล SO)",
                     text_color="gray").pack(pady=10)

        # ── ปุ่ม ────────────────────────────────────────────────────────
        btn_frame = CTkFrame(body, fg_color="transparent")
        btn_frame.pack(pady=(12, 4))
        CTkButton(btn_frame, text="❌  ยกเลิก",
                  width=140, height=38,
                  fg_color="#6B7280", hover_color="#4B5563",
                  command=self._cancel).pack(side="left", padx=8)
        CTkButton(btn_frame, text="✅  ยืนยันการจ่าย",
                  width=160, height=38,
                  fg_color="#16A34A", hover_color="#15803D",
                  command=self._confirm).pack(side="left", padx=8)

    # ------------------------------------------------------------------
    def _center(self, master, w, h):
        try:
            self.update_idletasks()
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            mw = master.winfo_width()
            mh = master.winfo_height()
            x = mx + (mw - w) // 2
            y = my + (mh - h) // 2
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _confirm(self):
        self.result = True
        self.grab_release()
        self.destroy()

    def _cancel(self):
        self.result = False
        self.grab_release()
        self.destroy()


class HRScreen(CTkFrame):
    
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master, corner_radius=0, fg_color=app_container.THEME["hr"]["bg"])
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.user_key = user_key
        self.user_name = user_name

        # ── Hardcode: user เหล่านี้ได้สิทธิ์เทียบเท่า Director ────────────
        # (เห็น column ตัวคูณ + แก้ไข cost_multiplier per-SO ได้)
        _DIRECTOR_LEVEL_USERS = {"LEK", "YUPIN"}
        self.user_role = 'Director' if (user_key in _DIRECTOR_LEVEL_USERS) else user_role

        # --- Fonts ---
        self.label_font = CTkFont(size=16, weight="bold", family="Roboto")
        self.entry_font = CTkFont(size=14, family="Roboto")
        self.header_font_table = CTkFont(size=14, weight="bold", family="Roboto")
        self.label_font_bold = CTkFont(size=12, weight="bold", family="Roboto")
        self.small_font = CTkFont(size=12, family="Roboto")

        self.header_map = app_container.HEADER_MAP
        self.sales_keys_list = self._get_sale_keys()

        # --- Variables ---
        self.db_df, self.uploaded_df, self.comparison_df, self.user_df, self.comparison_log_df = None, None, None, None, None
        self.initial_commission_result = None
        self.current_comm_df = None
        self.manual_entry_df = pd.DataFrame(columns=['so_number', 'sales_uploaded', 'cost_uploaded'])
        self.uploaded_file_path, self.sales_chart_canvas, self.po_chart_canvas, self.sales_target_chart_canvas = None, None, None, None
        self.selected_payout_ids = set(); self.select_all_var = tk.IntVar(value=0)
        self.theme = self.app_container.THEME["hr"]
        
        self.dropdown_style = {
            "fg_color": "white",
            "text_color": "black",
            "button_color": self.theme.get("primary", "#3B82F6"),
            "button_hover_color": self.theme.get("header", "#2563EB")
        }
        
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.period_options = ["ปีนี้", "เดือนนี้", "Q1", "Q2", "Q3", "Q4"] + self.thai_months + ["กำหนดช่วงเวลาเอง..."]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}

        # Pagination vars
        self.history_current_page, self.history_rows_per_page, self.history_total_rows = 0, 20, 0
        self.user_current_page, self.user_rows_per_page, self.user_total_rows = 0, 20, 0
        self.edit_data_current_page = 0
        self.edit_data_rows_per_page = 15

        # --- Layout Setup ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 0))
        CTkLabel(header_frame, text=f"หน้าจอสำหรับฝ่ายบุคคล (HR): {self.user_name}", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).pack(side="left")
        
        # --- [🔥 แก้ไข] เพิ่มปุ่มพิมพ์ใบปะหน้าไว้ข้างๆ ปุ่มออกจากระบบ ---
        from hr_windows import HRCoverSheetDialog  # Import Class ใหม่ที่สร้างไว้

        button_container = CTkFrame(header_frame, fg_color="transparent")
        button_container.pack(side="right")

        # ปุ่มพิมพ์ใบปะหน้า
        CTkButton(
            button_container, 
            text="🖨️ พิมพ์ใบปะหน้า (ค้นหา)", 
            command=lambda: HRCoverSheetDialog(self, self.app_container),
            fg_color="#7C3AED", # สีม่วง
            hover_color="#6D28D9",
            width=140
        ).pack(side="left", padx=(0, 10))

        # ปุ่มออกจากระบบ
        CTkButton(
            button_container, 
            text="ออกจากระบบ", 
            command=self.app_container.show_login_screen, 
            fg_color="transparent", 
            border_color="#D32F2F", 
            text_color="#D32F2F", 
            border_width=2, 
            hover_color="#FFEBEE"
        ).pack(side="left")
        # -------------------------------------------------------------

        # ========================================================================================
        # สร้าง Main TabView (แท็บแม่) แบ่งหมวดหมู่
        # ========================================================================================
        self.main_tab_view = CTkTabview(self, corner_radius=10, border_width=0, width=200, fg_color=self.cget("fg_color"))
        self.main_tab_view.grid(row=1, column=0, pady=10, padx=20, sticky="nsew")
        
        # เพิ่มหมวดหมู่หลัก (Parent Tabs)
        self.cat_analysis = self.main_tab_view.add("📊 วิเคราะห์ (Analysis)")
        self.cat_management = self.main_tab_view.add("⚙️ จัดการ (Management)")
        self.cat_entry = self.main_tab_view.add("⌨️ คีย์แทน (Data Entry)")
        self.cat_commission = self.main_tab_view.add("💰 ค่าคอมมิชชั่น (Commission)")

        # กำหนด Grid ให้แต่ละหมวดหมู่
        for tab_name in ["📊 วิเคราะห์ (Analysis)", "⚙️ จัดการ (Management)", "⌨️ คีย์แทน (Data Entry)", "💰 ค่าคอมมิชชั่น (Commission)"]:
            self.main_tab_view.tab(tab_name).grid_columnconfigure(0, weight=1)
            self.main_tab_view.tab(tab_name).grid_rowconfigure(0, weight=1)

        # ------------------------------------------------------------------
        # 1. หมวดวิเคราะห์ (Analytics) -> สร้าง Sub-Tabs
        # ------------------------------------------------------------------
        self.analysis_tabs = CTkTabview(self.cat_analysis, corner_radius=10, command=self._on_tab_selected)
        self.analysis_tabs.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        self.dashboard_tab = self.analysis_tabs.add("ภาพรวม (Dashboard)")
        self.sales_target_tab = self.analysis_tabs.add("เป้าการขาย")
        self.outstanding_tab = self.analysis_tabs.add("ยอดค้างชำระ")

        self._create_dashboard_tab(self.dashboard_tab)
        self._create_sales_target_tab(self.sales_target_tab)
        self.outstanding_dashboard = OutstandingDashboardTab(self.outstanding_tab, self.app_container)

        # ------------------------------------------------------------------
        # 2. หมวดจัดการ (Management) -> สร้าง Sub-Tabs
        # ------------------------------------------------------------------
        self.management_tabs = CTkTabview(self.cat_management, corner_radius=10, command=self._on_tab_selected)
        self.management_tabs.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        self.manage_users_tab = self.management_tabs.add("ผู้ใช้งาน")
        self.edit_data_tab = self.management_tabs.add("แก้ไขข้อมูล (Master Edit)")
        self.audit_log_tab = self.management_tabs.add("บันทึกระบบ (Log)")
        self.cancelled_so_tab = self.management_tabs.add("จัดการ SO ยกเลิก")
        self._create_cancelled_so_tab(self.cancelled_so_tab)

        self._create_manage_users_tab(self.manage_users_tab)
        self._create_edit_data_tab(self.edit_data_tab)
        self._create_audit_log_tab(self.audit_log_tab)

        # ------------------------------------------------------------------
        # 3. หมวดคีย์แทน (Data Entry) -> สร้าง Sub-Tabs
        # ------------------------------------------------------------------
        self.entry_tabs = CTkTabview(self.cat_entry, corner_radius=10, command=self._on_tab_selected)
        self.entry_tabs.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        self.sales_mode_tab = self.entry_tabs.add("แทนเซลส์")
        self.sales_mode_tab.grid_columnconfigure(0, weight=1); self.sales_mode_tab.grid_rowconfigure(0, weight=1)
        
        self.pu_mode_tab = self.entry_tabs.add("แทนจัดซื้อ")
        self.pu_mode_tab.grid_columnconfigure(0, weight=1); self.pu_mode_tab.grid_rowconfigure(0, weight=1)

        # ------------------------------------------------------------------
        # 4. หมวดค่าคอมมิชชั่น (Commission) -> สร้าง Sub-Tabs
        # ------------------------------------------------------------------
        self.commission_tabs = CTkTabview(self.cat_commission, corner_radius=10, command=self._on_tab_selected)
        self.commission_tabs.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        self.compare_commission_tab = self.commission_tabs.add("1. ตรวจสอบ (Verify)")
        self.process_commission_tab = self.commission_tabs.add("2. คำนวณ & จ่าย")
        self.payout_history_tab = self.commission_tabs.add("3. ประวัติการจ่าย")

        self._create_compare_commission_tab(self.compare_commission_tab)
        self._create_process_commission_tab(self.process_commission_tab)
        self._create_payout_history_tab(self.payout_history_tab)
        
        # ------------------------------------------------------------------
        
        # Load Status Flags
        self.after(100, self._initial_load)
        self._sales_mode_loaded = False 
        self._pu_mode_loaded = False 
        self._payout_history_loaded = False 
        self._dashboard_loaded, self._sales_target_loaded, self._users_loaded, self._compare_commission_loaded, self._process_commission_loaded, self._audit_log_loaded = False, False, False, False, False, False

    def _get_special_service_amounts(self, so_ids):
        """
        [NEW] ดึงยอดขาย (SO) และต้นทุน (PO) ของสินค้ากลุ่มพิเศษ
        (แก้ไข: นำ EXP-0194 ออก เพื่อให้คิดเป็นสินค้าปกติ)
        """
        if not so_ids:
            return pd.DataFrame()

        # 1. กำหนดรหัสสินค้ากลุ่มพิเศษ (Hardcode)
        
        # Cutting/Drilling Codes (คงเดิม)
        cutting_codes = ["'EXP-0079'", "'EXP-0128'"]
        
        # Other Service Codes (❌ เอา 'EXP-0194' ออกจากบรรทัดนี้ครับ)
        service_codes = ["'EXP-0006'", "'EXP-0049'", "'EXP-0077'", "'EXP-0174'"]

        # แปลง List so_ids เป็น String สำหรับ Query IN (...)
        so_ids_str = ', '.join(map(str, so_ids))

        # 2. เขียน SQL Query — match by both product_name patterns AND product_code
        sql = f"""
        WITH po_items AS (
            SELECT
                c.id AS comm_id,
                COALESCE(SUM(CASE
                    WHEN poi.product_name LIKE '%%ค่าตัด%%' OR poi.product_name LIKE '%%เจาะ%%'
                         OR COALESCE(poi.product_code, '') IN ('EXP-0079', 'EXP-0128')
                    THEN poi.total_price ELSE 0 END), 0) as po_cutting_cost,
                COALESCE(SUM(CASE
                    WHEN poi.product_name LIKE '%%ค่าบริการ%%'
                         OR COALESCE(poi.product_code, '') IN ('EXP-0006', 'EXP-0049', 'EXP-0077', 'EXP-0174')
                    THEN poi.total_price ELSE 0 END), 0) as po_service_cost
            FROM commissions c
            JOIN purchase_orders po ON c.so_number = po.so_number
            JOIN purchase_order_items poi ON po.id = poi.purchase_order_id
            WHERE c.id IN ({so_ids_str}) AND po.status = 'Approved'
            GROUP BY c.id
        )
        SELECT 
            c.id,
            c.so_number,
            COALESCE(c.cutting_drilling_fee, 0) as so_cutting_rev,
            COALESCE(c.other_service_fee, 0) as so_service_rev,
            COALESCE(p.po_cutting_cost, 0) as po_cutting_cost,
            COALESCE(p.po_service_cost, 0) as po_service_cost
        FROM commissions c
        LEFT JOIN po_items p ON c.id = p.comm_id
        WHERE c.id IN ({so_ids_str})
        """
        
        try:
            return pd.read_sql_query(sql, self.app_container.pg_engine)
        except Exception as e:
            print(f"Error fetching special services: {e}")
            return pd.DataFrame()
            

    def _cancel_so_logic(self, so_number, reason):
        """Logic การยกเลิก SO + PO + Noti + Log"""
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # 1. ดึงข้อมูล SO เพื่อหาเจ้าของ (Sale Key)
                cursor.execute("SELECT id, sale_key FROM commissions WHERE so_number = %s", (so_number,))
                result = cursor.fetchone()
                if not result:
                    messagebox.showerror("Error", "ไม่พบ SO นี้ในระบบ")
                    return
                so_id, sale_key = result

                # 2. อัปเดต SO เป็น Cancelled (และปิด Active ไม่ให้คำนวณคอมฯ)
                # เพิ่ม rejection_reason เพื่อเก็บสาเหตุ
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'Cancelled', 
                        is_active = 0, 
                        rejection_reason = %s 
                    WHERE so_number = %s
                """, (f"ยกเลิกโดย {self.user_role}: {reason}", so_number))

                # 3. อัปเดต PO ที่เกี่ยวข้องทั้งหมดเป็น Cancelled
                cursor.execute("""
                    UPDATE purchase_orders 
                    SET status = 'Cancelled', 
                        approval_status = 'Cancelled' 
                    WHERE so_number = %s
                """, (so_number,))

                # 4. ส่ง Notification หาเจ้าของ SO
                noti_msg = f"SO: {so_number} ถูกยกเลิกโดย {self.user_role}\nสาเหตุ: {reason}\n(รายการนี้จะไม่ถูกนำไปคิดค่าคอมมิชชั่น)"
                cursor.execute("""
                    INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id, timestamp)
                    VALUES (%s, %s, FALSE, %s, NOW())
                """, (sale_key, noti_msg, so_id))

                # 4b. แจ้งเตือน Sale Manager ทุกคน
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Sales Manager' AND status = 'Active'")
                manager_keys = [r[0] for r in cursor.fetchall()]
                manager_msg = (f"[HR_CANCEL] SO: {so_number} ถูกยกเลิกโดย HR\n"
                               f"เจ้าของ SO: {sale_key}\n"
                               f"สาเหตุ: {reason}")
                for mgr_key in manager_keys:
                    cursor.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id, timestamp)
                        VALUES (%s, %s, FALSE, %s, NOW())
                    """, (mgr_key, manager_msg, so_id))

                # 5. บันทึก Audit Log
                import json
                log_data = json.dumps({"reason": reason, "cancelled_by": self.user_role})
                cursor.execute("""
                    INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                    VALUES (%s, %s, %s, %s, %s, NOW())
                """, ('Cancel SO', 'commissions', so_id, self.user_name, log_data))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ยกเลิก SO: {so_number} เรียบร้อยแล้ว")
            
            # TODO: Refresh หน้าจอหลังจากทำเสร็จ

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    # วิธีเรียกใช้ (ผูกกับปุ่ม)
    def on_click_cancel_button(self):
        # สมมติได้ so_number มาจากการเลือกในตาราง
        selected_so = "SO-xxxx" 
        
        # เปิด Dialog
        from cancellation_dialog import CancellationReasonDialog
        CancellationReasonDialog(self, lambda reason: self._cancel_so_logic(selected_so, reason))

    def _open_sales_filter_dialog(self):
        # เตรียมรายชื่อ (Key, Name)
        sales_list = []
        if hasattr(self, 'sales_user_info'):
            for key, info in self.sales_user_info.items():
                sales_list.append((key, info.get('name', key)))
        else:
            # Fallback ถ้ายังไม่มีข้อมูล
            sales_list = [(k, k) for k in self.sales_keys_list]
        
        # เรียงตามชื่อ
        sales_list.sort(key=lambda x: x[1])

        # เปิด Dialog
        SalesFilterDialog(self, sales_list, self.selected_sales_filter, self._on_filter_confirmed)

    def _on_filter_confirmed(self, selected_keys):
        self.selected_sales_filter = selected_keys
        
        # อัปเดตข้อความบนปุ่ม
        count = len(selected_keys)
        total = len(self.sales_keys_list) if hasattr(self, 'sales_keys_list') else 0
        if count == total or total == 0:
            self.filter_btn.configure(text="👤 กรองพนักงาน (ทั้งหมด)")
            self.selected_sales_filter = None # Reset เป็น None เพื่อความง่าย
        else:
            self.filter_btn.configure(text=f"👤 กรองพนักงาน ({count} คน)")
        
        # รีเฟรชกราฟทันที
        self._on_target_filter_search()

    def _on_target_filter_search(self):
        """ค้นหาตามรอบเดือนค่าคอม (เดือน/ปี เท่านั้น — ไม่ใช้วันที่แน่นอน)"""
        try:
            import calendar

            def get_month_year(m_var, y_var):
                m = self.thai_months.index(m_var.get()) + 1
                y = int(y_var.get()) - 543  # พ.ศ. → ค.ศ.
                return m, y

            s_m, s_y = get_month_year(self.start_m_var, self.start_y_var)
            e_m, e_y = get_month_year(self.end_m_var,   self.end_y_var)

            # start = วันที่ 1 ของรอบเริ่ม
            start_date = datetime(s_y, s_m, 1)
            # end   = วันสุดท้ายของรอบสิ้นสุด (ครอบคลุมทั้งเดือน)
            last_day = calendar.monthrange(e_y, e_m)[1]
            end_date = datetime(e_y, e_m, last_day)

            if start_date > end_date:
                messagebox.showerror(
                    "รอบเดือนไม่ถูกต้อง",
                    "รอบเริ่มต้นต้องมาก่อนหรือเท่ากับรอบสิ้นสุด",
                    parent=self
                )
                return

            # บันทึกให้ _get_sales_vs_target_data ใช้
            self.custom_target_start = start_date
            self.custom_target_end   = end_date

            # target_multiplier = จำนวนเดือนในช่วง
            months_diff = (e_y - s_y) * 12 + (e_m - s_m) + 1

            self.sales_target_period_var.set("กำหนดช่วงเวลาเอง...")
            self._update_sales_target_dashboard()

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()

    def _on_sales_target_period_change(self, selected_period):
        if selected_period == "กำหนดช่วงเวลาเอง...":
            # เรียกใช้ Dialog ที่สร้างไว้
            dialog = DateRangeSelectionDialog(self)
            self.wait_window(dialog)
            
            if dialog.confirmed:
                # บันทึกวันที่เลือกไว้ในตัวแปร
                self.custom_target_start = dialog.start_date
                self.custom_target_end = dialog.end_date
                # อัปเดตกราฟ
                self._update_sales_target_dashboard()
            else:
                # ถ้ากดยกเลิก ให้กลับไปเป็นค่าเดิม (เช่น เดือนนี้)
                self.sales_target_period_var.set("เดือนนี้")
        else:
            # กรณีเลือกปกติ (Q1, เดือน, ปี)
            self.custom_target_start = None
            self.custom_target_end = None
            self._update_sales_target_dashboard()

    def _create_payout_history_table(self, df):
        """(เวอร์ชันแก้ไขสมบูรณ์) สร้างตารางประวัติการจ่ายเงิน พร้อมแก้ปัญหาคลิกไม่ติดและค่า 0.00"""
        
        # 1. ล้างข้อมูลเก่าใน Frame
        for widget in self.payout_history_frame.winfo_children():
            widget.destroy()

        if df is None or df.empty:
            CTkLabel(self.payout_history_frame, text="ไม่พบข้อมูลตามเงื่อนไขที่เลือก").pack(pady=20)
            return

        # 2. ตั้งค่า Style ของตาราง
        style = ttk.Style(self.payout_history_frame)
        style.theme_use("clam")
        
        style.configure("Payout.Treeview.Heading", 
                            font=self.label_font_bold, 
                            background="#065F46",
                            foreground="white", 
                            relief="flat", 
                            padding=(10, 8))
        style.map("Payout.Treeview.Heading", background=[('active', "#047857")])
        style.configure("Payout.Treeview", 
                            rowheight=32, 
                            font=self.small_font,
                            fieldbackground="#F9FAFB",
                            foreground="#1F2937")
        style.map("Payout.Treeview", 
                    background=[('selected', self.theme["primary"])], 
                    foreground=[('selected', 'white')])

        # 3. สร้าง Frame สำหรับวาง Treeview
        tree_frame = CTkFrame(self.payout_history_frame, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # กำหนดชื่อคอลัมน์และหัวข้อ
        columns = {
            'payout_period_text': 'รอบค่าคอม',
            'sale_key': 'รหัสพนักงาน', 'sale_name': 'ชื่อพนักงาน', 'plan_name': 'แผน',
            'sales_target': 'เป้าหมาย', 
            'avg_margin': 'Avg Margin (%)',   # <<< ✅ เพิ่มคอลัมน์นี้ตรงนี้
            'total_sales': 'ยอดขาย',
            'total_normal_sales': 'Normal',
            'total_below_sales': 'BelowT',
            'timestamp': 'วันที่ทำรายการจ่าย',
            'final_commission': 'ยอดคอม Gross',
            'incentives_total': 'Incentive', 'deductions_total': 'ยอดหัก', 
            'withholding_tax': 'หัก 3%', 'net_commission': 'ยอดโอนสุทธิ'
        }
        
        tree = ttk.Treeview(tree_frame, columns=list(columns.keys()), show='headings', style="Payout.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        # ตั้งค่าสีสลับแถว
        tree.tag_configure('oddrow', background='#FFFFFF')
        tree.tag_configure('evenrow', background='#F0F9FF')

        # 4. ตั้งค่าความกว้างและการจัดวางคอลัมน์
        for col_id, col_text in columns.items():
            anchor = 'w'
            width = 120 # ค่าเริ่มต้น

            if col_id == 'sale_name': 
                width = 200
            elif col_id == 'timestamp': 
                width = 110
            elif col_id == 'plan_name': 
                width = 80
            elif col_id == 'payout_period_text':
                width = 130
            
            # คอลัมน์ตัวเลขให้ชิดขวา
            elif col_id in ['sales_target', 'total_sales', 'total_normal_sales', 'total_below_sales',
                'final_commission', 'incentives_total', 'deductions_total', 
                'withholding_tax', 'net_commission']:
                anchor = 'e'
                width = 130
                if col_id in ['total_normal_sales', 'total_below_sales']:
                    width = 100 

            # <<< ✅ จุดที่เพิ่ม >>>
            elif col_id == 'avg_margin':
                anchor = 'e'
                width = 110
            
            tree.heading(col_id, text=col_text, anchor='center')
            tree.column(col_id, anchor=anchor, width=width, minwidth=60)

        # 5. วนลูปใส่ข้อมูล (Data Population)
        for i, row in df.iterrows():
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'

            values = []
            for col_id in columns.keys():
                value = row[col_id]
                
                # รายชื่อคอลัมน์ที่เป็นตัวเลขเงิน (ต้องแปลง Null เป็น 0.00 เสมอ)
                money_cols = [
                    'sales_target', 'total_sales', 'total_normal_sales', 'total_below_sales', 
                    'final_commission', 'incentives_total', 'deductions_total', 
                    'withholding_tax', 'net_commission'
                ]

                # <<< ✅ จัดเรียง if-elif ให้ถูกต้อง >>>
                if col_id == 'avg_margin':
                    try:
                        values.append(f"{float(value):.2f}%")
                    except (ValueError, TypeError):
                        values.append("0.00%")
                
                elif col_id in money_cols:
                    try:
                        # พยายามแปลงเป็น float (รองรับทั้ง None, "", และ string ที่มี comma)
                        if pd.isna(value) or str(value).strip() == "":
                            float_val = 0.0
                        else:
                            float_val = float(str(value).replace(',', ''))
                        values.append(f"{float_val:,.2f}")
                    except ValueError:
                        # ถ้าแปลงไม่ได้จริงๆ ให้เป็น 0.00
                        values.append("0.00")
                
                # ถ้าไม่ใช่คอลัมน์เงิน แต่มีค่า (ไม่ว่าง)
                elif pd.notna(value) and str(value).strip() != "":
                    if isinstance(value, datetime): 
                        values.append(value.strftime('%d/%m/%Y'))
                    elif isinstance(value, (float, np.floating, int)): 
                        values.append(f"{value:,.2f}")
                    else: 
                        values.append(str(value))
                else:
                    # ค่าว่างสำหรับคอลัมน์ทั่วไป
                    values.append("")
            
            # สำคัญ: iid ต้องเป็น ID ของ record ใน Database เพื่อให้คลิกแล้วดึงข้อมูลถูก
            tree.insert("", "end", values=values, iid=str(row['id']), tags=(tag,))
        
        # 6. Bind Event (จุดสำคัญที่แก้ไข!)
        # ใช้แค่ Double-1 เท่านั้น เพื่อเปิด Popup (ห้ามใส่ Button-1 ซ้อน)
        tree.bind("<Double-1>", lambda e: self._on_payout_history_double_click(e, tree))

        # 7. สร้าง Scrollbar
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        v_scroll.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=v_scroll.set)
        
    def _navigate_payout_month(self, direction):
        """ฟังก์ชันอัจฉริยะสำหรับคำนวณเดือนเดินหน้า/ถอยหลัง"""
        current_month_str = self.payout_month_var.get()
        current_year_str = self.payout_year_var.get()

        # ถ้าเลือก "ทุกเดือน" หรือ "ทุกปี" อยู่ ให้ตั้งต้นที่เดือน/ปี ปัจจุบัน
        if current_month_str == "ทุกเดือน" or current_year_str == "ทุกปี":
            now = datetime.now()
            m = now.month
            y = now.year
        else:
            m = self.thai_month_map[current_month_str]
            y = int(current_year_str)

        # คำนวณเลื่อนเดือน
        if direction == "prev":
            m -= 1
            if m < 1:
                m = 12
                y -= 1
        elif direction == "next":
            m += 1
            if m > 12:
                m = 1
                y += 1

        # อัปเดตค่ากลับไปที่ Dropdown
        self.payout_month_var.set(self.thai_months[m - 1])
        self.payout_year_var.set(str(y))

        # โหลดข้อมูลใหม่ทันที
        self._load_payout_history()

    def _payout_prev_page(self):
        self._navigate_payout_month("prev")

    def _payout_next_page(self):
        self._navigate_payout_month("next")

    def _create_edit_data_tab(self, parent_tab):
        """สร้าง UI สำหรับหน้า Master Edit SO/PO"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        # --- Frame สำหรับการค้นหา ---
        search_frame = CTkFrame(parent_tab)
        search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        search_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(search_frame, text="ค้นหา SO/PO Number:", font=self.label_font).grid(row=0, column=0, padx=10, pady=10)
        
        self.master_edit_search_entry = CTkEntry(search_frame, font=self.entry_font, placeholder_text="กรอก SO หรือ PO ที่ต้องการค้นหา...")
        self.master_edit_search_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        search_button = CTkButton(search_frame, text="ค้นหา", command=self._search_so_po)
        search_button.grid(row=0, column=2, padx=10, pady=10)
        
        # --- Frame สำหรับแสดงผลการค้นหา ---
        self.master_edit_results_frame = CTkScrollableFrame(parent_tab, label_text="ผลการค้นหา")
        self.master_edit_results_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.master_edit_results_frame.grid_columnconfigure(0, weight=1)

    def _search_so_po(self):
        """(เวอร์ชันแก้ไข) ค้นหาข้อมูล SO และ PO เฉพาะตาม Keyword และกรองตามเซลส์"""
        for widget in self.master_edit_results_frame.winfo_children():
            widget.destroy()
            
        keyword = self.master_edit_search_entry.get().strip().upper()
        if not keyword:
            self._load_hr_edit_queue()
            return

        search_term = keyword
        if search_term.startswith("SO"): search_term = search_term[2:]
        elif search_term.startswith("PO"): search_term = search_term[2:]
        
        # --- START: เพิ่ม Logic การกรองตามเซลส์ ---
        selected_sale = self.master_edit_sale_var.get()
        sale_filter_so = ""
        sale_filter_po = ""
        params_so = [f"%{search_term}%"]
        params_po = [f"%{search_term}%"]

        if selected_sale != "ทั้งหมด":
            sale_filter_so = " AND sale_key = %s"
            params_so.append(selected_sale)
            
            # สำหรับ PO ต้อง Join เพื่อหา sale_key จาก commissions
            sale_filter_po = " AND p.so_number IN (SELECT so_number FROM commissions WHERE sale_key = %s)"
            params_po.append(selected_sale)
        # --- END ---
        
        try:
            so_query = f"SELECT id, so_number, customer_name, sale_key FROM commissions WHERE so_number ILIKE %s AND is_active = 1 {sale_filter_so}"
            po_query = f"SELECT p.id, p.so_number, p.po_number, p.supplier_name FROM purchase_orders p WHERE p.po_number ILIKE %s {sale_filter_po}"

            so_df = pd.read_sql_query(so_query, self.pg_engine, params=tuple(params_so))
            po_df = pd.read_sql_query(po_query, self.pg_engine, params=tuple(params_po))

            if so_df.empty and po_df.empty:
                CTkLabel(self.master_edit_results_frame, text=f"ไม่พบข้อมูลสำหรับ '{keyword}' ของเซลส์ '{selected_sale}'").pack(pady=20)
                return

            if not so_df.empty:
                CTkLabel(self.master_edit_results_frame, text="ผลการค้นหา: Sales Orders (SO)", font=self.label_font).pack(anchor="w", padx=10, pady=(10,0))
                for _, row in so_df.iterrows():
                    self._create_so_card_for_editing(self.master_edit_results_frame, row.to_dict())

            if not po_df.empty:
                CTkLabel(self.master_edit_results_frame, text="ผลการค้นหา: Purchase Orders (PO)", font=self.label_font).pack(anchor="w", padx=10, pady=(10,0))
                for _, row in po_df.iterrows():
                    po_id = int(row['id'])
                    po_card = CTkFrame(self.master_edit_results_frame, border_width=1)
                    po_card.pack(fill="x", padx=10, pady=5)
                    info = f"PO: {row['po_number']} | SO: {row['so_number']} | Supplier: {row['supplier_name']}"
                    CTkLabel(po_card, text=info).pack(side="left", padx=10, pady=5)
                    CTkButton(po_card, text="แก้ไข PO", command=lambda pid=po_id: self._open_po_editor_for_hr(pid)).pack(side="right", padx=10, pady=5)

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการค้นหา: {e}", parent=self)

    def _open_so_editor_for_hr(self, so_id):
        """เปิดหน้าต่างแก้ไข SO สำหรับ HR (ฉบับแก้ไข เพิ่ม StringVars ที่ขาดไป)"""
        try:
            so_df = pd.read_sql_query("SELECT * FROM commissions WHERE id = %s", self.pg_engine, params=(so_id,))
            if so_df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูล SO ID: {so_id}", parent=self)
                return
            
            # เราต้องสร้าง StringVars จำลองที่ SOPopupWindow ต้องการ
            so_shared_vars = {}
            so_shared_vars['delivery_type_var'] = tk.StringVar()
            so_shared_vars['sales_service_vat_option'] = tk.StringVar()
            so_shared_vars['cutting_drilling_fee_vat_option'] = tk.StringVar()
            so_shared_vars['other_service_fee_vat_option'] = tk.StringVar()
            so_shared_vars['shipping_vat_option_var'] = tk.StringVar()
            so_shared_vars['credit_card_fee_vat_option_var'] = tk.StringVar()
            so_shared_vars['so_grand_total_var'] = tk.StringVar()
            so_shared_vars['so_vs_payment_result_var'] = tk.StringVar()
            so_shared_vars['difference_amount_var'] = tk.StringVar()
            so_shared_vars['cash_required_total_var'] = tk.StringVar()
            so_shared_vars['cash_verification_result_var'] = tk.StringVar()
            so_shared_vars['credit_term_var'] = tk.StringVar(value="เงินสด")
            
            
            # --- START: เพิ่ม StringVars สำหรับแสดง VAT ที่ขาดไป ---
            so_shared_vars['sales_vat_calc_var'] = tk.StringVar(value="0.00")
            so_shared_vars['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
            so_shared_vars['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
            so_shared_vars['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
            so_shared_vars['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
            so_shared_vars['relocation_vat_option_var'] = tk.StringVar(value="VAT")
            so_shared_vars['relocation_vat_calc_var'] = tk.StringVar(value="0.00")
            # --- END ---

            # เรียกใช้ SOPopupWindow จาก hr_windows.py
            SOPopupWindow(
                master=self,
                app_container=self.app_container,
                sales_data=so_df.iloc[0].to_dict(),
                so_shared_vars=so_shared_vars,
                sale_theme=self.app_container.THEME["sale"],
                on_save_callback=self._search_so_po
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างแก้ไข SO ได้: {e}", parent=self)
            traceback.print_exc()
    
    def _open_po_editor_for_hr(self, po_id):
        """เปิดหน้าต่างแก้ไข PO สำหรับ HR"""
        try:
            # ตรวจสอบว่ามี PO ID นี้อยู่จริงหรือไม่
            po_df = pd.read_sql_query("SELECT id FROM purchase_orders WHERE id = %s", self.pg_engine, params=(po_id,))
            if po_df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูล PO ID: {po_id}", parent=self)
                return

            # เรียกใช้ PurchaseDetailWindow จาก history_windows.py
            # และส่ง callback function ไปด้วยเพื่อให้หน้าจอ Refresh หลังบันทึก
            PurchaseDetailWindow(
                master=self,
                app_container=self.app_container,
                purchase_id=int(po_id),
                on_save_callback=self._search_so_po # ใช้ฟังก์ชันค้นหาเพื่อโหลดข้อมูลใหม่
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างแก้ไข PO ได้: {e}", parent=self)
            traceback.print_exc()

    def _initial_load_process_commission(self):
        """
        ฟังก์ชันสำหรับโหลดข้อมูลเริ่มต้นของแท็บ 'ประมวลผลค่าคอม'
        เมื่อถูกเรียกครั้งแรก
        """
        # เรียกใช้ฟังก์ชันที่มีอยู่แล้วซึ่งทำหน้าที่โหลดข้อมูลของเซลส์ที่เลือก
        self._on_sale_selected_for_process()

    def _on_tab_selected(self):
        """
        ตรวจสอบว่า Tab ไหนถูกเลือก (ทั้ง Tab หลักและ Tab ย่อย) และโหลดข้อมูลตามความเหมาะสม
        """
        main_tab = self.main_tab_view.get()
        selected_sub_tab = ""
        
        if main_tab == "📊 วิเคราะห์ (Analysis)":
            selected_sub_tab = self.analysis_tabs.get()
            if selected_sub_tab == "ภาพรวม (Dashboard)" and not self._dashboard_loaded:
                self._initial_load_dashboard(); self._dashboard_loaded = True
            elif selected_sub_tab == "เป้าการขาย" and not self._sales_target_loaded:
                self._initial_load_sales_target(); self._sales_target_loaded = True
            
            if selected_sub_tab == "ภาพรวม (Dashboard)": self._update_dashboard()
            elif selected_sub_tab == "เป้าการขาย": self._update_sales_target_dashboard()

        elif main_tab == "⚙️ จัดการ (Management)":
            selected_sub_tab = self.management_tabs.get()
            if selected_sub_tab == "ผู้ใช้งาน" and not self._users_loaded:
                self._populate_users_table(); self._users_loaded = True
            elif selected_sub_tab == "บันทึกระบบ (Log)" and not self._audit_log_loaded:
                self._populate_audit_log_table(); self._audit_log_loaded = True
            elif selected_sub_tab == "จัดการ SO ยกเลิก":
                self._load_cancelled_so_history()
            
            if selected_sub_tab == "ผู้ใช้งาน": self._populate_users_table()
            elif selected_sub_tab == "บันทึกระบบ (Log)": self._populate_audit_log_table()
            elif selected_sub_tab == "จัดการ SO ยกเลิก": self._load_cancelled_so_history()

        elif main_tab == "⌨️ คีย์แทน (Data Entry)":
            selected_sub_tab = self.entry_tabs.get()
            if selected_sub_tab == "แทนเซลส์" and not self._sales_mode_loaded:
                try:
                    from sales_proxy_screen import SalesProxyScreen
                    self.sales_proxy_screen_instance = SalesProxyScreen(
                        master=self.sales_mode_tab, app_container=self.app_container,
                        proxy_user_key=self.user_key, proxy_user_name=self.user_name,
                        user_role=self.user_role, role_to_proxy="Sale"
                    )
                    self.sales_proxy_screen_instance.grid(row=0, column=0, sticky="nsew")
                    self._sales_mode_loaded = True
                except Exception as e: messagebox.showerror("Error", f"Load Sales Mode failed: {e}")
            
            elif selected_sub_tab == "แทนจัดซื้อ" and not self._pu_mode_loaded:
                try:
                    from purchasing_proxy_screen import PurchasingProxyScreen
                    self.pu_proxy_screen_instance = PurchasingProxyScreen(
                        master=self.pu_mode_tab, app_container=self.app_container,
                        proxy_user_key=self.user_key, proxy_user_name=self.user_name,
                        role_to_proxy="Purchasing Staff"
                    )
                    self.pu_proxy_screen_instance.pack(fill="both", expand=True)
                    self._pu_mode_loaded = True
                except Exception as e: messagebox.showerror("Error", f"Load PU Mode failed: {e}")

        elif main_tab == "💰 ค่าคอมมิชชั่น (Commission)":
            selected_sub_tab = self.commission_tabs.get()
            
            if selected_sub_tab == "1. ตรวจสอบ (Verify)" and not self._compare_commission_loaded:
                self._compare_commission_loaded = True
            elif selected_sub_tab == "2. คำนวณ & จ่าย" and not self._process_commission_loaded:
                self._initial_load_process_commission(); self._process_commission_loaded = True
            elif selected_sub_tab == "3. ประวัติการจ่าย" and not self._payout_history_loaded:
                self._load_payout_history(); self._payout_history_loaded = True
            
            # Refresh
            if selected_sub_tab == "2. คำนวณ & จ่าย": 
                # 🔥 [จุดแก้ที่ 1] บังคับให้ Dropdown ชื่อเซลส์ เปลี่ยนเป็นคนที่เพิ่งตรวจเสร็จ
                if hasattr(self, 'last_verified_sale') and self.last_verified_sale in self.active_sales_keys:
                    self.selected_sale_for_process.set(self.last_verified_sale)
                self._on_sale_selected_for_process()
                
            elif selected_sub_tab == "3. ประวัติการจ่าย": self._load_payout_history()

    def _show_calculation_details(self):
        """แสดงรายละเอียดการคำนวณในหน้าต่างใหม่ (ฉบับกันตาย เปิดได้ 100%)"""
        
        if not hasattr(self, 'latest_commission_result') or not self.latest_commission_result:
            messagebox.showinfo("ไม่มีข้อมูล", "กรุณากด 'คำนวณขั้นสุดท้าย' ก่อนดูรายละเอียด", parent=self)
            return

        # 1. ดึงข้อมูลที่ได้จากการคำนวณออกมา
        debug_data = self.latest_commission_result.get('debug_df')
        breakdown_data = self.latest_commission_result.get('so_breakdown_df')
        
        # 2. บังคับแปลงข้อมูลขั้นตอน (Debug) ให้เป็น DataFrame เสมอ (แก้ปัญหา Error List)
        if isinstance(debug_data, pd.DataFrame) and not debug_data.empty:
            debug_df = debug_data
        elif isinstance(debug_data, list) and len(debug_data) > 0:
            debug_df = pd.DataFrame(debug_data)
        else:
            # ถ้าไม่มีมาให้เลย ให้สร้างข้อความจำลอง
            debug_df = pd.DataFrame([{"รายการ": "ข้อมูลขั้นตอนการคำนวณ", "ค่า": "กรุณาดูรายละเอียดในแท็บ SO Breakdown"}])

        # 3. บังคับแปลงข้อมูลราย SO (Breakdown) ให้เป็น DataFrame เสมอ
        if isinstance(breakdown_data, pd.DataFrame) and not breakdown_data.empty:
            so_breakdown_df = breakdown_data
        elif isinstance(breakdown_data, list) and len(breakdown_data) > 0:
            so_breakdown_df = pd.DataFrame(breakdown_data)
        else:
            # 🔥 ไม้ตาย: ถ้าใน logic ลืมส่งตารางกลับมา ให้ดึงข้อมูลดิบ (current_comm_df) ไปโชว์แทนเลย!
            if hasattr(self, 'current_comm_df') and not self.current_comm_df.empty:
                so_breakdown_df = self.current_comm_df.copy()
            else:
                so_breakdown_df = pd.DataFrame()

        # 4. ดึงชื่อแผนแล้วเรียกเปิด Popup
        sale_key = self.selected_sale_for_process.get()
        plan_name = self.sales_user_info.get(sale_key, {}).get('plan', 'Unknown Plan')
        
        try:
            from hr_windows import CalculationDetailViewer
            CalculationDetailViewer(
                master=self,
                debug_df=debug_df,
                so_breakdown_df=so_breakdown_df,
                plan_name=plan_name,
                comm_df=self.current_comm_df if hasattr(self, 'current_comm_df') else None,
                user_role=getattr(self, 'user_role', None),
                recalculate_callback=self._calculate_commission_for_period,
            )
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถเปิดหน้าต่างรายละเอียดได้: {e}", parent=self)
            traceback.print_exc()

    def _trial_export_data(self):
        """
        ฟังก์ชันสำหรับทดลอง Export ข้อมูลเป็น Excel เท่านั้น
        จะไม่มีการลบข้อมูลออกจากฐานข้อมูลอย่างเด็ดขาด
        """
        # 1. เปิดหน้าต่างถามปี/เดือน เหมือนเดิม
        dialog = AnnualArchiveDialog(self, datetime.now().year)
        self.wait_window(dialog)
        archive_config = dialog.result

        if archive_config is None:
            messagebox.showinfo("ยกเลิก", "การทดลอง Export ถูกยกเลิก", parent=self)
            return

        mode, year_to_archive, month_to_archive = archive_config["mode"], archive_config["year"], archive_config["month"]
        
        # แสดงข้อความว่ากำลังทำงาน
        loading_popup = CTkToplevel(self)
        loading_popup.geometry("300x100")
        loading_popup.title("โปรดรอ")
        loading_popup.transient(self)
        loading_popup.grab_set()
        CTkLabel(loading_popup, text="กำลัง Export ข้อมูลเป็น Excel...\nกรุณารอสักครู่", font=self.label_font).pack(expand=True)
        self.update_idletasks()

        try:
            # 2. เตรียมตำแหน่งและชื่อไฟล์
            archive_dir_base = os.path.join("archive", "annual_records", str(year_to_archive))
            os.makedirs(archive_dir_base, exist_ok=True)
            current_timestamp_for_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
            archive_suffix = f"_{year_to_archive}"

            if mode == "monthly":
                archive_suffix += f"_{month_to_archive:02d}"
                start_date_filter = datetime(year_to_archive, month_to_archive, 1, 0, 0, 0)
                end_date_filter = datetime(year_to_archive, month_to_archive, calendar.monthrange(year_to_archive, month_to_archive)[1], 23, 59, 59)
            else:
                start_date_filter = datetime(year_to_archive, 1, 1, 0, 0, 0)
                end_date_filter = datetime(year_to_archive, 12, 31, 23, 59, 59)
            
            start_date_str, end_date_str = start_date_filter.strftime('%Y-%m-%d %H:%M:%S'), end_date_filter.strftime('%Y-%m-%d %H:%M:%S')

            files_created = []

            # 3. Export ตาราง commissions
            comm_df = pd.read_sql_query(f"SELECT * FROM commissions WHERE timestamp BETWEEN %s AND %s", self.pg_engine, params=(start_date_str, end_date_str))
            if not comm_df.empty:
                comm_filename = f"TRIAL_commissions{archive_suffix}_{current_timestamp_for_filename}.xlsx"
                comm_path = os.path.join(archive_dir_base, comm_filename)
                comm_df.to_excel(comm_path, index=False)
                files_created.append(comm_filename)

            # 4. Export ตาราง purchase_orders
            po_df = pd.read_sql_query(f"SELECT * FROM purchase_orders WHERE timestamp BETWEEN %s AND %s", self.pg_engine, params=(start_date_str, end_date_str))
            if not po_df.empty:
                po_filename = f"TRIAL_purchase_orders{archive_suffix}_{current_timestamp_for_filename}.xlsx"
                po_path = os.path.join(archive_dir_base, po_filename)
                po_df.to_excel(po_path, index=False)
                files_created.append(po_filename)
            
            # *** ไม่มีการลบข้อมูลใดๆ ในฟังก์ชันนี้ ***

            loading_popup.destroy() # ปิดหน้าต่าง "โปรดรอ"

            # 5. แจ้งผลลัพธ์
            if not files_created:
                messagebox.showinfo("ไม่พบข้อมูล", "ไม่พบข้อมูลในช่วงเวลาที่เลือกสำหรับ Export", parent=self)
            else:
                file_list_str = "\n - ".join(files_created)
                success_message = (
                    "ทดลอง Export สำเร็จ!\n\n"
                    f"ไฟล์ถูกบันทึกที่โฟลเดอร์:\n{archive_dir_base}\n\n"
                    f"ไฟล์ที่สร้าง:\n - {file_list_str}\n\n"
                    "**ข้อมูลในระบบยังคงอยู่เหมือนเดิม ไม่มีการลบใดๆ เกิดขึ้น**"
                )
                messagebox.showinfo("สำเร็จ", success_message, parent=self)

        except Exception as e:
            loading_popup.destroy()
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดระหว่างการ Export: {e}\n{traceback.format_exc()}", parent=self)

    def _reset_payout_filters(self):
        """รีเซ็ตค่าในฟิลเตอร์และโหลดข้อมูลใหม่"""
        self.payout_month_var.set("ทุกเดือน")
        self.payout_year_var.set(str(datetime.now().year))
        self.payout_search_entry.delete(0, 'end')
        self.history_current_page = 0 # <-- สั่งให้กลับไปที่หน้าแรกเสมอ
        self._load_payout_history()

    def _open_comparison_history_window(self):
        from hr_windows import ComparisonHistoryWindow # Import ที่นี่เพื่อเลี่ยง Circular Import
        ComparisonHistoryWindow(master=self, app_container=self.app_container)

    def _create_payout_history_tab(self, parent_tab):
        """(เวอร์ชันปรับปรุง) สร้าง Layout สำหรับหน้าประวัติการจ่ายเงิน (เปลี่ยนเป็นเลื่อนเดือน)"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(2, weight=1) # แถวที่ 2 (ตาราง) จะขยายได้

        # --- Frame หลักสำหรับตัวกรองทั้งหมด ---
        filter_container = CTkFrame(parent_tab, fg_color="transparent")
        filter_container.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        filter_container.grid_columnconfigure(3, weight=1)

        # --- ตัวกรอง เดือน/ปี ---
        CTkLabel(filter_container, text="เลือกช่วงเวลา:").pack(side="left", padx=(5,2))
        
        month_options = ["ทุกเดือน"] + self.thai_months
        self.payout_month_var = tk.StringVar(value="ทุกเดือน")
        CTkOptionMenu(filter_container, variable=self.payout_month_var, values=month_options).pack(side="left", padx=5)

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        self.payout_year_var = tk.StringVar(value=str(current_year))
        CTkOptionMenu(filter_container, variable=self.payout_year_var, values=year_options).pack(side="left", padx=5)

        # --- ช่องค้นหา ---
        self.payout_search_entry = CTkEntry(filter_container, placeholder_text="ค้นหาจากรหัส หรือชื่อพนักงานขาย...")
        self.payout_search_entry.pack(side="left", padx=(20, 5), fill="x", expand=True)
        self.payout_search_entry.bind("<Return>", lambda e: self._load_payout_history())
        
        # --- ปุ่ม ---
        CTkButton(filter_container, text="ค้นหา", command=self._load_payout_history).pack(side="left", padx=(0, 5))
        CTkButton(filter_container, text="ล้างค่า", command=self._reset_payout_filters, fg_color="gray").pack(side="left")

        # --- Frame สำหรับเปลี่ยนเดือน (แทนที่ Pagination เดิม) ---
        pagination_frame = CTkFrame(parent_tab, fg_color="transparent")
        pagination_frame.grid(row=1, column=0, padx=10, pady=0, sticky="ew")

        # เปลี่ยนปุ่มเป็นเลื่อนเดือน (เอา state="disabled" ออก เพื่อให้กดได้ตลอด)
        self.payout_prev_button = CTkButton(pagination_frame, text="◀ เดือนก่อนหน้า", command=self._payout_prev_page, width=120)
        self.payout_prev_button.pack(side="left")
        
        self.payout_page_label = CTkLabel(pagination_frame, text="รอบบิล: ทุกเดือน", font=self.label_font_bold, text_color=self.theme["primary"])
        self.payout_page_label.pack(side="left", expand=True)
        
        self.payout_next_button = CTkButton(pagination_frame, text="เดือนถัดไป ▶", command=self._payout_next_page, width=120)
        self.payout_next_button.pack(side="right")
        
        # --- Frame สำหรับแสดงตาราง ---
        self.payout_history_frame = CTkFrame(parent_tab)
        self.payout_history_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.payout_history_frame.grid_columnconfigure(0, weight=1)
        self.payout_history_frame.grid_rowconfigure(0, weight=1)

    def _load_payout_history(self):
        """(เวอร์ชันแก้ไข) โหลดประวัติการจ่ายเงิน พร้อมดึงยอดขาย Normal/BelowT"""
        try:
            search_term = self.payout_search_entry.get().strip()
            selected_year = self.payout_year_var.get()
            selected_month = self.payout_month_var.get()

            params = []
            where_clauses = ["1=1"] 

            # <<< START: แก้ไข Query ตรงนี้ >>>
            base_query = """
                SELECT 
                    log.id, 
                    log.payout_period_text,
                    log.sale_key, 
                    u.sale_name, 
                    log.plan_name, 
                    u.sales_target,
                    -- คำนวณ Avg Margin สดๆ จากบิลทั้งหมดในรอบนี้ (เอาสัญลักษณ์เปอร์เซ็นต์ออกแล้ว)
                    (SELECT COALESCE(SUM(final_gp) / NULLIF(SUM(final_sales_amount), 0) * 100, 0) FROM commissions WHERE payout_id = log.id) AS avg_margin,
                    log.total_sales,
                    log.total_normal_sales,
                    log.total_below_sales,
                    log.timestamp,            
                    log.final_commission, 
                    log.incentives_total,
                    log.deductions_total, 
                    log.withholding_tax,
                    log.net_commission,
                    log.commission_year,     
                    log.commission_month     
                FROM commission_payout_logs log
                JOIN sales_users u ON log.sale_key = u.sale_key
            """
            # <<< END >>>

            if search_term:
                where_clauses.append("(u.sale_name ILIKE %s OR log.sale_key ILIKE %s)")
                params.extend([f"%{search_term}%", f"%{search_term}%"])

            if selected_year != "ทุกปี":
                where_clauses.append("log.commission_year = %s") # <--- ✅ แก้ไข
                params.append(int(selected_year))

            if selected_month != "ทุกเดือน":
                month_num = self.thai_month_map[selected_month]
                where_clauses.append("log.commission_month = %s") # <--- ✅ แก้ไข
                params.append(month_num)
            
            where_clause = " AND ".join(where_clauses)
            query = f"{base_query} WHERE {where_clause} ORDER BY log.commission_year DESC, log.commission_month DESC, log.timestamp DESC"
            
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
            self._create_payout_history_table(df)

            display_text = f"รอบบิล: {selected_month}"
            if selected_year != "ทุกปี": display_text += f" ปี {selected_year}"
            if hasattr(self, 'payout_page_label'):
                self.payout_page_label.configure(text=display_text)

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดประวัติ: {e}", parent=self)
            traceback.print_exc()
            self._create_payout_history_table(pd.DataFrame())

    def _on_payout_history_double_click(self, event, tree):
        """(เวอร์ชันแก้ไข) ตรวจสอบการคลิกให้แม่นยำขึ้น"""
        try:
            # 1. ตรวจสอบว่าคลิกที่ส่วนไหนของตาราง
            region = tree.identify("region", event.x, event.y)
            if region != "cell": 
                return # ถ้าคลิกที่หัวตารางหรือขอบ ให้ข้ามไป

            # 2. ดึง Item ที่ถูกเลือก
            selected_item_iid = tree.focus()
            if not selected_item_iid:
                return # ถ้าไม่มีการเลือก ให้ข้ามไป
            
            # 3. แปลง ID เป็นตัวเลข (iid ที่เราใส่ไว้คือ row['id'])
            payout_id = int(selected_item_iid)
            
            # 4. เปิดหน้าต่าง Popup
            # ตรวจสอบว่า Class PayoutDetailWindow ถูก Import มาแล้ว
            from hr_windows import PayoutDetailWindow 
            PayoutDetailWindow(master=self, app_container=self.app_container, payout_id=payout_id)

        except Exception as e:
            # ถ้ามี Error ให้แจ้งเตือนออกมา (จะได้รู้ว่าผิดตรงไหน)
            messagebox.showerror("Error", f"ไม่สามารถเปิดรายละเอียดได้: {e}", parent=self)
            traceback.print_exc()

    ### --- จุดที่แก้ไข --- ###
    # ผมได้รวมฟังก์ชัน _create_plan_a_summary_table และ _create_plan_b_summary_table
    # ให้เป็นฟังก์ชันเดียวคือ _create_commission_summary_table เพื่อลดความซ้ำซ้อนของโค้ด
    # และเพิ่มความยืดหยุ่นในการแสดงผล ไม่ว่าข้อมูลจะมี 2 หรือ 3 คอลัมน์ก็ตาม
    def _create_commission_summary_table(self, summary_df, container=None):
        """สร้างตารางสรุปผลการคำนวณค่าคอมมิชชั่นแบบไดนามิก (พร้อม Scrollbar แยก)"""
        if container is None:
            container = self.process_result_frame
            
        for widget in container.winfo_children(): widget.destroy()

        if summary_df is None or summary_df.empty:
            CTkLabel(container, text="ไม่พบข้อมูลสำหรับสร้างสรุป").pack(pady=20)
            return

        # สร้าง Frame สำหรับตาราง
        tree_frame = CTkFrame(container, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Summary.Treeview.Heading", font=self.header_font_table, background="#3B82F6", foreground="white")
        style.configure("Summary.Treeview", rowheight=30, font=self.entry_font)
        style.map("Summary.Treeview", background=[('selected', "#DBEAFE")])

        # สร้างคอลัมน์แบบไดนามิกจาก DataFrame
        columns_to_show = list(summary_df.columns)
        
        # [แก้ไข] กำหนด height=12 (ประมาณ 12 แถว) เพื่อจำกัดความสูงไม่ให้ยืดจนดันหน้าจอ
        tree = ttk.Treeview(tree_frame, columns=columns_to_show, show="headings", style="Summary.Treeview", height=12)
        
        # [แก้ไข] เพิ่ม Scrollbar แนวตั้งเฉพาะสำหรับตารางนี้
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=v_scroll.set)
        
        # Grid layout (วางตารางคู่กับ Scrollbar)
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")

        # กำหนดชื่อหัวตารางและความกว้าง
        header_map = {
            'description': 'รายการสรุป',
            'value': 'ยอดรวม (บาท)',
            'commission': 'ค่าคอมมิชชั่น (บาท)'
        }

        for col_id in columns_to_show:
            header_text = header_map.get(col_id, col_id)
            anchor = 'e' if col_id in ['value', 'commission'] else 'w'
            width = 400 if col_id == 'description' else 200
            tree.heading(col_id, text=header_text)
            tree.column(col_id, width=width, anchor=anchor)

        # ตั้งค่า Tag สำหรับแถวพิเศษ
        tree.tag_configure('summary_row', font=self.header_font_table, background="#F3F4F6")
        tree.tag_configure('final_row', font=self.header_font_table, background="#D1FAE5")

        # เพิ่มข้อมูลลงในตาราง
        for _, row in summary_df.iterrows():
            values_tuple = []
            for col in columns_to_show:
                val = row[col]
                # Format ตัวเลขให้มี comma และทศนิยม 2 ตำแหน่ง
                if isinstance(val, (int, float)):
                    values_tuple.append(f"{val:,.2f}")
                else:
                    values_tuple.append(val)
            
            desc = row['description']
            tags = ()
            if "สรุป" in desc or "รวม" in desc or "ขั้นต้น" in desc:
                tags = ('summary_row',)
            if "หลังหัก" in desc:
                tags = ('final_row',)
            
            tree.insert("", "end", values=tuple(values_tuple), tags=tags)

    def _get_po_data(self, so_number):
        if not so_number:
            return pd.DataFrame()
        try:
            query = """
                SELECT id, po_number, supplier_name, total_cost, status, 
                       shipping_to_stock_cost, shipping_to_site_cost, relocation_cost
                FROM purchase_orders 
                WHERE so_number = %s
            """
            df = pd.read_sql(query, self.app_container.pg_engine, params=(so_number,))
            return df
        except Exception as e:
            messagebox.showerror("Database Error", f"Could not fetch PO data: {e}", parent=self)
            return pd.DataFrame()
    
    def _open_verification_window(self, so_number):
        """
        เปิดหน้าต่างตรวจสอบข้อมูล (เวอร์ชันแก้ไขสมบูรณ์)
        - ดึงข้อมูล SO และ PO จากฐานข้อมูลโดยตรงและแยกจากกัน
        - ส่งข้อมูลที่ถูกต้องไปยัง HRVerificationWindow
        """
        try:
            # --- 1. ดึงข้อมูล SO ล่าสุดจาก DB ---
            so_query = """
                SELECT c.*, u.sale_name 
                FROM commissions c 
                JOIN sales_users u ON c.sale_key = u.sale_key 
                WHERE c.so_number = %s AND c.is_active = 1 
                ORDER BY c.id DESC LIMIT 1
            """
            system_data_df = pd.read_sql_query(so_query, self.pg_engine, params=(so_number,))

            if system_data_df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูลที่ Active สำหรับ SO: {so_number} ในระบบ", parent=self)
                return
            system_data = system_data_df.iloc[0].to_dict()

            # --- 2. ดึงข้อมูล PO ที่เกี่ยวข้องทั้งหมด โดยใช้ฟังก์ชันเดิมที่ถูกต้อง ---
            po_data = self._get_po_data(so_number)

            # --- 3. ดึงข้อมูลจากไฟล์ Excel (ถ้ามี) ---
            excel_data = {}
            if self.uploaded_df is not None and not self.uploaded_df.empty:
                # ทำให้ so_number เป็น string เพื่อให้เปรียบเทียบได้ถูกต้อง
                self.uploaded_df['so_number'] = self.uploaded_df['so_number'].astype(str)
                excel_data_row = self.uploaded_df[self.uploaded_df['so_number'].str.strip() == str(so_number).strip()]
                if not excel_data_row.empty:
                    excel_data = excel_data_row.iloc[0].to_dict()
            
            # --- 4. ส่งข้อมูลที่สะอาดและถูกต้องไปยังหน้าต่าง Verify ---
            self.app_container.show_hr_verification_window(
                system_data=system_data,
                excel_data=excel_data,
                po_data=po_data,
                refresh_callback=self._refresh_comparison_view,
                target_commission_month=getattr(self, 'current_comparison_month', None),
                target_commission_year=getattr(self, 'current_comparison_year', None),
                user_role=getattr(self, 'user_role', None),
            )
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถเปิดหน้าต่างตรวจสอบได้: {e}", parent=self)
            traceback.print_exc()

    def _on_tree_double_click(self, event, tree, df):
        """
        Callback เมื่อดับเบิลคลิกบน Treeview
        (เวอร์ชันแก้ไข: เพิ่ม df เป็น argument)
        """
        try:
            selected_item_iid = tree.focus()
            if not selected_item_iid: return

            values = tree.item(selected_item_iid, "values")
            if not values: return

            so_number = values[0] # SO Number อยู่คอลัมน์แรกเสมอ

            if not so_number or so_number == 'ยอดรวม (Total)': return
                
            self._open_verification_window(so_number)

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างตรวจสอบได้: {e}", parent=self)
            traceback.print_exc()

    def _initial_load(self):
        self._populate_users_table()
        self._users_loaded = True
        

    def _create_edit_data_tab(self, parent_tab):
        """(เวอร์ชันแก้ไข) สร้าง UI สำหรับหน้า Master Edit SO/PO พร้อมฟิลเตอร์และ Pagination"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(2, weight=1) # แถวที่ 2 (ScrollFrame) จะขยาย

        # --- Frame สำหรับฟิลเตอร์และการค้นหา ---
        search_frame = CTkFrame(parent_tab)
        search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        search_frame.grid_columnconfigure(3, weight=1)
        
        CTkLabel(search_frame, text="เลือกเซลส์:", font=self.label_font).grid(row=0, column=0, padx=(10,5), pady=10)
        self.master_edit_sale_var = tk.StringVar(value="ทั้งหมด")
        self.master_edit_sale_menu = CTkOptionMenu(
            search_frame, 
            variable=self.master_edit_sale_var, 
            values=["ทั้งหมด"] + self.sales_keys_list,
            command=lambda _: self._load_hr_edit_queue()
        )
        self.master_edit_sale_menu.grid(row=0, column=1, padx=5, pady=10)
        
        CTkLabel(search_frame, text="ค้นหาเฉพาะ SO/PO:", font=self.label_font).grid(row=0, column=2, padx=(20,5), pady=10)
        self.master_edit_search_entry = CTkEntry(search_frame, font=self.entry_font, placeholder_text="กรอก SO หรือ PO...")
        self.master_edit_search_entry.grid(row=0, column=3, padx=5, pady=10, sticky="ew")

        self.master_edit_search_entry.bind("<Return>", lambda event: self._search_so_po())
        self.master_edit_search_entry.bind("<KP_Enter>", lambda event: self._search_so_po()) # สำหรับ Enter บน Numpad
        
        search_button = CTkButton(search_frame, text="ค้นหา", command=self._search_so_po, width=80)
        search_button.grid(row=0, column=4, padx=5, pady=10)
        
        clear_button = CTkButton(search_frame, text="ล้าง / แสดงคิวงาน", command=self._load_hr_edit_queue, fg_color="gray", width=120)
        clear_button.grid(row=0, column=5, padx=5, pady=10)

        # +++ ส่วนของปุ่มแบ่งหน้าที่ขาดไป +++
        pagination_frame = CTkFrame(parent_tab, fg_color="transparent")
        pagination_frame.grid(row=1, column=0, padx=10, pady=0, sticky="ew")
        
        self.edit_data_prev_button = CTkButton(pagination_frame, text="<< หน้าก่อนหน้า", command=self._edit_data_prev_page, width=120, state="disabled")
        self.edit_data_prev_button.pack(side="left")

        self.edit_data_page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.edit_data_page_label.pack(side="left", expand=True)

        self.edit_data_next_button = CTkButton(pagination_frame, text="หน้าถัดไป >>", command=self._edit_data_next_page, width=120, state="disabled")
        self.edit_data_next_button.pack(side="right")
        # +++ สิ้นสุดส่วนที่ขาดไป +++

        self.master_edit_results_frame = CTkScrollableFrame(parent_tab, label_text="ผลการค้นหา / คิวงานที่ต้องตรวจสอบ")
        self.master_edit_results_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.master_edit_results_frame.grid_columnconfigure(0, weight=1)

    def _edit_data_prev_page(self):
        if self.edit_data_current_page > 0:
            self.edit_data_current_page -= 1
            self._load_hr_edit_queue()

    def _edit_data_next_page(self):
        self.edit_data_current_page += 1
        self._load_hr_edit_queue()

    def _load_hr_edit_queue(self):
        """(เวอร์ชันแก้ไข) โหลดคิวงาน SO ที่มีสถานะ 'PO Sent'"""
        for widget in self.master_edit_results_frame.winfo_children():
            widget.destroy()
        self.master_edit_search_entry.delete(0, 'end')

        try:
            # --- START: แก้ไข Query ให้มองหาสถานะ 'PO Sent' ---
            base_query = "FROM commissions WHERE status = 'PO Sent' AND is_active = 1"
            # --- END ---
            params = []
            
            selected_sale = self.master_edit_sale_var.get()
            if selected_sale != "ทั้งหมด":
                base_query += " AND sale_key = %s"
                params.append(selected_sale)

            count_query = f"SELECT COUNT(*) {base_query}"
            total_rows = pd.read_sql_query(count_query, self.pg_engine, params=tuple(params)).iloc[0,0]
            total_pages = (total_rows + self.edit_data_rows_per_page - 1) // self.edit_data_rows_per_page

            offset = self.edit_data_current_page * self.edit_data_rows_per_page
            data_query = f"SELECT id, so_number, customer_name, sale_key {base_query} ORDER BY timestamp DESC LIMIT %s OFFSET %s"
            final_params = params + [self.edit_data_rows_per_page, offset]
            
            df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(final_params))

            self.edit_data_page_label.configure(text=f"หน้า {self.edit_data_current_page + 1} / {max(1, total_pages)}")
            self.edit_data_prev_button.configure(state="normal" if self.edit_data_current_page > 0 else "disabled")
            self.edit_data_next_button.configure(state="normal" if self.edit_data_current_page < total_pages - 1 else "disabled")

            if df.empty:
                CTkLabel(self.master_edit_results_frame, text=f"ไม่พบ SO ที่รอการตรวจสอบในคิวงานของ: {selected_sale}").pack(pady=20)
                return

            for _, row in df.iterrows():
                self._create_so_card_for_editing(self.master_edit_results_frame, row.to_dict())

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดคิวงาน: {e}", parent=self)

    def _search_so_po(self):
        """(เวอร์ชันแก้ไข) ค้นหาข้อมูล SO และ PO เฉพาะตาม Keyword"""
        for widget in self.master_edit_results_frame.winfo_children():
            widget.destroy()
            
        keyword = self.master_edit_search_entry.get().strip().upper()
        if not keyword:
            self._load_hr_edit_queue()
            return

        search_term = keyword
        if search_term.startswith("SO"): search_term = search_term[2:]
        elif search_term.startswith("PO"): search_term = search_term[2:]
        
        try:
            so_query = "SELECT id, so_number, customer_name, sale_key FROM commissions WHERE so_number ILIKE %s AND is_active = 1"
            po_query = "SELECT id, so_number, po_number, supplier_name FROM purchase_orders WHERE po_number ILIKE %s"

            so_df = pd.read_sql_query(so_query, self.pg_engine, params=(f"%{search_term}%",))
            po_df = pd.read_sql_query(po_query, self.pg_engine, params=(f"%{search_term}%",))

            if so_df.empty and po_df.empty:
                CTkLabel(self.master_edit_results_frame, text=f"ไม่พบข้อมูลสำหรับ '{keyword}'").pack(pady=20)
                return

            if not so_df.empty:
                CTkLabel(self.master_edit_results_frame, text="ผลการค้นหา: Sales Orders (SO)", font=self.label_font).pack(anchor="w", padx=10, pady=(10,0))
                for _, row in so_df.iterrows():
                    self._create_so_card_for_editing(self.master_edit_results_frame, row.to_dict())

            if not po_df.empty:
                CTkLabel(self.master_edit_results_frame, text="ผลการค้นหา: Purchase Orders (PO)", font=self.label_font).pack(anchor="w", padx=10, pady=(10,0))
                for _, row in po_df.iterrows():
                    po_id = int(row['id'])
                    po_card = CTkFrame(self.master_edit_results_frame, border_width=1)
                    po_card.pack(fill="x", padx=10, pady=5)
                    info = f"PO: {row['po_number']} | SO: {row['so_number']} | Supplier: {row['supplier_name']}"
                    CTkLabel(po_card, text=info).pack(side="left", padx=10, pady=5)
                    CTkButton(po_card, text="แก้ไข PO", command=lambda pid=po_id: self._open_po_editor_for_hr(pid)).pack(side="right", padx=10, pady=5)

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการค้นหา: {e}", parent=self)

    def _create_so_card_for_editing(self, parent, so_data):
        """(เวอร์ชันปรับปรุง) Helper สำหรับสร้าง SO Card พร้อมใส่สีหากมียอดค้างชำระ"""
        so_id = int(so_data['id'])
        so_number = so_data['so_number']
        
        # --- START: แก้ไข Logic ตรวจสอบยอดค้างชำระ ---
        raw_diff = so_data.get('difference_amount', 0.0) or 0.0
        difference_amount = float(raw_diff) # แปลงเป็น float ให้ชัวร์

        # กำหนด Threshold (ค่าความคลาดเคลื่อนที่ยอมรับได้) เป็น -0.01
        # ถ้าน้อยกว่า -0.01 แปลว่าขาดจริง (เช่น -0.02, -100)
        is_short_payment = difference_amount < -0.01 
        
        # ถ้าโอนขาด ให้ใช้สีส้มอ่อน/แดง, ถ้าปกติ ให้ใช้สีฟ้าอ่อน
        card_color = "#FEF3C7" if is_short_payment else "#F0F9FF"
        info_text_color = "#92400E" if is_short_payment else "gray"
        
        so_card = CTkFrame(parent, border_width=1, fg_color=card_color)
        # --- END ---
        
        so_card.pack(fill="x", padx=10, pady=8)
        so_card.grid_columnconfigure(0, weight=1)
        
        header_frame = CTkFrame(so_card, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        header_frame.grid_columnconfigure(0, weight=1)
        
        # --- START: ปรับปรุงการแสดงข้อความ ---
        main_info_text = f"SO: {so_number}  |  ลูกค้า: {so_data.get('customer_name','N/A')}  |  เซลส์: {so_data.get('sale_key','N/A')}"
        CTkLabel(header_frame, text=main_info_text, font=self.entry_font).grid(row=0, column=0, sticky="w")

        # แสดงข้อความยอดค้างชำระ เฉพาะเมื่อขาดจริง (เกิน 0.01)
        if is_short_payment:
            due_text = f"⚠️ ยอดโอนขาด: {abs(difference_amount):,.2f} บาท"
            CTkLabel(so_card, text=due_text, text_color=info_text_color, font=CTkFont(size=12, weight="bold")).grid(row=1, column=0, sticky="w", padx=10, pady=(0,5))
        # --- END ---
        
        action_frame = CTkFrame(header_frame, fg_color="transparent")
        action_frame.grid(row=0, column=1, sticky="e")
        
        po_container = CTkFrame(so_card, fg_color="#FFFFFF")
        
        CTkButton(action_frame, text="แก้ไข SO", width=100, command=lambda sid=so_id: self._open_so_editor_for_hr(sid)).pack(side="left", padx=5)
        CTkButton(action_frame, text="แสดง/ซ่อน POs", width=120, fg_color="gray", command=lambda s_num=so_number, container=po_container: self._toggle_po_list(s_num, container)).pack(side="left", padx=5)

    def _toggle_po_list(self, so_number, container):
        """(ฟังก์ชันใหม่) แสดง/ซ่อน และโหลดรายการ PO ที่เกี่ยวข้องกับ SO"""
        if container.winfo_viewable():
            container.grid_forget()
        else:
            container.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
            if not container.winfo_children(): # โหลดข้อมูล PO เฉพาะครั้งแรกที่กด
                try:
                    query = "SELECT id, po_number, supplier_name, status FROM purchase_orders WHERE so_number = %s"
                    df = pd.read_sql_query(query, self.pg_engine, params=(so_number,))
                    if df.empty:
                        CTkLabel(container, text="- ไม่พบ PO ที่เกี่ยวข้อง -").pack(padx=10, pady=10)
                        return
                    for _, row in df.iterrows():
                        po_id = int(row['id'])
                        po_card = CTkFrame(container, border_width=1, border_color="#E2E8F0")
                        po_card.pack(fill="x", padx=10, pady=5)
                        status_th = STATUS_THAI_MAP.get(row['status'], row['status'])
                        info = f"  - PO: {row['po_number']} | Supplier: {row['supplier_name']} | สถานะ: {status_th}"
                        CTkLabel(po_card, text=info).pack(side="left", padx=10, pady=5)
                        CTkButton(po_card, text="แก้ไข PO", width=100, command=lambda pid=po_id: self._open_po_editor_for_hr(pid)).pack(side="right", padx=10, pady=5)
                except Exception as e:
                    messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลด PO ย่อย: {e}", parent=self)
    
    def _show_loading(self, frame_to_clear):
        for widget in frame_to_clear.winfo_children(): widget.destroy()
        loading_label = CTkLabel(frame_to_clear, text="กำลังโหลดข้อมูล...", font=CTkFont(size=18, slant="italic"), text_color="gray50"); loading_label.pack(expand=True, pady=20); self.update_idletasks(); return loading_label

    def _get_date_range_from_period(self, period):
        today = datetime.now(); year = today.year
        if period == "เดือนนี้": start_date, end_date = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0), today
        elif period == "ปีนี้": start_date, end_date = today.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0), today
        elif period in self.thai_month_map: month_num = self.thai_month_map[period]; start_date = datetime(year, month_num, 1, 0, 0, 0); last_day = calendar.monthrange(year, month_num)[1]; end_date = datetime(year, month_num, last_day, 23, 59, 59)
        else: start_date, end_date = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0), today
        return start_date, end_date

    # ในไฟล์ hr_screen.py ค้นหาฟังก์ชันเดิมแล้วแทนที่ด้วยอันนี้
    def _create_sales_target_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)
        
        # ตัวแปรเก็บช่วงเวลา
        self.sales_target_period_var = tk.StringVar(value="เดือนนี้")

        # ตัวแปรเก็บรายชื่อคนที่เลือก (เริ่มต้น = None แปลว่าเอาทุกคน)
        self.selected_sales_filter = None

        # โหมดแสดงผล: 'chart' หรือ 'table'
        self.sales_view_mode = 'chart'

        # --- 1. Filter Toolbar ---
        filter_frame = CTkFrame(parent_tab, fg_color="transparent")
        filter_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        # [ปุ่มกรองพนักงาน]
        self.filter_btn = CTkButton(filter_frame, text="👤 กรองพนักงาน (ทั้งหมด)", 
                                    command=self._open_sales_filter_dialog,
                                    fg_color="#6366F1") # สีม่วง
        self.filter_btn.pack(side="left", padx=(5, 15))

        # [ส่วนเลือกรอบเดือนค่าคอม — เดือน/ปี เท่านั้น (ไม่มีวัน)]
        months = self.thai_months
        current_year = datetime.now().year
        years = [str(y + 543) for y in range(current_year - 2, current_year + 3)]

        def create_month_picker(label_text):
            frame = CTkFrame(filter_frame, fg_color="transparent")
            frame.pack(side="left", padx=8)
            CTkLabel(frame, text=label_text,
                     font=self.label_font_bold).pack(side="left", padx=(0, 5))
            m_idx = datetime.now().month - 1
            m_var = tk.StringVar(value=months[m_idx])
            CTkOptionMenu(frame, variable=m_var, values=months,
                          width=110).pack(side="left", padx=2)
            y_var = tk.StringVar(value=str(current_year + 543))
            CTkOptionMenu(frame, variable=y_var, values=years,
                          width=80).pack(side="left", padx=2)
            return m_var, y_var

        self.start_m_var, self.start_y_var = create_month_picker("จากรอบ:")
        self.end_m_var,   self.end_y_var   = create_month_picker("ถึงรอบ:")

        # [ปุ่มค้นหา]
        search_btn = CTkButton(filter_frame, text="🔍 ค้นหา", width=100,
                               fg_color=self.theme["primary"],
                               command=self._on_target_filter_search)
        search_btn.pack(side="left", padx=20)

        # [Toggle กราฟ / ตาราง]
        toggle_frame = CTkFrame(filter_frame, fg_color="transparent")
        toggle_frame.pack(side="right", padx=(0, 5))

        def _set_view(mode):
            self.sales_view_mode = mode
            if mode == 'chart':
                btn_chart.configure(fg_color=self.theme["primary"], text_color="white")
                btn_table.configure(fg_color="#E2E8F0", text_color="#475569")
            else:
                btn_table.configure(fg_color=self.theme["primary"], text_color="white")
                btn_chart.configure(fg_color="#E2E8F0", text_color="#475569")
            self._update_sales_target_dashboard()

        btn_chart = CTkButton(toggle_frame, text="📊 กราฟ", width=90,
                              fg_color=self.theme["primary"], text_color="white",
                              corner_radius=6, command=lambda: _set_view('chart'))
        btn_chart.pack(side="left", padx=2)

        btn_table = CTkButton(toggle_frame, text="📋 ตาราง", width=90,
                              fg_color="#E2E8F0", text_color="#475569",
                              corner_radius=6, command=lambda: _set_view('table'))
        btn_table.pack(side="left", padx=2)
        
        # --- 2. Chart Area ---
        self.sales_target_chart_frame = CTkFrame(parent_tab, border_width=1, corner_radius=10)
        self.sales_target_chart_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    def _update_sales_target_dashboard(self):
        loading = self._show_loading(self.sales_target_chart_frame)
        try:
            period = self.sales_target_period_var.get()
            sales_vs_target_data = self._get_sales_vs_target_data(period)
            loading.destroy()
            mode = getattr(self, 'sales_view_mode', 'chart')
            if mode == 'table':
                self._create_sales_vs_target_table(self.sales_target_chart_frame, sales_vs_target_data)
            else:
                self._create_sales_vs_target_chart(self.sales_target_chart_frame, sales_vs_target_data)
        except Exception as e:
            loading.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอัปเดต Dashboard: {e}", parent=self)
            traceback.print_exc()
    
    def _initial_load_sales_target(self):
        """
        ฟังก์ชันสำหรับโหลดข้อมูลเริ่มต้นของแท็บ 'วิเคราะห์เป้าการขาย'
        เมื่อถูกเรียกครั้งแรก
        """
        # เรียกใช้ฟังก์ชันที่มีอยู่แล้วซึ่งทำหน้าที่โหลดและวาดกราฟ
        self._update_sales_target_dashboard()

    def _get_sales_vs_target_data(self, period):
        """
        ดึงยอดขายจาก SO ที่ผ่าน cal com แล้วเท่านั้น (status IN ('Paid','HR Verified'))
        ตัวเลขถูก lock — แก้ไขไม่ได้ หลังจาก cal com
        Filter ช่วงเวลาด้วย commission_month / commission_year
        """
        try:
            today = datetime.now()
            current_year = today.year
            params = []
            date_filter_clauses = []
            target_multiplier = 1.0

            # =========================================================
            # 1. กรองช่วงเวลาด้วย commission_month / commission_year
            #    (ยึดรอบที่ถูกคิดคอม ไม่ใช่วันที่สร้าง SO)
            # =========================================================
            if period == "กำหนดช่วงเวลาเอง...":
                if hasattr(self, 'custom_target_start') and self.custom_target_start:
                    s_date = self.custom_target_start
                    e_date = self.custom_target_end

                    if isinstance(s_date, str):
                        try:    s_date = datetime.strptime(s_date, "%d/%m/%Y")
                        except ValueError: s_date = datetime.strptime(s_date, "%Y-%m-%d")
                    if isinstance(e_date, str):
                        try:    e_date = datetime.strptime(e_date, "%d/%m/%Y")
                        except ValueError: e_date = datetime.strptime(e_date, "%Y-%m-%d")

                    # เปรียบเทียบด้วย commission period (ปี-เดือน-01)
                    date_filter_clauses.append(
                        "MAKE_DATE(c.commission_year, c.commission_month, 1) "
                        "BETWEEN %s::date AND %s::date"
                    )
                    params.extend([s_date.strftime("%Y-%m-%d"), e_date.strftime("%Y-%m-%d")])
                    # นับจำนวนเดือนในช่วง (เช่น ม.ค.–มี.ค. = 3 เดือน)
                    months_diff = (e_date.year - s_date.year) * 12 + (e_date.month - s_date.month) + 1
                    target_multiplier = max(1, months_diff)
                else:
                    # fallback → เดือนปัจจุบัน
                    date_filter_clauses.append("c.commission_month = %s")
                    params.append(today.month)
                    date_filter_clauses.append("c.commission_year = %s")
                    params.append(current_year)

            elif period == "เดือนนี้":
                date_filter_clauses.append("c.commission_month = %s")
                params.append(today.month)
                date_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
                target_multiplier = 1.0

            elif period == "ปีนี้":
                date_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
                target_multiplier = 12.0

            elif period in ["Q1", "Q2", "Q3", "Q4"]:
                quarters = {"Q1": (1,2,3), "Q2": (4,5,6), "Q3": (7,8,9), "Q4": (10,11,12)}
                months = quarters[period]
                date_filter_clauses.append(
                    f"c.commission_month IN ({','.join(map(str, months))})"
                )
                date_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
                target_multiplier = 3.0

            elif period in self.thai_month_map:
                month_num = self.thai_month_map[period]
                date_filter_clauses.append("c.commission_month = %s")
                params.append(month_num)
                date_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
                target_multiplier = 1.0

            else:  # fallback
                date_filter_clauses.append("c.commission_month = %s")
                params.append(today.month)
                date_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)

            date_filter_sql = " AND ".join(date_filter_clauses)

            # =========================================================
            # 2. จัดการกรองพนักงาน
            # =========================================================
            sale_filter_clause = ""
            if hasattr(self, 'selected_sales_filter') and self.selected_sales_filter:
                placeholders = ','.join(
                    ["REPLACE(LOWER(%s), ' ', '')"] * len(self.selected_sales_filter)
                )
                sale_filter_clause = (
                    f" AND REPLACE(LOWER(su.sale_key), ' ', '') IN ({placeholders})"
                )
                params.extend(self.selected_sales_filter)
            elif hasattr(self, 'custom_target_sale') and self.custom_target_sale != "ทั้งหมด":
                sale_filter_clause = (
                    " AND REPLACE(LOWER(su.sale_key), ' ', '') "
                    "= REPLACE(LOWER(%s), ' ', '')"
                )
                params.append(self.custom_target_sale)

            # =========================================================
            # 3. Query หลัก — ดึงจาก commission_payout_logs (locked snapshot)
            #    total_sales คือยอด ณ เวลาที่จ่ายค่าคอมจริง ไม่เปลี่ยนแปลงภายหลัง
            # =========================================================
            query = f"""
                SELECT
                    su.sale_name,
                    su.sale_key,
                    COALESCE(su.sales_target, 0) * %s AS sales_target,
                    COALESCE(SUM(c.total_sales), 0) AS total_sales,
                    0 AS total_outstanding
                FROM sales_users su
                LEFT JOIN commission_payout_logs c
                       ON REPLACE(LOWER(su.sale_key), ' ', '')
                          = REPLACE(LOWER(c.sale_key), ' ', '')
                      AND {date_filter_sql}
                WHERE su.status = 'Active'
                  {sale_filter_clause}
                GROUP BY su.sale_name, su.sale_key, su.sales_target, su.role
                HAVING (su.role = 'Sale'
                        OR COALESCE(SUM(c.total_sales), 0) > 0)
                ORDER BY su.sale_name ASC;
            """

            final_params = [target_multiplier] + params
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(final_params))

            df['sales_target']      = df['sales_target'].fillna(0)
            df['total_sales']       = df['total_sales'].fillna(0)
            df['total_outstanding'] = df['total_outstanding'].fillna(0)

            # ── Sale Center: ดึงจาก commissions table โดยตรง ──────────
            # (ไม่ผ่าน commission_payout_logs เพราะไม่มีการคิดค่าคอม)
            sc_date_filter = date_filter_sql.replace("c.", "sc.")
            sc_query = f"""
                SELECT
                    COALESCE(SUM(sc.sales_service_amount), 0) AS total_sales
                FROM commissions sc
                WHERE REPLACE(LOWER(sc.sale_key), ' ', '') = 'salecenter'
                  AND {sc_date_filter}
            """
            sc_params = params  # ใช้ params เดียวกัน (ไม่มี target_multiplier)
            sc_df = pd.read_sql_query(sc_query, self.pg_engine, params=tuple(sc_params))
            sc_total = float(sc_df['total_sales'].iloc[0]) if not sc_df.empty else 0.0

            # เพิ่ม Sale Center row เข้า df ถ้ามียอด
            if sc_total > 0:
                sc_row = pd.DataFrame([{
                    'sale_name':        'Sale Center',
                    'sale_key':         'Sale Center',
                    'sales_target':     0.0,
                    'total_sales':      sc_total,
                    'total_outstanding': 0.0,
                }])
                df = pd.concat([df, sc_row], ignore_index=True)

            return df

        except Exception as e:
            print(f"Error getting sales vs target data: {e}")
            messagebox.showerror(
                "Database Error",
                f"ไม่สามารถดึงข้อมูลเป้าหมายการขายได้: {e}",
                parent=self
            )
            traceback.print_exc()
            return pd.DataFrame(
                columns=['sale_name', 'sale_key', 'sales_target',
                         'total_sales', 'total_outstanding']
            )

    def _create_sales_vs_target_chart(self, parent_frame, data_df):
        from matplotlib.patches import FancyBboxPatch
        from matplotlib.lines import Line2D
        from matplotlib.patches import Patch
        import matplotlib.patheffects as pe

        # ── ล้างกราฟเก่า ──────────────────────────────────────────
        if hasattr(self, 'sales_target_chart_canvas') and self.sales_target_chart_canvas:
            self.sales_target_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children():
            widget.destroy()

        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูลพนักงานขาย",
                     font=self.header_font_table).pack(expand=True)
            return

        # ── 1. Config ─────────────────────────────────────────────────
        # sale_key ที่ต้องการซ่อน (test / admin accounts)
        EXCLUDE_SALE_KEYS = {
            's', 'd', 'p', 'mp', 'ms', 'hr', 'sm',
            'Pimhathai',
        }
        # Sale Center แสดงด้วยสีพิเศษ (ยอดบริษัท ไม่คิดค่าคอม)
        SALE_CENTER_KEY = 'Sale Center'

        # Merge config: sale_key → (ชื่อกลุ่ม, label ในแท่ง)
        # คนที่มีหลาย ID ให้เพิ่มคู่ตรงนี้
        PERSON_MERGE = {
            'VOW-P': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ภาณุพงศ์'),
            'VOW-S': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ฐรินทร์ญา'),
        }

        # กรอง test/admin accounts ออกก่อน
        df2 = data_df[~data_df['sale_key'].isin(EXCLUDE_SALE_KEYS)].copy()

        def _group_name(row):
            k = str(row['sale_key']).strip()
            return PERSON_MERGE[k][0] if k in PERSON_MERGE else row['sale_name']

        def _seg_label(row):
            k = str(row['sale_key']).strip()
            return PERSON_MERGE[k][1] if k in PERSON_MERGE else row['sale_name']

        df2['_group']     = df2.apply(_group_name, axis=1)
        df2['_seg_label'] = df2.apply(_seg_label,  axis=1)

        grouped = df2.groupby('_group', sort=False)
        people_data = []
        for name, grp in grouped:
            sub_items = []
            for _, row in grp.iterrows():
                sub_items.append({
                    'sale_key': row['sale_key'],
                    'label':    row['_seg_label'],
                    'sales':    float(row['total_sales']),
                })
            # เรียง: ก้อนใหญ่ด้านล่าง ก้อนเล็กด้านบน
            sub_items.sort(key=lambda s: s['sales'], reverse=True)
            people_data.append({
                'name':              name,
                'total_sales':       float(grp['total_sales'].sum()),
                'total_outstanding': float(grp['total_outstanding'].sum()),
                'target':            float(grp['sales_target'].sum()),
                'sub_items':         sub_items,
            })
        people_data.sort(key=lambda p: p['total_sales'], reverse=True)

        # กรองคนที่ไม่มียอดและไม่มีเป้าหมาย (N/A) ออกจากกราฟ
        people_data = [p for p in people_data if p['total_sales'] > 0 or p['target'] > 0]

        n            = len(people_data)
        names        = [p['name']              for p in people_data]
        sales        = [p['total_sales']       for p in people_data]
        targets      = [p['target']            for p in people_data]
        outstandings = [p['total_outstanding'] for p in people_data]

        # ── 2. สี achievement (ตามยอดรวม) ───────────────────────────
        # เขียว ≥100% | เหลือง 70–99% | แดง <70%
        # Sub-item ที่ 2 ใช้สีอ่อนกว่า เพื่อแยกให้เห็น
        ACHIEVE_COLORS = {
            'green':  ('#22C55E', '#86EFAC'),   # เข้ม, อ่อน
            'yellow': ('#F59E0B', '#FCD34D'),
            'red':    ('#EF4444', '#FCA5A5'),
            'gray':   ('#94A3B8', '#CBD5E1'),
        }

        def achievement_key(s, t):
            if t <= 0: return 'gray'
            pct = s / t
            if pct >= 1.0: return 'green'
            if pct >= 0.7: return 'yellow'
            return 'red'

        pct_labels = []
        for p, s, t in zip(people_data, sales, targets):
            is_sc = any(sub['sale_key'] == SALE_CENTER_KEY for sub in p['sub_items'])
            if is_sc:
                pct_labels.append("ยอดบริษัท")
            elif t > 0:
                pct_labels.append(f"{s/t*100:.0f}%")
            else:
                pct_labels.append("N/A")

        # ── 3. Figure & Axes ─────────────────────────────────────────
        BG     = '#F8FAFC'
        GRID_C = '#E2E8F0'

        chart_width = max(10, n * 1.6)
        fig = Figure(figsize=(chart_width, 7.2), dpi=100, facecolor=BG)
        ax  = fig.add_subplot(111)
        ax.set_facecolor(BG)

        # ── 4. กำหนด x positions & width ────────────────────────────
        x     = np.arange(n)
        width = 0.65
        max_t = max(targets or [1])

        # ── 5. วาดแท่งพื้นหลัง (target track) ────────────────────────
        # วาดเฉพาะส่วนที่ "ยังไม่ถึงเป้า" เพื่อไม่ทับแท่งยอดขาย
        for i, p in enumerate(people_data):
            gap = max(0.0, p['target'] - p['total_sales'])
            if gap > 0:
                ax.bar(x[i], gap, width,
                       bottom=p['total_sales'],
                       color='#E2E8F0', zorder=2, linewidth=0)

        # ── 6. วาดแท่งยอดขาย (stacked ถ้ามีหลาย ID) ─────────────────
        SALE_CENTER_COLOR = ('#0EA5E9', '#7DD3FC')  # sky blue สำหรับ Sale Center
        for i, p in enumerate(people_data):
            total_s  = p['total_sales']
            target   = p['target']
            is_sc    = any(s['sale_key'] == SALE_CENTER_KEY for s in p['sub_items'])
            if is_sc:
                colors = SALE_CENTER_COLOR
            else:
                akey   = achievement_key(total_s, target)
                colors = ACHIEVE_COLORS[akey]        # (dark, light)
            subs     = p['sub_items']
            multi_id = len([s for s in subs if s['sales'] > 0]) > 1

            if not subs:
                continue   # ยอด 0 ข้ามไป

            bottom = 0.0
            for idx, sub in enumerate(subs):
                seg_h = sub['sales']
                if seg_h <= 0:
                    continue
                color = colors[min(idx, 1)]        # dark=ก้อนแรก, light=ก้อนถัดไป
                # partner segment ใช้ hatch เพื่อแยกให้ชัด
                is_partner = (idx > 0)
                ax.bar(x[i], seg_h, width,
                       bottom=bottom,
                       color=color, zorder=3,
                       linewidth=0.8 if is_partner else 0,
                       edgecolor='white' if is_partner else 'none',
                       hatch='//' if is_partner else None,
                       alpha=0.92)

                # ป้ายภายใน segment
                mid_y = bottom + seg_h / 2
                # ถ้าเส้น target ตัดผ่าน segment → เลื่อน text ออกจากเส้น
                seg_top = bottom + seg_h
                t_line  = p['target']
                if t_line > 0 and bottom < t_line < seg_top:
                    # ใช้ midpoint ของครึ่งล่าง (ใต้ target line) ถ้ามีพื้นที่พอ
                    if (t_line - bottom) >= seg_h * 0.35:
                        mid_y = bottom + (t_line - bottom) / 2
                    else:
                        mid_y = t_line + (seg_top - t_line) / 2

                # partner segment (idx>0): สีเข้มอ่านง่ายบน hatch background อ่อน
                txt_color = '#1a5c1a' if is_partner else 'white'
                txt_bbox  = None

                # stroke สีขาวสำหรับ white text, สีอ่อนสำหรับ dark text
                stroke_color = 'white' if txt_color != 'white' else '#00000066'
                fx = [pe.withStroke(linewidth=3, foreground=stroke_color)]

                if multi_id:
                    if seg_h > max_t * 0.10:
                        ax.text(x[i], mid_y,
                                f"{sub['label']}\n{seg_h:,.0f}",
                                ha='center', va='center',
                                fontsize=14, weight='medium',
                                color=txt_color, zorder=7,
                                linespacing=1.4,
                                path_effects=fx)
                    elif seg_h > max_t * 0.04:
                        ax.text(x[i], mid_y,
                                f"{seg_h:,.0f}",
                                ha='center', va='center',
                                fontsize=13, weight='medium',
                                color=txt_color, zorder=7,
                                path_effects=fx)
                    # segment เล็กเกินไป → ไม่แสดงในแท่ง
                else:
                    if seg_h > max_t * 0.06:
                        ax.text(x[i], mid_y,
                                f"{seg_h:,.0f}",
                                ha='center', va='center',
                                fontsize=14, weight='medium',
                                color=txt_color, zorder=7,
                                path_effects=fx)

                bottom += seg_h

        # ── 7. เส้น target แนวนอนต่อคน (dashed) + Target label สีเดียวกับเส้น
        half = width / 2
        for i, t in enumerate(targets):
            if t > 0:
                ax.hlines(t, x[i] - half, x[i] + half,
                          colors='#6366F1', linewidths=2.2,
                          linestyles='--', zorder=5)
                ax.text(x[i], t + max_t * 0.012,
                        f"Target  {t:,.0f}",
                        ha='center', va='bottom',
                        fontsize=9.5, weight='bold',
                        color='#6366F1',
                        zorder=6)

        # ── 8. ป้าย % achievement บนหัวแท่ง ──────────────────────────
        for i, (p, pct) in enumerate(zip(people_data, pct_labels)):
            s    = p['total_sales']
            t    = p['target']
            akey = achievement_key(s, t)
            bar_top = max(s, 0)

            if s == 0 and t > 0:
                # ยังไม่มีข้อมูล SO — แสดง label ในพื้นที่ grey target bar
                ax.text(x[i], t * 0.5,
                        "ยังไม่มี\nข้อมูล SO",
                        ha='center', va='center',
                        fontsize=11, weight='bold',
                        color='#94A3B8', zorder=8,
                        style='italic', linespacing=1.4)
            else:
                # ดัน % badge ให้พ้น Target label เสมอ (Target label สูงสุด ≈ t + 14%)
                pct_y = max(bar_top, t) + max_t * 0.16

                ax.text(x[i], pct_y,
                        pct,
                        ha='center', va='bottom',
                        fontsize=16, weight='medium', color='black',
                        zorder=8,
                        path_effects=[pe.withStroke(linewidth=2, foreground='white')])

                # ถ้ายอดรวมเล็กมาก ให้แสดงตัวเลขนอกแท่งด้วย
                if s <= max_t * 0.08:
                    ax.text(x[i], s + max_t * 0.012,
                            f"{s:,.0f}",
                            ha='center', va='bottom',
                            fontsize=12, weight='bold', color='black', zorder=7)

        # ── 9. ยอดค้างชำระ — ซ่อนไว้ (ไม่แสดงในกราฟ)

        # ── 10. ตกแต่ง Axes ──────────────────────────────────────────
        max_sales = max((p['total_sales'] for p in people_data), default=0)
        y_top = max(max_sales, max_t) + max_t * 0.55
        ax.set_ylim(0, y_top)

        formatter = FuncFormatter(lambda v, _: f'{v:,.0f}')
        ax.yaxis.set_major_formatter(formatter)
        ax.tick_params(axis='y', labelsize=13)
        ax.tick_params(axis='x', pad=8)

        # x-axis label: ชื่อ (2 บรรทัด) + เป้า (บรรทัด 3)
        def _fmt_name(nm, tgt):
            # "/" → newline เพื่อให้ทุกชื่อแตกบรรทัดสม่ำเสมอ
            short = nm.replace(' / ', '\n').replace(' ', '\n', 1)
            if tgt > 0:
                return f"{short}\nเป้า {tgt:,.0f}"
            return short

        wrapped_names = [_fmt_name(nm, tgt) for nm, tgt in zip(names, targets)]
        ax.set_xticks(x)
        ax.set_xticklabels(wrapped_names, rotation=0, ha='center',
                           fontsize=10, weight='medium', color='black',
                           linespacing=1.35)

        ax.set_ylabel('จำนวนเงิน (บาท)', fontsize=12, weight='medium',
                      color='black', labelpad=10)

        # ── Summary KPI ─────────────────────────────────────────────
        team_sales  = sum(p['total_sales'] for p in people_data)
        team_target = sum(p['target']      for p in people_data)
        team_pct    = (team_sales / team_target * 100) if team_target > 0 else 0
        n_hit  = sum(1 for p in people_data if p['target'] > 0 and p['total_sales'] >= p['target'])
        n_miss = sum(1 for p in people_data if p['target'] > 0 and p['total_sales'] < p['target'])
        summary = (f"ทีมรวม  {team_sales:,.0f} / {team_target:,.0f} บาท"
                   f"  ({team_pct:.0f}%)     "
                   f"ถึงเป้า {n_hit} คน  ·  ยังไม่ถึง {n_miss} คน")

        ax.set_title(
            f"ยอดขาย vs เป้าหมาย  (เฉพาะ SO ที่คิดค่าคอมแล้ว)\n{summary}",
            fontsize=16, weight='bold', color='#0F172A',
            loc='left', pad=12,
            # ให้ subtitle เล็กและสีอ่อนกว่าด้วย MultiLineText trick ไม่ได้ใน mpl
            # → ใช้ title บรรทัดเดียวแล้ว overlay summary แยก
        )
        # override title ให้ชัดขึ้น แล้ว annotate summary ด้านล่าง
        ax.set_title('ยอดขาย vs เป้าหมาย  (เฉพาะ SO ที่คิดค่าคอมแล้ว)',
                     fontsize=16, weight='semibold', color='#0F172A',
                     loc='left', pad=36)
        ax.text(0, 1.015, summary,
                transform=ax.transAxes,
                fontsize=10.5, color='#64748B',
                ha='left', va='bottom', weight='medium')

        # Grid แนวนอนเบาๆ
        ax.yaxis.grid(True, color=GRID_C, linewidth=0.8, zorder=0)
        ax.set_axisbelow(True)

        # ลบ spine ด้านบน-ขวา (clean look)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(GRID_C)
        ax.spines['bottom'].set_color(GRID_C)
        ax.tick_params(colors='black')

        # ── 11. Legend ────────────────────────────────────────────────
        has_multi_id = any(len(p['sub_items']) > 1 for p in people_data)
        has_sale_center = any(
            any(s['sale_key'] == SALE_CENTER_KEY for s in p['sub_items'])
            for p in people_data
        )
        legend_items = [
            Patch(facecolor='#22C55E', label='≥ 100% เป้า'),
            Patch(facecolor='#F59E0B', label='70–99% เป้า'),
            Patch(facecolor='#EF4444', label='< 70% เป้า'),
            Line2D([0], [0], color='#6366F1', lw=2,
                   linestyle='--', label='เป้าหมาย'),
            Patch(facecolor='#E2E8F0', edgecolor='#CBD5E1',
                  label='ส่วนที่ยังไม่ถึงเป้า'),
        ]
        if has_multi_id:
            legend_items.append(
                Patch(facecolor='#86EFAC', label='ยอดขายของพาร์ทเนอร์')
            )
        if has_sale_center:
            legend_items.append(
                Patch(facecolor='#0EA5E9', label='ยอดบริษัท (Sale Center)')
            )
        leg = ax.legend(handles=legend_items,
                        loc='upper right',
                        ncol=1,
                        frameon=True, framealpha=0.95,
                        edgecolor='#CBD5E1', fontsize=10,
                        prop={'weight': 'bold', 'size': 10},
                        borderpad=0.7, labelspacing=0.4)

        try:
            fig.tight_layout(rect=[0, 0.08, 1, 1])
        except Exception:
            fig.subplots_adjust(left=0.08, right=0.92, top=0.92, bottom=0.10)

        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(
            side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.sales_target_chart_canvas = canvas

    def _create_sales_vs_target_table(self, parent_frame, data_df):
        import tkinter.ttk as ttk

        for w in parent_frame.winfo_children():
            w.destroy()

        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูล",
                     font=self.header_font_table).pack(expand=True)
            return

        # ── ประมวลผลข้อมูล (เหมือน chart) ───────────────────────────
        EXCLUDE_SALE_KEYS = {'s','d','p','mp','ms','hr','sm','Sale Center','Pimhathai'}
        PERSON_MERGE = {
            'VOW-P': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ภาณุพงศ์'),
            'VOW-S': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ฐรินทร์ญา'),
        }
        df2 = data_df[~data_df['sale_key'].isin(EXCLUDE_SALE_KEYS)].copy()
        df2['_group'] = df2.apply(
            lambda r: PERSON_MERGE[r['sale_key']][0] if r['sale_key'] in PERSON_MERGE else r['sale_name'],
            axis=1)
        rows = []
        for name, grp in df2.groupby('_group', sort=False):
            t = float(grp['sales_target'].sum())
            s = float(grp['total_sales'].sum())
            pct  = (s / t * 100) if t > 0 else 0.0
            diff = s - t
            rows.append({'name': name, 'target': t, 'sales': s, 'pct': pct, 'diff': diff})
        rows.sort(key=lambda r: r['sales'], reverse=True)

        total_t = sum(r['target'] for r in rows)
        total_s = sum(r['sales']  for r in rows)
        total_pct  = (total_s / total_t * 100) if total_t > 0 else 0.0
        total_diff = total_s - total_t

        # ── Period label ─────────────────────────────────────────────
        try:
            s_m = self.thai_months.index(self.start_m_var.get()) + 1
            e_m = self.thai_months.index(self.end_m_var.get())   + 1
            s_y = int(self.start_y_var.get())
            e_y = int(self.end_y_var.get())
            if s_m == e_m and s_y == e_y:
                period_label = f"{self.start_m_var.get()} {s_y}"
            else:
                period_label = f"{self.start_m_var.get()} {s_y} – {self.end_m_var.get()} {e_y}"
        except Exception:
            period_label = "รอบที่เลือก"

        # ── Layout ───────────────────────────────────────────────────
        outer = CTkScrollableFrame(parent_frame, fg_color="white", corner_radius=10)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        CTkLabel(outer, text=f"สรุปยอดขาย vs เป้าหมาย  —  {period_label}",
                 font=CTkFont(size=15, weight="bold"),
                 text_color="#0F172A").pack(anchor="w", padx=16, pady=(12, 8))

        # ── Treeview style ────────────────────────────────────────────
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("SalesTable.Treeview",
                        background="white", foreground="#1E293B",
                        rowheight=34, fieldbackground="white",
                        font=('TH Sarabun New', 13))
        style.configure("SalesTable.Treeview.Heading",
                        background="#D1FAE5", foreground="#065F46",
                        font=('TH Sarabun New', 13, 'bold'), relief="flat")
        style.map("SalesTable.Treeview",
                  background=[('selected', '#EDE9FE')],
                  foreground=[('selected', '#1E293B')])

        cols = ('name', 'target', 'sales', 'pct', 'diff')
        tree = ttk.Treeview(outer, columns=cols, show='headings',
                            style="SalesTable.Treeview", height=len(rows) + 1)

        tree.heading('name',   text='พนักงาน')
        tree.heading('target', text='เป้าหมาย (บาท)')
        tree.heading('sales',  text='ยอดขายจริง (บาท)')
        tree.heading('pct',    text='%')
        tree.heading('diff',   text='ส่วนต่าง (บาท)')

        tree.column('name',   width=200, anchor='w')
        tree.column('target', width=170, anchor='e')
        tree.column('sales',  width=170, anchor='e')
        tree.column('pct',    width=90,  anchor='center')
        tree.column('diff',   width=170, anchor='e')

        style.configure("SalesTable.Treeview", rowheight=34)

        # alternate row colors
        tree.tag_configure('odd',      background='#F8FAFC')
        tree.tag_configure('even',     background='white')
        tree.tag_configure('negative', foreground='#DC2626')
        tree.tag_configure('positive', foreground='#16A34A')
        tree.tag_configure('total',    background='#EFF6FF',
                           font=('TH Sarabun New', 13, 'bold'),
                           foreground='#1D4ED8')

        for idx, r in enumerate(rows):
            diff_str = f"+{r['diff']:,.0f}" if r['diff'] >= 0 else f"{r['diff']:,.0f}"
            tag_row  = 'odd' if idx % 2 else 'even'
            tag_diff = 'positive' if r['diff'] >= 0 else 'negative'
            tree.insert('', 'end', tags=(tag_row, tag_diff), values=(
                r['name'],
                f"{r['target']:,.0f}" if r['target'] > 0 else '—',
                f"{r['sales']:,.0f}",
                f"{r['pct']:.1f}%",
                diff_str if r['target'] > 0 else '—',
            ))

        # Total row
        total_diff_str = f"+{total_diff:,.0f}" if total_diff >= 0 else f"{total_diff:,.0f}"
        tree.insert('', 'end', tags=('total',), values=(
            'รวมทีม',
            f"{total_t:,.0f}",
            f"{total_s:,.0f}",
            f"{total_pct:.1f}%",
            total_diff_str,
        ))

        tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    def _get_po_status_summary(self, start_date, end_date):
        try: query = "SELECT status, COUNT(id) as count FROM purchase_orders WHERE timestamp BETWEEN %s AND %s GROUP BY status"; df = pd.read_sql_query(query, self.pg_engine, params=(start_date, end_date)); return df
        except Exception as e: print(f"Error getting PO status summary: {e}"); messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลสถานะ PO ได้: {e}", parent=self); return pd.DataFrame(columns=['status', 'count'])

    def _create_po_pie_chart(self, parent_frame, data_df):
        if hasattr(self, 'po_chart_canvas') and self.po_chart_canvas: self.po_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children(): widget.destroy()
        if data_df.empty: CTkLabel(parent_frame, text="ไม่พบข้อมูลใบสั่งซื้อในช่วงเวลานี้", font=self.header_font_table).pack(expand=True); return

        # <<< START: 1. เพิ่มโค้ดสำหรับเตรียมการ "ระเบิด" ชิ้นส่วนเล็กๆ >>>
        total = data_df['count'].sum()
        # สร้าง list ของ explode โดยให้ชิ้นที่น้อยกว่า 5% แยกตัวออกมา 0.2 หน่วย
        explode_values = [0.2 if (count / total) < 0.05 else 0 for count in data_df['count']]
        # <<< END >>>

        fig = Figure(figsize=(5, 4), dpi=100, facecolor=self.theme["bg"])
        ax = fig.add_subplot(111)
        
        status_colors_map = { 
            "Approved": "#BBF7D0", 
            "Pending Approval": "#FEF08A", 
            "Rejected": "#FECACA", 
            "Cancelled": "#FBCFE8", # เพิ่มสีสำหรับ Cancelled
            "Draft": "#E5E7EB" 
        }
        pie_colors = [status_colors_map.get(status, "#B0B0B0") for status in data_df['status']]

        # <<< START: 2. แก้ไขการเรียกใช้ ax.pie() ให้รองรับการปรับแต่งใหม่ >>>
        ax.pie(
            data_df['count'], 
            labels=data_df['status'], 
            # ฟังก์ชัน lambda นี้จะแสดง % ก็ต่อเมื่อชิ้นส่วนใหญ่กว่า 3%
            autopct=lambda pct: f'{pct:.1f}%' if pct > 3 else '',
            startangle=140,         # ปรับมุมเริ่มต้นเพื่อให้กลุ่มเล็กๆ อยู่ด้านบน
            colors=pie_colors, 
            textprops={'fontsize': 12},
            pctdistance=0.85,       # ปรับระยะห่างของ % ให้อยู่ด้านในมากขึ้น
            explode=explode_values  # สั่งให้ "ระเบิด" ชิ้นส่วนตามที่เราคำนวณไว้
        )
        # <<< END >>>

        ax.axis('equal')
        ax.set_title('สัดส่วนสถานะใบสั่งซื้อ (PO)', fontsize=16, weight="bold")
        
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.po_chart_canvas = canvas

    def _create_dashboard_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        filter_frame = CTkFrame(parent_tab, fg_color="transparent")
        filter_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        CTkLabel(filter_frame, text="ช่วงเวลา:", font=self.label_font).pack(side="left", padx=(5, 10))
        self.dashboard_period_var = tk.StringVar(value="เดือนนี้")
        period_menu = CTkOptionMenu(filter_frame, variable=self.dashboard_period_var, values=self.period_options, command=lambda _: self._update_dashboard())
        period_menu.pack(side="left", padx=5)

        refresh_button = CTkButton(filter_frame, text="Refresh", width=100, fg_color=self.theme["primary"], command=self._update_dashboard)
        refresh_button.pack(side="left", padx=20)
        
        # --- START: เพิ่มส่วนนี้เข้ามา ---
        # ตรวจสอบ Role ของผู้ใช้ที่ Login เข้ามา
        # สมมติว่า Role ถูกเก็บไว้ใน self.app_container.current_user_role
        
        # **โค้ดจะแสดงปุ่มนี้ ก็ต่อเมื่อ Role ของผู้ใช้เป็น 'Director' เท่านั้น**
        if self.user_role == 'Director':
            archive_button = CTkButton(
                filter_frame, 
                text="⚙️ บันทึกประจำปี (Archive)", 
                fg_color="#64748B", 
                hover_color="#475569",
                command=self._annual_archive_data
            )
            archive_button.pack(side="right", padx=10)
        # --- END: สิ้นสุดส่วนที่เพิ่ม ---
        trial_export_button = CTkButton(
                filter_frame,
                text="📄 ทดลอง Export (ไม่ลบข้อมูล)",
                fg_color="#3B82F6", # สีน้ำเงินเพื่อให้แตกต่าง
                hover_color="#2563EB",
                command=self._trial_export_data # << เรียกใช้ฟังก์ชันใหม่
            )
        trial_export_button.pack(side="right", padx=(0, 10))

        chart_container = CTkFrame(parent_tab, fg_color="transparent")
        chart_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        chart_container.grid_columnconfigure((0, 1), weight=1)
        chart_container.grid_rowconfigure(0, weight=1)

        self.sales_chart_frame = CTkFrame(chart_container, border_width=1, corner_radius=10)
        self.sales_chart_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        self.po_chart_frame = CTkFrame(chart_container, border_width=1, corner_radius=10)
        self.po_chart_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

    def _update_dashboard(self):
        loading1 = self._show_loading(self.sales_chart_frame)
        loading2 = self._show_loading(self.po_chart_frame)

        try:
            period = self.dashboard_period_var.get()
            start_date, end_date = self._get_date_range_from_period(period)
            start_date_str = start_date.strftime("%Y-%m-%d %H:%M:%S")
            end_date_str = end_date.strftime("%Y-%m-%d %H:%M:%S")

            # --- START: แก้ไขการส่งค่า ---
            # ส่ง 'period' (string) ไปแทน start_date, end_date
            sales_by_employee_data = self._get_sales_by_employee_data(period) 
            # --- END ---
            po_summary_data = self._get_po_status_summary(start_date_str, end_date_str)

            loading1.destroy()
            self._create_sales_by_employee_chart(self.sales_chart_frame, sales_by_employee_data)

            loading2.destroy()
            self._create_po_pie_chart(self.po_chart_frame, po_summary_data)

        except Exception as e:
            if 'loading1' in locals() and loading1.winfo_exists(): loading1.destroy()
            if 'loading2' in locals() and loading2.winfo_exists(): loading2.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอัปเดต Dashboard: {e}", parent=self)

    def _initial_load_dashboard(self):
        """
        ฟังก์ชันสำหรับโหลดข้อมูลเริ่มต้นของแท็บ 'Dashboard'
        เมื่อถูกเรียกครั้งแรก
        """
        # เรียกใช้ฟังก์ชันที่มีอยู่แล้วซึ่งทำหน้าที่โหลดและวาดกราฟ
        self._update_dashboard()

    def _get_sales_by_employee_data(self, period):
        try:
            today = datetime.now()
            current_year = today.year
            params = []
            date_filter_clauses = []
            target_multiplier = 1.0 

            # =========================================================
            # 1. จัดการเรื่องวันที่ (ใช้ ::date เพื่อความชัวร์ 100%)
            # =========================================================
            if period == "กำหนดช่วงเวลาเอง...":
                if hasattr(self, 'custom_target_start') and self.custom_target_start:
                    s_date = self.custom_target_start
                    e_date = self.custom_target_end

                    if isinstance(s_date, str):
                        try: s_date = datetime.strptime(s_date, "%d/%m/%Y")
                        except ValueError: s_date = datetime.strptime(s_date, "%Y-%m-%d")

                    if isinstance(e_date, str):
                        try: e_date = datetime.strptime(e_date, "%d/%m/%Y")
                        except ValueError: e_date = datetime.strptime(e_date, "%Y-%m-%d")

                    start_str = s_date.strftime("%Y-%m-%d")
                    end_str = e_date.strftime("%Y-%m-%d")

                    date_filter_clauses.append('c."timestamp"::date BETWEEN %s::date AND %s::date')
                    params.extend([start_str, end_str])

                    days_diff = (e_date - s_date).days + 1
                    target_multiplier = max(0.03, days_diff / 30.0)
                else:
                    date_filter_clauses.append('EXTRACT(MONTH FROM c."timestamp"::date) = %s')
                    params.append(today.month)
                    date_filter_clauses.append('EXTRACT(YEAR FROM c."timestamp"::date) = %s')
                    params.append(current_year)

            elif period == "เดือนนี้":
                date_filter_clauses.append('EXTRACT(MONTH FROM c."timestamp"::date) = %s')
                params.append(today.month)
                date_filter_clauses.append('EXTRACT(YEAR FROM c."timestamp"::date) = %s')
                params.append(current_year)
                target_multiplier = 1.0

            elif period == "ปีนี้":
                date_filter_clauses.append('EXTRACT(YEAR FROM c."timestamp"::date) = %s')
                params.append(current_year)
                target_multiplier = 12.0

            elif period in ["Q1", "Q2", "Q3", "Q4"]:
                quarters = {"Q1": (1, 2, 3), "Q2": (4, 5, 6), "Q3": (7, 8, 9), "Q4": (10, 11, 12)}
                months = quarters.get(period)
                date_filter_clauses.append(f'EXTRACT(MONTH FROM c."timestamp"::date) IN ({",".join(["%s"] * len(months))})')
                params.extend(months)
                date_filter_clauses.append('EXTRACT(YEAR FROM c."timestamp"::date) = %s')
                params.append(current_year)
                target_multiplier = 3.0

            elif period in self.thai_month_map:
                month_num = self.thai_month_map[period]
                date_filter_clauses.append('EXTRACT(MONTH FROM c."timestamp"::date) = %s')
                params.append(month_num)
                date_filter_clauses.append('EXTRACT(YEAR FROM c."timestamp"::date) = %s')
                params.append(current_year)
                target_multiplier = 1.0

            else:
                date_filter_clauses.append('EXTRACT(MONTH FROM c."timestamp"::date) = %s')
                params.append(today.month)
                date_filter_clauses.append('EXTRACT(YEAR FROM c."timestamp"::date) = %s')
                params.append(current_year)

            date_filter_sql = " AND ".join(date_filter_clauses)

            # =========================================================
            # 2. กรองสถานะ
            # =========================================================
            status_condition = """
                c.status NOT IN ('Draft', 'Cancelled', 'Rejected by SM', 'Rejected by HR', 'Original', 'Edited')
            """

            # =========================================================
            # 3. จัดการกรองพนักงาน
            # =========================================================
            sale_filter_clause = ""
            if hasattr(self, 'selected_sales_filter') and self.selected_sales_filter:
                placeholders = ",".join(["%s"] * len(self.selected_sales_filter))
                sale_filter_clause = f" AND REGEXP_REPLACE(LOWER(su.sale_key), '\s+', '', 'g') IN ({placeholders})"
                
                cleaned_keys = [key.replace(" ", "").lower() for key in self.selected_sales_filter]
                params.extend(cleaned_keys)

            elif hasattr(self, 'custom_target_sale') and self.custom_target_sale != "ทั้งหมด":
                sale_filter_clause = " AND REGEXP_REPLACE(LOWER(su.sale_key), '\s+', '', 'g') = REGEXP_REPLACE(LOWER(%s), '\s+', '', 'g')"
                params.append(self.custom_target_sale)

            # =========================================================
            # 4. Query หลัก (🔥 ไม้ตาย: ล้างช่องว่างทุกรูปแบบด้วย REGEXP_REPLACE)
            # =========================================================
            query = f"""
                SELECT 
                    su.sale_name, 
                    su.sale_key, 
                    COALESCE(su.sales_target, 0) * %s AS sales_target, 
                    -- ดึงยอดขายและค่าบริการทั้งหมดมารวมกันให้ครบเหมือน SO Grand Total
                    COALESCE(SUM(c.sales_service_amount + COALESCE(c.cutting_drilling_fee, 0) + COALESCE(c.other_service_fee, 0)), 0) AS total_sales,
                    COALESCE(SUM(CASE WHEN COALESCE(c.difference_amount, 0) < -1 THEN ABS(c.difference_amount) ELSE 0 END), 0) AS total_outstanding
                FROM sales_users su
                -- 🔥 จุดที่ทำให้รอด: ใช้ REGEXP_REPLACE กวาดล้างอักขระขยะทุกตัวก่อน JOIN
                LEFT JOIN commissions c 
                    ON REGEXP_REPLACE(LOWER(su.sale_key), '\s+', '', 'g') = REGEXP_REPLACE(LOWER(c.sale_key), '\s+', '', 'g')
                    AND c.is_active = 1
                    AND {status_condition}
                    AND {date_filter_sql}
                WHERE su.status = 'Active'
                    {sale_filter_clause}
                GROUP BY su.sale_name, su.sale_key, su.sales_target, su.role
                HAVING (su.role = 'Sale' OR COALESCE(SUM(c.sales_service_amount), 0) > 0)
                ORDER BY su.sale_name ASC;
            """

            final_params = [target_multiplier] + params

            df = pd.read_sql_query(query, self.pg_engine, params=tuple(final_params))

            df['sales_target'] = df['sales_target'].fillna(0)
            df['total_sales'] = df['total_sales'].fillna(0)
            df['total_outstanding'] = df['total_outstanding'].fillna(0)

            return df

        except Exception as e:
            print(f"Error getting sales vs target data: {e}")
            messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลเป้าหมายการขายได้: {e}", parent=self)
            traceback.print_exc()
            return pd.DataFrame(columns=['sale_name', 'sale_key', 'sales_target', 'total_sales', 'total_outstanding'])

    def _create_sales_by_employee_chart(self, parent_frame, data_df):
        if hasattr(self, 'sales_chart_canvas') and self.sales_chart_canvas:
            self.sales_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children():
            widget.destroy()

        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูลยอดขายตามพนักงานในช่วงเวลานี้", font=self.header_font_table).pack(expand=True)
            return
        
        fig = Figure(figsize=(6, 4), dpi=100, facecolor=self.theme["bg"])
        ax = fig.add_subplot(111)
        ax.set_facecolor(self.theme["bg"])

        # --- ส่วนการเรียงลำดับ (จากมากไปน้อย) ---
        person_totals = data_df.groupby('sale_name')['total_sales'].sum()
        data_df['person_total'] = data_df['sale_name'].map(person_totals)
        data_df = data_df.sort_values(by=['person_total', 'total_sales'], ascending=[False, False])
        
        unique_names = data_df['sale_name'].unique() 
        x = np.arange(len(unique_names))
        width = 0.6 
        
        # --- ดึงข้อมูล Target สำหรับวาดเส้น ---
        target_data = data_df.drop_duplicates(subset='sale_name')['sales_target'].values
        
        # --- (โค้ดส่วน "กราฟสายรุ้ง" เหมือนเดิม) ---
        base_colors = ['#2a9d8f', '#e76f51', '#3B82F6', '#8B5CF6', '#e9c46a', '#10B981', '#264653']
        stack_color = '#f4a261' 

        for i, sale_name in enumerate(unique_names):
            sales_data_for_person = data_df[data_df['sale_name'] == sale_name]
            current_bottom = 0 
            
            for j, (_, row) in enumerate(sales_data_for_person.iterrows()):
                sales_value = row['total_sales']
                sale_key = row['sale_key']

                if j == 0:
                    color = base_colors[i % len(base_colors)]
                else:
                    color = stack_color
                
                rects = ax.bar(x[i], sales_value, width, 
                               bottom=current_bottom, 
                               color=color)
                
                if sales_value > 0:
                    ax.bar_label(rects, labels=[f'{sales_value:,.0f}'], 
                                 label_type='center', color='white', weight='bold', fontsize=9)
                
                current_bottom += sales_value

            if current_bottom > 0:
                y_offset = current_bottom * 0.01
                ax.text(x[i], current_bottom + y_offset, f'{current_bottom:,.0f}',
                        ha='center', va='bottom', fontsize=10, weight='bold')

        # --- START: แก้ไข linestyle (เส้นประ -> เส้นทึบ) ---
        ax.plot(x, target_data, 
                color='#F97316',     # สีส้ม
                marker='o',          # จุดกลม
                linestyle='-',       # <--- แก้ไขตรงนี้เป็นเส้นทึบ
                linewidth=2.5,       # เส้นหนา
                label='Target')      
        # --- END ---

        ax.set_ylabel('ยอดขาย (บาท)', fontsize=12)
        ax.set_title('สรุปยอดขายตามพนักงาน (Active)', fontsize=16, weight="bold")
        
        ax.set_xticks(x)
        ax.set_xticklabels(unique_names, rotation=45, ha="right", fontsize=11)

        formatter = FuncFormatter(lambda y, pos: f'{int(y):,}')
        ax.yaxis.set_major_formatter(formatter)
        ax.grid(axis='y', linestyle='--', alpha=0.7)
        
        ax.legend(prop={'size': 12})
        
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.sales_chart_canvas = canvas

    def _create_audit_log_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1); parent_tab.grid_rowconfigure(0, weight=1)
        self.audit_log_frame = CTkScrollableFrame(parent_tab); self.audit_log_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

    def _populate_audit_log_table(self):
        if not hasattr(self, 'audit_log_frame') or not self.audit_log_frame.winfo_exists():
            print("Warning: audit_log_frame not yet created or visible. Skipping populate.")
            return

        loading = self._show_loading(self.audit_log_frame);
        try:
            df = pd.read_sql_query("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500", self.pg_engine)
            loading.destroy()
            self._create_styled_dataframe_table(self.audit_log_frame, df, "บันทึกกิจกรรม (500 รายการล่าสุด)")
        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            messagebox.showerror("Database Error", f"ยังไม่มีบันทึกกิจกรรม หรือเกิดข้อผิดพลาด: {e}", parent=self)

    def _populate_comparison_log_table(self):
        if not hasattr(self, 'results_frame') or not self.results_frame.winfo_exists():
            self.after(100, self._populate_comparison_log_table)
            return

        loading = self._show_loading(self.results_frame)

        def _do_load():
            try:
                frame_height = self.results_frame.winfo_height()
                row_height = 28
                header_height = 40
                
                if frame_height > header_height:
                    num_visible_rows = (frame_height - header_height) // row_height
                else:
                    num_visible_rows = 1
                
                query = f"SELECT * FROM comparison_logs ORDER BY id DESC LIMIT {max(1, num_visible_rows)}"
                df = pd.read_sql_query(query, self.pg_engine)
                
                if not df.empty and 'summary_json' in df.columns:
                    new_cols = ['matched_records', 'diff_records']
                    for col in new_cols:
                        df[col] = 0

                    for index, row in df.iterrows():
                        summary_str = row['summary_json']
                        if summary_str and isinstance(summary_str, str):
                            try:
                                summary_data = json.loads(summary_str)
                                df.loc[index, 'matched_records'] = summary_data.get('matched_records', 0)
                                df.loc[index, 'diff_records'] = summary_data.get('diff_records', 0)
                            except (json.JSONDecodeError, TypeError):
                                pass
                    
                    df = df.drop(columns=['summary_json'])

                self.comparison_log_df = df
                if loading.winfo_exists(): loading.destroy()
                self.results_frame_label.configure(text=f"ประวัติการเปรียบเทียบข้อมูล ({len(df)} รายการล่าสุด)")
                self._create_styled_dataframe_table(self.results_frame, df, "")

            except Exception as e:
                if loading.winfo_exists(): loading.destroy()
                messagebox.showerror("Database Error", f"ไม่สามารถโหลดบันทึกการเปรียบเทียบได้: {e}", parent=self)
        
        self.after(50, _do_load)

    def _create_manage_users_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(1, weight=1); parent_tab.grid_rowconfigure(0, weight=1)
        manage_frame = CTkFrame(parent_tab, corner_radius=10); manage_frame.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="ns"); manage_frame.grid_columnconfigure(0, weight=1)
        
        CTkLabel(manage_frame, text="จัดการข้อมูลผู้ใช้งาน", font=self.header_font_table, text_color=self.theme["header"]).grid(row=0, column=0, pady=10, padx=20, sticky="w", columnspan=2)
        CTkLabel(manage_frame, text="User Key:", font=self.label_font).grid(row=1, column=0, padx=20, pady=(10, 2), sticky="w", columnspan=2); self.user_key_entry = CTkEntry(manage_frame, font=self.entry_font, width=250); self.user_key_entry.grid(row=2, column=0, padx=20, pady=5, sticky="ew", columnspan=2)
        CTkLabel(manage_frame, text="User Name:", font=self.label_font).grid(row=3, column=0, padx=20, pady=(10, 2), sticky="w", columnspan=2); self.user_name_entry = CTkEntry(manage_frame, font=self.entry_font); self.user_name_entry.grid(row=4, column=0, padx=20, pady=5, sticky="ew", columnspan=2)
        CTkLabel(manage_frame, text="Password:", font=self.label_font).grid(row=5, column=0, padx=20, pady=(10, 2), sticky="w", columnspan=2); self.password_entry = CTkEntry(manage_frame, font=self.entry_font, show="*"); self.password_entry.grid(row=6, column=0, padx=20, pady=5, sticky="ew", columnspan=2)
        CTkLabel(manage_frame, text="ประเภท:", font=self.label_font).grid(row=7, column=0, padx=20, pady=(10, 2), sticky="w", columnspan=2)
        
        # +++ START: แก้ไขบรรทัดนี้ +++
        # เพิ่ม 'Sale Support' เข้าไปในลิสต์ของ Role
        all_roles = ["Sale", "Sale Support", "Sales Manager", "Purchasing Staff", "Purchasing Manager", "Director", "HR"]
        
        self.role_var = tk.StringVar(value="Sale")
        self.role_menu = CTkOptionMenu(manage_frame, variable=self.role_var, values=all_roles, command=self._on_role_changed)
        # +++ END +++
        
        self.role_menu.grid(row=8, column=0, padx=20, pady=5, sticky="ew", columnspan=2)
        
        self.sale_type_var = tk.StringVar(value="Outbound"); self.sale_type_frame = CTkFrame(manage_frame, fg_color="transparent"); CTkLabel(self.sale_type_frame, text="ประเภท Sale:", font=self.label_font).pack(side="left", padx=(0, 10)); CTkRadioButton(self.sale_type_frame, text="Outbound", variable=self.sale_type_var, value="Outbound").pack(side="left", padx=5); CTkRadioButton(self.sale_type_frame, text="Inbound", variable=self.sale_type_var, value="Inbound").pack(side="left", padx=5)
        self.plan_var = tk.StringVar(value="Plan A"); self.plan_frame = CTkFrame(manage_frame, fg_color="transparent"); CTkLabel(self.plan_frame, text="แผนค่าคอมฯ:", font=self.label_font).pack(side="left", padx=(0, 10)); self.plan_menu = CTkOptionMenu(self.plan_frame, variable=self.plan_var, values=["Plan A", "Plan B", "Plan C", "Plan D"]); self.plan_menu.pack(side="left", expand=True, fill="x")
        self.sales_target_label = CTkLabel(manage_frame, text="ยอดเป้าหมาย:", font=self.label_font); self.sales_target_entry = NumericEntry(manage_frame, font=self.entry_font, placeholder_text="0.00")
        
        button_frame_1 = CTkFrame(manage_frame, fg_color="transparent"); button_frame_1.grid(row=13, column=0, padx=20, pady=(10, 5), sticky="ew", columnspan=2); button_frame_1.grid_columnconfigure((0, 1), weight=1)
        
        CTkButton(button_frame_1, text="เพิ่ม", command=self._add_user, fg_color=self.theme["primary"]).grid(row=0, column=0, padx=(0, 2), sticky="ew")
        CTkButton(button_frame_1, text="อัปเดต", command=self._update_user, fg_color="#006EFF").grid(row=0, column=1, padx=(2, 0), sticky="ew")
        
        button_frame_2 = CTkFrame(manage_frame, fg_color="transparent"); button_frame_2.grid(row=14, column=0, padx=20, pady=5, sticky="ew", columnspan=2); button_frame_2.grid_columnconfigure((0, 1), weight=1)
        
        CTkButton(button_frame_2, text="ปิดใช้งาน", command=self._deactivate_user, fg_color="#F97316", hover_color="#EA580C").grid(row=0, column=0, padx=(0, 2), sticky="ew")
        CTkButton(button_frame_2, text="เปิดใช้งาน", command=self._activate_user, fg_color="#EAB308", hover_color="#CA8A04").grid(row=0, column=1, padx=(2, 0), sticky="ew")
        
        button_frame_3 = CTkFrame(manage_frame, fg_color="transparent"); button_frame_3.grid(row=15, column=0, padx=20, pady=(10, 5), sticky="ew", columnspan=2)
        CTkButton(button_frame_3, text="ลบถาวร", command=self._permanent_delete_user, fg_color="#D32F2F", hover_color="#B71C1C").pack(fill="x")
        
        self.table_container = CTkFrame(parent_tab, corner_radius=10); self.table_container.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="nsew"); self.table_container.grid_rowconfigure(1, weight=1); self.table_container.grid_columnconfigure(0, weight=1)
        user_pagination_frame = CTkFrame(self.table_container, fg_color="transparent"); user_pagination_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5); self.user_prev_button = CTkButton(user_pagination_frame, text="<<", command=self._user_prev_page, width=50); self.user_prev_button.pack(side="left"); self.user_page_label = CTkLabel(user_pagination_frame, text="Page 1/1"); self.user_page_label.pack(side="left", padx=10); self.user_next_button = CTkButton(user_pagination_frame, text=">>", command=self._user_next_page, width=50); self.user_next_button.pack(side="left")
        
        self._on_role_changed()

    def _on_role_changed(self, selected_role=None):
        is_sale = self.role_var.get() == "Sale"
        widgets_map = {'sale_type_frame': (is_sale, 9), 'plan_frame': (is_sale, 10), 'sales_target_label': (is_sale, 11), 'sales_target_entry': (is_sale, 12)}
        for widget_name, (visible, row) in widgets_map.items():
            widget = getattr(self, widget_name, None)
            if widget and hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                if visible:
                    pady = (10, 2) if "label" in widget_name else 5
                    widget.grid(row=row, column=0, padx=20, pady=pady, sticky="ew", columnspan=2)
                else:
                    widget.grid_forget()
        if not is_sale:
            if hasattr(self, 'sales_target_entry'): self.sales_target_entry.delete(0, "end")
            if hasattr(self, 'sale_type_var'): self.sale_type_var.set("Outbound")
            if hasattr(self, 'plan_var'): self.plan_var.set("Plan A")
    
    def _user_prev_page(self):
        if self.user_current_page > 0: self.user_current_page -= 1; self._populate_users_table()
        
    def _user_next_page(self):
        total_pages = (self.user_total_rows + self.user_rows_per_page - 1) // self.user_rows_per_page
        if self.user_current_page < total_pages - 1: self.user_current_page += 1; self._populate_users_table()
        
    def _populate_users_table(self):
        table_frame = CTkFrame(self.table_container)
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        loading_label = self._show_loading(table_frame)
        try:
            count_query = "SELECT COUNT(*) FROM sales_users"
            self.user_total_rows = pd.read_sql(count_query, self.pg_engine).iloc[0,0]
            total_pages = (self.user_total_rows + self.user_rows_per_page - 1) // self.user_rows_per_page
            offset = self.user_current_page * self.user_rows_per_page
            
            data_query = "SELECT sale_key, sale_name, role, sale_type, commission_plan, sales_target, status FROM sales_users ORDER BY id DESC LIMIT %s OFFSET %s"
            self.user_df = pd.read_sql_query(data_query, self.pg_engine, params=(self.user_rows_per_page, offset))
            self.user_df.rename(columns={"sale_key": "User Key", "sale_name": "User Name", "role": "ประเภท", "sale_type": "ประเภท Sale", "commission_plan": "แผนค่าคอมฯ", "sales_target": "ยอดเป้าหมาย", "status": "สถานะ"}, inplace=True)
            
            loading_label.destroy()

            role_colors = {
                "Sale": "#EFF6FF", "Sales Manager": "#DBEAFE",
                "Purchasing Staff": "#F5F3FF", "Purchasing Manager": "#EDE9FE",
                "HR": "#F0FDF4", "Director": "#F3F4F6", "กรรมการ": "#F3F4F6"
            }

            # <<< START: เพิ่ม iid_column เข้าไปตรงนี้ >>>
            self._create_styled_dataframe_table(
                parent=table_frame, 
                df=self.user_df, 
                title="ข้อมูลผู้ใช้งาน", 
                on_row_click=self._on_user_row_click_treeview,
                status_column="ประเภท",
                status_colors=role_colors,
                iid_column='User Key' # <-- เพิ่มบรรทัดนี้
            )
            # <<< END >>>
            
            self.user_page_label.configure(text=f"Page {self.user_current_page + 1} / {max(1, total_pages)}")
            self.user_prev_button.configure(state="normal" if self.user_current_page > 0 else "disabled")
            self.user_next_button.configure(state="normal" if self.user_current_page < total_pages - 1 else "disabled")
        except Exception as e:
            if loading_label.winfo_exists(): loading_label.destroy()
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดผู้ใช้งานได้: {e}", parent=self)
    
    def _on_user_row_click_treeview(self, event, tree, df):
        try:
            record_id = tree.focus() # tree.focus() จะคืนค่า iid ที่เราตั้งไว้ (User Key)
            if not record_id:
                return

            # ค้นหาข้อมูลจาก DataFrame ด้วย User Key
            filtered_df = df.loc[df['User Key'] == record_id]

            if not filtered_df.empty:
                row_data = filtered_df.iloc[0]
                self._on_user_row_click(row_data) # ส่งข้อมูลไปแสดงผลในฟอร์ม
        except Exception as e:
            print(f"An error occurred in _on_user_row_click_treeview: {e}")

    def _on_user_row_click(self, row_data):
        self.user_key_entry.delete(0, tk.END); self.user_key_entry.insert(0, row_data.get("User Key", ""))
        self.user_name_entry.delete(0, tk.END); self.user_name_entry.insert(0, row_data.get("User Name", ""))
        self.role_var.set(row_data.get("ประเภท", "Sale"))
        self.password_entry.delete(0, tk.END); self.password_entry.configure(placeholder_text="ปล่อยว่างไว้หากไม่ต้องการเปลี่ยน")
        self.sale_type_var.set(row_data.get("ประเภท Sale", "Outbound"))
        self.plan_var.set(row_data.get("แผนค่าคอมฯ", "Plan A"))
        self.sales_target_entry.delete(0, tk.END)
        target_value = row_data.get("ยอดเป้าหมาย", 0.0)
        if pd.notna(target_value): self.sales_target_entry.insert(0, f"{target_value:,.2f}")
        self._on_role_changed()

    def _add_user(self):
        key, name, role, password = self.user_key_entry.get().strip(), self.user_name_entry.get().strip(), self.role_var.get(), self.password_entry.get().strip()
        if not key or not name or not password: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก User Key, User Name และ Password สำหรับผู้ใช้ใหม่", parent=self); return
        sale_type, commission_plan, sales_target = None, None, 0.0
        if role == "Sale":
            sale_type, commission_plan = self.sale_type_var.get(), self.plan_var.get()
            try: sales_target = float(self.sales_target_entry.get().replace(",", "") or 0.0)
            except ValueError: messagebox.showwarning("ข้อมูลผิดพลาด", "ยอดเป้าหมายต้องเป็นตัวเลขเท่านั้น", parent=self); return
        hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor: 
                cursor.execute("INSERT INTO sales_users (sale_key, sale_name, password_hash, role, sale_type, commission_plan, sales_target, status) VALUES (%s, %s, %s, %s, %s, %s, %s, 'Active')", (key, name, hashed_password.decode('utf-8'), role, sale_type, commission_plan, sales_target))
                cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, new_value) VALUES (%s, %s, %s, %s, %s)",
                               ('Add', 'sales_users', None, self.user_key, json.dumps({'sale_key': key, 'sale_name': name, 'role': role})))
            conn.commit(); messagebox.showinfo("สำเร็จ", "เพิ่มผู้ใช้งานเรียบร้อยแล้ว", parent=self); self._refresh_all_data_views()
            self.user_key_entry.delete(0, "end"); self.user_name_entry.delete(0, "end"); self.sales_target_entry.delete(0, "end"); self.password_entry.delete(0, "end")
        except psycopg2.errors.UniqueViolation:
            if conn: conn.rollback(); messagebox.showerror("ผิดพลาด", "User Key นี้มีอยู่ในระบบแล้ว", parent=self)
        except Exception as e:
            if conn: conn.rollback(); messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}", parent=self)
        finally: self.app_container.release_connection(conn)

    def _update_user(self):
        key, name, role, password = self.user_key_entry.get().strip(), self.user_name_entry.get().strip(), self.role_var.get(), self.password_entry.get().strip()
        if not key or not name: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก User Key และ User Name", parent=self); return
        sale_type, commission_plan, sales_target = None, None, 0.0
        if role == "Sale":
            sale_type, commission_plan = self.sale_type_var.get(), self.plan_var.get()
            try: sales_target = float(self.sales_target_entry.get().replace(",", "") or 0.0)
            except ValueError: messagebox.showwarning("ข้อมูลผิดพลาด", "ยอดเป้าหมายต้องเป็นตัวเลขเท่านั้น", parent=self); return
        
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM sales_users WHERE sale_key = %s", (key,))
                old_data_row = cursor.fetchone()
                if not old_data_row:
                    messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบ User Key ที่ต้องการอัปเดต", parent=self); return
                
                old_data = dict(old_data_row)

                update_cols = ["sale_name", "role", "sale_type", "commission_plan", "sales_target"]
                update_values = [name, role, sale_type, commission_plan, sales_target]
                
                if password:
                    hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
                    update_cols.append("password_hash")
                    update_values.append(hashed_password.decode('utf-8'))

                set_clause = ", ".join([f"{col} = %s" for col in update_cols])
                sql_query = f"UPDATE sales_users SET {set_clause} WHERE sale_key = %s"
                cursor.execute(sql_query, (*update_values, key))
                
                if cursor.rowcount == 0:
                    messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบ User Key ที่ต้องการอัปเดต", parent=self)
                else: 
                    conn.commit()
                    cursor.execute("SELECT * FROM sales_users WHERE sale_key = %s", (key,))
                    new_data_row = cursor.fetchone()
                    new_data = dict(new_data_row)
                    
                    changes = {k: new_data[k] for k in new_data if k in old_data and new_data[k] != old_data[k]}
                    
                    old_value_json = json.dumps(old_data, default=str)
                    new_value_json = json.dumps(new_data, default=str)
                    changes_json = json.dumps(changes, default=str)
                    
                    cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, old_value, new_value, changes) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                                   ('Update', 'sales_users', int(old_data['id']), self.user_key, old_value_json, new_value_json, changes_json))
                    conn.commit()
                    messagebox.showinfo("สำเร็จ", "อัปเดตข้อมูลเรียบร้อย", parent=self)
            self._refresh_all_data_views()
        except psycopg2.errors.LockNotAvailable as e:
            if conn: conn.rollback()
            messagebox.showwarning("รายการถูกล็อค", "ไม่สามารถอัปเดตข้อมูลได้ในขณะนี้", parent=self)
        except Exception as e:
            if conn: conn.rollback()
            print("--- TRACEBACK START ---")
            print(traceback.format_exc())
            print("--- TRACEBACK END ---")
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}", parent=self)
        finally:
            self.app_container.release_connection(conn)

    def _update_user_status(self, key, status):
        if not key: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก User Key หรือเลือกจากตาราง", parent=self); return
        action_text = "เปิดการใช้งาน" if status == "Active" else "ปิดการใช้งาน"
        if messagebox.askyesno("ยืนยัน", f"คุณแน่ใจหรือไม่ที่จะ{action_text}ผู้ใช้งาน: {key}?", parent=self):
            conn = None
            try:
                conn = self.app_container.get_connection()
                with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                    cursor.execute("SELECT id, status FROM sales_users WHERE sale_key = %s", (key,))
                    old_data = cursor.fetchone()
                    if not old_data: messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบ User Key ที่ต้องการอัปเดต", parent=self); return
                    
                    cursor.execute("UPDATE sales_users SET status = %s WHERE sale_key = %s", (status, key))
                    if cursor.rowcount == 0: messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบ User Key ที่ต้องการอัปเดต", parent=self)
                    else: 
                        conn.commit()
                        cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, old_value, new_value, changes) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                                       (action_text, 'sales_users', old_data['id'], self.user_key, json.dumps({'status': old_data['status']}), json.dumps({'status': status}), json.dumps({'status': status})))
                        conn.commit()
                        messagebox.showinfo("สำเร็จ", f"{action_text}ผู้ใช้งานเรียบร้อยแล้ว", parent=self)
                self._refresh_all_data_views()
            except psycopg2.errors.LockNotAvailable as e:
                if conn: conn.rollback(); messagebox.showwarning("รายการถูกล็อค", "ไม่สามารถอัปเดตสถานะได้ในขณะนี้", parent=self)
            except Exception as e:
                if conn: conn.rollback(); messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}", parent=self)
            finally: self.app_container.release_connection(conn)

    def _deactivate_user(self): self._update_user_status(self.user_key_entry.get().strip(), 'Inactive')
    def _activate_user(self): self._update_user_status(self.user_key_entry.get().strip(), 'Active')
    
    def _permanent_delete_user(self):
        key = self.user_key_entry.get().strip()
        if not key: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก User Key ที่ต้องการลบถาวร", parent=self); return
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT id, status FROM sales_users WHERE sale_key = %s", (key,)); result = cursor.fetchone()
                if not result: messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบผู้ใช้งาน User Key: {key}", parent=self); return
                if result['status'] == 'Active': messagebox.showerror("เงื่อนไขไม่ถูกต้อง", "ไม่สามารถลบผู้ใช้งานที่ยัง 'Active' ได้\nกรุณา 'ปิดการใช้งาน' ผู้ใช้ก่อน", parent=self); return
                
                cursor.execute("SELECT 1 FROM commissions WHERE sale_key = %s LIMIT 1", (key,)); commission_history = cursor.fetchone()
                cursor.execute("SELECT 1 FROM purchase_orders WHERE user_key = %s LIMIT 1", (key,)); po_history = cursor.fetchone()
                
                if commission_history or po_history: 
                    messagebox.showerror("ไม่สามารถลบได้", "ผู้ใช้งานนี้มีประวัติการทำรายการอยู่ในระบบ\nไม่สามารถลบถาวรได้", parent=self)
                    return
                
                if messagebox.askyesno("ยืนยันการลบถาวร", f"คุณแน่ใจจริงๆ หรือไม่ที่จะลบผู้ใช้งาน '{key}' ออกจากระบบอย่างถาวร?\n**การกระทำนี้ไม่สามารถย้อนกลับได้!**", icon="warning", parent=self):
                    cursor.execute("DELETE FROM sales_users WHERE sale_key = %s", (key,)); 
                    conn.commit()
                    cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, old_value) VALUES (%s, %s, %s, %s, %s)",
                                   ('Delete Permanent', 'sales_users', result['id'], self.user_key, json.dumps({'sale_key': key})))
                    conn.commit()
                    messagebox.showinfo("สำเร็จ", f"ผู้ใช้งาน '{key}' ถูกลบออกจากระบบอย่างถาวรแล้ว", parent=self); self._refresh_all_data_views(); self.user_key_entry.delete(0, "end"); self.user_name_entry.delete(0, "end"); self.sales_target_entry.delete(0, "end")
        except psycopg2.errors.LockNotAvailable as e:
            if conn: conn.rollback(); messagebox.showwarning("รายการถูกล็อค", "ไม่สามารถลบข้อมูลได้ในขณะนี้", parent=self)
        except Exception as e:
            if conn: conn.rollback(); messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดที่ไม่คาดคิดระหว่างการลบ: {e}", parent=self)
        finally: self.app_container.release_connection(conn)

    def _refresh_all_data_views(self):
        if self._users_loaded: self._populate_users_table()
        if self._dashboard_loaded: self._update_dashboard()
        if self._sales_target_loaded: self._update_sales_target_dashboard()
        
        active_sales = self._get_active_sales_list()
        if hasattr(self, 'sale_process_dropdown'): self.sale_process_dropdown.configure(values=active_sales)
        
        if self._audit_log_loaded: self._populate_audit_log_table()
        if self._compare_commission_loaded and hasattr(self, 'results_frame') and self.results_frame.winfo_exists():
            for widget in self.results_frame.winfo_children(): widget.destroy()
            self.results_frame_label.configure(text="กรุณากด 'เริ่มต้นการเปรียบเทียบใหม่' เพื่อเริ่มใช้งาน")
            if hasattr(self, 'finalize_button'): self.finalize_button.pack_forget()
            if hasattr(self, 'export_button'): self.export_button.pack_forget()

        if self._process_commission_loaded: self._on_sale_selected_for_process()

    def _get_sale_keys(self):
        try: return pd.read_sql("SELECT sale_key FROM sales_users WHERE role = 'Sale' AND status = 'Active' ORDER BY sale_key", self.pg_engine)["sale_key"].tolist()
        except Exception as e: print(f"Error getting sale keys: {e}"); messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลรหัสพนักงานขายได้: {e}", parent=self); return []

    def _create_compare_commission_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(2, weight=1)

        # --- Frame ควบคุมด้านบน ---
        control_frame = CTkFrame(parent_tab)
        control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        CTkButton(control_frame, text="🚀 เริ่มต้นการเปรียบเทียบใหม่", command=self._start_new_comparison, font=CTkFont(size=16, weight="bold")).pack(side="left", padx=10, pady=10)
        CTkButton(control_frame, text="📖 แสดงประวัติการเปรียบเทียบ", command=self._open_comparison_history_window, fg_color="#64748B").pack(side="left", padx=10, pady=10)
        
        

        # --- [✨ ส่วนที่เพิ่มใหม่] Label แสดงชื่อ Sale ---
        self.active_sale_label = CTkLabel(control_frame, 
                                          text="สถานะ: ยังไม่ได้เลือกพนักงาน", 
                                          font=CTkFont(size=16, weight="bold"), 
                                          text_color="gray")
        self.active_sale_label.pack(side="left", padx=20)
        # -----------------------------------------------

        self.verify_passed_button = CTkButton(control_frame, text="ยืนยัน SO ที่ผ่านเกณฑ์ (Verify Passed SOs)", fg_color="#16A34A", hover_color="#15803D", command=self._verify_passed_sos)
        self.verify_passed_button.pack(side="right", padx=10, pady=10)
        self.verify_passed_button.pack_forget()

        self.export_button = CTkButton(control_frame, text="📄 Export ผลลัพธ์", command=self._export_comparison)
        self.export_button.pack(side="right", padx=10, pady=10)
        self.export_button.pack_forget()

        self.results_frame_label = CTkLabel(parent_tab, text="กรุณากด 'เริ่มต้นการเปรียบเทียบใหม่' เพื่อเริ่มใช้งาน", font=self.label_font, text_color="gray")
        self.results_frame_label.grid(row=1, column=0, padx=10, pady=(5, 0), sticky="w")
        
        # --- Frame สำหรับตาราง (แสดงผลเต็มพื้นที่) ---
        self.results_frame = CTkFrame(parent_tab)
        self.results_frame.grid(row=2, column=0, pady=(0, 10), padx=10, sticky="nsew")
        self.results_frame.grid_rowconfigure(0, weight=1)
        self.results_frame.grid_columnconfigure(0, weight=1)
    
    def _open_cancelled_history(self):
        """เปิดหน้าต่างดูประวัติ SO ที่ถูกยกเลิก"""
        from history_windows import CancelledHistoryWindow
        try:
            CancelledHistoryWindow(self, self.app_container)
        except Exception as e:
            tk.messagebox.showerror("Error", f"ไม่สามารถเปิดหน้าต่างได้: {e}")

    def _update_summary_pane(self):
        # +++ START: เพิ่ม Safety Check ตรงนี้ +++
        # ตรวจสอบก่อนว่า self.summary_pane ถูกสร้างขึ้นแล้วหรือยัง
        if not hasattr(self, 'summary_pane') or not self.summary_pane.winfo_exists():
            return # ถ้ายังไม่มี ให้หยุดการทำงานของฟังก์ชันนี้ไปเลย
        # +++ END +++

        # ล้างข้อมูลเก่าใน Pane
        for widget in self.summary_pane.winfo_children():
            widget.destroy()

        if self.comparison_df is None or self.comparison_df.empty:
            CTkLabel(self.summary_pane, text="ไม่มีข้อมูลสรุป").pack(expand=True)
            return
            
        # กรองเอาเฉพาะแถวข้อมูล ไม่รวมแถว 'ยอดรวม (Total)'
        df_data = self.comparison_df[self.comparison_df['เลขที่ SO'] != 'ยอดรวม (Total)']

        # --- ส่วนแสดงตัวเลขสรุป (KPIs) ---
        kpi_frame = CTkFrame(self.summary_pane, fg_color="transparent")
        kpi_frame.pack(fill="x", padx=15, pady=15)
        kpi_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(kpi_frame, text="ภาพรวมการเปรียบเทียบ", font=CTkFont(size=18, weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        kpi_font = CTkFont(size=14)
        kpi_value_font = CTkFont(size=14, weight="bold")

        total_sales = df_data['ยอดขายรวม (ระบบ)'].sum()
        total_cost = df_data['ต้นทุน (ระบบ)'].sum()
        total_diff_sales = df_data['ผลต่างยอดขาย'].sum()

        CTkLabel(kpi_frame, text="จำนวน SO ทั้งหมด:", font=kpi_font).grid(row=1, column=0, sticky="w", pady=2)
        CTkLabel(kpi_frame, text=f"{len(df_data)} รายการ", font=kpi_value_font).grid(row=1, column=1, sticky="e", pady=2)
        
        CTkLabel(kpi_frame, text="ยอดขายรวม (ระบบ):", font=kpi_font).grid(row=2, column=0, sticky="w", pady=2)
        CTkLabel(kpi_frame, text=f"{total_sales:,.2f}", font=kpi_value_font).grid(row=2, column=1, sticky="e", pady=2)

        CTkLabel(kpi_frame, text="ต้นทุนรวม (ระบบ):", font=kpi_font).grid(row=3, column=0, sticky="w", pady=2)
        CTkLabel(kpi_frame, text=f"{total_cost:,.2f}", font=kpi_value_font).grid(row=3, column=1, sticky="e", pady=2)
        
        diff_color = "red" if total_diff_sales < 0 else "green"
        CTkLabel(kpi_frame, text="ผลต่างยอดขายรวม:", font=kpi_font).grid(row=4, column=0, sticky="w", pady=2)
        CTkLabel(kpi_frame, text=f"{total_diff_sales:,.2f}", font=kpi_value_font, text_color=diff_color).grid(row=4, column=1, sticky="e", pady=2)

        # --- ส่วนแสดงกราฟวงกลม ---
        chart_frame = CTkFrame(self.summary_pane, fg_color="transparent")
        chart_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        status_counts = df_data['สถานะ'].value_counts()
        
        if not status_counts.empty:
            fig = Figure(figsize=(5, 4), dpi=100, facecolor="#F9FAFB")
            ax = fig.add_subplot(111)
            
            status_colors = {
                "ผ่านเกณฑ์": "#BBF7D0", "ยอดขายต่ำกว่า Express": "#FECACA", "ต้นทุนต่ำกว่า Express": "#FEF08A",
                "‼️ ขายขาดทุน (ตรวจสอบด่วน)": "#F87171", "‼️ ต้นทุน Express ผิดปกติ (<50%)": "#F97316",
                "ข้อมูลไม่ตรงกัน": "#FED7AA", "มีใน Express, ไม่มีในระบบ": "#E5E7EB", "มีในระบบ, ไม่มีใน Express": "#E5E7EB",
                "กำไรดี": "#16A34A", "กำไรน้อย": "#FCD34D", "ขาดทุน": "#EF4444", "ยืนยันแล้ว (รอผล)": "#9CA3AF",
            }
            colors = [status_colors.get(status, "#D1D5DB") for status in status_counts.index]

            ax.pie(status_counts, labels=status_counts.index, autopct='%1.1f%%', startangle=90, colors=colors, 
                   textprops={'fontsize': 10, 'fontfamily': 'Tahoma'})
            ax.set_title('สรุปสถานะรายการ', fontname='Tahoma', fontsize=14, weight="bold")
            fig.tight_layout()

            if hasattr(self, 'summary_pie_chart_canvas'):
                self.summary_pie_chart_canvas.get_tk_widget().destroy()

            self.summary_pie_chart_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            self.summary_pie_chart_canvas.draw()
            self.summary_pie_chart_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _start_new_comparison(self):
        active_sales_keys = self._get_sale_keys()
        config_dialog = ComparisonConfigDialog(self, sales_keys=active_sales_keys)
        self.wait_window(config_dialog)

        if not config_dialog.result:
            return

        config = config_dialog.result
        selected_salesperson = config["salesperson"]

        if hasattr(self, 'active_sale_label'):
             self.active_sale_label.configure(
                 text=f"กำลังตรวจสอบข้อมูลของ: {selected_salesperson}", 
                 text_color="#2563EB" # สีน้ำเงินให้เด่น
             )
        selected_month = config["month"]
        selected_year = config["year"]
        
        # บันทึกค่าไว้ใช้ตอน Refresh
        self.current_comparison_month = selected_month
        self.current_comparison_year = selected_year
        self.current_comparison_salesperson = selected_salesperson
        self.uploaded_df = config["imported_df"]
        self.manual_entry_df = config["manual_df"]

        # จัดการ mapping ชื่อคอลัมน์ของไฟล์ Excel (เหมือนเดิม)
        if self.uploaded_df is not None:
            self.uploaded_df.columns = [str(c).lower().strip() for c in self.uploaded_df.columns]

            column_mapping_options = {
                'so_number': ['so number', 'so_number', 'so no.', 'เลขที่ so', 'อ้างถึง'],
                'sales_uploaded': ['sales_service_amount', 'sales', 'amount', 'ยอดขาย*', 'ยอดขาย', 'ยอดขาย/บริการ', 'ยอดขายรวม'],
                'shipping_cost_uploaded': ['shipping_cost', 'ค่าจัดส่ง', 'ต้นทุนค่าจัดส่ง'],
                'relocation_cost_uploaded': ['relocation_cost', 'ค่าย้าย', 'ต้นทุนค่าย้าย'],
                'total_payment_amount': ['total_payment_amount', 'total payment', 'ยอดชำระ', 'ยอดชำระรวม'],
                'payment_date': ['payment_date', 'pay date', 'วันที่ชำระ'],
                'cost_uploaded': ['cost', 'cogs', 'ต้นทุน', 'ต้นทุนค่าสินค้า/บริการ', 'ต้นทุนรวม'],
                'brokerage_fee_uploaded': ['brokerage_fee', 'ค่านายหน้า', 'ต้นทุนค่านายหน้า'],
                'transfer_fee_uploaded': ['transfer_fee', 'ค่าธรรมเนียมโอน', 'ต้นทุนค่าธรรมเนียมโอน']
            }
            
            rename_map = {}
            for standard_name, possible_names in column_mapping_options.items():
                found_col = next((c for c in self.uploaded_df.columns if c in possible_names), None)
                if found_col:
                    rename_map[found_col] = standard_name
            
            self.uploaded_df.rename(columns=rename_map, inplace=True)
        
        loading = self._show_loading(self.results_frame)
        self.results_frame_label.configure(text=f"กำลังโหลดข้อมูลจากฐานข้อมูลสำหรับ: {selected_salesperson}...")
        
        try:
            self.current_comparison_salesperson = selected_salesperson
            
            # --- แก้ไข Query หลัก: ปรับเงื่อนไข Status และ Time ---
            base_query = """SELECT c.*, 
                       po.cogs_db, po.po_shipping_stock, po.po_shipping_site, po.po_relocation, 
                       u.sale_name,
                       ss.sale_name as support_user_name 
                FROM commissions c 
                JOIN sales_users u ON c.sale_key = u.sale_key
                LEFT JOIN sales_users ss ON c.support_user_key = ss.sale_key
                LEFT JOIN (
                        SELECT
                            p.so_number,
                            SUM(COALESCE(poi.total_price, 0)) as cogs_db,
                            SUM(p.shipping_to_stock_cost) as po_shipping_stock,
                            SUM(p.shipping_to_site_cost) as po_shipping_site,
                            SUM(p.relocation_cost) as po_relocation
                        FROM purchase_orders p
                        LEFT JOIN purchase_order_items poi ON p.id = poi.purchase_order_id
                        WHERE p.status = 'Approved'
                        GROUP BY p.so_number
                    ) po ON c.so_number = po.so_number
                WHERE c.is_active = 1
                  AND c.status NOT IN ('Paid', 'Cancelled', 'HR Verified')
            """

            params = []

            # --- Filter พนักงานขาย ---
            if selected_salesperson != "ทั้งหมด":
                base_query += " AND c.sale_key = %s"
                params.append(selected_salesperson)
            else:
                # กรณีเลือกทั้งหมด
                base_query += " AND c.sale_key IN (SELECT sale_key FROM sales_users WHERE status = 'Active' AND role = 'Sale')"
            
            # --- ✅ จุดแก้ไขที่ 2: เปลี่ยนเงื่อนไขเวลาเป็น 'ย้อนหลังทั้งหมด' ---
            # Logic: (ปีก่อนหน้า) OR (ปีเดียวกัน แต่เดือนน้อยกว่าหรือเท่ากับเดือนที่เลือก)
            base_query += " AND ((c.commission_year < %s) OR (c.commission_year = %s AND c.commission_month <= %s))"
            params.extend([selected_year, selected_year, selected_month])

            data_query = base_query + " ORDER BY c.timestamp DESC"
            
            self.db_df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(params))
            
            if loading.winfo_exists(): loading.destroy()
            
            self._compare_data()

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            print(traceback.format_exc())
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูลได้: {e}", parent=self)
            
    def _verify_passed_sos(self):
        if self.comparison_df is None or self.comparison_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลการเปรียบเทียบที่จะยืนยัน", parent=self)
            return

        df_to_verify = self.comparison_df[self.comparison_df['สถานะ'] == 'ผ่านเกณฑ์']
        so_numbers_to_verify = tuple(df_to_verify['เลขที่ SO'].tolist())

        if not so_numbers_to_verify:
            messagebox.showinfo("ไม่พบรายการ", "ไม่พบรายการที่ 'ผ่านเกณฑ์' ที่จะยืนยันได้ในขณะนี้", parent=self)
            return

        selected_month = getattr(self, 'current_comparison_month', None)
        selected_year = getattr(self, 'current_comparison_year', None)

        if not selected_month or not selected_year:
            messagebox.showerror("Error", "ไม่สามารถระบุเดือนที่กำลังเปรียบเทียบได้ โปรดเริ่มการเปรียบเทียบใหม่", parent=self)
            return

        msg = (f"คุณต้องการยืนยันข้อมูลสำหรับ {len(so_numbers_to_verify)} รายการที่ผ่านเกณฑ์ใช่หรือไม่?\n\n"
               f"โปรแกรมจะดึงข้อมูลล่าสุดมาคำนวณใหม่ และจัดรอบค่าคอมเป็น {selected_month}/{selected_year} ก่อนบันทึก")
        if not messagebox.askyesno("ยืนยันข้อมูล", msg, parent=self):
            return

        records_to_update = []
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                placeholders = ', '.join(['%s'] * len(so_numbers_to_verify))
                query = f"""
                    SELECT c.*, (COALESCE(po_items.total_item_cost, 0) - COALESCE(po_discounts.total_bill_discount, 0)) as cogs_db
                    FROM commissions c
                    LEFT JOIN (
                        SELECT p.so_number, SUM(COALESCE(poi.total_price, 0)) as total_item_cost
                        FROM purchase_orders p
                        LEFT JOIN purchase_order_items poi ON p.id = poi.purchase_order_id
                        WHERE p.status = 'Approved' AND p.so_number IN ({placeholders})
                        GROUP BY p.so_number
                    ) po_items ON c.so_number = po_items.so_number
                    LEFT JOIN (
                        SELECT so_number, SUM(COALESCE(bill_discount, 0)) as total_bill_discount
                        FROM purchase_orders
                        WHERE status = 'Approved' AND so_number IN ({placeholders})
                        GROUP BY so_number
                    ) po_discounts ON c.so_number = po_discounts.so_number
                    WHERE c.so_number IN ({placeholders}) AND c.is_active = 1
                """
                cursor.execute(query, so_numbers_to_verify * 3)
                latest_so_data = cursor.fetchall()

                for row_data in latest_so_data:
                    final_sale = (float(row_data.get('sales_service_amount', 0) or 0) +
                                  float(row_data.get('cutting_drilling_fee', 0) or 0) +
                                  float(row_data.get('other_service_fee', 0) or 0))
                    
                    final_cost = float(row_data.get('cogs_db', 0) or 0)
                    final_gp = final_sale - final_cost
                    final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
                    
                    records_to_update.append((int(row_data['id']), final_sale, final_cost, final_gp, final_margin))

            if not records_to_update:
                messagebox.showerror("ผิดพลาด", "ไม่สามารถเตรียมข้อมูลสำหรับอัปเดตได้", parent=self)
                return

            with conn.cursor() as cursor:
                update_query = f"""
                    UPDATE commissions SET 
                        status = 'HR Verified', 
                        final_sales_amount = data.final_sale,
                        final_cost_amount = data.final_cost,
                        final_gp = data.final_gp,
                        final_margin = data.final_margin,
                        payout_id = NULL,
                        commission_month = {selected_month},
                        commission_year = {selected_year}
                    FROM (VALUES %s) AS data(record_id, final_sale, final_cost, final_gp, final_margin)
                    WHERE commissions.id = data.record_id;
                """
                psycopg2.extras.execute_values(cursor, update_query, records_to_update,
                    template="(%s::int, %s::float, %s::float, %s::float, %s::float)", page_size=100)
                updated_rows = cursor.rowcount
            conn.commit()
            
            # 🔥 [จุดแก้ที่ 4] จำชื่อเซลส์และเดือนที่ทำเสร็จ
            self.last_verified_period = f"{self.thai_months[int(selected_month)-1]} {int(selected_year)+543}"
            self.last_verified_sale = self.current_comparison_salesperson

            messagebox.showinfo("สำเร็จ", f"ยืนยันข้อมูล {updated_rows} รายการเรียบร้อยแล้ว\n(อัปเดตรอบค่าคอมเป็นเดือน {selected_month}/{selected_year})", parent=self)
            self._refresh_comparison_view()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)


    def _verify_passed_sos(self):
        if self.comparison_df is None or self.comparison_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลการเปรียบเทียบที่จะยืนยัน", parent=self)
            return

        df_to_verify = self.comparison_df[self.comparison_df['สถานะ'] == 'ผ่านเกณฑ์']
        so_numbers_to_verify = tuple(df_to_verify['เลขที่ SO'].tolist())

        if not so_numbers_to_verify:
            messagebox.showinfo("ไม่พบรายการ", "ไม่พบรายการที่ 'ผ่านเกณฑ์' ที่จะยืนยันได้ในขณะนี้", parent=self)
            return

        selected_month = getattr(self, 'current_comparison_month', None)
        selected_year = getattr(self, 'current_comparison_year', None)

        if not selected_month or not selected_year:
            messagebox.showerror("Error", "ไม่สามารถระบุเดือนที่กำลังเปรียบเทียบได้ โปรดเริ่มการเปรียบเทียบใหม่", parent=self)
            return

        msg = (f"คุณต้องการยืนยันข้อมูลสำหรับ {len(so_numbers_to_verify)} รายการที่ผ่านเกณฑ์ใช่หรือไม่?\n\n"
               f"โปรแกรมจะดึงข้อมูลล่าสุดมาคำนวณใหม่ และจัดรอบค่าคอมเป็น {selected_month}/{selected_year} ก่อนบันทึก")
        if not messagebox.askyesno("ยืนยันข้อมูล", msg, parent=self):
            return

        records_to_update = []
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                placeholders = ', '.join(['%s'] * len(so_numbers_to_verify))
                query = f"""
                    SELECT c.*, (COALESCE(po_items.total_item_cost, 0) - COALESCE(po_discounts.total_bill_discount, 0)) as cogs_db
                    FROM commissions c
                    LEFT JOIN (
                        SELECT p.so_number, SUM(COALESCE(poi.total_price, 0)) as total_item_cost
                        FROM purchase_orders p
                        LEFT JOIN purchase_order_items poi ON p.id = poi.purchase_order_id
                        WHERE p.status = 'Approved' AND p.so_number IN ({placeholders})
                        GROUP BY p.so_number
                    ) po_items ON c.so_number = po_items.so_number
                    LEFT JOIN (
                        SELECT so_number, SUM(COALESCE(bill_discount, 0)) as total_bill_discount
                        FROM purchase_orders
                        WHERE status = 'Approved' AND so_number IN ({placeholders})
                        GROUP BY so_number
                    ) po_discounts ON c.so_number = po_discounts.so_number
                    WHERE c.so_number IN ({placeholders}) AND c.is_active = 1
                """
                cursor.execute(query, so_numbers_to_verify * 3)
                latest_so_data = cursor.fetchall()

                for row_data in latest_so_data:
                    final_sale = (float(row_data.get('sales_service_amount', 0) or 0) +
                                  float(row_data.get('cutting_drilling_fee', 0) or 0) +
                                  float(row_data.get('other_service_fee', 0) or 0))
                    
                    final_cost = float(row_data.get('cogs_db', 0) or 0)
                    final_gp = final_sale - final_cost
                    final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
                    
                    records_to_update.append((int(row_data['id']), final_sale, final_cost, final_gp, final_margin))

            if not records_to_update:
                messagebox.showerror("ผิดพลาด", "ไม่สามารถเตรียมข้อมูลสำหรับอัปเดตได้", parent=self)
                return

            with conn.cursor() as cursor:
                update_query = f"""
                    UPDATE commissions SET 
                        status = 'HR Verified', 
                        final_sales_amount = data.final_sale,
                        final_cost_amount = data.final_cost,
                        final_gp = data.final_gp,
                        final_margin = data.final_margin,
                        payout_id = NULL,
                        commission_month = {selected_month},
                        commission_year = {selected_year}
                    FROM (VALUES %s) AS data(record_id, final_sale, final_cost, final_gp, final_margin)
                    WHERE commissions.id = data.record_id;
                """
                psycopg2.extras.execute_values(cursor, update_query, records_to_update,
                    template="(%s::int, %s::float, %s::float, %s::float, %s::float)", page_size=100)
                updated_rows = cursor.rowcount
            conn.commit()
            
            # 🔥 [เพิ่มตรงนี้] สร้างตัวแปรเก็บรอบที่เพิ่งตรวจสอบเสร็จ
            month_name = self.thai_months[int(selected_month) - 1]
            year_buddhist = int(selected_year) + 543
            self.app_container.last_verified_period = f"{month_name} {year_buddhist}"
            self.app_container.last_verified_sale = so_numbers_to_verify[0] # เก็บเป็นเบาะแสไว้ (ไม่ใช้ก็ไม่เป็นไร)

            messagebox.showinfo("สำเร็จ", f"ยืนยันข้อมูล {updated_rows} รายการเรียบร้อยแล้ว\n(อัปเดตรอบค่าคอมเป็นเดือน {selected_month}/{selected_year})", parent=self)
            self._refresh_comparison_view()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _calculate_final_pu_cost(self, row):
        """
        (เวอร์ชันปรับปรุงตาม Feedback)
        คำนวณต้นทุนรวมสุดท้ายจากฝั่งระบบ (PU) โดยจะใช้เฉพาะยอดรวมค่าสินค้าจาก PO เท่านั้น
        """
        overrides = {}
        if pd.notna(row.get('hr_cost_overrides')):
            try:
                overrides = json.loads(row['hr_cost_overrides']) if isinstance(row['hr_cost_overrides'], str) else row['hr_cost_overrides']
            except (json.JSONDecodeError, TypeError):
                pass
        
        # --- START: แก้ไข Logic การคำนวณต้นทุน ---
        # 1. ดึงต้นทุนสินค้า (cogs_db คือ SUM(total_cost) จาก PO) มาเป็นค่าตั้งต้น
        total_system_cost = float(row.get('cogs_db', 0) or 0)
        
        # 2. ใช้ค่าที่ HR แก้ไขเอง (Overrides) ถ้ามี, ไม่เช่นนั้นก็ใช้ค่าจาก PO โดยตรง
        # ไม่นำค่าใช้จ่ายอื่นๆ (brokerage, transfer, giveaways) จาก SO มารวมแล้ว
        final_cost = float(overrides.get('ต้นทุนรวม', total_system_cost))
        # --- END ---
        
        return final_cost


    # hr_screen.py (ฟังก์ชัน _compare_data ที่แก้ไขแล้ว)

    def _compare_data(self):
        try:
            # 1. รวบรวมข้อมูล
            comparison_sources = []
            if self.uploaded_df is not None and not self.uploaded_df.empty:
                comparison_sources.append(self.uploaded_df)
            if self.manual_entry_df is not None and not self.manual_entry_df.empty:
                comparison_sources.append(self.manual_entry_df)
            
            if not comparison_sources:
                messagebox.showwarning("ไม่มีข้อมูลเปรียบเทียบ", "กรุณา Import ไฟล์ หรือ คีย์ข้อมูลด้วยมือ", parent=self)
                self._create_styled_dataframe_table(self.results_frame, self.db_df)
                return
                
            valid_sources = [df for df in comparison_sources if 'so_number' in df.columns]
            if not valid_sources:
                found_cols = []
                for df in comparison_sources:
                    found_cols.extend(list(df.columns))
                col_list = ', '.join(found_cols) if found_cols else '(ไม่มี)'
                messagebox.showwarning(
                    "ข้อมูลไม่ถูกต้อง",
                    f"ไม่พบคอลัมน์ 'so_number' ในข้อมูลที่นำเข้า\n\n"
                    f"คอลัมน์ที่พบในไฟล์: {col_list}\n\n"
                    f"ชื่อคอลัมน์ที่ระบบรับได้: so number, so_number, so no., เลขที่ so, อ้างถึง",
                    parent=self
                )
                return
            uploaded_compare_df = pd.concat(valid_sources, ignore_index=True).drop_duplicates(subset=['so_number'], keep='last')

            processed_so_query = "SELECT so_number FROM commissions WHERE status IN ('HR Verified', 'Paid', 'Deferred by HR', 'Cancelled')"
            processed_so_df = pd.read_sql_query(processed_so_query, self.pg_engine)
            
            if not processed_so_df.empty:
                processed_so_list = processed_so_df['so_number'].tolist()
                uploaded_compare_df['so_number'] = uploaded_compare_df['so_number'].astype(str).str.strip()
                uploaded_compare_df = uploaded_compare_df[~uploaded_compare_df['so_number'].isin(processed_so_list)]

            db_compare_df = self.db_df.copy()
            db_compare_df['so_number'] = db_compare_df['so_number'].astype(str).str.strip()
            
            # --- [Logic 1] คำนวณ Grand Total (ยอดบิล) สำรองไว้ กรณีใน DB เป็น 0 ---
            # รวม: สินค้า + บริการ + ขนส่ง + ค่าย้าย + บัตรเครดิต
            cols_to_sum = ['sales_service_amount', 'cutting_drilling_fee', 'other_service_fee', 
                           'shipping_cost', 'relocation_cost', 'credit_card_fee']
            
            db_compare_df['calc_base_total'] = 0.0
            for col in cols_to_sum:
                if col in db_compare_df.columns:
                    db_compare_df['calc_base_total'] += pd.to_numeric(db_compare_df[col], errors='coerce').fillna(0)
            
            # คูณ VAT 7% เพื่อหา "ยอดที่ต้องชำระจริง"
            db_compare_df['calc_grand_total'] = db_compare_df['calc_base_total'] * 1.07
            # ----------------------------------------------------------------

            # เตรียมข้อมูลเปรียบเทียบยอดขาย/ต้นทุน
            sales_revenue_keys = ['sales_service_amount', 'cutting_drilling_fee', 'other_service_fee']
            db_compare_df['sales_for_comparison'] = 0
            for key in sales_revenue_keys:
                if key in db_compare_df.columns:
                    db_compare_df['sales_for_comparison'] += pd.to_numeric(db_compare_df[key], errors='coerce').fillna(0)

            cogs = pd.to_numeric(db_compare_df['cogs_db'], errors='coerce').fillna(0)
            db_compare_df['cost_db'] = cogs

            # เตรียมข้อมูลไฟล์อัปโหลด
            uploaded_compare_df['sales_uploaded'] = pd.to_numeric(uploaded_compare_df.get('sales_uploaded'), errors='coerce').fillna(0)
            uploaded_compare_df['cost_uploaded'] = pd.to_numeric(uploaded_compare_df.get('cost_uploaded'), errors='coerce').fillna(0)

            # Merge
            merged_df = pd.merge(db_compare_df, uploaded_compare_df, on='so_number', how='outer', suffixes=('_db', '_uploaded'), indicator=True)
            
            # Adjust Sales Uploaded (หักค่าขนส่งออกเพื่อให้เทียบกับ Base System ได้)
            merged_df['sales_uploaded'] = pd.to_numeric(merged_df['sales_uploaded'], errors='coerce').fillna(0) - pd.to_numeric(merged_df['shipping_cost'], errors='coerce').fillna(0)

            # [Restore Layout] คืนค่าคอลัมน์แหล่งที่มา
            merged_df['แหล่งยอดขาย'] = merged_df['hr_sale_source'].apply(
                lambda x: 'ระบบ' if x == 'system' else ('Express' if x == 'express' else 'ยังไม่เลือก')
            )
            merged_df['แหล่งต้นทุน'] = merged_df['hr_cost_source'].apply(
                lambda x: 'ระบบ' if x == 'system' else ('Express' if x == 'express' else 'ยังไม่เลือก')
            )

            # ==============================================================================
            # ฟังก์ชันตัดสินสถานะ (Logic ใหม่: ลบกันตรงๆ ตามที่คุณต้องการ)
            # ==============================================================================
            def determine_status_and_color(row):
                so_num = str(row.get('so_number', ''))
                
                # 1. หา "ยอดบิลสุทธิ" (Grand Total)
                grand_total = float(row.get('so_grand_total', 0) or 0)
                calc_grand = float(row.get('calc_grand_total', 0) or 0)
                
                # ถ้าใน DB เป็น 0 ให้ใช้ค่าที่คำนวณเองเมื่อกี้
                if grand_total <= 1.0: 
                    grand_total = calc_grand

                # 2. หา "ยอดที่จ่ายมาจริง" (Total Payment)
                pay_uploaded = float(row.get('total_payment_amount', 0) or 0)
                pay_db = float(row.get('total_payment_amount_db', 0) or 0)
                total_pay = pay_uploaded if pay_uploaded > 0 else pay_db
                
                # 3. หา "ภาษีหัก ณ ที่จ่าย" (WHT)
                wht = float(row.get('wht_3_percent', 0) or 0)

                # 4. คำนวณผลต่าง: (ยอดโอน + หักภาษี) - ยอดบิล
                real_difference = (total_pay + wht) - grand_total

                # --- DEBUG LOG ---
                if abs(real_difference) > 1.0:
                    print(f"DEBUG {so_num}: Pay({total_pay}) + WHT({wht}) - Bill({grand_total}) = Diff({real_difference})")

                # ================= ตัดสินสถานะ =================
                
                # A. ตรวจสอบการโอนเงิน (สำคัญสุด)
                if real_difference < -0.05: # ติดลบ = โอนขาด
                    return f"⚠️ ยอดโอนขาด ({abs(real_difference):,.2f})"
                
                # B. ตรวจสอบสถานะ Data (มีในระบบ/ไม่มี)
                if row['_merge'] == 'right_only': return 'มีใน Express, ไม่มีในระบบ'
                if row['_merge'] == 'left_only': return 'มีในระบบ, ไม่มีใน Express'

                # C. ตรวจสอบ Sales/Cost Mismatch
                sys_sale = float(row.get('sales_for_comparison', 0) or 0)
                exp_sale = float(row.get('sales_uploaded', 0) or 0)
                sys_cost = float(row.get('cost_db', 0) or 0)
                exp_cost = float(row.get('cost_uploaded', 0) or 0)

                if exp_sale > (sys_sale + 1.0): return "ยอดขายต่ำกว่า Express"
                if exp_cost > (sys_cost + 1.0): return "ต้นทุนต่ำกว่า Express"

                # D. ถ้าผ่านหมด
                if real_difference > 0.05: # บวก = โอนเกิน
                    return f"ผ่านเกณฑ์ (โอนเกิน {real_difference:,.2f})"
                else:
                    return "ผ่านเกณฑ์" # พอดีเป๊ะ (ผลต่างเป็น 0)

            # Apply Logic
            merged_df['สถานะ'] = merged_df.apply(determine_status_and_color, axis=1)
            
            merged_df['ผลต่างยอดขาย'] = merged_df['sales_for_comparison'].fillna(0) - merged_df['sales_uploaded'].fillna(0)
            merged_df['ผลต่างต้นทุน'] = merged_df['cost_db'].fillna(0) - merged_df['cost_uploaded'].fillna(0)
            
            # [Restore Layout] ใส่คอลัมน์กลับมาให้ครบ 100%
            display_order_map = {
                'so_number': 'เลขที่ SO',
                'sales_service_amount': 'ยอดขาย/บริการ (ระบบ)',
                'shipping_cost': 'ค่าขนส่ง (ระบบ)',
                'relocation_cost': 'ค่าย้าย (ระบบ)',        # <--- ใส่กลับมาแล้ว
                'sales_for_comparison': 'ยอดขายรวม (ระบบ)',
                'sales_uploaded': 'ยอดขาย (Express)',
                'cost_db': 'ต้นทุน (ระบบ)',
                'cost_uploaded': 'ต้นทุน (Express)',
                'ผลต่างยอดขาย': 'ผลต่างยอดขาย',
                'ผลต่างต้นทุน': 'ผลต่างต้นทุน',
                'แหล่งยอดขาย': 'แหล่งยอดขาย',              # <--- ใส่กลับมาแล้ว
                'แหล่งต้นทุน': 'แหล่งต้นทุน',              # <--- ใส่กลับมาแล้ว
                'สถานะ': 'สถานะ'
            }

            for key in display_order_map.keys():
                if key not in merged_df.columns:
                    merged_df[key] = np.nan
            
            self.comparison_df = merged_df[list(display_order_map.keys())].copy()
            self.comparison_df.rename(columns=display_order_map, inplace=True)

            # เพิ่มแถวสรุปยอดรวม (Total)
            so_count = len(self.comparison_df)
            numeric_cols = ['ยอดขายรวม (ระบบ)', 'ยอดขาย (Express)', 'ต้นทุน (ระบบ)', 'ต้นทุน (Express)', 'ผลต่างยอดขาย', 'ผลต่างต้นทุน']
            summary_data = self.comparison_df[numeric_cols].sum().to_dict()
            summary_row = pd.Series(summary_data)
            summary_row['เลขที่ SO'] = 'ยอดรวม (Total)'
            summary_row['สถานะ'] = f"รวม {so_count} รายการ"
            self.comparison_df = pd.concat([self.comparison_df, summary_row.to_frame().T], ignore_index=True)
            
            status_colors = {
                "ผ่านเกณฑ์": "#D1FAE5", 
                "ยอดขายต่ำกว่า Express": "#FEF2F2", 
                "ต้นทุนต่ำกว่า Express": "#FEFCE8", 
                "ข้อมูลไม่ตรงกัน": "#FFF7ED", 
                "ยอดโอนเกิน": "#D1FAE5", 
                "ยอดโอนขาด": "#FEF3C7", 
                "มีใน Express, ไม่มีในระบบ": "#FEF2F2", 
                "มีในระบบ, ไม่มีใน Express": "#FEFCE8",
                "กำไรดี": "#D1FAE5", "กำไรน้อย": "#FEFCE8", "ขาดทุน": "#FEF2F2", 
                "ยืนยันแล้ว (รอผล)": "#E5E7EB", 
                "‼️ ขายขาดทุน (ตรวจสอบด่วน)": "#F87171", 
                "‼️ ต้นทุน Express ผิดปกติ (<50%)": "#F97316",
            }
            
            self.results_frame_label.configure(text="ผลลัพธ์การเปรียบเทียบ (ดับเบิลคลิกเพื่อตรวจสอบ)")
            self._create_styled_dataframe_table(self.results_frame, self.comparison_df, "", status_column="สถานะ", status_colors=status_colors, on_row_click=self._on_tree_double_click)
            self.export_button.pack(side="right", padx=10, pady=10)
            self.verify_passed_button.pack(side="right", padx=10, pady=10)
            self._update_summary_pane()

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการเปรียบเทียบข้อมูล: {e}\n\n{traceback.format_exc()}", parent=self)
    
    def _save_comparison_to_log(self):
        if self.comparison_df is None or self.comparison_df.empty:
            print("Warning: _save_comparison_to_log called with no data. Skipping.")
            return

        try:
            summary_stats = {
                "total_records": len(self.comparison_df),
                "matched_records": len(self.comparison_df[self.comparison_df['สถานะ'] == 'ผ่านเกณฑ์']),
                "diff_records": len(self.comparison_df[self.comparison_df['สถานะ'] != 'ผ่านเกณฑ์']),
                "in_system_only": len(self.comparison_df[self.comparison_df['สถานะ'] == 'มีในระบบ, ไม่มีในไฟล์']),
                "in_file_only": len(self.comparison_df[self.comparison_df['สถานะ'] == 'มีในไฟล์, ไม่มีในระบบ'])
            }
            summary_json = json.dumps(summary_stats)
            detail_json = self.comparison_df.to_json(orient='records')
            source_info = os.path.basename(self.uploaded_file_path) if hasattr(self, 'uploaded_file_path') and self.uploaded_file_path else "Manual Entry"
            salesperson_filter = self.current_comparison_salesperson
            
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO comparison_logs (hr_user_key, salesperson_filter, source_info, summary_json, detail_json)
                    VALUES (%s, %s, %s, %s, %s)
                """, (self.user_key, salesperson_filter, source_info, summary_json, detail_json))
            conn.commit()
            print("Comparison log saved automatically during finalization.") # เปลี่ยนข้อความเป็น Log ภายใน
        except Exception as e:
            print(f"Error during automatic log save: {e}") # แสดง Error ใน Console แทน Popup
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _export_comparison(self):
        if self.comparison_df is None: messagebox.showwarning("ไม่มีข้อมูล", "กรุณาเปรียบเทียบข้อมูลก่อน Export", parent=self); return
        try:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="บันทึกผลการเปรียบเทียบ", initialfile=f"comparison_result_{datetime.now().strftime('%Y%m%d')}.xlsx")
            if save_path: self.comparison_df.to_excel(save_path, index=False); messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้วที่:\n{save_path}", parent=self)
        except Exception as e: messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=self)

    def _on_commission_row_double_click(self, event, tree, df):
        record_id_str = tree.focus()
        if not record_id_str: return
        
        record_id = int(record_id_str)
        row_data = df.loc[df['id'] == record_id].iloc[0].to_dict()
        self.app_container.show_edit_commission_window(data=row_data, refresh_callback=self._load_sale_history)

    def _refresh_comparison_view(self):
        if not hasattr(self, 'current_comparison_salesperson'):
            return

        if hasattr(self, 'active_sale_label'):
             self.active_sale_label.configure(
                 text=f"กำลังตรวจสอบข้อมูลของ: {self.current_comparison_salesperson}", 
                 text_color="#2563EB"
             )

        loading = self._show_loading(self.results_frame)
        self.results_frame_label.configure(text=f"กำลังรีเฟรชข้อมูลสำหรับ: {self.current_comparison_salesperson}...")

        try:
            # ดึงค่าเดือนและปีที่เคยเลือกไว้
            selected_month = getattr(self, 'current_comparison_month', None)
            selected_year = getattr(self, 'current_comparison_year', None)
            
            # --- แก้ไข Query หลัก (เหมือนข้างบน) ---
            base_query = """SELECT c.*, 
                       po.cogs_db, po.po_shipping_stock, po.po_shipping_site, po.po_relocation, 
                       u.sale_name,
                       ss.sale_name as support_user_name 
                FROM commissions c 
                JOIN sales_users u ON c.sale_key = u.sale_key
                LEFT JOIN sales_users ss ON c.support_user_key = ss.sale_key
                LEFT JOIN (
                        SELECT
                            p.so_number,
                            SUM(COALESCE(poi.total_price, 0)) as cogs_db,
                            SUM(p.shipping_to_stock_cost) as po_shipping_stock,
                            SUM(p.shipping_to_site_cost) as po_shipping_site,
                            SUM(p.relocation_cost) as po_relocation
                        FROM purchase_orders p
                        LEFT JOIN purchase_order_items poi ON p.id = poi.purchase_order_id
                        WHERE p.status = 'Approved'
                        GROUP BY p.so_number
                    ) po ON c.so_number = po.so_number
                WHERE c.is_active = 1
                  AND c.status NOT IN ('Paid', 'Cancelled', 'HR Verified')
            """
            params = []

            if self.current_comparison_salesperson != "ทั้งหมด":
                base_query += " AND c.sale_key = %s"
                params.append(self.current_comparison_salesperson)
            else:
                base_query += " AND c.sale_key IN (SELECT sale_key FROM sales_users WHERE status = 'Active' AND role = 'Sale')"

            # --- ✅ แก้ไขตรงนี้: เปลี่ยนเงื่อนไขเวลาเป็น 'ย้อนหลังทั้งหมด' ---
            if selected_month and selected_year:
                base_query += " AND ((c.commission_year < %s) OR (c.commission_year = %s AND c.commission_month <= %s))"
                params.extend([selected_year, selected_year, selected_month])

            data_query = base_query + " ORDER BY c.timestamp DESC"

            self.db_df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(params))

            if loading.winfo_exists(): loading.destroy()

            self._compare_data()

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            traceback.print_exc()
            messagebox.showerror("Database Error", f"ไม่สามารถรีเฟรชข้อมูลได้: {e}", parent=self)

    def _get_active_sales_list(self):
        try:
            # แก้ไข Query ให้ดึง sales_target มาด้วย
            df = pd.read_sql("SELECT sale_key, sale_name, commission_plan, sales_target FROM sales_users WHERE status = 'Active' and role = 'Sale' ORDER BY sale_key", self.pg_engine)
            # แก้ไขการสร้าง Dictionary ให้เก็บ target ด้วย
            self.sales_user_info = {
                row['sale_key']: {
                    'name': row['sale_name'], 
                    'plan': row['commission_plan'],
                    'target': row['sales_target'] 
                } for idx, row in df.iterrows()
            }
            return list(self.sales_user_info.keys())
        except Exception as e: 
            messagebox.showerror("DB Error", f"ไม่สามารถดึงรายชื่อพนักงานขายได้: {e}")
            return []

    def _confirm_and_save_commissions(self, selected_ids, df_to_process):
        if df_to_process.empty: messagebox.showwarning("No Data", "ไม่มีข้อมูลผลลัพธ์ที่จะบันทึก", parent=self); return
        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการยืนยันการจ่ายเงินสำหรับ {len(selected_ids)} รายการที่เลือกใช่หรือไม่?"): return
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                update_query = "UPDATE commissions SET final_commission = %s, status = 'Paid' WHERE id = %s"
                for _, row in df_to_process.iterrows():
                    final_comm = row.get('ค่าคอมที่คำนวณได้', 0.0)
                    final_comm = 0.0 if pd.isna(final_comm) else final_comm
                    record_id = row['record_id']
                    cursor.execute(update_query, (int(final_comm), int(record_id)))
                    cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, new_value, changes) VALUES (%s, %s, %s, %s, %s, %s)", ('Payout', 'commissions', record_id, self.user_key, json.dumps({'final_commission': final_comm, 'status': 'Paid'}), json.dumps({'status': 'Paid', 'final_commission': final_comm})))
            conn.commit(); messagebox.showinfo("สำเร็จ", "บันทึกและยืนยันการจ่ายเงินเรียบร้อยแล้ว", parent=self); self._on_sale_selected_for_process(); self._populate_audit_log_table()
        except Exception as e:
            if conn: conn.rollback(); messagebox.showerror("Database Error", f"ไม่สามารถบันทึกข้อมูลได้: {e}\n{traceback.format_exc()}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
        
    def _create_process_commission_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)
        
        control_frame = CTkFrame(parent_tab)
        control_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        CTkLabel(control_frame, text="เลือกพนักงานขาย:").pack(side="left", padx=(10,5))
        self.active_sales_keys = self._get_active_sales_list()
        self.selected_sale_for_process = tk.StringVar()
        self.sale_process_dropdown = CTkOptionMenu(control_frame, variable=self.selected_sale_for_process, values=self.active_sales_keys, command=self._on_sale_selected_for_process)
        self.sale_process_dropdown.pack(side="left", padx=5)
        
        CTkLabel(control_frame, text="เลือกงวดที่ต้องการคำนวณ:").pack(side="left", padx=(20, 5))
        self.process_period_var = tk.StringVar()
        self.process_period_menu = CTkOptionMenu(control_frame, variable=self.process_period_var, values=["-ยังไม่ได้เลือก-"], command=self._calculate_commission_for_period)
        self.process_period_menu.pack(side="left", padx=5)

        # multiplier ปรับ per-SO ผ่าน popup "รายละเอียดตาม SO" แล้ว ไม่ต้องมี global control ที่นี่

        # --- [แก้ไข] เปลี่ยนเป็น CTkScrollableFrame ---
        # เพื่อให้หน้าจอเลื่อนลงได้เมื่อเนื้อหายาวเกินหน้าจอโน้ตบุ๊ก
        self.process_result_frame = CTkScrollableFrame(parent_tab, label_text="ผลลัพธ์การคำนวณ")
        self.process_result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.process_result_frame.grid_columnconfigure(0, weight=1)

        self.after(100, self._on_sale_selected_for_process)
        

    def _apply_bulk_multiplier(self):
        """Director: ตั้ง cost_multiplier ให้ทุก SO ในงวดพร้อมกัน แล้ว recalculate"""
        if not hasattr(self, 'current_comm_df') or self.current_comm_df.empty:
            messagebox.showwarning("ยังไม่มีข้อมูล", "กรุณาเลือกพนักงานและงวดก่อน", parent=self)
            return
        mult = float(self.director_multiplier_var.get())
        so_ids = self.current_comm_df['id'].tolist()
        if not so_ids:
            return
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE commissions SET cost_multiplier = %s WHERE id = ANY(%s)",
                    (mult, so_ids)
                )
            conn.commit()
            self.app_container.release_connection(conn)
            messagebox.showinfo(
                "ตั้งค่าสำเร็จ",
                f"ตั้ง ×{mult:.2f} ให้ {len(so_ids)} SO เรียบร้อยแล้ว\nกด 'คำนวณขั้นสุดท้าย' เพื่อเห็นผลลัพธ์",
                parent=self
            )
            self._calculate_commission_for_period()
        except Exception as e:
            messagebox.showerror("Database Error", str(e), parent=self)

    def _on_sale_selected_for_process(self, sale_key=None):
        if sale_key is None:
            sale_key = self.selected_sale_for_process.get()
        if not sale_key: return

        for widget in self.process_result_frame.winfo_children(): widget.destroy()
        
        try:
            query = """
                SELECT DISTINCT commission_year, commission_month 
                FROM commissions 
                WHERE sale_key = %s AND status = 'HR Verified' AND is_active = 1
                AND payout_id IS NULL 
                ORDER BY CAST(commission_year AS INTEGER) DESC, CAST(commission_month AS INTEGER) DESC
            """
            df_periods = pd.read_sql_query(query, self.pg_engine, params=(sale_key,))

            if df_periods.empty:
                self.process_period_menu.configure(values=["-ไม่มีข้อมูล-"], state="disabled")
                self.process_period_var.set("-ไม่มีข้อมูล-")
                CTkLabel(self.process_result_frame, text=f"ไม่พบข้อมูลที่ 'Verified' และยังไม่ได้จ่ายเงินสำหรับ: {sale_key}").pack(pady=20)
                return

            period_options = [f"{self.thai_months[int(month)-1]} {int(year)+543}" for year, month in zip(df_periods['commission_year'], df_periods['commission_month'])]
            self.process_period_menu.configure(values=period_options, state="normal")
            
            target_period = period_options[0] 
            
            # 🔥 [จุดแก้ที่ 2] อ่านจาก self.last_verified_period ให้ตรงกันเป๊ะๆ
            last_verified = getattr(self, 'last_verified_period', None)
            if last_verified and last_verified in period_options:
                target_period = last_verified
                
            self.process_period_var.set(target_period)
            self._calculate_commission_for_period()

        except Exception as e:
            messagebox.showerror("DB Error", f"เกิดข้อผิดพลาดในการค้นหางวดข้อมูล: {e}", parent=self)
            traceback.print_exc()

    def _calculate_commission_for_period(self, selected_period=None):
        if selected_period is None:
            selected_period = self.process_period_var.get()
        
        sale_key = self.selected_sale_for_process.get()
        if not selected_period or not sale_key or "-" in selected_period:
            return

        # แยกเดือนและปี
        month_name, year_be_str = selected_period.split()
        month_num = self.thai_month_map[month_name]
        year_ad = int(year_be_str) - 543

        # บันทึกค่างวดไว้ใช้ตอน Save
        self.current_period_text = selected_period
        self.selected_month = month_num
        self.selected_year = year_ad

        # ดึงข้อมูล Plan และ Target
        plan_info = self.sales_user_info.get(sale_key, {})
        plan = plan_info.get('plan', 'Plan A')
        sales_target = float(plan_info.get('target', 0.0))

        # เคลียร์หน้าจอและแสดง Loading
        for widget in self.process_result_frame.winfo_children(): widget.destroy()
        loading = self._show_loading(self.process_result_frame)

        try:
            # ==============================================================================
            # 🔥 [จุดแก้ไขสำคัญ 1] ปรับ SQL Query ให้ดึงข้อมูล Cost แยกประเภทให้ครบถ้วน
            # ==============================================================================
            query_comm = """
                SELECT
                    c.*,
                    -- ดึงค่าแยกย่อยออกมา เพื่อให้ Business Logic นำไปจับคู่คำนวณ Excess ได้ถูกต้อง
                    COALESCE(po_costs.shipping_to_stock_cost, 0) as shipping_to_stock_cost,
                    COALESCE(po_costs.shipping_to_site_cost, 0) as shipping_to_site_cost,
                    COALESCE(po_costs.po_cutting_cost, 0) as po_cutting_cost,
                    COALESCE(po_costs.po_service_cost, 0) as po_service_cost,
                    -- ค่าตัด/เจาะ แยก VAT vs CASH (สำหรับ match กับฝั่ง SO)
                    COALESCE(po_costs.po_cutting_vat_cost,  0) as po_cutting_vat_cost,
                    COALESCE(po_costs.po_cutting_cash_cost, 0) as po_cutting_cash_cost,
                    -- ค่าตัด/เจาะ ที่เป็น EXP items ในรายการสินค้า (VAT case)
                    COALESCE(poi_cutting.po_cutting_item_cost, 0) as po_cutting_item_cost,
                    -- live_cogs: ยอดรวม PO items ปัจจุบัน (ครบทุก item รวม EXP-0006)
                    COALESCE(live_cogs.total_item_cost, c.final_cost_amount) as live_cogs
                FROM commissions c
                LEFT JOIN (
                    SELECT
                        so_number,
                        -- รวมยอดแยกตามประเภท (เฉพาะ PO ที่ไม่ถูก Cancel/Reject)
                        SUM(COALESCE(shipping_to_stock_cost, 0)) as shipping_to_stock_cost,
                        SUM(COALESCE(shipping_to_site_cost, 0)) as shipping_to_site_cost,
                        SUM(COALESCE(cutting_cost, 0)) as po_cutting_cost,
                        -- แยก cutting_cost ตาม VAT type
                        SUM(CASE WHEN cutting_vat_type = 'VAT'  THEN COALESCE(cutting_cost, 0) ELSE 0 END) as po_cutting_vat_cost,
                        SUM(CASE WHEN cutting_vat_type = 'CASH' THEN COALESCE(cutting_cost, 0) ELSE 0 END) as po_cutting_cash_cost,
                        0 as po_service_cost
                    FROM purchase_orders
                    WHERE status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
                    GROUP BY so_number
                ) po_costs ON c.so_number = po_costs.so_number
                LEFT JOIN (
                    -- EXP items ที่ชื่อขึ้นต้นด้วย ค่าตัด / ค่าเจาะ (VAT case: จาก product list)
                    SELECT po.so_number,
                           SUM(COALESCE(poi.total_price, 0)) as po_cutting_item_cost
                    FROM purchase_order_items poi
                    JOIN purchase_orders po ON po.id = poi.purchase_order_id
                    WHERE (poi.product_name ILIKE '%%ค่าตัด%%' OR poi.product_name ILIKE '%%ค่าเจาะ%%')
                      AND po.status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
                    GROUP BY po.so_number
                ) poi_cutting ON poi_cutting.so_number = c.so_number
                LEFT JOIN (
                    -- ใช้ total_cost จาก purchase_orders (หักส่วนลดท้ายบิลแล้ว) แทน SUM items
                    SELECT so_number, COALESCE(SUM(total_cost), 0) as total_item_cost
                    FROM purchase_orders
                    WHERE status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
                    GROUP BY so_number
                ) live_cogs ON c.so_number = live_cogs.so_number
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
            
            params = (sale_key, year_ad, year_ad, month_num)
            
            self.current_comm_df = pd.read_sql_query(query_comm, self.pg_engine, params=params)

            # ใช้ live_cogs แทน final_cost_amount (เพื่อรวม EXP-0006 ที่อาจ approve หลัง HR verify)
            if 'live_cogs' in self.current_comm_df.columns:
                self.current_comm_df['final_cost_amount'] = self.current_comm_df['live_cogs'].fillna(
                    self.current_comm_df['final_cost_amount'])
                self.current_comm_df.drop(columns=['live_cogs'], inplace=True)

            # cost_multiplier ของแต่ละ SO ถูกดึงมาจาก DB โดยตรง (ผ่าน SELECT c.*)
            # Director แก้ per-SO ผ่าน popup "รายละเอียดตาม SO" → บันทึกลง DB แล้ว
            # ไม่ override ที่นี่ เพื่อให้ค่า per-SO จาก DB ถูกนำไปใช้จริง

            # บันทึก ID ของ SO ที่จะถูกประมวลผล
            self.current_so_ids = self.current_comm_df['id'].tolist()

            # Reconciliation Logic
            if not self.current_comm_df.empty:
                so_ids_list = self.current_comm_df['id'].tolist()
                special_amounts_df = self._get_special_service_amounts(so_ids_list)

                if not special_amounts_df.empty:
                    # ลบ columns ที่ซ้ำออกก่อน merge เพื่อป้องกัน _x/_y suffix bug
                    cols_to_drop = [c for c in ['po_cutting_cost', 'po_service_cost']
                                    if c in self.current_comm_df.columns]
                    if cols_to_drop:
                        self.current_comm_df = self.current_comm_df.drop(columns=cols_to_drop)

                    self.current_comm_df = pd.merge(
                        self.current_comm_df,
                        special_amounts_df[['id', 'so_cutting_rev', 'so_service_rev', 'po_cutting_cost', 'po_service_cost']],
                        on='id',
                        how='left'
                    )
                    # เติม 0 ในช่องที่ว่าง
                    cols_to_fill = ['so_cutting_rev', 'so_service_rev', 'po_cutting_cost', 'po_service_cost']
                    for col in cols_to_fill:
                        if col in self.current_comm_df.columns:
                            self.current_comm_df[col] = self.current_comm_df[col].fillna(0)

            # ยอดขายรวม — ใช้ sales_service_amount เหมือน CalculationDetailViewer (col_mapping ลำดับแรก)
            # fallback → final_sales_amount ถ้าไม่มี
            _sales_col = ('sales_service_amount'
                          if 'sales_service_amount' in self.current_comm_df.columns
                          else 'final_sales_amount')
            self.current_total_sales = pd.to_numeric(
                self.current_comm_df[_sales_col], errors='coerce').fillna(0).sum()

            # ต้นทุนรวม = final_cost_amount × cost_multiplier (ต้องตรงกับตาราง CalculationDetailViewer)
            _cost_amt  = pd.to_numeric(self.current_comm_df['final_cost_amount'], errors='coerce').fillna(0)
            _cost_mult = pd.to_numeric(self.current_comm_df.get('cost_multiplier', 1.03), errors='coerce').fillna(1.03)
            self.current_total_cost = (_cost_amt * _cost_mult).sum()

            if self.current_comm_df.empty:
                loading.destroy()
                CTkLabel(self.process_result_frame, text="ไม่พบข้อมูลในงวดที่เลือก").pack(pady=20)
                return
            
            # --- Auto Deduction Calculation (เพื่อโชว์บนหน้าจอ HR) ---
            # คิดแยกราย SO — เฉพาะ SO ที่ PO ค่าขนส่ง > SO ค่าขนส่ง เท่านั้น ไม่ให้ SO ดีช่วย offset SO แย่
            _so_ship = (pd.to_numeric(self.current_comm_df['shipping_cost'], errors='coerce').fillna(0)
                        + pd.to_numeric(self.current_comm_df['relocation_cost'], errors='coerce').fillna(0))
            _po_ship = (pd.to_numeric(self.current_comm_df['shipping_to_stock_cost'], errors='coerce').fillna(0)
                        + pd.to_numeric(self.current_comm_df['shipping_to_site_cost'], errors='coerce').fillna(0))
            shipping_diff = (_po_ship - _so_ship).clip(lower=0).sum()  # เอาเฉพาะ row ที่ PO > SO
            shipping_deduction = (shipping_diff / 0.2) * 0.0175 if shipping_diff > 0 else 0.0
            total_so_shipping = _so_ship.sum()
            total_po_shipping = _po_ship.sum()
            
            # Brokerage / Marketing Deduction (คงเดิม)
            total_brokerage = self.current_comm_df['brokerage_fee'].sum()
            total_difference = self.current_comm_df['difference_amount'].sum()
            difference_deduction = 0.0
            diff_base = total_brokerage - total_difference
            if diff_base < 0:
                difference_deduction = (abs(diff_base) / 0.2) * 0.0175

            total_marketing = self.current_comm_df['coupons'].sum() + self.current_comm_df['giveaways'].sum()
            marketing_deduction = (total_marketing / 0.2) * 0.0175 if total_marketing > 0 else 0.0

            final_auto_deduction = shipping_deduction + difference_deduction + marketing_deduction

            # --- DEBUG: แสดงรายละเอียด Auto Deduction ราย SO ---
            print("\n" + "="*70)
            print(f"[DEBUG] Auto Deduction Breakdown  (รวม = {final_auto_deduction:,.2f} บาท)")
            print("="*70)

            # 1. Shipping deduction
            print(f"\n[1] Shipping Deduction = {shipping_deduction:,.2f} บาท")
            print(f"    PO shipping รวม : {total_po_shipping:,.2f}  |  SO shipping รวม : {total_so_shipping:,.2f}")
            print(f"    ส่วนต่าง (PO-SO) : {total_po_shipping - total_so_shipping:,.2f}  {'(ไม่หัก)' if total_po_shipping <= total_so_shipping else ''}")
            _debug_ship_cols = ['so_number', 'shipping_cost', 'relocation_cost',
                                'shipping_to_stock_cost', 'shipping_to_site_cost']
            _avail = [c for c in _debug_ship_cols if c in self.current_comm_df.columns]
            _ship_df = self.current_comm_df[_avail].copy()
            for col in _avail[1:]:
                _ship_df[col] = pd.to_numeric(_ship_df[col], errors='coerce').fillna(0)
            _ship_df['so_ship'] = _ship_df.get('shipping_cost', 0) + _ship_df.get('relocation_cost', 0)
            _ship_df['po_ship'] = _ship_df.get('shipping_to_stock_cost', 0) + _ship_df.get('shipping_to_site_cost', 0)
            _ship_nonzero = _ship_df[(_ship_df['so_ship'] != 0) | (_ship_df['po_ship'] != 0)]
            if not _ship_nonzero.empty:
                print(f"    {'SO':<20} {'SO ship':>12} {'PO ship':>12}")
                print(f"    {'-'*44}")
                for _, r in _ship_nonzero.iterrows():
                    print(f"    {str(r.get('so_number','')):<20} {r['so_ship']:>12,.2f} {r['po_ship']:>12,.2f}")

            # 2. Brokerage/Difference deduction
            print(f"\n[2] Brokerage/Difference Deduction = {difference_deduction:,.2f} บาท")
            print(f"    brokerage รวม : {total_brokerage:,.2f}  |  difference รวม : {total_difference:,.2f}")
            print(f"    diff_base (broker-diff) : {diff_base:,.2f}  {'(ไม่หัก)' if diff_base >= 0 else ''}")
            _brok_cols = ['so_number', 'brokerage_fee', 'difference_amount']
            _avail2 = [c for c in _brok_cols if c in self.current_comm_df.columns]
            _brok_df = self.current_comm_df[_avail2].copy()
            for col in _avail2[1:]:
                _brok_df[col] = pd.to_numeric(_brok_df[col], errors='coerce').fillna(0)
            _brok_nonzero = _brok_df[(_brok_df.get('brokerage_fee', 0) != 0) | (_brok_df.get('difference_amount', 0) != 0)]
            if not _brok_nonzero.empty:
                print(f"    {'SO':<20} {'brokerage':>12} {'difference':>12}")
                print(f"    {'-'*44}")
                for _, r in _brok_nonzero.iterrows():
                    print(f"    {str(r.get('so_number','')):<20} {r.get('brokerage_fee',0):>12,.2f} {r.get('difference_amount',0):>12,.2f}")

            # 2.5 Cutting/Drilling SO vs PO
            _cut_so = pd.to_numeric(self.current_comm_df.get('cutting_drilling_fee', 0), errors='coerce').fillna(0)
            _cut_po = pd.to_numeric(self.current_comm_df.get('po_cutting_cost', 0), errors='coerce').fillna(0)
            _svc_po = pd.to_numeric(self.current_comm_df.get('po_service_cost', 0), errors='coerce').fillna(0)
            _so_cut_total = _cut_so.sum()
            _po_cut_total = (_cut_po + _svc_po).sum()
            print(f"\n[2.5] ค่าตัด/เจาะ SO vs PO (ข้อมูลเปรียบเทียบ)")
            print(f"    SO ค่าตัด+เจาะ รวม : {_so_cut_total:,.2f}  |  PO ค่าตัด+บริการ รวม : {_po_cut_total:,.2f}")
            print(f"    ส่วนต่าง (PO-SO) : {_po_cut_total - _so_cut_total:,.2f}  {'(PO แพงกว่า)' if _po_cut_total > _so_cut_total else '(SO สูงกว่า)'}")
            _cut_df = self.current_comm_df[['so_number']].copy()
            _cut_df['so_cut']  = _cut_so.values
            _cut_df['po_cut']  = (_cut_po + _svc_po).values
            _cut_df['diff']    = _cut_df['po_cut'] - _cut_df['so_cut']
            _cut_nonzero = _cut_df[(_cut_df['so_cut'] != 0) | (_cut_df['po_cut'] != 0)]
            if not _cut_nonzero.empty:
                print(f"    {'SO':<20} {'SO ตัด/เจาะ':>14} {'PO ตัด/บริการ':>14} {'ส่วนต่าง':>12}")
                print(f"    {'-'*62}")
                for _, r in _cut_nonzero.iterrows():
                    flag = " ⚠" if r['diff'] > 0 else ""
                    print(f"    {str(r.get('so_number','')):<20} {r['so_cut']:>14,.2f} {r['po_cut']:>14,.2f} {r['diff']:>12,.2f}{flag}")
            else:
                print(f"    (ไม่มี SO ที่มีค่าตัด/เจาะ)")

            # 3. Marketing deduction
            print(f"\n[3] Marketing Deduction (coupons+giveaways) = {marketing_deduction:,.2f} บาท")
            print(f"    total_marketing = {total_marketing:,.2f}")
            _mkt_cols = ['so_number', 'coupons', 'giveaways']
            _avail3 = [c for c in _mkt_cols if c in self.current_comm_df.columns]
            _mkt_df = self.current_comm_df[_avail3].copy()
            for col in _avail3[1:]:
                _mkt_df[col] = pd.to_numeric(_mkt_df[col], errors='coerce').fillna(0)
            _mkt_nonzero = _mkt_df[(_mkt_df.get('coupons', 0) != 0) | (_mkt_df.get('giveaways', 0) != 0)]
            if not _mkt_nonzero.empty:
                print(f"    {'SO':<20} {'coupons':>12} {'giveaways':>12}")
                print(f"    {'-'*44}")
                for _, r in _mkt_nonzero.iterrows():
                    print(f"    {str(r.get('so_number','')):<20} {r.get('coupons',0):>12,.2f} {r.get('giveaways',0):>12,.2f}")

            print("\n" + "="*70 + "\n")
            # --- END DEBUG ---

            # เตรียม Dataframe ส่งไปคำนวณจริง
            df_for_calc = self.current_comm_df.copy()
            df_for_calc['final_sales_amount'] = pd.to_numeric(df_for_calc['final_sales_amount'], errors='coerce').fillna(0.0)
            df_for_calc['total_revenue'] = df_for_calc['final_sales_amount']
            
            # ค่าดำเนินการ
            default_operating_fee = 0.0
            default_fees = {'Plan A': 25000.00, 'Plan B': 100000.00, 'Plan C': 100000.00, 'Plan D': 750000.00}
            standard_plan_fee = default_fees.get(plan, 0.0)

            try:
                if hasattr(self, 'operating_fee_entry') and self.operating_fee_entry and self.operating_fee_entry.winfo_exists():
                    fee_str = self.operating_fee_entry.get()
                    default_operating_fee = utils.convert_to_float(fee_str) if fee_str.strip() else 0.0
                else:
                    default_operating_fee = standard_plan_fee
            except:
                default_operating_fee = standard_plan_fee
            
            # --- ส่งคำนวณค่าคอมมิชชั่น ---
            if plan in ['Plan A', 'Plan B', 'Plan C', 'Plan D']:
                # 🔥 ส่งข้อมูลที่ครบถ้วน (มี column แยกย่อยครบ) ไปให้ Business Logic
                self.initial_commission_result = business_logic.calculate_monthly_commission(
                    plan_name=plan,
                    comm_df=df_for_calc,
                    sales_target=sales_target,
                    operating_fee=default_operating_fee,
                    incentives=None,
                    additional_deductions=None
                )
            else:
                self.initial_commission_result = {'type': 'error', 'message': f'ไม่รู้จัก Plan: {plan}'}

            self.latest_commission_result = self.initial_commission_result

            # --- แสดงผลลัพธ์ ---
            loading.destroy()
            result_type = self.initial_commission_result.get('type')

            if result_type in ['no_commission', 'error']:
                message = self.initial_commission_result.get('message', 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ')
                CTkLabel(self.process_result_frame, text=message, font=self.label_font, text_color="orange", wraplength=600).pack(pady=30, padx=20)
            else:
                if result_type == 'summary_plan_a':
                    self.commission_details_df = self.initial_commission_result.get('details')
                else:
                    self.commission_details_df = None
                
                # สร้างหน้าจอ Input พร้อมค่าที่คำนวณได้
                self._create_hr_input_interface(
                    auto_deduction_value=final_auto_deduction, # ส่งค่ายอดรวมที่คิดได้ไปแสดง
                    default_operating_fee_to_display=default_operating_fee 
                )

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            traceback.print_exc()
            messagebox.showerror("Calculation Error", f"เกิดข้อผิดพลาดในการคำนวณ: {e}", parent=self)

    def _create_hr_input_interface(self, auto_deduction_value=0.0, default_operating_fee_to_display=None):
        """
        สร้างหน้าจอสำหรับกรอกข้อมูลค่าคอมมิชชั่น (เพิ่มช่องยอดขายขั้นต่ำ และจัด Row ใหม่)
        """
        for widget in self.process_result_frame.winfo_children():
            widget.destroy()

        self.process_result_frame.grid_columnconfigure(0, weight=1)

        if not hasattr(self, 'initial_commission_result'):
             self.initial_commission_result = {}
             
        calculated_commission = (
            self.initial_commission_result.get('final_commission_pre_deductions') or 
            self.initial_commission_result.get('final_commission', 0.0)
        )   

        input_frame = CTkFrame(self.process_result_frame)
        input_frame.grid(row=0, column=0, pady=(10, 0), padx=10, sticky="ew")

        # Row 0: แผน
        plan_name = self.sales_user_info.get(self.selected_sale_for_process.get(), {}).get('plan', 'N/A')
        self.plan_display_label = CTkLabel(input_frame, text=f"แผนค่าคอมมิชชั่น: {plan_name}", font=self.header_font_table, text_color=self.theme["primary"])
        self.plan_display_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

        # Row 1: ยอดคอมมิชชั่นที่คำนวณได้
        CTkLabel(input_frame, text="ยอดคอมมิชชั่นที่คำนวณได้:", font=self.label_font).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.calculated_commission_label = CTkLabel(input_frame, text=f"{calculated_commission:,.2f} บาท", font=self.header_font_table)
        self.calculated_commission_label.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # Row 2: สถิติยอดขาย/ต้นทุน
        stats_frame = CTkFrame(input_frame, fg_color="transparent")
        stats_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(5,0))
        stats_frame.grid_columnconfigure((1, 3), weight=1)
        
        total_sales_display = getattr(self, 'current_total_sales', 0.0)
        total_cost_display = getattr(self, 'current_total_cost', 0.0)
        
        CTkLabel(stats_frame, text="ยอดขายรวม:", font=self.label_font, text_color="#2563EB").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        CTkLabel(stats_frame, text=f"{total_sales_display:,.2f}", font=self.entry_font).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        CTkLabel(stats_frame, text="ต้นทุนรวม:", font=self.label_font, text_color="#D97706").grid(row=0, column=2, padx=(20, 10), pady=5, sticky="w")
        CTkLabel(stats_frame, text=f"{total_cost_display:,.2f}", font=self.entry_font).grid(row=0, column=3, padx=10, pady=5, sticky="w")

        # --- Row 3: (-) ค่าดำเนินการ ---
        CTkLabel(input_frame, text="(-) ค่าดำเนินการ:", font=self.label_font).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.operating_fee_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.operating_fee_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        
        if default_operating_fee_to_display is not None:
            self.operating_fee_entry.insert(0, f"{default_operating_fee_to_display:,.2f}")
        else:
            default_fees = {'Plan A': 25000, 'Plan B': 100000, 'Plan C': 100000, 'Plan D': 750000}
            val = default_fees.get(plan_name, 0.0)
            self.operating_fee_entry.insert(0, f"{val:,.2f}")

        # --- Row 4: ยอดขายขั้นต่ำ ---
        CTkLabel(input_frame, text="ยอดขายขั้นต่ำ:", font=self.label_font).grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.min_sales_entry = NumericEntry(input_frame, placeholder_text="500,000")
        self.min_sales_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")
        
        default_min_sales = 750000 if plan_name == 'Plan D' else 500000
        self.min_sales_entry.insert(0, f"{default_min_sales:,.0f}")

        # --- Row 5: (+) Incentive ---
        CTkLabel(input_frame, text="(+) Incentive:", font=self.label_font).grid(row=5, column=0, padx=10, pady=10, sticky="w")
        self.incentive_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.incentive_entry.grid(row=5, column=1, padx=10, pady=10, sticky="ew")

        # --- Row 6: (-) หัก ค่าใช้จ่ายอื่นๆ ---
        CTkLabel(input_frame, text="(-) หัก ค่าใช้จ่ายอื่นๆ:", font=self.label_font).grid(row=6, column=0, padx=10, pady=10, sticky="w")
        self.deduction_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.deduction_entry.grid(row=6, column=1, padx=10, pady=10, sticky="ew")
        if auto_deduction_value > 0:
            self.deduction_entry.insert(0, f"{auto_deduction_value:,.2f}")

        # --- Row 7: หมายเหตุ ---
        CTkLabel(input_frame, text="หมายเหตุ/Incentive อื่นๆ:", font=self.label_font).grid(row=7, column=0, padx=10, pady=10, sticky="w")
        self.payout_notes_entry = CTkTextbox(input_frame, height=80)
        self.payout_notes_entry.grid(row=7, column=1, padx=10, pady=10, sticky="ew")
        
        # --- Row 8: ปุ่มกด ---
        calc_button_frame = CTkFrame(input_frame, fg_color="transparent")
        calc_button_frame.grid(row=8, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
        calc_button_frame.grid_columnconfigure((0, 1), weight=1)
        
        CTkButton(calc_button_frame, text="คำนวณขั้นสุดท้ายและแสดงสรุป", command=self._perform_final_calculation, fg_color=self.theme["primary"]).grid(row=0, column=0, padx=(0, 5), pady=10, sticky="ew")
        
        self.detail_button = CTkButton(calc_button_frame, text="แสดงการคิดแบบละเอียด", command=self._show_calculation_details)
        self.detail_button.grid(row=0, column=1, padx=(5, 0), pady=10, sticky="ew")

        # =================================================================
        # 🔥 แก้ไข Error: ตรวจสอบความมีอยู่ของข้อมูลอย่างปลอดภัย
        # =================================================================
        debug_data = self.initial_commission_result.get('debug_df')
        has_data = False
        
        if isinstance(debug_data, pd.DataFrame) and not debug_data.empty:
            has_data = True
        elif isinstance(debug_data, (list, dict)) and len(debug_data) > 0:
            has_data = True

        if has_data:
            self.detail_button.configure(state="normal")
        else:
            self.detail_button.configure(state="disabled")
        # =================================================================

        # พื้นที่แสดงตารางสรุป
        self.final_summary_frame = CTkFrame(self.process_result_frame, fg_color="transparent")
        self.final_summary_frame.grid(row=1, column=0, pady=10, padx=10, sticky="ew")
        
        # พื้นที่ปุ่มยืนยัน
        bottom_action_frame = CTkFrame(self.process_result_frame, fg_color="transparent")
        bottom_action_frame.grid(row=2, column=0, pady=(0, 10), padx=10, sticky="ew")
        
        self.confirm_payout_button = CTkButton(bottom_action_frame, text="✅ ยืนยันการจ่ายเงินและบันทึก",
                                command=self._confirm_payout_and_save,
                                fg_color="#16A34A", hover_color="#15803D",
                                font=CTkFont(size=16, weight="bold"))
        self.confirm_payout_button.pack(fill="x", padx=10, pady=5)

    def _get_incentives_data(self):
        try:
            val = float(self.incentive_entry.get().replace(",", "") or 0.0)
            return {"Incentive พิเศษ": val} if val > 0 else {}
        except: return {}

    def _get_deductions_data(self):
        try:
            val = float(self.deduction_entry.get().replace(",", "") or 0.0)
            return {"ค่าใช้จ่าย/ดำเนินการ": val} if val > 0 else {}
        except: return {}

    def _perform_final_calculation(self):
        """
        อ่านค่าจากหน้าจอและเรียก business_logic เพื่อคำนวณยอดสุทธิ
        """
        try:
            # --- 1. ระบุตัวตนพนักงานและแผน ---
            sale_key = self.selected_sale_for_process.get()
            plan = self.sales_user_info.get(sale_key, {}).get('plan', 'Plan A') 

            # --- 2. ดึงข้อมูลตัวเลขจากหน้าจอ ---
            sales_target_val = float(self.sales_target_entry.get().replace(",", "") or 0.0)
            operating_fee_val = float(self.operating_fee_entry.get().replace(",", "") or 0.0)
            min_sales_val = float(self.min_sales_entry.get().replace(",", "") or 500000.0)

            # --- 3. รวบรวมข้อมูล Incentive และ Deduction ---
            incentives_dict = self._get_incentives_data()
            deductions_dict = self._get_deductions_data()

            # ตรวจสอบข้อมูลพนักงาน
            if self.current_comm_df is None or self.current_comm_df.empty:
                messagebox.showwarning("คำเตือน", "ไม่พบข้อมูลพนักงานสำหรับคำนวณขั้นสุดท้าย")
                return

            # --- 4. เรียกใช้ Logic การคำนวณ ---
            final_result = business_logic.calculate_monthly_commission(
                plan_name=plan,
                comm_df=self.current_comm_df,
                sales_target=sales_target_val,
                operating_fee=operating_fee_val,
                incentives=incentives_dict,
                additional_deductions=deductions_dict,
                min_sales_target=min_sales_val
            )

            # [สำคัญ] บันทึกผลลัพธ์เก็บไว้ให้ปุ่ม Detail เรียกใช้
            self.latest_commission_result = final_result

            # --- 5. แสดงผลลัพธ์บนหน้าจอ ---
            self.final_summary_data = None 
            self.confirm_payout_button.pack_forget()

            for widget in self.final_summary_frame.winfo_children():
                widget.destroy()

            # ตรวจสอบว่าคำนวณสำเร็จหรือไม่
            if final_result.get('type') == 'error':
                 messagebox.showerror("คำนวณล้มเหลว", final_result.get('message', 'เกิดข้อผิดพลาดไม่ทราบสาเหตุ'))
                 # ปิดปุ่ม Detail ถ้าคำนวณพัง
                 self.detail_button.configure(state="disabled")
                 return

            # แยกประเภทผลลัพธ์เพื่อเลือก key ให้ถูก
            result_type = final_result.get('type')
            summary_df = None
            
            if result_type == 'summary_plan_a':
                summary_df = final_result.get('summary')
            else:
                # Plan B, C, D จะส่งมาใน key 'data'
                summary_df = final_result.get('data')

            if summary_df is not None:
                self._create_commission_summary_table(summary_df, container=self.final_summary_frame)
                
                # เปิดปุ่ม Detail ให้กดได้ เพราะมีข้อมูลแล้ว
                self.detail_button.configure(state="normal")

                # แสดงปุ่มยืนยัน
                self.final_summary_data = summary_df 
                self.confirm_payout_button.pack(pady=(10, 20), padx=20, ipady=10, side="bottom", anchor="se")
            else:
                messagebox.showerror("ผิดพลาด", "ไม่สามารถสร้างตารางสรุปได้ (ข้อมูลว่างเปล่า)")

        except ValueError:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกตัวเลขให้ถูกต้อง")
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาด: {str(e)}")
            traceback.print_exc()
    
    def _confirm_payout_and_save(self):
        """
        (ฉบับแก้ไข: คำนวณ Normal/Below Sales จาก DataFrame โดยตรง แก้ปัญหา Below T เป็น 0)
        """
        try:
            # --- 1. ตรวจสอบความพร้อม ---
            if not hasattr(self, 'latest_commission_result') or not self.latest_commission_result:
                messagebox.showwarning("ยังไม่พร้อม", "กรุณากด 'คำนวณขั้นสุดท้าย' ก่อนยืนยันการจ่ายเงิน", parent=self)
                return

            # ดึงข้อมูลสรุปผลการคำนวณ
            result_type = self.latest_commission_result.get('type')
            final_summary_df = None
            
            if result_type == 'summary_plan_a':
                final_summary_df = self.latest_commission_result.get('summary')
            elif result_type == 'summary_other':
                final_summary_df = self.latest_commission_result.get('data')

            if final_summary_df is None:
                messagebox.showerror("ผิดพลาด", "ไม่พบข้อมูลสรุปผลการคำนวณ", parent=self)
                return

            # --- 2. เตรียมตัวเลขเพื่อแสดงใน Popup ---
            def get_val(desc_keyword):
                try:
                    row = final_summary_df[final_summary_df['description'].str.contains(desc_keyword, case=False, na=False)]
                    if not row.empty:
                        return float(row['value'].iloc[0])
                    return 0.0
                except:
                    return 0.0

            val_gross = get_val("Gross|ขั้นต้น")
            val_wht = get_val("หัก ณ ที่จ่าย|3%")
            val_net = get_val("ยอดสรุปคอมหลังหัก|สุทธิ|Net")
            
            # นับจำนวน SO
            count_current = 0
            count_old = 0
            
            if hasattr(self, 'current_comm_df') and not self.current_comm_df.empty:
                df_temp = self.current_comm_df.copy()
                df_temp['commission_year'] = pd.to_numeric(df_temp['commission_year'], errors='coerce').fillna(0)
                df_temp['commission_month'] = pd.to_numeric(df_temp['commission_month'], errors='coerce').fillna(0)
                
                current_mask = (df_temp['commission_year'] == self.selected_year) & (df_temp['commission_month'] == self.selected_month)
                count_current = len(df_temp[current_mask])
                count_old = len(df_temp[~current_mask])
            
            total_items = count_current + count_old

            # --- 3. แสดง Popup ยืนยัน (Custom Dialog แบบละเอียด) ---
            comm_df_for_dialog = self.current_comm_df if hasattr(self, 'current_comm_df') else None
            dlg = PayoutConfirmDialog(
                self.winfo_toplevel(),
                period_text=self.current_period_text,
                val_gross=val_gross,
                val_wht=val_wht,
                val_net=val_net,
                comm_df=comm_df_for_dialog,
                selected_year=self.selected_year,
                selected_month=self.selected_month,
            )
            self.wait_window(dlg)
            if not dlg.result:
                return

            # =========================================================
            # เตรียมข้อมูลสำหรับบันทึก Log
            # =========================================================
            
            payout_notes = self.payout_notes_entry.get("1.0", "end-1c").strip()
            plan_name = self.sales_user_info.get(self.selected_sale_for_process.get(), {}).get('plan', 'N/A')
            
            # ดึง Incentive/Deduction
            incentives_df = final_summary_df[final_summary_df['description'].str.startswith('(+) ')]
            incentives_total = incentives_df['value'].sum()

            deductions_df = final_summary_df[
                final_summary_df['description'].str.startswith('(-) ')
                & (~final_summary_df['description'].str.contains('ดำเนินการ', case=False, na=False))
                & (~final_summary_df['description'].str.contains('หัก ณ ที่จ่าย', case=False, na=False))
            ]
            deductions_total = deductions_df['value'].sum()
            
            # คำนวณยอดขายรวม
            real_total_sales = 0.0
            if hasattr(self, 'current_comm_df') and not self.current_comm_df.empty:
                s_amount = pd.to_numeric(self.current_comm_df['final_sales_amount'], errors='coerce').fillna(0.0)
                real_total_sales = s_amount.sum()
            
            # [🔥 แก้ไข] คำนวณยอด Normal/Below จาก DataFrame โดยตรง
            val_normal_sales = 0.0
            val_below_sales = 0.0
            if hasattr(self, 'current_comm_df') and not self.current_comm_df.empty:
                margin_col = pd.to_numeric(self.current_comm_df['margin'], errors='coerce').fillna(0)
                sales_col = pd.to_numeric(self.current_comm_df['final_sales_amount'], errors='coerce').fillna(0)
                val_normal_sales = sales_col[margin_col >= 10].sum()
                val_below_sales = sales_col[margin_col < 10].sum()

            # --- [สำคัญ] เตรียมข้อมูล detail_json เพื่อบันทึก ---
            debug_df = self.latest_commission_result.get('debug_df')     
            breakdown_df = self.latest_commission_result.get('so_breakdown_df') 
            
            details_pack = {
                "debug": debug_df.to_dict(orient='records') if debug_df is not None else [],
                "breakdown": breakdown_df.to_dict(orient='records') if breakdown_df is not None else []
            }

            # Dictionary ข้อมูลที่จะบันทึก
            log_data = {
                "sale_key": self.selected_sale_for_process.get(),
                "plan_name": plan_name,
                "payout_period_text": self.current_period_text,
                "commission_month": self.selected_month,
                "commission_year": self.selected_year,
                "calculated_commission": float(self.latest_commission_result.get('final_commission_pre_deductions', 0.0)),
                "incentives_total": float(incentives_total),
                "deductions_total": float(deductions_total),
                "final_commission": float(val_gross),         
                "withholding_tax": float(val_wht),           
                "net_commission": float(val_net),            
                "notes": payout_notes,
                "summary_data_json": final_summary_df.to_json(orient='records'),
                "so_ids_json": json.dumps(self.current_so_ids),
                "total_sales": float(real_total_sales),
                "total_normal_sales": float(val_normal_sales),
                "total_below_sales": float(val_below_sales),
                "detail_json": json.dumps(details_pack, default=str)
            }
            
            log_data = {k: v for k, v in log_data.items() if v is not None}

            # บันทึกลง DB
            conn = self.app_container.get_connection()
            try:
                with conn.cursor() as cursor:
                    columns = ", ".join(log_data.keys())
                    placeholders = ", ".join(["%s"] * len(log_data))
                    sql_insert_log = f"""
                        INSERT INTO commission_payout_logs ({columns}, timestamp)
                        VALUES ({placeholders}, %s)
                        RETURNING id;
                    """
                    params = list(log_data.values()) + [datetime.now()]
                    cursor.execute(sql_insert_log, tuple(params))
                    payout_id = cursor.fetchone()[0]

                    so_ids_tuple = tuple(self.current_so_ids)
                    if so_ids_tuple:
                        # ถ้า Director เลือก multiplier ไว้ → save ลง commissions ด้วย
                        if (getattr(self, 'user_role', None) == 'Director'
                                and hasattr(self, 'director_multiplier_var')):
                            _mult = float(self.director_multiplier_var.get())
                            cursor.execute("""
                                UPDATE commissions
                                SET status = 'Paid', payout_id = %s, cost_multiplier = %s
                                WHERE id IN %s
                            """, (payout_id, _mult, so_ids_tuple))
                        else:
                            cursor.execute("""
                                UPDATE commissions
                                SET status = 'Paid', payout_id = %s
                                WHERE id IN %s
                            """, (payout_id, so_ids_tuple))
                
                    audit_msg = json.dumps({'payout_id': payout_id, 'net_amount': val_net})
                    cursor.execute("""
                        INSERT INTO audit_log (action, table_name, record_id, user_info, new_value, timestamp) 
                        VALUES (%s, %s, %s, %s, %s, NOW())
                    """, ('Confirm Payout', 'commission_payout_logs', payout_id, self.user_key, audit_msg))

                conn.commit()
                messagebox.showinfo("สำเร็จ", f"บันทึกการจ่ายเงิน (Payout ID: {payout_id}) เรียบร้อยแล้ว!", parent=self)
                
                self._on_sale_selected_for_process()
                self._load_payout_history()
            
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
                traceback.print_exc()
            finally:
                if conn: self.app_container.release_connection(conn)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการเตรียมข้อมูล: {e}", parent=self)
            traceback.print_exc()
    
    def _toggle_select_all_payouts(self):
        if not hasattr(self, 'payout_tree') or not self.payout_tree.winfo_exists():
            messagebox.showwarning("เกิดข้อผิดพลาด", "ไม่พบตารางผลลัพธ์ที่จะดำเนินการ\nกรุณาลองเลือกพนักงานขายอีกครั้ง", parent=self)
            self.select_all_var.set(0)
            return

        if self.select_all_var.get() == 1:
            self.selected_payout_ids.clear()
            for child_id in self.payout_tree.get_children():
                record_id_int = int(child_id); self.selected_payout_ids.add(record_id_int); current_values = list(self.payout_tree.item(child_id, "values")); current_values[0] = "☑"; self.payout_tree.item(child_id, values=current_values)
        else:
            for record_id_int in list(self.selected_payout_ids):
                if self.payout_tree.exists(str(record_id_int)):
                    current_values = list(self.payout_tree.item(str(record_id_int), "values")); current_values[0] = "☐"; self.payout_tree.item(str(record_id_int), values=current_values)
            self.selected_payout_ids.clear()
        self._update_payout_summary()

    def _create_payout_table(self, df, is_paid_view):
        for widget in self.process_result_frame.winfo_children(): widget.destroy()
        if df.empty: CTkLabel(self.process_result_frame, text="ไม่พบข้อมูล").pack(pady=20); self.select_all_checkbox.configure(state="disabled"); return
        cols = df.columns.tolist()
        if not is_paid_view:
            cols_to_display = ["เลือก"] + [c for c in cols if c != "record_id"]; display_df = df.copy(); display_df["เลือก"] = "☐"; display_df = display_df[["เลือก"] + [c for c in display_df.columns if c != "เลือก" and c != "record_id"]]
        else: cols_to_display = [c for c in cols if c != "record_id"]; display_df = df.copy()
        tree_frame = CTkFrame(self.process_result_frame, fg_color="transparent"); tree_frame.pack(fill="both", expand=True); tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)
        style = ttk.Style(); style.theme_use("default"); style.configure("Treeview.Heading", font=self.header_font_table); style.configure("Treeview", rowheight=28, font=self.entry_font); style.map("Treeview", background=[('selected', self.app_container.THEME["hr"].get("primary", "#16A34A"))])
        tree = ttk.Treeview(tree_frame, columns=cols_to_display, show='headings', style="Treeview"); tree.grid(row=0, column=0, sticky="nsew"); self.payout_tree = tree
        status_colors = {"Normal": "#D1FAE5", "Eligible": "#D1FAE5", "Below T": "#FEF9C3", "Under Tier": "#FEF9C3", "Eligible, but base <= 0": "#FEF9C3", "Base <= 0": "#FEF9C3", "No Comm (Margin Gap)": "#FEF9C3", "Not Eligible (<500K)": "#FEE2E2", "Not Eligible (<750K)": "#FEE2E2", "‼️ ขายขาดทุน (ตรวจสอบด่วน)": "#F87171", "จ่ายแล้ว": "#E5E7EB"}
        for status, color in status_colors.items(): tree.tag_configure(status, background=color)
        for col in cols_to_display:
            anchor = "center" if col in ["เลือก", "แผน"] else "w"; width = 60 if col == "เลือก" else 80 if col == "แผน" else 150; tree.heading(col, text=col, anchor=anchor)
            if any(s in col for s in ["ยอดขาย", "ต้นทุน", "กำไร", "Margin", "ค่าคอม", "ยอดรวม"]): tree.column(col, anchor="e", width=width)
            else: tree.column(col, anchor=anchor, width=width)
        for index, row in df.iterrows():
            tag = row.get("สถานะคำนวณ", ""); display_values = []
            if not is_paid_view: display_values.append("☐")
            for col_name in cols:
                if col_name == "record_id": continue
                value = row[col_name]
                if pd.notna(value):
                    if col_name in ['id', 'commission_month', 'commission_year', 'original_id']: display_values.append(int(value))
                    elif pd.api.types.is_float_dtype(value) or isinstance(value, (float, np.floating)): display_values.append(f"{value:,.2f}")
                    else: display_values.append(str(value))
                else: display_values.append("")
            tree.insert("", "end", values=display_values, tags=(tag,), iid=str(row['record_id']))
        if not is_paid_view: tree.bind("<Button-1>", self._on_payout_tree_click)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview); h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview); tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set); v_scroll.grid(row=0, column=1, sticky='ns'); h_scroll.grid(row=1, column=0, sticky='ew')
        
        
        self.select_all_checkbox.configure(state="normal" if not is_paid_view and not df.empty else "disabled")

    def _on_payout_tree_click(self, event):
        if self.payout_tree.identify_region(event.x, event.y) != "cell": return
        if self.payout_tree.identify_column(event.x) == "#1": 
            record_iid = self.payout_tree.focus();
            if not record_iid: return
            record_id_int = int(record_iid); current_values = list(self.payout_tree.item(record_iid, "values")); new_val = "☑" if current_values[0] == "☐" else "☐"
            if new_val == "☑": self.selected_payout_ids.add(record_id_int)
            else: self.selected_payout_ids.discard(record_id_int)
            current_values[0] = new_val; self.payout_tree.item(record_iid, values=current_values)
            self.select_all_var.set(1 if len(self.selected_payout_ids) == len(self.payout_tree.get_children()) else 0)
            self._update_payout_summary()

    def _update_payout_summary(self):
        count = len(self.selected_payout_ids)
        if count == 0: total_payout = 0.0; self.confirm_payout_button.configure(state="disabled"); self.export_payout_button.configure(state="disabled")
        else:
            selected_df = self.commission_results_df[self.commission_results_df['record_id'].isin(self.selected_payout_ids)]; total_payout = selected_df['ค่าคอมที่คำนวณได้'].sum()
            self.confirm_payout_button.configure(state="normal"); self.export_payout_button.configure(state="normal")
        self.total_selected_label.configure(text=f"จำนวนที่เลือก: {count} รายการ"); self.total_payout_label.configure(text=f"ยอดรวมที่จะจ่าย: {total_payout:,.2f} บาท")

    def _confirm_payout_for_selected(self):
        if not self.selected_payout_ids: messagebox.showwarning("ไม่ได้เลือกรายการ", "กรุณาเลือกรายการที่ต้องการยืนยันการจ่ายเงิน", parent=self); return
        df_to_process = self.commission_results_df[self.commission_results_df['record_id'].isin(self.selected_payout_ids)]
        self._confirm_and_save_commissions(self.selected_payout_ids, df_to_process)

    def _export_selected_payout(self):
        if not self.selected_payout_ids: messagebox.showwarning("ไม่ได้เลือกรายการ", "กรุณาเลือกรายการที่ต้องการ Export", parent=self); return
        df_to_export = self.commission_results_df[self.commission_results_df['record_id'].isin(self.selected_payout_ids)].copy()
        df_for_export_display = df_to_export.drop(columns=['record_id', 'เลือก'], errors='ignore')
        total_payout = df_for_export_display['ค่าคอมที่คำนวณได้'].sum()
        summary_values = [""] * len(df_for_export_display.columns)
        try:
            total_label_idx = df_for_export_display.columns.to_list().index('ค่าคอม (%)'); total_value_idx = df_for_export_display.columns.to_list().index('ค่าคอมที่คำนวณได้')
            summary_values[total_label_idx] = "ยอดรวม"; summary_values[total_value_idx] = total_payout
        except ValueError: pass
        summary_row = pd.Series(summary_values, index=df_for_export_display.columns)
        df_final_export = pd.concat([df_for_export_display, summary_row.to_frame().T], ignore_index=True)
        try:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="บันทึก Report การจ่ายค่าคอม", initialfile=f"payout_report_{datetime.now().strftime('%Y%m%d')}.xlsx")
            if save_path: df_final_export.to_excel(save_path, index=False); messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้วที่:\n{save_path}", parent=self)
        except Exception as e: messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=self); traceback.print_exc()

    # hr_screen.py (ภายในคลาส HRScreen)

    def _create_styled_dataframe_table(self, parent, df, title="", on_row_click=None, status_column=None, status_colors=None, iid_column=None):
        """Creates a styled ttk.Treeview table from a pandas DataFrame."""
        for widget in parent.winfo_children():
            widget.destroy()

        if title:
            label = CTkLabel(parent, text=title, font=CTkFont(size=14, weight="bold"))
            label.pack(pady=(5, 10))

        if df is None or df.empty:
            CTkLabel(parent, text="ไม่พบข้อมูล").pack(pady=20)
            return

        columns = df.columns.tolist()
        
        style = ttk.Style(parent)
        style.theme_use("clam")
        
        style.configure("Custom.Treeview.Heading", 
                        font=self.label_font_bold, 
                        background="#022c22",
                        foreground="white", 
                        relief="flat", 
                        padding=(10, 8))
        style.map("Custom.Treeview.Heading", background=[('active', "#065f46")])
        
        style.configure("Custom.Treeview", 
                        rowheight=32, 
                        font=self.small_font,
                        fieldbackground="#FFFFFF", 
                        foreground="#111827")
        style.map("Custom.Treeview", background=[('selected', "#3B82F6")])

        frame = CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=10, pady=5)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        tree = ttk.Treeview(frame, columns=columns, show='headings', style="Custom.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")
        
        tree.tag_configure('summary_row', background='#E5E7EB', font=CTkFont(size=12, weight="bold"))
        
        if status_colors:
            for tag_name, color in status_colors.items():
                tree.tag_configure(tag_name, background=color)

        column_configs = {
            'เลขที่ SO': {'width': 100, 'anchor': 'w'}, 'ยอดขาย/บริการ (ระบบ)': {'width': 165, 'anchor': 'e'},
            'ค่าขนส่ง (ระบบ)': {'width': 130, 'anchor': 'e'}, 'ค่าย้าย (ระบบ)': {'width': 130, 'anchor': 'e'},
            'ยอดขายรวม (ระบบ)': {'width': 140, 'anchor': 'e'}, 'ยอดขาย (Express)': {'width': 145, 'anchor': 'e'},
            'ต้นทุน (ระบบ)': {'width': 130, 'anchor': 'e'}, 'ต้นทุน (Express)': {'width': 140, 'anchor': 'e'},
            'ผลต่างยอดขาย': {'width': 120, 'anchor': 'e'}, 'ผลต่างต้นทุน': {'width': 110, 'anchor': 'e'},
            'แหล่งยอดขาย': {'width': 100, 'anchor': 'center'}, 'แหล่งต้นทุน': {'width': 100, 'anchor': 'center'},
            'สถานะ': {'width': 240, 'anchor': 'w'}
        }

        for col_id in columns:
            config = column_configs.get(col_id, {'width': 120, 'anchor': 'w'})
            tree.heading(col_id, text=col_id, anchor='center')
            can_stretch = col_id in ['เลขที่ SO', 'สถานะ']
            tree.column(col_id, width=config['width'], anchor=config['anchor'], stretch=can_stretch)

        for index, row in df.iterrows():
            tags_tuple = ()
            status_val_str = str(row.get(status_column, ''))

            # --- START: แก้ไข Logic การกำหนด Tag สี ---
            if 'ยอดรวม (Total)' in str(row.iloc[0]):
                tags_tuple += ('summary_row',)
            elif status_colors and status_column:
                tag_to_apply = None
                # ตรวจสอบสถานะแบบพิเศษก่อน
                if status_val_str.startswith("ผ่านเกณฑ์ (โอนเกิน"):
                    tag_to_apply = "ผ่านเกณฑ์" # ใช้ Tag สีเขียวของ "ผ่านเกณฑ์"
                elif status_val_str.startswith("⚠️ ยอดโอนขาด"):
                    tag_to_apply = "ยอดโอนขาด"
                elif status_val_str.startswith("✅ ยอดโอนเกิน"):
                    tag_to_apply = "ยอดโอนเกิน"
                # ถ้าไม่เข้าเงื่อนไขพิเศษ ให้ตรวจสอบแบบปกติ
                elif status_val_str in status_colors:
                    tag_to_apply = status_val_str
                
                if tag_to_apply:
                    tags_tuple = (tag_to_apply,)
            # --- END ---

            values = []
            for col_name in columns:
                value = row[col_name]
                if pd.notna(value):
                    if isinstance(value, (datetime, pd.Timestamp)):
                        values.append(value.strftime('%d/%m/%Y %H:%M'))
                    elif isinstance(value, (float, np.floating)):
                        values.append(f"{value:,.2f}")
                    else:
                        values.append(str(value))
                else:
                    values.append("")

            iid_value = row.get(iid_column, str(index)) if iid_column else str(index)
            tree.insert("", "end", values=values, tags=tags_tuple, iid=str(iid_value))
        
        v_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        if on_row_click:
            tree.bind("<Double-1>", lambda e: on_row_click(e, tree, df))

    def _get_archive_date_range(self, year, month=None):
        """สร้างช่วงวันที่เริ่มต้นและสิ้นสุดสำหรับการ Archive"""
        if month:
            # รายเดือน
            start_date = datetime(year, month, 1, 0, 0, 0)
            end_date = datetime(year, month, calendar.monthrange(year, month)[1], 23, 59, 59)
        else:
            # รายปี
            start_date = datetime(year, 1, 1, 0, 0, 0)
            end_date = datetime(year, 12, 31, 23, 59, 59)
        return start_date.strftime('%Y-%m-%d %H:%M:%S'), end_date.strftime('%Y-%m-%d %H:%M:%S')

    def _export_table_to_archive(self, table_name, archive_path, start_date, end_date, sale_key=None, columns='*'):
        """Export ข้อมูลจากตารางที่ระบุไปยังไฟล์ Excel พร้อมเปลี่ยนหัวข้อเป็นภาษาไทย และกรองตาม sale_key และ columns (ถ้ามี)"""
        
        # --- START: โค้ดที่แก้ไข ---
        # เปลี่ยน SELECT * เป็น SELECT {columns} เพื่อความยืดหยุ่น
        column_selection = ", ".join(columns) if isinstance(columns, list) else "*"
        query = f"SELECT {column_selection} FROM {table_name} WHERE timestamp BETWEEN %(start)s AND %(end)s"
        params = {'start': start_date, 'end': end_date}

        if sale_key and table_name == 'commissions':
            query += " AND sale_key = %(sale_key)s"
            params['sale_key'] = sale_key
            
        df = pd.read_sql_query(query, self.pg_engine, params=params)
        # --- END: สิ้นสุดโค้ดที่แก้ไข ---

        if not df.empty:
            header_map = self.app_container.HEADER_MAP
            rename_dict = {db_col: thai_name for db_col, thai_name in header_map.items() if db_col in df.columns}
            df.rename(columns=rename_dict, inplace=True)
            df.to_excel(archive_path, index=False)
            print(f"Exported {len(df)} rows from {table_name} to {archive_path}")
            return len(df)
        return 0

    def _delete_archived_data(self, conn, table_name, start_date, end_date):
        """ลบข้อมูลที่ถูก Archive แล้วออกจากฐานข้อมูล"""
        with conn.cursor() as cursor:
            cursor.execute(f"DELETE FROM {table_name} WHERE timestamp BETWEEN %s AND %s", 
                           (start_date, end_date))
            deleted_count = cursor.rowcount
            # บันทึกกิจกรรมการลบ
            summary = json.dumps({'period': f"{start_date} to {end_date}", 'deleted_count': deleted_count})
            cursor.execute("INSERT INTO audit_log (action, table_name, user_info, summary_json) VALUES (%s, %s, %s, %s)",
                           ('Annual Archive Delete', table_name, self.user_key, summary))
            print(f"Deleted {deleted_count} rows from {table_name}")
            return deleted_count
    
    # นำฟังก์ชันนี้ไปวางทับฟังก์ชัน _annual_archive_data เดิม
    def _annual_archive_data(self):
        dialog = AnnualArchiveDialog(self, datetime.now().year)
        self.wait_window(dialog)
        archive_config = dialog.result

        if not archive_config:
            return  # ผู้ใช้กดยกเลิก

        mode, year, month = archive_config["mode"], archive_config["year"], archive_config["month"]

        # --- ยืนยันการทำงาน ---
        period_text = ""
        if mode == "monthly": period_text = f"เดือน {self.thai_months[month - 1]} ปี {year}"
        elif mode == "annual": period_text = f"ทั้งปี {year} (ไฟล์รวม)"
        elif mode == "annual_by_month": period_text = f"ทั้งปี {year} (แยกไฟล์รายเดือน)"

        msg = (f"คุณต้องการ Export ข้อมูลของ '{period_text}' ใช่หรือไม่?\n\n"
               "**ขั้นตอนนี้จะยังไม่ลบข้อมูลออกจากระบบ**")
        if not messagebox.askyesno("ยืนยันการ Export ข้อมูล", msg, icon="question", parent=self):
            return

        # --- ถามผู้ใช้ว่าจะบันทึกไฟล์ไว้ที่ไหน ---
        messagebox.showinfo("เลือกโฟลเดอร์", "ขั้นตอนต่อไป โปรดเลือกโฟลเดอร์หลักที่จะใช้เก็บไฟล์ Archive", parent=self)
        base_archive_path = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์ Archive")

        if not base_archive_path:
            messagebox.showinfo("ยกเลิก", "การ Export ถูกยกเลิก", parent=self)
            return

        loading_popup = utils.show_loading_popup(self, "กำลัง Export ข้อมูล...")
        
        try:
            total_files_created = 0
            
            # --- Logic หลักในการ Export ---
            if mode == "annual_by_month":
                # โหมดใหม่: วนลูป 12 เดือน
                for m in range(1, 13):
                    loading_popup.lift()
                    month_name = self.thai_months[m - 1]
                    loading_popup.label.configure(text=f"กำลัง Export เดือน: {month_name}...")
                    self.update_idletasks()

                    month_folder_name = f"{m:02d}-{month_name}"
                    archive_dir = os.path.join(base_archive_path, str(year), month_folder_name)
                    os.makedirs(archive_dir, exist_ok=True)
                    
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    start_date, end_date = self._get_archive_date_range(year, m)
                    
                    # Export SO และ Commission ของแต่ละเซลส์ในเดือนนั้นๆ
                    sales_keys_df = pd.read_sql_query("SELECT DISTINCT sale_key FROM commissions WHERE timestamp BETWEEN %s AND %s", self.pg_engine, params=(start_date, end_date))
                    if not sales_keys_df.empty:
                        for sale_key in sales_keys_df['sale_key']:
                            period_suffix = f"_{sale_key}_{year}_{m:02d}"
                            so_path = os.path.join(archive_dir, f"SOs{period_suffix}_{ts}.xlsx")
                            if self._export_table_to_archive('commissions', so_path, start_date, end_date, sale_key=sale_key, columns=['so_number', 'bill_date', 'customer_name', 'sale_key', 'status', 'sales_service_amount', 'shipping_cost', 'cutting_drilling_fee', 'other_service_fee', 'credit_card_fee', 'transfer_fee', 'brokerage_fee', 'giveaways', 'coupons', 'wht_3_percent', 'total_payment_amount', 'payment_date', 'payment_before_vat', 'payment_no_vat', 'commission_month', 'commission_year', 'timestamp']) > 0:
                                total_files_created += 1
                            
                            comm_path = os.path.join(archive_dir, f"Commissions{period_suffix}_{ts}.xlsx")
                            if self._export_table_to_archive('commissions', comm_path, start_date, end_date, sale_key=sale_key, columns=['so_number', 'final_sales_amount', 'final_cost_amount', 'final_gp', 'final_margin', 'final_commission']) > 0:
                                total_files_created += 1

                    # Export PO ของเดือนนั้นๆ (ไฟล์รวม)
                    po_period_suffix = f"_{year}_{m:02d}"
                    po_path = os.path.join(archive_dir, f"purchase_orders{po_period_suffix}_{ts}.xlsx")
                    if self._export_table_to_archive('purchase_orders', po_path, start_date, end_date) > 0:
                        total_files_created += 1
            
            else: # โหมดเดิม (รายเดือน หรือ รายปีไฟล์รวม)
                archive_dir = os.path.join(base_archive_path, str(year))
                if mode == "monthly":
                    month_name = self.thai_months[month - 1]
                    archive_dir = os.path.join(archive_dir, f"{month:02d}-{month_name}")
                os.makedirs(archive_dir, exist_ok=True)
                
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                start_date, end_date = self._get_archive_date_range(year, month)
                
                sales_keys_df = pd.read_sql_query("SELECT DISTINCT sale_key FROM commissions WHERE timestamp BETWEEN %s AND %s", self.pg_engine, params=(start_date, end_date))
                if not sales_keys_df.empty:
                    for sale_key in sales_keys_df['sale_key']:
                        period_suffix = f"_{sale_key}_{year}" + (f"_{month:02d}" if month else "")
                        so_path = os.path.join(archive_dir, f"SOs{period_suffix}_{ts}.xlsx")
                        if self._export_table_to_archive('commissions', so_path, start_date, end_date, sale_key=sale_key, columns=['so_number', 'bill_date', 'customer_name', 'sale_key', 'status', 'sales_service_amount', 'shipping_cost', 'cutting_drilling_fee', 'other_service_fee', 'credit_card_fee', 'transfer_fee', 'brokerage_fee', 'giveaways', 'coupons', 'wht_3_percent', 'total_payment_amount', 'payment_date', 'payment_before_vat', 'payment_no_vat', 'commission_month', 'commission_year', 'timestamp']) > 0:
                            total_files_created += 1
                        
                        comm_path = os.path.join(archive_dir, f"Commissions{period_suffix}_{ts}.xlsx")
                        if self._export_table_to_archive('commissions', comm_path, start_date, end_date, sale_key=sale_key, columns=['so_number', 'final_sales_amount', 'final_cost_amount', 'final_gp', 'final_margin', 'final_commission']) > 0:
                            total_files_created += 1

                po_period_suffix = f"_{year}" + (f"_{month:02d}" if month else "")
                po_path = os.path.join(archive_dir, f"purchase_orders{po_period_suffix}_{ts}.xlsx")
                if self._export_table_to_archive('purchase_orders', po_path, start_date, end_date) > 0:
                    total_files_created += 1

            loading_popup.destroy()

            if total_files_created == 0:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบข้อมูลในช่วงเวลาที่เลือก ({period_text})", parent=self)
                return
            
            success_msg = (f"Export ข้อมูลสำเร็จ! (รวม {total_files_created} ไฟล์)\n\n"
                           f"ไฟล์ทั้งหมดถูกบันทึกที่:\n{base_archive_path}\n\n"
                           "**หมายเหตุ: ข้อมูลยังไม่ได้ถูกลบออกจากระบบ**")
            messagebox.showinfo("Export สำเร็จ", success_msg, parent=self)

        except Exception as e:
            if 'loading_popup' in locals() and loading_popup.winfo_exists():
                loading_popup.destroy()
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดระหว่างการ Export: {e}", parent=self)
            traceback.print_exc()

    def _create_cancelled_so_tab(self, parent_tab):
        # --- Grid Setup: แบ่งหน้าจอเป็น 2 ส่วน (บน-เล็ก / ล่าง-ใหญ่) ---
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(0, weight=0) # ส่วนค้นหา (ความสูงคงที่)
        parent_tab.grid_rowconfigure(1, weight=1) # ส่วนตาราง (ขยายเต็มที่)

        # =========================================================
        #  SECTION 1: Professional Action Bar (แถบเครื่องมือด้านบน)
        # =========================================================
        # ใช้ Frame ที่มีสีพื้นหลัง (เช่น สีเทาอ่อน) เพื่อให้ดูเป็นสัดส่วนเหมือน Toolbar
        action_bar = CTkFrame(parent_tab, height=60, fg_color=("gray90", "gray16"), corner_radius=6)
        action_bar.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        
        # ใช้ Grid ภายใน Action Bar เพื่อจัดกึ่งกลางแนวตั้ง (Vertical Center) ได้ง่ายกว่า Pack
        action_bar.grid_columnconfigure(3, weight=1) # ช่องว่างตรงกลางให้ยืดได้
        
        # [1] Label
        CTkLabel(action_bar, text="🔎 ค้นหา SO:", font=self.label_font_bold).grid(row=0, column=0, padx=(20, 5), pady=10, sticky="w")
        
        # [2] Search Input (ยาวขึ้นและสูงขึ้นเล็กน้อยให้กดง่าย)
        self.cancel_search_entry = CTkEntry(action_bar, placeholder_text="ระบุเลข SO... (เช่น SO6701-001)", width=250, height=34)
        self.cancel_search_entry.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        self.cancel_search_entry.bind("<Return>", lambda e: self._search_so_to_cancel())
        
        # [3] Search Button (สีฟ้าเด่น)
        CTkButton(action_bar, text="ค้นหา", command=self._search_so_to_cancel, 
                  width=100, height=34, fg_color="#3B82F6", hover_color="#2563EB", font=self.label_font_bold).grid(row=0, column=2, padx=10, pady=10, sticky="w")

        # [4] Inline Result Area (แสดงผลลัพธ์ต่อท้ายปุ่มค้นหาเลย)
        self.inline_result_frame = CTkFrame(action_bar, fg_color="transparent", height=34)
        self.inline_result_frame.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        # [5] Refresh Button (ขวาสุด)
        CTkButton(action_bar, text="⟳ รีเฟรช", command=self._load_cancelled_so_history, 
                  width=90, height=34, fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).grid(row=0, column=4, padx=20, pady=10, sticky="e")

        # =========================================================
        #  SECTION 2: Full-Width History Table (ตารางเต็มจอ)
        # =========================================================
        # Container สำหรับตาราง
        table_container = CTkFrame(parent_tab, fg_color="transparent")
        table_container.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        
        # Header + Toggle
        header_row = CTkFrame(table_container, fg_color="transparent", height=36)
        header_row.pack(fill="x", pady=(0, 5))
        CTkLabel(header_row, text="📜 ประวัติรายการยกเลิก", font=self.header_font_table, text_color="#EF4444").pack(side="left")

        self._cancel_view_mode = "normal"
        toggle_frame = CTkFrame(header_row, fg_color="transparent")
        toggle_frame.pack(side="right")
        self._toggle_normal_btn = CTkButton(
            toggle_frame, text="ยกเลิก SO", width=110, height=30,
            fg_color="#EF4444", hover_color="#DC2626", text_color="white",
            font=CTkFont(size=12, weight="bold"),
            command=lambda: self._switch_cancel_view("normal"))
        self._toggle_normal_btn.pack(side="left", padx=(0, 4))
        self._toggle_transport_btn = CTkButton(
            toggle_frame, text="🚚 SO ค่าขนส่ง", width=120, height=30,
            fg_color="transparent", border_width=1, border_color="#D97706",
            text_color="#D97706", font=CTkFont(size=12),
            command=lambda: self._switch_cancel_view("transport"))
        self._toggle_transport_btn.pack(side="left")

        # Frame สำหรับวาง Treeview
        self.cancelled_history_frame = CTkFrame(table_container, fg_color="transparent")
        self.cancelled_history_frame.pack(fill="both", expand=True)

        # โหลดข้อมูลเริ่มต้น
        self.after(100, self._load_cancelled_so_history)

    def _switch_cancel_view(self, mode):
        self._cancel_view_mode = mode
        if mode == "normal":
            self._toggle_normal_btn.configure(fg_color="#EF4444", text_color="white")
            self._toggle_transport_btn.configure(fg_color="transparent", text_color="#D97706")
            self._load_cancelled_so_history()
        else:
            self._toggle_transport_btn.configure(fg_color="#D97706", text_color="white")
            self._toggle_normal_btn.configure(fg_color="transparent", text_color="#EF4444")
            self._load_transport_so_history()

    def _ensure_transport_cancel_table(self):
        """สร้าง transport_cancel_requests ถ้ายังไม่มี"""
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS transport_cancel_requests (
                        id SERIAL PRIMARY KEY,
                        so_number VARCHAR(50) NOT NULL,
                        so_id INTEGER,
                        sale_key VARCHAR(50),
                        customer_name VARCHAR(200),
                        original_status VARCHAR(50),
                        requested_by VARCHAR(100),
                        requested_at TIMESTAMP DEFAULT NOW(),
                        status VARCHAR(20) DEFAULT 'pending',
                        reviewed_by VARCHAR(100),
                        reviewed_at TIMESTAMP
                    )
                """)
            conn.commit()
        except Exception:
            pass
        finally:
            if conn: self.app_container.release_connection(conn)

    def _load_transport_so_history(self):
        """โหลดประวัติ SO ค่าขนส่งที่ SM อนุมัติยกเลิกแล้ว"""
        for widget in self.cancelled_history_frame.winfo_children():
            widget.destroy()
        self._ensure_transport_cancel_table()
        try:
            query = """
                SELECT so_number, sale_key, customer_name,
                       requested_by, reviewed_by, requested_at
                FROM transport_cancel_requests
                WHERE status = 'approved'
                ORDER BY requested_at DESC
                LIMIT 100
            """
            df = pd.read_sql_query(query, self.pg_engine)
            if df.empty:
                CTkLabel(self.cancelled_history_frame, text="ยังไม่มีประวัติ SO ค่าขนส่ง",
                         font=self.entry_font, text_color="gray").pack(pady=40)
                return
            df.rename(columns={
                'so_number': 'เลขที่ SO',
                'sale_key': 'รหัสพนักงาน',
                'customer_name': 'ชื่อลูกค้า',
                'requested_by': 'ขอโดย HR',
                'reviewed_by': 'อนุมัติโดย SM',
                'requested_at': 'วันที่ขอ',
            }, inplace=True)
            self._create_styled_dataframe_table(self.cancelled_history_frame, df, title="")
            tree = None
            for widget in self.cancelled_history_frame.winfo_children():
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Treeview):
                        tree = child
                        break
                if tree: break
            if tree:
                tree.column('เลขที่ SO', width=120, stretch=False, anchor="center")
                tree.column('รหัสพนักงาน', width=100, stretch=False, anchor="center")
                tree.column('วันที่ขอ', width=150, stretch=False, anchor="center")
                tree.column('ชื่อลูกค้า', width=200, stretch=True)
                tree.column('ขอโดย HR', width=120, stretch=False, anchor="center")
                tree.column('อนุมัติโดย SM', width=120, stretch=False, anchor="center")
        except Exception as e:
            CTkLabel(self.cancelled_history_frame, text=f"โหลดข้อมูลล้มเหลว: {e}",
                     text_color="red").pack(pady=20)
            traceback.print_exc()

    # -------------------------------------------------------------------------
    #  ฟังก์ชันค้นหา (ปรับปรุงให้แสดงผลใน Inline Frame)
    # -------------------------------------------------------------------------
    def _search_so_to_cancel(self):
        so_number = self.cancel_search_entry.get().strip().upper()
        
        # ล้างผลเก่า
        for w in self.inline_result_frame.winfo_children(): w.destroy()
        
        if not so_number: return

        try:
            # ดึงข้อมูลเพื่อแสดงผล
            query = "SELECT id, customer_name, sale_key, status, sales_service_amount FROM commissions WHERE so_number = %s"
            df = pd.read_sql_query(query, self.pg_engine, params=(so_number,))
            
            if df.empty:
                CTkLabel(self.inline_result_frame, text=f"❌ ไม่พบ SO: {so_number}", text_color="#EF4444", font=self.label_font_bold).pack(side="left")
                return

            row = df.iloc[0]
            status = row['status']
            
            # การ์ดแสดงผลง่ายๆ แนวนอน (Compact Info)
            info_text = f"พบข้อมูล: {row['customer_name']} (ยอด: {row['sales_service_amount']:,.2f})  |  สถานะ: {status}"
            CTkLabel(self.inline_result_frame, text=info_text, font=self.small_font).pack(side="left", padx=(0, 15))

            # ปุ่ม Action (แสดงเฉพาะเมื่อยกเลิกได้)
            if status == 'Pending Transport Cancel':
                CTkLabel(self.inline_result_frame, text="🚚 รอ SM อนุมัติยกเลิกค่ารถ...",
                         text_color="#D97706").pack(side="left")
            elif status not in ['Paid', 'HR Verified', 'Cancelled']:
                CTkButton(self.inline_result_frame, text="⚠️ ยกเลิกรายการนี้",
                          fg_color="#DC2626", hover_color="#B91C1C", height=32,
                          command=lambda: self._confirm_cancel_so(so_number)).pack(side="left")
            else:
                reason_msg = "(ยกเลิกไม่ได้: จ่ายแล้ว/ยืนยันแล้ว)" if status != 'Cancelled' else "(ยกเลิกไปแล้ว)"
                CTkLabel(self.inline_result_frame, text=reason_msg, text_color="gray").pack(side="left")

        except Exception as e:
            messagebox.showerror("Error", f"{e}")

    def _confirm_cancel_so(self, so_number):
        """เรียก Dialog ถามเหตุผลและบันทึก"""
        CancellationReasonDialog(self, lambda reason: self._process_cancellation_callback(so_number, reason))

    def _process_cancellation_callback(self, so_number, reason):
        """Callback หลังจากกดตกลงใน Dialog"""
        if reason == "ค่ารถไม่อยู่ในเงื่อนไขคอมมิชชั่น":
            self._request_transport_cancel(so_number)
        else:
            self._cancel_so_logic(so_number, reason)

        self.cancel_search_entry.delete(0, "end")
        for widget in self.inline_result_frame.winfo_children(): widget.destroy()
        self._load_cancelled_so_history()

    def _request_transport_cancel(self, so_number):
        """ส่งคำขอยกเลิก SO ค่ารถ → รอ SM อนุมัติ (ไม่ยกเลิกทันที)"""
        self._ensure_transport_cancel_table()
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    "SELECT id, sale_key, customer_name, status FROM commissions WHERE so_number = %s",
                    (so_number,)
                )
                result = cursor.fetchone()
                if not result:
                    messagebox.showerror("Error", "ไม่พบ SO นี้ในระบบ")
                    return
                so_id, sale_key, customer_name, original_status = result

                cursor.execute("""
                    INSERT INTO transport_cancel_requests
                        (so_number, so_id, sale_key, customer_name, original_status, requested_by)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (so_number, so_id, sale_key, customer_name, original_status, self.user_name))

                # พัก SO ไว้ในสถานะรอ — ยังไม่ยกเลิก
                cursor.execute("""
                    UPDATE commissions SET status = 'Pending Transport Cancel'
                    WHERE so_number = %s
                """, (so_number,))

                # แจ้ง SM ทุกคนให้อนุมัติ
                cursor.execute(
                    "SELECT sale_key FROM sales_users WHERE role = 'Sales Manager' AND status = 'Active'"
                )
                manager_keys = [r[0] for r in cursor.fetchall()]
                msg = (f"[TRANSPORT_CANCEL] SO: {so_number} ขอยกเลิกเหตุผล: ค่ารถไม่อยู่ในเงื่อนไขคอมมิชชั่น\n"
                       f"เจ้าของ SO: {sale_key} | ลูกค้า: {customer_name}\n"
                       f"ขอโดย HR: {self.user_name} — กรุณาตรวจสอบในหน้ารายการรออนุมัติ")
                for mgr_key in manager_keys:
                    cursor.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id, timestamp)
                        VALUES (%s, %s, FALSE, %s, NOW())
                    """, (mgr_key, msg, so_id))

            conn.commit()
            messagebox.showinfo("ส่งคำขอสำเร็จ",
                                f"SO: {so_number} ส่งคำขอยกเลิกค่ารถให้ SM อนุมัติแล้ว\nรอ Sale Manager ตรวจสอบ")
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _load_cancelled_so_history(self):
        """โหลดตารางประวัติ พร้อมบังคับขยายคอลัมน์ให้เต็มจอ"""
        # ล้างตารางเก่า
        for widget in self.cancelled_history_frame.winfo_children(): widget.destroy()
        
        try:
            # 1. Query ข้อมูล
            query = """
                SELECT so_number, sale_key, customer_name, rejection_reason, timestamp 
                FROM commissions 
                WHERE status = 'Cancelled' 
                ORDER BY timestamp DESC 
                LIMIT 100
            """
            df = pd.read_sql_query(query, self.pg_engine)
            
            if df.empty:
                CTkLabel(self.cancelled_history_frame, text="ยังไม่มีประวัติการยกเลิก", font=self.entry_font, text_color="gray").pack(pady=40)
                return
            
            # เปลี่ยนชื่อคอลัมน์ให้สื่อความหมาย
            df.rename(columns={
                'so_number': 'เลขที่ SO',
                'sale_key': 'รหัสพนักงาน',
                'customer_name': 'ชื่อลูกค้า',
                'rejection_reason': 'สาเหตุการยกเลิก',
                'timestamp': 'วันที่ยกเลิก'
            }, inplace=True)

            # 2. สร้างตารางด้วย Helper Function เดิม
            self._create_styled_dataframe_table(
                self.cancelled_history_frame, 
                df, 
                title="" # ไม่ต้องใส่ Title ซ้ำ เพราะเราทำ Header ข้างนอกแล้ว
            )
            
            # =============================================================
            # [🔥 HERO FIX] เทคนิคแก้พื้นที่ขาว: เจาะเข้าไปแก้ Treeview ให้ยืดคอลัมน์
            # =============================================================
            # หาตัว Treeview widget ที่ถูกสร้างโดย _create_styled_dataframe_table
            tree = None
            for widget in self.cancelled_history_frame.winfo_children():
                # ปกติ Helper จะสร้าง Frame ครอบ Treeview อีกที
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Treeview):
                        tree = child
                        break
                if tree: break
            
            if tree:
                # กำหนดความกว้างและการยืดตัว (Stretch) ใหม่ให้สวยงาม
                # เลขที่ SO (Fixed)
                tree.column('เลขที่ SO', width=120, stretch=False, anchor="center")
                # รหัสพนักงาน (Fixed)
                tree.column('รหัสพนักงาน', width=100, stretch=False, anchor="center")
                # วันที่ (Fixed)
                tree.column('วันที่ยกเลิก', width=150, stretch=False, anchor="center")
                
                # *** คอลัมน์พระเอก: ให้ยืดกินพื้นที่ที่เหลือทั้งหมด ***
                tree.column('ชื่อลูกค้า', width=200, stretch=True) 
                tree.column('สาเหตุการยกเลิก', width=300, stretch=True)
            # =============================================================
            
        except Exception as e:
            CTkLabel(self.cancelled_history_frame, text=f"โหลดข้อมูลล้มเหลว: {e}", text_color="red").pack(pady=20)
            traceback.print_exc()

