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
from history_windows import PurchaseDetailWindow
from custom_widgets import NumericEntry, DateSelector
import utils
import business_logic
# --- DIALOG CLASSES ---

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
        
        try:
            if file_path.endswith('.csv'):
                with open(file_path, 'rb') as f: result = chardet.detect(f.read())
                df = pd.read_csv(file_path, encoding=result['encoding'])
            else:
                df = pd.read_excel(file_path)
            
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

class HRScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master, corner_radius=0, fg_color=app_container.THEME["hr"]["bg"])
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role

        self.label_font = CTkFont(size=16, weight="bold", family="Roboto")
        self.entry_font = CTkFont(size=14, family="Roboto")
        self.header_font_table = CTkFont(size=14, weight="bold", family="Roboto")
        self.label_font_bold = CTkFont(size=12, weight="bold", family="Roboto")
        self.small_font = CTkFont(size=12, family="Roboto")


        self.header_map = app_container.HEADER_MAP

        self.sales_keys_list = self._get_sale_keys()

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
        self.period_options = ["ปีนี้", "เดือนนี้", "Q1", "Q2", "Q3", "Q4"] + self.thai_months
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}

        self.history_current_page, self.history_rows_per_page, self.history_total_rows = 0, 20, 0
        self.user_current_page, self.user_rows_per_page, self.user_total_rows = 0, 20, 0

        self.edit_data_current_page = 0
        self.edit_data_rows_per_page = 15

        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        header_frame = CTkFrame(self, fg_color="transparent"); header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 0))
        CTkLabel(header_frame, text=f"หน้าจอสำหรับฝ่ายบุคคล (HR): {self.user_name}", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).pack(side="left")
        CTkButton(header_frame, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").pack(side="right")

        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1, segmented_button_selected_color=self.theme["primary"], segmented_button_unselected_hover_color="#A7F3D0", fg_color=self.cget("fg_color"), command=self._on_tab_selected)
        self.tab_view.grid(row=1, column=0, pady=10, padx=20, sticky="nsew")
        
        self.outstanding_tab = self.tab_view.add("ติดตามยอดค้างชำระ") 
        # สร้าง instance ของ Dashboard แล้วใส่เข้าไปใน Tab
        self.outstanding_dashboard = OutstandingDashboardTab(self.outstanding_tab, self.app_container)
        # --- ส่วนของการสร้าง Tab ที่แก้ไขลำดับแล้ว ---
        self.dashboard_tab = self.tab_view.add("Dashboard สรุปภาพรวม")
        self.sales_target_tab = self.tab_view.add("วิเคราะห์เป้าการขาย")
        self.manage_users_tab = self.tab_view.add("จัดการผู้ใช้งาน")

        # ✅ สร้าง Tab "Sales Mode" ขึ้นมาก่อน
        self.sales_mode_tab = self.tab_view.add("ลงข้อมูลแทนเซลส์ (Sales Mode)")
        # ✅ จากนั้นค่อยตั้งค่า (Configure) ให้กับ Tab ที่เพิ่งสร้าง
        self.sales_mode_tab.grid_rowconfigure(0, weight=1)
        self.sales_mode_tab.grid_columnconfigure(0, weight=1)

        self.pu_mode_tab = self.tab_view.add("ลงข้อมูลแทนจัดซื้อ (PU Mode)")
        self.pu_mode_tab.grid_rowconfigure(0, weight=1); self.pu_mode_tab.grid_columnconfigure(0, weight=1)

        self.edit_data_tab = self.tab_view.add("แก้ไขข้อมูล (SO/PO)")
        self.compare_commission_tab = self.tab_view.add("เปรียบเทียบ / ดูประวัติ")
        self.process_commission_tab = self.tab_view.add("ประมวลผลและจ่ายค่าคอม")
        self.payout_history_tab = self.tab_view.add("ประวัติการจ่ายค่าคอม")
        self.audit_log_tab = self.tab_view.add("บันทึกกิจกรรม")
        
        # --- สิ้นสุดส่วนที่แก้ไข ---
        
        self._create_dashboard_tab(self.dashboard_tab)
        self._create_sales_target_tab(self.sales_target_tab)
        self._create_manage_users_tab(self.manage_users_tab)
        self._create_edit_data_tab(self.edit_data_tab)
        self._create_compare_commission_tab(self.compare_commission_tab)
        self._create_process_commission_tab(self.process_commission_tab)
        self._create_payout_history_tab(self.payout_history_tab)
        self._create_audit_log_tab(self.audit_log_tab)

        self.tab_view.set("จัดการผู้ใช้งาน")
        self.after(100, self._initial_load)
        self._sales_mode_loaded = False 
        self._pu_mode_loaded = False 
        self._payout_history_loaded = False 
        self._dashboard_loaded, self._sales_target_loaded, self._users_loaded, self._compare_commission_loaded, self._process_commission_loaded, self._audit_log_loaded = False, False, False, False, False, False
    
    def _create_payout_history_table(self, df):
        """(เวอร์ชันปรับปรุง) เพิ่มคอลัมน์ ยอดขาย, Normal, BelowT"""
        for widget in self.payout_history_frame.winfo_children():
            widget.destroy()

        if df is None or df.empty:
            CTkLabel(self.payout_history_frame, text="ไม่พบข้อมูลตามเงื่อนไขที่เลือก").pack(pady=20)
            return

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

        tree_frame = CTkFrame(self.payout_history_frame, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # <<< START: 1. เพิ่มคอลัมน์ใหม่ 3 คอลัมน์ >>>
        columns = {
            'payout_period_text': 'รอบค่าคอม',  # <--- ✅ เพิ่ม
            'sale_key': 'รหัสพนักงาน', 'sale_name': 'ชื่อพนักงาน', 'plan_name': 'แผน',
            'sales_target': 'เป้าหมาย', 
            'total_sales': 'ยอดขาย',
            'total_normal_sales': 'Normal',
            'total_below_sales': 'BelowT',
            'timestamp': 'วันที่ทำรายการจ่าย', # <--- ✅ แก้ไขชื่อ
            'final_commission': 'ยอดคอม Gross',
            'incentives_total': 'Incentive', 'deductions_total': 'ยอดหัก', 'withholding_tax': 'หัก 3%', 'net_commission': 'ยอดโอนสุทธิ'
        }
        # <<< END >>>
        
        tree = ttk.Treeview(tree_frame, columns=list(columns.keys()), show='headings', style="Payout.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        tree.tag_configure('oddrow', background='#FFFFFF')
        tree.tag_configure('evenrow', background='#F0F9FF')

        for col_id, col_text in columns.items():
            anchor = 'w'
            width = 120 # Default

            if col_id == 'sale_name': 
                width = 200
            elif col_id == 'timestamp': 
                width = 110
            elif col_id == 'plan_name': 
                width = 80
            elif col_id == 'payout_period_text': # <--- ✅ เพิ่ม
                width = 130
            
            # <<< START: 2. ตั้งค่าให้คอลัมน์ใหม่เป็นชิดขวา และกำหนดขนาด >>>
            if col_id in ['sales_target', 'total_sales', 'final_commission',
                'incentives_total', 'deductions_total', 
                'withholding_tax', 'net_commission']: # <--- ✅ เพิ่ม 'withholding_tax' ที่นี่
                anchor = 'e'
                width = 130
            elif col_id in ['total_normal_sales', 'total_below_sales']:
                anchor = 'e'
                width = 100 
            # <<< END >>>
            
            tree.heading(col_id, text=col_text, anchor='center')
            tree.column(col_id, anchor=anchor, width=width, minwidth=60)

        for i, row in df.iterrows():
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'

            values = []
            for col_id in columns.keys():
                value = row[col_id]
                if pd.notna(value):
                    if isinstance(value, datetime): 
                        values.append(value.strftime('%d/%m/%Y')) # Format วันที่
                    elif isinstance(value, (float, np.floating, int)): # <<< 3. เพิ่ม int เข้าไปเผื่อ
                        values.append(f"{value:,.2f}")
                    else: 
                        values.append(str(value))
                else:
                    values.append("")
            
            tree.insert("", "end", values=values, iid=str(row['id']), tags=(tag,))
        
        tree.bind("<Double-1>", lambda e: self._on_payout_history_double_click(e, tree))

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        v_scroll.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=v_scroll.set)

    def _payout_prev_page(self):
        if self.history_current_page > 0:
            self.history_current_page -= 1
            self._load_payout_history()

    def _payout_next_page(self):
        total_pages = (self.history_total_rows + self.history_rows_per_page - 1) // self.history_rows_per_page
        if self.history_current_page < total_pages - 1:
            self.history_current_page += 1
            self._load_payout_history()

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
        selected_tab_name = self.tab_view.get()

        if selected_tab_name == "ลงข้อมูลแทนเซลส์ (Sales Mode)" and not self._sales_mode_loaded:
            try:
                from sales_proxy_screen import SalesProxyScreen
                self.sales_proxy_screen_instance = SalesProxyScreen(
                    master=self.sales_mode_tab,
                    app_container=self.app_container,
                    proxy_user_key=self.user_key,
                    proxy_user_name=self.user_name,
                    user_role=self.user_role, # <--- เพิ่มบรรทัดนี้เข้ามา
                    role_to_proxy="Sale"
                )
                self.sales_proxy_screen_instance.grid(row=0, column=0, sticky="nsew")
                self._sales_mode_loaded = True
            except ImportError:
                messagebox.showerror("ผิดพลาด", "ไม่พบไฟล์ sales_proxy_screen.py")
            except Exception as e:
                messagebox.showerror("ผิดพลาด", f"ไม่สามารถโหลดหน้าจอ Sales Mode ได้: {e}")
        
        elif selected_tab_name == "ลงข้อมูลแทนจัดซื้อ (PU Mode)" and not self._pu_mode_loaded:
            try:
                from purchasing_proxy_screen import PurchasingProxyScreen
                self.pu_proxy_screen_instance = PurchasingProxyScreen(
                    master=self.pu_mode_tab,
                    app_container=self.app_container,
                    proxy_user_key=self.user_key,
                    proxy_user_name=self.user_name,
                    role_to_proxy="Purchasing Staff"
                )
                self.pu_proxy_screen_instance.pack(fill="both", expand=True) # ใช้ .pack() แทน .grid()
                self._pu_mode_loaded = True
            except ImportError:
                messagebox.showerror("ผิดพลาด", "ไม่พบไฟล์ purchasing_proxy_screen.py")
            except Exception as e:
                messagebox.showerror("ผิดพลาด", f"ไม่สามารถโหลดหน้าจอ PU Mode ได้: {e}")

        elif selected_tab_name == "ประวัติการจ่ายค่าคอม" and not self._payout_history_loaded:
            self._load_payout_history()
            self._payout_history_loaded = True
        
        elif selected_tab_name == "Dashboard สรุปภาพรวม" and not self._dashboard_loaded:
            self._initial_load_dashboard()
            self._dashboard_loaded = True
            
        elif selected_tab_name == "วิเคราะห์เป้าการขาย" and not self._sales_target_loaded:
            self._initial_load_sales_target()
            self._sales_target_loaded = True
            
        elif selected_tab_name == "จัดการผู้ใช้งาน" and not self._users_loaded:
            self._populate_user_table()
            self._users_loaded = True

        elif selected_tab_name == "เปรียบเทียบ / ดูประวัติ" and not self._compare_commission_loaded:
            self._compare_commission_loaded = True
            self._compare_commission_loaded = True

        elif selected_tab_name == "ประมวลผลและจ่ายค่าคอม" and not self._process_commission_loaded:
            self._initial_load_process_commission()
            self._process_commission_loaded = True

        elif selected_tab_name == "บันทึกกิจกรรม" and not self._audit_log_loaded:
            self._populate_audit_log_table()
            self._audit_log_loaded = True

        if selected_tab_name == "Dashboard สรุปภาพรวม" and not self._dashboard_loaded:
            self._update_dashboard()
            self._dashboard_loaded = True
        elif selected_tab_name == "วิเคราะห์เป้าการขาย" and not self._sales_target_loaded:
            self._update_sales_target_dashboard()
            self._sales_target_loaded = True
        elif selected_tab_name == "จัดการผู้ใช้งาน" and not self._users_loaded:
            self._populate_users_table()
            self._users_loaded = True
        elif selected_tab_name == "เปรียบเทียบ / ดูประวัติ" and not self._compare_commission_loaded:
            self._compare_commission_loaded = True
        elif selected_tab_name == "ประมวลผลและจ่ายค่าคอม" and not self._process_commission_loaded:
            self._on_sale_selected_for_process()
            self._process_commission_loaded = True
        elif selected_tab_name == "ประวัติการจ่ายค่าคอม" and not self._payout_history_loaded:
            self._load_payout_history()
            self._payout_history_loaded = True
        elif selected_tab_name == "บันทึกกิจกรรม" and not self._audit_log_loaded:
            self._populate_audit_log_table()
            self._audit_log_loaded = True

    def _show_calculation_details(self):
        # --- แก้ไข 3 บรรทัดนี้ครับ ---
        if self.latest_commission_result: # <-- ✅ แก้ไขจาก initial_commission_result
            debug_df = self.latest_commission_result.get('debug_df') # <-- ✅ แก้ไข
            so_breakdown_df = self.latest_commission_result.get('so_breakdown_df') # <-- ✅ แก้ไข
            
            sale_key = self.selected_sale_for_process.get()
        # --- สิ้นสุดจุดแก้ไข ---
        
            plan_name = self.sales_user_info.get(sale_key, {}).get('plan', 'Unknown Plan')
            
            # ส่ง DataFrame ทั้งสองตัวไปที่หน้าต่าง Viewer
            CalculationDetailViewer(
                master=self, 
                debug_df=debug_df, 
                so_breakdown_df=so_breakdown_df, 
                plan_name=plan_name
            )
        else:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่พบข้อมูลการคำนวณ", parent=self)

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
        """(เวอร์ชันปรับปรุง) สร้าง Layout สำหรับหน้าประวัติการจ่ายเงิน"""
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

        # --- Frame สำหรับ Pagination ---
        pagination_frame = CTkFrame(parent_tab, fg_color="transparent")
        pagination_frame.grid(row=1, column=0, padx=10, pady=0, sticky="ew")

        self.payout_prev_button = CTkButton(pagination_frame, text="<<", command=self._payout_prev_page, width=50, state="disabled")
        self.payout_prev_button.pack(side="left")
        self.payout_page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.payout_page_label.pack(side="left", expand=True)
        self.payout_next_button = CTkButton(pagination_frame, text=">>", command=self._payout_next_page, width=50, state="disabled")
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
                    log.payout_period_text,   -- <<< ✅ เพิ่มบรรทัดนี้
                    log.sale_key, 
                    u.sale_name, 
                    log.plan_name, 
                    u.sales_target,
                    log.total_sales,
                    log.total_normal_sales,
                    log.total_below_sales,
                    log.timestamp,            -- (เก็บไว้เพื่อแสดงผล 'วันที่จ่าย')
                    log.final_commission, 
                    log.incentives_total,
                    log.deductions_total, 
                    log.withholding_tax,
                    log.net_commission,
                    log.commission_year,      -- <<< ✅ เพิ่มบรรทัดนี้
                    log.commission_month      -- <<< ✅ เพิ่มบรรทัดนี้
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

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดประวัติ: {e}", parent=self)
            traceback.print_exc()
            self._create_payout_history_table(pd.DataFrame())

    def _on_payout_history_double_click(self, event, tree):
        """(เวอร์ชันปรับปรุง) Callback เมื่อดับเบิลคลิกบนตารางประวัติการจ่ายเงิน"""
        selected_item_iid = tree.focus()
        if not selected_item_iid:
            return
        
        try:
            payout_id = int(selected_item_iid)
            PayoutDetailWindow(master=self, app_container=self.app_container, payout_id=payout_id)
        except (ValueError, TclError) as e:
            print(f"Invalid item selected: {selected_item_iid}, error: {e}")

    ### --- จุดที่แก้ไข --- ###
    # ผมได้รวมฟังก์ชัน _create_plan_a_summary_table และ _create_plan_b_summary_table
    # ให้เป็นฟังก์ชันเดียวคือ _create_commission_summary_table เพื่อลดความซ้ำซ้อนของโค้ด
    # และเพิ่มความยืดหยุ่นในการแสดงผล ไม่ว่าข้อมูลจะมี 2 หรือ 3 คอลัมน์ก็ตาม
    def _create_commission_summary_table(self, summary_df, container=None):
        """สร้างตารางสรุปผลการคำนวณค่าคอมมิชชั่นแบบไดนามิกตาม DataFrame ที่ได้รับ"""
        if container is None:
            container = self.process_result_frame
            
        for widget in container.winfo_children(): widget.destroy()

        if summary_df is None or summary_df.empty:
            CTkLabel(container, text="ไม่พบข้อมูลสำหรับสร้างสรุป").pack(pady=20)
            return

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
        tree = ttk.Treeview(tree_frame, columns=columns_to_show, show="headings", style="Summary.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

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
                refresh_callback=self._refresh_comparison_view
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
        
        # --- START: เพิ่ม Logic การตรวจสอบยอดค้างชำระเพื่อเปลี่ยนสี ---
        difference_amount = so_data.get('difference_amount', 0.0) or 0.0
        
        # ถ้า difference_amount > 0 (โอนขาด) ให้ใช้สีส้มอ่อน, ถ้าไม่ ให้ใช้สีฟ้าอ่อนปกติ
        card_color = "#FEF3C7" if difference_amount > 0 else "#F0F9FF"
        info_text_color = "#92400E" if difference_amount > 0 else "gray"
        
        so_card = CTkFrame(parent, border_width=1, fg_color=card_color)
        # --- END ---
        
        so_card.pack(fill="x", padx=10, pady=8)
        so_card.grid_columnconfigure(0, weight=1)
        
        header_frame = CTkFrame(so_card, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        header_frame.grid_columnconfigure(0, weight=1)
        
        # --- START: ปรับปรุงการแสดงข้อความให้มีข้อมูลยอดค้างชำระด้วย ---
        main_info_text = f"SO: {so_number}  |  ลูกค้า: {so_data.get('customer_name','N/A')}  |  เซลส์: {so_data.get('sale_key','N/A')}"
        CTkLabel(header_frame, text=main_info_text, font=self.entry_font).grid(row=0, column=0, sticky="w")

        # แสดงข้อความยอดค้างชำระ ถ้ามี
        if difference_amount < 0:
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
                        info = f"  - PO: {row['po_number']} | Supplier: {row['supplier_name']} | สถานะ: {row['status']}"
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

    def _create_sales_target_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1); parent_tab.grid_rowconfigure(1, weight=1); filter_frame = CTkFrame(parent_tab, fg_color="transparent"); filter_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew"); CTkLabel(filter_frame, text="ช่วงเวลา:", font=self.label_font).pack(side="left", padx=(5,10)); self.sales_target_period_var = tk.StringVar(value="เดือนนี้"); period_menu = CTkOptionMenu(filter_frame, variable=self.sales_target_period_var, values=self.period_options, command=lambda _: self._update_sales_target_dashboard()); period_menu.pack(side="left", padx=5); refresh_button = CTkButton(filter_frame, text="Refresh", width=100, fg_color=self.theme["primary"], command=self._update_sales_target_dashboard); refresh_button.pack(side="left", padx=20); self.sales_target_chart_frame = CTkFrame(parent_tab, border_width=1, corner_radius=10); self.sales_target_chart_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    def _update_sales_target_dashboard(self):
        loading = self._show_loading(self.sales_target_chart_frame)
        try:
            period = self.sales_target_period_var.get()
            
            # --- START: แก้ไขการส่งค่า ---
            # ลบ start_date, end_date ที่ไม่จำเป็นออก
            # ส่ง 'period' (string) ไปแทน
            sales_vs_target_data = self._get_sales_vs_target_data(period)
            # --- END ---
            
            loading.destroy()
            self._create_sales_vs_target_chart(self.sales_target_chart_frame, sales_vs_target_data)
        except Exception as e: 
            loading.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอัปเดต Dashboard: {e}", parent=self)
            traceback.print_exc() # เพิ่ม traceback
    
    def _initial_load_sales_target(self):
        """
        ฟังก์ชันสำหรับโหลดข้อมูลเริ่มต้นของแท็บ 'วิเคราะห์เป้าการขาย'
        เมื่อถูกเรียกครั้งแรก
        """
        # เรียกใช้ฟังก์ชันที่มีอยู่แล้วซึ่งทำหน้าที่โหลดและวาดกราฟ
        self._update_sales_target_dashboard()

    def _get_sales_vs_target_data(self, period): # <-- แก้ไข parameter
        try:
            # --- START: สร้าง logic การกรองใหม่ (เหมือนกับ _get_sales_by_employee_data) ---
            today = datetime.now()
            current_year = today.year
            params = []
            
            commission_filter_clauses = []
            if period == "เดือนนี้":
                commission_filter_clauses.append("c.commission_month = %s")
                params.append(today.month)
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "ปีนี้":
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            
            # --- START: เพิ่ม Logic ของ Q1-Q4 ตรงนี้ ---
            elif period == "Q1":
                commission_filter_clauses.append("c.commission_month IN (1, 2, 3)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "Q2":
                commission_filter_clauses.append("c.commission_month IN (4, 5, 6)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "Q3":
                commission_filter_clauses.append("c.commission_month IN (7, 8, 9)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "Q4":
                commission_filter_clauses.append("c.commission_month IN (10, 11, 12)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            # --- END: สิ้นสุด Logic Q1-Q4 ---
            
            elif period in self.thai_month_map:
                month_num = self.thai_month_map[period]
                commission_filter_clauses.append("c.commission_month = %s")
                params.append(month_num)
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year) # กรองตามปีปัจจุบัน
            else: # Fallback (เหมือน "เดือนนี้")
                commission_filter_clauses.append("c.commission_month = %s")
                params.append(today.month)
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
                
            commission_filter_sql = " AND ".join(commission_filter_clauses)
            # --- END ---

            query = f"""
                SELECT su.sale_name, su.sales_target, 
                       COALESCE(SUM(c.sales_service_amount), 0) as total_sales 
                FROM sales_users su 
                LEFT JOIN commissions c ON su.sale_key = c.sale_key 
                                     AND c.is_active = 1 
                                     AND {commission_filter_sql} -- <-- แก้ไขตรงนี้
                WHERE su.role = 'Sale' 
                  AND su.sales_target > 0 
                  AND su.status = 'Active' 
                GROUP BY su.sale_key, su.sale_name, su.sales_target 
                ORDER BY su.sale_name;
            """
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
            return df
        except Exception as e: 
            print(f"Error getting sales vs target data: {e}") 
            messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลเป้าหมายการขายได้: {e}", parent=self)
            traceback.print_exc() 
            return pd.DataFrame(columns=['sale_name', 'sales_target', 'total_sales'])

    def _create_sales_vs_target_chart(self, parent_frame, data_df):
        if hasattr(self, 'sales_target_chart_canvas') and self.sales_target_chart_canvas:
            self.sales_target_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children():
            widget.destroy()
            
        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูลพนักงานขายที่มีการตั้งเป้าหมาย", font=self.header_font_table).pack(expand=True)
            return

        # ไม่ต้องกำหนด font_name แล้ว เพราะเราตั้งค่า Default ไว้แล้ว
        fig = Figure(figsize=(10, 6), dpi=100, facecolor=self.theme["bg"])
        ax = fig.add_subplot(111)
        ax.set_facecolor(self.theme["bg"])
        
        formatter = FuncFormatter(lambda y, pos: f'{y:,.0f}')
        ax.yaxis.set_major_formatter(formatter)

        x = np.arange(len(data_df['sale_name']))
        width = 0.35
        rects1 = ax.bar(x - width/2, data_df['total_sales'], width, label='ยอดขายจริง', color=self.theme["primary"])
        rects2 = ax.bar(x + width/2, data_df['sales_target'], width, label='ยอดเป้าหมาย', color='#CBD5E1', edgecolor='#94A3B8', linewidth=1)

        ax.set_ylabel('ยอดขาย (บาท)', fontsize=12)
        ax.set_title('กราฟเปรียบเทียบยอดขายจริงกับยอดเป้าหมาย', fontsize=18, weight="bold")
        ax.set_xticks(x)
        ax.set_xticklabels(data_df['sale_name'], rotation=45, ha="right", fontsize=11)
        ax.legend(prop={'size': 12})
        ax.grid(axis='y', linestyle='--', alpha=0.7)

        ax.bar_label(rects1, padding=3, fmt='{:,.0f}', fontsize=9)
        ax.bar_label(rects2, padding=3, fmt='{:,.0f}', fontsize=9)
        
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.sales_target_chart_canvas = canvas

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

    def _get_sales_by_employee_data(self, period): # <-- parameter ถูกต้องแล้ว
        try:
            # --- START: สร้าง logic การกรองใหม่ ---
            today = datetime.now()
            current_year = today.year
            params = []
            
            # 1. สร้าง WHERE clause สำหรับ commission period
            commission_filter_clauses = []
            if period == "เดือนนี้":
                commission_filter_clauses.append("c.commission_month = %s")
                params.append(today.month)
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "ปีนี้":
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            
            # --- START: เพิ่ม Logic ของ Q1-Q4 ตรงนี้ ---
            elif period == "Q1":
                commission_filter_clauses.append("c.commission_month IN (1, 2, 3)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "Q2":
                commission_filter_clauses.append("c.commission_month IN (4, 5, 6)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "Q3":
                commission_filter_clauses.append("c.commission_month IN (7, 8, 9)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            elif period == "Q4":
                commission_filter_clauses.append("c.commission_month IN (10, 11, 12)")
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
            # --- END: สิ้นสุด Logic Q1-Q4 ---
            
            elif period in self.thai_month_map:
                month_num = self.thai_month_map[period]
                commission_filter_clauses.append("c.commission_month = %s")
                params.append(month_num)
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year) # กรองตามปีปัจจุบัน
            else: # Fallback (เหมือน "เดือนนี้")
                commission_filter_clauses.append("c.commission_month = %s")
                params.append(today.month)
                commission_filter_clauses.append("c.commission_year = %s")
                params.append(current_year)
                
            commission_filter_sql = " AND ".join(commission_filter_clauses)
            # --- END ---

            query = f"""
                SELECT 
                    su.sale_name, 
                    su.sale_key, 
                    su.sales_target, 
                    COALESCE(SUM(c.sales_service_amount), 0) as total_sales
                FROM sales_users su
                LEFT JOIN commissions c ON su.sale_key = c.sale_key
                                     AND c.is_active = 1
                                     AND {commission_filter_sql}
                WHERE su.role = 'Sale' AND su.status = 'Active'
                GROUP BY su.sale_name, su.sale_key, su.sales_target
                HAVING COALESCE(SUM(c.sales_service_amount), 0) > 0
                ORDER BY su.sale_name, total_sales DESC;
            """
            
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params)) 
            return df
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลยอดขายตามพนักงานได้: {e}", parent=self)
            traceback.print_exc()
            return pd.DataFrame(columns=['sale_name', 'sale_key', 'sales_target', 'total_sales'])

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
        selected_month = config["month"] # <-- รับค่าเดือน
        selected_year = config["year"]   # <-- รับค่าปี
        self.current_comparison_month = selected_month
        self.current_comparison_year = selected_year
        self.current_comparison_salesperson = selected_salesperson
        self.uploaded_df = config["imported_df"]
        self.manual_entry_df = config["manual_df"]

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
            # --- START: จุดที่แก้ไข Query ---
            base_query = """SELECT c.*, 
                       po.cogs_db, po.po_shipping_stock, po.po_shipping_site, po.po_relocation, 
                       u.sale_name,
                       ss.sale_name as support_user_name 
                FROM commissions c 
                JOIN sales_users u ON c.sale_key = u.sale_key
                LEFT JOIN sales_users ss ON c.support_user_key = ss.sale_key
                LEFT JOIN (
                        /* <<< จุดแก้ไขที่สำคัญที่สุด >>> */
                        SELECT
                            p.so_number,
                            -- 1. ต้นทุนสินค้า (cogs_db): คือผลรวม total_price จาก items ที่อยู่ใน PO ที่ Approved แล้วเท่านั้น
                            SUM(COALESCE(poi.total_price, 0)) as cogs_db,
                            -- 2. ค่าขนส่ง: คือผลรวมจากตาราง PO หลักเหมือนเดิม
                            SUM(p.shipping_to_stock_cost) as po_shipping_stock,
                            SUM(p.shipping_to_site_cost) as po_shipping_site,
                            SUM(p.relocation_cost) as po_relocation
                        FROM purchase_orders p
                        LEFT JOIN purchase_order_items poi ON p.id = poi.purchase_order_id
                        WHERE p.status = 'Approved'
                        GROUP BY p.so_number
                    ) po ON c.so_number = po.so_number
                WHERE c.is_active = 1 
                  AND c.status NOT IN ('HR Verified', 'Paid', 'Deferred by HR', 'Cancelled')
"""
            # --- END: สิ้นสุดการแก้ไข Query ---
            params = []

            ### START: แก้ไขส่วนนี้ ###
            # ตอนนี้ selected_salesperson จะเป็นรหัสของเซลส์เสมอ ไม่ใช่ "ทั้งหมด"
            # เราจึงสามารถลดรูปเงื่อนไขได้
            base_query += " AND c.sale_key = %s"
            params.append(selected_salesperson)
            ### END ###
            
            base_query += " AND c.commission_year = %s AND c.commission_month = %s AND c.sale_key = %s"
            params.extend([selected_year, selected_month, selected_salesperson])

            data_query = base_query + " ORDER BY c.timestamp DESC"
            
            self.db_df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(params))
            
            if loading.winfo_exists(): loading.destroy()
            
            self._compare_data()

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            print(traceback.format_exc())
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูลได้: {e}", parent=self)
            
    def _finalize_comparison(self):
        if self.comparison_df is None or self.comparison_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลการเปรียบเทียบที่จะยืนยัน", parent=self)
            return

        good_statuses = ["ผ่านเกณฑ์"]
        df_to_finalize = self.comparison_df[self.comparison_df['สถานะ'].isin(good_statuses)].copy()

        if df_to_finalize.empty:
            messagebox.showinfo("ไม่พบรายการ", "ไม่พบรายการที่ 'ผ่านเกณฑ์' ที่จะส่งต่อได้ในขณะนี้", parent=self)
            return
        
        self._save_comparison_to_log()

        records_to_update = []
        for index, row in df_to_finalize.iterrows():
            so_number = row['เลขที่ SO']
            
            full_row_data = self.comparison_df.loc[self.comparison_df['เลขที่ SO'] == so_number].iloc[0]

            sales_db_pure = full_row_data.get('ยอดขาย/บริการ (ระบบ)', 0)
            sales_uploaded = full_row_data.get('ยอดขาย (Express)', 0)
            cost_db = full_row_data.get('ต้นทุน (ระบบ)', 0)
            cost_uploaded = full_row_data.get('ต้นทุน (Express)', 0)
            
            sales_db_pure_cleaned = utils.convert_to_float(sales_db_pure)
            sales_uploaded_cleaned = utils.convert_to_float(sales_uploaded)
            cost_db_cleaned = utils.convert_to_float(cost_db)
            cost_uploaded_cleaned = utils.convert_to_float(cost_uploaded)

            # <<< ส่วนที่แก้ไข >>>
            so_record_from_db = self.db_df[self.db_df['so_number'] == so_number]
            sale_source = 'system'
            cost_source = 'system'
            
            if not so_record_from_db.empty:
                record = so_record_from_db.iloc[0]
                sale_source = record.get('hr_sale_source') if pd.notna(record.get('hr_sale_source')) else 'system'
                cost_source = record.get('hr_cost_source') if pd.notna(record.get('hr_cost_source')) else 'system'

            if sale_source == 'express':
                final_sale = sales_uploaded_cleaned
            else:
                final_sale = sales_db_pure_cleaned

            if cost_source == 'express':
                final_cost = cost_uploaded_cleaned
            else:
                final_cost = cost_db_cleaned
            # <<< สิ้นสุดส่วนที่แก้ไข >>>

            final_gp = final_sale - final_cost
            final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0

            so_record = self.db_df[self.db_df['so_number'] == so_number]
            if not so_record.empty:
                record_id = so_record.iloc[0]['id']
                records_to_update.append((
                    int(record_id),
                    final_sale,
                    final_cost,
                    final_gp,
                    final_margin
                ))

        if not records_to_update:
            messagebox.showinfo("ไม่พบรายการ", "ไม่สามารถหา ID ของรายการที่ต้องการอัปเดตได้", parent=self)
            return

        msg = (f"คุณต้องการยืนยันข้อมูลสำหรับ {len(records_to_update)} รายการที่ผ่านเกณฑ์ใช่หรือไม่?\n\n"
               f"การกระทำนี้จะอัปเดตสถานะและบันทึกยอดขาย/ต้นทุนสุดท้ายเข้าระบบ เพื่อให้พร้อมสำหรับประมวลผลค่าคอมมิชชั่นต่อไป")

        if not messagebox.askyesno("ยืนยันการส่งต่อข้อมูล", msg, parent=self):
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                update_query = """
                    UPDATE commissions 
                    SET 
                        status = 'HR Verified', 
                        final_sales_amount = data.final_sale,
                        final_cost_amount = data.final_cost,
                        final_gp = data.final_gp,
                        final_margin = data.final_margin
                    FROM (VALUES %s) AS data(record_id, final_sale, final_cost, final_gp, final_margin)
                    WHERE commissions.id = data.record_id;
                """
                psycopg2.extras.execute_values(
                    cursor,
                    update_query,
                    records_to_update,
                    template="(%s::int, %s::float, %s::float, %s::float, %s::float)",
                    page_size=100
                )
                updated_rows = cursor.rowcount
            conn.commit()
            
            messagebox.showinfo("สำเร็จ", f"อัปเดตข้อมูล {updated_rows} รายการเป็น 'HR Verified' เรียบร้อยแล้ว", parent=self)

            self._refresh_comparison_view()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการยืนยันข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    # hr_screen.py (เพิ่มฟังก์ชันนี้เข้าไปในคลาส HRScreen)

    def _verify_passed_sos(self):
        """
        (เวอร์ชันปรับปรุง) ยืนยัน SO ที่ 'ผ่านเกณฑ์' ทั้งหมด โดยการดึงข้อมูลล่าสุดจาก DB
        มาคำนวณใหม่ทั้งหมดก่อนทำการอัปเดต เพื่อความถูกต้อง 100%
        """
        if self.comparison_df is None or self.comparison_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลการเปรียบเทียบที่จะยืนยัน", parent=self)
            return

        # 1. รวบรวม SO ที่ 'ผ่านเกณฑ์'
        df_to_verify = self.comparison_df[self.comparison_df['สถานะ'] == 'ผ่านเกณฑ์']
        so_numbers_to_verify = tuple(df_to_verify['เลขที่ SO'].tolist())

        if not so_numbers_to_verify:
            messagebox.showinfo("ไม่พบรายการ", "ไม่พบรายการที่ 'ผ่านเกณฑ์' ที่จะยืนยันได้ในขณะนี้", parent=self)
            return

        # 2. ถามเพื่อยืนยันการทำงาน
        msg = (f"คุณต้องการยืนยันข้อมูลสำหรับ {len(so_numbers_to_verify)} รายการที่ผ่านเกณฑ์ใช่หรือไม่?\n\n"
               f"โปรแกรมจะดึงข้อมูลล่าสุดมาคำนวณใหม่ทั้งหมดก่อนบันทึก")
        if not messagebox.askyesno("ยืนยันข้อมูล", msg, parent=self):
            return

        records_to_update = []
        conn = None
        try:
            # 3. Query ข้อมูลล่าสุดของ SO ที่เลือกจาก DB
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                placeholders = ', '.join(['%s'] * len(so_numbers_to_verify))
                
                # Query ที่สมบูรณ์เพื่อดึงทั้งข้อมูล SO และต้นทุนล่าสุดจาก PO Items
                query = f"""
                    SELECT
                        c.*, 
                        -- 3. นำยอดรวมสินค้า มาลบกับ ยอดรวมส่วนลด
                        (COALESCE(po_items.total_item_cost, 0) - COALESCE(po_discounts.total_bill_discount, 0)) as cogs_db
                    FROM commissions c
                    LEFT JOIN (
                        -- 1. รวมยอดราคาสินค้าทั้งหมด (จะได้ 5,737)
                        SELECT 
                            p.so_number,
                            SUM(COALESCE(poi.total_price, 0)) as total_item_cost
                        FROM purchase_orders p
                        LEFT JOIN purchase_order_items poi ON p.id = poi.purchase_order_id
                        WHERE p.status = 'Approved' AND p.so_number IN ({placeholders})
                        GROUP BY p.so_number
                    ) po_items ON c.so_number = po_items.so_number
                    LEFT JOIN (
                        -- 2. รวมยอดส่วนลดท้ายบิลทั้งหมด (จะได้ 37)
                        SELECT
                            so_number,
                            SUM(COALESCE(bill_discount, 0)) as total_bill_discount
                        FROM purchase_orders
                        WHERE status = 'Approved' AND so_number IN ({placeholders})
                        GROUP BY so_number
                    ) po_discounts ON c.so_number = po_discounts.so_number
                    WHERE c.so_number IN ({placeholders}) AND c.is_active = 1
                """
                cursor.execute(query, so_numbers_to_verify)
                latest_so_data = cursor.fetchall()

                # 4. วนลูปคำนวณค่าสุดท้ายใหม่ทั้งหมดด้วย Logic ล่าสุด
                for row_data in latest_so_data:
                    # คำนวณ Final Sales (รายรับรวม)
                    final_sale = (float(row_data.get('sales_service_amount', 0) or 0) +
                                  float(row_data.get('cutting_drilling_fee', 0) or 0) +
                                  float(row_data.get('other_service_fee', 0) or 0))
                    
                    # Final Cost คือ cogs_db ที่เราดึงมาใหม่ล่าสุด
                    final_cost = float(row_data.get('cogs_db', 0) or 0)

                    final_gp = final_sale - final_cost
                    final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
                    
                    records_to_update.append((
                        int(row_data['id']),
                        final_sale,
                        final_cost,
                        final_gp,
                        final_margin
                    ))

            if not records_to_update:
                messagebox.showerror("ผิดพลาด", "ไม่สามารถเตรียมข้อมูลสำหรับอัปเดตได้ (อาจไม่พบข้อมูลล่าสุดใน DB)", parent=self)
                return

            # 5. อัปเดตฐานข้อมูล (Bulk Update)
            with conn.cursor() as cursor:
                update_query = """
                    UPDATE commissions SET 
                        status = 'HR Verified', 
                        final_sales_amount = data.final_sale,
                        final_cost_amount = data.final_cost,
                        final_gp = data.final_gp,
                        final_margin = data.final_margin,
                        payout_id = NULL
                    FROM (VALUES %s) AS data(record_id, final_sale, final_cost, final_gp, final_margin)
                    WHERE commissions.id = data.record_id;
                """
                psycopg2.extras.execute_values(
                    cursor, update_query, records_to_update,
                    template="(%s::int, %s::float, %s::float, %s::float, %s::float)",
                    page_size=100
                )
                updated_rows = cursor.rowcount
            conn.commit()
            
            messagebox.showinfo("สำเร็จ", f"ยืนยันข้อมูล {updated_rows} รายการเรียบร้อยแล้ว", parent=self)
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
            comparison_sources = []
            if self.uploaded_df is not None and not self.uploaded_df.empty:
                comparison_sources.append(self.uploaded_df)
            if self.manual_entry_df is not None and not self.manual_entry_df.empty:
                comparison_sources.append(self.manual_entry_df)
            
            if not comparison_sources:
                messagebox.showwarning("ไม่มีข้อมูลเปรียบเทียบ", "กรุณา Import ไฟล์ หรือ คีย์ข้อมูลด้วยมือ", parent=self)
                self._create_styled_dataframe_table(self.results_frame, self.db_df)
                return
                
            uploaded_compare_df = pd.concat(comparison_sources, ignore_index=True).drop_duplicates(subset=['so_number'], keep='last')

            processed_so_query = "SELECT so_number FROM commissions WHERE status IN ('HR Verified', 'Paid', 'Deferred by HR', 'Cancelled')"
            processed_so_df = pd.read_sql_query(processed_so_query, self.pg_engine)
            
            if not processed_so_df.empty:
                processed_so_list = processed_so_df['so_number'].tolist()
                initial_count = len(uploaded_compare_df)
                uploaded_compare_df['so_number'] = uploaded_compare_df['so_number'].astype(str).str.strip()
                uploaded_compare_df = uploaded_compare_df[~uploaded_compare_df['so_number'].isin(processed_so_list)]
                if initial_count > len(uploaded_compare_df):
                    print(f"Filtered out {initial_count - len(uploaded_compare_df)} already processed SO(s) from the uploaded data.")

            db_compare_df = self.db_df.copy()
            db_compare_df['so_number'] = db_compare_df['so_number'].astype(str).str.strip()
            
            sales_revenue_keys = ['sales_service_amount', 'cutting_drilling_fee', 'other_service_fee']
            db_compare_df['sales_for_comparison'] = 0
            for key in sales_revenue_keys:
                if key in db_compare_df.columns:
                    db_compare_df['sales_for_comparison'] += pd.to_numeric(db_compare_df[key], errors='coerce').fillna(0)

            brokerage = pd.to_numeric(db_compare_df['brokerage_fee'], errors='coerce').fillna(0)
            transfer = pd.to_numeric(db_compare_df['transfer_fee'], errors='coerce').fillna(0)
            giveaways = pd.to_numeric(db_compare_df['giveaways'], errors='coerce').fillna(0)
            cogs = pd.to_numeric(db_compare_df['cogs_db'], errors='coerce').fillna(0)
            db_compare_df['cost_db'] = cogs

            db_compare_df['gp_db'] = db_compare_df['sales_service_amount'] - db_compare_df['cost_db']
            db_compare_df['margin_db'] = (db_compare_df['gp_db'] / db_compare_df['sales_service_amount'].replace(0, np.nan)) * 100

            uploaded_compare_df['so_number'] = uploaded_compare_df['so_number'].astype(str).str.strip()
            uploaded_compare_df['sales_uploaded'] = pd.to_numeric(uploaded_compare_df.get('sales_uploaded'), errors='coerce').fillna(0)
            uploaded_compare_df['cost_uploaded'] = pd.to_numeric(uploaded_compare_df.get('cost_uploaded'), errors='coerce').fillna(0)
            uploaded_compare_df['gp_uploaded'] = uploaded_compare_df['sales_uploaded'] - uploaded_compare_df['cost_uploaded']
            uploaded_compare_df['margin_uploaded'] = (uploaded_compare_df['gp_uploaded'] / uploaded_compare_df['sales_uploaded'].replace(0, np.nan)) * 100

            merged_df = pd.merge(db_compare_df, uploaded_compare_df, on='so_number', how='outer', suffixes=('_db', '_uploaded'), indicator=True)
            
            merged_df['sales_uploaded'] = merged_df['sales_uploaded'].fillna(0) - merged_df['shipping_cost'].fillna(0)

            merged_df['แหล่งยอดขาย'] = merged_df['hr_sale_source'].apply(
                lambda x: 'ระบบ' if x == 'system' else ('Express' if x == 'express' else 'ยังไม่เลือก')
            )
            merged_df['แหล่งต้นทุน'] = merged_df['hr_cost_source'].apply(
                lambda x: 'ระบบ' if x == 'system' else ('Express' if x == 'express' else 'ยังไม่เลือก')
            )

            def determine_status_and_color(row):
                difference_amount = row.get('difference_amount', 0.0) or 0.0
                
                # --- START: แก้ไข Logic การตรวจสอบสถานะทั้งหมด ---
                
                # 1. ตรวจสอบยอดโอนขาดก่อน (เป็นปัญหาสำคัญสุด)
                if difference_amount < -0.01:
                    return f"⚠️ ยอดโอนขาด ({abs(difference_amount):,.2f})"

                # 2. ตรวจสอบสถานะอื่นๆ ที่ควรแสดงผลทันที
                if row['status'] == 'HR Verified':
                    final_margin = row['final_margin']
                    if pd.isna(final_margin): return 'ยืนยันแล้ว (รอผล)'
                    if final_margin < 0: return 'ขาดทุน'
                    elif final_margin < 10: return 'กำไรน้อย'
                    else: return 'กำไรดี'

                if row['_merge'] == 'right_only': return 'มีใน Express, ไม่มีในระบบ'
                if row['_merge'] == 'left_only': return 'มีในระบบ, ไม่มีใน Express'

                # 3. ตรวจสอบปัญหาข้อมูลอื่นๆ
                final_system_sale = row['sales_for_comparison']
                final_system_cost = row['cost_db']
                cost_uploaded = row['cost_uploaded']

                if pd.notna(final_system_sale) and pd.notna(final_system_cost) and final_system_cost > final_system_sale:
                    return "‼️ ขายขาดทุน (ตรวจสอบด่วน)"
                if pd.notna(final_system_cost) and final_system_cost > 0 and pd.notna(cost_uploaded) and cost_uploaded < (final_system_cost * 0.5):
                    return "‼️ ต้นทุน Express ผิดปกติ (<50%)"
                
                # 4. ตรวจสอบข้อมูล Sales/Cost ระหว่างระบบกับ Express
                sale_system_rounded = round(float(final_system_sale) if pd.notna(final_system_sale) else 0.0, 2)
                sale_express_rounded = round(float(row['sales_uploaded']) if pd.notna(row['sales_uploaded']) else 0.0, 2)
                cost_system_rounded = round(float(final_system_cost) if pd.notna(final_system_cost) else 0.0, 2)
                cost_express_rounded = round(float(row['cost_uploaded']) if pd.notna(row['cost_uploaded']) else 0.0, 2)

                sale_ok = sale_system_rounded >= sale_express_rounded
                cost_ok = cost_system_rounded >= cost_express_rounded
                
                # 5. ถ้าทุกอย่างถูกต้อง (รวมถึงยอดโอนไม่ขาด) ให้ถือว่า "ผ่านเกณฑ์"
                if sale_ok and cost_ok:
                    # ถ้ามียอดโอนเกิน ให้แสดงข้อความบอก แต่ยังคงสถานะให้ผ่านได้
                    if difference_amount > 0.01:
                        return f"ผ่านเกณฑ์ (โอนเกิน {difference_amount:,.2f})"
                    else:
                        return "ผ่านเกณฑ์"
                
                # 6. ถ้าข้อมูลไม่ตรงกัน ให้แสดงตามปกติ
                elif not sale_ok: return "ยอดขายต่ำกว่า Express"
                elif not cost_ok: return "ต้นทุนต่ำกว่า Express"
                else: return "ข้อมูลไม่ตรงกัน"

            merged_df['สถานะ'] = merged_df.apply(determine_status_and_color, axis=1)
            
            merged_df['ผลต่างยอดขาย'] = merged_df['sales_for_comparison'].fillna(0) - merged_df['sales_uploaded'].fillna(0)
            merged_df['ผลต่างต้นทุน'] = merged_df['cost_db'].fillna(0) - merged_df['cost_uploaded'].fillna(0)
            
            display_order_map = {
                'so_number': 'เลขที่ SO',
                'sales_service_amount': 'ยอดขาย/บริการ (ระบบ)',
                'shipping_cost': 'ค่าขนส่ง (ระบบ)',
                'relocation_cost': 'ค่าย้าย (ระบบ)',
                'sales_for_comparison': 'ยอดขายรวม (ระบบ)',
                'sales_uploaded': 'ยอดขาย (Express)',
                'cost_db': 'ต้นทุน (ระบบ)',
                'cost_uploaded': 'ต้นทุน (Express)',
                'ผลต่างยอดขาย': 'ผลต่างยอดขาย',
                'ผลต่างต้นทุน': 'ผลต่างต้นทุน',
                'แหล่งยอดขาย': 'แหล่งยอดขาย',
                'แหล่งต้นทุน': 'แหล่งต้นทุน',
                'สถานะ': 'สถานะ'
            }

            for key in display_order_map.keys():
                if key not in merged_df.columns:
                    merged_df[key] = np.nan
            
            self.comparison_df = merged_df[list(display_order_map.keys())].copy()
            self.comparison_df.rename(columns=display_order_map, inplace=True)

            # --- START: แก้ไขส่วนสรุปยอดตรงนี้ ---
            # 1. นับจำนวน SO ก่อนที่จะเพิ่มแถวสรุป
            so_count = len(self.comparison_df)

            # 2. คำนวณผลรวม (เหมือนเดิม)
            numeric_cols = ['ยอดขายรวม (ระบบ)', 'ยอดขาย (Express)', 'ต้นทุน (ระบบ)', 'ต้นทุน (Express)', 'ผลต่างยอดขาย', 'ผลต่างต้นทุน']
            summary_data = self.comparison_df[numeric_cols].sum().to_dict()
            
            summary_row = pd.Series(summary_data)
            summary_row['เลขที่ SO'] = 'ยอดรวม (Total)'
            
            # 3. นำจำนวน SO ที่นับได้มาใส่ในคอลัมน์ 'สถานะ' ของแถวสรุป
            summary_row['สถานะ'] = f"รวม {so_count} รายการ"
            # --- END ---
            
            self.comparison_df = pd.concat([self.comparison_df, summary_row.to_frame().T], ignore_index=True)
                
            status_colors = {
                "ผ่านเกณฑ์": "#D1FAE5",                 # สีเขียวอ่อน
                "ยอดขายต่ำกว่า Express": "#FEF2F2",      # สีแดงอ่อน
                "ต้นทุนต่ำกว่า Express": "#FEFCE8",       # สีเหลืองอ่อน
                "ข้อมูลไม่ตรงกัน": "#FFF7ED",          # สีส้มอ่อน
                
                # ใช้ Key แบบง่ายๆ และกำหนดสีที่ต้องการ
                "ยอดโอนเกิน": "#D1FAE5",              # สีเขียวอ่อน
                "ยอดโอนขาด": "#FEF3C7",              # สีเหลืองเข้ม

                # สถานะอื่นๆ
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

        loading = self._show_loading(self.results_frame)
        self.results_frame_label.configure(text=f"กำลังรีเฟรชข้อมูลสำหรับ: {self.current_comparison_salesperson}...")

        try:
            # --- START: โค้ดส่วนที่เพิ่มเข้ามา ---
            # ดึงค่าเดือนและปีที่เคยเลือกไว้จาก Dialog (ถ้ามี)
            # เราจะใช้ค่าที่เก็บไว้ในตัวแปรจากฟังก์ชัน _start_new_comparison
            selected_month = getattr(self, 'current_comparison_month', None)
            selected_year = getattr(self, 'current_comparison_year', None)
            # --- END ---
            
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
                  AND c.status NOT IN ('HR Verified', 'Paid', 'Deferred by HR', 'Cancelled')
            """
            params = []

            if self.current_comparison_salesperson != "ทั้งหมด":
                base_query += " AND c.sale_key = %s"
                params.append(self.current_comparison_salesperson)
            else:
                base_query += " AND c.sale_key IN (SELECT sale_key FROM sales_users WHERE status = 'Active' AND role = 'Sale')"

            # --- START: เพิ่มเงื่อนไขการกรองเดือนและปี ---
            if selected_month and selected_year:
                base_query += " AND c.commission_month = %s AND c.commission_year = %s"
                params.extend([selected_month, selected_year])
            # --- END ---

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
        
        self.process_result_frame = CTkFrame(parent_tab)
        self.process_result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.process_result_frame.grid_rowconfigure(1, weight=1)
        self.process_result_frame.grid_columnconfigure(0, weight=1)

        self.after(100, self._on_sale_selected_for_process)
        

    def _on_sale_selected_for_process(self, sale_key=None):
        """เมื่อเลือกพนักงานขาย จะค้นหางวดที่มีข้อมูล 'ที่ยังไม่เคยจ่าย' มาให้คำนวณ"""
        if sale_key is None:
            sale_key = self.selected_sale_for_process.get()
        if not sale_key: return

        for widget in self.process_result_frame.winfo_children(): widget.destroy()
        
        try:
            # --- จุดที่แก้ไข: เพิ่มเงื่อนไข AND payout_id IS NULL ---
            query = """
                SELECT DISTINCT commission_year, commission_month 
                FROM commissions 
                WHERE sale_key = %s AND status = 'HR Verified' AND is_active = 1
                AND payout_id IS NULL 
                ORDER BY commission_year DESC, commission_month DESC
            """
            df_periods = pd.read_sql_query(query, self.pg_engine, params=(sale_key,))

            if df_periods.empty:
                self.process_period_menu.configure(values=["-ไม่มีข้อมูล-"], state="disabled")
                self.process_period_var.set("-ไม่มีข้อมูล-")
                CTkLabel(self.process_result_frame, text=f"ไม่พบข้อมูลที่ 'Verified' และยังไม่ได้จ่ายเงินสำหรับ: {sale_key}").pack(pady=20)
                return

            period_options = [f"{self.thai_months[month-1]} {year+543}" for year, month in zip(df_periods['commission_year'], df_periods['commission_month'])]
            self.process_period_menu.configure(values=period_options, state="normal")
            self.process_period_var.set(period_options[0])
            self._calculate_commission_for_period()

        except Exception as e:
            messagebox.showerror("DB Error", f"เกิดข้อผิดพลาดในการค้นหางวดข้อมูล: {e}", parent=self)
    
    

    def _calculate_commission_for_period(self, selected_period=None):
        if selected_period is None:
            selected_period = self.process_period_var.get()
        
        sale_key = self.selected_sale_for_process.get()
        if not selected_period or not sale_key or "-" in selected_period:
            return

        month_name, year_be_str = selected_period.split()
        month_num = self.thai_month_map[month_name]
        year_ad = int(year_be_str) - 543

        # <<< START: เพิ่ม 3 บรรทัดนี้เพื่อบันทึกค่างวดไว้ใช้ตอน Save >>>
        self.current_period_text = selected_period
        self.selected_month = month_num
        self.selected_year = year_ad
        # <<< END >>>

        plan_info = self.sales_user_info.get(sale_key, {})
        plan = plan_info.get('plan', 'Plan A')
        sales_target = float(plan_info.get('target', 0.0))

        for widget in self.process_result_frame.winfo_children(): widget.destroy()
        loading = self._show_loading(self.process_result_frame)

        try:
            query_comm = """
                SELECT c.*, COALESCE(po_costs.total_po_shipping_cost, 0) as total_po_shipping_cost
                FROM commissions c
                LEFT JOIN (
                    SELECT so_number, SUM(COALESCE(shipping_to_stock_cost, 0) + COALESCE(shipping_to_site_cost, 0)) as total_po_shipping_cost
                    FROM purchase_orders WHERE status = 'Approved' GROUP BY so_number
                ) po_costs ON c.so_number = po_costs.so_number
                WHERE c.sale_key = %s 
                    AND c.status = 'HR Verified' 
                    AND c.payout_id IS NULL
                    AND c.commission_month = %s 
                    AND c.commission_year = %s
                    AND c.is_active = 1
            """
            params = (sale_key, month_num, year_ad)
            self.current_comm_df = pd.read_sql_query(query_comm, self.pg_engine, params=params)

            # <<< START: บันทึก ID ของ SO ที่จะถูกประมวลผล >>>
            self.current_so_ids = self.current_comm_df['id'].tolist()
            # <<< END >>>
            
            self.current_total_sales = self.current_comm_df['final_sales_amount'].sum()
            self.current_total_cost = self.current_comm_df['final_cost_amount'].sum()

            if self.current_comm_df.empty:
                loading.destroy()
                CTkLabel(self.process_result_frame, text="ไม่พบข้อมูลในงวดที่เลือก").pack(pady=20)
                return
            
            # --- ส่วนที่ 1: คำนวณค่าหักจาก "ส่วนต่างค่ารถ" ---
            total_so_shipping = self.current_comm_df['shipping_cost'].sum()
            total_po_shipping = self.current_comm_df['total_po_shipping_cost'].sum()
            shipping_deduction = 0.0
            
            print("\n" + "="*25)
            print("### DEBUG: Auto-Deduction Calculation ###")
            print("-" * 15)
            print("Part 1: Shipping Cost Difference")
            print(f"  - Total PO Shipping: {total_po_shipping:,.2f}")
            print(f"  - Total SO Shipping: {total_so_shipping:,.2f}")

            if total_po_shipping > total_so_shipping:
                shipping_diff = total_po_shipping - total_so_shipping
                shipping_deduction = (shipping_diff / 0.2) * 0.0175
                
                print(f"  - Difference (PO > SO): {shipping_diff:,.2f}")
                print(f"  - Formula: ({shipping_diff:,.2f} / 0.2) * 0.0175")
                print(f"  - Shipping Deduction = {shipping_deduction:,.2f}")
            else:
                print("  - Condition not met (PO Shipping <= SO Shipping)")
                print(f"  - Shipping Deduction = 0.00")

            # --- ส่วนที่ 2: คำนวณค่าหักจาก "ส่วนต่างนายหน้า" ---
            print("-" * 15)
            print("Part 2: Brokerage Difference")
            
            print("  --- Breakdown of Difference Amount ---")
            df_with_diff = self.current_comm_df[self.current_comm_df['difference_amount'] != 0]
            if not df_with_diff.empty:
                for index, row in df_with_diff.iterrows():
                    print(f"    -> SO: {row['so_number']}, Difference: {row['difference_amount']:,.2f}")
            else:
                print("    -> No SOs with a non-zero difference amount.")
            print("  ------------------------------------")

            total_brokerage = self.current_comm_df['brokerage_fee'].sum()
            total_difference = self.current_comm_df['difference_amount'].sum()
            difference_deduction = 0.0
            diff_base = total_brokerage - total_difference
            
            print(f"  - Total Brokerage Fee: {total_brokerage:,.2f}")
            print(f"  - Total Difference Amount: {total_difference:,.2f}")
            print(f"  - Base Calculation (Brokerage - Difference): {diff_base:,.2f}")
            
            if diff_base < 0:
                positive_diff = abs(diff_base)
                difference_deduction = (positive_diff / 0.2) * 0.0175
                
                print(f"  - Condition met (Base < 0)")
                print(f"  - Formula: ({positive_diff:,.2f} / 0.2) * 0.0175")
                print(f"  - Difference Deduction = {difference_deduction:,.2f}")
            else:
                print("  - Condition not met (Base >= 0)")
                print(f"  - Difference Deduction = 0.00")

            # --- 3. รวมยอดหักอัตโนมัติทั้งหมด ---
            final_auto_deduction = shipping_deduction + difference_deduction

            print("-" * 15)
            print("Part 3: Final Summary")
            print(f"  - Total Auto Deduction = (Shipping) {shipping_deduction:,.2f} + (Difference) {difference_deduction:,.2f}")
            print(f"  - FINAL AUTO DEDUCTION: {final_auto_deduction:,.2f}")
            print("="*25 + "\n")

            df_for_calc = self.current_comm_df.copy()
            df_for_calc['total_revenue'] = df_for_calc['final_sales_amount']
            
            # <<< START: แก้ไข Logic การดึงค่าดำเนินการ >>>
            default_operating_fee = 0.0
            try:
                fee_str = self.operating_fee_entry.get()
                default_operating_fee = utils.convert_to_float(fee_str)
                print(f"-> [DEBUG] อ่านค่า Operating Fee จาก Textbox: {default_operating_fee} (จาก string: '{fee_str}')")
            except AttributeError:
                default_fees = {'Plan A': 25000.00, 'Plan B': 100000.00, 'Plan C': 100000.00, 'Plan D': 750000.00}
                default_operating_fee = default_fees.get(plan, 0.0)
                print(f"-> [DEBUG] Textbox ยังไม่ถูกสร้าง (ครั้งแรก), ใช้ค่า Hardcode: {default_operating_fee}")
            except Exception as e:
                print(f"-> [DEBUG] Error อ่านค่า Operating Fee, ใช้ค่า 0.0. (Error: {e})")
                default_operating_fee = 0.0
            # <<< END: สิ้นสุดการแก้ไข >>>
            
            # --- คำนวณค่าคอม ---
            if plan == 'Plan A':
                print(f"-> [DEBUG] กำลังส่งค่า Operating Fee: {default_operating_fee} ไปคำนวณ (Plan A)")
                print("="*30 + "\n")
                self.initial_commission_result = business_logic.calculate_monthly_commission(
                    plan_name=plan,
                    comm_df=df_for_calc, # Plan A ใช้ df_for_calc
                    sales_target=sales_target,
                    operating_fee=default_operating_fee,
                    incentives=None, 
                    additional_deductions=None
                )
            
            elif plan in ["Plan B", "Plan C", "Plan D"]:
                print(f"-> [DEBUG] กำลังส่งค่า Operating Fee: {default_operating_fee} ไปคำนวณ (Plan {plan})")
                print("="*30 + "\n")
                self.initial_commission_result = business_logic.calculate_monthly_commission(
                    plan_name=plan,
                    comm_df=df_for_calc, # Plan B,C,D ก็ควรใช้ df_for_calc
                    sales_target=sales_target,
                    operating_fee=default_operating_fee,
                    incentives=None,
                    additional_deductions=None
                )
            
            else:
                self.initial_commission_result = {'type': 'error', 'message': f'ไม่รู้จัก Plan: {plan}'}

            self.latest_commission_result = self.initial_commission_result

            result_type = self.initial_commission_result.get('type')
            loading.destroy()

            if result_type in ['no_commission', 'error']:
                message = self.initial_commission_result.get('message', 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ')
                CTkLabel(self.process_result_frame, text=message, font=self.label_font, text_color="orange", wraplength=600).pack(pady=30, padx=20)
            else:
                if result_type == 'summary_plan_a':
                    self.commission_details_df = self.initial_commission_result.get('details')
                else:
                    self.commission_details_df = None
                
                self._create_hr_input_interface(
                    auto_deduction_value=final_auto_deduction,
                    default_operating_fee_to_display=default_operating_fee 
                )

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            traceback.print_exc()
            messagebox.showerror("Calculation Error", f"เกิดข้อผิดพลาดในการคำนวณ: {e}", parent=self)

    def _create_hr_input_interface(self, auto_deduction_value=0.0, default_operating_fee_to_display=None):
        """
        สร้างหน้าจอสำหรับกรอก Incentive/Deduction และแสดงผลสรุปค่าคอม
        (ฉบับแก้ไข: ดึงค่าคอมที่ถูกต้องมาแสดง)
        """
        for widget in self.process_result_frame.winfo_children():
            widget.destroy()

        self.process_result_frame.grid_rowconfigure(1, weight=1)
        self.process_result_frame.grid_columnconfigure(0, weight=1)

        if not hasattr(self, 'initial_commission_result'):
             self.initial_commission_result = {}
             
        # <<< START: แก้ไข Key ที่ดึงข้อมูลตรงนี้ >>>
        # เปลี่ยนจาก 'final_commission' เป็น 'final_commission_pre_deductions'
        calculated_commission = (
        self.initial_commission_result.get('final_commission_pre_deductions') or 
        self.initial_commission_result.get('final_commission', 0.0)
        )   
        # <<< END >>>

        input_frame = CTkFrame(self.process_result_frame)
        input_frame.grid(row=0, column=0, pady=(10, 0), padx=10, sticky="ew")

        plan_name = self.sales_user_info.get(self.selected_sale_for_process.get(), {}).get('plan', 'N/A')
        self.plan_display_label = CTkLabel(input_frame, text=f"แผนค่าคอมมิชชั่น: {plan_name}", font=self.header_font_table, text_color=self.theme["primary"])
        self.plan_display_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

        CTkLabel(input_frame, text="ยอดคอมมิชชั่นที่คำนวณได้:", font=self.label_font).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.calculated_commission_label = CTkLabel(input_frame, text=f"{calculated_commission:,.2f} บาท", font=self.header_font_table)
        self.calculated_commission_label.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        stats_frame = CTkFrame(input_frame, fg_color="transparent")
        stats_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(5,0))
        stats_frame.grid_columnconfigure((1, 3), weight=1)

        # ดึงค่าจาก self.current_total_sales และ self.current_total_cost แทน
        # (ค่าเหล่านี้ถูกคำนวณไว้แล้วใน _calculate_commission_for_period)
        total_sales_display = getattr(self, 'current_total_sales', 0.0)
        total_cost_display = getattr(self, 'current_total_cost', 0.0)

        CTkLabel(stats_frame, text="ยอดขายรวม (ที่ใช้คำนวณ):", font=self.label_font, text_color="#2563EB").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        CTkLabel(stats_frame, text=f"{total_sales_display:,.2f} บาท", font=self.entry_font).grid(row=0, column=1, padx=10, pady=5, sticky="w")

        CTkLabel(stats_frame, text="ต้นทุนรวม (ที่ใช้คำนวณ):", font=self.label_font, text_color="#D97706").grid(row=0, column=2, padx=(20, 10), pady=5, sticky="w")
        CTkLabel(stats_frame, text=f"{total_cost_display:,.2f} บาท", font=self.entry_font).grid(row=0, column=3, padx=10, pady=5, sticky="w")

        CTkLabel(input_frame, text="(-) ค่าดำเนินการ:", font=self.label_font).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.operating_fee_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.operating_fee_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        
        if default_operating_fee_to_display is not None:
            self.operating_fee_entry.insert(0, f"{default_operating_fee_to_display:,.2f}")
        else:
            default_fees = {'Plan A': 25000, 'Plan B': 100000, 'Plan C': 100000, 'Plan D': 750000}
            default_fee = default_fees.get(plan_name, 0.0)
            self.operating_fee_entry.insert(0, f"{default_fee:,.2f}")
        
        CTkLabel(input_frame, text="(+) Incentive:", font=self.label_font).grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.incentive_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.incentive_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

        CTkLabel(input_frame, text="(-) หัก ค่าใช้จ่ายอื่นๆ:", font=self.label_font).grid(row=5, column=0, padx=10, pady=10, sticky="w")
        self.deduction_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.deduction_entry.grid(row=5, column=1, padx=10, pady=10, sticky="ew")
        if auto_deduction_value > 0:
            self.deduction_entry.insert(0, f"{auto_deduction_value:,.2f}")

        CTkLabel(input_frame, text="หมายเหตุ/Incentive อื่นๆ:", font=self.label_font).grid(row=6, column=0, padx=10, pady=10, sticky="w")
        self.payout_notes_entry = CTkTextbox(input_frame, height=80)
        self.payout_notes_entry.grid(row=6, column=1, padx=10, pady=10, sticky="ew")
        
        calc_button_frame = CTkFrame(input_frame, fg_color="transparent")
        calc_button_frame.grid(row=7, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
        calc_button_frame.grid_columnconfigure((0, 1), weight=1)
        
        CTkButton(calc_button_frame, text="คำนวณขั้นสุดท้ายและแสดงสรุป", command=self._perform_final_calculation, fg_color=self.theme["primary"]).grid(row=0, column=0, padx=(0, 5), pady=10, sticky="ew")
        
        self.detail_button = CTkButton(calc_button_frame, text="แสดงการคิดแบบละเอียด", command=self._show_calculation_details)
        self.detail_button.grid(row=0, column=1, padx=(5, 0), pady=10, sticky="ew")

        if not self.initial_commission_result.get('debug_df', pd.DataFrame()).empty:
            self.detail_button.configure(state="normal")
        else:
            self.detail_button.configure(state="disabled")

        self.final_summary_frame = CTkScrollableFrame(self.process_result_frame, fg_color="transparent")
        self.final_summary_frame.grid(row=1, column=0, pady=10, padx=10, sticky="nsew")
        
        bottom_action_frame = CTkFrame(self.process_result_frame, fg_color="transparent")
        bottom_action_frame.grid(row=2, column=0, pady=(0, 10), padx=10, sticky="ew")
        
        self.confirm_payout_button = CTkButton(bottom_action_frame, text="✅ ยืนยันการจ่ายเงินและบันทึก",
                                command=self._confirm_payout_and_save,
                                fg_color="#16A34A", hover_color="#15803D",
                                font=CTkFont(size=16, weight="bold"))

    def _perform_final_calculation(self):
        try:
            operating_fee_val = float(self.operating_fee_entry.get().replace(",", "") or 0.0)
            incentive_val = float(self.incentive_entry.get().replace(",", "") or 0.0)
            deduction_val = float(self.deduction_entry.get().replace(",", "") or 0.0)
        except ValueError:
            messagebox.showerror("ข้อมูลผิดพลาด", "กรุณากรอก Incentive และ Deduction เป็นตัวเลข", parent=self)
            return

        incentives_dict = {"Incentive พิเศษ": incentive_val} if incentive_val > 0 else None
        deductions_dict = {"ค่าใช้จ่าย/ดำเนินการ": deduction_val} if deduction_val > 0 else None
        
        sale_key = self.selected_sale_for_process.get()
        plan = self.sales_user_info.get(sale_key, {}).get('plan', 'Plan A')

        df_for_final_calc = self.current_comm_df.copy()
        df_for_final_calc['total_revenue'] = df_for_final_calc['final_sales_amount']

        final_result = business_logic.calculate_monthly_commission(
            plan_name=plan,
            comm_df=df_for_final_calc,
            operating_fee=operating_fee_val, # <-- ส่งค่าจากช่องกรอก
            incentives=incentives_dict,
            additional_deductions=deductions_dict
        )
        self.latest_commission_result = final_result

        self.final_summary_data = None 
        self.confirm_payout_button.pack_forget()

        result_type = final_result.get('type')
        summary_df = None
        details_df = None

        if result_type == 'summary_plan_a':
            summary_df = final_result.get('summary')
            details_df = final_result.get('details') 
        elif result_type == 'summary_other':
            summary_df = final_result.get('data')

        if summary_df is not None:
            # แสดงตารางสรุปบนหน้าจอ (GUI)
            self._create_commission_summary_table(summary_df, container=self.final_summary_frame)
            
            # --- START: ส่วนที่แก้ไข ---
            # ถ้ามีตารางรายละเอียด (เป็น Plan A) ให้ print ออกทาง Terminal
            if details_df is not None:
                print("\n" + "="*40)
                print("  DEBUG: Commission Calculation Details (Plan A)")
                print("="*40)
                # ใช้ .to_string() เพื่อให้แสดงผลสวยงามใน Terminal
                print(details_df.to_string())
                print("="*40 + "\n")
            # --- END: สิ้นสุดส่วนที่แก้ไข ---

            self.final_summary_data = summary_df 
            self.confirm_payout_button.pack(pady=(10, 20), padx=20, ipady=10, side="bottom", anchor="se")
            self.confirm_payout_button.tkraise()
        else:
            for widget in self.final_summary_frame.winfo_children(): widget.destroy()
            message = final_result.get('message', 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ')
            CTkLabel(self.final_summary_frame, text=message).pack(pady=20)
    
    def _confirm_payout_and_save(self):
        """
        (เวอร์ชันปรับปรุงล่าสุด)
        ✅ ยืนยันการจ่ายเงิน, อัปเดตสถานะ SO, และบันทึก Log การจ่ายเงิน
        ✅ คำนวณยอดหักเฉพาะ Decentive (ไม่นับค่าดำเนินการ / หัก ณ ที่จ่าย)
        ✅ รองรับชื่อ Gross Commission ที่แตกต่างกันในแต่ละ Plan
        ✅ รองรับชื่อยอดขาย Normal / Below ของทุกแผน (รวม Tier 1/2/3)
        """
        try:
            # --- ตรวจสอบความพร้อมก่อน ---
            if not hasattr(self, 'latest_commission_result') or not self.latest_commission_result:
                messagebox.showwarning("ยังไม่พร้อม", "กรุณากด 'คำนวณขั้นสุดท้าย' ก่อนยืนยันการจ่ายเงิน", parent=self)
                return

            if not messagebox.askyesno(
                "ยืนยันการจ่ายเงิน",
                "คุณยืนยันที่จะบันทึกการจ่ายเงินนี้ใช่หรือไม่?\n"
                "การดำเนินการนี้จะอัปเดตสถานะ SO ทั้งหมดเป็น 'Paid' และไม่สามารถย้อนกลับได้จากหน้านี้",
                parent=self
            ):
                return
            
            # --- 1. ดึงข้อมูลสรุปทั้งหมด ---
            payout_notes = self.payout_notes_entry.get("1.0", "end-1c").strip()
            plan_name = self.sales_user_info.get(self.selected_sale_for_process.get(), {}).get('plan', 'N/A')
            
            final_summary_df = None
            initial_summary_df = None 

            result_type = self.latest_commission_result.get('type')
            if result_type == 'summary_plan_a':
                final_summary_df = self.latest_commission_result.get('summary')
                initial_summary_df = self.latest_commission_result.get('summary') 
            elif result_type == 'summary_other':
                final_summary_df = self.latest_commission_result.get('data') 
                initial_summary_df = self.latest_commission_result.get('data') 

            if final_summary_df is None or initial_summary_df is None:
                messagebox.showerror("ผิดพลาด", "ไม่พบข้อมูลสรุปผลการคำนวณ (Summary DF is None)", parent=self)
                return
            
            # --- ฟังก์ชันช่วย ---
            def get_final_summary_value(key_name, default=0.0):
                try:
                    value = final_summary_df.loc[final_summary_df['description'] == key_name, 'value'].values[0]
                    return float(value)
                except (IndexError, KeyError):
                    return default

            def get_initial_summary_value(key_name, default=0.0):
                if initial_summary_df is None:
                    return default
                try:
                    if 'description' in initial_summary_df.columns:
                        value = initial_summary_df.loc[initial_summary_df['description'] == key_name, 'value'].values[0]
                        return float(value)
                    else:
                        print("Warning: 'description' column not found in initial_summary_df.")
                        return default
                except (IndexError, KeyError):
                    return default

            # --- ✅ Logic ยืดหยุ่นในการดึงยอดขาย Normal / Below ---
            try:
                # Normal (Tier 1)
                normal_row = initial_summary_df[
                    initial_summary_df['description'].str.contains("ปกติ|Normal|Tier 1", case=False, na=False)
                ]
                total_normal_sales = float(normal_row['value'].iloc[0]) if not normal_row.empty else 0.0

                # Below Tier (Tier 2 หรือ Tier 3)
                below_row = initial_summary_df[
                    initial_summary_df['description'].str.contains("Below|นอกเงื่อนไข|Tier 2|Tier 3", case=False, na=False)
                ]
                total_below_sales = below_row['value'].sum() if not below_row.empty else 0.0

                total_sales = total_normal_sales + total_below_sales

                print(f"\n💡 ตรวจพบยอดขายจาก Summary:")
                print(f"   - Normal (Tier 1): {total_normal_sales:,.2f}")
                print(f"   - Below (Tier 2+3): {total_below_sales:,.2f}")
                print(f"   - Total Sales: {total_sales:,.2f}")

            except Exception:
                total_sales = total_normal_sales = total_below_sales = 0.0
                print("⚠️ ไม่พบข้อมูลยอดขาย Normal/Below ใน summary dataframe")

            # --- ✅ START: ปรับ logic การคำนวณยอดหัก ---
            # 1. Incentive รวม
            incentives_df = final_summary_df[
                final_summary_df['description'].str.startswith('(+) ')
            ]
            incentives_total = incentives_df['value'].sum()

            # 2. Deduction (เฉพาะรายการจริง)
            deductions_df = final_summary_df[
                final_summary_df['description'].str.startswith('(-) ')
                & (~final_summary_df['description'].str.contains('ดำเนินการ', case=False, na=False))
                & (~final_summary_df['description'].str.contains('หัก ณ ที่จ่าย', case=False, na=False))
            ]
            deductions_total = deductions_df['value'].sum()

            print("\n🟡 รายการที่ถูกนับเป็น (-) Deductions (หลังกรอง):")
            print(deductions_df[['description', 'value']])
            print(f"✅ Deductions Total (ไม่รวมค่าดำเนินการ/WHT): {deductions_total:,.2f}")
            # --- ✅ END ---

            # --- ✅ Logic ตรวจหา Gross Commission แบบยืดหยุ่น ---
            try:
                gross_row = final_summary_df[
                    final_summary_df['description'].str.contains("ขั้นต้น|ก่อนหัก|Gross", case=False, na=False)
                ]
                final_commission_val = float(gross_row['value'].iloc[0]) if not gross_row.empty else 0.0
            except Exception:
                final_commission_val = 0.0

            # --- 2. เตรียมข้อมูลสำหรับบันทึกลง log ---
            log_data = {
                "sale_key": self.selected_sale_for_process.get(),
                "plan_name": plan_name,
                "payout_period_text": self.current_period_text,
                "commission_month": self.selected_month,
                "commission_year": self.selected_year,
                "calculated_commission": float(self.latest_commission_result.get('final_commission_pre_deductions', 0.0)),
                "incentives_total": float(incentives_total),
                "deductions_total": float(deductions_total),
                "final_commission": final_commission_val,  # ✅ ใช้ค่าที่หาได้แบบยืดหยุ่น
                "withholding_tax": get_final_summary_value("(-) หัก ณ ที่จ่าย 3%"),
                "net_commission": get_final_summary_value("ยอดสรุปคอมหลังหัก ณ ที่จ่าย"),
                "notes": payout_notes,
                "summary_data_json": final_summary_df.to_json(orient='records'),
                "so_ids_json": json.dumps(self.current_so_ids),
                "total_sales": float(total_sales),
                "total_normal_sales": float(total_normal_sales),
                "total_below_sales": float(total_below_sales)
            }
            log_data = {k: v for k, v in log_data.items() if v is not None}

            # --- Debug: แสดงค่าที่จะบันทึกลงฐานข้อมูล ---
            print("\n🧾 ข้อมูล log_data ที่จะถูกบันทึก:")
            for k, v in log_data.items():
                print(f"  {k}: {v}")

            # --- 3. บันทึกลงฐานข้อมูล ---
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
                        cursor.execute("""
                            UPDATE commissions 
                            SET status = 'Paid', payout_id = %s
                            WHERE id IN %s
                        """, (payout_id, so_ids_tuple))
                
                conn.commit()
                messagebox.showinfo(
                    "สำเร็จ",
                    f"บันทึกการจ่ายเงิน (Payout ID: {payout_id}) เรียบร้อยแล้ว!\n"
                    f"อัปเดตสถานะ SO จำนวน {len(self.current_so_ids)} รายการเป็น 'Paid'",
                    parent=self
                )
                
                self._on_sale_selected_for_process()
                self._load_payout_history()
            
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดระหว่างบันทึกข้อมูล: {e}", parent=self)
                traceback.print_exc()
            finally:
                if conn:
                    self.app_container.release_connection(conn)

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