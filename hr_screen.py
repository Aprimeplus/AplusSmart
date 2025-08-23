# hr_screen.py (ฉบับแก้ไขและปรับปรุง)

import tkinter as tk
from tkinter import ttk
from customtkinter import (CTkFrame, CTkLabel, CTkEntry, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkTabview, filedialog,
                           CTkInputDialog, CTkOptionMenu, CTkCheckBox, CTkTextbox, CTkComboBox, CTkRadioButton, CTkToplevel)
from tkinter import messagebox
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
from hr_windows import HRVerificationWindow, PayoutDetailWindow, PayoutCalculationViewer
from tkinter import ttk, filedialog


import matplotlib
matplotlib.use('TkAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import FuncFormatter, MaxNLocator

from sqlalchemy import create_engine

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
        self.geometry("500x400")
        self.grab_set()
        self.transient(master)

        self.result = None
        self.imported_df = None
        self.manual_df = pd.DataFrame(columns=['so_number', 'sales_uploaded', 'cost_uploaded'])

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        sales_frame = CTkFrame(self)
        sales_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        CTkLabel(sales_frame, text="1. เลือกพนักงานขาย:", font=master.label_font).pack(side="left", padx=10)
        self.selected_sale = tk.StringVar(value="ทั้งหมด")
        self.sale_dropdown = CTkOptionMenu(sales_frame, variable=self.selected_sale, values=["ทั้งหมด"] + sales_keys, command=self._check_run_button_state)
        self.sale_dropdown.pack(side="left", padx=10, pady=10)

        source_frame = CTkFrame(self)
        source_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        source_frame.grid_columnconfigure(1, weight=1)
        CTkLabel(source_frame, text="2. เลือกแหล่งข้อมูล (อย่างน้อย 1 อย่าง):", font=master.label_font).grid(row=0, column=0, columnspan=2, sticky="w", padx=10)
        
        self.import_button = CTkButton(source_frame, text="นำเข้าไฟล์ Excel/CSV", command=self._on_import_file)
        self.import_button.grid(row=1, column=0, padx=10, pady=10)
        self.file_label = CTkLabel(source_frame, text="ยังไม่ได้เลือกไฟล์", text_color="gray", anchor="w")
        self.file_label.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.manual_button = CTkButton(source_frame, text="เพิ่มข้อมูลด้วยมือ...", command=self._on_add_manual)
        self.manual_button.grid(row=2, column=0, padx=10, pady=10)

        manual_display_frame = CTkScrollableFrame(self, label_text="รายการที่คีย์ด้วยมือ")
        manual_display_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.manual_display_label = CTkLabel(manual_display_frame, text="ไม่มีข้อมูล")
        self.manual_display_label.pack(pady=10)

        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=3, column=0, pady=20)
        self.run_button = CTkButton(button_frame, text="เริ่มการเปรียบเทียบ", command=self._on_run_comparison, state="disabled")
        self.run_button.pack(side="left", padx=10)
        CTkButton(button_frame, text="ยกเลิก", command=self.destroy, fg_color="gray").pack(side="left", padx=10)

    def _check_run_button_state(self, *args):
        has_data = (self.imported_df is not None) or (not self.manual_df.empty)
        if has_data:
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
        self.result = {
            "salesperson": self.selected_sale.get(),
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
        self.user_role = user_role # << เพิ่มบรรทัดนี้เพื่อเก็บค่า role ไว้

        self.label_font = CTkFont(size=16, weight="bold", family="Roboto")
        self.entry_font = CTkFont(size=14, family="Roboto")
        self.header_font_table = CTkFont(size=14, weight="bold", family="Roboto")

        self.header_map = app_container.HEADER_MAP

        self.db_df, self.uploaded_df, self.comparison_df, self.user_df, self.comparison_log_df = None, None, None, None, None
        self.initial_commission_result = None
        self.current_comm_df = None
        self.manual_entry_df = pd.DataFrame(columns=['so_number', 'sales_uploaded', 'cost_uploaded'])
        self.uploaded_file_path, self.sales_chart_canvas, self.po_chart_canvas, self.sales_target_chart_canvas = None, None, None, None
        self.selected_payout_ids = set(); self.select_all_var = tk.IntVar(value=0)
        self.theme = self.app_container.THEME["hr"]
        
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.period_options = ["ปีนี้", "เดือนนี้"] + self.thai_months
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}

        self.history_current_page, self.history_rows_per_page, self.history_total_rows = 0, 20, 0
        self.user_current_page, self.user_rows_per_page, self.user_total_rows = 0, 20, 0

        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        header_frame = CTkFrame(self, fg_color="transparent"); header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 0))
        CTkLabel(header_frame, text=f"หน้าจอสำหรับฝ่ายบุคคล (HR): {self.user_name}", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).pack(side="left")
        CTkButton(header_frame, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").pack(side="right")

        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1, segmented_button_selected_color=self.theme["primary"], segmented_button_unselected_hover_color="#A7F3D0", fg_color=self.cget("fg_color"), command=self._on_tab_selected)
        self.tab_view.grid(row=1, column=0, pady=10, padx=20, sticky="nsew")

        self.dashboard_tab = self.tab_view.add("Dashboard สรุปภาพรวม"); self.sales_target_tab = self.tab_view.add("วิเคราะห์เป้าการขาย"); self.manage_users_tab = self.tab_view.add("จัดการผู้ใช้งาน"); self.compare_commission_tab = self.tab_view.add("เปรียบเทียบ / ดูประวัติ"); self.process_commission_tab = self.tab_view.add("ประมวลผลและจ่ายค่าคอม"); self.payout_history_tab = self.tab_view.add("ประวัติการจ่ายค่าคอม"); self.audit_log_tab = self.tab_view.add("บันทึกกิจกรรม")
        self._create_dashboard_tab(self.dashboard_tab);self._create_payout_history_tab(self.payout_history_tab)  ;self._create_sales_target_tab(self.sales_target_tab); self._create_manage_users_tab(self.manage_users_tab); self._create_compare_commission_tab(self.compare_commission_tab); self._create_process_commission_tab(self.process_commission_tab); self._create_audit_log_tab(self.audit_log_tab)
        self.tab_view.set("จัดการผู้ใช้งาน")
        self.after(100, self._initial_load)
        self._payout_history_loaded = False 
        self._dashboard_loaded, self._sales_target_loaded, self._users_loaded, self._compare_commission_loaded, self._process_commission_loaded, self._audit_log_loaded = False, False, False, False, False, False
    
    # อยู่ในไฟล์ hr_screen.py ภายในคลาส HRScreen
    
    

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
        self._load_payout_history()

    def _open_comparison_history_window(self):
        from hr_windows import ComparisonHistoryWindow # Import ที่นี่เพื่อเลี่ยง Circular Import
        ComparisonHistoryWindow(master=self, app_container=self.app_container)

    def _create_payout_history_tab(self, parent_tab):
        """สร้าง Layout สำหรับหน้าประวัติการจ่ายเงิน (ฉบับปรับปรุงมีฟิลเตอร์)"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        # --- Frame หลักสำหรับตัวกรองทั้งหมด ---
        filter_container = CTkFrame(parent_tab, fg_color="transparent")
        filter_container.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

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
        CTkLabel(filter_container, text="ค้นหาพนักงานขาย:").pack(side="left", padx=(20, 2))
        self.payout_search_entry = CTkEntry(filter_container, placeholder_text="พิมพ์รหัสพนักงานขาย...")
        self.payout_search_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # --- ปุ่ม ---
        CTkButton(filter_container, text="ค้นหา", command=self._load_payout_history).pack(side="left", padx=10)
        CTkButton(filter_container, text="ล้างค่า", command=self._reset_payout_filters).pack(side="left")

        # --- Frame สำหรับแสดงตาราง ---
        self.payout_history_frame = CTkFrame(parent_tab)
        self.payout_history_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.payout_history_frame.grid_columnconfigure(0, weight=1)
        self.payout_history_frame.grid_rowconfigure(0, weight=1)

    def _load_payout_history(self):
        """โหลดข้อมูลจากตาราง commission_payout_logs และแสดงผลตามฟิลเตอร์ (ฉบับคำนวณ Margin)"""
        loading = self._show_loading(self.payout_history_frame)
        try:
            # ดึงคอลัมน์ที่จำเป็นทั้งหมดจากฐานข้อมูล
            base_query = "SELECT id, timestamp, hr_user_key, sale_key, plan_name, summary_json, notes, total_sales_for_margin, total_cost_for_margin, sales_target_at_payout FROM commission_payout_logs"
            where_clauses = []
            params = []

            # สร้างเงื่อนไขจากฟิลเตอร์
            selected_month = self.payout_month_var.get()
            if selected_month != "ทุกเดือน":
                month_num = self.thai_month_map[selected_month]
                where_clauses.append("EXTRACT(MONTH FROM timestamp) = %s")
                params.append(month_num)
            selected_year = self.payout_year_var.get()
            if selected_year != "ทุกปี":
                year_num = int(selected_year)
                where_clauses.append("EXTRACT(YEAR FROM timestamp) = %s")
                params.append(year_num)
            search_term = self.payout_search_entry.get().strip()
            if search_term:
                where_clauses.append("sale_key ILIKE %s")
                params.append(f"%{search_term}%")
            if where_clauses:
                base_query += " WHERE " + " AND ".join(where_clauses)
            base_query += " ORDER BY timestamp DESC"
            
            df = pd.read_sql_query(base_query, self.pg_engine, params=tuple(params))

            if df.empty:
                loading.destroy()
                CTkLabel(self.payout_history_frame, text="ไม่พบข้อมูลตามเงื่อนไขที่เลือก").pack(pady=20)
                return
            
            # เตรียม List ว่างสำหรับเก็บค่าที่จะคำนวณ
            avg_margin_list = []
            hit_target_percent_list = [] 

            # วนลูปเพื่อคำนวณค่า Margin และ %Hit Target
            for index, row in df.iterrows():
                sales = row['total_sales_for_margin']
                cost = row['total_cost_for_margin']
                target = row['sales_target_at_payout']

                # คำนวณ Margin
                if pd.notna(sales) and pd.notna(cost) and sales > 0:
                    margin = ((sales - cost) / sales) * 100
                    avg_margin_list.append(margin)
                else:
                    avg_margin_list.append("N/A")

                # คำนวณ % Hit Target
                if pd.notna(sales) and pd.notna(target) and target > 0:
                    percent_achieved = (sales / target) * 100
                    hit_target_percent_list.append(percent_achieved)
                else:
                    hit_target_percent_list.append(0.0)

            # เพิ่มคอลัมน์ที่คำนวณใหม่เข้าไปใน DataFrame
            df['avg_margin'] = avg_margin_list
            df['hit_target_percent'] = hit_target_percent_list
            
            # เตรียม List ว่างสำหรับค่าที่ดึงจาก JSON
            pre_tax_comm_list = []
            incentive_list = []
            gross_comm_list = []
            net_commission_list = []

            # วนลูปเพื่อดึงข้อมูลจาก summary_json
            for summary_data in df['summary_json']:
                try:
                    def get_value(desc_keyword, default=0.0):
                        item = next((i for i in summary_data if desc_keyword in i['description']), None)
                        return item.get('value') if item else default
                    pre_tax_comm_list.append(get_value("ยอดคอมมิชชั่นก่อนหักภาษี"))
                    total_incentive = sum(item.get('value', 0.0) for item in summary_data if '(+) Incentive' in item.get('description', ''))
                    incentive_list.append(total_incentive)
                    gross_comm_list.append(get_value("ยอดคอมมิชชั่นขั้นต้น"))
                    net_commission_list.append(get_value("หลังหัก ณ ที่จ่าย"))
                except (TypeError, KeyError):
                    [lst.append(0.0) for lst in [pre_tax_comm_list, incentive_list, gross_comm_list, net_commission_list]]

            # เพิ่มคอลัมน์ที่ดึงจาก JSON เข้าไปใน DataFrame
            df['pre_tax_comm'] = pre_tax_comm_list
            df['incentives'] = incentive_list
            df['gross_comm'] = gross_comm_list
            df['net_commission'] = net_commission_list

            # สร้าง DataFrame สุดท้ายสำหรับแสดงผล
            df_display = df[[
                'id', 'timestamp', 'sale_key', 'total_sales_for_margin', 'sales_target_at_payout', 
                'hit_target_percent', 'avg_margin', 'incentives', 'gross_comm', 
                'pre_tax_comm', 'net_commission', 'notes'
            ]].copy()

            # เปลี่ยนชื่อคอลัมน์เป็นภาษาไทย
            df_display.rename(columns={
                'id': 'ID', 'timestamp': 'วันที่จ่าย', 'sale_key': 'พนักงานขาย', 
                'total_sales_for_margin': 'ยอดขายรวม', 'sales_target_at_payout': 'ยอดเป้าหมาย',
                'hit_target_percent': 'Hit Target (%)', 'avg_margin': 'Margin เฉลี่ย (%)',
                'incentives': 'Incentive', 'gross_comm': 'คอม+Incentive',
                'pre_tax_comm': 'ยอดก่อนหักภาษี', 'net_commission': 'ยอดสุทธิ', 'notes': 'หมายเหตุ'
            }, inplace=True)

            loading.destroy()
            # สร้างตารางจากข้อมูลที่เตรียมไว้
            self._create_styled_dataframe_table(
                self.payout_history_frame, 
                df_display, 
                label_text="",
                on_row_click=self._on_payout_history_double_click
            )

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดประวัติการจ่ายเงินได้: {e}", parent=self)
            traceback.print_exc()

    def _on_payout_history_double_click(self, event, tree, df):
        """เมื่อดับเบิลคลิกที่ประวัติ จะเปิดหน้าต่างแสดงรายละเอียดการคำนวณ"""
        try:
            record_id_str = tree.focus()
            if not record_id_str: return

            payout_id = int(tree.item(record_id_str, "values")[0])
            
            # เรียกเปิดหน้าต่างใหม่ที่เราสร้าง
            PayoutCalculationViewer(
                master=self, 
                app_container=self.app_container, 
                payout_id=payout_id
            )

        except (ValueError, IndexError) as e:
            messagebox.showwarning("ผิดพลาด", "ไม่สามารถอ่าน Payout ID จากแถวที่เลือกได้", parent=self)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างรายละเอียดได้: {e}", parent=self)
            traceback.print_exc()

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
        เปิดหน้าต่างตรวจสอบข้อมูล โดยดึงข้อมูลล่าสุดจาก DB โดยตรง
        (เวอร์ชันแก้ไขให้ทำงานได้จากทุกที่)
        """
        try:
            # --- จุดที่แก้ไข: Query ข้อมูล SO ล่าสุดจาก DB โดยตรง ---
            query = """
                SELECT c.*, po.cogs_db, po.po_shipping_stock, po.po_shipping_site, po.po_relocation, u.sale_name 
                FROM commissions c 
                JOIN sales_users u ON c.sale_key = u.sale_key
                LEFT JOIN (
                    SELECT so_number, SUM(total_cost) as cogs_db,
                           SUM(shipping_to_stock_cost) as po_shipping_stock,
                           SUM(shipping_to_site_cost) as po_shipping_site,
                           SUM(relocation_cost) as po_relocation
                    FROM purchase_orders WHERE status = 'Approved' GROUP BY so_number
                ) po ON c.so_number = po.so_number
                WHERE c.is_active = 1 AND c.so_number = %s
            """
            system_data_df = pd.read_sql_query(query, self.app_container.pg_engine, params=(so_number,))

            if system_data_df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูลที่ Active สำหรับ SO: {so_number} ในระบบ", parent=self)
                return

            system_data = system_data_df.iloc[0].to_dict()
            
            # --- โค้ดส่วนที่เหลือเหมือนเดิม แต่จะใช้ข้อมูลที่เพิ่งดึงมา ---
            excel_data = {}
            if self.uploaded_df is not None and not self.uploaded_df.empty:
                excel_data_row = self.uploaded_df[self.uploaded_df['so_number'].astype(str).str.strip() == str(so_number).strip()]
                if not excel_data_row.empty:
                    excel_data = excel_data_row.iloc[0].to_dict()

            po_data = self._get_po_data(so_number)
            
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
        Callback เมื่อมีการดับเบิลคลิกบนตารางเปรียบเทียบ
        จะดึง SO Number และเปิดหน้าต่าง HRVerificationWindow
        """
        try:
            # 1. หา item id (iid) ของแถวที่ถูกคลิก
            selected_item_iid = tree.focus()
            if not selected_item_iid:
                return

            # 2. ดึงข้อมูลทั้งหมดในแถวนั้นออกมา
            values = tree.item(selected_item_iid, "values")
            if not values:
                return

            # 3. ดึง SO Number จากคอลัมน์แรก (index 0)
            so_number = values[0]

            # 4. ตรวจสอบว่า so_number ไม่ใช่ค่าว่างหรือค่าที่ไม่ถูกต้อง
            if not so_number or so_number == '0':
                messagebox.showwarning("ข้อมูลผิดพลาด", "ไม่สามารถระบุ SO Number จากแถวที่เลือกได้", parent=self)
                return
                
            # 5. เปิดหน้าต่างตรวจสอบข้อมูล
            self._open_verification_window(so_number)

        except (IndexError, ValueError) as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถอ่านข้อมูลจากแถวที่เลือกได้: {e}", parent=self)
            traceback.print_exc()
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างตรวจสอบได้: {e}", parent=self)
            traceback.print_exc()

    def _initial_load(self):
        self._populate_users_table()
        self._users_loaded = True
        
    def _on_tab_selected(self):
        selected_tab_name = self.tab_view.get()
        if selected_tab_name == "Dashboard สรุปภาพรวม" and not self._dashboard_loaded: self._update_dashboard(); self._dashboard_loaded = True
        elif selected_tab_name == "วิเคราะห์เป้าการขาย" and not self._sales_target_loaded: self._update_sales_target_dashboard(); self._sales_target_loaded = True
        elif selected_tab_name == "จัดการผู้ใช้งาน" and not self._users_loaded: self._populate_users_table(); self._users_loaded = True
        elif selected_tab_name == "เปรียบเทียบ / ดูประวัติ" and not self._compare_commission_loaded: self._compare_commission_loaded = True
        elif selected_tab_name == "ประมวลผลและจ่ายค่าคอม" and not self._process_commission_loaded: self._on_sale_selected_for_process(); self._process_commission_loaded = True
        elif selected_tab_name == "ประวัติการจ่ายค่าคอม" and not self._payout_history_loaded:
            self._load_payout_history()
            self._payout_history_loaded = True
        elif selected_tab_name == "บันทึกกิจกรรม" and not self._audit_log_loaded: self._populate_audit_log_table(); self._audit_log_loaded = True
    
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
            period = self.sales_target_period_var.get(); start_date, end_date = self._get_date_range_from_period(period); start_date_str, end_date_str = start_date.strftime("%Y-%m-%d %H:%M:%S"), end_date.strftime("%Y-%m-%d %H:%M:%S"); sales_vs_target_data = self._get_sales_vs_target_data(start_date_str, end_date_str); loading.destroy(); self._create_sales_vs_target_chart(self.sales_target_chart_frame, sales_vs_target_data)
        except Exception as e: loading.destroy(); messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอัปเดต Dashboard: {e}", parent=self)

    def _get_sales_vs_target_data(self, start_date, end_date):
        try:
            query = "SELECT su.sale_name, su.sales_target, COALESCE(SUM(c.sales_service_amount), 0) as total_sales FROM sales_users su LEFT JOIN commissions c ON su.sale_key = c.sale_key AND c.is_active = 1 AND c.timestamp BETWEEN %s AND %s WHERE su.role = 'Sale' AND su.sales_target > 0 AND su.status = 'Active' GROUP BY su.sale_key, su.sale_name, su.sales_target ORDER BY su.sale_name;"; df = pd.read_sql_query(query, self.pg_engine, params=(start_date, end_date)); return df
        except Exception as e: print(f"Error getting sales vs target data: {e}"); messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลเป้าหมายการขายได้: {e}", parent=self); return pd.DataFrame(columns=['sale_name', 'sales_target', 'total_sales'])

    def _create_sales_vs_target_chart(self, parent_frame, data_df):
        if hasattr(self, 'sales_target_chart_canvas') and self.sales_target_chart_canvas: self.sales_target_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children(): widget.destroy()
        if data_df.empty: CTkLabel(parent_frame, text="ไม่พบข้อมูลพนักงานขายที่มีการตั้งเป้าหมาย", font=self.header_font_table).pack(expand=True); return
        fig = Figure(figsize=(10, 6), dpi=100, facecolor=self.theme["bg"]); ax = fig.add_subplot(111); ax.set_facecolor(self.theme["bg"]); font_name = 'Tahoma'; formatter = FuncFormatter(lambda y, pos: f'{y:,.0f}'); ax.yaxis.set_major_formatter(formatter); x = np.arange(len(data_df['sale_name'])); width = 0.35; rects1 = ax.bar(x - width/2, data_df['total_sales'], width, label='ยอดขายจริง', color=self.theme["primary"]); rects2 = ax.bar(x + width/2, data_df['sales_target'], width, label='ยอดเป้าหมาย', color='#CBD5E1', edgecolor='#94A3B8', linewidth=1); ax.set_ylabel('ยอดขาย (บาท)', fontname=font_name, fontsize=12); ax.set_title('กราฟเปรียบเทียบยอดขายจริงกับยอดเป้าหมาย', fontname=font_name, fontsize=18, weight="bold"); ax.set_xticks(x); ax.set_xticklabels(data_df['sale_name'], rotation=45, ha="right", fontname=font_name, fontsize=11); ax.legend(prop={'family': font_name, 'size': 12}); ax.grid(axis='y', linestyle='--', alpha=0.7); ax.bar_label(rects1, padding=3, fmt='{:,.0f}', fontsize=9); ax.bar_label(rects2, padding=3, fmt='{:,.0f}', fontsize=9); fig.tight_layout(); canvas = FigureCanvasTkAgg(fig, master=parent_frame); canvas.draw(); canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10); self.sales_target_chart_canvas = canvas

    def _get_po_status_summary(self, start_date, end_date):
        try: query = "SELECT status, COUNT(id) as count FROM purchase_orders WHERE timestamp BETWEEN %s AND %s GROUP BY status"; df = pd.read_sql_query(query, self.pg_engine, params=(start_date, end_date)); return df
        except Exception as e: print(f"Error getting PO status summary: {e}"); messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลสถานะ PO ได้: {e}", parent=self); return pd.DataFrame(columns=['status', 'count'])

    def _create_po_pie_chart(self, parent_frame, data_df):
        if hasattr(self, 'po_chart_canvas') and self.po_chart_canvas: self.po_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children(): widget.destroy()
        if data_df.empty: CTkLabel(parent_frame, text="ไม่พบข้อมูลใบสั่งซื้อในช่วงเวลานี้", font=self.header_font_table).pack(expand=True); return
        fig = Figure(figsize=(5, 4), dpi=100, facecolor=self.theme["bg"]); ax = fig.add_subplot(111); status_colors_map = { "Approved": "#BBF7D0", "Pending Approval": "#FEF08A", "Rejected": "#FECACA", "Draft": "#E5E7EB" }; pie_colors = [status_colors_map.get(status, "#B0B0B0") for status in data_df['status']]; ax.pie(data_df['count'], labels=data_df['status'], autopct='%1.1f%%', startangle=90, colors=pie_colors, textprops={'fontname': 'Tahoma', 'fontsize': 12}); ax.axis('equal'); ax.set_title('สัดส่วนสถานะใบสั่งซื้อ (PO)', fontname='Tahoma', fontsize=16, weight="bold"); fig.tight_layout(); canvas = FigureCanvasTkAgg(fig, master=parent_frame); canvas.draw(); canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5); self.po_chart_canvas = canvas
    
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

            sales_by_employee_data = self._get_sales_by_employee_data(start_date_str, end_date_str)
            po_summary_data = self._get_po_status_summary(start_date_str, end_date_str)

            loading1.destroy()
            self._create_sales_by_employee_chart(self.sales_chart_frame, sales_by_employee_data)

            loading2.destroy()
            self._create_po_pie_chart(self.po_chart_frame, po_summary_data)

        except Exception as e:
            if 'loading1' in locals() and loading1.winfo_exists(): loading1.destroy()
            if 'loading2' in locals() and loading2.winfo_exists(): loading2.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอัปเดต Dashboard: {e}", parent=self)

    def _get_sales_by_employee_data(self, start_date, end_date):
        try:
            query = """
                SELECT su.sale_name, COALESCE(SUM(c.sales_service_amount), 0) as total_sales
                FROM sales_users su
                LEFT JOIN commissions c ON su.sale_key = c.sale_key
                                     AND c.is_active = 1
                                     AND c.timestamp BETWEEN %s AND %s
                WHERE su.role = 'Sale' AND su.status = 'Active'
                GROUP BY su.sale_key, su.sale_name
                HAVING COALESCE(SUM(c.sales_service_amount), 0) > 0
                ORDER BY total_sales DESC;
            """
            df = pd.read_sql_query(query, self.pg_engine, params=(start_date, end_date))
            return df
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลยอดขายตามพนักงานได้: {e}", parent=self)
            return pd.DataFrame(columns=['sale_name', 'total_sales'])

    def _create_sales_by_employee_chart(self, parent_frame, data_df):
        if hasattr(self, 'sales_chart_canvas') and self.sales_chart_canvas:
            self.sales_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children():
            widget.destroy()

        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูลยอดขายตามพนักงานในช่วงเวลานี้", font=self.header_font_table).pack(expand=True)
            return

        colors = ['#2a9d8f', '#e9c46a', '#f4a261', '#e76f51', '#264653']
        bar_colors = [colors[i % len(colors)] for i in range(len(data_df))]

        fig = Figure(figsize=(6, 4), dpi=100, facecolor=self.theme["bg"])
        ax = fig.add_subplot(111)
        ax.set_facecolor(self.theme["bg"])

        bars = ax.bar(data_df['sale_name'], data_df['total_sales'], color=bar_colors)
        
        ax.set_ylabel('ยอดขาย (บาท)', fontname='Tahoma', fontsize=12)
        ax.set_title('สรุปยอดขายตามพนักงาน (Active)', fontname='Tahoma', fontsize=16, weight="bold")
        
        if len(data_df) > 4:
            ax.tick_params(axis='x', rotation=45)

        formatter = FuncFormatter(lambda y, pos: f'{int(y):,}')
        ax.yaxis.set_major_formatter(formatter)
        ax.grid(axis='y', linestyle='--', alpha=0.7)

        ax.bar_label(bars, fmt='{:,.0f}', padding=3, fontname='Tahoma', fontsize=9)
        
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
        
        self.role_var = tk.StringVar(value="Sale")
        self.role_menu = CTkOptionMenu(manage_frame, variable=self.role_var, values=["Sale", "Sales Manager", "Purchasing Staff", "Purchasing Manager", "Director", "HR"], command=self._on_role_changed)
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

            # --- จุดที่แก้ไข: กำหนดชุดสีตามแผนก (Role) ---
            role_colors = {
                "Sale": "#EFF6FF",               # สีฟ้าอ่อน (Theme ฝ่ายขาย)
                "Sales Manager": "#DBEAFE",       # สีฟ้าเข้มขึ้น
                "Purchasing Staff": "#F5F3FF",    # สีม่วงอ่อน (Theme ฝ่ายจัดซื้อ)
                "Purchasing Manager": "#EDE9FE",  # สีม่วงเข้มขึ้น
                "HR": "#F0FDF4",               # สีเขียวอ่อน (Theme HR)
                "Director": "#F3F4F6",            # สีเทาอ่อน (สำหรับกรรมการ)
                "กรรมการ": "#F3F4F6"
            }

            # --- จุดที่แก้ไข: ส่ง status_column และ status_colors เข้าไป ---
            self._create_styled_dataframe_table(
                parent=table_frame, 
                df=self.user_df, 
                label_text="ข้อมูลผู้ใช้งาน", 
                on_row_click=self._on_user_row_click_treeview,
                status_column="ประเภท",      # บอกให้ใช้คอลัมน์ "ประเภท" เป็นเงื่อนไข
                status_colors=role_colors   # บอกให้ใช้ชุดสีที่เพิ่งกำหนด
            )
            
            self.user_page_label.configure(text=f"Page {self.user_current_page + 1} / {max(1, total_pages)}")
            self.user_prev_button.configure(state="normal" if self.user_current_page > 0 else "disabled")
            self.user_next_button.configure(state="normal" if self.user_current_page < total_pages - 1 else "disabled")
        except Exception as e:
            if loading_label.winfo_exists(): loading_label.destroy()
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดผู้ใช้งานได้: {e}", parent=self)
    
    def _on_user_row_click_treeview(self, event, tree, df):
        try:
            record_id = tree.focus()
            if not record_id:
                return

            filtered_df = df.loc[df['User Key'] == record_id]

            if not filtered_df.empty:
                row_data = filtered_df.iloc[0]
                self._on_user_row_click(row_data)
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

        control_frame = CTkFrame(parent_tab)
        control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        CTkButton(control_frame, text="🚀 เริ่มต้นการเปรียบเทียบใหม่", command=self._start_new_comparison, font=CTkFont(size=16, weight="bold")).pack(side="left", padx=10, pady=10)
        CTkButton(control_frame, text="📖 แสดงประวัติการเปรียบเทียบ", command=self._open_comparison_history_window, fg_color="#64748B").pack(side="left", padx=10, pady=10)

        # --- ปุ่ม "บันทึกผลการเปรียบเทียบนี้" ถูกลบออกจากส่วนนี้แล้ว ---

        self.finalize_button = CTkButton(control_frame, text="✅ ยืนยันข้อมูลและส่งต่อเพื่อคำนวณค่าคอม", 
                                         fg_color="#16A34A", hover_color="#15803D", 
                                         command=self._finalize_comparison)
        self.finalize_button.pack(side="right", padx=10, pady=10)
        self.finalize_button.pack_forget()

        self.export_button = CTkButton(control_frame, text="📄 Export ผลลัพธ์", command=self._export_comparison)
        self.export_button.pack(side="right", padx=10, pady=10)
        self.export_button.pack_forget()

        self.results_frame_label = CTkLabel(parent_tab, text="กรุณากด 'เริ่มต้นการเปรียบเทียบใหม่' เพื่อเริ่มใช้งาน", font=self.label_font, text_color="gray")
        self.results_frame_label.grid(row=1, column=0, padx=10, pady=(5, 0), sticky="w")

        self.results_frame = CTkFrame(parent_tab)
        self.results_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.results_frame.grid_rowconfigure(0, weight=1)
        self.results_frame.grid_columnconfigure(0, weight=1)

    def _start_new_comparison(self):
        active_sales_keys = self._get_sale_keys()
        config_dialog = ComparisonConfigDialog(self, sales_keys=active_sales_keys)
        self.wait_window(config_dialog)

        if not config_dialog.result:
            return

        config = config_dialog.result
        selected_salesperson = config["salesperson"]
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
            base_query = """SELECT c.*, po.cogs_db, po.po_shipping_stock, po.po_shipping_site, po.po_relocation, u.sale_name 
                          FROM commissions c 
                          JOIN sales_users u ON c.sale_key = u.sale_key
                          LEFT JOIN (
                                SELECT 
                                    so_number, 
                                    SUM(total_cost) as cogs_db,
                                    SUM(shipping_to_stock_cost) as po_shipping_stock,
                                    SUM(shipping_to_site_cost) as po_shipping_site,
                                    SUM(relocation_cost) as po_relocation
                                FROM purchase_orders 
                                WHERE status = 'Approved' 
                                GROUP BY so_number
                            ) po ON c.so_number = po.so_number
                          WHERE c.is_active = 1 AND c.status = 'Forwarded_To_HR'"""
            # --- END: สิ้นสุดการแก้ไข Query ---
            params = []

            if selected_salesperson != "ทั้งหมด":
                base_query += " AND c.sale_key = %s"
                params.append(selected_salesperson)
            else:
                base_query += " AND c.sale_key IN (SELECT sale_key FROM sales_users WHERE status = 'Active' AND role = 'Sale')"
            
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

        # --- START: จุดแก้ไขที่สำคัญ ---
        # 1. กรองข้อมูลจาก "ตารางข้อมูล" (self.comparison_df) ไม่ใช่ "กรอบแสดงผล"
        # 2. เลือกเฉพาะรายการที่มีสถานะ 'ผ่านเกณฑ์' ซึ่งพร้อมที่จะถูกส่งต่อ
        good_statuses = ["ผ่านเกณฑ์"]
        df_to_finalize = self.comparison_df[self.comparison_df['สถานะ'].isin(good_statuses)].copy()
        # --- END: สิ้นสุดการแก้ไข ---

        if df_to_finalize.empty:
            messagebox.showinfo("ไม่พบรายการ", "ไม่พบรายการที่ 'ผ่านเกณฑ์' ที่จะส่งต่อได้ในขณะนี้", parent=self)
            return
        
        self._save_comparison_to_log()

        records_to_update = []
        for index, row in df_to_finalize.iterrows():
            so_number = row['เลขที่ SO'] # <--- แก้ไขจุดที่ 1
            
            # ดึงข้อมูลจาก DataFrame ที่ merge แล้ว เพื่อให้ได้ข้อมูลครบถ้วน
            full_row_data = self.comparison_df.loc[self.comparison_df['เลขที่ SO'] == so_number].iloc[0] # <--- แก้ไขจุดที่ 2

            sales_db_pure = full_row_data.get('ยอดขาย/บริการ (ระบบ)', 0) # <--- แก้ไขจุดที่ 3
            sales_uploaded = full_row_data.get('ยอดขาย (Express)', 0)
            cost_db = full_row_data.get('ต้นทุน (ระบบ)', 0)
            cost_uploaded = full_row_data.get('ต้นทุน (Express)', 0)
            
            sales_db_pure_cleaned = utils.convert_to_float(sales_db_pure)
            sales_uploaded_cleaned = utils.convert_to_float(sales_uploaded)
            cost_db_cleaned = utils.convert_to_float(cost_db)
            cost_uploaded_cleaned = utils.convert_to_float(cost_uploaded)

            final_sale = max(sales_db_pure_cleaned, sales_uploaded_cleaned)
            final_cost = min(cost_db_cleaned, cost_uploaded_cleaned) if cost_db_cleaned > 0 and cost_uploaded_cleaned > 0 else max(cost_db_cleaned, cost_uploaded_cleaned)

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
            
            for widget in self.results_frame.winfo_children(): widget.destroy()
            self.results_frame_label.configure(text="กรุณากด 'เริ่มต้นการเปรียบเทียบใหม่' เพื่อเริ่มใช้งาน")
            self.finalize_button.pack_forget()
            self.export_button.pack_forget()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการยืนยันข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _calculate_final_pu_cost(self, row):
        overrides = {}
        if pd.notna(row.get('hr_cost_overrides')):
            try:
                overrides = json.loads(row['hr_cost_overrides']) if isinstance(row['hr_cost_overrides'], str) else row['hr_cost_overrides']
            except (json.JSONDecodeError, TypeError):
                pass

        # --- START: แก้ไขวิธีจัดการค่าว่างให้ถูกต้อง ---
        cogs_val = float(row.get('cogs_db', 0) or 0)
        po_shipping_stock_val = float(row.get('po_shipping_stock', 0) or 0)
        po_shipping_site_val = float(row.get('po_shipping_site', 0) or 0)
        po_relocation_val = float(row.get('po_relocation', 0) or 0)
        brokerage_fee_val = float(row.get('brokerage_fee', 0) or 0)
        transfer_fee_val = float(row.get('transfer_fee', 0) or 0)

        # 2. คำนวณต้นทุนตั้งต้นจากข้อมูล PO
        # ต้นทุนสินค้า = ยอดรวม PO - ค่าส่งทั้งหมดที่รวมอยู่ใน PO
        po_product_cost = cogs_val - po_shipping_stock_val - po_shipping_site_val
        po_shipping = po_shipping_stock_val + po_shipping_site_val
        po_relocation = po_relocation_val
        
        # 3. ตรวจสอบค่าที่ถูกแก้ไขโดย HR (Overrides)
        cost_product = float(overrides.get('ต้นทุนค่าสินค้า/บริการ', po_product_cost))
        cost_shipping = float(overrides.get('ต้นทุนค่าจัดส่ง', po_shipping))
        cost_relocation = float(overrides.get('ต้นทุนค่าย้าย', po_relocation))
        cost_brokerage = float(overrides.get('ต้นทุนค่านายหน้า', brokerage_fee_val))
        cost_transfer = float(overrides.get('ต้นทุนค่าธรรมเนียมโอน', transfer_fee_val))

        # 4. รวมเป็นต้นทุนสุดท้าย
        return cost_product + cost_shipping + cost_relocation + cost_transfer + cost_brokerage

    # อยู่ในไฟล์ hr_screen.py ภายในคลาส HRScreen

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
            
            required_cols = [
                'id', 'so_number', 'sales_service_amount', 'shipping_cost', 'relocation_cost', 
                'brokerage_fee', 'transfer_fee', 'status', 'final_margin', 
                'hr_cost_overrides', 'cogs_db', 'po_shipping_stock', 'po_shipping_site', 'po_relocation',
                'cutting_drilling_fee', 'other_service_fee', 'credit_card_fee'
            ]
            for col in required_cols:
                if col not in self.db_df.columns: self.db_df[col] = np.nan
            db_compare_df = self.db_df.copy()
            
            db_compare_df['so_number'] = db_compare_df['so_number'].astype(str).str.strip()
            
            revenue_cols = [
                'sales_service_amount', 'shipping_cost', 'relocation_cost',
                'cutting_drilling_fee', 'other_service_fee', 'credit_card_fee'
            ]
            for col in revenue_cols:
                db_compare_df[col] = pd.to_numeric(db_compare_df[col], errors='coerce').fillna(0)

            db_compare_df['sales_for_comparison'] = (
                db_compare_df['sales_service_amount'] + 
                db_compare_df['shipping_cost'] + 
                db_compare_df['relocation_cost'] +
                db_compare_df['cutting_drilling_fee'] +
                db_compare_df['other_service_fee'] +
                db_compare_df['credit_card_fee']
            )
            
            db_compare_df['cost_db'] = db_compare_df.apply(self._calculate_final_pu_cost, axis=1)
            db_compare_df['payment_before_vat_po'] = db_compare_df['po_shipping_stock'].fillna(0) + db_compare_df['po_shipping_site'].fillna(0)

           

            db_compare_df['gp_db'] = db_compare_df['sales_service_amount'] - db_compare_df['cost_db']
            db_compare_df['margin_db'] = (db_compare_df['gp_db'] / db_compare_df['sales_service_amount'].replace(0, np.nan)) * 100

            uploaded_compare_df['so_number'] = uploaded_compare_df['so_number'].astype(str).str.strip()
            uploaded_compare_df['sales_uploaded'] = pd.to_numeric(uploaded_compare_df.get('sales_uploaded'), errors='coerce').fillna(0)
            uploaded_compare_df['cost_uploaded'] = pd.to_numeric(uploaded_compare_df.get('cost_uploaded'), errors='coerce').fillna(0)
            uploaded_compare_df['gp_uploaded'] = uploaded_compare_df['sales_uploaded'] - uploaded_compare_df['cost_uploaded']
            uploaded_compare_df['margin_uploaded'] = (uploaded_compare_df['gp_uploaded'] / uploaded_compare_df['sales_uploaded'].replace(0, np.nan)) * 100

            merged_df = pd.merge(db_compare_df, uploaded_compare_df, on='so_number', how='outer', suffixes=('_db', '_uploaded'), indicator=True)

            def determine_status_and_color(row):
                if row['status'] == 'HR Verified':
                    final_margin = row['final_margin']
                    if pd.isna(final_margin): return 'ยืนยันแล้ว (รอผล)'
                    if final_margin < 0: return 'ขาดทุน'
                    elif final_margin < 10: return 'กำไรน้อย'
                    else: return 'กำไรดี'
                if row['_merge'] == 'right_only': return 'มีใน Express, ไม่มีในระบบ'
                if row['_merge'] == 'left_only': return 'มีในระบบ, ไม่มีใน Express'
                
                sale_ok = row['sales_for_comparison'] >= row['sales_uploaded']
                cost_ok = row['cost_db'] >= row['cost_uploaded']
                
                if sale_ok and cost_ok: 
                    return "ผ่านเกณฑ์"
                elif not sale_ok: 
                    return "ยอดขายต่ำกว่า Express"
                elif not cost_ok: 
                    return "ต้นทุนต่ำกว่า Express"
                else: 
                    return "ข้อมูลไม่ตรงกัน"

            merged_df['สถานะ'] = merged_df.apply(determine_status_and_color, axis=1)
            merged_df['ผลต่างยอดขาย'] = merged_df['sales_service_amount'].fillna(0) - merged_df['sales_uploaded'].fillna(0)
            merged_df['ผลต่างต้นทุน'] = merged_df['cost_db'].fillna(0) - merged_df['cost_uploaded'].fillna(0)
            
            display_order_map = {
                'so_number': 'เลขที่ SO',
                'sales_for_comparison': 'ยอดขาย (ระบบ)',
                'sales_uploaded': 'ยอดขาย (Express)',
                'cost_db': 'ต้นทุน (ระบบ)',
                'cost_uploaded': 'ต้นทุน (Express)',
                'ผลต่างยอดขาย': 'ผลต่างยอดขาย',
                'ผลต่างต้นทุน': 'ผลต่างต้นทุน',
                'สถานะ': 'สถานะ',
                'sales_service_amount': 'ยอดขาย/บริการ (ระบบ)'
            }

            # --- START: แทนที่โค้ดเก่าด้วยส่วนนี้ ---
            # 1. ตรวจสอบให้แน่ใจว่าคอลัมน์ทั้งหมดที่เราต้องการมีอยู่ใน DataFrame
            for key in display_order_map.keys():
                if key not in merged_df.columns:
                    merged_df[key] = np.nan
            
            # 2. เลือกเฉพาะคอลัมน์ที่ต้องการตามลำดับ และสร้าง DataFrame ใหม่
            self.comparison_df = merged_df[list(display_order_map.keys())].copy()

            # 3. เปลี่ยนชื่อคอลัมน์เป็นภาษาไทย
            self.comparison_df.rename(columns=display_order_map, inplace=True)

            # (ลบบรรทัด self.comparison_df = merged_df[list...].copy() ที่ซ้ำซ้อนออกไปแล้ว)
            
            self.comparison_df.rename(columns=display_order_map, inplace=True)
            
            status_colors = {
                "ผ่านเกณฑ์": "#BBF7D0", 
                "ยอดขายต่ำกว่า Express": "#FECACA", 
                "ต้นทุนต่ำกว่า Express": "#FEF08A",
                "มีใน Express, ไม่มีในระบบ": "#FECACA", 
                "มีในระบบ, ไม่มีใน Express": "#FEF08A",
                "ข้อมูลไม่ตรงกัน": "#FED7AA",
                "กำไรดี": "#BBF7D0", 
                "กำไรน้อย": "#FEF08A", 
                "ขาดทุน": "#FECACA", 
                "ยืนยันแล้ว (รอผล)": "#E5E7EB"
            }
            
            self.results_frame_label.configure(text="ผลลัพธ์การเปรียบเทียบ (ดับเบิลคลิกเพื่อตรวจสอบ)")
            self._create_styled_dataframe_table(self.results_frame, self.comparison_df, "", status_column="สถานะ", status_colors=status_colors, on_row_click=self._on_tree_double_click)
            self.export_button.pack(side="right", padx=10, pady=10)
            self.finalize_button.pack(side="right", padx=10, pady=10)
            
           
            # --- END: สิ้นสุดการแก้ไข ---

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
            # --- START: จุดที่แก้ไข Query ---
            base_query = """SELECT c.*, po.cogs_db, po.po_shipping_stock, po.po_shipping_site, po.po_relocation, u.sale_name 
                        FROM commissions c 
                        JOIN sales_users u ON c.sale_key = u.sale_key
                        LEFT JOIN (
                                SELECT 
                                    so_number, 
                                    SUM(total_cost) as cogs_db,
                                    SUM(shipping_to_stock_cost) as po_shipping_stock,
                                    SUM(shipping_to_site_cost) as po_shipping_site,
                                    SUM(relocation_cost) as po_relocation
                                FROM purchase_orders 
                                WHERE status = 'Approved' 
                                GROUP BY so_number
                            ) po ON c.so_number = po.so_number
                        WHERE c.is_active = 1 AND c.status = 'Forwarded_To_HR'"""
            # --- END: สิ้นสุดการแก้ไข Query ---
            params = []

            if self.current_comparison_salesperson != "ทั้งหมด":
                base_query += " AND c.sale_key = %s"
                params.append(self.current_comparison_salesperson)
            else:
                base_query += " AND c.sale_key IN (SELECT sale_key FROM sales_users WHERE status = 'Active' AND role = 'Sale')"

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
        """ดึงข้อมูลและคำนวณค่าคอมตามงวดที่เลือก"""
        if selected_period is None:
            selected_period = self.process_period_var.get()
        
        sale_key = self.selected_sale_for_process.get()
        if not selected_period or not sale_key or "-" in selected_period:
            return

        month_name, year_be_str = selected_period.split()
        month_num = self.thai_month_map[month_name]
        year_ad = int(year_be_str) - 543

        plan_info = self.sales_user_info.get(sale_key, {})
        plan = plan_info.get('plan', 'Plan A')

        for widget in self.process_result_frame.winfo_children(): widget.destroy()
        loading = self._show_loading(self.process_result_frame)

        try:
            # (โค้ดส่วน Query ดึงข้อมูลเหมือนเดิม)
            query_comm = """
                SELECT 
                    id, so_number, sale_key, commission_month, commission_year,
                    difference_amount, final_sales_amount, final_cost_amount,
                    sales_service_amount, payment_no_vat, shipping_cost,
                    giveaways, brokerage_fee, transfer_fee,
                    separate_shipping_charge, payment_before_vat
                FROM commissions 
                WHERE sale_key = %s AND status = 'HR Verified' AND payout_id IS NULL
                AND commission_month = %s AND commission_year = %s
            """
            params = (sale_key, month_num, year_ad)
            self.current_comm_df = pd.read_sql_query(query_comm, self.pg_engine, params=params)

            self.current_total_sales = self.current_comm_df['final_sales_amount'].sum()
            self.current_total_cost = self.current_comm_df['final_cost_amount'].sum()

            if self.current_comm_df.empty:
                loading.destroy()
                CTkLabel(self.process_result_frame, text="ไม่พบข้อมูลในงวดที่เลือก").pack(pady=20)
                return
            
            # vvvvvvvvvvvvvv จุดสำคัญ vvvvvvvvvvvvvv
            # เรียกใช้ business_logic เหมือนเดิม แต่จะเก็บผลลัพธ์ไว้
            total_giveaways = 0.0
            if 'giveaways' in self.current_comm_df.columns:
                total_giveaways = self.current_comm_df['giveaways'].sum()

            self.initial_commission_result = business_logic.calculate_monthly_commission(plan, self.current_comm_df)
            
            # ถ้ามีตารางรายละเอียด (เป็น Plan A) ให้เก็บไว้
            if self.initial_commission_result.get('type') == 'summary_plan_a':
                self.commission_details_df = self.initial_commission_result.get('details')
            else:
                self.commission_details_df = None # เคลียร์ค่าถ้าไม่ใช่ Plan A
            # ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

            loading.destroy()
            self._create_hr_input_interface(auto_deduction_value=total_giveaways)

        except Exception as e:
            if loading.winfo_exists(): loading.destroy()
            traceback.print_exc()
            messagebox.showerror("Calculation Error", f"เกิดข้อผิดพลาดในการคำนวณ: {e}", parent=self)

    def _create_hr_input_interface(self, auto_deduction_value=0.0):
        """
        สร้างหน้าจอสำหรับกรอก Incentive/Deduction และแสดงผลสรุปค่าคอม
        (ฉบับแก้ไข: ใช้ .grid() เพื่อควบคุม Layout ให้แม่นยำยิ่งขึ้น)
        """
        # --- START: โค้ดที่แก้ไข ---
        # 1. ล้าง Frame เดิมและตั้งค่า Grid Layout ใหม่ให้ชัดเจน
        for widget in self.process_result_frame.winfo_children():
            widget.destroy()

        # กำหนดให้แถวที่ 1 (แถวของตารางสรุป) เป็นแถวที่จะขยายตัว
        self.process_result_frame.grid_rowconfigure(1, weight=1)
        self.process_result_frame.grid_columnconfigure(0, weight=1)

        calculated_commission = self.initial_commission_result.get('final_commission', 0.0)

        # 2. สร้าง Frame สำหรับกรอกข้อมูล (แถวบนสุด) และใช้ .grid()
        input_frame = CTkFrame(self.process_result_frame)
        input_frame.grid(row=0, column=0, pady=(10, 0), padx=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        # (โค้ดส่วนแสดงผลข้อมูลใน input_frame เหมือนเดิมทุกอย่าง)
        plan_name = self.sales_user_info.get(self.selected_sale_for_process.get(), {}).get('plan', 'N/A')
        self.plan_display_label = CTkLabel(input_frame, text=f"แผนค่าคอมมิชชั่น: {plan_name}", font=self.header_font_table, text_color=self.theme["primary"])
        self.plan_display_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        CTkLabel(input_frame, text="ยอดคอมมิชชั่นที่คำนวณได้:", font=self.label_font).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.calculated_commission_label = CTkLabel(input_frame, text=f"{calculated_commission:,.2f} บาท", font=self.header_font_table)
        self.calculated_commission_label.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        stats_frame = CTkFrame(input_frame, fg_color="transparent")
        stats_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(5,0))
        stats_frame.grid_columnconfigure((1, 3), weight=1)
        CTkLabel(stats_frame, text="ยอดขายรวม (ที่ใช้คำนวณ):", font=self.label_font, text_color="#2563EB").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        CTkLabel(stats_frame, text=f"{getattr(self, 'current_total_sales', 0.0):,.2f} บาท", font=self.entry_font).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        CTkLabel(stats_frame, text="ต้นทุนรวม (ที่ใช้คำนวณ):", font=self.label_font, text_color="#D97706").grid(row=0, column=2, padx=(20, 10), pady=5, sticky="w")
        CTkLabel(stats_frame, text=f"{getattr(self, 'current_total_cost', 0.0):,.2f} บาท", font=self.entry_font).grid(row=0, column=3, padx=10, pady=5, sticky="w")
        CTkLabel(input_frame, text="(+) Incentive:", font=self.label_font).grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.incentive_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.incentive_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        CTkLabel(input_frame, text="(-) หัก ค่าใช้จ่าย/ดำเนินการ:", font=self.label_font).grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.deduction_entry = NumericEntry(input_frame, placeholder_text="0.00")
        self.deduction_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")
        if auto_deduction_value > 0:
            self.deduction_entry.insert(0, f"{auto_deduction_value:,.2f}")
        CTkLabel(input_frame, text="หมายเหตุ/Incentive อื่นๆ:", font=self.label_font).grid(row=5, column=0, padx=10, pady=10, sticky="w")
        self.payout_notes_entry = CTkTextbox(input_frame, height=80)
        self.payout_notes_entry.grid(row=5, column=1, padx=10, pady=10, sticky="ew")
        CTkButton(input_frame, text="คำนวณขั้นสุดท้ายและแสดงสรุป", command=self._perform_final_calculation, fg_color=self.theme["primary"]).grid(row=6, column=0, columnspan=2, pady=20)

        # 3. สร้าง Frame สำหรับตารางสรุป (แถวกลาง) และใช้ .grid()
        #    - ทำให้เป็น CTkScrollableFrame เพื่อให้เลื่อนได้
        #    - sticky="nsew" เพื่อให้มันขยายเต็มพื้นที่แถวที่ 1 ที่เรากำหนด weight ไว้
        self.final_summary_frame = CTkScrollableFrame(self.process_result_frame, fg_color="transparent")
        self.final_summary_frame.grid(row=1, column=0, pady=10, padx=10, sticky="nsew")

        # 4. สร้าง Frame สำหรับปุ่มยืนยัน (แถวล่างสุด) และใช้ .grid()
        #    - แถวนี้จะ "ไม่ขยาย" และจะอยู่ที่ด้านล่างเสมอ
        bottom_action_frame = CTkFrame(self.process_result_frame, fg_color="transparent")
        bottom_action_frame.grid(row=2, column=0, pady=(0, 10), padx=10, sticky="ew")

        # 5. สร้างปุ่มยืนยันให้อยู่ใน Frame ล่างสุด
        self.confirm_payout_button = CTkButton(bottom_action_frame, text="✅ ยืนยันการจ่ายเงินและบันทึก",
                            command=self._confirm_payout_and_save,
                            fg_color="#16A34A", hover_color="#15803D",
                            font=CTkFont(size=16, weight="bold"))


    def _perform_final_calculation(self):
        try:
            incentive_val = float(self.incentive_entry.get().replace(",", "") or 0.0)
            deduction_val = float(self.deduction_entry.get().replace(",", "") or 0.0)
        except ValueError:
            messagebox.showerror("ข้อมูลผิดพลาด", "กรุณากรอก Incentive และ Deduction เป็นตัวเลข", parent=self)
            return

        incentives_dict = {"Incentive พิเศษ": incentive_val} if incentive_val > 0 else None
        deductions_dict = {"ค่าใช้จ่าย/ดำเนินการ": deduction_val} if deduction_val > 0 else None
        
        sale_key = self.selected_sale_for_process.get()
        plan = self.sales_user_info.get(sale_key, {}).get('plan', 'Plan A')

        final_result = business_logic.calculate_monthly_commission(
            plan_name=plan,
            comm_df=self.current_comm_df,
            incentives=incentives_dict,
            additional_deductions=deductions_dict
        )
        
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
        if self.final_summary_data is None or self.current_comm_df is None:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลสรุปที่จะบันทึก", parent=self)
            return
            
        sale_key = self.selected_sale_for_process.get()
        user_info = self.sales_user_info.get(sale_key, {})
        plan_name = user_info.get('plan', 'N/A')
        sales_target = user_info.get('target', 0.0)
        
        net_commission_row = self.final_summary_data[self.final_summary_data['description'].str.contains("หลังหัก ณ ที่จ่าย")]
        net_commission_value = net_commission_row['value'].iloc[0] if not net_commission_row.empty else 0.0

        msg = (f"คุณต้องการยืนยันการจ่ายค่าคอมมิชชั่นสำหรับ '{sale_key}' ใช่หรือไม่?\n\n"
            f"ยอดสุทธิ: {net_commission_value:,.2f} บาท\n\n"
            f"การกระทำนี้จะบันทึก Log และอัปเดตสถานะรายการทั้งหมดเป็น 'Paid'")

        if not messagebox.askyesno("ยืนยันการจ่ายเงิน", msg, parent=self):
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 1. เตรียมข้อมูลสำหรับ Log
                so_numbers_list = self.current_comm_df['so_number'].tolist()
                so_numbers_json = json.dumps(so_numbers_list)
                summary_json = self.final_summary_data.to_json(orient='records')
                payout_notes = self.payout_notes_entry.get("1.0", "end-1c").strip()

                total_sales = float(self.current_comm_df['final_sales_amount'].sum())
                total_cost = float(self.current_comm_df['final_cost_amount'].sum())

                # 2. บันทึก Log ลงตารางใหม่
                cursor.execute("""
                    INSERT INTO commission_payout_logs 
                    (hr_user_key, sale_key, plan_name, so_numbers_json, summary_json, notes, total_sales_for_margin, total_cost_for_margin, sales_target_at_payout) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    RETURNING id
                """, (self.user_key, sale_key, plan_name, so_numbers_json, summary_json, payout_notes, total_sales, total_cost, sales_target))
            
                # --- START: แก้ไขการจัดย่อหน้าตรงนี้ ---
                # ย้ายโค้ดส่วนนี้ทั้งหมดเข้ามาอยู่ในบล็อก with conn.cursor()
                
                # รับ ID ของประวัติการจ่ายเงิน (payout_id) ที่เพิ่งสร้าง
                new_payout_id = cursor.fetchone()[0]

                # 3. อัปเดตสถานะ และ "ประทับตรา" payout_id ลงบน SO ที่เกี่ยวข้อง
                record_ids_to_update = tuple(self.current_comm_df['id'].tolist())
                if record_ids_to_update:
                    cursor.execute("""
                        UPDATE commissions 
                        SET status = 'Paid', payout_id = %s
                        WHERE id IN %s
                    """, (new_payout_id, record_ids_to_update))
                
                # --- END: สิ้นสุดการแก้ไข ---
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกการจ่ายค่าคอมมิชชั่นเรียบร้อยแล้ว", parent=self)
            
            # 4. รีเฟรชหน้าจอ
            self._on_sale_selected_for_process()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึก: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)
    
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

    def _create_styled_dataframe_table(self, parent, df, label_text="", on_row_click=None, status_colors=None, status_column=None):
        for widget in parent.winfo_children():
            widget.destroy()

        if df is None or df.empty:
            CTkLabel(parent, text=f"ไม่พบข้อมูลสำหรับ '{label_text}'").pack(pady=20)
            return
        
        container = CTkFrame(parent, fg_color="transparent")
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        if label_text:
            CTkLabel(container, text=label_text, font=self.header_font_table).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        tree_frame = CTkFrame(container, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = df.columns.tolist()
        
        # --- START: โค้ดส่วนสไตล์ที่ปรับปรุงใหม่ ---
        style = ttk.Style(self)
        style.theme_use("clam") # <--- 1. ใช้ Theme 'clam' เพื่อให้ปรับแต่งได้เต็มที่

        # 2. ตั้งค่าสไตล์สำหรับ "หัวตาราง" (Header)
        style.configure("Modern.Treeview.Heading", 
                        font=self.header_font_table, 
                        background=self.theme.get("header", "#1E40AF"), # สีพื้นหลังหัวตาราง
                        foreground="white",                             # สีตัวอักษร
                        relief="flat",                                  # ทำให้ขอบแบน
                        padding=(10, 10))                               # เพิ่มช่องว่างภายในซ้ายขวาและบนล่าง
        style.map("Modern.Treeview.Heading",
                  background=[('active', self.theme.get("primary", "#3B82F6"))]) # สีตอนเอาเมาส์ไปชี้

        # 3. ตั้งค่าสไตล์สำหรับ "แถวข้อมูล" (Rows)
        style.configure("Modern.Treeview", 
                        rowheight=32,                                   # <--- เพิ่มความสูงของแถวให้อ่านง่าย
                        font=self.entry_font,
                        background="#FFFFFF",                           # สีพื้นหลังปกติ
                        fieldbackground="#FFFFFF",
                        foreground="#111827")                           # สีตัวอักษรปกติ
        
        style.map("Modern.Treeview",
                  background=[('selected', self.theme.get("primary", "#3B82F6"))], # สีตอนที่เลือกแถว
                  foreground=[('selected', "white")])
        # --- END: สิ้นสุดโค้ดส่วนสไตล์ ---

        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Modern.Treeview") # <--- 4. เรียกใช้สไตล์ใหม่
        tree.grid(row=0, column=0, sticky="nsew")
        
        # กำหนดสีพื้นหลังของแถวตามสถานะ (โค้ดส่วนนี้ทำงานเหมือนเดิม)
        if status_colors:
            for tag_name, color in status_colors.items():
                tree.tag_configure(tag_name, background=color)

        # ตั้งค่าคอลัมน์และการจัดวางข้อมูล
        for col_id in columns:
            tree.heading(col_id, text=col_id, anchor='center') # <--- จัดให้หัวข้ออยู่กึ่งกลาง
            width = 150 
            anchor = 'w' # การจัดวางข้อมูลปกติ (ชิดซ้าย)
            
            # จัดข้อมูลที่เป็นตัวเลขให้ชิดขวา
            if any(s in col_id for s in ['ยอด', 'ต้นทุน', 'ผลต่าง', 'Margin']): 
                width = 140
                anchor = 'e'
            elif any(s in col_id for s in ['ID', 'Key']): 
                width = 80
                anchor = 'center'
            elif 'สถานะ' in col_id:
                width = 200
            
            tree.column(col_id, width=width, anchor=anchor)

        # เพิ่มข้อมูลลงในตาราง (โค้ดส่วนนี้ทำงานเหมือนเดิม)
        for index, row in df.iterrows():
            tags_tuple = ()
            if status_colors and status_column and status_column in df.columns:
                status_val = str(row.get(status_column, ''))
                if status_val in status_colors:
                    tags_tuple = (status_val,)

            values = []
            for col_name in columns:
                value = row[col_name]
                if pd.notna(value):
                    # --- START: เพิ่มเงื่อนไขนี้เข้ามา ---
                    if isinstance(value, (datetime, pd.Timestamp)):
                        values.append(value.strftime('%d/%m/%Y %H:%M'))
                    # --- END: สิ้นสุดส่วนที่เพิ่ม ---
                    elif isinstance(value, (float, np.floating)):
                        values.append(f"{value:,.2f}")
                    else:
                        values.append(str(value))
                else:
                    values.append("")

            row_id_for_iid = row.get('เลขที่ SO', str(index))
            tree.insert("", "end", values=values, tags=tags_tuple, iid=str(row_id_for_iid))
        
        # สร้าง Scrollbars (โค้ดส่วนนี้ทำงานเหมือนเดิม)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
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