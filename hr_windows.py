import tkinter as tk
from tkinter import ttk, filedialog, TclError 
from customtkinter import (CTkToplevel, CTkTextbox, CTkScrollableFrame, CTkLabel, CTkFont, CTkFrame, CTkButton, CTkEntry, CTkRadioButton, CTkOptionMenu, CTkInputDialog)
from tkinter import messagebox
import json
import customtkinter as ctk 
import pandas as pd
from datetime import datetime
import traceback
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
import utils
from export_utils import export_commission_details_to_excel, export_payout_so_list_to_excel
import psycopg2.errors
import psycopg2.extras
import numpy as np
from utils import RejectionReasonDialog
from sqlalchemy import create_engine
from history_windows import SOPopupWindow
from history_windows import print_transport_pdf_wrapper



class SalesDataViewerWindow(CTkToplevel):
    def __init__(self, master, app_container, so_number):
        super().__init__(master)
        self.app_container = app_container
        self.so_number = so_number
        self.so_data = None

        self.sale_theme = self.app_container.THEME.get("sale", {"primary": "#3B82F6", "header": "#1E40AF"})
        self.title(f"รายละเอียดข้อมูล SO: {self.so_number}")
        self.geometry("700x800")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkLabel(header_frame, text=f"ข้อมูลสำหรับ SO Number: {self.so_number}", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(header_frame, text="ปิด", command=self.destroy, width=80).pack(side="right")

        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.load_data()
        self.create_widgets()
        
        self.transient(master)
        self.grab_set()

    def load_data(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1", (self.so_number,))
                self.so_data = cursor.fetchone()
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล SO ได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_section_frame(self, parent, title):
        frame = ctk.CTkFrame(parent, corner_radius=10, border_width=1, border_color=self.sale_theme['primary'])
        frame.pack(fill="x", pady=(10, 5), padx=5)
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=18, weight="bold"), text_color=self.sale_theme["header"]).grid(
            row=0, column=0, columnspan=2, padx=15, pady=(10, 5), sticky="w")
        return frame

    def _add_display_row(self, parent, row_index, label_text, value):
        if value is None or pd.isna(value):
            value_text = "-"
        elif isinstance(value, (int, float, np.floating)):
            value_text = f"{value:,.2f}"
        elif isinstance(value, (datetime, pd.Timestamp)):
            value_text = value.strftime('%Y-%m-%d')
        else:
            value_text = str(value)
            
        ctk.CTkLabel(parent, text=label_text, font=ctk.CTkFont(size=14)).grid(
            row=row_index, column=0, padx=(15, 10), pady=4, sticky="w")
        ctk.CTkLabel(parent, text=value_text, font=ctk.CTkFont(size=14), wraplength=400, justify="left").grid(
            row=row_index, column=1, padx=(10, 15), pady=4, sticky="ew")

    def create_widgets(self):
        if not self.so_data:
            ctk.CTkLabel(self.main_frame, text="ไม่พบข้อมูล").pack(pady=20)
            return
        
        header_map = self.app_container.HEADER_MAP

        # --- Section 1: รายละเอียดการขาย ---
        f1 = self._create_section_frame(self.main_frame, "รายละเอียดการขาย")
        self._add_display_row(f1, 1, header_map.get('so_number', 'เลขที่ SO'), self.so_data.get('so_number'))
        self._add_display_row(f1, 2, header_map.get('bill_date', 'วันที่เปิด SO'), self.so_data.get('bill_date'))
        self._add_display_row(f1, 3, header_map.get('customer_id', 'รหัสลูกค้า'), self.so_data.get('customer_id'))
        self._add_display_row(f1, 4, header_map.get('customer_name', 'ชื่อลูกค้า'), self.so_data.get('customer_name'))
        self._add_display_row(f1, 5, header_map.get('credit_term', 'เครดิต'), self.so_data.get('credit_term'))
        self._add_display_row(f1, 6, header_map.get('commission_month', 'เดือนคอมมิชชั่น'), self.so_data.get('commission_month'))
        self._add_display_row(f1, 7, header_map.get('commission_year', 'ปีคอมมิชชั่น'), self.so_data.get('commission_year') + 543 if self.so_data.get('commission_year') else None)

        # --- Section 2: ยอดขายและบริการ ---
        f2 = self._create_section_frame(self.main_frame, "ยอดขายและบริการ")
        self._add_display_row(f2, 1, f"{header_map.get('sales_service_amount', 'ยอดขาย/บริการ')} ({self.so_data.get('sales_service_vat_option')})", self.so_data.get('sales_service_amount'))
        self._add_display_row(f2, 2, f"{header_map.get('cutting_drilling_fee', 'ค่าบริการตัด/เจาะ')} ({self.so_data.get('cutting_drilling_fee_vat_option')})", self.so_data.get('cutting_drilling_fee'))
        self._add_display_row(f2, 3, f"{header_map.get('other_service_fee', 'ค่าบริการอื่นๆ')} ({self.so_data.get('other_service_fee_vat_option')})", self.so_data.get('other_service_fee'))

        # --- Section 3: ค่าจัดส่ง ---
        f3 = self._create_section_frame(self.main_frame, "ค่าจัดส่ง")
        self._add_display_row(f3, 1, f"{header_map.get('shipping_cost', 'ค่าขนส่ง')} ({self.so_data.get('shipping_vat_option')})", self.so_data.get('shipping_cost'))
        self._add_display_row(f3, 2, header_map.get('delivery_date', 'วันที่จัดส่ง'), self.so_data.get('delivery_date'))
        
        # +++ START: เพิ่ม Section ที่ขาดไป +++
        
        # --- Section 4: Delivery Note ---
        f4 = self._create_section_frame(self.main_frame, "Delivery Note")
        self._add_display_row(f4, 1, header_map.get('delivery_type', 'การจัดส่ง'), self.so_data.get('delivery_type'))
        self._add_display_row(f4, 2, header_map.get('pickup_location', 'Location เข้ารับ'), self.so_data.get('pickup_location'))
        self._add_display_row(f4, 3, header_map.get('relocation_cost', 'ค่าย้าย'), self.so_data.get('relocation_cost'))
        self._add_display_row(f4, 4, header_map.get('date_to_warehouse', 'วันที่เข้าคลัง'), self.so_data.get('date_to_warehouse'))
        self._add_display_row(f4, 5, header_map.get('date_to_customer', 'วันที่ส่งลูกค้า'), self.so_data.get('date_to_customer'))
        self._add_display_row(f4, 6, header_map.get('pickup_registration', 'ทะเบียนเข้ารับ'), self.so_data.get('pickup_registration'))

        # --- Section 5: ค่าธรรมเนียมและส่วนลด ---
        f5 = self._create_section_frame(self.main_frame, "ค่าธรรมเนียมและส่วนลด")
        self._add_display_row(f5, 1, f"{header_map.get('credit_card_fee', 'ค่าธรรมเนียมบัตร')} ({self.so_data.get('credit_card_fee_vat_option')})", self.so_data.get('credit_card_fee'))
        self._add_display_row(f5, 2, header_map.get('transfer_fee', 'ค่าธรรมเนียมโอน'), self.so_data.get('transfer_fee'))
        self._add_display_row(f5, 3, header_map.get('wht_3_percent', 'หัก ณ ที่จ่าย 3%'), self.so_data.get('wht_3_percent'))
        self._add_display_row(f5, 4, header_map.get('brokerage_fee', 'ค่านายหน้า'), self.so_data.get('brokerage_fee'))
        self._add_display_row(f5, 5, header_map.get('giveaways', 'ของแถม'), self.so_data.get('giveaways'))
        self._add_display_row(f5, 6, header_map.get('coupons', 'คูปอง'), self.so_data.get('coupons'))

        # --- Section 6: รายละเอียดการชำระเงิน ---
        f6 = self._create_section_frame(self.main_frame, "รายละเอียดการชำระเงิน")
        self._add_display_row(f6, 1, header_map.get('total_payment_amount', 'ยอดชำระรวม'), self.so_data.get('total_payment_amount'))
        self._add_display_row(f6, 2, header_map.get('payment_date', 'วันที่ชำระ'), self.so_data.get('payment_date'))
        self._add_display_row(f6, 3, header_map.get('payment_before_vat', 'ชำระก่อน VAT'), self.so_data.get('payment_before_vat'))
        self._add_display_row(f6, 4, header_map.get('payment_no_vat', 'ชำระ NV'), self.so_data.get('payment_no_vat'))


    def load_and_populate_data(self):
        self._show_loading()
        try:
            query = "SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1"
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.so_number,))
            if df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูลสำหรับ SO Number: {self.so_number}", parent=self)
                self.after(100, self.destroy)
                return
            
            data = df.iloc[0]

            for key, (parent, label, row, col_name) in self.fields.items():
                entry_widget = getattr(self, f"{key}_entry")
                value = data.get(col_name)
                display_text = ""
                if pd.notna(value):
                    if isinstance(value, (int, float)):
                        display_text = f"{value:,.2f}"
                    else:
                        display_text = str(value)

                entry_widget.configure(state="normal")
                entry_widget.insert(0, display_text)
                entry_widget.configure(state="readonly")
        
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล SO ได้: {e}", parent=self)
            self.after(100, self.destroy)
        finally:
            self._hide_loading()
# +++ END: เพิ่ม Class ใหม่ +++

# hr_windows.py (คลาส HRVerificationWindow ที่แก้ไขสมบูรณ์)

class HRVerificationWindow(CTkToplevel):

    def __init__(self, master, app_container, system_data, excel_data, po_data, refresh_callback=None):
        super().__init__(master)
        self.master = master
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.system_data = system_data
        self.excel_data = excel_data
        self.po_data = po_data
        self.refresh_callback = refresh_callback
        self.so_number = self.system_data.get('so_number', 'N/A')
        self.record_id = self.system_data.get('id')
        self.cost_multiplier_var = tk.StringVar(value="1.03") # สร้างตัวแปรพร้อมค่าเริ่มต้น

        # --- ตัวแปรสำหรับเก็บค่าที่ User Override และค่าที่คำนวณแล้ว ---
        self.cost_overrides = {} # เก็บค่า cost ที่ HR อาจแก้ไขเอง
        self.calculated_values = {} # เก็บผลรวมที่คำนวณล่าสุด
        self.final_sale_source = tk.StringVar(value="system")
        self.final_cost_source = tk.StringVar(value="system")
        self.final_sale_source.trace_add("write", self._update_selection_display)
        self.final_cost_source.trace_add("write", self._update_selection_display)
        self.final_sale_source = tk.StringVar(value="system")
        self.final_cost_source = tk.StringVar(value="system")
        self.final_sale_source.trace_add("write", self._update_selection_display)
        self.final_cost_source.trace_add("write", self._update_selection_display)
        self.final_sale_source = tk.StringVar(value="system")
        self.final_cost_source = tk.StringVar(value="system")
        self.final_sale_source.trace_add("write", self._update_selection_display)
        self.final_cost_source.trace_add("write", self._update_selection_display)

        self._so_create_string_vars() # สร้าง StringVars สำหรับหน้าต่างแก้ไข SO
        
        

        self.title(f"สรุปข้อมูล SO: {self.so_number}")
        self.geometry("950x750")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- UI Layout ใหม่ทั้งหมด ---
        self._create_new_ui_layout()

        initial_multiplier = self.system_data.get('cost_multiplier')
        if initial_multiplier and f"{initial_multiplier:.2f}" in ["1.01", "1.02", "1.03", "1.04", "1.05"]:
            self.cost_multiplier_var.set(f"{initial_multiplier:.2f}")
        else:
            self.cost_multiplier_var.set("1.03") # ถ้าไม่มีข้อมูล ให้ใช้ค่าเริ่มต้น

        # --- โหลดและคำนวณข้อมูล ---
        self.after(50, self._update_all_calculations_and_ui)
        
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.transient(master)
        self.grab_set()
     
    def _create_new_ui_layout(self):
        """สร้าง UI Layout ใหม่ทั้งหมดสำหรับหน้าต่างนี้"""
        # --- Header ---
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 10))
        ctk.CTkLabel(header_frame, text=f"SO Number: {self.so_number}", font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")
        
        detail_button_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        detail_button_frame.pack(side="right")
        ctk.CTkButton(detail_button_frame, text="ดูข้อมูล SO", command=self._view_so_data).pack(side="left", padx=(0, 5))
        ctk.CTkButton(detail_button_frame, text="✏️ แก้ไขข้อมูล SO", command=self._open_so_editor_popup).pack(side="left", padx=(5, 0))

        # --- Main Scrollable Frame (สำหรับเนื้อหาทั้งหมด) ---
        scroll_frame = ctk.CTkScrollableFrame(self, fg_color="#F0F2F5")
        scroll_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        scroll_frame.grid_columnconfigure((0, 1), weight=1)

        sales_card = CTkFrame(scroll_frame, corner_radius=10)
        sales_card.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self._create_summary_card(sales_card, "ยอดขายรวมสุดท้าย (Final Sales)", "sales")

        cost_card = CTkFrame(scroll_frame, corner_radius=10)
        cost_card.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self._create_summary_card(cost_card, "ยอดต้นทุนรวมสุดท้าย (Final Cost)", "cost")
        # +++ END +++
        
        # +++ START: สร้างส่วนให้ HR เลือกแหล่งข้อมูล +++
        self._create_final_summary_section(scroll_frame) # <--- เรียกฟังก์ชันใหม่
        # +++ END +++

        self.po_container_frame = CTkFrame(scroll_frame, fg_color="transparent")

        # การ์ดสรุปยอดขาย
        sales_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        sales_card.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self._create_summary_card(sales_card, "ยอดขายรวมสุดท้าย (Final Sales)", "sales")

        # การ์ดสรุปต้นทุน
        cost_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        cost_card.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self._create_summary_card(cost_card, "ยอดต้นทุนรวมสุดท้าย (Final Cost)", "cost")
        
        # ส่วนเลือกแหล่งข้อมูล (Radio buttons)
        self._create_final_summary_section(scroll_frame)
        
        # ส่วนแสดงรายการ PO
        self.po_container_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        self.po_container_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
        self.po_container_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.po_container_frame, text="ใบสั่งซื้อ (PO) ที่เกี่ยวข้อง", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 5))

        # --- Action Buttons Frame (ย้ายมาไว้ด้านล่างสุด นอก ScrollFrame) ---
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(10, 15))
        action_frame.grid_columnconfigure((0, 1, 2, 3), weight=1) # ทำให้ปุ่มขยายเต็มพื้นที่เท่าๆ กัน

        ctk.CTkButton(action_frame, text="ตีกลับให้ฝ่ายขาย (Reject)", height=40, fg_color="#D97706", hover_color="#B45309", command=self._reject_to_salesperson).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(action_frame, text="เลื่อนไปเดือนถัดไป (Defer)", height=40, fg_color="#64748B", hover_color="#475569", command=self._defer_so).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(action_frame, text="บันทึกการแก้ไข", height=40, fg_color="#3B82F6", hover_color="#2563EB", command=self._save_intermediate_changes).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(action_frame, text="ยืนยันข้อมูลถูกต้อง (Verify)", height=40, fg_color="#16A34A", hover_color="#15803D", command=self._verify_and_save_data).grid(row=0, column=3, padx=5, pady=5, sticky="ew")

    
    def _create_comparison_table(self, parent_frame):
        """สร้างตารางเปรียบเทียบแบบ Excel-light (มี zebra striping + row divider)"""
        columns = (
            "so_number", "system_sale", "express_sale",
            "system_cost", "express_cost",
            "diff_sale", "diff_cost", "status"
        )

        # --- สร้าง Style สำหรับ Treeview ---
        style = ttk.Style()
        style.theme_use("default")

        style.configure(
            "Treeview",
            background="white",
            foreground="black",
            rowheight=25,
            fieldbackground="white",
            borderwidth=1
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#4CAF50",
            foreground="white"
        )

        style.map("Treeview", background=[("selected", "#90CAF9")])

        # --- สร้าง Treeview ---
        tree = ttk.Treeview(
            parent_frame,
            columns=columns,
            show="headings",
            height=20
        )

        headers = [
            "เลขที่ SO", "ยอดขาย (ระบบ)", "ยอดขาย (Express)",
            "ต้นทุน (ระบบ)", "ต้นทุน (Express)",
            "ผลต่างยอดขาย", "ผลต่างต้นทุน", "สถานะ"
        ]

        for col, header in zip(columns, headers):
            tree.heading(col, text=header)
            anchor = "center" if col in ["so_number", "status"] else "e"
            tree.column(col, width=120, anchor=anchor, stretch=True)

        # --- Scrollbar ---
        vsb = ttk.Scrollbar(parent_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        parent_frame.grid_rowconfigure(0, weight=1)
        parent_frame.grid_columnconfigure(0, weight=1)

        # --- ตั้งค่า zebra striping ---
        tree.tag_configure("oddrow", background="white")
        tree.tag_configure("evenrow", background="#F2F2F2")

        # --- ตัวอย่างข้อมูล (จริง ๆ คุณจะ insert จาก system/excel data) ---
        example_data = [
            ("SO6809AM005", "6,840.00", "6,840.00", "4,000.00", "3,738.32", "0.00", "261.68", "ผ่านการตรวจ"),
            ("SO6809AM010", "17,415.00", "19,315.00", "16,167.17", "12,124.65", "-1,900.00", "4,042.52", "ผ่านการตรวจ"),
        ]

        for i, row in enumerate(example_data):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            tree.insert("", "end", values=row, tags=(tag,))

        self.comparison_tree = tree

    
    
    def _create_summary_card(self, parent, title, card_type):
        """Helper function สำหรับสร้างการ์ดสรุป"""
        parent.grid_columnconfigure(0, weight=1)
        CTkLabel(parent, text=title, font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(10, 0), sticky="w")
        
        value_label = CTkLabel(parent, text="0.00", font=CTkFont(size=32, weight="bold"))
        value_label.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="w")

        source_label = CTkLabel(parent, text="Source: System", font=CTkFont(size=12), text_color="gray50")
        source_label.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="w")

        if card_type == "sales":
            self.final_sales_label = value_label
            self.final_sales_source_label = source_label
        else: # cost
            self.final_cost_label = value_label
            self.final_cost_source_label = source_label 

            # --- เพิ่มโค้ดส่วนนี้เข้าไป ---
            CTkLabel(parent, text="ตัวคูณต้นทุน (Cost Multiplier):", font=CTkFont(size=12)).grid(row=3, column=0, padx=(15, 5), pady=(10, 0), sticky="w")
            multiplier_options = ["1.01", "1.02", "1.03", "1.04", "1.05"]

            # self.cost_multiplier_var ถูกสร้างใน __init__ แล้ว
            self.cost_multiplier_menu = CTkOptionMenu(parent, variable=self.cost_multiplier_var, values=multiplier_options)
            self.cost_multiplier_menu.grid(row=4, column=0, padx=15, pady=(0, 10), sticky="w")

    # (วางฟังก์ชันนี้ต่อจาก _create_summary_card)
    

    def _populate_po_cards(self):
        """สร้างการ์ดสำหรับ PO แต่ละใบ (ฉบับแก้ไข: ใช้หน้าต่าง Master แบบเดียวกับ HR)"""
        # ล้างการ์ดเก่า
        for widget in self.po_container_frame.winfo_children():
            if isinstance(widget, ctk.CTkLabel) and "ใบสั่งซื้อ" in widget.cget("text"):
                continue
            widget.destroy()

        if self.po_data.empty:
            ctk.CTkLabel(self.po_container_frame, text="ไม่พบข้อมูล PO ที่เกี่ยวข้อง").pack(pady=10)
            return

        for index, row in self.po_data.iterrows():
            po_card = ctk.CTkFrame(self.po_container_frame, border_width=1, corner_radius=8)
            po_card.pack(fill="x", expand=True, padx=0, pady=4)
            po_card.grid_columnconfigure(0, weight=1)

            info_frame = ctk.CTkFrame(po_card, fg_color="transparent")
            info_frame.grid(row=0, column=0, sticky="w", padx=10, pady=5)
            
            grand_total = row.get('grand_total', 0) or 0
            info_text = f"PO: {row['po_number']}  |  Supplier: {row['supplier_name']}  |  ยอดรวม: {grand_total:,.2f} บาท"

            ctk.CTkLabel(info_frame, text=info_text).pack(anchor="w")
            
            status_color = "#16A34A" if row['status'] == 'Approved' else 'gray'
            ctk.CTkLabel(info_frame, text=f"สถานะ: {row['status']}", text_color=status_color).pack(anchor="w")
                
            action_frame = ctk.CTkFrame(po_card, fg_color="transparent")
            action_frame.grid(row=0, column=1, padx=10, pady=5, sticky="e")

            # 🔥 จุดแก้ไข: เรียก PurchaseDetailWindow แทน (หน้าต่าง Master ของ HR/PU)
            from history_windows import PurchaseDetailWindow
            edit_button = ctk.CTkButton(
                action_frame, 
                text="ดูรายละเอียด / แก้ไข", 
                width=150,
                command=lambda po_id=row['id']: PurchaseDetailWindow(
                    self.master, 
                    self.app_container, 
                    int(po_id), 
                    on_save_callback=self._update_all_calculations_and_ui
                )
            )
            edit_button.pack(pady=2, padx=2)

    def _create_summary_card(self, parent, title, card_type):
        """Helper function สำหรับสร้างการ์ดสรุป"""
        parent.grid_columnconfigure(0, weight=1)
        CTkLabel(parent, text=title, font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(10, 0), sticky="w")
        
        value_label = CTkLabel(parent, text="0.00", font=CTkFont(size=32, weight="bold"))
        value_label.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="w")

        source_label = CTkLabel(parent, text="Source: System", font=CTkFont(size=12), text_color="gray50")
        source_label.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="w")

        if card_type == "sales":
            self.final_sales_label = value_label
            self.final_sales_source_label = source_label
        else: # cost
            self.final_cost_label = value_label
            self.final_cost_source_label = source_label 

            # --- START: เพิ่มโค้ดส่วนนี้เข้าไป ---
            CTkLabel(parent, text="ตัวคูณต้นทุน (Cost Multiplier):", font=CTkFont(size=12)).grid(row=3, column=0, padx=(15, 5), pady=(10, 0), sticky="w")
            multiplier_options = ["1.01", "1.02", "1.03", "1.04", "1.05"]

            # self.cost_multiplier_var ถูกสร้างใน __init__ แล้ว
            self.cost_multiplier_menu = CTkOptionMenu(parent, variable=self.cost_multiplier_var, values=multiplier_options)
            self.cost_multiplier_menu.grid(row=4, column=0, padx=15, pady=(0, 10), sticky="w")

    def _reload_data(self):
        """(เวอร์ชันแก้ไข) ดึงข้อมูล SO และ PO ที่ Active ล่าสุดจากฐานข้อมูล"""
        try:
            # --- START: แก้ไข Query ตรงนี้ ---
            # เปลี่ยนจากการค้นหาด้วย ID เก่า (self.record_id)
            # มาเป็นการค้นหาด้วย SO Number และเลือกเฉพาะรายการที่ is_active = 1
            so_query = """
                SELECT c.*, u.sale_name 
                FROM commissions c 
                JOIN sales_users u ON c.sale_key = u.sale_key 
                WHERE c.so_number = %s AND c.is_active = 1 
                ORDER BY c.id DESC LIMIT 1
            """
            so_df = pd.read_sql_query(so_query, self.pg_engine, params=(self.so_number,))
            # --- END ---

            if not so_df.empty:
                # อัปเดตข้อมูล system_data และ record_id ให้เป็นของใหม่ล่าสุด
                self.system_data = so_df.iloc[0].to_dict()
                self.record_id = self.system_data['id']

            # ดึงข้อมูล POs (เหมือนเดิม)
            self.po_data = pd.read_sql_query("SELECT * FROM purchase_orders WHERE so_number = %s ORDER BY id", self.pg_engine, params=(self.so_number,))
            print(f"Data reloaded for SO {self.so_number}. Found {len(self.po_data)} POs.")

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถรีโหลดข้อมูลได้: {e}", parent=self)
            self.system_data = {}
            self.po_data = pd.DataFrame()

    def _so_create_string_vars(self):
        """สร้าง StringVars ที่จำเป็นสำหรับ SOPopupWindow"""
        self.so_form_widgets = {}
        self.sale_theme = self.app_container.THEME["sale"]
        self.dropdown_style = {
            "fg_color": "white", "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }
        self.so_form_widgets['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        self.so_form_widgets['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_form_widgets['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_verification_result_var'] = tk.StringVar(value="-")
        
        # --- START: เพิ่ม 5 บรรทัดนี้เข้าไป ---
        # สร้าง StringVars สำหรับแสดง VAT ของแต่ละรายการที่ขาดไป
        self.so_form_widgets['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")

    # hr_windows.py (ภายในคลาส HRVerificationWindow)

    def _view_so_data(self):
        if self.system_data and self.system_data.get('so_number'):
            # <<< แก้ไข: เรียกใช้ SODetailViewer ที่ถูกต้อง >>>
            SODetailViewer(master=self, app_container=self.app_container, so_number=self.system_data['so_number'])
        else:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบ SO Number สำหรับแสดงข้อมูล", parent=self)

    def _view_po_data(self):
        if not self.po_data.empty:
            # ดึงค่า ID จากแถวแรก
            po_id_value = self.po_data.iloc[0].get('id')

            # ตรวจสอบว่ามีค่า ID จริงๆ
            if po_id_value is not None:
                # <<< START: แก้ไขจุดนี้ >>>
                # แปลงค่า numpy.int64 ให้เป็น int ปกติก่อนส่งไปใช้งาน
                first_po_id = int(po_id_value)
                self.app_container.show_purchase_detail_window(first_po_id)
                # <<< END: สิ้นสุดการแก้ไข >>>
            else:
                messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบ ID ของ PO ในข้อมูลที่แสดง", parent=self)
        else:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูล PO ที่เกี่ยวข้อง", parent=self)

    def _edit_data(self):
        if self.system_data and self.system_data.get('id_db'):
            record_id_to_load = self.system_data.get('id_db')
            # ต้องสร้าง DataFrame ชั่วคราวเพื่อให้เข้ากับ Format ที่ EditCommissionWindow คาดหวัง
            data_for_edit = pd.DataFrame([self.system_data])
            row_to_edit = data_for_edit.iloc[0]
            
            # ต้องมี refresh_callback ที่ส่งมาจาก hr_screen
            self.app_container.show_edit_commission_window(row_to_edit, self.refresh_callback, user_role="HR")
            self.destroy() # ปิดหน้าต่างปัจจุบันหลังเปิดหน้าแก้ไข
        else:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่สามารถหาข้อมูลสำหรับแก้ไขได้", parent=self)

    def _create_revenue_table(self):
        revenue_keys = ['sales_service_amount', 'shipping_cost', 'cutting_drilling_fee', 'other_service_fee', 'credit_card_fee']
        revenue_headers = self.app_container.HEADER_MAP

        # --- Table Header ---
        headers = ["หัวข้อ", "ข้อมูลในระบบ", "ข้อมูลจาก Express"]
        for col, text in enumerate(headers):
            header_cell = ctk.CTkFrame(self.revenue_table_frame, border_width=1, corner_radius=0)
            header_cell.grid(row=0, column=col, sticky="nsew")
            ctk.CTkLabel(header_cell, text=text, font=ctk.CTkFont(weight="bold")).pack(padx=5, pady=5)

        # --- Table Rows ---
        for i, key in enumerate(revenue_keys, start=1):
            header = revenue_headers.get(key, key)
            system_val = self.system_data.get(key, 0)
            excel_val = self.excel_data.get(key, 'N/A')

            row_values = [
                header,
                f"{system_val:,.2f}",
                excel_val if isinstance(excel_val, str) else f"{excel_val:,.2f}"
            ]

            for col, value in enumerate(row_values):
                bg_color = "#F9FAFB" if i % 2 == 0 else "white"  # zebra stripe
                cell = ctk.CTkFrame(self.revenue_table_frame, border_width=1, fg_color=bg_color, corner_radius=0)
                cell.grid(row=i, column=col, sticky="nsew")
                ctk.CTkLabel(cell, text=value, anchor="w").pack(padx=5, pady=5)

    def _create_cost_table(self):
        cost_keys = ['final_cost_amount', 'giveaways', 'brokerage_fee', 'wht_3_percent', 'transfer_fee']
        cost_headers = self.app_container.HEADER_MAP

        # Header
        CTkLabel(self.cost_table_frame, text="หัวข้อ", font=CTkFont(weight="bold")).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        CTkLabel(self.cost_table_frame, text="ข้อมูลในระบบ", font=CTkFont(weight="bold")).grid(row=1, column=1, padx=5, pady=2, sticky="e")
        CTkLabel(self.cost_table_frame, text="ข้อมูลจาก Express", font=CTkFont(weight="bold")).grid(row=1, column=2, padx=5, pady=2, sticky="e")

        # Data Rows
        for i, key in enumerate(cost_keys):
            header = cost_headers.get(key, key)
            system_val = self.system_data.get(key) # ดึงค่ามาก่อน
            excel_val = self.excel_data.get(key, 0)

            # --- START: เพิ่มการตรวจสอบค่า None ---
            system_text = f"{system_val:,.2f}" if system_val is not None else "0.00"
            # --- END: สิ้นสุดการแก้ไข ---
            
            CTkLabel(self.cost_table_frame, text=header).grid(row=i+2, column=0, padx=5, pady=2, sticky="w")
            CTkLabel(self.cost_table_frame, text=system_text).grid(row=i+2, column=1, padx=5, pady=2, sticky="e")
            
            val_for_comparison = system_val if system_val is not None else 0
            color = "red" if val_for_comparison < excel_val else "green"
            CTkLabel(self.cost_table_frame, text=f"{excel_val:,.2f}", text_color=color).grid(row=i+2, column=2, padx=5, pady=2, sticky="e")
    
    def _create_section_frame(self, parent, title, col):
        frame = CTkFrame(parent, corner_radius=10, border_width=1)
        frame.grid(row=0, column=col, sticky="nsew", padx=5, pady=5)
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        label = CTkLabel(frame, text=title, font=CTkFont(size=14, weight="bold"))
        label.grid(row=0, column=0, padx=10, pady=(5, 10), sticky="w")
        
        return frame
    
    
    def _on_verify(self):
        conn = None
        try:
            # --- Determine Final Values ---
            if self.sales_choice_var.get() == 'system':
                # <<< จุดแก้ไข: เปลี่ยนจากการดึงค่าบนหน้าจอ มาเป็นค่าจากข้อมูลดิบ 'system_data'
                final_sale = utils.convert_to_float(self.system_data.get('sales_service_amount', 0))
            else: # 'excel'
                final_sale = utils.convert_to_float(self.excel_data.get('sales_uploaded', 0))

            if self.cost_choice_var.get() == 'system':
                final_cost = utils.convert_to_float(self.system_data.get('cost_db', 0))
            else: # 'excel'
                final_cost = utils.convert_to_float(self.excel_data.get('cost_uploaded', 0))

            final_gp = final_sale - final_cost
            final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
            
            # --- Database Operation ---
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                update_query = """
                    UPDATE commissions SET
                        final_sales_amount = %s,
                        final_cost_amount = %s,
                        final_gp = %s,
                        final_margin = %s,
                        status = %s
                    WHERE id = %s
                """
                cursor.execute(update_query, (final_sale, final_cost, final_gp, final_margin, 'HR Verified', self.record_id))
            conn.commit()

            messagebox.showinfo("สำเร็จ", "ยืนยันและบันทึกข้อมูลเรียบร้อยแล้ว", parent=self.master)

            if self.refresh_callback:
                self.refresh_callback()
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_close(self):
        if self.refresh_callback:
            self.refresh_callback()
        self.destroy()
    
    def _update_all_calculations_and_ui(self):
        """
        อัปเดตข้อมูลที่คำนวณได้ทั้งหมด และรีเฟรช UI ที่เกี่ยวข้อง
        (เวอร์ชันแก้ไขสำหรับ UI ใหม่)
        """
        self._reload_data() # <-- เพิ่มบรรทัดนี้เพื่อดึงข้อมูลล่าสุดก่อน
        self._recalculate_summaries()
        self._populate_po_cards() # รีเฟรชรายการ PO
        self._update_selection_display() # อัปเดตการ์ดสรุปยอดขายและต้นทุน

    # อยู่ในไฟล์ hr_windows.py ภายในคลาส HRVerificationWindow

    def _update_selection_display(self, *args):
        """
        อัปเดตตัวเลขในหน้าจอสรุป (ฉบับแก้ไข: แสดงผลรวม Multiplier แต่เก็บค่าดิบไว้ข้างหลัง)
        """
        # 1. ดึงค่าจากตัวแปรกลางที่คำนวณไว้ (ซึ่งตอนนี้เก็บเป็น 'ทุนดิบ' ตามฟังก์ชัน _recalculate_summaries ใหม่)
        total_sale_system = self.calculated_values.get('total_sale_system', 0.0)
        total_sale_express = self.calculated_values.get('total_sale_express', 0.0)
        
        # ดึงทุนดิบมาเตรียมคำนวณเพื่อการแสดงผล
        raw_cost_system = self.calculated_values.get('total_cost_system', 0.0)
        total_cost_express = self.calculated_values.get('total_cost_express', 0.0)
        
        # ดึงตัวคูณปัจจุบัน (เช่น 1.03)
        multiplier = float(self.cost_multiplier_var.get())
        
        # คำนวณยอดที่รวมตัวคูณแล้ว "เพื่อใช้แสดงผลบนหน้าจอเท่านั้น"
        display_cost_system = raw_cost_system * multiplier

        # 2. อัปเดตข้อความบน Radio Buttons (โชว์ยอดที่รวม 1.03 ให้ HR สบายใจ)
        self.sales_system_radio.configure(text=f"จากระบบ (System): {total_sale_system:,.2f} บาท")
        self.sales_express_radio.configure(text=f"จากไฟล์ (Express): {total_sale_express:,.2f} บาท")
        self.cost_system_radio.configure(text=f"จากระบบ (System): {display_cost_system:,.2f} บาท")
        self.cost_express_radio.configure(text=f"จากไฟล์ (Express): {total_cost_express:,.2f} บาท")

        # 3. อัปเดตการ์ดสรุปยอดขาย (Big Card)
        source_sales = self.final_sale_source.get()
        final_sales_val = total_sale_system if source_sales == "system" else total_sale_express
        
        if hasattr(self, 'final_sales_label'):
            self.final_sales_label.configure(text=f"{final_sales_val:,.2f}")
            self.final_sales_source_label.configure(text=f"ที่มา: {'System' if source_sales == 'system' else 'Express'}")

        # 4. อัปเดตการ์ดสรุปต้นทุน (Big Card)
        source_cost = self.final_cost_source.get()
        
        # หากเลือก System ให้แสดงยอดที่คูณ 1.03 แล้วบนหน้าจอ
        if source_cost == "system":
            final_cost_display_val = display_cost_system
        else:
            final_cost_display_val = total_cost_express
            
        if hasattr(self, 'final_cost_label'):
            self.final_cost_label.configure(text=f"{final_cost_display_val:,.2f}")
            self.final_cost_source_label.configure(text=f"ที่มา: {'System' if source_cost == 'system' else 'Express'}")
            
        # (Optional) คำนวณ Profit/Margin ชั่วคราวเพื่อแสดงผลการตรวจสอบ
        final_profit = final_sales_val - final_cost_display_val
        final_margin = (final_profit / final_sales_val * 100) if final_sales_val else 0.0

    def _so_create_string_vars(self):
        """สร้าง StringVars ที่จำเป็นสำหรับ SOPopupWindow"""
        self.so_form_widgets = {}
        self.sale_theme = self.app_container.THEME["sale"]
        self.dropdown_style = {
            "fg_color": "white", "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }
        self.so_form_widgets['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        self.so_form_widgets['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_form_widgets['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_verification_result_var'] = tk.StringVar(value="-")
        
        # --- START: เพิ่ม 5 บรรทัดนี้เข้าไป ---
        # สร้าง StringVars สำหรับแสดง VAT ของแต่ละรายการที่ขาดไป
        self.so_form_widgets['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['relocation_cost_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['relocation_vat_calc_var'] = tk.StringVar(value="0.00")

    def _save_so_changes_from_popup(self, so_id, so_shared_vars_data, current_popup_widgets_ref):
        """
        [ฉบับแก้ไขสมบูรณ์: แยกยอดเงินโอนและเงินสดออกจากกันเด็ดขาด] 
        เพื่อให้ยอด 'ผลต่าง' ในส่วน VAT คิดเฉพาะยอดที่โอนผ่านธนาคารเท่านั้น
        """
        updated_data = {}
        
        # 1. Mapping ชื่อ Widget -> ชื่อคอลัมน์ DB
        key_map = {
            'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost',
            'relocation_cost_entry': 'relocation_cost', 'credit_card_fee_entry': 'credit_card_fee',
            'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons',
            'giveaway_value_entry': 'giveaways', 'cash_product_input_entry': 'cash_product_input',
            'bill_date_selector': 'bill_date',
            'delivery_date_selector': 'delivery_date', 'payment_date_selector': 'payment_date',
            'date_to_wh_selector': 'date_to_warehouse', 'date_to_customer_selector': 'date_to_customer',
            'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id',
            'credit_term_entry': 'credit_term', 'pickup_location_entry': 'pickup_location',
            'pickup_rego_entry': 'pickup_registration'
        }

        # 2. รวบรวมข้อมูลพื้นฐานจาก Widgets
        for widget_key, data_key in key_map.items():
            if widget_key in current_popup_widgets_ref:
                widget = current_popup_widgets_ref[widget_key]
                try:
                    if not widget.winfo_exists(): continue
                    if isinstance(widget, (NumericEntry, CTkEntry)):
                        raw_value = widget.get()
                        is_numeric = any(x in data_key for x in ['amount', 'cost', 'fee', 'wht', 'coupons', 'giveaways', 'input'])
                        updated_data[data_key] = utils.convert_to_float(raw_value) if is_numeric else raw_value
                    elif isinstance(widget, DateSelector):
                        updated_data[data_key] = widget.get_date()
                except: continue

        # 3. [🔥 จุดแก้ไขหลัก] แยกยอดเงินโอน และ ยอดเงินสด
        # ดึงยอดโอน 1 และ 2 (ฝั่งที่ต้องเทียบกับยอดรวม VAT)
        p1 = utils.convert_to_float(current_popup_widgets_ref.get('payment1_amount_entry').get()) if current_popup_widgets_ref.get('payment1_amount_entry') else 0.0
        p2 = utils.convert_to_float(current_popup_widgets_ref.get('payment2_amount_entry').get()) if current_popup_widgets_ref.get('payment2_amount_entry') else 0.0
        total_transfer = p1 + p2 # ยอดโอนรวม

        # ดึงยอดเงินสด (ฝั่ง Cash)
        cash_paid = 0.0
        if 'cash_actual_payment_entry' in current_popup_widgets_ref:
            cash_paid = utils.convert_to_float(current_popup_widgets_ref['cash_actual_payment_entry'].get())

        # บันทึกลงคอลัมน์ที่แยกกัน
        updated_data['total_payment_amount'] = total_transfer  # ยอดโอนชำระ (VAT)
        updated_data['cash_actual_payment'] = cash_paid       # ยอดชำระเงินสดจริง (CASH)

        # 4. [🔥 แก้ไข Logic ส่วนต่าง] คำนวณเฉพาะส่วนต่างของยอดโอน (ไม่เอาเงินสดมาเกี่ยว)
        try:
            # ยอดที่ระบบคำนวณว่าควรจะโอน (Grand Total ของสินค้าฝั่ง VAT)
            grand_total_vat = utils.convert_to_float(so_shared_vars_data['so_grand_total_var'].get())
            
            # ผลต่าง VAT = ยอดโอนรวม - ยอดเต็มบิลที่ต้องออก VAT
            # บรรทัดนี้สำคัญ: ต้องไม่มีการบวก cash_paid เข้าไปเด็ดขาด
            new_difference = total_transfer - grand_total_vat
            updated_data['difference_amount'] = new_difference
            
            # อัปเดตยอดโชว์บนหน้าจอทันที
            if 'difference_amount_var' in so_shared_vars_data:
                so_shared_vars_data['difference_amount_var'].set(f"{new_difference:,.2f}")
                
        except Exception as e:
            print(f"Error calculating VAT difference: {e}")

        # 5. บันทึกลงฐานข้อมูล
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'commissions'")
                db_columns = {row[0] for row in cursor.fetchall()}
                final_data_to_save = {k: v for k, v in updated_data.items() if k in db_columns}

                if final_data_to_save:
                    set_clauses = [f'"{k}" = %s' for k in final_data_to_save.keys()]
                    params = list(final_data_to_save.values())
                    params.append(so_id)
                    
                    update_query = f'UPDATE commissions SET {", ".join(set_clauses)} WHERE id = %s'
                    cursor.execute(update_query, tuple(params))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกและแยกยอดเงินสด/โอนเรียบร้อย", parent=self)
            
            # รีเฟรชข้อมูลในหน้า Verification หลังบันทึก
            self._update_all_calculations_and_ui()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"บันทึกไม่สำเร็จ: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
            
    def _open_so_editor_popup(self):
        """เปิดหน้าต่าง SOPopupWindow สำหรับให้ HR แก้ไขข้อมูล SO โดยละเอียด"""
        SOPopupWindow(
            master=self,
            app_container=self.app_container, # <-- เพิ่มบรรทัดนี้ที่ขาดไป
            sales_data=self.system_data,
            so_shared_vars=self.so_form_widgets,
            sale_theme=self.sale_theme,
            on_save_callback=self._update_all_calculations_and_ui
        )

    def _on_payout_history_double_click(self, event, tree):
        """(เวอร์ชัน Debug) ตรวจสอบตำแหน่งคลิกและเปิด Popup"""
        print("DEBUG: Double Click Detected!") # เช็คว่า event ทำงานไหม

        try:
            # 1. ตรวจสอบ region
            region = tree.identify("region", event.x, event.y)
            print(f"DEBUG: Region = {region}")
            if region != "cell": 
                return 

            # 2. ดึง Item ID
            selected_item_iid = tree.focus()
            print(f"DEBUG: Selected IID = {selected_item_iid}")
            if not selected_item_iid:
                return 

            # 3. แปลง ID
            payout_id = int(selected_item_iid)
            print(f"DEBUG: Payout ID = {payout_id}")

            # 4. เปิด Popup
            from hr_windows import PayoutDetailWindow 
            PayoutDetailWindow(
                master=self, 
                app_container=self.app_container, 
                payout_id=payout_id
            )
            print("DEBUG: Popup Opened Successfully")

        except Exception as e:
            # ใช้ print ธรรมดาแทน traceback เผื่อลืม import
            print(f"ERROR in double click: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"เปิดหน้าต่างไม่ได้: {e}", parent=self)

    def _save_cell_edit(self, tree, item_id, column_index, new_value_str):
        """บันทึกค่าที่แก้ไขและคำนวณสรุปใหม่"""
        try:
            new_value = float(str(new_value_str).replace(",", ""))
        except (ValueError, TypeError):
            messagebox.showwarning("ข้อมูลผิดพลาด", "กรุณาใส่ค่าเป็นตัวเลขเท่านั้น", parent=self)
            self._recalculate_summaries() # Refresh to original state
            return

        # หาว่าเป็นตาราง Sales หรือ Cost
        is_sales_tree = (tree == self.sales_tree)
        field_name = self.sales_tree.item(item_id, "values")[0] if is_sales_tree else self.costing_tree.item(item_id, "values")[0]
        field_name = field_name.replace(" *", "").strip()

        # Map ชื่อที่แสดงผลกลับไปเป็นชื่อ Key ใน Dictionary
        field_to_key_map = {
            'รายได้ค่าสินค้า/บริการ': ('sales_service_amount', 'sales_service_amount'),
            'รายได้ค่าการจัดส่ง': ('shipping_cost', 'shipping_cost_uploaded'),
            'ต้นทุนค่าสินค้า/บริการ': (None, 'cost_uploaded'),
            'ต้นทุนค่าจัดส่ง': (None, 'shipping_cost_uploaded'),
            'ต้นทุนค่าย้าย': (None, 'relocation_cost_uploaded'),
            'ต้นทุนค่านายหน้า': ('brokerage_fee', 'brokerage_fee_uploaded'),
            'ต้นทุนค่าธรรมเนียมโอน': ('transfer_fee', 'transfer_fee_uploaded'),
        }
        
        keys = field_to_key_map.get(field_name)
        if not keys:
            return

        # อัปเดตข้อมูลใน Memory
# อัปเดตข้อมูลใน Memory
        if column_index == 1:  # System Data
           if is_sales_tree:
              target_dict = self.system_data
              key = keys[0]
              if key:
                 target_dict[key] = new_value
           else:  # ถ้าเป็นตาราง Costing
            self.cost_overrides[field_name] = new_value

        elif column_index == 2:  # Express Data
             target_dict = self.excel_data
             key = keys[1]
             if key:
                target_dict[key] = new_value

        
        # คำนวณทุกอย่างใหม่
        self._recalculate_summaries()
        self._refresh_sales_comparison_table()
        self._refresh_costing_comparison_table()
        self._update_all_calculations_and_ui()
        self._update_all_calculations_and_ui()
# อยู่ในไฟล์ hr_windows.py ภายในคลาส HRVerificationWindow

    def _recalculate_summaries(self):
        """
        คำนวณต้นทุนและยอดขาย (แก้ไขเพื่อป้องกันการคูณค่าบริหารจัดการเบิ้ล)
        """
        po_product_cost_raw = 0.0
        approved_po_df = self.po_data[self.po_data['status'] == 'Approved']
        
        if not approved_po_df.empty:
            approved_po_ids = tuple(approved_po_df['id'].tolist())
            try:
                query = "SELECT product_name, total_price FROM purchase_order_items WHERE purchase_order_id IN %s"
                items_df = pd.read_sql(query, self.pg_engine, params=(approved_po_ids,))
                shipping_keywords = ['ค่ารถ', 'shipping', 'delivery', 'ขนส่ง', 'ค่าขนย้าย', 'relocation', 'ค่าส่ง']
                
                for _, row in items_df.iterrows():
                    p_name = str(row['product_name']).lower()
                    price = float(row['total_price'] or 0)
                    if not any(k in p_name for k in shipping_keywords):
                        po_product_cost_raw += price
            except: pass

        # 1. เก็บ "ทุนดิบ" ลงในระบบเพื่อรอการ Save (Business Logic จะไปคูณ 1.03 เอง)
        total_sale_express = float(self.excel_data.get('sales_uploaded', 0) or 0)
        total_cost_express = float(self.excel_data.get('cost_uploaded', 0) or 0)
        sales_keys = ['sales_service_amount', 'cutting_drilling_fee', 'other_service_fee']
        total_sale_system = sum(float(self.system_data.get(key, 0) or 0) for key in sales_keys)

        self.calculated_values = {
            'total_sale_system': total_sale_system,
            'total_sale_express': total_sale_express,
            'total_cost_system': po_product_cost_raw,  # ส่ง 23,336.00 (ทุนดิบ)
            'total_cost_express': total_cost_express
        }

        # 2. ปรับการตั้งค่าเริ่มต้น (ถ้ามี)
        for key, source in [('hr_sale_source', self.final_sale_source), ('hr_cost_source', self.final_cost_source)]:
            val = self.system_data.get(key)
            source.set(val if val in ['system', 'express'] else "system")
            
        # 3. สั่ง Update หน้าจอ (ฟังก์ชันนี้จะทำหน้าที่แสดงผลเป็น 24,036.08 ให้เอง)
        self._update_selection_display()

    def _create_so_info_section(self):
        frame = CTkFrame(self, fg_color="#F0F0F0")
        frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        font_bold=CTkFont(size=14, weight="bold"); font_normal=CTkFont(size=14)
        
        so_info_frame = CTkFrame(frame, fg_color="transparent")
        so_info_frame.pack(side="left", padx=10, pady=5)
        
        CTkLabel(so_info_frame, text="SO Number:", font=font_bold).grid(row=0, column=0, sticky="w")
        CTkLabel(so_info_frame, text=self.system_data.get('so_number', 'N/A'), font=font_normal).grid(row=0, column=1, padx=5, sticky="w")
        
        sale_name = self.system_data.get('sale_name', self.system_data.get('sale_key', 'N/A'))
        CTkLabel(so_info_frame, text="Sale Name:", font=font_bold).grid(row=0, column=2, padx=10, sticky="w")
        CTkLabel(so_info_frame, text=sale_name, font=font_normal).grid(row=0, column=3, padx=5, sticky="w")
        
        CTkLabel(so_info_frame, text="Customer:", font=font_bold).grid(row=1, column=0, sticky="w")
        CTkLabel(so_info_frame, text=self.system_data.get('customer_name', 'N/A'), font=font_normal, wraplength=400).grid(row=1, column=1, columnspan=3, sticky="w")
        
        CTkButton(frame, text="ดู/แก้ไขข้อมูล SO", 
          command=self._open_so_editor_popup
        ).pack(side="left", padx=20)

    def _create_main_paned_window(self):
        paned_window = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=8, bg="#D1D5DB")
        paned_window.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        left_pane = CTkFrame(paned_window, fg_color="transparent")
        left_pane.grid_rowconfigure(0, weight=1); left_pane.grid_columnconfigure(0, weight=1)
        paned_window.add(left_pane, width=700)
        self._create_sales_info_column(left_pane)
        right_pane = CTkFrame(paned_window, fg_color="transparent")
        right_pane.grid_rowconfigure(0, weight=1); right_pane.grid_columnconfigure(0, weight=1)
        paned_window.add(right_pane, width=700)
        self._create_costing_info_column(right_pane)

    def _create_final_summary_section(self, parent_scroll_frame):
        """(ฉบับปรับปรุง UI) สร้าง Widget ในส่วนสรุปและเลือกแหล่งข้อมูล"""
        
        # Frame หลักสำหรับส่วนนี้ทั้งหมด
        frame = CTkFrame(parent_scroll_frame, border_width=1, corner_radius=10)
        frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        # --- ตั้งค่าให้ Frame หลักมี 2 คอลัมน์ที่ขยายเท่ากัน ---
        frame.grid_columnconfigure((0, 1), weight=1)

        # หัวข้อหลัก (อยู่ตรงกลางด้านบน)
        CTkLabel(frame, text="สรุปและเลือกข้อมูลเพื่อคำนวณ Margin/Commission", font=CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=10, pady=10)

        # --- การ์ดฝั่งซ้าย (ยอดขาย) ---
        sales_frame = CTkFrame(frame, fg_color=("gray90", "gray20"))
        sales_frame.grid(row=1, column=0, padx=(10, 5), pady=10, sticky="nsew")
        sales_frame.grid_columnconfigure(0, weight=1) # ทำให้ widget ภายในขยายเต็ม

        CTkLabel(sales_frame, text="ยอดขายรวมทั้งหมด", font=CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 5))
        
        self.sales_system_radio = CTkRadioButton(sales_frame, text="จากระบบ: 0.00", variable=self.final_sale_source, value="system")
        self.sales_system_radio.pack(anchor="w", padx=20, pady=5)
        
        self.sales_express_radio = CTkRadioButton(sales_frame, text="จาก Express: 0.00", variable=self.final_sale_source, value="express")
        self.sales_express_radio.pack(anchor="w", padx=20, pady=(5, 15))

        # --- การ์ดฝั่งขวา (ต้นทุน) ---
        cost_frame = CTkFrame(frame, fg_color=("gray90", "gray20"))
        cost_frame.grid(row=1, column=1, padx=(5, 10), pady=10, sticky="nsew")
        cost_frame.grid_columnconfigure(0, weight=1) # ทำให้ widget ภายในขยายเต็ม
        
        CTkLabel(cost_frame, text="ยอดต้นทุนรวมทั้งหมด", font=CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 5))
        
        self.cost_system_radio = CTkRadioButton(cost_frame, text="จากระบบ: 0.00", variable=self.final_cost_source, value="system")
        self.cost_system_radio.pack(anchor="w", padx=20, pady=5)
        
        self.cost_express_radio = CTkRadioButton(cost_frame, text="จาก Express: 0.00", variable=self.final_cost_source, value="express")
        self.cost_express_radio.pack(anchor="w", padx=20, pady=(5, 15))
        # +++ END: สิ้นสุดการแก้ไข +++

    def _create_sales_info_column(self, parent):
        frame = CTkFrame(parent); frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(1, weight=1); frame.grid_columnconfigure(0, weight=1)
        CTkLabel(frame, text="เปรียบเทียบข้อมูลขาย (ในระบบ vs. Express)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=5)
        self.sales_tree = ttk.Treeview(frame, columns=("Field", "System", "Express"), show="headings")
        self.sales_tree.grid(row=1, column=0, sticky="nsew")
        self.sales_tree.heading("Field", text="ฟิลด์ข้อมูล"); self.sales_tree.heading("System", text="ข้อมูลในระบบ"); self.sales_tree.heading("Express", text="ข้อมูลจาก Express")
        self.sales_tree.column("Field", width=180); self.sales_tree.column("System", width=180, anchor="e"); self.sales_tree.column("Express", width=180, anchor="e")
        self.sales_tree.tag_configure('mismatch', background='#FEE2E2'); self.sales_tree.tag_configure('match', background='#F0FDF4')
        self._refresh_sales_comparison_table()
        self.sales_tree.bind("<Double-1>", self._on_cell_double_click)

    def _create_costing_info_column(self, parent):
        frame = CTkFrame(parent); frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(1, weight=1); frame.grid_columnconfigure(0, weight=1)
        CTkLabel(frame, text="เปรียบเทียบข้อมูลทุน (ในระบบ vs. Express)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=5)
        self.costing_tree = ttk.Treeview(frame, columns=("Field", "System", "Express"), show="headings")
        self.costing_tree.grid(row=1, column=0, sticky="nsew")
        self.costing_tree.heading("Field", text="ฟิลด์ข้อมูล"); self.costing_tree.heading("System", text="ข้อมูลในระบบ (PU)"); self.costing_tree.heading("Express", text="ข้อมูลจาก Express")
        self.costing_tree.column("Field", width=180); self.costing_tree.column("System", width=180, anchor="e"); self.costing_tree.column("Express", width=180, anchor="e")
        self.costing_tree.tag_configure('mismatch', background='#FEE2E2'); self.costing_tree.tag_configure('match', background='#F0FDF4')
        self._refresh_costing_comparison_table()
        self.costing_tree.bind("<Double-1>", self._on_cell_double_click)
        self.costing_tree.bind("<Double-1>", self._on_cell_double_click)

    def _refresh_sales_comparison_table(self):
        for item in self.sales_tree.get_children(): self.sales_tree.delete(item)
        sales_fields_map = {
            'รายได้ค่าสินค้า/บริการ': ('sales_service_amount', 'sales_uploaded'),
            'รายได้ค่าการจัดส่ง': ('shipping_cost', 'shipping_cost_uploaded'),
        }
        for display_name, (sys_key, exp_key) in sales_fields_map.items():
            sys_val = self.system_data.get(sys_key)
            exp_val = self.excel_data.get(exp_key)
            sys_str = f"{sys_val:,.2f}" if isinstance(sys_val, (int, float)) else str(sys_val or 'N/A')
            exp_str = f"{exp_val:,.2f}" if isinstance(exp_val, (int, float)) else str(exp_val or 'N/A')
            tag = 'match' if sys_str == exp_str else 'mismatch'
            self.sales_tree.insert("", "end", values=(display_name, sys_str, exp_str), tags=(tag,))

    def _refresh_costing_comparison_table(self):
        for item in self.costing_tree.get_children(): self.costing_tree.delete(item)
        total_po_shipping_cost = self.po_data['shipping_to_stock_cost'].sum() + self.po_data['shipping_to_site_cost'].sum()
        total_po_product_cost = self.po_data['total_cost'].sum() - total_po_shipping_cost
        total_po_relocation_cost = self.po_data['relocation_cost'].sum() if 'relocation_cost' in self.po_data.columns else 0
        brokerage_cost = self.system_data.get('brokerage_fee', 0) or 0
        transfer_fee_cost = self.system_data.get('transfer_fee', 0) or 0
        cost_fields_map = {
            'ต้นทุนค่าสินค้า/บริการ': (total_po_product_cost, 'cost_uploaded'),
            'ต้นทุนค่าจัดส่ง': (total_po_shipping_cost, 'shipping_cost_uploaded'),
            'ต้นทุนค่าย้าย': (total_po_relocation_cost, 'relocation_cost_uploaded'),
            'ต้นทุนค่านายหน้า': (brokerage_cost, 'brokerage_fee_uploaded'),
            'ต้นทุนค่าธรรมเนียมโอน': (transfer_fee_cost, 'transfer_fee_uploaded'),
        }
        for display_name, (original_sys_val, exp_key) in cost_fields_map.items():
           # ตรวจสอบว่ามีค่าที่ถูกแก้ไขเก็บไว้หรือไม่
           sys_val = self.cost_overrides.get(display_name, original_sys_val)

           display_name_with_star = f"{display_name} *" if display_name in self.cost_overrides else display_name

           exp_val = self.excel_data.get(exp_key, 0) or 0
           tag = 'match' if f"{sys_val:,.2f}" == f"{exp_val:,.2f}" else 'mismatch'
           self.costing_tree.insert("", "end", values=(display_name_with_star, f"{sys_val:,.2f}", f"{exp_val:,.2f}"), tags=(tag,))

    def _create_po_summary_table(self):
        main_frame = CTkScrollableFrame(self, label_text="ใบสั่งซื้อ (PO) ที่เกี่ยวข้อง")
        main_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        if self.po_data.empty: CTkLabel(main_frame, text="ไม่พบข้อมูล Purchase Order ที่เกี่ยวข้อง").pack(pady=20); return
        for _, po_row in self.po_data.iterrows():
            po_id = po_row['id']; card = CTkFrame(main_frame, border_width=1); card.pack(fill="x", padx=5, pady=5)
            header_frame = CTkFrame(card, fg_color="#F3F4F6", corner_radius=6); header_frame.pack(fill="x", padx=3, pady=3); header_frame.grid_columnconfigure(1, weight=1)
            toggle_label = CTkLabel(header_frame, text="▶", font=CTkFont(size=14), cursor="hand2"); toggle_label.grid(row=0, column=0, padx=(5,2))
            po_text = f"PO: {po_row['po_number']}  |  ซัพพลายเออร์: {po_row['supplier_name']}  |  ยอดรวมต้นทุน: {po_row['total_cost']:,.2f}"; info_label = CTkLabel(header_frame, text=po_text, font=CTkFont(size=14)); info_label.grid(row=0, column=1, padx=5, pady=8, sticky="w")
            status_color = "#16A34A" if po_row['status'] == 'Approved' else 'gray'; status_label = CTkLabel(header_frame, text=f"สถานะ: {po_row['status']}", font=CTkFont(size=12, weight="bold"), text_color=status_color); status_label.grid(row=0, column=2, padx=10, sticky="e")
            detail_button = CTkButton(header_frame, text="ดูรายละเอียด PO", width=120, command=lambda p_id=po_id: self.app_container.show_purchase_detail_window(p_id)); detail_button.grid(row=0, column=3, padx=10)
            detail_frame = CTkFrame(card, fg_color="#FAFAFA")
            toggle_widgets = [toggle_label, info_label] 
            for widget in toggle_widgets:
                widget.bind("<Button-1>", lambda e, df=detail_frame, pid=po_id, tl=toggle_label: self._toggle_po_items(df, pid, tl))
    
    def _toggle_po_items(self, detail_frame, po_id, toggle_label):
        if detail_frame.winfo_viewable(): detail_frame.pack_forget(); toggle_label.configure(text="▶")
        else:
            detail_frame.pack(fill="both", expand=True, padx=15, pady=(0,5)); toggle_label.configure(text="▼")
            if not detail_frame.winfo_children():
                try:
                    query = "SELECT product_name, quantity, unit_price, total_price FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id"
                    items_df = pd.read_sql_query(query, self.app_container.pg_engine, params=(po_id,))
                    if items_df.empty: CTkLabel(detail_frame, text="ไม่พบรายการสินค้าใน PO นี้").pack(pady=5); return
                    tree_container = CTkFrame(detail_frame, fg_color="transparent"); tree_container.pack(fill="both", expand=True, pady=5); tree_container.grid_columnconfigure(0, weight=1); tree_container.grid_rowconfigure(0, weight=1)
                    tree = ttk.Treeview(tree_container, columns=("name", "qty", "price", "total"), show="headings"); tree.grid(row=0, column=0, sticky="nsew")
                    scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview); scrollbar.grid(row=0, column=1, sticky="ns"); tree.configure(yscrollcommand=scrollbar.set)
                    tree.heading("name", text="ชื่อสินค้า"); tree.heading("qty", text="จำนวน"); tree.heading("price", text="ราคา/หน่วย"); tree.heading("total", text="ราคารวม")
                    tree.column("name", width=300, anchor="w"); tree.column("qty", width=80, anchor="e"); tree.column("price", width=100, anchor="e"); tree.column("total", width=100, anchor="e")
                    for _, item_row in items_df.iterrows(): tree.insert("", "end", values=(item_row['product_name'], f"{item_row['quantity']:,.2f}", f"{item_row['unit_price']:,.2f}", f"{item_row['total_price']:,.2f}"))
                except Exception as e: CTkLabel(detail_frame, text=f"Error loading items: {e}", text_color="red").pack(pady=5)
    
    def _create_action_buttons(self):
        frame = CTkFrame(self, fg_color="transparent")
        frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew") # <<< แก้ไข: sticky="ew"
        frame.grid_columnconfigure((0,1,2), weight=1) # <<< เพิ่ม: ทำให้ปุ่มขยายเท่ากัน

        CTkButton(frame, text="ตีกลับให้ฝ่ายจัดซื้อ (Reject to PU)", fg_color="#F97316", hover_color="#EA580C", command=self._reject_to_purchasing).grid(row=0, column=0, padx=5, pady=5, sticky="ew")

    # <<< เพิ่ม: ปุ่มบันทึกการแก้ไข >>>
        CTkButton(frame, text="💾 บันทึกการแก้ไข", fg_color="#3B82F6", hover_color="#2563EB", command=self._save_intermediate_changes).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        CTkButton(frame, text="ยืนยันข้อมูลถูกต้อง (Verify Data)", fg_color="#16A34A", hover_color="#15803D", command=self._verify_and_save_data).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
    # hr_windows.py (ฟังก์ชัน _reject_to_purchasing ที่แก้ไขแล้ว)

    def _reject_to_salesperson(self):
        """
        ตีกลับ SO และ PO ทั้งหมดที่เกี่ยวข้อง ไปยังเซลส์เจ้าของ
        """
        dialog = CTkInputDialog(text="กรุณาระบุเหตุผลที่ตีกลับ (จะถูกส่งไปให้เซลส์):", title="ตีกลับ SO และ PO ทั้งหมด")
        reason = dialog.get_input()
        
        if not reason or not reason.strip():
            return

        so_id = self.system_data.get('id')
        sale_key_to_notify = self.system_data.get('sale_key')

        if not sale_key_to_notify:
            messagebox.showerror("ผิดพลาด", "ไม่สามารถหารหัสของเซลส์เจ้าของ SO นี้ได้", parent=self)
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 1. อัปเดตสถานะ SO และบันทึกเหตุผล
                cursor.execute(
                    "UPDATE commissions SET status = 'Rejected by HR', rejection_reason = %s WHERE id = %s",
                    (reason.strip(), so_id)
                )

                # <<< START: เพิ่มโค้ดส่วนนี้ >>>
                # 2. อัปเดตสถานะของ PO ทุกใบที่เกี่ยวข้องกับ SO นี้ให้เป็น 'Rejected' ด้วย
                cursor.execute(
                    "UPDATE purchase_orders SET status = 'Rejected', approval_status = 'Rejected' WHERE so_number = %s",
                    (self.so_number,)
                )
                # <<< END >>>

                # 3. สร้าง Notification แจ้งเตือนเซลส์เจ้าของโดยตรง
                message = f"SO: {self.so_number} ของคุณถูกตีกลับโดย HR\nเหตุผล: {reason.strip()}"
                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)",
                    (sale_key_to_notify, message, so_id)
                )
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ตีกลับ SO และ PO ที่เกี่ยวข้องทั้งหมดไปยังฝ่ายขายเรียบร้อยแล้ว", parent=self.master)
            
            self._on_close()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _defer_so(self):
        """
        (แก้ไขใหม่) ส่งคำขอเลื่อนจ่าย (Defer Request) ให้ Sale พิจารณา (แทนการบังคับเลื่อน)
        """
        so_number = self.system_data.get('so_number')
        
        # 1. ให้ HR ระบุเหตุผล (ใช้ RejectionReasonDialog หรือ InputDialog ก็ได้)
        # แนะนำใช้ RejectionReasonDialog เพื่อความสวยงามและ UX ที่เหมือนกัน
        from utils import RejectionReasonDialog # ตรวจสอบว่า import แล้ว
        dialog = RejectionReasonDialog(self)
        self.wait_window(dialog)
        reason = getattr(dialog, '_reason_string', None)
        
        if reason is None: return # กดยกเลิก

        # ถามยืนยันอีกครั้ง
        if not messagebox.askyesno("ยืนยัน", f"ต้องการส่งคำขอเลื่อนจ่าย SO: {so_number} ให้ฝ่ายขายพิจารณาใช่หรือไม่?"):
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 2. อัปเดตสถานะเป็น 'Defer Requested' 
                # (ยังไม่เปลี่ยนเดือน/ปี เพราะต้องรอ Sale เลือกเดือนเอง)
                cursor.execute("""
                    UPDATE commissions 
                    SET 
                        status = 'Defer Requested', 
                        rejection_reason = %s
                    WHERE id = %s
                """, (f"HR Request: {reason.strip()}", self.system_data['id']))
                
                # 3. แจ้งเตือน Sale
                sale_key = self.system_data.get('sale_key')
                if sale_key:
                    msg = f"HR ขอเลื่อนจ่าย SO: {so_number}\nเหตุผล: {reason.strip()}\n(กรุณาไปที่เมนู 'งานของฉัน' เพื่อตอบรับหรือปฏิเสธ)"
                    cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)", (sale_key, msg, self.system_data['id']))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ส่งคำขอเลื่อนจ่าย SO: {so_number} ให้ฝ่ายขายพิจารณาเรียบร้อยแล้ว", parent=self.master)
            
            self._on_close() # ปิดหน้าต่างและ Refresh หน้าหลัก

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _save_intermediate_changes(self):
      """บันทึกข้อมูล SO และ PU overrides ลง DB โดยไม่เปลี่ยนสถานะ"""
      if not messagebox.askyesno("ยืนยัน", "คุณต้องการบันทึกการเปลี่ยนแปลงทั้งหมดลงฐานข้อมูลใช่หรือไม่?", parent=self):
          return

      so_id = self.system_data.get('id')
      conn = None
      try:
          conn = self.app_container.get_connection()
          with conn.cursor() as cursor:
              columns_to_update = [
                  "sales_service_amount", "shipping_cost", "relocation_cost", "brokerage_fee",
                  "transfer_fee", "cutting_drilling_fee", "other_service_fee", "credit_card_fee",
                  "wht_3_percent", "coupons", "giveaways", "total_payment_amount",
                  "cash_product_input", "cash_actual_payment", "bill_date", "delivery_date",
                  "payment_date", "date_to_warehouse", "date_to_customer", "customer_name",
                  "customer_id", "credit_term", "delivery_type", "pickup_location", "pickup_registration",
                  "payment_before_vat", 
                  "payment_no_vat"      
              ]
              
              set_clauses = [f"{col} = %s" for col in columns_to_update]
              
              params = [self.system_data.get(col) for col in columns_to_update]

              cost_overrides_json = json.dumps(self.cost_overrides)
              set_clauses.append("hr_cost_overrides = %s")
              params.append(cost_overrides_json)

              set_clauses.append("hr_sale_source = %s")
              params.append(self.final_sale_source.get())
              set_clauses.append("hr_cost_source = %s")
              params.append(self.final_cost_source.get())

              set_clauses.append("cost_multiplier = %s")
              params.append(float(self.cost_multiplier_var.get()))

              update_query = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
              params.append(so_id)

              cursor.execute(update_query, tuple(params))

          conn.commit()
          messagebox.showinfo("สำเร็จ", "บันทึกการแก้ไขข้อมูลเรียบร้อยแล้ว", parent=self)
          
          # +++ เพิ่มโค้ด 2 บรรทัดนี้เข้าไปท้ายสุดของ try block +++
          if self.refresh_callback:
              self.refresh_callback() # สั่งให้หน้าจอหลัก Refresh ทันที!

      except Exception as e:
          if conn: conn.rollback()
          messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึก: {e}", parent=self)
          traceback.print_exc()
      finally:
          if conn: self.app_container.release_connection(conn)


    def _verify_and_save_data(self):
        # --- ดึงค่าจากตัวแปรกลางที่คำนวณไว้แล้ว ---
        final_sale_source = self.final_sale_source.get()
        final_cost_source = self.final_cost_source.get()

        total_sale_system = self.calculated_values.get('total_sale_system', 0.0)
        total_sale_express = self.calculated_values.get('total_sale_express', 0.0)
        total_cost_system = self.calculated_values.get('total_cost_system', 0.0)
        total_cost_express = self.calculated_values.get('total_cost_express', 0.0)

        final_sale = float(total_sale_system if final_sale_source == "system" else total_sale_express)
        final_cost = float(total_cost_system if final_cost_source == "system" else total_cost_express)
        
        # +++ START: เพิ่มโค้ด Debug ตอนกด Save +++
        print("\n--- DEBUGGING SAVE ACTION ---")
        print(f"User selected cost source: '{final_cost_source}'")
        print(f"Value of total_cost_system WAS: {total_cost_system:,.2f}")
        print(f"Value of total_cost_express WAS: {total_cost_express:,.2f}")
        print(f"FINAL cost value to be saved IS: {final_cost:,.2f}")
        print("-----------------------------\n")
        
        so_id = self.system_data.get('id')
        final_gp = final_sale - final_cost
        final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
        
        msg = (f"คุณต้องการยืนยันข้อมูลสำหรับ SO นี้ใช่หรือไม่?\n\n"
               f"ยอดขายสุดท้ายที่เลือก: {final_sale:,.2f} บาท\n"
               f"ยอดต้นทุนสุดท้ายที่เลือก: {final_cost:,.2f} บาท\n"
               # ... (ข้อความที่เหลือ) ...
              )
        if not messagebox.askyesno("ยืนยันข้อมูล", msg, parent=self):
            return
            
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                update_query = """
                    UPDATE commissions 
                    SET 
                        status = 'HR Verified', 
                        final_sales_amount = %s, 
                        final_cost_amount = %s,
                        final_gp = %s,
                        final_margin = %s,
                        hr_sale_source = %s,
                        hr_cost_source = %s,
                        payout_id = NULL -- <--- เพิ่มบรรทัดนี้
                    WHERE id = %s
                """
                params = (final_sale, final_cost, final_gp, final_margin, 
                          self.final_sale_source.get(), self.final_cost_source.get(), so_id)
                cursor.execute(update_query, params)
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ยืนยันและบันทึกข้อมูลเรียบร้อยแล้ว", parent=self.master)
            self._on_close()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    

class PayoutDetailWindow(CTkToplevel):
    """
    หน้าต่างแสดงรายละเอียด Payout (มี Navigation เลื่อนเดือนได้)
    """
    def __init__(self, master, app_container, payout_id):
        super().__init__(master)
        self.app_container = app_container
        self.payout_id = payout_id
        self.payout_log_data = None
        
        # Theme Setting
        try:
            self.theme = self.app_container.THEME["hr"]
        except (AttributeError, KeyError):
            self.theme = {"primary": "#3B82F6", "header": "#1E40AF"}

        self.title("รายละเอียดการจ่ายค่าคอมมิชชั่น")
        self.geometry("1000x900")
        
        # จัด Layout หลัก: แถว 0 คือเมนู, แถว 1 คือเนื้อหา
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 

        # --- ส่วนที่ 1: Navigation Bar (ปุ่มเลื่อนซ้าย-ขวา) ---
        self.nav_frame = CTkFrame(self, fg_color="transparent")
        self.nav_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,0))
        self._create_navigation_bar() # <--- เรียกฟังก์ชันสร้างปุ่ม

        # --- ส่วนที่ 2: Content หลัก (ที่จะเปลี่ยนไปเรื่อยๆ) ---
        self.content_frame = CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.content_frame.grid_columnconfigure(0, weight=1)
        
        # เริ่มต้นโหลดข้อมูล
        self._refresh_content() # <--- ฟังก์ชันพระเอกที่คอยวาดหน้าจอใหม่

        self.transient(master)
        self.grab_set()
    
    def _create_navigation_bar(self):
        """สร้างปุ่มเลื่อนซ้ายขวา และชื่อ Title ตรงกลาง"""
        self.nav_frame.grid_columnconfigure(1, weight=1) # ให้ Title อยู่กลาง

        # ปุ่มย้อนกลับ (<)
        self.btn_prev = CTkButton(self.nav_frame, text="◀ รอบก่อนหน้า", width=120, 
                                  command=lambda: self._navigate_payout("prev"),
                                  fg_color="#64748B", hover_color="#475569")
        self.btn_prev.grid(row=0, column=0, padx=10)

        # Title ตรงกลาง
        self.lbl_title = CTkLabel(self.nav_frame, text=f"Payout ID: {self.payout_id}", font=("Arial", 20, "bold"))
        self.lbl_title.grid(row=0, column=1)

        # ปุ่มถัดไป (>)
        self.btn_next = CTkButton(self.nav_frame, text="รอบถัดไป ▶", width=120, 
                                  command=lambda: self._navigate_payout("next"),
                                  fg_color="#64748B", hover_color="#475569")
        self.btn_next.grid(row=0, column=2, padx=10)

    def _navigate_payout(self, direction):
        """ค้นหา ID ถัดไปหรือก่อนหน้า แล้วรีเฟรชหน้าจอ"""
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                if direction == "prev":
                    # หา ID ที่น้อยกว่าปัจจุบัน (ตัวล่าสุด)
                    cursor.execute("SELECT id FROM commission_payout_logs WHERE id < %s ORDER BY id DESC LIMIT 1", (self.payout_id,))
                else:
                    # หา ID ที่มากกว่าปัจจุบัน (ตัวแรกสุด)
                    cursor.execute("SELECT id FROM commission_payout_logs WHERE id > %s ORDER BY id ASC LIMIT 1", (self.payout_id,))
                
                result = cursor.fetchone()
                
                if result:
                    new_id = result[0]
                    self.payout_id = new_id # อัปเดต ID
                    self._refresh_content() # โหลดหน้าจอใหม่
                else:
                    messagebox.showinfo("สุดทางแล้ว", "ไม่มีข้อมูลของรอบการจ่ายนี้แล้ว", parent=self)

        except Exception as e:
            print(f"Navigation Error: {e}")
        finally:
            self.app_container.release_connection(conn)

    def _refresh_content(self):
        """โหลดข้อมูลใหม่และวาดหน้าจอใหม่ทั้งหมด (ไม่ต้องปิดหน้าต่าง)"""
        
        # 1. อัปเดต Title ด้านบน
        self.lbl_title.configure(text=f"Payout ID: {self.payout_id}")
        
        # 2. ล้าง Widget เก่าทิ้งให้หมด
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        # 3. โหลดข้อมูลใหม่
        self._load_data() 

        # 4. วาด UI ใหม่
        # --- ส่วน A: สถิติ ---
        self._create_header_statistics()

        # --- ส่วน B: ข้อมูลทั่วไปและปุ่ม ---
        top_sub_frame = CTkFrame(self.content_frame, fg_color="transparent")
        top_sub_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        top_sub_frame.grid_columnconfigure(0, weight=1)

        info_text = "กำลังโหลด..."
        if self.payout_log_data:
            ts = self.payout_log_data.get('timestamp')
            date_str = pd.to_datetime(ts).strftime('%d/%m/%Y %H:%M') if ts else "N/A"
            info_text = f"วันที่จ่าย: {date_str} | Sale Key: {self.payout_log_data.get('sale_key','-')} | Plan: {self.payout_log_data.get('plan_name','-')}"

        CTkLabel(top_sub_frame, text=info_text, font=('Roboto', 14), anchor="w").grid(row=0, column=0, sticky="w")

        btn_container = CTkFrame(top_sub_frame, fg_color="transparent")
        btn_container.grid(row=0, column=1, sticky="e")

        CTkButton(btn_container, text="แสดงสรุปการคำนวณ", command=self._show_calculation_summary, 
                  fg_color=self.theme["primary"]).pack(side="left", padx=5)
        
        CTkButton(btn_container, text="Export Excel", command=self._on_export_excel, 
                  fg_color="#16A34A", hover_color="#15803D").pack(side="left", padx=5)

        # --- ส่วน C: สรุปสินค้า ---
        self._create_product_summary()

        # --- ส่วน D: หมายเหตุ ---
        notes_frame = CTkFrame(self.content_frame, fg_color="transparent")
        notes_frame.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        CTkLabel(notes_frame, text="หมายเหตุ:", font=('Roboto', 12, 'bold')).pack(anchor="w")
        notes_text = CTkTextbox(notes_frame, height=50, font=('Roboto', 12), state="normal", fg_color="#F3F4F6", text_color="#333")
        notes_text.insert("1.0", self.payout_log_data.get('notes') or "")
        notes_text.configure(state="disabled")
        notes_text.pack(fill="x")

        # --- ส่วน E: ตารางรายการ SO ---
        self.content_frame.grid_rowconfigure(4, weight=1) # ให้แถวตารางขยาย
        so_list_frame = CTkScrollableFrame(self.content_frame, label_text="รายการ SO ทั้งหมดในรอบนี้")
        so_list_frame.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

        if self.payout_log_data:
            so_list_df = self._prepare_so_dataframe()
            self._create_so_list_view(so_list_frame, so_list_df)

    def _get_info_text(self):
        """Helper สร้างข้อความ Info"""
        if not self.payout_log_data: return "กำลังโหลด..."
        
        ts = self.payout_log_data.get('timestamp')
        date_str = pd.to_datetime(ts).strftime('%d/%m/%Y %H:%M') if ts else "N/A"
        sale_name = self.payout_log_data.get('sale_key','-')
        plan = self.payout_log_data.get('plan_name','-')
        
        # ดึงยอด Net มาโชว์ตรงนี้ด้วย
        net = float(self.payout_log_data.get('net_commission') or 0.0)
        
        return f"วันที่จ่าย: {date_str} | Sale: {sale_name} | Plan: {plan} | Net: {net:,.2f} บาท"
        
    def _create_header_statistics(self):
        """สร้างการ์ดแสดงสถิติ (Updated: เพิ่มยอดโอนสุทธิ)"""
        stats_frame = CTkFrame(self.content_frame, fg_color="transparent")
        stats_frame.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")
        # ปรับให้มี 5 คอลัมน์ (Sales, GP, Margin, Net, Count)
        stats_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        conn = None
        try:
            # 1. หาค่า Net Commission (ยอดโอนสุทธิ)
            net_commission = 0.0
            try:
                # พยายามดึงจาก JSON Summary ก่อน (เผื่อมีการปรับปรุงตัวเลข)
                summary_data = self.payout_log_data.get('summary_data_json')
                if isinstance(summary_data, str): summary_data = json.loads(summary_data)
                
                found_in_json = False
                if isinstance(summary_data, list):
                    for item in summary_data:
                        desc = item.get('description', '')
                        if 'Net' in desc or 'สุทธิ' in desc or 'หลังหัก' in desc:
                            net_commission = float(item.get('value') or 0.0)
                            found_in_json = True
                            break
                
                # ถ้าไม่เจอใน JSON ให้ใช้ค่าจาก Column โดยตรง
                if not found_in_json:
                    net_commission = float(self.payout_log_data.get('net_commission') or 0.0)
            except:
                net_commission = float(self.payout_log_data.get('net_commission') or 0.0)

            # 2. ดึงข้อมูลยอดขายรวมจาก Database
            if not self.payout_log_data or not self.payout_log_data.get('so_ids_json'): return
            so_id_list = json.loads(self.payout_log_data['so_ids_json'])
            if not so_id_list: return
            
            placeholders = ', '.join(['%s'] * len(so_id_list))
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                query = f"""
                    SELECT SUM(final_sales_amount) as total_sales, SUM(final_gp) as total_gp, COUNT(so_number) as total_so
                    FROM commissions WHERE id IN ({placeholders})
                """
                cursor.execute(query, tuple(so_id_list))
                res = cursor.fetchone()
                
                total_sales = res['total_sales'] or 0
                total_gp = res['total_gp'] or 0
                total_so = res['total_so'] or 0
                avg_margin = (total_gp / total_sales * 100) if total_sales > 0 else 0

                # สร้างการ์ด 5 ใบ (เพิ่ม Net Payout เป็นใบที่ 4 สีฟ้าเด่นๆ)
                self._create_stat_card(stats_frame, 0, "Total Sales", f"{total_sales:,.2f}", "#FFFFFF")
                self._create_stat_card(stats_frame, 1, "Total GP", f"{total_gp:,.2f}", "#10B981") # เขียว
                self._create_stat_card(stats_frame, 2, "Avg. Margin", f"{avg_margin:.2f}%", "#F59E0B") # ส้ม
                self._create_stat_card(stats_frame, 3, "ยอดโอนสุทธิ (Net)", f"{net_commission:,.2f}", "#3B82F6") # ฟ้า (พระเอก)
                self._create_stat_card(stats_frame, 4, "Total SO", f"{total_so} ใบ", "#FFFFFF")

        except Exception as e:
            print(f"Stats Error: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_stat_card(self, parent, col, title, value, text_color):
        """Helper สร้างการ์ดเล็กๆ"""
        frame = CTkFrame(parent, corner_radius=10, border_width=1, border_color="#666")
        frame.grid(row=0, column=col, padx=5, pady=5, sticky="ew")
        CTkLabel(frame, text=title, font=("Arial", 12, "bold"), text_color="gray").pack(pady=(10,0))
        CTkLabel(frame, text=value, font=("Arial", 20, "bold"), text_color=text_color).pack(pady=(0,10))

    def _create_product_summary(self):
        """สร้างตารางสรุปสินค้า (Product Mix) พร้อมปุ่ม Debug"""
        frame = CTkFrame(self.content_frame)
        frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        
        # --- Header Section (Title + Debug Button) ---
        header_frame = CTkFrame(frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=5)
        
        CTkLabel(header_frame, text="สรุปยอดขายแยกตามสินค้า (Product Mix)", font=("Arial", 16, "bold")).pack(side="left")
        
        # [NEW] ปุ่มกดดูไส้ใน
        debug_btn = CTkButton(header_frame, text="🔍 ดูที่มาตัวเลข (Debug)", width=120, height=24,
                              font=("Arial", 12), fg_color="#64748B", hover_color="#475569",
                              command=self._open_product_debug_window) # <--- เรียกฟังก์ชันใหม่
        debug_btn.pack(side="right")
        # ---------------------------------------------

        table_frame = CTkScrollableFrame(frame, height=150, fg_color="transparent")
        table_frame.pack(fill="x", padx=5, pady=5)
        table_frame.grid_columnconfigure(0, weight=3)
        table_frame.grid_columnconfigure(1, weight=1)
        table_frame.grid_columnconfigure(2, weight=1)
        table_frame.grid_columnconfigure(3, weight=1)
        
        headers = ["ชื่อสินค้า", "จำนวนรวม", "ยอดขายรวม (บาท)", "% Share"]
        for idx, h in enumerate(headers):
            CTkLabel(table_frame, text=h, font=("Arial", 12, "bold")).grid(row=0, column=idx, sticky="w" if idx==0 else "e", padx=10)

        conn = None
        try:
            if not self.payout_log_data or not self.payout_log_data.get('so_ids_json'): return
            so_id_list = json.loads(self.payout_log_data['so_ids_json'])
            if not so_id_list: return
            placeholders = ', '.join(['%s'] * len(so_id_list))

            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # Query นี้คือการ "รวมยอด (Group By)"
                cursor.execute(f"""
                    SELECT product_name, SUM(quantity) as qty, SUM(total_price) as amt 
                    FROM purchase_order_items 
                    WHERE purchase_order_id IN (
                        SELECT id FROM purchase_orders WHERE so_number IN (
                            SELECT so_number FROM commissions WHERE id IN ({placeholders})
                        )
                    )
                    GROUP BY product_name ORDER BY amt DESC
                """, tuple(so_id_list))
                
                products = cursor.fetchall()
                grand_total = sum(p['amt'] for p in products) if products else 1

                for i, p in enumerate(products):
                    r = i + 1
                    share = (p['amt'] / grand_total * 100)
                    CTkLabel(table_frame, text=p['product_name']).grid(row=r, column=0, sticky="w", padx=10, pady=2)
                    CTkLabel(table_frame, text=f"{p['qty']:,.0f}").grid(row=r, column=1, sticky="e", padx=10, pady=2)
                    CTkLabel(table_frame, text=f"{p['amt']:,.2f}").grid(row=r, column=2, sticky="e", padx=10, pady=2)
                    CTkLabel(table_frame, text=f"{share:.1f}%").grid(row=r, column=3, sticky="e", padx=10, pady=2)
        except Exception as e:
            print(f"Prod Summary Err: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _load_data(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM commission_payout_logs WHERE id = %s", (self.payout_id,))
                self.payout_log_data = cursor.fetchone()
                if not self.payout_log_data:
                    messagebox.showerror("Error", "ไม่พบข้อมูล Payout ID นี้", parent=self)
                    self.destroy()
        except Exception as e:
            messagebox.showerror("DB Error", f"{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _prepare_so_dataframe(self):
        """
        เตรียมข้อมูลสำหรับแสดงในตาราง (ฉบับแก้ไข: บังคับคำนวณ Margin ใหม่โดยคูณ 1.03 เสมอ เพื่อให้ตรงกับ Popup)
        """
        if not self.payout_log_data or not self.payout_log_data.get('so_ids_json'): 
            return pd.DataFrame()
            
        try:
            so_id_list = json.loads(self.payout_log_data['so_ids_json'])
            if not so_id_list: return pd.DataFrame()

            placeholders = ', '.join(['%s']*len(so_id_list))
            
            # [🔥 แก้ไข Query] ตัด total_po_shipping_cost ออกเพื่อแก้ Error
            query = f"""
                SELECT so_number, sales_service_amount, final_cost_amount, cost_multiplier, 
                       difference_amount 
                FROM commissions 
                WHERE id IN ({placeholders}) 
                ORDER BY so_number DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=tuple(so_id_list))
            
            if df.empty: return pd.DataFrame()
            
            # 1. แปลงข้อมูลเป็นตัวเลข
            sales = pd.to_numeric(df['sales_service_amount'], errors='coerce').fillna(0)
            cost = pd.to_numeric(df['final_cost_amount'], errors='coerce').fillna(0)
            diff = pd.to_numeric(df['difference_amount'], errors='coerce').fillna(0)
            
            # 2. ดึงตัวคูณ (ถ้าไม่มีใน DB ให้ใช้ 1.03)
            if 'cost_multiplier' in df.columns:
                mult = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
                # ถ้าตัวคูณเป็น 0 หรือ 1 (หลุดมา) ให้บังคับเป็น 1.03 ไว้ก่อนเพื่อความปลอดภัยในมุมมอง HR
                mult = mult.apply(lambda x: 1.03 if x < 1.01 else x)
            else:
                mult = 1.03

            # 3. [🔥 สำคัญ] คำนวณกำไรใหม่สดๆ (เหมือนใน Popup)
            # Profit = Sales - (Cost * Multiplier) + Diff
            profit = (sales - (cost * mult)) + diff
            
            # 4. คำนวณ Margin (%)
            df['calculated_margin'] = (profit / sales.replace(0, np.nan)) * 100
            df['calculated_margin'] = df['calculated_margin'].fillna(0.0)
            
            # 5. กำหนดสถานะ (Normal / Below Tier)
            df['status'] = df['calculated_margin'].apply(lambda x: 'Normal' if x >= 10.0 else 'Below Tier')
            
            return df[['so_number', 'status', 'sales_service_amount', 'calculated_margin']].rename(columns={
                'so_number': 'SO Number', 
                'status': 'สถานะ', 
                'sales_service_amount': 'ยอดขายสินค้า', 
                'calculated_margin': 'Margin (%)'
            })

        except Exception as e:
            print(f"Prepare DF Error: {e}")
            return pd.DataFrame()
            
    def _create_so_list_view(self, parent, df):
        if df.empty: 
            CTkLabel(parent, text="ไม่มีข้อมูล").pack(pady=10)
            return
        
        style = ttk.Style(parent)
        style.theme_use("clam")
        style.configure("SOList.Treeview.Heading", font=('Roboto', 12, 'bold'), background="#E0E7FF")
        tree = ttk.Treeview(parent, columns=list(df.columns), show='headings', style="SOList.Treeview")
        tree.pack(fill="both", expand=True)
        
        tree.tag_configure('Normal', background='#F0FDF4')
        tree.tag_configure('Below Tier', background='#FEF2F2')

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor='center')

        for _, row in df.iterrows():
            vals = [f"{v:,.2f}" if isinstance(v, float) else v for v in row]
            tree.insert("", "end", values=vals, tags=(row['สถานะ'],))
        
        # Double click to view details (ถ้ามี class SODetailViewer)
        tree.bind("<Double-1>", lambda e: SODetailViewer(self, self.app_container, tree.item(tree.focus(), "values")[0]))

    def _show_calculation_summary(self):
        """เปิดหน้าต่างสรุปการคำนวณ (PayoutCalculationViewer)"""
        try:
            # เรียกใช้ Class PayoutCalculationViewer
            PayoutCalculationViewer(
                master=self, 
                app_container=self.app_container, 
                payout_id=self.payout_id
            )
        except NameError:
            # กันเหนียวเผื่อยังไม่ได้ประกาศ Class หรือ Import
            messagebox.showerror("Error", "ไม่พบ Class PayoutCalculationViewer กรุณาตรวจสอบว่ามีโค้ดส่วนนี้ในไฟล์แล้ว", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถเปิดหน้าต่างสรุปได้: {e}", parent=self)

    def _on_export_excel(self):
        # โค้ดเดิมสำหรับ export
        try:
            export_payout_so_list_to_excel(self, self.app_container, self.payout_id)
        except Exception as e:
            messagebox.showerror("Error", f"{e}")

    def _open_product_debug_window(self):
        """เปิดหน้าต่างแสดงข้อมูลดิบ (Raw Data) พร้อมวิเคราะห์ส่วนต่างราคา"""
        
        debug_win = CTkToplevel(self)
        debug_win.title("Debug: เจาะลึกที่มาของยอดขาย (Advanced Breakdown)")
        debug_win.geometry("1300x650") # ขยายให้กว้างขึ้น
        
        # Header
        header_frame = CTkFrame(debug_win, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=10)
        CTkLabel(header_frame, text="ตารางตรวจสอบความผิดปกติของราคา (Price Auditing)", font=("Arial", 18, "bold")).pack()
        CTkLabel(header_frame, text="เช็คว่า ราคารวม (Total) ตรงกับ จำนวน x ราคาต่อหน่วย หรือไม่?", text_color="gray").pack()

        # ตาราง
        tree_frame = CTkFrame(debug_win)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # เพิ่มคอลัมน์สำหรับการตรวจสอบ
        columns = [
            "SO Number", "PO Number", "สินค้า", 
            "จำนวน (A)", "ราคา/หน่วย (B)", 
            "ราคาสูตร (A x B)", "ราคาจริงใน DB", 
            "ส่วนต่าง (ลด/เพิ่ม)", "% ที่หายไป"
        ]
        
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(fill="both", expand=True)

        # จัด Style
        style = ttk.Style()
        style.configure("Debug.Treeview.Heading", font=('Arial', 10, 'bold'))
        tree.configure(style="Debug.Treeview")

        # Config Columns
        for col in columns:
            tree.heading(col, text=col)
            width = 250 if "สินค้า" in col else 100
            anchor = "e" if any(x in col for x in ["จำนวน", "ราคา", "ส่วนต่าง", "%"]) else "w"
            tree.column(col, width=width, anchor=anchor)

        # Tag สี
        tree.tag_configure('match', background='#F0FDF4') # สีเขียว (ตรงเป๊ะ)
        tree.tag_configure('mismatch', background='#FEF2F2', foreground="#DC2626") # สีแดง (ราคาไม่ตรงสูตร)
        tree.tag_configure('total', background="#E0E7FF", font=("Arial", 10, "bold"))

        conn = None
        try:
            if not self.payout_log_data or not self.payout_log_data.get('so_ids_json'): return
            so_id_list = json.loads(self.payout_log_data['so_ids_json'])
            placeholders = ', '.join(['%s'] * len(so_id_list))

            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                query = f"""
                    SELECT 
                        c.so_number, po.po_number, poi.product_name,
                        poi.quantity, poi.unit_price, poi.total_price
                    FROM purchase_order_items poi
                    JOIN purchase_orders po ON poi.purchase_order_id = po.id
                    JOIN commissions c ON po.so_number = c.so_number
                    WHERE c.id IN ({placeholders})
                    ORDER BY c.so_number, poi.product_name
                """
                cursor.execute(query, tuple(so_id_list))
                raw_rows = cursor.fetchall()

                grand_db_total = 0
                grand_calc_total = 0

                for row in raw_rows:
                    qty = float(row['quantity'] or 0)
                    unit_price = float(row['unit_price'] or 0)
                    db_total = float(row['total_price'] or 0)
                    
                    # 1. คำนวณราคาตามสูตร (Quantity x Unit Price)
                    calc_total = qty * unit_price
                    
                    # 2. หาความแตกต่าง
                    diff = db_total - calc_total
                    
                    # 3. หา % ส่วนลด (ถ้ามี)
                    if calc_total > 0:
                        percent_diff = (diff / calc_total) * 100
                    else:
                        percent_diff = 0.0

                    # 4. กำหนดสี (ถ้าต่างกันเกิน 0.01 บาท ให้แดง)
                    tag = 'match' if abs(diff) < 0.01 else 'mismatch'

                    vals = (
                        row['so_number'],
                        row['po_number'],
                        row['product_name'],
                        f"{qty:,.2f}",
                        f"{unit_price:,.2f}",
                        f"{calc_total:,.2f}",   # ราคาสูตร
                        f"{db_total:,.2f}",     # ราคาจริง
                        f"{diff:,.2f}",         # ส่วนต่าง
                        f"{percent_diff:,.1f}%" # %
                    )
                    tree.insert("", "end", values=vals, tags=(tag,))
                    
                    grand_db_total += db_total
                    grand_calc_total += calc_total

                # Footer
                diff_total = grand_db_total - grand_calc_total
                tree.insert("", "end", values=("", "", "=== GRAND TOTAL ===", "", "", 
                                               f"{grand_calc_total:,.2f}", 
                                               f"{grand_db_total:,.2f}", 
                                               f"{diff_total:,.2f}", ""), tags=('total',))

        except Exception as e:
            messagebox.showerror("Debug Error", f"{e}", parent=debug_win)
        finally:
            if conn: self.app_container.release_connection(conn)
            
        debug_win.transient(self)
        debug_win.grab_set()

class ComparisonDetailViewer(CTkToplevel):
    def __init__(self, master, detail_df):
        super().__init__(master)
        self.title("รายละเอียดผลการเปรียบเทียบ")
        self.geometry("1100x500")
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_frame = CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # ใช้ฟังก์ชันสร้างตารางสวยงามที่คัดลอกมา
        self.theme = master.app_container.THEME["hr"]
        self.header_font_table = master.header_font_table
        self.entry_font = CTkFont(size=14)
        self._create_styled_dataframe_table(main_frame, detail_df)

        self.transient(master)
        self.grab_set()

    def _create_styled_dataframe_table(self, parent, df, label_text="", on_row_click=None, status_colors=None, status_column=None):
        for widget in parent.winfo_children():
            widget.destroy()
        if df is None or df.empty:
            CTkLabel(parent, text=f"ไม่พบข้อมูลสำหรับ '{label_text}'").pack(pady=20)
            return
        
        # --- START: โค้ดส่วนที่แก้ไข ---
        # 1. สร้าง Dictionary สำหรับแปลงชื่อหัวข้อ
        header_map = {
            'so_number': 'SO Number',
            'sales_service_amount': 'ยอดขาย (ระบบ)',
            'sales_uploaded': 'ยอดขาย (Express)',
            'cost_db': 'ต้นทุน (ระบบ)',
            'cost_uploaded': 'ต้นทุน (Express)',
            'margin_db': 'Margin (ระบบ) %',
            'margin_uploaded': 'Margin (Express) %',
            'ผลต่างยอดขาย': 'ผลต่างยอดขาย',
            'ผลต่างต้นทุน': 'ผลต่างต้นทุน',
            'สถานะ': 'สถานะ'
        }
        
        # 2. เปลี่ยนชื่อคอลัมน์ใน DataFrame ก่อนนำไปแสดงผล
        df_display = df.rename(columns=header_map)
        # --- END: สิ้นสุดโค้ดส่วนที่แก้ไข ---

        container = CTkFrame(parent, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        if label_text:
            CTkLabel(container, text=label_text, font=self.header_font_table).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        tree_frame = CTkFrame(container, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = df_display.columns.tolist()
        
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Custom.Treeview.Heading", 
                        font=self.header_font_table, 
                        background=self.theme["primary"],
                        foreground="white",
                        relief="flat")
        style.map("Custom.Treeview.Heading",
                background=[('active', self.theme["header"])])
        
        style.configure("Custom.Treeview", 
                        rowheight=28, 
                        font=self.entry_font
                        )
        
        style.map("Custom.Treeview",
                background=[('selected', self.theme["primary"])],
                foreground=[('selected', "white")])

        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Custom.Treeview", height=15)
        tree.grid(row=0, column=0, sticky="nsew")

        # กำหนดสีตามสถานะที่อาจมี
        status_colors_map = {
            "ผ่านเกณฑ์": "#BBF7D0", "ยอดขายต่ำกว่า Express": "#FECACA",
            "ต้นทุนต่ำกว่า Express": "#FEF08A", "มีในไฟล์, ไม่มีในระบบ": "#FECACA",
            "มีในระบบ, ไม่มีในไฟล์": "#FEF08A", "ข้อมูลไม่ตรงกัน": "#FED7AA",
            "กำไรดี": "#BBF7D0", "กำไรน้อย": "#FEF08A",
            "ขาดทุน": "#FECACA", "ยืนยันแล้ว (รอผล)": "#E5E7EB"
        }
        for status, color in status_colors_map.items():
            tree.tag_configure(status, background=color)

        for col_id in columns:
            header_text = col_id
            tree.heading(col_id, text=header_text)
            width = 150
            if "ยอดขาย" in col_id or "ต้นทุน" in col_id:
                tree.column(col_id, width=120, anchor='e')
            elif "Margin" in col_id:
                tree.column(col_id, width=100, anchor='e')
            else:
                tree.column(col_id, width=width, anchor='w')

        for index, row in df_display.iterrows():
            # ดึงค่า 'สถานะ' จาก df_display เพื่อใช้กำหนดสี
            status_value = row.get('สถานะ', '')
            tags = [status_value] if status_value in status_colors_map else []
            
            values = []
            for col_name in columns:
                value = row[col_name]
                if pd.notna(value):
                    if isinstance(value, (float, np.floating)): values.append(f"{value:,.2f}")
                    else: values.append(str(value))
                else:
                    values.append("")
            
            # ใช้ so_number จาก DataFrame ต้นฉบับ (df) เพื่อเป็น ID ที่ไม่ซ้ำกัน
            iid_value = df.iloc[index]['SO Number']
            tree.insert("", "end", values=values, tags=tuple(tags), iid=str(iid_value))
        
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        if on_row_click: 
            tree.bind("<Double-1>", lambda e: on_row_click(e, tree, df))

class ComparisonHistoryWindow(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.master_screen = master
        self.all_logs_df = pd.DataFrame()
        self.detail_df = pd.DataFrame()
        self.current_view = 'summary'
        
        self.theme = self.app_container.THEME["hr"]
        self.header_font_table = ctk.CTkFont(size=14, weight="bold")
        self.entry_font = ctk.CTkFont(size=14)

        # --- START: เพิ่มตัวแปรสำหรับฟิลเตอร์เดือน/ปี ---
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")
        # --- END ---

        self.title("ประวัติการเปรียบเทียบข้อมูล")
        self.geometry("1200x700")
        
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- ส่วนควบคุมและฟิลเตอร์ ---
        self.control_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.back_button = ctk.CTkButton(self.control_frame, text="◀ กลับไปหน้าสรุป", command=self._show_summary_view)
        self.back_button.pack(side="left", padx=5)
        self.back_button.pack_forget() 

        # --- START: เพิ่ม Dropdown สำหรับเดือนและปี ---
        month_options = ["ทุกเดือน"] + self.thai_months
        self.month_menu = ctk.CTkOptionMenu(self.control_frame, variable=self.month_var, values=month_options)
        self.month_menu.pack(side="left", padx=(10, 5))

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        self.year_menu = ctk.CTkOptionMenu(self.control_frame, variable=self.year_var, values=year_options)
        self.year_menu.pack(side="left", padx=5)
        # --- END ---

        self.search_entry = ctk.CTkEntry(self.control_frame, placeholder_text="ค้นหา...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(10, 5))
        self.search_entry.bind("<KeyRelease>", lambda e: self._apply_filter())
        
        ctk.CTkButton(self.control_frame, text="ค้นหา", width=80, command=self._apply_filter).pack(side="left")

        # --- Label แสดงสถานะ ---
        status_bar_frame = ctk.CTkFrame(self, fg_color="transparent")
        status_bar_frame.grid(row=1, column=0, padx=10, sticky="ew")

        self.title_label = ctk.CTkLabel(status_bar_frame, text="ภาพรวมการเปรียบเทียบ (ดับเบิลคลิกเพื่อดูรายละเอียด)", font=ctk.CTkFont(size=16, weight="bold"))
        self.title_label.pack(side="left")

        self.count_label = ctk.CTkLabel(status_bar_frame, text="กำลังโหลด...", text_color="gray")
        self.count_label.pack(side="right")

        # --- Frame สำหรับตาราง ---
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.grid(row=2, column=0, padx=10, pady=(0,10), sticky="nsew")
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        self.after(50, self._load_and_display_history)
        self.transient(master)
        self.grab_set()

    def _show_summary_view(self):
        """สลับกลับไปแสดงผลหน้าสรุป Log"""
        self.current_view = 'summary'
        self.title_label.configure(text="ภาพรวมการเปรียบเทียบ (ดับเบิลคลิกเพื่อดูรายละเอียด)")
        self.search_entry.configure(placeholder_text="ค้นหาไฟล์หรือผู้จัดทำ...")
        self.back_button.pack_forget()
        self.month_menu.pack(side="left", padx=(10, 5))
        self.year_menu.pack(side="left", padx=5)
        self._apply_filter() # <--- แก้ไขให้เรียกใช้ _apply_filter แทน

    def _show_log_details(self, log_id):
        """สลับไปแสดงผลหน้ารายละเอียดของ Log ที่เลือก"""
        self.current_view = 'detail'
        log_row = self.all_logs_df[self.all_logs_df['ID'] == log_id].iloc[0]
        timestamp = pd.to_datetime(log_row['เวลาที่ทำรายการ']).strftime('%Y-%m-%d %H:%M')
        
        self.title_label.configure(text=f"รายละเอียด Log ID: {log_id} (เวลา: {timestamp}) - ดับเบิลคลิก SO เพื่อตรวจสอบ")
        self.search_entry.configure(placeholder_text="ค้นหา SO...")
        # แสดงปุ่ม Back และซ่อนฟิลเตอร์เดือน/ปี
        self.back_button.pack(side="left", padx=5)
        self.month_menu.pack_forget()
        self.year_menu.pack_forget()

        details_list = log_row['detail_json_hidden']
        if details_list and isinstance(details_list, list):
            self.detail_df = pd.DataFrame(details_list)
        else:
            self.detail_df = pd.DataFrame()
        self._populate_treeview(self.detail_df)

    def _load_and_display_history(self):
        """โหลดข้อมูล Log ทั้งหมด และแปลง JSON ให้เป็น DataFrame ที่ใช้งานได้"""
        try:
            query = "SELECT id, timestamp, hr_user_key, salesperson_filter, source_info, summary_json, detail_json FROM comparison_logs ORDER BY timestamp DESC LIMIT 200"
            logs_df = pd.read_sql_query(query, self.app_container.pg_engine)

            if logs_df.empty:
                self.all_logs_df = pd.DataFrame()
                self._populate_treeview(self.all_logs_df) # <--- แก้ไขโดยส่ง df เข้าไป
                return

            def safe_json_normalize(series):
                processed_data = []
                for item in series:
                    if isinstance(item, str):
                        try: processed_data.append(json.loads(item))
                        except json.JSONDecodeError: processed_data.append({})
                    elif isinstance(item, dict):
                        processed_data.append(item)
                    else:
                        processed_data.append({})
                return pd.json_normalize(processed_data)

            summary_df = safe_json_normalize(logs_df['summary_json'])
            logs_df = pd.concat([logs_df.drop(columns=['summary_json']), summary_df], axis=1)
            
            self.all_logs_df = logs_df.rename(columns={
                'id': 'ID', 'timestamp': 'เวลาที่ทำรายการ', 'hr_user_key': 'ทำโดย (HR)',
                'salesperson_filter': 'ข้อมูลของเซลส์', 'source_info': 'ไฟล์/แหล่งข้อมูล',
                'total_records': 'ยอดรวม', 'matched_records': 'ตรงกัน', 'diff_records': 'แตกต่าง',
                'detail_json': 'detail_json_hidden'
            })
            self._show_summary_view()

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถโหลดประวัติได้: {e}", parent=self)
            traceback.print_exc()

    def _apply_filter(self, *args):
        """กรองข้อมูลใน DataFrame ตามเงื่อนไขที่เลือกและอัปเดตตาราง"""
        if self.all_logs_df.empty:
            self._populate_treeview(self.all_logs_df)
            return
            
        df = self.all_logs_df.copy()

        # --- START: เพิ่ม Logic การกรองด้วยเดือนและปี ---
        df['เวลาที่ทำรายการ'] = pd.to_datetime(df['เวลาที่ทำรายการ'])
        
        selected_month_str = self.month_var.get()
        if selected_month_str != "ทุกเดือน":
            month_num = self.thai_month_map[selected_month_str]
            df = df[df['เวลาที่ทำรายการ'].dt.month == month_num]

        selected_year_str = self.year_var.get()
        if selected_year_str != "ทุกปี":
            year_num = int(selected_year_str)
            df = df[df['เวลาที่ทำรายการ'].dt.year == year_num]
        # --- END ---
            
        search_term = self.search_entry.get().strip().lower()
        if search_term:
             if self.current_view == 'summary':
                df = df[
                    df['ทำโดย (HR)'].str.lower().str.contains(search_term, na=False) |
                    df['ไฟล์/แหล่งข้อมูล'].str.lower().str.contains(search_term, na=False)
                ]
             else: # detail view
                if not self.detail_df.empty:
                    df = self.detail_df[self.detail_df['SO Number'].str.lower().str.contains(search_term, na=False)]

        self._populate_treeview(df)

    def _populate_treeview(self, df):
        # ... (โค้ดส่วนนี้เหมือนเดิม ไม่มีการเปลี่ยนแปลง) ...
        for widget in self.tree_frame.winfo_children():
            widget.destroy()

        self.count_label.configure(text=f"พบ {len(df)} รายการ")
        if df.empty:
            ctk.CTkLabel(self.tree_frame, text="ไม่พบข้อมูล").pack(pady=20)
            return

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("History.Treeview.Heading", font=self.header_font_table, background=self.theme['primary'], foreground="white", relief="flat")
        style.configure("History.Treeview", rowheight=28, font=self.entry_font, fieldbackground="#FFFFFF")
        style.map("History.Treeview", background=[('selected', self.theme["header"])])
        
        tree = ttk.Treeview(self.tree_frame, show="headings", style="History.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        if self.current_view == 'summary':
            columns = ['ID', 'เวลาที่ทำรายการ', 'ทำโดย (HR)', 'ข้อมูลของเซลส์', 'ไฟล์/แหล่งข้อมูล', 'ยอดรวม', 'ตรงกัน', 'แตกต่าง']
            tree['columns'] = columns
            for col in columns:
                width = 80 if col in ['ID', 'ยอดรวม', 'ตรงกัน', 'แตกต่าง'] else 180
                tree.heading(col, text=col)
                tree.column(col, width=width, anchor='w')
            
            tree.tag_configure('matched_ok', background='#F0FDF4')
            tree.tag_configure('mismatched', background='#FEFCE8')

            for _, row in df.iterrows():
                diff_count = row.get('แตกต่าง', 0)
                tag = 'mismatched' if pd.to_numeric(diff_count, errors='coerce') > 0 else 'matched_ok'
                timestamp = pd.to_datetime(row['เวลาที่ทำรายการ']).strftime('%Y-%m-%d %H:%M')
                values = (row['ID'], timestamp, row['ทำโดย (HR)'], row['ข้อมูลของเซลส์'], row['ไฟล์/แหล่งข้อมูล'], 
                          row.get('total_records', 0), row.get('matched_records', 0), diff_count)
                tree.insert("", "end", values=values, iid=row['ID'], tags=(tag,))
        else: # detail view
            columns = ['SO Number', 'ยอดขาย (ระบบ)', 'ยอดขาย (Express)', 'ต้นทุน (ระบบ)', 'ต้นทุน (Express)', 'สถานะ']
            tree['columns'] = columns
            status_colors_map = {
                "ผ่านเกณฑ์": "#F0FDF4", "ยอดขายต่ำกว่า Express": "#FEF2F2",
                "ต้นทุนต่ำกว่า Express": "#FEFCE8", "มีในไฟล์, ไม่มีในระบบ": "#FEF2F2",
                "มีในระบบ, ไม่มีในไฟล์": "#FEFCE8", "ข้อมูลไม่ตรงกัน": "#FFF7ED",
            }
            for status, color in status_colors_map.items():
                tree.tag_configure(status, background=color)

            for col in columns:
                anchor = 'e' if 'ยอด' in col or 'ต้นทุน' in col else 'w'
                width = 150
                tree.heading(col, text=col)
                tree.column(col, anchor=anchor, width=width)
            
            for _, row in df.iterrows():
                tag = row.get('สถานะ', '')
                tags_tuple = (tag,) if tag else ()
                values = (
                    row.get('SO Number'), f"{row.get('ยอดขาย (ระบบ)', 0):,.2f}", f"{row.get('ยอดขาย (Express)', 0):,.2f}",
                    f"{row.get('ต้นทุน (ระบบ)', 0):,.2f}", f"{row.get('ต้นทุน (Express)', 0):,.2f}", row.get('สถานะ')
                )
                unique_iid = f"{row.get('log_id')}-{row.get('SO Number')}"
                tree.insert("", "end", values=values, tags=tags_tuple, iid=unique_iid)
        
        tree.bind("<Double-1>", self._on_row_double_click)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)

    def _on_row_double_click(self, event):
        tree = event.widget
        selected_iid = tree.focus()
        if not selected_iid: return

        if self.current_view == 'summary':
            log_id = int(tree.item(selected_iid, "values")[0])
            self._show_log_details(log_id)
        else: # detail view
            so_number = tree.item(selected_iid, "values")[0]
            self.master_screen._open_verification_window(so_number)
# hr_windows.py (เพิ่มคลาสใหม่นี้ต่อท้ายไฟล์)

class PayoutCalculationViewer(CTkToplevel):
    def __init__(self, master, app_container, payout_id):
        super().__init__(master)
        self.app_container = app_container
        self.payout_id = payout_id
        
        self.theme = self.app_container.THEME["hr"]
        self.header_font_table = CTkFont(size=12, weight="bold")
        self.entry_font = CTkFont(size=12)

        self.title(f"ที่มาการคำนวณค่าคอม (Payout ID: {payout_id})")
        self.geometry("1100x750") 
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # --- Header ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.summary_label = CTkLabel(top_frame, text="กำลังโหลดข้อมูล...", font=CTkFont(size=16, weight="bold"))
        self.summary_label.pack(side="left")

        # --- Tab View ---
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        
        self.tab_steps = self.tab_view.add("ขั้นตอนการคำนวณ")  
        self.tab_so = self.tab_view.add("รายละเอียดตาม SO")    

        self.after(100, self._load_and_display_data) 
        
        self.transient(master)
        self.grab_set()

    def _load_and_display_data(self):
        try:
            # 1. ดึงข้อมูลจาก Log
            query = """
                SELECT summary_json, so_numbers_json, sale_key, plan_name, timestamp, detail_json 
                FROM commission_payout_logs 
                WHERE id = %s
            """
            log_data = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.payout_id,)).iloc[0]

            # อัปเดต Label
            ts = pd.to_datetime(log_data['timestamp']).strftime('%d/%m/%Y %H:%M')
            self.summary_label.configure(text=f"การคำนวณของ: {log_data['sale_key']} (แผน: {log_data['plan_name']}) เมื่อ {ts}")

            # 2. แกะข้อมูล JSON
            debug_df = pd.DataFrame()
            breakdown_df = pd.DataFrame()

            if 'detail_json' in log_data and log_data['detail_json']:
                try:
                    details = json.loads(log_data['detail_json'])
                    
                    if isinstance(details, dict):
                        if 'debug' in details: debug_df = pd.DataFrame(details['debug'])
                        if 'breakdown' in details: breakdown_df = pd.DataFrame(details['breakdown'])
                    elif isinstance(details, list):
                        debug_df = pd.DataFrame(details)
                        
                except Exception as e:
                    print(f"Error parsing detail_json: {e}")

            # 3. แสดงผล Tab 1: ขั้นตอนการคำนวณ
            if not debug_df.empty:
                self._populate_calc_steps_tab(self.tab_steps, debug_df)
            else:
                self._clear_frame(self.tab_steps)
                ctk.CTkLabel(self.tab_steps, text="ไม่พบข้อมูลขั้นตอนการคำนวณ (อาจเป็นข้อมูลเก่า)", font=("Arial", 16)).pack(pady=40)

            # 4. แสดงผล Tab 2: รายละเอียด SO
            if not breakdown_df.empty:
                self._populate_so_breakdown_tab(self.tab_so, breakdown_df) 
            else:
                 self._clear_frame(self.tab_so)
                 ctk.CTkLabel(self.tab_so, text="ไม่พบข้อมูล Breakdown ราย SO (อาจเป็นข้อมูลเก่า)", font=("Arial", 16)).pack(pady=20)
                 ctk.CTkButton(self.tab_so, text="คลิกเพื่อดูรายละเอียด SO ทั้งหมด (ดึงข้อมูลสด)", command=self._open_so_list_viewer).pack(pady=10)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถโหลดรายละเอียดได้: {e}", parent=self)
            self.destroy()

    def _populate_calc_steps_tab(self, tab, df):
        """แสดงข้อมูลขั้นตอนการคำนวณ (Debug Steps)"""
        self._clear_frame(tab)

        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        
        tree_frame = ctk.CTkFrame(tab, fg_color="transparent")
        tree_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style(self)
        style.theme_use("clam")
        
        header_font = ("Tahoma", 11, "bold")
        content_font = ("Tahoma", 11)

        style.configure("Steps.Treeview.Heading", font=header_font, background="#64748B", foreground="white", relief="flat")
        style.configure("Steps.Treeview", rowheight=30, font=content_font)

        tree = ttk.Treeview(tree_frame, columns=("item", "value"), show="headings", style="Steps.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        tree.heading("item", text="รายการ / ขั้นตอน")
        tree.heading("value", text="ค่า / ผลลัพธ์")
        
        tree.column("item", width=500, anchor="w")
        tree.column("value", width=200, anchor="e")

        tree.tag_configure('header', background='#E2E8F0', font=header_font)
        tree.tag_configure('separator', background='#F1F5F9')
        tree.tag_configure('highlight', background='#FEF9C3', font=header_font)
        tree.tag_configure('success', background='#DCFCE7', foreground="#166534")
        tree.tag_configure('fail', background='#FEE2E2', foreground="#991B1B")

        for _, row in df.iterrows():
            item_text = str(row.get('รายการ', '')).strip()
            value_text = str(row.get('ค่า', '')).strip()
            
            tags = []
            if item_text.startswith('##'):
                item_text = item_text.replace('##', '').strip()
                tags.append('header')
            elif item_text == '---':
                item_text = ''
                value_text = ''
                tags.append('separator')
            elif 'ยอดรวม' in item_text or 'สุทธิ' in item_text:
                tags.append('highlight')
            elif 'ผ่าน' in value_text and '✅' in value_text:
                tags.append('success')
            elif 'ไม่ผ่าน' in value_text or '❌' in value_text:
                tags.append('fail')
            
            tree.insert("", "end", values=(item_text, value_text), tags=tuple(tags))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)

    def _populate_so_breakdown_tab(self, tab, df):
        """
        แสดงรายการ SO (Breakdown) แบบ Robust พร้อมสีและ Format
        (อัปเดต: ให้รองรับชื่อคอลัมน์ใหม่ๆ จาก Business Logic ทุก Plan)
        """
        self._clear_frame(tab)

        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        tree_frame = ctk.CTkFrame(tab, fg_color="transparent")
        tree_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        if df is None or df.empty:
            ctk.CTkLabel(tree_frame, text="ไม่พบข้อมูลรายละเอียด SO").pack(pady=20)
            return

        df_display = df.copy()

        # 🔥 [จุดแก้ไข] เพิ่มคำพ้องความหมายให้ครอบคลุมทุก Plan ('ยอดขายรวม', 'ต้นทุนสินค้า', 'กำไรสุทธิ')
        col_mapping = {
            'so_number': ['so_number', 'เลขที่ SO', 'SO Number'],
            'sales': ['sales_service_amount', 'ยอดขาย', 'ยอดขาย (Base)', 'final_sales_amount', 'ยอดขายสินค้า', 'ยอดขายรวม'],
            'cost': ['final_cost_amount', 'ต้นทุน', 'ต้นทุน (Net)', 'cost', 'final_cost', 'ต้นทุนสินค้า', 'ต้นทุนรวม'],
            'profit': ['profit', 'กำไร', 'final_gp', 'กำไร (Profit)', 'กำไรสุทธิ'],
            'margin': ['margin', 'Margin (%)', 'final_margin', 'Margin %'],
            'status': ['status', 'Status', 'สถานะ'],
            'multiplier': ['cost_multiplier', 'ตัวคูณ']
        }

        for target_col, possible_cols in col_mapping.items():
            found_col = next((c for c in df_display.columns if c in possible_cols), None)
            if found_col:
                if target_col not in ['so_number', 'status']:
                    df_display[target_col] = df_display[found_col].astype(str).str.replace(',', '', regex=False)
                    df_display[target_col] = pd.to_numeric(df_display[target_col], errors='coerce').fillna(0)
                else:
                    df_display[target_col] = df_display[found_col]
            else:
                # Default Multiplier 1.03 ถ้าหาไม่เจอ
                if target_col == 'multiplier':
                     df_display[target_col] = 1.03 
                else:
                    df_display[target_col] = 0.0 if target_col not in ['so_number', 'status'] else '-'

        # --- คำนวณต้นทุนรวมตัวคูณ และตั้งชื่อหัวข้อ ---
        cost_header = "ต้นทุน (Net)"
        
        if df_display['multiplier'].max() > 1.001:
            max_mult = df_display['multiplier'].max()
            percent_add = int((max_mult - 1) * 100)
            cost_header = f"ต้นทุน (+{percent_add}%)"

            # คูณตัวเลขต้นทุนจริงๆ
            df_display['cost'] = df_display['cost'] * df_display['multiplier']
        
        # คำนวณ Profit/Margin สดๆ อีกรอบ
        df_display['profit'] = df_display['sales'] - df_display['cost']
        df_display['margin'] = df_display.apply(
            lambda x: (x['profit'] / x['sales'] * 100) if x['sales'] != 0 else 0, axis=1
        )

        final_columns = ['so_number', 'sales', 'cost', 'profit', 'margin', 'status']
        header_labels = {
            'so_number': 'เลขที่ SO', 'sales': 'ยอดขาย', 'cost': cost_header,
            'profit': 'กำไร', 'margin': 'Margin %', 'status': 'สถานะ'
        }

        style = ttk.Style(self)
        style.theme_use("clam")
        header_font = ("Tahoma", 11, "bold")
        content_font = ("Tahoma", 11)

        style.configure("Breakdown.Treeview.Heading", font=header_font, background="#3B82F6", foreground="white", relief="flat")
        style.configure("Breakdown.Treeview", rowheight=30, font=content_font)

        tree = ttk.Treeview(tree_frame, columns=final_columns, show="headings", style="Breakdown.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        tree.tag_configure('Normal', background='#DCFCE7', foreground='#166534')      
        tree.tag_configure('Below Tier', background='#FEE2E2', foreground='#991B1B')  
        tree.tag_configure('Paid', background='#E0E7FF', foreground='#3730A3')        
        tree.tag_configure('Default', background='white')

        for col in final_columns:
            anchor = 'center' if col in ['so_number', 'status'] else 'e'
            width = 150 if col == 'so_number' else 120
            tree.heading(col, text=header_labels.get(col, col))
            tree.column(col, width=width, anchor=anchor)

        for _, row in df_display.iterrows():
            values = []
            status_val = str(row['status'])
            tag = 'Default'
            if any(x in status_val for x in ['Normal', 'ผ่านเกณฑ์', '>=10']): tag = 'Normal'
            elif any(x in status_val for x in ['Below', 'ต่ำกว่า', '<']): tag = 'Below Tier'
            elif 'Paid' in status_val: tag = 'Paid'

            values.append(row['so_number'])
            values.append(f"{row['sales']:,.2f}")
            values.append(f"{row['cost']:,.2f}")
            values.append(f"{row['profit']:,.2f}")
            values.append(f"{row['margin']:,.2f}%")
            values.append(status_val)
            
            tree.insert("", "end", values=tuple(values), tags=(tag,))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)
        
        # ผูก Double Click เปิดดู SO Detail (ถ้ามี)
        if hasattr(self, '_on_so_row_double_click'):
            tree.bind("<Double-1>", lambda e: self._on_so_row_double_click(e))

    def _open_so_list_viewer(self):
        PayoutDetailWindow(master=self, app_container=self.app_container, payout_id=self.payout_id)

    def _clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()


class CalculationDetailViewer(CTkToplevel):
    def __init__(self, master, debug_df, so_breakdown_df, plan_name):
        super().__init__(master)
        self.app_container = master.app_container
        self.title(f"รายละเอียดการคำนวณ - {plan_name}")
        self.plan_name = plan_name
        self.geometry("1100x700")
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tab_view = ctk.CTkTabview(self, corner_radius=10)
        self.tab_view.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.tab_steps = self.tab_view.add("ขั้นตอนการคำนวณ")
        self.tab_so = self.tab_view.add("รายละเอียดตาม SO")

        # เรียกใช้ฟังก์ชันแสดงผล (ใช้ Logic ใหม่)
        self._populate_calc_steps_tab(self.tab_steps, debug_df)
        self._populate_so_breakdown_tab(self.tab_so, so_breakdown_df)

        self.transient(master)
        self.grab_set()
    
    def _populate_calc_steps_tab(self, tab, df):
        """แสดงข้อมูลขั้นตอนการคำนวณ"""
        self._clear_frame(tab)
        
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        
        tree_frame = ctk.CTkFrame(tab, fg_color="transparent")
        tree_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        if df is None or df.empty:
            ctk.CTkLabel(tree_frame, text="ไม่พบข้อมูลขั้นตอนการคำนวณ").pack(pady=20)
            return

        style = ttk.Style(self)
        style.theme_use("clam")
        header_font = ("Tahoma", 11, "bold")
        content_font = ("Tahoma", 11)

        style.configure("Steps.Treeview.Heading", font=header_font, background="#64748B", foreground="white", relief="flat")
        style.configure("Steps.Treeview", rowheight=30, font=content_font)

        tree = ttk.Treeview(tree_frame, columns=("item", "value"), show="headings", style="Steps.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        tree.heading("item", text="รายการ / ขั้นตอน")
        tree.heading("value", text="ค่า / ผลลัพธ์")
        tree.column("item", width=500, anchor="w")
        tree.column("value", width=200, anchor="e")

        tree.tag_configure('header', background='#E2E8F0', font=header_font)
        tree.tag_configure('separator', background='#F1F5F9')
        tree.tag_configure('highlight', background='#FEF9C3', font=header_font)
        tree.tag_configure('success', background='#DCFCE7', foreground="#166534")
        tree.tag_configure('fail', background='#FEE2E2', foreground="#991B1B")

        for _, row in df.iterrows():
            item_text = str(row.get('รายการ', '')).strip()
            value_text = str(row.get('ค่า', '')).strip()
            
            tags = []
            if item_text.startswith('##'):
                item_text = item_text.replace('##', '').strip()
                tags.append('header')
            elif item_text == '---':
                item_text = ''
                value_text = ''
                tags.append('separator')
            elif 'ยอดรวม' in item_text or 'สุทธิ' in item_text:
                tags.append('highlight')
            elif 'ผ่าน' in value_text and '✅' in value_text:
                tags.append('success')
            elif 'ไม่ผ่าน' in value_text or '❌' in value_text:
                tags.append('fail')
            
            tree.insert("", "end", values=(item_text, value_text), tags=tuple(tags))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)

    def _populate_so_breakdown_tab(self, tab, df):
        """
        แสดงรายการ SO (Breakdown) แบบ Robust พร้อมสีและ Format
        (อัปเดต: ให้รองรับชื่อคอลัมน์ใหม่ๆ จาก Business Logic ทุก Plan)
        """
        self._clear_frame(tab)

        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        tree_frame = ctk.CTkFrame(tab, fg_color="transparent")
        tree_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        if df is None or df.empty:
            ctk.CTkLabel(tree_frame, text="ไม่พบข้อมูลรายละเอียด SO").pack(pady=20)
            return

        df_display = df.copy()

        # 🔥 [จุดแก้ไข] เพิ่มคำพ้องความหมายให้ครอบคลุมทุก Plan ('ยอดขายรวม', 'ต้นทุนสินค้า', 'กำไรสุทธิ')
        col_mapping = {
            'so_number': ['so_number', 'เลขที่ SO', 'SO Number'],
            'sales': ['sales_service_amount', 'ยอดขาย', 'ยอดขาย (Base)', 'final_sales_amount', 'ยอดขายสินค้า', 'ยอดขายรวม'],
            'cost': ['final_cost_amount', 'ต้นทุน', 'ต้นทุน (Net)', 'cost', 'final_cost', 'ต้นทุนสินค้า', 'ต้นทุนรวม'],
            'profit': ['profit', 'กำไร', 'final_gp', 'กำไร (Profit)', 'กำไรสุทธิ'],
            'margin': ['margin', 'Margin (%)', 'final_margin', 'Margin %'],
            'status': ['status', 'Status', 'สถานะ'],
            'multiplier': ['cost_multiplier', 'ตัวคูณ']
        }

        for target_col, possible_cols in col_mapping.items():
            found_col = next((c for c in df_display.columns if c in possible_cols), None)
            if found_col:
                if target_col not in ['so_number', 'status']:
                    df_display[target_col] = df_display[found_col].astype(str).str.replace(',', '', regex=False)
                    df_display[target_col] = pd.to_numeric(df_display[target_col], errors='coerce').fillna(0)
                else:
                    df_display[target_col] = df_display[found_col]
            else:
                # Default Multiplier 1.03 ถ้าหาไม่เจอ
                if target_col == 'multiplier':
                     df_display[target_col] = 1.03 
                else:
                    df_display[target_col] = 0.0 if target_col not in ['so_number', 'status'] else '-'

        # --- คำนวณต้นทุนรวมตัวคูณ และตั้งชื่อหัวข้อ ---
        cost_header = "ต้นทุน (Net)"
        
        if df_display['multiplier'].max() > 1.001:
            max_mult = df_display['multiplier'].max()
            percent_add = int((max_mult - 1) * 100)
            cost_header = f"ต้นทุน (+{percent_add}%)"

            # คูณตัวเลขต้นทุนจริงๆ
            df_display['cost'] = df_display['cost'] * df_display['multiplier']
        
        # คำนวณ Profit/Margin สดๆ อีกรอบ
        df_display['profit'] = df_display['sales'] - df_display['cost']
        df_display['margin'] = df_display.apply(
            lambda x: (x['profit'] / x['sales'] * 100) if x['sales'] != 0 else 0, axis=1
        )

        final_columns = ['so_number', 'sales', 'cost', 'profit', 'margin', 'status']
        header_labels = {
            'so_number': 'เลขที่ SO', 'sales': 'ยอดขาย', 'cost': cost_header,
            'profit': 'กำไร', 'margin': 'Margin %', 'status': 'สถานะ'
        }

        style = ttk.Style(self)
        style.theme_use("clam")
        header_font = ("Tahoma", 11, "bold")
        content_font = ("Tahoma", 11)

        style.configure("Breakdown.Treeview.Heading", font=header_font, background="#3B82F6", foreground="white", relief="flat")
        style.configure("Breakdown.Treeview", rowheight=30, font=content_font)

        tree = ttk.Treeview(tree_frame, columns=final_columns, show="headings", style="Breakdown.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        tree.tag_configure('Normal', background='#DCFCE7', foreground='#166534')      
        tree.tag_configure('Below Tier', background='#FEE2E2', foreground='#991B1B')  
        tree.tag_configure('Paid', background='#E0E7FF', foreground='#3730A3')        
        tree.tag_configure('Default', background='white')

        for col in final_columns:
            anchor = 'center' if col in ['so_number', 'status'] else 'e'
            width = 150 if col == 'so_number' else 120
            tree.heading(col, text=header_labels.get(col, col))
            tree.column(col, width=width, anchor=anchor)

        for _, row in df_display.iterrows():
            values = []
            status_val = str(row['status'])
            tag = 'Default'
            if any(x in status_val for x in ['Normal', 'ผ่านเกณฑ์', '>=10']): tag = 'Normal'
            elif any(x in status_val for x in ['Below', 'ต่ำกว่า', '<']): tag = 'Below Tier'
            elif 'Paid' in status_val: tag = 'Paid'

            values.append(row['so_number'])
            values.append(f"{row['sales']:,.2f}")
            values.append(f"{row['cost']:,.2f}")
            values.append(f"{row['profit']:,.2f}")
            values.append(f"{row['margin']:,.2f}%")
            values.append(status_val)
            
            tree.insert("", "end", values=tuple(values), tags=(tag,))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)
        
        # ผูก Double Click เปิดดู SO Detail (ถ้ามี)
        if hasattr(self, '_on_so_row_double_click'):
            tree.bind("<Double-1>", lambda e: self._on_so_row_double_click(e))

    def _on_so_row_double_click(self, event):
        tree = event.widget
        selected = tree.focus()
        if not selected: return
        vals = tree.item(selected, "values")
        so_num = vals[0]
        
        try:
            # [🔥 แก้ไข] ลบบรรทัด import ออก แล้วเรียกใช้ Class ได้เลย
            # เพราะ SODetailViewer อยู่ในไฟล์นี้แล้ว
            SODetailViewer(self, self.app_container, so_num)
            
        except NameError:
            # กันเหนียว: ถ้าหาไม่เจอจริงๆ ลอง import จากไฟล์ตัวเอง
            try:
                from hr_windows import SODetailViewer
                SODetailViewer(self, self.app_container, so_num)
            except Exception as e:
                 messagebox.showerror("Error", f"ไม่สามารถเปิดหน้าต่างรายละเอียดได้: {e}")
                 
        except Exception as e:
            print(f"Error opening detail: {e}")
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")

    def _clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

class SODetailViewer(CTkToplevel):
    def __init__(self, master, app_container, so_number):
        super().__init__(master)
        self.app_container = app_container
        self.so_number = so_number
        
        self.title(f"ข้อมูล SO: {so_number} และ PO ที่เกี่ยวข้อง")
        self.geometry("800x600")
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        CTkLabel(header_frame, text=f"SO Number: {so_number}", font=CTkFont(size=18, weight="bold")).pack(side="left")

        self.main_frame = CTkScrollableFrame(self)
        self.main_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
        self.after(50, self._load_and_display_data)
        
        self.transient(master)
        self.grab_set()

    def _load_and_display_data(self):
        try:
            so_query = "SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1"
            so_df = pd.read_sql_query(so_query, self.app_container.pg_engine, params=(self.so_number,))
            
            # <<< START: แก้ไข Query ให้ดึง id ของ PO มาด้วย >>>
            po_query = "SELECT id, po_number, supplier_name, status, total_cost FROM purchase_orders WHERE so_number = %s ORDER BY timestamp DESC"
            # <<< END >>>
            po_df = pd.read_sql_query(po_query, self.app_container.pg_engine, params=(self.so_number,))

            if not so_df.empty:
                self._display_so_details(so_df.iloc[0])
            else:
                CTkLabel(self.main_frame, text="ไม่พบข้อมูล SO นี้").pack(pady=10)

            self._display_po_details(po_df)

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูลได้: {e}", parent=self)
            self.destroy()

    def _create_detail_section_frame(self, parent, title):
        frame = CTkFrame(parent, corner_radius=10, border_width=1)
        frame.pack(fill="x", pady=(10, 5), padx=5)
        frame.grid_columnconfigure(1, weight=1)
        CTkLabel(frame, text=title, font=CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=15, pady=(10, 5), sticky="w")
        return frame
    
    def _add_detail_row(self, parent, row_index, label_text, value, sub_text=""):
        if value is None or pd.isna(value):
            value_text = "-"
        elif isinstance(value, (int, float, np.floating)):
            value_text = f"{value:,.2f}"
        elif isinstance(value, (datetime, pd.Timestamp)):
            value_text = value.strftime('%d/%m/%Y')
        else:
            value_text = str(value)

        if sub_text:
            value_text += f" ({sub_text})"
            
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(
            row=row_index, column=0, padx=(15, 10), pady=4, sticky="w")
        CTkLabel(parent, text=value_text, font=CTkFont(size=14), wraplength=400, justify="left").grid(
            row=row_index, column=1, padx=(10, 15), pady=4, sticky="ew")

    def _add_detail_row_with_vat(self, parent, row_index, label_text, value, vat_option):
        # แสดงแถวข้อมูลหลักเหมือนเดิม
        self._add_detail_row(parent, row_index, label_text, value, sub_text=vat_option)
        
        # คำนวณและแสดงแถว VAT เฉพาะเมื่อเป็น "VAT"
        if vat_option == 'VAT' and isinstance(value, (int, float)):
            vat_amount = value * 0.07
            
            # สร้าง Label เยื้องเข้าไปเล็กน้อยเพื่อความสวยงาม
            CTkLabel(parent, text="  └─ ยอด VAT 7%", font=CTkFont(size=12, slant="italic"), text_color="gray50").grid(
                row=row_index + 1, column=0, padx=(25, 10), pady=(0, 4), sticky="w")
            CTkLabel(parent, text=f"{vat_amount:,.2f}", font=CTkFont(size=12, slant="italic"), text_color="gray50").grid(
                row=row_index + 1, column=1, padx=(10, 15), pady=(0, 4), sticky="w")
            return 2 # คืนค่า 2 เพื่อบอกว่าใช้ไป 2 แถว
        return 1 # คืนค่า 1 ถ้าใช้ไปแถวเดียว
  

    def _display_so_details(self, so_data):
        header_map = self.app_container.HEADER_MAP

        # --- Section 1: รายละเอียดการขาย ---
        f1 = self._create_detail_section_frame(self.main_frame, "รายละเอียดการขาย")
        self._add_detail_row(f1, 1, header_map.get('bill_date', 'วันที่เปิด SO'), so_data.get('bill_date'))
        self._add_detail_row(f1, 2, header_map.get('customer_name', 'ชื่อลูกค้า'), so_data.get('customer_name'))
        self._add_detail_row(f1, 3, header_map.get('credit_term', 'เครดิต'), so_data.get('credit_term'))

        # --- Section 2: ยอดขายและบริการ ---
        f2 = self._create_detail_section_frame(self.main_frame, "ยอดขายและบริการ")
        self._add_detail_row(f2, 1, header_map.get('sales_service_amount', 'ยอดขาย/บริการ'), so_data.get('sales_service_amount'), sub_text=so_data.get('sales_service_vat_option'))
        self._add_detail_row(f2, 2, header_map.get('cutting_drilling_fee', 'ค่าบริการตัด/เจาะ'), so_data.get('cutting_drilling_fee'), sub_text=so_data.get('cutting_drilling_fee_vat_option'))
        self._add_detail_row(f2, 3, header_map.get('other_service_fee', 'ค่าบริการอื่นๆ'), so_data.get('other_service_fee'), sub_text=so_data.get('other_service_fee_vat_option'))
        
        # --- Section 3: ค่าจัดส่ง ---
        f3 = self._create_detail_section_frame(self.main_frame, "ค่าจัดส่ง")
        self._add_detail_row(f3, 1, header_map.get('shipping_cost', 'ค่าขนส่ง'), so_data.get('shipping_cost'), sub_text=so_data.get('shipping_vat_option'))
        self._add_detail_row(f3, 2, header_map.get('relocation_cost', 'ค่าย้าย'), so_data.get('relocation_cost'))

        # --- Section 4: ค่าธรรมเนียมและส่วนลด ---
        f4 = self._create_detail_section_frame(self.main_frame, "ค่าธรรมเนียมและส่วนลด")
        self._add_detail_row(f4, 1, header_map.get('brokerage_fee', 'ค่านายหน้า'), so_data.get('brokerage_fee'))
        self._add_detail_row(f4, 2, header_map.get('giveaways', 'ของแถม'), so_data.get('giveaways'))
        self._add_detail_row(f4, 3, header_map.get('coupons', 'คูปอง'), so_data.get('coupons'))

        # --- Section 5: สรุปข้อมูลที่ยืนยันโดย HR ---
        f5 = self._create_detail_section_frame(self.main_frame, "สรุปข้อมูลที่ยืนยันโดย HR")
        self._add_detail_row(f5, 1, 'ยอดขายสุดท้าย', so_data.get('final_sales_amount'))
        self._add_detail_row(f5, 2, 'ต้นทุนสุดท้าย', so_data.get('final_cost_amount'))
        self._add_detail_row(f5, 3, 'กำไร (GP)', so_data.get('final_gp'))
        self._add_detail_row(f5, 4, 'Margin สุดท้าย (%)', so_data.get('final_margin'))
        self._add_detail_row(f5, 5, 'สถานะล่าสุด', so_data.get('status'))
    
    def _on_po_row_double_click(self, event):
        """จัดการเมื่อดับเบิลคลิกที่แถว PO (ฉบับแก้ไข: สั่งเปิดหน้าต่างเลย ไม่ต้องสร้างปุ่ม)"""
        tree = event.widget
        selected_item_iid = tree.focus()
        if not selected_item_iid:
            return
        
        try:
            item_values = tree.item(selected_item_iid, "values")
            po_id = int(float(item_values[0])) 
            
            # 🔥 จุดแก้ไข: สั่งเปิดหน้าต่าง PurchaseDetailWindow ขึ้นมาทันที
            from history_windows import PurchaseDetailWindow
            PurchaseDetailWindow(
                self.master, 
                self.app_container, 
                po_id,
                on_save_callback=self._load_and_display_data 
            )
            
        except (IndexError, ValueError, TclError) as e:
            print(f"Could not get PO ID from selected row: {e}")

    def _display_po_details(self, po_df):
        po_frame = CTkFrame(self.main_frame, corner_radius=10)
        po_frame.pack(fill="both", expand=True, padx=10, pady=10)
        po_frame.grid_rowconfigure(1, weight=1)
        po_frame.grid_columnconfigure(0, weight=1)
        
        CTkLabel(po_frame, text="Purchase Orders ที่เกี่ยวข้อง (ดับเบิลคลิกเพื่อดูรายละเอียด)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)

        if po_df.empty:
            CTkLabel(po_frame, text="- ไม่มี PO ที่เกี่ยวข้อง -").grid(row=1, column=0, pady=10)
            return

        tree_frame = CTkFrame(po_frame, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("PO.Treeview.Heading", font=CTkFont(size=12, weight="bold"))
        style.configure("PO.Treeview", rowheight=28, font=CTkFont(size=12))

        # <<< START: แก้ไขการสร้างตาราง PO >>>
        # รวม 'id' เข้ามาใน columns แต่ซ่อนไว้ไม่ให้ผู้ใช้เห็น
        columns_with_id = list(po_df.columns)
        columns_to_display = [col for col in columns_with_id if col != 'id']

        tree = ttk.Treeview(tree_frame, columns=columns_with_id, displaycolumns=columns_to_display, show="headings", style="PO.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        for col in columns_to_display: # วนลูปเฉพาะคอลัมน์ที่จะแสดงผล
            tree.heading(col, text=col.replace('_', ' ').title())
            tree.column(col, anchor='e' if 'cost' in col else 'w', width=150)

        for _, row in po_df.iterrows():
            values = [f"{v:,.2f}" if isinstance(v, (int, float)) else v for v in row]
            tree.insert("", "end", values=tuple(values))

        # ผูก Event การดับเบิลคลิกเข้ากับฟังก์ชันใหม่
        tree.bind("<Double-1>", self._on_po_row_double_click)
        # <<< END >>>

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)
    

class EditPOWindowByHR(CTkToplevel):
    def __init__(self, master, app_container, po_id, on_close_callback=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.po_id = po_id
        self.on_close_callback = on_close_callback
        self.item_widgets = []

        self.title(f"HR: แก้ไขข้อมูล PO ID: {self.po_id}")
        self.geometry("900x600")

        # --- [แก้ไข] ใช้ grid layout หลัก ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # แถวที่ 1 (ScrollFrame) จะขยาย

        # --- สร้าง UI ของฟอร์ม ---
        self._create_widgets()
        
        # --- โหลดข้อมูลมาใส่ฟอร์ม ---
        self.after(50, self._load_data)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.transient(master)
        self.grab_set()

    def _create_widgets(self):
        """สร้าง UI (ฉบับอัปเดต: ตัดช่อง Relocation ออก)"""
        
        # 1. สร้าง Main Scrollable Frame
        self.main_scroll = ctk.CTkScrollableFrame(self)
        self.main_scroll.pack(fill="both", expand=True, padx=0, pady=0)
        self.main_scroll.grid_columnconfigure(0, weight=1)

        # --- ส่วนที่ 1: ข้อมูล Header (PO) ---
        header_frame = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(3, weight=1)
        
        ctk.CTkLabel(header_frame, text="Supplier Name:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_supplier = ctk.CTkEntry(header_frame)
        self.entry_supplier.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(header_frame, text="PO Number:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.entry_po_number = ctk.CTkEntry(header_frame)
        self.entry_po_number.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        # --- ส่วนที่ 2: ข้อมูลขนส่ง (Shipping Info) ---
        shipping_grp = ctk.CTkFrame(self.main_scroll, border_width=1)
        shipping_grp.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        shipping_grp.grid_columnconfigure((1, 3), weight=1)
        
        ctk.CTkLabel(shipping_grp, text="ข้อมูลการจัดส่ง (Shipping)", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=5)

        # 2.1 เข้าสต๊อก
        ctk.CTkLabel(shipping_grp, text="[1. เข้าสต๊อก] คนขับ:", text_color="#B91C1C").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_stock_driver = ctk.CTkEntry(shipping_grp, placeholder_text="ชื่อคนขับ/ขนส่ง")
        self.entry_stock_driver.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(shipping_grp, text="ทะเบียน:", text_color="#B91C1C").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.entry_stock_plate = ctk.CTkEntry(shipping_grp, placeholder_text="ทะเบียนรถ")
        self.entry_stock_plate.grid(row=1, column=3, padx=5, pady=5, sticky="ew")

        # 2.2 เข้าไซต์
        ctk.CTkLabel(shipping_grp, text="[2. เข้าไซต์] คนขับ:", text_color="#1D4ED8").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_site_driver = ctk.CTkEntry(shipping_grp, placeholder_text="ชื่อคนขับ/ขนส่ง")
        self.entry_site_driver.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(shipping_grp, text="ทะเบียน:", text_color="#1D4ED8").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.entry_site_plate = ctk.CTkEntry(shipping_grp, placeholder_text="ทะเบียนรถ")
        self.entry_site_plate.grid(row=2, column=3, padx=5, pady=5, sticky="ew")

        # --- ส่วนที่ 3: รายการสินค้า (Items) ---
        items_frame = ctk.CTkFrame(self.main_scroll, fg_color="transparent") 
        items_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        items_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(items_frame, text="รายการสินค้า (PO Items)", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=5)
        
        header_item = ctk.CTkFrame(items_frame, fg_color="#E5E7EB", corner_radius=5)
        header_item.pack(fill="x")
        header_item.grid_columnconfigure(0, weight=4)
        header_item.grid_columnconfigure(1, weight=1)
        header_item.grid_columnconfigure(2, weight=2)
        
        ctk.CTkLabel(header_item, text="ชื่อสินค้า").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(header_item, text="จำนวน").grid(row=0, column=1, padx=5, pady=2)
        ctk.CTkLabel(header_item, text="ราคา/หน่วย").grid(row=0, column=2, padx=5, pady=2)

        self.items_content_frame = ctk.CTkFrame(items_frame, fg_color="transparent")
        self.items_content_frame.pack(fill="x", pady=2)
        self.items_content_frame.grid_columnconfigure(0, weight=4)
        self.items_content_frame.grid_columnconfigure(1, weight=1)
        self.items_content_frame.grid_columnconfigure(2, weight=2)

        # --- ส่วนที่ 4: ค่าบริการตัด/เจาะ (Cutting) ---
        cutting_frame = ctk.CTkFrame(self.main_scroll, border_width=1)
        cutting_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        cutting_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(cutting_frame, text="[3. ค่าบริการตัด/เจาะ]", font=ctk.CTkFont(size=14, weight="bold"), text_color="#7E22CE").grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=5)

        ctk.CTkLabel(cutting_frame, text="ยอดเงิน:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.cutting_cost_entry = NumericEntry(cutting_frame)
        self.cutting_cost_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # VAT/CASH
        self.cutting_vat_var = tk.StringVar(value="VAT")
        vat_frame = ctk.CTkFrame(cutting_frame, fg_color="transparent")
        vat_frame.grid(row=1, column=2, padx=5, sticky="w")
        ctk.CTkRadioButton(vat_frame, text="VAT", variable=self.cutting_vat_var, value="VAT").pack(side="left", padx=5)
        ctk.CTkRadioButton(vat_frame, text="CASH", variable=self.cutting_vat_var, value="CASH").pack(side="left", padx=5)

        # WHT
        ctk.CTkLabel(cutting_frame, text="หัก WHT:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.cutting_wht_var = tk.StringVar(value="No")
        wht_frame = ctk.CTkFrame(cutting_frame, fg_color="transparent")
        wht_frame.grid(row=2, column=1, columnspan=2, padx=5, sticky="w")
        ctk.CTkRadioButton(wht_frame, text="ไม่หัก", variable=self.cutting_wht_var, value="No").pack(side="left", padx=5)
        ctk.CTkRadioButton(wht_frame, text="1%", variable=self.cutting_wht_var, value="1%").pack(side="left", padx=5)
        ctk.CTkRadioButton(wht_frame, text="3%", variable=self.cutting_wht_var, value="3%").pack(side="left", padx=5)

        # Note
        ctk.CTkLabel(cutting_frame, text="หมายเหตุ:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.cutting_remark_entry = ctk.CTkEntry(cutting_frame)
        self.cutting_remark_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # --- ส่วนที่ 5: ปุ่มควบคุม ---
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=10, pady=10, side="bottom")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(button_frame, text="ยกเลิก", command=self._on_close, fg_color="gray", height=40).grid(row=0, column=0, padx=5, sticky="ew")
        ctk.CTkButton(button_frame, text="บันทึกการแก้ไข", command=self._save_changes, height=40, fg_color="#16A34A").grid(row=0, column=1, padx=5, sticky="ew")

    def _load_data(self):
        """ดึงข้อมูล PO และ Items จาก DB (รองรับคอลัมน์ใหม่ครบถ้วน)"""
        try:
            # ดึงข้อมูล PO
            po_df = pd.read_sql("""
                SELECT *, 
                       cutting_cost, cutting_vat_type, cutting_wht_type, cutting_remark,
                       shipping_to_stock_driver, shipping_to_stock_plate,
                       shipping_to_site_driver, shipping_to_site_plate
                FROM purchase_orders 
                WHERE id = %s
            """, self.pg_engine, params=(self.po_id,))
            
            if po_df.empty:
                messagebox.showerror("ผิดพลาด", "ไม่พบข้อมูล PO", parent=self)
                self.destroy()
                return
            
            po_data = po_df.iloc[0]
            self.entry_supplier.insert(0, po_data.get('supplier_name', ''))
            self.entry_po_number.insert(0, po_data.get('po_number', ''))
            
            # โหลด Shipping Info (คนขับ/ทะเบียน)
            self.entry_stock_driver.insert(0, po_data.get('shipping_to_stock_driver', '') or '')
            self.entry_stock_plate.insert(0, po_data.get('shipping_to_stock_plate', '') or '')
            self.entry_site_driver.insert(0, po_data.get('shipping_to_site_driver', '') or '')
            self.entry_site_plate.insert(0, po_data.get('shipping_to_site_plate', '') or '')

            # โหลด Cutting Info
            cutting_cost = po_data.get('cutting_cost', 0) or 0
            self.cutting_cost_entry.insert(0, f"{cutting_cost:.2f}")
            self.cutting_vat_var.set(po_data.get('cutting_vat_type', 'VAT') or 'VAT')
            self.cutting_wht_var.set(po_data.get('cutting_wht_type', 'No') or 'No')
            self.cutting_remark_entry.insert(0, po_data.get('cutting_remark', '') or '')
            
            # โหลดรายการสินค้า
            items_df = pd.read_sql("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(self.po_id,))
            self.item_widgets = [] # Reset list
            for _, item in items_df.iterrows():
                self._add_item_row(item.to_dict())

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดข้อมูล: {e}", parent=self)

    def _add_item_row(self, item_data):
        row_index = len(self.item_widgets) + 1 # +1 เพราะแถว 0 คือ Header

        # [แก้ไข] ใช้ grid ในการวาง widget ของแถว
        entry_name = ctk.CTkEntry(self.items_content_frame)
        entry_name.insert(0, item_data.get('product_name', ''))
        entry_name.grid(row=row_index, column=0, padx=(0,2), pady=2, sticky="ew")

        entry_qty = NumericEntry(self.items_content_frame)
        entry_qty.insert(0, f"{item_data.get('quantity', 0):.2f}")
        entry_qty.grid(row=row_index, column=1, padx=2, pady=2, sticky="ew")

        entry_price = NumericEntry(self.items_content_frame)
        entry_price.insert(0, f"{item_data.get('unit_price', 0):.2f}")
        entry_price.grid(row=row_index, column=2, padx=2, pady=2, sticky="ew")
        
        self.item_widgets.append({
            'id': item_data['id'],
            'name_entry': entry_name,
            'qty_entry': entry_qty,
            'price_entry': entry_price
        })
        
    def _save_changes(self):
        """บันทึกข้อมูล (ตัด Relocation ออกจาก Logic)"""
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # 1. ค่า Cutting
                cut_cost = utils.convert_to_float(self.cutting_cost_entry.get())
                cut_vat_type = self.cutting_vat_var.get()
                cut_wht_type = self.cutting_wht_var.get()
                cut_remark = self.cutting_remark_entry.get().strip()
                cut_vat_amt = cut_cost * 0.07 if cut_vat_type == 'VAT' else 0.0
                
                cut_wht_amt = 0.0
                if cut_wht_type == '1%': cut_wht_amt = cut_cost * 0.01
                elif cut_wht_type == '3%': cut_wht_amt = cut_cost * 0.03

                # [🔥 เพิ่ม] Driver/Plate
                stock_drv = self.entry_stock_driver.get().strip()
                stock_plt = self.entry_stock_plate.get().strip()
                site_drv = self.entry_site_driver.get().strip()
                site_plt = self.entry_site_plate.get().strip()

                # 2. อัปเดต PO (Set relocation_cost = 0.0)
                cursor.execute("""
                    UPDATE purchase_orders 
                    SET supplier_name = %s, 
                        po_number = %s,
                        cutting_cost = %s, cutting_vat_type = %s, cutting_vat_amount = %s,
                        cutting_wht_type = %s, cutting_wht_amount = %s, cutting_remark = %s,
                        
                        shipping_to_stock_driver = %s, shipping_to_stock_plate = %s,
                        shipping_to_site_driver = %s, shipping_to_site_plate = %s,
                        
                        relocation_cost = 0.0
                        
                    WHERE id = %s
                """, (
                    self.entry_supplier.get(), self.entry_po_number.get(),
                    cut_cost, cut_vat_type, cut_vat_amt, cut_wht_type, cut_wht_amt, cut_remark,
                    stock_drv, stock_plt, site_drv, site_plt,
                    self.po_id
                ))

                # 3. อัปเดต Items
                for item_row in self.item_widgets:
                    new_name = item_row['name_entry'].get()
                    new_qty = utils.convert_to_float(item_row['qty_entry'].get())
                    new_price = utils.convert_to_float(item_row['price_entry'].get())
                    new_total = new_qty * new_price
                    
                    cursor.execute("""
                        UPDATE purchase_order_items 
                        SET product_name = %s, quantity = %s, unit_price = %s, total_price = %s
                        WHERE id = %s
                    """, (new_name, new_qty, new_price, new_total, item_row['id']))

                # 4. Audit Log
                log_msg = f"Edited PO: {self.entry_po_number.get()} by HR"
                cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp) VALUES (%s, %s, %s, %s, %s, %s)",
                               ('PO Edited by HR', 'purchase_orders', self.po_id, self.app_container.current_user_key, json.dumps({"msg": log_msg}), datetime.now()))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลเรียบร้อย", parent=self)
            self._on_close()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_close(self):
        if self.on_close_callback:
            self.on_close_callback() # สั่งให้หน้าต่าง HRVerificationWindow รีเฟรชตัวเอง
        self.destroy()

# --- [🔥 NEW CLASS] หน้าต่างค้นหาและพิมพ์ใบปะหน้าสำหรับ HR ---
class HRCoverSheetDialog(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        
        self.title("🖨️ ระบบพิมพ์ใบปะหน้า (HR Special)")
        self.geometry("800x600") # ขยายขนาดให้กว้างขึ้นเพื่อใส่ตาราง
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # ให้พื้นที่ตารางขยายตัว

        # --- 1. ส่วนค้นหา ---
        search_frame = CTkFrame(self, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        CTkLabel(search_frame, text="ค้นหา SO / PO / ลูกค้า:", font=CTkFont(size=14, weight="bold")).pack(side="left", padx=(0, 10))
        
        self.search_entry = CTkEntry(search_frame, placeholder_text="พิมพ์คำค้นหา...", width=300)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.search_entry.bind("<Return>", lambda e: self._search_data())
        self.search_entry.bind("<KeyRelease>", self._on_key_release) # ค้นหาทันทีที่พิมพ์ (Optional)
        
        CTkButton(search_frame, text="🔍 ค้นหา", width=100, command=self._search_data).pack(side="left")

        # --- 2. ส่วนแสดงตารางผลลัพธ์ (Treeview) ---
        table_frame = CTkFrame(self)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # สร้าง Style สำหรับ Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=30, font=("Roboto", 12))
        style.configure("Treeview.Heading", font=("Roboto", 12, "bold"))
        
        # กำหนดคอลัมน์
        columns = ("so_number", "customer", "po_count", "total_amount", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("so_number", text="SO Number")
        self.tree.heading("customer", text="ชื่อลูกค้า")
        self.tree.heading("po_count", text="จำนวน PO")
        self.tree.heading("total_amount", text="ยอดรวมต้นทุน PO")
        self.tree.heading("status", text="สถานะ SO")

        self.tree.column("so_number", width=120, anchor="center")
        self.tree.column("customer", width=250, anchor="w")
        self.tree.column("po_count", width=80, anchor="center")
        self.tree.column("total_amount", width=120, anchor="e")
        self.tree.column("status", width=100, anchor="center")
        
        # Scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Bind Double Click -> Print
        self.tree.bind("<Double-1>", lambda e: self._print_action())

        # --- 3. ส่วนปุ่มดำเนินการ ---
        action_frame = CTkFrame(self, fg_color="transparent")
        action_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=20)
        
        self.status_label = CTkLabel(action_frame, text="พบข้อมูล 0 รายการ", text_color="gray50")
        self.status_label.pack(side="left")

        self.print_btn = CTkButton(action_frame, text="🖨️ พิมพ์ใบปะหน้า (Selected)", 
                                   command=self._print_action, 
                                   fg_color="#7C3AED", hover_color="#6D28D9",
                                   state="disabled", width=200, height=40)
        self.print_btn.pack(side="right")

        self.print_transport_btn = CTkButton(action_frame, text="🖨️ ใบค่ารถ (Transport)", 
                                   command=self._print_transport_action, 
                                   fg_color="#059669", hover_color="#047857", # สีเขียว
                                   state="disabled", width=180, height=40)
        self.print_transport_btn.pack(side="right", padx=10)
        
        # โหลดข้อมูลเริ่มต้น (Optional)
        self._search_data(initial=True)
 
    def _print_transport_action(self):
        """Action สำหรับปุ่มพิมพ์ใบค่ารถในหน้า HR"""
        selected = self.tree.selection()
        if not selected:
            from tkinter import messagebox
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกรายการในตารางก่อน", parent=self)
            return

        # ดึง ID ของ PO จากตาราง (ปกติจะอยู่ที่ values[0])
        item_data = self.tree.item(selected[0])
        po_id = item_data['values'][0]
 
    def _print_transport_action(self):
        selected_item = self.tree.selection()
        if not selected_item: return
        item_values = self.tree.item(selected_item[0], "values")
        so_number = item_values[0]
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. Header
                cursor.execute("""
                    SELECT c.so_number, c.customer_name, u.sale_name 
                    FROM commissions c 
                    LEFT JOIN sales_users u ON c.sale_key = u.sale_key 
                    WHERE c.so_number = %s LIMIT 1
                """, (so_number,))
                header = cursor.fetchone()
                if not header: return

                # 2. Items
                cursor.execute("""
                    SELECT 
                        po_number, 
                        supplier_name,
                        
                        -- Stock (ค่าย้ายเข้าโกดัง)
                        shipping_to_stock_cost,
                        shipping_to_stock_driver,
                        shipping_to_stock_plate,
                        shipping_to_stock_notes,
                        shipping_to_stock_shipper,
                        shipping_to_stock_vat_type,
                        shipping_to_stock_wht_type,
                        shipping_to_stock_date,
                        
                        -- Site (ค่าส่งหน้างาน)
                        shipping_to_site_cost,
                        shipping_to_site_driver,
                        shipping_to_site_plate,
                        shipping_to_site_notes,
                        shipping_to_site_shipper,
                        shipping_to_site_vat_type,
                        shipping_to_site_wht_type,
                        shipping_to_site_date
                        
                    FROM purchase_orders 
                    WHERE so_number = %s AND status != 'Cancelled'
                """, (so_number,))
                pos = cursor.fetchall()
                
                transport_list = []
                for po in pos:
                    transport_list.append({
                        'po_number': po['po_number'],

                        # === STOCK (ค่าย้าย) ===
                        'stock_cost': po['shipping_to_stock_cost'] or 0,
                        'stock_driver': po['shipping_to_stock_driver'] or '-',
                        'stock_plate': po['shipping_to_stock_plate'] or '-',
                        'stock_notes': po['shipping_to_stock_notes'] or '-',  
                        'stock_supplier': po['shipping_to_stock_shipper'] or '-',  # ✅ ชื่อบริษัทจริง
                        'stock_vat': po['shipping_to_stock_vat_type'] or 'CASH',
                        'stock_wht': po['shipping_to_stock_wht_type'] or 'ไม่มีหัก',
                        'stock_date': po['shipping_to_stock_date'],

                        # === SITE (ค่าส่งหน้างาน) ===
                        'site_cost': po['shipping_to_site_cost'] or 0,
                        'site_driver': po['shipping_to_site_driver'] or '-',
                        'site_plate': po['shipping_to_site_plate'] or '-',
                        'site_notes': po['shipping_to_site_notes'] or '-',  
                        'site_supplier': po['shipping_to_site_shipper'] or '-',     # ✅ ชื่อบริษัทจริง
                        'site_vat': po['shipping_to_site_vat_type'] or 'CASH',
                        'site_wht': po['shipping_to_site_wht_type'] or 'ไม่มีหัก',
                        'site_date': po['shipping_to_site_date'],
                    })
                
                # เรียกฟังก์ชันสร้าง PDF
                from po_document_generator import generate_transport_fee_pdf
                generate_transport_fee_pdf(dict(header), transport_list)

        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Error", str(e))
            import traceback
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_key_release(self, event):
        # ทำ Debounce นิดหน่อยหรือค้นหาเลยก็ได้
        if self.search_entry.get().strip() == "":
            self._search_data(initial=True) # ถ้าลบหมดให้โหลดล่าสุด

    def _search_data(self, initial=False):
        keyword = self.search_entry.get().strip().upper()
        
        # ล้างข้อมูลเก่าในตาราง
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # เริ่มต้น Query หลัก
                base_query = """
                    SELECT 
                        c.so_number, 
                        c.customer_name, 
                        c.status,
                        COUNT(p.id) as po_count, 
                        COALESCE(SUM(p.grand_total), 0) as total_amt
                    FROM commissions c
                    LEFT JOIN purchase_orders p ON c.so_number = p.so_number AND p.status != 'Cancelled'
                    WHERE c.is_active = 1
                """
                
                params = []
                
                # ถ้ามีการค้นหา ให้เพิ่มเงื่อนไข WHERE
                if not initial and keyword:
                    base_query += """
                        AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s OR p.po_number ILIKE %s)
                    """
                    search_term = f"%{keyword}%"
                    params = [search_term, search_term, search_term]

                # [🔥 แก้ไขสำคัญ] ย้าย GROUP BY ออกมาข้างนอก เพื่อให้ทำงานเสมอ (แก้ Error SQL)
                # ต้อง Group ตามคอลัมน์ที่ไม่ได้ใช้ Aggregate function (COUNT/SUM)
                base_query += " GROUP BY c.so_number, c.customer_name, c.status, c.timestamp"
                
                # เรียงลำดับจากล่าสุดไปเก่าสุด
                base_query += " ORDER BY c.timestamp DESC"

                # ถ้าเป็นการโหลดครั้งแรก หรือไม่ได้พิมพ์คำค้นหา ให้จำกัดจำนวน 20 รายการ
                if initial or not keyword:
                    base_query += " LIMIT 20"

                cursor.execute(base_query, tuple(params))
                results = cursor.fetchall()
                
                # วนลูปใส่ข้อมูลลงตาราง
                for row in results:
                    so_num = row[0]
                    cust = row[1] or "N/A"
                    status = row[2]
                    po_cnt = row[3]
                    total = row[4]
                    
                    self.tree.insert("", "end", values=(so_num, cust, po_cnt, f"{total:,.2f}", status))
                
                self.status_label.configure(text=f"พบข้อมูล {len(results)} รายการ")
                
                # ตรวจสอบว่าเจอข้อมูลหรือไม่ เพื่อเปิด/ปิดปุ่ม
                if len(results) > 0:
                    self.print_btn.configure(state="normal")
                    
                    # ถ้ามีปุ่มพิมพ์ค่ารถ ให้เปิดใช้งานด้วย
                    if hasattr(self, 'print_transport_btn'):
                        self.print_transport_btn.configure(state="normal")
                    
                    # Select รายการแรกให้อัตโนมัติ เพื่อความสะดวก
                    child_id = self.tree.get_children()[0]
                    self.tree.focus(child_id)
                    self.tree.selection_set(child_id)
                else:
                    self.print_btn.configure(state="disabled")
                    
                    # ถ้ามีปุ่มพิมพ์ค่ารถ ให้ปิดใช้งานด้วย
                    if hasattr(self, 'print_transport_btn'):
                        self.print_transport_btn.configure(state="disabled")

        except Exception as e:
            print(f"Search error: {e}")
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการดึงข้อมูล: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _print_action(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกรายการที่ต้องการพิมพ์")
            return
            
        item_values = self.tree.item(selected_item[0], "values")
        so_number = item_values[0] # ดึงเลข SO จากคอลัมน์แรก
        
        # เรียกใช้ Logic การพิมพ์เดิม
        self._execute_print(so_number)

    def _execute_print(self, so_number):
        try:
            # 1. ดึง SO Header
            so_df = pd.read_sql_query("""
                SELECT c.*, u_so.sale_name AS sale_name
                FROM commissions c
                LEFT JOIN sales_users u_so ON c.sale_key = u_so.sale_key
                WHERE c.so_number = %s AND c.is_active = 1 LIMIT 1
            """, self.pg_engine, params=(so_number,))
            
            if so_df.empty: return
            so_header_data = so_df.iloc[0].to_dict()

            # 2. ดึง PO List (ไม่สน Approved, เอาหมดที่ไม่ใช่ Cancelled)
            po_query = """
                SELECT po.*, u_po.sale_name AS user_name,
                    m1.sale_name AS approver_1, m2.sale_name AS approver_2, d.sale_name AS approver_3
                FROM purchase_orders po
                LEFT JOIN sales_users u_po ON po.user_key = u_po.sale_key
                LEFT JOIN sales_users m1 ON po.approver_manager1_key = m1.sale_key
                LEFT JOIN sales_users m2 ON po.approver_manager2_key = m2.sale_key
                LEFT JOIN sales_users d ON po.approver_director_key = d.sale_key
                WHERE po.so_number = %s AND po.status != 'Cancelled'
            """
            all_po_df = pd.read_sql_query(po_query, self.pg_engine, params=(so_number,))
            
            if all_po_df.empty:
                messagebox.showinfo("แจ้งเตือน", "รายการนี้ไม่มีใบสั่งซื้อ (PO)", parent=self)
                return

            # 3. เตรียม Data List
            all_po_data_list = []
            
            # [🔥 เพิ่ม] เตรียม Connection เพื่อดึงข้อมูล Supplier Bank
            conn = self.app_container.get_connection()
            
            try:
                for _, po_row in all_po_df.iterrows():
                    po_id = po_row['id']
                    
                    # ดึง Items
                    items_df = pd.read_sql("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
                    
                    # ดึง Payments (รวม bank_account_type)
                    payments_df = pd.read_sql("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
                    
                    po_dict = po_row.to_dict()
                    
                    # =========================================================
                    # [🔥 เพิ่ม] ดึงข้อมูลธนาคารจาก Master Data (Suppliers) มาใส่
                    # เพื่อใช้กรณีที่ยังไม่มีประวัติการจ่ายเงินใน payments_df
                    # =========================================================
                    supplier_name = po_dict.get('supplier_name')
                    if supplier_name:
                        with conn.cursor() as cur:
                            cur.execute("""
                                SELECT bank_name, bank_account_number, bank_account_type 
                                FROM suppliers 
                                WHERE supplier_name = %s LIMIT 1
                            """, (supplier_name,))
                            sup_res = cur.fetchone()
                            if sup_res:
                                po_dict['supplier_bank_name'] = sup_res[0]
                                po_dict['supplier_account_number'] = sup_res[1]
                                po_dict['supplier_account_type'] = sup_res[2] # ประเภทบัญชีจาก Master
                    # =========================================================

                    # คำนวณยอดจ่ายแล้ว
                    deposit = sum(p['amount'] for p in payments_df.to_dict('records') if p['payment_type'] in ['Payment 1', 'Payment 2'])
                    full_pay = sum(p['amount'] for p in payments_df.to_dict('records') if p['payment_type'] == 'Full Payment')
                    
                    po_dict['deposit_amount'] = deposit
                    po_dict['full_payment_amount'] = full_pay
                    po_dict['balance_due_po'] = (po_dict.get('grand_total', 0) or 0) - deposit - full_pay
                    
                    # Mapping Shipping/Approver
                    po_dict['shipping_cost_1'] = po_dict.get('shipping_to_stock_cost', 0.0)
                    po_dict['shipping_vat_type_1'] = po_dict.get('shipping_to_stock_vat_type', 'CASH')
                    po_dict['shipper_1'] = po_dict.get('shipping_to_stock_shipper', '')
                    po_dict['shipping_cost_2'] = po_dict.get('shipping_to_site_cost', 0.0)
                    po_dict['shipping_vat_type_2'] = po_dict.get('shipping_to_site_vat_type', 'CASH')
                    po_dict['shipper_2'] = po_dict.get('shipping_to_site_shipper', '')
                    
                    po_dict['creator_user'] = po_dict.get('user_name', '')
                    po_dict['approver_1'] = po_dict.get('approver_1', '')
                    po_dict['approver_2'] = po_dict.get('approver_2', '')
                    po_dict['approver_3'] = po_dict.get('approver_3', '')

                    all_po_data_list.append({
                        "header": po_dict,
                        "items": items_df.to_dict('records'),
                        "payments": payments_df.to_dict('records')
                    })
            finally:
                if conn: self.app_container.release_connection(conn)

            # 4. Print
            from po_document_generator import generate_multi_po_pdf
            generate_multi_po_pdf(so_header_data=so_header_data, all_po_data=all_po_data_list)
            
            # ไม่ต้องปิดหน้าต่าง (เผื่ออยากพิมพ์ใบอื่นต่อ)

        except Exception as e:
            messagebox.showerror("Error", f"Print Failed: {e}", parent=self)
            traceback.print_exc()