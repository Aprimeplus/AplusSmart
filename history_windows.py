# history_windows.py

import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkToplevel, CTkTextbox, CTkScrollableFrame, CTkLabel, CTkFont, CTkFrame, CTkButton, CTkEntry, CTkRadioButton, CTkOptionMenu, CTkTabview)
from tkinter import messagebox
import json
import customtkinter
from customtkinter import CTkLabel, CTkFont, CTkCheckBox
import pandas as pd
from datetime import datetime
import traceback
import psycopg2.errors
import psycopg2.extras
import numpy as np

# --- ตรวจสอบว่า import ถูกต้องตามนี้ ---
import utils
from utils import FormattedNumericEntry, RejectionReasonDialog
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
# ---

from sqlalchemy import create_engine

class SalesDataViewerWindow(CTkToplevel):
    def __init__(self, master, app_container, so_number):
        super().__init__(master)
        self.app_container = app_container
        self.so_number = so_number
        self.so_data = None

        self.title(f"รายละเอียดข้อมูล SO: {self.so_number}")
        self.geometry("800x650")

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Header ---
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        CTkLabel(header_frame, text=f"ข้อมูลสำหรับ SO Number: {self.so_number}", font=CTkFont(size=18, weight="bold")).pack(side="left")

        # --- Main Frame ---
        self.main_frame = CTkScrollableFrame(self)
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
                # ดึงข้อมูลล่าสุดที่มี is_active = 1
                cursor.execute("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1", (self.so_number,))
                self.so_data = cursor.fetchone()
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล SO ได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def create_widgets(self):
        if not self.so_data:
            CTkLabel(self.main_frame, text="ไม่พบข้อมูล").pack(pady=20)
            return

        self.main_frame.grid_columnconfigure(1, weight=1)
        header_map = self.app_container.HEADER_MAP
        current_row = 0

        # --- ส่วนแสดงรายละเอียด SO ---
        title_label = CTkLabel(self.main_frame, text="รายละเอียด SO", font=CTkFont(size=16, weight="bold"), anchor="w")
        title_label.grid(row=current_row, column=0, columnspan=2, padx=10, pady=(15, 5), sticky="w")
        current_row += 1

        # --- START: แก้ไข Logic การแสดงผลตรงนี้ ---
        # รายการฟิลด์ที่จะแสดงผล
        fields_to_display = [
            'bill_date', 'customer_id', 'customer_name', 'credit_term',
            'sales_service_amount', 'cutting_drilling_fee', 'other_service_fee',
            'shipping_cost', 'delivery_date', 'relocation_cost',
            'credit_card_fee', 'brokerage_fee', 'giveaways', 'coupons',
            'total_payment_amount', 'payment_date'
        ]

        # รายการฟิลด์ที่ต้องตรวจสอบ VAT
        vat_fields = {
            'sales_service_amount': 'sales_service_vat_option',
            'cutting_drilling_fee': 'cutting_drilling_fee_vat_option',
            'other_service_fee': 'other_service_fee_vat_option',
            'shipping_cost': 'shipping_vat_option',
            'credit_card_fee': 'credit_card_fee_vat_option'
        }

        for col in fields_to_display:
            if col not in self.so_data: continue

            display_name = header_map.get(col, col)
            value = self.so_data[col]

            # ถ้าฟิลด์นี้อยู่ในกลุ่มที่ต้องเช็ค VAT ให้เรียกใช้ฟังก์ชันใหม่
            if col in vat_fields:
                vat_option_key = vat_fields[col]
                vat_option = self.so_data.get(vat_option_key, 'CASH')
                rows_used = self._add_display_row_with_vat(self.main_frame, current_row, display_name, value, vat_option)
                current_row += rows_used

            # ถ้าไม่ใช่ ให้แสดงผลแบบปกติ
            else:
                if isinstance(value, (int, float)): value_text = f"{value:,.2f}"
                elif isinstance(value, datetime): value_text = value.strftime('%d/%m/%Y')
                else: value_text = str(value) if value is not None else "-"
                self._add_display_row(self.main_frame, current_row, display_name, value_text)
                current_row += 1


class PurchaseDetailWindow(CTkToplevel):
    def __init__(self, master, app_container, purchase_id, on_save_callback=None, **kwargs):
        super().__init__(master)
        self.title(f"รายละเอียด/แก้ไขใบสั่งซื้อ (PO ID: {purchase_id})")
        self.geometry("100x100")
        
        self.app_container = app_container
        self.purchase_id = purchase_id
        self.on_save_callback = on_save_callback
        self._load_supplier_data_for_autocomplete()
        
        self.user_role = self.app_container.current_user_role
        
        self.po_entries = {}
        self.item_entries = []
        self.deleted_item_ids = []
        self.payment_entries = []
        self.deleted_payment_ids = []
        
        # +++ START: แก้ไข Layout หลัก +++
        # กำหนดให้แถวที่ 0 (ScrollFrame) ขยายตัว แต่แถวที่ 1 (Buttons) ไม่ขยาย
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0) 
        self.grid_columnconfigure(0, weight=1)

        # ScrollFrame จะอยู่ในแถวที่ 0
        self.scroll_frame = CTkScrollableFrame(self)
        self.scroll_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)
        
        # สร้างปุ่ม Action ต่างๆ (จะถูกวางในแถวที่ 1)
        self._create_action_buttons()
        # +++ END +++
        
        

        self.after(50, self._load_and_display_data)
        self.transient(master)
        self.grab_set()
    
    def _position_window(self):
        """จัดตำแหน่งหน้าต่างให้อยู่กึ่งกลางแนวนอน และยึดตำแหน่งบนสุดไว้"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        
        # จัดกึ่งกลางแนวนอนของ "หน้าจอ"
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        
        # ยึดตำแหน่งบนสุดของหน้าต่าง (แกน Y) ไว้ที่ 40 pixels จากขอบจอบนเสมอ
        y = 40
        
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _load_supplier_data_for_autocomplete(self):
        """โหลดข้อมูล Supplier ทั้งหมดมาเตรียมไว้สำหรับ AutoComplete"""
        try:
            df = pd.read_sql("SELECT id, supplier_name, supplier_code, credit_term FROM suppliers ORDER BY supplier_name", self.app_container.pg_engine)
            self.supplier_completion_data = []
            for _, row in df.iterrows():
                self.supplier_completion_data.append({
                    "id": row['id'],
                    "name": row['supplier_name'],
                    "code": row.get('supplier_code', ''),
                    "term": row.get('credit_term', 'เงินสด')
                })
        except Exception as e:
            print(f"Error loading supplier data for history window: {e}")
            self.supplier_completion_data = []

    def _on_supplier_selected_in_detail(self, selection_dict):
        """Callback เมื่อมีการเลือก Supplier จาก AutoComplete ในหน้ารายละเอียด"""
        if not selection_dict:
            return

        # อัปเดตข้อมูลในช่อง Supplier Code ผ่าน self.po_entries
        supplier_code_entry = self.po_entries.get('supplier_code')
        if supplier_code_entry and supplier_code_entry.winfo_exists():
            supplier_code_entry.delete(0, tk.END)
            supplier_code_entry.insert(0, selection_dict.get('code', ''))

        # อัปเดตข้อมูลในช่อง Credit Term ผ่าน self.po_entries
        credit_term_entry = self.po_entries.get('credit_term')
        if credit_term_entry and credit_term_entry.winfo_exists():
            credit_term_map = {'เงินสด': 'เงินสด', '0': 'เงินสด', '7': 'Cr 7', '15': 'Cr 15', '30': 'Cr 30'}
            term_value = str(selection_dict.get('term', 'เงินสด')).strip()
            credit_term_entry.delete(0, tk.END)
            credit_term_entry.insert(0, credit_term_map.get(term_value, term_value))

    def _get_supplier_names(self):
        """(เวอร์ชันแก้ไข) ดึงรายการชื่อซัพพลายเออร์และทำความสะอาดข้อมูล"""
        conn = None
        names = []
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # ดึงข้อมูลเฉพาะชื่อซัพพลายเออร์ที่ไม่ใช่ค่าว่าง
                cursor.execute("SELECT supplier_name FROM suppliers WHERE supplier_name IS NOT NULL AND supplier_name != '' ORDER BY supplier_name")
                
                fetched_rows = cursor.fetchall()
                cleaned_items = set() # ใช้ set เพื่อกรองข้อมูลซ้ำ

                for row in fetched_rows:
                    original_name = row[0]
                    if isinstance(original_name, str):
                        # .strip() เพื่อลบช่องว่างหน้า-หลัง
                        clean_name = original_name.strip()
                        if clean_name: # ตรวจสอบว่าชื่อไม่เป็นค่าว่างหลัง clean
                            cleaned_items.add(clean_name)
                
                names = sorted(list(cleaned_items)) # แปลงกลับเป็น List และเรียงลำดับ

        except Exception as e:
            print(f"!!! ERROR in _get_supplier_names: {e}")
        finally:
            if conn:
                self.app_container.release_connection(conn)
        
        return names

    def _on_supplier_selected(self, supplier_name_or_code):
        """เมื่อเลือกซัพพลายเออร์จากรายการ Autocomplete ให้ดึงข้อมูลมาเติม"""
        input_value = supplier_name_or_code.strip()
        if not input_value:
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. ลองค้นหาจาก "ชื่อ" ก่อน
                cursor.execute("SELECT supplier_code, credit_term FROM suppliers WHERE supplier_name = %s LIMIT 1", (input_value,))
                supplier_info = cursor.fetchone()

                if supplier_info:
                    # พบข้อมูลจากชื่อ -> เติมรหัสและ Credit Term
                    utils.set_entry_text(self.po_entries['supplier_code'], supplier_info['supplier_code'])
                    utils.set_entry_text(self.po_entries['credit_term'], supplier_info['credit_term'])
                else:
                    # 2. ถ้าไม่พบ ให้ลองค้นหาจาก "รหัส"
                    cursor.execute("SELECT supplier_name, credit_term FROM suppliers WHERE supplier_code = %s LIMIT 1", (input_value,))
                    supplier_info_by_code = cursor.fetchone()
                    if supplier_info_by_code:
                        # พบข้อมูลจากรหัส -> เติมชื่อและ Credit Term (ข้อมูลจะ Sync กัน)
                        utils.set_entry_text(self.po_entries['supplier_name'], supplier_info_by_code['supplier_name'])
                        utils.set_entry_text(self.po_entries['credit_term'], supplier_info_by_code['credit_term'])

        except Exception as e:
            print(f"Error in _on_supplier_selected: {e}")
        finally:
            if conn:
                self.app_container.release_connection(conn)

    def _create_dropdown_row(self, parent, row_index, label, value, key, options):
        """ฟังก์ชัน Helper ใหม่สำหรับสร้างแถว Dropdown ให้สวยงามขึ้น"""
        CTkLabel(parent, text=f"{label}:").grid(row=row_index, column=0, padx=10, pady=5, sticky="w")
        
        entry_var = tk.StringVar()
        initial_value = str(value) if value is not None and str(value) in options else options[0]
        entry_var.set(initial_value)
        
        # กำหนดความกว้างของ Dropdown และผูก command
        entry = CTkOptionMenu(parent, variable=entry_var, values=options, width=250, # <-- กำหนดความกว้าง
                            command=self._recalculate_summary_totals)
        
        entry.grid(row=row_index, column=1, padx=10, pady=5, sticky="w") # <-- เปลี่ยนเป็น sticky="w"
        self.po_entries[key] = entry_var
        return entry
    
    def _on_bank_selected(self, selected_bank, account_entry_widget):
        """
        เมื่อเลือกธนาคาร ให้ค้นหาเลขบัญชีล่าสุดที่เคยใช้กับซัพพลายเออร์และธนาคารนี้
        """
        if not selected_bank or selected_bank == "ระบุเอง":
            return

        supplier_name = self.po_entries.get('supplier_name').get()
        if not supplier_name:
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                query = """
                    SELECT p_pay.bank_account_number
                    FROM purchase_order_payments p_pay
                    JOIN purchase_orders po ON p_pay.purchase_order_id = po.id
                    WHERE po.supplier_name = %s 
                    AND p_pay.bank_name = %s
                    AND p_pay.bank_account_number IS NOT NULL AND p_pay.bank_account_number != ''
                    ORDER BY p_pay.payment_date DESC, p_pay.id DESC
                    LIMIT 1;
                """
                cursor.execute(query, (supplier_name, selected_bank))
                result = cursor.fetchone()

                if result and account_entry_widget.winfo_exists():
                    account_number = result[0]
                    utils.set_entry_text(account_entry_widget, account_number)

        except Exception as e:
            print(f"Error looking up last bank account: {e}")
            # ไม่ต้องแสดง popup error เพื่อไม่ให้รบกวนการทำงาน
        finally:
            if conn:
                self.app_container.release_connection(conn)
    
    def _load_and_display_data(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (self.purchase_id,))
                po_data = cursor.fetchone()
                
                if not po_data:
                    messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบ PO ID: {self.purchase_id}", parent=self)
                    self.destroy()
                    return

                self.po_data = dict(po_data)
                
                # ดึง supplier_code และ credit_term จากชื่อซัพพลายเออร์
                supplier_name = self.po_data.get('supplier_name')
                if supplier_name:
                    cursor.execute("SELECT supplier_code, credit_term FROM suppliers WHERE supplier_name = %s LIMIT 1", (supplier_name,))
                    supplier_info = cursor.fetchone()
                    if supplier_info:
                        self.po_data['supplier_code'] = supplier_info['supplier_code']
                        self.po_data['credit_term'] = supplier_info['credit_term']
                    else:
                        self.po_data['supplier_code'] = ""
                        self.po_data['credit_term'] = ""

                cursor.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", (self.purchase_id,))
                items_data = cursor.fetchall()
                cursor.execute("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", (self.purchase_id,))
                payments_data = cursor.fetchall()

            self.items_data = [dict(item) for item in items_data]
            self.payments_data = [dict(payment) for payment in payments_data]

            self._create_formatted_view()

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)


    def _create_formatted_view(self):
        # ฟังก์ชันนี้จะวาดเนื้อหาลงใน self.scroll_frame
        for widget in self.scroll_frame.winfo_children(): widget.destroy()
        
        self.item_entries, self.deleted_item_ids = [], []
        self.payment_entries, self.deleted_payment_ids = [], []

        self._create_info_section(self.scroll_frame, self.po_data)
        self._create_summary_section(self.scroll_frame, self.po_data)
        self._create_items_section(self.scroll_frame, self.items_data)
        self._create_shipping_section(self.scroll_frame, self.po_data)
        self._create_payments_section(self.scroll_frame, self.payments_data)
        self._create_approval_info_section(self.scroll_frame, self.po_data)
        
        self._recalculate_summary_totals()

        # --- เพิ่มบรรทัดนี้เข้าไปท้ายสุด ---
        # หน่วงเวลาเล็กน้อยเพื่อให้แน่ใจว่า UI วาดเสร็จแล้วจึงค่อยปรับขนาดและตำแหน่ง
        self.after(100, self._position_window)

    def _create_section(self, parent, title):
        section_frame = CTkFrame(parent, corner_radius=10, border_width=1)
        section_frame.pack(fill="x", padx=10, pady=8)
        section_frame.grid_columnconfigure(1, weight=1)
        CTkLabel(section_frame, text=title, font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(5,10), sticky="w")
        return section_frame
    
    def _create_editable_row(self, parent, row_index, label, value, key, is_numeric=False, widget_class=None, options=None):
        """Helper สำหรับสร้างแถวที่สามารถแก้ไขข้อมูลได้"""
        CTkLabel(parent, text=f"{label}:").grid(row=row_index, column=0, padx=10, pady=5, sticky="w")
        
        if widget_class == CTkOptionMenu:
            entry_var = tk.StringVar()
            initial_value = str(value) if value is not None and str(value) in (options or []) else (options[0] if options else "")
            entry_var.set(initial_value)
            entry = CTkOptionMenu(parent, variable=entry_var, values=options or [], command=self._recalculate_summary_totals)
            self.po_entries[key] = entry_var 
        elif isinstance(self.po_entries.get(key), AutoCompleteEntry): # ถ้าเป็น AutoCompleteEntry ให้ข้ามไป
            return
        elif is_numeric:
            entry = FormattedNumericEntry(parent, command=self._recalculate_summary_totals)
            entry.set(value if value is not None else 0.0)
            self.po_entries[key] = entry
        else:
            entry = CTkEntry(parent)
            entry.insert(0, str(value) if value is not None else "")
            self.po_entries[key] = entry
            
        entry.grid(row=row_index, column=1, padx=10, pady=5, sticky="ew")
        return entry

    def _add_display_row(self, parent, row_index, label, value):
        if value is None or pd.isna(value): value_text = "-"
        elif isinstance(value, (int, float, np.floating)): value_text = f"{value:,.2f}"
        elif isinstance(value, datetime): value_text = value.strftime('%d/%m/%Y %H:%M')
        else: value_text = str(value)

        CTkLabel(parent, text=f"{label}:", anchor="w").grid(row=row_index, column=0, padx=10, pady=3, sticky="w")
        CTkLabel(parent, text=value_text, wraplength=400, justify="left", anchor="w").grid(row=row_index, column=1, padx=10, pady=3, sticky="w")
    
    def _add_display_row_with_vat(self, parent, row_index, label_text, value, vat_option):
        """ฟังก์ชัน Helper สำหรับสร้างแถวข้อมูลพร้อมคำนวณและแสดง VAT"""
        # 1. แสดงแถวข้อมูลหลัก พร้อมวงเล็บ (VAT) หรือ (CASH)
        value_text = f"{value:,.2f} ({vat_option})" if isinstance(value, (int, float)) else f"{value} ({vat_option})"
        self._add_display_row(parent, row_index, label_text, value_text)

        # 2. คำนวณและแสดงแถว VAT ถ้าเป็น 'VAT'
        if vat_option == 'VAT' and isinstance(value, (int, float)) and value != 0:
            vat_amount = value * 0.07
            vat_font = CTkFont(size=12, slant="italic")

            CTkLabel(parent, text="  └─ ยอด VAT 7%", font=vat_font, text_color="gray50").grid(
                row=row_index + 1, column=0, padx=(30, 10), pady=(0, 4), sticky="w")
            CTkLabel(parent, text=f"{vat_amount:,.2f}", font=vat_font, text_color="gray50").grid(
                row=row_index + 1, column=1, padx=(10, 15), pady=(0, 4), sticky="w")
            return 2 # บอกว่าใช้ไป 2 แถว

        return 1 # บอกว่าใช้ไป 1 แถว

    def _create_info_section(self, parent, data):
        info_frame = self._create_section(parent, "ข้อมูลทั่วไป")
        current_row = 1

        # --- ส่วนของ Supplier Name ที่ใช้ AutoCompleteEntry ---
        CTkLabel(info_frame, text="ชื่อซัพพลายเออร์:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        supplier_entry = AutoCompleteEntry(
            info_frame,
            completion_list=self.supplier_completion_data,  # ใช้ List of Dictionaries
            display_key='name',                             # ระบุ Key ที่จะแสดงผล
            command=self._on_supplier_selected_in_detail    # เรียกใช้ Callback command
        )
        supplier_entry.insert(0, data.get("supplier_name", ""))
        supplier_entry.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries['supplier_name'] = supplier_entry
        current_row += 1

        # --- ส่วนของข้อมูลอื่นๆ ที่สามารถแก้ไขได้ ---
        # เราจะเก็บ reference ของ supplier_code และ credit_term ไว้ใน self.po_entries
        # เพื่อให้ callback สามารถอัปเดตค่าในช่องเหล่านี้ได้
        self._create_editable_row(info_frame, current_row, "PO Number", data.get("po_number"), key="po_number")
        current_row += 1
        
        # เมื่อสร้างแถว รหัสซัพพลายเออร์ เราเก็บ widget ที่ return กลับมาไว้
        supplier_code_widget = self._create_editable_row(info_frame, current_row, "รหัสซัพพลายเออร์", data.get("supplier_code"), key="supplier_code")
        self.po_entries['supplier_code'] = supplier_code_widget
        current_row += 1
        
        # เมื่อสร้างแถว Credit Term เราเก็บ widget ที่ return กลับมาไว้
        credit_term_widget = self._create_editable_row(info_frame, current_row, "Credit Term", data.get("credit_term"), key="credit_term")
        self.po_entries['credit_term'] = credit_term_widget
        current_row += 1
        
        self._create_editable_row(info_frame, current_row, "SO Number", data.get("so_number"), key="so_number")
        current_row += 1
        
        self._create_editable_row(info_frame, current_row, "ประเภท PO", data.get("po_mode"), key="po_mode", widget_class=CTkOptionMenu, options=["Single-PO", "Multiple-PO"])
        current_row += 1
        
        self._add_display_row(info_frame, current_row, "สถานะ", data.get("status"))
        current_row += 1

    def _create_items_section(self, parent, items_list):
        self.items_frame = self._create_section(parent, "รายการสินค้า")
        
        # --- START: เพิ่ม 'คลัง' เข้าไปใน Headers ---
        headers = ["รหัสสินค้า", "ชื่อสินค้า", "คลัง", "น้ำหนัก", "จำนวน", "ราคา/หน่วย", "ส่วนลด", "ราคารวม"]
        # --- END ---

        header_container = CTkFrame(self.items_frame, fg_color="transparent")
        header_container.grid(row=1, column=0, sticky="ew")
        
        # --- START: ปรับสัดส่วนคอลัมน์ใหม่ ---
        col_weights = [2, 4, 2, 1, 1, 2, 3, 2] 
        # --- END ---

        for i, header_text in enumerate(headers):
            header_container.grid_columnconfigure(i, weight=col_weights[i])
            CTkLabel(header_container, text=header_text, font=CTkFont(size=14, weight="bold")).grid(row=0, column=i, padx=5, pady=5)
        
        self.items_content_frame = CTkFrame(self.items_frame, fg_color="transparent")
        self.items_content_frame.grid(row=2, column=0, sticky="ew")
        for item in items_list:
            self._add_item_row(item)
        
        add_button = CTkButton(self.items_frame, text="+ เพิ่มรายการ", command=self._add_new_item_row)
        add_button.grid(row=3, column=0, padx=5, pady=10, sticky="w")

    def _add_item_row(self, item_data=None, is_new=False):
        if item_data is None: item_data = {}
        row_frame = CTkFrame(self.items_content_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=2)
        
        # --- START: ปรับสัดส่วนและเพิ่มคอลัมน์ 'คลัง' ---
        col_weights = [2, 4, 2, 1, 1, 2, 3, 2, 0] 
        for i, w in enumerate(col_weights):
            row_frame.grid_columnconfigure(i, weight=w)

        entry_code = CTkEntry(row_frame); entry_code.insert(0, item_data.get('product_code', '')); entry_code.grid(row=0, column=0, padx=5, sticky="ew")
        entry_name = CTkEntry(row_frame); entry_name.insert(0, item_data.get('product_name', '')); entry_name.grid(row=0, column=1, padx=5, sticky="ew")
        
        # เพิ่มช่องกรอกสำหรับ 'คลัง'
        entry_warehouse = CTkEntry(row_frame); entry_warehouse.insert(0, item_data.get('warehouse', '')); entry_warehouse.grid(row=0, column=2, padx=5, sticky="ew")

        entry_weight = FormattedNumericEntry(row_frame, command=self._recalculate_summary_totals); entry_weight.set(item_data.get('total_weight', 0)); entry_weight.grid(row=0, column=3, padx=5, sticky="ew")
        entry_qty = FormattedNumericEntry(row_frame, command=self._recalculate_summary_totals); entry_qty.set(item_data.get('quantity', 0)); entry_qty.grid(row=0, column=4, padx=5, sticky="ew")
        entry_price = FormattedNumericEntry(row_frame, command=self._recalculate_summary_totals); entry_price.set(item_data.get('unit_price', 0)); entry_price.grid(row=0, column=5, padx=5, sticky="ew")
        
        discount_frame = CTkFrame(row_frame, fg_color="transparent")
        discount_frame.grid(row=0, column=6, padx=5, sticky="ew")
        discount_frame.grid_columnconfigure(0, weight=1)
        
        entry_discount = FormattedNumericEntry(discount_frame, command=self._recalculate_summary_totals)
        entry_discount.set(item_data.get('discount_value', 0))
        entry_discount.pack(side="left", fill="x", expand=True, padx=(0,2))

        discount_type_var = tk.StringVar(value=item_data.get('discount_type', 'บาท'))
        discount_type_menu = CTkOptionMenu(discount_frame, variable=discount_type_var, values=["บาท", "%"], width=70, command=self._recalculate_summary_totals)
        discount_type_menu.pack(side="left")
        
        label_total = CTkLabel(row_frame, text="0.00", anchor="e"); label_total.grid(row=0, column=7, padx=5, sticky="ew")
        delete_button = CTkButton(row_frame, text="ลบ", width=40, fg_color="#DC2626", hover_color="#B91C1C", command=lambda r=row_frame, i=item_data.get('id'): self._remove_item_row(r, i)); delete_button.grid(row=0, column=8, padx=(5,0))

        self.item_entries.append({
            'id': item_data.get('id'), 'frame': row_frame, 
            'widgets': {
                'product_code': entry_code, 'product_name': entry_name, 
                'warehouse': entry_warehouse, # <-- เก็บ widget ใหม่
                'total_weight': entry_weight, 'quantity': entry_qty, 
                'unit_price': entry_price, 'discount_value': entry_discount,
                'discount_type_var': discount_type_var,
                'total_price_label': label_total
            }
        })
        self._recalculate_summary_totals()

    def _add_new_item_row(self):
        self._add_item_row(is_new=True)

    def _remove_item_row(self, row_frame, item_id):
        if item_id is not None:
            self.deleted_item_ids.append(item_id)
        index_to_remove = -1
        for i, item_entry in enumerate(self.item_entries):
            if item_entry['frame'] == row_frame:
                index_to_remove = i
                break
        if index_to_remove != -1:
            self.item_entries.pop(index_to_remove)
        row_frame.destroy()
        self._recalculate_summary_totals()

    # (ในไฟล์ history_windows.py ภายในคลาส PurchaseDetailWindow)
# ให้นำฟังก์ชันนี้ไปวางทับของเดิม

    def _create_shipping_section(self, parent, data):
        shipping_frame = self._create_section(parent, "ข้อมูลการจัดส่ง")
        current_row = 1

        # --- Sub-section: ค่าส่งเข้าสต๊อก ---
        CTkLabel(shipping_frame, text="--- ค่าจัดส่งเข้าสต๊อก ---", font=CTkFont(weight="bold")).grid(row=current_row, column=0, columnspan=2, pady=(5,2), sticky="w", padx=10)
        current_row += 1

        self._create_editable_row(shipping_frame, current_row, "ค่าส่งเข้าสต๊อก:", data.get("shipping_to_stock_cost"), key="shipping_to_stock_cost", is_numeric=True)
        current_row += 1

        CTkLabel(shipping_frame, text="VAT 7%:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        stock_vat_display = CTkEntry(shipping_frame, state="readonly", fg_color="gray85")
        stock_vat_display.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries["shipping_to_stock_vat_display"] = stock_vat_display
        current_row += 1

        CTkLabel(shipping_frame, text="วันที่ส่งเข้าสต๊อก:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        stock_date_selector = DateSelector(shipping_frame)
        stock_date_selector.set_date(data.get("shipping_to_stock_date"))
        stock_date_selector.grid(row=current_row, column=1, padx=10, pady=5, sticky="w")
        self.po_entries["shipping_to_stock_date"] = stock_date_selector
        current_row += 1

        self._create_dropdown_row(shipping_frame, current_row, "ประเภท VAT", data.get("shipping_to_stock_vat_type"), key="shipping_to_stock_vat_type", options=["VAT", "CASH"])
        current_row += 1
        
        shipper_options = ["ซัพพลายเออร์จัดส่ง", "Aplus Logistic", "Lalamove/Others"]
        self._create_dropdown_row(shipping_frame, current_row, "ผู้จัดส่ง", data.get("shipping_to_stock_shipper"), key="shipping_to_stock_shipper", options=shipper_options)
        current_row += 1

        wht_options = ["ไม่มีหัก", "1%", "3%"]
        self._create_dropdown_row(shipping_frame, current_row, "หัก ณ ที่จ่าย", data.get("shipping_to_stock_wht_type"), key="shipping_to_stock_wht_type", options=wht_options)
        current_row += 1
        
        # --- START: เพิ่มช่องแสดงยอดหัก ณ ที่จ่าย (สต๊อก) ---
        CTkLabel(shipping_frame, text="ยอดหัก ณ ที่จ่าย:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        stock_wht_display = CTkEntry(shipping_frame, state="readonly", fg_color="gray85")
        stock_wht_display.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries["shipping_to_stock_wht_display"] = stock_wht_display
        current_row += 1
        # --- END ---
        
        self._create_editable_row(shipping_frame, current_row, "หมายเหตุ:", data.get("shipping_to_stock_notes"), key="shipping_to_stock_notes")
        current_row += 1

        # --- Sub-section: ค่าส่งเข้าไซต์ ---
        CTkLabel(shipping_frame, text="--- ค่าจัดส่งเข้าไซต์ ---", font=CTkFont(weight="bold")).grid(row=current_row, column=0, columnspan=2, pady=(10,2), sticky="w", padx=10)
        current_row += 1

        self._create_editable_row(shipping_frame, current_row, "ค่าส่งเข้าไซต์:", data.get("shipping_to_site_cost"), key="shipping_to_site_cost", is_numeric=True)
        current_row += 1
        
        CTkLabel(shipping_frame, text="VAT 7%:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        site_vat_display = CTkEntry(shipping_frame, state="readonly", fg_color="gray85")
        site_vat_display.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries["shipping_to_site_vat_display"] = site_vat_display
        current_row += 1

        CTkLabel(shipping_frame, text="วันที่ส่งเข้าไซต์:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        site_date_selector = DateSelector(shipping_frame)
        site_date_selector.set_date(data.get("shipping_to_site_date"))
        site_date_selector.grid(row=current_row, column=1, padx=10, pady=5, sticky="w")
        self.po_entries["shipping_to_site_date"] = site_date_selector
        current_row += 1
        
        self._create_dropdown_row(shipping_frame, current_row, "ประเภท VAT", data.get("shipping_to_site_vat_type"), key="shipping_to_site_vat_type", options=["VAT", "CASH"])
        current_row += 1
        
        self._create_dropdown_row(shipping_frame, current_row, "ผู้จัดส่ง", data.get("shipping_to_site_shipper"), key="shipping_to_site_shipper", options=shipper_options)
        current_row += 1

        self._create_dropdown_row(shipping_frame, current_row, "หัก ณ ที่จ่าย", data.get("shipping_to_site_wht_type"), key="shipping_to_site_wht_type", options=wht_options)
        current_row += 1
        
        # --- START: เพิ่มช่องแสดงยอดหัก ณ ที่จ่าย (ไซต์) ---
        CTkLabel(shipping_frame, text="ยอดหัก ณ ที่จ่าย:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        site_wht_display = CTkEntry(shipping_frame, state="readonly", fg_color="gray85")
        site_wht_display.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries["shipping_to_site_wht_display"] = site_wht_display
        current_row += 1
        # --- END ---
        
        self._create_editable_row(shipping_frame, current_row, "หมายเหตุ:", data.get("shipping_to_site_notes"), key="shipping_to_site_notes")
        current_row += 1

        self._create_editable_row(shipping_frame, current_row, "ค่าย้าย:", data.get("relocation_cost"), key="relocation_cost", is_numeric=True)

    def _create_summary_section(self, parent, data):
        summary_frame = self._create_section(parent, "สรุปยอด")
        
        # --- กำหนดสไตล์สำหรับป้ายแสดงผล ---
        display_font = CTkFont(size=14, weight="bold")
        grand_total_font = CTkFont(size=18, weight="bold")
        display_fg_color = ("gray85", "gray18") # สีพื้นหลังสำหรับโหมดสว่าง/มืด

        # --- ยอดรวมต้นทุนสินค้า ---
        CTkLabel(summary_frame, text="ยอดรวมต้นทุนสินค้า:", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.total_cost_label = CTkLabel(summary_frame, text="0.00", font=display_font, fg_color=display_fg_color, corner_radius=5)
        self.total_cost_label.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        # --- ส่วนลดท้ายบิล ---
        self._create_editable_row(summary_frame, 2, "ส่วนลดท้ายบิล:", data.get("bill_discount"), key="bill_discount", is_numeric=True)
        
        # --- น้ำหนักรวม ---
        CTkLabel(summary_frame, text="น้ำหนักรวม:", anchor="w").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.total_weight_label = CTkLabel(summary_frame, text="0.00 kg", font=display_font, fg_color=display_fg_color, corner_radius=5)
        self.total_weight_label.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # --- ภาษีหัก ณ ที่จ่าย (3%) ---
        self.wht_checkbox = CTkCheckBox(summary_frame, text="ภาษีหัก ณ ที่จ่าย (3%):", command=self._recalculate_summary_totals)
        self.wht_checkbox.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        wht_entry = FormattedNumericEntry(summary_frame, command=self._recalculate_summary_totals)
        wht_entry.set(data.get("wht_3_percent", 0.0))
        wht_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew") 
        self.po_entries['wht_3_percent'] = wht_entry
        if data.get("wht_3_percent_checked"):
            self.wht_checkbox.select()

        # --- ภาษีมูลค่าเพิ่ม (7%) ---
        self.vat_checkbox = CTkCheckBox(summary_frame, text="ภาษีมูลค่าเพิ่ม (7%):", command=self._recalculate_summary_totals)
        self.vat_checkbox.grid(row=5, column=0, padx=10, pady=5, sticky="w")
        vat_entry = FormattedNumericEntry(summary_frame, command=self._recalculate_summary_totals)
        vat_entry.set(data.get("vat_7_percent", 0.0))
        vat_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries['vat_7_percent'] = vat_entry
        if data.get("vat_7_percent_checked") is not False:
            self.vat_checkbox.select()

        # --- ยอดรวมที่ต้องชำระ ---
        CTkLabel(summary_frame, text="ยอดรวมที่ต้องชำระ:", anchor="w", font=CTkFont(size=16, weight="bold")).grid(row=6, column=0, padx=10, pady=(10, 5), sticky="w")
        self.grand_total_label = CTkLabel(summary_frame, text="0.00", font=grand_total_font, fg_color="#16A34A", text_color="white", corner_radius=5)
        self.grand_total_label.grid(row=6, column=1, padx=10, pady=(10, 5), sticky="ew")


    # <<< START: เพิ่ม Section ใหม่สำหรับข้อมูลการอนุมัติ >>>
    def _create_approval_info_section(self, parent, data):
        approval_frame = self._create_section(parent, "ข้อมูลการอนุมัติและประวัติ")
        self._add_display_row(approval_frame, 1, "สร้างโดย", data.get("user_key"))
        self._add_display_row(approval_frame, 2, "สร้างเมื่อ", data.get("timestamp"))
        
        if data.get("approver_manager1_key"):
            self._add_display_row(approval_frame, 3, "อนุมัติโดย Manager 1", f"{data.get('approver_manager1_key')} (เมื่อ: {pd.to_datetime(data.get('approval_date_manager1')).strftime('%d/%m/%Y %H:%M') if pd.notna(data.get('approval_date_manager1')) else '-'})")
        if data.get("approver_manager2_key"):
            self._add_display_row(approval_frame, 4, "อนุมัติโดย Manager 2", f"{data.get('approver_manager2_key')} (เมื่อ: {pd.to_datetime(data.get('approval_date_manager2')).strftime('%d/%m/%Y %H:%M') if pd.notna(data.get('approval_date_manager2')) else '-'})")
        if data.get("approver_director_key"):
            self._add_display_row(approval_frame, 5, "อนุมัติโดย Director", f"{data.get('approver_director_key')} (เมื่อ: {pd.to_datetime(data.get('approval_date_director')).strftime('%d/%m/%Y %H:%M') if pd.notna(data.get('approval_date_director')) else '-'})")
        if data.get("rejection_reason"):
             self._add_display_row(approval_frame, 6, "เหตุผลที่ถูกปฏิเสธ", data.get("rejection_reason"))
    # <<< END >>>

    def _recalculate_summary_totals(self, *args):
        if not self.winfo_exists():
            return
            
        total_cost = 0.0
        total_weight = 0.0
        
        for item_row in self.item_entries:
            try:
                widgets = item_row['widgets']
                if not widgets['quantity'].winfo_exists(): continue
                qty = widgets['quantity'].get_value()
                price = widgets['unit_price'].get_value()
                discount = widgets['discount_value'].get_value()
                weight = widgets['total_weight'].get_value()
                discount_type = widgets['discount_type_var'].get()
                
                line_total = qty * price
                discount_amount = (line_total * (discount / 100.0)) if discount_type == '%' else discount
                item_total = line_total - discount_amount

                if widgets['total_price_label'].winfo_exists():
                    widgets['total_price_label'].configure(text=f"{item_total:,.2f}")
                
                total_cost += item_total
                total_weight += weight
            except (ValueError, TypeError, KeyError):
                total_price_label = item_row.get('widgets', {}).get('total_price_label')
                if total_price_label and total_price_label.winfo_exists():
                    total_price_label.configure(text="Error")

        if hasattr(self, 'total_cost_label') and self.total_cost_label.winfo_exists():
            self.total_cost_label.configure(text=f"{total_cost:,.2f}")
        if hasattr(self, 'total_weight_label') and self.total_weight_label.winfo_exists():
            self.total_weight_label.configure(text=f"{total_weight:,.2f} kg")
        
        try:
            if not self.po_entries['shipping_to_stock_cost'].winfo_exists(): return
            
            shipping_stock = self.po_entries['shipping_to_stock_cost'].get_value()
            shipping_stock_vat_type = self.po_entries['shipping_to_stock_vat_type'].get()
            shipping_stock_wht_type = self.po_entries['shipping_to_stock_wht_type'].get()
            
            shipping_site = self.po_entries['shipping_to_site_cost'].get_value()
            shipping_site_vat_type = self.po_entries['shipping_to_site_vat_type'].get()
            shipping_site_wht_type = self.po_entries['shipping_to_site_wht_type'].get()

            relocation_cost = self.po_entries['relocation_cost'].get_value()
            bill_discount = self.po_entries['bill_discount'].get_value()
            wht_entry = self.po_entries.get('wht_3_percent')
            vat_entry = self.po_entries.get('vat_7_percent')
        
        except (KeyError, ValueError, TypeError):
            return

        stock_vat_amount = shipping_stock * 0.07 if shipping_stock_vat_type == 'VAT' else 0.0
        site_vat_amount = shipping_site * 0.07 if shipping_site_vat_type == 'VAT' else 0.0

        if self.po_entries.get("shipping_to_stock_vat_display"): utils.set_entry_text(self.po_entries["shipping_to_stock_vat_display"], f"{stock_vat_amount:,.2f}")
        if self.po_entries.get("shipping_to_site_vat_display"): utils.set_entry_text(self.po_entries["shipping_to_site_vat_display"], f"{site_vat_amount:,.2f}")

        shipping_stock_wht_amount = shipping_stock * (0.01 if shipping_stock_wht_type == '1%' else 0.03 if shipping_stock_wht_type == '3%' else 0)
        shipping_site_wht_amount = shipping_site * (0.01 if shipping_site_wht_type == '1%' else 0.03 if shipping_site_wht_type == '3%' else 0)

        if self.po_entries.get("shipping_to_stock_wht_display"): utils.set_entry_text(self.po_entries["shipping_to_stock_wht_display"], f"{shipping_stock_wht_amount:,.2f}")
        if self.po_entries.get("shipping_to_site_wht_display"): utils.set_entry_text(self.po_entries["shipping_to_site_wht_display"], f"{shipping_site_wht_amount:,.2f}")

        base_for_tax = total_cost - bill_discount
        if shipping_stock_vat_type == 'VAT': base_for_tax += shipping_stock
        if shipping_site_vat_type == 'VAT': base_for_tax += shipping_site

        vat_amount_total = base_for_tax * 0.07 if hasattr(self, 'vat_checkbox') and self.vat_checkbox.get() == 1 else 0.0
        wht_amount_products = base_for_tax * 0.03 if hasattr(self, 'wht_checkbox') and self.wht_checkbox.get() == 1 else 0.0

        if wht_entry and wht_entry.winfo_exists(): wht_entry.set(wht_amount_products)
        if vat_entry and vat_entry.winfo_exists(): vat_entry.set(vat_amount_total)

        non_vat_costs = 0.0
        if shipping_stock_vat_type == 'CASH': non_vat_costs += shipping_stock
        if shipping_site_vat_type == 'CASH': non_vat_costs += shipping_site
        non_vat_costs += relocation_cost
        
        total_wht_deduction = wht_amount_products + shipping_stock_wht_amount + shipping_site_wht_amount
        grand_total = base_for_tax + vat_amount_total - total_wht_deduction + non_vat_costs
        
        if hasattr(self, 'grand_total_label') and self.grand_total_label.winfo_exists():
            self.grand_total_label.configure(text=f"{grand_total:,.2f}")

    
    def _create_payments_section(self, parent, payments_list):
        """แก้ไขส่วนการชำระเงินให้แสดงผลถูกต้องและไม่ซ้อนทับ"""
        self.payments_frame = self._create_section(parent, "การชำระเงิน")
        self.payments_frame.grid_columnconfigure(0, weight=1)

        if payments_list:
            # มีข้อมูล Payment: สร้าง Headers และแสดง Rows
            self._create_payment_headers()
            
            # สร้าง Content Frame เมื่อมีข้อมูลเท่านั้น
            self.payments_content_frame = CTkFrame(self.payments_frame, fg_color="transparent")
            self.payments_content_frame.grid(row=2, column=0, sticky="ew", padx=10)
            self.payments_content_frame.grid_columnconfigure(0, weight=1)
            
            for payment in payments_list:
                self._add_payment_row(payment)
                
            # ปุ่มเพิ่ม - เมื่อมีข้อมูลแล้ว
            self._create_add_button(3)
        else:
            # ไม่มีข้อมูล Payment: แสดงข้อความสั้นๆ
            self.payment_empty_label = CTkLabel(
                self.payments_frame, 
                text="ยังไม่มีรายการชำระเงิน", 
                text_color="gray50", 
                font=CTkFont(size=14, slant="italic"),
                height=30
            )
            self.payment_empty_label.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
            
            # ปุ่มเพิ่ม - เมื่อยังไม่มีข้อมูล
            self._create_add_button(2)
    
    def _create_add_button(self, row_position):
        """สร้างปุ่ม 'เพิ่มการชำระเงิน' ในตำแหน่งที่ถูกต้อง"""
        self.add_payment_button = CTkButton(
            self.payments_frame, 
            text="+ เพิ่มการชำระเงิน", 
            command=self._handle_add_payment_button_click,
            width=180,
            height=35
        )
        self.add_payment_button.grid(row=row_position, column=0, pady=(10, 5), padx=10, sticky="w")


    def _create_payment_headers(self):
        """สร้างหัวตารางสำหรับ Payment - เรียกใช้เฉพาะเมื่อมีข้อมูล"""
        headers = ["ประเภทการชำระ", "ยอดเงิน", "วันที่ชำระ", "ธนาคาร", "เลขที่บัญชี"]
        col_weights = [2, 2, 2, 2, 3]
        
        self.payments_header_container = CTkFrame(self.payments_frame, fg_color="transparent")
        self.payments_header_container.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))

        for i, header_text in enumerate(headers):
            self.payments_header_container.grid_columnconfigure(i, weight=col_weights[i])
            CTkLabel(
                self.payments_header_container, 
                text=header_text, 
                font=CTkFont(size=14, weight="bold")
            ).grid(row=0, column=i, padx=5, pady=5, sticky="w")

    def _handle_add_payment_button_click(self):
        """จัดการการเพิ่ม Payment แถวใหม่ - แก้ไขการจัดวาง"""
        # ตรวจสอบว่าเป็นการเพิ่มครั้งแรกหรือไม่
        is_first_add = (hasattr(self, 'payment_empty_label') and 
                        self.payment_empty_label is not None and 
                        self.payment_empty_label.winfo_exists())

        if is_first_add:
            # ลบป้าย "ยังไม่มีข้อมูล" ทิ้ง
            self.payment_empty_label.destroy()
            self.payment_empty_label = None
            
            # ลบปุ่มเพิ่มเก่า
            if hasattr(self, 'add_payment_button'):
                self.add_payment_button.destroy()
            
            # สร้าง Headers ขึ้นมาใหม่
            self._create_payment_headers()
            
            # สร้าง Content Frame สำหรับแถวข้อมูล
            self.payments_content_frame = CTkFrame(self.payments_frame, fg_color="transparent")
            self.payments_content_frame.grid(row=2, column=0, sticky="ew", padx=10)
            self.payments_content_frame.grid_columnconfigure(0, weight=1)
            
            # สร้างปุ่มเพิ่มในตำแหน่งใหม่
            self._create_add_button(3)

        # เพิ่มข้อมูลเปล่าเข้าไปใน data list
        if not hasattr(self, 'payments_data'):
            self.payments_data = []
        self.payments_data.append({})

        # เพิ่มแถว Widget ใหม่
        self._add_payment_row()

        
    def _add_new_payment_and_redraw(self):
        """
        เพิ่มข้อมูลการชำระเงินใหม่เข้าไปใน data list และสั่ง redraw UI ใหม่ทั้งหมด
        เพื่อให้ UI แสดงผลตามสถานะข้อมูลล่าสุด
        """
        if not hasattr(self, 'payments_data'):
            self.payments_data = []
        self.payments_data.append({})

        # --- START: จุดแก้ไข ---
        # เปลี่ยนจากการเรียกตรงๆ มาเป็นการหน่วงเวลาเล็กน้อยด้วย after_idle
        # เพื่อรอให้ event อื่นๆ จัดการตัวเองให้เสร็จก่อน
        self.after_idle(self._create_formatted_view)
    # --- END ---

    def _add_payment_row(self, payment_data=None):
        if payment_data is None: 
            payment_data = {}

        row_frame = CTkFrame(self.payments_content_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=2)
        
        # แก้ไข: กำหนดน้ำหนักคอลัมน์ให้ถูกต้องและเพิ่มคอลัมน์สำหรับปุ่มลบ
        row_frame.grid_columnconfigure(0, weight=2)  # ประเภทการชำระ
        row_frame.grid_columnconfigure(1, weight=2)  # ยอดเงิน  
        row_frame.grid_columnconfigure(2, weight=2)  # วันที่ชำระ
        row_frame.grid_columnconfigure(3, weight=2)  # ธนาคาร
        row_frame.grid_columnconfigure(4, weight=3)  # เลขที่บัญชี
        row_frame.grid_columnconfigure(5, weight=0)  # ปุ่มลบ (ไม่ขยาย)

        # ประเภทการชำระ
        payment_types = ["Payment 1", "Payment 2", "Full Payment", "CN Refund"]
        type_var = tk.StringVar(value=payment_data.get('payment_type', payment_types[0]))
        type_menu = CTkOptionMenu(row_frame, variable=type_var, values=payment_types, width=120)
        type_menu.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        # ยอดเงิน
        amount_entry = FormattedNumericEntry(row_frame)
        amount_entry.set(payment_data.get('amount', 0.0))
        amount_entry.grid(row=0, column=1, padx=5, sticky="ew")
        
        # วันที่ชำระ
        date_selector = DateSelector(row_frame)
        date_selector.set_date(payment_data.get('payment_date'))
        date_selector.grid(row=0, column=2, padx=5, sticky="ew")
        
        # ธนาคาร
        bank_list = ["ระบุเอง", "BBL", "KBANK", "KTB", "SCB", "TTB", "BAY", "GSB", "BAAC", "UOB", "CIMB"]
        bank_var = tk.StringVar(value=payment_data.get('bank_name', bank_list[0]))
        bank_menu = CTkOptionMenu(row_frame, 
                                variable=bank_var, 
                                values=bank_list,
                                width=100,
                                command=lambda bank=bank_var.get(), acc_entry=None: self._on_bank_selected(bank, acc_entry))
        bank_menu.grid(row=0, column=3, padx=5, sticky="ew")
        
        # เลขที่บัญชี
        account_entry = CTkEntry(row_frame)
        account_number = payment_data.get('bank_account_number')
        if account_number is not None and pd.notna(account_number):
            account_entry.insert(0, str(account_number))
        account_entry.grid(row=0, column=4, padx=5, sticky="ew")
        
        # แก้ไข trace สำหรับ bank selection
        bank_var.trace_add("write", lambda *args, bv=bank_var, ae=account_entry: self._on_bank_selected(bv.get(), ae))
        
        # ปุ่มลบ - แก้ไขตำแหน่งและขนาด
        delete_button = CTkButton(row_frame, 
                                text="ลบ", 
                                width=50,  # กำหนดความกว้างคงที่
                                height=32, # กำหนดความสูงให้พอดีกับ entry
                                fg_color="#DC2626", 
                                hover_color="#B91C1C",
                                command=lambda r=row_frame, p_id=payment_data.get('id'): self._remove_payment_row(r, p_id))
        delete_button.grid(row=0, column=5, padx=(5, 0), sticky="")  # ไม่ใช้ sticky="ew"

        # เก็บ reference ของ widgets
        self.payment_entries.append({
            'id': payment_data.get('id'),
            'frame': row_frame,
            'widgets': {
                'type_var': type_var,
                'amount_entry': amount_entry,
                'date_selector': date_selector,
                'bank_var': bank_var,
                'account_entry': account_entry
            }
        })
    
    def _remove_payment_row(self, row_frame, payment_id):
        if payment_id is not None:
            self.deleted_payment_ids.append(payment_id)
        
        index_to_remove = -1
        for i, entry in enumerate(self.payment_entries):
            if entry['frame'] == row_frame:
                index_to_remove = i
                break
        
        if index_to_remove != -1:
            self.payment_entries.pop(index_to_remove)
        
        row_frame.destroy()
        self._recalculate_summary_totals()

    def _create_action_buttons(self):
        # สร้าง Frame ใหม่สำหรับวางปุ่ม และวางไว้ที่แถวที่ 1 ของหน้าต่างหลัก (self)
        self.button_frame = CTkFrame(self, fg_color=("gray85", "gray18"))
        self.button_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        
        # --- โค้ดเวอร์ชันเดิมที่ไม่มีปุ่มลบ ---
        if self.user_role in ['Purchasing Manager', 'Director', 'HR']:
            self.button_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
            
            approve_button = CTkButton(self.button_frame, text="อนุมัติ (Approve)", command=self._approve_po, fg_color="#16A34A", hover_color="#15803D")
            approve_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

            reject_button = CTkButton(self.button_frame, text="ปฏิเสธ (Reject)", command=self._reject_po, fg_color="#DC2626", hover_color="#B91C1C")
            reject_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

            save_button = CTkButton(self.button_frame, text="บันทึกการแก้ไข", command=self._save_changes, fg_color="#3B82F6", hover_color="#2563EB")
            save_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

            close_button = CTkButton(self.button_frame, text="ปิด", command=self.destroy, fg_color="gray")
            close_button.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        else:
            self.button_frame.grid_columnconfigure((0, 1), weight=1)
            save_button = CTkButton(self.button_frame, text="บันทึกการแก้ไข", command=self._save_changes)
            save_button.grid(row=0, column=0, padx=(0,5), pady=5, sticky="ew")
        
            close_button = CTkButton(self.button_frame, text="ปิด", command=self.destroy, fg_color="gray")
            close_button.grid(row=0, column=1, padx=(5,0), pady=5, sticky="ew")
    
    def _approve_po(self):
        # Logic การอนุมัติ (ยกมาจาก purchasing_manager_screen.py)
        # สามารถนำ Logic ที่ซับซ้อนเกี่ยวกับการอนุมัติตามลำดับขั้นและยอดเงินมาใส่ที่นี่ได้
        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการอนุมัติ PO ID: {self.purchase_id} ใช่หรือไม่?", parent=self):
            return
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ตัวอย่าง Logic การอนุมัติแบบง่าย
                cursor.execute(
                    "UPDATE purchase_orders SET status = 'Approved', approval_status = 'Approved', approver_manager1_key = %s, approval_date_manager1 = %s WHERE id = %s",
                    (self.app_container.current_user_key, datetime.now(), self.purchase_id)
                )
            conn.commit()
            messagebox.showinfo("สำเร็จ", "อนุมัติ PO เรียบร้อยแล้ว", parent=self)
            if self.on_save_callback:
                self.on_save_callback() # Refresh หน้าจอหลัก
            self.destroy()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _reject_po(self):
        # Logic การปฏิเสธ (ยกมาจาก purchasing_manager_screen.py)
        dialog = RejectionReasonDialog(self)
        self.wait_window(dialog)
        reason = getattr(dialog, '_reason_string', None)
        if reason is None:
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    "UPDATE purchase_orders SET status = 'Rejected', approval_status = 'Rejected', rejection_reason = %s, last_modified_by = %s WHERE id = %s",
                    (reason.strip(), self.app_container.current_user_key, self.purchase_id)
                )
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ปฏิเสธ PO เรียบร้อยแล้ว", parent=self)
            if self.on_save_callback:
                self.on_save_callback() # Refresh หน้าจอหลัก
            self.destroy()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _adjust_window_height_to_content(self):
        """ปรับตำแหน่งหน้าต่างให้อยู่ตรงกลาง (ไม่ปรับขนาดอัตโนมัติแล้ว)"""
        # เราจะลบโค้ดคำนวณขนาดที่ซับซ้อนออกทั้งหมด
        # ให้ฟังก์ชันนี้ทำหน้าที่แค่ 'จัดตำแหน่ง' หน้าต่างหลังจากที่ UI ถูกวาดเสร็จแล้วเท่านั้น
        self._position_window()

        # --- 1. คำนวณความสูงที่เหมาะสม (เหมือนเดิม) ---
        content_height = self.scroll_frame.winfo_reqheight()
        buttons_height = self.button_frame.winfo_reqheight()
        total_needed_height = content_height + buttons_height + 40
        
        screen_height = self.winfo_screenheight()
        max_height = screen_height - 80
        final_height = min(total_needed_height, max_height)
        
        ### START: เพิ่มโค้ดคำนวณความกว้าง ###
        # --- 2. คำนวณความกว้างที่เหมาะสม ---
        # ดึงความกว้างที่เนื้อหาต้องการ และบวกเผื่อระยะขอบและ scrollbar
        content_width = self.scroll_frame.winfo_reqwidth() + 40 
        
        screen_width = self.winfo_screenwidth()
        # กำหนดความกว้างสูงสุดไม่ให้เกินขอบจอ
        max_width = screen_width - 80
        
        # เลือกใช้ความกว้างที่เนื้อหาต้องการ แต่ต้องไม่น้อยกว่า 950 และไม่เกินขอบจอ
        final_width = max(950, min(content_width, max_width))
        ### END ###

        # 3. ปรับขนาดหน้าต่างด้วยค่าที่คำนวณได้ใหม่ทั้งหมด
        self.geometry(f"{final_width}x{final_height}")
        
        # 4. จัดตำแหน่งหน้าต่างให้อยู่ตรงกลาง (เหมือนเดิม)
        self._position_window()

    def _position_window(self):
        """
        (เวอร์ชันใหม่) คำนวณขนาดที่เหมาะสมไม่ให้เกิน 90% ของหน้าจอ และจัดตำแหน่งให้อยู่ตรงกลาง
        """
        self.update_idletasks()

        # --- 1. คำนวณขนาดที่เหมาะสม ---
        # หาขนาดหน้าจอที่หน้าต่างนี้ปรากฏอยู่
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # กำหนดขนาดหน้าต่างสูงสุดไม่ให้เกิน 90% ของหน้าจอ
        max_width = int(screen_width * 0.9)
        max_height = int(screen_height * 0.9)
        
        # --- 2. ตั้งค่าขนาดหน้าต่าง ---
        # ใช้ขนาดคงที่ที่เราคิดว่าเหมาะสม (เช่น 1200x850)
        # แต่ถ้ามันใหญ่กว่าขนาดสูงสุดที่คำนวณได้ ให้ใช้ขนาดสูงสุดแทน
        final_width = min(1350, max_width)
        final_height = min(850, max_height)
        
        self.geometry(f"{final_width}x{final_height}")

        # --- 3. จัดตำแหน่งให้อยู่กลางจอ ---
        # ใช้ขนาดสุดท้ายที่คำนวณได้มาจัดตำแหน่ง
        x = (screen_width // 2) - (final_width // 2)
        y = (screen_height // 2) - (final_height // 2)
        
        # ตั้งค่าตำแหน่ง
        self.geometry(f"+{x}+{y}")

    def _save_changes(self):
        self._recalculate_summary_totals()
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # --- START: เพิ่มโค้ดสำหรับตรวจสอบการเปลี่ยนแปลง SO ---
                original_so_number = self.po_data.get('so_number', '').strip()
                new_so_number = self.po_entries['so_number'].get().strip()
                # --- END ---

                # --- START: เพิ่ม Logic การอัปเดตข้อมูล Supplier ก่อน ---
                supplier_name = self.po_entries['supplier_name'].get().strip()
                supplier_code = self.po_entries['supplier_code'].get().strip()
                
                if supplier_name and supplier_code:
                    cursor.execute("SELECT id FROM suppliers WHERE supplier_name = %s", (supplier_name,))
                    existing_supplier = cursor.fetchone()
                    if existing_supplier:
                        cursor.execute("UPDATE suppliers SET supplier_code = %s, credit_term = %s WHERE id = %s", 
                                    (supplier_code, self.po_entries['credit_term'].get(), existing_supplier[0]))
                    else:
                        cursor.execute("INSERT INTO suppliers (supplier_name, supplier_code, credit_term) VALUES (%s, %s, %s)",
                                    (supplier_name, supplier_code, self.po_entries['credit_term'].get()))

                total_cost = utils.convert_to_float(self.total_cost_label.cget("text"))
                grand_total = utils.convert_to_float(self.grand_total_label.cget("text"))

                cursor.execute("""
                    UPDATE purchase_orders SET 
                        so_number = %s, po_number = %s, supplier_name = %s, 
                        credit_term = %s, po_mode = %s,
                        shipping_to_stock_cost = %s, shipping_to_stock_date = %s, 
                        shipping_to_site_cost = %s, shipping_to_site_date = %s, 
                        relocation_cost = %s, total_cost = %s, grand_total = %s,
                        shipping_to_stock_vat_type = %s, shipping_to_stock_shipper = %s,
                        shipping_to_stock_wht_type = %s, shipping_to_stock_notes = %s,
                        shipping_to_site_vat_type = %s, shipping_to_site_shipper = %s,
                        shipping_to_site_wht_type = %s, shipping_to_site_notes = %s,
                        wht_3_percent_checked = %s, vat_7_percent_checked = %s,
                        bill_discount = %s
                    WHERE id = %s
                """, (
                    new_so_number, # <-- ใช้ตัวแปรใหม่
                    self.po_entries['po_number'].get(),
                    supplier_name,
                    self.po_entries['credit_term'].get(), self.po_entries['po_mode'].get(),
                    self.po_entries['shipping_to_stock_cost'].get_value(),
                    self.po_entries['shipping_to_stock_date'].get_date(),
                    self.po_entries['shipping_to_site_cost'].get_value(),
                    self.po_entries['shipping_to_site_date'].get_date(),
                    self.po_entries['relocation_cost'].get_value(),
                    total_cost, grand_total,
                    self.po_entries['shipping_to_stock_vat_type'].get(),
                    self.po_entries['shipping_to_stock_shipper'].get(),
                    self.po_entries['shipping_to_stock_wht_type'].get(),
                    self.po_entries['shipping_to_stock_notes'].get(),
                    self.po_entries['shipping_to_site_vat_type'].get(),
                    self.po_entries['shipping_to_site_shipper'].get(),
                    self.po_entries['shipping_to_site_wht_type'].get(),
                    self.po_entries['shipping_to_site_notes'].get(),
                    bool(self.wht_checkbox.get()),
                    bool(self.vat_checkbox.get()),
                    self.po_entries['bill_discount'].get_value(),
                    self.purchase_id
                ))
                
                if self.deleted_item_ids:
                    cursor.execute("DELETE FROM purchase_order_items WHERE id IN %s", (tuple(self.deleted_item_ids),))

                for item_row in self.item_entries:
                    widgets, item_id = item_row['widgets'], item_row['id']
                    code, name = widgets['product_code'].get(), widgets['product_name'].get()
                    warehouse = widgets['warehouse'].get()
                    weight, qty = widgets['total_weight'].get_value(), widgets['quantity'].get_value()
                    price, discount = widgets['unit_price'].get_value(), widgets['discount_value'].get_value()
                    discount_type = widgets['discount_type_var'].get()
                    line_total = qty * price
                    discount_amount = (line_total * (discount / 100.0)) if discount_type == "%" else discount
                    total = line_total - discount_amount
                    
                    if item_id:
                        cursor.execute("""
                            UPDATE purchase_order_items SET 
                                product_code = %s, product_name = %s, warehouse = %s, total_weight = %s, 
                                quantity = %s, unit_price = %s, discount_value = %s, 
                                discount_type = %s, total_price = %s 
                            WHERE id = %s
                        """, (code, name, warehouse, weight, qty, price, discount, discount_type, total, item_id))
                    else:
                        cursor.execute("""
                            INSERT INTO purchase_order_items 
                            (purchase_order_id, product_code, product_name, warehouse, total_weight, 
                            quantity, unit_price, discount_value, discount_type, total_price) 
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (self.purchase_id, code, name, warehouse, weight, qty, price, discount, discount_type, total))

                if self.deleted_payment_ids:
                    cursor.execute("DELETE FROM purchase_order_payments WHERE id IN %s", (tuple(self.deleted_payment_ids),))
                for payment_row in self.payment_entries:
                    widgets, payment_id = payment_row['widgets'], payment_row['id']
                    p_type, p_amount, p_date = widgets['type_var'].get(), widgets['amount_entry'].get_value(), widgets['date_selector'].get_date()
                    p_bank, p_account = widgets['bank_var'].get(), widgets['account_entry'].get()
                    if payment_id:
                        cursor.execute("UPDATE purchase_order_payments SET payment_type = %s, amount = %s, payment_date = %s, bank_name = %s, bank_account_number = %s WHERE id = %s", (p_type, p_amount, p_date, p_bank, p_account, payment_id))
                    else:
                        cursor.execute("INSERT INTO purchase_order_payments (purchase_order_id, payment_type, amount, payment_date, bank_name, bank_account_number) VALUES (%s, %s, %s, %s, %s, %s)", (self.purchase_id, p_type, p_amount, p_date, p_bank, p_account))
                
                log_details = { "message": f"PO ID {self.purchase_id} edited by {self.user_role} ({self.app_container.current_user_key})" }
                cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp) VALUES (%s, %s, %s, %s, %s, %s)", ('PO Edited', 'purchase_orders', self.purchase_id, self.app_container.current_user_key, json.dumps(log_details, default=str), datetime.now()))
            
            conn.commit()

            # --- START: ปรับปรุงข้อความยืนยัน ---
            if original_so_number != new_so_number and new_so_number != '':
                po_number_display = self.po_entries['po_number'].get()
                success_message = (f"บันทึกสำเร็จ!\n\n"
                                   f"PO '{po_number_display}' ถูกย้ายจาก SO '{original_so_number}' "
                                   f"ไปยัง SO '{new_so_number}' เรียบร้อยแล้ว")
            else:
                success_message = "บันทึกการแก้ไข PO เรียบร้อยแล้ว"
            
            messagebox.showinfo("สำเร็จ", success_message, parent=self)
            # --- END ---
            
            if self.on_save_callback: 
                self.on_save_callback()
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)


class PurchaseHistoryWindow(CTkToplevel):

    def _debounce_search(self, event=None):
        """ยกเลิกการค้นหาเก่าและตั้งเวลาใหม่ทุกครั้งที่พิมพ์"""
        # หากมี job ที่ตั้งเวลาไว้ก่อนหน้า ให้ยกเลิกไป
        if self._debounce_job:
            self.after_cancel(self._debounce_job)

        # ตั้งเวลาเพื่อเรียกฟังก์ชันค้นหาจริงในอีก 500 มิลลิวินาที (0.5 วินาที)
        self._debounce_job = self.after(500, self._apply_filters)

    def __init__(self, master, app_container, on_save_callback=None, **kwargs):
        super().__init__(master)
        self.title("ประวัติใบสั่งซื้อ (PO History)")
        self.geometry("1200x700")
        
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.on_save_callback = on_save_callback
        self.user_role = self.app_container.current_user_role
        
        # --- ตัวแปรสำหรับ Pagination และ Filter ---
        self.all_po_df = None
        self.filtered_df = None
        self.current_page = 0
        self.rows_per_page = 50
        
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # --- สร้าง UI Layout ใหม่ ---
        self._create_new_layout()
        
        # --- เริ่มโหลดข้อมูล ---
        self.after(50, self._load_initial_data)
        
        self.transient(master)
        self.grab_set()
    
    def _create_new_layout(self):
        """สร้าง UI Layout ใหม่ทั้งหมดสำหรับหน้าต่างประวัติ PO"""
        # --- Top Frame (Filter & Pagination) ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")

        # --- Filter Section ---
        filter_frame = CTkFrame(top_frame, fg_color="transparent")
        filter_frame.pack(side="left")

        month_options = ["ทุกเดือน"] + self.thai_months
        CTkOptionMenu(filter_frame, variable=self.month_var, values=month_options).pack(side="left", padx=5)

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        CTkOptionMenu(filter_frame, variable=self.year_var, values=year_options).pack(side="left", padx=5)

        self.search_entry = CTkEntry(filter_frame, placeholder_text="ค้นหา SO, PO, Supplier...")
        self.search_entry.pack(side="left", padx=10, fill="x", expand=True)
        # --- แก้ไข: เปลี่ยน event จาก KeyRelease เป็น Debounce เพื่อประสิทธิภาพที่ดีกว่า ---
        self._debounce_job = None 
        self.search_entry.bind("<KeyRelease>", self._debounce_search)

        CTkButton(filter_frame, text="ค้นหา", command=self._apply_filters, width=80).pack(side="left")

        # --- Pagination Section ---
        pagination_frame = CTkFrame(top_frame, fg_color="transparent")
        pagination_frame.pack(side="right")

        self.prev_button = CTkButton(pagination_frame, text="<<", command=self._prev_page, width=50, state="disabled")
        self.prev_button.pack(side="left", padx=5)
        self.page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.page_label.pack(side="left", padx=5)
        self.next_button = CTkButton(pagination_frame, text=">>", command=self._next_page, width=50, state="disabled")
        self.next_button.pack(side="left", padx=5)

        # --- Main Frame for the Treeview ---
        # (เราจะสร้าง Treeview ข้างในฟังก์ชัน _update_treeview_display)
        self.history_frame = CTkFrame(self)
        self.history_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.history_frame.grid_rowconfigure(0, weight=1)
        self.history_frame.grid_columnconfigure(0, weight=1)

        # --- Loading Label ---
        self.loading_label = CTkLabel(self, text="กำลังโหลดข้อมูล...", font=CTkFont(size=18, slant="italic"), text_color="gray50")

    def _next_page(self):
        total_pages = (len(self.filtered_df) + self.rows_per_page - 1) // self.rows_per_page if hasattr(self, 'filtered_df') else 0
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._update_treeview_display()

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._update_treeview_display()

    def _show_loading(self):
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.update_idletasks()

    def _hide_loading(self):
        self.loading_label.place_forget()
    
    def _load_initial_data(self):
        for widget in self.history_frame.winfo_children(): widget.destroy()
        self._show_loading()
        try:
            # แก้ไข Query ให้ JOIN ตาราง user และแสดงชื่อเจ้าของ/ผู้สร้างแทน
            query = """
                SELECT 
                    po.id, 
                    po.timestamp, 
                    po.so_number, 
                    po.po_number, 
                    po.supplier_name,
                    owner.sale_name as owner_name,
                    proxy.sale_name as proxy_name
                FROM purchase_orders po
                LEFT JOIN sales_users owner ON po.user_key = owner.sale_key
                LEFT JOIN sales_users proxy ON po.proxy_user_key = proxy.sale_key
                ORDER BY po.timestamp DESC
            """
            self.all_po_df = pd.read_sql_query(query, self.pg_engine)
            self.all_po_df['timestamp'] = pd.to_datetime(self.all_po_df['timestamp'])
            self._hide_loading()
            self._apply_filters()
        except Exception as e:
            self._hide_loading()
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดประวัติ PO ได้: {e}", parent=self)
            traceback.print_exc()

    def _apply_filters(self):
        if self.all_po_df is None:
            return

        self.current_page = 0
        df = self.all_po_df.copy()

        selected_month_str = self.month_var.get()
        if selected_month_str != "ทุกเดือน":
            month_num = self.thai_month_map[selected_month_str]
            df = df[df['timestamp'].dt.month == month_num]

        selected_year_str = self.year_var.get()
        if selected_year_str != "ทุกปี":
            year_num = int(selected_year_str)
            df = df[df['timestamp'].dt.year == year_num]

        search_term = self.search_entry.get().strip().lower()
        if search_term:
            df = df[
                df['so_number'].str.lower().str.contains(search_term, na=False) |
                df['po_number'].str.lower().str.contains(search_term, na=False) |
                df['supplier_name'].str.lower().str.contains(search_term, na=False)
            ]

        self.filtered_df = df
        self._update_treeview_display()

    def _update_treeview_display(self):
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        if self.filtered_df.empty:
            CTkLabel(self.history_frame, text="ไม่พบข้อมูลตามเงื่อนไขที่เลือก").pack(pady=20)
            self.page_label.configure(text="Page 0 / 0")
            self.prev_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            return
            
        total_rows = len(self.filtered_df)
        total_pages = (total_rows + self.rows_per_page - 1) // self.rows_per_page
        
        start_row = self.current_page * self.rows_per_page
        end_row = start_row + self.rows_per_page
        df_page = self.filtered_df.iloc[start_row:end_row]

        self._create_styled_dataframe_table(self.history_frame, df_page)

        self.page_label.configure(text=f"Page {self.current_page + 1} / {max(1, total_pages)}")
        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.configure(state="normal" if self.current_page < total_pages - 1 else "disabled")

    def _on_row_double_click(self, event, tree):
        try:
            item_id = tree.focus()
            if not item_id: return
            item_values = tree.item(item_id)['values']
            
            # <<< START: แก้ไขจุดนี้ >>>
            # แปลงค่าที่ดึงมาให้เป็น int ปกติของ Python ก่อน
            purchase_id = int(item_values[0])
            # <<< END: สิ้นสุดการแก้ไข >>>
            
            self.app_container.show_purchase_detail_window(purchase_id)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดดูรายละเอียดได้: {e}", parent=self)

    def _create_styled_dataframe_table(self, parent, df):
        # สร้างคอลัมน์ใหม่สำหรับแสดงผล
        df['display_owner'] = df.apply(
            lambda row: f"{row['owner_name']}" if pd.isna(row['proxy_name']) else f"{row['owner_name']} (โดย {row['proxy_name']})",
            axis=1
        )
        
        # กำหนดคอลัมน์ที่จะแสดงในตาราง
        display_columns = {
            'timestamp': 'เวลาบันทึก',
            'so_number': 'SO Number',
            'po_number': 'PO Number',
            'supplier_name': 'Supplier',
            'display_owner': 'เจ้าของ PO (ผู้สร้าง)'
        }
        
        columns_to_show = list(display_columns.keys())
        df_display = df[columns_to_show]

        tree = ttk.Treeview(parent, columns=columns_to_show, show='headings')
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Roboto', 14, 'bold'))
        style.configure("Treeview", rowheight=25, font=('Roboto', 12))
        
        # ตั้งค่าหัวตารางและขนาด
        for col_id, col_text in display_columns.items():
            tree.heading(col_id, text=col_text)
            width = 200 # default
            if col_id == 'timestamp': width = 180
            if col_id == 'display_owner': width = 250
            tree.column(col_id, width=width, anchor='w')

        # ใส่ข้อมูลลงในตาราง (ใช้ df_display)
        for index, row in df_display.iterrows():
            # เก็บ id เดิมไว้ในตัวแปร iid เพื่อใช้ตอนดับเบิลคลิก
            original_id = df.loc[index, 'id']
            values = list(row)
            values[0] = row['timestamp'].strftime('%Y-%m-%d %H:%M:%S') # Format วันที่
            tree.insert("", "end", values=values, iid=original_id)

        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        # แก้ไข _on_row_double_click ให้ใช้ iid ที่เราเก็บไว้
        tree.bind("<Double-1>", lambda e: self._on_row_double_click(e, tree, use_iid=True))


class CommissionHistoryWindow(CTkToplevel):
    def __init__(self, master, app_container, sale_key_filter=None, on_row_double_click=None, support_user_key_filter=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.sale_key_filter = sale_key_filter
        self.on_row_double_click_callback = on_row_double_click
        self.support_user_key_filter = support_user_key_filter # <-- บรรทัดนี้จะทำงานได้ถูกต้องแล้ว
        self.df = None
        
        # --- ตัวแปรสำหรับ Pagination และ Filter ---
        self.current_page = 0
        self.rows_per_page = 50
        self.total_rows = 0
        self.total_pages = 0
        self.active_tab = "drafts"

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")
        
        self.title(f"ประวัติการบันทึกของ: {self.sale_key_filter}")
        self.geometry("1400x700")
        try: self.theme = self.app_container.THEME["sale"]
        except (AttributeError, KeyError): self.theme = {"header": "#1D4ED8", "primary": "#3B82F6"}
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._create_new_layout()
        
        self.after(50, self._populate_history_table)
        self.transient(master)
        self.grab_set()
        self.focus()
    
    def _on_tree_row_double_click(self, event, tree):
        """Callback เมื่อดับเบิลคลิกที่แถวใน Treeview"""
        try:
            selected_item = tree.focus()
            if not selected_item:
                return
            
            # ดึงข้อมูลจาก DataFrame โดยใช้ ID ที่เป็น iid
            record_id = int(selected_item)
            row_data = self.df[self.df['id'] == record_id].iloc[0]

            if self.on_row_double_click_callback:
                self.on_row_double_click_callback(row_data)
        except (ValueError, IndexError) as e:
            print(f"Could not process double click: {e}")

    def _cancel_selected_record(self):
        try:
            # ตรวจสอบว่ามี Treeview และมีรายการที่ถูกเลือกหรือไม่
            if not hasattr(self, 'tree') or not self.tree.focus():
                messagebox.showwarning("ไม่ได้เลือกรายการ", "กรุณาเลือกรายการที่ต้องการยกเลิก", parent=self)
                return

            item_id = self.tree.focus()
            selected_index = self.tree.index(item_id)
            record_data = self.df.iloc[selected_index]

            record_id = record_data['id']
            record_status = record_data['status']
            record_so = record_data['so_number']

            # อนุญาตให้ยกเลิกได้เฉพาะสถานะที่ยังไม่ได้ส่งเท่านั้น
            if record_status not in ['Original', 'Edited']:
                messagebox.showerror("ไม่สามารถยกเลิกได้", 
                                     f"ไม่สามารถยกเลิกรายการนี้ได้ เนื่องจากมีสถานะเป็น '{record_status}'\n"
                                     "(ยกเลิกได้เฉพาะรายการที่เป็นฉบับร่างเท่านั้น)", 
                                     parent=self)
                return

            if messagebox.askyesno("ยืนยันการยกเลิก", 
                                   f"คุณต้องการยกเลิก SO Number: {record_so} ใช่หรือไม่?\n"
                                   "(รายการจะถูกซ่อนจากประวัติ แต่ยังสามารถตรวจสอบได้โดยแอดมิน)", 
                                   parent=self, icon="warning"):
                conn = None
                try:
                    conn = self.app_container.get_connection()
                    with conn.cursor() as cursor:
                        # อัปเดต is_active=0 เพื่อซ่อน และเปลี่ยนสถานะเป็น Cancelled
                        cursor.execute("UPDATE commissions SET is_active = 0, status = 'Cancelled' WHERE id = %s", (int(record_id),))
                    conn.commit()
                    messagebox.showinfo("สำเร็จ", "ยกเลิกรายการเรียบร้อยแล้ว", parent=self)
                    # โหลดข้อมูลใหม่เพื่อรีเฟรชตาราง
                    self._populate_history_table()
                except Exception as e:
                    if conn: conn.rollback()
                    messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
                finally:
                    if conn: self.app_container.release_connection(conn)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถยกเลิกรายการได้: {e}", parent=self)

    def _export_history(self):
        """
        (เวอร์ชันอัปเกรด) Export ข้อมูลทั้งหมดในช่วงเวลาที่เลือก
        """
        # --- เรียกใช้หน้าต่างเลือกเวลา ---
        from export_utils import DateRangeDialog # Import เข้ามาเฉพาะกิจ
        dialog = DateRangeDialog(self)
        self.wait_window(dialog)

        start_date = dialog.start_date
        end_date = dialog.end_date

        if not start_date or not end_date:
            print("Export canceled by user.")
            return

        try:
            # --- สร้าง Query เพื่อดึงข้อมูลทั้งหมดในช่วงเวลาที่เลือก ---
            query = """
            SELECT * FROM commissions 
            WHERE sale_key = %s 
              AND is_active = 1
              AND timestamp::date BETWEEN %s AND %s
            ORDER BY timestamp DESC
            """
            
            # ใช้ pg_engine และ params ในการดึงข้อมูล
            df_to_export = pd.read_sql_query(query, self.pg_engine, params=(self.sale_key_filter, start_date, end_date))

            if df_to_export.empty:
                messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูล Commission ในช่วงเวลาที่เลือก", parent=self)
                return

            # --- ส่วนที่เหลือคือการจัดรูปแบบและบันทึกไฟล์ ---
            default_filename = f"commission_full_history_{self.sale_key_filter}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="บันทึกไฟล์ประวัติ Commission ทั้งหมด",
                initialfile=default_filename,
                parent=self
            )

            if not save_path:
                return

            # แปลงชื่อคอลัมน์ทั้งหมดเป็นภาษาไทยโดยใช้ HEADER_MAP ตัวหลัก
            header_map = self.app_container.HEADER_MAP
            df_to_export.rename(columns=lambda c: header_map.get(c, c), inplace=True)

            # บันทึกเป็นไฟล์ Excel
            df_to_export.to_excel(save_path, index=False)
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลทั้งหมดเรียบร้อยแล้วที่:\n{save_path}", parent=self)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=self)
            traceback.print_exc()

    def _show_loading(self):
        """แสดง Label 'กำลังโหลดข้อมูล...'"""
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.update_idletasks()

    def _hide_loading(self):
        """ซ่อน Label 'กำลังโหลดข้อมูล...'"""
        self.loading_label.place_forget()

    def _create_new_layout(self):
        """สร้าง UI Layout ใหม่ทั้งหมดที่มี Tabs และฟิลเตอร์"""
        # --- Top Frame (Filter & Pagination) ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")
        
        # --- START: เพิ่ม UI สำหรับฟิลเตอร์เดือน/ปี ---
        filter_frame = CTkFrame(top_frame, fg_color="transparent")
        filter_frame.pack(side="left")

        month_options = ["ทุกเดือน"] + self.thai_months
        CTkOptionMenu(filter_frame, variable=self.month_var, values=month_options).pack(side="left", padx=5)

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        CTkOptionMenu(filter_frame, variable=self.year_var, values=year_options).pack(side="left", padx=5)
        
        CTkButton(filter_frame, text="ค้นหา", command=self._populate_history_table, width=80).pack(side="left", padx=10)
        # --- END ---

        pagination_frame = CTkFrame(top_frame, fg_color="transparent")
        pagination_frame.pack(side="right")
        
        self.prev_button = CTkButton(pagination_frame, text="<<", command=self._prev_page, width=50, state="disabled")
        self.prev_button.pack(side="left", padx=5)
        self.page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.page_label.pack(side="left", padx=5)
        self.next_button = CTkButton(pagination_frame, text=">>", command=self._next_page, width=50, state="disabled")
        self.next_button.pack(side="left", padx=5)

        # --- Tab View ---
        self.tab_view = CTkTabview(self, command=self._on_tab_change)
        self.tab_view.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
        self.draft_tab = self.tab_view.add("ฉบับร่าง / ตีกลับ")
        self.submitted_tab = self.tab_view.add("รายการที่ส่งแล้ว")
        
        self.draft_tab.grid_columnconfigure(0, weight=1); self.draft_tab.grid_rowconfigure(0, weight=1)
        self.submitted_tab.grid_columnconfigure(0, weight=1); self.submitted_tab.grid_rowconfigure(0, weight=1)

        self.draft_frame = CTkFrame(self.draft_tab, fg_color="transparent"); self.draft_frame.grid(row=0, column=0, sticky="nsew")
        self.submitted_frame = CTkFrame(self.submitted_tab, fg_color="transparent"); self.submitted_frame.grid(row=0, column=0, sticky="nsew")
        
        self.button_frame = CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="e")
        
        self.cancel_button = CTkButton(self.button_frame, text="ยกเลิกรายการที่เลือก", command=self._cancel_selected_record, fg_color="#DC2626", hover_color="#B91C1C")
        self.cancel_button.pack(side="left", padx=10)
        
        self.export_button = CTkButton(self.button_frame, text="Export to Excel", command=self._export_history, fg_color=self.theme["primary"])
        self.export_button.pack(side="left")

        self.loading_label = CTkLabel(self, text="กำลังโหลดข้อมูล...", font=CTkFont(size=18, slant="italic"), text_color="gray50")


    def _on_tab_change(self):
        """Callback เมื่อมีการเปลี่ยน Tab"""
        selected_tab = self.tab_view.get()
        self.active_tab = "drafts" if selected_tab == "ฉบับร่าง / ตีกลับ" else "submitted"
        
        if self.active_tab == "drafts": self.cancel_button.pack(side="left", padx=10)
        else: self.cancel_button.pack_forget()
        
        self.current_page = 0
        self._populate_history_table()

    def _populate_history_table(self):
        """โหลดและแสดงข้อมูลตาม Tab และฟิลเตอร์ที่เลือก (เวอร์ชันแก้ไขไวยากรณ์ SQL)"""
        target_frame = self.draft_frame if self.active_tab == "drafts" else self.submitted_frame
        for widget in target_frame.winfo_children(): widget.destroy()
        self._show_loading()

        try:
            # --- START: แก้ไข Logic การสร้าง Query ทั้งหมด ---
            if self.active_tab == "drafts":
                status_condition = "c.status IN ('Original', 'Edited', 'Rejected by SM', 'Rejected by HR', 'Deferred by HR', 'Deferred by SM')"
            else:
                status_condition = "c.status NOT IN ('Original', 'Edited', 'Rejected by SM', 'Rejected by HR', 'Cancelled', 'Deferred by HR', 'Deferred by SM')"

            where_clauses = ["c.is_active = 1", status_condition]
            params = []

            if self.support_user_key_filter:
                where_clauses.append("c.support_user_key = %s")
                params.append(self.support_user_key_filter)
            elif self.sale_key_filter:
                where_clauses.append("c.sale_key = %s")
                params.append(self.sale_key_filter)

            selected_month_str = self.month_var.get()
            if selected_month_str != "ทุกเดือน":
                month_num = self.thai_month_map[selected_month_str]
                where_clauses.append("EXTRACT(MONTH FROM c.timestamp::timestamp) = %s")
                params.append(month_num)

            selected_year_str = self.year_var.get()
            if selected_year_str != "ทุกปี":
                year_num = int(selected_year_str)
                where_clauses.append("EXTRACT(YEAR FROM c.timestamp::timestamp) = %s")
                params.append(year_num)
            
            # 1. สร้างส่วน FROM และ JOIN ทั้งหมดก่อน
            query_body = """
                FROM commissions c
                LEFT JOIN sales_users ss ON c.support_user_key = ss.sale_key
                LEFT JOIN sales_users su_owner ON c.sale_key = su_owner.sale_key
            """
            
            # 2. สร้างส่วน WHERE แยกต่างหาก
            where_string = f"WHERE {' AND '.join(where_clauses)}"

            # 3. ประกอบร่าง Query สำหรับนับจำนวนแถว
            count_query = f"SELECT COUNT(c.id) {query_body} {where_string}"
            
            # --- END: สิ้นสุดการแก้ไข Logic ---

            count_df = pd.read_sql_query(count_query, self.pg_engine, params=tuple(params))
            self.total_rows = count_df.iloc[0, 0] if not count_df.empty else 0
            self.total_pages = (self.total_rows + self.rows_per_page - 1) // self.rows_per_page

            offset = self.current_page * self.rows_per_page
            
            # สร้าง params สำหรับ data_query โดยเพิ่ม limit และ offset
            data_params = params + [self.rows_per_page, offset]

            # 4. ประกอบร่าง Query สำหรับดึงข้อมูลมาแสดงผล
            data_query = f"""
                SELECT c.*,
                    ss.sale_name as support_user_name,
                    su_owner.sale_name as owner_name
                {query_body}
                {where_string}
                ORDER BY c.timestamp DESC
                LIMIT %s OFFSET %s
            """
            
            self.df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(data_params))

            self.df['customer_display'] = self.df.apply(
                lambda row: f"{row['customer_name']} (คีย์โดย: {row['support_user_name']})" if pd.notna(row['support_user_name']) else f"{row['customer_name']} (คีย์โดย: {row.get('owner_name', 'N/A')})",
                axis=1
            )

            self._hide_loading()

            if self.df.empty and self.current_page == 0:
                CTkLabel(target_frame, text="ไม่พบข้อมูล").pack(pady=20)
            else:
                self._create_styled_treeview(target_frame, self.df)

            self._update_pagination_controls()

        except Exception as e:
            self._hide_loading()
            traceback.print_exc()
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดประวัติได้: {e}", parent=self)

    def _create_styled_treeview(self, parent, df):
        """สร้าง Treeview และเติมข้อมูล (ปรับปรุงให้แสดงชื่อผู้คีย์)"""
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # +++ START: แก้ไขคอลัมน์ที่แสดงผล +++
        columns = ['id', 'timestamp', 'status', 'so_number', 'customer_display', 'sales_service_amount', 'shipping_cost', 'rejection_reason']
        display_columns = ['ID', 'เวลาบันทึก', 'สถานะ', 'SO Number', 'ชื่อลูกค้า (ผู้คีย์)', 'ยอดขาย/บริการ', 'ค่าขนส่ง', 'เหตุผลที่ถูกตีกลับ']
        # +++ END +++
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("History.Treeview.Heading", font=('Roboto', 11, 'bold'), relief="flat", background="#E5E7EB")
        style.configure("History.Treeview", rowheight=28, font=('Roboto', 11))
        style.map("History.Treeview", background=[('selected', self.theme.get("primary", "#3B82F6"))])
        
        self.tree = ttk.Treeview(parent, columns=columns, show='headings', style="History.Treeview")
        
        self.tree.tag_configure('Draft', background='#FEFCE8')
        self.tree.tag_configure('Rejected', background='#FEF2F2')
        self.tree.tag_configure('Submitted', background='#F0FDF4')
        self.tree.tag_configure('Default', background='white')

        for i, col_id in enumerate(columns):
            width = 100
            anchor = 'w'
            if col_id in ['id', 'status']: width = 80
            elif col_id in ['sales_service_amount', 'shipping_cost']: width = 120; anchor = 'e'
            # +++ START: แก้ไขความกว้างคอลัมน์ +++
            elif col_id == 'customer_display': width = 300
            elif col_id == 'so_number': width = 150
            # +++ END +++
            elif col_id == 'timestamp': width = 160
            elif col_id == 'rejection_reason': width = 250
            
            self.tree.heading(col_id, text=display_columns[i])
            self.tree.column(col_id, anchor=anchor, width=width)
        
        for index, row in df.iterrows():
            status = row['status']
            tag = 'Default'
            if 'Reject' in status or 'Defer' in status: tag = 'Rejected'
            elif status in ['Original', 'Edited']: tag = 'Draft'
            else: tag = 'Submitted'
            
            values = []
            for col_name in columns:
                value = row.get(col_name)
                if pd.notna(value):
                    if isinstance(value, (float, np.floating)):
                        values.append(f"{value:,.2f}")
                    elif isinstance(value, (datetime, pd.Timestamp)):
                        values.append(value.strftime('%Y-%m-%d %H:%M'))
                    else:
                        values.append(str(value))
                else:
                    values.append("")
            self.tree.insert("", "end", values=values, tags=(tag,), iid=str(row['id']))
        
        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(parent, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')

        if self.on_row_double_click_callback:
            self.tree.bind("<Double-1>", lambda event: self._on_tree_row_double_click(event, self.tree))

    def _update_pagination_controls(self):
        """อัปเดตสถานะของปุ่ม Pagination"""
        self.page_label.configure(text=f"Page {self.current_page + 1} / {max(1, self.total_pages)}")
        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._populate_history_table()

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._populate_history_table()

class SOPopupWindow(CTkToplevel):
    def __init__(self, master, app_container, sales_data, so_shared_vars, sale_theme, on_save_callback=None):
        super().__init__(master)
        self.master = master
        self.app_container = app_container # <<< เพิ่มบรรทัดนี้
        self.sales_data = sales_data
        self.so_shared_vars = so_shared_vars
        self.sale_theme = sale_theme
        self.on_save_callback = on_save_callback # <<< เพิ่มบรรทัดนี้
        self.popup_widgets = {}
        self.trace_ids_for_so_calc = []

        # เคลียร์และสร้าง StringVar ที่จำเป็นสำหรับ Pop-up นี้โดยเฉพาะ
        self.so_shared_vars['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['other_service_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        # <<< เพิ่มเติม: StringVar สำหรับค่าย้ายใน Pop-up >>>
        self.so_shared_vars['relocation_vat_calc_var'] = tk.StringVar(value="0.00")
        
        self.title(f"ข้อมูล Sales Order (SO: {sales_data.get('so_number', 'N/A')})")
        self.geometry("700x750") # ขยายความสูงเผื่อ
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.so_data_form_frame = CTkScrollableFrame(self, corner_radius=10, label_text="ข้อมูล Sales Order (แก้ไขได้)", label_fg_color=self.sale_theme["bg"], label_text_color=self.sale_theme["header"])
        self.so_data_form_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self._create_so_data_form_content(self.so_data_form_frame)
        self._so_bind_events()
        self.after(100, lambda: self._populate_so_form(self.sales_data))

        self.protocol("WM_DELETE_WINDOW", self._on_popup_close)
        self.transient(master)
        self.grab_set()

    def _on_popup_close(self):
        for var, trace_id in self.trace_ids_for_so_calc:
            try:
                var.trace_vdelete("write", trace_id)
            except tk.TclError as e:
                print(f"คำเตือน: ไม่สามารถยกเลิก trace '{trace_id}' ได้: {e}")
        self.trace_ids_for_so_calc = []
        self.destroy()

    def _create_so_section_frame(self, parent, title):
        frame = CTkFrame(parent, corner_radius=10, border_width=1, border_color=self.sale_theme['primary'])
        frame.pack(fill="x", pady=(10, 5), padx=5)
        frame.grid_columnconfigure(1, weight=1)
        CTkLabel(frame, text=title, font=CTkFont(size=18, weight="bold"), text_color=self.sale_theme["header"]).grid(row=0, column=0, columnspan=3, padx=15, pady=(10, 5), sticky="w")
        return frame
            
    def _add_form_row(self, parent, label_text, widget, key, row_index):
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=(15, 10), pady=4, sticky="w")
        widget.grid(row=row_index, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")
        self.popup_widgets[key] = widget
    
    def _add_item_row_with_vat(self, parent, label_text, entry_key, vat_option_key, vat_display_var_key, row_index):
        """
        (เวอร์ชันแก้ไข) ฟังก์ชัน Helper ที่จะเพิ่มป้ายแสดง VAT เข้าไปทางด้านขวา
        """
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(
            row=row_index, column=0, padx=(15, 10), pady=4, sticky="w"
        )
        
        item_frame = CTkFrame(parent, fg_color="transparent")
        item_frame.grid(row=row_index, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")
        item_frame.grid_columnconfigure(0, weight=1)

        amount_entry = NumericEntry(item_frame)
        amount_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.popup_widgets[entry_key] = amount_entry

        vat_frame = CTkFrame(item_frame, fg_color="transparent")
        vat_frame.pack(side="left")
        CTkRadioButton(vat_frame, text="VAT", variable=self.so_shared_vars[vat_option_key], value="VAT").pack(side="left")
        CTkRadioButton(vat_frame, text="CASH", variable=self.so_shared_vars[vat_option_key], value="CASH").pack(side="left", padx=5)

        # --- START: เพิ่มโค้ดส่วนนี้ ---
        # เพิ่ม Label สำหรับแสดงยอด VAT ของรายการนี้โดยเฉพาะ
        vat_label = CTkLabel(item_frame, textvariable=self.so_shared_vars[vat_display_var_key], font=CTkFont(size=12), text_color="gray50")
        vat_label.pack(side="left", padx=(10, 0))
        # --- END ---
            
    def _create_so_data_form_content(self, parent_frame):
        # Section 1: Sales Details
        f1 = self._create_so_section_frame(parent_frame, "รายละเอียดการขาย")
        self._add_form_row(f1, "วันที่เปิด SO:", DateSelector(f1, dropdown_style=self.master.dropdown_style), 'bill_date_selector', 1)
        self._add_form_row(f1, "ชื่อลูกค้า:", CTkEntry(f1), 'customer_name_entry', 2)
        self._add_form_row(f1, "รหัสลูกค้า:", CTkEntry(f1), 'customer_id_entry', 3)
        self._add_form_row(f1, "Credit Term:", CTkEntry(f1), 'credit_term_entry', 4)

        # Section 2: Sales and Services
        f2 = self._create_so_section_frame(parent_frame, "ยอดขายและบริการ")
        # --- แก้ไขการเรียกใช้ฟังก์ชัน 3 บรรทัดนี้ ---
        self._add_item_row_with_vat(f2, "ยอดขายสินค้า/บริการ:", 'sales_amount_entry', 'sales_service_vat_option', 'sales_vat_calc_var', 1)
        self._add_item_row_with_vat(f2, "ค่าบริการตัด/เจาะ:", 'cutting_drilling_fee_entry', 'cutting_drilling_fee_vat_option', 'cutting_drilling_vat_calc_var', 2)
        self._add_item_row_with_vat(f2, "ค่าบริการอื่นๆ:", 'other_service_fee_entry', 'other_service_fee_vat_option', 'other_service_vat_calc_var', 3)
        
        # Section 3: Shipping Cost
        f3 = self._create_so_section_frame(parent_frame, "ค่าจัดส่ง")
        # --- แก้ไขการเรียกใช้ฟังก์ชัน 1 บรรทัดนี้ ---
        self._add_item_row_with_vat(f3, "ค่าจัดส่ง:", 'shipping_cost_entry', 'shipping_vat_option_var', 'shipping_vat_calc_var', 1)
        self._add_form_row(f3, "วันที่จัดส่ง:", DateSelector(f3, dropdown_style=self.master.dropdown_style), 'delivery_date_selector', 2)

        # Section 4: Delivery Note
        f4 = self._create_so_section_frame(parent_frame, "Delivery Note")
        delivery_options = [
            "ซัพพลายเออร์จัดส่ง", "Aplus Logistic ส่งหน้างาน", "ลูกค้ารับเองที่ซัพ",
            "ลูกค้ารับเองที่คลัง 132", "ย้ายเข้าคลัง Aplus Logistic รอลูกค้ารับที่คลัง",
            "ย้ายเข้าคลัง Aplus Logistic รอ Aplus Logistic จัดส่ง",
            "ย้ายเข้าคลัง Lalamove รอลูกค้ารับที่คลัง 132", "ส่ง Lalamove ให้ลูกค้าหน้างาน",
            "Aplus Logistic+ฝากส่งขนส่ง", "Lalamove +ฝากส่งขนส่ง"
        ]
        self._add_form_row(f4, "การจัดส่ง:", CTkOptionMenu(f4, variable=self.so_shared_vars['delivery_type_var'], values=delivery_options, **self.master.dropdown_style), 'delivery_type_menu', 1)
        self._add_form_row(f4, "Location เข้ารับ:", CTkEntry(f4, placeholder_text="ใส่ อำเภอ, จังหวัด หรือ Google map link"), 'pickup_location_entry', 2)

        # --- เพิ่มส่วน "ค่าย้าย" พร้อม VAT ---
        self._add_item_row_with_vat(f4, "ค่าย้าย:", 'relocation_cost_entry', 'relocation_cost_vat_option', 'relocation_vat_calc_var', 3)

        self._add_form_row(f4, "วันที่ย้ายเข้าคลัง:", DateSelector(f4, dropdown_style=self.master.dropdown_style), 'date_to_wh_selector', 4)
        self._add_form_row(f4, "วันที่จัดส่งลูกค้า:", DateSelector(f4, dropdown_style=self.master.dropdown_style), 'date_to_customer_selector', 5)
        self._add_form_row(f4, "ทะเบียนเข้ารับ:", CTkEntry(f4), 'pickup_rego_entry', 6)

        # Section 5: Fees and Discounts
        f5 = self._create_so_section_frame(parent_frame, "ค่าธรรมเนียมและส่วนลด")
        # --- แก้ไขการเรียกใช้ฟังก์ชัน 1 บรรทัดนี้ ---
        self._add_item_row_with_vat(f5, "ค่าธรรมเนียมบัตร:", 'credit_card_fee_entry', 'credit_card_fee_vat_option_var', 'card_fee_vat_calc_var', 1)
        self._add_form_row(f5, "ค่าธรรมเนียมโอน:", NumericEntry(f5), 'transfer_fee_entry', 2)
        self._add_form_row(f5, "ภาษีหัก ณ ที่จ่าย:", NumericEntry(f5), 'wht_fee_entry', 3)
        self._add_form_row(f5, "ค่านายหน้า:", NumericEntry(f5), 'brokerage_fee_entry', 4)
        self._add_form_row(f5, "คูปอง:", NumericEntry(f5), 'coupon_value_entry', 5)
        self._add_form_row(f5, "ของแถม:", NumericEntry(f5), 'giveaway_value_entry', 6)

        # Section 6: Payment Details
        f6 = self._create_so_section_frame(parent_frame, "รายละเอียดการโอนชำระ")
        self._add_form_row(f6, "ยอดโอนชำระ 1:", NumericEntry(f6), 'payment1_amount_entry', 1)
        self._add_form_row(f6, "ยอดโอนชำระ 2:", NumericEntry(f6), 'payment2_amount_entry', 2)
        self._add_form_row(f6, "วันที่ชำระ:", DateSelector(f6, dropdown_style=self.master.dropdown_style), 'payment_date_selector', 3)
        
        # Section 7: SO Summary
        f7 = self._create_so_section_frame(parent_frame, "SO สรุปยอดรวม VAT")
        self._add_form_row(f7, "ยอดรวมที่ต้องชำระ:", CTkLabel(f7, textvariable=self.so_shared_vars['so_grand_total_var']), 'grand_total_display', 1)
        self._add_form_row(f7, "ตรวจสอบยอด SO vs โอน:", CTkLabel(f7, textvariable=self.so_shared_vars['so_vs_payment_result_var']), 'so_check_display', 2)
        self._add_form_row(f7, "ผลต่าง:", CTkLabel(f7, textvariable=self.so_shared_vars['difference_amount_var']), 'difference_display', 3)

        # Section 8: Cash Verification
        f8 = self._create_so_section_frame(parent_frame, "ตรวจสอบยอดชำระเงินสด")
        self._add_form_row(f8, "ยอดค่าสินค้าเงินสด:", NumericEntry(f8), 'cash_product_input_entry', 1)
        self._add_form_row(f8, "ยอดที่ต้องชำระเงินสด:", CTkLabel(f8, textvariable=self.so_shared_vars['cash_required_total_var']), 'cash_required_display', 2)
        self._add_form_row(f8, "ยอดชำระจริงเงินสด:", NumericEntry(f8), 'cash_actual_payment_entry', 3)
        self._add_form_row(f8, "ตรวจสอบยอดเงินสด:", CTkLabel(f8, textvariable=self.so_shared_vars['cash_verification_result_var']), 'cash_check_display', 4)
        
        # --- Single Save Button at the very end ---
        save_button = CTkButton(parent_frame, text="บันทึกข้อมูล SO", command=self._save_so_changes, fg_color="#16A34A", hover_color="#15803D", font=CTkFont(size=16, weight="bold"))
        save_button.pack(fill="x", padx=10, pady=20)
        

    def _so_bind_events(self):
        self.trace_ids_for_so_calc = []
        widgets_to_bind_keys = [
            "sales_amount_entry", "cutting_drilling_fee_entry", "other_service_fee_entry",
            "shipping_cost_entry", "credit_card_fee_entry", "transfer_fee_entry",
            "wht_fee_entry", "coupon_value_entry", "giveaway_value_entry",
            "brokerage_fee_entry", "payment1_amount_entry", "payment2_amount_entry",
            "cash_product_input_entry", "cash_actual_payment_entry",
            "relocation_cost_entry" # <<< เพิ่มเติม: relocation_cost_entry
        ]
        for key in widgets_to_bind_keys:
            if key in self.popup_widgets and isinstance(self.popup_widgets[key], (CTkEntry, NumericEntry)):
                self.popup_widgets[key].bind("<KeyRelease>", self._so_update_final_calculations)
            
        radio_vars_keys = [
            'sales_service_vat_option', 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option', 'shipping_vat_option_var',
            'credit_card_fee_vat_option_var',
            'relocation_cost_vat_option' # <<< เพิ่มเติม: relocation_cost_vat_option
        ]
        for key in radio_vars_keys:
            if key in self.so_shared_vars and isinstance(self.so_shared_vars[key], tk.StringVar):
                # ใช้ 'w' แทน 'write'
                trace_id = self.so_shared_vars[key].trace_add("write", self._so_update_final_calculations)
                self.trace_ids_for_so_calc.append((self.so_shared_vars[key], trace_id))

    def _so_update_final_calculations(self, *args):
        if not self.winfo_exists(): return

        w_vars = self.so_shared_vars
        w_widgets = self.popup_widgets

        def get_float_from_entry(entry_key):
            entry_widget = w_widgets.get(entry_key)
            if entry_widget and entry_widget.winfo_exists():
                try: return utils.convert_to_float(entry_widget.get())
                except (tk.TclError, ValueError): return 0.0
            return 0.0

        # --- 1. ดึงข้อมูลตัวเลขจากฟอร์มทั้งหมด ---
        sales = get_float_from_entry('sales_amount_entry')
        shipping = get_float_from_entry('shipping_cost_entry')
        card_fee = get_float_from_entry('credit_card_fee_entry')
        cutting_drilling = get_float_from_entry('cutting_drilling_fee_entry')
        other_service = get_float_from_entry('other_service_fee_entry')
        
        # ดึงค่าภาษีหัก ณ ที่จ่าย (WHT)
        wht = get_float_from_entry('wht_fee_entry')
        
        # --- 2. แยกรายการที่ต้องคิด VAT ---
        total_vatable_revenue = 0.0
        total_cashable_services_and_fees = 0.0
        items_to_process = [
            (get_float_from_entry('sales_amount_entry'), w_vars['sales_service_vat_option'].get(), w_vars['sales_vat_calc_var']),
            (get_float_from_entry('cutting_drilling_fee_entry'), w_vars['cutting_drilling_fee_vat_option'].get(), w_vars['cutting_drilling_vat_calc_var']),
            (get_float_from_entry('other_service_fee_entry'), w_vars['other_service_fee_vat_option'].get(), w_vars['other_service_vat_calc_var']),
            (get_float_from_entry('shipping_cost_entry'), w_vars['shipping_vat_option_var'].get(), w_vars['shipping_vat_calc_var']),
            (get_float_from_entry('credit_card_fee_entry'), w_vars['credit_card_fee_vat_option_var'].get(), w_vars['card_fee_vat_calc_var'])
        ]

        for amount, option, vat_display_var in items_to_process:
            item_vat = 0.0
            if option == "VAT":
                total_vatable_revenue += amount
                item_vat = amount * 0.07  # คำนวณ VAT ของรายการนี้
            else: 
                total_cashable_services_and_fees += amount
            
            # อัปเดตค่าใน StringVar ของป้ายแสดง VAT แต่ละรายการ
            vat_display_var.set(f"VAT: {item_vat:,.2f}")
                
        # --- 3. [นี่คือสูตรที่ถูกต้องและจะถูกเรียกใช้งานจริง] ---
        # ยอดที่ต้องชำระ = (ยอดรวมรายการ VAT ทั้งหมด * 1.07) - ยอดหักภาษี ณ ที่จ่าย (WHT)
        # ส่วนลดอื่นๆ (คูปอง, ของแถม) จะไม่ถูกนำมาคำนวณในยอดที่ลูกค้าต้องจ่าย
        final_grand_total = (total_vatable_revenue * 1.07) - wht
        w_vars['so_grand_total_var'].set(f"{final_grand_total:,.2f}")

        # --- 4. คำนวณส่วนต่างการชำระ ---
        payment1 = get_float_from_entry('payment1_amount_entry')
        payment2 = get_float_from_entry('payment2_amount_entry')
        so_vs_payment_diff = (payment1 + payment2) - final_grand_total
        w_vars['difference_amount_var'].set(f"{so_vs_payment_diff:,.2f}")

        # --- 5. อัปเดต UI แสดงผล ---
        def set_check_result(label_widget_key, var, diff_val, plus_text, minus_text):
            label_widget_ref = w_widgets.get(label_widget_key)
            if not (label_widget_ref and label_widget_ref.winfo_exists()): return
            color_map = {"-": ("gray85", "black"), "ok": ("#BBF7D0", "#15803D"), "bad": ("#FECACA", "#B91C1C")}
            if abs(diff_val) < 0.01: state, text = "ok", "ถูกต้อง"
            elif diff_val > 0: state, text = "ok", f"{plus_text} (+{abs(diff_val):,.2f})"
            else: state, text = "bad", f"{minus_text} ({abs(diff_val):,.2f})"
            var.set(text)
            label_widget_ref.configure(fg_color=color_map[state][0], text_color=color_map[state][1], text=text)

        set_check_result('so_check_display', w_vars.get('so_vs_payment_result_var'), so_vs_payment_diff, "ยอดโอนเกิน", "ยอดโอนขาด")

        # --- 6. คำนวณยอดเงินสด ---
        cash_product_val = get_float_from_entry('cash_product_input_entry')
        cash_required_total = cash_product_val + total_cashable_services_and_fees
        w_vars['cash_required_total_var'].set(f"{cash_required_total:,.2f}")
        
        actual_cash_payment = get_float_from_entry('cash_actual_payment_entry')
        cash_diff = actual_cash_payment - cash_required_total
        
        set_check_result('cash_check_display', w_vars.get('cash_verification_result_var'), cash_diff, "เงินสดเกิน", "เงินสดขาด")
        
    def _populate_so_form(self, data):
        if not self.winfo_exists(): return

        def set_val(widget_or_var, value):
            if not widget_or_var: return
            
            if not (hasattr(widget_or_var, 'winfo_exists') and widget_or_var.winfo_exists()) and not isinstance(widget_or_var, tk.StringVar):
                return

            if isinstance(widget_or_var, (CTkEntry, NumericEntry, AutoCompleteEntry)):
                state = widget_or_var.cget("state")
                widget_or_var.configure(state="normal")
                widget_or_var.delete(0, "end")
                if pd.notna(value):
                    widget_or_var.insert(0, f"{value:,.2f}" if isinstance(value, (float, int)) else str(value))
                widget_or_var.configure(state=state)
            elif isinstance(widget_or_var, DateSelector):
                dt = pd.to_datetime(value, errors='coerce')
                widget_or_var.set_date(dt.to_pydatetime() if pd.notna(dt) else None)
            elif isinstance(widget_or_var, tk.StringVar):
                widget_or_var.set(str(value) if pd.notna(value) and value else "")
            elif isinstance(widget_or_var, CTkLabel):
                widget_or_var.configure(text=f"{value:,.2f}" if isinstance(value, (float, int)) else str(value) if value is not None and value != "" else "")
            elif isinstance(widget_or_var, CTkOptionMenu):
                widget_or_var.set(str(value) if pd.notna(value) and value else widget_or_var.cget("values")[0])

        key_map = {
            'bill_date': 'bill_date_selector', 'customer_name': 'customer_name_entry', 'customer_id': 'customer_id_entry',
            'credit_term': 'credit_term_entry', 'sales_service_amount': 'sales_amount_entry', 'cutting_drilling_fee': 'cutting_drilling_fee_entry',
            'other_service_fee': 'other_service_fee_entry', 'shipping_cost': 'shipping_cost_entry', 'delivery_date': 'delivery_date_selector',
            'credit_card_fee': 'credit_card_fee_entry', 'transfer_fee': 'transfer_fee_entry', 'wht_3_percent': 'wht_fee_entry',
            'brokerage_fee': 'brokerage_fee_entry', 'coupons': 'coupon_value_entry', 'giveaways': 'giveaway_value_entry',
            'payment_date': 'payment_date_selector', 'cash_product_input': 'cash_product_input_entry', 'cash_actual_payment': 'cash_actual_payment_entry',
            'sales_service_vat_option': 'sales_service_vat_option', 'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option': 'other_service_fee_vat_option', 'shipping_vat_option': 'shipping_vat_option_var',
            'credit_card_fee_vat_option': 'credit_card_fee_vat_option_var', 'so_grand_total': 'so_grand_total_var',
            'so_vs_payment_result': 'so_vs_payment_result_var', 'difference_amount': 'difference_amount_var',
            'cash_required_total': 'cash_required_total_var', 'cash_verification_result': 'cash_verification_result_var',
            'delivery_type': 'delivery_type_var', 'pickup_location': 'pickup_location_entry',
            'relocation_cost': 'relocation_cost_entry', 'date_to_warehouse': 'date_to_wh_selector','payment_before_vat_entry': 'payment_before_vat', 
            'payment_no_vat_entry': 'payment_no_vat',
            'date_to_customer': 'date_to_customer_selector', 'pickup_registration': 'pickup_rego_entry'
        }
        
        for key, widget in self.popup_widgets.items():
            if isinstance(widget, (CTkEntry, NumericEntry, AutoCompleteEntry)): set_val(widget, "")
            elif isinstance(widget, DateSelector): set_val(widget, None)
            elif isinstance(widget, CTkLabel): widget.configure(text="")
        
        for key, var in self.so_shared_vars.items():
            if isinstance(var, tk.StringVar): var.set("")
        
        if data is not None:
            for db_key, w_key in key_map.items():
                widget_or_var = self.so_shared_vars.get(w_key) or self.popup_widgets.get(w_key)
                if widget_or_var:
                    set_val(widget_or_var, data.get(db_key))
            
            payment1_entry = self.popup_widgets.get('payment1_amount_entry')
            if payment1_entry: set_val(payment1_entry, data.get('total_payment_amount'))
        
        self.update_idletasks()
        self._so_update_final_calculations()

    def _save_so_changes(self):
        """(เวอร์ชันใหม่) รวบรวมข้อมูลและบันทึกลงฐานข้อมูลโดยตรงจาก Pop-up"""
        if self.sales_data is None: 
            messagebox.showerror("ข้อผิดพลาด", "ไม่มีข้อมูล SO ให้บันทึก", parent=self)
            return

        so_id = self.sales_data.get('id')
        updated_data = {}
        
        # --- Logic การรวบรวมข้อมูลจากฟอร์ม ---
        key_map = {
            'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id', 'credit_term_entry': 'credit_term',
            'pickup_location_entry': 'pickup_location', 'pickup_rego_entry': 'pickup_registration',
            'bill_date_selector': 'bill_date', 'delivery_date_selector': 'delivery_date', 'payment_date_selector': 'payment_date',
            'date_to_wh_selector': 'date_to_warehouse', 'date_to_customer_selector': 'date_to_customer',
            'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost',
            'relocation_cost_entry': 'relocation_cost', 'credit_card_fee_entry': 'credit_card_fee',
            'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons',
            'giveaway_value_entry': 'giveaways', 'cash_product_input_entry': 'cash_product_input',
            'cash_actual_payment_entry': 'cash_actual_payment'
        }

        for widget_key, data_key in key_map.items():
            value = None
            if widget_key in self.popup_widgets:
                widget = self.popup_widgets[widget_key]
                if widget and widget.winfo_exists():
                    if isinstance(widget, DateSelector): value = widget.get_date()
                    elif isinstance(widget, (NumericEntry, CTkEntry)):
                        value = widget.get()
                        if any(k in data_key for k in ['amount', 'cost', 'fee', 'wht', 'percent', 'coupons', 'giveaways']):
                            value = utils.convert_to_float(value)
            if value is not None: updated_data[data_key] = value

        shared_vars_map = {
            'delivery_type_var': 'delivery_type', 'sales_service_vat_option': 'sales_service_vat_option',
            'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option', 'other_service_fee_vat_option': 'other_service_fee_vat_option',
            'shipping_vat_option_var': 'shipping_vat_option', 'credit_card_fee_vat_option_var': 'credit_card_fee_vat_option'
        }
        for var_key, data_key in shared_vars_map.items():
            if var_key in self.so_shared_vars: updated_data[data_key] = self.so_shared_vars[var_key].get()

        p1 = utils.convert_to_float(self.popup_widgets.get('payment1_amount_entry').get())
        p2 = utils.convert_to_float(self.popup_widgets.get('payment2_amount_entry').get())
        updated_data['total_payment_amount'] = p1 + p2

        # --- Logic การบันทึกลงฐานข้อมูล ---
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                set_clauses = [f'"{k}" = %s' for k in updated_data.keys()]
                params = list(updated_data.values()) + [so_id]
                sql_update = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
                cursor.execute(sql_update, tuple(params))
            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกการแก้ไข SO เรียบร้อยแล้ว", parent=self)
            if self.on_save_callback:
                self.on_save_callback() # เรียก callback เพื่อ refresh หน้าจอหลัก
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)
