# history_windows.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pandas as pd
from datetime import datetime
import traceback
import psycopg2.errors
import psycopg2.extras
import numpy as np

# --- CustomTkinter Imports (กู้คืนส่วนนี้ที่หายไป) ---
from customtkinter import (
    CTkToplevel, CTkTextbox, CTkScrollableFrame, CTkLabel, CTkFont, 
    CTkFrame, CTkButton, CTkEntry, CTkRadioButton, CTkOptionMenu, CTkTabview,
    CTkCheckBox 
)

import utils
from utils import FormattedNumericEntry, RejectionReasonDialog
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry

# --- Import ฟังก์ชัน PDF จากไฟล์ที่เราเพิ่งแก้ ---
from po_document_generator import generate_transport_fee_pdf

# ========================================================================================
#  PRINT WRAPPER FUNCTION
# ========================================================================================

def print_transport_pdf_wrapper(app_container, po_id):
    conn = app_container.get_connection()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
            # 1. ดึงข้อมูล
            cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (po_id,))
            po_data = cursor.fetchone()
            
            if not po_data:
                messagebox.showerror("Error", "ไม่พบข้อมูล PO นี้")
                return

            po_dict = dict(po_data)
            
            # --- DEBUG LOG (จะขึ้นใน Terminal เมื่อกดปุ่ม Print) ---
            print("\n" + "="*40)
            print(f"DEBUG PRINT: {po_dict.get('po_number')}")
            print(f"Stock Driver: '{po_dict.get('shipping_to_stock_driver')}'")
            print(f"Stock Plate:  '{po_dict.get('shipping_to_stock_plate')}'")
            print("-" * 20)
            print(f"Site Driver:  '{po_dict.get('shipping_to_site_driver')}'")
            print(f"Site Plate:   '{po_dict.get('shipping_to_site_plate')}'")
            print("="*40 + "\n")

            header_data = {
                'so_number': po_dict.get('so_number', '-'),
                'customer_name': po_dict.get('customer_name', '-'),
                'sale_name': po_dict.get('user_key', '-') 
            }

            # เรียกฟังก์ชันจาก po_document_generator.py
            generate_transport_fee_pdf(header_data, [po_dict])

    except Exception as e:
        messagebox.showerror("Error", f"Error: {e}")
        print(traceback.format_exc())
    finally:
        app_container.release_connection(conn)

# ========================================================================================
#  EXISTING CLASSES START HERE
# ========================================================================================


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
            'unloading_status', 'special_request', # 🟢 เพิ่มใหม่
            'credit_card_fee', 'brokerage_fee', 'giveaway_vat', 'giveaway_no_vat', 'coupons', # 🟢 แก้ไข
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
        self._load_product_master_data()
        
        self.user_role = self.app_container.current_user_role
        
        # --- START: แก้ไขจุดนี้ ---
        # Initialize self.po_data ให้เป็น dict ว่างไว้ก่อน
        self.po_data = {}
        # --- END ---
        
        self.po_entries = {}
        self.item_entries = []
        self.deleted_item_ids = []
        self.payment_entries = []
        self.deleted_payment_ids = []
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0) 
        self.grid_columnconfigure(0, weight=1)

        self.scroll_frame = CTkScrollableFrame(self)
        self.scroll_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)
        
        # --- START: แก้ไขจุดนี้ ---
        # ย้ายการสร้างปุ่มไปไว้หลังจากโหลดข้อมูลเสร็จแล้ว
        # เราจะสร้างแค่ Frame เปล่าๆ ไว้รอก่อน
        self.button_frame = CTkFrame(self, fg_color=("gray85", "gray18"))
        self.button_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        # --- END ---

        self.after(50, self._load_and_display_data)
        self.transient(master)
        self.grab_set()
    
    def _load_product_master_data(self):
        try:
            query = "SELECT product_code, product_name, warehouse, last_unit_price, last_weight_per_unit FROM products ORDER BY product_code"
            df = pd.read_sql(query, self.app_container.pg_engine)
            
            self.product_completion_data = []
            self.product_data_map = {}
            
            MAX_NAME_LENGTH = 50
            for _, row in df.iterrows():
                name = row['product_name'] or ""
                display_name = name[:MAX_NAME_LENGTH] + '...' if len(name) > MAX_NAME_LENGTH else name
                display_text = f"{row['product_code']} - {display_name}"

                item_data = {
                    "name": name,
                    "code": row['product_code'],
                    "warehouse": row.get('warehouse', ''),
                    "display": display_text,
                    "last_price": row.get('last_unit_price'),
                    "last_weight": row.get('last_weight_per_unit')
                }
                self.product_completion_data.append(item_data)
                self.product_data_map[item_data['code']] = item_data
                
        except Exception as e: 
            print(f"Error loading product master data: {e}")
            self.product_completion_data = []
            self.product_data_map = {}

    # [🔥 เพิ่มใหม่ 2] ฟังก์ชันทำงานเมื่อเลือกสินค้า (Auto-fill ชื่อ/ราคา)
    def _on_product_code_selected(self, selection_dict, row_widgets):
        if not selection_dict: return
        
        # อัปเดตค่าในช่องรหัส (เผื่อเลือกจาก Dropdown)
        row_widgets['product_code'].delete(0, "end")
        row_widgets['product_code'].insert(0, selection_dict.get('code', ''))

        # ดึงข้อมูลสินค้ามาเติม
        code = selection_dict.get('code')
        product_data = self.product_data_map.get(code)
        
        if product_data:
            # เติมชื่อสินค้า
            row_widgets['product_name'].delete(0, "end")
            row_widgets['product_name'].insert(0, product_data.get('name', ''))
            
            # เติมคลัง
            row_widgets['warehouse'].delete(0, "end")
            row_widgets['warehouse'].insert(0, product_data.get('warehouse', ''))
            
            # เติมราคาล่าสุด (ถ้ามี)
            if product_data.get('last_price'):
                row_widgets['unit_price'].set(product_data.get('last_price'))
            
            # เติมน้ำหนัก (ถ้ามี)
            if product_data.get('last_weight'):
                row_widgets['total_weight'].set(product_data.get('last_weight'))

            # คำนวณยอดรวมใหม่
            self._recalculate_summary_totals()

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
                # 1. ดึงข้อมูล Header ของ PO
                # ใช้ SELECT * เพื่อให้ได้คอลัมน์ใหม่ (shipping_to_stock_driver ฯลฯ) มาด้วยอัตโนมัติ
                cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (self.purchase_id,))
                po_data = cursor.fetchone()
                
                if not po_data:
                    messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบ PO ID: {self.purchase_id}", parent=self)
                    self.destroy()
                    return

                self.po_data = dict(po_data)
                
                # 2. ดึงข้อมูล Supplier เพิ่มเติม (รหัส และ เครดิต)
                supplier_name = self.po_data.get('supplier_name')
                if supplier_name:
                    cursor.execute("SELECT supplier_code, credit_term FROM suppliers WHERE supplier_name = %s LIMIT 1", (supplier_name,))
                    supplier_info = cursor.fetchone()
                    if supplier_info:
                        self.po_data['supplier_code'] = supplier_info['supplier_code']
                        # ถ้าใน PO ไม่มีเครดิต (เป็นค่าว่าง) ให้ดึงจาก Supplier Master มาใช้
                        if not self.po_data.get('credit_term'):
                            self.po_data['credit_term'] = supplier_info['credit_term']
                    else:
                        self.po_data['supplier_code'] = ""
                        # self.po_data['credit_term'] คงค่าเดิมไว้

                # 3. ดึงรายการสินค้า (Items)
                cursor.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", (self.purchase_id,))
                items_data = cursor.fetchall()

                # 4. ดึงรายการชำระเงิน (Payments)
                cursor.execute("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", (self.purchase_id,))
                payments_data = cursor.fetchall()

            # แปลงข้อมูลเป็น List of Dict
            self.items_data = [dict(item) for item in items_data]
            self.payments_data = [dict(payment) for payment in payments_data]

            # --- [สำคัญ] สร้างปุ่มและหน้าจอ ---
            # ต้องสร้างปุ่มหลังจากโหลด self.po_data เสร็จแล้ว เพราะต้องเช็ค Status
            self._create_action_buttons()
            
            # สร้างหน้าจอแสดงผล (Shipping, Items, Cutting ฯลฯ)
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
        
        # --- [🔥 เพิ่ม] สร้างส่วนค่าบริการตัด/เจาะ ---
        self._create_cutting_section(self.scroll_frame, self.po_data)
        # ----------------------------------------

        self._create_payments_section(self.scroll_frame, self.payments_data)
        self._create_approval_info_section(self.scroll_frame, self.po_data)
        
        self._recalculate_summary_totals()

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

        # [🔥 แก้ไข] ใช้ AutoCompleteEntry แทน CTkEntry ปกติ
        # (ต้องมั่นใจว่าเรียก _load_product_master_data แล้วนะ)
        entry_code = AutoCompleteEntry(
            row_frame, 
            completion_list=getattr(self, 'product_completion_data', []), # กัน Error ถ้ายังไม่โหลด
            display_key='display',
            placeholder_text="รหัส"
        )
        entry_code.insert(0, item_data.get('product_code', ''))
        entry_code.grid(row=0, column=0, padx=5, sticky="ew")

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

        # รวม Widgets เพื่อเก็บ Reference
        widgets_dict = {
            'product_code': entry_code, 'product_name': entry_name, 
            'warehouse': entry_warehouse, 
            'total_weight': entry_weight, 'quantity': entry_qty, 
            'unit_price': entry_price, 'discount_value': entry_discount,
            'discount_type_var': discount_type_var,
            'total_price_label': label_total
        }

        # [🔥 เพิ่ม] ผูก Command เมื่อเลือกสินค้า
        entry_code.command = lambda sel, w=widgets_dict: self._on_product_code_selected(sel, w)

        self.item_entries.append({
            'id': item_data.get('id'), 'frame': row_frame, 
            'widgets': widgets_dict
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
        """สร้าง Section สำหรับข้อมูลการจัดส่ง (แก้ไข: เพิ่มคนขับ/ทะเบียน)"""
        shipping_frame = self._create_section(parent, "ข้อมูลการจัดส่ง (Shipping)")
        current_row = 1

        # --- 1. ค่าจัดส่งเข้าสต๊อก ---
        CTkLabel(shipping_frame, text="--- [1] ค่าจัดส่งเข้าสต๊อก ---", font=CTkFont(weight="bold"), text_color="#B91C1C").grid(row=current_row, column=0, columnspan=2, pady=(10,5), sticky="w", padx=10)
        current_row += 1

        self._create_editable_row(shipping_frame, current_row, "ค่าส่งเข้าสต๊อก:", data.get("shipping_to_stock_cost"), key="shipping_to_stock_cost", is_numeric=True)
        current_row += 1
        
        # [🔥 เพิ่ม] คนขับ + ทะเบียน (Stock)
        self._create_editable_row(shipping_frame, current_row, "คนขับ (Stock):", data.get("shipping_to_stock_driver"), key="shipping_to_stock_driver")
        current_row += 1
        self._create_editable_row(shipping_frame, current_row, "ทะเบียน (Stock):", data.get("shipping_to_stock_plate"), key="shipping_to_stock_plate")
        current_row += 1

        # วันที่
        CTkLabel(shipping_frame, text="วันที่ส่งเข้าสต๊อก:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        stock_date_selector = DateSelector(shipping_frame)
        stock_date_selector.set_date(data.get("shipping_to_stock_date"))
        stock_date_selector.grid(row=current_row, column=1, padx=10, pady=5, sticky="w")
        self.po_entries["shipping_to_stock_date"] = stock_date_selector
        current_row += 1

        # ตัวเลือกอื่นๆ
        shipper_options = ["ซัพพลายเออร์จัดส่ง", "Aplus Logistic", "Lalamove/Others"]
        self._create_dropdown_row(shipping_frame, current_row, "ผู้จัดส่ง", data.get("shipping_to_stock_shipper"), key="shipping_to_stock_shipper", options=shipper_options)
        current_row += 1
        
        self._create_dropdown_row(shipping_frame, current_row, "ประเภท VAT", data.get("shipping_to_stock_vat_type"), key="shipping_to_stock_vat_type", options=["VAT", "CASH"])
        current_row += 1
        
        wht_options = ["ไม่มีหัก", "1%", "3%"]
        self._create_dropdown_row(shipping_frame, current_row, "หัก ณ ที่จ่าย", data.get("shipping_to_stock_wht_type"), key="shipping_to_stock_wht_type", options=wht_options)
        current_row += 1
        
        self._create_editable_row(shipping_frame, current_row, "หมายเหตุ:", data.get("shipping_to_stock_notes"), key="shipping_to_stock_notes")
        current_row += 1

        # --- 2. ค่าจัดส่งเข้าไซต์ ---
        CTkLabel(shipping_frame, text="--- [2] ค่าจัดส่งเข้าไซต์ ---", font=CTkFont(weight="bold"), text_color="#1D4ED8").grid(row=current_row, column=0, columnspan=2, pady=(15,5), sticky="w", padx=10)
        current_row += 1

        self._create_editable_row(shipping_frame, current_row, "ค่าส่งเข้าไซต์:", data.get("shipping_to_site_cost"), key="shipping_to_site_cost", is_numeric=True)
        current_row += 1
        
        # [🔥 เพิ่ม] คนขับ + ทะเบียน (Site)
        self._create_editable_row(shipping_frame, current_row, "คนขับ (Site):", data.get("shipping_to_site_driver"), key="shipping_to_site_driver")
        current_row += 1
        self._create_editable_row(shipping_frame, current_row, "ทะเบียน (Site):", data.get("shipping_to_site_plate"), key="shipping_to_site_plate")
        current_row += 1

        # วันที่
        CTkLabel(shipping_frame, text="วันที่ส่งเข้าไซต์:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        site_date_selector = DateSelector(shipping_frame)
        site_date_selector.set_date(data.get("shipping_to_site_date"))
        site_date_selector.grid(row=current_row, column=1, padx=10, pady=5, sticky="w")
        self.po_entries["shipping_to_site_date"] = site_date_selector
        current_row += 1

        # ตัวเลือกอื่นๆ
        self._create_dropdown_row(shipping_frame, current_row, "ผู้จัดส่ง", data.get("shipping_to_site_shipper"), key="shipping_to_site_shipper", options=shipper_options)
        current_row += 1
        
        self._create_dropdown_row(shipping_frame, current_row, "ประเภท VAT", data.get("shipping_to_site_vat_type"), key="shipping_to_site_vat_type", options=["VAT", "CASH"])
        current_row += 1
        
        self._create_dropdown_row(shipping_frame, current_row, "หัก ณ ที่จ่าย", data.get("shipping_to_site_wht_type"), key="shipping_to_site_wht_type", options=wht_options)
        current_row += 1
        
        self._create_editable_row(shipping_frame, current_row, "หมายเหตุ:", data.get("shipping_to_site_notes"), key="shipping_to_site_notes")
        current_row += 1

        # ค่าย้าย (Relocation) - แยกออกมาต่างหาก
    
    def _create_cutting_section(self, parent, data):
        """สร้าง Section สำหรับค่าบริการตัด/เจาะ"""
        cutting_frame = self._create_section(parent, "ค่าบริการตัด/เจาะ (Cutting/Drilling)")
        current_row = 1

        # ค่าบริการ (Cost)
        self._create_editable_row(cutting_frame, current_row, "ค่าบริการตัด/เจาะ:", data.get("cutting_cost"), key="cutting_cost", is_numeric=True)
        current_row += 1

        # แสดง VAT 7% (คำนวณอัตโนมัติ)
        CTkLabel(cutting_frame, text="VAT 7%:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        cutting_vat_display = CTkEntry(cutting_frame, state="readonly", fg_color="gray85")
        cutting_vat_display.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries["cutting_vat_display"] = cutting_vat_display
        current_row += 1

        # ตัวเลือกประเภท VAT / หัก ณ ที่จ่าย
        self._create_dropdown_row(cutting_frame, current_row, "ประเภท VAT", data.get("cutting_vat_type"), key="cutting_vat_type", options=["VAT", "CASH"])
        current_row += 1

        wht_options = ["No", "1%", "3%"] # ตรงกับ Database Default 'No'
        self._create_dropdown_row(cutting_frame, current_row, "หัก ณ ที่จ่าย", data.get("cutting_wht_type"), key="cutting_wht_type", options=wht_options)
        current_row += 1

        # แสดงยอดหัก ณ ที่จ่าย (คำนวณอัตโนมัติ)
        CTkLabel(cutting_frame, text="ยอดหัก ณ ที่จ่าย:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
        cutting_wht_display = CTkEntry(cutting_frame, state="readonly", fg_color="gray85")
        cutting_wht_display.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries["cutting_wht_display"] = cutting_wht_display
        current_row += 1

        # หมายเหตุ
        self._create_editable_row(cutting_frame, current_row, "หมายเหตุ:", data.get("cutting_remark"), key="cutting_remark")
        
        # Note แจ้งเตือน
        CTkLabel(cutting_frame, text="(หมายเหตุ: ค่าตัด/เจาะ จะถูกรวมในต้นทุน แต่ไม่ถูกรวมในยอดที่ต้องจ่ายให้ซัพพลายเออร์)", 
                 font=CTkFont(size=11, slant="italic"), text_color="gray50").grid(row=current_row+1, column=0, columnspan=2, padx=10, pady=5, sticky="w")

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
            
        total_cost = 0.0 # ยอดรวมย่อยสินค้า
        total_weight = 0.0
        
        # 1. คำนวณยอดสินค้าจากตาราง (ส่วนนี้ทำงานได้ปกติ)
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
                pass

        # 2. ดึงค่าอื่นๆ
        try:
            # Helper function เพื่อดึงค่าอย่างปลอดภัย
            def get_val(key):
                entry = self.po_entries.get(key)
                return entry.get_value() if entry and entry.winfo_exists() else 0.0
            
            def get_opt(key, default="CASH"):
                obj = self.po_entries.get(key)
                if obj:
                    if isinstance(obj, tk.StringVar): return obj.get()
                    elif hasattr(obj, 'winfo_exists') and obj.winfo_exists(): return obj.get()
                return default

            bill_discount = get_val('bill_discount')
            
            # --- Shipping Info ---
            raw_shipping_stock = get_val('shipping_to_stock_cost')
            shipping_stock_vat_type = get_opt('shipping_to_stock_vat_type')
            shipper_stock = get_opt('shipping_to_stock_shipper') # 🔥 ดึงผู้จัดส่ง (Stock)

            raw_shipping_site = get_val('shipping_to_site_cost')
            shipping_site_vat_type = get_opt('shipping_to_site_vat_type')
            shipper_site = get_opt('shipping_to_site_shipper')   # 🔥 ดึงผู้จัดส่ง (Site)

            # --- Relocation Info ---
            raw_relocation_cost = get_val('relocation_cost') 
            # ค่าย้าย: ดูจาก Delivery Type หลักของ PO (เพราะหน้านี้ไม่มี Dropdown แยก)
            main_delivery_type = self.po_data.get('delivery_type', '')

            # --- Cutting Info ---
            cutting_cost = get_val('cutting_cost')
            cutting_vat_type = get_opt('cutting_vat_type')
            cutting_wht_type = get_opt('cutting_wht_type', 'No')

            wht_entry = self.po_entries.get('wht_3_percent')
            vat_entry = self.po_entries.get('vat_7_percent')
        
        except Exception as e:
            print(f"Recalc Error (Data Fetch): {e}")
            return 

        # =========================================================================
        # 🟢 [LOGIC ใหม่] กรองยอดที่จะจ่ายให้ซัพพลายเออร์เท่านั้น
        # =========================================================================
        
        # 1. ค่าส่งเข้า Stock: จ่ายซัพฯ เฉพาะเมื่อเลือก "ซัพพลายเออร์จัดส่ง"
        payable_shipping_stock = raw_shipping_stock if shipper_stock == 'ซัพพลายเออร์จัดส่ง' else 0.0
        
        # 2. ค่าส่งเข้า Site: จ่ายซัพฯ เฉพาะเมื่อเลือก "ซัพพลายเออร์จัดส่ง"
        payable_shipping_site = raw_shipping_site if shipper_site == 'ซัพพลายเออร์จัดส่ง' else 0.0
        
        # 3. ค่าย้าย: จ่ายซัพฯ เฉพาะเมื่อ Delivery Type หลักเป็น "ซัพพลายเออร์จัดส่ง"
        payable_relocation = raw_relocation_cost if main_delivery_type == 'ซัพพลายเออร์จัดส่ง' else 0.0

        # =========================================================================

        # 3. คำนวณต้นทุนรวม (Net PO Cost) *เพื่อแสดงผล*
        # (ตรงนี้ยังโชว์ยอดเต็มรวมค่ารถทุกอย่าง เพื่อให้รู้ต้นทุนจริงของโปรเจกต์)
        product_base_cost = total_cost - bill_discount
        
        # 🔥 [แก้ตรงนี้] ตัดค่ารถออกจากการคำนวณต้นทุนสินค้า
        net_po_cost_display = product_base_cost + cutting_cost 

        # อัปเดต Label ต้นทุน
        if hasattr(self, 'total_cost_label') and self.total_cost_label.winfo_exists():
            self.total_cost_label.configure(text=f"{net_po_cost_display:,.2f}")
        
        if hasattr(self, 'total_weight_label') and self.total_weight_label.winfo_exists():
            self.total_weight_label.configure(text=f"{total_weight:,.2f} kg")

        # 4. คำนวณ VAT/WHT (ใช้ยอดที่ Payable เท่านั้นมาคิดภาษีในบิลนี้)
        
        # (Shipping VAT Display - แสดงตามยอดจริงของช่องนั้นๆ)
        stock_vat_display = raw_shipping_stock * 0.07 if shipping_stock_vat_type == 'VAT' else 0.0
        site_vat_display = raw_shipping_site * 0.07 if shipping_site_vat_type == 'VAT' else 0.0
        
        if self.po_entries.get("shipping_to_stock_vat_display"): 
            utils.set_entry_text(self.po_entries["shipping_to_stock_vat_display"], f"{stock_vat_display:,.2f}")
        if self.po_entries.get("shipping_to_site_vat_display"): 
            utils.set_entry_text(self.po_entries["shipping_to_site_vat_display"], f"{site_vat_display:,.2f}")
        
        # (Cutting VAT/WHT)
        cutting_vat_amount = cutting_cost * 0.07 if cutting_vat_type == 'VAT' else 0.0
        if self.po_entries.get("cutting_vat_display"): 
            utils.set_entry_text(self.po_entries["cutting_vat_display"], f"{cutting_vat_amount:,.2f}")
        
        cutting_wht_rate = 0.01 if '1' in cutting_wht_type else 0.03 if '3' in cutting_wht_type else 0
        cutting_wht_amount = cutting_cost * cutting_wht_rate
        if self.po_entries.get("cutting_wht_display"): 
            utils.set_entry_text(self.po_entries["cutting_wht_display"], f"{cutting_wht_amount:,.2f}")

        # 5. คำนวณยอดที่ต้องจ่ายซัพพลายเออร์ (Grand Total) จริงๆ
        # สูตร: (สินค้า + ค่าบริการที่ซัพเก็บ) + VAT - WHT
        
        # ฐานภาษี (สินค้า)
        base_for_tax = product_base_cost 
        
        # บวกค่าขนส่งเข้าฐาน VAT (เฉพาะส่วนที่จ่ายซัพฯ และเป็น VAT)
        if shipping_stock_vat_type == 'VAT': base_for_tax += payable_shipping_stock
        if shipping_site_vat_type == 'VAT': base_for_tax += payable_shipping_site
        # (ถ้าค่าตัดมี VAT ก็รวมด้วย แต่มันแยกคิดข้างล่าง หรือรวมตรงนี้แล้วแต่ระบบเก่า)
        # ตามโค้ดเดิม ค่าตัดแยกคิดต่างหาก หรือรวม? 
        # ปกติค่าตัดรวมใน PO เดียวกัน ถ้าซัพฯ ทำ
        
        # คำนวณ VAT รวมของ PO นี้
        vat_amount_total = base_for_tax * 0.07 if hasattr(self, 'vat_checkbox') and self.vat_checkbox.get() == 1 else 0.0
        
        # รวม VAT ค่าตัด (ถ้ามี)
        vat_amount_total += cutting_vat_amount

        # คำนวณ WHT รวม
        wht_amount_products = base_for_tax * 0.03 if hasattr(self, 'wht_checkbox') and self.wht_checkbox.get() == 1 else 0.0
        
        # รวม WHT ค่าขนส่ง (เฉพาะที่จ่ายซัพ)
        stock_wht_rate = 0.01 if '1' in get_opt('shipping_to_stock_wht_type') else 0.03 if '3' in get_opt('shipping_to_stock_wht_type') else 0
        site_wht_rate = 0.01 if '1' in get_opt('shipping_to_site_wht_type') else 0.03 if '3' in get_opt('shipping_to_site_wht_type') else 0
        
        total_wht_deduction = wht_amount_products + (payable_shipping_stock * stock_wht_rate) + (payable_shipping_site * site_wht_rate) + cutting_wht_amount

        if wht_entry and wht_entry.winfo_exists(): wht_entry.set(total_wht_deduction)
        if vat_entry and vat_entry.winfo_exists(): vat_entry.set(vat_amount_total)

        # ส่วนที่ไม่มี VAT (และต้องจ่ายซัพฯ)
        non_vat_payable = 0.0
        if shipping_stock_vat_type == 'CASH': non_vat_payable += payable_shipping_stock
        if shipping_site_vat_type == 'CASH': non_vat_payable += payable_shipping_site
        non_vat_payable += payable_relocation # ค่าย้ายที่จ่ายซัพ
        
        # ยอดสุทธิที่ต้องจ่าย (Grand Total)
        # = (สินค้า + ค่ารถที่ซัพเก็บ) + VAT - WHT + ค่าตัด + ค่ารถ(Cash)
        # (ต้องระวังอย่าบวกซ้ำ)
        
        grand_total = (base_for_tax + non_vat_payable + cutting_cost) + vat_amount_total - total_wht_deduction
        
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
        """สร้างหัวตารางสำหรับ Payment (แก้ไข: เพิ่มคอลัมน์ประเภทบัญชี)"""
        # [🔥 แก้ไข] เพิ่ม 'ประเภทบัญชี' ใน list
        headers = ["ประเภทการชำระ", "ยอดเงิน", "วันที่ชำระ", "ธนาคาร", "เลขที่บัญชี", "ประเภทบัญชี"]
        # [🔥 แก้ไข] ปรับน้ำหนักคอลัมน์ให้สมดุล
        col_weights = [2, 2, 2, 2, 3, 2]
        
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
        """เพิ่มแถวการชำระเงิน (แก้ไข: เพิ่ม Dropdown ประเภทบัญชี)"""
        if payment_data is None: 
            payment_data = {}

        row_frame = CTkFrame(self.payments_content_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=2)
        
        # [🔥 แก้ไข] กำหนด Grid ใหม่ (เพิ่มคอลัมน์ที่ 5)
        row_frame.grid_columnconfigure(0, weight=2)  # ประเภท
        row_frame.grid_columnconfigure(1, weight=2)  # ยอดเงิน
        row_frame.grid_columnconfigure(2, weight=2)  # วันที่
        row_frame.grid_columnconfigure(3, weight=2)  # ธนาคาร
        row_frame.grid_columnconfigure(4, weight=3)  # เลขบัญชี
        row_frame.grid_columnconfigure(5, weight=2)  # [ใหม่] ประเภทบัญชี
        row_frame.grid_columnconfigure(6, weight=0)  # ปุ่มลบ

        # 0. ประเภทการชำระ
        payment_types = ["Payment 1", "Payment 2", "Full Payment", "CN Refund"]
        type_var = tk.StringVar(value=payment_data.get('payment_type', payment_types[0]))
        type_menu = CTkOptionMenu(row_frame, variable=type_var, values=payment_types, width=120)
        type_menu.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        # 1. ยอดเงิน
        amount_entry = FormattedNumericEntry(row_frame)
        amount_entry.set(payment_data.get('amount', 0.0))
        amount_entry.grid(row=0, column=1, padx=5, sticky="ew")
        
        # 2. วันที่ชำระ
        date_selector = DateSelector(row_frame)
        date_selector.set_date(payment_data.get('payment_date'))
        date_selector.grid(row=0, column=2, padx=5, sticky="ew")
        
        # 3. ธนาคาร
        bank_list = ["ระบุเอง", "BBL", "KBANK", "KTB", "SCB", "TTB", "BAY", "GSB", "BAAC", "UOB", "CIMB"]
        bank_var = tk.StringVar(value=payment_data.get('bank_name', bank_list[0]))
        bank_menu = CTkOptionMenu(row_frame, 
                                variable=bank_var, 
                                values=bank_list,
                                width=100,
                                command=lambda bank=bank_var.get(), acc_entry=None: self._on_bank_selected(bank, acc_entry))
        bank_menu.grid(row=0, column=3, padx=5, sticky="ew")
        
        # 4. เลขที่บัญชี
        account_entry = CTkEntry(row_frame)
        account_number = payment_data.get('bank_account_number')
        if account_number is not None and pd.notna(account_number):
            account_entry.insert(0, str(account_number))
        account_entry.grid(row=0, column=4, padx=5, sticky="ew")
        
        # [🔥 เพิ่มใหม่] 5. ประเภทบัญชี (Dropdown)
        acc_type_val = payment_data.get('bank_account_type', 'ออมทรัพย์')
        # กันเหนียวเผื่อค่าเป็น None
        if not acc_type_val: acc_type_val = 'ออมทรัพย์'
            
        acc_type_var = tk.StringVar(value=acc_type_val)
        acc_type_menu = CTkOptionMenu(row_frame, variable=acc_type_var, values=["ออมทรัพย์", "กระแสรายวัน"], width=100)
        acc_type_menu.grid(row=0, column=5, padx=5, sticky="ew")

        # แก้ไข trace สำหรับ bank selection (เพื่อให้ auto-fill เลขบัญชีทำงาน)
        bank_var.trace_add("write", lambda *args, bv=bank_var, ae=account_entry: self._on_bank_selected(bv.get(), ae))
        
        # 6. ปุ่มลบ (ขยับไป column 6)
        delete_button = CTkButton(row_frame, 
                                text="ลบ", 
                                width=50, 
                                height=32,
                                fg_color="#DC2626", 
                                hover_color="#B91C1C",
                                command=lambda r=row_frame, p_id=payment_data.get('id'): self._remove_payment_row(r, p_id))
        delete_button.grid(row=0, column=6, padx=(5, 0), sticky="")

        # เก็บ reference ของ widgets ลง list
        self.payment_entries.append({
            'id': payment_data.get('id'),
            'frame': row_frame,
            'widgets': {
                'type_var': type_var,
                'amount_entry': amount_entry,
                'date_selector': date_selector,
                'bank_var': bank_var,
                'account_entry': account_entry,
                'acc_type_var': acc_type_var  # [🔥 สำคัญ] เก็บตัวแปรนี้ไว้ใช้ตอนบันทึก
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
        """สร้างปุ่ม Action แบบเรียบง่าย (ตัดปุ่มพิเศษออก) และเปิดสิทธิ์ให้ HR แก้ไขได้"""
        # ล้าง Frame เดิมก่อนสร้างใหม่
        for widget in self.button_frame.winfo_children():
            widget.destroy()

        po_status = self.po_data.get('status')
        current_user = self.app_container.current_user_key
        po_owner = self.po_data.get('user_key')
        
        # --- ตรวจสอบสิทธิ์ ---
        is_owner = (current_user == po_owner)
        is_manager_or_director = self.user_role in ['Purchasing Manager', 'Director']
        is_hr = (self.user_role == 'HR')

        # ใครบ้างที่มีสิทธิ์ "บันทึก" ได้เสมอ (แม้ PO จะ Approved แล้ว)
        can_edit_always = is_owner or is_manager_or_director or is_hr

        # กำหนดค่าความสูงและระยะห่าง
        button_height = 40
        vertical_padding = 10

        # ==============================================================================
        # ส่วนปุ่ม Action หลัก (Approve/Reject/Save/Close)
        # ==============================================================================
        
        # กรณี A: Manager/Director เปิด PO Approved (เห็นปุ่ม Revert)
        if po_status == 'Approved' and is_manager_or_director:
            self.button_frame.grid_columnconfigure((0, 1, 2), weight=1)

            revert_button = CTkButton(self.button_frame, text="ตีกลับ (Revert)", command=self._revert_to_draft, fg_color="#F97316", hover_color="#EA580C", height=button_height)
            revert_button.grid(row=0, column=0, padx=5, pady=vertical_padding, sticky="ew")

            save_button = CTkButton(self.button_frame, text="บันทึกการแก้ไข", command=self._save_changes, fg_color="#3B82F6", hover_color="#2563EB", height=button_height)
            save_button.grid(row=0, column=1, padx=5, pady=vertical_padding, sticky="ew")

            close_button = CTkButton(self.button_frame, text="ปิด", command=self.destroy, fg_color="gray", height=button_height)
            close_button.grid(row=0, column=2, padx=5, pady=vertical_padding, sticky="ew")

        # กรณี B: PO รออนุมัติ (เห็นปุ่ม Approve/Reject) - เฉพาะคนที่มีสิทธิ์อนุมัติ
        elif po_status == 'Pending Approval' and (is_manager_or_director or is_hr):
            self.button_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
            
            approve_button = CTkButton(self.button_frame, text="อนุมัติ", command=self._approve_po, fg_color="#16A34A", hover_color="#15803D", height=button_height)
            approve_button.grid(row=0, column=0, padx=5, pady=vertical_padding, sticky="ew")

            reject_button = CTkButton(self.button_frame, text="ปฏิเสธ", command=self._reject_po, fg_color="#DC2626", hover_color="#B91C1C", height=button_height)
            reject_button.grid(row=0, column=1, padx=5, pady=vertical_padding, sticky="ew")

            save_button = CTkButton(self.button_frame, text="บันทึก", command=self._save_changes, fg_color="#3B82F6", hover_color="#2563EB", height=button_height)
            save_button.grid(row=0, column=2, padx=5, pady=vertical_padding, sticky="ew")

            close_button = CTkButton(self.button_frame, text="ปิด", command=self.destroy, fg_color="gray", height=button_height)
            close_button.grid(row=0, column=3, padx=5, pady=vertical_padding, sticky="ew")
        
        # กรณี C: ทั่วไป (HR หรือ Owner ดู PO ปกติ หรือ PO Approved แล้ว)
        else:
            self.button_frame.grid_columnconfigure((0, 1), weight=1)
            
            # เช็คว่าจะให้ปุ่มบันทึกกดได้หรือไม่?
            if can_edit_always:
                save_text = "บันทึกการแก้ไข"
                save_state = "normal"
                save_color = "#3B82F6" # สีฟ้า
                save_hover = "#2563EB"
            else:
                save_text = "บันทึก (Disabled)"
                save_state = "disabled"
                save_color = "gray"
                save_hover = "gray"

            save_button = CTkButton(
                self.button_frame, 
                text=save_text, 
                state=save_state,
                fg_color=save_color,
                hover_color=save_hover,
                command=self._save_changes, 
                height=button_height
            )
            save_button.grid(row=0, column=0, padx=(0,5), pady=vertical_padding, sticky="ew")
            
            close_button = CTkButton(self.button_frame, text="ปิด", command=self.destroy, fg_color="gray", height=button_height)
            close_button.grid(row=0, column=1, padx=(5,0), pady=vertical_padding, sticky="ew")

    def _open_transport_edit_dialog(self):
        """เปิดหน้าต่างแก้ไขค่าขนส่ง"""
        try:
            TransportEditDialog(
                master=self,
                app_container=self.app_container,
                po_id=self.purchase_id,
                on_save_callback=self._load_and_display_data # รีโหลดหน้าจอหลัก
            )
        except NameError:
            messagebox.showerror("Error", "ไม่พบคลาส TransportEditDialog (โปรดตรวจสอบว่าได้วางโค้ด Class ไว้ท้ายไฟล์แล้ว)")

    def _open_transport_log_viewer(self):
        """เปิดหน้าต่างดู Log ค่าขนส่ง"""
        try:
            TransportLogViewer(self, self.app_container)
        except NameError:
             messagebox.showerror("Error", "ไม่พบคลาส TransportLogViewer")

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

    def _return_so_to_queue(self):
        """
        ยกเลิกการเชื่อมโยง PO ปัจจุบัน และคืนสถานะ SO กลับไปที่คิวงานของฝ่ายจัดซื้อ
        """
        # ดึงข้อมูล SO ที่เชื่อมโยงอยู่
        so_number_to_return = self.po_data.get('so_number')

        if not so_number_to_return:
            messagebox.showwarning("ไม่พบข้อมูล", "PO ใบนี้ไม่ได้เชื่อมโยงกับ SO ใดๆ", parent=self)
            return

        # ถามเพื่อยืนยันการทำงาน
        msg = (f"คุณต้องการยกเลิกการเชื่อมโยง PO นี้\n"
               f"และส่ง SO: {so_number_to_return} กลับไปที่คิวงานใช่หรือไม่?")
        
        if not messagebox.askyesno("ยืนยันการคืน SO", msg, icon="warning", parent=self):
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 1. อัปเดต PO ใบปัจจุบันให้ไม่ผูกกับ SO ใดๆ
                cursor.execute(
                    "UPDATE purchase_orders SET so_number = NULL WHERE id = %s",
                    (self.purchase_id,)
                )

                # 2. คืนสถานะของ SO กลับไปที่ 'Approved by SM' (หรือสถานะอื่นที่ถูกต้องสำหรับคิวงาน PU)
                cursor.execute(
                    "UPDATE commissions SET status = 'Approved by SM' WHERE so_number = %s",
                    (so_number_to_return,)
                )
            
            # ยืนยันการเปลี่ยนแปลงทั้งหมดในฐานข้อมูล
            conn.commit()
            messagebox.showinfo("สำเร็จ", "คืน SO กลับสู่คิวงานเรียบร้อยแล้ว", parent=self)
            
            # เรียก Callback เพื่อ Refresh หน้าจอหลัก (ถ้ามี)
            if self.on_save_callback:
                self.on_save_callback()
            
            # ปิดหน้าต่างปัจจุบัน
            self.destroy()

        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn:
                self.app_container.release_connection(conn)

    def _revert_to_draft(self):
        """สำหรับ Manager: ตีกลับ PO ที่ 'Approved' แล้วกลับไปเป็น 'Draft'"""
        dialog = RejectionReasonDialog(self)
        self.wait_window(dialog)
        reason = getattr(dialog, '_reason_string', None)
        if reason is None:
            return

        po_number = self.po_data.get('po_number', 'N/A')
        po_creator_key = self.po_data.get('user_key')

        msg = (f"คุณต้องการตีกลับ PO: {po_number} กลับไปเป็นฉบับร่างใช่หรือไม่?\n\n"
               f"PO จะถูกส่งกลับไปให้ {po_creator_key} เพื่อแก้ไข")
        if not messagebox.askyesno("ยืนยันการตีกลับ", msg, icon="warning", parent=self):
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    UPDATE purchase_orders 
                    SET 
                        status = 'Draft', 
                        approval_status = 'Draft',
                        approver_manager1_key = NULL, approval_date_manager1 = NULL,
                        approver_manager2_key = NULL, approval_date_manager2 = NULL,
                        approver_director_key = NULL, approval_date_director = NULL,
                        rejection_reason = %s,
                        last_modified_by = %s
                    WHERE id = %s
                """, (f"Reverted by {self.user_role}: {reason}", self.app_container.current_user_key, self.purchase_id))

                so_number = self.po_data.get('so_number')
                if so_number:
                    # คืนสถานะ SO กลับไปเป็น 'PO In Progress' เพื่อให้ PU รู้ว่าต้องจัดการต่อ
                    cursor.execute("UPDATE commissions SET status = 'PO In Progress' WHERE so_number = %s AND is_active = 1", (so_number,))
                
                if po_creator_key:
                    # สร้าง Notification แจ้งเตือนคนสร้าง PO
                    notif_msg = f"PO: {po_number} ที่อนุมัติแล้ว ถูกตีกลับโดย Manager เพื่อให้แก้ไข\nเหตุผล: {reason}"
                    cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)",
                                   (po_creator_key, notif_msg, self.purchase_id))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "ตีกลับ PO เป็นฉบับร่างเรียบร้อยแล้ว", parent=self)
            if self.on_save_callback:
                self.on_save_callback() # Refresh หน้าหลัก
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _save_changes(self):
        """บันทึกการแก้ไขข้อมูล PO (Header, Items, Payments) แบบสมบูรณ์"""
        
        # คำนวณยอดเงินล่าสุดก่อนบันทึก
        self._recalculate_summary_totals()
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ========================================================================================
                # PART 1: HEADER (ข้อมูลหลัก PO)
                # ========================================================================================
                new_supplier = self.po_entries['supplier_name'].get().strip()
                new_po_number = self.po_entries['po_number'].get().strip()
                
                total_cost = utils.convert_to_float(self.total_cost_label.cget("text"))
                grand_total = utils.convert_to_float(self.grand_total_label.cget("text"))

                cut_cost = 0.0
                if 'cutting_cost' in self.po_entries:
                    cut_cost = self.po_entries['cutting_cost'].get_value()

                relocation_val = 0.0
                if 'relocation_cost' in self.po_entries:
                    relocation_val = self.po_entries['relocation_cost'].get_value()

                # Helper ดึงค่า String จาก Entry แบบปลอดภัย
                def get_str(key):
                    entry = self.po_entries.get(key)
                    if entry:
                        return str(entry.get()).strip()
                    return ""

                # Helper ดึงค่าวันที่
                def get_date(key):
                    entry = self.po_entries.get(key)
                    if entry and hasattr(entry, 'get_date'):
                        return entry.get_date()
                    return None

                # Helper ดึงค่าตัวเลข
                def get_num(key):
                    entry = self.po_entries.get(key)
                    if not entry: return 0.0
                    
                    # 1. ถ้า widget มีเมธอด get_value() (เช่น NumericEntry) ให้ใช้เลย
                    if hasattr(entry, 'get_value'):
                        try:
                            return float(entry.get_value())
                        except:
                            return 0.0
                            
                    # 2. ถ้าเป็น CTkEntry หรือ StringVar
                    try:
                        val_str = str(entry.get()).replace(',', '').strip()
                        return float(val_str) if val_str else 0.0
                    except:
                        return 0.0

                cursor.execute("""
                    UPDATE purchase_orders SET 
                        po_number = %s, supplier_name = %s, 
                        credit_term = %s, po_mode = %s,
                        
                        shipping_to_stock_cost = %s, shipping_to_stock_date = %s, 
                        shipping_to_site_cost = %s, shipping_to_site_date = %s, 
                        relocation_cost = %s, total_cost = %s, grand_total = %s,
                        
                        shipping_to_stock_vat_type = %s, shipping_to_stock_shipper = %s,
                        shipping_to_stock_wht_type = %s, shipping_to_stock_notes = %s,
                        
                        shipping_to_site_vat_type = %s, shipping_to_site_shipper = %s,
                        shipping_to_site_wht_type = %s, shipping_to_site_notes = %s,
                        
                        shipping_to_stock_driver = %s, shipping_to_stock_plate = %s,
                        shipping_to_site_driver = %s, shipping_to_site_plate = %s,

                        cutting_cost = %s, cutting_vat_type = %s, cutting_wht_type = %s, cutting_remark = %s,

                        wht_3_percent_checked = %s, vat_7_percent_checked = %s,
                        bill_discount = %s
                    WHERE id = %s
                """, (
                    new_po_number, new_supplier,
                    get_str('credit_term'), get_str('po_mode'),
                    
                    get_num('shipping_to_stock_cost'),
                    get_date('shipping_to_stock_date'),
                    get_num('shipping_to_site_cost'),
                    get_date('shipping_to_site_date'),
                    
                    relocation_val,
                    total_cost, grand_total,
                    
                    get_str('shipping_to_stock_vat_type'),
                    get_str('shipping_to_stock_shipper'),
                    get_str('shipping_to_stock_wht_type'),
                    get_str('shipping_to_stock_notes'),
                    
                    get_str('shipping_to_site_vat_type'),
                    get_str('shipping_to_site_shipper'),
                    get_str('shipping_to_site_wht_type'),
                    get_str('shipping_to_site_notes'),

                    # --- [จุดสำคัญ] ต้องดึงให้ตรงกับ Key ที่สร้างไว้ใน _create_shipping_section ---
                    get_str('shipping_to_stock_driver'), 
                    get_str('shipping_to_stock_plate'),
                    get_str('shipping_to_site_driver'), 
                    get_str('shipping_to_site_plate'),
                    # ------------------------------------------------------------------------

                    cut_cost, 
                    get_str('cutting_vat_type'),
                    get_str('cutting_wht_type'),
                    get_str('cutting_remark'),

                    bool(self.wht_checkbox.get()),
                    bool(self.vat_checkbox.get()),
                    get_num('bill_discount'),
                    self.purchase_id
                ))

                # ... (ส่วน PART 2: ITEMS และ PART 3: PAYMENTS ปล่อยไว้เหมือนเดิมได้เลยครับ ไม่ต้องแก้) ...
                # ... แต่เพื่อให้ชัวร์ คุณก๊อปปี้ส่วนล่างของฟังก์ชันเดิมมาต่อท้ายตรงนี้ได้เลย ...
                
                # ========================================================================================
                # PART 2: ITEMS (สินค้า)
                # ========================================================================================
                
                if self.deleted_item_ids:
                    cursor.execute("DELETE FROM purchase_order_items WHERE id IN %s", (tuple(self.deleted_item_ids),))

                for item_row in self.item_entries:
                    widgets = item_row['widgets']
                    
                    qty = widgets['quantity'].get_value()
                    price = widgets['unit_price'].get_value()
                    discount = widgets['discount_value'].get_value()
                    discount_type = widgets['discount_type_var'].get()
                    
                    line_total = qty * price
                    discount_amount = (line_total * (discount / 100.0)) if discount_type == '%' else discount
                    item_total = line_total - discount_amount

                    item_id = item_row.get('id')

                    if item_id: 
                        cursor.execute("""
                            UPDATE purchase_order_items 
                            SET product_code = %s, product_name = %s, warehouse = %s,
                                total_weight = %s, quantity = %s, unit_price = %s, 
                                discount_value = %s, discount_type = %s, total_price = %s
                            WHERE id = %s
                        """, (
                            widgets['product_code'].get(), widgets['product_name'].get(), widgets['warehouse'].get(),
                            widgets['total_weight'].get_value(), qty, price,
                            discount, discount_type, item_total,
                            item_id
                        ))
                    else:
                        cursor.execute("""
                            INSERT INTO purchase_order_items 
                            (purchase_order_id, product_code, product_name, warehouse, total_weight, quantity, unit_price, discount_value, discount_type, total_price)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (
                            self.purchase_id,
                            widgets['product_code'].get(), widgets['product_name'].get(), widgets['warehouse'].get(),
                            widgets['total_weight'].get_value(), qty, price,
                            discount, discount_type, item_total
                        ))

                # ========================================================================================
                # PART 3: PAYMENTS (การชำระเงิน)
                # ========================================================================================
                
                if self.deleted_payment_ids:
                    cursor.execute("DELETE FROM purchase_order_payments WHERE id IN %s", (tuple(self.deleted_payment_ids),))

                for pay_row in self.payment_entries:
                    widgets = pay_row['widgets']
                    
                    p_type = widgets['type_var'].get()
                    amount = widgets['amount_entry'].get_value()
                    p_date = widgets['date_selector'].get_date()
                    bank = widgets['bank_var'].get()
                    acc_num = widgets['account_entry'].get()
                    acc_type = widgets['acc_type_var'].get()
                    
                    pay_id = pay_row.get('id')

                    if pay_id:
                        cursor.execute("""
                            UPDATE purchase_order_payments
                            SET payment_type = %s, amount = %s, payment_date = %s,
                                bank_name = %s, bank_account_number = %s, bank_account_type = %s
                            WHERE id = %s
                        """, (p_type, amount, p_date, bank, acc_num, acc_type, pay_id))
                    else:
                        cursor.execute("""
                            INSERT INTO purchase_order_payments
                            (purchase_order_id, payment_type, amount, payment_date, bank_name, bank_account_number, bank_account_type)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """, (self.purchase_id, p_type, amount, p_date, bank, acc_num, acc_type))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูล PO เรียบร้อยแล้ว", parent=self)
            
            if self.on_save_callback: 
                self.on_save_callback()
            
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

class PurchaseHistoryWindow(CTkToplevel):

    def _debounce_search(self, event=None):
        """ยกเลิกการค้นหาเก่าและตั้งเวลาใหม่ทุกครั้งที่พิมพ์"""
        if self._debounce_job:
            self.after_cancel(self._debounce_job)
        self._debounce_job = self.after(500, self._apply_filters)

    def __init__(self, master, app_container, on_save_callback=None, **kwargs):
        super().__init__(master)
        self.title("ประวัติใบสั่งซื้อ (Approved PO History)")
        self.geometry("1200x700")
        
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.on_save_callback = on_save_callback
        self.user_role = self.app_container.current_user_role
        self.theme = self.app_container.THEME.get("purchasing", {"primary": "#3B82F6"}) # กำหนด Theme
        
        # --- ตัวแปรสำหรับ Pagination และ Filter ---
        self.all_po_df = None
        self.filtered_df = None
        self.current_page = 0
        self.rows_per_page = 50
        self._debounce_job = None
        
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
        # --- Top Frame (Filter & Pagination) ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")
        
        filter_frame = CTkFrame(top_frame, fg_color="transparent")
        filter_frame.pack(side="left")

        month_options = ["ทุกเดือน"] + self.thai_months
        CTkOptionMenu(filter_frame, variable=self.month_var, values=month_options).pack(side="left", padx=5)

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        CTkOptionMenu(filter_frame, variable=self.year_var, values=year_options).pack(side="left", padx=5)
        
        # [แก้ไข] สร้างช่องค้นหาที่ขาดหายไป
        self.search_entry = CTkEntry(filter_frame, placeholder_text="ค้นหา SO / PO / Supplier...", width=200)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", lambda event: self._apply_filters())
        
        # [แก้ไข] เปลี่ยน command ให้ถูกต้อง (เป็น _apply_filters)
        CTkButton(filter_frame, text="ค้นหา", command=self._apply_filters, width=80).pack(side="left", padx=10)

        # Pagination Controls
        pagination_frame = CTkFrame(top_frame, fg_color="transparent")
        pagination_frame.pack(side="right")
        
        self.prev_button = CTkButton(pagination_frame, text="<<", command=self._prev_page, width=50, state="disabled")
        self.prev_button.pack(side="left", padx=5)
        self.page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.page_label.pack(side="left", padx=5)
        self.next_button = CTkButton(pagination_frame, text=">>", command=self._next_page, width=50, state="disabled")
        self.next_button.pack(side="left", padx=5)

        # --- Main Content Frame (ปรับปรุง Layout ให้เรียบง่าย) ---
        # ใช้ Frame เดียวแทน TabView เพราะหน้านี้แสดงแค่ประวัติที่ Approved แล้ว
        self.history_frame = CTkFrame(self, fg_color="transparent")
        self.history_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.history_frame.grid_rowconfigure(0, weight=1)
        self.history_frame.grid_columnconfigure(0, weight=1)
        
        # Bottom Frame (ปุ่ม Export)
        self.button_frame = CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="e")

        self.export_button = CTkButton(self.button_frame, text="Export to Excel", command=self._export_history, fg_color=self.theme.get("primary", "#3B82F6"))
        self.export_button.pack(side="left")

        self.loading_label = CTkLabel(self, text="กำลังโหลดข้อมูล...", font=CTkFont(size=18, slant="italic"), text_color="gray50")

    def _next_page(self):
        total_pages = (len(self.filtered_df) + self.rows_per_page - 1) // self.rows_per_page if self.filtered_df is not None else 0
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
            # ดึงเฉพาะ PO ที่ Approved แล้ว
            query = """
                SELECT 
                    po.id, 
                    po.timestamp, 
                    po.so_number, 
                    po.po_number, 
                    po.supplier_name,
                    owner.sale_name as owner_name,
                    proxy.sale_name as proxy_name,
                    po.status
                FROM purchase_orders po
                LEFT JOIN sales_users owner ON po.user_key = owner.sale_key
                LEFT JOIN sales_users proxy ON po.proxy_user_key = proxy.sale_key
                WHERE po.status = 'Approved'
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

        # Filter by Month
        selected_month_str = self.month_var.get()
        if selected_month_str != "ทุกเดือน":
            month_num = self.thai_month_map[selected_month_str]
            df = df[df['timestamp'].dt.month == month_num]

        # Filter by Year
        selected_year_str = self.year_var.get()
        if selected_year_str != "ทุกปี":
            year_num = int(selected_year_str)
            df = df[df['timestamp'].dt.year == year_num]

        # Filter by Search Term
        if hasattr(self, 'search_entry'):
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

        if self.filtered_df is None or self.filtered_df.empty:
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

    def _on_row_double_click(self, event, tree, use_iid=False):
        try:
            purchase_id_to_view = tree.focus()
            if not purchase_id_to_view: 
                return
            
            purchase_id = int(purchase_id_to_view)
            
            # เปิดหน้าต่างรายละเอียด PO
            self.app_container.show_purchase_detail_window(
                purchase_id=purchase_id,
                on_save_callback=self._load_initial_data
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดดูรายละเอียดได้: {e}", parent=self)

    def _create_styled_dataframe_table(self, parent, df):
        df = df.copy()
        df['display_owner'] = df.apply(
            lambda row: f"{row['owner_name']}" if pd.isna(row['proxy_name']) else f"{row['owner_name']} (โดย {row['proxy_name']})",
            axis=1
        )
        
        display_columns = {
            'timestamp': 'เวลาบันทึก',
            'so_number': 'SO Number',
            'po_number': 'PO Number',
            'supplier_name': 'Supplier',
            'display_owner': 'เจ้าของ PO (ผู้สร้าง)',
            'status': 'สถานะ'
        }
        
        columns_to_show = list(display_columns.keys())
        
        tree = ttk.Treeview(parent, columns=columns_to_show, show='headings')
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Roboto', 14, 'bold'))
        style.configure("Treeview", rowheight=25, font=('Roboto', 12))
        
        for col_id, col_text in display_columns.items():
            tree.heading(col_id, text=col_text)
            width = 200
            if col_id == 'timestamp': width = 180
            if col_id == 'display_owner': width = 250
            if col_id == 'status': width = 120
            tree.column(col_id, width=width, anchor='w')

        for index, row in df.iterrows():
            original_id = row['id']
            values = [row[col] for col in columns_to_show]
            if isinstance(values[0], pd.Timestamp):
                 values[0] = values[0].strftime('%Y-%m-%d %H:%M:%S')
            
            tree.insert("", "end", values=values, iid=str(original_id))

        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        # ผูก Event Double Click
        tree.bind("<Double-1>", lambda e: self._on_row_double_click(e, tree, use_iid=True))
        
    def _export_history(self):
        """Export ข้อมูลในตารางปัจจุบันเป็น Excel"""
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลที่จะ Export", parent=self)
            return

        try:
            default_filename = f"approved_po_history_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="บันทึกประวัติ PO",
                initialfile=default_filename,
                parent=self
            )

            if save_path:
                df_to_export = self.filtered_df.copy()
                # เปลี่ยนชื่อคอลัมน์ให้สวยงาม
                rename_map = {
                    'timestamp': 'เวลาบันทึก',
                    'so_number': 'SO Number',
                    'po_number': 'PO Number',
                    'supplier_name': 'Supplier',
                    'owner_name': 'เจ้าของ PO',
                    'proxy_name': 'ผู้สร้างแทน',
                    'status': 'สถานะ'
                }
                df_to_export.rename(columns=rename_map, inplace=True)
                
                # เลือกเฉพาะคอลัมน์ที่ต้องการ
                cols_to_keep = [c for c in rename_map.values() if c in df_to_export.columns]
                df_to_export = df_to_export[cols_to_keep]

                df_to_export.to_excel(save_path, index=False)
                messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้วที่:\n{save_path}", parent=self)
        
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=self)
            traceback.print_exc()


class CommissionHistoryWindow(CTkToplevel):
    def __init__(self, master, app_container, sale_key_filter=None, on_row_double_click=None, support_user_key_filter=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.sale_key_filter = sale_key_filter
        self.on_row_double_click_callback = on_row_double_click
        self.support_user_key_filter = support_user_key_filter
        self.df = None
        
        self.current_page = 0
        self.rows_per_page = 50
        self.total_rows = 0
        self.total_pages = 0
        self.active_tab = "deferral" # Default เป็นแท็บงานด่วน

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")
        
        title_text = f"ประวัติการบันทึกของ: {self.sale_key_filter}" if self.sale_key_filter else "ประวัติการบันทึก (Admin View)"
        self.title(title_text)
        self.geometry("1400x800")
        
        try: self.theme = self.app_container.THEME["sale"]
        except (AttributeError, KeyError): self.theme = {"header": "#1D4ED8", "primary": "#3B82F6"}
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._create_new_layout()
        
        self.after(50, self._populate_history_table)
        self.transient(master)
        self.grab_set()
        self.focus()

    def _create_new_layout(self):
        # Top Frame (Filter)
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")
        
        filter_frame = CTkFrame(top_frame, fg_color="transparent")
        filter_frame.pack(side="left")

        month_options = ["ทุกเดือน"] + self.thai_months
        CTkOptionMenu(filter_frame, variable=self.month_var, values=month_options).pack(side="left", padx=5)

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        CTkOptionMenu(filter_frame, variable=self.year_var, values=year_options).pack(side="left", padx=5)
        
        self.search_entry = CTkEntry(filter_frame, placeholder_text="ค้นหา SO / ลูกค้า...", width=200)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", lambda event: self._populate_history_table())
        
        CTkButton(filter_frame, text="ค้นหา", command=self._populate_history_table, width=80).pack(side="left", padx=10)

        # Pagination
        pagination_frame = CTkFrame(top_frame, fg_color="transparent")
        pagination_frame.pack(side="right")
        self.prev_button = CTkButton(pagination_frame, text="<<", command=self._prev_page, width=50, state="disabled")
        self.prev_button.pack(side="left", padx=5)
        self.page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.page_label.pack(side="left", padx=5)
        self.next_button = CTkButton(pagination_frame, text=">>", command=self._next_page, width=50, state="disabled")
        self.next_button.pack(side="left", padx=5)

        # --- Tab View (3 แท็บใหม่) ---
        self.tab_view = CTkTabview(self, command=self._on_tab_change)
        self.tab_view.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
        self.deferral_tab = self.tab_view.add("⚠️ รายการรอตัดสินใจ & ถูกเลื่อน")
        self.deferral_frame = CTkFrame(self.deferral_tab, fg_color="transparent"); self.deferral_frame.pack(fill="both", expand=True)

        self.payout_tab = self.tab_view.add("💰 ประวัติการรับเงิน")
        self.payout_frame = CTkFrame(self.payout_tab, fg_color="transparent"); self.payout_frame.pack(fill="both", expand=True)

        self.draft_tab = self.tab_view.add("📝 รายการฉบับร่าง / ส่งแล้ว")
        self.draft_frame = CTkFrame(self.draft_tab, fg_color="transparent"); self.draft_frame.pack(fill="both", expand=True)
        
        # Bottom Buttons
        self.button_frame = CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="e")
        
        # ปุ่ม Action (โชว์เฉพาะแท็บ Deferral)
        self.action_button = CTkButton(self.button_frame, text="⚡ จัดการรายการที่เลือก", command=self._open_deferral_action_window, 
                                       fg_color="#F59E0B", hover_color="#D97706", font=CTkFont(weight="bold"))
        self.action_button.pack(side="left", padx=10)
        
        self.cancel_button = CTkButton(self.button_frame, text="ยกเลิกรายการที่เลือก", command=self._cancel_selected_record, 
                                       fg_color="#DC2626", hover_color="#B91C1C")
        self.cancel_button.pack(side="left", padx=10)
        self.cancel_button.pack_forget() # ซ่อนก่อน
        
        self.export_button = CTkButton(self.button_frame, text="Export to Excel", command=self._export_history, fg_color=self.theme["primary"])
        self.export_button.pack(side="left")

        self.loading_label = CTkLabel(self, text="กำลังโหลดข้อมูล...", font=CTkFont(size=18, slant="italic"), text_color="gray50")

    def _on_tab_change(self):
        selected_tab = self.tab_view.get()
        if selected_tab == "⚠️ รายการรอตัดสินใจ & ถูกเลื่อน":
            self.active_tab = "deferral"
            self.action_button.pack(side="left", padx=10)
            self.cancel_button.pack_forget()
        elif selected_tab == "💰 ประวัติการรับเงิน":
            self.active_tab = "payout"
            self.action_button.pack_forget()
            self.cancel_button.pack_forget()
        else: # Drafts
            self.active_tab = "drafts"
            self.action_button.pack_forget()
            self.cancel_button.pack(side="left", padx=10)
        
        self.current_page = 0
        self._populate_history_table()

    def _populate_history_table(self):
        if self.active_tab == "deferral": target_frame = self.deferral_frame
        elif self.active_tab == "payout": target_frame = self.payout_frame
        else: target_frame = self.draft_frame

        for widget in target_frame.winfo_children(): widget.destroy()
        self._show_loading()

        try:
            # เงื่อนไข SQL ตาม Tab
            if self.active_tab == "deferral":
                status_condition = "c.status IN ('Defer Requested', 'Deferred')"
            elif self.active_tab == "payout":
                status_condition = "c.status IN ('Paid', 'HR Verified')"
            else: # drafts
                status_condition = "c.status IN ('Original', 'Edited', 'Draft', 'Rejected by SM', 'Rejected by HR', 'Cancelled', 'Deferred by HR', 'Deferred by SM', 'PO Sent')"

            where_clauses = ["c.is_active = 1", status_condition]
            params = []

            if hasattr(self, 'search_entry') and self.search_entry.get().strip():
                search_text = self.search_entry.get().strip()
                where_clauses.append("(c.so_number ILIKE %s OR c.customer_name ILIKE %s)")
                params.extend([f"%{search_text}%", f"%{search_text}%"])

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
            
            query_body = """
                FROM commissions c
                LEFT JOIN sales_users ss ON c.support_user_key = ss.sale_key
                LEFT JOIN sales_users su_owner ON c.sale_key = su_owner.sale_key
                LEFT JOIN commission_payout_logs cr ON c.payout_id = cr.id  
            """
            where_string = f"WHERE {' AND '.join(where_clauses)}"

            count_query = f"SELECT COUNT(c.id) {query_body} {where_string}"
            count_df = pd.read_sql_query(count_query, self.pg_engine, params=tuple(params))
            self.total_rows = count_df.iloc[0, 0] if not count_df.empty else 0
            self.total_pages = (self.total_rows + self.rows_per_page - 1) // self.rows_per_page

            offset = self.current_page * self.rows_per_page
            data_params = params + [self.rows_per_page, offset]

            data_query = f"""
                SELECT c.*,
                    ss.sale_name as support_user_name,
                    su_owner.sale_name as owner_name,
                    cr.commission_month AS paid_month, 
                    cr.commission_year AS paid_year
                {query_body}
                {where_string}
                ORDER BY c.timestamp DESC
                LIMIT %s OFFSET %s
            """
            self.df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(data_params))

            self.df['customer_display'] = self.df.apply(lambda row: f"{row['customer_name']} (คีย์โดย: {row['support_user_name']})" if pd.notna(row['support_user_name']) else f"{row['customer_name']} (คีย์โดย: {row.get('owner_name', 'N/A')})", axis=1)
            self.df['payment_period_display'] = self.df.apply(lambda row: f"{int(row['paid_month'])}/{int(row['paid_year'])}" if pd.notna(row['paid_month']) else "-", axis=1)

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
        columns = ['id', 'timestamp', 'status', 'so_number', 'customer_display', 'sales_service_amount']
        display_columns = ['ID', 'เวลาบันทึก', 'สถานะ', 'SO Number', 'ชื่อลูกค้า', 'ยอดขาย']
        
        if self.active_tab == "payout":
            columns.append('payment_period_display')
            display_columns.append('รอบจ่าย')
        elif self.active_tab == "deferral":
            columns.append('rejection_reason')
            display_columns.append('เหตุผล/สถานะการเลื่อน')

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("History.Treeview.Heading", font=('Roboto', 11, 'bold'), background="#E5E7EB")
        style.configure("History.Treeview", rowheight=28)
        
        self.tree = ttk.Treeview(parent, columns=columns, show='headings', style="History.Treeview")
        self.tree.pack(fill="both", expand=True)
        
        # Tags สี
        self.tree.tag_configure('Draft', background='#FEFCE8')
        self.tree.tag_configure('Rejected', background='#FEF2F2')
        self.tree.tag_configure('Defer Requested', background='#FFF7ED') # สีส้มอ่อน
        self.tree.tag_configure('Deferred', background='#F3F4F6') # สีเทา (ดอง)
        self.tree.tag_configure('Paid', background='#DCFCE7')

        for i, col_id in enumerate(columns):
            width = 100
            if col_id == 'so_number': width = 150
            elif col_id == 'customer_display': width = 250
            elif col_id == 'rejection_reason': width = 200
            self.tree.heading(col_id, text=display_columns[i])
            self.tree.column(col_id, width=width)

        for index, row in df.iterrows():
            status = row['status']
            tag = 'Default'
            if 'Reject' in status: tag = 'Rejected'
            elif status == 'Defer Requested': tag = 'Defer Requested'
            elif status == 'Deferred': tag = 'Deferred'
            elif status in ['Paid', 'HR Verified']: tag = 'Paid'
            elif status in ['Original', 'Draft']: tag = 'Draft'
            
            values = [row.get(col, '') for col in columns]
            # Format วันที่และตัวเลข
            if pd.notna(values[1]): values[1] = pd.to_datetime(values[1]).strftime('%Y-%m-%d %H:%M')
            if pd.notna(values[5]): values[5] = f"{float(values[5]):,.2f}"
            
            self.tree.insert("", "end", values=values, tags=(tag,), iid=str(row['id']))

        if self.on_row_double_click_callback:
            self.tree.bind("<Double-1>", lambda event: self._on_tree_row_double_click(event, self.tree))

    def _open_deferral_action_window(self):
        if not hasattr(self, 'tree') or not self.tree.focus():
            messagebox.showwarning("เตือน", "กรุณาเลือกรายการก่อน", parent=self)
            return
        
        item_id = self.tree.focus()
        index = self.tree.index(item_id)
        data = self.df.iloc[index]
        
        if data['status'] not in ['Defer Requested', 'Deferred']:
            messagebox.showinfo("ข้อมูล", "รายการนี้ไม่ได้อยู่ในสถานะที่ต้องจัดการเลื่อน", parent=self)
            return
            
        DeferralActionDialog(self, self.app_container, data, callback=self._populate_history_table)

    # --- Pagination & Utils ---
    def _update_pagination_controls(self):
        self.page_label.configure(text=f"Page {self.current_page + 1} / {max(1, self.total_pages)}")
        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")

    def _next_page(self):
        if self.current_page < self.total_pages - 1: self.current_page += 1; self._populate_history_table()

    def _prev_page(self):
        if self.current_page > 0: self.current_page -= 1; self._populate_history_table()

    def _show_loading(self):
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center"); self.update_idletasks()

    def _hide_loading(self):
        self.loading_label.place_forget()
        
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

    def _on_tree_row_double_click(self, event, tree):
        try:
            selected_item = tree.focus()
            if not selected_item: return
            record_id = int(selected_item)
            row_data = self.df[self.df['id'] == record_id].iloc[0]
            if self.on_row_double_click_callback: self.on_row_double_click_callback(row_data)
        except: pass
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

class DeferralActionDialog(CTkToplevel):
    """หน้าต่างสำหรับ Sale เลือกจัดการรายการที่ HR ขอเลื่อน (Defer)"""
    def __init__(self, master, app_container, record_data, callback=None):
        super().__init__(master)
        self.app_container = app_container
        self.record_data = record_data
        self.callback = callback
        
        so_number = record_data.get('so_number')
        reason = record_data.get('rejection_reason', '-').replace('HR Request:', '').strip()
        
        self.title(f"จัดการรายการขอเลื่อนจ่าย: {so_number}")
        
        # --- แก้ไข 1: ขยายขนาดให้สูงขึ้น (จาก 450/650 -> 700) และกว้างขึ้นเล็กน้อย ---
        self.geometry("550x700")
        
        # --- แก้ไข 2: ใช้คำสั่ง attributes("-topmost", True) เพื่อบังคับให้ลอยหน้าสุดเสมอ ---
        self.attributes("-topmost", True)
        
        # --- UI Layout ---
        main_frame = CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # ส่วนแสดงข้อมูล
        CTkLabel(main_frame, text="⚠️ HR ได้ส่งคำขอ 'เลื่อนจ่าย' รายการนี้", font=CTkFont(size=16, weight="bold"), text_color="#F59E0B").pack(pady=(0, 10))
        
        info_frame = CTkFrame(main_frame, fg_color=("gray90", "gray20"))
        info_frame.pack(fill="x", pady=10)
        
        CTkLabel(info_frame, text=f"SO Number: {so_number}", font=CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 2))
        CTkLabel(info_frame, text=f"ลูกค้า: {record_data.get('customer_name')}", font=CTkFont(size=12)).pack(anchor="w", padx=10, pady=2)
        CTkLabel(info_frame, text=f"ยอดคอมฯ: {record_data.get('sales_service_amount', 0):,.2f}", font=CTkFont(size=12)).pack(anchor="w", padx=10, pady=(2, 10))
        
        CTkLabel(main_frame, text="เหตุผลจาก HR:", font=CTkFont(weight="bold")).pack(anchor="w", pady=(10, 0))
        reason_box = CTkTextbox(main_frame, height=80, fg_color="transparent", border_width=1) # เพิ่มความสูงกล่องข้อความเล็กน้อย
        reason_box.insert("1.0", reason)
        reason_box.configure(state="disabled")
        reason_box.pack(fill="x", pady=5)
        
        CTkLabel(main_frame, text="กรุณาเลือกการดำเนินการ:", font=CTkFont(weight="bold")).pack(anchor="w", pady=(20, 10))
        
        # --- ปุ่มทางเลือก 1: ไม่ยอมเลื่อน (รับเงินเลย) ---
        self.btn_reject = CTkButton(main_frame, text="💸 ยืนยันรับเงินรอบนี้ (ไม่เลื่อน)", 
                                    command=self._on_reject_deferral,
                                    fg_color="#10B981", hover_color="#059669", # สีเขียว
                                    height=45, font=CTkFont(size=14, weight="bold")) # เพิ่มความสูงปุ่ม
        self.btn_reject.pack(fill="x", pady=5)
        
        # --- ปุ่มทางเลือก 2: ยอมเลื่อน (เลือกเดือน) ---
        defer_frame = CTkFrame(main_frame, fg_color="transparent")
        defer_frame.pack(fill="x", pady=15) # เพิ่มระยะห่าง
        
        CTkLabel(defer_frame, text="หรือ ยอมให้เลื่อนไปเดือน:").pack(side="left")
        
        # สร้างตัวเลือกเดือน: เริ่มจากเดือนปัจจุบัน (0) ไปจนถึง 12 เดือนข้างหน้า
        next_months = []
        curr = datetime.now()
        for i in range(0, 13): 
            next_date = curr + pd.DateOffset(months=i)
            next_months.append(next_date.strftime("%m/%Y"))
            
        self.target_month_var = tk.StringVar(value=next_months[0]) # Default เดือนปัจจุบัน
        self.month_menu = CTkOptionMenu(defer_frame, variable=self.target_month_var, values=next_months, width=140)
        self.month_menu.pack(side="left", padx=10)
        
        self.btn_confirm = CTkButton(defer_frame, text="ยืนยันการเลื่อน", 
                                     command=self._on_confirm_deferral,
                                     fg_color="#F59E0B", hover_color="#D97706", # สีส้ม
                                     width=120)
        self.btn_confirm.pack(side="left")

        # บังคับ Focus มาที่หน้าต่างนี้
        self.transient(master)
        self.grab_set()
        self.focus_force()

    def _on_reject_deferral(self):
        """Sale กดรับเงินเลย (Reject Deferral) -> สถานะกลับไปเป็น PO Sent (เพื่อให้ HR ตรวจสอบ/เทียบใหม่)"""
        if not messagebox.askyesno("ยืนยัน", "คุณต้องการรับเงินในรอบนี้ทันที ใช่หรือไม่?\n(รายการจะถูกส่งกลับไปให้ HR ตรวจสอบอีกครั้ง)"):
            return
            
        self._update_status(
            new_status='PO Sent',  # <--- แก้ไขตรงนี้: เปลี่ยนจาก 'HR Verified' เป็น 'PO Sent'
            log_msg="Sale Rejected Deferral (Requesting Payment Now)",
            new_month=None, new_year=None
        )
    def _on_confirm_deferral(self):
        """Sale ยอมเลื่อน (Confirm Deferral) -> สถานะเป็น Deferred + ย้ายเดือน"""
        target_str = self.target_month_var.get() # "MM/YYYY"
        try:
            t_month, t_year = map(int, target_str.split('/'))
            
            msg = f"คุณยืนยันที่จะเลื่อนการรับเงินรายการนี้ ไปเป็นเดือน {t_month}/{t_year} ใช่หรือไม่?"
            if not messagebox.askyesno("ยืนยันการเลื่อน", msg):
                return

            self._update_status(
                new_status='Deferred',
                log_msg=f"Sale Confirmed Deferral to {t_month}/{t_year}",
                new_month=t_month, new_year=t_year
            )
        except ValueError:
            messagebox.showerror("Error", "รูปแบบเดือนไม่ถูกต้อง")

    def _update_status(self, new_status, log_msg, new_month=None, new_year=None):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # สร้าง SQL Update
                sql = "UPDATE commissions SET status = %s, rejection_reason = %s"
                params = [new_status, log_msg]
                
                # ถ้ามีการย้ายเดือน (กรณีเลื่อน) ให้แก้ commission_month/year ด้วย
                if new_month and new_year:
                    sql += ", commission_month = %s, commission_year = %s"
                    params.extend([new_month, new_year])
                
                sql += " WHERE id = %s"
                params.append(self.record_data['id'])
                
                cursor.execute(sql, tuple(params))
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกการตัดสินใจเรียบร้อยแล้ว", parent=self)
            
            if self.callback:
                self.callback() # Refresh หน้าจอประวัติ
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

# ในไฟล์ history_windows.py

# ในไฟล์ history_windows.py

class SOPopupWindow(CTkToplevel):
    def __init__(self, master, app_container, sales_data, so_shared_vars, sale_theme, on_save_callback=None):
        super().__init__(master)
        self.master = master
        self.app_container = app_container
        self.sales_data = sales_data
        self.so_shared_vars = so_shared_vars.copy() # ใช้ .copy() เพื่อไม่ให้กระทบตัวแปรต้นทาง
        self.sale_theme = sale_theme
        self.on_save_callback = on_save_callback
        self.popup_widgets = {}
        self.trace_ids_for_so_calc = []

        # --- [แก้ไข 1] กำหนด Style ภายในตัวเอง ---
        self.dropdown_style = {
            "fg_color": "white",
            "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }

        # --- [แก้ไข 2] สร้างตัวแปร StringVar ที่ขาดหายไปให้ครบถ้วน (ป้องกัน KeyError) ---
        required_vars = [
            'sales_service_vat_option', 
            'cutting_drilling_fee_vat_option', 
            'other_service_fee_vat_option', 
            'shipping_vat_option_var', 
            'credit_card_fee_vat_option_var',
            'relocation_cost_vat_option',  # <--- ตัวที่ทำให้เกิด Error
            'delivery_type_var'
        ]
        
        for key in required_vars:
            if key not in self.so_shared_vars:
                # ถ้าไม่มี ให้สร้างใหม่เป็นค่า Default 'VAT'
                self.so_shared_vars[key] = tk.StringVar(value="VAT")
        
        # กำหนดค่าเริ่มต้นให้กับ Delivery Type ถ้าเพิ่งสร้างใหม่
        if self.so_shared_vars['delivery_type_var'].get() == "VAT": 
             self.so_shared_vars['delivery_type_var'].set("ซัพพลายเออร์จัดส่ง")

        # สร้างตัวแปรสำหรับเก็บผลลัพธ์การคำนวณ (Calculation Vars)
        calc_vars = [
            'sales_vat_calc_var', 'cutting_drilling_vat_calc_var',
            'other_service_fee_vat_calc_var', 'shipping_vat_calc_var',
            'card_fee_vat_calc_var', 'relocation_vat_calc_var',
            'payment_total_var', 'so_grand_total_var', 'difference_amount_var',
            'cash_required_total_var', 'so_vs_payment_result_var', 'cash_verification_result_var',
            'cash_service_total_var'  # 🟢 เพิ่มตัวแปรนี้เข้ามาบรรทัดนี้ครับ!
        ]
        for var_name in calc_vars:
            if var_name not in self.so_shared_vars:
                self.so_shared_vars[var_name] = tk.StringVar(value="0.00" if "result" not in var_name else "-")

        self.title(f"ข้อมูล Sales Order (SO: {sales_data.get('so_number', 'N/A')})")
        self.geometry("700x750")
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
            except tk.TclError:
                pass
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
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=(15, 10), pady=4, sticky="w")
        
        item_frame = CTkFrame(parent, fg_color="transparent")
        item_frame.grid(row=row_index, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")
        item_frame.grid_columnconfigure(0, weight=1)

        amount_entry = NumericEntry(item_frame)
        amount_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.popup_widgets[entry_key] = amount_entry

        vat_frame = CTkFrame(item_frame, fg_color="transparent")
        vat_frame.pack(side="left")
        
        # ตรงนี้คือจุดที่เคย Error เพราะ vat_option_key ไม่ได้อยู่ใน so_shared_vars
        # แต่ตอนนี้เราแก้ใน __init__ แล้ว จึงปลอดภัย
        CTkRadioButton(vat_frame, text="VAT", variable=self.so_shared_vars[vat_option_key], value="VAT").pack(side="left")
        CTkRadioButton(vat_frame, text="CASH", variable=self.so_shared_vars[vat_option_key], value="CASH").pack(side="left", padx=5)

        vat_label = CTkLabel(item_frame, textvariable=self.so_shared_vars[vat_display_var_key], font=CTkFont(size=12), text_color="gray50")
        vat_label.pack(side="left", padx=(10, 0))
            
    def _create_so_data_form_content(self, parent_frame):
        # Section 1: Sales Details
        f1 = self._create_so_section_frame(parent_frame, "รายละเอียดการขาย")
        self._add_form_row(f1, "วันที่เปิด SO:", DateSelector(f1, dropdown_style=self.dropdown_style), 'bill_date_selector', 1)
        self._add_form_row(f1, "ชื่อลูกค้า:", CTkEntry(f1), 'customer_name_entry', 2)
        self._add_form_row(f1, "รหัสลูกค้า:", CTkEntry(f1), 'customer_id_entry', 3)
        self._add_form_row(f1, "Credit Term:", CTkEntry(f1), 'credit_term_entry', 4)

        # Section 2: Sales and Services
        f2 = self._create_so_section_frame(parent_frame, "ยอดขายและบริการ")
        self._add_item_row_with_vat(f2, "ยอดขายสินค้า/บริการ:", 'sales_amount_entry', 'sales_service_vat_option', 'sales_vat_calc_var', 1)
        self._add_item_row_with_vat(f2, "ค่าบริการตัด/เจาะ:", 'cutting_drilling_fee_entry', 'cutting_drilling_fee_vat_option', 'cutting_drilling_vat_calc_var', 2)
        self._add_item_row_with_vat(f2, "ค่าบริการอื่นๆ:", 'other_service_fee_entry', 'other_service_fee_vat_option', 'other_service_vat_calc_var', 3)
        
        # Section 3: Shipping Cost
        f3 = self._create_so_section_frame(parent_frame, "ค่าจัดส่ง")
        self._add_item_row_with_vat(f3, "ค่าจัดส่ง:", 'shipping_cost_entry', 'shipping_vat_option_var', 'shipping_vat_calc_var', 1)
        self._add_form_row(f3, "วันที่จัดส่ง:", DateSelector(f3, dropdown_style=self.dropdown_style), 'delivery_date_selector', 2)

        # Section 4: Delivery Note
        f4 = self._create_so_section_frame(parent_frame, "Delivery Note")
        delivery_options = [
            "ซัพพลายเออร์จัดส่ง", "Aplus Logistic ส่งหน้างาน", "ลูกค้ารับเองที่ซัพ",
            "ลูกค้ารับเองที่คลัง 132", "ย้ายเข้าคลัง Aplus Logistic รอลูกค้ารับที่คลัง",
            "ย้ายเข้าคลัง Aplus Logistic รอ Aplus Logistic จัดส่ง",
            "ย้ายเข้าคลัง Lalamove รอลูกค้ารับที่คลัง 132", "ส่ง Lalamove ให้ลูกค้าหน้างาน",
            "Aplus Logistic+ฝากส่งขนส่ง", "Lalamove +ฝากส่งขนส่ง"
        ]
        self._add_form_row(f4, "การจัดส่ง:", CTkOptionMenu(f4, variable=self.so_shared_vars['delivery_type_var'], values=delivery_options, **self.dropdown_style), 'delivery_type_menu', 1)
        self._add_form_row(f4, "Location เข้ารับ:", CTkEntry(f4, placeholder_text="ใส่ อำเภอ, จังหวัด หรือ Google map link"), 'pickup_location_entry', 2)
        self._add_item_row_with_vat(f4, "ค่าย้าย:", 'relocation_cost_entry', 'relocation_cost_vat_option', 'relocation_vat_calc_var', 3)
        self._add_form_row(f4, "วันที่ย้ายเข้าคลัง:", DateSelector(f4, dropdown_style=self.dropdown_style), 'date_to_wh_selector', 4)
        self._add_form_row(f4, "วันที่จัดส่งลูกค้า:", DateSelector(f4, dropdown_style=self.dropdown_style), 'date_to_customer_selector', 5)
        self._add_form_row(f4, "ทะเบียนเข้ารับ:", CTkEntry(f4), 'pickup_rego_entry', 6)
        
        # 🟢 [เพิ่มใหม่] เงื่อนไขลงสินค้า และ Special Request
        if 'unloading_status_var' not in self.so_shared_vars:
            self.so_shared_vars['unloading_status_var'] = tk.StringVar(value="ไม่รวมลง")
        unloading_frame = CTkFrame(f4, fg_color="transparent")
        CTkRadioButton(unloading_frame, text="รวมลง", variable=self.so_shared_vars['unloading_status_var'], value="รวมลง").pack(side="left", padx=5)
        CTkRadioButton(unloading_frame, text="ไม่รวมลง", variable=self.so_shared_vars['unloading_status_var'], value="ไม่รวมลง").pack(side="left", padx=5)
        self._add_form_row(f4, "เงื่อนไขลงสินค้า:", unloading_frame, 'unloading_status_radio', 7)
        self._add_form_row(f4, "Special Request:", CTkEntry(f4), 'special_request_entry', 8)

        # Section 5: Fees and Discounts
        f5 = self._create_so_section_frame(parent_frame, "ค่าธรรมเนียมและส่วนลด")
        self._add_item_row_with_vat(f5, "ค่าธรรมเนียมบัตร:", 'credit_card_fee_entry', 'credit_card_fee_vat_option_var', 'card_fee_vat_calc_var', 1)
        self._add_form_row(f5, "ค่าธรรมเนียมโอน:", NumericEntry(f5), 'transfer_fee_entry', 2)
        self._add_form_row(f5, "ภาษีหัก ณ ที่จ่าย:", NumericEntry(f5), 'wht_fee_entry', 3)
        self._add_form_row(f5, "ค่านายหน้า:", NumericEntry(f5), 'brokerage_fee_entry', 4)
        
        # 🟢 [แก้ไข] เปลี่ยนของแถมเป็น CTkEntry ให้พิมพ์ตัวหนังสือได้
        self._add_form_row(f5, "ของแถมใน SO (Vat):", CTkEntry(f5), 'giveaway_vat_entry', 5)
        self._add_form_row(f5, "ของแถมนอก SO (No Vat):", CTkEntry(f5), 'giveaway_no_vat_entry', 6)
        self._add_form_row(f5, "คูปอง:", NumericEntry(f5), 'coupon_value_entry', 7)                # 🟢 แก้ไข (เลื่อนเป็นแถว 7)

        # Section 6: Payment Details [แยกวันที่ 1 และ 2]
        f6 = self._create_so_section_frame(parent_frame, "รายละเอียดการโอนชำระ")
        
        self._add_form_row(f6, "ยอดโอนชำระ 1:", NumericEntry(f6), 'payment1_amount_entry', 1)
        self._add_form_row(f6, "วันที่ชำระ 1:", DateSelector(f6, dropdown_style=self.dropdown_style), 'payment1_date_selector', 2)
        
        self._add_form_row(f6, "ยอดโอนชำระ 2:", NumericEntry(f6), 'payment2_amount_entry', 3)
        self._add_form_row(f6, "วันที่ชำระ 2:", DateSelector(f6, dropdown_style=self.dropdown_style), 'payment2_date_selector', 4)
        
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
        
        # Save Button
        save_button = CTkButton(parent_frame, text="บันทึกข้อมูล SO", command=self._save_so_changes, fg_color="#16A34A", hover_color="#15803D", font=CTkFont(size=16, weight="bold"))
        save_button.pack(fill="x", padx=10, pady=20)

    def _so_bind_events(self):
        self.trace_ids_for_so_calc = []
        widgets_to_bind_keys = [
            "sales_amount_entry", "cutting_drilling_fee_entry", "other_service_fee_entry",
            "shipping_cost_entry", "credit_card_fee_entry", "transfer_fee_entry",
            "wht_fee_entry", "coupon_value_entry",
            "brokerage_fee_entry", "payment1_amount_entry", "payment2_amount_entry",
            "cash_product_input_entry", "cash_actual_payment_entry",
            "relocation_cost_entry" 
        ]
        for key in widgets_to_bind_keys:
            if key in self.popup_widgets and isinstance(self.popup_widgets[key], (CTkEntry, NumericEntry)):
                self.popup_widgets[key].bind("<KeyRelease>", self._so_update_final_calculations)
            
        radio_vars_keys = [
            'sales_service_vat_option', 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option', 'shipping_vat_option_var',
            'credit_card_fee_vat_option_var',
            'relocation_cost_vat_option' 
        ]
        for key in radio_vars_keys:
            if key in self.so_shared_vars and isinstance(self.so_shared_vars[key], tk.StringVar):
                trace_id = self.so_shared_vars[key].trace_add("write", self._so_update_final_calculations)
                self.trace_ids_for_so_calc.append((self.so_shared_vars[key], trace_id))

    def _so_update_final_calculations(self, *args):
        if not self.winfo_exists(): return

        w_vars = self.so_shared_vars
        w_widgets = self.popup_widgets

        # Helper สำหรับดึงค่าตัวเลขจาก Entry
        def get_float_from_entry(entry_key):
            entry_widget = w_widgets.get(entry_key)
            if entry_widget and entry_widget.winfo_exists():
                try: return utils.convert_to_float(entry_widget.get())
                except (tk.TclError, ValueError): return 0.0
            return 0.0

        # --- 1. ดึงข้อมูลตัวเลขจากฟอร์ม ---
        sales_amt = get_float_from_entry('sales_amount_entry')
        cutting_fee = get_float_from_entry('cutting_drilling_fee_entry')
        other_fee = get_float_from_entry('other_service_fee_entry')
        shipping_cost = get_float_from_entry('shipping_cost_entry')
        relocation_cost = get_float_from_entry('relocation_cost_entry')
        card_fee = get_float_from_entry('credit_card_fee_entry')
        transfer_fee = get_float_from_entry('transfer_fee_entry')
        wht = get_float_from_entry('wht_fee_entry')
        coupons = get_float_from_entry('coupon_value_entry')

        # --- 2. คัดแยกยอด VAT กับ ยอด CASH ---
        items_to_process = [
            (sales_amt, w_vars.get('sales_service_vat_option'), w_vars.get('sales_vat_calc_var')),
            (cutting_fee, w_vars.get('cutting_drilling_fee_vat_option'), w_vars.get('cutting_drilling_vat_calc_var')),
            (other_fee, w_vars.get('other_service_fee_vat_option'), w_vars.get('other_service_vat_calc_var')),
            (shipping_cost, w_vars.get('shipping_vat_option_var'), w_vars.get('shipping_vat_calc_var')),
            (card_fee, w_vars.get('credit_card_fee_vat_option_var'), w_vars.get('card_fee_vat_calc_var')),
            (relocation_cost, w_vars.get('relocation_cost_vat_option'), w_vars.get('relocation_vat_calc_var'))
        ]

        total_vatable_base = 0.0  
        total_vat = 0.0           
        total_cashable_services = 0.0 

        for amount, option_var, display_var in items_to_process:
            if not option_var: continue
            
            item_vat = 0.0
            if option_var.get() == "VAT":
                total_vatable_base += amount
                item_vat = amount * 0.07
                total_vat += item_vat
            else:
                # ถ้าระบุเป็น CASH (หรือ NO VAT) จะถูกบวกเก็บไว้ที่นี่
                total_cashable_services += amount 
            
            if display_var:
                display_var.set(f"VAT: {item_vat:,.2f}")

        # --- 3. คำนวณยอดรวมที่ต้องชำระโอน (Grand Total ฝั่ง VAT) ---
        final_grand_total = (total_vatable_base + total_vat + transfer_fee) - wht
        w_vars['so_grand_total_var'].set(f"{final_grand_total:,.2f}")

        # --- 4. คำนวณส่วนต่างการโอน ---
        payment1 = get_float_from_entry('payment1_amount_entry')
        payment2 = get_float_from_entry('payment2_amount_entry')
        total_payment = payment1 + payment2
        w_vars['payment_total_var'].set(f"{total_payment:,.2f}")
        
        so_vs_payment_diff = total_payment - final_grand_total
        w_vars['difference_amount_var'].set(f"{so_vs_payment_diff:,.2f}")

        # --- 5. อัปเดต Label สีเขียว/แดง (โอนขาด/เกิน) ---
        def set_check_result(label_widget_key, var, diff_val, plus_text, minus_text):
            label_widget_ref = w_widgets.get(label_widget_key)
            if not (label_widget_ref and label_widget_ref.winfo_exists()): return
            
            color_map = {"-": ("gray85", "black"), "ok": ("#BBF7D0", "#15803D"), "bad": ("#FECACA", "#B91C1C")}
            
            if abs(diff_val) < 0.01: state, text = "ok", "ถูกต้อง"
            elif diff_val > 0: state, text = "ok", f"{plus_text} (+{abs(diff_val):,.2f})"
            else: state, text = "bad", f"{minus_text} ({abs(diff_val):,.2f})"
            
            var.set(text)
            label_widget_ref.configure(fg_color=color_map[state][0], text_color=color_map[state][1])

        set_check_result('so_check_display', w_vars.get('so_vs_payment_result_var'), so_vs_payment_diff, "ยอดโอนเกิน", "ยอดโอนขาด")

        # --- 6. คำนวณยอดเงินสด (Cash) 🟢 แก้ไข: เพิ่มการแสดงผลยอดรวมบริการเงินสด ---
        
        # แสดงยอดรวมค่าบริการที่ติ๊ก "CASH" ทั้งหมด ออกมาที่หน้าจอ
        w_vars['cash_service_total_var'].set(f"{total_cashable_services:,.2f}")
        
        cash_product_val = get_float_from_entry('cash_product_input_entry')
        
        # ยอดเงินสดที่ต้องจ่ายทั้งหมด = ค่าสินค้าเงินสด + ค่าบริการเงินสด
        cash_required_total = cash_product_val + total_cashable_services
        w_vars['cash_required_total_var'].set(f"{cash_required_total:,.2f}")
        
        # ตรวจสอบส่วนต่างเงินสด
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

        # --- Logic: ดึงข้อมูลเก่ามาใส่ ถ้าข้อมูลใหม่ว่างเปล่า ---
        val_p1_date = data.get('payment1_date')
        val_p2_date = data.get('payment2_date')
        
        if pd.isna(val_p1_date) and pd.isna(val_p2_date):
            val_p1_date = data.get('payment_date')
            
        def to_float_safe(x):
            try:
                return float(x) if x is not None else 0.0
            except:
                return 0.0

        val_p1_amt = to_float_safe(data.get('payment1_amount'))
        val_p2_amt = to_float_safe(data.get('payment2_amount'))
        total_pay_db = to_float_safe(data.get('total_payment_amount'))

        if val_p1_amt == 0 and val_p2_amt == 0 and total_pay_db > 0:
             val_p1_amt = total_pay_db
             if pd.isna(val_p1_date):
                 val_p1_date = data.get('payment_date')

        key_map = {
            'bill_date': 'bill_date_selector', 'customer_name': 'customer_name_entry', 'customer_id': 'customer_id_entry',
            'credit_term': 'credit_term_entry', 'sales_service_amount': 'sales_amount_entry', 'cutting_drilling_fee': 'cutting_drilling_fee_entry',
            'other_service_fee': 'other_service_fee_entry', 'shipping_cost': 'shipping_cost_entry', 'delivery_date': 'delivery_date_selector',
            'credit_card_fee': 'credit_card_fee_entry', 'transfer_fee': 'transfer_fee_entry', 'wht_3_percent': 'wht_fee_entry',
            'brokerage_fee': 'brokerage_fee_entry', 'coupons': 'coupon_value_entry', 
            
            # 🟢 แก้ไขของแถมและหน้างาน
            'giveaway_vat': 'giveaway_vat_entry', 'giveaway_no_vat': 'giveaway_no_vat_entry', 
            'special_request': 'special_request_entry', 'unloading_status': 'unloading_status_var',
            
            'cash_product_input': 'cash_product_input_entry', 'cash_actual_payment': 'cash_actual_payment_entry',
            'sales_service_vat_option': 'sales_service_vat_option', 'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option': 'other_service_fee_vat_option', 'shipping_vat_option': 'shipping_vat_option_var',
            'credit_card_fee_vat_option': 'credit_card_fee_vat_option_var', 'so_grand_total': 'so_grand_total_var',
            'so_vs_payment_result': 'so_vs_payment_result_var', 'difference_amount': 'difference_amount_var',
            'cash_required_total': 'cash_required_total_var', 'cash_verification_result': 'cash_verification_result_var',
            'delivery_type': 'delivery_type_var', 'pickup_location': 'pickup_location_entry',
            'relocation_cost': 'relocation_cost_entry', 'date_to_warehouse': 'date_to_wh_selector',
            
            'payment_before_vat': 'payment_before_vat_entry', 
            'payment_no_vat': 'payment_no_vat_entry',
            
            'date_to_customer': 'date_to_customer_selector', 'pickup_registration': 'pickup_rego_entry',
            'relocation_cost_vat_option': 'relocation_cost_vat_option'
        }
        
        for key, widget in self.popup_widgets.items():
            if isinstance(widget, (CTkEntry, NumericEntry, AutoCompleteEntry)): set_val(widget, "")
            elif isinstance(widget, DateSelector): set_val(widget, None)
            elif isinstance(widget, CTkLabel): widget.configure(text="")
        
        for key, var in self.so_shared_vars.items():
            if isinstance(var, tk.StringVar): var.set("")
        
        if data is not None:
            # 1. วนลูปใส่ข้อมูลทั่วไป
            for db_key, w_key in key_map.items():
                widget_or_var = self.so_shared_vars.get(w_key) or self.popup_widgets.get(w_key)
                if widget_or_var:
                    set_val(widget_or_var, data.get(db_key))
            
            # 2. [สำคัญ] ใส่ข้อมูล Payment ที่เราเตรียมไว้ (Fallback logic)
            set_val(self.popup_widgets.get('payment1_amount_entry'), val_p1_amt)
            set_val(self.popup_widgets.get('payment2_amount_entry'), val_p2_amt)
            set_val(self.popup_widgets.get('payment1_date_selector'), val_p1_date)
            set_val(self.popup_widgets.get('payment2_date_selector'), val_p2_date)
        
        self.update_idletasks()
        self._so_update_final_calculations()
    
    def _save_so_changes(self):
        """
        บันทึกข้อมูล SO (อัปเดตการคำนวณ Grand Total ก่อนบันทึกให้ตรงกับหน้าจอ)
        """
        if self.sales_data is None: 
            messagebox.showerror("ข้อผิดพลาด", "ไม่มีข้อมูล SO ให้บันทึก", parent=self)
            return

        so_id = self.sales_data.get('id')
        updated_data = {}
        
        # 1. Mapping ชื่อ Widget -> ชื่อคอลัมน์ DB
        key_map = {
            'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id', 'credit_term_entry': 'credit_term',
            'pickup_location_entry': 'pickup_location', 'pickup_rego_entry': 'pickup_registration',
            'bill_date_selector': 'bill_date', 'delivery_date_selector': 'delivery_date', 
            'payment1_date_selector': 'payment1_date', 'payment2_date_selector': 'payment2_date',
            'payment1_amount_entry': 'payment1_amount', 'payment2_amount_entry': 'payment2_amount',
            'date_to_wh_selector': 'date_to_warehouse', 'date_to_customer_selector': 'date_to_customer',
            'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost',
            'relocation_cost_entry': 'relocation_cost', 'credit_card_fee_entry': 'credit_card_fee',
            'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons',
            
            # 🟢 [แก้ไข] ของแถม และหน้างาน
            'giveaway_vat_entry': 'giveaway_vat', 'giveaway_no_vat_entry': 'giveaway_no_vat', 
            'special_request_entry': 'special_request',
            
            'cash_product_input_entry': 'cash_product_input',
            'cash_actual_payment_entry': 'cash_actual_payment'
        }

        # 2. Loop เก็บข้อมูล
        for widget_key, data_key in key_map.items():
            value = None
            if widget_key in self.popup_widgets:
                widget = self.popup_widgets[widget_key]
                try:
                    if not widget.winfo_exists(): continue
                except: continue

                if isinstance(widget, DateSelector): 
                    value = widget.get_date()
                elif isinstance(widget, (NumericEntry, CTkEntry)):
                    raw_val = widget.get()
                    # 🟢 [แก้ไข] ลบ giveaways ออกจาก list นี้
                    numeric_keywords = ['amount', 'cost', 'fee', 'wht', 'percent', 'coupons', 'input', 'payment']
                    is_numeric = any(k in data_key for k in numeric_keywords)

                    if is_numeric:
                        if raw_val is None or str(raw_val).strip() == "":
                            value = 0.0 
                        else:
                            try:
                                value = float(str(raw_val).replace(",", ""))
                            except ValueError:
                                value = 0.0
                    else:
                        value = raw_val
            
            if value is not None: 
                updated_data[data_key] = value

        # 3. จัดการ Dropdown (VAT Options)
        shared_vars_map = {
            'delivery_type_var': 'delivery_type', 
            'sales_service_vat_option': 'sales_service_vat_option',
            'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option', 
            'other_service_fee_vat_option': 'other_service_fee_vat_option',
            'shipping_vat_option_var': 'shipping_vat_option', 
            'credit_card_fee_vat_option_var': 'credit_card_fee_vat_option',
            'relocation_cost_vat_option': 'relocation_cost_vat_option',
            'unloading_status_var': 'unloading_status' # 🟢 เพิ่มใหม่
        }
        for var_key, data_key in shared_vars_map.items():
            if var_key in self.so_shared_vars: 
                updated_data[data_key] = self.so_shared_vars[var_key].get()

        # --- 4. 🟢 แก้ไข: คำนวณยอด Grand Total โดยแยก VAT/CASH ออกจากกัน ---
        sales = updated_data.get('sales_service_amount', 0.0)
        cutting = updated_data.get('cutting_drilling_fee', 0.0)
        other = updated_data.get('other_service_fee', 0.0)
        shipping = updated_data.get('shipping_cost', 0.0)
        relocation = updated_data.get('relocation_cost', 0.0) 
        card_fee = updated_data.get('credit_card_fee', 0.0)
        transfer_fee = updated_data.get('transfer_fee', 0.0)
        coupons = updated_data.get('coupons', 0.0)
        wht = updated_data.get('wht_3_percent', 0.0)

        def get_vatable_base(amount, option_key):
            """คืนค่าจำนวนเงินเฉพาะถ้าเลือกเป็น VAT"""
            opt = updated_data.get(option_key, 'No VAT')
            return amount if opt == 'VAT' else 0.0

        def is_cash(option_key):
            """เช็คว่าเลือกเป็น CASH หรือไม่"""
            return updated_data.get(option_key, 'No VAT') != 'VAT'

        # รวมเฉพาะยอดฐานที่เป็น VAT
        vatable_base_sum = (
            get_vatable_base(sales, 'sales_service_vat_option') +
            get_vatable_base(cutting, 'cutting_drilling_fee_vat_option') +
            get_vatable_base(other, 'other_service_fee_vat_option') +
            get_vatable_base(shipping, 'shipping_vat_option') +
            get_vatable_base(relocation, 'relocation_cost_vat_option') +
            get_vatable_base(card_fee, 'credit_card_fee_vat_option')
        )
        
        vat_sum = vatable_base_sum * 0.07

        # คำนวณยอดรวมโอน (เอาเฉพาะฐาน VAT + VAT + โอน - คูปอง - WHT)
        grand_total_calc = (vatable_base_sum + vat_sum + transfer_fee) - wht
        
        # 5. คำนวณยอดชำระโอน
        p1 = updated_data.get('payment1_amount', 0.0)
        p2 = updated_data.get('payment2_amount', 0.0)
        total_paid = p1 + p2
        updated_data['total_payment_amount'] = total_paid
        updated_data['difference_amount'] = total_paid - grand_total_calc

        # 6. คำนวณยอดเงินสด (รวมเฉพาะรายการที่เป็น CASH)
        cash_product_input = updated_data.get('cash_product_input', 0.0)
        cash_services = (
            (sales if is_cash('sales_service_vat_option') else 0) +
            (cutting if is_cash('cutting_drilling_fee_vat_option') else 0) +
            (other if is_cash('other_service_fee_vat_option') else 0) +
            (shipping if is_cash('shipping_vat_option') else 0) +
            (relocation if is_cash('relocation_cost_vat_option') else 0) +
            (card_fee if is_cash('credit_card_fee_vat_option') else 0)
        )
        updated_data['cash_required_total'] = cash_product_input + cash_services

        # 7. จัดการวันที่จ่าย
        p1_date = updated_data.get('payment1_date')
        p2_date = updated_data.get('payment2_date')
        main_payment_date = None
        if p1_date and p2_date:
            try: main_payment_date = max(p1_date, p2_date)
            except: main_payment_date = p1_date
        elif p1_date: main_payment_date = p1_date
        elif p2_date: main_payment_date = p2_date
        updated_data['payment_date'] = main_payment_date
        updated_data['timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # =====================================================================
        # เริ่มกระบวนการบันทึก
        # =====================================================================
        success = False      
        error_message = None 

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'commissions'")
                valid_columns = {row[0] for row in cursor.fetchall()}

                final_data = {k: v for k, v in updated_data.items() if k in valid_columns}

                if final_data:
                    set_clauses = [f'"{k}" = %s' for k in final_data.keys()]
                    params = list(final_data.values()) + [so_id]
                    
                    sql_update = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
                    cursor.execute(sql_update, tuple(params))

                    cursor.execute("""
                        INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                        VALUES (%s, %s, %s, %s, %s, NOW())
                    """, ('Edit SO (Fix Cash Calc)', 'commissions', so_id, self.app_container.current_user_key, "Saved with corrected calculations"))

            conn.commit()
            success = True 
        
        except Exception as e:
            if conn: conn.rollback()
            error_message = str(e) 
            traceback.print_exc()
        
        finally:
            if conn: self.app_container.release_connection(conn)

        if success:
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูล SO และการชำระเงินเรียบร้อยแล้ว", parent=self)
            if self.on_save_callback: 
                self.on_save_callback()
            self.destroy()
            
        elif error_message:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด:\n{error_message}", parent=self)
            
class SOReassignmentDialog(CTkToplevel):
    """
    หน้าต่างสำหรับ Sale Support ใช้ย้ายความเป็นเจ้าของ SO (Reassign Sale Key)
    * รองรับการย้ายทุกสถานะ (Draft, Pending, PO In Progress, etc.)
    """
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.title("เครื่องมือย้ายเจ้าของ SO (Reassign Owner)")
        self.geometry("500x500")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- ส่วนที่ 1: ค้นหา SO ---
        search_frame = CTkFrame(self)
        search_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        search_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(search_frame, text="ระบุเลขที่ SO:", font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=10, pady=10)
        self.so_entry = CTkEntry(search_frame, placeholder_text="เช่น SO6811AM001")
        self.so_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.so_entry.bind("<Return>", lambda e: self._search_so())
        
        search_btn = CTkButton(search_frame, text="ค้นหา", width=100, command=self._search_so)
        search_btn.grid(row=0, column=2, padx=10, pady=10)
        
        # --- ส่วนที่ 2: แสดงรายละเอียดและเลือกเจ้าของใหม่ ---
        self.detail_frame = CTkFrame(self)
        self.detail_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.detail_frame.grid_columnconfigure(1, weight=1)

        # ซ่อนไว้ก่อน จนกว่าจะกดค้นหาเจอ
        self.detail_frame.grid_remove() 

        # ข้อมูลปัจจุบัน
        CTkLabel(self.detail_frame, text="ลูกค้า:", font=CTkFont(size=12)).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.lbl_customer = CTkLabel(self.detail_frame, text="-", font=CTkFont(weight="bold"))
        self.lbl_customer.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        CTkLabel(self.detail_frame, text="เจ้าของปัจจุบัน:", font=CTkFont(size=12)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.lbl_current_sale = CTkLabel(self.detail_frame, text="-", font=CTkFont(weight="bold"), text_color="#D97706") # สีส้ม
        self.lbl_current_sale.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        CTkLabel(self.detail_frame, text="สถานะปัจจุบัน:", font=CTkFont(size=12)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.lbl_status = CTkLabel(self.detail_frame, text="-", font=CTkFont(weight="bold"))
        self.lbl_status.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        # ส่วนเลือกเจ้าของใหม่
        CTkLabel(self.detail_frame, text="เลือกเจ้าของใหม่:", font=CTkFont(size=14, weight="bold")).grid(row=3, column=0, padx=10, pady=(20,5), sticky="w")
        
        self.new_sale_var = tk.StringVar(value="เลือกเซลส์...")
        self.sale_dropdown = CTkOptionMenu(self.detail_frame, variable=self.new_sale_var, values=[])
        self.sale_dropdown.grid(row=3, column=1, padx=10, pady=(20,5), sticky="ew")

        # --- ส่วนที่ 3: ปุ่มบันทึก ---
        self.action_frame = CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        self.action_frame.grid_columnconfigure(0, weight=1)
        
        self.save_btn = CTkButton(self.action_frame, text="บันทึกการย้าย (Reassign)", fg_color="#16A34A", hover_color="#15803D", command=self._save_reassignment, height=40, font=CTkFont(size=16, weight="bold"))
        self.save_btn.pack(fill="x")
        
        self.action_frame.grid_remove() # ซ่อนไว้ก่อน

        self.current_so_data = None
        self._load_sales_list()
        
        self.transient(master)
        self.grab_set()

    def _load_sales_list(self):
        """ดึงรายชื่อเซลส์ทั้งหมดมาใส่ Dropdown"""
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = 'Sale' AND status = 'Active' ORDER BY sale_key"
            df = pd.read_sql_query(query, self.app_container.pg_engine)
            
            if not df.empty:
                # สร้าง list แบบ "SALE01 : ชื่อเซลส์"
                sale_options = [f"{row['sale_key']} : {row['sale_name']}" for _, row in df.iterrows()]
                self.sale_dropdown.configure(values=sale_options)
        except Exception as e:
            print(f"Error loading sales list: {e}")

    def _search_so(self):
        so_num = self.so_entry.get().strip().upper()
        if not so_num: return

        try:
            # ดึงข้อมูล SO พร้อมชื่อเซลส์ปัจจุบัน (ดึง status มาเช็คด้วย)
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.status, c.sale_key, u.sale_name
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.so_number = %s AND c.is_active = 1
                LIMIT 1
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(so_num,))

            if df.empty:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบ SO หมายเลข: {so_num} ในระบบ", parent=self)
                self.detail_frame.grid_remove()
                self.action_frame.grid_remove()
                self.current_so_data = None
                return

            # พบข้อมูล
            row = df.iloc[0]
            self.current_so_data = row
            
            self.lbl_customer.configure(text=row['customer_name'])
            sale_display = f"{row['sale_key']} ({row['sale_name']})" if row['sale_name'] else row['sale_key']
            self.lbl_current_sale.configure(text=sale_display)
            
            # แสดงสถานะและเปลี่ยนสีถ้าสถานะซีเรียส
            status = row['status']
            self.lbl_status.configure(text=status)
            if status in ['Paid', 'HR Verified', 'Cancelled']:
                self.lbl_status.configure(text_color="#B91C1C") # สีแดงเข้ม
            else:
                self.lbl_status.configure(text_color="black")

            # แสดง Frame
            self.detail_frame.grid()
            self.action_frame.grid()

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)

    def _save_reassignment(self):
        if self.current_so_data is None: return
        
        selected_str = self.new_sale_var.get()
        if not selected_str or "เลือก" in selected_str:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกเจ้าของใหม่ (Salesperson)", parent=self)
            return

        # แยกเอาเฉพาะ Sale Key (เช่น "SALE01 : Somchai" -> "SALE01")
        new_sale_key = selected_str.split(":")[0].strip()
        so_number = self.current_so_data['so_number']
        old_sale_key = self.current_so_data['sale_key']
        so_id = int(self.current_so_data['id'])
        current_status = self.current_so_data['status']

        if new_sale_key == old_sale_key:
            messagebox.showwarning("แจ้งเตือน", "คุณเลือกเจ้าของเดิม ไม่มีการเปลี่ยนแปลง", parent=self)
            return

        # --- START: Logic การแจ้งเตือนตามสถานะ ---
        warning_msg = ""
        # ถ้าสถานะเป็น Paid หรือ HR Verified ให้เตือนแรงหน่อย เพราะอาจกระทบยอดเงิน
        if current_status in ['Paid', 'HR Verified']:
            warning_msg = (f"\n⚠️ คำเตือน: SO นี้อยู่ในสถานะ '{current_status}' แล้ว\n"
                           "การย้ายเจ้าของอาจส่งผลต่อการคำนวณค่าคอมมิชชั่นที่ทำไปแล้ว\n"
                           "กรุณาตรวจสอบให้แน่ใจก่อนดำเนินการ")
        # ถ้าส่งไปแล้ว (Pending...) ก็เตือนปกติ
        elif current_status not in ['Original', 'Edited', 'Draft', 'Rejected by SM', 'Rejected by HR']:
            warning_msg = (f"\nℹ️ หมายเหตุ: SO นี้ถูกส่งเข้าระบบแล้ว (สถานะ: {current_status})\n"
                           "ระบบจะทำการเปลี่ยนชื่อเจ้าของให้ทันที")

        if not messagebox.askyesno("ยืนยันการย้าย", f"ยืนยันย้าย SO: {so_number}\n\nจาก: {old_sale_key}\nไปเป็น: {new_sale_key}\n{warning_msg}", icon="warning" if warning_msg else "question"):
            return
        # --- END ---

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 1. อัปเดตเจ้าของในตาราง commissions (ทำได้เลย ไม่ติด WHERE status)
                cursor.execute("""
                    UPDATE commissions 
                    SET sale_key = %s 
                    WHERE id = %s
                """, (new_sale_key, so_id))

                # 2. บันทึก Log
                log_detail = {
                    "message": "Reassign Owner by Sale Support",
                    "from_sale": old_sale_key,
                    "to_sale": new_sale_key,
                    "so_number": so_number,
                    "original_status": current_status
                }
                cursor.execute("""
                    INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, ('Reassign SO', 'commissions', so_id, self.app_container.current_user_key, json.dumps(log_detail), datetime.now()))
                
                # 3. แจ้งเตือน Sale คนใหม่
                msg = f"SO: {so_number} ถูกย้ายมาเป็นของคุณ โดย Sale Support (สถานะปัจจุบัน: {current_status})"
                cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read) VALUES (%s, %s, FALSE)", (new_sale_key, msg))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ย้าย SO: {so_number} ไปยัง {new_sale_key} เรียบร้อยแล้ว", parent=self)
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึก: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)


class CancelledHistoryWindow(CTkToplevel): # <-- แก้จาก ctk.CTkToplevel เป็น CTkToplevel
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.title("ประวัติ SO ที่ถูกยกเลิก (Cancelled History)")
        self.geometry("1100x600")
        
        self.user_role = self.app_container.current_user_role
        self.user_key = self.app_container.current_user_key

        # Header
        # แก้จาก ctk.CTkLabel เป็น CTkLabel และ ctk.CTkFont เป็น CTkFont
        CTkLabel(self, text="รายการ SO ที่ถูกยกเลิก (ไม่นำมาคิดค่าคอมมิชชั่น)", font=CTkFont(size=20, weight="bold")).pack(pady=10)
        
        # Filter Frame
        filter_frame = CTkFrame(self)
        filter_frame.pack(fill="x", padx=10, pady=5)
        
        self.search_entry = CTkEntry(filter_frame, placeholder_text="ค้นหา SO Number...")
        self.search_entry.pack(side="left", padx=10)
        
        CTkButton(filter_frame, text="ค้นหา", command=self.load_data).pack(side="left")
        CTkButton(filter_frame, text="Refresh", command=self.load_data, fg_color="gray").pack(side="left", padx=5)

        # Table Frame
        self.tree_frame = CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.create_table()
        self.load_data()

    def create_table(self):
        columns = ["SO Number", "Sale Key", "Customer", "สาเหตุการยกเลิก", "วันที่ยกเลิก", "สถานะ"]
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", height=20)
        
        # Config columns
        self.tree.heading("SO Number", text="SO Number")
        self.tree.column("SO Number", width=120)
        
        self.tree.heading("Sale Key", text="พนักงานขาย")
        self.tree.column("Sale Key", width=100)
        
        self.tree.heading("Customer", text="ลูกค้า")
        self.tree.column("Customer", width=200)
        
        self.tree.heading("สาเหตุการยกเลิก", text="สาเหตุการยกเลิก")
        self.tree.column("สาเหตุการยกเลิก", width=300)

        self.tree.heading("วันที่ยกเลิก", text="วันที่อัปเดต")
        self.tree.column("วันที่ยกเลิก", width=150)

        self.tree.heading("สถานะ", text="สถานะ")
        self.tree.column("สถานะ", width=100)

        # Scrollbar
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def load_data(self):
        # Clear table
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        search_txt = self.search_entry.get().strip()
        
        # Base Query
        query = """
            SELECT so_number, sale_key, customer_name, rejection_reason, timestamp, status 
            FROM commissions 
            WHERE status = 'Cancelled'
        """
        params = []

        # Role Based Filtering
        if self.user_role == 'Sale':
            query += " AND sale_key = %s"
            params.append(self.user_key)
        
        # Search Logic
        if search_txt:
            query += " AND so_number ILIKE %s"
            params.append(f"%{search_txt}%")
            
        query += " ORDER BY timestamp DESC"

        try:
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=tuple(params))
            
            if df.empty:
                return

            # +++ [แก้ไขตรงนี้] แปลงข้อมูล timestamp ให้เป็น datetime object ให้ชัวร์ก่อน +++
            df['timestamp'] = pd.to_datetime(df['timestamp']) 
            # +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            for _, row in df.iterrows():
                # Format Date
                ts = row['timestamp']
                
                # ตรวจสอบว่าเป็นวันเวลาที่ถูกต้องหรือไม่
                if pd.notna(ts):
                    try:
                        date_str = ts.strftime('%d/%m/%Y %H:%M')
                    except:
                        date_str = str(ts) # ถ้าแปลงไม่ได้จริงๆ ให้แสดงตามเดิม
                else:
                    date_str = "-"
                
                values = (
                    row['so_number'],
                    row['sale_key'],
                    row['customer_name'],
                    row['rejection_reason'], 
                    date_str,
                    row['status']
                )
                self.tree.insert("", "end", values=values)
                
        except Exception as e:
            print(f"Error loading cancelled history: {e}")
            import traceback
            traceback.print_exc()

class TransportEditDialog(CTkToplevel):
    def __init__(self, master, app_container, po_id, on_save_callback=None):
        super().__init__(master)
        self.app_container = app_container
        self.po_id = po_id
        self.on_save_callback = on_save_callback
        self.old_data = {} 

        # ตัวแปรสำหรับ Radio Button
        self.cut_vat_var = tk.StringVar(value="VAT")
        self.cut_wht_var = tk.StringVar(value="No")

        self.title("แก้ไขข้อมูลขนส่ง/ค่าตัด (Special Edit)")
        self.geometry("700x800")
        
        self.main_frame = CTkScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame.grid_columnconfigure(1, weight=1)

        self._create_ui()
        self.after(50, self._load_current_data)
        
        self.transient(master)
        self.grab_set()

    def _create_ui(self):
        r = 0
        CTkLabel(self.main_frame, text="แก้ไขค่าขนส่ง/ค่าตัด (Backdoor)", font=CTkFont(size=18, weight="bold"), text_color="#B91C1C").grid(row=r, column=0, columnspan=2, pady=(10, 5)); r+=1
        
        # --- 1. Stock ---
        self._add_header(r, "[1] ค่าย้าย / เข้า Stock"); r+=1
        self.stock_cost = self._add_row(r, "ยอดเงิน (ค่าย้าย):", is_numeric=True); r+=1
        self.stock_driver = self._add_row(r, "คนขับ/ขนส่ง:"); r+=1
        self.stock_plate = self._add_row(r, "ทะเบียนรถ:"); r+=1
        self.stock_note = self._add_row(r, "หมายเหตุ:"); r+=1

        # --- 2. Site ---
        self._add_header(r, "[2] ค่ารถ / ส่ง Site (ลูกค้า)"); r+=1
        self.site_cost = self._add_row(r, "ยอดเงิน (ค่ารถ):", is_numeric=True); r+=1
        self.site_driver = self._add_row(r, "คนขับ/ขนส่ง:"); r+=1
        self.site_plate = self._add_row(r, "ทะเบียนรถ:"); r+=1
        self.site_note = self._add_row(r, "หมายเหตุ:"); r+=1

        # --- 3. Cutting ---
        self._add_header(r, "[3] ค่าบริการตัด/เจาะ (Cutting)"); r+=1
        self.cutting_cost = self._add_row(r, "ยอดเงิน (ค่าตัด):", is_numeric=True); 
        self.cutting_cost.bind("<KeyRelease>", self._update_cutting_summary)
        r+=1

        # ตัวเลือก VAT / WHT
        opt_frame = CTkFrame(self.main_frame, fg_color="transparent")
        opt_frame.grid(row=r, column=1, sticky="w", padx=10, pady=2)
        CTkLabel(opt_frame, text="ประเภท:").pack(side="left", padx=(0,5))
        CTkRadioButton(opt_frame, text="VAT", variable=self.cut_vat_var, value="VAT", command=self._update_cutting_summary).pack(side="left", padx=5)
        CTkRadioButton(opt_frame, text="CASH", variable=self.cut_vat_var, value="CASH", command=self._update_cutting_summary).pack(side="left", padx=5)
        CTkLabel(opt_frame, text="|", text_color="gray").pack(side="left", padx=10)
        CTkLabel(opt_frame, text="หัก ณ ที่จ่าย:").pack(side="left", padx=(0,5))
        CTkRadioButton(opt_frame, text="ไม่หัก", variable=self.cut_wht_var, value="No", command=self._update_cutting_summary).pack(side="left", padx=5)
        CTkRadioButton(opt_frame, text="1%", variable=self.cut_wht_var, value="1%", command=self._update_cutting_summary).pack(side="left", padx=5)
        CTkRadioButton(opt_frame, text="3%", variable=self.cut_wht_var, value="3%", command=self._update_cutting_summary).pack(side="left", padx=5)
        r+=1

        # สรุปการคำนวณ
        summary_frame = CTkFrame(self.main_frame, fg_color="#F3F4F6", corner_radius=6)
        summary_frame.grid(row=r, column=1, sticky="ew", padx=10, pady=5)
        self.lbl_vat_amt = CTkLabel(summary_frame, text="VAT: 0.00", font=CTkFont(size=12), text_color="gray")
        self.lbl_vat_amt.pack(side="left", padx=10, pady=5)
        self.lbl_wht_amt = CTkLabel(summary_frame, text="WHT: 0.00", font=CTkFont(size=12), text_color="red")
        self.lbl_wht_amt.pack(side="left", padx=10, pady=5)
        self.lbl_net_amt = CTkLabel(summary_frame, text="สุทธิ: 0.00", font=CTkFont(size=14, weight="bold"), text_color="green")
        self.lbl_net_amt.pack(side="right", padx=10, pady=5)
        r+=1

        self.cutting_note = self._add_row(r, "หมายเหตุ (ตัด):"); r+=1

        # ======================================================================
        # ปุ่ม Action (พิมพ์ และ บันทึก)
        # ======================================================================
        btn_frame = CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.grid(row=r, column=0, columnspan=2, pady=30, sticky="ew")
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        # ปุ่มพิมพ์ (เรียก Wrapper เพื่อดึงข้อมูลใหม่เสมอ)
        CTkButton(btn_frame, 
                  text="🖨️ พิมพ์ใบสรุป (PDF)", 
                  fg_color="#3B82F6", hover_color="#2563EB", 
                  height=45, font=CTkFont(size=16, weight="bold"), 
                  command=self._print_action_safe
        ).grid(row=0, column=0, padx=5, sticky="ew")

        # ปุ่มบันทึก
        CTkButton(btn_frame, 
                  text="บันทึกการแก้ไข", 
                  fg_color="#16A34A", hover_color="#15803D", 
                  height=45, font=CTkFont(size=16, weight="bold"), 
                  command=self._save_transport_changes
        ).grid(row=0, column=1, padx=5, sticky="ew")

    def _add_header(self, row, text):
        CTkLabel(self.main_frame, text=text, font=CTkFont(size=14, weight="bold"), fg_color="gray90", corner_radius=6, text_color="black").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5), padx=5)

    def _add_row(self, row, label, is_numeric=False):
        CTkLabel(self.main_frame, text=label).grid(row=row, column=0, padx=10, pady=5, sticky="e")
        if is_numeric:
            entry = NumericEntry(self.main_frame)
        else:
            entry = CTkEntry(self.main_frame)
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        return entry

    def _update_cutting_summary(self, *args):
        try:
            base_cost = utils.convert_to_float(self.cutting_cost.get())
            vat_amt = base_cost * 0.07 if self.cut_vat_var.get() == "VAT" else 0.0
            wht_rate = 0.01 if self.cut_wht_var.get() == "1%" else (0.03 if self.cut_wht_var.get() == "3%" else 0.0)
            wht_amt = base_cost * wht_rate
            net = (base_cost + vat_amt) - wht_amt
            self.lbl_vat_amt.configure(text=f"VAT: {vat_amt:,.2f}")
            self.lbl_wht_amt.configure(text=f"หัก: {wht_amt:,.2f}")
            self.lbl_net_amt.configure(text=f"สุทธิ: {net:,.2f}")
        except: pass

    def _load_current_data(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (self.po_id,))
                data = cursor.fetchone()
                if data:
                    self.old_data = dict(data)
                    utils.set_entry_text(self.stock_cost, data.get('shipping_to_stock_cost', 0))
                    utils.set_entry_text(self.stock_driver, data.get('shipping_to_stock_driver', ''))
                    utils.set_entry_text(self.stock_plate, data.get('shipping_to_stock_plate', ''))
                    utils.set_entry_text(self.stock_note, data.get('shipping_to_stock_notes', ''))
                    
                    utils.set_entry_text(self.site_cost, data.get('shipping_to_site_cost', 0))
                    utils.set_entry_text(self.site_driver, data.get('shipping_to_site_driver', ''))
                    utils.set_entry_text(self.site_plate, data.get('shipping_to_site_plate', ''))
                    utils.set_entry_text(self.site_note, data.get('shipping_to_site_notes', ''))
                    
                    utils.set_entry_text(self.cutting_cost, data.get('cutting_cost', 0))
                    utils.set_entry_text(self.cutting_note, data.get('cutting_remark', ''))
                    
                    self.cut_vat_var.set(data.get('cutting_vat_type', 'VAT'))
                    wht_val = data.get('cutting_wht_type', 'No')
                    self.cut_wht_var.set(wht_val if wht_val in ['No','1%','3%'] else 'No')
                    self._update_cutting_summary()
        except Exception as e: messagebox.showerror("Error", f"Load failed: {e}")
        finally: self.app_container.release_connection(conn)

    # ในคลาส TransportEditDialog
    def _save_transport_changes(self):
        """บันทึกข้อมูลและคำนวณส่วนต่างเพื่ออัปเดต Grand Total"""
        
        # 1. รับค่าใหม่จาก GUI (ใช้ .get() และ .strip() ให้ชัวร์)
        new_stock_cost = utils.convert_to_float(self.stock_cost.get())
        new_site_cost = utils.convert_to_float(self.site_cost.get())
        new_cut_cost = utils.convert_to_float(self.cutting_cost.get())
        
        # [สำคัญ] ดึงข้อมูล Text ให้ครบ
        stock_drv = str(self.stock_driver.get()).strip()
        stock_plt = str(self.stock_plate.get()).strip()
        stock_nte = str(self.stock_note.get()).strip()
        
        site_drv = str(self.site_driver.get()).strip()
        site_plt = str(self.site_plate.get()).strip()
        site_nte = str(self.site_note.get()).strip()
        
        cut_rem = str(self.cutting_note.get()).strip()
        
        # 2. เตรียม Log และข้อมูลเดิม
        user = self.app_container.current_user_key
        po_num = self.old_data.get('po_number', '-')

        # คำนวณค่าเดิม (Old Values)
        old_stock_cost = float(self.old_data.get('shipping_to_stock_cost') or 0)
        old_site_cost = float(self.old_data.get('shipping_to_site_cost') or 0)
        old_cut_cost = float(self.old_data.get('cutting_cost') or 0)

        # [🔥 จุดสำคัญ 1] ดึงประเภทการจัดส่งจากข้อมูลเดิม เพื่อใช้ตรวจสอบเงื่อนไข
        delivery_type = self.old_data.get('delivery_type', '')

        if not messagebox.askyesno("ยืนยัน", f"ยืนยันบันทึกการแก้ไข PO: {po_num}?"): return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ฟังก์ชันคำนวณยอดสุทธิ (Net Logic: Base + VAT - WHT)
                def calc_net(base, vat, wht):
                    v = base * 0.07 if vat == 'VAT' else 0
                    # แปลง WHT text เป็นตัวเลข
                    w_rate = 0.0
                    if wht == '1%': w_rate = 0.01
                    elif wht == '3%': w_rate = 0.03
                    return (base + v) - (base * w_rate)

                # ดึง Vat/Wht เดิม
                o_stock_v = self.old_data.get('shipping_to_stock_vat_type')
                o_stock_w = self.old_data.get('shipping_to_stock_wht_type')
                o_site_v = self.old_data.get('shipping_to_site_vat_type')
                o_site_w = self.old_data.get('shipping_to_site_wht_type')
                o_cut_v = self.old_data.get('cutting_vat_type')
                o_cut_w = self.old_data.get('cutting_wht_type')

                # ค่า Vat/Wht ใหม่ (สำหรับ Cutting ที่แก้ได้ในหน้านี้)
                n_cut_v = self.cut_vat_var.get()
                n_cut_w = self.cut_wht_var.get()

                # คำนวณ Net เดิม vs ใหม่
                # (สมมติ Vat/Wht ขนส่งไม่เปลี่ยนในหน้านี้ ใช้ค่าเดิม o_stock_v/w)
                old_net_stock = calc_net(old_stock_cost, o_stock_v, o_stock_w)
                new_net_stock = calc_net(new_stock_cost, o_stock_v, o_stock_w)
                
                old_net_site = calc_net(old_site_cost, o_site_v, o_site_w)
                new_net_site = calc_net(new_site_cost, o_site_v, o_site_w)
                
                old_net_cut = calc_net(old_cut_cost, o_cut_v, o_cut_w)
                new_net_cut = calc_net(new_cut_cost, n_cut_v, n_cut_w)

                # =================================================================
                # [🔥 จุดสำคัญ 2] คำนวณส่วนต่างที่จะไปกระทบ Grand Total (ยอดจ่ายซัพ)
                # =================================================================
                
                transport_diff = 0.0
                
                # ถ้าเป็น "ซัพพลายเออร์จัดส่ง" -> ส่วนต่างค่ารถจะไปกระทบยอดจ่ายรวม
                if delivery_type == "ซัพพลายเออร์จัดส่ง":
                    transport_diff = (new_net_stock - old_net_stock) + (new_net_site - old_net_site)
                else:
                    # ถ้าไม่ใช่ซัพส่ง (เช่น รับเอง, เอกชน) -> ค่ารถไม่เกี่ยวกับยอดจ่ายซัพ
                    transport_diff = 0.0

                # ค่าตัด (Cutting) กระทบ Grand Total เสมอ (ตาม Logic เดิม)
                cutting_diff = (new_net_cut - old_net_cut)

                # ส่วนต่างสุทธิที่จะไปบวก/ลบ ใน Grand Total
                diff_grand_total = transport_diff + cutting_diff
                
                # ส่วนต่าง Total Cost (เฉพาะค่าของ+ค่าบริการ ไม่เกี่ยวกับว่าใครจ่าย)
                diff_total_cost = (new_cut_cost - old_cut_cost) 

                # 5. Update SQL (อัปเดตทุกฟิลด์ให้ครบ)
                cursor.execute("""
                    UPDATE purchase_orders SET 
                        shipping_to_stock_cost = %s, 
                        shipping_to_stock_driver = %s, 
                        shipping_to_stock_plate = %s, 
                        shipping_to_stock_notes = %s,
                        
                        shipping_to_site_cost = %s, 
                        shipping_to_site_driver = %s, 
                        shipping_to_site_plate = %s, 
                        shipping_to_site_notes = %s,
                        
                        cutting_cost = %s, 
                        cutting_remark = %s, 
                        cutting_vat_type = %s, 
                        cutting_wht_type = %s,
                        
                        total_cost = COALESCE(total_cost, 0) + %s, 
                        grand_total = COALESCE(grand_total, 0) + %s
                    WHERE id = %s
                """, (
                    new_stock_cost, stock_drv, stock_plt, stock_nte,
                    new_site_cost, site_drv, site_plt, site_nte,
                    new_cut_cost, cut_rem, n_cut_v, n_cut_w,
                    diff_total_cost, diff_grand_total, self.po_id
                ))

                # Log
                cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp) VALUES (%s, %s, %s, %s, %s, NOW())", 
                               ('Edit PO Transport', 'purchase_orders', self.po_id, user, f"Updated Transport/Cutting (Diff GT: {diff_grand_total})"))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลเรียบร้อยแล้ว", parent=self)
            
            if self.on_save_callback: self.on_save_callback()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"{e}")
            print(traceback.format_exc())
        finally:
            if conn: self.app_container.release_connection(conn)

    def _print_action_safe(self):
        """ปุ่มพิมพ์ในหน้าแก้ไข"""
        # บังคับถามก่อนพิมพ์ เพื่อเตือนสติให้ Save
        if messagebox.askyesno("พิมพ์เอกสาร", "ระบบจะพิมพ์ข้อมูลล่าสุดจากฐานข้อมูล\nหากเพิ่งแก้ไข กรุณากด 'บันทึก' ก่อน\n\nต้องการพิมพ์เลยหรือไม่?"):
            try:
                # เรียก Wrapper ที่อยู่ไฟล์เดียวกัน
                print_transport_pdf_wrapper(self.app_container, self.po_id)
            except Exception as e:
                messagebox.showerror("Error", f"Print Failed: {e}")

class TransportLogViewer(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.title("ประวัติการแก้ไขค่าขนส่ง (Transport Logs)")
        self.geometry("1000x600")

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        filter_frame = CTkFrame(self)
        filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        
        self.month_var = tk.StringVar(value="ทุกเดือน")
        month_menu = CTkOptionMenu(filter_frame, variable=self.month_var, values=["ทุกเดือน"] + self.thai_months, command=self.load_logs)
        month_menu.pack(side="left", padx=5)

        self.year_var = tk.StringVar(value=str(datetime.now().year))
        year_menu = CTkOptionMenu(filter_frame, variable=self.year_var, values=[str(y) for y in range(2024, 2030)], command=self.load_logs)
        year_menu.pack(side="left", padx=5)

        self.search_entry = CTkEntry(filter_frame, placeholder_text="ค้นหา PO / User...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)
        CTkButton(filter_frame, text="ค้นหา", command=self.load_logs).pack(side="left", padx=5)

        self.tree_frame = CTkFrame(self)
        self.tree_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        self._create_treeview()
        CTkLabel(self, text="* ดับเบิลคลิกที่รายการเพื่อดูรายละเอียดทั้งหมด", text_color="gray50", font=CTkFont(size=12)).grid(row=2, column=0, pady=5)
        self.after(100, self.load_logs)
        self.transient(master)
        self.grab_set()

    def _create_treeview(self):
        columns = ("timestamp", "user", "po_number", "details")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings")
        self.tree.heading("timestamp", text="เวลาแก้ไข"); self.tree.column("timestamp", width=150, anchor="center")
        self.tree.heading("user", text="ผู้แก้ไข"); self.tree.column("user", width=100, anchor="center")
        self.tree.heading("po_number", text="เลขที่ PO"); self.tree.column("po_number", width=120, anchor="center")
        self.tree.heading("details", text="รายละเอียดการเปลี่ยนแปลง"); self.tree.column("details", width=600, anchor="w")

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self._on_double_click)

    def _on_double_click(self, event):
        item_id = self.tree.focus()
        if not item_id: return
        item = self.tree.item(item_id)
        formatted_text = str(item['values'][3]).replace(" | ", "\n")
        
        popup = CTkToplevel(self)
        popup.title("รายละเอียด Log")
        popup.geometry("500x400")
        textbox = CTkTextbox(popup, font=CTkFont(size=14))
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        textbox.insert("1.0", formatted_text)
        textbox.configure(state="disabled")
        CTkButton(popup, text="ปิด", command=popup.destroy).pack(pady=10)
        popup.transient(self); popup.grab_set()

    def load_logs(self, *args):
        for item in self.tree.get_children(): self.tree.delete(item)
        month_str, year_str = self.month_var.get(), self.year_var.get()
        search_txt = self.search_entry.get().strip().lower()

        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                query = """
                    SELECT a.timestamp, a.user_info, p.po_number, a.changes
                    FROM audit_log a
                    LEFT JOIN purchase_orders p ON a.record_id = p.id
                    WHERE a.action = 'Edit PO Transport'
                """
                params = []
                if month_str != "ทุกเดือน":
                    query += " AND EXTRACT(MONTH FROM a.timestamp::timestamp) = %s"
                    params.append(self.thai_month_map[month_str])
                if year_str:
                    query += " AND EXTRACT(YEAR FROM a.timestamp::timestamp) = %s"
                    params.append(int(year_str))
                query += " ORDER BY a.timestamp::timestamp DESC"

                cursor.execute(query, tuple(params))
                rows = cursor.fetchall()
                for row in rows:
                    po_num, user, changes = row['po_number'] or "N/A", row['user_info'], row['changes']
                    if search_txt and search_txt not in str(po_num).lower() and search_txt not in str(user).lower() and search_txt not in str(changes).lower(): continue
                    
                    try: ts_str = pd.to_datetime(row['timestamp']).strftime("%d/%m/%Y %H:%M")
                    except: ts_str = str(row['timestamp'])
                    
                    self.tree.insert("", "end", values=(ts_str, user, po_num, str(changes).replace('\n', ' | ')))
        except Exception as e: messagebox.showerror("Error", f"Load logs failed: {e}")
        finally: self.app_container.release_connection(conn)

class TransportPOSearchDialog(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.title("ค้นหา PO เพื่อแก้ไขค่าขนส่ง (Transport Manager)")
        self.geometry("1000x600")
        
        # Grid Configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 1. Header & Search
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        CTkLabel(top_frame, text="เลือก PO ของคุณเพื่อแก้ไขค่ารถ", font=CTkFont(size=16, weight="bold")).pack(side="left")
        
        self.search_entry = CTkEntry(top_frame, placeholder_text="ค้นหา PO / SO / Supplier...", width=250)
        self.search_entry.pack(side="right", padx=5)
        self.search_entry.bind("<Return>", lambda e: self._load_data())
        
        CTkButton(top_frame, text="ค้นหา", width=80, command=self._load_data).pack(side="right")

        # 2. Table List
        self.tree_frame = CTkFrame(self)
        self.tree_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)

        self._create_treeview()
        
        # 3. Action Button
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=20)
        
        # ปุ่มดู Log
        CTkButton(btn_frame, text="ดูประวัติการแก้ไข (Logs)", fg_color="#64748B", hover_color="#475569", command=self._open_log_viewer).pack(side="left")
        
        # ปุ่มแก้ไข (Enable เมื่อเลือกรายการ)
        self.btn_edit = CTkButton(btn_frame, text="🚚 แก้ไขค่ารถรายการที่เลือก", 
                                  fg_color="#8B5CF6", hover_color="#7C3AED", # สีม่วง
                                  font=CTkFont(size=16, weight="bold"),
                                  state="disabled",
                                  command=self._open_edit_dialog)
        self.btn_edit.pack(side="right")

        self.print_transport_btn = CTkButton(
            btn_frame, # ตรวจสอบชื่อตัวแปรนี้ให้ตรงกับของคุณ (btn_frame หรือ action_frame)
            text="🖨️ ใบค่ารถ (Transport)", 
            fg_color="#059669", hover_color="#047857", # สีเขียว
            state="disabled", 
            width=180, height=40,
            command=self._print_transport_action # <--- ต้องเรียกฟังก์ชันนี้
        )
        self.print_transport_btn.pack(side="right", padx=10)

        # Initial Load
        self.after(100, self._load_data)
        self.transient(master)
        self.grab_set()

    def _print_transport_action(self):
        """เมื่อกดปุ่มใบค่ารถ"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("เตือน", "กรุณาเลือกรายการก่อนครับ")
            return
        
        # ดึง ID ของ PO ที่เลือก
        item = self.tree.item(selected[0])
        po_id = item['values'][0] 
        
        print(f">>> User Clicked Print for PO ID: {po_id}")

        # เรียก Wrapper เพื่อดึงข้อมูลสดแล้วพิมพ์
        print_transport_pdf_wrapper(self.app_container, po_id)
 
    def _create_treeview(self):
        columns = ("id", "po_number", "so_number", "supplier", "status", "transport_cost")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings")
        
        self.tree.heading("id", text="ID")
        self.tree.heading("po_number", text="PO Number")
        self.tree.heading("so_number", text="SO Number")
        self.tree.heading("supplier", text="Supplier")
        self.tree.heading("status", text="สถานะ")
        self.tree.heading("transport_cost", text="ค่ารถปัจจุบัน (Stock+Site)")

        self.tree.column("id", width=0, stretch=False) # ซ่อน ID
        self.tree.column("po_number", width=120, anchor="center")
        self.tree.column("so_number", width=120, anchor="center")
        self.tree.column("supplier", width=250, anchor="w")
        self.tree.column("status", width=100, anchor="center")
        self.tree.column("transport_cost", width=150, anchor="e")

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True)
        
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda e: self._open_edit_dialog())

    def _load_data(self):
        # Clear
        for item in self.tree.get_children(): self.tree.delete(item)
        self.btn_edit.configure(state="disabled")

        search_txt = self.search_entry.get().strip().lower()
        user_key = self.app_container.current_user_key

        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # Query: ดึง PO ของ User นี้ (และไม่เอาที่ Cancelled)
                query = """
                    SELECT id, po_number, so_number, supplier_name, status, 
                           (COALESCE(shipping_to_stock_cost,0) + COALESCE(shipping_to_site_cost,0)) as total_transport
                    FROM purchase_orders 
                    WHERE user_key = %s 
                    AND status != 'Cancelled'
                """
                params = [user_key]

                if search_txt:
                    query += " AND (LOWER(po_number) LIKE %s OR LOWER(so_number) LIKE %s OR LOWER(supplier_name) LIKE %s)"
                    wildcard = f"%{search_txt}%"
                    params.extend([wildcard, wildcard, wildcard])
                
                query += " ORDER BY id DESC"
                
                cursor.execute(query, tuple(params))
                rows = cursor.fetchall()

                for row in rows:
                    vals = (
                        row['id'],
                        row['po_number'],
                        row['so_number'],
                        row['supplier_name'],
                        row['status'],
                        f"{row['total_transport']:,.2f}"
                    )
                    self.tree.insert("", "end", values=vals)

        except Exception as e:
            messagebox.showerror("Error", f"{e}")
        finally:
            self.app_container.release_connection(conn)

    def _on_select(self, event):
        if self.tree.selection():
            self.btn_edit.configure(state="normal")
            self.btn_print.configure(state="normal") 
        else:
            self.btn_edit.configure(state="disabled")
            self.btn_print.configure(state="disabled")

    def _open_edit_dialog(self):
        selected = self.tree.selection()
        if not selected: return
        
        item = self.tree.item(selected[0])
        po_id = item['values'][0] # ID hidden at index 0
        
        # เรียกใช้ TransportEditDialog (ต้องมี Class นี้อยู่ในไฟล์เดียวกันแล้ว)
        try:
            TransportEditDialog(self, self.app_container, po_id, on_save_callback=self._load_data)
        except NameError:
             messagebox.showerror("Error", "ไม่พบ Class TransportEditDialog ในไฟล์ history_windows.py")

    def _print_selected_action(self):
        """ดึง ID รายการที่เลือกแล้วสั่งพิมพ์"""
        selected = self.tree.selection()
        if not selected: return
        
        item = self.tree.item(selected[0])
        po_id = item['values'][0] # ดึง ID จากคอลัมน์แรกที่ซ่อนอยู่
        
        # เรียก Wrapper เพื่อดึงข้อมูลสดๆ จาก DB แล้วพิมพ์
        # (ต้องมั่นใจว่า print_transport_pdf_wrapper ถูก import หรือประกาศไว้ด้านบนไฟล์แล้ว)
        try:
            print_transport_pdf_wrapper(self.app_container, po_id)
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการสั่งพิมพ์: {e}")

    def _open_log_viewer(self):
        # เรียกใช้ TransportLogViewer (ต้องมี Class นี้อยู่ในไฟล์เดียวกันแล้ว)
        try:
            TransportLogViewer(self, self.app_container)
        except NameError:
             messagebox.showerror("Error", "ไม่พบ Class TransportLogViewer ในไฟล์ history_windows.py")