# history_windows.py (ฉบับแก้ไขล่าสุด)

import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkToplevel, CTkTextbox, CTkScrollableFrame, CTkLabel, CTkFont, CTkFrame, CTkButton, CTkEntry, CTkRadioButton, CTkOptionMenu, CTkTabview)
from tkinter import messagebox
import json
import pandas as pd
from datetime import datetime
import traceback
from custom_widgets import NumericEntry, DateSelector , AutoCompleteEntry
import utils
import psycopg2.errors
import psycopg2.extras
import numpy as np

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
        
        # --- เปลี่ยนจากการ pack frame มาใช้ grid เพื่อสร้างตาราง ---
        self.main_frame.grid_columnconfigure(1, weight=1)

        sections = {
            "รายละเอียดการขาย": ['bill_date', 'customer_id', 'customer_name', 'so_number', 'credit_term', 'commission_month', 'commission_year'],
            "ยอดขายและบริการ": ['sales_service_amount', 'cutting_drilling_fee', 'other_service_fee'],
            "ค่าจัดส่ง": ['shipping_cost', 'delivery_date', 'relocation_cost'],
            "ค่าธรรมเนียมและส่วนลด": ['credit_card_fee', 'transfer_fee', 'wht_3_percent', 'brokerage_fee', 'giveaways', 'coupons'],
            "รายละเอียดการชำระเงิน": ['total_payment_amount', 'payment_date', 'payment_before_vat', 'payment_no_vat'],
            "ข้อมูลที่คำนวณแล้ว (Final)": ['final_sales_amount', 'final_cost_amount', 'final_gp', 'final_margin']
        }
        
        header_map = self.app_container.HEADER_MAP
        normal_font = CTkFont(size=14)
        current_row = 0

        for title, columns in sections.items():
            # Section Title
            title_label = CTkLabel(self.main_frame, text=title, font=CTkFont(size=16, weight="bold"), anchor="w")
            title_label.grid(row=current_row, column=0, columnspan=2, padx=10, pady=(15, 5), sticky="w")
            current_row += 1
            
            # Table Rows
            for col in columns:
                if col in self.so_data:
                    display_name = header_map.get(col, col)
                    value = self.so_data[col]
                    
                    if isinstance(value, (int, float)): 
                        value_text = f"{value:,.2f}"
                    elif isinstance(value, datetime): 
                        value_text = value.strftime('%d/%m/%Y')
                    else: 
                        value_text = str(value) if value is not None else "-"
                    
                    CTkLabel(self.main_frame, text=display_name, font=normal_font, anchor="w").grid(row=current_row, column=0, padx=20, pady=3, sticky="w")
                    CTkLabel(self.main_frame, text=value_text, font=normal_font, wraplength=400, justify="left", anchor="w").grid(row=current_row, column=1, padx=10, pady=3, sticky="w")
                    current_row += 1

# history_windows.py (นำไปวางทับคลาส PurchaseDetailWindow เดิมทั้งหมด)

# history_windows.py

class PurchaseDetailWindow(CTkToplevel):
    def __init__(self, master, app_container, purchase_id, approve_callback=None, reject_callback=None, on_save_callback=None):
        super().__init__(master)
        self.title(f"รายละเอียด/แก้ไขใบสั่งซื้อ (PO ID: {purchase_id})")
        self.geometry("900x700")
        
        self.app_container = app_container
        self.purchase_id = purchase_id
        self.on_save_callback = on_save_callback
        self.user_role = self.app_container.current_user_role 
        
        # ตัวแปรสำหรับเก็บ Widget ที่แก้ไขได้
        self.po_entries = {}
        self.item_widgets = []

        # Layout หลัก
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.scroll_frame = CTkScrollableFrame(self)
        self.scroll_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)
        
        # สร้างปุ่ม Action ไว้ด้านล่างสุด
        self._create_action_buttons()

        self.after(50, self._load_and_display_data)
        self.transient(master)
        self.grab_set()

    def _load_and_display_data(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (self.purchase_id,))
                po_data = cursor.fetchone()
                cursor.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", (self.purchase_id,))
                items_data = cursor.fetchall()
                cursor.execute("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", (self.purchase_id,))
                payments_data = cursor.fetchall()
            
            if not po_data:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบ PO ID: {self.purchase_id}", parent=self)
                self.destroy(); return

            self.create_formatted_view(po_data, items_data, payments_data)

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def create_formatted_view(self, po_data, items_data, payments_data):
        """สร้าง UI ทั้งหมดแบบละเอียด และเป็นฟอร์มที่แก้ไขได้"""
        self._create_info_section(self.scroll_frame, po_data)
        self._create_items_section(self.scroll_frame, items_data)
        self._create_summary_section(self.scroll_frame, po_data)
        self._create_payments_section(self.scroll_frame, payments_data)
        self._create_shipping_section(self.scroll_frame, po_data)

    def _create_section(self, parent, title):
        section_frame = CTkFrame(parent, corner_radius=10, border_width=1)
        section_frame.pack(fill="x", padx=10, pady=10)
        section_frame.grid_columnconfigure(1, weight=1)
        title_label = CTkLabel(section_frame, text=title, font=CTkFont(size=16, weight="bold"))
        title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(5,10), sticky="w")
        return section_frame

    def _create_editable_row(self, parent, row, label, value, key):
        """สร้างแถวข้อมูลที่แก้ไขได้ และเก็บ reference ไว้"""
        CTkLabel(parent, text=label).grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        is_numeric_key = any(s in key for s in ['cost', 'total', 'grand'])
        if isinstance(value, (int, float, np.floating)) or is_numeric_key:
            entry = NumericEntry(parent)
            entry.insert(0, f"{float(value):,.2f}" if value is not None else "0.00")
        else:
            entry = CTkEntry(parent)
            entry.insert(0, str(value) if value is not None else "")

        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries[key] = entry

    def _create_info_section(self, parent, data):
        info_frame = self._create_section(parent, "ข้อมูลทั่วไป")
        self._create_editable_row(info_frame, 1, "SO Number:", data.get("so_number"), key="so_number")
        self._create_editable_row(info_frame, 2, "PO Number:", data.get("po_number"), key="po_number")
        self._create_editable_row(info_frame, 3, "ชื่อซัพพลายเออร์:", data.get("supplier_name"), key="supplier_name")
        self._create_editable_row(info_frame, 4, "Credit Term:", data.get("credit_term"), key="credit_term")

    def _create_items_section(self, parent, items_list):
        items_frame = self._create_section(parent, "รายการสินค้า")
        
        # --- [แก้ไข] กำหนด Grid ให้กับ items_frame ---
        items_frame.grid_columnconfigure(0, weight=4) # Name
        items_frame.grid_columnconfigure(1, weight=1) # Qty
        items_frame.grid_columnconfigure(2, weight=2) # Price
        
        # --- [แก้ไข] Header (ใช้ .grid() และกำหนด row=1 เพราะ row=0 คือ Title ของ Section) ---
        header = CTkFrame(items_frame, fg_color="#E5E7EB", corner_radius=0)
        header.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=(0, 2))
        header.grid_columnconfigure(0, weight=4)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=2)
        CTkLabel(header, text="Product Name").grid(row=0, column=0, padx=5, sticky="w")
        CTkLabel(header, text="Quantity").grid(row=0, column=1, padx=5)
        CTkLabel(header, text="Unit Price").grid(row=0, column=2, padx=5)

        if not items_list:
            CTkLabel(items_frame, text="ไม่มีรายการสินค้า").grid(row=2, column=0, columnspan=3, pady=10)
            return
        
        # Item Rows (จะถูกเพิ่มด้วย grid ใน _add_item_row)
        for item in items_list:
            self._add_item_row(items_frame, item)

    def _add_item_row(self, parent, item_data):
        # row_index จะเริ่มที่ 2 เพราะ row 0 คือ Title, row 1 คือ Header
        row_index = len(self.item_widgets) + 2

        # --- [แก้ไข] สร้าง Frame และวางด้วย .grid() ---
        row_frame = CTkFrame(parent, fg_color="transparent")
        row_frame.grid(row=row_index, column=0, columnspan=3, sticky="ew", padx=5, pady=1)
        row_frame.grid_columnconfigure(0, weight=4) # Name
        row_frame.grid_columnconfigure(1, weight=1) # Qty
        row_frame.grid_columnconfigure(2, weight=2) # Price
        
        entry_name = CTkEntry(row_frame)
        entry_name.insert(0, item_data.get('product_name', ''))
        entry_name.grid(row=0, column=0, padx=(0,2), sticky="ew")

        entry_qty = NumericEntry(row_frame)
        entry_qty.insert(0, f"{item_data.get('quantity', 0):.2f}")
        entry_qty.grid(row=0, column=1, padx=2, sticky="ew")

        entry_price = NumericEntry(row_frame)
        entry_price.insert(0, f"{item_data.get('unit_price', 0):.2f}")
        entry_price.grid(row=0, column=2, padx=(2,0), sticky="ew")
        
        self.item_widgets.append({
            'id': item_data['id'],
            'name_entry': entry_name,
            'qty_entry': entry_qty,
            'price_entry': entry_price
        })

    def _create_summary_section(self, parent, summary_data):
        summary_frame = self._create_section(parent, "สรุปยอด")
        self._create_editable_row(summary_frame, 1, "ยอดรวมต้นทุนสินค้า:", summary_data.get("total_cost"), key="total_cost")
        self._create_editable_row(summary_frame, 2, "ยอดรวมที่ต้องชำระ:", summary_data.get("grand_total"), key="grand_total")

    def _create_payments_section(self, parent, payments_list):
        payments_frame = self._create_section(parent, "การชำระเงิน")
        if not payments_list:
            CTkLabel(payments_frame, text="ไม่มีข้อมูลการชำระเงิน").grid(row=1, column=0, pady=10)
        else:
            for i, payment in enumerate(payments_list):
                p_text = f"{payment.get('payment_type', 'N/A')}: {payment.get('amount'):,.2f} บาท (วันที่: {payment.get('payment_date')})"
                CTkLabel(payments_frame, text=p_text).grid(row=i+1, column=0, padx=10, pady=2, sticky="w")

    def _create_shipping_section(self, parent, shipping_data):
        shipping_frame = self._create_section(parent, "ข้อมูลการจัดส่ง")
        CTkLabel(shipping_frame, text="--- การจัดส่งเข้าสต๊อก ---", font=CTkFont(slant="italic")).grid(row=1, column=0, columnspan=2, pady=(5,2), sticky="w", padx=10)
        self._create_editable_row(shipping_frame, 2, "ค่าจัดส่ง (สต๊อก):", shipping_data.get("shipping_to_stock_cost"), key="shipping_to_stock_cost")
        CTkLabel(shipping_frame, text="--- การจัดส่งเข้าไซต์งาน ---", font=CTkFont(slant="italic")).grid(row=3, column=0, columnspan=2, pady=(10,2), sticky="w", padx=10)
        self._create_editable_row(shipping_frame, 4, "ค่าจัดส่ง (ไซต์):", shipping_data.get("shipping_to_site_cost"), key="shipping_to_site_cost")
    
    def _create_action_buttons(self):
        button_frame = CTkFrame(self)
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        button_frame.grid_columnconfigure(0, weight=1)

        # --- ตรวจสอบ Role ของผู้ใช้ตรงนี้ ---
        if self.user_role == 'HR':
            # ถ้าเป็น HR, ให้แสดงปุ่ม "บันทึกการแก้ไข"
            save_button = CTkButton(button_frame, text="บันทึกการแก้ไข", command=self._save_changes)
            save_button.pack(pady=10)
        else:
            # ถ้าเป็น Role อื่น (เช่น PU), ให้แสดงแค่ปุ่ม "ปิด"
            close_button = CTkButton(button_frame, text="ปิด", command=self.destroy, fg_color="gray")
            close_button.pack(pady=10)

    def _save_changes(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # 1. อัปเดตข้อมูล PO หลัก
                cursor.execute("""
                    UPDATE purchase_orders SET 
                        so_number = %s, po_number = %s, supplier_name = %s, 
                        credit_term = %s, total_cost = %s, grand_total = %s,
                        shipping_to_stock_cost = %s, shipping_to_site_cost = %s
                    WHERE id = %s
                """, (
                    self.po_entries['so_number'].get(), self.po_entries['po_number'].get(),
                    self.po_entries['supplier_name'].get(), self.po_entries['credit_term'].get(),
                    utils.convert_to_float(self.po_entries['total_cost'].get()),
                    utils.convert_to_float(self.po_entries['grand_total'].get()),
                    utils.convert_to_float(self.po_entries['shipping_to_stock_cost'].get()),
                    utils.convert_to_float(self.po_entries['shipping_to_site_cost'].get()),
                    self.purchase_id
                ))

                # 2. อัปเดตรายการสินค้า
                for item_row in self.item_widgets:
                    item_id = item_row['id']
                    new_name = item_row['name_entry'].get()
                    new_qty = utils.convert_to_float(item_row['qty_entry'].get())
                    new_price = utils.convert_to_float(item_row['price_entry'].get())
                    new_total = new_qty * new_price
                    cursor.execute("""
                        UPDATE purchase_order_items 
                        SET product_name = %s, quantity = %s, unit_price = %s, total_price = %s
                        WHERE id = %s
                    """, (new_name, new_qty, new_price, new_total, item_id))

                # 3. บันทึก Log
                log_details = { "message": f"Edited PO: {self.po_entries['po_number'].get()} by HR ({self.app_container.current_user_key})" }
                cursor.execute("""
                    INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, ('PO Edited by HR', 'purchase_orders', self.purchase_id, self.app_container.current_user_key, json.dumps(log_details), datetime.now()))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกการแก้ไข PO เรียบร้อยแล้ว", parent=self)
            
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
        # หากมี job ที่ตั้งเวลาไว้ก่อนหน้า ให้ยกเลิกไป
        if self._debounce_job:
            self.after_cancel(self._debounce_job)

        # ตั้งเวลาเพื่อเรียกฟังก์ชันค้นหาจริงในอีก 500 มิลลิวินาที (0.5 วินาที)
        self._debounce_job = self.after(500, self._apply_filters)

    def __init__(self, master, app_container, sale_key_filter=None, on_row_double_click=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.sale_key_filter = sale_key_filter
        self.on_row_double_click_callback = on_row_double_click
        self.df = None
        
        # --- ตัวแปรสำหรับ Pagination และ Filter ---
        self.current_page = 0
        self.rows_per_page = 50
        self.total_rows = 0
        self.total_pages = 0
        self.active_tab = "drafts"
        
        # --- ตัวแปรสำหรับ UI และฟิลเตอร์เดือน/ปี ---
        self.title(f"ประวัติการบันทึกของ: {self.sale_key_filter}")
        self.geometry("1400x700")
        try: self.theme = master.THEME["sale"]
        except (AttributeError, KeyError): self.theme = {"header": "#1D4ED8", "primary": "#3B82F6"}
        
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- สร้าง UI Layout ใหม่ ---
        self._create_new_layout()
        
        self.after(50, self._load_initial_data)
        self.transient(master)
        self.grab_set()
        self.focus()
    
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
            query = "SELECT id, timestamp, user_key, so_number, po_number, supplier_name, approved_by FROM purchase_orders WHERE status = 'Approved' ORDER BY timestamp DESC"
            self.all_po_df = pd.read_sql_query(query, self.pg_engine)
            self.all_po_df['timestamp'] = pd.to_datetime(self.all_po_df['timestamp'])
            self._hide_loading()
            self._apply_filters()
        except Exception as e:
            self._hide_loading()
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดประวัติได้: {e}", parent=self)

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
        columns = df.columns.tolist()
        tree = ttk.Treeview(parent, columns=columns, show='headings')
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Roboto', 14, 'bold'))
        style.configure("Treeview", rowheight=25, font=('Roboto', 12))
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor='w')
        for index, row in df.iterrows():
            row['timestamp'] = row['timestamp'].strftime('%Y-%m-%d %H:%M:%S')
            tree.insert("", "end", values=list(row))
        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        tree.bind("<Double-1>", lambda e: self._on_row_double_click(e, tree))


class CommissionHistoryWindow(CTkToplevel):
    def __init__(self, master, app_container, sale_key_filter=None, on_row_double_click=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.sale_key_filter = sale_key_filter
        self.on_row_double_click_callback = on_row_double_click
        self.df = None
        
        # --- ตัวแปรสำหรับ Pagination และ Filter ---
        self.current_page = 0
        self.rows_per_page = 50
        self.total_rows = 0
        self.total_pages = 0
        self.active_tab = "drafts" # 'drafts' or 'submitted'

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")
        
        # --- ตัวแปรสำหรับ UI ---
        self.title(f"ประวัติการบันทึกของ: {self.sale_key_filter}")
        self.geometry("1400x700")
        try: self.theme = master.THEME["sale"]
        except (AttributeError, KeyError): self.theme = {"header": "#1D4ED8", "primary": "#3B82F6"}
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- สร้าง UI Layout ใหม่ ---
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
        """โหลดและแสดงข้อมูลตาม Tab และฟิลเตอร์ที่เลือก"""
        target_frame = self.draft_frame if self.active_tab == "drafts" else self.submitted_frame
        for widget in target_frame.winfo_children(): widget.destroy()
        self._show_loading()

        try:
            if self.active_tab == "drafts": status_condition = "status IN ('Original', 'Edited', 'Rejected by SM', 'Rejected by HR')"
            else: status_condition = "status NOT IN ('Original', 'Edited', 'Rejected by SM', 'Rejected by HR', 'Cancelled')"
            
            # --- START: เพิ่ม Logic การกรองใน Query ---
            base_query = f"FROM commissions WHERE sale_key = %s AND is_active = 1 AND {status_condition}"
            params = [self.sale_key_filter]

            selected_month_str = self.month_var.get()
            if selected_month_str != "ทุกเดือน":
                month_num = self.thai_month_map[selected_month_str]
                base_query += " AND EXTRACT(MONTH FROM timestamp::timestamp) = %s"
                params.append(month_num)

            selected_year_str = self.year_var.get()
            if selected_year_str != "ทุกปี":
                year_num = int(selected_year_str)
                base_query += " AND EXTRACT(YEAR FROM timestamp::timestamp) = %s"
                params.append(year_num)
            # --- END ---

            count_query = f"SELECT COUNT(*) {base_query}"
            count_df = pd.read_sql_query(count_query, self.pg_engine, params=tuple(params))
            self.total_rows = count_df.iloc[0, 0] if not count_df.empty else 0
            self.total_pages = (self.total_rows + self.rows_per_page - 1) // self.rows_per_page

            offset = self.current_page * self.rows_per_page
            data_query = f"SELECT * {base_query} ORDER BY timestamp DESC LIMIT %s OFFSET %s"
            params.extend([self.rows_per_page, offset])
            
            self.df = pd.read_sql_query(data_query, self.pg_engine, params=tuple(params))
            
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
        """สร้าง Treeview และเติมข้อมูล (ปรับปรุงให้มีสีสัน)"""
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        columns = ['id', 'timestamp', 'status', 'so_number', 'customer_name', 'sales_service_amount', 'shipping_cost', 'rejection_reason']
        display_columns = ['ID', 'เวลาบันทึก', 'สถานะ', 'SO Number', 'ชื่อลูกค้า', 'ยอดขาย/บริการ', 'ค่าขนส่ง', 'เหตุผลที่ถูกตีกลับ']
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("History.Treeview.Heading", font=('Roboto', 11, 'bold'), relief="flat", background="#E5E7EB")
        style.configure("History.Treeview", rowheight=28, font=('Roboto', 11))
        style.map("History.Treeview", background=[('selected', self.theme.get("primary", "#3B82F6"))])
        
        self.tree = ttk.Treeview(parent, columns=columns, show='headings', style="History.Treeview")
        
        # กำหนดสีสำหรับแต่ละสถานะ
        self.tree.tag_configure('Draft', background='#FEFCE8') # เหลืองอ่อน
        self.tree.tag_configure('Rejected', background='#FEF2F2') # แดงอ่อน
        self.tree.tag_configure('Submitted', background='#F0FDF4') # เขียวอ่อน
        self.tree.tag_configure('Default', background='white')

        for i, col_id in enumerate(columns):
            width = 100
            anchor = 'w'
            if col_id in ['id', 'status']: width = 80
            elif col_id in ['sales_service_amount', 'shipping_cost']: width = 120; anchor = 'e'
            elif col_id in ['customer_name', 'so_number']: width = 200
            elif col_id == 'timestamp': width = 160
            elif col_id == 'rejection_reason': width = 250
            
            self.tree.heading(col_id, text=display_columns[i])
            self.tree.column(col_id, anchor=anchor, width=width)
        
        for index, row in df.iterrows():
            status = row['status']
            tag = 'Default'
            if 'Reject' in status: tag = 'Rejected'
            elif status in ['Original', 'Edited']: tag = 'Draft'
            else: tag = 'Submitted'
            
            values = []
            for col_name in columns:
                value = row.get(col_name) # ใช้ .get() เพื่อความปลอดภัย
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
    def __init__(self, master, sales_data, so_shared_vars, sale_theme):
        super().__init__(master)
        self.master = master
        self.sales_data = sales_data
        self.so_shared_vars = so_shared_vars
        self.sale_theme = sale_theme
        self.popup_widgets = {}
        self.trace_ids_for_so_calc = []
        self.title(f"ข้อมูล Sales Order (SO: {sales_data.get('so_number', 'N/A')})")
        self.geometry("700x700")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.so_data_form_frame = CTkScrollableFrame(self, corner_radius=10, label_text="ข้อมูล Sales Order (แก้ไขได้)", label_fg_color=self.sale_theme["bg"], label_text_color=self.sale_theme["header"])
        self.so_data_form_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self._create_so_data_form_content(self.so_data_form_frame)
        self._so_bind_events()
        self.after(100, lambda: self._populate_so_form(self.sales_data))

        # ปุ่มปิดจะอยู่ด้านล่างปุ่ม Save
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
    
    def _add_item_row_with_vat(self, parent, label_text, entry_key, vat_option_key, row_index):
        """
        ฟังก์ชัน Helper สำหรับสร้างแถวข้อมูลที่มีช่องกรอกตัวเลขและตัวเลือก VAT/CASH
        """
        # 1. สร้าง Label แสดงชื่อรายการ (เช่น "ยอดขายสินค้า/บริการ:")
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(
            row=row_index, column=0, padx=(15, 10), pady=4, sticky="w"
        )
        
        # 2. สร้าง Frame เพื่อจัดกลุ่มช่องกรอกตัวเลขและปุ่ม Radio ให้อยู่ด้วยกัน
        item_frame = CTkFrame(parent, fg_color="transparent")
        item_frame.grid(row=row_index, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")
        item_frame.grid_columnconfigure(0, weight=1) # ทำให้ช่องกรอกตัวเลขขยายได้

        # 3. สร้างช่องกรอกตัวเลข (NumericEntry)
        amount_entry = NumericEntry(item_frame)
        amount_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # 4. เก็บ Widget ที่สร้างขึ้นไว้ใน Dictionary เพื่อใช้อ้างอิงภายหลัง
        self.popup_widgets[entry_key] = amount_entry

        # 5. สร้าง Frame สำหรับปุ่ม Radio (VAT/CASH)
        vat_frame = CTkFrame(item_frame, fg_color="transparent")
        vat_frame.pack(side="left")

        # 6. สร้างปุ่ม Radio "VAT" และ "CASH"
        # โดยใช้ตัวแปร (StringVar) จาก so_shared_vars ที่ส่งเข้ามา
        CTkRadioButton(
            vat_frame, 
            text="VAT", 
            variable=self.so_shared_vars[vat_option_key], 
            value="VAT"
        ).pack(side="left")
        
        CTkRadioButton(
            vat_frame, 
            text="CASH", 
            variable=self.so_shared_vars[vat_option_key], 
            value="CASH"
        ).pack(side="left", padx=5)
            
    def _create_so_data_form_content(self, parent_frame):
        # Section 1: Sales Details
        f1 = self._create_so_section_frame(parent_frame, "รายละเอียดการขาย")
        self._add_form_row(f1, "วันที่เปิด SO:", DateSelector(f1, dropdown_style=self.master.dropdown_style), 'bill_date_selector', 1)
        self._add_form_row(f1, "ชื่อลูกค้า:", CTkEntry(f1), 'customer_name_entry', 2)
        self._add_form_row(f1, "รหัสลูกค้า:", CTkEntry(f1), 'customer_id_entry', 3)
        self._add_form_row(f1, "Credit Term:", CTkEntry(f1), 'credit_term_entry', 4)

        # Section 2: Sales and Services
        f2 = self._create_so_section_frame(parent_frame, "ยอดขายและบริการ")
        self._add_item_row_with_vat(f2, "ยอดขายสินค้า/บริการ:", 'sales_amount_entry', 'sales_service_vat_option', 1)
        self._add_item_row_with_vat(f2, "ค่าบริการตัด/เจาะ:", 'cutting_drilling_fee_entry', 'cutting_drilling_fee_vat_option', 2)
        self._add_item_row_with_vat(f2, "ค่าบริการอื่นๆ:", 'other_service_fee_entry', 'other_service_fee_vat_option', 3)
        
        # Section 3: Shipping Cost
        f3 = self._create_so_section_frame(parent_frame, "ค่าจัดส่ง")
        self._add_item_row_with_vat(f3, "ค่าจัดส่ง:", 'shipping_cost_entry', 'shipping_vat_option_var', 1)
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
        self._add_form_row(f4, "ค่าย้าย:", NumericEntry(f4), 'relocation_cost_entry', 3)
        self._add_form_row(f4, "วันที่ย้ายเข้าคลัง:", DateSelector(f4, dropdown_style=self.master.dropdown_style), 'date_to_wh_selector', 4)
        self._add_form_row(f4, "วันที่จัดส่งลูกค้า:", DateSelector(f4, dropdown_style=self.master.dropdown_style), 'date_to_customer_selector', 5)
        self._add_form_row(f4, "ทะเบียนเข้ารับ:", CTkEntry(f4), 'pickup_rego_entry', 6)

        # Section 5: Fees and Discounts
        f5 = self._create_so_section_frame(parent_frame, "ค่าธรรมเนียมและส่วนลด")
        self._add_item_row_with_vat(f5, "ค่าธรรมเนียมบัตร:", 'credit_card_fee_entry', 'credit_card_fee_vat_option_var', 1)
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
            "cash_product_input_entry", "cash_actual_payment_entry"
        ]
        for key in widgets_to_bind_keys:
            if key in self.popup_widgets and isinstance(self.popup_widgets[key], (CTkEntry, NumericEntry)):
                self.popup_widgets[key].bind("<KeyRelease>", self._so_update_final_calculations)
            
        radio_vars_keys = [
            'sales_service_vat_option', 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option', 'shipping_vat_option_var',
            'credit_card_fee_vat_option_var'
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
            if entry_widget:
                if hasattr(entry_widget, '_entry') and entry_widget._entry.winfo_exists():
                    try: return utils.convert_to_float(entry_widget.get())
                    except (tk.TclError, ValueError): return 0.0
                else: return 0.0
            return 0.0

        sales = get_float_from_entry('sales_amount_entry')
        shipping = get_float_from_entry('shipping_cost_entry')
        card_fee = get_float_from_entry('credit_card_fee_entry')
        cutting_drilling = get_float_from_entry('cutting_drilling_fee_entry')
        other_service = get_float_from_entry('other_service_fee_entry')
        
        total_vatable_subtotal, total_cashable_services_and_fees = 0.0, 0.0
        
        items_to_process = [(sales, w_vars['sales_service_vat_option'].get()), (cutting_drilling, w_vars['cutting_drilling_fee_vat_option'].get()), (other_service, w_vars['other_service_fee_vat_option'].get()), (shipping, w_vars['shipping_vat_option_var'].get()), (card_fee, w_vars['credit_card_fee_vat_option_var'].get())]
        for amount, option in items_to_process:
            if option == "VAT": total_vatable_subtotal += amount
            else: total_cashable_services_and_fees += amount
            
        so_grand_total = total_vatable_subtotal * 1.07
        w_vars['so_grand_total_var'].set(f"{so_grand_total:,.2f}")

        payment1 = get_float_from_entry('payment1_amount_entry')
        payment2 = get_float_from_entry('payment2_amount_entry')
        so_vs_payment_diff = (payment1 + payment2) - so_grand_total
        w_vars['difference_amount_var'].set(f"{so_vs_payment_diff:,.2f}")

        def set_check_result(label_widget_key, var, diff_val, plus_text, minus_text):
            label_widget_ref = w_widgets.get(label_widget_key)
            color_map = {"-": ("gray85", "black"), "ok": ("#BBF7D0", "#15803D"), "bad": ("#FECACA", "#B91C1C")}
            if abs(diff_val) < 0.01: state, text = "ok", "ถูกต้อง"
            elif diff_val > 0: state, text = "ok", f"{plus_text} (+{abs(diff_val):,.2f})"
            else: state, text = "bad", f"{minus_text} (-{abs(diff_val):,.2f})"
            if label_widget_ref and label_widget_ref.winfo_exists():
                var.set(text)
                if isinstance(label_widget_ref, CTkLabel): label_widget_ref.configure(fg_color=color_map[state][0], text_color=color_map[state][1], text=text)
                elif isinstance(label_widget_ref, CTkEntry):
                    current_state = label_widget_ref.cget("state")
                    if current_state == "readonly": label_widget_ref.configure(state="normal")
                    label_widget_ref.delete(0, "end"); label_widget_ref.insert(0, text)
                    label_widget_ref.configure(fg_color=color_map[state][0], text_color=color_map[state][1])
                    if current_state == "readonly": label_widget_ref.configure(state="readonly")

        set_check_result('so_check_display', w_vars.get('so_vs_payment_result_var'), so_vs_payment_diff, "ยอดโอนเกิน", "ยอดโอนขาด")

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
        if self.sales_data is None: 
            messagebox.showerror("ข้อผิดพลาด", "ไม่มีข้อมูล SO ให้บันทึก", parent=self)
            return
        self.master._save_so_changes_from_popup(self.sales_data.get('id'), self.so_shared_vars, self.popup_widgets)

