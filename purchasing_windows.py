# purchasing_windows.py (ฉบับแก้ไขสมบูรณ์)

import customtkinter as ctk
from tkinter import messagebox, StringVar, TclError
import psycopg2
import pandas as pd
from datetime import datetime
import traceback

from custom_widgets import AutoCompleteEntry, DateSelector
import utils
import psycopg2.extras


class PurchaseOrderWindow(ctk.CTkToplevel):
    def __init__(self, master, app_container, po_id=None, on_close_callback=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = self.app_container.pg_engine
        self.po_id = po_id
        self.on_close_callback = on_close_callback
        self.item_widgets = []
        self.deleted_item_ids = []
        self.po_entries = {}
        
        # --- START: แก้ไขการโหลดข้อมูล ---
        self.supplier_completion_data = []
        self._load_supplier_data_for_autocomplete() # เรียกใช้ฟังก์ชันใหม่
        # --- END ---

        self.title(f"ใบสั่งซื้อ (PO) - {'แก้ไข' if self.po_id else 'สร้างใหม่'}")
        self.geometry("900x800")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.total_cost_var = StringVar(value="0.00")

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_scroll_frame = ctk.CTkScrollableFrame(self)
        main_scroll_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        main_scroll_frame.grid_columnconfigure(0, weight=1)

        self._create_po_details_section(main_scroll_frame)
        self._create_shipping_section(main_scroll_frame)
        self._create_items_section(main_scroll_frame)

        self.submit_button = ctk.CTkButton(self, text="บันทึกและส่งอนุมัติ (Submit PO)", command=self._submit_po)
        self.submit_button.grid(row=2, column=0, pady=10, padx=10, sticky="ew")

        if self.po_id:
            self._load_po_data()
        else:
            self._add_item_row()

        self.transient(master)
        self.grab_set()

    def _create_section_frame(self, parent, title):
        frame = ctk.CTkFrame(parent, corner_radius=10, border_width=1)
        frame.pack(fill="x", padx=5, pady=8)
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="w")
        return frame

    def _create_po_details_section(self, parent):
        po_details_frame = self._create_section_frame(parent, "ข้อมูลทั่วไป")

        # --- START: แก้ไขการสร้าง AutoCompleteEntry ---
        ctk.CTkLabel(po_details_frame, text="Supplier:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry_supplier = AutoCompleteEntry(
            po_details_frame,
            completion_list=self.supplier_completion_data,
            display_key='name',
            command=self._on_supplier_selected
        )
        entry_supplier.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries['supplier_name'] = entry_supplier
        # --- END ---

        ctk.CTkLabel(po_details_frame, text="PO Number:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        entry_po_number = ctk.CTkEntry(po_details_frame)
        entry_po_number.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries['po_number'] = entry_po_number
        
        ctk.CTkLabel(po_details_frame, text="SO Number:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        entry_so_number = ctk.CTkEntry(po_details_frame)
        entry_so_number.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.po_entries['so_number'] = entry_so_number

        ctk.CTkLabel(po_details_frame, text="Total Cost (excl. Shipping):").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(po_details_frame, textvariable=self.total_cost_var, font=ctk.CTkFont(size=12, weight="bold")).grid(row=4, column=1, padx=10, pady=5, sticky="w")

    def _create_shipping_section(self, parent):
        shipping_frame = self._create_section_frame(parent, "ข้อมูลการจัดส่ง (ภาพรวม PO)")
        
        ctk.CTkLabel(shipping_frame, text="--- ค่าจัดส่งเข้าสต๊อก ---", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, columnspan=2, pady=(5,2), sticky="w", padx=10)
        ctk.CTkLabel(shipping_frame, text="ค่าส่ง:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.po_entries['shipping_to_stock_cost'] = utils.FormattedNumericEntry(shipping_frame)
        self.po_entries['shipping_to_stock_cost'].grid(row=2, column=1, sticky="ew", padx=10, pady=5)

        ctk.CTkLabel(shipping_frame, text="วันที่:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.po_entries['shipping_to_stock_date'] = DateSelector(shipping_frame)
        self.po_entries['shipping_to_stock_date'].grid(row=3, column=1, sticky="w", padx=10, pady=5)

        ctk.CTkLabel(shipping_frame, text="--- ค่าจัดส่งเข้าไซต์ ---", font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, columnspan=2, pady=(10,2), sticky="w", padx=10)
        ctk.CTkLabel(shipping_frame, text="ค่าส่ง:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.po_entries['shipping_to_site_cost'] = utils.FormattedNumericEntry(shipping_frame)
        self.po_entries['shipping_to_site_cost'].grid(row=5, column=1, sticky="ew", padx=10, pady=5)

        ctk.CTkLabel(shipping_frame, text="วันที่:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.po_entries['shipping_to_site_date'] = DateSelector(shipping_frame)
        self.po_entries['shipping_to_site_date'].grid(row=6, column=1, sticky="w", padx=10, pady=5)

    def _create_items_section(self, parent):
        items_frame = self._create_section_frame(parent, "รายการสินค้า")
        header_frame = ctk.CTkFrame(items_frame, fg_color="transparent")
        header_frame.grid(row=1, column=0, sticky="ew")
        header_frame.grid_columnconfigure(0, weight=4)
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(2, weight=2)
        ctk.CTkLabel(header_frame, text="ชื่อสินค้า/รายละเอียด").grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkLabel(header_frame, text="จำนวน").grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkLabel(header_frame, text="ราคา/หน่วย").grid(row=0, column=2, padx=5, pady=5)

        self.items_content_frame = ctk.CTkFrame(items_frame, fg_color="transparent")
        self.items_content_frame.grid(row=2, column=0, sticky="ew")
        
        add_button = ctk.CTkButton(items_frame, text="+ เพิ่มรายการ", command=self._add_item_row)
        add_button.grid(row=3, column=0, padx=5, pady=10, sticky="w")

    # --- START: เพิ่มฟังก์ชันใหม่สำหรับโหลดข้อมูล ---
    def _load_supplier_data_for_autocomplete(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            df = pd.read_sql("SELECT id, supplier_name, supplier_code, credit_term FROM suppliers ORDER BY supplier_name", conn)
            self.supplier_completion_data = []
            for _, row in df.iterrows():
                self.supplier_completion_data.append({
                    "id": row['id'],
                    "name": row['supplier_name'],
                    "code": row.get('supplier_code', ''),
                    "term": row.get('credit_term', 'เงินสด')
                })
        except Exception as e:
            print(f"Error fetching supplier data for PO window: {e}")
            self.supplier_completion_data = []
        finally:
            if conn: self.app_container.release_connection(conn)
    # --- END ---
    
    # --- START: แก้ไขฟังก์ชัน Callback ---
    def _on_supplier_selected(self, selection_dict):
        """เมื่อเลือกซัพพลายเออร์, selection_dict คือ dictionary ที่ถูกส่งกลับมา"""
        # ฟังก์ชันนี้ถูกเรียกใช้โดย command ของ AutoCompleteEntry
        # ในหน้านี้เราอาจจะไม่ต้องทำอะไรเป็นพิเศษ แต่ต้องมีฟังก์ชันรองรับ
        if selection_dict:
            print(f"PO Window: Selected supplier -> {selection_dict.get('name')}")
        pass
    # --- END ---

    def _load_po_data(self):
        try:
            query_po = "SELECT * FROM purchase_orders WHERE id = %s"
            po_df = pd.read_sql(query_po, self.pg_engine, params=(self.po_id,))
            if po_df.empty:
                messagebox.showerror("Error", "PO not found.", parent=self)
                return
            po_data = po_df.iloc[0]

            # Set ข้อมูลพื้นฐาน
            utils.set_entry_text(self.po_entries['supplier_name'], po_data.get('supplier_name', ''))
            utils.set_entry_text(self.po_entries['po_number'], po_data.get('po_number', ''))
            utils.set_entry_text(self.po_entries['so_number'], po_data.get('so_number', ''))
            
            # ==============================================================================
            # [LOGIC ใหม่] ดึงค่ารถจากระบบขนส่ง (PX) มาใส่ถ้ามี
            # ==============================================================================
            current_po_num = po_data.get('po_number', '')
            stock_cost_val = po_data.get('shipping_to_stock_cost', 0) or 0
            
            if current_po_num:
                # เรียกฟังก์ชันใน AppContainer เพื่อเช็คและดึงค่ารถ
                # (ฟังก์ชันนี้จะคืนค่า > 0 ถ้าเจอ PX ที่สถานะ Pending Match)
                px_cost = self.app_container.sync_transport_cost_to_po(current_po_num)
                
                if px_cost > 0:
                    # ถ้าเจอค่ารถ และใน PO ปัจจุบันยังเป็น 0 (หรืออยากให้ทับ)
                    if stock_cost_val == 0:
                        stock_cost_val = px_cost
                        messagebox.showinfo("Auto Sync", f"🚚 พบค่ารถจากฝ่ายขนส่ง: {px_cost:,.2f} บาท\nระบบนำมาใส่ใน 'ค่าส่งเข้าสต๊อก' ให้แล้ว", parent=self)
            # ==============================================================================

            # นำค่าที่ได้ (เดิม หรือ ใหม่จาก PX) มาใส่ในช่อง
            self.po_entries['shipping_to_stock_cost'].set(stock_cost_val)
            
            self.po_entries['shipping_to_stock_date'].set_date(po_data.get('shipping_to_stock_date'))
            self.po_entries['shipping_to_site_cost'].set(po_data.get('shipping_to_site_cost', 0))
            self.po_entries['shipping_to_site_date'].set_date(po_data.get('shipping_to_site_date'))

            query_items = "SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id"
            items_df = pd.read_sql(query_items, self.pg_engine, params=(self.po_id,))
            for index, row in items_df.iterrows():
                self._add_item_row(row.to_dict())
            
            self._update_totals() 
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to load PO data: {e}", parent=self)
            traceback.print_exc()

    def _add_item_row(self, item=None):
        item_frame = ctk.CTkFrame(self.items_content_frame, fg_color="transparent")
        item_frame.pack(fill="x", expand=True, pady=2)
        item_frame.grid_columnconfigure(0, weight=4)
        item_frame.grid_columnconfigure(1, weight=1)
        item_frame.grid_columnconfigure(2, weight=2)
        item_frame.grid_columnconfigure(3, weight=0)

        entry_desc = ctk.CTkEntry(item_frame)
        entry_desc.grid(row=0, column=0, padx=5, pady=2, sticky="ew")
        
        entry_qty = utils.FormattedNumericEntry(item_frame, command=self._update_totals)
        entry_qty.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

        entry_cost = utils.FormattedNumericEntry(item_frame, command=self._update_totals)
        entry_cost.grid(row=0, column=2, padx=5, pady=2, sticky="ew")

        delete_button = ctk.CTkButton(item_frame, text="ลบ", width=40, fg_color="#DC2626", hover_color="#B91C1C", 
                                      command=lambda f=item_frame, i=item.get('id') if item else None: self._delete_item_row(f, i))
        delete_button.grid(row=0, column=3, padx=5, pady=2)

        if item:
            entry_desc.insert(0, item.get('product_name', ''))
            entry_qty.set(item.get('quantity', 1))
            entry_cost.set(item.get('unit_price', 0))

        self.item_widgets.append({
            'frame': item_frame, 'desc_entry': entry_desc, 'qty_entry': entry_qty, 'cost_entry': entry_cost,
            'id': item.get('id') if item else None
        })
        self._update_totals()

    def _delete_item_row(self, frame_to_delete, item_id):
        if item_id is not None:
            self.deleted_item_ids.append(item_id)
        
        self.item_widgets = [row for row in self.item_widgets if row['frame'] != frame_to_delete]
        frame_to_delete.destroy()
        self._update_totals()

    def _update_totals(self, *args):
        total_cost = 0.0
        for item_row in self.item_widgets:
            try:
                qty = item_row['qty_entry'].get_value()
                cost_per_unit = item_row['cost_entry'].get_value()
                total_cost += qty * cost_per_unit
            except (ValueError, TclError):
                continue
        self.total_cost_var.set(f"{total_cost:,.2f}")

    def _submit_po(self):
        supplier_name = self.po_entries['supplier_name'].get().strip()
        po_number = self.po_entries['po_number'].get().strip()
        so_number = self.po_entries['so_number'].get().strip()

        if not supplier_name or not so_number:
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก Supplier และ SO Number", parent=self)
            return

        total_cost = utils.convert_to_float(self.total_cost_var.get())
        header_data = {
            'supplier_name': supplier_name, 'po_number': po_number, 'so_number': so_number,
            'total_cost': total_cost, 'user_key': self.app_container.current_user_key,
            'shipping_to_stock_cost': self.po_entries['shipping_to_stock_cost'].get_value(),
            'shipping_to_stock_date': self.po_entries['shipping_to_stock_date'].get_date(),
            'shipping_to_site_cost': self.po_entries['shipping_to_site_cost'].get_value(),
            'shipping_to_site_date': self.po_entries['shipping_to_site_date'].get_date()
        }
        
        items_data = []
        for row in self.item_widgets:
            desc = row['desc_entry'].get().strip()
            if not desc: continue
            items_data.append({
                'id': row['id'], 'product_name': desc, 'quantity': row['qty_entry'].get_value(),
                'unit_price': row['cost_entry'].get_value(),
                'total_price': row['qty_entry'].get_value() * row['cost_entry'].get_value()
            })
        
        if not items_data:
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณาเพิ่มรายการสินค้าอย่างน้อย 1 รายการ", parent=self)
            return
            
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                if self.po_id:
                    set_clauses = [f"{key} = %s" for key in header_data.keys()]
                    params = list(header_data.values()) + [self.po_id]
                    cursor.execute(f"UPDATE purchase_orders SET {', '.join(set_clauses)} WHERE id = %s", tuple(params))
                else:
                    columns = header_data.keys()
                    values = [f"%({key})s" for key in columns]
                    cursor.execute(f"INSERT INTO purchase_orders ({', '.join(columns)}) VALUES ({', '.join(values)}) RETURNING id", header_data)
                    self.po_id = cursor.fetchone()[0]

                if self.deleted_item_ids:
                    cursor.execute("DELETE FROM purchase_order_items WHERE id IN %s", (tuple(self.deleted_item_ids),))

                for item in items_data:
                    item['purchase_order_id'] = self.po_id
                    if item.get('id'):
                        cursor.execute("""
                            UPDATE purchase_order_items SET product_name=%(product_name)s, quantity=%(quantity)s, unit_price=%(unit_price)s, total_price=%(total_price)s
                            WHERE id=%(id)s """, item)
                    else:
                        cursor.execute("""
                            INSERT INTO purchase_order_items (purchase_order_id, product_name, quantity, unit_price, total_price)
                            VALUES (%(purchase_order_id)s, %(product_name)s, %(quantity)s, %(unit_price)s, %(total_price)s) """, item)

            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกใบสั่งซื้อเรียบร้อยแล้ว", parent=self)
            self._on_close()
        except psycopg2.Error as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"ไม่สามารถบันทึก PO ได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_close(self):
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()
    
class SOFinderDialog(ctk.CTkToplevel):
    """
    หน้าต่างสำหรับค้นหา SO และแสดงผลลัพธ์พร้อมปุ่ม Action ตามสถานะของ SO
    """
    def __init__(self, master, so_number):
        super().__init__(master)
        self.master = master # master ในที่นี้คือ PurchasingScreen instance
        self.app_container = master.app_container
        self.so_number_to_find = so_number
        self.so_data = None

        self.title(f"ผลการค้นหาสำหรับ SO: {self.so_number_to_find}")
        self.geometry("600x300")
        self.grid_columnconfigure(1, weight=1)

        self.after(50, self._fetch_and_display_so)
        
        self.transient(master)
        self.grab_set()

    def _fetch_and_display_so(self):
        try:
            # Query ข้อมูล SO พร้อมทั้ง JOIN เพื่อเอาชื่อ Sale และชื่อคนที่ Claim งาน
            query = """
                SELECT 
                    c.*, 
                    u_sale.sale_name,
                    u_pu.sale_name as pu_claimer_name
                FROM commissions c
                LEFT JOIN sales_users u_sale ON c.sale_key = u_sale.sale_key
                LEFT JOIN sales_users u_pu ON c.user_key = u_pu.sale_key
                WHERE c.so_number = %s AND c.is_active = 1 LIMIT 1
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.so_number_to_find,))

            if df.empty:
                ctk.CTkLabel(self, text=f"ไม่พบข้อมูล SO Number: {self.so_number_to_find}", font=ctk.CTkFont(size=16, weight="bold"), text_color="orange").pack(pady=50)
                return

            self.so_data = df.iloc[0].to_dict()
            self._populate_ui()

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการค้นหา SO: {e}", parent=self)
            self.destroy()

    def _populate_ui(self):
        # --- แสดงรายละเอียดของ SO ที่เจอ ---
        details_frame = ctk.CTkFrame(self, fg_color="transparent")
        details_frame.pack(fill="x", padx=20, pady=20)
        details_frame.grid_columnconfigure(1, weight=1)

        def create_detail_row(row, label, value):
            ctk.CTkLabel(details_frame, text=label, font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, sticky="w", padx=5, pady=3)
            ctk.CTkLabel(details_frame, text=f":  {value}", wraplength=400, justify="left").grid(row=row, column=1, sticky="w", padx=5, pady=3)

        create_detail_row(0, "SO Number", self.so_data.get('so_number', 'N/A'))
        create_detail_row(1, "ลูกค้า", self.so_data.get('customer_name', 'N/A'))
        create_detail_row(2, "พนักงานขาย", self.so_data.get('sale_name', 'N/A'))
        
        status = self.so_data.get('status')
        status_label = ctk.CTkLabel(details_frame, text=":  " + status, font=ctk.CTkFont(weight="bold"))
        status_label.grid(row=3, column=1, sticky="w", padx=5, pady=3)
        ctk.CTkLabel(details_frame, text="สถานะปัจจุบัน", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, sticky="w", padx=5, pady=3)
        
        # --- สร้างปุ่ม Action ตามเงื่อนไขของสถานะ ---
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.pack(fill="x", padx=20, pady=10)

        if status == 'Pending PU':
            status_label.configure(text_color="#22C55E") # สีเขียว
            ctk.CTkButton(action_frame, text="รับงานและเริ่มสร้าง PO", command=self._claim_and_load, height=40).pack(fill="x")
        elif status == 'PO In Progress' and self.so_data.get('user_key') == self.master.user_key:
            status_label.configure(text_color="#F59E0B") # สีเหลือง
            claimer = self.so_data.get('pu_claimer_name', self.so_data.get('user_key'))
            create_detail_row(4, "ดำเนินการโดย", f"คุณ ({claimer})")
            ctk.CTkButton(action_frame, text="ทำต่อ (Continue)", command=self._claim_and_load, height=40).pack(fill="x")
        else:
            status_label.configure(text_color="#EF4444") # สีแดง
            claimer = self.so_data.get('pu_claimer_name', self.so_data.get('user_key', 'Unknown'))
            create_detail_row(4, "ดำเนินการโดย", claimer)

    def _claim_and_load(self):
        """
        เรียกใช้ฟังก์ชันบนหน้าจอหลักเพื่อโหลด SO นี้ และปิดหน้าต่างนี้
        """
        if self.so_data:
            so_num = self.so_data['so_number']
            # เราใช้ฟังก์ชัน select_so_from_task เพราะมันทำงานแบบเดียวกัน
            self.master.select_so_from_task(so_num)
            self.destroy()