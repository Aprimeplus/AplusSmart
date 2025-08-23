# edit_commission_window.py (ฉบับแก้ไข: แยก บันทึก กับ ส่ง)

import tkinter as tk
from customtkinter import (CTkToplevel, CTkScrollableFrame, CTkLabel, CTkEntry, CTkButton,
                           CTkFont, CTkFrame)
from tkinter import messagebox
import psycopg2
import psycopg2.errors
import psycopg2.extras
from datetime import datetime
import pandas as pd
import traceback
import numpy as np

# สันนิษฐานว่าคุณมีไฟล์ custom_widgets และ utils อยู่
try:
    from custom_widgets import NumericEntry, DateSelector
    import utils
except ImportError:
    messagebox.showwarning("คำเตือน", "ไม่พบไฟล์ custom_widgets หรือ utils กรุณาตรวจสอบ")
    # สร้างคลาสจำลองเพื่อไม่ให้โปรแกรมพัง
    class NumericEntry(tk.Entry): pass
    class DateSelector(tk.Entry): pass
    class utils:
        @staticmethod
        def convert_to_float(s):
            try: return float(str(s).replace(',', ''))
            except (ValueError, TypeError): return 0.0


class EditCommissionWindow(CTkToplevel):
    def __init__(self, parent, app_container, data, refresh_callback, user_role=None):
        super().__init__(parent)
        self.parent = parent
        self.app_container = app_container
        self.data_row = data
        self.refresh_callback = refresh_callback
        self.user_role = user_role

        self.readonly_cols = ['sale_key', 'product_vat_7', 'shipping_vat_7', 'difference_amount']
        self.date_cols = ['bill_date', 'payment_date', 'delivery_date', 'payment1_date', 'payment2_date', 'date_to_warehouse', 'date_to_customer']
        self.numeric_cols = [
            'sales_service_amount', 'shipping_cost', 'total_payment_amount', 'brokerage_fee',
            'giveaways', 'coupons', 'transfer_fee', 'credit_card_fee', 'wht_3_percent',
            'relocation_cost', 'cash_product_input', 'cash_actual_payment'
        ]
        self.integer_cols = ['commission_month', 'commission_year']
        self.so_number = self.data_row.get('so_number')

        if not self.so_number:
            messagebox.showerror("ข้อมูลผิดพลาด", "ไม่สามารถหา SO Number จากข้อมูลที่เลือกได้", parent=self)
            self.after(10, self.destroy)
            return

        self.original_record_id = None
        self.title(f"แก้ไขข้อมูล SO: {self.so_number} (โดย {self.user_role or 'Sale'})")
        self.geometry("600x800")
        self.entries = {}

        self.main_frame = CTkScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.create_form()
        self.load_data()

        self.transient(parent)
        self.grab_set()

    def create_form(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'commissions' ORDER BY ordinal_position")
                db_column_info = cursor.fetchall()

            self.db_column_names = [info['column_name'] for info in db_column_info]

            form_frame = CTkFrame(self.main_frame, fg_color="transparent")
            form_frame.pack(fill="x", expand=True, padx=5, pady=5)
            form_frame.grid_columnconfigure(1, weight=1)
            self.entries = {}

            header_map_thai = self.app_container.HEADER_MAP

            excluded_cols = ['id', 'status', 'is_active', 'original_id', 'timestamp', 'rejection_reason', 'approver_sale_manager_key', 'approval_date_sale_manager', 'claim_timestamp']
            row_index = 0
            for col_name in self.db_column_names:
                if col_name in excluded_cols: continue

                display_name = header_map_thai.get(col_name, col_name)
                label = CTkLabel(form_frame, text=f"{display_name}:", font=("Roboto", 14))
                label.grid(row=row_index, column=0, padx=10, pady=5, sticky="w")

                if col_name in self.date_cols: entry = DateSelector(form_frame)
                elif col_name in self.numeric_cols: entry = NumericEntry(form_frame, font=("Roboto", 14))
                else: entry = CTkEntry(form_frame, width=300, font=("Roboto", 14))

                entry.grid(row=row_index, column=1, padx=10, pady=5, sticky="ew")
                self.entries[col_name] = entry
                row_index += 1

                if col_name in self.readonly_cols:
                    entry.configure(state="disabled")
                    if not isinstance(entry, DateSelector):
                        entry.configure(fg_color="gray85")
            
            # --- START: 1. เปลี่ยนข้อความบนปุ่ม ---
            self.save_button = CTkButton(self, text="บันทึกการแก้ไข", command=self._save_changes, font=("Roboto", 16, "bold"))
            # --- END: 1. เปลี่ยนข้อความบนปุ่ม ---
            self.save_button.pack(pady=(5, 10), padx=10)

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถสร้างฟอร์มได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def load_data(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # แก้ไขให้ดึงข้อมูลจาก id แทน so_number เพื่อความแม่นยำ
                record_id_to_load = self.data_row.get('id')
                cursor.execute("SELECT * FROM commissions WHERE id = %s AND is_active = 1", (record_id_to_load,))
                db_data_dict = cursor.fetchone()

            if db_data_dict:
                self.original_record_id = db_data_dict.get('id')
                for col_name, entry_widget in self.entries.items():
                    value = db_data_dict.get(col_name)
                    if isinstance(entry_widget, DateSelector): entry_widget.set_date(value)
                    elif isinstance(entry_widget, CTkEntry):
                        if entry_widget.winfo_exists():
                            entry_widget.delete(0, tk.END)
                            if value is not None:
                                if isinstance(value, (int, float, np.floating)):
                                    entry_widget.insert(0, f"{value:,.2f}")
                                else:
                                    entry_widget.insert(0, str(value))
            else:
                 messagebox.showerror("ผิดพลาด", f"ไม่พบข้อมูล SO: {self.so_number} ในระบบ", parent=self)
                 self.destroy()

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูลได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _save_changes(self):
        so_from_entry = self.entries['so_number'].get().strip()
        if not so_from_entry:
            messagebox.showerror("ผิดพลาด", "SO Number ไม่สามารถเป็นค่าว่างได้", parent=self)
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                if self.original_record_id is not None:
                    # Deactivate the old record, marking it as corrected
                    cursor.execute("UPDATE commissions SET is_active = 0, status = 'Corrected by SM' WHERE id = %s", (self.original_record_id,))

                new_data = {}
                for col_name, entry_widget in self.entries.items():
                    if isinstance(entry_widget, DateSelector): new_data[col_name] = entry_widget.get_date()
                    elif col_name in self.integer_cols:
                        value = entry_widget.get()
                        new_data[col_name] = int(float(value.replace(",", ""))) if value else None
                    elif col_name in self.numeric_cols: new_data[col_name] = utils.convert_to_float(entry_widget.get())
                    else: new_data[col_name] = entry_widget.get().strip() if entry_widget.get() else None

                # --- START: 2. ปรับ Logic การบันทึก ---
                if self.user_role == 'Sales Manager':
                    # เมื่อ Manager แก้ไข ให้สถานะกลับไปเป็น "รออนุมัติ" จาก Manager อีกครั้ง
                    # เพื่อให้ Manager เป็นคนกดยืนยัน (Approve) จากหน้าจอหลักเอง
                    new_data["status"] = 'Awaiting SM Approval'
                else: # กรณี Sale เป็นคนแก้ไขเอง (จากงานที่ถูกตีกลับ)
                    new_data["status"] = 'Edited' # หรืออาจจะเป็น 'Pending Sale Manager Approval' ก็ได้ ขึ้นอยู่กับ Flow
                # --- END: 2. ปรับ Logic การบันทึก ---

                new_data["is_active"] = 1
                new_data["original_id"] = self.original_record_id
                new_data["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                all_db_columns = [col for col in self.db_column_names if col != 'id']
                ordered_values = [new_data.get(col) for col in all_db_columns]

                cols_str = ', '.join([f'"{k}"' for k in all_db_columns])
                placeholders_str = ', '.join(['%s'] * len(all_db_columns))
                cursor.execute(f"INSERT INTO commissions ({cols_str}) VALUES ({placeholders_str}) RETURNING id", tuple(ordered_values))
                
                # --- START: 3. เอาส่วนการส่ง Notification ออกจากหน้านี้ ---
                # การส่ง Notification จะเกิดขึ้นเมื่อ Manager กดปุ่ม "รวบรวมส่ง HR" ที่หน้าจอหลัก
                # --- END: 3. เอาส่วนการส่ง Notification ออก ---

            conn.commit()
            # --- START: 4. เปลี่ยนข้อความยืนยัน ---
            messagebox.showinfo("สำเร็จ", "บันทึกการแก้ไขเรียบร้อยแล้ว\nกรุณากด 'อนุมัติ' จากหน้าจอหลักอีกครั้งเพื่อยืนยันและส่งต่อ", parent=self)
            # --- END: 4. เปลี่ยนข้อความยืนยัน ---
            
            self.refresh_callback() # รีเฟรชหน้าจอ Sale Manager
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)