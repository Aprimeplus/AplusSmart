# commission_app.py (ฉบับสมบูรณ์ รวมทุกฟังก์ชัน)

import tkinter as tk
from tkinter import messagebox, filedialog
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton, CTkRadioButton, 
                           CTkEntry, CTkOptionMenu, CTkScrollableFrame, CTkToplevel, CTkTabview, CTkCheckBox)
import pandas as pd
from datetime import datetime, timedelta
import traceback
import psycopg2.errors
import psycopg2.extras
import os 

from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
import utils
from export_utils import DateRangeDialog

class PaymentUpdateWindow(CTkToplevel):
    """หน้าต่าง Pop-up สำหรับอัปเดตข้อมูลการชำระเงินโดยเฉพาะ"""
    def __init__(self, master, app_container, so_data, on_save_callback):
        super().__init__(master)
        self.app_container = app_container
        self.so_data = so_data
        self.on_save_callback = on_save_callback
        self.so_id = self.so_data['id']

        self.title(f"อัปเดตยอดชำระ SO: {self.so_data['so_number']}")
        self.geometry("500x400")
        self.grid_columnconfigure(0, weight=1)

        # --- Display Info ---
        info_frame = CTkFrame(self, fg_color="transparent")
        info_frame.grid(row=0, column=0, padx=20, pady=15, sticky="ew")
        info_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(info_frame, text="SO Number:", font=CTkFont(weight="bold")).grid(row=0, column=0, sticky="w")
        CTkLabel(info_frame, text=self.so_data['so_number']).grid(row=0, column=1, sticky="w", padx=5)
        
        # <<< START: แก้ไขการดึงยอดที่ต้องชำระ >>>
        # ใช้สูตรคำนวณยอดที่ต้องชำระที่แท้จริง เพื่อความแม่นยำ
        original_payment = self.so_data.get('total_payment_amount', 0.0) or 0.0
        original_difference = self.so_data.get('difference_amount', 0.0) or 0.0
        actual_grand_total = original_payment - original_difference

        CTkLabel(info_frame, text="ยอดที่ต้องชำระ:", font=CTkFont(weight="bold")).grid(row=1, column=0, sticky="w")
        CTkLabel(info_frame, text=f"{actual_grand_total:,.2f} บาท").grid(row=1, column=1, sticky="w", padx=5)
        # <<< END >>>

        CTkLabel(info_frame, text="ยอดโอนขาดปัจจุบัน:", font=CTkFont(weight="bold"), text_color="#D97706").grid(row=2, column=0, sticky="w")
        CTkLabel(info_frame, text=f"{abs(original_difference):,.2f} บาท", text_color="#D97706").grid(row=2, column=1, sticky="w", padx=5)

        # --- Payment Entries ---
        payment_frame = CTkFrame(self)
        payment_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        payment_frame.grid_columnconfigure(1, weight=1)
        
        CTkLabel(payment_frame, text="ยอดโอนชำระ 1:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.payment1_entry = NumericEntry(payment_frame)
        # แสดงยอดโอนเดิมในช่องแรก
        self.payment1_entry.insert(0, f"{original_payment:,.2f}")
        self.payment1_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        CTkLabel(payment_frame, text="ยอดโอนชำระ 2 (เพิ่มเติม):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.payment2_entry = NumericEntry(payment_frame, placeholder_text="กรอกยอดที่โอนเพิ่ม...")
        self.payment2_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # --- Buttons ---
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=20, pady=15, sticky="ew")
        button_frame.grid_columnconfigure((0,1), weight=1)
        
        CTkButton(button_frame, text="ยกเลิก", fg_color="gray", command=self.destroy).grid(row=0, column=0, padx=(0,5), sticky="ew")
        CTkButton(button_frame, text="บันทึกยอดชำระ", command=self._save_payment_update).grid(row=0, column=1, padx=(5,0), sticky="ew")

        self.transient(master)
        self.grab_set()

    def _save_payment_update(self):
        try:
            p1 = utils.convert_to_float(self.payment1_entry.get())
            p2 = utils.convert_to_float(self.payment2_entry.get())
            new_total_payment = p1 + p2
            
            # <<< START: แก้ไข Logic การคำนวณให้แม่นยำ >>>
            # 1. คำนวณยอดที่ต้องชำระที่แท้จริง
            original_payment = self.so_data.get('total_payment_amount', 0.0) or 0.0
            original_difference = self.so_data.get('difference_amount', 0.0) or 0.0
            actual_grand_total = original_payment - original_difference
            
            # 2. คำนวณส่วนต่างใหม่ (ยอดโอนใหม่ - ยอดที่ต้องชำระ)
            new_difference = new_total_payment - actual_grand_total
            # <<< END >>>
            
            if not messagebox.askyesno("ยืนยัน", "คุณต้องการอัปเดตยอดชำระเงินใช่หรือไม่?", parent=self):
                return
            
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("""
                    UPDATE commissions SET total_payment_amount = %s, difference_amount = %s
                    WHERE id = %s
                """, (new_total_payment, new_difference, self.so_id))
            conn.commit()

            messagebox.showinfo("สำเร็จ", "อัปเดตยอดชำระเรียบร้อยแล้ว", parent=self)
            
            if self.on_save_callback:
                self.on_save_callback()
            self.destroy()

        except Exception as e:
            if 'conn' in locals() and conn: conn.rollback()
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถบันทึกได้: {e}", parent=self)
        finally:
            if 'conn' in locals() and conn: self.app_container.release_connection(conn)

class SalesTasksWindow(CTkToplevel):
    def __init__(self, master, app_container, sale_key):
        super().__init__(master)
        self.commission_app = master
        self.app_container = app_container
        self.sale_key = sale_key
        
        self.title("งานของฉัน (My Tasks)")
        self.geometry("900x600")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        CTkLabel(header, text="งานของฉัน", font=CTkFont(size=18, weight="bold")).pack(side="left")
        CTkButton(header, text="Refresh", command=self.load_tasks, width=80).pack(side="right")

        self.task_tab_view = CTkTabview(self, corner_radius=10)
        self.task_tab_view.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        # --- START: เพิ่มแท็บใหม่สำหรับยอดค้างชำระ ---
        self.payment_due_tab = self.task_tab_view.add("⚠️ SO ค้างชำระ (แก้ไขยอดโอน)")
        # --- END ---
        
        self.rejected_tab = self.task_tab_view.add("งานที่ถูกตีกลับ (Rejected)")
        self.draft_tab = self.task_tab_view.add("ฉบับร่าง (ยังไม่นำส่ง)")
        
        # --- START: เพิ่ม Frame สำหรับแท็บใหม่ ---
        self.payment_due_frame = CTkScrollableFrame(self.payment_due_tab, label_text="รายการที่ยอดโอนชำระไม่ครบ (แก้ไขยอดโอนแล้วกดบันทึก)")
        self.payment_due_frame.pack(fill="both", expand=True, padx=5, pady=5)
        # --- END ---

        self.rejected_frame = CTkScrollableFrame(self.rejected_tab, label_text="รายการที่ต้องแก้ไข")
        self.rejected_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.draft_frame = CTkScrollableFrame(self.draft_tab, label_text="ดับเบิลคลิกรายการเพื่อแก้ไข/ทำต่อ")
        self.draft_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.after(50, self.load_tasks)
        self.transient(master)
        self.grab_set()
    
    def _load_payment_due_tasks(self):
        """(ฉบับแก้ไข) โหลด SO ที่มียอดโอนขาดเท่านั้น"""
        for widget in self.payment_due_frame.winfo_children():
            widget.destroy()

        try:
            # ✅ Query เฉพาะรายการที่ difference_amount > 0 (โอนขาด)
            query = """
                SELECT * FROM commissions 
                WHERE sale_key = %s 
                AND difference_amount < 0
                AND is_active = 1
                AND status != 'Paid'
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))

            if df.empty:
                CTkLabel(self.payment_due_frame, text="ไม่พบรายการโอนขาด").pack(pady=20)
                return

            for _, row_data in df.iterrows():
                total_payment = row_data.get('total_payment_amount', 0.0) or 0.0
                difference = row_data.get('difference_amount', 0.0) or 0.0

                # ✅ แสดงเฉพาะยอดโอนขาด
                card_color = "#FFFBEB"  # สีเหลือง
                balance_text = f"ยอดโอนขาด: {difference:,.2f} บาท"
                text_color = "#B45309"
                button_text = "แก้ไขยอดชำระ"

                card = CTkFrame(self.payment_due_frame, border_width=1, fg_color=card_color)
                card.pack(fill="x", padx=5, pady=4)
                card.grid_columnconfigure(0, weight=1)

                top_frame = CTkFrame(card, fg_color="transparent")
                top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 0))

                info = f"SO: {row_data['so_number']} | ลูกค้า: {row_data['customer_name']} | สถานะปัจจุบัน: {row_data['status']}"
                CTkLabel(top_frame, text=info, font=CTkFont(size=14, weight="bold")).pack(side="left")

                edit_button = CTkButton(
                    top_frame,
                    text=button_text,
                    width=120,
                    command=lambda r=row_data.to_dict(): self._open_payment_updater(r)
                )
                edit_button.pack(side="right")

                reason_label = CTkLabel(
                    card,
                    text=balance_text,
                    text_color=text_color,
                    wraplength=700,
                    justify="left",
                    anchor="w"
                )
                reason_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))

        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดรายการค้างชำระได้: {e}", parent=self)
    
    def _open_payment_updater(self, so_data_dict):
        """(ฟังก์ชันใหม่) เปิดหน้าต่างสำหรับอัปเดตยอดชำระโดยเฉพาะ"""
        PaymentUpdateWindow(
            master=self, 
            app_container=self.app_container, 
            so_data=so_data_dict, 
            on_save_callback=self.load_tasks
        )

    def on_close(self):
        self.commission_app.tasks_window = None
        self.destroy()

    def load_tasks(self):
        self._load_payment_due_tasks() # <-- เพิ่มบรรทัดนี้
        self._load_rejected_tasks()
        self._load_draft_tasks()

    def _load_rejected_tasks(self):
        for widget in self.rejected_frame.winfo_children(): widget.destroy()
        try:
            # <<< START: แก้ไข Query ให้รองรับสถานะ Deferred ด้วย >>>
            query = """
                SELECT * FROM commissions 
                WHERE sale_key = %s AND status IN ('Rejected by SM', 'Rejected by HR', 'Deferred by SM', 'Deferred by HR', 'Draft') AND is_active = 1 
                ORDER BY timestamp DESC
            """
            # <<< END >>>
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))
            if df.empty:
                CTkLabel(self.rejected_frame, text="ไม่มีงานที่ถูกตีกลับหรือถูกเลื่อน").pack(pady=20)
                return
            
            for _, row in df.iterrows():
                # <<< START: แก้ไขการแสดงผล Card ให้แยกสีและข้อความ >>>
                status = row['status']
                if 'Reject' in status:
                    card_color = "#FEF2F2" # แดงอ่อน
                    reason_color = "#B91C1C"
                    status_prefix = "ตีกลับ"
                elif 'Defer' in status: # Deferred
                    card_color = "#FEFCE8" # เหลืองอ่อน
                    reason_color = "#A16207"
                    status_prefix = "เลื่อน"
                else: # ถ้าไม่ใช่ Reject หรือ Defer ก็ต้องเป็น Draft
                    card_color = "#FEF2F2" # แดงอ่อน (เหมือนตีกลับ)
                    reason_color = "#B91C1C"
                    status_prefix = "ตีกลับ (โดย Manager)"
                
                card = CTkFrame(self.rejected_frame, border_width=1, fg_color=card_color)
                # ... (โค้ดสร้าง Card ส่วนที่เหลือเหมือนเดิม) ...
                card.pack(fill="x", padx=5, pady=4)
                card.grid_columnconfigure(0, weight=1)

                top_frame = CTkFrame(card, fg_color="transparent")
                top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(5,0))
                
                info = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}"
                CTkLabel(top_frame, text=info, font=CTkFont(size=14, weight="bold")).pack(side="left")

                edit_button = CTkButton(top_frame, text="แก้ไข", width=80, command=lambda r=row: self._edit_and_close(r))
                edit_button.pack(side="right")
                
                rejected_by = "SM" if "SM" in status else "HR"
                reason_text = row['rejection_reason'] or "ไม่มีเหตุผลระบุ"
                reason_label = CTkLabel(card, text=f"เหตุผลที่{status_prefix} (จาก {rejected_by}): {reason_text}", text_color=reason_color, wraplength=700, justify="left", anchor="w")
                reason_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))
                # <<< END >>>

        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดงานที่ถูกตีกลับได้: {e}", parent=self)

    def _load_draft_tasks(self):
        for widget in self.draft_frame.winfo_children(): widget.destroy()
        try:
            # VVVV เพิ่มบรรทัดนี้เข้าไปเพื่อ Debug VVVV
            print(f"--- DEBUG: กำลังโหลดฉบับร่างสำหรับ sale_key: {self.sale_key} ---")
            
            query = """
                SELECT *
                FROM commissions 
                WHERE sale_key = %s AND status IN ('Original', 'Edited') AND is_active = 1 
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))
            if df.empty:
                CTkLabel(self.draft_frame, text="ไม่มีฉบับร่าง").pack(pady=20)
                return
            
            for _, row in df.iterrows():
                card = CTkFrame(self.draft_frame, border_width=1)
                card.pack(fill="x", padx=5, pady=3)
                info = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}"
                CTkLabel(card, text=info).pack(side="left", padx=10, pady=5)

                edit_callback = lambda e, r=row: self._edit_and_close(r)
                card.bind("<Double-1>", edit_callback)
                for child in card.winfo_children(): child.bind("<Double-1>", edit_callback)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดฉบับร่างได้: {e}", parent=self)
            
    def _edit_and_close(self, row_data):
        self.commission_app._edit_history_item(row_data.to_dict())
        self.on_close()

class SubmitSODialog(CTkToplevel):
    def __init__(self, master, app_container, sale_key, sale_name):
        super().__init__(master)
        self.commission_app = master
        self.app_container = app_container
        self.sale_key = sale_key
        self.sale_name = sale_name
        self.checkbox_list = []

        self.title("เลือกรายการ SO ที่จะนำส่ง")
        self.geometry("700x500")
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Top Frame for Select All ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=15, pady=(10, 0), sticky="ew")
        self.select_all_var = tk.IntVar(value=0)
        self.select_all_checkbox = CTkCheckBox(top_frame, text="เลือกทั้งหมด", variable=self.select_all_var, command=self._toggle_all_checkboxes, font=CTkFont(weight="bold"))
        self.select_all_checkbox.pack(anchor="w")

        # --- Scrollable Frame for SO List ---
        self.scroll_frame = CTkScrollableFrame(self, label_text="รายการ SO ที่เป็นฉบับร่าง")
        self.scroll_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        
        # --- Bottom Frame for Buttons ---
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        button_frame.grid_columnconfigure((0,1), weight=1)
        
        self.submit_button = CTkButton(button_frame, text="ยืนยันการนำส่ง (0)", command=self._confirm_submission, state="disabled")
        self.submit_button.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        CTkButton(button_frame, text="ยกเลิก", fg_color="gray", command=self.destroy).grid(row=0, column=1, padx=(5,0), sticky="ew")

        self.after(50, self._populate_so_list)
        self.transient(master)
        self.grab_set()

    def _populate_so_list(self):
        try:
            query = "SELECT id, so_number, customer_name FROM commissions WHERE sale_key = %s AND status IN ('Original', 'Edited') AND is_active = 1 ORDER BY timestamp DESC"
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))

            if df.empty:
                CTkLabel(self.scroll_frame, text="ไม่พบรายการที่เป็นฉบับร่าง").pack(pady=20)
                self.select_all_checkbox.configure(state="disabled")
                return

            for _, row in df.iterrows():
                checkbox_var = tk.IntVar(value=0)
                checkbox_var.trace_add("write", self._update_submit_button_state)
                
                so_id = row['id']
                so_text = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}"
                
                cb = CTkCheckBox(self.scroll_frame, text=so_text, variable=checkbox_var)
                cb.pack(anchor="w", padx=10, pady=5)
                self.checkbox_list.append((checkbox_var, so_id, row.to_dict()))

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดรายการ SO ได้: {e}", parent=self)
            self.destroy()

    def _toggle_all_checkboxes(self):
        is_selected = self.select_all_var.get()
        for var, _, _ in self.checkbox_list:
            var.set(is_selected)

    def _update_submit_button_state(self, *args):
        selected_count = sum(var.get() for var, _, _ in self.checkbox_list)
        self.submit_button.configure(text=f"ยืนยันการนำส่ง ({selected_count})")
        if selected_count > 0:
            self.submit_button.configure(state="normal")
        else:
            self.submit_button.configure(state="disabled")

    def _confirm_submission(self):
        selected_records = [(so_id, record_data) for var, so_id, record_data in self.checkbox_list if var.get() == 1]
        
        if not selected_records:
            messagebox.showwarning("ยังไม่ได้เลือก", "กรุณาเลือก SO อย่างน้อย 1 รายการ", parent=self)
            return

        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการนำส่ง SO จำนวน {len(selected_records)} รายการใช่หรือไม่?", parent=self):
            return
            
        selected_ids = [so_id for so_id, _ in selected_records]
        
        # เพิ่มบรรทัดนี้: สร้าง tuple ของ IDs จาก selected_ids
        ids_tuple = tuple(selected_ids) 

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                new_status = 'Pending PU'
                # บรรทัดนี้ควรถูกลบออก หรือเป็นคอมเมนต์ ถ้าไม่ใช่การใช้ execute_values
                # psycopg2.extras.execute_values 
                cursor.execute(
                    f"UPDATE commissions SET status = '{new_status}' WHERE id IN %s",
                    (ids_tuple,) 
                )
                
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Staff' AND status = 'Active'")
                pu_keys = [row[0] for row in cursor.fetchall()]

                notif_data = []
                for _, record_data in selected_records:
                    message = f"SO ใหม่รอสร้าง PO: {record_data['so_number']}"
                    for pu_key in pu_keys:
                        notif_data.append((pu_key, message, False, record_data['id']))
                
                psycopg2.extras.execute_values(
                    cursor,
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES %s",
                    notif_data
                )
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", f"นำส่งข้อมูลจำนวน {len(selected_ids)} รายการไปยังฝ่ายจัดซื้อเรียบร้อยแล้ว", parent=self.commission_app)
            
            self.commission_app._update_tasks_badge()
            self.commission_app._refresh_history_if_open()
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการนำส่งข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

class CommissionApp(CTkFrame):
    def __init__(self, master, sale_key=None, sale_name=None, app_container=None, show_logout_button=True, user_role=None, create_default_header=True):
        super().__init__(master, corner_radius=0, fg_color=app_container.THEME["sale"]["bg"])
        self.master = master
        self.app_container = app_container
        self.sale_key = sale_key or "UNKNOWN_SALE_KEY"
        self.sale_name = sale_name or "Unknown Sales User"
        self.user_role = user_role # <--- เพิ่มบรรทัดนี้
        self.theme = app_container.THEME["sale"]
        self.pg_engine = app_container.pg_engine
        self.show_logout_button = show_logout_button

        self.dropdown_style = {
            "fg_color": "white",
            "text_color": "black",
            "button_color": self.theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }

        self.editing_record_id = None
        self.history_window = None
        self.customer_data = {}
        self.customer_codes = []
        self.customer_completion_data = [] 
        self.so_form_widgets = {}
        self.header_map = app_container.HEADER_MAP

        self._create_string_vars()
        self.tasks_window = None
        self.tasks_button = None
        self.polling_job_id = None
        
        # 1. กำหนด Layout หลักของ Frame นี้ก่อน
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 2. สร้าง Header และวางในแถวที่ 0 (ถ้าจำเป็น)
        if create_default_header:
            self._create_header()

        # 3. สร้าง Scrollable Container และวางในแถวที่ 1
        self.scrollable_main_container = CTkScrollableFrame(self, fg_color="transparent")
        self.scrollable_main_container.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # 4. กำหนด Layout ภายใน Scrollable Container
        self.scrollable_main_container.grid_columnconfigure(0, weight=1, uniform="group1")
        self.scrollable_main_container.grid_columnconfigure(1, weight=1, uniform="group1")
        self.scrollable_main_container.grid_rowconfigure(0, weight=1)

        # 5. สร้าง Frame ซ้าย-ขวา และวางใน Scrollable Container
        self.left_frame = CTkFrame(self.scrollable_main_container, fg_color="transparent")
        self.left_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")

        self.right_frame = CTkFrame(self.scrollable_main_container, fg_color="transparent")
        self.right_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")

        # 6. สร้างฟอร์มทั้งหมดลงใน Frame ซ้าย-ขวา
        self._populate_all_forms()
        
        # 7. โหลดข้อมูลและผูก Event ต่างๆ
        self._load_customer_data()
        self._bind_events()
        
        # 8. เริ่มการทำงานเบื้องหลัง
        self._start_polling()
        self.bind("<Destroy>", self._on_destroy)

    def _open_my_tasks_window(self):
        if self.tasks_window is None or not self.tasks_window.winfo_exists():
            self.tasks_window = SalesTasksWindow(self, app_container=self.app_container, sale_key=self.sale_key)
        else:
            self.tasks_window.focus()

    def _update_tasks_badge(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM commissions WHERE sale_key = %s AND status = 'Rejected by SM' AND is_active = 1", (self.sale_key,))
                rejected_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM commissions WHERE sale_key = %s AND status IN ('Original', 'Edited') AND is_active = 1", (self.sale_key,))
                draft_count = cursor.fetchone()[0]
            
            total_tasks = rejected_count + draft_count
            
            if hasattr(self, 'tasks_button') and self.tasks_button and self.tasks_button.winfo_exists():
                self.tasks_button.configure(text=f"งานของฉัน 🔔 ({total_tasks})")
                if total_tasks > 0:
                    self.tasks_button.configure(fg_color="#F59E0B", hover_color="#D97706")
                else:
                    self.tasks_button.configure(fg_color=("#3B8ED0", "#1F6AA5"), hover_color=("#36719F", "#144870"))
        except Exception as e:
            print(f"Error updating tasks badge: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _start_polling(self):
        self._update_tasks_badge()
        self.polling_job_id = self.after(30000, self._start_polling)

    def _on_destroy(self, event):
        if hasattr(event, 'widget') and event.widget is self:
            if self.polling_job_id:
                self.after_cancel(self.polling_job_id)

    def _create_header(self):
        self.header_frame = CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        # --- START: แก้ไขส่วนนี้ทั้งหมด ---
        # กำหนดให้คอลัมน์ซ้าย (ชื่อ) ขยายตัว และคอลัมน์ขวา (ปุ่ม) ไม่ขยาย
        self.header_frame.grid_columnconfigure(0, weight=1)
        
        CTkLabel(self.header_frame, text=f"ฝ่ายขาย: {self.sale_name} ({self.sale_key})", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).grid(row=0, column=0, sticky="w")
        
        # สร้าง Frame สำหรับปุ่มต่างๆ ทางขวา และใช้ grid วาง
        button_container = CTkFrame(self.header_frame, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="e")
        
        # ใช้ grid วางปุ่มภายใน button_container
        self.tasks_button = CTkButton(button_container, text="งานของฉัน 🔔 (0)", command=self._open_my_tasks_window)
        self.tasks_button.grid(row=0, column=0, padx=10)

        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").grid(row=0, column=1, padx=(0, 10))
        
        if self.show_logout_button:
            CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").grid(row=0, column=1, padx=(0, 10))
    
    
    def _refresh_history_if_open(self):
        if self.history_window and self.history_window.winfo_exists():
            if hasattr(self.history_window, '_populate_history_table'):
                self.history_window._populate_history_table()

    def _edit_history_item(self, row_data):
        record_status = row_data.get('status')

        # --- START: แก้ไขเงื่อนไขการตรวจสอบสถานะตรงนี้ ---
        # สถานะที่ "ห้ามแก้ไขเด็ดขาด" คือเมื่อ HR ตรวจสอบแล้ว หรือจ่ายเงินไปแล้ว
        uneditable_statuses = ('Paid', 'HR Verified', 'Cancelled')

        if record_status in uneditable_statuses:
            messagebox.showwarning("ไม่สามารถแก้ไขได้", 
                                 f"รายการนี้มีสถานะ '{record_status}' ซึ่งเป็นขั้นตอนสุดท้ายแล้ว จึงไม่สามารถแก้ไขได้", 
                                 parent=self)
            return
        # --- END ---

        self._clear_form(confirm=False)
        self.editing_record_id = int(row_data.get('id'))
        if self.history_window and self.history_window.winfo_exists():
            self.history_window.destroy()
            self.history_window = None
        if self.tasks_window and self.tasks_window.winfo_exists():
             self.tasks_window.destroy()
             self.tasks_window = None
        self._populate_form_from_data(row_data)
        messagebox.showinfo("โหลดข้อมูลเพื่อแก้ไข", f"โหลด SO Number: {row_data.get('so_number')} สำหรับแก้ไข", parent=self)

    def _on_history_so_select(self, row_data):
        """
        Callback function เมื่อมีการดับเบิลคลิกที่ SO ในหน้าต่างประวัติ
        (เวอร์ชันแก้ไข: เพิ่มการตั้งค่า editing_record_id ให้ถูกต้อง)
        """
        if row_data is None:
            return

        so_status = row_data.get('status')
        so_number = row_data.get('so_number')
        
        editable_statuses = ['Original', 'Edited', 'Rejected by SM', 'Rejected by HR', 'Deferred by SM', 'Deferred by HR']

        if so_status in editable_statuses:
            if messagebox.askyesno("โหลดข้อมูล", f"คุณต้องการโหลดข้อมูล SO: {so_number} เพื่อแก้ไขหรือไม่?"):
                
                # <<< START: แก้ไข Logic ทั้งหมดตรงนี้ >>>
                # 1. ล้างข้อมูลในฟอร์มให้เป็นค่าเริ่มต้นก่อน เพื่อป้องกันข้อมูลเก่าค้าง
                self._clear_form(confirm=False)
                
                # 2. (สำคัญที่สุด) ตั้งค่า ID ที่กำลังแก้ไข เพื่อให้โปรแกรมเข้าสู่ "โหมดแก้ไข"
                self.editing_record_id = int(row_data.get('id'))
                
                # 3. โหลดข้อมูลจากแถวที่เลือกมาใส่ในฟอร์ม
                self._populate_form_from_data(row_data.to_dict()) 
                
                # 4. ปิดหน้าต่างประวัติ (เหมือนเดิม)
                if self.history_window and self.history_window.winfo_exists():
                    self.history_window.destroy()
                # <<< END: สิ้นสุดการแก้ไข >>>
        else:
            messagebox.showinfo(
                "ไม่สามารถแก้ไขได้",
                f"SO: {so_number} อยู่ในสถานะ '{so_status}'\n\n"
                "ข้อมูลได้ถูกส่งต่อไปในกระบวนการแล้ว จึงไม่สามารถแก้ไขได้ในขั้นตอนนี้",
                parent=self.history_window
            )

    def _save_data(self):
        form_data = self._gather_data_from_form()
        is_valid, message = self._validate_form(form_data)
        if not is_valid:
            messagebox.showerror("ข้อมูลไม่ถูกต้อง", message, parent=self)
            return
        if self.editing_record_id:
            if not messagebox.askyesno("ยืนยันการแก้ไข", "คุณต้องการบันทึกการเปลี่ยนแปลงนี้ใช่หรือไม่?", parent=self):
                return
        if form_data['customer_type'] == "ลูกค้าใหม่" and not self.editing_record_id:
            try:
                self._handle_new_customer(form_data)
            except Exception as e:
                messagebox.showerror("ผิดพลาด", f"ไม่สามารถเพิ่มลูกค้าใหม่ได้:\n{e}", parent=self)
                return
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                if self.editing_record_id:
                    cursor.execute("UPDATE commissions SET is_active = 0 WHERE id = %s", (self.editing_record_id,))
                    form_data['status'] = 'Edited'
                    form_data['original_id'] = self.editing_record_id
                self._perform_db_insert(form_data)
            conn.commit()
            action = "อัปเดต" if self.editing_record_id else "บันทึก"
            messagebox.showinfo("สำเร็จ", f"{action}ข้อมูลเรียบร้อยแล้ว", parent=self)
            self._clear_form(confirm=False)
            self._load_customer_data()
            self._refresh_history_if_open()
            self._update_tasks_badge()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_section_frame(self, parent, title):
        frame = CTkFrame(parent, corner_radius=10, border_width=1)
        frame.grid_columnconfigure(1, weight=1)
        CTkLabel(frame, text=title, font=CTkFont(size=18, weight="bold"), text_color=self.theme["header"]).grid(row=0, column=0, columnspan=3, padx=15, pady=(10, 5), sticky="w")
        return frame

    def _add_form_row(self, parent, label_text, widget, row_index, columnspan=2, padx=(10, 15), pady=4):
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=(15, 10), pady=pady, sticky="w")
        widget.grid(row=row_index, column=1, columnspan=columnspan, padx=padx, pady=pady, sticky="ew")

    def _add_item_row_with_vat(self, parent, label_text, entry_widget, radio_var, row_index, padx=(10, 15), pady=5):
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=15, pady=pady, sticky="w")
        entry_widget.grid(row=row_index, column=1, padx=padx, pady=pady, sticky="ew")

        radio_frame = CTkFrame(parent, fg_color="transparent")
        radio_frame.grid(row=row_index, column=2, padx=padx, pady=pady, sticky="w")
        CTkRadioButton(radio_frame, text="VAT", variable=radio_var, value="VAT").pack(side="left", padx=5)
        CTkRadioButton(radio_frame, text="CASH", variable=radio_var, value="NO VAT").pack(side="left", padx=5)

    def _populate_all_forms(self):
        self._populate_sales_details_form(self.left_frame)
        self._populate_sales_services_frame(self.left_frame)
        self._populate_shipping_frame(self.left_frame)
        self._populate_fees_frame(self.left_frame)

        self._populate_other_expenses_frame(self.right_frame)
        self._populate_payment_frame(self.right_frame)
        self._populate_so_summary_frame(self.right_frame)
        self._populate_cash_verification_frame(self.right_frame)
        self._populate_action_frame(self.right_frame)

    def _create_string_vars(self):
        now = datetime.now()
        thai_months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_months = thai_months_list
        self.thai_month_map = {name: i+1 for i, name in enumerate(thai_months_list)}

        self.customer_type_var = tk.StringVar(value="ลูกค้าเก่า")
        self.credit_term_var = tk.StringVar(value="เงินสด")
        self.commission_month_var = tk.StringVar(value=self.thai_months[now.month - 1])
        self.commission_year_var = tk.StringVar(value=str(now.year + 543))
        self.payment1_percent_var = tk.StringVar(value="ระบุยอดเอง")
        self.payment2_percent_var = tk.StringVar(value="ระบุยอดเอง")
        self.sales_vat_calc_var = tk.StringVar(value="0.00")
        self.cutting_drilling_vat_calc_var = tk.StringVar(value="0.00")
        self.other_service_vat_calc_var = tk.StringVar(value="0.00")
        self.shipping_vat_calc_var = tk.StringVar(value="0.00")
        self.card_fee_vat_calc_var = tk.StringVar(value="0.00")
        self.relocation_vat_option_var = tk.StringVar(value="VAT")
        self.relocation_vat_calc_var = tk.StringVar(value="0.00")
        self.payment_total_var = tk.StringVar(value="0.00")
        self.so_subtotal_var = tk.StringVar(value="0.00")
        self.so_vat_var = tk.StringVar(value="0.00")
        self.so_grand_total_var = tk.StringVar(value="0.00")
        self.so_vs_payment_result_var = tk.StringVar(value="-")

        self.so_number_var = tk.StringVar(value="SO")

        self.difference_amount_var = tk.StringVar(value="0.00")
        self.balance_due_var = tk.StringVar(value="0.00")
        self.cash_product_input_var = tk.StringVar(value="0.00")
        self.cash_service_total_var = tk.StringVar(value="0.00")
        self.cash_required_total_var = tk.StringVar(value="0.00")
        self.cash_actual_payment_var = tk.StringVar(value="0.00")
        self.cash_verification_result_var = tk.StringVar(value="-")
        self.sales_service_vat_option = tk.StringVar(value="VAT")
        self.cutting_drilling_fee_vat_option = tk.StringVar(value="VAT")
        self.other_service_fee_vat_option = tk.StringVar(value="VAT")
        self.shipping_vat_option_var = tk.StringVar(value="VAT")
        self.credit_card_fee_vat_option_var = tk.StringVar(value="VAT")

        self.payment1_method_var = tk.StringVar(value="ชำระสด")
        self.payment2_method_var = tk.StringVar(value="ชำระสด")
        self.delivery_type_var = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")

    def _force_uppercase_so_number(self, *args):
        current_text = self.so_number_var.get()
        new_text = current_text.upper()
        if new_text != current_text:
            self.so_number_var.set(new_text)

    def _on_delivery_type_change(self, *args):
        selected_option = self.delivery_type_var.get()

    # ตรวจสอบว่าวิดเจ็ตถูกสร้างขึ้นแล้วหรือยังก่อนที่จะใช้งาน
        if hasattr(self, 'date_to_wh_label') and self.date_to_wh_label.winfo_exists() and \
           hasattr(self, 'date_to_wh_selector') and self.date_to_wh_selector.winfo_exists():

         if "คลัง" in selected_option:
        # ถ้ามีคำว่า "คลัง" ในตัวเลือก ให้แสดงช่องวันที่
           self.date_to_wh_label.grid()
           self.date_to_wh_selector.grid()
         else:
        # ถ้าไม่มี ให้ซ่อน
           self.date_to_wh_label.grid_remove()
           self.date_to_wh_selector.grid_remove()

    def _bind_events(self):
        widgets_to_bind_names = [
            "sales_amount_entry", "cutting_drilling_fee_entry", "other_service_fee_entry",
            "shipping_cost_entry", "credit_card_fee_entry", "transfer_fee_entry",
            "wht_fee_entry", "coupon_value_entry", "giveaway_value_entry",
            "brokerage_fee_entry", "payment1_amount_entry", "payment2_amount_entry",
            "cash_product_input_entry", "cash_actual_payment_entry",
            # <<< เพิ่มเติม: เพิ่ม relocation_cost_entry เข้าไปใน list นี้ >>>
            "relocation_cost_entry"
        ]

        for widget_name in widgets_to_bind_names:
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                widget.bind("<KeyRelease>", self._update_final_calculations)

        for var in [
            self.sales_service_vat_option, self.cutting_drilling_fee_vat_option,
            self.other_service_fee_vat_option, self.shipping_vat_option_var,
            self.credit_card_fee_vat_option_var,self.relocation_vat_option_var
        ]:
            var.trace_add("write", self._update_final_calculations)

        self.so_number_var.trace_add("write", self._force_uppercase_so_number)
        self.so_number_entry.bind("<FocusIn>", lambda e: self.so_number_entry.icursor(tk.END))

        self._on_delivery_type_change()

    def _show_history(self):
        try:
            # vvvv แก้ไขบรรทัดนี้ vvvv
            if self.user_role and self.user_role.lower() == 'sale support':
                self.history_window = self.app_container.show_history_window(
                    support_user_key_filter=self.app_container.current_user_key,
                    sale_key_filter=None,
                    edit_callback=self._on_history_so_select
                )
            else:
                self.history_window = self.app_container.show_history_window(
                    sale_key_filter=self.sale_key,
                    support_user_key_filter=None,
                    edit_callback=self._on_history_so_select
                )
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถเปิดหน้าต่างประวัติได้: {e}", parent=self)

    def _populate_form_from_data(self, data):
        def set_entry_value(entry_widget, value):
            if isinstance(entry_widget, (NumericEntry, CTkEntry, AutoCompleteEntry)):
                is_readonly = entry_widget.cget("state") == "readonly"
                if is_readonly: entry_widget.configure(state="normal")
                entry_widget.delete(0, tk.END)
                if pd.notna(value) and value is not None:
                    if isinstance(value, (int, float)):
                        entry_widget.insert(0, f"{value:,.2f}")
                    else:
                        entry_widget.insert(0, str(value))
                if is_readonly: entry_widget.configure(state="readonly")

        def set_date_selector(selector_widget, date_str):
            if pd.notna(date_str) and date_str is not None:
                try:
                    if isinstance(date_str, datetime): dt_obj = date_str
                    elif isinstance(date_str, pd.Timestamp): dt_obj = date_str.to_pydatetime()
                    else: dt_obj = datetime.strptime(str(date_str), '%Y-%m-%d')
                    selector_widget.set_date(dt_obj)
                except (ValueError, TypeError):
                    selector_widget.set_date(None)
            else:
                selector_widget.set_date(None)

        def set_radio_button(radio_var, value, default="VAT"):
            if pd.notna(value) and value is not None:
                radio_var.set(str(value))
            else:
                radio_var.set(default)

        set_date_selector(self.bill_date_selector, data.get('bill_date'))
        month_from_data = data.get('commission_month')
        if pd.notna(month_from_data):
            try:
                month_int = int(utils.convert_to_float(month_from_data))
                if 1 <= month_int <= 12: self.commission_month_var.set(self.thai_months[month_int - 1])
            except (ValueError, TypeError): pass

        year_from_data = data.get('commission_year')
        if pd.notna(year_from_data):
            try:
                year_int = int(utils.convert_to_float(year_from_data))
                self.commission_year_var.set(str(year_int + 543))
            except (ValueError, TypeError): pass

        customer_type = data.get('customer_type', 'ลูกค้าเก่า')
        self.customer_type_var.set(customer_type)
        self._toggle_customer_fields()

        if customer_type == "ลูกค้าเก่า":
            set_entry_value(self.customer_id_entry, data.get('customer_id'))
            self._on_customer_id_selected(data.get('customer_id'))
        else:
            set_entry_value(self.new_customer_id_entry, data.get('customer_id'))
            set_entry_value(self.new_customer_name_entry, data.get('customer_name'))

        self.credit_term_var.set(data.get('credit_term', 'เงินสด'))
        self.so_number_var.set(data.get('so_number', 'SO'))

        set_entry_value(self.sales_amount_entry, data.get('sales_service_amount'))
        set_radio_button(self.sales_service_vat_option, data.get('sales_service_vat_option'))
        set_entry_value(self.cutting_drilling_fee_entry, data.get('cutting_drilling_fee'))
        set_radio_button(self.cutting_drilling_fee_vat_option, data.get('cutting_drilling_fee_vat_option'))
        set_entry_value(self.other_service_fee_entry, data.get('other_service_fee'))
        set_radio_button(self.other_service_fee_vat_option, data.get('other_service_fee_vat_option'))
        set_entry_value(self.shipping_cost_entry, data.get('shipping_cost'))
        set_radio_button(self.shipping_vat_option_var, data.get('shipping_vat_option'))
        set_date_selector(self.delivery_date_selector, data.get('delivery_date'))
        set_entry_value(self.credit_card_fee_entry, data.get('credit_card_fee'))
        set_radio_button(self.credit_card_fee_vat_option_var, data.get('credit_card_fee_vat_option'))
        set_entry_value(self.transfer_fee_entry, data.get('transfer_fee'))
        set_entry_value(self.wht_fee_entry, data.get('wht_3_percent'))
        set_entry_value(self.brokerage_fee_entry, data.get('brokerage_fee'))
        set_entry_value(self.coupon_value_entry, data.get('coupons'))
        set_entry_value(self.giveaway_value_entry, data.get('giveaways'))
        set_radio_button(self.relocation_vat_option_var, data.get('relocation_cost_vat_option'))


        payment1 = data.get('total_payment_amount', 0.0) # Simplified for now
        set_entry_value(self.payment1_amount_entry, payment1)
        self.payment1_percent_var.set("ระบุยอดเอง")
        self.payment2_amount_entry.delete(0, tk.END)
        self.payment2_percent_var.set("ระบุยอดเอง")

        set_date_selector(self.payment1_date_selector, data.get('payment1_date'))
        set_date_selector(self.payment2_date_selector, data.get('payment2_date'))

        set_entry_value(self.cash_product_input_entry, data.get('cash_product_input'))
        set_entry_value(self.cash_actual_payment_entry, data.get('cash_actual_payment'))

        set_radio_button(self.delivery_type_var, data.get('delivery_type', 'ซัพพลายเออร์จัดส่ง'))
        set_entry_value(self.pickup_location_entry, data.get('pickup_location'))
        set_entry_value(self.relocation_cost_entry, data.get('relocation_cost'))
        set_date_selector(self.date_to_wh_selector, data.get('date_to_warehouse'))
        set_date_selector(self.date_to_customer_selector, data.get('date_to_customer'))
        set_entry_value(self.pickup_rego_entry, data.get('pickup_registration'))


        self.payment1_method_var.set(data.get('payment1_method', 'ไม่เลือก'))
        self.payment2_method_var.set(data.get('payment2_method', 'ไม่เลือก'))

        self._update_final_calculations()

    def _load_customer_data(self):
        try:
            df = pd.read_sql("SELECT customer_name, customer_code, credit_term FROM customers ORDER BY customer_name", self.pg_engine)
            
            # --- สร้างข้อมูลโครงสร้างใหม่ ---
            self.customer_completion_data = []
            for _, row in df.iterrows():
                display_text = f"{row['customer_code']} - {row['customer_name']}"
                self.customer_completion_data.append({
                    "name": row['customer_name'],
                    "code": row['customer_code'],
                    "term": row.get('credit_term', 'เงินสด'),
                    "display": display_text
                })
            
            # สร้าง Map สำหรับการอ้างอิงข้อมูล
            self.customer_data_map = {item['display']: item for item in self.customer_completion_data}

            # <<< START: เพิ่มเติมส่วนที่ขาดไป >>>
            # ตรวจสอบว่า widget ถูกสร้างแล้วหรือยัง ก่อนที่จะอัปเดตข้อมูลเข้าไป
            if hasattr(self, 'customer_id_entry') and self.customer_id_entry.winfo_exists():
                self.customer_id_entry.update_completion_list(self.customer_completion_data)
            # <<< END: สิ้นสุดการเพิ่มเติม >>>

        except Exception as e:
            print(f"Error loading customer data: {e}")
            self.customer_completion_data = []
            self.customer_data_map = {}

    def _on_customer_id_selected(self, selection_data):
        """
        ฟังก์ชันนี้ถูกออกแบบใหม่ให้รองรับข้อมูล 2 รูปแบบ:
        1. Dictionary: เมื่อผู้ใช้เลือกรายการจาก AutoComplete suggestion.
        2. String (customer_code): เมื่อฟังก์ชันถูกเรียกตอนโหลดข้อมูลเก่ามาแสดง (populate_form).
        """
        customer_name = ''
        credit_term = 'เงินสด'
        customer_code = ''

        if isinstance(selection_data, dict):
            # กรณีที่ 1: ผู้ใช้เลือกจาก AutoComplete (ได้ข้อมูลมาเป็น dict)
            customer_name = selection_data.get('name', '')
            credit_term = selection_data.get('term', 'เงินสด')
            customer_code = selection_data.get('code', '')

            self.customer_id_entry.delete(0, tk.END)
            self.customer_id_entry.insert(0, customer_code)

        elif isinstance(selection_data, str) and selection_data:
            # กรณีที่ 2: โหลดข้อมูลเก่า (ได้ข้อมูลมาเป็น string ของ customer_code)
            # เราจะค้นหาข้อมูลลูกค้าทั้งหมดจาก customer_code ที่ได้รับมา
            customer_code_to_find = selection_data
            found_customer = next((item for item in self.customer_completion_data if item.get('code') == customer_code_to_find), None)
            
            if found_customer:
                customer_name = found_customer.get('name', '')
                credit_term = found_customer.get('term', 'เงินสด')

        # --- ส่วนของการอัปเดตหน้าจอ (เหมือนเดิม) ---
        self.customer_name_entry.configure(state="normal")
        self.customer_name_entry.delete(0, tk.END)
        self.customer_name_entry.insert(0, customer_name)
        self.customer_name_entry.configure(state="readonly")
        
        self.credit_term_var.set(credit_term)

    def _populate_sales_details_form(self, parent):
        
        frame = self._create_section_frame(parent, "รายละเอียดการขาย")
        frame.pack(fill="x", pady=(0, 10))
        
        # --- Commission Period is now on top ---
        commission_period_outer_frame = CTkFrame(frame, fg_color="transparent")
        commission_period_outer_frame.grid(row=1, column=0, columnspan=3, padx=15, pady=4, sticky="ew") # <-- Changed to row=1
        CTkLabel(commission_period_outer_frame, text="รอบคอมมิชชั่น:", font=CTkFont(size=14)).pack(side="left")

        month_year_frame = CTkFrame(commission_period_outer_frame, fg_color="transparent")
        month_year_frame.pack(side="left")

        self.commission_month_menu = CTkOptionMenu(month_year_frame, variable=self.commission_month_var, values=list(self.thai_month_map.keys()), width=120, **self.dropdown_style)
        self.commission_month_menu.pack(side="left", padx=5)

        current_year_be = datetime.now().year + 543
        year_list = [str(y) for y in range(current_year_be - 2, current_year_be + 5)]
        self.commission_year_menu = CTkOptionMenu(month_year_frame, variable=self.commission_year_var, values=year_list, width=90, **self.dropdown_style)
        self.commission_year_menu.pack(side="left", padx=5)
        
        # --- Bill Date is now below ---
        bill_date_outer_frame = CTkFrame(frame, fg_color="transparent")
        bill_date_outer_frame.grid(row=2, column=0, columnspan=3, padx=15, pady=4, sticky="ew") # <-- Changed to row=2
        CTkLabel(bill_date_outer_frame, text="วันที่เปิด SO:", font=CTkFont(size=14)).pack(side="left")
        self.bill_date_selector = DateSelector(bill_date_outer_frame, dropdown_style=self.dropdown_style)
        self.bill_date_selector.pack(side="left", padx=5)

        # --- The rest of the function remains the same ---
        customer_type_radio_frame = CTkFrame(frame, fg_color="transparent")
        CTkRadioButton(customer_type_radio_frame, text="ลูกค้าเก่า", variable=self.customer_type_var, value="ลูกค้าเก่า", command=self._toggle_customer_fields).pack(side="left", padx=5)
        CTkRadioButton(customer_type_radio_frame, text="ลูกค้าใหม่", variable=self.customer_type_var, value="ลูกค้าใหม่", command=self._toggle_customer_fields).pack(side="left", padx=20)
        self._add_form_row(frame, "ประเภทลูกค้า:", customer_type_radio_frame, 3)

        self.old_customer_frame = CTkFrame(frame, fg_color="transparent")
        self.old_customer_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=0, pady=0)
        self.old_customer_frame.grid_columnconfigure(1, weight=1)

        # <<< START: แก้ไขการเรียกใช้ AutoCompleteEntry ตรงนี้ >>>
        self.customer_id_entry = AutoCompleteEntry(
            master=self.old_customer_frame, 
            completion_list=self.customer_completion_data,
            display_key='display',  # แก้ไข: เพิ่ม display_key ที่จำเป็น
            command=self._on_customer_id_selected, # แก้ไข: เปลี่ยนชื่อ argument เป็น command
            placeholder_text="ค้นหารหัสหรือชื่อลูกค้า..."
        )
        # <<< END: สิ้นสุดการแก้ไข >>>

        self._add_form_row(self.old_customer_frame, "รหัสลูกค้า:", self.customer_id_entry, 0)

        self.customer_name_entry = CTkEntry(self.old_customer_frame, state="readonly")
        self._add_form_row(self.old_customer_frame, "ชื่อลูกค้า:", self.customer_name_entry, 1)

        self.new_customer_frame = CTkFrame(frame, fg_color="transparent")
        self.new_customer_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=0, pady=0)
        self.new_customer_frame.grid_columnconfigure(1, weight=1)

        self.new_customer_id_entry = CTkEntry(self.new_customer_frame, placeholder_text="กำหนดรหัสลูกค้าใหม่")
        self._add_form_row(self.new_customer_frame, "รหัสลูกค้า (ใหม่):", self.new_customer_id_entry, 0)

        self.new_customer_name_entry = CTkEntry(self.new_customer_frame, placeholder_text="กรอกชื่อลูกค้าใหม่")
        self._add_form_row(self.new_customer_frame, "ชื่อลูกค้า (ใหม่):", self.new_customer_name_entry, 1)

        credit_so_frame = CTkFrame(frame, fg_color="transparent")
        credit_so_frame.grid(row=5, column=0, columnspan=3, padx=15, pady=4, sticky="ew")
        credit_so_frame.grid_columnconfigure((1, 3), weight=1)

        CTkLabel(credit_so_frame, text="Credit:", font=CTkFont(size=14)).grid(row=0, column=0, sticky="w")
        self.credit_term_menu = CTkOptionMenu(credit_so_frame, variable=self.credit_term_var, values=["เงินสด", "CR"], **self.dropdown_style)
        self.credit_term_menu.grid(row=0, column=1, sticky="ew", padx=(10, 20))

        CTkLabel(credit_so_frame, text="เลขที่ใบสั่งขาย:", font=CTkFont(size=14)).grid(row=0, column=2, padx=(20, 10), sticky="w")
        self.so_number_entry = CTkEntry(credit_so_frame, textvariable=self.so_number_var)
        self.so_number_entry.grid(row=0, column=3, sticky="ew")

        self._toggle_customer_fields()

    def _toggle_customer_fields(self):
        if self.customer_type_var.get() == "ลูกค้าเก่า":
            self.new_customer_frame.grid_remove()
            self.old_customer_frame.grid()
            if hasattr(self, 'customer_id_entry'):
                self.customer_id_entry.delete(0, tk.END)
                self.customer_name_entry.configure(state="normal")
                self.customer_name_entry.delete(0, tk.END)
                self.customer_name_entry.configure(state="readonly")
        else:
            self.old_customer_frame.grid_remove()
            self.new_customer_frame.grid()
            self.credit_term_var.set("เงินสด")

    # <<< START: CODE REPLACEMENT >>>
    def _populate_other_expenses_frame(self, parent):
        # <<< แก้ไข: จัด Layout ของ Frame นี้ใหม่ทั้งหมด >>>
        details_container = CTkFrame(parent, fg_color="transparent")
        details_container.pack(fill="x", expand=True, pady=10)
        details_container.grid_columnconfigure(0, weight=1)
        details_container.grid_columnconfigure(1, weight=1)

        # --- คอลัมน์ซ้าย: ส่วนลด/รายการเพิ่มเติม ---
        discounts_frame = self._create_section_frame(details_container, "ส่วนลด/รายการเพิ่มเติม")
        discounts_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.brokerage_fee_entry = NumericEntry(discounts_frame)
        self._add_form_row(discounts_frame, "ค่านายหน้า:", self.brokerage_fee_entry, 1)
        self.coupon_value_entry = NumericEntry(discounts_frame)
        self._add_form_row(discounts_frame, "คูปอง:", self.coupon_value_entry, 2)
        self.giveaway_value_entry = NumericEntry(discounts_frame)
        self._add_form_row(discounts_frame, "ของแถม:", self.giveaway_value_entry, 3)

        # --- คอลัมน์ขวา: Delivery Note ---
        delivery_frame = self._create_section_frame(details_container, "Delivery Note")
        delivery_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        delivery_frame.grid_columnconfigure(1, weight=1) # ทำให้คอลัมน์ที่ 2 ขยายได้

        delivery_options = [
            "ซัพพลายเออร์จัดส่ง", "Aplus Logistic ส่งหน้างาน", "ลูกค้ารับเองที่ซัพ",
            "ลูกค้ารับเองที่คลัง 132", "ย้ายเข้าคลัง Aplus Logistic รอลูกค้ารับที่คลัง",
            "ย้ายเข้าคลัง Aplus Logistic รอ Aplus Logistic จัดส่ง",
            "ย้ายเข้าคลัง Lalamove รอลูกค้ารับที่คลัง 132", "ส่ง Lalamove ให้ลูกค้าหน้างาน",
            "Aplus Logistic+ฝากส่งขนส่ง", "Lalamove +ฝากส่งขนส่ง"
        ]
        self.delivery_type_menu = CTkOptionMenu(delivery_frame, variable=self.delivery_type_var, values=delivery_options, command=self._on_delivery_type_change, **self.dropdown_style)
        self._add_form_row(delivery_frame, "การจัดส่ง:", self.delivery_type_menu, 1)

        self.pickup_location_entry = CTkEntry(delivery_frame, placeholder_text="ใส่ อำเภอ, จังหวัด หรือ Google map link")
        self._add_form_row(delivery_frame, "Location เข้ารับ:", self.pickup_location_entry, 2)

        # --- ส่วนของ "ค่าย้าย" ที่แก้ไขใหม่ ---
        CTkLabel(delivery_frame, text="ค่าย้าย:", font=CTkFont(size=14)).grid(row=3, column=0, padx=15, pady=5, sticky="w")
        relocation_entry_frame = CTkFrame(delivery_frame, fg_color="transparent")
        relocation_entry_frame.grid(row=3, column=1, padx=(10,15), pady=5, sticky="ew")
        self.relocation_cost_entry = NumericEntry(relocation_entry_frame)
        self.relocation_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        
        relocation_radio_frame = CTkFrame(relocation_entry_frame, fg_color="transparent")
        relocation_radio_frame.pack(side="left")
        CTkRadioButton(relocation_radio_frame, text="VAT", variable=self.relocation_vat_option_var, value="VAT").pack(side="left", padx=5)
        CTkRadioButton(relocation_radio_frame, text="CASH", variable=self.relocation_vat_option_var, value="NO VAT").pack(side="left", padx=5)

        # --- ช่องแสดงผล VAT ของ "ค่าย้าย" ---
        self.relocation_vat_display = CTkEntry(delivery_frame, textvariable=self.relocation_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(delivery_frame, "VAT 7% (ค่าย้าย):", self.relocation_vat_display, 4)

        # --- Widget ที่เหลือ ถูกเลื่อนลำดับแถวลงมา ---
        self.date_to_wh_label = CTkLabel(delivery_frame, text="วันที่ย้ายเข้าคลัง:", font=CTkFont(size=14))
        self.date_to_wh_label.grid(row=5, column=0, padx=15, pady=4, sticky="w")
        self.date_to_wh_selector = DateSelector(delivery_frame, dropdown_style=self.dropdown_style)
        self.date_to_wh_selector.grid(row=5, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")

        self.date_to_customer_selector = DateSelector(delivery_frame, dropdown_style=self.dropdown_style)
        self._add_form_row(delivery_frame, "วันที่จัดส่งลูกค้า:", self.date_to_customer_selector, 6)

        self.pickup_rego_entry = CTkEntry(delivery_frame)
        self._add_form_row(delivery_frame, "ทะเบียนเข้ารับ:", self.pickup_rego_entry, 7)
    # <<< END: CODE REPLACEMENT >>>


    def _populate_sales_services_frame(self, parent):
        frame = self._create_section_frame(parent, "ยอดขายและบริการ")
        frame.pack(fill="x", pady=(0,10))
        frame.grid_columnconfigure(1, weight=1)

        CTkLabel(frame, text="รายการ", font=CTkFont(size=14, weight="bold")).grid(row=1, column=0, padx=15)
        CTkLabel(frame, text="ยอดขาย", font=CTkFont(size=14, weight="bold")).grid(row=1, column=1, padx=10)
        CTkLabel(frame, text="ประเภท", font=CTkFont(size=14, weight="bold")).grid(row=1, column=2, padx=10)

        self.sales_amount_entry = NumericEntry(frame)
        self._add_item_row_with_vat(frame, "ยอดขายสินค้า/บริการ:", self.sales_amount_entry, self.sales_service_vat_option, 2)

        note_font = CTkFont(size=12, slant="italic")
        note_label = CTkLabel(frame, text="หมายเหตุ: ไม่รวมค่าส่ง/ค่ารถ", font=note_font, text_color="#FF0000")
        note_label.grid(row=3, column=1, columnspan=2, padx=(10, 15), pady=(0, 5), sticky="w")

        self.sales_vat_var_display = CTkEntry(frame, textvariable=self.sales_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "VAT 7% (สินค้า/บริการ):", self.sales_vat_var_display, 4)

        self.cutting_drilling_fee_entry = NumericEntry(frame)
        self._add_item_row_with_vat(frame, "ค่าบริการตัด/เจาะ:", self.cutting_drilling_fee_entry, self.cutting_drilling_fee_vat_option, 5)

        self.cutting_drilling_vat_var_display = CTkEntry(frame, textvariable=self.cutting_drilling_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "VAT 7% (ตัด/เจาะ):", self.cutting_drilling_vat_var_display, 6)

        self.other_service_fee_entry = NumericEntry(frame)
        self._add_item_row_with_vat(frame, "ค่าบริการอื่นๆ:", self.other_service_fee_entry, self.other_service_fee_vat_option, 7)

        self.other_service_vat_var_display = CTkEntry(frame, textvariable=self.other_service_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "VAT 7% (บริการอื่นๆ):", self.other_service_vat_var_display, 8)

    def _populate_shipping_frame(self, parent):
        frame = self._create_section_frame(parent, "ค่าจัดส่ง"); frame.pack(fill="x", pady=(0,10)); frame.grid_columnconfigure(1, weight=1)
        self.shipping_cost_entry = NumericEntry(frame); self._add_item_row_with_vat(frame, "ค่าจัดส่ง:", self.shipping_cost_entry, self.shipping_vat_option_var, 1)
        self.shipping_vat_var_display = CTkEntry(frame, textvariable=self.shipping_vat_calc_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "VAT 7% (ค่าจัดส่ง):", self.shipping_vat_var_display, 2)
        self.delivery_date_selector = DateSelector(frame, dropdown_style=self.dropdown_style); self._add_form_row(frame, "วันที่จัดส่ง:", self.delivery_date_selector, 3, columnspan=2)

    def _populate_fees_frame(self, parent):
        frame = self._create_section_frame(parent, "ค่าธรรมเนียม"); frame.pack(fill="x", pady=(0,10)); frame.grid_columnconfigure(1, weight=1)
        self.credit_card_fee_entry = NumericEntry(frame); self._add_item_row_with_vat(frame, "ค่าธรรมเนียมบัตรเครดิต:", self.credit_card_fee_entry, self.credit_card_fee_vat_option_var, 1)
        self.card_fee_vat_var_display = CTkEntry(frame, textvariable=self.card_fee_vat_calc_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "VAT 7% (ค่าธรรมเนียม):", self.card_fee_vat_var_display, 2)
        self.transfer_fee_entry = NumericEntry(frame, placeholder_text="หากมี"); self._add_form_row(frame, "ค่าธรรมเนียมโอน:", self.transfer_fee_entry, 3)
        self.wht_fee_entry = NumericEntry(frame, placeholder_text="หากมี"); self._add_form_row(frame, "ภาษีหัก ณ ที่จ่าย:", self.wht_fee_entry, 4)

    def _populate_payment_frame(self, parent):
        frame = self._create_section_frame(parent, "รายละเอียดการโอนชำระ")
        frame.pack(fill="x", expand=True, pady=10)
        
        payment_options = ["ชำระสด", "โอน KBANK", "โอน TTB - ออมทรัพย์", "โอน TTB - กระแส", "โอน กรรมการ", "ชำระผ่านบัตรเครดิต"]

        # --- Payment 1 Row ---
        payment1_frame = CTkFrame(frame, fg_color="transparent")
        payment1_frame.grid(row=1, column=1, columnspan=2, padx=(10,15), pady=4, sticky="ew")

        self.payment1_percent_menu = CTkOptionMenu(payment1_frame, variable=self.payment1_percent_var, values=["ระบุยอดเอง", "30%", "50%", "100%"], width=120, **self.dropdown_style)
        self.payment1_percent_menu.pack(side="left", padx=(0, 5))

        self.payment1_amount_entry = NumericEntry(payment1_frame, placeholder_text="ระบุยอดโอนตามสลิป")
        self.payment1_amount_entry.pack(side="left", fill="x", expand=True, padx=(0,5))

        self.payment1_method_menu = CTkOptionMenu(payment1_frame, variable=self.payment1_method_var, values=payment_options, width=160, **self.dropdown_style)
        self.payment1_method_menu.pack(side="left", padx=(0,5))

        self.payment1_date_selector = DateSelector(payment1_frame, dropdown_style=self.dropdown_style)
        self.payment1_date_selector.pack(side="left")

        self._add_form_row(frame, "1. มัดจำ/ชำระเต็ม:", payment1_frame, 1)

        # --- Payment 2 Row ---
        payment2_frame = CTkFrame(frame, fg_color="transparent")
        payment2_frame.grid(row=2, column=1, columnspan=2, padx=(10,15), pady=4, sticky="ew")

        self.payment2_percent_menu = CTkOptionMenu(payment2_frame, variable=self.payment2_percent_var, values=["ระบุยอดเอง", "30%", "50%", "70%", "100%"], width=120, **self.dropdown_style)
        self.payment2_percent_menu.pack(side="left", padx=(0, 5))

        self.payment2_amount_entry = NumericEntry(payment2_frame, placeholder_text="ระบุยอดโอนตามสลิป")
        self.payment2_amount_entry.pack(side="left", fill="x", expand=True, padx=(0,5))

        self.payment2_method_menu = CTkOptionMenu(payment2_frame, variable=self.payment2_method_var, values=payment_options, width=160, **self.dropdown_style)
        self.payment2_method_menu.pack(side="left", padx=(0,5))

        self.payment2_date_selector = DateSelector(payment2_frame, dropdown_style=self.dropdown_style)
        self.payment2_date_selector.pack(side="left")

        self._add_form_row(frame, "2. มัดจำ/ชำระเต็ม:", payment2_frame, 2)

        # --- Other rows ---
        self.payment1_percent_menu.configure(command=self._on_payment1_select)
        self.payment2_percent_menu.configure(command=self._on_payment2_select)

        payment_total_output = CTkEntry(frame, textvariable=self.payment_total_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "ยอดโอนชำระรวม VAT:", payment_total_output, 3)

        self.balance_due_entry = CTkEntry(frame, textvariable=self.balance_due_var, state="readonly", fg_color="gray85", font=CTkFont(weight="bold"))
        self._add_form_row(frame, "ค้างชำระ:", self.balance_due_entry, 4)

    def _populate_so_summary_frame(self, parent):
        frame = self._create_section_frame(parent, "SO ยอดขายสินค้า/ค่าจัดส่ง รวมภาษีมูลค่าเพิ่ม"); frame.pack(fill="x", pady=(0, 10))
        frame.grid_columnconfigure(1, weight=1)
        self._add_form_row(frame, "รวมยอดขาย SO:", CTkEntry(frame, textvariable=self.so_subtotal_var, state="readonly", fg_color="gray85"), 1, columnspan=1)
        self._add_form_row(frame, "VAT 7%:", CTkEntry(frame, textvariable=self.so_vat_var, state="readonly", fg_color="gray85"), 2, columnspan=1)
        self._add_form_row(frame, "ยอดที่ต้องชำระ:", CTkEntry(frame, textvariable=self.so_grand_total_var, state="readonly", fg_color="gray85", font=CTkFont(weight="bold")), 3, columnspan=1)
        self.so_vs_payment_result_entry = CTkEntry(frame, textvariable=self.so_vs_payment_result_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "ตรวจสอบยอด SO VS ชำระ:", self.so_vs_payment_result_entry, 4, columnspan=1)
        self.difference_amount_entry = CTkEntry(frame, textvariable=self.difference_amount_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "ผลต่าง:", self.difference_amount_entry, 5, columnspan=1)

        self.balance_due_entry = CTkEntry(frame, textvariable=self.balance_due_var, state="readonly", fg_color="gray85", font=CTkFont(weight="bold"))
        self._add_form_row(frame, "ค้างชำระ:", self.balance_due_entry, 6, columnspan=1)

    def _populate_cash_verification_frame(self, parent):
        frame = self._create_section_frame(parent, "ตรวจสอบยอดชำระเงินสด CASH"); frame.pack(fill="x", pady=0)
        frame.grid_columnconfigure(1, weight=1)

        CTkLabel(frame, text="ยอดค่าสินค้าเงินสด:", font=CTkFont(size=14)).grid(row=1, column=0, padx=15, pady=5, sticky="w")
        self.cash_product_input_entry = NumericEntry(frame, textvariable=self.cash_product_input_var, placeholder_text="0.00")
        self.cash_product_input_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.cash_product_input_entry.bind("<KeyRelease>", self._update_final_calculations)

        CTkLabel(frame, text="ยอดรวมค่าบริการเงินสด:", font=CTkFont(size=14)).grid(row=2, column=0, padx=15, pady=5, sticky="w")
        self.cash_service_total_entry = CTkEntry(frame, textvariable=self.cash_service_total_var, state="readonly", fg_color="gray85")
        self.cash_service_total_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        CTkLabel(frame, text="ยอดที่ต้องชำระเงินสด:", font=CTkFont(size=14)).grid(row=3, column=0, padx=15, pady=5, sticky="w")
        self.cash_required_total_entry = CTkEntry(frame, textvariable=self.cash_required_total_var, state="readonly", fg_color="gray85")
        self.cash_required_total_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        CTkLabel(frame, text="ยอดชำระจริงเงินสด:", font=CTkFont(size=14)).grid(row=4, column=0, padx=15, pady=5, sticky="w")
        self.cash_actual_payment_entry = NumericEntry(frame, textvariable=self.cash_actual_payment_var, placeholder_text="0.00")
        self.cash_actual_payment_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.cash_actual_payment_entry.bind("<KeyRelease>", self._update_final_calculations)

        CTkLabel(frame, text="ตรวจสอบยอดชำระเงินสด:", font=CTkFont(size=14)).grid(row=5, column=0, padx=15, pady=5, sticky="w")
        self.cash_verification_result_entry = CTkEntry(frame, textvariable=self.cash_verification_result_var, state="readonly", fg_color="gray85")
        self.cash_verification_result_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

    def _populate_action_frame(self, parent):
        frame = CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=(10,0))
        frame.grid_rowconfigure((0,1), weight=1)
        frame.grid_columnconfigure((0,1,2), weight=1)
        
        # +++ START: เพิ่ม self. เพื่อเก็บ Reference ของปุ่ม +++
        self.btn_clear = CTkButton(frame, text="ล้างข้อมูล", fg_color="#F97316", command=self._clear_form)
        self.btn_clear.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        self.btn_edit = CTkButton(frame, text="แก้ไขข้อมูลจากประวัติ", fg_color="#EAB308", command=self._show_history)
        self.btn_edit.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        self.save_button = CTkButton(frame, text="บันทึกข้อมูล", command=self._save_data)
        self.save_button.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")
        
        btn_history = CTkButton(frame, text="แสดงประวัติ", command=self._show_history)
        btn_history.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        
        btn_export = CTkButton(frame, text="นำไฟล์ออก EXCEL", fg_color="#1F2937", command=self._export_history_to_excel)
        btn_export.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        
        btn_submit = CTkButton(frame, text="นำส่งข้อมูล...", fg_color="#16A34A", command=self._open_submit_dialog)
        btn_submit.grid(row=1, column=2, padx=5, pady=5, sticky="nsew")
        # +++ END +++
        
        note_text = "หมายเหตุ **\nนำส่งข้อมูลเข้าระบบคำนวณคอมมิชชั่นแล้วไม่สามารถแก้ไขได้"
        note_label = CTkLabel(frame, text=note_text, font=CTkFont(size=13), text_color="#D32F2F", justify="left")
        note_label.grid(row=2, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

    def _open_submit_dialog(self):
        # ฟังก์ชันใหม่สำหรับเปิดหน้าต่างเลือกส่ง SO
        SubmitSODialog(self, self.app_container, self.sale_key, self.sale_name)

    def _update_final_calculations(self, *args):
        # --- 1. รวบรวมข้อมูล (ส่วนนี้ทำงานถูกต้องอยู่แล้ว) ---
        sales = utils.convert_to_float(self.sales_amount_entry.get())
        shipping = utils.convert_to_float(self.shipping_cost_entry.get())
        card_fee = utils.convert_to_float(self.credit_card_fee_entry.get())
        cutting_drilling = utils.convert_to_float(self.cutting_drilling_fee_entry.get())
        other_service = utils.convert_to_float(self.other_service_fee_entry.get())
        relocation = utils.convert_to_float(self.relocation_cost_entry.get())
        wht = utils.convert_to_float(self.wht_fee_entry.get())

        total_vatable_revenue = 0.0
        total_cashable_services_and_fees = 0.0
        items_to_process = [
            (sales, self.sales_service_vat_option.get(), self.sales_vat_calc_var),
            (cutting_drilling, self.cutting_drilling_fee_vat_option.get(), self.cutting_drilling_vat_calc_var),
            (other_service, self.other_service_fee_vat_option.get(), self.other_service_vat_calc_var),
            (shipping, self.shipping_vat_option_var.get(), self.shipping_vat_calc_var),
            (card_fee, self.credit_card_fee_vat_option_var.get(), self.card_fee_vat_calc_var),
            (relocation, self.relocation_vat_option_var.get(), self.relocation_vat_calc_var)
        ]
        for amount, option, var_display in items_to_process:
            item_vat = 0.0
            if option == "VAT":
                total_vatable_revenue += amount
                item_vat = amount * 0.07
            else:
                total_cashable_services_and_fees += amount
            if var_display:
                var_display.set(f"{item_vat:,.2f}")

        final_amount_due = (total_vatable_revenue * 1.07) - wht
        self.so_grand_total_var.set(f"{final_amount_due:,.2f}")

        payment1 = utils.convert_to_float(self.payment1_amount_entry.get())
        payment2 = utils.convert_to_float(self.payment2_amount_entry.get())
        total_payment = payment1 + payment2
        self.payment_total_var.set(f"{total_payment:,.2f}")

        # --- 2. คำนวณส่วนต่าง (Difference) ---
        # แก้ไขสูตรเป็น: ยอดโอน - ยอดที่ต้องชำระ
        # - ถ้าเป็นบวก (+) แปลว่า "โอนเกิน"
        # - ถ้าเป็นลบ (-) แปลว่า "โอนขาด"
        difference = total_payment - final_amount_due
        self.difference_amount_var.set(f"{difference:,.2f}")
        
        # ยอดค้างชำระ (Balance Due) จะเป็นค่าบวกเสมอ
        self.balance_due_var.set(f"{abs(difference):,.2f}")

        def set_check_result(entry, var, diff_val, plus_text, minus_text):
            if not entry or not entry.winfo_exists(): return
            color_map = {"ok": ("#BBF7D0", "#15803D"), "bad": ("#FECACA", "#B91C1C")}
            if abs(diff_val) < 0.01:
                state, text = "ok", "ถูกต้อง"
                entry.configure(fg_color=color_map["ok"][0], text_color=color_map["ok"][1])
            elif diff_val > 0: # โอนเกิน (ค่าเป็นบวก)
                state, text = "ok", f"{plus_text} (+{abs(diff_val):,.2f})"
                entry.configure(fg_color=color_map["ok"][0], text_color=color_map["ok"][1])
            else: # โอนขาด (ค่าเป็นลบ)
                state, text = "bad", f"{minus_text} ({abs(diff_val):,.2f})"
                entry.configure(fg_color=color_map["bad"][0], text_color=color_map["bad"][1])
            var.set(text)

        # --- 3. เรียกใช้ฟังก์ชันแสดงผลด้วย Logic ที่ถูกต้อง ---
        # เมื่อ difference > 0 (บวก) ให้แสดง plus_text ("ยอดโอนเกิน")
        # เมื่อ difference < 0 (ลบ) ให้แสดง minus_text ("ยอดโอนขาด")
        set_check_result(self.so_vs_payment_result_entry, self.so_vs_payment_result_var, difference,
                         plus_text="ยอดโอนเกิน",
                         minus_text="ยอดโอนขาด")

        # --- 4. คำนวณส่วนของเงินสด (Cash) ---
        cash_product_val = utils.convert_to_float(self.cash_product_input_entry.get())
        cash_required_total = cash_product_val + total_cashable_services_and_fees
        self.cash_required_total_var.set(f"{cash_required_total:,.2f}")

        actual_cash_payment = utils.convert_to_float(self.cash_actual_payment_entry.get())
        cash_difference = actual_cash_payment - cash_required_total
        
        set_check_result(self.cash_verification_result_entry, self.cash_verification_result_var, cash_difference,
                         plus_text="เงินสดเกิน",
                         minus_text="เงินสดขาด")

        cash_product_val = utils.convert_to_float(self.cash_product_input_entry.get())
        cash_required_total = cash_product_val + total_cashable_services_and_fees
        self.cash_required_total_var.set(f"{cash_required_total:,.2f}")
        actual_cash_payment = utils.convert_to_float(self.cash_actual_payment_entry.get())
        cash_diff = cash_required_total - actual_cash_payment
        set_check_result(self.cash_verification_result_entry, self.cash_verification_result_var, cash_diff, "เงินสดขาด", "เงินสดเกิน")

    def _gather_data_from_form(self):
        is_new_customer = self.customer_type_var.get() == "ลูกค้าใหม่"

        customer_id = self.new_customer_id_entry.get().strip() if is_new_customer else self.customer_id_entry.get().strip()
        customer_name = self.new_customer_name_entry.get().strip() if is_new_customer else self.customer_name_entry.get().strip()

        p1_date = self.payment1_date_selector.get_date()
        p2_date = self.payment2_date_selector.get_date()
        main_payment_date = None
        if p1_date and p2_date: main_payment_date = max(p1_date, p2_date)
        elif p1_date: main_payment_date = p1_date
        elif p2_date: main_payment_date = p2_date

        data = {
            # Sales Details
            "bill_date": self.bill_date_selector.get_date(),
            "commission_month": self.thai_month_map.get(self.commission_month_var.get()),
            "commission_year": int(self.commission_year_var.get()) - 543 if self.commission_year_var.get().isdigit() else None,
            "customer_type": self.customer_type_var.get(),
            "customer_name": customer_name,
            "customer_id": customer_id,
            "credit_term": self.credit_term_var.get(),
            "so_number": self.so_number_var.get().strip(),

            # Sales and Services
            "sales_service_amount": utils.convert_to_float(self.sales_amount_entry.get()),
            "sales_service_vat_option": self.sales_service_vat_option.get(),
            "cutting_drilling_fee": utils.convert_to_float(self.cutting_drilling_fee_entry.get()),
            "cutting_drilling_fee_vat_option": self.cutting_drilling_fee_vat_option.get(),
            "other_service_fee": utils.convert_to_float(self.other_service_fee_entry.get()),
            "other_service_fee_vat_option": self.other_service_fee_vat_option.get(),

            # Shipping
            "shipping_cost": utils.convert_to_float(self.shipping_cost_entry.get()),
            "shipping_vat_option": self.shipping_vat_option_var.get(),
            "delivery_date": self.delivery_date_selector.get_date(),

            # Fees
            "credit_card_fee": utils.convert_to_float(self.credit_card_fee_entry.get()),
            "credit_card_fee_vat_option": self.credit_card_fee_vat_option_var.get(),
            "transfer_fee": utils.convert_to_float(self.transfer_fee_entry.get()),
            "wht_3_percent": utils.convert_to_float(self.wht_fee_entry.get()),

            # Other Expenses / Delivery Note
            "brokerage_fee": utils.convert_to_float(self.brokerage_fee_entry.get()),
            "coupons": utils.convert_to_float(self.coupon_value_entry.get()),
            "giveaways": utils.convert_to_float(self.giveaway_value_entry.get()),
            "relocation_cost_vat_option": self.relocation_vat_option_var.get(),
            # <<< START: ADD NEW DATA GATHERING >>>
            "delivery_type": self.delivery_type_var.get(),
            "pickup_location": self.pickup_location_entry.get().strip(),
            "relocation_cost": utils.convert_to_float(self.relocation_cost_entry.get()),
            "date_to_warehouse": self.date_to_wh_selector.get_date(),
            "date_to_customer": self.date_to_customer_selector.get_date(),
            "pickup_registration": self.pickup_rego_entry.get().strip(),
            # <<< END: ADD NEW DATA GATHERING >>>

            # Payment
            "total_payment_amount": utils.convert_to_float(self.payment_total_var.get()),
            "payment_date": main_payment_date,
            "payment1_date": p1_date,
            "payment2_date": p2_date,
            "payment1_method": self.payment1_method_var.get(),
            "payment2_method": self.payment2_method_var.get(),

            # Cash Verification
            "cash_product_input": utils.convert_to_float(self.cash_product_input_var.get()),
            "cash_service_total": utils.convert_to_float(self.cash_service_total_var.get()),
            "cash_required_total": utils.convert_to_float(self.cash_required_total_var.get()),
            "cash_actual_payment": utils.convert_to_float(self.cash_actual_payment_var.get()),

            # Other Calculated
            "difference_amount": utils.convert_to_float(self.difference_amount_var.get()),

            # System Info
            "sale_key": self.sale_key,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "Original",
            "is_active": 1,
        }

        if hasattr(self, 'proxy_user_key') and self.proxy_user_key:
            data['support_user_key'] = self.proxy_user_key

        return data

    def _validate_form(self, data):
        if not data["so_number"] or data["so_number"] == "SO":
            return False, "กรุณากรอก 'เลขที่ใบสั่งขาย (SO)'"

        if data['customer_type'] == "ลูกค้าใหม่":
            if not data["customer_name"] or not data["customer_id"]:
                return False, "สำหรับ 'ลูกค้าใหม่' กรุณากรอก 'ชื่อ' และ 'รหัส' ให้ครบถ้วน"
        else:  # ลูกค้าเก่า
            if not data["customer_id"]:
                return False, "กรุณาเลือก 'รหัสลูกค้า' สำหรับลูกค้าเก่า"

        # <<< START: แก้ไข Logic การตรวจสอบ SO ซ้ำทั้งหมด >>>
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # สร้าง Query และ Parameters พื้นฐาน
                query = "SELECT id FROM commissions WHERE so_number = %s AND is_active = 1"
                params = [data["so_number"]]

                # ถ้าอยู่ในโหมดแก้ไข (self.editing_record_id มีค่า)
                # ให้เพิ่มเงื่อนไขว่า "ไม่ต้องตรวจสอบกับ ID ของตัวเอง"
                if self.editing_record_id:
                    query += " AND id != %s"
                    params.append(self.editing_record_id)

                # ทำการ Query ด้วยเงื่อนไขที่ถูกต้อง
                cursor.execute(query, tuple(params))
                if cursor.fetchone():
                    return False, f"เลขที่ SO '{data['so_number']}' นี้มีอยู่ในระบบแล้ว"
        except Exception as e:
            # กรณีเกิดข้อผิดพลาดในการเชื่อมต่อฐานข้อมูล
            return False, f"เกิดข้อผิดพลาดในการตรวจสอบข้อมูล: {e}"
        finally:
            if conn:
                self.app_container.release_connection(conn)
        # <<< END: สิ้นสุดการแก้ไข >>>

        return True, ""

    def _handle_new_customer(self, data):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("SELECT id FROM customers WHERE customer_code = %s", (data['customer_id'],))
                if cursor.fetchone():
                    raise ValueError(f"รหัสลูกค้า '{data['customer_id']}' นี้มีอยู่แล้วในระบบ\nกรุณาใช้รหัสอื่น หรือเลือกจากเมนูลูกค้าเก่า")

                insert_query = "INSERT INTO customers (customer_code, customer_name, credit_term) VALUES (%s, %s, %s)"
                cursor.execute(insert_query, (data['customer_id'], data['customer_name'], data['credit_term']))
            conn.commit()
            print(f"Added new customer: {data['customer_name']}")
        except Exception as e:
            if conn: conn.rollback()
            raise e
        finally:
            if conn: self.app_container.release_connection(conn)

    def _perform_db_insert(self, data):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'commissions'")
                db_columns = {row[0] for row in cursor.fetchall()}

                filtered_data = {k: v for k, v in data.items() if k in db_columns}

                filtered_data.pop('id', None)

                cols = ', '.join([f'"{k}"' for k in filtered_data.keys()])
                placeholders = ', '.join(['%s'] * len(filtered_data))

                sql = f"INSERT INTO commissions ({cols}) VALUES ({placeholders})"
                cursor.execute(sql, list(filtered_data.values()))
            conn.commit()
        except Exception as e:
            if conn: conn.rollback()
            raise e
        finally:
            if conn: self.app_container.release_connection(conn)

    def _refresh_history_if_open(self):
        if self.history_window and self.history_window.winfo_exists():
            if hasattr(self.history_window, '_populate_history_table'):
                self.history_window._populate_history_table()

    def _validate_so_number(self, *args):
        current_value = self.so_number_var.get()
        new_value = current_value.upper()
        if new_value != current_value:
            self.so_number_var.set(new_value)

    def _create_header(self):
        self.header_frame = CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        # --- START: แก้ไขส่วนนี้ทั้งหมด ---
        # กำหนดให้คอลัมน์ซ้าย (ชื่อ) ขยายตัว และคอลัมน์ขวา (ปุ่ม) ไม่ขยาย
        self.header_frame.grid_columnconfigure(0, weight=1)
        
        CTkLabel(self.header_frame, text=f"ฝ่ายขาย: {self.sale_name} ({self.sale_key})", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).grid(row=0, column=0, sticky="w")
        
        # สร้าง Frame สำหรับปุ่มต่างๆ ทางขวา และใช้ grid วาง
        button_container = CTkFrame(self.header_frame, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="e")
        
        # ใช้ grid วางปุ่มภายใน button_container
        self.tasks_button = CTkButton(button_container, text="งานของฉัน 🔔 (0)", command=self._open_my_tasks_window)
        self.tasks_button.grid(row=0, column=0, padx=10)

        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").grid(row=0, column=1, padx=(0, 10))

    def _on_payment1_select(self, selected_value: str):
        self._calculate_payment_from_percentage(self.payment1_percent_var, self.payment1_amount_entry)
    def _on_payment2_select(self, selected_value: str):
        self._calculate_payment_from_percentage(self.payment2_percent_var, self.payment2_amount_entry)

    def _calculate_payment_from_percentage(self, percent_var, amount_entry):
        try:
            selected_option = percent_var.get()
            if selected_option == "ระบุยอดเอง": return

            grand_total = utils.convert_to_float(self.so_grand_total_var.get())
            if grand_total <= 0:
                amount_entry.delete(0, tk.END)
                self._update_final_calculations()
                return

            percent = float(selected_option.replace('%', '')) / 100.0
            calculated_amount = grand_total * percent

            amount_entry.delete(0, tk.END)
            amount_entry.insert(0, f"{calculated_amount:,.2f}")

            if selected_option == "100%":
                if amount_entry == self.payment1_amount_entry:
                    self.payment2_percent_var.set("ระบุยอดเอง")
                    self.payment2_amount_entry.delete(0, tk.END)
                else:
                    self.payment1_percent_var.set("ระบุยอดเอง")
                    self.payment1_amount_entry.delete(0, tk.END)

            self._update_final_calculations()
        except (ValueError, TypeError) as e:
            print(f"Error calculating payment from percentage: {e}")
            self._update_final_calculations()

    def _clear_form(self, confirm=True):
        if confirm and not messagebox.askyesno("ยืนยัน", "คุณต้องการล้างข้อมูลทั้งหมดในฟอร์มใช่หรือไม่?", parent=self):
            return

        widget_attributes = [
            "so_number_entry", "customer_id_entry", "customer_name_entry",
            "new_customer_id_entry", "new_customer_name_entry", "sales_amount_entry",
            "cutting_drilling_fee_entry", "other_service_fee_entry", "shipping_cost_entry",
            "credit_card_fee_entry", "transfer_fee_entry", "wht_fee_entry", "brokerage_fee_entry",
            "coupon_value_entry", "giveaway_value_entry", "payment1_amount_entry",
            "payment2_amount_entry", "cash_product_input_entry", "cash_actual_payment_entry",
            "pickup_location_entry", "relocation_cost_entry", "pickup_rego_entry"
        ]

        for attr_name in widget_attributes:
            if hasattr(self, attr_name):
                widget = getattr(self, attr_name)
                if isinstance(widget, (CTkEntry, NumericEntry, AutoCompleteEntry)):
                    if widget.cget("state") == "readonly":
                        widget.configure(state="normal")
                        widget.delete(0, tk.END)
                        widget.configure(state="readonly")
                    else:
                        widget.delete(0, tk.END)

        today = datetime.now()
        for selector in [self.bill_date_selector, self.delivery_date_selector, 
                         self.payment1_date_selector, self.payment2_date_selector,
                         self.date_to_wh_selector, self.date_to_customer_selector]:
            selector.set_date(today)

        self.so_number_var.set("SO")
        self.customer_type_var.set("ลูกค้าเก่า")
        self.credit_term_var.set("เงินสด")

        self.commission_month_var.set(self.thai_months[today.month - 1])
        self.commission_year_var.set(str(today.year + 543))
        self.payment1_percent_var.set("ระบุยอดเอง")
        self.payment2_percent_var.set("ระบุยอดเอง")

        self.sales_service_vat_option.set("VAT")
        self.cutting_drilling_fee_vat_option.set("VAT")
        self.other_service_fee_vat_option.set("VAT")
        self.shipping_vat_option_var.set("VAT")
        self.credit_card_fee_vat_option_var.set("VAT")
        self.relocation_vat_option_var.set("VAT")

        self.payment1_method_var.set("ไม่เลือก")
        self.payment2_method_var.set("ไม่เลือก")
        self.delivery_type_var.set("ซัพพลายเออร์จัดส่ง")

        self.editing_record_id = None
        self._toggle_customer_fields()
        self._update_final_calculations()

        if confirm: messagebox.showinfo("สำเร็จ", "ข้อมูลถูกล้างเรียบร้อยแล้ว", parent=self)

    def _save_data(self):
        form_data = self._gather_data_from_form()
        is_valid, message = self._validate_form(form_data)
        if not is_valid:
            messagebox.showerror("ข้อมูลไม่ถูกต้อง", message, parent=self)
            return

        if self.editing_record_id:
            if not messagebox.askyesno("ยืนยันการแก้ไข", "คุณต้องการบันทึกการเปลี่ยนแปลงนี้ใช่หรือไม่?", parent=self):
                return

        if form_data['customer_type'] == "ลูกค้าใหม่" and not self.editing_record_id:
            try:
                self._handle_new_customer(form_data)
            except Exception as e:
                messagebox.showerror("ผิดพลาด", f"ไม่สามารถเพิ่มลูกค้าใหม่ได้:\n{e}", parent=self)
                return

        # <<< START: แก้ไข Logic การบันทึกข้อมูลทั้งหมด >>>
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # ดึงรายชื่อคอลัมน์จากฐานข้อมูลเพื่อความแม่นยำ
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'commissions'")
                db_columns = {row[0] for row in cursor.fetchall()}
                
                # กรองข้อมูลจากฟอร์มให้เหลือเฉพาะที่มีในฐานข้อมูล
                filtered_data = {k: v for k, v in form_data.items() if k in db_columns}

                if self.editing_record_id:
                    # --- โหมดแก้ไข: ใช้คำสั่ง UPDATE ---
                    filtered_data['status'] = 'Edited'
                    
                    # สร้างส่วน SET ของคำสั่ง SQL แบบไดนามิก
                    set_clauses = [f'"{key}" = %s' for key in filtered_data.keys()]
                    params = list(filtered_data.values())
                    params.append(self.editing_record_id) # เพิ่ม ID สำหรับ WHERE clause

                    update_query = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
                    cursor.execute(update_query, tuple(params))
                    action = "อัปเดต"

                else:
                    # --- โหมดสร้างใหม่: ใช้คำสั่ง INSERT (เหมือนเดิม) ---
                    self._perform_db_insert(filtered_data) # ส่ง filtered_data ไปแทน
                    action = "บันทึก"

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"{action}ข้อมูลเรียบร้อยแล้ว", parent=self)

            self._clear_form(confirm=False)
            self._load_customer_data()
            self._refresh_history_if_open()
            self._update_tasks_badge()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _export_history_to_excel(self):
        # 1. เปิด DateRangeDialog เพื่อให้ผู้ใช้เลือกช่วงเวลา
        dialog = DateRangeDialog(self)
        self.master.wait_window(dialog)

        start_date_raw = dialog.start_date # รับค่าวันที่แบบ 'YYYY-MM-DD' หรือ None
        end_date_raw = dialog.end_date     # รับค่าวันที่แบบ 'YYYY-MM-DD' หรือ None

        # ถ้าผู้ใช้กดยกเลิกใน dialog
        if not start_date_raw or not end_date_raw:
            messagebox.showinfo("ยกเลิก", "การ Export ถูกยกเลิก", parent=self)
            return

        try:
            # สร้าง datetime objects โดยกำหนดเวลาให้เป็น 00:00:00 สำหรับวันเริ่มต้น
            # และ 23:59:59 สำหรับวันสิ้นสุด เพื่อให้ครอบคลุมทั้งวัน
            start_datetime = datetime.strptime(start_date_raw, '%Y-%m-%d')
            end_datetime = datetime.strptime(end_date_raw, '%Y-%m-%d') + timedelta(hours=23, minutes=59, seconds=59)

        except ValueError as e:
            messagebox.showerror("ข้อผิดพลาดวันที่", f"รูปแบบวันที่ไม่ถูกต้อง: {e}", parent=self)
            return

        if not messagebox.askyesno("ยืนยัน", "คุณต้องการดึงข้อมูลและ Export เป็นไฟล์ Excel หรือไม่?", parent=self):
            return

        try:
            # แก้ไข query ให้รวมเงื่อนไขของ Sale Key และช่วงเวลา
            # commissions.timestamp is TEXT, so direct BETWEEN might not work as expected with TIMESTAMP objects.
            # It's better to cast timestamp to TIMESTAMP or DATE in SQL for comparison with TIMESTAMP objects.
            # Assuming commissions.timestamp is stored as TEXT in 'YYYY-MM-DD HH:MM:SS' format.
            # If it's a real TIMESTAMP WITH TIME ZONE in DB, psycopg2 will handle the datetime objects.
            # If it's TEXT, we pass strings in the format expected by the DB.

            query = """
                SELECT * FROM commissions
                WHERE sale_key = %s AND is_active = 1
                AND timestamp::timestamp BETWEEN %s::timestamp AND %s::timestamp
                ORDER BY timestamp DESC
            """
            # ใช้ params เพื่อป้องกัน SQL Injection
            params = (self.sale_key, start_datetime.strftime('%Y-%m-%d %H:%M:%S'), end_datetime.strftime('%Y-%m-%d %H:%M:%S'))
            df = pd.read_sql_query(query, self.pg_engine, params=params)

            if df.empty:
                messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลสำหรับ Export ในช่วงเวลาที่เลือก", parent=self)
                return

            df_to_export = df.copy()
            df_to_export.rename(columns=self.header_map, inplace=True)

            default_filename = f"commission_history_{self.sale_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="บันทึกไฟล์ประวัติ Commission",
                initialfile=default_filename,
                parent=self
            )

            if save_path:
                df_to_export.to_excel(save_path, index=False)
                messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้วที่:\n{save_path}", parent=self)
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=self)
            traceback.print_exc() # สำหรับ Debugging