# purchasing_screen.py (ฉบับสมบูรณ์ แก้ไขทั้งหมด)

import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkFrame, CTkLabel, CTkEntry, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkOptionMenu, CTkCheckBox, CTkTabview, CTkComboBox,
                           CTkToplevel, CTkRadioButton, CTkSegmentedButton, CTkInputDialog)
from tkinter import messagebox
from datetime import datetime
import json
import psycopg2
import psycopg2.errors
import psycopg2.extras
import pandas as pd
import numpy as np
from PIL import Image, ImageTk
import traceback
import re
from history_windows import SOPopupWindow
from hr_windows import SODetailViewer 
from export_utils import export_approved_pos_to_excel


# --- แก้ไข: ลบ import ที่เป็นปัญหาออก และย้าย Dialog class ไปไว้ในไฟล์ของตัวเอง (ถ้ามี) ---
# from history_windows import SOPopupWindow 
# from export_utils import export_approved_pos_to_excel
from pdf_utils import export_approved_pos_to_pdf
from po_selection_dialog import POSelectionDialog

# --- แก้ไข: import Dialog ที่ถูกต้อง ---
from po_selection_dialog import POSelectionDialog
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
import utils

class SubmitPODialog(CTkToplevel):
    def __init__(self, master, purchasing_screen_instance):
        super().__init__(master)
        self.purchasing_screen = purchasing_screen_instance
        self.app_container = purchasing_screen_instance.app_container
        self.user_key = purchasing_screen_instance.user_key
        self.checkbox_list = []

        self.title("เลือกรายการ PO ที่จะส่งอนุมัติ")
        self.geometry("800x600")
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=15, pady=(10, 0), sticky="ew")
        self.select_all_var = tk.IntVar(value=0)
        self.select_all_checkbox = CTkCheckBox(top_frame, text="เลือกทั้งหมด", variable=self.select_all_var, command=self._toggle_all_checkboxes, font=CTkFont(weight="bold"))
        self.select_all_checkbox.pack(anchor="w")

        self.scroll_frame = CTkScrollableFrame(self, label_text="รายการ PO ที่เป็นฉบับร่าง")
        self.scroll_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        button_frame.grid_columnconfigure((0,1), weight=1)
        
        self.submit_button = CTkButton(button_frame, text="ยืนยันการส่งอนุมัติ (0)", command=self._confirm_submission, state="disabled")
        self.submit_button.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        CTkButton(button_frame, text="ยกเลิก", fg_color="gray", command=self.destroy).grid(row=0, column=1, padx=(5,0), sticky="ew")

        self.after(50, self._populate_po_list)
        self.transient(master)
        self.grab_set()
    
      # purchasing_screen.py (เพิ่มฟังก์ชันนี้เข้าไปในคลาส PurchasingScreen)


    def _populate_po_list(self):
        try:
            query = "SELECT id, po_number, so_number, supplier_name FROM purchase_orders WHERE user_key = %s AND status = 'Draft' ORDER BY timestamp DESC"
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.user_key,))

            if df.empty:
                CTkLabel(self.scroll_frame, text="ไม่พบรายการที่เป็นฉบับร่าง").pack(pady=20)
                self.select_all_checkbox.configure(state="disabled")
                return

            for _, row in df.iterrows():
                checkbox_var = tk.IntVar(value=0)
                checkbox_var.trace_add("write", self._update_submit_button_state)
                
                po_id = row['id']
                po_text = f"PO: {row['po_number']} | SO: {row['so_number']} | Supplier: {row['supplier_name']}"
                
                cb = CTkCheckBox(self.scroll_frame, text=po_text, variable=checkbox_var)
                cb.pack(anchor="w", padx=10, pady=5)
                self.checkbox_list.append((checkbox_var, po_id, row.to_dict()))

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดรายการ PO ได้: {e}", parent=self)
            self.destroy()

    def _toggle_all_checkboxes(self):
        is_selected = self.select_all_var.get()
        for var, _, _ in self.checkbox_list:
            var.set(is_selected)

    def _update_submit_button_state(self, *args):
        selected_count = sum(var.get() for var, _, _ in self.checkbox_list)
        self.submit_button.configure(text=f"ยืนยันการส่งอนุมัติ ({selected_count})")
        self.submit_button.configure(state="normal" if selected_count > 0 else "disabled")

    def _confirm_submission(self):
        selected_records = [(po_id, record_data) for var, po_id, record_data in self.checkbox_list if var.get() == 1]
        
        if not selected_records:
            messagebox.showwarning("ยังไม่ได้เลือก", "กรุณาเลือก PO อย่างน้อย 1 รายการ", parent=self)
            return

        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการส่ง PO จำนวน {len(selected_records)} รายการเพื่อขออนุมัติใช่หรือไม่?", parent=self):
            return
            
        selected_ids = [po_id for po_id, _ in selected_records]
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                
                # --- START: แก้ไข Logic การอัปเดตค่าขนส่ง (ฉบับสมบูรณ์) ---

                # 1. อัปเดตสถานะ PO ที่เลือกทั้งหมดให้เป็น 'Pending Approval' ก่อน
                ids_tuple = tuple(selected_ids)
                update_query = """
                    UPDATE purchase_orders 
                    SET status = 'Pending Approval', approval_status = 'Pending Mgr 1' 
                    WHERE id IN %s
                """
                cursor.execute(update_query, (ids_tuple,))

                # 2. รวบรวม SO ที่เกี่ยวข้องทั้งหมดจากการส่งครั้งนี้ (เพื่อไม่ให้ทำงานซ้ำซ้อน)
                affected_so_numbers = list(set(rec['so_number'] for _, rec in selected_records))

                # 3. วนลูปเพื่อ "คำนวณยอดรวมค่าขนส่งใหม่ทั้งหมด" ของแต่ละ SO
                for so_number in affected_so_numbers:
                    # 3.1 Query ยอดรวมค่าขนส่งจาก PO "ทุกใบ" (ที่อนุมัติแล้วหรือกำลังรออนุมัติ) ของ SO นี้
                    cursor.execute("""
                        SELECT SUM(COALESCE(shipping_to_stock_cost, 0) + COALESCE(shipping_to_site_cost, 0))
                        FROM purchase_orders
                        WHERE so_number = %s AND status IN ('Pending Approval', 'Approved')
                    """, (so_number,))
                    
                    new_total_shipping_cost = cursor.fetchone()[0] or 0.0

                    # 3.2 นำยอดรวมใหม่ "อัปเดตทับ" ค่าเก่าในตาราง commissions
                    # วิธีนี้จะทำให้ข้อมูลถูกต้องเสมอ ไม่มีการบวกซ้ำ
                    cursor.execute("""
                        UPDATE commissions
                        SET payment_before_vat = %s
                        WHERE so_number = %s AND is_active = 1
                    """, (new_total_shipping_cost, so_number))

                # --- END: สิ้นสุดการแก้ไข Logic ---

                # 4. สร้าง Notification (เหมือนเดิม)
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Manager' AND status = 'Active'")
                manager_keys = [row[0] for row in cursor.fetchall()]

                notif_data = []
                for po_id, record_data in selected_records:
                    message = f"PO ใหม่ ({record_data['po_number']}) รอการอนุมัติจากผู้จัดการ"
                    for manager_key in manager_keys:
                        notif_data.append((manager_key, message, False, po_id))
                
                psycopg2.extras.execute_values(
                    cursor,
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES %s",
                    notif_data
                )
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ส่ง PO จำนวน {len(selected_ids)} รายการเพื่อขออนุมัติเรียบร้อยแล้ว", parent=self.purchasing_screen)
            
            self.purchasing_screen._update_tasks_badge()
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการส่งข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

# --- Helper Window Classes ---
class MyTasksWindow(CTkToplevel):
    def __init__(self, master, purchasing_screen_instance):
        super().__init__(master)
        self.purchasing_screen = purchasing_screen_instance
        self.app_container = purchasing_screen_instance.app_container
        self.user_key = purchasing_screen_instance.user_key
        self.label_font = purchasing_screen_instance.label_font
        self.title("งานของฉัน (My Tasks)")
        self.geometry("900x600")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self._create_my_tasks_view(self)
        self.after(50, self.load_tasks)
        
        self.transient(master)
        self.grab_set()

    def _create_my_tasks_view(self, parent):
        header = CTkFrame(parent, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,5))
        CTkLabel(header, text="งานของฉัน (My Tasks)", font=CTkFont(size=18, weight="bold")).pack(side="left")
        CTkButton(header, text="Refresh", command=self.load_tasks, width=80).pack(side="right")
        
        self.task_tab_view = CTkTabview(parent, corner_radius=10)
        self.task_tab_view.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        self.new_so_tab = self.task_tab_view.add("SO ใหม่รอสร้าง PO")
        self.in_progress_tab = self.task_tab_view.add("งานที่กำลังดำเนินการ (SO/PO Drafts)")
        self.rejected_tab = self.task_tab_view.add("งานที่ถูกปฏิเสธ (Rejected)")

        # --- Layout ใหม่สำหรับแท็บ "SO ใหม่" ---
        self.new_so_scroll_frame = CTkScrollableFrame(self.new_so_tab, label_text="คลิกเพื่อเริ่มสร้าง PO")
        self.new_so_scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Layout ใหม่สำหรับแท็บ "งานที่กำลังดำเนินการ" ---
        self.in_progress_scroll_frame = CTkScrollableFrame(self.in_progress_tab)
        self.in_progress_scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.in_progress_scroll_frame.grid_columnconfigure(0, weight=1)

        # "โซน" สำหรับ SO
        so_zone = CTkFrame(self.in_progress_scroll_frame, fg_color="#F0F9FF", corner_radius=10)
        so_zone.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        so_zone.grid_columnconfigure(0, weight=1)
        CTkLabel(so_zone, text="SO ที่กำลังดำเนินการ", font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=15, pady=(10, 0))
        CTkLabel(so_zone, text="(SO ที่คุณ Claim มาแล้ว แต่ยังไม่ได้สร้าง PO)", font=CTkFont(size=12, slant="italic"), text_color="gray50").pack(anchor="w", padx=15, pady=(0, 10))
        self.so_in_progress_content_frame = CTkFrame(so_zone, fg_color="transparent") # Frame เปล่าสำหรับใส่ Card
        self.so_in_progress_content_frame.pack(fill="x", expand=True, padx=5, pady=(0, 5))

        # "โซน" สำหรับ PO
        po_zone = CTkFrame(self.in_progress_scroll_frame, fg_color="#F1F5F9", corner_radius=10)
        po_zone.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
        po_zone.grid_columnconfigure(0, weight=1)
        CTkLabel(po_zone, text="PO ฉบับร่าง", font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=15, pady=(10, 0))
        CTkLabel(po_zone, text="(PO ที่คุณสร้างและบันทึกร่างไว้ แต่ยังไม่ได้ส่งอนุมัติ)", font=CTkFont(size=12, slant="italic"), text_color="gray50").pack(anchor="w", padx=15, pady=(0, 10))
        self.po_draft_content_frame = CTkFrame(po_zone, fg_color="transparent") # Frame เปล่าสำหรับใส่ Card
        self.po_draft_content_frame.pack(fill="x", expand=True, padx=5, pady=(0, 5))

        # --- Layout ใหม่สำหรับแท็บ "งานที่ถูกปฏิเสธ" ---
        self.rejected_scroll_frame = CTkScrollableFrame(self.rejected_tab, label_text="รายการที่ต้องแก้ไข")
        self.rejected_scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
    def on_close(self):
        self.purchasing_screen._update_tasks_badge()
        self.purchasing_screen.tasks_window = None
        self.destroy()

    def load_tasks(self):
        self._load_new_so_tasks()
        self._load_in_progress_tasks()
        self._load_rejected_po_tasks()

    def _load_new_so_tasks(self):
        frame = self.new_so_scroll_frame
        for widget in frame.winfo_children(): widget.destroy()
        try:
            query = """
                SELECT c.id, c.so_number, c.timestamp, c.customer_name, u.sale_name 
                FROM commissions c JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Pending PU' AND c.is_active = 1 ORDER BY c.timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine)
            if df.empty: CTkLabel(frame, text="ไม่มี SO ใหม่").pack(pady=20); return
            for _, row in df.iterrows():
                card = CTkFrame(frame, border_width=1, fg_color="#F0FDF4")
                card.pack(fill="x", padx=5, pady=3)

                # กำหนด Grid ภายใน Card: คอลัมน์ 0 (ข้อมูล) จะขยาย, คอลัมน์ 1 (ปุ่ม) จะคงที่
                card.grid_columnconfigure(0, weight=1)
                card.grid_columnconfigure(1, weight=0)

                # สร้าง Frame สำหรับรวมข้อมูลตัวอักษรไว้ด้วยกันในคอลัมน์ที่ 0
                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.grid(row=0, column=0, sticky="w", padx=10, pady=5)

                # เตรียมข้อมูล
                ts = pd.to_datetime(row['timestamp']).strftime("%Y-%m-%d %H:%M") if pd.notna(row['timestamp']) else "N/A"
                info_text = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']} (ส่งโดย: {row['sale_name']})"

                # แสดงข้อมูลหลัก (บรรทัดบน)
                CTkLabel(info_frame, text=info_text, font=self.label_font, justify="left").pack(anchor="w")
                
                # แสดงเวลา (บรรทัดล่าง)
                CTkLabel(info_frame, text=f"เวลาที่ส่ง: {ts}", font=CTkFont(size=11), text_color="gray").pack(anchor="w")

                # สร้างปุ่มและวางในคอลัมน์ที่ 1
                start_button = CTkButton(card, text="เริ่มสร้าง PO", command=lambda s=row['so_number']: self._select_so_and_close(s))
                start_button.grid(row=0, column=1, sticky="e", padx=10, pady=5)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดรายการ SO ใหม่ได้: {e}", parent=self)

    def _load_in_progress_tasks(self):
        # ล้างข้อมูลเก่าออกจาก content frames
        for widget in self.so_in_progress_content_frame.winfo_children(): widget.destroy()
        for widget in self.po_draft_content_frame.winfo_children(): widget.destroy()
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # --- START: แก้ไข Query ตรงนี้ ---
                # เพิ่มเงื่อนไข NOT EXISTS เพื่อตรวจสอบว่ายังไม่มี PO ของ SO นี้ถูกสร้างขึ้น
                so_query = """
                    SELECT c.id, c.so_number, c.timestamp, c.customer_name 
                    FROM commissions c
                    WHERE c.status = 'PO In Progress' AND c.user_key = %s
                    AND NOT EXISTS (
                        SELECT 1 FROM purchase_orders po WHERE po.so_number = c.so_number
                    )
                    ORDER BY c.timestamp DESC
                """
                # --- END: สิ้นสุดการแก้ไข ---
                cursor.execute(so_query, (self.user_key,))
                claimed_sos = cursor.fetchall()

                # Query สำหรับ PO Drafts เหมือนเดิม
                po_query = "SELECT id, timestamp, so_number, po_number, supplier_name FROM purchase_orders WHERE user_key = %s AND status = 'Draft' ORDER BY timestamp DESC"
                cursor.execute(po_query, (self.user_key,))
                draft_pos = cursor.fetchall()

            if not claimed_sos:
                CTkLabel(self.so_in_progress_content_frame, text="ไม่มี SO ที่รอสร้าง PO ใบแรก").pack(pady=10)
            else:
                for so_data in claimed_sos:
                    card = CTkFrame(self.so_in_progress_content_frame, border_width=1)
                    card.pack(fill="x", padx=5, pady=3)
                    info = f"SO: {so_data['so_number']} - ลูกค้า: {so_data['customer_name']}"
                    CTkLabel(card, text=info, font=self.label_font).pack(side="left", padx=10, pady=5)
                    continue_button = CTkButton(card, text="ทำต่อ", command=lambda s=so_data['so_number']: self._continue_so_task(s))
                    continue_button.pack(side="right", padx=10, pady=5)

            if not draft_pos:
                CTkLabel(self.po_draft_content_frame, text="ไม่มี PO ฉบับร่าง").pack(pady=10)
            else:
                for po_data in draft_pos:
                    po_id = po_data['id']
                    card = CTkFrame(self.po_draft_content_frame, border_width=1); card.pack(fill="x", padx=5, pady=3)
                    card.grid_columnconfigure(0, weight=1); card.grid_columnconfigure(1, weight=0)
                    info_frame = CTkFrame(card, fg_color="transparent"); info_frame.grid(row=0, column=0, sticky="w", padx=10, pady=5)
                    action_frame = CTkFrame(card, fg_color="transparent"); action_frame.grid(row=0, column=1, sticky="e", padx=10, pady=5)
                    info = f"SO: {po_data['so_number']} | PO: {po_data['po_number']} | Supplier: {po_data['supplier_name']}"; CTkLabel(info_frame, text=info).pack(anchor="w")
                    CTkLabel(info_frame, text="สถานะ: ฉบับร่าง (ค้างส่ง)", font=CTkFont(size=12, weight="bold")).pack(anchor="w")
                    CTkButton(action_frame, text="แก้ไข", width=60, command=lambda p=po_id: self._edit_and_close(p)).pack(side="left", padx=2)
                    CTkButton(action_frame, text="ส่งอนุมัติ", width=80, fg_color="#16A34A", command=lambda p=po_id: self._submit_draft(p)).pack(side="left", padx=2)
                    CTkButton(action_frame, text="ลบ", width=40, fg_color="#D32F2F", hover_color="#B71C1C", command=lambda p=po_id: self._delete_draft(p)).pack(side="left", padx=2)
                    callback = lambda e, p=po_id: self._edit_and_close(p); card.bind("<Double-1>", callback)
                    for child in card.winfo_children(): child.bind("<Double-1>", callback)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดงานที่กำลังดำเนินการได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _load_rejected_po_tasks(self):
        frame = self.rejected_scroll_frame
        for widget in frame.winfo_children(): widget.destroy()
        try:
            query = "SELECT id, timestamp, so_number, po_number, supplier_name, status, rejection_reason FROM purchase_orders WHERE user_key = %s AND status = %s ORDER BY timestamp DESC"
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.user_key, "Rejected"))
            if df.empty: CTkLabel(frame, text="ไม่มีรายการที่ถูกปฏิเสธ").pack(pady=10); return
            for index, row in df.iterrows():
                po_id = row['id']
                card = CTkFrame(frame, border_width=1, fg_color="#FECACA"); card.pack(fill="x", padx=5, pady=3)
                info_frame = CTkFrame(card, fg_color="transparent"); info_frame.pack(fill="x", padx=10, pady=5)
                info = f"SO: {row['so_number']} | PO: {row['po_number']} | Supplier: {row['supplier_name']}"; CTkLabel(info_frame, text=info).pack(anchor="w")
                CTkLabel(info_frame, text=f"Last Update: {row['timestamp']}", font=CTkFont(size=11), text_color="gray50").pack(anchor="w")
                if pd.notna(row.get('rejection_reason')):
                    CTkLabel(card, text=f"เหตุผล: {row['rejection_reason']}", text_color="#B91C1C", wraplength=800, justify="left").pack(anchor="w", padx=10, pady=(0,5))
                edit_callback = lambda e, p=po_id: self._edit_and_close(p); card.bind("<Double-1>", edit_callback)
                for child in card.winfo_children(): child.bind("<Double-1>", edit_callback)
        except Exception as e: messagebox.showerror("Error", f"Error loading rejected PO tasks: {e}", parent=self)

    def _select_so_and_close(self, so_number):
        self.purchasing_screen.after(50, lambda: self.purchasing_screen.select_so_from_task(so_number))
        self.on_close()

    def _continue_so_task(self, so_number):
       self.purchasing_screen.so_entry.set(so_number) 
       self.purchasing_screen.after(50, lambda: self.purchasing_screen._on_so_selected(so_number, is_editing=True)) 
       self.on_close()

    def _edit_and_close(self, po_id):
        self.purchasing_screen.after(50, lambda: self.purchasing_screen._load_po_to_edit(po_id))
        self.purchasing_screen.tasks_window = None
        self.destroy()
            
    def _submit_draft(self, po_id):
        if not messagebox.askyesno("ยืนยันการส่ง", "คุณแน่ใจหรือไม่ที่จะส่งรายการนี้เพื่อขออนุมัติ?", icon="question", parent=self): return
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("UPDATE purchase_orders SET status = 'Pending Approval', approval_status = 'Pending Mgr 1', timestamp = %s WHERE id = %s", (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), po_id))
            conn.commit(); self.load_tasks()
        except Exception as e:
            if conn: conn.rollback(); messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _delete_draft(self, po_id):
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณแน่ใจหรือไม่ที่จะลบฉบับร่าง ID: {po_id}?", icon="warning", parent=self): return
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("DELETE FROM purchase_orders WHERE id = %s AND status = 'Draft'", (po_id,))
            conn.commit(); self.load_tasks()
        except Exception as e:
            if conn: conn.rollback(); messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการลบ: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

class ProductManagementWindow(CTkToplevel):
    def __init__(self, master, purchasing_screen_instance):
        super().__init__(master)
        self.purchasing_screen = purchasing_screen_instance
        self.app_container = purchasing_screen_instance.app_container
        self.user_key = purchasing_screen_instance.user_key
        self.label_font = purchasing_screen_instance.label_font
        self.entry_font = purchasing_screen_instance.entry_font

        self.title("จัดการข้อมูลสินค้าหลัก (Product Management)")
        self.geometry("800x600")
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(20, self.load_products)

        self.transient(master)
        self.grab_set()

    def on_close(self):
        self.purchasing_screen.product_management_window = None
        self.destroy()
        self.purchasing_screen._load_product_master_data()

    def _create_widgets(self):
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        CTkLabel(header_frame, text="จัดการข้อมูลสินค้าหลัก", font=CTkFont(size=18, weight="bold"))\
            .pack(side="left")

        button_frame = CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side="right")
        CTkButton(button_frame, text="เพิ่มสินค้าใหม่", command=self._add_product).pack(side="left", padx=5)
        CTkButton(button_frame, text="แก้ไขสินค้าที่เลือก", command=self._edit_product).pack(side="left", padx=5)
        CTkButton(button_frame, text="ลบสินค้าที่เลือก", command=self._delete_product, fg_color="#D32F2F", hover_color="#B71C1C").pack(side="left", padx=5)
        CTkButton(button_frame, text="Refresh", command=self.load_products).pack(side="left", padx=5)

        self.search_entry = CTkEntry(self, placeholder_text="ค้นหาสินค้า (รหัส/ชื่อ)", font=self.entry_font)
        self.search_entry.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        self.search_entry.bind("<KeyRelease>", self._filter_products)

        self.tree_frame = CTkFrame(self, fg_color="transparent")
        self.tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("id", "product_code", "product_name", "warehouse")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("id", text="ID", anchor="center")
        self.tree.heading("product_code", text="รหัสสินค้า", anchor="center")
        self.tree.heading("product_name", text="ชื่อสินค้า", anchor="center")
        self.tree.heading("warehouse", text="คลัง", anchor="center")

        self.tree.column("id", width=50, anchor="center")
        self.tree.column("product_code", width=150, anchor="w")
        self.tree.column("product_name", width=300, anchor="w")
        self.tree.column("warehouse", width=100, anchor="center")

        self.tree.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background="#F5F5F5",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="#F5F5F5")
        style.map('Treeview', background=[('selected', '#3B82F6')])
        style.configure("Treeview.Heading",
                        font=CTkFont(size=12, weight="bold"),
                        background="#E0E0E0",
                        foreground="black")

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)

    def load_products(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = self.app_container.get_connection()
        try:
            cursor_query = "SELECT id, product_code, product_name, warehouse FROM products ORDER BY product_code"
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute(cursor_query)
                products = cursor.fetchall()
                for prod in products:
                    self.tree.insert("", "end", values=(
                        prod['id'],
                        prod['product_code'],
                        prod['product_name'],
                        prod['warehouse']
                    ))
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดข้อมูลสินค้าได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn:
                self.app_container.release_connection(conn)

    def _filter_products(self, event):
        search_term = self.search_entry.get().strip().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = self.app_container.get_connection()
        try:
            query = "SELECT id, product_code, product_name, warehouse FROM products WHERE LOWER(product_code) LIKE %s OR LOWER(product_name) LIKE %s ORDER BY product_code"
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute(query, (f"%{search_term}%", f"%{search_term}%"))
                products = cursor.fetchall()
                for prod in products:
                    self.tree.insert("", "end", values=(
                        prod['id'],
                        prod['product_code'],
                        prod['product_name'],
                        prod['warehouse']
                    ))
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถค้นหาข้อมูลสินค้าได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn:
                self.app_container.release_connection(conn)


    def _add_product(self):
        ProductEditDialog(self, product_data=None, pm_window=self)

    def _edit_product(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("เลือกสินค้า", "กรุณาเลือกสินค้าที่ต้องการแก้ไข", parent=self)
            return

        values = self.tree.item(selected_item, 'values')
        product_id = values[0]

        conn = self.app_container.get_connection()
        try:
            cursor_query = "SELECT id, product_code, product_name, warehouse FROM products WHERE id = %s"
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute(cursor_query, (product_id,))
                product_data = cursor.fetchone()
                if product_data:
                    ProductEditDialog(self, product_data=product_data, pm_window=self)
                else:
                    messagebox.showerror("ข้อผิดพลาด", "ไม่พบข้อมูลสินค้าที่เลือก", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถดึงข้อมูลสินค้าเพื่อแก้ไขได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn:
                self.app_container.release_connection(conn)

    def _delete_product(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("เลือกสินค้า", "กรุณาเลือกสินค้าที่ต้องการลบ", parent=self)
            return

        values = self.tree.item(selected_item, 'values')
        product_id = values[0]
        product_code = values[1]
        product_name = values[2]

        if not messagebox.askyesno("ยืนยันการลบ", f"คุณแน่ใจหรือไม่ที่จะลบสินค้า '{product_code} - {product_name}'?", icon="warning", parent=self):
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("DELETE FROM products WHERE id = %s", (product_id,))
                conn.commit()
                messagebox.showinfo("สำเร็จ", f"ลบสินค้า '{product_code}' เรียบร้อยแล้ว", parent=self)
                self.load_products()
        except psycopg2.Error as e:
            if conn:
                conn.rollback()
            messagebox.showerror("Database Error", f"ไม่สามารถลบสินค้าได้: {e}\nอาจมีข้อมูล PO อ้างอิงถึงสินค้านี้", parent=self)
            traceback.print_exc()
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการลบ: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn:
                self.app_container.release_connection(conn)

# --- Product Edit Dialog ---
class ProductEditDialog(CTkToplevel):
    def __init__(self, master, product_data, pm_window):
        super().__init__(master)
        self.pm_window = pm_window
        self.app_container = pm_window.app_container
        self.product_data = product_data
        self.editing_mode = product_data is not None
        self.title("แก้ไขข้อมูลสินค้า" if self.editing_mode else "เพิ่มสินค้าใหม่")
        self.geometry("400x250")
        self.grid_columnconfigure(1, weight=1)

        self._create_widgets()
        if self.editing_mode:
            self._populate_form()

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(master)
        self.grab_set()

    def on_close(self):
        self.destroy()

    def _create_widgets(self):
        row = 0
        CTkLabel(self, text="รหัสสินค้า:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.product_code_entry = CTkEntry(self)
        self.product_code_entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        row += 1

        CTkLabel(self, text="ชื่อสินค้า:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.product_name_entry = CTkEntry(self)
        self.product_name_entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        row += 1

        CTkLabel(self, text="คลัง:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.warehouse_entry = CTkEntry(self)
        self.warehouse_entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        row += 1

        save_button_text = "บันทึกการแก้ไข" if self.editing_mode else "เพิ่มสินค้า"
        CTkButton(self, text=save_button_text, command=self._save_product).grid(row=row, column=0, columnspan=2, pady=20)

    def _populate_form(self):
        if self.product_data:
            self.product_code_entry.insert(0, self.product_data.get('product_code', ''))
            self.product_name_entry.insert(0, self.product_data.get('product_name', ''))
            self.warehouse_entry.insert(0, self.product_data.get('warehouse', ''))
            if self.editing_mode:
                self.product_code_entry.configure(state="readonly")

    def _save_product(self):
        code = self.product_code_entry.get().strip()
        name = self.product_name_entry.get().strip()
        warehouse = self.warehouse_entry.get().strip()
        if not code or not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกรหัสและชื่อสินค้า", parent=self)
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                if self.editing_mode:
                    product_id = self.product_data['id']
                    cursor.execute("""
                        UPDATE products SET product_name = %s, warehouse = %s, last_updated = %s
                        WHERE id = %s
                    """, (name, warehouse, datetime.now(), product_id))
                    messagebox.showinfo("สำเร็จ", f"อัปเดตสินค้า '{name}' เรียบร้อยแล้ว", parent=self)
                else:
                    cursor.execute("SELECT id FROM products WHERE product_code = %s", (code,))
                    if cursor.fetchone():
                        messagebox.showerror("ข้อมูลซ้ำ", "รหัสสินค้านี้มีอยู่ในระบบแล้ว", parent=self)
                        return
                    cursor.execute("""
                        INSERT INTO products (product_code, product_name, warehouse, last_updated)
                        VALUES (%s, %s, %s, %s)
                    """, (code, name, warehouse, datetime.now()))
                    messagebox.showinfo("สำเร็จ", f"เพิ่มสินค้าใหม่ '{name}' เรียบร้อยแล้ว", parent=self)
            conn.commit()
            self.pm_window.load_products()
            self.on_close()
        except psycopg2.Error as db_error:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {db_error}", parent=self)
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

# ==============================================================================
# PurchasingScreen Main Class
# ==============================================================================
class PurchasingScreen(CTkFrame):
    def __init__(self, master, user_key=None, user_name=None):
        self.master, self.app_container = master, master
        self.user_key, self.user_name = user_key, user_name
        self.theme = self.app_container.THEME["purchasing"]
        self.sale_theme = self.app_container.THEME["sale"]
        
        super().__init__(master, corner_radius=0, fg_color="#EDE9FE")
        self.shipping_to_stock_vat_var = tk.StringVar(value="VAT")
        self.shipping_to_site_vat_var = tk.StringVar(value="VAT")
        
        self.dropdown_style = {
            "fg_color": "white",
            "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }
        
        self.label_font = CTkFont(size=14, weight="bold", family="Roboto"); self.entry_font = CTkFont(size=14, family="Roboto"); self.header_font_table = CTkFont(size=14, weight="bold", family="Roboto")
        self.product_rows = []; self.payment_entries = {}
        self.editing_po_id, self.pg_engine = None, self.app_container.pg_engine
        self.current_commission_data = None
        self.supplier_data_map, self.supplier_display_list = {}, []
        self.editing_supplier_id = None
        self.product_data_map = {}; self.product_display_list = []
        self.po_mode_var = tk.StringVar(value="Single-PO")
        
        self.payment1_percent_var = tk.StringVar(value="ระบุยอดเอง")
        self.payment2_percent_var = tk.StringVar(value="ระบุยอดเอง")
        self.total_deposit_var = tk.StringVar(value="0.00")
        self.balance_due_var = tk.StringVar(value="0.00")
        
        self.sales_data_popup = None
        self.so_form_widgets = {}
        self._so_create_string_vars()
        self.shipping_to_stock_vat_display_var = tk.StringVar(value="0.00")
        self.shipping_to_site_vat_display_var = tk.StringVar(value="0.00")
        self.tasks_window = None
        self.product_management_window = None
        self.polling_job_id = None

        self._load_supplier_data()
        self._load_product_master_data()

        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        self._create_header()


        # --- START: โค้ดส่วนที่แก้ไข ---
        # ส่วนนี้คือส่วนที่แก้ไขแล้ว โดยการลบ PanedWindow ออกไป
        # และวาง po_pane ลงบนหน้าจอหลัก (self) โดยตรง
        self.po_pane = CTkFrame(self, fg_color="transparent")
        self.po_pane.grid_rowconfigure(0, weight=1)
        self.po_pane.grid_columnconfigure(0, weight=1)
        self.po_pane.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self._create_po_form_layout(self.po_pane)
        

        self._poll_and_update_tasks_badge()
        self.bind("<Destroy>", self._on_destroy)
    
    def _lookup_so_details(self):
        """
        เปิดหน้าต่างให้ผู้ใช้กรอก SO Number เพื่อค้นหาและแสดงรายละเอียด
        """
        dialog = CTkInputDialog(text="กรุณาใส่ SO Number ที่ต้องการค้นหา:", title="ค้นหาข้อมูล Sales Order")
        so_to_find = dialog.get_input()

        if so_to_find and so_to_find.strip():
            # เรียกใช้หน้าต่าง SODetailViewer ที่เราสร้างไว้ใน hr_windows.py
            SODetailViewer(
                master=self, 
                app_container=self.app_container, 
                so_number=so_to_find.strip().upper()
            )   
        elif so_to_find is not None: # ถ้าผู้ใช้กด OK แต่ไม่กรอกอะไร
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก SO Number", parent=self)

    def _update_summary(self, *args):
    # --- 1. คำนวณยอดรวมจากรายการสินค้า (Product Subtotal) ---
      product_subtotal = 0
      overall_total_weight = 0
      for row_dict in self.product_rows:
          try:
              if not row_dict["name"].winfo_exists(): continue
            
              qty = utils.convert_to_float(row_dict["qty"].get())
              price = utils.convert_to_float(row_dict["price"].get())
              weight = utils.convert_to_float(row_dict["weight"].get())
              discount_val = utils.convert_to_float(row_dict["discount_entry"].get())
              discount_type = row_dict["discount_type_var"].get()

              line_total = qty * price
              discount_amount = (line_total * (discount_val / 100.0)) if discount_type == "%" else discount_val
              row_final_price = line_total - discount_amount
              row_final_weight = qty * weight

              product_subtotal += row_final_price
              overall_total_weight += row_final_weight

            # อัปเดต UI ของแต่ละแถว
              for entry, value in [(row_dict["total_price"], row_final_price), (row_dict["total_weight"], row_final_weight)]:
                  entry.configure(state="normal")
                  entry.delete(0, tk.END)
                  entry.insert(0, f"{value:,.2f}")
                  entry.configure(state="readonly")
          except (ValueError, tk.TclError):
              continue

    # --- 2. ดึงข้อมูลตัวเลขทั้งหมดจากฟอร์ม ---
      shipping_stock_cost = utils.convert_to_float(self.shipping_to_stock_cost_entry.get())
      shipping_site_cost = utils.convert_to_float(self.shipping_to_site_cost_entry.get())
      end_of_bill_discount = utils.convert_to_float(self.end_of_bill_discount_entry.get())
      p1 = utils.convert_to_float(self.payment_entries["Payment 1"]["amount"].get())
      p2 = utils.convert_to_float(self.payment_entries["Payment 2"]["amount"].get())
      full_payment = utils.convert_to_float(self.payment_entries["Full Payment"]["amount"].get())
    
    # --- 3. คำนวณยอดที่ต้องชำระให้ซัพพลายเออร์ และค่าส่งที่จ่ายแยก ---
      supplier_payable_vatable = product_subtotal - end_of_bill_discount
      supplier_payable_non_vatable = 0.0
      separate_shipping_cost = 0.0

    # ตรวจสอบค่าส่งเข้าสต๊อก
      if self.shipping_to_stock_type_var.get() == 'ซัพพลายเออร์จัดส่ง':
          if self.shipping_to_stock_vat_var.get() == 'VAT':
              supplier_payable_vatable += shipping_stock_cost
          else:
              supplier_payable_non_vatable += shipping_stock_cost
      else:
          separate_shipping_cost += shipping_stock_cost

    # ตรวจสอบค่าส่งเข้าไซต์
      if self.shipping_to_site_type_var.get() == 'ซัพพลายเออร์จัดส่ง':
          if self.shipping_to_site_vat_var.get() == 'VAT':
              supplier_payable_vatable += shipping_site_cost
          else:
              supplier_payable_non_vatable += shipping_site_cost
      else:
          separate_shipping_cost += shipping_site_cost

    # --- 4. คำนวณ VAT, WHT, ยอดสุทธิ และยอดค้างชำระ ---
      vat7_amount = supplier_payable_vatable * 0.07 if hasattr(self, 'vat_checkbox') and self.vat_checkbox.get() else 0.0
      wht3_amount = supplier_payable_vatable * 0.03 if hasattr(self, 'vat3_checkbox') and self.vat3_checkbox.get() else 0.0
    
      grand_total_payable_to_supplier = (supplier_payable_vatable + vat7_amount - wht3_amount) + supplier_payable_non_vatable
      total_deposit = p1 + p2
      balance_due = grand_total_payable_to_supplier - total_deposit - full_payment

    # --- 5. อัปเดต UI ทั้งหมดในส่วนสรุป ---
      def set_readonly_val(entry, value):
          if entry and entry.winfo_exists():
             entry.configure(state="normal")
             entry.delete(0, "end")
             entry.insert(0, f"{value:,.2f}")
             entry.configure(state="readonly")
 
      total_po_cost = (product_subtotal - end_of_bill_discount) + shipping_stock_cost + shipping_site_cost
      set_readonly_val(self.total_cost_entry, total_po_cost)
      set_readonly_val(self.total_weight_summary_entry, overall_total_weight)
      set_readonly_val(self.vat7_entry, vat7_amount)
      set_readonly_val(self.vat3_entry, wht3_amount)
      set_readonly_val(self.grand_total_with_vat_entry, supplier_payable_vatable + supplier_payable_non_vatable + vat7_amount)
      set_readonly_val(self.grand_total_payable_entry, grand_total_payable_to_supplier)
      set_readonly_val(self.separate_shipping_entry, separate_shipping_cost)
  
      self.total_deposit_var.set(f"{total_deposit:,.2f}")
    
      stock_vat_display = shipping_stock_cost * 0.07 if self.shipping_to_stock_vat_var.get() == 'VAT' else 0.0
      site_vat_display = shipping_site_cost * 0.07 if self.shipping_to_site_vat_var.get() == 'VAT' else 0.0
      self.shipping_to_stock_vat_display_var.set(f"{stock_vat_display:,.2f}")
      self.shipping_to_site_vat_display_var.set(f"{site_vat_display:,.2f}")

      if hasattr(self, 'balance_due_entry') and self.balance_due_entry.winfo_exists():
          if abs(balance_due) < 0.01:
              text, text_color, bg_color = "ยอดชำระครบถ้วน", "#15803D", "#BBF7D0"
          elif balance_due < 0:
              text, text_color, bg_color = f"ชำระเกิน {abs(balance_due):,.2f}", "#15803D", "#BBF7D0"
          else:
              text, text_color, bg_color = f"ยอดค้างชำระ {balance_due:,.2f}", "#B91C1C", "#FECACA"
        
          self.balance_due_var.set(text)
          self.balance_due_entry.configure(text_color=text_color, fg_color=bg_color)
      else:
          self.balance_due_var.set(f"{balance_due:,.2f}")

    def _so_create_string_vars(self):
        now = datetime.now()
        self.so_form_widgets['thai_months'] = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.so_form_widgets['thai_month_map'] = {name: i+1 for i, name in enumerate(self.so_form_widgets['thai_months'])}
        
        self.so_form_widgets['customer_type_var'] = tk.StringVar(value="ลูกค้าเก่า")
        self.so_form_widgets['credit_term_var'] = tk.StringVar(value="เงินสด")
        self.so_form_widgets['commission_month_var'] = tk.StringVar(value=self.so_form_widgets['thai_months'][now.month - 1])
        self.so_form_widgets['commission_year_var'] = tk.StringVar(value=str(now.year + 543))
        self.so_form_widgets['payment1_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_form_widgets['payment2_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_form_widgets['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['payment_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_subtotal_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vat_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_form_widgets['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_product_input_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_service_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_actual_payment_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_verification_result_var'] = tk.StringVar(value="-")

        self.so_form_widgets['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")

        self.rr_number_var = tk.StringVar(value="RR")
        self.payment1_method_var = tk.StringVar(value="ชำระสด")
        self.payment2_method_var = tk.StringVar(value="ชำระสด")
        self.delivery_type_var = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        
        self.so_form_widgets['rr_number_var'] = self.rr_number_var
        self.so_form_widgets['payment1_method_var'] = self.payment1_method_var
        self.so_form_widgets['payment2_method_var'] = self.payment2_method_var
        self.so_form_widgets['delivery_type_var'] = self.delivery_type_var

    def _create_header(self):
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,5))
        CTkLabel(header_frame, text=f"ฝ่ายจัดซื้อ: {self.user_name} (ID: {self.user_key})", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).pack(side="left")

        button_container = CTkFrame(header_frame, fg_color="transparent")
        button_container.pack(side="right")
        
        self.tasks_button = CTkButton(button_container, text="My Tasks 🔔 (0)", command=self._open_my_tasks_window)
        self.tasks_button.pack(side="left", padx=(0, 5))

        # <<< START: เพิ่มปุ่มใหม่ตรงนี้ >>>
        CTkButton(button_container, text="🔍 ค้นหา SO", command=self._lookup_so_details, fg_color="#0891B2").pack(side="left", padx=5)
        # <<< END >>>

        CTkButton(button_container, text="📖 ดูประวัติ PO", command=lambda: self.app_container.show_history_window(), fg_color="#64748B").pack(side="left", padx=5)
        CTkButton(button_container, text="🔧 จัดการสินค้า", command=self._open_product_management_window, fg_color="#6D28D9", hover_color="#5B21B6").pack(side="left", padx=5)
        
        # ... (โค้ดส่วนที่เหลือของฟังก์ชันเหมือนเดิม) ...
        CTkButton(button_container, text="Export PDF (PO อนุมัติ)", command=lambda: export_approved_pos_to_pdf(self, self.pg_engine), fg_color="#c026d3", hover_color="#a21caf").pack(side="left", padx=5)
        export_button = CTkButton(button_container, text="Export Excel (PO อนุมัติ)", command=lambda: export_approved_pos_to_excel(self, self.pg_engine), fg_color="#107C41", hover_color="#0B532B")
        export_button.pack(side="left", padx=5)
        CTkButton(button_container, text="(ล้างฟอร์ม PO)", command=self.handle_clear_button_press, fg_color="#E11D48").pack(side="left", padx=5)
        self.toggle_so_data_button = CTkButton(button_container, text="ดูข้อมูล SO", command=self._open_so_popup, fg_color=self.sale_theme.get("primary", "#3B82F6"))
        self.toggle_so_data_button.pack(side="left", padx=5)
        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").pack(side="right", padx=(5, 0))
    
    def _open_so_selection_dialog(self):
        """
        ฟังก์ชันนี้จะเรียกใช้ฟังก์ชันหลักใน main_app.py
        เพื่อเปิดหน้าต่างสำหรับเลือก SO มาพิมพ์
        """
        self.app_container.open_so_print_dialog()

    def _open_po_selection_dialog(self):
        """
        ฟังก์ชันนี้จะเปิดหน้าต่างสำหรับเลือก PO ใบเดียวมาพิมพ์
        และส่ง callback ไปยังฟังก์ชันที่ถูกต้องใน main_app.py
        """
        dialog = POSelectionDialog(
            master=self, 
            pg_engine=self.app_container.pg_engine, 
            print_callback=self.app_container.generate_single_po_document 
        )

    def _on_destroy(self, event):
        if hasattr(event, 'widget') and event.widget is self:
            self._stop_polling()
            if self.sales_data_popup and self.sales_data_popup.winfo_exists():
                self.sales_data_popup._on_popup_close()
                self.sales_data_popup = None
            if self.tasks_window and self.tasks_window.winfo_exists():
                self.tasks_window.destroy()
                self.tasks_window = None
            if self.product_management_window and self.product_management_window.winfo_exists():
                self.product_management_window.destroy()
                self.product_management_window = None

    def _stop_polling(self):
        if self.polling_job_id: self.after_cancel(self.polling_job_id); self.polling_job_id = None
            
    def _poll_and_update_tasks_badge(self):
        self._update_tasks_badge()
        self.polling_job_id = self.after(30000, self._poll_and_update_tasks_badge)

    def _update_tasks_badge(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM notifications WHERE user_key_to_notify = %s AND is_read = FALSE AND message LIKE 'SO ใหม่รอสร้าง PO%%'", (self.user_key,)); new_so_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM purchase_orders WHERE user_key = %s AND status = 'Rejected'", (self.user_key,)); rejected_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM purchase_orders WHERE user_key = %s AND status = 'Draft'", (self.user_key,)); draft_count = cursor.fetchone()[0]
            total_tasks = new_so_count + rejected_count + draft_count
            if hasattr(self, 'tasks_button') and self.tasks_button.winfo_exists():
                self.tasks_button.configure(text=f"My Tasks 🔔 ({total_tasks})")
                if total_tasks > 0: self.tasks_button.configure(fg_color="#F59E0B", hover_color="#D97706")
                else: self.tasks_button.configure(fg_color=("#3B8ED0", "#1F6AA5"), hover_color=("#36719F", "#144870"))
        except Exception as e: print(f"Error updating tasks badge: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _open_my_tasks_window(self):
        try:
            if self.tasks_window and self.tasks_window.winfo_exists():
                self.tasks_window.lift()
                self.tasks_window.focus()
                return
        except (tk.TclError, AttributeError):
            self.tasks_window = None

        self.tasks_window = MyTasksWindow(self, purchasing_screen_instance=self)

    def _open_product_management_window(self):
        if self.product_management_window is None or not self.product_management_window.winfo_exists():
            self.product_management_window = ProductManagementWindow(self, purchasing_screen_instance=self)
        else:
            self.product_management_window.focus()

    def select_so_from_task(self, so_number):
        if self.so_entry.winfo_exists():
            self.so_entry.set(so_number)
        
        self._on_so_selected(so_number)
        
        self._update_tasks_badge()
    
    def _open_so_popup(self):
        if self.current_commission_data is None: messagebox.showinfo("ข้อมูล SO", "กรุณาเลือก SO Number ก่อน", parent=self); return
        if self.sales_data_popup and self.sales_data_popup.winfo_exists(): self.sales_data_popup.focus(); return
        self.sales_data_popup = SOPopupWindow(self, sales_data=self.current_commission_data, so_shared_vars=self.so_form_widgets, sale_theme=self.sale_theme)
    
    def _save_so_changes_from_popup(self, so_id, so_shared_vars_data, current_popup_widgets_ref):
        updated_data = {}
        db_key_map = {
            'bill_date_selector': 'bill_date', 'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id',
            'credit_term_entry': 'credit_term', 'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost', 'delivery_date': 'delivery_date',
            'credit_card_fee_entry': 'credit_card_fee', 'transfer_fee_entry': 'transfer_fee', 'wht_3_percent': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons', 'giveaways': 'giveaways',
            'payment_date': 'payment_date_selector', 'cash_product_input': 'cash_product_input_entry', 'cash_actual_payment': 'cash_actual_payment',
            'sales_service_vat_option': 'sales_service_vat_option', 'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option': 'other_service_fee_vat_option', 'shipping_vat_option_var': 'shipping_vat_option',
            'credit_card_fee_vat_option': 'credit_card_fee_vat_option_var', 'so_grand_total_var': 'so_grand_total',
            'total_payment_amount': 'total_payment_amount', 'difference_amount_var': 'difference_amount'
        }

        def _safe_get_float(entry_widget):
            if entry_widget and hasattr(entry_widget, 'winfo_exists') and entry_widget.winfo_exists():
                try: return utils.convert_to_float(entry_widget.get())
                except (ValueError, tk.TclError): return 0.0
            return 0.0

        for widget_key, db_col_name in db_key_map.items():
            value = None
            if widget_key in current_popup_widgets_ref:
                widget_instance = current_popup_widgets_ref[widget_key]
                if isinstance(widget_instance, NumericEntry): value = _safe_get_float(widget_instance)
                elif isinstance(widget_instance, DateSelector): value = widget_instance.get_date() if widget_instance.winfo_exists() else None
                elif isinstance(widget_instance, (CTkEntry, AutoCompleteEntry)): value = widget_instance.get().strip() or None if widget_instance.winfo_exists() else None
                elif isinstance(widget_instance, CTkLabel):
                    if db_col_name in so_shared_vars_data and isinstance(so_shared_vars_data[db_col_name], tk.StringVar): value = so_shared_vars_data[db_col_name].get()
                    elif widget_instance.winfo_exists(): value = widget_instance.cget("text").strip() or None
                    else: value = None
            elif widget_key in so_shared_vars_data and isinstance(so_shared_vars_data[widget_key], tk.StringVar): value = so_shared_vars_data[widget_key].get()
            if value is not None: updated_data[db_col_name] = value
        
        p1 = _safe_get_float(current_popup_widgets_ref.get('payment1_amount_entry'))
        p2 = _safe_get_float(current_popup_widgets_ref.get('payment2_amount_entry'))
        updated_data['total_payment_amount'] = p1 + p2

        updated_data['so_grand_total'] = utils.convert_to_float(so_shared_vars_data['so_grand_total_var'].get())
        updated_data['difference_amount'] = utils.convert_to_float(so_shared_vars_data['difference_amount_var'].get())

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'commissions'")
                db_columns = {row[0] for row in cursor.fetchall()}
            
            final_data_to_save = {k: v for k, v in updated_data.items() if k in db_columns}
            set_clauses = [f'"{k}" = %s' for k, v in final_data_to_save.items()]
            params = list(final_data_to_save.values()) + [so_id]
            sql_update = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
            
            with conn.cursor() as cursor: cursor.execute(sql_update, tuple(params))
            conn.commit()
            
            messagebox.showinfo("สำเร็จ", f"บันทึกทับข้อมูล SO Number: {self.current_commission_data.get('so_number')} เรียบร้อยแล้ว", parent=self)
            
            reloaded_df = pd.read_sql_query("SELECT * FROM commissions WHERE id = %s", self.app_container.pg_engine, params=(so_id,))
            if not reloaded_df.empty: self.current_commission_data = reloaded_df.iloc[0].to_dict()
            else: self.current_commission_data = None
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล SO จาก Pop-up:\n{e}\n{traceback.format_exc()}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)


    def _on_so_selected(self, so_number: str, is_editing: bool = False):
        if not so_number: self.handle_clear_button_press(confirm=False); return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT id, status, user_key FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", (so_number,))
                so_id_in_commissions, so_status, so_user_key = cursor.fetchone() if cursor.rowcount > 0 else (None, None, None)

                if not is_editing:
                    if so_id_in_commissions is None:
                        messagebox.showwarning("ไม่พบ SO", f"ไม่พบ SO Number: {so_number} ในสถานะที่พร้อมดำเนินการ", parent=self); self.so_entry.set(""); return
                    if so_status == 'Pending PU':
                        # <<< เพิ่มการบันทึก claim_timestamp >>>
                        cursor.execute("UPDATE commissions SET status = 'PO In Progress', user_key = %s, claim_timestamp = %s WHERE id = %s", (self.user_key, datetime.now(), so_id_in_commissions)); conn.commit()
                        messagebox.showinfo("Claim SO", f"คุณได้ Claim SO: {so_number} เพื่อดำเนินการสร้าง PO แล้ว", parent=self)
                    elif so_status == 'PO In Progress' and so_user_key == self.user_key:
                        pass
                    elif so_status == 'PO In Progress' and so_user_key != self.user_key:
                        messagebox.showwarning("SO ถูกเลือกไปแล้ว", f"SO: {so_number} ถูกผู้ใช้งานอื่น (User ID: {so_user_key}) เลือกไปแล้ว", parent=self)
                        self.so_entry.configure(values=self._get_commission_so_numbers()); self.so_entry.set(""); return
                    else:
                        messagebox.showwarning("SO ไม่พร้อม", f"SO: {so_number} อยู่ในสถานะ '{so_status}' ไม่สามารถสร้าง PO ได้", parent=self)
                        self.so_entry.configure(values=self._get_commission_so_numbers()); self.so_entry.set(""); return
                
            df = pd.read_sql_query("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", self.pg_engine, params=(so_number,))
            if not df.empty:
                self.current_commission_data = df.iloc[0].to_dict(); self._open_so_popup()
            else:
                self.current_commission_data = None; messagebox.showerror("ไม่พบข้อมูล SO", f"ไม่พบข้อมูลสำหรับ SO Number: {so_number}", parent=self)
                self.so_entry.set("")
                if self.sales_data_popup and self.sales_data_popup.winfo_exists(): self.sales_data_popup.destroy(); self.sales_data_popup = None
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดระหว่างการเลือก SO: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def handle_clear_button_press(self, confirm=True):
        if confirm and not messagebox.askyesno("ยืนยัน", "คุณต้องการล้างข้อมูลทั้งหมดในฟอร์มใช่หรือไม่?", parent=self): return
        if self.current_commission_data:
            so_id_to_release = self.current_commission_data.get('id')
            so_number_to_release = self.current_commission_data.get('so_number')
            conn = self.app_container.get_connection()
            try:
                with conn.cursor() as cursor:
                    # <<< เพิ่มการลบ claim_timestamp เมื่อยกเลิกการทำงาน >>>
                    cursor.execute("UPDATE commissions SET status = 'Pending PU', user_key = NULL, claim_timestamp = NULL WHERE id = %s AND status = 'PO In Progress' AND user_key = %s", (so_id_to_release, self.user_key)); conn.commit()
                    self.so_entry.configure(values=self._get_commission_so_numbers())
            except Exception as e:
                if conn: conn.rollback(); print(f"Error releasing SO status: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)
        
        self._clear_form(confirm=False)

    def _clear_form(self, confirm=True, keep_so=False):
        if confirm and not messagebox.askyesno("ยืนยัน", "คุณต้องการล้างข้อมูลทั้งหมดในฟอร์มใช่หรือไม่?", parent=self):
            return
        
        if not keep_so:
            self.so_entry.set("")
            self.current_commission_data = None
            if self.sales_data_popup and self.sales_data_popup.winfo_exists():
                self.sales_data_popup.destroy()
                self.sales_data_popup = None

        self.editing_po_id = None
        self.shipping_to_stock_vat_var.set("VAT")
        self.shipping_to_site_vat_var.set("VAT")
        self.department_entry.delete(0, 'end')
        self.pur_order_entry.delete(0, 'end')
        self.po_number_type_var.set("PO")
        self.po_number_input_var.set("")
        self.rr_number_var.set("RR")
        self._validate_po_input()

        self.supplier_name_combo.delete(0, 'end')
        self.supplier_code_entry.delete(0, 'end')
        self.credit_term_entry.delete(0, 'end')
        
        for row in self.product_rows:
            for widget in row["widgets"]:
                widget.destroy()
        self.product_rows.clear()
        self._add_product_row()
        
        entries_to_clear = [
            self.shipping_to_stock_cost_entry, self.shipping_to_stock_notes_entry,
            self.shipping_to_site_cost_entry, self.shipping_to_site_notes_entry,
            self.total_weight_summary_entry, self.total_cost_entry,
            self.end_of_bill_discount_entry,
            self.vat3_entry, self.vat7_entry, self.grand_total_with_vat_entry, self.grand_total_payable_entry
        ]
        for entry in entries_to_clear:
            if hasattr(entry, 'winfo_exists') and entry.winfo_exists():
                is_readonly = entry.cget("state") == "readonly"
                if is_readonly: entry.configure(state="normal")
                entry.delete(0, "end")
                if is_readonly: entry.configure(state="readonly")
        
        # Clear payment fields
        for p_type in ["Payment 1", "Payment 2", "Full Payment"]:
            p_dict = self.payment_entries.get(p_type)
            if p_dict and p_dict["amount"].winfo_exists():
                p_dict['amount'].delete(0, "end")
            if p_dict and p_dict.get("date") and p_dict["date"].winfo_exists():
                p_dict["date"].set_date(None)
            if p_dict and p_dict.get("percent_var"):
                p_dict["percent_var"].set("ระบุยอดเอง")

        if hasattr(self, 'cn_refund_amount_entry') and self.cn_refund_amount_entry.winfo_exists():
            self.cn_refund_amount_entry.delete(0, 'end')
            self.cn_refund_date_selector.set_date(None)
            
        self.total_deposit_var.set("0.00")
        self.balance_due_var.set("0.00")

        if hasattr(self, 'vat3_checkbox'): self.vat3_checkbox.deselect()
        if hasattr(self, 'vat_checkbox'): self.vat_checkbox.deselect()
            
        self._update_summary()

    def _create_po_form_layout(self, parent):
        self.purchasing_form_frame = CTkScrollableFrame(parent, corner_radius=10, fg_color="#D6D7D8", label_text="ฟอร์มใบสั่งซื้อ (PO)")
        self.purchasing_form_frame.pack(fill="both", expand=True)
        self.purchasing_form_frame.grid_columnconfigure(0, weight=1)
        self._create_top_info_frame(self.purchasing_form_frame)
        self._create_product_grid(self.purchasing_form_frame)
        self._create_bottom_summary_frame(self.purchasing_form_frame)
        self._create_footer_frame(self.purchasing_form_frame)
    
    def _create_top_info_frame(self, parent):
        top_frame = CTkFrame(parent, fg_color="#F9FAFB")
        top_frame.pack(fill="x", padx=10, pady=10)
        
        top_frame.grid_columnconfigure(1, weight=1)
        top_frame.grid_columnconfigure(3, weight=1)
        top_frame.grid_columnconfigure(5, weight=1)
        top_frame.grid_columnconfigure(7, weight=1)

        CTkLabel(top_frame, text="SO Number:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.so_entry = CTkComboBox(top_frame, values=self._get_commission_so_numbers(), command=self._on_so_selected)
        self.so_entry.set("")
        self.so_entry.grid(row=0, column=1, columnspan=7, sticky="ew", padx=10, pady=5)

        CTkLabel(top_frame, text="เอกสาร PO/ST:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        po_st_frame = CTkFrame(top_frame, fg_color="transparent")
        po_st_frame.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        po_st_frame.grid_columnconfigure(1, weight=1)

        self.po_number_type_var = tk.StringVar(value="PO")
        self.po_number_type_var.trace_add("write", self._on_po_number_type_changed)
        self.po_type_dropdown = CTkOptionMenu(po_st_frame, variable=self.po_number_type_var, values=["PO", "ST"], width=80, **self.dropdown_style)
        self.po_type_dropdown.grid(row=0, column=0, sticky="w", padx=(0,5))

        self.po_number_input_var = tk.StringVar()
        self.po_number_input_var.trace_add("write", self._validate_po_input)
        self.po_number_entry = CTkEntry(po_st_frame, font=self.entry_font, textvariable=self.po_number_input_var)
        self.po_number_entry.grid(row=0, column=1, sticky="ew")
        self.po_number_entry.bind("<FocusIn>", self._on_po_focus_in)
        
        CTkLabel(top_frame, text="RR Number:").grid(row=1, column=2, sticky="w", padx=10, pady=5)
        self.rr_number_var.trace_add("write", self._force_uppercase_rr)
        self.rr_number_entry = CTkEntry(top_frame, font=self.entry_font, textvariable=self.rr_number_var)
        self.rr_number_entry.grid(row=1, column=3, sticky="ew", padx=10, pady=5)

        CTkLabel(top_frame, text="แผนก:").grid(row=1, column=4, sticky="w", padx=(20, 10), pady=5)
        self.department_entry = CTkEntry(top_frame, font=self.entry_font)
        self.department_entry.grid(row=1, column=5, sticky="ew", padx=10, pady=5)

        CTkLabel(top_frame, text="PUR Order :").grid(row=1, column=6, sticky="w", padx=(20, 10), pady=5)
        self.pur_order_entry = CTkEntry(top_frame, font=self.entry_font)
        self.pur_order_entry.grid(row=1, column=7, sticky="ew", padx=10, pady=5)

        sup_frame = CTkFrame(top_frame, fg_color="transparent")
        sup_frame.grid(row=2, column=0, columnspan=8, sticky="ew", padx=5, pady=5)
        sup_frame.grid_columnconfigure(1, weight=4)
        sup_frame.grid_columnconfigure(3, weight=2)
        sup_frame.grid_columnconfigure(5, weight=2)
        sup_frame.grid_columnconfigure(6, weight=1)
        sup_frame.grid_columnconfigure(7, weight=1)

        CTkLabel(sup_frame, text="Supplier Name:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        
        # --- START: แก้ไขบรรทัดนี้ ---
        # เปลี่ยนจาก self.supplier_display_list เป็น self.supplier_completion_data
        self.supplier_name_combo = AutoCompleteEntry(sup_frame, 
                                                     completion_list=self.supplier_completion_data, 
                                                     command_on_select=self._on_supplier_selected, 
                                                     placeholder_text="พิมพ์เพื่อค้นหาซัพพลายเออร์...")
        # --- END: สิ้นสุดการแก้ไข ---
        
        self.supplier_name_combo.grid(row=0, column=1, sticky="ew", padx=(0,10), pady=3)
        
        CTkLabel(sup_frame, text="Supplier Code:").grid(row=0, column=2, sticky="w", padx=5, pady=3)
        self.supplier_code_entry = CTkEntry(sup_frame, font=self.entry_font)
        self.supplier_code_entry.grid(row=0, column=3, sticky="ew", padx=(0,10), pady=3)

        CTkLabel(sup_frame, text="Credit Term:").grid(row=0, column=4, sticky="w", padx=5, pady=3)
        self.credit_term_entry = CTkEntry(sup_frame, font=self.entry_font)
        self.credit_term_entry.grid(row=0, column=5, sticky="ew", padx=(0,10), pady=3)

        self.update_supplier_button = CTkButton(sup_frame, text="บันทึก/อัปเดต", width=120, command=self._save_or_update_supplier)
        self.update_supplier_button.grid(row=0, column=6, sticky="e", padx=(5, 10), pady=3)

        mode_frame = CTkFrame(sup_frame, fg_color="transparent")
        mode_frame.grid(row=0, column=7, sticky="e", padx=(10, 5), pady=3)
        CTkRadioButton(mode_frame, text="Single PO/ST", variable=self.po_mode_var, value="Single-PO").pack(side="left")
        CTkRadioButton(mode_frame, text="Multiple PO/ST", variable=self.po_mode_var, value="Multiple-PO").pack(side="left", padx=10)
    
    def _on_po_number_type_changed(self, *args): self._validate_po_input()

    def _validate_po_input(self, *args):
        current_text = self.po_number_input_var.get(); selected_type = self.po_number_type_var.get(); new_text = current_text.upper()
        if not new_text: new_text = selected_type
        elif new_text.startswith(selected_type): pass
        elif (selected_type == "PO" and new_text.startswith("ST")) or (selected_type == "ST" and new_text.startswith("PO")): new_text = selected_type + new_text[2:]
        else: new_text = selected_type + new_text
        if new_text != current_text: self.po_number_input_var.set(new_text); self.after(10, lambda: self.po_number_entry.icursor(tk.END))

    def _force_uppercase_rr(self, *args):
        current_text = self.rr_number_var.get(); new_text = current_text.upper()
        if not new_text.startswith("RR"): new_text = "RR" if new_text == "" else "RR" + new_text
        if new_text != current_text: self.rr_number_var.set(new_text); self.after(10, lambda: self.rr_number_entry.icursor(tk.END))

    def _on_po_focus_in(self, event):
        current_text = self.po_number_input_var.get(); selected_type = self.po_number_type_var.get()
        if not current_text.startswith(selected_type): self.po_number_input_var.set(selected_type + current_text)
        self.po_number_entry.icursor(tk.END)
            
    def _on_supplier_selected(self, selection_dict: dict):
        # --- START: เพิ่ม Logic แก้ไขข้อความในช่องค้นหา ---
        # 1. ลบข้อความยาวๆ (ชื่อ + รหัส) ที่ AutoComplete ใส่เข้ามาทิ้งไป
        self.supplier_name_combo.delete(0, tk.END)
        # 2. ใส่เฉพาะ "ชื่อ" กลับเข้าไปใหม่
        self.supplier_name_combo.insert(0, selection_dict.get('name', ''))
        # --- END: สิ้นสุดการเพิ่ม Logic ---

        selected_display_name = selection_dict.get('display', '')
        
        self.supplier_code_entry.delete(0, tk.END)
        self.credit_term_entry.delete(0, tk.END)
        self.editing_supplier_id = None

        if selected_display_name in self.supplier_data_map:
            supplier = self.supplier_data_map[selected_display_name]
            self.editing_supplier_id = supplier.get('id')
            
            self.supplier_code_entry.insert(0, supplier.get('code', ''))
            
            credit_term_map = {'เงินสด': 'เงินสด', '0': 'เงินสด', '7': 'Cr 7', '15': 'Cr 15', '30': 'Cr 30'}
            term_value = str(supplier.get('term', 'เงินสด')).strip()
            self.credit_term_entry.insert(0, credit_term_map.get(term_value, term_value))
            
    def _save_or_update_supplier(self):
        name, code, term = self.supplier_name_combo.get().strip(), self.supplier_code_entry.get().strip(), self.credit_term_entry.get().strip()
        if not name or not code: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกชื่อและรหัสซัพพลายเออร์", parent=self); return
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT id FROM suppliers WHERE supplier_code = %s", (code,)); existing_by_code = cursor.fetchone()
                is_update = False
                for k, v in self.supplier_data_map.items():
                    if v['supplier_name'] == name: self.editing_supplier_id = v['id']; is_update = True; break
                if is_update:
                    if existing_by_code and existing_by_code['id'] != self.editing_supplier_id: messagebox.showerror("ข้อมูลซ้ำ", "รหัสซัพพลายเออร์นี้ถูกใช้แล้ว", parent=self); return
                    cursor.execute("UPDATE suppliers SET supplier_name = %s, supplier_code = %s, credit_term = %s WHERE id = %s", (name, code, term, self.editing_supplier_id)); messagebox.showinfo("สำเร็จ", f"อัปเดตข้อมูล '{name}' เรียบร้อยแล้ว", parent=self)
                else:
                    if existing_by_code: messagebox.showerror("ข้อมูลซ้ำ", "รหัสซัพพลายเออร์นี้มีอยู่แล้ว", parent=self); return
                    cursor.execute("INSERT INTO suppliers (supplier_name, supplier_code, credit_term) VALUES (%s, %s, %s)", (name, code, term)); messagebox.showinfo("สำเร็จ", f"เพิ่มซัพพลายเออร์ใหม่ '{name}' เรียบร้อยแล้ว", parent=self)
            conn.commit()
        except Exception as e:
            if conn: conn.rollback(); messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally: self.app_container.release_connection(conn); self._load_supplier_data()
    
    def _on_product_selected(self, selection_dict, row_widgets):
        if not selection_dict: return

        # --- START: เพิ่ม Logic แก้ไขข้อความในช่องค้นหา ---
        # 1. ดึงวิดเจ็ตช่องค้นหาของแถวนั้นๆ ออกมา
        code_entry_widget = row_widgets.get("code")
        # 2. ลบข้อความยาวๆ ทิ้งไป แล้วใส่เฉพาะ "รหัสสินค้า" กลับเข้าไป
        if code_entry_widget:
            code_entry_widget.delete(0, tk.END)
            code_entry_widget.insert(0, selection_dict.get('code', ''))
        # --- END: สิ้นสุดการเพิ่ม Logic ---
        
        code = selection_dict.get('code')
        product_data = self.product_data_map.get(code)
        
        if product_data:
            row_widgets['master_data'] = {
                "product_name": product_data.get("name", ""),
                "warehouse": product_data.get("warehouse", "")
            }
            
            row_widgets["name_var"].set(str(product_data.get("name") or ""))
            row_widgets["warehouse_var"].set(str(product_data.get("warehouse") or ""))
            
            if row_widgets["price"].winfo_exists(): row_widgets["price"].delete(0, "end")
            if row_widgets["weight"].winfo_exists(): row_widgets["weight"].delete(0, "end")
            
            self._update_summary()
            self._check_for_override(row_widgets)

    def _check_for_override(self, row_dict):
        master_data = row_dict.get('master_data')
        if not master_data:
            return

        current_name = row_dict["name_var"].get()
        original_name = master_data.get("product_name", "")
        is_name_changed = current_name != original_name

        current_warehouse = row_dict["warehouse_var"].get()
        original_warehouse = master_data.get("warehouse", "")
        is_warehouse_changed = current_warehouse != original_warehouse

        if is_name_changed or is_warehouse_changed:
            row_dict["warning_label"].configure(text="*แก้ไข", text_color="orange")
        else:
            row_dict["warning_label"].configure(text="")
    
    def _get_commission_so_numbers(self):
        try:
            query = "SELECT DISTINCT so_number FROM commissions WHERE (status = 'Pending PU' OR (status = 'PO In Progress' AND user_key = %s)) AND is_active = 1 ORDER BY so_number;"
            df = pd.read_sql_query(query, self.pg_engine, params=(self.user_key,)); return [""] + df['so_number'].tolist()
        except Exception as e: print(f"Error fetching available SO numbers: {e}"); messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูล SO ได้: {e}", parent=self); return [""]
    
    def _load_supplier_data(self):
        try:
            df = pd.read_sql("SELECT id, supplier_name, supplier_code, credit_term FROM suppliers ORDER BY supplier_name", self.pg_engine)
            
            # --- START: แก้ไขส่วนนี้ ---
            # สร้างข้อมูลที่มีโครงสร้างสำหรับ AutoComplete ที่ฉลาดขึ้น
            self.supplier_completion_data = []
            for _, row in df.iterrows():
                display_text = f"{row['supplier_name']} (Code: {row.get('supplier_code', '')})"
                self.supplier_completion_data.append({
                    "id": row['id'],
                    "name": row['supplier_name'],
                    "code": row.get('supplier_code', ''),
                    "term": row.get('credit_term', 'เงินสด'),
                    "display": display_text
                })

            # สำหรับการอ้างอิงข้อมูลเมื่อผู้ใช้เลือก
            self.supplier_data_map = {item['display']: item for item in self.supplier_completion_data}
            
            # ส่งข้อมูลที่มีโครงสร้างใหม่นี้ไปยัง widget
            if hasattr(self, 'supplier_name_combo'):
                self.supplier_name_combo.completion_list = self.supplier_completion_data
            # --- END: สิ้นสุดการแก้ไข ---

        except Exception as e: 
            print(f"Error loading supplier data: {e}")
            self.supplier_completion_data = []
            self.supplier_data_map = {}
    
    def _load_product_master_data(self):
        try:
            df = pd.read_sql("SELECT product_code, product_name, warehouse FROM products ORDER BY product_code", self.pg_engine)
            
            # --- START: แก้ไขส่วนนี้ ---
            # สร้างข้อมูลที่มีโครงสร้างสำหรับ AutoComplete ที่ฉลาดขึ้น
            self.product_completion_data = []
            MAX_NAME_LENGTH = 50
            for _, row in df.iterrows():
                name = row['product_name'] or "" # จัดการกรณีชื่อเป็นค่าว่าง
                # ย่อชื่อที่ยาวเกินไปสำหรับแสดงผลใน Dropdown
                display_name = name[:MAX_NAME_LENGTH] + '...' if len(name) > MAX_NAME_LENGTH else name
                display_text = f"{row['product_code']} - {display_name}"

                self.product_completion_data.append({
                    "name": name, # เก็บชื่อเต็มไว้
                    "code": row['product_code'],
                    "warehouse": row.get('warehouse', ''),
                    "display": display_text
                })

            # สร้าง Map สำหรับดึงข้อมูลเมื่อผู้ใช้เลือก (เหมือนเดิมแต่ใช้ข้อมูลใหม่)
            self.product_data_map = {item['code']: item for item in self.product_completion_data}

            # อัปเดตรายการสินค้าในแถวที่ถูกสร้างไปแล้วให้ใช้ข้อมูลชุดใหม่
            for row_dict in self.product_rows:
                if "code" in row_dict and isinstance(row_dict["code"], AutoCompleteEntry):
                    row_dict["code"].completion_list = self.product_completion_data
            # --- END: สิ้นสุดการแก้ไข ---

        except Exception as e: 
            print(f"Error loading product master data: {e}")
            self.product_completion_data = []
            self.product_data_map = {}

    def _create_product_grid(self, parent):
        product_container = CTkFrame(parent, fg_color="#D6D7D8")
        product_container.pack(fill="x", expand=True, padx=10, pady=5)
        CTkLabel(product_container, text="รายการสินค้าและต้นทุน", font=self.header_font_table).pack(anchor="w", pady=5, padx=10)
        
        self.products_frame = CTkFrame(product_container, fg_color="transparent")
        self.products_frame.pack(fill="x", expand=True, padx=10, pady=5)
        
        headers = ["สถานะ", "รหัสสินค้า", "ชื่อสินค้า", "คลัง", "แก้ไข", "จำนวน", "ต้นทุนหน่วย (ไม่รวม VAT)", "ส่วนลด", "น้ำหนัก/หน่วย (กก.)", "น้ำหนักรวม (กก.)", "ต้นทุนรวม"]
        col_weights = [2, 4, 6, 2, 1, 2, 2, 3, 2, 2, 3]

        for i, h_text in enumerate(headers):
            self.products_frame.grid_columnconfigure(i, weight=col_weights[i])
            CTkLabel(self.products_frame, text=h_text, font=self.header_font_table, fg_color="#E0E0E0").grid(row=0, column=i, padx=1, pady=1, sticky="nsew")
        
        self.product_rows.clear()
        self._add_product_row()
        
        buttons_frame = CTkFrame(product_container, fg_color="transparent")
        buttons_frame.pack(anchor="e", pady=10, padx=10)
        CTkButton(buttons_frame, text="เพิ่มรายการสินค้า", command=self._add_product_row).pack(side="left", padx=5)
        self.delete_row_button = CTkButton(buttons_frame, text="ลบรายการล่าสุด", command=self._delete_last_product_row, fg_color="#D32F2F", hover_color="#B71C1C")
        self.delete_row_button.pack(side="left", padx=5)

    def _delete_last_product_row(self):
        if len(self.product_rows) > 1:
            last_row = self.product_rows.pop();
            for widget in last_row["widgets"]: widget.destroy()
            self._update_summary()
        else: messagebox.showwarning("ไม่สามารถลบได้", "ต้องมีรายการสินค้าอย่างน้อย 1 แถว", parent=self)
        
    def _add_product_row(self):
        row_num = len(self.product_rows) + 1
        
        product_name_var = tk.StringVar()
        warehouse_var = tk.StringVar()

        status_var = tk.StringVar(value="Stock")
        status_menu = CTkOptionMenu(self.products_frame, variable=status_var, values=["Stock", "Trade"], **self.dropdown_style)
        
        # --- START: แก้ไขส่วนนี้ ---
        # เปลี่ยนจากการใช้ self.product_display_list เป็น self.product_completion_data
        product_code_entry = AutoCompleteEntry(self.products_frame, 
                                               completion_list=self.product_completion_data, 
                                               placeholder_text="Code")
        # --- END: สิ้นสุดการแก้ไข ---

        product_name_entry = CTkEntry(self.products_frame, placeholder_text="Name", textvariable=product_name_var)
        warehouse_entry = CTkEntry(self.products_frame, placeholder_text="คลัง", textvariable=warehouse_var)

        warning_label = CTkLabel(self.products_frame, text="", width=10, font=CTkFont(size=12, slant="italic"), text_color="orange")
        qty_entry = NumericEntry(self.products_frame, placeholder_text="Qty")
        weight_entry = NumericEntry(self.products_frame, placeholder_text="kg/unit")
        price_entry = NumericEntry(self.products_frame, placeholder_text="price/unit")
        
        discount_frame = CTkFrame(self.products_frame, fg_color="transparent")
        discount_value_entry = NumericEntry(discount_frame)
        discount_value_entry.pack(side="left", fill="x", expand=True, padx=(0, 2))
        discount_type_var = tk.StringVar(value="บาท")
        discount_type_menu = CTkOptionMenu(discount_frame, variable=discount_type_var, values=["บาท", "%"], width=70, **self.dropdown_style)
        discount_type_menu.pack(side="left")

        total_weight_entry = CTkEntry(self.products_frame, state="readonly", fg_color="gray85")
        total_price_entry = CTkEntry(self.products_frame, state="readonly", fg_color="gray85")
        
        widgets = [
            status_menu, product_code_entry, product_name_entry,
            warehouse_entry, warning_label, qty_entry, price_entry, 
            discount_frame, weight_entry, total_weight_entry, total_price_entry
        ]

        for col, widget in enumerate(widgets):
            widget.grid(row=row_num, column=col, padx=1, pady=1, sticky="ew")
        
        row_dict = {
            "status_var": status_var, "code": product_code_entry, "name": product_name_entry,
            "warehouse": warehouse_entry, "qty": qty_entry, "weight": weight_entry,
            "price": price_entry, "discount_entry": discount_value_entry,
            "discount_type_var": discount_type_var, "total_weight": total_weight_entry,
            "total_price": total_price_entry, "widgets": widgets,
            "warning_label": warning_label, "master_data": None,
            "name_var": product_name_var,
            "warehouse_var": warehouse_var
        }
        
        self.product_rows.append(row_dict)
        product_code_entry.command_on_select = lambda selection, r=row_dict: self._on_product_selected(selection, r)
        
        product_name_var.trace_add("write", lambda *args, r=row_dict: self._check_for_override(r))
        warehouse_var.trace_add("write", lambda *args, r=row_dict: self._check_for_override(r))
        
        for entry in [qty_entry, weight_entry, price_entry, discount_value_entry]:
            entry.bind("<KeyRelease>", self._update_summary)
        discount_type_menu.configure(command=self._update_summary)

    def _create_bottom_summary_frame(self, parent):
        bottom_container = CTkFrame(parent, fg_color="transparent")
        bottom_container.pack(fill="x", expand=True, padx=10, pady=10)
        bottom_container.grid_columnconfigure((0, 1, 2), weight=1, uniform="group1")

        shipping_frame = CTkFrame(bottom_container, fg_color="#D6D7D8", border_width=1)
        shipping_frame.grid(row=0, column=0, padx=(0, 5), sticky="nsew")
        shipping_frame.grid_columnconfigure(1, weight=1)
        self._populate_shipping_column(shipping_frame)

        summary_frame = CTkFrame(bottom_container, fg_color="#D6D7D8", border_width=1)
        summary_frame.grid(row=0, column=1, padx=5, sticky="nsew")
        summary_frame.grid_columnconfigure(1, weight=1)
        self._populate_summary_column(summary_frame)

        payment_frame = CTkFrame(bottom_container, fg_color="#D6D7D8", border_width=1)
        payment_frame.grid(row=0, column=2, padx=(5, 0), sticky="nsew")
        payment_frame.grid_columnconfigure(1, weight=1)
        self._populate_payment_column(payment_frame)

    def _populate_shipping_column(self, parent_frame):
        CTkLabel(parent_frame, text="ค่าจัดส่ง", font=self.header_font_table).grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        # --- Section 1: Shipping to Stock ---
        CTkLabel(parent_frame, text="1.ค่าจัดส่งเข้าสต๊อก").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        
        stock_cost_frame = CTkFrame(parent_frame, fg_color="transparent")
        stock_cost_frame.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        
        self.shipping_to_stock_cost_entry = NumericEntry(stock_cost_frame)
        self.shipping_to_stock_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.shipping_to_stock_cost_entry.bind("<KeyRelease>", self._update_summary)

        stock_vat_radio_frame = CTkFrame(stock_cost_frame, fg_color="transparent")
        stock_vat_radio_frame.pack(side="left")
        CTkRadioButton(stock_vat_radio_frame, text="VAT", variable=self.shipping_to_stock_vat_var, value="VAT").pack(side="left")
        CTkRadioButton(stock_vat_radio_frame, text="CASH", variable=self.shipping_to_stock_vat_var, value="CASH").pack(side="left", padx=5)
        self.shipping_to_stock_vat_var.trace_add("write", self._update_summary)

        CTkLabel(parent_frame, text="VAT 7%:", font=self.entry_font).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_stock_vat_display_entry = CTkEntry(parent_frame, textvariable=self.shipping_to_stock_vat_display_var, state="readonly", fg_color="gray85")
        self.shipping_to_stock_vat_display_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        self.shipping_to_stock_date_selector = DateSelector(parent_frame, dropdown_style=self.dropdown_style)
        self.shipping_to_stock_date_selector.grid(row=3, column=1, sticky="w", padx=5, pady=2)
        
        # --- ตรวจสอบส่วนนี้ให้แน่ใจว่าตัวแปรถูกต้อง ---
        self.shipping_to_stock_type_var = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        stock_shipper_radio_frame = CTkFrame(parent_frame, fg_color="transparent")
        stock_shipper_radio_frame.grid(row=4, column=1, sticky="w", padx=5, pady=2)
        CTkRadioButton(stock_shipper_radio_frame, text="ซัพพลายเออร์จัดส่ง", variable=self.shipping_to_stock_type_var, value="ซัพพลายเออร์จัดส่ง", command=self._update_summary).pack(side="left")
        CTkRadioButton(stock_shipper_radio_frame, text="Aplus Logistic", variable=self.shipping_to_stock_type_var, value="Aplus Logistic", command=self._update_summary).pack(side="left", padx=5)
        CTkRadioButton(stock_shipper_radio_frame, text="Lalamove/Others", variable=self.shipping_to_stock_type_var, value="Lalamove/Others", command=self._update_summary).pack(side="left", padx=5)
        
        self.shipping_to_stock_notes_entry = CTkEntry(parent_frame, placeholder_text="หมายเหตุ...")
        self.shipping_to_stock_notes_entry.grid(row=5, column=1, sticky="ew", padx=5, pady=2)

        CTkFrame(parent_frame, height=2, fg_color="gray90").grid(row=6, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

        # --- Section 2: Shipping to Site ---
        CTkLabel(parent_frame, text="2.ค่าจัดส่งเข้าไซต์").grid(row=7, column=0, padx=10, pady=5, sticky="w")

        site_cost_frame = CTkFrame(parent_frame, fg_color="transparent")
        site_cost_frame.grid(row=7, column=1, sticky="ew", padx=5, pady=2)

        self.shipping_to_site_cost_entry = NumericEntry(site_cost_frame)
        self.shipping_to_site_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.shipping_to_site_cost_entry.bind("<KeyRelease>", self._update_summary)

        site_vat_radio_frame = CTkFrame(site_cost_frame, fg_color="transparent")
        site_vat_radio_frame.pack(side="left")
        CTkRadioButton(site_vat_radio_frame, text="VAT", variable=self.shipping_to_site_vat_var, value="VAT").pack(side="left")
        CTkRadioButton(site_vat_radio_frame, text="CASH", variable=self.shipping_to_site_vat_var, value="CASH").pack(side="left", padx=5)
        self.shipping_to_site_vat_var.trace_add("write", self._update_summary)

        CTkLabel(parent_frame, text="VAT 7%:", font=self.entry_font).grid(row=8, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_site_vat_display_entry = CTkEntry(parent_frame, textvariable=self.shipping_to_site_vat_display_var, state="readonly", fg_color="gray85")
        self.shipping_to_site_vat_display_entry.grid(row=8, column=1, sticky="ew", padx=5, pady=2)

        self.shipping_to_site_date_selector = DateSelector(parent_frame, dropdown_style=self.dropdown_style)
        self.shipping_to_site_date_selector.grid(row=9, column=1, sticky="w", padx=5, pady=2)
        
        # --- ตรวจสอบส่วนนี้ให้แน่ใจว่าตัวแปรถูกต้อง ---
        self.shipping_to_site_type_var = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        site_shipper_radio_frame = CTkFrame(parent_frame, fg_color="transparent")
        site_shipper_radio_frame.grid(row=10, column=1, sticky="w", padx=5, pady=2)
        CTkRadioButton(site_shipper_radio_frame, text="ซัพพลายเออร์จัดส่ง", variable=self.shipping_to_site_type_var, value="ซัพพลายเออร์จัดส่ง", command=self._update_summary).pack(side="left")
        CTkRadioButton(site_shipper_radio_frame, text="Aplus Logistic", variable=self.shipping_to_site_type_var, value="Aplus Logistic", command=self._update_summary).pack(side="left", padx=5)
        CTkRadioButton(site_shipper_radio_frame, text="Lalamove/Others", variable=self.shipping_to_site_type_var, value="Lalamove/Others", command=self._update_summary).pack(side="left", padx=5)

        self.shipping_to_site_notes_entry = CTkEntry(parent_frame, placeholder_text="หมายเหตุ...")
        self.shipping_to_site_notes_entry.grid(row=11, column=1, sticky="ew", padx=5, pady=2)
    
    def _populate_payment_column(self, parent_frame):
        parent_frame.grid_columnconfigure(1, weight=1)
        self.payment_entries.clear()

        CTkLabel(parent_frame, text="การชำระซัพพลายเออร์", font=self.header_font_table).grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        # --- Row Generation Helper ---
        def create_payment_row(label_text, row_index, has_percent_dropdown=False):
            CTkLabel(parent_frame, text=label_text).grid(row=row_index, column=0, padx=5, pady=5, sticky="w")
            
            p_frame = CTkFrame(parent_frame, fg_color="transparent")
            p_frame.grid(row=row_index, column=1, sticky="ew")
            
            percent_var = None
            if has_percent_dropdown:
                percent_var = tk.StringVar(value="ระบุยอดเอง")
                p_percent = CTkOptionMenu(p_frame, variable=percent_var, values=["ระบุยอดเอง", "30%", "50%", "100%"], width=120)
                p_percent.pack(side="left", padx=(0, 5))
            
            p_amount = NumericEntry(p_frame)
            p_amount.pack(side="left", fill="x", expand=True, padx=(0, 5))
            
            p_date = DateSelector(p_frame, dropdown_style=self.dropdown_style)
            p_date.pack(side="left")
            
            p_amount.bind("<KeyRelease>", self._update_summary)
            if has_percent_dropdown:
                p_percent.configure(command=lambda val, pv=percent_var, pa=p_amount: self._calculate_payment_from_percentage(val, pv, pa))

            return p_amount, p_date, percent_var

        # --- Create Rows ---
        p1_amount, p1_date, self.payment1_percent_var = create_payment_row("1.มัดจำ:", 1, has_percent_dropdown=True)
        p2_amount, p2_date, self.payment2_percent_var = create_payment_row("2.มัดจำ:", 2, has_percent_dropdown=True)

        CTkLabel(parent_frame, text="ยอดรวมมัดจำ:", font=self.label_font).grid(row=3, column=0, padx=5, pady=8, sticky="w")
        total_deposit_entry = CTkEntry(parent_frame, textvariable=self.total_deposit_var, state="readonly", fg_color="gray85")
        total_deposit_entry.grid(row=3, column=1, sticky="ew", pady=8, padx=5)

        CTkLabel(parent_frame, text="ยอดค้าง:", font=self.label_font).grid(row=4, column=0, padx=5, pady=8, sticky="w")
        
        self.balance_due_entry = CTkEntry(parent_frame, textvariable=self.balance_due_var, state="readonly", fg_color="gray85")
        self.balance_due_entry.grid(row=4, column=1, sticky="ew", pady=8, padx=5)
        

        fp_amount, fp_date, _ = create_payment_row("ชำระเต็ม:", 5)
        
        cn_amount, cn_date, _ = create_payment_row("CN/คืนส่วนลด:", 6)

        # --- Store Widgets in self.payment_entries ---
        self.payment_entries["Payment 1"] = {"amount": p1_amount, "date": p1_date, "percent_var": self.payment1_percent_var}
        self.payment_entries["Payment 2"] = {"amount": p2_amount, "date": p2_date, "percent_var": self.payment2_percent_var}
        self.payment_entries["Full Payment"] = {"amount": fp_amount, "date": fp_date}
        self.payment_entries["CN Refund"] = {"amount": cn_amount, "date": cn_date}
    
    def _populate_summary_column(self, parent_frame):
        CTkLabel(parent_frame, text="สรุปต้นทุน", font=self.header_font_table).grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        row_idx = 1

        def _create_summary_row(parent, label_text, row):
            CTkLabel(parent, text=label_text).grid(row=row, column=0, sticky="w", padx=10, pady=5)
            entry = CTkEntry(parent, state="readonly", fg_color="gray85")
            entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
            return entry

        self.total_weight_summary_entry = _create_summary_row(parent_frame, "น้ำหนักรวมทั้งหมด (กก.)", row_idx)
        self.total_cost_entry = _create_summary_row(parent_frame, "ยอดต้นทุนรวมของ PO (ไม่รวม VAT)", row_idx + 1)

        CTkLabel(parent_frame, text="ส่วนลดท้ายบิล:").grid(row=row_idx + 2, column=0, sticky="w", padx=10, pady=5)
        self.end_of_bill_discount_entry = NumericEntry(parent_frame)
        self.end_of_bill_discount_entry.grid(row=row_idx + 2, column=1, sticky="ew", padx=5, pady=5)
        self.end_of_bill_discount_entry.bind("<KeyRelease>", self._update_summary)

        self.vat3_checkbox = CTkCheckBox(parent_frame, text="หัก ณ ที่จ่าย 3%", command=self._update_summary)
        self.vat3_checkbox.grid(row=row_idx + 3, column=0, sticky="w", padx=10, pady=2)
        self.vat3_entry = CTkEntry(parent_frame, state="readonly", fg_color="gray85")
        self.vat3_entry.grid(row=row_idx + 3, column=1, sticky="ew", padx=5, pady=5)

        self.vat_checkbox = CTkCheckBox(parent_frame, text="Vat 7%", command=self._update_summary)
        self.vat_checkbox.grid(row=row_idx + 4, column=0, sticky="w", padx=10, pady=2)
        self.vat7_entry = CTkEntry(parent_frame, state="readonly", fg_color="gray85")
        self.vat7_entry.grid(row=row_idx + 4, column=1, sticky="ew", padx=5, pady=5)

        self.grand_total_with_vat_entry = _create_summary_row(parent_frame, "ยอดรวมใบแจ้งหนี้ (รวม VAT)", row_idx + 5)

        self.separate_shipping_entry = _create_summary_row(parent_frame, "ค่าจัดส่งต้นทุน - ชำระแยก", row_idx + 6)
        self.separate_shipping_entry.configure(text_color="#F97316", font=(self.entry_font.cget("family"), 12, "bold"))
        
        self.grand_total_payable_entry = CTkEntry(parent_frame, state="readonly", fg_color="#D1FAE5", text_color="#065F46", font=(self.entry_font.cget("family"), 16, "bold"), border_color="#10B981", border_width=2)
        CTkLabel(parent_frame, text="ยอดรวมที่ต้องชำระซัพพลายเออร์").grid(row=row_idx+6, column=0, sticky="w", padx=10, pady=5)
        self.grand_total_payable_entry.grid(row=row_idx+6, column=1, sticky="ew", padx=5, pady=5)

    def _create_footer_frame(self, parent):
        footer = CTkFrame(parent, fg_color="transparent")
        footer.pack(fill="x", expand=True, padx=10, pady=15)
        btn_config = {"corner_radius": 8, "font": (self.label_font.cget("family"), 12)}

        CTkButton(footer, text="📄 พิมพ์ใบสั่งซื้อ (PO)", command=self._open_so_selection_dialog, fg_color="#7C3AED", **btn_config).pack(side="left", padx=5, expand=True, fill="x")
        CTkButton(footer, text="💾 บันทึกฉบับร่าง (Save Draft)", command=lambda: self._save_po('Draft'), **btn_config).pack(side="left", padx=5, expand=True, fill="x")
        
        # เปลี่ยน command ของปุ่ม "ขออนุมัติ"
        CTkButton(footer, text="📤 ขออนุมัติ...", command=self._open_submit_po_dialog, fg_color="#16A34A", **btn_config).pack(side="left", padx=5, expand=True, fill="x")
    
    def _open_po_selection_dialog(self):
     try:
        POSelectionDialog(self, self.pg_engine, print_callback=self._print_selected_po)
     except Exception as e:
        messagebox.showerror("Error", f"Could not open PO selection window: {e}", parent=self)
        traceback.print_exc()

    def _open_submit_po_dialog(self):
        SubmitPODialog(self, self)

        

    def _print_selected_po(self, po_id):
        conn = self.app_container.get_connection()
        try:
            query = """
                SELECT
                    -- Fields from purchase_orders (po)
                    po.po_number,
                    po.rr_number,
                    po.department,
                    po.supplier_name,
                    po.credit_term,
                    po.po_mode,
                    po.wht_3_percent_amount AS wht_3_percent_po,
                    po.vat_7_percent_amount AS vat_7_percent_po,
                    po.grand_total AS grand_total_vat_po,
                    
                    -- Fields from commissions (c)
                    c.so_number,
                    c.bill_date,
                    c.commission_month,
                    c.commission_year,
                    c.customer_name,
                    c.credit_term,
                    c.sales_service_amount,
                    c.credit_card_fee,
                    c.cutting_drilling_fee,
                    c.transfer_fee,
                    c.wht_3_percent,
                    c.other_service_fee,
                    c.marketing_fee,
                    c.brokerage_fee,
                    c.shipping_cost,
                    c.coupons,
                    c.giveaways,
                    c.payment1_amount,
                    c.payment1_date,
                    c.payment1_method,
                    c.payment2_amount,
                    c.payment2_date,
                    c.payment2_method,
                    c.total_payment_amount,
                    c.balance_due,
                    c.cash_actual_payment,
                    c.cash_product_input,
                    c.cash_service_total,
                    c.cash_required_total,
                    c.delivery_type,
                    c.pickup_location,
                    c.relocation_cost,
                    c.date_to_warehouse,
                    c.date_to_customer,
                    c.pickup_registration,
                    c.sales_service_vat_option,
                    c.credit_card_fee_vat_option,
                    c.cutting_drilling_fee_vat_option,
                    c.other_service_fee_vat_option,
                    c.shipping_vat_option,

                    -- Fields from sales_users (u_po and u_so)
                    u_po.sale_name AS user_name,
                    u_so.sale_name AS sale_name
                    
                FROM purchase_orders po
                LEFT JOIN commissions c ON po.so_number = c.so_number
                LEFT JOIN sales_users u_po ON po.user_key = u_po.sale_key
                LEFT JOIN sales_users u_so ON c.sale_key = u_so.sale_key
                WHERE po.id = %s
                LIMIT 1;
            """
            po_df = pd.read_sql_query(query, self.pg_engine, params=(po_id,))

            if po_df.empty:
                messagebox.showerror("Error", "ไม่พบข้อมูล PO ที่เลือก", parent=self)
                return
            header_data = po_df.iloc[0].to_dict()

            items_df = pd.read_sql_query("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
            items_data = items_df.to_dict('records')

            payments_df = pd.read_sql_query("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
            payments_data = payments_df.to_dict('records')
            
            formatted_data = {
                "header": header_data,
                "items": items_data,
                "payments": payments_data
            }
            
            self.app_container.generate_single_po_document(po_id)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการดึงข้อมูลเพื่อพิมพ์: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    # purchasing_screen.py (ฟังก์ชัน _update_summary ที่แก้ไขแล้ว)

    # purchasing_screen.py (ฟังก์ชัน _update_summary ฉบับปรับปรุงใหม่)


        
    def _gather_form_data(self, *args):
        self._update_summary()
        
        header_data = {
            'so_number': self.so_entry.get(),
            'po_number': self.po_number_input_var.get(),
            'rr_number': self.rr_number_var.get(),
            'department': self.department_entry.get().strip(),
            'pur_order': self.pur_order_entry.get().strip(),
            'supplier_name': self.supplier_name_combo.get(),
            'supplier_code': self.supplier_code_entry.get(),
            'credit_term': self.credit_term_entry.get(),
            'po_mode': self.po_mode_var.get(), 
            'user_key': self.user_key,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'shipping_to_stock_cost': utils.convert_to_float(self.shipping_to_stock_cost_entry.get()),
            'shipping_to_site_cost': utils.convert_to_float(self.shipping_to_site_cost_entry.get()),
            'shipping_to_stock_cost': utils.convert_to_float(self.shipping_to_stock_cost_entry.get()),
            'shipping_to_stock_vat_type': self.shipping_to_stock_vat_var.get(),
            'shipping_to_stock_date': self.shipping_to_stock_date_selector.get_date(),
            # --- ตรวจสอบ 2 บรรทัดล่างนี้ให้แน่ใจว่าถูกต้อง ---
            'shipping_to_stock_shipper': self.shipping_to_stock_type_var.get(),
            'shipping_to_stock_notes': self.shipping_to_stock_notes_entry.get(),
            'shipping_to_site_cost': utils.convert_to_float(self.shipping_to_site_cost_entry.get()),
            'shipping_to_site_vat_type': self.shipping_to_site_vat_var.get(),
            'shipping_to_site_date': self.shipping_to_site_date_selector.get_date(),
            'shipping_to_site_shipper': self.shipping_to_site_type_var.get(),
            'shipping_to_site_notes': self.shipping_to_site_notes_entry.get(),
            # --- สิ้นสุดส่วนที่ต้องตรวจสอบ ---
            'total_cost': utils.convert_to_float(self.total_cost_entry.get()),
            'total_weight': utils.convert_to_float(self.total_weight_summary_entry.get()),
            'wht_3_percent_checked': bool(self.vat3_checkbox.get()),
            'wht_3_percent_amount': utils.convert_to_float(self.vat3_entry.get()),
            'vat_7_percent_checked': bool(self.vat_checkbox.get()),
            'vat_7_percent_amount': utils.convert_to_float(self.vat7_entry.get()),
            'grand_total': utils.convert_to_float(self.grand_total_payable_entry.get())
        }
        
        items_data = []
        for row in self.product_rows:
            if row["name"].get().strip():
                items_data.append({
                    "product_name": row["name"].get().strip(),
                    "status": row["status_var"].get(),
                    "product_code": row["code"].get().strip(),
                    "warehouse": row["warehouse"].get().strip(),
                    "quantity": utils.convert_to_float(row["qty"].get()),
                    "weight_per_unit": utils.convert_to_float(row["weight"].get()),
                    "unit_price": utils.convert_to_float(row["price"].get()),
                    "discount_value": utils.convert_to_float(row["discount_entry"].get()),
                    "discount_type": row["discount_type_var"].get(),
                    "total_weight": utils.convert_to_float(row["total_weight"].get()),
                    "total_price": utils.convert_to_float(row["total_price"].get())
                })
        
        payments_data = []
        for p_type, p_widgets in self.payment_entries.items():
            amount = utils.convert_to_float(p_widgets["amount"].get())
            if amount > 0:
                payments_data.append({
                    "payment_type": p_type,
                    "amount": amount,
                    "payment_date": p_widgets["date"].get_date()
                })

        return {"header": header_data, "items": items_data, "payments": payments_data}

    def _save_po(self, status):
        form_data = self._gather_form_data()
        header, items, payments = form_data.get('header', {}), form_data.get('items', []), form_data.get('payments', [])

        if status == 'Pending Approval' and (not header.get("so_number") or not header.get("supplier_name") or not items or not header.get("po_number")):
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก SO, PO/ST Number, Supplier, และเพิ่มสินค้าอย่างน้อย 1 รายการก่อนส่ง", parent=self)
            return
        
        conn = self.app_container.get_connection()
        try:
            if status == 'Pending Approval':
                header['status'] = 'Pending Approval'
                header['approval_status'] = 'Pending Mgr 1'
            else:
                header['status'] = 'Draft'
                header['approval_status'] = 'Draft'

            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'purchase_orders'")
                db_columns = {row[0] for row in cursor.fetchall()}
                
                if self.editing_po_id:
                    header.pop('user_key', None); header.pop('timestamp', None)
                    filtered_header = {k: v for k, v in header.items() if k in db_columns}
                    set_clause = ", ".join([f'"{k}"' for k in filtered_header.keys()])
                    params = list(filtered_header.values()) + [self.editing_po_id]
                    # แก้ไข: เปลี่ยน '"{k}" = %s' เป็น f'"{k}" = %s'
                    set_clause_formatted = ", ".join([f'"{k}" = %s' for k in filtered_header.keys()])
                    cursor.execute(f"UPDATE purchase_orders SET {set_clause_formatted} WHERE id = %s", params)
                    new_po_id = self.editing_po_id
                else:
                    filtered_header = {k: v for k, v in header.items() if k in db_columns}
                    cols = ", ".join([f'"{k}"' for k in filtered_header.keys()])
                    placeholders = ", ".join(["%s"] * len(filtered_header))
                    cursor.execute(f"INSERT INTO purchase_orders ({cols}) VALUES ({placeholders}) RETURNING id", list(filtered_header.values()))
                    new_po_id = cursor.fetchone()[0]
                
                cursor.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = %s", (new_po_id,))
                cursor.execute("DELETE FROM purchase_order_payments WHERE purchase_order_id = %s", (new_po_id,))
                
                # --- START: แก้ไขส่วนนี้โดยใช้ execute_values ---
                if items:
                    item_cols_str = ", ".join([f'"{k}"' for k in items[0].keys()])
                    insert_query_items = f"INSERT INTO purchase_order_items (purchase_order_id, {item_cols_str}) VALUES %s"
                    values_list_items = [ (new_po_id,) + tuple(item.values()) for item in items ]
                    psycopg2.extras.execute_values(cursor, insert_query_items, values_list_items)

                if payments:
                    payment_cols_str = ", ".join([f'"{k}"' for k in payments[0].keys()])
                    insert_query_payments = f"INSERT INTO purchase_order_payments (purchase_order_id, {payment_cols_str}) VALUES %s"
                    values_list_payments = [ (new_po_id,) + tuple(p.values()) for p in payments ]
                    psycopg2.extras.execute_values(cursor, insert_query_payments, values_list_payments)
                # --- END: สิ้นสุดการแก้ไข ---

                if status == 'Pending Approval':
                    self._create_initial_approval_notification(cursor, new_po_id)
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", f"บันทึก PO เป็น '{status}' สำเร็จ", parent=self)
            
            so_number_in_po = header.get("so_number")
            if so_number_in_po and status != 'Draft':
                # คำนวณค่าขนส่งที่ต้องเป็น payment_before_vat (เป็น CASH)
                payment_before_vat_cost = 0.0
                payment_before_vat_cost += header.get('shipping_to_stock_cost', 0)
                payment_before_vat_cost += header.get('shipping_to_site_cost', 0)

                # อัปเดตตาราง commissions ด้วยค่า payment_before_vat
                cursor.execute("""
                    UPDATE commissions
                    SET payment_before_vat = COALESCE(payment_before_vat, 0) + %s
                    WHERE so_number = %s AND is_active = TRUE
                """, (payment_before_vat_cost, so_number_in_po))

            if so_number_in_po and (status == 'Pending Approval' or status == 'Draft'):
                try:
                    with conn.cursor() as cursor:
                        new_comm_status = 'PO In Progress' if status == 'Draft' else 'PO Sent'
                        # <<< เพิ่ม user_key เข้าไปใน UPDATE เพื่อให้รู้ว่าใครกำลังทำ PO นี้ >>>
                        cursor.execute("UPDATE commissions SET status = %s, user_key = %s WHERE so_number = %s", (new_comm_status, self.user_key, so_number_in_po))
                        conn.commit()
                except Exception as e:
                    print(f"Error updating commissions status after PO save: {e}")

            if header.get('po_mode') == 'Multiple-PO':
                self._clear_form(keep_so=True)
            else:
                self._clear_form(keep_so=False)

        except psycopg2.Error as db_error:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"ไม่สามารถบันทึกได้: {db_error}", parent=self)
            traceback.print_exc()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_initial_approval_notification(self, cursor, po_id):
        try:
            cursor.execute("SELECT po_number, user_key FROM purchase_orders WHERE id = %s", (po_id,))
            po_info = cursor.fetchone()
            if not po_info: return
            po_number = po_info[0]

            cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Manager' AND status = 'Active'")
            manager_keys = [row[0] for row in cursor.fetchall()]

            message = f"PO ใหม่ ({po_number}) รอการอนุมัติจากผู้จัดการ"

            for manager_key in manager_keys:
                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, related_po_id, is_read) VALUES (%s, %s, %s, FALSE)",
                    (manager_key, message, po_id)
                )
        except Exception as e:
            print(f"Error creating initial PO approval notification: {e}")
            traceback.print_exc()

    def _load_po_to_edit(self, po_id):
        conn = self.app_container.get_connection()
        try:
            self.handle_clear_button_press(confirm=False)
            self.editing_po_id = po_id
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (po_id,))
                po_data = cursor.fetchone()
                if not po_data:
                    messagebox.showerror("Error", "ไม่พบ PO ที่ต้องการแก้ไข", parent=self)
                    return
                cursor.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", (po_id,))
                items_data = cursor.fetchall()
                cursor.execute("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", (po_id,))
                payments_data = cursor.fetchall()

            so_num = po_data.get("so_number", "")
            self.so_entry.set(so_num)
            self._on_so_selected(so_num, is_editing=True)
            
            # --- START: แก้ไข Logic การเติมข้อมูล Supplier ---
            supplier_name_from_db = str(po_data.get("supplier_name") or "")
            
            # 1. ค้นหาข้อมูล Supplier ทั้งหมด (Dictionary) จากชื่อที่บันทึกไว้
            supplier_dict_to_pass = None
            if hasattr(self, 'supplier_completion_data'):
                for item_dict in self.supplier_completion_data:
                    if item_dict.get('name') == supplier_name_from_db:
                        supplier_dict_to_pass = item_dict
                        break
            
            # 2. เรียกใช้ฟังก์ชัน _on_supplier_selected ด้วยข้อมูลที่ถูกต้อง
            if supplier_dict_to_pass:
                # เติมข้อความในช่องค้นหาก่อน
                self.supplier_name_combo.delete(0, 'end')
                self.supplier_name_combo.insert(0, supplier_dict_to_pass.get('name', ''))
                # จากนั้นเรียกฟังก์ชันเพื่อเติมข้อมูลส่วนที่เหลือ
                self._on_supplier_selected(supplier_dict_to_pass)
            else:
                # กรณีหาไม่เจอ (เช่น ซัพพลายเออร์ถูกลบไปแล้ว) ให้แสดงแค่ชื่อ
                self.supplier_name_combo.delete(0, 'end')
                self.supplier_name_combo.insert(0, supplier_name_from_db)
            # --- END: สิ้นสุดการแก้ไข ---
            
            # ... (โค้ดส่วนที่เหลือของฟังก์ชันเหมือนเดิมทั้งหมด) ...
            self.shipping_to_stock_vat_var.set(po_data.get("shipping_to_stock_vat_type", "VAT"))
            self.shipping_to_site_vat_var.set(po_data.get("shipping_to_site_vat_type", "VAT"))
            
            if po_data.get("vat_7_percent_checked"): self.vat_checkbox.select()
            po_full_number = po_data.get("po_number", "PO")
            self.po_number_type_var.set("PO")
            if po_full_number.startswith("ST"):
                self.po_number_type_var.set("ST")
            self.po_number_input_var.set(po_full_number)
            self.rr_number_var.set(po_data.get("rr_number", ""))
            
            self.department_entry.delete(0, 'end')
            self.department_entry.insert(0, po_data.get("department", ""))
            self.pur_order_entry.delete(0, 'end')
            self.pur_order_entry.insert(0, po_data.get("pur_order", ""))
            self.po_mode_var.set(po_data.get("po_mode", "Single-PO")) 

            self.shipping_to_stock_cost_entry.delete(0, 'end')
            self.shipping_to_stock_cost_entry.insert(0, f"{po_data.get('shipping_to_stock_cost', 0):.2f}")
            self.shipping_to_site_cost_entry.delete(0, 'end')
            self.shipping_to_site_cost_entry.insert(0, f"{po_data.get('shipping_to_site_cost', 0):.2f}")
            
            for row in self.product_rows:
                for widget in row["widgets"]:
                    widget.destroy()
            self.product_rows.clear()
            
            if not items_data:
                self._add_product_row()
            else:
                for item in items_data:
                    self._add_product_row()
                    last_row = self.product_rows[-1]
                    last_row["name_var"].set(str(item.get("product_name") or ""))
                    last_row["status_var"].set(str(item.get("status") or "Stock"))
                    last_row["code"].insert(0, str(item.get("product_code") or ""))
                    last_row["warehouse_var"].set(str(item.get("warehouse") or ""))
                    last_row["qty"].insert(0, f"{(item.get('quantity') or 0):.2f}")
                    last_row["weight"].insert(0, f"{(item.get('weight_per_unit') or 0):.2f}")
                    last_row["price"].insert(0, f"{(item.get('unit_price') or 0):.2f}")
                    last_row["discount_entry"].insert(0, f"{(item.get('discount_value') or 0):.2f}")
                    last_row["discount_type_var"].set(str(item.get("discount_type") or "บาท"))
            
            if po_data.get("vat_7_percent_checked"): self.vat_checkbox.select()
            else: self.vat_checkbox.deselect()
            
            if po_data.get("wht_3_percent_checked"): self.vat3_checkbox.select()
            else: self.vat3_checkbox.deselect()

            for p_data in payments_data:
                p_type = p_data.get('payment_type')
                if p_type in self.payment_entries:
                    p_widgets = self.payment_entries[p_type]
                    p_widgets["amount"].insert(0, f"{p_data.get('amount', 0):,.2f}")
                    if p_widgets.get("date"):
                        p_widgets["date"].set_date(p_data.get("payment_date"))
            
            self._update_summary()
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการโหลดข้อมูล: {e}\n{traceback.format_exc()}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _calculate_payment_from_percentage(self, selected_value, percent_var, amount_entry):
        try:
            if selected_value == "ระบุยอดเอง":
                return

            grand_total = utils.convert_to_float(self.grand_total_payable_entry.get())
            if grand_total <= 0:
                amount_entry.delete(0, tk.END)
                self._update_summary()
                return

            percent = float(selected_value.replace('%', '')) / 100.0
            calculated_amount = grand_total * percent
            
            amount_entry.delete(0, tk.END)
            amount_entry.insert(0, f"{calculated_amount:,.2f}")
            self._update_summary()

        except (ValueError, TypeError) as e:
            print(f"Error calculating payment from percentage: {e}")
            self._update_summary()

    