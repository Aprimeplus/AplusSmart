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
import time 
import re
from history_windows import SOPopupWindow
from hr_windows import SOPopupWindow
from hr_windows import SODetailViewer 
from export_utils import export_approved_pos_to_excel
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
from simple_async import SimpleAsyncHelper, show_loading_message, hide_loading_message
from purchasing_windows import SOFinderDialog
from daily_report_widget import DailyReportWidget
from cost_benchmark import CostBenchmarkScreen
from dashboard_cost import DashboardCostScreen




# --- แก้ไข: ลบ import ที่เป็นปัญหาออก และย้าย Dialog class ไปไว้ในไฟล์ของตัวเอง (ถ้ามี) ---
# from history_windows import SOPopupWindow 
# from export_utils import export_approved_pos_to_excel
from pdf_utils import export_approved_pos_to_pdf
from po_selection_dialog import POSelectionDialog

# --- แก้ไข: import Dialog ที่ถูกต้อง ---
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
            
        print(f"\n--- DEBUG: Starting _confirm_submission for {len(selected_records)} POs ---") # DEBUG

        selected_ids = [po_id for po_id, _ in selected_records]
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                
                # 1. อัปเดตสถานะ PO
                ids_tuple = tuple(selected_ids)
                update_query = """
                    UPDATE purchase_orders 
                    SET status = 'Pending Approval', approval_status = 'Pending Mgr 1' 
                    WHERE id IN %s
                """
                cursor.execute(update_query, (ids_tuple,))
                print(f"DEBUG: Updated status for IDs: {selected_ids}") # DEBUG

                # 2. คำนวณยอดค่าขนส่งใหม่ (Sync Logic)
                affected_so_numbers = list(set(rec['so_number'] for _, rec in selected_records))
                for so_number in affected_so_numbers:
                    # ... (Logic เดิม) ...
                    cursor.execute("""
                        SELECT SUM(COALESCE(shipping_to_stock_cost, 0) + COALESCE(shipping_to_site_cost, 0))
                        FROM purchase_orders
                        WHERE so_number = %s AND status IN ('Pending Approval', 'Approved')
                    """, (so_number,))
                    new_total_shipping_cost = cursor.fetchone()[0] or 0.0
                    
                    cursor.execute("""
                        UPDATE commissions
                        SET payment_before_vat = %s
                        WHERE so_number = %s AND is_active = 1
                    """, (new_total_shipping_cost, so_number))
                    print(f"DEBUG: Synced shipping cost for SO {so_number}: {new_total_shipping_cost}") # DEBUG

                # 3. สร้าง Notification
                cursor.execute("SELECT sale_key, role FROM sales_users WHERE role IN ('Purchasing Manager', 'Manager', 'Director') AND status = 'Active'")
                managers = cursor.fetchall()
                print(f"DEBUG: Found managers for notification: {managers}") # DEBUG

                if not managers:
                    print("⚠️ WARNING: No Manager found! Notifications skipped.")
                
                notif_data = []
                for po_id, record_data in selected_records:
                    message = f"PO ใหม่ ({record_data['po_number']}) รอการอนุมัติจากผู้จัดการ"
                    for sale_key, role in managers:
                        notif_data.append((sale_key, message, False, po_id))
                
                if notif_data:
                    psycopg2.extras.execute_values(
                        cursor,
                        "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES %s",
                        notif_data
                    )
                    print(f"DEBUG: Inserted {len(notif_data)} notifications") # DEBUG
            
            conn.commit()
            print("DEBUG: Commit successful") # DEBUG
            
            messagebox.showinfo("สำเร็จ", f"ส่ง PO จำนวน {len(selected_ids)} รายการเพื่อขออนุมัติเรียบร้อยแล้ว", parent=self.purchasing_screen)
            
            self.purchasing_screen._update_tasks_badge()
            self.destroy()

        except Exception as e:
            print(f"❌ ERROR in _confirm_submission: {e}") # DEBUG
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
        
        self.new_so_current_page = 0
        self.new_so_rows_per_page = 15
        self.new_so_search_term = ""

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
        CTkButton(header, text="Refresh All", command=self.load_tasks, width=100).pack(side="right")
        
        self.task_tab_view = CTkTabview(parent, corner_radius=10)
        self.task_tab_view.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        self.new_so_tab = self.task_tab_view.add("SO ใหม่รอสร้าง PO")
        self.in_progress_tab = self.task_tab_view.add("งานที่กำลังดำเนินการ (SO/PO Drafts)")
        self.rejected_tab = self.task_tab_view.add("งานที่ถูกปฏิเสธ (Rejected)")

        self.new_so_tab.grid_columnconfigure(0, weight=1)
        self.new_so_tab.grid_rowconfigure(2, weight=1)

        # 1. Search Frame
        search_frame = CTkFrame(self.new_so_tab)
        search_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        search_frame.grid_columnconfigure(0, weight=1)
        self.new_so_search_entry = CTkEntry(search_frame, placeholder_text="ค้นหา SO หรือ ชื่อลูกค้า...")
        self.new_so_search_entry.grid(row=0, column=0, sticky="ew", padx=(10,5), pady=10)
        CTkButton(search_frame, text="ค้นหา", command=self._search_new_so_tasks, width=80).grid(row=0, column=1, padx=5, pady=10)
        CTkButton(search_frame, text="ล้างค่า", command=self._clear_search_and_refresh, fg_color="gray", width=80).grid(row=0, column=2, padx=5, pady=10)

        # 2. Pagination Frame
        pagination_frame = CTkFrame(self.new_so_tab, fg_color="transparent")
        pagination_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=0)
        self.new_so_prev_button = CTkButton(pagination_frame, text="<< หน้าก่อนหน้า", command=self._new_so_prev_page, state="disabled")
        self.new_so_prev_button.pack(side="left")
        self.new_so_page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.new_so_page_label.pack(side="left", expand=True)
        self.new_so_next_button = CTkButton(pagination_frame, text="หน้าถัดไป >>", command=self._new_so_next_page, state="disabled")
        self.new_so_next_button.pack(side="right")

        # 3. Scroll Frame
        self.new_so_scroll_frame = CTkScrollableFrame(self.new_so_tab)
        self.new_so_scroll_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        
        # --- Layout เดิมสำหรับแท็บอื่นๆ ---
        self.in_progress_scroll_frame = CTkScrollableFrame(self.in_progress_tab)
        self.in_progress_scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.in_progress_scroll_frame.grid_columnconfigure(0, weight=1)
        so_zone = CTkFrame(self.in_progress_scroll_frame, fg_color="#F0F9FF", corner_radius=10); so_zone.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5)); so_zone.grid_columnconfigure(0, weight=1)
        CTkLabel(so_zone, text="SO ที่กำลังดำเนินการ", font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=15, pady=(10, 0))
        CTkLabel(so_zone, text="(SO ที่คุณ Claim มาแล้ว แต่ยังไม่ได้สร้าง PO)", font=CTkFont(size=12, slant="italic"), text_color="gray50").pack(anchor="w", padx=15, pady=(0, 10))
        self.so_in_progress_content_frame = CTkFrame(so_zone, fg_color="transparent"); self.so_in_progress_content_frame.pack(fill="x", expand=True, padx=5, pady=(0, 5))
        po_zone = CTkFrame(self.in_progress_scroll_frame, fg_color="#F1F5F9", corner_radius=10); po_zone.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10)); po_zone.grid_columnconfigure(0, weight=1)
        CTkLabel(po_zone, text="PO ฉบับร่าง", font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=15, pady=(10, 0))
        CTkLabel(po_zone, text="(PO ที่คุณสร้างและบันทึกร่างไว้ แต่ยังไม่ได้ส่งอนุมัติ)", font=CTkFont(size=12, slant="italic"), text_color="gray50").pack(anchor="w", padx=15, pady=(0, 10))
        self.po_draft_content_frame = CTkFrame(po_zone, fg_color="transparent"); self.po_draft_content_frame.pack(fill="x", expand=True, padx=5, pady=(0, 5))
        self.rejected_scroll_frame = CTkScrollableFrame(self.rejected_tab, label_text="รายการที่ต้องแก้ไข"); self.rejected_scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    def on_close(self):
        self.purchasing_screen._update_tasks_badge()
        self.purchasing_screen.tasks_window = None
        self.destroy()

    def load_tasks(self):
        self._load_new_so_tasks()
        self._load_in_progress_tasks()
        self._load_rejected_po_tasks()

    def _search_new_so_tasks(self):
        self.new_so_search_term = self.new_so_search_entry.get().strip()
        self.new_so_current_page = 0
        self._load_new_so_tasks()

    def _new_so_prev_page(self):
        if self.new_so_current_page > 0:
            self.new_so_current_page -= 1
            self._load_new_so_tasks()
        
    def _new_so_next_page(self):
        self.new_so_current_page += 1
        self._load_new_so_tasks()

    def _clear_search_and_refresh(self):
        self.new_so_search_entry.delete(0, 'end')
        self.new_so_search_term = ""
        self.new_so_current_page = 0
        self._load_new_so_tasks()

    def _load_new_so_tasks(self):
        frame = self.new_so_scroll_frame
        for widget in frame.winfo_children(): widget.destroy()

        try:
            base_query = "FROM commissions c JOIN sales_users u ON c.sale_key = u.sale_key WHERE c.status = 'Pending PU' AND c.is_active = 1"
            params = []
            
            if self.new_so_search_term:
                base_query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                search_like = f"%{self.new_so_search_term}%"
                params.extend([search_like, search_like])

            count_query = f"SELECT COUNT(c.id) {base_query}"
            total_rows = pd.read_sql_query(count_query, self.app_container.pg_engine, params=tuple(params)).iloc[0,0]
            total_pages = (total_rows + self.new_so_rows_per_page - 1) // self.new_so_rows_per_page

            offset = self.new_so_current_page * self.new_so_rows_per_page
            data_query = f"SELECT c.id, c.so_number, c.timestamp, c.customer_name, u.sale_name {base_query} ORDER BY c.timestamp DESC LIMIT %s OFFSET %s"
            final_params = params + [self.new_so_rows_per_page, offset]
            
            df = pd.read_sql_query(data_query, self.app_container.pg_engine, params=tuple(final_params))

            self.new_so_page_label.configure(text=f"หน้า {self.new_so_current_page + 1} / {max(1, total_pages)}")
            self.new_so_prev_button.configure(state="normal" if self.new_so_current_page > 0 else "disabled")
            self.new_so_next_button.configure(state="normal" if self.new_so_current_page < total_pages - 1 else "disabled")

            if df.empty:
                message = "ไม่พบ SO ใหม่" if not self.new_so_search_term else f"ไม่พบผลลัพธ์สำหรับ '{self.new_so_search_term}'"
                CTkLabel(frame, text=message).pack(pady=20)
                return
            
            for _, row in df.iterrows():
                card = CTkFrame(frame, border_width=1, fg_color="#F0FDF4")
                card.pack(fill="x", padx=5, pady=3)
                card.grid_columnconfigure(0, weight=1)

                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.grid(row=0, column=0, sticky="w", padx=10, pady=5)

                ts = pd.to_datetime(row['timestamp']).strftime("%Y-%m-%d %H:%M") if pd.notna(row['timestamp']) else "N/A"
                info_text = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']} (ส่งโดย: {row['sale_name']})"

                CTkLabel(info_frame, text=info_text, font=self.label_font, justify="left").pack(anchor="w")
                CTkLabel(info_frame, text=f"เวลาที่ส่ง: {ts}", font=CTkFont(size=11), text_color="gray").pack(anchor="w")

                start_button = CTkButton(card, text="เริ่มสร้าง PO", command=lambda s=row['so_number']: self._select_so_and_close(s))
                start_button.grid(row=0, column=1, sticky="e", padx=10, pady=5)
        
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดรายการ SO ใหม่ได้: {e}", parent=self)
            traceback.print_exc()
 
    def _return_so_from_task(self, so_number):
        """ส่ง SO กลับไปที่คิว 'Pending PU' จากหน้า My Tasks"""
        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการส่ง SO: {so_number} กลับไปที่คิวงานใช่หรือไม่?", parent=self):
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # อัปเดตสถานะ, ลบ user ที่ claim และลบเวลาที่ claim ออก
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'Pending PU', user_key = NULL, claim_timestamp = NULL 
                    WHERE so_number = %s AND user_key = %s AND status = 'PO In Progress'
                """, (so_number, self.user_key))
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", f"SO: {so_number} ถูกส่งกลับไปที่คิวงานเรียบร้อยแล้ว", parent=self)
            
            # โหลดข้อมูลในหน้า Tasks ใหม่ทั้งหมด
            self.load_tasks()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _load_in_progress_tasks(self):
        for widget in self.so_in_progress_content_frame.winfo_children(): widget.destroy()
        for widget in self.po_draft_content_frame.winfo_children(): widget.destroy()
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                so_query = """
                    SELECT c.id, c.so_number, c.timestamp, c.customer_name 
                    FROM commissions c
                    WHERE c.status = 'PO In Progress' AND c.user_key = %s
                    AND NOT EXISTS (
                        SELECT 1 FROM purchase_orders po WHERE po.so_number = c.so_number
                    )
                    ORDER BY c.timestamp DESC
                """
                cursor.execute(so_query, (self.user_key,))
                claimed_sos = cursor.fetchall()

                po_query = """
                    SELECT 
                        po.id, po.timestamp, po.so_number, po.po_number, po.supplier_name,
                        owner.sale_name as owner_name,
                        proxy.sale_name as proxy_name
                    FROM purchase_orders po
                    LEFT JOIN sales_users owner ON po.user_key = owner.sale_key
                    LEFT JOIN sales_users proxy ON po.proxy_user_key = proxy.sale_key
                    WHERE po.user_key = %s AND po.status = 'Draft' 
                    ORDER BY po.timestamp DESC
                """
                cursor.execute(po_query, (self.user_key,))
                draft_pos = cursor.fetchall()

            if not claimed_sos:
                CTkLabel(self.so_in_progress_content_frame, text="ไม่มี SO ที่รอสร้าง PO ใบแรก").pack(pady=10)
            else:
                for so_data in claimed_sos:
                    card = CTkFrame(self.so_in_progress_content_frame, border_width=1)
                    card.pack(fill="x", padx=5, pady=3)
                    card.grid_columnconfigure(0, weight=1)

                    info = f"SO: {so_data['so_number']} - ลูกค้า: {so_data['customer_name']} (ดำเนินการโดย: คุณ)"
                    CTkLabel(card, text=info, font=self.label_font).grid(row=0, column=0, sticky="w", padx=10, pady=5)

                    action_frame = CTkFrame(card, fg_color="transparent")
                    action_frame.grid(row=0, column=1, sticky="e", padx=10, pady=5)

                    return_button = CTkButton(
                        action_frame, 
                        text="คืน SO", 
                        command=lambda s=so_data['so_number']: self._return_so_from_task(s),
                        fg_color="#F97316", # สีเหลือง/ส้ม
                        hover_color="#EA580C",
                        width=80
                    )
                    return_button.pack(side="left", padx=(0, 5))

                    continue_button = CTkButton(
                        action_frame, 
                        text="ทำต่อ", 
                        command=lambda s=so_data['so_number']: self._continue_so_task(s)
                    )
                    continue_button.pack(side="left")

            if not draft_pos:
                CTkLabel(self.po_draft_content_frame, text="ไม่มี PO ฉบับร่าง").pack(pady=10)
            else:
                for po_data in draft_pos:
                    po_id = po_data['id']
                    card = CTkFrame(self.po_draft_content_frame, border_width=1); card.pack(fill="x", padx=5, pady=3)
                    card.grid_columnconfigure(0, weight=1); card.grid_columnconfigure(1, weight=0)
                    info_frame = CTkFrame(card, fg_color="transparent"); info_frame.grid(row=0, column=0, sticky="w", padx=10, pady=5)
                    action_frame = CTkFrame(card, fg_color="transparent"); action_frame.grid(row=0, column=1, sticky="e", padx=10, pady=5)
                    
                    info = f"SO: {po_data['so_number']} | PO: {po_data['po_number']} | Supplier: {po_data['supplier_name']}"
                    CTkLabel(info_frame, text=info).pack(anchor="w")
                    
                    owner_name = po_data.get('owner_name', 'N/A')
                    proxy_name = po_data.get('proxy_name')

                    if pd.notna(proxy_name) and proxy_name != owner_name:
                        owner_text = f"Owner: {owner_name} (สร้างโดย: {proxy_name})"
                        text_color = "#6D28D9" # สีม่วง
                    else:
                        owner_text = f"Owner: {owner_name}"
                        text_color = "gray30"
                    
                    CTkLabel(info_frame, text=owner_text, font=CTkFont(size=12, slant="italic"), text_color=text_color).pack(anchor="w")
                    
                    CTkButton(action_frame, text="แก้ไข", width=60, command=lambda p=po_id: self._edit_and_close(p)).pack(side="left", padx=2)
                    
                    # [🔥 แก้ไข] ปุ่มนี้เคยเรียก _submit_draft ที่ไม่มี Notification ตอนนี้แก้ฟังก์ชันนั้นแล้ว
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
        
        print(f"\n--- DEBUG: Starting _submit_draft for PO ID: {po_id} ---") # DEBUG
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # 1. Update Status
                cursor.execute("UPDATE purchase_orders SET status = 'Pending Approval', approval_status = 'Pending Mgr 1', timestamp = %s WHERE id = %s RETURNING po_number", (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), po_id))
                po_num_res = cursor.fetchone()
                po_number = po_num_res[0] if po_num_res else "N/A"
                print(f"DEBUG: Status updated for PO: {po_number}") # DEBUG

                # 2. ค้นหา Manager
                cursor.execute("SELECT sale_key, role FROM sales_users WHERE role IN ('Purchasing Manager', 'Manager', 'Director') AND status = 'Active'")
                managers = cursor.fetchall()
                print(f"DEBUG: Found {len(managers)} managers: {managers}") # DEBUG

                if not managers:
                    print("⚠️ WARNING: No Manager found in DB! Notification will NOT be sent.")
                    messagebox.showwarning("แจ้งเตือน", "ไม่พบรายชื่อผู้จัดการ (Manager) ในระบบ\nสถานะ PO เปลี่ยนแล้ว แต่จะไม่มีการแจ้งเตือน")

                # 3. ส่ง Notification
                for sale_key, role in managers:
                     msg = f"PO ใหม่ ({po_number}) รอการอนุมัติจากผู้จัดการ"
                     cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, related_po_id, is_read) VALUES (%s, %s, %s, FALSE)",
                        (sale_key, msg, po_id)
                    )
                     print(f"DEBUG: Notification sent to {sale_key} ({role})") # DEBUG

            conn.commit()
            print("DEBUG: Commit successful") # DEBUG
            
            self.load_tasks()
            messagebox.showinfo("สำเร็จ", f"ส่ง PO: {po_number} เรียบร้อยแล้ว", parent=self)

        except Exception as e:
            print(f"❌ ERROR in _submit_draft: {e}") # DEBUG
            if conn: conn.rollback(); 
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
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

# ค้นหา class ProductManagementWindow และวางทับด้วยโค้ดนี้

class ProductManagementWindow(CTkToplevel):
    def __init__(self, master, purchasing_screen_instance):
        super().__init__(master)
        self.purchasing_screen = purchasing_screen_instance
        self.app_container = purchasing_screen_instance.app_container
        self.user_key = purchasing_screen_instance.user_key
        self.label_font = purchasing_screen_instance.label_font
        self.entry_font = purchasing_screen_instance.entry_font

        self.title("จัดการข้อมูลสินค้าหลัก (Product Management)")
        self.geometry("1000x700")
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
        
        CTkButton(button_frame, text="เพิ่มสินค้าใหม่", width=100, command=self._add_product).pack(side="left", padx=5)
        CTkButton(button_frame, text="แก้ไข", width=80, command=self._edit_product).pack(side="left", padx=5)
        CTkButton(button_frame, text="ลบ", width=60, command=self._delete_product, fg_color="#D32F2F", hover_color="#B71C1C").pack(side="left", padx=5)
        
        CTkLabel(button_frame, text="|", font=CTkFont(size=20), text_color="gray").pack(side="left", padx=5)
        
        CTkButton(button_frame, text="📥 Export Excel", width=100, command=self._export_products, 
                  fg_color="#107C41", hover_color="#0B532B").pack(side="left", padx=5)
                  
        CTkButton(button_frame, text="📤 Import Excel", width=100, command=self._import_products, 
                  fg_color="#D97706", hover_color="#B45309").pack(side="left", padx=5)

        CTkButton(button_frame, text="🔄", width=40, command=self.load_products).pack(side="left", padx=5)

        self.search_entry = CTkEntry(self, placeholder_text="ค้นหาสินค้า (รหัส/ชื่อ)", font=self.entry_font)
        self.search_entry.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        self.search_entry.bind("<KeyRelease>", self._filter_products)

        self.tree_frame = CTkFrame(self, fg_color="transparent")
        self.tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("id", "product_code", "product_name", "warehouse", "price", "weight")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("id", text="ID", anchor="center")
        self.tree.heading("product_code", text="รหัสสินค้า", anchor="center")
        self.tree.heading("product_name", text="ชื่อสินค้า", anchor="center")
        self.tree.heading("warehouse", text="คลัง", anchor="center")
        self.tree.heading("price", text="ราคาล่าสุด", anchor="e")
        self.tree.heading("weight", text="นน.ล่าสุด", anchor="e")

        self.tree.column("id", width=50, anchor="center")
        self.tree.column("product_code", width=150, anchor="w")
        self.tree.column("product_name", width=350, anchor="w")
        self.tree.column("warehouse", width=100, anchor="center")
        self.tree.column("price", width=100, anchor="e")
        self.tree.column("weight", width=100, anchor="e")

        self.tree.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#F5F5F5", foreground="black", rowheight=25, fieldbackground="#F5F5F5")
        style.map('Treeview', background=[('selected', '#3B82F6')])
        style.configure("Treeview.Heading", font=CTkFont(size=12, weight="bold"), background="#E0E0E0", foreground="black")

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)

    def load_products(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = self.app_container.get_connection()
        try:
            cursor_query = "SELECT id, product_code, product_name, warehouse, last_unit_price, last_weight_per_unit FROM products ORDER BY product_code"
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute(cursor_query)
                products = cursor.fetchall()
                for prod in products:
                    price = f"{prod['last_unit_price']:,.2f}" if prod['last_unit_price'] else "-"
                    weight = f"{prod['last_weight_per_unit']:,.2f}" if prod['last_weight_per_unit'] else "-"
                    
                    self.tree.insert("", "end", values=(
                        prod['id'],
                        prod['product_code'],
                        prod['product_name'],
                        prod['warehouse'],
                        price,
                        weight
                    ))
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดข้อมูลสินค้าได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _filter_products(self, event):
        search_term = self.search_entry.get().strip().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = self.app_container.get_connection()
        try:
            query = """
                SELECT id, product_code, product_name, warehouse, last_unit_price, last_weight_per_unit 
                FROM products 
                WHERE LOWER(product_code) LIKE %s OR LOWER(product_name) LIKE %s 
                ORDER BY product_code
            """
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute(query, (f"%{search_term}%", f"%{search_term}%"))
                products = cursor.fetchall()
                for prod in products:
                    price = f"{prod['last_unit_price']:,.2f}" if prod['last_unit_price'] else "-"
                    weight = f"{prod['last_weight_per_unit']:,.2f}" if prod['last_weight_per_unit'] else "-"
                    
                    self.tree.insert("", "end", values=(
                        prod['id'],
                        prod['product_code'],
                        prod['product_name'],
                        prod['warehouse'],
                        price,
                        weight
                    ))
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถค้นหาข้อมูลสินค้าได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    # --- ฟังก์ชัน Export ที่แก้ไขแล้ว (เพิ่ม ID) ---
    def _export_products(self):
        """Export ข้อมูลสินค้าทั้งหมดเป็น Excel"""
        try:
            conn = self.app_container.get_connection()
            # [แก้ไข] เพิ่ม id ใน Query
            query = "SELECT id, product_code, product_name, warehouse, last_unit_price, last_weight_per_unit FROM products ORDER BY product_code"
            df = pd.read_sql_query(query, conn)
            
            if df.empty:
                messagebox.showinfo("ไม่มีข้อมูล", "ไม่พบข้อมูลสินค้าที่จะ Export", parent=self)
                return

            # [แก้ไข] เพิ่ม System_ID ในไฟล์ Excel
            df.rename(columns={
                'id': 'System_ID', # คอลัมน์สำคัญ! ห้ามเปลี่ยนชื่อในไฟล์ Excel ถ้าจะใช้อัปเดต
                'product_code': 'รหัสสินค้า',
                'product_name': 'ชื่อสินค้า',
                'warehouse': 'คลัง',
                'last_unit_price': 'ราคาล่าสุด',
                'last_weight_per_unit': 'น้ำหนักล่าสุด'
            }, inplace=True)

            default_filename = f"products_master_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="บันทึกไฟล์ข้อมูลสินค้า",
                initialfile=default_filename,
                parent=self
            )

            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("สำเร็จ", f"Export ข้อมูลสินค้าเรียบร้อยแล้วที่:\n{save_path}\n\n*ห้ามแก้ไขคอลัมน์ System_ID หากต้องการอัปเดตข้อมูลเดิม*", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถ Export ได้: {e}", parent=self)
        finally:
             if conn: self.app_container.release_connection(conn)

    # --- ฟังก์ชัน Import ที่แก้ไขแล้ว (เช็ค ID ก่อน) ---
    def _import_products(self):
        """Import ข้อมูลสินค้าจาก Excel (Update by ID or Code, or Insert)"""
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel ข้อมูลสินค้า",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            parent=self
        )
        
        if not file_path:
            return

        conn = None
        try:
            # 1. อ่านไฟล์
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            # 2. Map ชื่อคอลัมน์
            column_map = {
                'System_ID': 'id', 'id': 'id', 'ID': 'id', # รองรับชื่อ ID หลายแบบ
                'รหัสสินค้า': 'product_code', 'product_code': 'product_code', 'code': 'product_code',
                'ชื่อสินค้า': 'product_name', 'product_name': 'product_name', 'name': 'product_name',
                'คลัง': 'warehouse', 'warehouse': 'warehouse',
                'ราคาล่าสุด': 'last_unit_price', 'last_unit_price': 'last_unit_price', 'price': 'last_unit_price',
                'น้ำหนักล่าสุด': 'last_weight_per_unit', 'last_weight_per_unit': 'last_weight_per_unit', 'weight': 'last_weight_per_unit'
            }
            
            df.columns = [str(c).strip() for c in df.columns]
            rename_dict = {}
            for col in df.columns:
                if col in column_map:
                    rename_dict[col] = column_map[col]
            
            df.rename(columns=rename_dict, inplace=True)
            
            # 3. ตรวจสอบคอลัมน์ที่จำเป็น
            if 'product_code' not in df.columns or 'product_name' not in df.columns:
                messagebox.showerror("รูปแบบไฟล์ไม่ถูกต้อง", "ไฟล์ต้องมีคอลัมน์ 'รหัสสินค้า' และ 'ชื่อสินค้า' เป็นอย่างน้อย", parent=self)
                return

            # 4. ทำความสะอาดข้อมูล
            df['product_code'] = df['product_code'].astype(str).str.strip()
            df['product_name'] = df['product_name'].astype(str).str.strip()
            if 'warehouse' in df.columns:
                df['warehouse'] = df['warehouse'].fillna('').astype(str).str.strip()
            else:
                df['warehouse'] = ''
            
            if 'last_unit_price' in df.columns:
                df['last_unit_price'] = pd.to_numeric(df['last_unit_price'], errors='coerce').fillna(0)
            else: df['last_unit_price'] = 0
            
            if 'last_weight_per_unit' in df.columns:
                df['last_weight_per_unit'] = pd.to_numeric(df['last_weight_per_unit'], errors='coerce').fillna(0)
            else: df['last_weight_per_unit'] = 0

            # 5. เริ่มกระบวนการ Import
            if not messagebox.askyesno("ยืนยันการนำเข้า", f"พบข้อมูล {len(df)} รายการ\nต้องการนำเข้าและอัปเดตข้อมูลหรือไม่?", parent=self):
                return

            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                updated_count = 0
                inserted_count = 0
                
                for _, row in df.iterrows():
                    code = row['product_code']
                    name = row['product_name']
                    if not code or not name: continue
                    
                    wh = row.get('warehouse', '')
                    price = row.get('last_unit_price', 0)
                    weight = row.get('last_weight_per_unit', 0)
                    row_id = row.get('id') # ลองดึง ID มาดู

                    target_id = None
                    
                    # [Logic ใหม่]
                    # 1. ถ้ามี ID มาในไฟล์ ให้ลองหาด้วย ID ก่อน
                    if pd.notna(row_id):
                        cursor.execute("SELECT id FROM products WHERE id = %s", (row_id,))
                        res = cursor.fetchone()
                        if res:
                            target_id = res[0]
                    
                    # 2. ถ้าไม่มี ID หรือหา ID ไม่เจอ -> ให้ลองหาด้วย Product Code
                    if not target_id:
                        cursor.execute("SELECT id FROM products WHERE product_code = %s", (code,))
                        res = cursor.fetchone()
                        if res:
                            target_id = res[0]

                    if target_id:
                        # Update (รวมถึงกรณีเปลี่ยนรหัสสินค้า ก็จะทำได้ถ้าอิงตาม ID)
                        cursor.execute("""
                            UPDATE products 
                            SET product_code = %s, product_name = %s, warehouse = %s, last_unit_price = %s, last_weight_per_unit = %s, last_updated = NOW()
                            WHERE id = %s
                        """, (code, name, wh, price, weight, target_id))
                        updated_count += 1
                    else:
                        # Insert
                        cursor.execute("""
                            INSERT INTO products (product_code, product_name, warehouse, last_unit_price, last_weight_per_unit, last_updated)
                            VALUES (%s, %s, %s, %s, %s, NOW())
                        """, (code, name, wh, price, weight))
                        inserted_count += 1
                
                conn.commit()
                messagebox.showinfo("สำเร็จ", f"นำเข้าข้อมูลเรียบร้อยแล้ว\n- เพิ่มใหม่: {inserted_count} รายการ\n- อัปเดต: {updated_count} รายการ", parent=self)
                self.load_products()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการนำเข้า: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

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
            # [แก้ไข] เอาบรรทัดที่ล็อก readonly ออกแล้ว เพื่อให้แก้ไขได้

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
                    
                    # [เพิ่ม] ตรวจสอบรหัสซ้ำ (กรณีเปลี่ยนรหัส)
                    cursor.execute("SELECT id FROM products WHERE product_code = %s AND id != %s", (code, product_id))
                    if cursor.fetchone():
                        messagebox.showerror("ข้อมูลซ้ำ", f"รหัสสินค้า '{code}' มีอยู่ในระบบแล้ว", parent=self)
                        return

                    # [แก้ไข] อัปเดต product_code ด้วย
                    cursor.execute("""
                        UPDATE products 
                        SET product_code = %s, product_name = %s, warehouse = %s, last_updated = %s
                        WHERE id = %s
                    """, (code, name, warehouse, datetime.now(), product_id))
                    
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
    
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None, initial_so_number=None):
        self.master = master
        self.app_container = app_container
        self.user_key, self.user_name = user_key, user_name
        self.theme = self.app_container.THEME["purchasing"]
        self.sale_theme = self.app_container.THEME["sale"]
        self.async_helper = SimpleAsyncHelper(self)

        self.is_running = True

        super().__init__(master, corner_radius=0, fg_color="#EDE9FE")
        self.shipping_to_stock_vat_var = tk.StringVar(value="VAT")
        self.shipping_to_site_vat_var = tk.StringVar(value="VAT")
        self.shipping_to_stock_wht_var = tk.StringVar(value="ไม่มีหัก")
        self.shipping_to_site_wht_var = tk.StringVar(value="ไม่มีหัก")
        self.shipping_to_stock_wht_display_var = tk.StringVar(value="0.00")
        self.shipping_to_site_wht_display_var = tk.StringVar(value="0.00")
        
        # --- [🔥 เพิ่ม] ตัวแปรสำหรับค่าตัด/เจาะ ---
        self.cutting_vat_var = tk.StringVar(value="CASH") 
        self.cutting_wht_var = tk.StringVar(value="No")
        self.cutting_vat_display_var = tk.StringVar(value="0.00")
        self.cutting_wht_display_var = tk.StringVar(value="0.00")
        self.cutting_total_display_var = tk.StringVar(value="0.00")
        # ------------------------------------
        
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
        
        self.supplier_completion_data = []
        self.product_completion_data = []

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
        

        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)

        self._create_header()

        self.tab_view = CTkTabview(self, text_color="black")
        self.tab_view.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.tab_view.add("สร้างใบสั่งซื้อ (PO)")
        self.po_pane = self.tab_view.tab("สร้างใบสั่งซื้อ (PO)")
        self.po_pane.grid_rowconfigure(0, weight=1)
        self.po_pane.grid_columnconfigure(0, weight=1)
        
        self._create_po_form_layout(self.po_pane)
        
        self.tab_view.add("Daily Report")
        report_tab = self.tab_view.tab("Daily Report")
        report_tab.grid_columnconfigure(0, weight=1)
        report_tab.grid_rowconfigure(0, weight=1)
        
        self.daily_report = DailyReportWidget(report_tab, self.app_container)
        self.daily_report.pack(fill="both", expand=True)

        self.tab_view.add("เทียบราคา (Cost Benchmark)")
        benchmark_tab = self.tab_view.tab("เทียบราคา (Cost Benchmark)")
        benchmark_tab.grid_columnconfigure(0, weight=1)
        benchmark_tab.grid_rowconfigure(0, weight=1)
        
        # ดึงหน้าจอจาก cost_benchmark.py มาฝังในแท็บนี้
        self.cost_benchmark_view = CostBenchmarkScreen(benchmark_tab, self.app_container)
        self.cost_benchmark_view.pack(fill="both", expand=True)

        self.tab_view.add("Dashboard เทียบราคา")
        dashboard_cost_tab = self.tab_view.tab("Dashboard เทียบราคา")
        dashboard_cost_tab.grid_columnconfigure(0, weight=1)
        dashboard_cost_tab.grid_rowconfigure(0, weight=1)
        
        # ดึงหน้าจอ Dashboard มาฝัง
        self.dashboard_cost_view = DashboardCostScreen(dashboard_cost_tab, self.app_container)
        self.dashboard_cost_view.pack(fill="both", expand=True)
        # ==========================================================

        self._load_supplier_data()
        self._load_product_master_data()

        self._poll_and_update_tasks_badge()
        self.bind("<Destroy>", self._on_destroy)
        
    def _open_transport_manager(self):
        """เปิดหน้าต่างค้นหาและจัดการค่าขนส่ง"""
        # ต้อง Import Class มาก่อน (ระวัง Circular Import)
        try:
            from history_windows import TransportPOSearchDialog
            TransportPOSearchDialog(self, self.app_container)
        except ImportError:
            messagebox.showerror("Error", "ไม่พบโมดูล TransportPOSearchDialog", parent=self)
        except Exception as e:
            print(traceback.format_exc())
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
    
    def _edit_so_number(self):
        if not self.current_commission_data:
            messagebox.showwarning("ยังไม่ได้เลือก SO", "กรุณาเลือก SO ที่ต้องการแก้ไขก่อน", parent=self)
            return

        old_so_number = self.current_commission_data.get('so_number')
        record_id = self.current_commission_data.get('id')

        dialog = CTkInputDialog(text=f"กรุณาใส่เลข SO ใหม่สำหรับ '{old_so_number}':", title="แก้ไขเลขที่ Sales Order")
        new_so_number = dialog.get_input()

        if not new_so_number or not new_so_number.strip():
            return
        
        new_so_number = new_so_number.strip().upper()
        if new_so_number == old_so_number:
            return

        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการเปลี่ยนเลข SO จาก '{old_so_number}' เป็น '{new_so_number}' ใช่หรือไม่?\n\n(PO ทั้งหมดที่เกี่ยวข้องกับ SO นี้จะถูกอัปเดตด้วย)", parent=self):
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT id FROM commissions WHERE so_number = %s AND is_active = 1", (new_so_number,))
                if cursor.fetchone():
                    raise ValueError(f"เลขที่ SO '{new_so_number}' นี้มีอยู่แล้วในระบบ ไม่สามารถใช้ซ้ำได้")

                cursor.execute("UPDATE commissions SET so_number = %s WHERE id = %s", (new_so_number, record_id))
                cursor.execute("UPDATE purchase_orders SET so_number = %s WHERE so_number = %s", (new_so_number, old_so_number))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "แก้ไขเลขที่ SO เรียบร้อยแล้ว", parent=self)
            
            self.handle_clear_button_press(confirm=False)
            self._refresh_so_list()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("เกิดข้อผิดพลาด", str(e), parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _cancel_so_record(self):
        if not self.current_commission_data:
            messagebox.showwarning("ยังไม่ได้เลือก SO", "กรุณาเลือก SO ที่ต้องการยกเลิกก่อน", parent=self)
            return
            
        so_number = self.current_commission_data.get('so_number')
        record_id = self.current_commission_data.get('id')
        
        msg = (f"คุณต้องการยกเลิก SO: '{so_number}' ใช่หรือไม่?\n\n"
               "การกระทำนี้จะเปลี่ยนสถานะ SO เป็น 'Cancelled' และซ่อนจากคิวงานปกติ "
               "รวมถึงยกเลิก PO ที่เกี่ยวข้องทั้งหมดด้วย\n\n**การกระทำนี้ไม่สามารถย้อนกลับได้**")

        if not messagebox.askyesno("ยืนยันการยกเลิก SO", msg, icon="warning", parent=self):
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("UPDATE commissions SET status = 'Cancelled by PU', is_active = 0 WHERE id = %s", (record_id,))
                cursor.execute("UPDATE purchase_orders SET status = 'Cancelled' WHERE so_number = %s", (so_number,))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ยกเลิก SO: '{so_number}' เรียบร้อยแล้ว", parent=self)
            
            self.handle_clear_button_press(confirm=False)
            self._refresh_so_list()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการยกเลิก SO: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _lookup_so_details(self):
        """
        (เวอร์ชันแก้ไข) เปิดหน้าต่าง Input Dialog และส่ง SO Number ไปให้ SOFinderDialog
        """
        dialog = CTkInputDialog(text="กรุณาใส่ SO Number ที่ต้องการค้นหา:", title="ค้นหาข้อมูล Sales Order")
        so_to_find = dialog.get_input()

        if so_to_find and so_to_find.strip():
            # <<< แก้ไข: เปลี่ยนจากการเรียก SODetailViewer มาเป็น SOFinderDialog ที่เราสร้างใหม่ >>>
            SOFinderDialog(master=self, so_number=so_to_find.strip().upper())
            
        elif so_to_find is not None: # ถ้าผู้ใช้กด OK แต่ไม่กรอกอะไร
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก SO Number", parent=self)

    def _refresh_so_list(self):
        print("Refreshing SO ComboBox list...")
        new_so_list = self._get_commission_so_numbers_formatted()
        
        if hasattr(self, 'so_entry') and isinstance(self.so_entry, CTkComboBox):
            self.so_entry.configure(values=new_so_list)
            self.so_entry.set("") 
            messagebox.showinfo("รีเฟรช", f"อัปเดตรายการ SO เรียบร้อยแล้ว\nพบ {len(new_so_list) - 1} รายการที่พร้อมดำเนินการ", parent=self)
        else:
            messagebox.showwarning("ผิดพลาด", "ไม่สามารถรีเฟรชรายการได้ Widget ไม่ถูกต้อง", parent=self)
            
    def _get_commission_so_numbers_formatted(self):
        try:
            query = """
                SELECT c.so_number, c.customer_name, u.sale_name
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE (c.status = 'Pending PU' OR (c.status = 'PO In Progress' AND c.user_key = %s)) 
                AND c.is_active = 1 
                ORDER BY c.timestamp DESC
            """
            df = pd.read_sql_query(query, self.pg_engine, params=(self.user_key,))
            
            if df.empty:
                return [""]

            formatted_list = [""] 
            MAX_CUST_NAME_LEN = 35 

            for _, row in df.iterrows():
                so = row['so_number']
                cust = str(row['customer_name'] or '')
                sale = str(row.get('sale_name') or 'N/A')

                if len(cust) > MAX_CUST_NAME_LEN:
                    cust = cust[:MAX_CUST_NAME_LEN] + "..."
                
                display_string = f"{so} | {cust} (เซลส์: {sale})"
                formatted_list.append(display_string)
            
            return formatted_list
        except Exception as e:
            print(f"Error fetching SO list for ComboBox: {e}")
            return [""]

    def _lookup_so_details(self):
        dialog = CTkInputDialog(text="กรุณาใส่ SO Number ที่ต้องการค้นหา:", title="ค้นหาข้อมูล Sales Order")
        so_to_find = dialog.get_input()

        if so_to_find and so_to_find.strip():
            SOFinderDialog(master=self, so_number=so_to_find.strip().upper())
            
        elif so_to_find is not None:
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก SO Number", parent=self)
    
    def sync_transport_cost_to_po(self, so_number):
        """
        ดึงข้อมูลค่าขนส่งและค่าย้ายจาก SO มาใส่ในช่องค่าใช้จ่ายของ PO โดยอัตโนมัติ
        """
        if not so_number:
            return

        # ตัดส่วนที่เกินออก เช่น "SO123 | Customer" -> "SO123"
        if "|" in so_number:
            so_number = so_number.split("|")[0].strip()

        print(f"DEBUG: Syncing transport cost for SO: {so_number}")
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ดึงทั้ง 'shipping_cost' (ค่าส่ง Site) และ 'relocation_cost' (ค่าย้าย Stock)
                cursor.execute("""
                    SELECT shipping_cost, relocation_cost 
                    FROM commissions 
                    WHERE so_number = %s AND is_active = 1 
                    LIMIT 1
                """, (so_number,))
                
                result = cursor.fetchone()
                
                if result:
                    shipping_site_val = result[0] or 0.0
                    relocation_stock_val = result[1] or 0.0
                    
                    print(f"DEBUG: Found info -> Site: {shipping_site_val}, Stock: {relocation_stock_val}")

                    # 1. อัปเดตช่อง "ค่าจัดส่งเข้าไซต์" (Section 2)
                    if hasattr(self, 'shipping_to_site_cost_entry'):
                        current_val = self.shipping_to_site_cost_entry.get()
                        # ถ้าช่องว่าง หรือเป็น 0 ให้เติมค่า
                        if not current_val or utils.convert_to_float(current_val) == 0:
                            utils.set_entry_text(self.shipping_to_site_cost_entry, f"{shipping_site_val:.2f}")
                            if shipping_site_val > 0 and hasattr(self, 'shipping_to_site_type_var'):
                                self.shipping_to_site_type_var.set("Aplus Logistic") 

                    # 2. อัปเดตช่อง "ค่าจัดส่งเข้าสต๊อก" (Section 1 - มาจากค่าย้าย)
                    if hasattr(self, 'shipping_to_stock_cost_entry'):
                        current_val = self.shipping_to_stock_cost_entry.get()
                        if not current_val or utils.convert_to_float(current_val) == 0:
                            utils.set_entry_text(self.shipping_to_stock_cost_entry, f"{relocation_stock_val:.2f}")
                            if relocation_stock_val > 0 and hasattr(self, 'shipping_to_stock_type_var'):
                                self.shipping_to_stock_type_var.set("Aplus Logistic")

                    # คำนวณยอดรวมใหม่ทันที
                    self._update_summary()
                else:
                    print(f"ℹ️ Not found transport info for {so_number}")

        except Exception as e:
            print(f"Error syncing transport cost: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _update_summary(self, *args):
        # --- 1. คำนวณยอดรวมจากรายการสินค้า (Product Subtotal) ---
        product_subtotal = 0  # ยอดรวมต้นทุนสินค้าทั้งหมด (รวมตัวฟรีด้วย)
        supplier_payable_product_base = 0 # ยอดฐานสินค้าที่จะเอาไปคิดเงินจ่ายซัพฯ
        
        overall_total_weight = 0
        
        for row_dict in self.product_rows:
            try:
                if not row_dict["name"].winfo_exists(): continue
                
                # ดึงข้อมูลรหัสสินค้ามาเช็ค
                code = ""
                if hasattr(row_dict["code"], "get"):
                     code = row_dict["code"].get().strip().upper()

                qty = utils.convert_to_float(row_dict["qty"].get())
                price = utils.convert_to_float(row_dict["price"].get())
                weight = utils.convert_to_float(row_dict["weight"].get())
                discount_val = utils.convert_to_float(row_dict["discount_entry"].get())
                discount_type = row_dict["discount_type_var"].get()

                line_total = qty * price
                discount_amount = (line_total * (discount_val / 100.0)) if discount_type == "%" else discount_val
                row_final_price = line_total - discount_amount
                row_final_weight = qty * weight

                # บวกเข้าต้นทุนรวมเสมอ (เพราะถือเป็นต้นทุนของ PO)
                product_subtotal += row_final_price
                overall_total_weight += row_final_weight
                
                # --- [🔥 Logic พิเศษ] EXP-0079A ไม่นำไปคิดเงินจ่ายซัพฯ ---
                if code == 'EXP-0079A':
                    pass # ไม่บวกเข้า supplier_payable
                else:
                    supplier_payable_product_base += row_final_price
                # -----------------------------------------------------

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
        cutting_cost = utils.convert_to_float(self.cutting_cost_entry.get())
        end_of_bill_discount = utils.convert_to_float(self.end_of_bill_discount_entry.get())
        
        p1 = utils.convert_to_float(self.payment_entries["Payment 1"]["amount"].get())
        p2 = utils.convert_to_float(self.payment_entries["Payment 2"]["amount"].get())
        full_payment = utils.convert_to_float(self.payment_entries["Full Payment"]["amount"].get())
      
        # --- 3. คำนวณยอดที่ต้องชำระให้ซัพพลายเออร์ ---
        # ใช้ supplier_payable_product_base (ที่หัก EXP-0079A แล้ว) มาเป็นฐานตั้งต้น
        supplier_payable_vatable = supplier_payable_product_base - end_of_bill_discount
        supplier_payable_non_vatable = 0.0
        separate_shipping_cost = 0.0
        
        # WHT (Stock/Site)
        shipping_stock_wht_amount = 0.0
        shipping_site_wht_amount = 0.0
        cutting_wht_amount = 0.0 # สำหรับแสดงผล

        # จัดการค่าส่ง Stock (EXP-0174 Mapping)
        if self.shipping_to_stock_type_var.get() == 'ซัพพลายเออร์จัดส่ง':
            stock_wht_type = self.shipping_to_stock_wht_var.get()
            if stock_wht_type == "1%": shipping_stock_wht_amount = shipping_stock_cost * 0.01
            elif stock_wht_type == "3%": shipping_stock_wht_amount = shipping_stock_cost * 0.03
            
            if self.shipping_to_stock_vat_var.get() == 'VAT': supplier_payable_vatable += shipping_stock_cost
            else: supplier_payable_non_vatable += shipping_stock_cost
        else:
            separate_shipping_cost += shipping_stock_cost

        # จัดการค่าส่ง Site (EXP-0006 Mapping)
        if self.shipping_to_site_type_var.get() == 'ซัพพลายเออร์จัดส่ง':
            site_wht_type = self.shipping_to_site_wht_var.get()
            if site_wht_type == "1%": shipping_site_wht_amount = shipping_site_cost * 0.01
            elif site_wht_type == "3%": shipping_site_wht_amount = shipping_site_cost * 0.03

            if self.shipping_to_site_vat_var.get() == 'VAT': supplier_payable_vatable += shipping_site_cost
            else: supplier_payable_non_vatable += shipping_site_cost
        else:
            separate_shipping_cost += shipping_site_cost
            
        # จัดการค่าตัด/เจาะ (Cutting Logic)
        cutting_wht_type = self.cutting_wht_var.get()
        if cutting_wht_type == "1%": cutting_wht_amount = cutting_cost * 0.01
        elif cutting_wht_type == "3%": cutting_wht_amount = cutting_cost * 0.03
        
        # ค่าตัดเจาะยังไงก็ต้องจ่ายซัพฯ (ตามที่คุณบอกว่าแยกจากยอดต้องชำระซัพไม่ได้) 
        # *แก้ไข: คุณบอกว่า "ไม่ไปรวมกับช่องที่ต้องชำระซับ แต่จะไปรวมอยู่ในทุนแทน"
        # ดังนั้น Code บรรทัดนี้ต้องเอาออก หรือต้องเช็คดีๆ ว่าตกลงจ่ายใคร
        # ถ้าจ่ายซัพฯเจ้านี้ ต้องบวก If ไม่จ่าย (จ่ายเงินสดหน้างาน) ไม่ต้องบวก
        if self.cutting_vat_var.get() == 'VAT': 
             pass # สมมติว่าไม่รวมในยอดบิลซัพพลายเออร์หลัก (ตามที่คุณแจ้งล่าสุด)
             # supplier_payable_vatable += cutting_cost 
        else: 
             pass
             # supplier_payable_non_vatable += cutting_cost

        # --- 4. คำนวณ VAT, WHT, ยอดสุทธิ ---
        vat7_amount = supplier_payable_vatable * 0.07 if hasattr(self, 'vat_checkbox') and self.vat_checkbox.get() else 0.0
        product_wht3_amount = supplier_payable_vatable * 0.03 if hasattr(self, 'vat3_checkbox') and self.vat3_checkbox.get() else 0.0
        
        # รวมยอดหัก ณ ที่จ่ายทั้งหมด (เฉพาะส่วนที่จ่ายผ่านบิลนี้)
        total_wht_deduction = product_wht3_amount + shipping_stock_wht_amount + shipping_site_wht_amount 
        
        grand_total_payable_to_supplier = (supplier_payable_vatable + vat7_amount - total_wht_deduction) + supplier_payable_non_vatable
        total_deposit = p1 + p2
        balance_due = grand_total_payable_to_supplier - total_deposit - full_payment

        # --- 5. อัปเดต UI ---
        def set_readonly_val(entry, value):
            if entry and entry.winfo_exists():
               entry.configure(state="normal"); entry.delete(0, "end")
               entry.insert(0, f"{value:,.2f}"); entry.configure(state="readonly")
   
        # ต้นทุนรวม PO = (สินค้าทั้งหมด รวม EXP-0079A - ส่วนลด) + ค่าตัด/เจาะ
        total_po_cost = (product_subtotal - end_of_bill_discount) + cutting_cost
        set_readonly_val(self.total_cost_entry, total_po_cost)

        set_readonly_val(self.total_weight_summary_entry, overall_total_weight)
        set_readonly_val(self.vat7_entry, vat7_amount)
        set_readonly_val(self.vat3_entry, product_wht3_amount)
        set_readonly_val(self.grand_total_with_vat_entry, supplier_payable_vatable + supplier_payable_non_vatable + vat7_amount)
        set_readonly_val(self.grand_total_payable_entry, grand_total_payable_to_supplier)
        set_readonly_val(self.separate_shipping_entry, separate_shipping_cost)
    
        self.total_deposit_var.set(f"{total_deposit:,.2f}")
      
        # Display VAT/WHT
        stock_vat_display = shipping_stock_cost * 0.07 if self.shipping_to_stock_vat_var.get() == 'VAT' else 0.0
        site_vat_display = shipping_site_cost * 0.07 if self.shipping_to_site_vat_var.get() == 'VAT' else 0.0
        cutting_vat_display = cutting_cost * 0.07 if self.cutting_vat_var.get() == 'VAT' else 0.0
        
        self.shipping_to_stock_vat_display_var.set(f"{stock_vat_display:,.2f}")
        self.shipping_to_site_vat_display_var.set(f"{site_vat_display:,.2f}")
        self.cutting_vat_display_var.set(f"{cutting_vat_display:,.2f}")
        
        self.shipping_to_stock_wht_display_var.set(f"{shipping_stock_wht_amount:,.2f}")
        self.shipping_to_site_wht_display_var.set(f"{shipping_site_wht_amount:,.2f}")
        self.cutting_wht_display_var.set(f"{cutting_wht_amount:,.2f}")
        
        cutting_total_val = cutting_cost + cutting_vat_display
        self.cutting_total_display_var.set(f"{cutting_total_val:,.2f}")

        if hasattr(self, 'balance_due_entry') and self.balance_due_entry.winfo_exists():
            if abs(balance_due) < 0.01: text, text_color, bg_color = "ยอดชำระครบถ้วน", "#15803D", "#BBF7D0"
            elif balance_due < 0: text, text_color, bg_color = f"ชำระเกิน {abs(balance_due):,.2f}", "#15803D", "#BBF7D0"
            else: text, text_color, bg_color = f"ยอดค้างชำระ {balance_due:,.2f}", "#B91C1C", "#FECACA"
            self.balance_due_var.set(text); self.balance_due_entry.configure(text_color=text_color, fg_color=bg_color)
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
        self.so_form_widgets['relocation_cost_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['relocation_vat_calc_var'] = tk.StringVar(value="0.00")
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

        # --- [🔥 ส่วนที่เพิ่มใหม่] ปุ่มจัดการค่าขนส่ง ---
        CTkButton(button_container, 
                  text="🚚 จัดการค่าขนส่ง & ค่าตัด", 
                  command=self._open_transport_manager, 
                  fg_color="#8B5CF6", hover_color="#7C3AED" # สีม่วง
        ).pack(side="left", padx=5)
        # ---------------------------------------------

        CTkButton(button_container, text="🔍 ค้นหา SO", command=self._lookup_so_details, fg_color="#0891B2").pack(side="left", padx=5)

        CTkButton(button_container, text="📖 ดูประวัติ PO", command=lambda: self.app_container.show_history_window(), fg_color="#64748B").pack(side="left", padx=5)
        CTkButton(button_container, text="🔧 จัดการสินค้า", command=self._open_product_management_window, fg_color="#6D28D9", hover_color="#5B21B6").pack(side="left", padx=5)
        
        CTkButton(button_container, text="Export PDF (PO อนุมัติ)", command=lambda: export_approved_pos_to_pdf(self, self.pg_engine), fg_color="#c026d3", hover_color="#a21caf").pack(side="left", padx=5)
        export_button = CTkButton(button_container, text="Export Excel (PO อนุมัติ)", command=lambda: export_approved_pos_to_excel(self, self.pg_engine), fg_color="#107C41", hover_color="#0B532B")
        export_button.pack(side="left", padx=5)
        CTkButton(button_container, text="(ล้างฟอร์ม PO)", command=self.handle_clear_button_press, fg_color="#E11D48").pack(side="left", padx=5)
        self.toggle_so_data_button = CTkButton(button_container, text="ดูข้อมูล SO", command=self._open_so_popup, fg_color=self.sale_theme.get("primary", "#3B82F6"))
        self.toggle_so_data_button.pack(side="left", padx=5)
        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").pack(side="right", padx=(5, 0))
    
    def _open_so_selection_dialog(self):
        self.app_container.open_so_print_dialog()

    def _open_po_selection_dialog(self):
     try:
        POSelectionDialog(self, self.pg_engine, print_callback=self._print_selected_po)
     except Exception as e:
        messagebox.showerror("Error", f"Could not open PO selection window: {e}", parent=self)
        traceback.print_exc()

    def _on_destroy(self, event):
        # ตรวจสอบว่า Event นี้เกิดจากตัว PurchasingScreen เองหรือไม่ (ไม่ใช่จาก Widget ลูก)
        if hasattr(event, 'widget') and event.widget is self:
            
            # [🔥 จุดสำคัญ] สั่งหยุด Loop การทำงานทันที
            self.is_running = False 
            
            # หยุด Polling
            self._stop_polling()
            
            # เคลียร์และปิดหน้าต่างย่อย (Popup) ทั้งหมดเพื่อคืนหน่วยความจำ
            if self.sales_data_popup and self.sales_data_popup.winfo_exists():
                self.sales_data_popup._on_popup_close() # เรียก cleanup ของ popup ถ้ามี
                self.sales_data_popup.destroy()
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
        # ถ้าถูกสั่งปิดแล้ว ให้หยุดทันที ไม่ต้องทำต่อ
        if not self.is_running:
            return

        try:
            if not self.winfo_exists():
                return
        except Exception:
            return

        self._update_tasks_badge()
        
        # ตั้งเวลาทำงานรอบถัดไป
        if self.is_running:
            try:
                self.polling_job_id = self.after(30000, self._poll_and_update_tasks_badge)
            except Exception:
                pass

    def _update_tasks_badge(self):
        if not self.is_running: return
        
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM notifications WHERE user_key_to_notify = %s AND is_read = FALSE AND message LIKE 'SO ใหม่รอสร้าง PO%%'", (self.user_key,)); new_so_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM purchase_orders WHERE user_key = %s AND status = 'Rejected'", (self.user_key,)); rejected_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM purchase_orders WHERE user_key = %s AND status = 'Draft'", (self.user_key,)); draft_count = cursor.fetchone()[0]
            
            total_tasks = new_so_count + rejected_count + draft_count
            
            # อัปเดต UI ถ้ายังเปิดอยู่
            if self.is_running and hasattr(self, 'tasks_button'):
                try:
                    self.tasks_button.configure(text=f"My Tasks 🔔 ({total_tasks})")
                    if total_tasks > 0: self.tasks_button.configure(fg_color="#F59E0B", hover_color="#D97706")
                    else: self.tasks_button.configure(fg_color=("#3B8ED0", "#1F6AA5"), hover_color=("#36719F", "#144870"))
                except Exception:
                    pass # กัน Error ถ้าปุ่มหายไปแล้ว

        except Exception as e:
            if "application has been destroyed" not in str(e):
                print(f"Error updating tasks badge: {e}")
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
        if self.current_commission_data is None:
            messagebox.showinfo("ข้อมูล SO", "กรุณาเลือก SO Number ก่อน", parent=self)
            return
        if self.sales_data_popup and self.sales_data_popup.winfo_exists():
            self.sales_data_popup.focus()
            return
        
        self.sales_data_popup = SOPopupWindow(
            master=self,
            app_container=self.app_container,
            sales_data=self.current_commission_data,
            so_shared_vars=self.so_form_widgets,
            sale_theme=self.sale_theme,
            on_save_callback=lambda: self._on_so_selected(self.so_entry.get(), is_editing=True)
        )
    
    def _save_so_changes_from_popup(self, so_id, so_shared_vars_data, current_popup_widgets_ref):
        updated_data = {}
        
        key_map = {
            'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id',
            'credit_term_entry': 'credit_term', 'pickup_location_entry': 'pickup_location',
            'pickup_rego_entry': 'pickup_registration', 'bill_date_selector': 'bill_date', 
            'delivery_date_selector': 'delivery_date', 'payment_date_selector': 'payment_date', 
            'date_to_wh_selector': 'date_to_warehouse', 'date_to_customer_selector': 'date_to_customer',
            'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost',
            'relocation_cost_entry': 'relocation_cost', 'credit_card_fee_entry': 'credit_card_fee',
            'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons',
            'giveaway_value_entry': 'giveaways', 'cash_product_input_entry': 'cash_product_input',
            'cash_actual_payment_entry': 'cash_actual_payment'
        }

        numeric_keywords = ['amount', 'cost', 'fee', 'wht', 'price', 'percent', 'coupons', 'giveaways']

        for widget_key, data_key in key_map.items():
            value = None
            if widget_key in current_popup_widgets_ref:
                widget = current_popup_widgets_ref[widget_key]
                if widget and widget.winfo_exists():
                    if isinstance(widget, DateSelector):
                        value = widget.get_date()
                    elif isinstance(widget, (NumericEntry, CTkEntry)):
                        value = widget.get()
                        if any(keyword in data_key for keyword in numeric_keywords):
                            value = utils.convert_to_float(value)
            
            if value is not None:
                updated_data[data_key] = value

        shared_vars_map = {
            'delivery_type_var': 'delivery_type',
            'sales_service_vat_option': 'sales_service_vat_option',
            'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option': 'other_service_fee_vat_option',
            'shipping_vat_option_var': 'shipping_vat_option',
            'credit_card_fee_vat_option_var': 'credit_card_fee_vat_option',
        }
        for var_key, data_key in shared_vars_map.items():
             if var_key in so_shared_vars_data:
                updated_data[data_key] = so_shared_vars_data[var_key].get()

        p1 = utils.convert_to_float(current_popup_widgets_ref.get('payment1_amount_entry').get())
        p2 = utils.convert_to_float(current_popup_widgets_ref.get('payment2_amount_entry').get())
        updated_data['total_payment_amount'] = p1 + p2

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                set_clauses = [f'"{k}" = %s' for k in updated_data.keys()]
                params = list(updated_data.values()) + [so_id]
                sql_update = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
                
                cursor.execute(sql_update, tuple(params))
            conn.commit()
            
            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล SO Number: {self.current_commission_data.get('so_number')} เรียบร้อยแล้ว", parent=self)
            
            reloaded_df = pd.read_sql_query("SELECT * FROM commissions WHERE id = %s", self.pg_engine, params=(so_id,))
            if not reloaded_df.empty:
                self.current_commission_data = reloaded_df.iloc[0].to_dict()
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล SO:\n{e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)


        def _safe_get_float(entry_widget):
            if entry_widget and hasattr(entry_widget, 'winfo_exists') and entry_widget.winfo_exists():
                try: return utils.convert_to_float(entry_widget.get())
                except (ValueError, tk.TclError): return 0.0
            return 0.0

        for widget_key, db_col_name in key_map.items():
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


    def _on_so_selected(self, selection_string: str, is_editing: bool = False):
        so_number = ""
        if selection_string and '|' in selection_string:
            so_number = selection_string.split('|')[0].strip()
        else:
            so_number = selection_string.strip()

        if not so_number:
            self.handle_clear_button_press(confirm=False)
            return
        conn = self.app_container.get_connection()
        self.sync_transport_cost_to_po(so_number)
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT id, status, user_key FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", (so_number,))
                so_id_in_commissions, so_status, so_user_key = cursor.fetchone() if cursor.rowcount > 0 else (None, None, None)

                if not is_editing:
                    if so_id_in_commissions is None:
                        messagebox.showwarning("ไม่พบ SO", f"ไม่พบ SO Number: {so_number} ในสถานะที่พร้อมดำเนินการ", parent=self); self.so_entry.set(""); return
                    if so_status == 'Pending PU':
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

    def handle_clear_button_press(self, confirm=True):
        if confirm and not messagebox.askyesno("ยืนยัน", "คุณต้องการล้างข้อมูลทั้งหมดในฟอร์มใช่หรือไม่?", parent=self): return
        if self.current_commission_data:
            so_id_to_release = self.current_commission_data.get('id')
            so_number_to_release = self.current_commission_data.get('so_number')
            conn = self.app_container.get_connection()
            try:
                with conn.cursor() as cursor:
                    cursor.execute("UPDATE commissions SET status = 'Pending PU', user_key = NULL, claim_timestamp = NULL WHERE id = %s AND status = 'PO In Progress' AND user_key = %s", (so_id_to_release, self.user_key)); conn.commit()
                    self.so_entry.configure(values=self._get_commission_so_numbers())
            except Exception as e:
                if conn: conn.rollback(); print(f"Error releasing SO status: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)
        
        self._clear_form(confirm=False)

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

        so_selection_frame = CTkFrame(top_frame, fg_color="transparent")
        so_selection_frame.grid(row=0, column=1, columnspan=7, sticky="ew", padx=10, pady=5)
        so_selection_frame.grid_columnconfigure(0, weight=1)

        self.so_entry = CTkComboBox(
            so_selection_frame,
            values=self._get_commission_so_numbers_formatted(),
            command=self._on_so_selected
        )
        self.so_entry.set("")
        self.so_entry.grid(row=0, column=0, sticky="ew")

        # +++ START: แก้ไขการจัดวางปุ่มใหม่ทั้งหมด +++
        # 1. สร้าง Frame ใหม่สำหรับวางปุ่มโดยเฉพาะ
        action_buttons_frame = CTkFrame(so_selection_frame, fg_color="transparent")
        action_buttons_frame.grid(row=0, column=1, sticky="e", padx=(10, 0))

        # 2. ใช้ .pack() เพื่อเรียงปุ่มจากซ้ายไปขวาภายใน Frame ใหม่
        edit_so_button = CTkButton(action_buttons_frame, text="แก้ไข SO", width=100, command=self._edit_so_number, fg_color="#EAB308", hover_color="#CA8A04")
        edit_so_button.pack(side="left", padx=5)

        cancel_so_button = CTkButton(action_buttons_frame, text="ยกเลิก SO", width=100, command=self._cancel_so_record, fg_color="#DC2626", hover_color="#B91C1C")
        cancel_so_button.pack(side="left", padx=5)

        refresh_button = CTkButton(action_buttons_frame, text="🔄", width=35, command=self._refresh_so_list)
        refresh_button.pack(side="left", padx=5)
        # +++ END +++

        # --- ส่วนที่เหลือของฟังก์ชันเหมือนเดิม ---
        CTkLabel(top_frame, text="เอกสาร PO/ST:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        po_st_frame = CTkFrame(top_frame, fg_color="transparent")
        po_st_frame.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        po_st_frame.grid_columnconfigure(1, weight=1)
        self.po_number_type_var = tk.StringVar(value="PO"); self.po_number_type_var.trace_add("write", self._on_po_number_type_changed)
        self.po_type_dropdown = CTkOptionMenu(po_st_frame, variable=self.po_number_type_var, values=["PO", "ST"], width=80, **self.dropdown_style)
        self.po_type_dropdown.grid(row=0, column=0, sticky="w", padx=(0,5))
        self.po_number_input_var = tk.StringVar(); self.po_number_input_var.trace_add("write", self._validate_po_input)
        self.po_number_entry = CTkEntry(po_st_frame, font=self.entry_font, textvariable=self.po_number_input_var)
        self.po_number_entry.grid(row=0, column=1, sticky="ew")
        self.po_number_entry.bind("<FocusIn>", self._on_po_focus_in)
        self.po_number_entry.bind("<FocusOut>", self._check_px_on_po_entry)
        self.po_number_entry.bind("<Return>", self._check_px_on_po_entry)
        CTkLabel(top_frame, text="RR Number:").grid(row=1, column=2, sticky="w", padx=10, pady=5)
        self.rr_number_var.trace_add("write", self._force_uppercase_rr)
        self.rr_number_entry = CTkEntry(top_frame, font=self.entry_font, textvariable=self.rr_number_var)
        self.rr_number_entry.grid(row=1, column=3, sticky="ew", padx=10, pady=5)
        CTkLabel(top_frame, text="แผนก:").grid(row=1, column=4, sticky="w", padx=(20, 10), pady=5)
        self.department_entry = CTkEntry(top_frame, font=self.entry_font); self.department_entry.grid(row=1, column=5, sticky="ew", padx=10, pady=5)
        CTkLabel(top_frame, text="PUR Order :").grid(row=1, column=6, sticky="w", padx=(20, 10), pady=5)
        self.pur_order_entry = CTkEntry(top_frame, font=self.entry_font); self.pur_order_entry.grid(row=1, column=7, sticky="ew", padx=10, pady=5)
        sup_frame = CTkFrame(top_frame, fg_color="transparent"); sup_frame.grid(row=2, column=0, columnspan=8, sticky="ew", padx=5, pady=5)
        sup_frame.grid_columnconfigure(1, weight=4); sup_frame.grid_columnconfigure(3, weight=2); sup_frame.grid_columnconfigure(5, weight=2); sup_frame.grid_columnconfigure(6, weight=1); sup_frame.grid_columnconfigure(7, weight=1)
        CTkLabel(sup_frame, text="Supplier Name:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.supplier_name_combo = AutoCompleteEntry(master=sup_frame, completion_list=self.supplier_completion_data, display_key='name', command=self._on_supplier_selected, placeholder_text="พิมพ์เพื่อค้นหาซัพพลายเออร์...")
        self.supplier_name_combo.grid(row=0, column=1, sticky="ew", padx=(0,10), pady=3)
        CTkLabel(sup_frame, text="Supplier Code:").grid(row=0, column=2, sticky="w", padx=5, pady=3)
        self.supplier_code_entry = CTkEntry(sup_frame, font=self.entry_font); self.supplier_code_entry.grid(row=0, column=3, sticky="ew", padx=(0,10), pady=3)
        CTkLabel(sup_frame, text="Credit Term:").grid(row=0, column=4, sticky="w", padx=5, pady=3)
        self.credit_term_entry = CTkEntry(sup_frame, font=self.entry_font); self.credit_term_entry.grid(row=0, column=5, sticky="ew", padx=(0,10), pady=3)
        self.update_supplier_button = CTkButton(sup_frame, text="บันทึก/อัปเดต", width=120, command=self._save_or_update_supplier)
        self.update_supplier_button.grid(row=0, column=6, sticky="e", padx=(5, 10), pady=3)
        mode_frame = CTkFrame(sup_frame, fg_color="transparent"); mode_frame.grid(row=0, column=7, sticky="e", padx=(10, 5), pady=3)
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
    
    # ในไฟล์ purchasing_screen.py

    def _check_px_on_po_entry(self, event=None):
        """
        ฟังก์ชันตรวจสอบและ Sync ค่าขนส่ง (Trigger เมื่อกรอกเลข PO เสร็จ หรือกด Enter)
        """
        try:
            # 1. ดึงเลข PO
            po_num = self.po_number_input_var.get().strip().upper()
            
            # 2. ดึงเลข SO ปัจจุบัน
            current_so_string = self.so_entry.get()
            so_number = ""
            if "|" in current_so_string:
                so_number = current_so_string.split("|")[0].strip()
            else:
                so_number = current_so_string.strip()

            print(f"DEBUG: Triggering sync for PO: '{po_num}'")

            # --- [🔥 แก้ไข] เรียก 2 ฟังก์ชัน ---
            
            # 1. ดึงงบประมาณจาก SO (Commissions) เหมือนเดิม (เผื่อไม่มีใบรถจริง)
            if so_number:
                self.sync_transport_cost_to_po(so_number)
            
            # 2. ดึงค่ารถจริงจาก Transport Admin (Transport Orders) มาทับ ถ้ามี
            if po_num:
                self._sync_from_transport_orders(po_num)

        except Exception as e:
            print(f"Error in _check_px_on_po_entry: {e}")
            traceback.print_exc()
            
    def _sync_from_transport_orders(self, po_number):
        """
        ดึงข้อมูลค่ารถจริง (Actual Cost) จากหน้า Transport Admin มาใส่ใน PO
        แก้ไข: ดึงครบทุกอย่าง (Cost, Driver, Plate, VAT, WHT, Remark)
        """
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # [🔥 แก้ไข] เพิ่ม wht_percent และ remarks ใน query
                cursor.execute("""
                    SELECT transport_type, transport_cost, transporter_name, license_plate, 
                           vat_amount, wht_percent, remarks
                    FROM transport_orders 
                    WHERE ref_po_number = %s
                """, (po_number,))
                
                rows = cursor.fetchall()
                
                if not rows:
                    return 

                print(f"found {len(rows)} transport records for {po_number}")
                
                for row in rows:
                    t_type = row['transport_type'] # Stock หรือ Site
                    driver = row['transporter_name'] or ""
                    plate = row['license_plate'] or ""
                    cost = row['transport_cost'] or 0.0
                    vat_amt = row['vat_amount'] or 0.0
                    wht_pct = row['wht_percent'] or 0.0
                    remark = row['remarks'] or ""
                    
                    # 1. Logic VAT
                    vat_option = "VAT" if vat_amt > 0 else "CASH"

                    # 2. Logic WHT (แปลงตัวเลขเป็น Text ที่ RadioButton เข้าใจ)
                    wht_option = "ไม่มีหัก"
                    if wht_pct == 1.0: wht_option = "1%"
                    elif wht_pct == 3.0: wht_option = "3%"

                    if t_type == 'Stock':
                        # อัปเดตข้อมูล Text
                        utils.set_entry_text(self.shipping_to_stock_cost_entry, f"{cost:.2f}")
                        utils.set_entry_text(self.shipping_to_stock_driver_entry, driver)
                        utils.set_entry_text(self.shipping_to_stock_plate_entry, plate)
                        utils.set_entry_text(self.shipping_to_stock_notes_entry, remark) # [🔥 เพิ่มหมายเหตุ]
                        
                        # อัปเดตตัวเลือก Radio
                        self.shipping_to_stock_vat_var.set(vat_option)
                        self.shipping_to_stock_wht_var.set(wht_option) # [🔥 เพิ่ม WHT]
                        
                        self.shipping_to_stock_type_var.set("Aplus Logistic")
                        
                    elif t_type == 'Site':
                        # อัปเดตข้อมูล Text
                        utils.set_entry_text(self.shipping_to_site_cost_entry, f"{cost:.2f}")
                        utils.set_entry_text(self.shipping_to_site_driver_entry, driver)
                        utils.set_entry_text(self.shipping_to_site_plate_entry, plate)
                        utils.set_entry_text(self.shipping_to_site_notes_entry, remark) # [🔥 เพิ่มหมายเหตุ]
                        
                        # อัปเดตตัวเลือก Radio
                        self.shipping_to_site_vat_var.set(vat_option)
                        self.shipping_to_site_wht_var.set(wht_option) # [🔥 เพิ่ม WHT]
                        
                        self.shipping_to_site_type_var.set("Aplus Logistic")

                # คำนวณยอดรวมใหม่ทันที
                self._update_summary()
                
                messagebox.showinfo("Sync ข้อมูล", f"ดึงข้อมูลค่ารถจริงจาก Admin (พร้อม VAT/WHT/หมายเหตุ) สำหรับ {po_number} เรียบร้อยแล้ว", parent=self)

        except Exception as e:
            print(f"Error syncing from transport orders: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_supplier_selected(self, selection_dict):
        if not selection_dict:
            return

        self.editing_supplier_id = None
        
        # เคลียร์ช่องกรอกข้อมูลก่อน
        self.supplier_name_combo.delete(0, tk.END)
        self.supplier_code_entry.delete(0, tk.END)
        self.credit_term_entry.delete(0, tk.END)

        # นำข้อมูลจาก Dictionary มาใส่ในช่องต่างๆ
        self.editing_supplier_id = selection_dict.get('id')
        self.supplier_name_combo.insert(0, selection_dict.get('name', ''))
        self.supplier_code_entry.insert(0, selection_dict.get('code', ''))
        
        # แปลง credit term ให้อยู่ในรูปแบบที่อ่านง่าย
        credit_term_map = {'เงินสด': 'เงินสด', '0': 'เงินสด', '7': 'Cr 7', '15': 'Cr 15', '30': 'Cr 30'}
        term_value = str(selection_dict.get('term', 'เงินสด')).strip()
        self.credit_term_entry.insert(0, credit_term_map.get(term_value, term_value))

        # ======================================================================
        # [🔥 เพิ่มใหม่] Auto-fill ประเภทบัญชีลงในช่องชำระเงิน (Payments)
        # ======================================================================
        default_acc_type = selection_dict.get('bank_account_type', 'ออมทรัพย์')
        
        # วนลูปทุกช่องการจ่ายเงิน (Payment 1, 2, Full, CN) แล้วเซ็ตค่า Dropdown
        for p_type, widgets in self.payment_entries.items():
            if 'acc_type_var' in widgets:
                # ถ้าค่าใน DB เป็น None หรือว่าง ให้ใช้ 'ออมทรัพย์'
                if not default_acc_type: 
                    default_acc_type = 'ออมทรัพย์'
                
                widgets['acc_type_var'].set(default_acc_type)
            
    def _save_or_update_supplier(self):
        name, code, term = self.supplier_name_combo.get().strip(), self.supplier_code_entry.get().strip(), self.credit_term_entry.get().strip()
        if not name or not code: messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกชื่อและรหัสซัพพลายเออร์", parent=self); return
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT id FROM suppliers WHERE supplier_code = %s", (code,)); existing_by_code = cursor.fetchone()
                is_update = False
                
                # --- START: แก้ไข Logic การค้นหา ---
                # วนลูปเพื่อหาว่าชื่อที่กรอกมามีอยู่ในระบบแล้วหรือยัง
                for k, v in self.supplier_data_map.items():
                    # <<< แก้ไข: เปลี่ยนจาก v['supplier_name'] เป็น v['name'] ให้ตรงกับตอนโหลดข้อมูล >>>
                    if v['name'] == name:
                        self.editing_supplier_id = v['id']
                        is_update = True
                        break
                # --- END: สิ้นสุดการแก้ไข ---
                
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

        code_entry_widget = row_widgets.get("code")
        if code_entry_widget:
            code_entry_widget.delete(0, tk.END)
            code_entry_widget.insert(0, selection_dict.get('code', ''))
        
        code = selection_dict.get('code')
        product_data = self.product_data_map.get(code)
        
        if product_data:
            row_widgets['master_data'] = {
                "product_name": product_data.get("name", ""),
                "warehouse": product_data.get("warehouse", "")
            }
            
            row_widgets["name_var"].set(str(product_data.get("name") or ""))
            row_widgets["warehouse_var"].set(str(product_data.get("warehouse") or ""))
            
            # <<< START: เพิ่ม Logic การใส่ราคาและน้ำหนักอัตโนมัติ >>>
            last_price = product_data.get("last_price")
            last_weight = product_data.get("last_weight")

            price_entry = row_widgets.get("price")
            if price_entry and price_entry.winfo_exists():
                price_entry.delete(0, "end")
                if pd.notna(last_price):
                    price_entry.insert(0, f"{last_price:.2f}")

            weight_entry = row_widgets.get("weight")
            if weight_entry and weight_entry.winfo_exists():
                weight_entry.delete(0, "end")
                if pd.notna(last_weight):
                    weight_entry.insert(0, f"{last_weight:.2f}")
            # <<< END >>>
            
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
            # [🔥 แก้ไข] เพิ่ม bank_account_type ใน Query
            df = pd.read_sql("""
                SELECT id, supplier_name, supplier_code, credit_term, bank_account_type 
                FROM suppliers 
                ORDER BY supplier_name
            """, self.pg_engine)
            
            # เคลียร์ข้อมูลเก่า
            self.supplier_completion_data = []
            self.supplier_data_map = {} 
            
            # วนลูปเพื่อเตรียมข้อมูล
            for _, row in df.iterrows():
                name = str(row['supplier_name'] or '').strip()
                code = str(row['supplier_code'] or '').strip()
                
                if not name:
                    continue

                if code:
                    display_text = f"{name} ({code})" 
                else:
                    display_text = name

                item_data = {
                    "id": row['id'],
                    "name": name,
                    "code": code,
                    "term": row.get('credit_term', 'เงินสด'),
                    
                    # [🔥 เพิ่ม] เก็บประเภทบัญชีไว้ใช้ (ถ้าไม่มีให้เป็น 'ออมทรัพย์')
                    "bank_account_type": row.get('bank_account_type', 'ออมทรัพย์'), 
                    
                    "display": display_text 
                }

                self.supplier_completion_data.append(item_data)
                self.supplier_data_map[name] = item_data

            # อัปเดตข้อมูลใน widget AutoCompleteEntry
            if hasattr(self, 'supplier_name_combo') and self.supplier_name_combo.winfo_exists():
                self.supplier_name_combo.display_key = 'display'  
                self.supplier_name_combo.update_completion_list(self.supplier_completion_data)

            print(f"✅ โหลดข้อมูลซัพพลายเออร์สำเร็จ {len(self.supplier_completion_data)} รายการ")

        except Exception as e: 
            print(f"❌ ERROR: เกิดข้อผิดพลาดในการโหลดข้อมูลซัพพลายเออร์: {e}")
            self.supplier_completion_data = []
            self.supplier_data_map = {}
    
    def _load_product_master_data(self):
        try:
            query = "SELECT product_code, product_name, warehouse, last_unit_price, last_weight_per_unit FROM products ORDER BY product_code"
            df = pd.read_sql(query, self.pg_engine)
            
            self.product_completion_data = []
            MAX_NAME_LENGTH = 50
            for _, row in df.iterrows():
                name = row['product_name'] or ""
                display_name = name[:MAX_NAME_LENGTH] + '...' if len(name) > MAX_NAME_LENGTH else name
                display_text = f"{row['product_code']} - {display_name}"

                self.product_completion_data.append({
                    "name": name,
                    "code": row['product_code'],
                    "warehouse": row.get('warehouse', ''),
                    "display": display_text,
                    "last_price": row.get('last_unit_price'),
                    "last_weight": row.get('last_weight_per_unit')
                })

            self.product_data_map = {item['code']: item for item in self.product_completion_data}

            # --- START: แก้ไขการอัปเดตข้อมูลใน widget ---
            for row_dict in self.product_rows:
                if "code" in row_dict and isinstance(row_dict["code"], AutoCompleteEntry) and row_dict["code"].winfo_exists():
                    row_dict["code"].update_completion_list(self.product_completion_data)
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
        
        # +++ เพิ่ม "ลบ" ในรายการ headers +++
        headers = ["สถานะ", "รหัสสินค้า", "ชื่อสินค้า", "คลัง", "แก้ไข", "จำนวน", "ต้นทุนหน่วย (ไม่รวม VAT)", "ส่วนลด", "น้ำหนัก/หน่วย (กก.)", "น้ำหนักรวม (กก.)", "ต้นทุนรวม", "ลบ"]
        col_weights = [2, 4, 6, 2, 1, 2, 2, 3, 2, 2, 3, 1]  # เพิ่ม weight สำหรับคอลัมน์ลบ

        for i, h_text in enumerate(headers):
            self.products_frame.grid_columnconfigure(i, weight=col_weights[i])
            CTkLabel(self.products_frame, text=h_text, font=self.header_font_table, fg_color="#E0E0E0").grid(row=0, column=i, padx=1, pady=1, sticky="nsew")
        
        self.product_rows.clear()
        self._add_product_row()
        
        buttons_frame = CTkFrame(product_container, fg_color="transparent")
        buttons_frame.pack(anchor="e", pady=10, padx=10)
        CTkButton(buttons_frame, text="เพิ่มรายการสินค้า", command=self._add_product_row).pack(side="left", padx=5)
    
    def _delete_product_row_by_index(self, index):
        """ลบแถวสินค้าตาม index ที่ระบุ"""
        if len(self.product_rows) <= 1:
            messagebox.showwarning("ไม่สามารถลบได้", "ต้องมีรายการสินค้าอย่างน้อย 1 แถว", parent=self)
            return
        
        if 0 <= index < len(self.product_rows):
            # ลบ widgets ทั้งหมดในแถวนั้น
            row_to_delete = self.product_rows[index]
            for widget in row_to_delete["widgets"]:
                widget.destroy()
            
            # ลบออกจาก list
            self.product_rows.pop(index)
            
            # Re-arrange grid positions สำหรับแถวที่เหลือ
            self._rearrange_product_rows()
            
            # อัปเดตปุ่มลบทั้งหมด (เพราะ index เปลี่ยน)
            self._update_delete_button_commands()
            
            # อัปเดตสถานะปุ่มลบ
            self._update_delete_buttons_state()
            
            # คำนวณยอดใหม่
            self._update_summary()

    def _rearrange_product_rows(self):
        """จัดเรียง grid positions ของแถวสินค้าใหม่หลังจากลบ"""
        for idx, row_dict in enumerate(self.product_rows):
            row_num = idx + 1
            for col, widget in enumerate(row_dict["widgets"]):
                if widget.winfo_exists():
                    widget.grid(row=row_num, column=col, padx=1, pady=1, sticky="ew")

    def _update_delete_button_commands(self):
        """อัปเดต command ของปุ่มลบทั้งหมดให้ตรงกับ index ปัจจุบัน"""
        for idx, row_dict in enumerate(self.product_rows):
            if "delete_button" in row_dict and row_dict["delete_button"].winfo_exists():
                row_dict["delete_button"].configure(
                    command=lambda i=idx: self._delete_product_row_by_index(i)
                )

    def _update_delete_buttons_state(self):
        """อัปเดตสถานะของปุ่มลบ (ถ้ามีแถวเดียวให้ปิดการใช้งานปุ่มลบ)"""
        has_only_one_row = len(self.product_rows) <= 1
        
        for row_dict in self.product_rows:
            if "delete_button" in row_dict and row_dict["delete_button"].winfo_exists():
                if has_only_one_row:
                    row_dict["delete_button"].configure(state="disabled")
                else:
                    row_dict["delete_button"].configure(state="normal")

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
        
        product_code_entry = AutoCompleteEntry(
            self.products_frame, 
            completion_list=self.product_completion_data, 
            display_key='display',
            placeholder_text="Code"
        )

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
        
        # +++ START: เพิ่มปุ่มลบในแต่ละแถว +++
        delete_button = CTkButton(
            self.products_frame, 
            text="❌", 
            width=40,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=lambda: self._delete_product_row_by_index(row_num - 1)  # ส่ง index ของแถว
        )
        # +++ END +++
        
        widgets = [
            status_menu, product_code_entry, product_name_entry,
            warehouse_entry, warning_label, qty_entry, price_entry, 
            discount_frame, weight_entry, total_weight_entry, total_price_entry,
            delete_button  # เพิ่ม delete_button เข้าไปใน widgets list
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
            "warehouse_var": warehouse_var,
            "delete_button": delete_button  # เก็บ reference ของปุ่มลบ
        }
        
        self.product_rows.append(row_dict)
        
        product_code_entry.command = lambda selection_dict, r=row_dict: self._on_product_selected(selection_dict, r)

        product_name_var.trace_add("write", lambda *args, r=row_dict: self._check_for_override(r))
        warehouse_var.trace_add("write", lambda *args, r=row_dict: self._check_for_override(r))
        
        for entry in [qty_entry, weight_entry, price_entry, discount_value_entry]:
            entry.bind("<KeyRelease>", self._update_summary)
        discount_type_menu.configure(command=self._update_summary)
        
        # อัปเดตสถานะปุ่มลบ
        self._update_delete_buttons_state()

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

        # --- Section 1: Shipping to Stock (ค่าย้าย) ---
        CTkLabel(parent_frame, text="1.ค่าย้ายเข้าคลัง").grid(row=1, column=0, padx=10, pady=(5,0), sticky="w")
        CTkLabel(parent_frame, text="(ค่าย้าย )", font=CTkFont(size=11), text_color="#EF4444").grid(row=2, column=0, padx=10, pady=(0,5), sticky="nw")
        
        stock_cost_frame = CTkFrame(parent_frame, fg_color="transparent")
        stock_cost_frame.grid(row=1, column=1, rowspan=2, sticky="ew", padx=5, pady=2) 
        
        self.shipping_to_stock_cost_entry = NumericEntry(stock_cost_frame)
        self.shipping_to_stock_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.shipping_to_stock_cost_entry.bind("<KeyRelease>", self._update_summary)

        stock_vat_radio_frame = CTkFrame(stock_cost_frame, fg_color="transparent")
        stock_vat_radio_frame.pack(side="left")
        CTkRadioButton(stock_vat_radio_frame, text="VAT", variable=self.shipping_to_stock_vat_var, value="VAT").pack(side="left")
        CTkRadioButton(stock_vat_radio_frame, text="CASH", variable=self.shipping_to_stock_vat_var, value="CASH").pack(side="left", padx=5)
        self.shipping_to_stock_vat_var.trace_add("write", self._update_summary)

        # VAT 7%
        CTkLabel(parent_frame, text="VAT 7%:", font=self.entry_font).grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_stock_vat_display_entry = CTkEntry(parent_frame, textvariable=self.shipping_to_stock_vat_display_var, state="readonly", fg_color="gray85")
        self.shipping_to_stock_vat_display_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        # WHT
        CTkLabel(parent_frame, text="หัก ณ ที่จ่าย:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_stock_wht_var.trace_add("write", self._update_summary)
        stock_wht_frame = CTkFrame(parent_frame, fg_color="transparent")
        stock_wht_frame.grid(row=4, column=1, sticky="ew", padx=5, pady=2)
        CTkRadioButton(stock_wht_frame, text="ไม่มีหัก", variable=self.shipping_to_stock_wht_var, value="ไม่มีหัก").pack(side="left", padx=(0,5))
        CTkRadioButton(stock_wht_frame, text="1%", variable=self.shipping_to_stock_wht_var, value="1%").pack(side="left", padx=5)
        CTkRadioButton(stock_wht_frame, text="3%", variable=self.shipping_to_stock_wht_var, value="3%").pack(side="left", padx=5)

        # WHT Amount
        CTkLabel(parent_frame, text="ยอดหัก ณ ที่จ่าย:", font=self.entry_font).grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_stock_wht_display_entry = CTkEntry(parent_frame, textvariable=self.shipping_to_stock_wht_display_var, state="readonly", fg_color="gray85")
        self.shipping_to_stock_wht_display_entry.grid(row=5, column=1, sticky="ew", padx=5, pady=2)

        # Date
        self.shipping_to_stock_date_selector = DateSelector(parent_frame, dropdown_style=self.dropdown_style)
        self.shipping_to_stock_date_selector.grid(row=6, column=1, sticky="w", padx=5, pady=2)
        
        # Shipper Type
        self.shipping_to_stock_type_var = tk.StringVar(value="Aplus Logistic")
        stock_shipper_radio_frame = CTkFrame(parent_frame, fg_color="transparent")
        stock_shipper_radio_frame.grid(row=7, column=1, sticky="w", padx=5, pady=2)
        CTkRadioButton(stock_shipper_radio_frame, text="ซัพพลายเออร์จัดส่ง", variable=self.shipping_to_stock_type_var, value="ซัพพลายเออร์จัดส่ง", command=self._update_summary).pack(side="left")
        CTkRadioButton(stock_shipper_radio_frame, text="Aplus Logistic", variable=self.shipping_to_stock_type_var, value="Aplus Logistic", command=self._update_summary).pack(side="left", padx=5)
        CTkRadioButton(stock_shipper_radio_frame, text="Lalamove/Others", variable=self.shipping_to_stock_type_var, value="Lalamove/Others", command=self._update_summary).pack(side="left", padx=5)
        
        # [🔥 เพิ่ม] Driver & Plate Fields (Stock)
        stock_driver_frame = CTkFrame(parent_frame, fg_color="transparent")
        stock_driver_frame.grid(row=8, column=1, sticky="ew", padx=5, pady=2)
        
        self.shipping_to_stock_driver_entry = CTkEntry(stock_driver_frame, placeholder_text="ชื่อคนขับ / บริษัทขนส่ง")
        self.shipping_to_stock_driver_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.shipping_to_stock_plate_entry = CTkEntry(stock_driver_frame, placeholder_text="ทะเบียนรถ", width=120)
        self.shipping_to_stock_plate_entry.pack(side="left")

        # Notes
        self.shipping_to_stock_notes_entry = CTkEntry(parent_frame, placeholder_text="หมายเหตุเพิ่มเติม...")
        self.shipping_to_stock_notes_entry.grid(row=9, column=1, sticky="ew", padx=5, pady=2)

        CTkFrame(parent_frame, height=2, fg_color="gray90").grid(row=10, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

        # --- Section 2: Shipping to Site (ค่าขนส่ง) ---
        CTkLabel(parent_frame, text="2.ค่าจัดส่งลูกค้า").grid(row=11, column=0, padx=10, pady=(5,0), sticky="w")
        CTkLabel(parent_frame, text="(ค่ารถ)", font=CTkFont(size=11), text_color="#EF4444").grid(row=12, column=0, padx=10, pady=(0,5), sticky="nw")

        site_cost_frame = CTkFrame(parent_frame, fg_color="transparent")
        site_cost_frame.grid(row=11, column=1, rowspan=2, sticky="ew", padx=5, pady=2)

        self.shipping_to_site_cost_entry = NumericEntry(site_cost_frame)
        self.shipping_to_site_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.shipping_to_site_cost_entry.bind("<KeyRelease>", self._update_summary)

        site_vat_radio_frame = CTkFrame(site_cost_frame, fg_color="transparent")
        site_vat_radio_frame.pack(side="left")
        CTkRadioButton(site_vat_radio_frame, text="VAT", variable=self.shipping_to_site_vat_var, value="VAT").pack(side="left")
        CTkRadioButton(site_vat_radio_frame, text="CASH", variable=self.shipping_to_site_vat_var, value="CASH").pack(side="left", padx=5)
        self.shipping_to_site_vat_var.trace_add("write", self._update_summary)

        # VAT 7%
        CTkLabel(parent_frame, text="VAT 7%:", font=self.entry_font).grid(row=13, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_site_vat_display_entry = CTkEntry(parent_frame, textvariable=self.shipping_to_site_vat_display_var, state="readonly", fg_color="gray85")
        self.shipping_to_site_vat_display_entry.grid(row=13, column=1, sticky="ew", padx=5, pady=2)

        # WHT
        CTkLabel(parent_frame, text="หัก ณ ที่จ่าย:").grid(row=14, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_site_wht_var.trace_add("write", self._update_summary)
        site_wht_frame = CTkFrame(parent_frame, fg_color="transparent")
        site_wht_frame.grid(row=14, column=1, sticky="ew", padx=5, pady=2)
        CTkRadioButton(site_wht_frame, text="ไม่มีหัก", variable=self.shipping_to_site_wht_var, value="ไม่มีหัก").pack(side="left", padx=(0,5))
        CTkRadioButton(site_wht_frame, text="1%", variable=self.shipping_to_site_wht_var, value="1%").pack(side="left", padx=5)
        CTkRadioButton(site_wht_frame, text="3%", variable=self.shipping_to_site_wht_var, value="3%").pack(side="left", padx=5)

        # WHT Amount
        CTkLabel(parent_frame, text="ยอดหัก ณ ที่จ่าย:", font=self.entry_font).grid(row=15, column=0, padx=10, pady=5, sticky="w")
        self.shipping_to_site_wht_display_entry = CTkEntry(parent_frame, textvariable=self.shipping_to_site_wht_display_var, state="readonly", fg_color="gray85")
        self.shipping_to_site_wht_display_entry.grid(row=15, column=1, sticky="ew", padx=5, pady=2)

        # Date
        self.shipping_to_site_date_selector = DateSelector(parent_frame, dropdown_style=self.dropdown_style)
        self.shipping_to_site_date_selector.grid(row=16, column=1, sticky="w", padx=5, pady=2)
        
        # Shipper Type
        self.shipping_to_site_type_var = tk.StringVar(value="Aplus Logistic")
        site_shipper_radio_frame = CTkFrame(parent_frame, fg_color="transparent")
        site_shipper_radio_frame.grid(row=17, column=1, sticky="w", padx=5, pady=2)
        CTkRadioButton(site_shipper_radio_frame, text="ซัพพลายเออร์จัดส่ง", variable=self.shipping_to_site_type_var, value="ซัพพลายเออร์จัดส่ง", command=self._update_summary).pack(side="left")
        CTkRadioButton(site_shipper_radio_frame, text="Aplus Logistic", variable=self.shipping_to_site_type_var, value="Aplus Logistic", command=self._update_summary).pack(side="left", padx=5)
        CTkRadioButton(site_shipper_radio_frame, text="Lalamove/Others", variable=self.shipping_to_site_type_var, value="Lalamove/Others", command=self._update_summary).pack(side="left", padx=5)

        # [🔥 เพิ่ม] Driver & Plate Fields (Site)
        site_driver_frame = CTkFrame(parent_frame, fg_color="transparent")
        site_driver_frame.grid(row=18, column=1, sticky="ew", padx=5, pady=2)
        
        self.shipping_to_site_driver_entry = CTkEntry(site_driver_frame, placeholder_text="ชื่อคนขับ / บริษัทขนส่ง")
        self.shipping_to_site_driver_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.shipping_to_site_plate_entry = CTkEntry(site_driver_frame, placeholder_text="ทะเบียนรถ", width=120)
        self.shipping_to_site_plate_entry.pack(side="left")

        # Notes
        self.shipping_to_site_notes_entry = CTkEntry(parent_frame, placeholder_text="หมายเหตุเพิ่มเติม...")
        self.shipping_to_site_notes_entry.grid(row=19, column=1, sticky="ew", padx=5, pady=2)
        
        CTkFrame(parent_frame, height=2, fg_color="gray90").grid(row=20, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

    def _populate_payment_column(self, parent_frame):
        parent_frame.grid_columnconfigure(1, weight=1)
        self.payment_entries.clear()

        self.bank_list = [
            "ระบุเอง", "BBL", "KBANK", "KTB", "SCB", "TTB", "BAY", "GSB", "BAAC", "UOB", "CIMB"
        ]

        CTkLabel(parent_frame, text="การชำระซัพพลายเออร์", font=self.header_font_table).grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        # --- Helper Function ---
        def create_payment_entry_frame(label_text, row_index, payment_type, has_percent_dropdown=False):
            p_frame = CTkFrame(parent_frame, fg_color="transparent")
            p_frame.grid(row=row_index, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
            p_frame.grid_columnconfigure(1, weight=1)

            # Row 0: Label & Amount
            CTkLabel(p_frame, text=label_text, font=self.label_font).grid(row=0, column=0, sticky="w", pady=2)
            amount_frame = CTkFrame(p_frame, fg_color="transparent")
            amount_frame.grid(row=0, column=1, sticky="ew")
            
            percent_var = None
            if has_percent_dropdown:
                percent_var = tk.StringVar(value="ระบุยอดเอง")
                p_percent = CTkOptionMenu(amount_frame, variable=percent_var, values=["ระบุยอดเอง", "30%", "50%", "100%"], width=120)
                p_percent.pack(side="left", padx=(0, 5))
            
            p_amount = NumericEntry(amount_frame)
            p_amount.pack(side="left", fill="x", expand=True)
            
            # Row 1: Bank & Account
            CTkLabel(p_frame, text="ธนาคาร:").grid(row=1, column=0, sticky="w", pady=2)
            bank_frame = CTkFrame(p_frame, fg_color="transparent")
            bank_frame.grid(row=1, column=1, sticky="ew")
            bank_frame.grid_columnconfigure(0, weight=1); bank_frame.grid_columnconfigure(1, weight=1)

            p_bank = CTkOptionMenu(bank_frame, values=self.bank_list, **self.dropdown_style)
            p_bank.grid(row=0, column=0, sticky="ew", padx=(0, 5))

            p_account = CTkEntry(bank_frame, placeholder_text="เลขที่บัญชี...")
            p_account.grid(row=0, column=1, sticky="ew")
            
            # ==================================================================
            # [🔥 เพิ่มใหม่] Row 2: ประเภทบัญชี (Dropdown)
            # ==================================================================
            CTkLabel(p_frame, text="ประเภทบัญชี:").grid(row=2, column=0, sticky="w", pady=2)
            
            acc_type_var = tk.StringVar(value="ออมทรัพย์")
            p_acc_type = CTkOptionMenu(
                p_frame, 
                variable=acc_type_var, 
                values=["ออมทรัพย์", "กระแสรายวัน"],
                **self.dropdown_style
            )
            p_acc_type.grid(row=2, column=1, sticky="w", pady=2) # sticky="w" เพื่อให้อยู่ซ้าย
            # ==================================================================

            # Row 3: Date (เลื่อนลงมาจาก Row 2)
            p_date = None
            if payment_type in ["Payment 1", "Payment 2"]:
                CTkLabel(p_frame, text="วันที่ชำระ:").grid(row=3, column=0, sticky="w", pady=2)
                p_date = DateSelector(p_frame, dropdown_style=self.dropdown_style)
                p_date.grid(row=3, column=1, sticky="ew")

            # Bind Events
            p_amount.bind("<KeyRelease>", self._update_summary)
            if has_percent_dropdown:
                p_percent.configure(command=lambda val, pv=percent_var, pa=p_amount: self._calculate_payment_from_percentage(val, pv, pa))

            # [สำคัญ] ส่ง acc_type_var กลับไปด้วย
            return p_amount, p_date, percent_var, p_bank, p_account, acc_type_var 

        # --- สร้าง Widget โดยรับตัวแปรเพิ่ม ---
        p1_amount, p1_date, self.payment1_percent_var, p1_bank, p1_account, p1_acc_type = create_payment_entry_frame("1.มัดจำ:", 1, payment_type="Payment 1", has_percent_dropdown=True)
        p2_amount, p2_date, self.payment2_percent_var, p2_bank, p2_account, p2_acc_type = create_payment_entry_frame("2.มัดจำ:", 2, payment_type="Payment 2", has_percent_dropdown=True)

        CTkLabel(parent_frame, text="ยอดรวมมัดจำ:", font=self.label_font).grid(row=3, column=0, padx=5, pady=8, sticky="w")
        total_deposit_entry = CTkEntry(parent_frame, textvariable=self.total_deposit_var, state="readonly", fg_color="gray85")
        total_deposit_entry.grid(row=3, column=1, sticky="ew", pady=8, padx=5)

        CTkLabel(parent_frame, text="ยอดค้าง:", font=self.label_font).grid(row=4, column=0, padx=5, pady=8, sticky="w")
        self.balance_due_entry = CTkEntry(parent_frame, textvariable=self.balance_due_var, state="readonly", fg_color="gray85")
        self.balance_due_entry.grid(row=4, column=1, sticky="ew", pady=8, padx=5)
        
        fp_amount, fp_date, _, fp_bank, fp_account, fp_acc_type = create_payment_entry_frame("ชำระเต็ม:", 5, payment_type="Full Payment")
        cn_amount, cn_date, _, cn_bank, cn_account, cn_acc_type = create_payment_entry_frame("CN/คืนส่วนลด:", 6, payment_type="CN Refund")

        # [🔥 สำคัญ] เก็บ acc_type_var ลงใน Dictionary เพื่อให้ฟังก์ชันอื่นดึงไปใช้ได้
        self.payment_entries["Payment 1"] = {
            "amount": p1_amount, "date": p1_date, "percent_var": self.payment1_percent_var, 
            "bank_menu": p1_bank, "account_entry": p1_account, "acc_type_var": p1_acc_type
        }
        self.payment_entries["Payment 2"] = {
            "amount": p2_amount, "date": p2_date, "percent_var": self.payment2_percent_var, 
            "bank_menu": p2_bank, "account_entry": p2_account, "acc_type_var": p2_acc_type
        }
        self.payment_entries["Full Payment"] = {
            "amount": fp_amount, "date": fp_date, 
            "bank_menu": fp_bank, "account_entry": fp_account, "acc_type_var": fp_acc_type
        }
        self.payment_entries["CN Refund"] = {
            "amount": cn_amount, "date": cn_date, 
            "bank_menu": cn_bank, "account_entry": cn_account, "acc_type_var": cn_acc_type
        }
        # --- END: สิ้นสุดการเรียกใช้ Helper ---
    
    def _populate_summary_column(self, parent_frame):
        # ==============================================================================
        # [ส่วนที่ 1] ค่าบริการตัด/เจาะ (ย้ายมาไว้บนสุด ก่อนคำว่าสรุปต้นทุน)
        # ==============================================================================
        
        # 1.1 ช่องกรอกราคา + ตัวเลือก VAT/CASH
        CTkLabel(parent_frame, text="ค่าบริการตัด/เจาะ:").grid(row=0, column=0, padx=10, pady=2, sticky="w")
        
        cutting_cost_frame = CTkFrame(parent_frame, fg_color="transparent")
        cutting_cost_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        self.cutting_cost_entry = NumericEntry(cutting_cost_frame)
        self.cutting_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.cutting_cost_entry.bind("<KeyRelease>", self._update_summary)

        cutting_vat_radio_frame = CTkFrame(cutting_cost_frame, fg_color="transparent")
        cutting_vat_radio_frame.pack(side="left")
        CTkRadioButton(cutting_vat_radio_frame, text="VAT", variable=self.cutting_vat_var, value="VAT").pack(side="left")
        CTkRadioButton(cutting_vat_radio_frame, text="CASH", variable=self.cutting_vat_var, value="CASH").pack(side="left", padx=5)
        self.cutting_vat_var.trace_add("write", self._update_summary)

        # 1.2 ช่องแสดงยอด VAT
        CTkLabel(parent_frame, text="VAT 7%:").grid(row=1, column=0, padx=10, pady=2, sticky="w")
        self.cutting_vat_display_entry = CTkEntry(parent_frame, textvariable=self.cutting_vat_display_var, state="readonly", fg_color="gray85")
        self.cutting_vat_display_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # 1.3 ตัวเลือกหัก ณ ที่จ่าย
        CTkLabel(parent_frame, text="หัก ณ ที่จ่าย:").grid(row=2, column=0, padx=10, pady=2, sticky="w")
        
        cutting_wht_frame = CTkFrame(parent_frame, fg_color="transparent")
        cutting_wht_frame.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        CTkRadioButton(cutting_wht_frame, text="ไม่มีหัก", variable=self.cutting_wht_var, value="No").pack(side="left")
        CTkRadioButton(cutting_wht_frame, text="1%", variable=self.cutting_wht_var, value="1%").pack(side="left", padx=5)
        CTkRadioButton(cutting_wht_frame, text="3%", variable=self.cutting_wht_var, value="3%").pack(side="left", padx=5)
        self.cutting_wht_var.trace_add("write", self._update_summary)

        # 1.4 ช่องแสดงยอดหัก ณ ที่จ่าย
        CTkLabel(parent_frame, text="ยอดหัก ณ ที่จ่าย:").grid(row=3, column=0, padx=10, pady=2, sticky="w")
        self.cutting_wht_display_entry = CTkEntry(parent_frame, textvariable=self.cutting_wht_display_var, state="readonly", fg_color="gray85")
        self.cutting_wht_display_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        # 1.5 ช่องรวมค่าบริการ
        CTkLabel(parent_frame, text="รวมค่าบริการ:").grid(row=4, column=0, padx=10, pady=2, sticky="w")
        self.cutting_total_display_entry = CTkEntry(parent_frame, textvariable=self.cutting_total_display_var, state="readonly", fg_color="#F3E8FF", font=CTkFont(weight="bold"))
        self.cutting_total_display_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=2)

        # 1.6 ช่องหมายเหตุ
        self.cutting_remark_entry = CTkEntry(parent_frame, placeholder_text="หมายเหตุตัด/เจาะ...")
        self.cutting_remark_entry.grid(row=5, column=1, sticky="ew", padx=5, pady=5)

        # --- เส้นคั่นสวยงาม เพื่อแยกส่วน ---
        CTkFrame(parent_frame, height=2, fg_color="gray90").grid(row=6, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

        # ==============================================================================
        # [ส่วนที่ 2] สรุปต้นทุน PO (เริ่มที่ Row 7)
        # ==============================================================================
        
        CTkLabel(parent_frame, text="สรุปต้นทุน", font=self.header_font_table).grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        # กำหนด row เริ่มต้นสำหรับส่วนถัดไป (นับต่อจากด้านบน)
        row_idx = 8 

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
        self.vat_checkbox.select()
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
        self.save_draft_button = CTkButton(footer, text="💾 บันทึกฉบับร่าง (Save Draft)", command=lambda: self._save_po('Draft'), **btn_config) # <--- เพิ่ม self.save_draft_button
        self.save_draft_button.pack(side="left", padx=5, expand=True, fill="x")
        
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
            # --- START: แก้ไข Query ---
            # เพิ่ม cutting_cost และ field ที่เกี่ยวข้อง
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
                    po.total_cost,
                    
                    -- Shipping info
                    po.shipping_to_stock_cost,
                    po.shipping_to_site_cost,
                    po.shipping_to_stock_shipper,
                    po.shipping_to_site_shipper,
                    po.shipping_to_stock_wht_type,
                    po.shipping_to_site_wht_type,
                    
                    -- [🔥 NEW] Cutting Info
                    po.cutting_cost,
                    po.cutting_vat_type,
                    po.cutting_vat_amount,
                    po.cutting_wht_type,
                    po.cutting_wht_amount,
                    po.cutting_remark,
                    
                    -- Fields from commissions (c)
                    c.so_number,
                    c.bill_date,
                    c.commission_month,
                    c.commission_year,
                    c.customer_name,
                    c.credit_term,
                    c.sales_service_amount,
                    c.credit_card_fee,
                    c.cutting_drilling_fee, -- อันนี้ของ SO
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
            # --- END: แก้ไข Query ---

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
            
            # เรียก PDF Generator (ต้องแน่ใจว่า function นี้รองรับ field ใหม่แล้ว หรือเดี๋ยวค่อยไปแก้ไฟล์นั้นอีกที)
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
            'so_number': self.so_entry.get().split('|')[0].strip() if '|' in self.so_entry.get() else self.so_entry.get().strip(),
            'po_number': self.po_number_input_var.get(),
            'rr_number': self.rr_number_var.get(),
            'department': self.department_entry.get().strip(),
            'pur_order': self.pur_order_entry.get().strip(),
            'supplier_name': self.supplier_name_combo.get(),
            'supplier_code': self.supplier_code_entry.get(),
            'credit_term': self.credit_term_entry.get(),
            'po_mode': self.po_mode_var.get(), 
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            
            # --- Stock ---
            'shipping_to_stock_cost': utils.convert_to_float(self.shipping_to_stock_cost_entry.get()),
            'shipping_to_stock_vat_type': self.shipping_to_stock_vat_var.get(),
            'shipping_to_stock_wht_type': self.shipping_to_stock_wht_var.get(),
            'shipping_to_stock_wht_amount': utils.convert_to_float(self.shipping_to_stock_wht_display_var.get()),
            'shipping_to_stock_date': self.shipping_to_stock_date_selector.get_date(),
            'shipping_to_stock_shipper': self.shipping_to_stock_type_var.get(),
            'shipping_to_stock_driver': self.shipping_to_stock_driver_entry.get(), # [🔥 เพิ่ม]
            'shipping_to_stock_plate': self.shipping_to_stock_plate_entry.get(),   # [🔥 เพิ่ม]
            'shipping_to_stock_notes': self.shipping_to_stock_notes_entry.get(),
            
            # --- Site ---
            'shipping_to_site_cost': utils.convert_to_float(self.shipping_to_site_cost_entry.get()),
            'shipping_to_site_vat_type': self.shipping_to_site_vat_var.get(),
            'shipping_to_site_wht_type': self.shipping_to_site_wht_var.get(),
            'shipping_to_site_wht_amount': utils.convert_to_float(self.shipping_to_site_wht_display_var.get()),
            'shipping_to_site_date': self.shipping_to_site_date_selector.get_date(),
            'shipping_to_site_shipper': self.shipping_to_site_type_var.get(),
            'shipping_to_site_driver': self.shipping_to_site_driver_entry.get(),   # [🔥 เพิ่ม]
            'shipping_to_site_plate': self.shipping_to_site_plate_entry.get(),     # [🔥 เพิ่ม]
            'shipping_to_site_notes': self.shipping_to_site_notes_entry.get(),
            
            # --- Cutting ---
            'cutting_cost': utils.convert_to_float(self.cutting_cost_entry.get()),
            'cutting_vat_type': self.cutting_vat_var.get(),
            'cutting_vat_amount': utils.convert_to_float(self.cutting_vat_display_var.get()),
            'cutting_wht_type': self.cutting_wht_var.get(),
            'cutting_wht_amount': utils.convert_to_float(self.cutting_wht_display_var.get()),
            'cutting_remark': self.cutting_remark_entry.get(),
            
            # --- Totals ---
            'total_cost': utils.convert_to_float(self.total_cost_entry.get()),
            'total_weight': utils.convert_to_float(self.total_weight_summary_entry.get()),
            'bill_discount': utils.convert_to_float(self.end_of_bill_discount_entry.get()),
            'wht_3_percent_checked': bool(self.vat3_checkbox.get()),
            'wht_3_percent_amount': utils.convert_to_float(self.vat3_entry.get()),
            'vat_7_percent_checked': bool(self.vat_checkbox.get()),
            'vat_7_percent_amount': utils.convert_to_float(self.vat7_entry.get()),
            'grand_total': utils.convert_to_float(self.grand_total_payable_entry.get())
        }

        if hasattr(self, 'proxy_user_key') and self.proxy_user_key:
            header_data['user_key'] = self.user_key
            header_data['proxy_user_key'] = self.proxy_user_key
        else:
            header_data['user_key'] = self.user_key
            header_data['proxy_user_key'] = None
        
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
                payment_date = p_widgets["date"].get_date() if p_widgets.get("date") else datetime.now()
                payments_data.append({
                    "payment_type": p_type,
                    "amount": amount,
                    "payment_date": payment_date,
                    "bank_name": p_widgets["bank_menu"].get(),
                    "bank_account_number": p_widgets["account_entry"].get(),
                    
                    # [🔥 เพิ่ม] ดึงค่าประเภทบัญชี
                    "bank_account_type": p_widgets["acc_type_var"].get() 
                })
        
        return {"header": header_data, "items": items_data, "payments": payments_data}


    def _save_po(self, status):
        """(ฉบับแก้ไข) บันทึก PO พร้อมบันทึกประเภทบัญชีธนาคาร"""
        
        form_data = self._gather_form_data()
        header, items, payments = form_data.get('header', {}), form_data.get('items', []), form_data.get('payments', [])

        if not header.get("so_number"):
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณาเลือก SO Number ก่อนทำการบันทึก", parent=self)
            return
        if not header.get("supplier_name"):
             messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก Supplier ก่อนบันทึก", parent=self)
             return
        if not items:
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณาเพิ่มสินค้าอย่างน้อย 1 รายการก่อนบันทึก", parent=self)
            return
        if status == 'Pending Approval' and not header.get("po_number"):
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณากรอก PO/ST Number ก่อนส่งอนุมัติ", parent=self)
            return

        loading_label = show_loading_message(self, f"กำลังบันทึก PO สถานะ '{status}'...")
        
        save_buttons = self._find_save_buttons(); self._disable_buttons(save_buttons)
    
        def save_work():
            conn = self.app_container.get_connection()
            try:
                with conn.cursor() as cursor:
                    # 1. ตรวจสอบ PO ซ้ำ (เฉพาะสร้างใหม่)
                    if not self.editing_po_id:
                        cursor.execute("""
                            SELECT id FROM purchase_orders 
                            WHERE so_number = %s AND supplier_name = %s
                        """, (header.get("so_number"), header.get("supplier_name")))
                        
                        if cursor.fetchone():
                            raise ValueError(f"มี PO สำหรับ SO '{header.get('so_number')}' และ Supplier '{header.get('supplier_name')}' นี้อยู่แล้ว")

                    # 2. เตรียมข้อมูล Header
                    if status == 'Pending Approval':
                        header['status'] = 'Pending Approval'; header['approval_status'] = 'Pending Mgr 1'
                    else:
                        header['status'] = 'Draft'; header['approval_status'] = 'Draft'

                    cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'purchase_orders'")
                    db_columns = {row[0] for row in cursor.fetchall()}
                    
                    if self.editing_po_id:
                        # Update
                        header.pop('user_key', None); header.pop('timestamp', None)
                        filtered_header = {k: v for k, v in header.items() if k in db_columns}
                        set_clause_formatted = ", ".join([f'"{k}" = %s' for k in filtered_header.keys()])
                        params = list(filtered_header.values()) + [self.editing_po_id]
                        cursor.execute(f"UPDATE purchase_orders SET {set_clause_formatted} WHERE id = %s", params)
                        new_po_id = self.editing_po_id
                    else:
                        # Insert
                        filtered_header = {k: v for k, v in header.items() if k in db_columns}
                        cols = ", ".join([f'"{k}"' for k in filtered_header.keys()])
                        placeholders = ", ".join(["%s"] * len(filtered_header))
                        cursor.execute(f"INSERT INTO purchase_orders ({cols}) VALUES ({placeholders}) RETURNING id", list(filtered_header.values()))
                        new_po_id = cursor.fetchone()[0]
                    
                    # 3. ลบข้อมูลเก่า (Items & Payments) แล้วลงใหม่
                    cursor.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = %s", (new_po_id,))
                    cursor.execute("DELETE FROM purchase_order_payments WHERE purchase_order_id = %s", (new_po_id,))
                    
                    # 4. Insert Items
                    if items:
                        items_values = [(new_po_id, item['product_name'], item['status'], item['product_code'], item['warehouse'], item['quantity'], item['weight_per_unit'], item['unit_price'], item['discount_value'], item['discount_type'], item['total_weight'], item['total_price']) for item in items]
                        psycopg2.extras.execute_values(cursor, "INSERT INTO purchase_order_items (purchase_order_id, product_name, status, product_code, warehouse, quantity, weight_per_unit, unit_price, discount_value, discount_type, total_weight, total_price) VALUES %s", items_values)

                    # 5. Insert Payments (รวม bank_account_type)
                    if payments:
                        # [🔥 แก้ไข] เพิ่ม p.get('bank_account_type') ใน values และ query
                        payments_values = [(
                            new_po_id, 
                            payment['payment_type'], 
                            payment['amount'], 
                            payment['payment_date'], 
                            payment['bank_name'], 
                            payment['bank_account_number'],
                            payment.get('bank_account_type', 'ออมทรัพย์') # เพิ่มตรงนี้
                        ) for payment in payments]
                        
                        psycopg2.extras.execute_values(
                            cursor, 
                            """INSERT INTO purchase_order_payments 
                               (purchase_order_id, payment_type, amount, payment_date, bank_name, bank_account_number, bank_account_type) 
                               VALUES %s""", 
                            payments_values
                        )

                    # 6. Update Commission Status (ถ้าไม่ใช่ Draft)
                    if status != 'Draft':
                        so_number_to_update = header.get("so_number")
                        if so_number_to_update:
                            cursor.execute("UPDATE commissions SET status = 'PO Sent' WHERE so_number = %s AND is_active = 1", (so_number_to_update,))
                            print(f"Updated commissions status to 'PO Sent' for SO: {so_number_to_update}")
                    
                    # 7. Notification (ถ้าส่งอนุมัติ)
                    if status == 'Pending Approval':
                        self._create_initial_approval_notification(cursor, new_po_id)
                
                conn.commit()
                
                # 8. Update Product Master Price/Weight (Optional)
                try:
                    with conn.cursor() as cursor:
                        price_updates = [(item['unit_price'], item['weight_per_unit'], item['product_code']) for item in items if item.get('product_code') and item.get('unit_price') is not None]
                        if price_updates:
                            psycopg2.extras.execute_values(cursor, "UPDATE products SET last_unit_price = data.price, last_weight_per_unit = data.weight, last_updated = NOW() FROM (VALUES %s) AS data (price, weight, code) WHERE product_code = data.code", price_updates)
                            conn.commit()
                except Exception as update_err:
                    print(f"Could not update last price/weight: {update_err}")

                return {"success": True, "po_id": new_po_id, "status": status}

            except Exception as e:
                if conn: conn.rollback()
                raise e
            finally:
                if conn: self.app_container.release_connection(conn)
    
        def on_success(result):
            hide_loading_message(loading_label)
            
            self._enable_buttons(save_buttons)
            messagebox.showinfo("สำเร็จ", f"บันทึก PO เป็น '{result['status']}' สำเร็จ", parent=self)
            try: self._load_product_master_data()
            except: pass
            
            if header.get('po_mode') == 'Multiple-PO':
                self._clear_form(keep_so=True, confirm=False)
            else:
                self._clear_form(keep_so=False, confirm=False)
    
        def on_error(error):
            hide_loading_message(loading_label)

            self._enable_buttons(save_buttons)
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกได้: {str(error)}", parent=self)
            traceback.print_exc()
        
        self.async_helper.run_in_background(save_work, on_success, on_error)
    
        def on_success(result):
            # <<< จุดที่แก้ไข: เรียกใช้ฟังก์ชันที่ถูกต้อง >>>
            hide_loading_message(loading_label)
            
            self._enable_buttons(save_buttons)
            messagebox.showinfo("สำเร็จ", f"บันทึก PO เป็น '{result['status']}' สำเร็จ", parent=self)
            try: self._load_product_master_data()
            except: pass
            
            if header.get('po_mode') == 'Multiple-PO':
                self._clear_form(keep_so=True, confirm=False)
            else:
                self._clear_form(keep_so=False, confirm=False)
    
        def on_error(error):
            # <<< จุดที่แก้ไข: เรียกใช้ฟังก์ชันที่ถูกต้อง >>>
            hide_loading_message(loading_label)

            self._enable_buttons(save_buttons)
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกได้: {str(error)}", parent=self)
            traceback.print_exc()
        
        self.async_helper.run_in_background(save_work, on_success, on_error)
        
    def _find_save_buttons(self):
        """หาปุ่มบันทึกทั้งหมดในฟอร์ม"""
        save_buttons = []
        # หาปุ่มที่มีคำว่า 'บันทึก' หรือ 'Save' ใน text
        for child in self.winfo_children():
            self._find_buttons_recursive(child, save_buttons, ["บันทึก", "Save", "ขออนุมัติ"])
        return save_buttons

    def _find_buttons_recursive(self, widget, button_list, search_texts):
        """หาปุ่มแบบ recursive"""
        try:
            # ตรวจสอบว่าเป็นปุ่มหรือไม่
            if hasattr(widget, 'cget') and hasattr(widget, 'configure'):
                try:
                    text = widget.cget('text')
                    if text and any(search_text in text for search_text in search_texts):
                        button_list.append(widget)
                except:
                    pass
            
            # ตรวจสอบ children
            for child in widget.winfo_children():
                self._find_buttons_recursive(child, button_list, search_texts)
        except:
            pass

    def _disable_buttons(self, buttons):
        """ปิดใช้งานปุ่ม"""
        for button in buttons:
            try:
                if button.winfo_exists():
                    button.configure(state="disabled")
            except:
                pass

    def _enable_buttons(self, buttons):
        """เปิดใช้งานปุ่ม"""
        for button in buttons:
            try:
                if button.winfo_exists():
                    button.configure(state="normal")
            except:
                pass

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
            self._clear_form(confirm=False)
            self.editing_po_id = po_id
            
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. Load Header
                cursor.execute("SELECT * FROM purchase_orders WHERE id = %s", (po_id,))
                po_data = cursor.fetchone()
                if not po_data: 
                    messagebox.showerror("Error", "ไม่พบ PO ที่ต้องการแก้ไข", parent=self)
                    return
                
                # 2. Load Items
                cursor.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", (po_id,))
                items_data = cursor.fetchall()
                
                # 3. Load Payments
                cursor.execute("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", (po_id,))
                payments_data = cursor.fetchall()

            # --- Populate Form ---

            # SO Number
            so_num = po_data.get("so_number", "")
            if so_num:
                values = self.so_entry.cget("values")
                matching_so_string = next((s for s in values if s.startswith(so_num)), None)
                if matching_so_string:
                    self.so_entry.set(matching_so_string)
                    self._on_so_selected(matching_so_string, is_editing=True)
                else:
                    self.so_entry.set(so_num)

            # Supplier Info
            supplier_name_from_db = str(po_data.get("supplier_name") or "")
            self.supplier_name_combo.delete(0, 'end')
            self.supplier_name_combo.insert(0, supplier_name_from_db)
            
            supplier_dict_to_load = next((item for item in self.supplier_completion_data if item.get('name') == supplier_name_from_db), None)
            if supplier_dict_to_load:
                self._on_supplier_selected(supplier_dict_to_load)

            # Basic Info
            po_full_number = po_data.get("po_number", "PO")
            self.po_number_type_var.set("ST" if po_full_number.startswith("ST") else "PO")
            self.po_number_input_var.set(po_full_number)
            self.rr_number_var.set(po_data.get("rr_number", ""))
            utils.set_entry_text(self.department_entry, po_data.get("department", ""))
            utils.set_entry_text(self.pur_order_entry, po_data.get("pur_order", ""))
            self.po_mode_var.set(po_data.get("po_mode", "Single-PO")) 

            # Shipping (Stock)
            stock_cost_val = po_data.get('shipping_to_stock_cost', 0) or 0
            utils.set_entry_text(self.shipping_to_stock_cost_entry, f"{stock_cost_val:.2f}")
            self.shipping_to_stock_vat_var.set(po_data.get("shipping_to_stock_vat_type", "VAT"))
            self.shipping_to_stock_wht_var.set(po_data.get("shipping_to_stock_wht_type", "ไม่มีหัก"))
            self.shipping_to_stock_date_selector.set_date(po_data.get("shipping_to_stock_date"))
            self.shipping_to_stock_type_var.set(po_data.get("shipping_to_stock_shipper", "ซัพพลายเออร์จัดส่ง"))
            utils.set_entry_text(self.shipping_to_stock_driver_entry, po_data.get("shipping_to_stock_driver", ""))
            utils.set_entry_text(self.shipping_to_stock_plate_entry, po_data.get("shipping_to_stock_plate", ""))
            utils.set_entry_text(self.shipping_to_stock_notes_entry, po_data.get("shipping_to_stock_notes", ""))

            # Shipping (Site)
            site_cost_val = po_data.get('shipping_to_site_cost', 0) or 0
            utils.set_entry_text(self.shipping_to_site_cost_entry, f"{site_cost_val:.2f}")
            self.shipping_to_site_vat_var.set(po_data.get("shipping_to_site_vat_type", "VAT"))
            self.shipping_to_site_wht_var.set(po_data.get("shipping_to_site_wht_type", "ไม่มีหัก"))
            self.shipping_to_site_date_selector.set_date(po_data.get("shipping_to_site_date"))
            self.shipping_to_site_type_var.set(po_data.get("shipping_to_site_shipper", "ซัพพลายเออร์จัดส่ง"))
            utils.set_entry_text(self.shipping_to_site_driver_entry, po_data.get("shipping_to_site_driver", ""))
            utils.set_entry_text(self.shipping_to_site_plate_entry, po_data.get("shipping_to_site_plate", ""))
            utils.set_entry_text(self.shipping_to_site_notes_entry, po_data.get("shipping_to_site_notes", ""))

            # Call Sync (Optional but good)
            if so_num: self.sync_transport_cost_to_po(so_num)

            # Cutting
            utils.set_entry_text(self.cutting_cost_entry, f"{po_data.get('cutting_cost', 0):.2f}")
            self.cutting_vat_var.set(po_data.get("cutting_vat_type", "VAT"))
            self.cutting_wht_var.set(po_data.get("cutting_wht_type", "No"))
            utils.set_entry_text(self.cutting_remark_entry, po_data.get("cutting_remark", ""))

            # Items
            for row in self.product_rows:
                for widget in row["widgets"]: widget.destroy()
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
            
            # Checkboxes & Discount
            if po_data.get("vat_7_percent_checked"): self.vat_checkbox.select()
            else: self.vat_checkbox.deselect()
            if po_data.get("wht_3_percent_checked"): self.vat3_checkbox.select()
            else: self.vat3_checkbox.deselect()

            utils.set_entry_text(self.end_of_bill_discount_entry, f"{po_data.get('bill_discount', 0):.2f}")

            # Payments
            for p_data in payments_data:
                p_type = p_data.get('payment_type')
                if p_type in self.payment_entries:
                    p_widgets = self.payment_entries[p_type]
                    
                    utils.set_entry_text(p_widgets["amount"], f"{p_data.get('amount', 0):,.2f}")
                    
                    if p_widgets.get("date"): 
                        p_widgets["date"].set_date(p_data.get("payment_date"))
                    
                    if p_widgets.get("bank_menu"): 
                        p_widgets["bank_menu"].set(p_data.get("bank_name", "ระบุเอง"))
                    
                    if p_widgets.get("account_entry"): 
                        utils.set_entry_text(p_widgets["account_entry"], p_data.get("bank_account_number", ""))
                    
                    # [🔥 แก้ไข] โหลดค่าประเภทบัญชี (ถ้ามีใน DB)
                    if p_widgets.get("acc_type_var"):
                        acc_type = p_data.get("bank_account_type")
                        if acc_type and acc_type in ["ออมทรัพย์", "กระแสรายวัน"]:
                            p_widgets["acc_type_var"].set(acc_type)
                        else:
                            p_widgets["acc_type_var"].set("ออมทรัพย์") # Default
            
            self._update_summary()

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการโหลดข้อมูล: {e}\n{traceback.format_exc()}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
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
        self.shipping_to_stock_wht_var.set("ไม่มีหัก")
        self.shipping_to_site_wht_var.set("ไม่มีหัก")
        
        # Reset Cutting Vars
        self.cutting_vat_var.set("VAT")
        self.cutting_wht_var.set("No")
        
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
            # Stock
            self.shipping_to_stock_cost_entry, self.shipping_to_stock_notes_entry,
            self.shipping_to_stock_driver_entry, self.shipping_to_stock_plate_entry, # [🔥 เพิ่ม]
            
            # Site
            self.shipping_to_site_cost_entry, self.shipping_to_site_notes_entry,
            self.shipping_to_site_driver_entry, self.shipping_to_site_plate_entry,   # [🔥 เพิ่ม]
            
            # Cutting
            self.cutting_cost_entry, self.cutting_remark_entry,
            
            # Totals
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
        
        for p_type in ["Payment 1", "Payment 2", "Full Payment", "CN Refund"]:
            p_dict = self.payment_entries.get(p_type)
            if p_dict:
                if p_dict.get("amount") and p_dict["amount"].winfo_exists():
                    p_dict['amount'].delete(0, "end")
                if p_dict.get("date") and p_dict["date"].winfo_exists():
                    p_dict["date"].set_date(None)
                if p_dict.get("percent_var"):
                    p_dict["percent_var"].set("ระบุยอดเอง")
                if p_dict.get("account_entry") and p_dict["account_entry"].winfo_exists():
                    p_dict["account_entry"].delete(0, "end")
                if p_dict.get("bank_menu") and p_dict["bank_menu"].winfo_exists():
                    p_dict["bank_menu"].set(self.bank_list[0])
            
        self.total_deposit_var.set("0.00")
        self.balance_due_var.set("0.00")

        if hasattr(self, 'vat3_checkbox'): self.vat3_checkbox.deselect()
        if hasattr(self, 'vat_checkbox'): self.vat_checkbox.deselect()
            
        self._update_summary()

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

    def _find_save_buttons(self):
        """หาปุ่มบันทึกทั้งหมดในฟอร์ม"""
        save_buttons = []
        # หาปุ่มที่มีคำว่า 'บันทึก' หรือ 'Save' ใน text
        for child in self.winfo_children():
            self._find_buttons_recursive(child, save_buttons, ["บันทึก", "Save", "ขออนุมัติ"])
        return save_buttons

    def _find_buttons_recursive(self, widget, button_list, search_texts):
        """หาปุ่มแบบ recursive"""
        try:
            # ตรวจสอบว่าเป็นปุ่มหรือไม่
            if hasattr(widget, 'cget') and hasattr(widget, 'configure'):
                try:
                    text = widget.cget('text')
                    if text and any(search_text in text for search_text in search_texts):
                        button_list.append(widget)
                except:
                    pass
            
            # ตรวจสอบ children
            for child in widget.winfo_children():
                self._find_buttons_recursive(child, button_list, search_texts)
        except:
            pass

    def _disable_buttons(self, buttons):
        """ปิดใช้งานปุ่ม"""
        for button in buttons:
            try:
                if button.winfo_exists():
                    button.configure(state="disabled")
            except:
                pass

    def _enable_buttons(self, buttons):
        """เปิดใช้งานปุ่ม"""
        for button in buttons:
            try:
                if button.winfo_exists():
                    button.configure(state="normal")
            except:
                pass

