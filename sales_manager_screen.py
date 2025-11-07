# sales_manager_screen.py (ฉบับปรับปรุง เพิ่มแท็บ Master Edit)

import tkinter as tk
from tkinter import ttk
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                           CTkOptionMenu, CTkRadioButton, CTkTabview) # <-- เพิ่ม CTkTabview
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import psycopg2.errors
import psycopg2.extras
import traceback
import utils

# --- นำเข้า Class ที่จำเป็น ---\
from history_windows import SOPopupWindow
from custom_widgets import NumericEntry, DateSelector

class SalesManagerScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role
        
        self.label_font = CTkFont(size=14, weight="bold")
        self.entry_font = CTkFont(size=14)
        
        # --- START: เพิ่มตัวแปรและ Theme ที่จำเป็นสำหรับ SOPopupWindow ---\
        self.so_popup = None
        self.so_form_widgets = {}
        self._so_create_string_vars() # สร้าง StringVars ที่จำเป็น
        
        # กำหนด Theme (อาจปรับตามความเหมาะสม)
        self.sale_theme = self.app_container.THEME.get("sale", {"bg": "white", "primary": "blue"})
        # --- END ---
        
        self.pg_engine = self.app_container.pg_engine
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # <-- แถวที่ 1 (Tabview) จะขยาย

        # --- 1. สร้าง Header (เหมือนเดิม) ---
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        CTkLabel(header_frame, text=f"หน้าจอผู้จัดการฝ่ายขาย: {self.user_name}", font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        button_frame = CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side="right", padx=10)
        
        CTkButton(button_frame, text="Refresh", command=self._load_pending_so).pack(side="left", padx=5)
        CTkButton(button_frame, text="ออกจากระบบ", command=self.app_container.show_login_screen, 
                    fg_color="transparent", border_color="#D32F2F", 
                    text_color="#D32F2F", border_width=2, 
                    hover_color="#FFEBEE").pack(side="left", padx=5)

        # --- 2. สร้าง TabView ---
        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.pending_tab = self.tab_view.add("รายการรออนุมัติ")
        self.master_edit_tab = self.tab_view.add("ค้นหา & ตีกลับ SO (Master)")

        # --- 3. สร้างเนื้อหาสำหรับแต่ละแท็บ ---
        self._create_pending_tab(self.pending_tab)
        self._create_master_edit_tab(self.master_edit_tab)
        
        # --- 4. โหลดข้อมูลเริ่มต้นสำหรับแท็บแรก ---
        self._load_pending_so() 

    def _create_pending_tab(self, parent_tab):
        """สร้าง UI สำหรับแท็บรออนุมัติ (โค้ดเดิม)"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(0, weight=1)
        self.main_frame = CTkScrollableFrame(parent_tab, label_text="SO ที่รอการอนุมัติ (Pending Sale Manager Approval)")
        self.main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

    def _create_master_edit_tab(self, parent_tab):
        """(ฟังก์ชันใหม่) สร้าง UI สำหรับแท็บ Master Edit"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1) # แถวที่ 1 (ScrollFrame) จะขยาย

        # --- Frame สำหรับฟิลเตอร์และการค้นหา ---
        search_frame = CTkFrame(parent_tab)
        search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)
        
        self.sm_master_search_entry = CTkEntry(search_frame, font=self.entry_font, placeholder_text="กรอก SO Number ที่ต้องการค้นหา...")
        self.sm_master_search_entry.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="ew")
        
        self.sm_master_search_entry.bind("<Return>", lambda event: self._sm_master_search())
        self.sm_master_search_entry.bind("<KP_Enter>", lambda event: self._sm_master_search())
        
        search_button = CTkButton(search_frame, text="ค้นหา", command=self._sm_master_search, width=100)
        search_button.grid(row=0, column=1, padx=5, pady=10)

        # --- Frame สำหรับแสดงผลลัพธ์ ---
        self.sm_master_results_frame = CTkScrollableFrame(parent_tab, label_text="ผลการค้นหา")
        self.sm_master_results_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.sm_master_results_frame.grid_columnconfigure(0, weight=1)

    def _sm_master_search(self):
        """(ฟังก์ชันใหม่) ค้นหา SO ทั้งหมดจาก Keyword"""
        for widget in self.sm_master_results_frame.winfo_children():
            widget.destroy()
            
        keyword = self.sm_master_search_entry.get().strip().upper()
        if not keyword:
            return

        search_term = keyword.replace("SO", "") # ค้นหาแบบง่ายๆ
        
        try:
            # ค้นหา SO (ดึงข้อมูลเซลส์และสถานะมาด้วย)
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status, u.sale_name
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.so_number ILIKE %s AND c.is_active = 1
            """
            so_df = pd.read_sql_query(query, self.pg_engine, params=(f"%{search_term}%",))

            if so_df.empty:
                CTkLabel(self.sm_master_results_frame, text=f"ไม่พบข้อมูลสำหรับ '{keyword}'").pack(pady=20)
                return

            for _, row in so_df.iterrows():
                # สร้าง Card แสดงผลลัพธ์
                self._create_sm_master_so_card(self.sm_master_results_frame, row.to_dict())

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการค้นหา: {e}", parent=self)

    def _create_sm_master_so_card(self, parent, so_data):
        """(ฟังก์ชันใหม่) สร้าง Card แสดงผลลัพธ์ SO ในแท็บ Master Edit"""
        so_id = so_data['id']
        so_number = so_data['so_number']
        so_status = so_data.get('status', 'N/A')
        
        # กำหนดสีพื้นหลังตามสถานะ
        status_colors = {
            'Paid': '#D1FAE5', 'HR Verified': '#A7F3D0', 
            'PO Sent': '#67E8F9', 'Pending Sale Manager Approval': '#FDE047',
            'PO In Progress': '#FEF3C7', 'Draft': '#E5E7EB',
            'Rejected': '#FECACA', 'Cancelled': '#FECACA', 'Rejected by SM': '#FECACA'
        }
        card_color = status_colors.get(so_status, "#F9FAFB")

        so_card = CTkFrame(parent, border_width=1, fg_color=card_color)
        so_card.pack(fill="x", padx=10, pady=5)
        
        info_text = f"SO: {so_number} | ลูกค้า: {so_data.get('customer_name','N/A')} | สถานะ: {so_status} | เซลส์: {so_data.get('sale_name', 'N/A')}"
        CTkLabel(so_card, text=info_text, font=self.entry_font).pack(side="left", padx=10, pady=5, anchor="w")
        
        action_frame = CTkFrame(so_card, fg_color="transparent")
        action_frame.pack(side="right", padx=10, pady=5)

        # ปุ่มที่ 1: "แก้ไข SO" (เรียกใช้ฟังก์ชันเดิม)
        CTkButton(
            action_frame, 
            text="แก้ไข SO", 
            width=100, 
            command=lambda s_num=so_number: self._open_so_editor_for_sm(s_num)
        ).pack(side="left", padx=5)
        
        # ปุ่มที่ 2: "ตีกลับ SO" (เรียกใช้ฟังก์ชันเดิม)
        # เราจะแสดงปุ่มนี้ ตราบใดที่ SO ยังไม่ถูกตีกลับไปแล้ว
        if so_status not in ('Draft', 'Rejected by SM', 'Cancelled'):
            CTkButton(
                action_frame, 
                text="ตีกลับ SO (Revert)", 
                width=140, 
                fg_color="#D32F2F", hover_color="#B71C1C", # สีแดง
                command=lambda s_id=so_id, s_num=so_number: self._reject_so(s_id, s_num)
            ).pack(side="left", padx=5)


    def _so_create_string_vars(self):
        """สร้าง StringVars ทั้งหมดที่ SOPopupWindow ต้องการ"""
        self.so_shared_vars = {}
        now = datetime.now()
        thai_months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        self.so_shared_vars['thai_months'] = thai_months_list
        self.so_shared_vars['thai_month_map'] = {name: i + 1 for i, name in enumerate(thai_months_list)}
        self.so_shared_vars['customer_type_var'] = tk.StringVar(value="ลูกค้าเก่า")
        self.so_shared_vars['credit_term_var'] = tk.StringVar(value="เงินสด")
        self.so_shared_vars['commission_month_var'] = tk.StringVar(value=thai_months_list[now.month - 1])
        self.so_shared_vars['commission_year_var'] = tk.StringVar(value=str(now.year + 543))
        self.so_shared_vars['payment1_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_shared_vars['payment2_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_shared_vars['payment_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_subtotal_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_vat_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_shared_vars['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['balance_due_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_product_input_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_service_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_actual_payment_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_verification_result_var'] = tk.StringVar(value="-")
        
        self.so_shared_vars['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['relocation_vat_calc_var'] = tk.StringVar(value="0.00")

        self.so_shared_vars['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_shared_vars['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_shared_vars['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_shared_vars['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_shared_vars['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_shared_vars['relocation_cost_vat_option'] = tk.StringVar(value="VAT")

        self.so_shared_vars['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")

    def _open_so_editor_for_sm(self, so_number):
        """(ปรับปรุง) เปิดหน้าต่างแก้ไข SO และตั้งค่า callback ให้ refresh ทั้งสองแท็บ"""
        if self.so_popup is not None and self.so_popup.winfo_exists():
            self.so_popup.focus()
            return

        try:
            so_df = pd.read_sql_query("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", self.pg_engine, params=(so_number,))
            if so_df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูล SO: {so_number}", parent=self)
                return
            
            # --- สร้างฟังก์ชัน Callback ---
            def _refresh_on_save():
                self._load_pending_so() # Refresh แท็บรออนุมัติ
                if hasattr(self, 'sm_master_search_entry'): # Refresh แท็บค้นหา
                    self._sm_master_search()
            
            self.so_popup = SOPopupWindow(
                master=self,
                app_container=self.app_container,
                sales_data=so_df.iloc[0].to_dict(),
                so_shared_vars=self.so_shared_vars,
                sale_theme=self.sale_theme,
                on_save_callback=_refresh_on_save # <-- ใช้ callback ที่เราสร้าง
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างแก้ไข SO ได้: {e}", parent=self)
            traceback.print_exc()
            if self.so_popup:
                self.so_popup.destroy()
                self.so_popup = None

    def _load_pending_so(self):
        """โหลด SO ที่รอการอนุมัติ (สำหรับแท็บแรก)"""
        # ตรวจสอบว่า main_frame ถูกสร้างแล้วหรือยัง
        if not hasattr(self, 'main_frame') or not self.main_frame.winfo_exists():
            self.after(100, self._load_pending_so) # ถ้ายังไม่สร้าง ให้รอ 100ms แล้วลองใหม่
            return
            
        for widget in self.main_frame.winfo_children():
            widget.destroy()
            
        try:
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, u.sale_name, c.timestamp
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Pending Sale Manager Approval' AND c.is_active = 1
                ORDER BY c.timestamp DESC
            """
            pending_so_df = pd.read_sql_query(query, self.pg_engine)
            
            if pending_so_df.empty:
                CTkLabel(self.main_frame, text="ไม่พบรายการที่รอการอนุมัติในขณะนี้").pack(pady=20)
                return
                
            for _, row in pending_so_df.iterrows():
                self._create_pending_so_card(row.to_dict())
                
        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดข้อมูล: {e}", parent=self)
            traceback.print_exc()

    def _create_pending_so_card(self, so_data):
        """สร้าง Card สำหรับ SO ที่รออนุมัติ (สำหรับแท็บแรก)"""
        so_id = so_data['id']
        so_number = so_data['so_number']
        
        card = CTkFrame(self.main_frame, border_width=1, corner_radius=10)
        card.pack(fill="x", padx=10, pady=5)
        
        card.grid_columnconfigure(0, weight=1)
        
        info_frame = CTkFrame(card, fg_color="transparent")
        info_frame.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        
        CTkLabel(info_frame, text=f"SO: {so_number}", font=CTkFont(size=16, weight="bold")).pack(anchor="w")
        CTkLabel(info_frame, text=f"ลูกค้า: {so_data.get('customer_name', 'N/A')}").pack(anchor="w")
        CTkLabel(info_frame, text=f"เซลส์: {so_data.get('sale_name', 'N/A')} ({so_data.get('sale_key', 'N/A')})").pack(anchor="w")
        CTkLabel(info_frame, text=f"วันที่ส่ง: {so_data.get('timestamp', pd.Timestamp.now()).strftime('%Y-%m-%d %H:%M')}").pack(anchor="w")
        
        action_frame = CTkFrame(card, fg_color="transparent")
        action_frame.grid(row=0, column=1, padx=15, pady=10, sticky="e")
        
        # --- ปุ่มสำหรับแท็บรออนุมัติ ---
        approve_button = CTkButton(action_frame, text="อนุมัติ", 
                                   command=lambda s_id=so_id, s_num=so_number: self._approve_so(s_id, s_num),
                                   fg_color="#16A34A", hover_color="#15803D")
        approve_button.pack(pady=5, fill="x")
        
        reject_button = CTkButton(action_frame, text="ตีกลับ (Reject)", 
                                  command=lambda s_id=so_id, s_num=so_number: self._reject_so(s_id, s_num),
                                  fg_color="#D32F2F", hover_color="#B71C1C")
        reject_button.pack(pady=5, fill="x")
        
        # ผูก Double-click กับการ์ด (เหมือนเดิม)
        card.bind("<Double-1>", lambda e, s_num=so_number: self._on_so_card_double_click(e, s_num))
        for widget in info_frame.winfo_children():
            widget.bind("<Double-1>", lambda e, s_num=so_number: self._on_so_card_double_click(e, s_num))

    def _on_so_card_double_click(self, event, so_number):
        """เมื่อดับเบิลคลิกที่ SO Card ให้เปิดหน้าต่างแก้ไข"""
        self._open_so_editor_for_sm(so_number)

    def _approve_so(self, so_id_to_approve, so_number):
        """อนุมัติ SO (ฟังก์ชันเดิม)"""
        if not messagebox.askyesno("ยืนยันการอนุมัติ", f"คุณต้องการอนุมัติ SO: {so_number} ใช่หรือไม่?\nSO จะถูกส่งต่อไปยังฝ่ายจัดซื้อ (PU)", parent=self):
            return
            
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 1. อัปเดตสถานะ SO
                cursor.execute(
                    "UPDATE commissions SET status = 'PO In Progress', approver_sale_manager_key = %s, approval_date_sale_manager = %s WHERE id = %s",
                    (self.user_key, datetime.now(), so_id_to_approve)
                )
                
                # 2. สร้าง Notification แจ้งเตือนฝ่ายจัดซื้อ (PU)
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Staff' AND status = 'Active'")
                pu_keys = [row[0] for row in cursor.fetchall()]
                
                message = f"SO: {so_number} ได้รับการอนุมัติแล้ว กรุณาสร้าง PO"
                for pu_key in pu_keys:
                    cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) VALUES (%s, %s, FALSE, %s)",
                        (pu_key, message, so_id_to_approve)
                    )
                
            conn.commit()
            messagebox.showinfo("สำเร็จ", "อนุมัติ SO เรียบร้อยแล้ว\nแจ้งเตือนฝ่ายจัดซื้อเพื่อดำเนินการต่อ", parent=self)
            self._load_pending_so() # Refresh หน้าจอ
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการอนุมัติ: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _reject_so(self, so_id_to_reject, so_number):
        """(ปรับปรุง) ตีกลับ SO (ใช้สำหรับทั้งสองแท็บ)"""
        dialog = CTkInputDialog(text=f"กรุณาระบุเหตุผลที่ตีกลับ SO: {so_number}", title="ตีกลับ SO")
        reason = dialog.get_input()
        if not reason or not reason.strip():
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute(
                    "UPDATE commissions SET status = 'Rejected by SM', rejection_reason = %s WHERE id = %s",
                    (reason.strip(), so_id_to_reject)
                )
                
                # หา sale_key และสร้าง Notification
                cursor.execute("SELECT sale_key FROM commissions WHERE id = %s", (so_id_to_reject,))
                sale_key = cursor.fetchone()[0]
                
                message = f"SO: {so_number} ถูกตีกลับโดย SM\nเหตุผล: {reason.strip()}"
                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) VALUES (%s, %s, FALSE, %s)",
                    (sale_key, message, so_id_to_reject)
                )
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ตีกลับ SO เรียบร้อยแล้ว", parent=self)
            
            # --- START: Refresh ทั้งสองส่วน ---
            self._load_pending_so() # Refresh แท็บรออนุมัติ
            if hasattr(self, 'sm_master_search_entry'): # Refresh แท็บค้นหา (ถ้ามี)
                self._sm_master_search()
            # --- END ---
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)