# sales_proxy_screen.py (เนื้อหาที่ถูกต้อง)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd

# Import คลาสหลักที่เราจะสืบทอด
from commission_app import CommissionApp

class SalesProxyScreen(CommissionApp):
    """
    หน้าจอสำหรับให้ Role อื่น (เช่น Sale Support, HR) ทำงานในนามของ Sale
    สืบทอดความสามารถทั้งหมดมาจาก CommissionApp แต่เพิ่มส่วนของการเลือกคนทำงานแทน
    """
    def __init__(self, master, app_container, proxy_user_key, proxy_user_name, user_role, role_to_proxy="Sale", show_logout_button=False):
        
        self.proxy_user_key = proxy_user_key
        self.proxy_user_name = proxy_user_name
        self.role_to_proxy = role_to_proxy
        self.show_logout_button = show_logout_button
        self.sale_key_owner = None

        # --- การควบคุมการสร้าง Header ---
        # เราจะบอกให้ CommissionApp สร้าง Header Frame เสมอ
        # แต่จะแสดงปุ่ม Logout หรือไม่ ขึ้นอยู่กับค่า show_logout_button
        super().__init__(master=master, 
                         app_container=app_container, 
                         sale_key=proxy_user_key, 
                         sale_name=proxy_user_name,
                         user_role=user_role,
                         show_logout_button=show_logout_button,
                         create_default_header=True) # <--- บอกให้สร้าง Header Frame เสมอ

        self.active_sales_list = self._get_all_active_sales()
        
        # ใช้ .after() เพื่อรอให้ UI หลักของ CommissionApp วาดเสร็จก่อน
        self.after(1, self._setup_proxy_ui)

    def _setup_proxy_ui(self):
        """สร้าง UI พิเศษสำหรับหน้า Proxy และซ่อนฟอร์มหลักไว้ก่อน"""
        self._create_sale_selection_header()
        self._toggle_main_form(state="disabled")

    def _create_sale_selection_header(self):
        """สร้างแถบ Header สีเหลืองสำหรับเลือกเซลส์"""
        # ล้างวิดเจ็ตเก่าทั้งหมดใน header_frame ทิ้งไปก่อน
        for widget in self.header_frame.winfo_children():
            widget.destroy()

        # สร้างแถบสีเหลืองใส่เข้าไปใน header_frame ที่ว่างเปล่า
        selection_header = CTkFrame(self.header_frame, fg_color="#FFFBEB", border_width=1, border_color="#FBBF24")
        selection_header.pack(side="left", fill="x", expand=True, pady=(0, 10))
        
        CTkLabel(selection_header, text=f"ผู้ดำเนินการ: {self.proxy_user_name} ({self.proxy_user_key})", font=CTkFont(size=16, weight="bold")).pack(side="left", padx=10, pady=10)
        
        sale_selection_frame = CTkFrame(selection_header, fg_color="transparent")
        sale_selection_frame.pack(side="left", padx=10, pady=10)
        
        CTkLabel(sale_selection_frame, text="ทำงานในนามของ (เซลส์):", font=CTkFont(size=14)).pack(side="left")
        
        sale_display_names = [f"{sale['sale_name']} ({sale['sale_key']})" for sale in self.active_sales_list]
        self.selected_sale_key_full = tk.StringVar(value="- กรุณาเลือกเซลส์ -")
        
        self.sale_selector_menu = CTkOptionMenu(
            sale_selection_frame,
            variable=self.selected_sale_key_full,
            values=["- กรุณาเลือกเซลส์ -"] + sale_display_names,
            command=self._on_sale_selected
        )
        self.sale_selector_menu.pack(side="left", padx=10)
    
    def _get_all_active_sales(self):
        """ดึงรายชื่อเซลส์ที่ยัง Active อยู่ทั้งหมด"""
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = %s AND status = 'Active' ORDER BY sale_name"
            df = pd.read_sql(query, self.pg_engine, params=(self.role_to_proxy,))
            return df.to_dict('records')
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงรายชื่อเซลส์ได้: {e}", parent=self)
            return []

    def _on_sale_selected(self, selected_display_name):
        """Callback เมื่อมีการเลือกเซลส์จาก Dropdown"""
        if "- กรุณาเลือกเซลส์ -" in selected_display_name:
            self.sale_key_owner = None
            self._toggle_main_form(state="disabled")
        else:
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                # อัปเดต sale_key ของคลาสแม่ (CommissionApp) ให้เป็นของคนที่ถูกเลือก
                self.sale_key = self.sale_key_owner
                self.sale_name = selected_sale_data['sale_name']
                self._toggle_main_form(state="normal")
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="disabled")
    
    def _toggle_main_form(self, state):
        """เปิด/ปิด การใช้งานฟอร์มหลัก"""
        if hasattr(self, 'scrollable_main_container') and self.scrollable_main_container.winfo_exists():
            for child in self.scrollable_main_container.winfo_children():
                try:
                    child.configure(state=state)
                except (tk.TclError, AttributeError):
                    pass