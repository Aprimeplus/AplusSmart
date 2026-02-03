# sales_proxy_screen.py (ฉบับ Debug: บังคับโชว์ปุ่ม)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd
from sales_support_features import SalesSupportOutstandingManager
# Import คลาสแม่
from commission_app import CommissionApp

class SalesProxyScreen(CommissionApp):
    """
    หน้าจอสำหรับให้ Role อื่น (เช่น Sale Support, HR) ทำงานในนามของ Sale
    """
    def __init__(self, master, app_container, proxy_user_key, proxy_user_name, user_role, role_to_proxy="Sale", show_logout_button=False):
        
        self.proxy_user_key = proxy_user_key
        self.proxy_user_name = proxy_user_name
        self.role_to_proxy = role_to_proxy
        self.show_logout_button = show_logout_button
        self.sale_key_owner = None
        self.user_role = user_role 

        # เรียกใช้ Class แม่ (CommissionApp)
        super().__init__(master=master, 
                           app_container=app_container, 
                           sale_key=proxy_user_key, 
                           sale_name=proxy_user_name,
                           user_role=user_role,
                           show_logout_button=show_logout_button,
                           create_default_header=True)

        # ======================================================================
        # [🔥 โค้ดที่ต้องเพิ่ม] สร้างแท็บติดตามหนี้ (วางไว้หลัง super().__init__)
        # ======================================================================
        target_roles = ['Sales Support', 'Admin', 'Manager', 'Director']
        
        if self.user_role in target_roles:
            # ชื่อแท็บใหม่
            tab_name = "ติดตามหนี้ (Support)"
            
            # ตรวจสอบว่า self.tab_view มีอยู่จริง (มาจาก CommissionApp)
            if hasattr(self, 'tab_view'):
                # สร้างแท็บใหม่ (ถ้ายังไม่มี)
                try:
                    self.tab_view.tab(tab_name)
                except ValueError: # กรณีไม่มีแท็บชื่อนี้
                    self.tab_view.add(tab_name)
                
                # สร้างหน้าจอ Outstanding Manager ใส่ลงในแท็บนั้น
                self.outstanding_manager = SalesSupportOutstandingManager(
                    master=self.tab_view.tab(tab_name), 
                    app_container=self.app_container
                )
                self.outstanding_manager.pack(fill="both", expand=True)
            else:
                print("Warning: self.tab_view not found in CommissionApp")
        # ======================================================================

        self.active_sales_list = self._get_all_active_sales()
        
        # รอให้ UI วาดเสร็จแล้วค่อยสร้าง Header พิเศษด้านบน
        self.after(10, self._setup_proxy_ui)

    def _setup_proxy_ui(self):
        """สร้าง UI พิเศษสำหรับหน้า Proxy"""
        self._create_sale_selection_header()
        self._toggle_main_form(state="disabled")

    def _create_sale_selection_header(self):
        """สร้างแถบ Header สีเหลืองสำหรับเลือกเซลส์"""
        for widget in self.header_frame.winfo_children():
            widget.destroy()

        selection_header = CTkFrame(self.header_frame, fg_color="#FFFBEB", border_width=1, border_color="#FBBF24")
        selection_header.pack(side="left", fill="x", expand=True, pady=(0, 10))
        
        CTkLabel(selection_header, text=f"ผู้ดำเนินการ: {self.proxy_user_name} ({self.proxy_user_key})", font=CTkFont(size=16, weight="bold"), text_color="black").pack(side="left", padx=10, pady=10)
        
        sale_selection_frame = CTkFrame(selection_header, fg_color="transparent")
        sale_selection_frame.pack(side="left", padx=10, pady=10)
        
        CTkLabel(sale_selection_frame, text="ทำงานในนามของ (เซลส์):", font=CTkFont(size=14), text_color="black").pack(side="left")
        
        sale_display_names = [f"{sale['sale_name']} ({sale['sale_key']})" for sale in self.active_sales_list]
        self.selected_sale_key_full = tk.StringVar(value="- กรุณาเลือกเซลส์ -")
        
        self.sale_selector_menu = CTkOptionMenu(
            sale_selection_frame,
            variable=self.selected_sale_key_full,
            values=["- กรุณาเลือกเซลส์ -"] + sale_display_names,
            command=self._on_sale_selected
        )
        self.sale_selector_menu.pack(side="left", padx=10)

    # -----------------------------------------------------------
    #  ★ จุดสำคัญ: บังคับโชว์ปุ่มโดยไม่สนเงื่อนไข ★
    # -----------------------------------------------------------
    def _populate_action_frame(self, parent):
        # 1. สร้างปุ่มมาตรฐานของหน้า Sales ก่อน
        super()._populate_action_frame(parent)

        # -----------------------------------------------------------
        # ส่วนเครื่องมือพิเศษสำหรับ Sale Support / Admin
        # -----------------------------------------------------------
        target_roles = ['Sale Support', 'Admin', 'Manager', 'Director']
        
        # เช็คสิทธิ์: ถ้าเป็น Role ที่กำหนด ให้แสดงเครื่องมือพิเศษ
        if self.user_role in target_roles:
            
            # เส้นคั่นสวยๆ
            separator = tk.Frame(parent, height=2, bd=1, relief="sunken")
            separator.pack(fill="x", padx=20, pady=(15, 10))

            # หัวข้อ
            tool_label = CTkLabel(parent, text=f"เครื่องมือพิเศษ ({self.user_role}):", 
                                  font=CTkFont(size=14, weight="bold"), text_color="gray50")
            tool_label.pack(anchor="w", padx=20, pady=(0, 5))

            # --- [ปุ่มที่ 1] ติดตามหนี้ (เพิ่มใหม่) ---
            debt_btn = CTkButton(
                parent,
                text="💰 ติดตามยอดค้างชำระ (Debt Tracker)",
                fg_color="#059669",  # สีเขียว
                hover_color="#047857",
                height=40,
                font=CTkFont(size=16, weight="bold"),
                command=self._open_debt_tracking_window
            )
            debt_btn.pack(fill="x", padx=20, pady=(0, 10))

            # --- [ปุ่มที่ 2] ย้ายเจ้าของ SO (ของเดิม) ---
            reassign_btn = CTkButton(
                parent, 
                text="🔄 ย้ายเจ้าของ SO (Reassign Owner)", 
                fg_color="#8B5CF6", # สีม่วง
                hover_color="#7C3AED",
                height=40,
                font=CTkFont(size=16, weight="bold"),
                command=self._open_reassign_window
            )
            reassign_btn.pack(fill="x", padx=20, pady=(0, 20))

    def _open_debt_tracking_window(self):
        """เปิดหน้าต่างติดตามหนี้แบบ Popup"""
        try:
            # สร้างหน้าต่างใหม่ (Popup)
            debt_window = tk.Toplevel(self)
            debt_window.title("ระบบติดตามยอดค้างชำระ - Sales Support")
            debt_window.geometry("1100x700")
            
            # ทำให้หน้าต่างอยู่ตรงกลาง
            x = self.winfo_x() + 50
            y = self.winfo_y() + 50
            debt_window.geometry(f"+{x}+{y}")

            # เรียกใช้ Class ที่เราสร้างไว้
            # (ต้องแน่ใจว่า import SalesSupportOutstandingManager มาแล้วที่หัวไฟล์)
            debt_manager = SalesSupportOutstandingManager(
                master=debt_window, 
                app_container=self.app_container
            )
            debt_manager.pack(fill="both", expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถเปิดหน้าต่างได้: {e}")

    def _open_reassign_window(self):
        """เปิดหน้าต่างย้ายเจ้าของ SO"""
        try:
            from history_windows import SOReassignmentDialog
            SOReassignmentDialog(self, self.app_container)
        except ImportError:
            messagebox.showerror("Error", "ไม่พบ Class 'SOReassignmentDialog' ใน history_windows.py")
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")

    # --- ฟังก์ชัน Helper อื่นๆ ---
    def _get_all_active_sales(self):
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = %s AND status = 'Active' ORDER BY sale_name"
            df = pd.read_sql(query, self.pg_engine, params=(self.role_to_proxy,))
            return df.to_dict('records')
        except Exception as e:
            print(f"Database Error: {e}")
            return []

    def _on_sale_selected(self, selected_display_name):
        if "- กรุณาเลือกเซลส์ -" in selected_display_name:
            self.sale_key_owner = None
            self._toggle_main_form(state="disabled")
        else:
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                self.sale_key = self.sale_key_owner
                self.sale_name = selected_sale_data['sale_name']
                self._toggle_main_form(state="normal")
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="disabled")
    
    def _toggle_main_form(self, state):
        if hasattr(self, 'scrollable_main_container') and self.scrollable_main_container.winfo_exists():
            self._recursive_toggle_state(self.scrollable_main_container, state)

    def _recursive_toggle_state(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass
            if child.winfo_children():
                self._recursive_toggle_state(child, state)