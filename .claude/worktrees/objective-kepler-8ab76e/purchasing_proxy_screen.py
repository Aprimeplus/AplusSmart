# purchasing_proxy_screen.py (เนื้อหาที่ถูกต้อง)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu
from tkinter import messagebox
import pandas as pd

from purchasing_screen import PurchasingScreen

class PurchasingProxyScreen(CTkFrame):
    def __init__(self, master, app_container, proxy_user_key, proxy_user_name, role_to_proxy="Purchasing Staff"):
        super().__init__(master) 

        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.proxy_user_key = proxy_user_key
        self.proxy_user_name = proxy_user_name
        self.role_to_proxy = role_to_proxy
        self.pu_key_owner = None 

        self.grid_rowconfigure(1, weight=1)    
        self.grid_columnconfigure(0, weight=1) 

        self.header_container = CTkFrame(self, fg_color="transparent")
        self.header_container.grid(row=0, column=0, sticky="ew")

        content_container = CTkFrame(self, fg_color="transparent")
        content_container.grid(row=1, column=0, sticky="nsew")

        self.purchasing_screen = PurchasingScreen(
            master=content_container,
            app_container=app_container,
            user_key=proxy_user_key,
            user_name=proxy_user_name,
            user_role=app_container.current_user_role
        )
        self.purchasing_screen.pack(fill="both", expand=True)

        self.after(50, self._reconfigure_buttons)
        
        self.active_pu_list = self._get_all_active_pus()
        self._setup_proxy_ui()

    def _reconfigure_buttons(self):
        if hasattr(self.purchasing_screen, 'save_draft_button'):
            self.purchasing_screen.save_draft_button.configure(command=self._proxy_save_draft)
        else:
            print("WARNING: ไม่พบ 'save_draft_button' ใน purchasing_screen instance")

    def _setup_proxy_ui(self):
        self._create_pu_selection_header()
        self._toggle_main_form(state="disabled")

    def _create_pu_selection_header(self):
        selection_header = CTkFrame(self.header_container, fg_color="#F0FDF4", border_width=1, border_color="#4ADE80")
        selection_header.pack(side="top", fill="x", padx=20, pady=(10, 5))
        
        CTkLabel(selection_header, text=f"ผู้ดำเนินการ: {self.proxy_user_name} ({self.proxy_user_key})", font=CTkFont(size=16, weight="bold")).pack(side="left", padx=10, pady=10)
        
        pu_selection_frame = CTkFrame(selection_header, fg_color="transparent")
        pu_selection_frame.pack(side="left", padx=10, pady=10)
        
        CTkLabel(pu_selection_frame, text="ทำงานในนามของ (จัดซื้อ):", font=CTkFont(size=14)).pack(side="left")
        
        pu_display_names = [f"{pu['sale_name']} ({pu['sale_key']})" for pu in self.active_pu_list]
        self.selected_pu_key_full = tk.StringVar(value="- กรุณาเลือกพนักงาน -")
        
        self.pu_selector_menu = CTkOptionMenu(
            pu_selection_frame,
            variable=self.selected_pu_key_full,
            values=["- กรุณาเลือกพนักงาน -"] + pu_display_names,
            command=self._on_pu_selected
        )
        self.pu_selector_menu.pack(side="left", padx=10)

    def _get_all_active_pus(self):
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = %s AND status = 'Active' ORDER BY sale_name"
            df = pd.read_sql(query, self.pg_engine, params=(self.role_to_proxy,))
            return df.to_dict('records')
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงรายชื่อพนักงานจัดซื้อได้: {e}", parent=self)
            return []

    def _on_pu_selected(self, selected_display_name):
        if "- กรุณาเลือกพนักงาน -" in selected_display_name:
            self.pu_key_owner = None
            self._toggle_main_form(state="disabled")
        else:
            selected_pu_data = next((pu for pu in self.active_pu_list if f"{pu['sale_name']} ({pu['sale_key']})" == selected_display_name), None)
            if selected_pu_data:
                self.pu_key_owner = selected_pu_data['sale_key']
                self.purchasing_screen.user_key = self.pu_key_owner 
                self.purchasing_screen.user_name = selected_pu_data['sale_name']
                self._toggle_main_form(state="normal")
            else:
                self.pu_key_owner = None
                self._toggle_main_form(state="disabled")

    def _toggle_main_form(self, state):
        widgets_to_toggle = []
        if hasattr(self.purchasing_screen, 'control_frame'): widgets_to_toggle.append(self.purchasing_screen.control_frame)
        if hasattr(self.purchasing_screen, 'main_content_frame'): widgets_to_toggle.append(self.purchasing_screen.main_content_frame)

        for container in widgets_to_toggle:
            if hasattr(container, 'winfo_children'):
                for child in container.winfo_children():
                    try:
                        child.configure(state=state)
                    except (tk.TclError, AttributeError):
                        pass

    def _proxy_save_draft(self):
        if not self.pu_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกพนักงาน", "กรุณาเลือกพนักงานจัดซื้อที่ต้องการทำงานแทน", parent=self)
            return

        self.purchasing_screen.proxy_user_key = self.proxy_user_key
        self.purchasing_screen._save_po(status='Draft')

        if hasattr(self.purchasing_screen, 'proxy_user_key'):
            del self.purchasing_screen.proxy_user_key