# sale_support_screen.py (ฉบับแก้ไขสมบูรณ์)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd

from commission_app import CommissionApp, SubmitSODialog # <-- เพิ่ม SubmitSODialog

class SaleSupportApp(CommissionApp):
    def __init__(self, master, app_container, user_key, user_name, user_role):
        super().__init__(master, 
                         sale_key=user_key, 
                         sale_name=user_name, 
                         app_container=app_container, 
                         show_logout_button=True, 
                         user_role=user_role)

        self.support_user_key = user_key
        self.selected_sale_key_full = tk.StringVar()
        self.sale_key_owner = None

        self.active_sales_list = self._get_all_active_sales()
        self._create_sale_selection_header()
        self._toggle_main_form(state="disabled")

    def _get_all_active_sales(self):
        try:
            df = pd.read_sql("SELECT sale_key, sale_name FROM sales_users WHERE role = 'Sale' AND status = 'Active' ORDER BY sale_name", self.pg_engine)
            return df.to_dict('records')
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงรายชื่อพนักงานขายได้: {e}", parent=self)
            return []

    def _create_sale_selection_header(self):
        selection_header = CTkFrame(self, fg_color="#FFFBEB", border_width=1, border_color="#FBBF24")
        selection_header.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        selection_header.grid_columnconfigure(1, weight=1)
        CTkLabel(selection_header, text=f"Sale Support: {self.sale_name}", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        sale_selection_frame = CTkFrame(selection_header, fg_color="transparent")
        sale_selection_frame.grid(row=0, column=1, sticky="e", padx=10, pady=10)
        CTkLabel(sale_selection_frame, text="ทำงานในนามของ:", font=CTkFont(size=14)).pack(side="left")
        
        sale_display_names = [f"{sale['sale_name']} ({sale['sale_key']})" for sale in self.active_sales_list]
        self.sale_selector_menu = CTkOptionMenu(
            sale_selection_frame,
            variable=self.selected_sale_key_full,
            values=["- กรุณาเลือกเซลส์ -"] + sale_display_names,
            command=self._on_sale_selected
        )
        self.sale_selector_menu.pack(side="left", padx=10)
        CTkButton(selection_header, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2).grid(row=0, column=2, padx=10, pady=10)

    def _on_sale_selected(self, selected_display_name):
        if "- กรุณาเลือกเซลส์ -" in selected_display_name:
            self.sale_key_owner = None
            self._toggle_main_form(state="disabled")
        else:
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                self._toggle_main_form(state="normal")
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="disabled")

    def _toggle_main_form(self, state):
        try:
            if hasattr(self, 'save_button'): self.save_button.configure(state=state)
            if hasattr(self, 'btn_clear'): self.btn_clear.configure(state=state)
            if hasattr(self, 'btn_edit'): self.btn_edit.configure(state=state)
            for child in self.scrollable_main_container.winfo_children():
                for grandchild in child.winfo_children():
                    for widget in grandchild.winfo_children():
                        try: widget.configure(state=state)
                        except Exception: pass
        except Exception as e:
            print(f"Error toggling form state: {e}")

    def _save_data(self):
        if not self.sale_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกเซลส์", "กรุณาเลือกเซลส์ที่ต้องการทำงานแทนจากเมนูด้านบนก่อน", parent=self)
            return
        super()._save_data()

    def _gather_data_from_form(self):
        form_data = super()._gather_data_from_form()
        form_data['sale_key'] = self.sale_key_owner
        form_data['support_user_key'] = self.support_user_key
        return form_data

    # +++ START: เพิ่มฟังก์ชันใหม่ 2 ฟังก์ชันนี้ +++
    def _show_history(self):
        """
        (เวอร์ชัน Sale Support) เปิดหน้าต่างประวัติโดยส่งฟิลเตอร์ของ "เซลส์ที่ถูกเลือก" ไป
        """
        if not self.sale_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกเซลส์", "กรุณาเลือกเซลส์ก่อนดูประวัติ", parent=self)
            return
            
        try:
            self.history_window = self.app_container.show_history_window(
                sale_key_filter=self.sale_key_owner, # <-- ใช้รหัสเซลส์ที่เลือก
                edit_callback=self._on_history_so_select
            )
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถเปิดหน้าต่างประวัติได้: {e}", parent=self)

    def _open_submit_dialog(self):
        """
        (เวอร์ชัน Sale Support) เปิดหน้าต่างนำส่งข้อมูลโดยใช้ "เซลส์ที่ถูกเลือก"
        """
        if not self.sale_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกเซลส์", "กรุณาเลือกเซลส์ก่อนนำส่งข้อมูล", parent=self)
            return
            
        # ค้นหาชื่อของเซลส์ที่ถูกเลือก
        sale_name_owner = next((s['sale_name'] for s in self.active_sales_list if s['sale_key'] == self.sale_key_owner), "N/A")

        SubmitSODialog(self, self.app_container, self.sale_key_owner, sale_name_owner)
    # +++ END +++