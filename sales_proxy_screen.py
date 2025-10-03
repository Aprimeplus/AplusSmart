# sales_proxy_screen.py

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd



# Import คลาสหลักที่เราจะสืบทอด
from commission_app import CommissionApp, SubmitSODialog

class SalesProxyScreen(CommissionApp):
    """
    หน้าจอสำหรับให้ User Role อื่น (เช่น HR, Sale Support) ทำงานในนามของ Salesperson
    """
    def __init__(self, master, app_container, proxy_user_key, proxy_user_name, role_to_proxy="Sale"):
        # เรียก __init__ ของ CommissionApp (เหมือนเดิม)
        super().__init__(master, 
                         sale_key=proxy_user_key, 
                         sale_name=proxy_user_name, 
                         app_container=app_container, 
                         show_logout_button=False,
                         user_role=app_container.current_user_role)

        # +++++++++++++++++++++++++++++++
        # ✅✅✅ เพิ่ม 2 บรรทัดนี้เข้าไปครับ ✅✅✅
        # บอกให้ Frame หลัก (ตัวเอง) รู้จักขยายแถวและคอลัมน์
        self.grid_rowconfigure(1, weight=1) # ให้แถวที่ 1 (พื้นที่ฟอร์ม) ขยายตัว
        self.grid_columnconfigure(0, weight=1) # ให้คอลัมน์ที่ 0 ขยายตัว
        # +++++++++++++++++++++++++++++++

        self.proxy_user_key = proxy_user_key
        self.proxy_user_name = proxy_user_name
        self.role_to_proxy = role_to_proxy
        self.sale_key_owner = None

        self.active_sales_list = self._get_all_active_sales()
        
        self._create_sale_selection_header()
        
        self._toggle_main_form(state="disabled")

    def _get_all_active_sales(self):
        """ดึงรายชื่อเซลส์ที่ยัง Active อยู่ทั้งหมด"""
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = %s AND status = 'Active' ORDER BY sale_name"
            df = pd.read_sql(query, self.pg_engine, params=(self.role_to_proxy,))
            return df.to_dict('records')
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงรายชื่อพนักงานขายได้: {e}", parent=self)
            return []

    def _create_sale_selection_header(self):
        """สร้าง UI ส่วนหัวสำหรับเลือกเซลส์ที่จะทำงานแทน (ฉบับแก้ไข)"""
        # ซ่อน Header เดิมของ CommissionApp ที่แสดงชื่อผู้ Login
        if hasattr(self, 'header_frame'):
            # ✅ เปลี่ยนมาใช้ .pack_forget() เพื่อให้สอดคล้องกัน
            self.header_frame.pack_forget()

        selection_header = CTkFrame(self, fg_color="#FFFBEB", border_width=1, border_color="#FBBF24")
        
        # ✅ เปลี่ยนมาใช้ .pack() และใช้ before= เพื่อจัดลำดับให้ถูกต้อง
        #    ให้ header ใหม่นี้ แสดงผล 'ก่อน' ส่วนที่เป็น scrollable_main_container
        selection_header.pack(fill="x", padx=10, pady=10, before=self.scrollable_main_container)
        
        # ส่วนที่เหลือใช้ .pack() หรือ .grid() ภายใน selection_header ได้ตามปกติ
        # เพราะเป็นคนละ "กล่อง" กันแล้ว
        selection_header.grid_columnconfigure(1, weight=1)
        
        CTkLabel(selection_header, text=f"ผู้ดำเนินการ: {self.proxy_user_name} ({self.proxy_user_key})", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        sale_selection_frame = CTkFrame(selection_header, fg_color="transparent")
        sale_selection_frame.grid(row=0, column=1, sticky="e", padx=10, pady=10)
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

    def _on_sale_selected(self, selected_display_name):
        """Callback เมื่อมีการเลือกเซลส์จาก Dropdown"""
        if "- กรุณาเลือกเซลส์ -" in selected_display_name:
            self.sale_key_owner = None
            self._toggle_main_form(state="disabled")
        else:
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                # เปลี่ยนค่า sale_key ใน CommissionApp ให้เป็นของเซลส์ที่ถูกเลือก
                self.sale_key = self.sale_key_owner 
                self._toggle_main_form(state="normal")
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="normal")

    def _toggle_main_form(self, state):
        """เปิด/ปิด การใช้งานฟอร์มหลัก"""
        # ทำให้ฟังก์ชันนี้ปลอดภัยมากขึ้น โดยการเช็คว่า widget มีอยู่จริงหรือไม่
        widgets_to_toggle = [
            getattr(self, name, None) for name in 
            ['save_button', 'btn_clear', 'btn_edit', 'scrollable_main_container']
        ]
        for widget in widgets_to_toggle:
            if widget and hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                try:
                    if widget == self.scrollable_main_container:
                        for child in widget.winfo_children():
                            for grandchild in child.winfo_children():
                                grandchild.configure(state=state)
                    else:
                        widget.configure(state=state)
                except Exception:
                    pass

    def _gather_data_from_form(self):
        """
        (Override) รวบรวมข้อมูลจากฟอร์ม และใส่ Key ของผู้ทำรายการ (Proxy)
        และเจ้าของรายการ (Owner) ให้ถูกต้อง
        """
        form_data = super()._gather_data_from_form()
        form_data['sale_key'] = self.sale_key_owner         # เจ้าของ SO คือเซลส์ที่ถูกเลือก
        form_data['support_user_key'] = self.proxy_user_key  # ผู้คีย์ข้อมูลคือ HR/Sale Support ที่ Login
        return form_data

    # Override ฟังก์ชันที่ต้องมีการตรวจสอบการเลือกเซลส์ก่อน
    def _save_data(self):
        if not self.sale_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกเซลส์", "กรุณาเลือกเซลส์ที่ต้องการทำงานแทนจากเมนูด้านบนก่อน", parent=self)
            return
        super()._save_data()

    def _show_history(self):
        if not self.sale_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกเซลส์", "กรุณาเลือกเซลส์ก่อนดูประวัติ", parent=self)
            return
        super()._show_history()

    def _open_submit_dialog(self):
        if not self.sale_key_owner:
            messagebox.showwarning("ยังไม่ได้เลือกเซลส์", "กรุณาเลือกเซลส์ก่อนนำส่งข้อมูล", parent=self)
            return
        sale_name_owner = self.selected_sale_key_full.get().split(' (')[0]
        SubmitSODialog(self, self.app_container, self.sale_key_owner, sale_name_owner)