# sales_proxy_screen.py (ฉบับแก้ไข Layout: รับประกันการแสดงปุ่ม)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd
target_roles = ['Sales Support', 'Admin', 'Manager', 'Director']
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

        self.active_sales_list = self._get_all_active_sales()
        
        # รอให้ UI วาดเสร็จแล้วค่อยสร้าง Header พิเศษด้านบน
        self.after(10, self._setup_proxy_ui)

    def _setup_proxy_ui(self):
        """สร้าง UI พิเศษสำหรับหน้า Proxy"""
        # 1. สร้าง Header สีเหลือง
        self._create_sale_selection_header()
        
        # 2. ปิดการใช้งานฟอร์มหลักก่อน (จนกว่าจะเลือกเซลส์)
        self._toggle_main_form(state="disabled")
        
        # 3. [สำคัญ] บังคับให้สร้างและแสดงปุ่ม Action Frame ทันที
        if hasattr(self, 'action_frame'):
             # ถ้ามี action_frame อยู่แล้ว ให้เคลียร์และสร้างใหม่ เพื่อให้ปุ่มของเราเข้าไปอยู่ด้วย
             for widget in self.action_frame.winfo_children():
                 widget.destroy()
             self._populate_action_frame(self.action_frame)
        else:
             # ถ้ายังไม่มี (ซึ่งแปลกมาก เพราะ CommissionApp ควรสร้างให้) ให้สร้างใหม่
             self.action_frame = CTkFrame(self.scrollable_main_container) # หรือ parent ที่เหมาะสม
             self.action_frame.pack(fill="x", pady=10)
             self._populate_action_frame(self.action_frame)


    def _create_sale_selection_header(self):
        """สร้างแถบ Header สีเหลืองสำหรับเลือกเซลส์"""
        # ล้าง Header เดิมทิ้ง
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
    #  ★ Override ส่วนปุ่มด้านล่าง เพื่อเพิ่มปุ่มพิเศษให้ Sale Support ★
    # -----------------------------------------------------------
    def _populate_action_frame(self, parent):
        # 1. สร้างปุ่มมาตรฐาน (บันทึก, ล้างค่า, นำส่ง) โดยเรียกฟังก์ชันแม่
        # (CommissionApp._populate_action_frame จะสร้างปุ่มมาตรฐานให้)
        super()._populate_action_frame(parent)

        # 2. เพิ่มปุ่มพิเศษต่อท้าย สำหรับ Sale Support เท่านั้น
        if self.user_role == 'Sale Support':
            
            # เส้นคั่น
            separator = tk.Frame(parent, height=2, bd=1, relief="sunken")
            separator.pack(fill="x", padx=20, pady=(20, 10))

            # หัวข้อเครื่องมือ
            tool_label = CTkLabel(parent, text="เครื่องมือสำหรับ Sale Support:", font=CTkFont(size=14, weight="bold"), text_color="gray50")
            tool_label.pack(anchor="w", padx=20, pady=(0, 5))

            # ปุ่มสีม่วง (ย้ายเจ้าของ SO)
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
            # หา Sale Key จากชื่อที่เลือก
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                # อัปเดต key ของฟอร์มหลัก
                self.sale_key = self.sale_key_owner
                self.sale_name = selected_sale_data['sale_name']
                # เปิดใช้งานฟอร์ม
                self._toggle_main_form(state="normal")
                
                # --- [เพิ่ม] รีโหลดข้อมูลปุ่ม Action Frame ใหม่ เพื่อให้แน่ใจว่าปุ่มยังอยู่ ---
                if hasattr(self, 'action_frame'):
                     for widget in self.action_frame.winfo_children():
                         widget.destroy()
                     self._populate_action_frame(self.action_frame)
                     
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="disabled")
    
    def _toggle_main_form(self, state):
        if hasattr(self, 'scrollable_main_container') and self.scrollable_main_container.winfo_exists():
            self._recursive_toggle_state(self.scrollable_main_container, state)

    def _recursive_toggle_state(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            try:
                # ยกเว้น Action Frame ไม่ต้องปิด (เพื่อให้กดปุ่มย้าย SO ได้ตลอด)
                if child == getattr(self, 'action_frame', None):
                    continue
                
                child.configure(state=state)
            except Exception:
                pass
            if child.winfo_children():
                self._recursive_toggle_state(child, state)