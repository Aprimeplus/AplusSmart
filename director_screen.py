import tkinter as tk
from customtkinter import CTkFrame, CTkTabview, CTkLabel, CTkFont

# --- Import หน้าจอทั้ง 3 ส่วน ---
from purchasing_manager_screen import PurchasingManagerScreen
from hr_screen import HRScreen
from sales_manager_screen import SalesManagerScreen  # ✅ เพิ่มการ Import

class DirectorScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- 1. Header ---
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 0))
        
        title = f"Dashboard ผู้บริหารระดับสูง (Director): {self.user_name}"
        CTkLabel(header_frame, text=title, font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        # --- 2. สร้าง Tab View (เพิ่มเป็น 3 แท็บ) ---
        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # สร้างแท็บ
        self.sale_tab = self.tab_view.add("ยอดขายและการอนุมัติ (Sales View)") # ✅ แท็บใหม่
        self.purchase_tab = self.tab_view.add("จัดซื้อและค่าใช้จ่าย (Purchasing View)")
        self.hr_tab = self.tab_view.add("พนักงานและคอมมิชชั่น (HR View)")

        # --- 3. ติดตั้งหน้าจอลงในแต่ละแท็บ ---

        # 🟢 TAB 1: Sales Manager Screen (หน้านี้จะเห็น Dashboard 40/60 ที่เราเพิ่งแก้)
        self.sales_screen = SalesManagerScreen(
            master=self.sale_tab,
            app_container=self.app_container,
            user_key=self.user_key,
            user_name=self.user_name,
            user_role=self.user_role
        )
        self.sales_screen.pack(fill="both", expand=True)

        # 🔵 TAB 2: Purchasing Manager Screen
        self.purchase_screen = PurchasingManagerScreen(
            master=self.purchase_tab, 
            app_container=self.app_container, 
            user_key=self.user_key, 
            user_role=self.user_role
        )
        self.purchase_screen.pack(fill="both", expand=True)

        # 🟣 TAB 3: HR Screen
        self.hr_screen = HRScreen(
            master=self.hr_tab,
            app_container=self.app_container,
            user_key=self.user_key,
            user_name=self.user_name,
            user_role=self.user_role
        )
        self.hr_screen.pack(fill="both", expand=True)

        # ตั้งค่าแท็บเริ่มต้น (เลือกได้ว่าจะให้เปิดหน้าไหนก่อน)
        self.tab_view.set("ยอดขายและการอนุมัติ (Sales View)")