# director_screen.py

import tkinter as tk
from customtkinter import CTkFrame, CTkTabview, CTkLabel, CTkFont

# Import หน้าจอที่เราต้องการจะนำมาใส่ในแท็บ
from purchasing_manager_screen import PurchasingManagerScreen
from hr_screen import HRScreen

class DirectorScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- 1. สร้าง Header ---
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 0))
        
        title = f"หน้าจอสำหรับกรรมการ: {self.user_name}"
        CTkLabel(header_frame, text=title, font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        # --- 2. สร้าง Tab View ---
        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # --- 3. เพิ่มแท็บต่างๆ ---
        self.manager_view_tab = self.tab_view.add("ภาพรวมและอนุมัติ (Manager View)")
        self.hr_view_tab = self.tab_view.add("วิเคราะห์ข้อมูล (HR View)")

        # --- 4. นำหน้าจอต่างๆ มาใส่ในแต่ละแท็บ ---
        
        # --- START: จุดที่แก้ไข ---
        # ลบ user_role="กรรมการ" ที่ซ้ำซ้อนออกไป
        self.manager_screen = PurchasingManagerScreen(
            self.manager_view_tab, 
            app_container=self.app_container, 
            user_key=self.user_key, 
            user_role=self.user_role # <-- เหลือไว้แค่ตัวนี้
        )
        # --- END: สิ้นสุดการแก้ไข ---
        self.manager_screen.pack(fill="both", expand=True)

        # สร้าง HRScreen ภายในแท็บที่สอง
        self.hr_screen = HRScreen(
            master=self.hr_view_tab,
            app_container=self.app_container,
            user_key=self.user_key,
            user_name=self.user_name,
            user_role=self.user_role
        )
        self.hr_screen.pack(fill="both", expand=True)
