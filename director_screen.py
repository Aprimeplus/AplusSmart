import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkButton

# --- Import หน้าจอทั้ง 4 ส่วน ---
from purchasing_manager_screen import PurchasingManagerScreen
from hr_screen import HRScreen
from sales_manager_screen import SalesManagerScreen
from management_report_screen import ManagementReportScreen

class DirectorScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # =========================================================
        # 1. Main Containers (สลับหน้าจอไปมา)
        # =========================================================
        self.home_frame = CTkFrame(self, fg_color="transparent")
        self.content_frame = CTkFrame(self, fg_color="transparent")

        self._build_home_menu()
        self.show_home() # เริ่มต้นด้วยการโชว์หน้า Home

    def _build_home_menu(self):
        """สร้างหน้าต่างเมนูหลัก (Big Picture)"""
        self.home_frame.grid_columnconfigure(0, weight=1)
        self.home_frame.grid_rowconfigure(1, weight=1)

        # --- Header ---
        header = CTkFrame(self.home_frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=30, pady=(30, 10))
        
        CTkLabel(header, text="📊 Management Dashboard", font=CTkFont(size=28, weight="bold")).pack(side="left")
        CTkLabel(header, text=f"ยินดีต้อนรับ, {self.user_name}", font=CTkFont(size=16), text_color="gray50").pack(side="left", padx=15, pady=(8,0))
        
        CTkButton(header, text="ออกจากระบบ", command=self.app_container.show_login_screen, 
                  fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", 
                  border_width=2, hover_color="#FFEBEE").pack(side="right")

        # --- Menu Container ---
        menu_container = CTkFrame(self.home_frame, fg_color="transparent")
        menu_container.grid(row=1, column=0, sticky="nsew", padx=30, pady=20)
        
        # จัดให้อยู่กึ่งกลาง
        menu_container.grid_columnconfigure((0, 1, 2), weight=1, uniform="col")
        menu_container.grid_rowconfigure((0, 1), weight=1)

        # --- สไตล์ของปุ่มการ์ด ---
        btn_kwargs = {
            "font": CTkFont(size=20, weight="bold"),
            "height": 180,
            "corner_radius": 15,
            "border_spacing": 10
        }

        # 1. การ์ดฝ่ายขาย (Sales)
        CTkButton(menu_container, text="📈\n\nฝ่ายขายและการอนุมัติ\n(Sales)", 
                  fg_color="#2563EB", hover_color="#1D4ED8", **btn_kwargs,
                  command=lambda: self.open_module("sales", "📈 ฝ่ายขายและการอนุมัติ")).grid(row=0, column=0, padx=15, pady=20, sticky="ew")

        # 2. การ์ดฝ่ายจัดซื้อ (Purchasing)
        CTkButton(menu_container, text="🛒\n\nฝ่ายจัดซื้อและสต๊อก\n(Purchasing & Inventory)", 
                  fg_color="#059669", hover_color="#047857", **btn_kwargs,
                  command=lambda: self.open_module("purchase", "🛒 ฝ่ายจัดซื้อและสต๊อก")).grid(row=0, column=1, padx=15, pady=20, sticky="ew")

        # 3. การ์ดฝ่ายบุคคล/บัญชี (HR & Accounting)
        CTkButton(menu_container, text="👥\n\nฝ่ายบุคคลและคอมมิชชั่น\n(HR & Commission)",
                  fg_color="#7C3AED", hover_color="#6D28D9", **btn_kwargs,
                  command=lambda: self.open_module("hr", "👥 ฝ่ายบุคคลและคอมมิชชั่น")).grid(row=0, column=2, padx=15, pady=20, sticky="ew")

        # 4. การ์ด Management Report
        CTkButton(menu_container, text="📑\n\nรายงานผู้บริหาร\n(Management Report)",
                  fg_color="#0F172A", hover_color="#1E293B", **btn_kwargs,
                  command=lambda: self.open_module("report", "📑 รายงานผู้บริหาร")).grid(row=1, column=0, padx=15, pady=20, sticky="ew")

    def show_home(self):
        """แสดงหน้า Home Menu และซ่อนหน้าเนื้อหา"""
        self.content_frame.grid_forget()
        
        # เคลียร์หน้าจอเดิมทิ้งเพื่อคืน Memory
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
        self.home_frame.grid(row=0, column=0, sticky="nsew")

    def open_module(self, module_name, title_text):
        """โหลดหน้าจอของแต่ละแผนกมาแสดง"""
        self.home_frame.grid_forget()
        self.content_frame.grid(row=0, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)

        # --- Top Navigation Bar ---
        top_bar = CTkFrame(self.content_frame, height=50, fg_color="gray90", corner_radius=0)
        top_bar.grid(row=0, column=0, sticky="ew")
        
        CTkButton(top_bar, text="⬅️ กลับหน้าเมนูหลัก", font=CTkFont(weight="bold"), 
                  fg_color="#4B5563", hover_color="#374151", height=36,
                  command=self.show_home).pack(side="left", padx=20, pady=10)
                  
        CTkLabel(top_bar, text=title_text, font=CTkFont(size=18, weight="bold"), text_color="#1F2937").pack(side="left", padx=10)

        # --- Module Container ---
        module_container = CTkFrame(self.content_frame, fg_color="transparent")
        module_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        # โชว์ "กำลังโหลด..." ให้เห็นก่อนทันที แล้วค่อยสร้างหน้าจอจริง (ซึ่งมัก query ฐานข้อมูลหนักๆ
        # จนบล็อก UI ชั่วขณะ) ในติ๊กถัดไป — ไม่ได้ทำให้โหลดเร็วขึ้น แต่กัน "เหมือนค้าง" เพราะอย่างน้อย
        # user เห็นว่าระบบตอบสนองแล้ว ไม่ใช่จอนิ่งเงียบๆ ระหว่างรอ
        loading_label = CTkLabel(module_container, text="⏳ กำลังโหลด...",
                                  font=CTkFont(size=15), text_color="#6B7280")
        loading_label.pack(expand=True)
        self.update_idletasks()

        def _build_module_screen():
            if not loading_label.winfo_exists():
                return  # user กดสลับหน้าหนีไปแล้วก่อนโหลดเสร็จ
            loading_label.destroy()

            # --- โหลด Class ตามที่เลือก ---
            if module_name == "sales":
                screen = SalesManagerScreen(
                    master=module_container,
                    app_container=self.app_container,
                    user_key=self.user_key,
                    user_name=self.user_name,
                    user_role=self.user_role
                )
            elif module_name == "purchase":
                screen = PurchasingManagerScreen(
                    master=module_container,
                    app_container=self.app_container,
                    user_key=self.user_key,
                    user_role=self.user_role
                )
            elif module_name == "hr":
                screen = HRScreen(
                    master=module_container,
                    app_container=self.app_container,
                    user_key=self.user_key,
                    user_name=self.user_name,
                    user_role=self.user_role
                )
            elif module_name == "report":
                screen = ManagementReportScreen(
                    master=module_container,
                    app_container=self.app_container,
                    user_key=self.user_key,
                    user_name=self.user_name,
                    user_role=self.user_role
                )
            else:
                return

            screen.pack(fill="both", expand=True)

        self.after(30, _build_module_screen)