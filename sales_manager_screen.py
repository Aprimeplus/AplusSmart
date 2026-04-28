import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                               CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                               CTkOptionMenu, CTkRadioButton, CTkTabview, CTkCheckBox)
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import psycopg2.errors
import psycopg2.extras
import traceback
import utils
import os

# เพิ่ม matplotlib สำหรับกราฟ
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
matplotlib.use('TkAgg')

# --- นำเข้า Class ที่จำเป็น ---
from history_windows import SOPopupWindow, DeferralHistoryWindow, ManagerDeferApprovalDialog
from daily_report_widget import DailyReportWidget

STATUS_THAI_MAP = {
    'Draft': 'ฉบับร่าง',
    'Edited': 'แก้ไข/บันทึกร่าง',
    'Pending Sale Manager Approval': 'รอ ผจก.ฝ่ายขายอนุมัติ',
    'Rejected by SM': 'ผจก.ขาย ตีกลับ',
    'Pending PU': 'รอฝ่ายจัดซื้อรับงาน',
    'PO In Progress': 'จัดซื้อกำลังดำเนินการ',
    'Pending Approval': 'รออนุมัติ PO',
    'Approved': 'อนุมัติแล้ว',
    'Rejected': 'ถูกตีกลับให้แก้ไข',
    'PO Sent': 'สั่งซื้อ/เปิด PO เรียบร้อย',
    'Forwarded_To_HR': 'ส่งต่อให้ HR',
    'HR Verified': 'HR ตรวจสอบแล้ว',
    'Paid': 'จ่ายค่าคอมฯ แล้ว',
    'Defer Requested': 'HR ขอเลื่อนจ่าย',
    'Deferred': 'ถูกเลื่อนการจ่าย',
    'Cancelled': 'ยกเลิก',
    'Cancelled by PU': 'ยกเลิกโดยจัดซื้อ'
}

class ToastNotification(CTkToplevel):
    """หน้าต่างแจ้งเตือนมุมขวาล่าง เด้งโชว์แล้วหายไปเอง (ไม่บล็อกการทำงาน)"""
    def __init__(self, master, title, message, duration=10000, color="#16A34A"):
        super().__init__(master)
        
        # ลบกรอบหน้าต่างออก (ไร้ขอบ)
        self.overrideredirect(True)
        self.attributes("-topmost", True) # ให้อยู่บนสุดเสมอ
        
        # ตั้งค่าสีและ Layout
        self.configure(fg_color=color)
        
        # จัดข้อความ
        CTkLabel(self, text=title, font=CTkFont(size=14, weight="bold"), text_color="white").pack(padx=20, pady=(10, 2), anchor="w")
        CTkLabel(self, text=message, font=CTkFont(size=12), text_color="white", justify="left").pack(padx=20, pady=(0, 10), anchor="w")

        self.update_idletasks()
        width = self.winfo_reqwidth()
        height = self.winfo_reqheight()

        # คำนวณตำแหน่งมุมขวาล่างของหน้าจอ
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # วางที่มุมขวาล่าง (ห่างขอบนิดหน่อย)
        x = screen_width - width - 30
        y = screen_height - height - 60 
        
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # ทำให้คลิกที่แจ้งเตือนแล้วหายไปทันที (เผื่อไม่อยากรอ 10 วิ)
        self.bind("<Button-1>", lambda e: self.destroy())
        for child in self.winfo_children():
            child.bind("<Button-1>", lambda e: self.destroy())

        # ตั้งเวลาทำลายตัวเอง (10 วินาที = 10000 ms)
        self.after(duration, self.destroy)

class SalesManagerScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role
        
        self.label_font = CTkFont(size=14, weight="bold")
        self.entry_font = CTkFont(size=14)

        self.label_font_bold = CTkFont(size=14, weight="bold")
        self.header_font_table = CTkFont(size=16, weight="bold")
        
        # --- เตรียมตัวแปรสำหรับ SOPopupWindow ---
        self.so_popup = None
        self._so_create_string_vars()
        self.sale_theme = self.app_container.THEME.get("sale", {"bg": "white", "primary": "#3B82F6"})
        
        self.pg_engine = self.app_container.pg_engine
        
        # ตัวแปรสำหรับ Filter กราฟ
        self._init_chart_filters()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 

        # --- 1. Header ---
        self._create_header()

        # --- 2. TabView ---
        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # ✅ เพิ่มแท็บ 1: รายการรออนุมัติ (เป็นหน้าแรก)
        self.approval_tab = self.tab_view.add("🗳️ รายการรออนุมัติ (SM Approval)")
        self.defer_approval_tab = self.tab_view.add("⏳ คำขอเลื่อนรอบคอมฯ (Defer)")
        # แท็บเดิม
        self.daily_report_tab = self.tab_view.add("📅 รายงานประจำวัน (SO Report)")
        self.master_tab = self.tab_view.add("🛠️ ค้นหาและจัดการ (Master)")
        self.cancelled_tab = self.tab_view.add("❌ ยกเลิก SO (Cancel)")

        # สร้างเนื้อหาในแต่ละ Tab
        self._create_approval_tab(self.approval_tab)
        self._create_defer_approval_tab(self.defer_approval_tab)
        self._create_daily_report_widget(self.daily_report_tab) 
        self._create_master_tab(self.master_tab)           
        self._create_cancelled_so_tab(self.cancelled_tab) 
        
        # ตั้งค่าหน้าแรกที่เปิดขึ้นมา
        self.tab_view.set("🗳️ รายการรออนุมัติ (SM Approval)")

        self.known_pending_so_ids = set() # เก็บ ID ของ SO ที่เคยแจ้งเตือนไปแล้ว
        self.reminder_timer_count = 0     # ตัวนับเวลาสำหรับเตือนทุก 10 นาที
        self.noti_job_id = None
        
        # เริ่มทำงานระบบ Noti เบื้องหลัง (หน่วงเวลา 3 วินาทีหลังเปิดหน้าจอ)
        self.after(3000, self._start_notification_system)
        
        # ดักจับตอนปิดหน้าจอให้หยุด Noti ด้วย
        self.bind("<Destroy>", self._on_destroy)
    
    def _create_cancelled_so_tab(self, parent_tab):
        # --- Grid Setup: แบ่งหน้าจอเป็น 2 ส่วน (บน-เล็ก / ล่าง-ใหญ่) ---
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(0, weight=0) # ส่วนค้นหา (ความสูงคงที่)
        parent_tab.grid_rowconfigure(1, weight=1) # ส่วนตาราง (ขยายเต็มที่)

        # =========================================================
        #  SECTION 1: Professional Action Bar (แถบเครื่องมือด้านบน)
        # =========================================================
        action_bar = CTkFrame(parent_tab, height=60, fg_color=("gray90", "gray16"), corner_radius=6)
        action_bar.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        
        action_bar.grid_columnconfigure(3, weight=1) 
        
        CTkLabel(action_bar, text="🔎 ค้นหา SO:", font=self.label_font_bold).grid(row=0, column=0, padx=(20, 5), pady=10, sticky="w")
        
        self.cancel_search_entry = CTkEntry(action_bar, placeholder_text="ระบุเลข SO... (เช่น SO6701-001)", width=250, height=34)
        self.cancel_search_entry.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        self.cancel_search_entry.bind("<Return>", lambda e: self._search_so_to_cancel())
        
        CTkButton(action_bar, text="ค้นหา", command=self._search_so_to_cancel, 
                  width=100, height=34, fg_color="#3B82F6", hover_color="#2563EB", font=self.label_font_bold).grid(row=0, column=2, padx=10, pady=10, sticky="w")

        # Inline Result Area
        self.inline_result_frame = CTkFrame(action_bar, fg_color="transparent", height=34)
        self.inline_result_frame.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        CTkButton(action_bar, text="⟳ รีเฟรช", command=self._load_cancelled_so_history, 
                  width=90, height=34, fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).grid(row=0, column=4, padx=20, pady=10, sticky="e")

        # =========================================================
        #  SECTION 2: Full-Width History Table (ตารางเต็มจอ)
        # =========================================================
        table_container = CTkFrame(parent_tab, fg_color="transparent")
        table_container.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        
        header_row = CTkFrame(table_container, fg_color="transparent", height=30)
        header_row.pack(fill="x", pady=(0, 5))
        CTkLabel(header_row, text="📜 ประวัติรายการที่ถูกยกเลิก (Cancelled History)", font=self.header_font_table, text_color="#EF4444").pack(side="left")

        self.cancelled_history_frame = CTkFrame(table_container, fg_color="transparent")
        self.cancelled_history_frame.pack(fill="both", expand=True)

        self.after(100, self._load_cancelled_so_history)

    def _start_notification_system(self):
        """ลูปตรวจสอบ Noti ทุกๆ 1 นาที"""
        self._check_pending_approvals()
        
        # ลูปเรียกตัวเองทุกๆ 60 วินาที (60,000 ms)
        self.noti_job_id = self.after(60000, self._start_notification_system)

    def _check_pending_approvals(self):
        """เช็ค DB ว่ามี SO ใหม่ หรือต้องทวงงานหรือไม่"""
        try:
            query = """
                SELECT id, so_number, customer_name 
                FROM commissions 
                WHERE status = 'Pending Sale Manager Approval' AND is_active = 1
            """
            df = pd.read_sql_query(query, self.pg_engine)
            
            current_pending_ids = set(df['id'].tolist())
            pending_count = len(current_pending_ids)
            
            # 1. เช็คว่ามี SO "ใหม่" เข้ามาหรือไม่ (ID ที่เพิ่งโผล่มาใหม่)
            new_so_ids = current_pending_ids - self.known_pending_so_ids
            
            if new_so_ids:
                # มีงานใหม่เข้า! เด้งแจ้งเตือนแบบ 10 วินาที (สีเขียว/ฟ้า)
                new_count = len(new_so_ids)
                ToastNotification(
                    self.winfo_toplevel(), 
                    title="🔔 มี SO ใหม่รออนุมัติ!", 
                    message=f"มีคำขออนุมัติ SO เข้ามาใหม่จำนวน {new_count} รายการ\nกรุณาตรวจสอบในแท็บ 'รายการรออนุมัติ'",
                    duration=10000, 
                    color="#2563EB" # สีฟ้า
                )
                # อัปเดตรายการที่รู้จักแล้ว และรีเฟรชตารางโชว์งานใหม่
                self.known_pending_so_ids = current_pending_ids
                
                # --- ย้ายบรรทัดนี้ไปไว้ใน after() เพื่อไม่ให้ขัดจังหวะ ---
                self.after(100, self._load_approval_data)
                
            # 2. ระบบทวงงาน (Reminder) ถ้ายังมีงานค้าง
            elif pending_count > 0:
                self.reminder_timer_count += 1
                
                # เช็คทุก 1 นาที ดังนั้น 10 รอบ = 10 นาที
                if self.reminder_timer_count >= 10: 
                    ToastNotification(
                        self.winfo_toplevel(), 
                        title="⚠️ แจ้งเตือนงานค้าง!", 
                        message=f"คุณยังมี SO รอการอนุมัติค้างอยู่ {pending_count} รายการ",
                        duration=10000, 
                        color="#EA580C" # สีส้ม
                    )
                    self.reminder_timer_count = 0 # รีเซ็ตตัวนับ
            else:
                # ไม่มีงานค้างเลย รีเซ็ตเวลา
                self.reminder_timer_count = 0
                self.known_pending_so_ids.clear()
                
        except Exception as e:
            print(f"Noti System Error: {e}")

    def _on_destroy(self, event):
        """หยุด Loop เมื่อปิดหน้าจอ"""
        if hasattr(event, 'widget') and event.widget is self:
            if self.noti_job_id:
                self.after_cancel(self.noti_job_id)

    def _init_chart_filters(self):
        """สร้างตัวแปรสำหรับ Filter กราฟ"""
        now = datetime.now()
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                           "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.chart_month_var = tk.StringVar(value=self.thai_months[now.month - 1])
        self.chart_year_var = tk.StringVar(value=str(now.year + 543))

    def _create_header(self):
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        CTkLabel(header_frame, text=f"Sale Manager Dashboard: {self.user_name}", font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        button_frame = CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side="right", padx=10)
        
        # 🟢 [เพิ่มตรงนี้] ปุ่ม Export ข้อมูลแบบมี Popup
        CTkButton(button_frame, text="📥 Export Data SO", 
                  fg_color="#10B981", hover_color="#059669", font=CTkFont(weight="bold"),
                  command=self._open_sm_export_dialog).pack(side="left", padx=(0, 10))
        
        CTkButton(button_frame, text="🔄 Refresh All", command=self._refresh_all_tabs).pack(side="left", padx=5)
        
        CTkButton(button_frame, text="ออกจากระบบ", command=self.app_container.show_login_screen, 
                  fg_color="transparent", border_color="#D32F2F", 
                  text_color="#D32F2F", border_width=2, 
                  hover_color="#FFEBEE").pack(side="left", padx=5)

    # 🟢 [เพิ่มฟังก์ชันนี้เข้าไปในคลาส SalesManagerScreen ด้วย]
    def _open_sm_export_dialog(self):
        """เรียกเปิดหน้าต่าง Popup สำหรับ Export"""
        SMExportDialog(self, self.app_container)

    def _refresh_all_tabs(self):
        """รีโหลดข้อมูลทุกแท็บ"""
        self._load_approval_data() # รีโหลดหน้าอนุมัติ
        self._update_rejection_chart() # รีโหลดกราฟ
        
        if hasattr(self, 'daily_report_widget'):
            self.daily_report_widget.load_report_data()
            if hasattr(self.daily_report_widget, 'dashboard_view'):
                 self.daily_report_widget.dashboard_view._update_chart()
        
        self._load_master_data() # รีโหลด Master Tab

    # =========================================================================
    # ✅ NEW TAB: SM APPROVAL พร้อม Dashboard กราฟ
    # =========================================================================
    def _create_approval_tab(self, parent):
        """สร้างแท็บรออนุมัติ แบ่งเป็น Dashboard 40% + List 60%"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=3)  # Dashboard ~30%
        parent.grid_rowconfigure(1, weight=7)  # List ~70%

        # ========== ส่วนบน 40%: Dashboard กราฟ ==========
        dashboard_container = CTkFrame(parent, fg_color="#F3F4F6", corner_radius=10)
        dashboard_container.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="nsew")
        dashboard_container.grid_columnconfigure(0, weight=1)
        dashboard_container.grid_rowconfigure(1, weight=1)

        # Header + Filters
        chart_header = CTkFrame(dashboard_container, fg_color="transparent")
        chart_header.grid(row=0, column=0, sticky="ew", padx=15, pady=10)
        
        CTkLabel(chart_header, text="📊 สถิติการตีกลับงานรายบุคคล", 
                font=CTkFont(size=16, weight="bold")).pack(side="left")

        filter_frame = CTkFrame(chart_header, fg_color="transparent")
        filter_frame.pack(side="right")
        
        # ✅ ปุ่ม Export Excel
        CTkButton(filter_frame, 
                 text="📥 Export Excel",
                 width=120,
                 height=32,
                 fg_color="#16A34A",
                 hover_color="#15803D",
                 font=CTkFont(size=12, weight="bold"),
                 command=self._export_rejection_to_excel).pack(side="left", padx=(0, 10))
        
        CTkLabel(filter_frame, text="รอบเดือน:", font=CTkFont(size=12)).pack(side="left", padx=5)
        CTkOptionMenu(filter_frame, variable=self.chart_month_var, values=self.thai_months, 
                      width=110, height=32,
                      command=lambda e: self._update_rejection_chart()).pack(side="left", padx=3)
        
        years = [str(y + 543) for y in range(datetime.now().year - 1, datetime.now().year + 2)]
        CTkOptionMenu(filter_frame, variable=self.chart_year_var, values=years, 
                      width=90, height=32,
                      command=lambda e: self._update_rejection_chart()).pack(side="left", padx=3)

        # พื้นที่กราฟ
        self.chart_area = CTkFrame(dashboard_container, fg_color="white", corner_radius=8)
        self.chart_area.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")

        # ========== ส่วนล่าง 60%: รายการ SO ==========
        list_container = CTkFrame(parent, fg_color="transparent")
        list_container.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew")
        list_container.grid_columnconfigure(0, weight=1)
        list_container.grid_rowconfigure(1, weight=1)

        # Search Bar
        search_frame = CTkFrame(list_container, fg_color="#F3F4F6", corner_radius=8, height=60)
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        search_frame.grid_propagate(False)
        
        CTkLabel(search_frame, text="🗳️ รายการรออนุมัติ", 
                font=CTkFont(size=14, weight="bold")).pack(side="left", padx=20)

        search_right = CTkFrame(search_frame, fg_color="transparent")
        search_right.pack(side="right", padx=20)
        
        self.approval_search_var = tk.StringVar()
        self.approval_search_entry = CTkEntry(
            search_right, 
            placeholder_text="🔍 ค้นหา SO หรือชื่อลูกค้า...", 
            width=300,
            height=36,
            textvariable=self.approval_search_var
        )
        self.approval_search_entry.pack(side="left", padx=(0, 8))
        self.approval_search_entry.bind("<Return>", lambda e: self._load_approval_data())

        CTkButton(search_right, text="ค้นหา", width=80, height=36,
                  command=self._load_approval_data).pack(side="left", padx=3)
        CTkButton(search_right, text="🔄 รีเฟรช", width=80, height=36, fg_color="gray",
                  command=self._load_approval_data).pack(side="left", padx=3)
        
        # รายการ SO
        self.approval_results_frame = CTkScrollableFrame(list_container, 
                                                         fg_color="white",
                                                         corner_radius=8)
        self.approval_results_frame.grid(row=1, column=0, sticky="nsew")
        self.approval_results_frame.grid_columnconfigure(0, weight=1)

        # โหลดข้อมูลครั้งแรก
        self.after(200, lambda: [self._update_rejection_chart(), self._load_approval_data()])

    def _update_rejection_chart(self):
        """วาดกราฟแท่งแนวนอนแสดงจำนวนการตีกลับของแต่ละ Sale (KPI)"""
        for widget in self.chart_area.winfo_children():
            widget.destroy()

        try:
            plt.close('all')
            try:
                plt.rcParams['font.family'] = 'TH Sarabun New'
            except:
                plt.rcParams['font.family'] = 'Tahoma'
            
            month_idx = self.thai_months.index(self.chart_month_var.get()) + 1
            year_val = int(self.chart_year_var.get()) - 543

            # 🟢 แก้ไขคิวรี่: ใช้ SUM(sm_reject_count) ดึงยอดตีกลับสะสมทั้งหมดของเดือนนั้น
            query = """
                SELECT u.sale_name, SUM(COALESCE(c.sm_reject_count, 0)) as reject_count
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.commission_month = %s 
                  AND c.commission_year = %s
                  AND c.is_active = 1
                  AND c.sm_reject_count > 0
                GROUP BY u.sale_name
                ORDER BY reject_count ASC
            """
            
            df = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year_val))
            
            if df.empty:
                empty_frame = CTkFrame(self.chart_area, fg_color="transparent")
                empty_frame.pack(expand=True, pady=30)
                CTkLabel(empty_frame, text="✅", font=CTkFont(size=48)).pack()
                CTkLabel(empty_frame, text="ไม่มีข้อมูลความผิดพลาดในรอบเดือนนี้ (KPI ดีเยี่ยม)",
                        font=CTkFont(size=15, weight="bold"), text_color="#16A34A").pack(pady=(5, 0))
                return

            num_sales = len(df)
            fig_height = max(2.0, min(3.5, num_sales * 0.8))  # จำกัดความสูงไว้ที่ 3.5 นิ้ว
            fig, ax = plt.subplots(figsize=(9.5, fig_height), dpi=80)  # ลด dpi ลง

            bars = ax.barh(df['sale_name'], df['reject_count'], color='#EF4444', height=0.65) # เปลี่ยนสีเป็นแดงสถิติ KPI
            
            ax.set_title(f"KPI สถิติความผิดพลาด (รอบ {self.chart_month_var.get()} {self.chart_year_var.get()})",
                        fontweight='bold', fontsize=16, pad=15)
            ax.set_xlabel("จำนวนครั้งที่ถูกตีกลับสะสม", fontsize=14, fontweight='bold')
            ax.set_ylabel("เซลล์", fontsize=14, fontweight='bold')
            
            ax.tick_params(axis='both', which='major', labelsize=13)
            ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
            
            ax.grid(axis='x', alpha=0.3, linestyle='--', linewidth=1.2)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_linewidth(1.5)
            ax.spines['bottom'].set_linewidth(1.5)

            for bar in bars:
                width = bar.get_width()
                if width > 0:
                    ax.text(width + 0.15, bar.get_y() + bar.get_height()/2,
                           f'{int(width)} ครั้ง', va='center', fontsize=13, fontweight='bold')

            plt.tight_layout(pad=1.5)
            canvas = FigureCanvasTkAgg(fig, master=self.chart_area)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="x", expand=False, padx=10, pady=5)
            
            plt.close(fig)
            
        except Exception as e:
            print(f"Chart Error: {traceback.format_exc()}")

    def _export_rejection_to_excel(self):
        """Export ข้อมูลสถิติ KPI การตีกลับเป็น Excel"""
        try:
            month_idx = self.thai_months.index(self.chart_month_var.get()) + 1
            year_val = int(self.chart_year_var.get()) - 543
            
            # 🟢 อัปเดตคิวรี่: ดึงข้อมูล SO ที่มีประวัติเคยถูกตีกลับ
            query = """
                SELECT 
                    u.sale_name as "ชื่อเซลล์",
                    c.so_number as "เลขที่ SO",
                    c.customer_name as "ชื่อลูกค้า",
                    c.sales_service_amount as "ยอดขาย (บาท)",
                    c.sm_reject_count as "จำนวนครั้งที่ถูกตีกลับ",
                    c.rejection_reason as "เหตุผลล่าสุดที่ถูกตีกลับ"
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.commission_month = %s 
                  AND c.commission_year = %s
                  AND c.is_active = 1
                  AND c.sm_reject_count > 0
                ORDER BY u.sale_name ASC, c.sm_reject_count DESC
            """
            
            df_detail = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year_val))
            
            if df_detail.empty:
                messagebox.showinfo("แจ้งเตือน", f"ไม่มีสถิติการตีกลับในรอบ {self.chart_month_var.get()} {self.chart_year_var.get()}")
                return
            
            # สร้างหน้าสรุป KPI รวบยอด
            summary_df = df_detail.groupby('ชื่อเซลล์')['จำนวนครั้งที่ถูกตีกลับ'].sum().reset_index()
            summary_df = summary_df.sort_values('จำนวนครั้งที่ถูกตีกลับ', ascending=False)
            
            default_filename = f"KPI_รายงานการตีกลับ_{self.chart_month_var.get()}_{self.chart_year_var.get()}.xlsx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_filename,
                title="บันทึกรายงานสถิติ KPI การตีกลับ"
            )
            
            if not file_path: return
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='สรุป KPI รายบุคคล', index=False)
                df_detail.to_excel(writer, sheet_name='รายละเอียด SO ที่ผิดพลาด', index=False)
                
                # จัด Format Excel
                worksheet1 = writer.sheets['สรุป KPI รายบุคคล']
                worksheet1.column_dimensions['A'].width = 30
                worksheet1.column_dimensions['B'].width = 25
                
                worksheet2 = writer.sheets['รายละเอียด SO ที่ผิดพลาด']
                worksheet2.column_dimensions['A'].width = 25
                worksheet2.column_dimensions['B'].width = 20
                worksheet2.column_dimensions['C'].width = 35
                worksheet2.column_dimensions['D'].width = 18
                worksheet2.column_dimensions['E'].width = 25
                worksheet2.column_dimensions['F'].width = 50
                
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูล KPI เรียบร้อย!\nบันทึกที่: {file_path}")
            if os.name == 'nt': os.startfile(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Export Failed: {str(e)}")
            print(traceback.format_exc())

    def _load_approval_data(self):
        """ดึงรายการรออนุมัติ พร้อมระบบค้นหา"""
        for widget in self.approval_results_frame.winfo_children(): 
            widget.destroy()
        
        search_txt = self.approval_search_var.get().strip().upper()
        
        try:
            # Base Query
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status, u.sale_name, c.sales_service_amount
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Pending Sale Manager Approval' AND c.is_active = 1
            """
            params = []

            # ✅ ถ้ามีการพิมพ์ค้นหา ให้เพิ่มเงื่อนไข SQL
            if search_txt:
                term = search_txt.replace("SO", "")
                query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                params.extend([f"%{term}%", f"%{term}%"])

            query += " ORDER BY c.timestamp ASC"
            
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                # Empty State
                empty_frame = CTkFrame(self.approval_results_frame, fg_color="transparent")
                empty_frame.pack(expand=True, pady=50)
                
                icon = "✅" if not search_txt else "🔍"
                CTkLabel(empty_frame, text=icon, font=CTkFont(size=48)).pack(pady=(0, 10))
                
                msg = "ไม่มีรายการรออนุมัติในขณะนี้" if not search_txt else "ไม่พบรายการที่ตรงกับคำค้นหา"
                CTkLabel(empty_frame, text=msg,
                        font=CTkFont(size=14),
                        text_color="#6B7280").pack()
                return

            for _, row in df.iterrows():
                self._create_so_card(self.approval_results_frame, row.to_dict(), is_approval_mode=True)
                
        except Exception as e:
            print(f"Load Approval Error: {e}")
            CTkLabel(self.approval_results_frame,
                    text=f"⚠️ เกิดข้อผิดพลาด: {str(e)[:100]}",
                    font=CTkFont(size=12),
                    text_color="#DC2626").pack(pady=20)

    # =========================================================================
    # TAB 1: DAILY REPORT WIDGET
    # =========================================================================
    def _create_daily_report_widget(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.daily_report_widget = DailyReportWidget(parent, self.app_container)
        self.daily_report_widget.pack(fill="both", expand=True)

    # =========================================================================
    # TAB 2: MASTER EDIT & SEARCH
    # =========================================================================
    def _create_master_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        filter_frame = CTkFrame(parent_tab)
        filter_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        
        current_year = datetime.now().year
        self.mst_year_var = tk.StringVar(value="ทุกปี")
        self.mst_month_var = tk.StringVar(value="ทุกเดือน")
        self.mst_day_var = tk.StringVar(value="ทุกวัน")
        self.mst_sale_var = tk.StringVar(value="All Sales")
        
        years = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 3, -1)]
        months = ["ทุกเดือน"] + ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                                  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        days = ["ทุกวัน"] + [str(d) for d in range(1, 32)]
        sales_list = self._get_sale_list()

        CTkLabel(filter_frame, text="ปี:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_year_var, values=years, width=75, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เดือน:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_month_var, values=months, width=110, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="วัน:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_day_var, values=days, width=70, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="Sale:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_sale_var, values=sales_list, width=120, command=self._load_master_data).pack(side="left", padx=5)

        self.sm_master_search_entry = CTkEntry(filter_frame, font=self.entry_font, placeholder_text="SO / ลูกค้า...", width=150)
        self.sm_master_search_entry.pack(side="left", padx=(15, 5), fill="x", expand=True)
        self.sm_master_search_entry.bind("<Return>", lambda e: self._load_master_data())
        
        CTkButton(filter_frame, text="🔍 ค้นหา", width=80, command=self._load_master_data).pack(side="left", padx=5)
        CTkButton(filter_frame, text="↺ รีเซ็ต", width=60, fg_color="gray", command=self._reset_master_filter).pack(side="left", padx=5)

        self.sm_master_results_frame = CTkScrollableFrame(parent_tab, label_text="รายการ SO ทั้งหมด (จัดการประวัติ)")
        self.sm_master_results_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.sm_master_results_frame.grid_columnconfigure(0, weight=1)

        self.after(200, self._load_master_data)

    def _get_sale_list(self):
        try:
            df = pd.read_sql_query("SELECT sale_key FROM sales_users WHERE role='Sale' ORDER BY sale_key", self.pg_engine)
            return ["All Sales"] + df['sale_key'].tolist()
        except: return ["All Sales"]

    def _reset_master_filter(self):
        self.mst_year_var.set("ทุกปี")
        self.mst_month_var.set("ทุกเดือน")
        self.mst_day_var.set("ทุกวัน")
        self.mst_sale_var.set("All Sales")
        self.sm_master_search_entry.delete(0, "end")
        self._load_master_data()

    def _load_master_data(self, event=None):
        for widget in self.sm_master_results_frame.winfo_children(): widget.destroy()
        try:
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status, u.sale_name, c.sales_service_amount
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.is_active = 1
            """
            params = []
            if self.mst_year_var.get() != "ทุกปี":
                query += " AND EXTRACT(YEAR FROM c.timestamp::timestamp) = %s"
                params.append(int(self.mst_year_var.get()))
            if self.mst_month_var.get() != "ทุกเดือน":
                thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
                params.append(thai_months.index(self.mst_month_var.get()) + 1)
                query += " AND EXTRACT(MONTH FROM c.timestamp::timestamp) = %s"
            if self.mst_day_var.get() != "ทุกวัน":
                query += " AND EXTRACT(DAY FROM c.timestamp::timestamp) = %s"; params.append(int(self.mst_day_var.get()))
            if self.mst_sale_var.get() != "All Sales":
                query += " AND c.sale_key = %s"; params.append(self.mst_sale_var.get())
            
            search_txt = self.sm_master_search_entry.get().strip().upper()
            if search_txt:
                term = search_txt.replace("SO", "")
                query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                params.extend([f"%{term}%", f"%{term}%"])

            query += " ORDER BY c.timestamp DESC LIMIT 20" 
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                CTkLabel(self.sm_master_results_frame, text="ไม่พบข้อมูล").pack(pady=20)
                return
            for _, row in df.iterrows():
                self._create_so_card(self.sm_master_results_frame, row.to_dict())
        except Exception as e: messagebox.showerror("Error", str(e))

    # =========================================================================
    # SHARED: SO CARD & ACTIONS
    # =========================================================================
    def _create_so_card(self, parent, so_data, is_approval_mode=False):
        """สร้าง Card แสดงข้อมูล SO พร้อมปุ่มต่างๆ"""
        so_id = so_data['id']
        so_number = so_data['so_number']
        status = so_data.get('status', 'N/A')
        amount = so_data.get('sales_service_amount', 0) or 0

        status_colors = {
            'PO In Progress': '#E0F2FE', 'Approved': '#DCFCE7', 'Paid': '#D1FAE5',
            'Rejected by SM': '#FEE2E2', 'Cancelled': '#F3F4F6', 'Draft': '#FEF3C7',
            'Pending Sale Manager Approval': '#FEF9C3'
        }
        bg_color = status_colors.get(status, "#FFFFFF")

        card = CTkFrame(parent, border_width=2, border_color="#E5E7EB",
                       corner_radius=10, fg_color=bg_color, height=70)
        card.pack(fill="x", padx=8, pady=6)
        card.pack_propagate(False)
        
        # ข้อมูล
        info_frame = CTkFrame(card, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, padx=15, pady=10)
        
        line1 = f"📋 SO: {so_number}  |  👤 {so_data.get('customer_name', '-')}"
        CTkLabel(info_frame, text=line1, 
                font=CTkFont(size=13, weight="bold"),
                anchor="w").pack(anchor="w")
        
        # 🟢 แปลงสถานะเป็นภาษาไทยเฉพาะตอนโชว์ข้อความ (สีพื้นหลังจะได้ไม่พัง)
        status_th = STATUS_THAI_MAP.get(status, status)
        
        line2 = f"🎯 เซลล์: {so_data.get('sale_name', '-')}  |  💰 {amount:,.2f} บาท  |  📌 {status_th}"
        CTkLabel(info_frame, text=line2,
                font=CTkFont(size=12),
                text_color="#6B7280",
                anchor="w").pack(anchor="w", pady=(3, 0))

        # ปุ่ม
        btn_frame = CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=12, pady=10)

        if is_approval_mode:
            CTkButton(btn_frame, text="✅ อนุมัติ", width=85, height=32,
                      fg_color="#16A34A", hover_color="#15803D",
                      font=CTkFont(size=12, weight="bold"),
                      command=lambda: self._approve_so(so_id, so_number)).pack(side="left", padx=3)

        CTkButton(btn_frame, text="🛠️ แก้ไข", width=80, height=32,
                  fg_color="#4F46E5", hover_color="#4338CA",
                  font=CTkFont(size=12, weight="bold"),
                  command=lambda: self._open_so_editor_for_sm(so_number)).pack(side="left", padx=3)

        if status not in ['Cancelled', 'Rejected by SM']:
            CTkButton(btn_frame, text="❌ ตีกลับ", width=80, height=32,
                      fg_color="#DC2626", hover_color="#B91C1C",
                      font=CTkFont(size=12, weight="bold"),
                      command=lambda: self._reject_so(so_id, so_number)).pack(side="left", padx=3)

    # ✅ ฟังก์ชันอนุมัติ SO
    def _approve_so(self, so_id, so_number):
        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการอนุมัติ SO: {so_number} ใช่หรือไม่?"):
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'Pending PU', 
                        approver_sale_manager_key = %s, 
                        approval_date_sale_manager = CURRENT_TIMESTAMP,
                        claim_timestamp = NULL
                    WHERE id = %s
                """, (self.user_key, so_id))
                
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Staff' AND status = 'Active'")
                pu_keys = [row[0] for row in cursor.fetchall()]
                
                for pu_key in pu_keys:
                    cursor.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id)
                        VALUES (%s, %s, FALSE, %s)
                    """, (pu_key, f"มี SO ใหม่ ({so_number}) ผ่านการอนุมัติแล้ว รอคุณ Claim งาน", so_id))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"อนุมัติ SO: {so_number} เรียบร้อยแล้ว")
            self._refresh_all_tabs()
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"Approve Failed: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _reject_so(self, so_id, so_number):
        """เปิดหน้าต่างเลือกเหตุผลการตีกลับ"""
        def save_rejection(reason):
            conn = None
            try:
                conn = self.app_container.get_connection()
                with conn.cursor() as cursor:
                    # 🟢 แก้ไขตรงนี้: เพิ่ม sm_reject_count = COALESCE(sm_reject_count, 0) + 1
                    cursor.execute("""
                        UPDATE commissions 
                        SET status = 'Rejected by SM', 
                            rejection_reason = %s,
                            sm_reject_count = COALESCE(sm_reject_count, 0) + 1 
                        WHERE id = %s
                    """, (reason, so_id))
                    
                    cursor.execute("SELECT sale_key FROM commissions WHERE id = %s", (so_id,))
                    res = cursor.fetchone()
                    
                    if res:
                        cursor.execute("""
                            INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) 
                            VALUES (%s, %s, FALSE, %s)
                        """, (res[0], f"SO: {so_number} ถูกตีกลับ: {reason}", so_id))
                
                conn.commit()
                messagebox.showinfo("สำเร็จ", f"ตีกลับ SO: {so_number} เรียบร้อยแล้ว\nระบบบันทึกสถิติความผิดพลาด +1 ครั้ง")
                self._refresh_all_tabs()
                
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Error", f"Reject Failed: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)

        SORejectionDialog(self, so_number, save_rejection)

    def _open_so_editor_for_sm(self, so_number):
        if self.so_popup is not None and self.so_popup.winfo_exists():
            self.so_popup.focus()
            return
        try:
            so_df = pd.read_sql_query("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", 
                                     self.pg_engine, params=(so_number,))
            if so_df.empty:
                messagebox.showerror("Error", "ไม่พบข้อมูล SO")
                return
            
            def _refresh_on_save():
                self._refresh_all_tabs()
            
            self.so_popup = SOPopupWindow(
                master=self,
                app_container=self.app_container,
                sales_data=so_df.iloc[0].to_dict(),
                so_shared_vars=self.so_shared_vars,
                sale_theme=self.sale_theme,
                on_save_callback=_refresh_on_save
            )
        except Exception as e:
            messagebox.showerror("Error", f"Open Editor Failed: {e}")
            print(traceback.format_exc())

    def _so_create_string_vars(self):
        """สร้าง StringVars สำหรับหน้าจอแก้ไข SO"""
        self.so_shared_vars = {}
        now = datetime.now()
        thai_months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        self.so_shared_vars.update({
            'thai_months': thai_months_list,
            'thai_month_map': {name: i + 1 for i, name in enumerate(thai_months_list)},
            'customer_type_var': tk.StringVar(value="ลูกค้าเก่า"),
            'credit_term_var': tk.StringVar(value="เงินสด"),
            'commission_month_var': tk.StringVar(value=thai_months_list[now.month - 1]),
            'commission_year_var': tk.StringVar(value=str(now.year + 543)),
            'payment1_percent_var': tk.StringVar(value="ระบุยอดเอง"),
            'payment2_percent_var': tk.StringVar(value="ระบุยอดเอง"),
            'delivery_type_var': tk.StringVar(value="ซัพพลายเออร์จัดส่ง"),
            'payment_total_var': tk.StringVar(value="0.00"),
            'so_subtotal_var': tk.StringVar(value="0.00"),
            'so_vat_var': tk.StringVar(value="0.00"),
            'so_grand_total_var': tk.StringVar(value="0.00"),
            'so_vs_payment_result_var': tk.StringVar(value="-"),
            'difference_amount_var': tk.StringVar(value="0.00"),
            'balance_due_var': tk.StringVar(value="0.00"),
            'cash_product_input_var': tk.StringVar(value="0.00"),
            'cash_service_total_var': tk.StringVar(value="0.00"),
            'cash_required_total_var': tk.StringVar(value="0.00"),
            'cash_actual_payment_var': tk.StringVar(value="0.00"),
            'cash_verification_result_var': tk.StringVar(value="-"),
            'sales_vat_calc_var': tk.StringVar(value="0.00"),
            'cutting_drilling_vat_calc_var': tk.StringVar(value="0.00"),
            'other_service_vat_calc_var': tk.StringVar(value="0.00"),
            'shipping_vat_calc_var': tk.StringVar(value="0.00"),
            'card_fee_vat_calc_var': tk.StringVar(value="0.00"),
            'relocation_vat_calc_var': tk.StringVar(value="0.00"),
            'sales_service_vat_option': tk.StringVar(value="VAT"),
            'cutting_drilling_fee_vat_option': tk.StringVar(value="VAT"),
            'other_service_fee_vat_option': tk.StringVar(value="VAT"),
            'shipping_vat_option_var': tk.StringVar(value="VAT"),
            'credit_card_fee_vat_option_var': tk.StringVar(value="VAT"),
            'relocation_cost_vat_option': tk.StringVar(value="VAT")
        })

    # =========================================================================
    # ระบบยกเลิก SO (Cancel Logic)
    # =========================================================================
    def _search_so_to_cancel(self):
        """ค้นหา SO และแสดงปุ่มยกเลิกในแถบเดียวกัน"""
        # ล้างผลลัพธ์เดิม
        for widget in self.inline_result_frame.winfo_children():
            widget.destroy()
            
        so_number = self.cancel_search_entry.get().strip().upper()
        if not so_number:
            CTkLabel(self.inline_result_frame, text="⚠️ กรุณาระบุเลข SO", text_color="#DC2626", font=self.label_font_bold).pack(side="left")
            return
            
        try:
            # ค้นหา SO ล่าสุดที่ Active อยู่
            query = "SELECT id, so_number, customer_name, status FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1"
            df = pd.read_sql_query(query, self.pg_engine, params=(so_number,))
            
            if df.empty:
                CTkLabel(self.inline_result_frame, text="❌ ไม่พบข้อมูล SO นี้ในระบบ", text_color="#DC2626", font=self.label_font_bold).pack(side="left")
                return
                
            row = df.iloc[0]
            current_status = row['status']
            
            # เช็คสถานะก่อนว่ายกเลิกได้ไหม
            if current_status == 'Cancelled':
                CTkLabel(self.inline_result_frame, text="⚠️ SO นี้ถูกยกเลิกไปแล้ว", text_color="#D97706", font=self.label_font_bold).pack(side="left")
                return
            if current_status == 'Paid':
                CTkLabel(self.inline_result_frame, text="⚠️ ไม่สามารถยกเลิกได้ (SO นี้ชำระเงินเรียบร้อยแล้ว)", text_color="#D97706", font=self.label_font_bold).pack(side="left")
                return
                
            # ถ้าสามารถยกเลิกได้ แสดงข้อมูล + ปุ่มยกเลิก
            info_text = f"✅ พบ SO: {row['so_number']} | ลูกค้า: {row['customer_name']} | สถานะ: {current_status}"
            CTkLabel(self.inline_result_frame, text=info_text, text_color="#16A34A", font=self.label_font_bold).pack(side="left", padx=(0, 15))
            
            CTkButton(self.inline_result_frame, text="🗑️ ยืนยันยกเลิก SO", fg_color="#DC2626", hover_color="#B91C1C", 
                      width=130, height=30, font=CTkFont(weight="bold"),
                      command=lambda: self._execute_cancel_so(row['id'], row['so_number'])).pack(side="left")
                      
        except Exception as e:
            CTkLabel(self.inline_result_frame, text=f"Error: {str(e)[:50]}", text_color="red").pack(side="left")

    def _execute_cancel_so(self, so_id, so_number):
        """เรียก Popup ระบุเหตุผล และดำเนินการยกเลิก SO"""
        
        def process_cancel(reason):
            # ถามย้ำอีกครั้งเพื่อความชัวร์ ป้องกันกดพลาด
            if not messagebox.askyesno("ยืนยันครั้งสุดท้าย", f"คุณแน่ใจหรือไม่ที่จะยกเลิก SO: {so_number} อย่างถาวร?"):
                return
                
            conn = None
            try:
                conn = self.app_container.get_connection()
                with conn.cursor() as cursor:
                    # เปลี่ยนสถานะเป็น Cancelled และเก็บเหตุผลลง DB
                    cursor.execute("""
                        UPDATE commissions 
                        SET status = 'Cancelled', 
                            rejection_reason = %s,
                            sm_reject_count = COALESCE(sm_reject_count, 0) + 1
                        WHERE id = %s
                    """, (f"ยกเลิกโดย SM: {reason}", so_id))
                    
                    # แจ้งเตือนเซลล์เจ้าของ SO ผ่าน Noti
                    cursor.execute("SELECT sale_key FROM commissions WHERE id = %s", (so_id,))
                    res = cursor.fetchone()
                    if res:
                        cursor.execute("""
                            INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) 
                            VALUES (%s, %s, FALSE, %s)
                        """, (res[0], f"SO: {so_number} ถูกยกเลิกโดย Manager เหตุผล: {reason}", so_id))
                        
                conn.commit()
                messagebox.showinfo("สำเร็จ", f"ยกเลิก SO: {so_number} เรียบร้อยแล้ว")
                
                # ล้างหน้าจอและโหลดข้อมูลใหม่
                self.cancel_search_entry.delete(0, tk.END)
                for widget in self.inline_result_frame.winfo_children(): widget.destroy()
                self._load_cancelled_so_history()
                self._refresh_all_tabs() # อัปเดตตารางอื่นๆ ด้วย
                
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Error", f"ไม่สามารถยกเลิกได้: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)

        # เปิดหน้าต่าง Popup SOCancelDialog ขึ้นมาให้เลือกเหตุผล
        SOCancelDialog(self, so_number, process_cancel)

    def _load_cancelled_so_history(self):
        """โหลดข้อมูลประวัติ SO ที่ถูกยกเลิกลง Treeview"""
        for widget in self.cancelled_history_frame.winfo_children():
            widget.destroy()
            
        try:
            query = """
                SELECT c.timestamp, c.so_number, c.customer_name, u.sale_name, c.rejection_reason
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Cancelled' AND c.is_active = 1
                ORDER BY c.timestamp DESC LIMIT 100
            """
            df = pd.read_sql_query(query, self.pg_engine)
            
            # ใช้ Treeview เพื่อความสะอาดและโหลดไว
            columns = ("วันที่", "SO Number", "ชื่อลูกค้า", "เซลล์", "เหตุผลที่ยกเลิก")
            tree = ttk.Treeview(self.cancelled_history_frame, columns=columns, show="headings", height=15)
            
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview.Heading", font=('Tahoma', 12, 'bold'), background="#F3F4F6")
            style.configure("Treeview", font=('Tahoma', 11), rowheight=30)
            
            tree.heading("วันที่", text="วันที่ / เวลา")
            tree.heading("SO Number", text="SO Number")
            tree.heading("ชื่อลูกค้า", text="ชื่อลูกค้า")
            tree.heading("เซลล์", text="พนักงานขาย")
            tree.heading("เหตุผลที่ยกเลิก", text="เหตุผลที่ยกเลิก")
            
            tree.column("วันที่", width=150, anchor="center")
            tree.column("SO Number", width=150, anchor="center")
            tree.column("ชื่อลูกค้า", width=350, anchor="w")
            tree.column("เซลล์", width=150, anchor="center")
            tree.column("เหตุผลที่ยกเลิก", width=400, anchor="w")
            
            vsb = ttk.Scrollbar(self.cancelled_history_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vsb.set)
            
            tree.pack(side="left", fill="both", expand=True)
            vsb.pack(side="right", fill="y")
            
            if df.empty:
                # ถ้าไม่มีข้อมูล ให้แทรกแถวว่างๆ บอกไว้
                tree.insert("", "end", values=("-", "ไม่มีประวัติการยกเลิก SO", "-", "-", "-"))
                return
                
            for _, row in df.iterrows():
                ts = row['timestamp']
                ts_str = str(ts)[:16] if pd.notna(ts) else "-"
                reason = row['rejection_reason'] or "ไม่ได้ระบุเหตุผล"
                
                tree.insert("", "end", values=(
                    ts_str,
                    row['so_number'],
                    row['customer_name'] or "-",
                    row['sale_name'] or "-",
                    reason
                ))
                
        except Exception as e:
            CTkLabel(self.cancelled_history_frame, text=f"Error loading history: {e}", text_color="red").pack()

    def _create_defer_approval_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        # Header
        header_frame = CTkFrame(parent_tab, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=10)
        CTkLabel(header_frame, text="⏳ รายการที่ HR/บัญชี ขอเลื่อนรอบจ่ายคอมมิชชั่น", font=CTkFont(size=16, weight="bold")).pack(side="left")
        CTkButton(header_frame, text="⟳ รีเฟรช", command=self._load_defer_requests, width=90, fg_color="gray").pack(side="right")
        CTkButton(header_frame, text="📋 ประวัติการเลื่อน SO",
                  command=lambda: DeferralHistoryWindow(self, self.app_container),
                  fg_color="#2563EB", hover_color="#1D4ED8", width=160).pack(side="right", padx=(0, 8))

        # Scrollable Frame
        self.defer_list_frame = CTkScrollableFrame(parent_tab, fg_color="white", corner_radius=8)
        self.defer_list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.after(200, self._load_defer_requests)

    def _load_defer_requests(self):
        for widget in self.defer_list_frame.winfo_children(): widget.destroy()

        try:
            # 🟢 [แก้ไข] เพิ่มการดึง commission_month และ commission_year มาจาก DB ด้วย
            query = """
                SELECT c.id, c.so_number, c.customer_name, u.sale_name, c.sale_key, c.rejection_reason,
                       c.commission_month, c.commission_year
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Defer Requested' AND c.is_active = 1
                ORDER BY c.timestamp DESC
            """
            df = pd.read_sql_query(query, self.pg_engine)

            if df.empty:
                CTkLabel(self.defer_list_frame, text="✅ ไม่มีรายการที่ขอเลื่อน SO ในขณะนี้", font=CTkFont(size=14)).pack(pady=40)
                return

            thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                           "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

            for _, row in df.iterrows():
                card = CTkFrame(self.defer_list_frame, fg_color="#FFF7ED", border_width=1, border_color="#FDBA74", corner_radius=8)
                card.pack(fill="x", padx=10, pady=5)

                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.pack(side="left", fill="both", expand=True, padx=15, pady=10)

                # 🟢 [เพิ่มใหม่] ลอจิกคำนวณเดือนที่ถูกเลื่อน (จากเดือนไหน -> ไปเดือนไหน)
                try:
                    m_current = int(row['commission_month'])
                    y_current = int(row['commission_year']) + 543
                    
                    from_month_str = thai_months[m_current - 1] if 1 <= m_current <= 12 else str(m_current)
                    from_year_str = str(y_current)
                    
                    # คำนวณเดือนถัดไป
                    m_next = m_current + 1
                    y_next = y_current
                    if m_next > 12:
                        m_next = 1
                        y_next += 1
                        
                    to_month_str = thai_months[m_next - 1]
                    to_year_str = str(y_next)
                    
                    defer_period_text = f"🔄 ขอเลื่อนจาก: {from_month_str} {from_year_str}  ➔  ไปเป็น: {to_month_str} {to_year_str}"
                except Exception:
                    defer_period_text = "🔄 ขอเลื่อนไปเดือนถัดไป"

                # แสดงบรรทัดที่ 1: เลข SO + ชื่อลูกค้า
                CTkLabel(info_frame, text=f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}", font=CTkFont(size=14, weight="bold")).pack(anchor="w")
                
                # แสดงบรรทัดที่ 2: ข้อความบอกรอบเดือน (สีน้ำเงิน) 🟢 [เพิ่มใหม่]
                CTkLabel(info_frame, text=defer_period_text, text_color="#2563EB", font=CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(2, 0))
                
                # แสดงบรรทัดที่ 3: ชื่อเซลล์ + เหตุผล (สีส้ม)
                CTkLabel(info_frame, text=f"👤 เซลล์: {row['sale_name']} | 💬 เหตุผลที่บัญชีขอเลื่อน: {row['rejection_reason']}", text_color="#C2410C", font=CTkFont(size=13)).pack(anchor="w")

                # ส่วนปุ่มกด
                btn_frame = CTkFrame(card, fg_color="transparent")
                btn_frame.pack(side="right", padx=15, pady=10)

                CTkButton(btn_frame, text="✅ อนุมัติให้เลื่อน", fg_color="#16A34A", hover_color="#15803D", width=120,
                          command=lambda r=row: self._action_defer(r, approve=True)).pack(side="left", padx=5)

                CTkButton(btn_frame, text="❌ ไม่อนุมัติ (บังคับจ่าย)", fg_color="#DC2626", hover_color="#B91C1C", width=140,
                          command=lambda r=row: self._action_defer(r, approve=False)).pack(side="left", padx=5)

        except Exception as e:
            print(f"Error loading defer requests: {e}")

    def _action_defer(self, row, approve):
        try:
            cur_m = int(row.get('commission_month', datetime.now().month))
            cur_y = int(row.get('commission_year', datetime.now().year))
        except Exception:
            cur_m, cur_y = datetime.now().month, datetime.now().year

        dialog = ManagerDeferApprovalDialog(self, row['so_number'], cur_m, cur_y, approve=approve)
        self.wait_window(dialog)

        if not dialog.confirmed:
            return

        reason = dialog.reason or ("Manager อนุมัติการเลื่อน" if approve else "Manager ไม่อนุมัติการเลื่อน บังคับจ่ายรอบนี้")

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                manager_key = getattr(self.app_container, 'current_user_key', 'Manager')

                if approve:
                    new_status = 'Deferred'
                    defer_decision = 'อนุมัติ'
                    t_m, t_y = dialog.target_month, dialog.target_year
                    thai_months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
                    month_label = f"{thai_months[t_m-1]} {t_y+543}"
                    msg_for_sale = (f"[DEFER] ✅ Manager อนุมัติเลื่อนคอม SO: {row['so_number']}\n"
                                    f"📅 รอบคอมที่จะนำกลับมาคิด: {month_label}\n"
                                    f"💬 เหตุผล: {reason}")
                    cursor.execute("""
                        UPDATE commissions
                        SET status = %s, rejection_reason = %s,
                            defer_decision = %s, defer_decision_reason = %s,
                            defer_approved_by = %s,
                            commission_month = %s, commission_year = %s
                        WHERE id = %s
                    """, (new_status, f"Manager Decision: {reason}", defer_decision, reason,
                          manager_key, t_m, t_y, row['id']))
                else:
                    new_status = 'Pending HR Approval'
                    defer_decision = 'ไม่อนุมัติ'
                    msg_for_sale = (f"[DEFER] ❌ Manager ไม่อนุมัติเลื่อนคอม SO: {row['so_number']}\n"
                                    f"บังคับจ่ายรอบปัจจุบัน\n💬 เหตุผล: {reason}")
                    cursor.execute("""
                        UPDATE commissions
                        SET status = %s, rejection_reason = %s,
                            defer_decision = %s, defer_decision_reason = %s,
                            defer_approved_by = %s
                        WHERE id = %s
                    """, (new_status, f"Manager Decision: {reason}", defer_decision, reason,
                          manager_key, row['id']))

                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) VALUES (%s, %s, FALSE, %s)",
                    (row['sale_key'], msg_for_sale, row['id']))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"บันทึกการตัดสินใจ SO: {row['so_number']} เรียบร้อยแล้ว")
            self._load_defer_requests()
            self._refresh_all_tabs()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", str(e))
        finally:
            if conn: self.app_container.release_connection(conn)

class SOCancelDialog(CTkToplevel):
    def __init__(self, master, so_number, on_confirm_callback):
        super().__init__(master)
        self.title(f"ยกเลิก SO: {so_number}")
        self.geometry("450x550")
        self.on_confirm_callback = on_confirm_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True) 

        CTkLabel(self, text=f"ระบุเหตุผลที่ยกเลิก SO: {so_number}", 
                font=CTkFont(size=16, weight="bold")).pack(pady=15)

        # รายการ Checkbox 
        self.reasons = [
            "1. สินค้าไม่พอ/หาของไม่ได้",
            "2. ลูกค้าเปลี่ยนสเปค/สั่งผิด",
            "3. แจ้งราคาผิดพลาด",
            "4. ราคาปรับขึ้น/ลูกค้าไม่สู้",
            "5. ขนส่งช้า/ไม่ตรงนัด",
            "6. งานด่วน/ส่งไม่ทัน",
            "7. ข้อมูลสเปคคลาดเคลื่อน",
            "8. เปิด SO ซ้ำ/ผิดพลาด"
        ]
        
        self.check_vars = []
        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=30)

        for reason in self.reasons:
            var = tk.BooleanVar(value=False)
            cb = CTkCheckBox(container, text=reason, variable=var, font=CTkFont(size=13))
            cb.pack(anchor="w", pady=5)
            self.check_vars.append((var, reason))

        # ช่องกรอกข้อมูลเพิ่มเติม (อื่นๆ)
        CTkLabel(self, text="อื่นๆ / ระบุเพิ่มเติม:", 
                font=CTkFont(size=13, weight="bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.other_reason_entry = CTkEntry(self, placeholder_text="พิมพ์เหตุผลเพิ่มเติมที่นี่...", width=380)
        self.other_reason_entry.pack(padx=30, pady=(0, 20))

        # ปุ่มกดยกเลิก/ตกลง
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=10)
        
        CTkButton(btn_frame, text="ปิด", fg_color="gray", width=100, 
                 command=self.destroy).pack(side="left", padx=5)
        CTkButton(btn_frame, text="ตกลง (ยกเลิก SO)", fg_color="#DC2626", hover_color="#B91C1C", 
                 width=130, command=self._on_confirm).pack(side="right", padx=5)

    def _on_confirm(self):
        selected_reasons = [text for var, text in self.check_vars if var.get()]
        other_text = self.other_reason_entry.get().strip()
        
        if other_text:
            selected_reasons.append(other_text)

        if not selected_reasons:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกหรือระบุเหตุผลอย่างน้อย 1 ข้อ", parent=self)
            return

        final_reason = ", ".join(selected_reasons)
        self.on_confirm_callback(final_reason)
        self.destroy()

class SORejectionDialog(CTkToplevel):
    def __init__(self, master, so_number, on_confirm_callback):
        super().__init__(master)
        self.title(f"ตีกลับ SO: {so_number}")
        self.geometry("450x550")
        self.on_confirm_callback = on_confirm_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True) 

        CTkLabel(self, text=f"ระบุเหตุผลที่ตีกลับ SO: {so_number}", 
                font=CTkFont(size=16, weight="bold")).pack(pady=15)

        # รายการ Checkbox สำหรับตีกลับ
        self.reasons = [
            "1. เอกสารแนบไม่ครบถ้วน",
            "2. ข้อมูลลูกค้า/ที่อยู่ ไม่ถูกต้อง",
            "3. ยอดเงินไม่ตรงกับใบเสนอราคา",
            "4. เครดิต/เงื่อนไขการชำระเงินไม่ถูกต้อง",
            "5. รายการสินค้า/ราคา ไม่ถูกต้อง",
            "6. ข้อมูลการจัดส่งไม่ชัดเจน",
            "7. รอเอกสารจากทางฝั่งลูกค้า",
            "8. อื่นๆ (ระบุเพิ่มเติมด้านล่าง)"
        ]
        
        self.check_vars = []
        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=30)

        for reason in self.reasons:
            var = tk.BooleanVar(value=False)
            cb = CTkCheckBox(container, text=reason, variable=var, font=CTkFont(size=13))
            cb.pack(anchor="w", pady=5)
            self.check_vars.append((var, reason))

        # ช่องกรอกข้อมูลเพิ่มเติม (อื่นๆ)
        CTkLabel(self, text="อื่นๆ / ระบุเพิ่มเติม:", 
                font=CTkFont(size=13, weight="bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.other_reason_entry = CTkEntry(self, placeholder_text="พิมพ์เหตุผลเพิ่มเติมที่นี่...", width=380)
        self.other_reason_entry.pack(padx=30, pady=(0, 20))

        # ปุ่มกด
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=10)
        
        CTkButton(btn_frame, text="ปิด", fg_color="gray", width=100, 
                 command=self.destroy).pack(side="left", padx=5)
        CTkButton(btn_frame, text="ตกลง (ส่งกลับไปแก้)", fg_color="#DC2626", hover_color="#B91C1C", 
                 width=140, command=self._on_confirm).pack(side="right", padx=5)

        self.transient(master)
        self.grab_set()
        self.focus_force()

    def _on_confirm(self):
        selected_reasons = [text for var, text in self.check_vars if var.get()]
        other_text = self.other_reason_entry.get().strip()
        
        if other_text:
            selected_reasons.append(other_text)

        if not selected_reasons:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกหรือระบุเหตุผลอย่างน้อย 1 ข้อ", parent=self)
            return

        final_reason = ", ".join(selected_reasons)
        self.on_confirm_callback(final_reason)
        self.destroy()

class SMExportDialog(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        
        self.title("📥 Export ข้อมูล SO ตามรอบคอมมิชชั่น")
        self.geometry("450x420") # ขยายความสูงหน้าต่างนิดหน่อยเผื่อที่ให้ข้อความ
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True)

        # --- 1. เตรียมข้อมูล Dropdown ---
        now = datetime.now()
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        
        current_year_be = now.year + 543
        self.years = [str(y) for y in range(current_year_be - 2, current_year_be + 2)]
        
        self.sales_list = ["ทั้งหมด (All Sales)"]
        self.sale_mapping = {}
        self._load_sales_users()

        # --- 2. ตัวแปรเก็บค่าที่เลือก ---
        self.selected_year = tk.StringVar(value=str(current_year_be))
        self.selected_month = tk.StringVar(value=self.thai_months[now.month - 1])
        self.selected_sale = tk.StringVar(value="ทั้งหมด (All Sales)")

        # --- 3. สร้าง UI ---
        # หัวข้อหลัก (ปรับ pady ด้านล่างให้ชิดข้อความ Note มากขึ้น)
        CTkLabel(self, text="เลือกเงื่อนไขเพื่อ Export ข้อมูล", font=CTkFont(size=16, weight="bold")).pack(pady=(20, 5))
        
        # 🟢 [เพิ่มตรงนี้] ข้อความ Note อธิบายเพิ่มเติม
        CTkLabel(self, text="* Export Report เพื่อตรวจสอบค่าคอมมิชชั่น ประจำงวด", 
                 font=CTkFont(size=12), text_color="gray50").pack(pady=(0, 15))

        form_frame = CTkFrame(self, fg_color="transparent")
        form_frame.pack(fill="x", padx=40)
        form_frame.grid_columnconfigure(1, weight=1)

        # เปลี่ยนคำให้ชัดเจนขึ้นว่าเป็น "รอบคิดคอมมิชชั่น"
        CTkLabel(form_frame, text="รอบคอมมิชชั่น (ปี พ.ศ.):", font=CTkFont(size=14)).grid(row=0, column=0, sticky="w", pady=10)
        CTkOptionMenu(form_frame, variable=self.selected_year, values=self.years).grid(row=0, column=1, sticky="ew", padx=(10, 0))

        CTkLabel(form_frame, text="รอบคอมมิชชั่น (เดือน):", font=CTkFont(size=14)).grid(row=1, column=0, sticky="w", pady=10)
        CTkOptionMenu(form_frame, variable=self.selected_month, values=self.thai_months).grid(row=1, column=1, sticky="ew", padx=(10, 0))

        # Sale
        CTkLabel(form_frame, text="พนักงานขาย:", font=CTkFont(size=14)).grid(row=2, column=0, sticky="w", pady=10)
        CTkOptionMenu(form_frame, variable=self.selected_sale, values=self.sales_list).grid(row=2, column=1, sticky="ew", padx=(10, 0))

        # ปุ่มกด
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=40, pady=30)
        
        CTkButton(btn_frame, text="ยกเลิก", fg_color="gray", width=120, command=self.destroy).pack(side="left")
        CTkButton(btn_frame, text="📥 Export Excel", fg_color="#10B981", hover_color="#059669", 
                  width=120, command=self._execute_export).pack(side="right")
    def _load_sales_users(self):
        """ดึงรายชื่อเซลส์ทั้งหมดมาใส่ใน Dropdown"""
        try:
            df = pd.read_sql_query("SELECT sale_key, sale_name FROM sales_users WHERE role = 'Sale' ORDER BY sale_key", self.pg_engine)
            for _, row in df.iterrows():
                display_name = f"[{row['sale_key']}] {row['sale_name']}"
                self.sales_list.append(display_name)
                self.sale_mapping[display_name] = row['sale_key'] # เก็บ mapping ไว้ค้นหา
        except Exception as e:
            print(f"Error loading sales users: {e}")

    def _execute_export(self):
        # 1. แปลงค่าจาก Dropdown เป็นค่าสำหรับ Database
        month_num = self.month_map[self.selected_month.get()]
        year_ad = int(self.selected_year.get()) - 543
        sale_selection = self.selected_sale.get()

        # 2. สร้าง SQL Query (เอาคอลัมน์สถานะและอัปเดตล่าสุดออก)
        query = """
            SELECT 
                c.so_number AS "เลขที่ SO",
                c.customer_name AS "ชื่อลูกค้า",
                c.sales_service_amount AS "ยอดขาย (บาท)",
                COALESCE(u.sale_name, c.sale_key) AS "พนักงานขาย"
            FROM commissions c
            LEFT JOIN sales_users u ON c.sale_key = u.sale_key
            WHERE c.is_active = 1
              AND c.commission_month = %s
              AND c.commission_year = %s
        """
        params = [month_num, year_ad]

        # ถ้าเลือกเซลส์คนใดคนหนึ่ง ให้เพิ่มเงื่อนไข
        if sale_selection != "ทั้งหมด (All Sales)":
            target_sale_key = self.sale_mapping[sale_selection]
            query += " AND c.sale_key = %s"
            params.append(target_sale_key)
            
        query += " ORDER BY c.so_number ASC"

        # 3. ดึงข้อมูล
        try:
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
            
            if df.empty:
                messagebox.showinfo("แจ้งเตือน", "ไม่พบข้อมูล SO ตามเงื่อนไขที่เลือก", parent=self)
                return

            # 4. บันทึกไฟล์
            sale_str = "All" if sale_selection == "ทั้งหมด (All Sales)" else target_sale_key
            default_filename = f"SO_Report_{year_ad}_{month_num:02d}_{sale_str}.xlsx"
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_filename,
                title="บันทึกไฟล์ Excel",
                parent=self
            )
            
            if not file_path: return # กดยกเลิก
            
            # ตกแต่ง Excel เล็กน้อย (เหลือแค่ 4 คอลัมน์)
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='SO Data', index=False)
                worksheet = writer.sheets['SO Data']
                worksheet.column_dimensions['A'].width = 20 # SO Number
                worksheet.column_dimensions['B'].width = 35 # Customer
                worksheet.column_dimensions['C'].width = 18 # Sales
                worksheet.column_dimensions['D'].width = 30 # Sale Name
            
            self.destroy()
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้ว!\n\nจำนวน: {len(df)} รายการ")
            
            # เปิดไฟล์ให้ดูเลย (เฉพาะ Windows)
            if os.name == 'nt':
                os.startfile(file_path)

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            print(traceback.format_exc())