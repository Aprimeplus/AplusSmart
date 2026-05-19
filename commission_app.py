# commission_app.py (ฉบับสมบูรณ์ รวมทุกฟังก์ชัน)

import tkinter as tk
from tkinter import messagebox, filedialog
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton, CTkRadioButton, 
                           CTkEntry, CTkOptionMenu, CTkScrollableFrame, CTkToplevel, CTkTabview, CTkCheckBox)
import pandas as pd
from datetime import datetime, timedelta
import traceback
import psycopg2.errors
import psycopg2.extras
import os 
from history_windows import PurchaseDetailWindow, DeferralActionDialog
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
import utils
from export_utils import DateRangeDialog
from tkinter import ttk
from daily_report_widget import DailyReportWidget
from outstanding_dashboard_tab import OutstandingDashboardTab

class PaymentUpdateWindow(CTkToplevel):
    """หน้าต่าง Pop-up สำหรับอัปเดตข้อมูลการชำระเงินโดยเฉพาะ"""
    def __init__(self, master, app_container, so_data, on_save_callback):
        super().__init__(master)
        self.app_container = app_container
        self.so_data = so_data
        self.on_save_callback = on_save_callback
        self.so_id = self.so_data['id']

        self.title(f"อัปเดตยอดชำระ SO: {self.so_data['so_number']}")
        self.geometry("560x500")
        self.grid_columnconfigure(0, weight=1)

        # --- Display Info ---
        info_frame = CTkFrame(self, fg_color="transparent")
        info_frame.grid(row=0, column=0, padx=20, pady=15, sticky="ew")
        info_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(info_frame, text="SO Number:", font=CTkFont(weight="bold")).grid(row=0, column=0, sticky="w")
        CTkLabel(info_frame, text=self.so_data['so_number']).grid(row=0, column=1, sticky="w", padx=5)
        
        # <<< START: แก้ไขการดึงยอดที่ต้องชำระ >>>
        # ใช้สูตรคำนวณยอดที่ต้องชำระที่แท้จริง เพื่อความแม่นยำ
        original_payment = self.so_data.get('total_payment_amount', 0.0) or 0.0
        original_difference = self.so_data.get('difference_amount', 0.0) or 0.0
        actual_grand_total = original_payment - original_difference

        CTkLabel(info_frame, text="ยอดที่ต้องชำระ:", font=CTkFont(weight="bold")).grid(row=1, column=0, sticky="w")
        CTkLabel(info_frame, text=f"{actual_grand_total:,.2f} บาท").grid(row=1, column=1, sticky="w", padx=5)
        # <<< END >>>

        CTkLabel(info_frame, text="ยอดโอนขาดปัจจุบัน:", font=CTkFont(weight="bold"), text_color="#D97706").grid(row=2, column=0, sticky="w")
        CTkLabel(info_frame, text=f"{abs(original_difference):,.2f} บาท", text_color="#D97706").grid(row=2, column=1, sticky="w", padx=5)

        # --- Payment Entries ---
        payment_frame = CTkFrame(self)
        payment_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        payment_frame.grid_columnconfigure(1, weight=1)
        payment_frame.grid_columnconfigure(3, weight=1)

        CTkLabel(payment_frame, text="ยอดโอนชำระ 1:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.payment1_entry = NumericEntry(payment_frame)
        self.payment1_entry.insert(0, f"{original_payment:,.2f}")
        self.payment1_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        CTkLabel(payment_frame, text="วันที่โอน:").grid(row=0, column=2, padx=(10, 2), pady=5, sticky="w")
        self.payment1_date = DateSelector(payment_frame)
        self.payment1_date.grid(row=0, column=3, padx=10, pady=5, sticky="ew")

        CTkLabel(payment_frame, text="ยอดโอนชำระ 2 (เพิ่มเติม):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.payment2_entry = NumericEntry(payment_frame, placeholder_text="กรอกยอดที่โอนเพิ่ม...")
        self.payment2_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        CTkLabel(payment_frame, text="วันที่โอน:").grid(row=1, column=2, padx=(10, 2), pady=5, sticky="w")
        self.payment2_date = DateSelector(payment_frame)
        self.payment2_date.grid(row=1, column=3, padx=10, pady=5, sticky="ew")

        raw_date = self.so_data.get('payment_date')
        if raw_date:
            try:
                dt = raw_date if not isinstance(raw_date, str) else datetime.strptime(raw_date, "%Y-%m-%d")
                self.payment1_date.set_date(dt)
            except (ValueError, TypeError):
                pass

        # --- Buttons ---
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=20, pady=15, sticky="ew")
        button_frame.grid_columnconfigure((0,1), weight=1)
        
        CTkButton(button_frame, text="ยกเลิก", fg_color="gray", command=self.destroy).grid(row=0, column=0, padx=(0,5), sticky="ew")
        CTkButton(button_frame, text="บันทึกยอดชำระ", command=self._save_payment_update).grid(row=0, column=1, padx=(5,0), sticky="ew")

        self.transient(master)
        self.grab_set()

    def _save_payment_update(self):
        try:
            p1 = utils.convert_to_float(self.payment1_entry.get())
            p2 = utils.convert_to_float(self.payment2_entry.get())
            new_total_payment = p1 + p2
            
            # 1. คำนวณยอดที่ต้องชำระที่แท้จริง
            original_payment = self.so_data.get('total_payment_amount', 0.0) or 0.0
            original_difference = self.so_data.get('difference_amount', 0.0) or 0.0
            actual_grand_total = original_payment - original_difference
            
            # 2. คำนวณส่วนต่างใหม่ (ยอดโอนใหม่ - ยอดที่ต้องชำระ)
            new_difference = new_total_payment - actual_grand_total
            
            if not messagebox.askyesno("ยืนยัน", "คุณต้องการอัปเดตยอดชำระเงินใช่หรือไม่?", parent=self):
                return
            
            # 3. กำหนดสถานะใหม่ (ถ้าโอนครบแล้ว หรือโอนเกิน ให้เด้งกลับไปเป็น Edited เพื่อให้เซลส์กดนำส่งใหม่)
            current_status = self.so_data.get('status', 'Draft')
            new_status = current_status
            if new_difference >= 0:
                new_status = 'Edited' # <--- เปลี่ยนตรงนี้ได้ตาม Flow ของบริษัทคุณ

            date1 = self.payment1_date.get_date()
            date2 = self.payment2_date.get_date() if p2 > 0 else None

            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 🟢 [แก้ไข] อัปเดตให้ครบทุกคอลัมน์ที่เกี่ยวข้อง
                cursor.execute("""
                    UPDATE commissions SET
                        total_payment_amount = %s,
                        payment1_amount = %s,
                        payment2_amount = %s,
                        payment2_date = %s,
                        payment_date = %s,
                        difference_amount = %s,
                        status = %s,
                        timestamp = %s
                    WHERE id = %s
                """, (
                    new_total_payment,
                    p1,
                    p2,
                    date2,
                    date1,
                    new_difference,
                    new_status,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    self.so_id
                ))
            conn.commit()

            messagebox.showinfo("สำเร็จ", "อัปเดตยอดชำระเรียบร้อยแล้ว", parent=self)
            
            if self.on_save_callback:
                self.on_save_callback()
            self.destroy()

        except Exception as e:
            if 'conn' in locals() and conn: conn.rollback()
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถบันทึกได้: {e}", parent=self)
        finally:
            if 'conn' in locals() and conn: self.app_container.release_connection(conn)

class SalesTasksWindow(CTkToplevel):
    def __init__(self, master, app_container, sale_key):
        super().__init__(master)
        self.commission_app = master
        self.app_container = app_container
        self.sale_key = sale_key
        
        self.title("งานของฉัน (My Tasks)")
        
        # --- ปรับขนาดหน้าต่างให้ใหญ่ขึ้น (1200x800) ---
        self.geometry("1200x650")
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Header Section ---
        header = CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        # ปรับขนาดฟอนต์หัวข้อและปุ่มให้เหมาะสมกับหน้าจอที่ใหญ่ขึ้น
        CTkLabel(header, text="งานของฉัน", font=CTkFont(size=20, weight="bold")).pack(side="left")
        CTkButton(header, text="Refresh", command=self.load_tasks, width=100, height=35).pack(side="right")

        # --- Tab View ---
        self.task_tab_view = CTkTabview(self, corner_radius=10)
        self.task_tab_view.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        # 🟢 [เพิ่มใหม่] 4. แท็บนำส่งแล้ว (รอดำเนินการ)
        self.submitted_tab = self.task_tab_view.add("นำส่งแล้ว (รอดำเนินการ)")
        
        # 1. แท็บ SO ค้างชำระ
        self.payment_due_tab = self.task_tab_view.add("⚠️ SO ค้างชำระ (แก้ไขยอดโอน)")
        
        # 2. แท็บงานที่ถูกตีกลับ
        self.rejected_tab = self.task_tab_view.add("งานที่ถูกตีกลับ (Rejected)")
        
        # 3. แท็บฉบับร่าง
        self.draft_tab = self.task_tab_view.add("ฉบับร่าง (ยังไม่นำส่ง)")
        
        # 4. แท็บติดตามสถานะค่าคอมฯ (Tracking)
        self.comm_status_tab = self.task_tab_view.add("📊 ติดตามค่าคอมฯ (Tracking)")
        
        # --- Create Frames for Each Tab ---

        # Frame 1: Payment Due (ใช้ Scrollable)
        self.payment_due_frame = CTkScrollableFrame(self.payment_due_tab, label_text="รายการที่ยอดโอนชำระไม่ครบ (แก้ไขยอดโอนแล้วกดบันทึก)")
        self.payment_due_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Frame 2: Rejected / Deferred (ใช้ Scrollable)
        self.rejected_frame = CTkScrollableFrame(self.rejected_tab, label_text="รายการที่ต้องแก้ไข หรือ ตัดสินใจ")
        self.rejected_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Frame 3: Drafts (ใช้ Scrollable)
        self.draft_frame = CTkScrollableFrame(self.draft_tab, label_text="ดับเบิลคลิกรายการเพื่อแก้ไข/ทำต่อ")
        self.draft_frame.pack(fill="both", expand=True, padx=5, pady=5)

        current_date = datetime.now()
        thai_months = ["ทั้งหมด", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        year_list = ["ทั้งหมด"] + [str(y + 543) for y in range(current_date.year - 2, current_date.year + 2)]

        self.submit_search_var = tk.StringVar(value="")
        self.submit_month_var = tk.StringVar(value="ทั้งหมด")
        self.submit_year_var = tk.StringVar(value="ทั้งหมด")

        # 1. แถบเครื่องมือ Filter ด้านบน
        submit_filter_frame = CTkFrame(self.submitted_tab, fg_color="transparent")
        submit_filter_frame.pack(fill="x", padx=5, pady=5)

        CTkLabel(submit_filter_frame, text="🔍 ค้นหา:").pack(side="left", padx=(5, 2))
        CTkEntry(submit_filter_frame, textvariable=self.submit_search_var, placeholder_text="SO / ชื่อลูกค้า...", width=160).pack(side="left", padx=5)

        CTkLabel(submit_filter_frame, text="เดือน:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(submit_filter_frame, variable=self.submit_month_var, values=thai_months, width=100).pack(side="left", padx=5)

        CTkLabel(submit_filter_frame, text="ปี:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(submit_filter_frame, variable=self.submit_year_var, values=year_list, width=80).pack(side="left", padx=5)

        CTkButton(submit_filter_frame, text="ค้นหา", command=self._load_submitted_tasks, width=70, fg_color="#2563EB").pack(side="left", padx=(15, 5))
        CTkButton(submit_filter_frame, text="ล้างค่า", command=self._clear_submit_filter, width=70, fg_color="gray").pack(side="left", padx=5)

        # 2. ตารางแสดงผลด้านล่าง Filter
        self.submitted_frame = CTkScrollableFrame(self.submitted_tab, label_text="รายการที่ส่งเข้าระบบแล้ว (กำลังรอการตรวจสอบ/จัดซื้อ)")
        self.submitted_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Frame 4: Commission Tracking
        # [สำคัญ] ใช้ CTkFrame ธรรมดา (ไม่ Scrollable) เพื่อตรึงหัวข้อ และให้ตารางมี Scrollbar ของตัวเอง
        self.comm_status_frame = CTkFrame(self.comm_status_tab, fg_color="transparent")
        self.comm_status_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Setup UI ภายในแท็บ Tracking
        self._setup_commission_status_tab()
        
        # --- Final Steps ---
        self.after(50, self.load_tasks)
        self.transient(master)
        self.grab_set()
    
    def _setup_commission_status_tab(self):
        """สร้าง UI สำหรับหน้าติดตามค่าคอมมิชชั่น พร้อมตัวกรองเดือน/ปี (เวอร์ชันปรับขนาดตารางให้เล็กลง)"""
        
        # --- กำหนด Style สำหรับตารางหน้านี้โดยเฉพาะ ---
        style = ttk.Style()
        style.theme_use("clam")
        
        # คงขนาดฟอนต์และความสูงแถวไว้เพื่อให้ "อ่านง่าย" เหมือนเดิม
        style.configure("Tracking.Treeview", 
                        font=('Segoe UI', 14),      
                        rowheight=40,               
                        foreground="black",         
                        background="white")
        
        style.configure("Tracking.Treeview.Heading", 
                        font=('Segoe UI', 16, 'bold'), 
                        padding=(5, 10))

        # --- 1. ส่วนตัวกรอง (Filter) ---
        filter_frame = CTkFrame(self.comm_status_frame, fg_color="transparent")
        filter_frame.pack(fill="x", padx=10, pady=(5, 5))
        
        CTkLabel(filter_frame, text="เลือกเดือนที่ต้องการดู:", font=CTkFont(size=16)).pack(side="left", padx=(0, 5))

        current_date = datetime.now()
        thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        try:
            default_month = thai_months[current_date.month - 1]
        except IndexError:
            default_month = thai_months[0]
            
        self.track_month_var = tk.StringVar(value=default_month)
        self.track_year_var = tk.StringVar(value=str(current_date.year + 543))

        month_menu = CTkOptionMenu(filter_frame, variable=self.track_month_var, values=thai_months, width=140, font=CTkFont(size=14))
        month_menu.pack(side="left", padx=5)

        year_list = [str(y + 543) for y in range(current_date.year - 2, current_date.year + 2)]
        year_menu = CTkOptionMenu(filter_frame, variable=self.track_year_var, values=year_list, width=100, font=CTkFont(size=14))
        year_menu.pack(side="left", padx=5)

        btn_search = CTkButton(filter_frame, text="ค้นหา", command=self._load_commission_status, width=100, font=CTkFont(size=14, weight="bold"))
        btn_search.pack(side="left", padx=10)

        # ========================================================================================
        
        # --- 2. ส่วนตารางรายการที่ได้ค่าคอมฯ (บน) ---
        top_container = CTkFrame(self.comm_status_frame, fg_color="transparent")
        top_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.comm_title_label = CTkLabel(top_container, 
                           text="รายการที่ถูกคิดค่าคอมมิชชั่น (กำลังโหลด...)", 
                           font=CTkFont(size=18, weight="bold")) 
        self.comm_title_label.pack(anchor="w", pady=(0, 5))

        tree_frame_1 = CTkFrame(top_container, fg_color="transparent")
        tree_frame_1.pack(fill="both", expand=True)

        columns_comm = ("SO Number", "ลูกค้า", "ยอดขาย", "สถานะ")
        
        # [แก้ไข] ลด height จาก 15 เหลือ 10 แถว
        self.tree_comm = ttk.Treeview(tree_frame_1, columns=columns_comm, show="headings", style="Tracking.Treeview", height=10)
        
        vsb_1 = ttk.Scrollbar(tree_frame_1, orient="vertical", command=self.tree_comm.yview)
        self.tree_comm.configure(yscrollcommand=vsb_1.set)
        
        self.tree_comm.heading("SO Number", text="SO Number")
        self.tree_comm.heading("ลูกค้า", text="ลูกค้า")
        self.tree_comm.heading("ยอดขาย", text="ยอดขายสุทธิ")
        self.tree_comm.heading("สถานะ", text="สถานะค่าคอม")

        self.tree_comm.column("SO Number", width=180)
        self.tree_comm.column("ลูกค้า", width=500)
        self.tree_comm.column("ยอดขาย", width=150, anchor="e")
        self.tree_comm.column("สถานะ", width=160, anchor="center")
        
        self.tree_comm.pack(side="left", fill="both", expand=True)
        vsb_1.pack(side="right", fill="y")


        # --- 3. ส่วนรายการที่ถูกเลื่อน (Deferred) (ล่าง) ---
        bottom_container = CTkFrame(self.comm_status_frame, fg_color="transparent")
        bottom_container.pack(fill="both", expand=True, padx=10, pady=10)

        self.defer_title_label = CTkLabel(bottom_container, 
                           text="รายการที่ถูกเลื่อน (Deferred SOs) - กำลังโหลด...", 
                           font=CTkFont(size=18, weight="bold"))
        self.defer_title_label.pack(anchor="w", pady=(10, 5))

        tree_frame_2 = CTkFrame(bottom_container, fg_color="transparent")
        tree_frame_2.pack(fill="both", expand=True)

        columns_defer = ("SO Number", "ลูกค้า", "เลื่อนไปเดือน", "เหตุผล")
        
        # [แก้ไข] ลด height จาก 8 เหลือ 5 แถว
        self.tree_defer = ttk.Treeview(tree_frame_2, columns=columns_defer, show="headings", style="Tracking.Treeview", height=5)
        
        vsb_2 = ttk.Scrollbar(tree_frame_2, orient="vertical", command=self.tree_defer.yview)
        self.tree_defer.configure(yscrollcommand=vsb_2.set)
        
        self.tree_defer.heading("SO Number", text="SO Number")
        self.tree_defer.heading("ลูกค้า", text="ลูกค้า")
        self.tree_defer.heading("เลื่อนไปเดือน", text="🎯 เลื่อนไปเดือน")
        self.tree_defer.heading("เหตุผล", text="เหตุผล")
        
        self.tree_defer.column("SO Number", width=180)
        self.tree_defer.column("ลูกค้า", width=400)
        self.tree_defer.column("เลื่อนไปเดือน", width=180, anchor="center")
        self.tree_defer.column("เหตุผล", width=350)
        
        self.tree_defer.pack(side="left", fill="both", expand=True)
        vsb_2.pack(side="right", fill="y")

    def _open_deferral_dialog(self, row_data):
        """เปิดหน้าต่างให้ Sale ตัดสินใจเรื่องการเลื่อนจ่าย (Defer)"""
        # แปลง row_data เป็น dict ถ้าจำเป็น
        record_data = row_data.to_dict() if hasattr(row_data, 'to_dict') else row_data
        
        # เปิด Dialog โดยส่ง callback เป็น self.load_tasks เพื่อให้รีเฟรชหน้าจอหลังจากปิด
        DeferralActionDialog(self, self.app_container, record_data, callback=self.load_tasks)

    def _load_payment_due_tasks(self):
        """(ฉบับแก้ไข) โหลด SO ที่มียอดโอนขาดเท่านั้น"""
        for widget in self.payment_due_frame.winfo_children():
            widget.destroy()

        try:
            # ROUND ใน query ป้องกัน floating point -0.0001 หลุดผ่าน < 0
            query = """
                SELECT * FROM commissions
                WHERE sale_key = %s
                AND ROUND(difference_amount::numeric, 2) < 0
                AND is_active = 1
                AND status != 'Paid'
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))

            if df.empty:
                CTkLabel(self.payment_due_frame, text="ไม่พบรายการโอนขาด").pack(pady=20)
                return

            for _, row_data in df.iterrows():
                total_payment = row_data.get('total_payment_amount', 0.0) or 0.0
                difference = round(row_data.get('difference_amount', 0.0) or 0.0, 2)

                # Python-side filter: ถ้า round แล้ว >= 0 ถือว่าชำระครบ ข้ามไป
                if difference >= 0:
                    continue

                card_color = "#FFFBEB"  # สีเหลือง
                balance_text = f"ยอดโอนขาด: {difference:,.2f} บาท"
                text_color = "#B45309"
                button_text = "แก้ไขยอดชำระ"

                card = CTkFrame(self.payment_due_frame, border_width=1, fg_color=card_color)
                card.pack(fill="x", padx=5, pady=4)
                card.grid_columnconfigure(0, weight=1)

                top_frame = CTkFrame(card, fg_color="transparent")
                top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 0))

                info = f"SO: {row_data['so_number']} | ลูกค้า: {row_data['customer_name']} | สถานะปัจจุบัน: {row_data['status']}"
                CTkLabel(top_frame, text=info, font=CTkFont(size=14, weight="bold")).pack(side="left")

                edit_button = CTkButton(
                    top_frame,
                    text=button_text,
                    width=120,
                    command=lambda r=row_data.to_dict(): self._open_payment_updater(r)
                )
                edit_button.pack(side="right")

                copy_button = CTkButton(
                    top_frame, 
                    text="📋 Copy Shortnote", 
                    width=120, fg_color="#22C55E", hover_color="#16A34A",
                    command=lambda r=row_data.to_dict(): self._copy_so_shortnote(r)
                )
                copy_button.pack(side="right", padx=5)

                reason_label = CTkLabel(
                    card,
                    text=balance_text,
                    text_color=text_color,
                    wraplength=700,
                    justify="left",
                    anchor="w"
                )
                reason_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))

        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดรายการค้างชำระได้: {e}", parent=self)
    
    def _open_payment_updater(self, so_data_dict):
        """(ฟังก์ชันใหม่) เปิดหน้าต่างสำหรับอัปเดตยอดชำระโดยเฉพาะ"""
        PaymentUpdateWindow(
            master=self, 
            app_container=self.app_container, 
            so_data=so_data_dict, 
            on_save_callback=self.load_tasks
        )

    def on_close(self):
        self.commission_app.tasks_window = None
        self.destroy()

    def load_tasks(self):
        self._load_payment_due_tasks()
        self._load_rejected_tasks()
        self._load_draft_tasks()
        self._load_submitted_tasks()
        self._load_commission_status()

    def _load_submitted_tasks(self):
        for widget in self.submitted_frame.winfo_children(): widget.destroy()
        
        # ดึงค่าจาก Filter
        search_text = self.submit_search_var.get().strip().lower()
        selected_month = self.submit_month_var.get()
        selected_year = self.submit_year_var.get()
        
        thai_months_only = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

        try:
            # Base Query
            query = """
                SELECT * FROM commissions 
                WHERE sale_key = %s 
                AND status IN ('Pending Sale Manager Approval', 'Pending PU', 'PO In Progress', 'Pending HR Approval') 
                AND is_active = 1 
            """
            params = [self.sale_key]

            # ต่อเติม Query ตามเงื่อนไขที่เลือก
            if search_text:
                query += " AND (LOWER(so_number) LIKE %s OR LOWER(customer_name) LIKE %s)"
                params.extend([f"%{search_text}%", f"%{search_text}%"])

            if selected_month != "ทั้งหมด":
                month_num = thai_months_only.index(selected_month) + 1
                query += " AND commission_month = %s"
                params.append(month_num)

            if selected_year != "ทั้งหมด":
                year_num = int(selected_year) - 543 # แปลง พ.ศ. กลับเป็น ค.ศ. (ถ้าใน DB เก็บเป็น ค.ศ.)
                query += " AND commission_year = %s"
                params.append(year_num)

            query += " ORDER BY timestamp DESC"

            # Execute Query
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=tuple(params))
            
            if df.empty:
                CTkLabel(self.submitted_frame, text="ไม่มีรายการรอดำเนินการ หรือ ไม่พบข้อมูลที่ค้นหา").pack(pady=20)
                return
                
            for _, row in df.iterrows():
                card = CTkFrame(self.submitted_frame, border_width=1, fg_color="#F0F9FF") 
                card.pack(fill="x", padx=5, pady=3)
                
                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.pack(side="left", fill="both", expand=True)
                
                info_text = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']} | สถานะ: {row['status']}"
                CTkLabel(info_frame, text=info_text, font=CTkFont(weight="bold")).pack(side="left", padx=10, pady=10)

                btn_frame = CTkFrame(card, fg_color="transparent")
                btn_frame.pack(side="right", padx=10)
                
                CTkButton(
                    btn_frame, text="📋 Copy Shortnote", width=120, 
                    fg_color="#22C55E", hover_color="#16A34A", 
                    command=lambda r=row: self._copy_so_shortnote(r.to_dict())
                ).pack(side="left", padx=2)

        except Exception as e:
            print(f"Error loading submitted tasks: {e}")
            messagebox.showerror("Error", "ไม่สามารถโหลดรายการที่นำส่งแล้วได้", parent=self)

    def _clear_submit_filter(self):
        """ล้างค่า Filter ทั้งหมดแล้วโหลดตารางใหม่"""
        self.submit_search_var.set("")
        self.submit_month_var.set("ทั้งหมด")
        self.submit_year_var.set("ทั้งหมด")
        self._load_submitted_tasks()

    def _copy_so_shortnote(self, so_data):
        """ลอจิก Copy Shortnote (รองรับทศนิยม .86 สมบูรณ์แบบ)"""
        if not so_data:
            messagebox.showwarning("แจ้งเตือน", "ไม่มีข้อมูล SO สำหรับคัดลอก", parent=self)
            return

        try:
            so_number = so_data.get('so_number', '-')
            pickup_loc = so_data.get('pickup_location') or '-'
            
            # 🟢 ฟังก์ชันจัดการตัวเลข (ไม่ใช้ utils และไม่ปัดเศษทิ้ง)
            def format_money(val):
                try:
                    if pd.isna(val) or val is None or str(val).strip() == '': 
                        return "-"
                    # ลบลูกน้ำออกและแปลงเป็น Float ตรงๆ
                    if isinstance(val, str): 
                        val = val.replace(',', '')
                    v = float(val)
                except:
                    return "-"
                    
                if v <= 0: 
                    return "-"
                
                # เช็คทศนิยม: ถ้ายอดกลมๆ ให้โชว์ 0 ตำแหน่ง, ถ้ามีเศษสตางค์ให้โชว์ 2 ตำแหน่ง
                if v % 1 == 0:
                    return f"{v:,.0f}"
                else:
                    return f"{v:,.2f}"

            sales_amount = format_money(so_data.get('sales_service_amount'))
            shipping_cost = format_money(so_data.get('shipping_cost'))
            relocation_cost = format_money(so_data.get('relocation_cost'))
            cutting_fee = format_money(so_data.get('cutting_drilling_fee'))
            discount = format_money(so_data.get('coupons'))

            date_to_wh = utils.format_date_safe(so_data.get('date_to_warehouse'), '%d/%m')
            date_to_cust = utils.format_date_safe(so_data.get('date_to_customer'), '%d/%m')

            delivery_type = so_data.get('delivery_type') or '-'
            order_pur_val = so_data.get('order_pur') or '-'
            rego = so_data.get('pickup_registration') or '-'
            
            delivery_map = so_data.get('delivery_map') or '-'
            contact_name = so_data.get('onsite_contact_name') or '-'
            contact_phone = so_data.get('onsite_contact_phone') or '-'
            vehicle_type = so_data.get('vehicle_type') or '-'
            
            # 🟢 นำ format_money ตัวใหม่มาครอบยอดชำระ
            total_paid = so_data.get('total_payment_amount') or 0
            difference = so_data.get('difference_amount') or 0
            
            try: total_val = float(str(total_paid).replace(',', ''))
            except: total_val = 0.0
            
            try: diff_val = float(str(difference).replace(',', ''))
            except: diff_val = 0.0

            # คำนวณยอดเต็มที่แท้จริง
            grand_total = total_val - diff_val

            if total_val <= 0:
                payment_display = f"ยังไม่ชำระ (ยอดที่ต้องชำระ {format_money(grand_total)})"
            elif diff_val < -0.01:
                # กรณีติดลบ = โอนขาด หรือ มัดจำ
                payment_display = f"มัดจำ {format_money(total_val)} (ค้างชำระ {format_money(abs(diff_val))})"
            elif diff_val > 0.01:
                # กรณีเป็นบวก = โอนเกิน
                payment_display = f"เต็มจำนวน {format_money(total_val)} (โอนเกิน {format_money(diff_val)})"
            else:
                # กรณีเป็น 0 = จ่ายพอดีเป๊ะ
                payment_display = f"เต็มจำนวน {format_money(total_val)}"
            
            remark_text = so_data.get('credit_term', 'เงินสด')

            # หาตัวแปรชื่อผู้จัดทำ (รองรับทั้งหน้า Sale และ Support)
            maker_name = so_data.get('owner_sale_name', 'Unknown')
            if maker_name == 'Unknown' and hasattr(self, 'commission_app'):
                maker_name = getattr(self.commission_app, 'sale_name', 'Unknown')

            brokerage_fee = format_money(so_data.get('brokerage_fee'))
            coupon_val = format_money(so_data.get('coupons'))
            giveaway_vat = so_data.get('giveaway_vat') or '-'
            giveaway_no_vat = so_data.get('giveaway_no_vat') or '-'

            credit_card_fee = format_money(so_data.get('credit_card_fee'))
            transfer_fee = format_money(so_data.get('transfer_fee'))
            wht_fee = format_money(so_data.get('wht_3_percent'))
            
            # (ถ้ามีคอลัมน์ส่วนลดโปรโมชั่นแยกต่างหาก ให้ใช้ promotion_discount ถ้าไม่มีระบบจะแสดงเป็น '-')
            discount = format_money(so_data.get('promotion_discount'))

            special_req = so_data.get('special_request') or '-'
            unloading_stat = so_data.get('unloading_status') or '-'

            # สร้างเส้นคั่น
            separator = "-" * 10

            shortnote_text = (
                f"เลขที่ {so_number}\n"
                f"ยอดขาย : {sales_amount}\n"
                f"ค่าส่ง  : {shipping_cost}\n"
                f"ค่าย้าย : {relocation_cost}\n"
                f"ค่าตัด : {cutting_fee}\n"
                f"ยอดชำระ : {payment_display}\n"
                f"ค่าธรรมเนียมบัตรเครดิต : {credit_card_fee}\n"
                f"ค่าธรรมเนียมโอน : {transfer_fee}\n"
                f"ภาษีหัก ณ ที่จ่าย : {wht_fee}\n"
                f"ค่านายหน้า : {brokerage_fee}\n"
                f"คูปอง : {coupon_val}\n"
                f"ของแถมใน so (vat) : {giveaway_vat}\n"
                f"ของแถมนอก so (no vat) : {giveaway_no_vat}\n"
                f"{separator}\n"
                f"วันที่ย้ายสินค้าเข้าคลัง132 : {date_to_wh}\n"
                f"วันที่จัดส่งลูกค้า : {date_to_cust}\n"
                f"Order Pur : {order_pur_val}\n"
                f"Payment : {remark_text}\n"
                f"อนุมัติโอนยอดค้างส่วนที่เหลือ วันจัดส่งสินค้า ก่อนลงสินค้า\n"
                f"{separator}\n"
                f"การจัดส่ง : {delivery_type}\n"
                f"แผนที่จัดส่ง : {delivery_map}\n"
                f"Location เข้ารับ : {pickup_loc}\n"
                f"ประเภทรถ : {vehicle_type}\n"
                f"เงื่อนไขลงสินค้า : {unloading_stat}\n"
                f"ทะเบียนรถ : {rego}\n"
                f"ชื่อผู้ติดต่อหน้างาน : {contact_name}\n"
                f"เบอร์ติดต่อหน้างาน : {contact_phone}\n"
                f"Special Request : {special_req}\n"
                f"{separator}\n"
                f"อ้างอิงจาก Aplus Smart\n"
                f"ผู้จัดทำ: {maker_name}"
            )

            self.clipboard_clear()
            self.clipboard_append(shortnote_text)
            self.update() 

            messagebox.showinfo("คัดลอกสำเร็จ", f"คัดลอก Shortnote ของ {so_number} แล้ว!", parent=self)

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถสร้าง Shortnote ได้: {e}", parent=self)
            import traceback
            traceback.print_exc()
            
    # 2. เพิ่มฟังก์ชันโหลดข้อมูล
    def _load_commission_status(self):
        """โหลดข้อมูลใส่ตารางติดตามค่าคอมฯ (เวอร์ชันสำหรับตารางที่ไม่มี Margin)"""
        # ล้างข้อมูลเก่า
        for item in self.tree_comm.get_children(): self.tree_comm.delete(item)
        for item in self.tree_defer.get_children(): self.tree_defer.delete(item)
        
        try:
            thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
            selected_month_str = self.track_month_var.get()
            selected_year_str = self.track_year_var.get()

            target_month = thai_months.index(selected_month_str) + 1
            target_year = int(selected_year_str) - 543
        except:
            target_month = datetime.now().month
            target_year = datetime.now().year
            selected_month_str = thai_months[target_month-1]
            selected_year_str = str(target_year + 543)

        try:
            # --- 1. ตารางบน: Commissioned SOs ---
            # ดึง sales_service_amount มาเผื่อไว้แสดงผลกรณีที่ยังไม่มี final
            query_comm = """
                SELECT so_number, customer_name, final_sales_amount, sales_service_amount, status 
                FROM commissions 
                WHERE sale_key = %s 
                  AND status NOT IN ('Draft', 'Cancelled', 'Deferred')
                  AND commission_month = %s 
                  AND commission_year = %s
                  AND is_active = 1
                ORDER BY so_number DESC
            """
            df_comm = pd.read_sql_query(query_comm, self.app_container.pg_engine, params=(self.sale_key, target_month, target_year))
            
            count_comm = len(df_comm)
            self.comm_title_label.configure(
                text=f"รายการค่าคอมมิชชั่น (เดือน {selected_month_str} {selected_year_str})  —  รวม {count_comm} รายการ"
            )
            
            for _, row in df_comm.iterrows():
                # Logic การแสดงยอดเงิน: ถ้ามี Final ให้ใช้ Final, ถ้าไม่มีให้ใช้ยอดเบื้องต้น
                if pd.notna(row['final_sales_amount']) and row['final_sales_amount'] > 0:
                    sales_val = row['final_sales_amount']
                else:
                    sales_val = row['sales_service_amount']
                
                sales_str = f"{sales_val:,.2f}" if pd.notna(sales_val) else "0.00"
                
                # แปลง status → จ่ายแล้ว / ยังไม่จ่าย
                raw_status = row['status']
                if raw_status == 'Paid':
                    display_status = '✅ จ่ายแล้ว'
                    tag = 'paid'
                else:
                    display_status = '⏳ ยังไม่จ่าย'
                    tag = 'pending'

                self.tree_comm.tag_configure('paid',    foreground="#16A34A", font=('TH Sarabun New', 13, 'bold'))
                self.tree_comm.tag_configure('pending', foreground="black")

                self.tree_comm.insert("", "end", values=(
                    row['so_number'],
                    row['customer_name'],
                    sales_str,
                    display_status,
                ), tags=(tag,))

            # --- 2. ตารางล่าง: Deferred SOs ---
            query_defer = """
                SELECT so_number, customer_name, commission_month, commission_year, rejection_reason 
                FROM commissions 
                WHERE sale_key = %s 
                  AND status = 'Deferred'
                  AND is_active = 1
                ORDER BY commission_year, commission_month
            """
            df_defer = pd.read_sql_query(query_defer, self.app_container.pg_engine, params=(self.sale_key,))
            
            count_defer = len(df_defer)
            self.defer_title_label.configure(
                text=f"รายการที่ถูกเลื่อน (Deferred SOs - ทั้งหมดที่ค้างอยู่)  —  รวม {count_defer} รายการ"
            )
            
            for _, row in df_defer.iterrows():
                m = int(row['commission_month'])
                y = int(row['commission_year']) + 543
                month_str = thai_months[m-1]
                target_period = f"{month_str} {y}"
                reason = row['rejection_reason'] or "-"
                
                self.tree_defer.tag_configure('normal_text', foreground="black")

                self.tree_defer.insert("", "end", values=(
                    row['so_number'], row['customer_name'], target_period, reason
                ), tags=('normal_text',))

        except Exception as e:
            print(f"Error loading commission status: {e}")

    def _load_rejected_tasks(self):
        for widget in self.rejected_frame.winfo_children(): widget.destroy()
        try:
            # 1. แก้ไข Query: เพิ่ม 'Defer Requested' เข้าไปในรายการสถานะ
            query = """
                SELECT * FROM commissions 
                WHERE sale_key = %s 
                AND status IN ('Rejected by SM', 'Rejected by HR', 'Deferred by SM', 'Deferred by HR', 'Draft', 'Defer Requested') 
                AND is_active = 1 
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))
            
            if df.empty:
                CTkLabel(self.rejected_frame, text="ไม่มีงานที่ต้องดำเนินการ").pack(pady=20)
                return
            
            for _, row in df.iterrows():
                status = row['status']
                
                # 2. ตั้งค่าสีและข้อความตามสถานะ
                if status == 'Defer Requested':
                    # 🟢 [แก้ไขใหม่] กรณีบัญชีขอเลื่อน: เซลล์กดตัดสินใจไม่ได้แล้ว ต้องรอ SM อย่างเดียว
                    card_color = "#FFF7ED" 
                    reason_color = "#C2410C" 
                    status_prefix = "รอ Manager อนุมัติเลื่อนจ่าย"
                    button_text = "รอผลตัดสินใจ" # เปลี่ยนข้อความปุ่ม
                    button_cmd = lambda: messagebox.showinfo("รอการอนุมัติ", "รายการนี้บัญชีขอเลื่อนจ่าย\nกรุณารอ Sale Manager เป็นผู้พิจารณาอนุมัติ/ไม่อนุมัติ")
                    
                elif 'Reject' in status:
                    card_color = "#FEF2F2" # สีแดงอ่อน
                    reason_color = "#B91C1C"
                    status_prefix = "ตีกลับ"
                    button_text = "แก้ไข"
                    button_cmd = lambda r=row: self._edit_and_close(r)
                    
                elif 'Defer' in status: # Deferred (รายการที่เลื่อนไปแล้ว แต่อาจจะค้างอยู่)
                    card_color = "#FEFCE8" 
                    reason_color = "#A16207"
                    status_prefix = "ถูกเลื่อน (Deferred)"
                    button_text = "ดูรายละเอียด"
                    button_cmd = lambda r=row: self._edit_and_close(r)
                    
                else: # Draft
                    card_color = "#F3F4F6" 
                    reason_color = "gray"
                    status_prefix = "ฉบับร่าง"
                    button_text = "แก้ไข"
                    button_cmd = lambda r=row: self._edit_and_close(r)
                
                # ... (ส่วนการสร้าง Card UI ให้คงเดิม แต่เปลี่ยนการเรียกใช้ตัวแปร) ...
                
                card = CTkFrame(self.rejected_frame, border_width=1, fg_color=card_color)
                card.pack(fill="x", padx=5, pady=4)
                card.grid_columnconfigure(0, weight=1)

                top_frame = CTkFrame(card, fg_color="transparent")
                top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(5,0))
                
                info = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}"
                CTkLabel(top_frame, text=info, font=CTkFont(size=14, weight="bold")).pack(side="left")

                # ใช้ button_text และ button_cmd ที่เรากำหนดไว้ข้างบน
                edit_button = CTkButton(top_frame, text=button_text, width=100, command=button_cmd)
                edit_button.pack(side="right")
                
                # แสดงเหตุผล
                reason_text = row['rejection_reason'] or "-"
                reason_label = CTkLabel(card, text=f"{status_prefix}: {reason_text}", text_color=reason_color, wraplength=700, justify="left", anchor="w")
                reason_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))

        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดรายการได้: {e}", parent=self)

    def _load_draft_tasks(self):
        for widget in self.draft_frame.winfo_children(): widget.destroy()
        try:
            # ✅ แก้ไข: เพิ่ม 'Draft' ใน Query เพื่อให้ดึงงานใหม่มาโชว์ได้
            query = """
                SELECT * FROM commissions 
                WHERE sale_key = %s AND status IN ('Draft', 'Original', 'Edited') AND is_active = 1 
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))
            if df.empty:
                CTkLabel(self.draft_frame, text="ไม่มีฉบับร่าง").pack(pady=20)
                return
            for _, row in df.iterrows():
                card = CTkFrame(self.draft_frame, border_width=1)
                card.pack(fill="x", padx=5, pady=3)
                
                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.pack(side="left", fill="both", expand=True)
                CTkLabel(info_frame, text=f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}").pack(side="left", padx=10, pady=10)
                info_frame.bind("<Double-1>", lambda e, r=row: self._edit_and_close(r))

                btn_frame = CTkFrame(card, fg_color="transparent")
                btn_frame.pack(side="right", padx=10)
                
                # ปุ่มคลิกแก้ไข (เพื่อความชัดเจน)
                CTkButton(btn_frame, text="✏️ แก้ไข", width=60, command=lambda r=row: self._edit_and_close(r)).pack(side="left", padx=2)
                
                # ปุ่ม Copy Shortnote (สีเขียว LINE)
                CTkButton(btn_frame, text="📋 Copy Shortnote", width=120, fg_color="#22C55E", hover_color="#16A34A", command=lambda r=row: self._copy_so_shortnote(r.to_dict())).pack(side="left", padx=2)
        except Exception as e: print(e)
            
    def _edit_and_close(self, row_data):
        self.commission_app._edit_history_item(row_data.to_dict())
        self.on_close()

class SubmitSODialog(CTkToplevel):
    def __init__(self, master, app_container, sale_key, sale_name):
        super().__init__(master)
        self.commission_app = master
        self.app_container = app_container
        self.sale_key = sale_key
        self.sale_name = sale_name
        self.checkbox_list = []

        self.title("เลือกรายการ SO ที่จะนำส่ง")
        self.geometry("500x500")
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Top Frame for Select All ---
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=15, pady=(10, 0), sticky="ew")
        self.select_all_var = tk.IntVar(value=0)
        self.select_all_checkbox = CTkCheckBox(top_frame, text="เลือกทั้งหมด", variable=self.select_all_var, command=self._toggle_all_checkboxes, font=CTkFont(weight="bold"))
        self.select_all_checkbox.pack(anchor="w")

        # --- Scrollable Frame for SO List ---
        self.scroll_frame = CTkScrollableFrame(self, label_text="รายการ SO ที่เป็นฉบับร่าง")
        self.scroll_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        
        # --- Bottom Frame for Buttons ---
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        button_frame.grid_columnconfigure((0,1), weight=1)
        
        self.submit_button = CTkButton(button_frame, text="ยืนยันการนำส่ง (0)", command=self._confirm_submission, state="disabled")
        self.submit_button.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        CTkButton(button_frame, text="ยกเลิก", fg_color="gray", command=self.destroy).grid(row=0, column=1, padx=(5,0), sticky="ew")

        self.after(50, self._populate_so_list)
        self.transient(master)
        self.grab_set()

    def _populate_so_list(self):
        """ดึงรายการ SO ฉบับร่างมาแสดง พร้อมผูก Event ให้ปุ่มกดส่งทำงานได้"""
        try:
            # 1. เคลียร์วิดเจ็ตเก่าและล้าง list ก่อนโหลดใหม่ (กันข้อมูลซ้ำ)
            for widget in self.scroll_frame.winfo_children():
                widget.destroy()
            self.checkbox_list = []

            # 2. Query ข้อมูล (รวมสถานะ 'Draft' ตามแผนงานใหม่)
            query = """
                SELECT id, so_number, customer_name 
                FROM commissions 
                WHERE sale_key = %s 
                AND status IN ('Draft', 'Original', 'Edited') 
                AND is_active = 1 
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.sale_key,))

            if df.empty:
                CTkLabel(self.scroll_frame, text="ไม่พบรายการที่เป็นฉบับร่าง").pack(pady=20)
                self.select_all_checkbox.configure(state="disabled")
                return

            # 3. สร้างรายการ Checkbox
            for _, row in df.iterrows():
                var = tk.IntVar(value=0)
                
                # 🔥 [จุดสำคัญ] ต้องมีบรรทัดนี้! เพื่อสั่งให้ปุ่ม "ยืนยันการนำส่ง" อัปเดตสถานะ (Enabled/Disabled)
                var.trace_add("write", self._update_submit_button_state)

                so_text = f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}"
                
                cb = CTkCheckBox(self.scroll_frame, text=so_text, variable=var)
                cb.pack(anchor="w", padx=10, pady=5)
                
                # เก็บตัวแปร var และ ID ไว้เพื่อใช้ตรวจสอบตอนกดส่ง
                self.checkbox_list.append((var, row['id'], row.to_dict()))

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดรายการ SO ได้: {e}", parent=self)
            print(f"Debug Error: {e}")

    def _toggle_all_checkboxes(self):
        is_selected = self.select_all_var.get()
        for var, _, _ in self.checkbox_list:
            var.set(is_selected)

    def _update_submit_button_state(self, *args):
        """นับจำนวนที่เลือก และเปิด/ปิดการใช้งานปุ่มส่งข้อมูล"""
        selected_count = sum(var.get() for var, _, _ in self.checkbox_list)
        self.submit_button.configure(text=f"ยืนยันการนำส่ง ({selected_count})")
        
        if selected_count > 0:
            self.submit_button.configure(state="normal") # ติ๊กแล้ว ปุ่มจะกดได้
        else:
            self.submit_button.configure(state="disabled") # ไม่ติ๊ก ปุ่มจะกดไม่ได้

    def _confirm_submission(self):
        selected_records = [(so_id, data) for var, so_id, data in self.checkbox_list if var.get() == 1]
        if not selected_records: return
        if not messagebox.askyesno("ยืนยัน", f"ส่ง SO จำนวน {len(selected_records)} รายการให้ผู้จัดการอนุมัติ?", parent=self): return
        
        ids = tuple(r[0] for r in selected_records)
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 1. เปลี่ยนสถานะ
                cursor.execute("UPDATE commissions SET status = 'Pending Sale Manager Approval' WHERE id IN %s", (ids,))
                # 2. ✅ แก้ไข: ดึง Sale Manager มาแจ้งเตือน
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Sale Manager' AND status = 'Active'")
                sm_keys = [row[0] for row in cursor.fetchall()]
                notif_data = []
                for _, data in selected_records:
                    for sm_key in sm_keys:
                        notif_data.append((sm_key, f"SO ใหม่รออนุมัติ: {data['so_number']}", False, data['id']))
                if notif_data:
                    psycopg2.extras.execute_values(cursor, "INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) VALUES %s", notif_data)
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ส่งรายการให้ผู้จัดการเรียบร้อยแล้ว", parent=self)
            self.commission_app._update_tasks_badge()
            self.destroy()
            
        except Exception as e: 
            if conn: conn.rollback()
            messagebox.showerror("Error", str(e), parent=self)
        finally:
            # สำคัญที่สุด: คืน Connection กลับเข้า Pool
            if conn: 
                self.app_container.release_connection(conn)

class SOSummaryExportDialog(CTkToplevel):
    """Dialog ให้ Sale เลือกช่วงข้อมูลแล้ว Export รายงานสรุป SO เป็น Excel"""

    THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    THAI_MONTHS_SHORT = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.",
                         "ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."]

    def __init__(self, master, app_container, sale_key, sale_name):
        super().__init__(master)
        self.app_container = app_container
        self.sale_key = sale_key
        self.sale_name = sale_name
        self.title("📊 สรุปรอบคอม — Export SO")
        self.geometry("540x480")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)

        now = datetime.now()
        year_be = now.year + 543

        # ── Header ──
        CTkLabel(self, text="📊 สรุปรอบคอม", font=CTkFont(size=16, weight="bold")).pack(pady=(20, 4))
        CTkLabel(self, text="เลือกช่วงข้อมูลที่ต้องการ Export", font=CTkFont(size=12),
                 text_color="gray50").pack(pady=(0, 14))

        # ── Mode selection ──
        self._mode_var = tk.StringVar(value="month")
        mode_frame = CTkFrame(self, fg_color="#F9FAFB", corner_radius=8)
        mode_frame.pack(fill="x", padx=20, pady=(0, 12))
        CTkRadioButton(mode_frame, text="เลือกตามรอบคอม (เดือน)", variable=self._mode_var,
                       value="month", command=self._on_mode_change).pack(anchor="w", padx=16, pady=(10, 4))
        CTkRadioButton(mode_frame, text="เลือกตามช่วงวันที่บันทึก", variable=self._mode_var,
                       value="daterange", command=self._on_mode_change).pack(anchor="w", padx=16, pady=(0, 10))

        # ── Dynamic input area ──
        self._input_frame = CTkFrame(self, fg_color="#EFF6FF", corner_radius=8)
        self._input_frame.pack(fill="x", padx=20, pady=(0, 12))
        self._input_frame.grid_columnconfigure(1, weight=1)

        # สร้าง widgets สำหรับ month mode
        self._month_var = tk.StringVar(value=self.THAI_MONTHS[now.month - 1])
        self._year_var  = tk.StringVar(value=str(year_be))
        year_list = [str(y) for y in range(year_be - 3, year_be + 2)]

        self._month_widgets = CTkFrame(self._input_frame, fg_color="transparent")
        CTkLabel(self._month_widgets, text="รอบคอมเดือน:", font=CTkFont(size=13)).pack(side="left", padx=(0, 8))
        CTkOptionMenu(self._month_widgets, variable=self._month_var,
                      values=self.THAI_MONTHS, width=150).pack(side="left", padx=(0, 8))
        CTkOptionMenu(self._month_widgets, variable=self._year_var,
                      values=year_list, width=90).pack(side="left")

        # สร้าง widgets สำหรับ daterange mode
        self._daterange_widgets = CTkFrame(self._input_frame, fg_color="transparent")
        CTkLabel(self._daterange_widgets, text="ตั้งแต่:", font=CTkFont(size=13)).grid(row=0, column=0, padx=(0, 8), pady=4, sticky="w")
        self._from_date = DateSelector(self._daterange_widgets)
        self._from_date.grid(row=0, column=1, sticky="w")
        CTkLabel(self._daterange_widgets, text="ถึง:", font=CTkFont(size=13)).grid(row=1, column=0, padx=(0, 8), pady=4, sticky="w")
        self._to_date = DateSelector(self._daterange_widgets)
        self._to_date.grid(row=1, column=1, sticky="w")

        self._on_mode_change()

        # ── Column preview ──
        preview = CTkFrame(self, fg_color="#F0FDF4", corner_radius=8)
        preview.pack(fill="x", padx=20, pady=(0, 14))
        CTkLabel(preview, text="คอลัมน์ที่จะ Export:", font=CTkFont(size=12, weight="bold"),
                 text_color="#15803D").pack(anchor="w", padx=14, pady=(8, 2))
        CTkLabel(preview, text="เลขที่ SO  |  ชื่อลูกค้า  |  ยอดขาย (บาท)  |  พนักงานขาย  |  สถานะ",
                 font=CTkFont(size=11), text_color="#374151").pack(anchor="w", padx=14, pady=(0, 8))

        # ── Buttons ──
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=4)
        CTkButton(btn_frame, text="📥 Export Excel", font=CTkFont(size=14, weight="bold"),
                  fg_color="#2563EB", hover_color="#1D4ED8", height=42, width=160,
                  command=self._export).pack(side="left", padx=8)
        CTkButton(btn_frame, text="ยกเลิก", fg_color="#6B7280", hover_color="#4B5563",
                  height=42, width=100, command=self.destroy).pack(side="left", padx=8)

        self.focus()

    def _on_mode_change(self):
        self._month_widgets.pack_forget()
        self._daterange_widgets.pack_forget()
        if self._mode_var.get() == "month":
            self._month_widgets.pack(padx=14, pady=12)
        else:
            self._daterange_widgets.pack(padx=14, pady=12)

    def _export(self):
        import pandas as pd
        from tkinter import filedialog

        try:
            engine = self.app_container.pg_engine
            month_num_map = {m: i+1 for i, m in enumerate(self.THAI_MONTHS)}

            if self._mode_var.get() == "month":
                m = month_num_map.get(self._month_var.get(), 0)
                y = int(self._year_var.get()) - 543
                if not m:
                    messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกรอบเดือน", parent=self)
                    return
                query = """
                    SELECT c.so_number, c.customer_name,
                           COALESCE(c.final_sales_amount, c.sales_service_amount, 0) AS amount,
                           su.sale_name, c.status,
                           c.commission_month, c.commission_year, c.rejection_reason
                    FROM commissions c
                    LEFT JOIN sales_users su ON c.sale_key = su.sale_key
                    WHERE c.sale_key = %(sk)s
                      AND c.commission_month = %(m)s AND c.commission_year = %(y)s
                      AND c.is_active = 1
                    ORDER BY c.timestamp DESC
                """
                df = pd.read_sql_query(query, engine, params={"sk": self.sale_key, "m": m, "y": y})
                period_label = f"{self._month_var.get()} {self._year_var.get()}"
            else:
                f_date = self._from_date.get_date()
                t_date = self._to_date.get_date()
                if not f_date or not t_date:
                    messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกช่วงวันที่", parent=self)
                    return
                query = """
                    SELECT c.so_number, c.customer_name,
                           COALESCE(c.final_sales_amount, c.sales_service_amount, 0) AS amount,
                           su.sale_name, c.status,
                           c.commission_month, c.commission_year, c.rejection_reason
                    FROM commissions c
                    LEFT JOIN sales_users su ON c.sale_key = su.sale_key
                    WHERE c.sale_key = %(sk)s
                      AND c.timestamp::date BETWEEN %(fd)s AND %(td)s
                      AND c.is_active = 1
                    ORDER BY c.timestamp DESC
                """
                df = pd.read_sql_query(query, engine, params={"sk": self.sale_key, "f_date": f_date, "t_date": t_date, "fd": f_date, "td": t_date})
                period_label = f"{f_date} ถึง {t_date}"

            if df.empty:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบรายการ SO ในช่วงที่เลือก", parent=self)
                return

            # แปลง status
            def map_status(row):
                s = str(row.get("status", ""))
                if s == "Paid":
                    return "จ่ายค่าคอมแล้ว"
                if s == "HR Verified":
                    return "HR ตรวจสอบแล้ว (รอจ่าย)"
                if s == "Defer Requested":
                    return "ขอเลื่อน (รออนุมัติ)"
                if s == "Deferred":
                    m_num = int(row.get("commission_month") or 0)
                    y_num = int(row.get("commission_year") or 0)
                    if m_num and y_num:
                        # commission_month คือเดือนเป้าหมาย, ต้นทางคือเดือนก่อนหน้า
                        if m_num > 1:
                            prev_m, prev_y = m_num - 1, y_num
                        else:
                            prev_m, prev_y = 12, y_num - 1
                        return (f"โดนเลื่อน {self.THAI_MONTHS_SHORT[prev_m-1]} {prev_y+543}"
                                f" → {self.THAI_MONTHS_SHORT[m_num-1]} {y_num+543}")
                    return "โดนเลื่อน"
                if s in ("Cancelled",):
                    return "ยกเลิก"
                return "รอจ่ายค่าคอม"

            df["สถานะ"] = df.apply(map_status, axis=1)
            df_export = df.rename(columns={
                "so_number":     "เลขที่ SO",
                "customer_name": "ชื่อลูกค้า",
                "amount":        "ยอดขาย (บาท)",
                "sale_name":     "พนักงานขาย",
            })[["เลขที่ SO", "ชื่อลูกค้า", "ยอดขาย (บาท)", "พนักงานขาย", "สถานะ"]]

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"SO_Summary_{self.sale_name}_{period_label}.xlsx",
                parent=self
            )
            if not save_path:
                return

            df_export.to_excel(save_path, index=False)
            messagebox.showinfo("สำเร็จ", f"Export เรียบร้อยแล้ว\n{save_path}", parent=self)
            self.destroy()

        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)


class DeferralNoticeDialog(CTkToplevel):
    """Popup บังคับให้ Sale กด 'รับทราบ' ก่อนใช้งานระบบได้ เมื่อมี SO ที่ Manager ตัดสินใจเรื่องเลื่อนคอม"""
    def __init__(self, master, app_container, notifications):
        super().__init__(master)
        self.app_container = app_container
        self.notifications = notifications
        self.title("⚠️ แจ้งเตือนการเลื่อนรอบคอม")
        self.geometry("620x480")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # บังคับต้องกด รับทราบ
        self.grid_columnconfigure(0, weight=1)

        CTkLabel(self, text="⚠️ มี SO ของคุณที่ Manager ตัดสินใจแล้ว",
                 font=CTkFont(size=17, weight="bold"), text_color="#DC2626").pack(pady=(22, 4))
        CTkLabel(self, text="กรุณาอ่านและกดรับทราบเพื่อยืนยันว่าคุณได้รับทราบข้อมูลแล้ว",
                 font=CTkFont(size=12), text_color="gray40").pack(pady=(0, 12))

        scroll = CTkScrollableFrame(self, height=260)
        scroll.pack(fill="x", padx=20, pady=(0, 8))
        scroll.grid_columnconfigure(0, weight=1)

        for i, notif in enumerate(notifications):
            msg = str(notif['message']).replace('[DEFER]', '').strip()
            bg = "#FEF9C3" if "อนุมัติ" in msg else "#FEE2E2"
            card = CTkFrame(scroll, fg_color=bg, corner_radius=8)
            card.pack(fill="x", pady=4, padx=2)
            CTkLabel(card, text=msg, font=CTkFont(size=12), justify="left",
                     wraplength=540, text_color="#1E293B").pack(padx=12, pady=10, anchor="w")

        CTkButton(self, text="✅ รับทราบแล้ว", font=CTkFont(size=15, weight="bold"),
                  fg_color="#16A34A", hover_color="#15803D", height=48,
                  command=self._acknowledge).pack(pady=16, padx=40, fill="x")

    def _acknowledge(self):
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                ids = [n['id'] for n in self.notifications]
                cursor.executemany("UPDATE notifications SET is_read = TRUE WHERE id = %s",
                                   [(i,) for i in ids])
            conn.commit()
            self.app_container.release_connection(conn)
        except Exception as e:
            print(f"[DeferralNotice] acknowledge error: {e}")
        self.destroy()

    def _revert_withdraw_after_windows_set_titlebar_color(self):
        try:
            if self.winfo_exists():
                super()._revert_withdraw_after_windows_set_titlebar_color()
        except Exception:
            pass


class MissedSONoticeDialog(CTkToplevel):
    """Popup แจ้งเตือน SO ที่รอบเดือนผ่านไปแล้วแต่ยังไม่ถูกจ่ายค่าคอม"""

    STATUS_THAI = {
        'Forwarded_To_HR': 'ส่งต่อให้ HR',
        'HR Verified':     'HR ตรวจสอบแล้ว',
        'Defer Requested': 'ขอเลื่อนจ่าย',
        'Deferred':        'ถูกเลื่อนการจ่าย',
    }
    THAI_MONTHS = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                   "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

    def __init__(self, master, rows):
        super().__init__(master)
        self.title("⚠️ พบ SO ที่ยังไม่ได้คิดค่าคอม")
        self.geometry("780x520")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # บังคับกด รับทราบ

        # ── Header ──────────────────────────────────────────────────
        CTkLabel(self, text="⚠️ พบ SO ที่รอบเดือนผ่านไปแล้วแต่ยังไม่ได้คิดค่าคอม",
                 font=CTkFont(size=16, weight="bold"),
                 text_color="#DC2626").pack(pady=(20, 4))
        CTkLabel(self,
                 text=f"มีทั้งหมด {len(rows)} SO  |  กรุณาแจ้ง HR เพื่อดำเนินการต่อ",
                 font=CTkFont(size=12), text_color="gray40").pack(pady=(0, 12))

        # ── Table ────────────────────────────────────────────────────
        from tkinter import ttk
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("MissedSO.Treeview",
                        background="white", foreground="#1E293B",
                        rowheight=30, fieldbackground="white",
                        font=('TH Sarabun New', 12))
        style.configure("MissedSO.Treeview.Heading",
                        background="#FEE2E2", foreground="#991B1B",
                        font=('TH Sarabun New', 12, 'bold'), relief="flat")
        style.map("MissedSO.Treeview",
                  background=[('selected', '#FEF2F2')],
                  foreground=[('selected', '#1E293B')])

        # ── Acknowledge button (pack ก่อนเพื่อให้ติดล่างเสมอ) ──────────
        CTkButton(self, text="✅ รับทราบแล้ว",
                  font=CTkFont(size=15, weight="bold"),
                  fg_color="#DC2626", hover_color="#B91C1C",
                  height=46, command=self.destroy
                  ).pack(side="bottom", pady=14, padx=40, fill="x")

        # ── Table ────────────────────────────────────────────────────
        frame = CTkFrame(self, fg_color="white", corner_radius=8)
        frame.pack(side="top", fill="both", expand=True, padx=20, pady=(0, 4))

        cols = ('so_number', 'customer', 'amount', 'period', 'status')
        tree = ttk.Treeview(frame, columns=cols, show='headings',
                            style="MissedSO.Treeview", height=12)
        tree.heading('so_number', text='เลขที่ SO')
        tree.heading('customer',  text='ลูกค้า')
        tree.heading('amount',    text='ยอดขาย (บาท)')
        tree.heading('period',    text='รอบคอมที่ค้าง')
        tree.heading('status',    text='สถานะปัจจุบัน')

        tree.column('so_number', width=130, anchor='center')
        tree.column('customer',  width=200, anchor='w')
        tree.column('amount',    width=130, anchor='e')
        tree.column('period',    width=110, anchor='center')
        tree.column('status',    width=160, anchor='center')

        tree.tag_configure('odd',  background='#FFF7F7')
        tree.tag_configure('even', background='white')

        for idx, r in enumerate(rows):
            amount = (float(r.get('sales_service_amount') or 0)
                      + float(r.get('cutting_drilling_fee') or 0)
                      + float(r.get('other_service_fee') or 0))
            m = int(r.get('commission_month') or 0)
            y = int(r.get('commission_year') or 0)
            period_str = f"{self.THAI_MONTHS[m] if 0 < m <= 12 else m} {y + 543 if y else ''}"
            status_str = self.STATUS_THAI.get(r.get('status', ''), r.get('status', ''))
            tag = 'odd' if idx % 2 else 'even'
            tree.insert('', 'end', tags=(tag,), values=(
                r.get('so_number', '-'),
                r.get('customer_name', '-'),
                f"{amount:,.0f}",
                period_str,
                status_str,
            ))

        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)

    def _revert_withdraw_after_windows_set_titlebar_color(self):
        try:
            if self.winfo_exists():
                super()._revert_withdraw_after_windows_set_titlebar_color()
        except Exception:
            pass


class CommissionApp(CTkFrame):
    def __init__(self, master, sale_key=None, sale_name=None, app_container=None, show_logout_button=True, user_role=None, create_default_header=True):
        super().__init__(master, corner_radius=0, fg_color=app_container.THEME["sale"]["bg"])
        self.master = master
        self.app_container = app_container
        self.sale_key = sale_key or "UNKNOWN_SALE_KEY"
        self.sale_name = sale_name or "Unknown Sales User"
        self.user_role = user_role 
        self.theme = app_container.THEME["sale"]
        self.pg_engine = app_container.pg_engine
        self.show_logout_button = show_logout_button

        self.dropdown_style = {
            "fg_color": "white",
            "text_color": "black",
            "button_color": self.theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }

        self.editing_record_id = None
        self.history_window = None
        self.customer_data = {}
        self.customer_codes = []
        self.customer_completion_data = [] 
        self.so_form_widgets = {}
        self.header_map = app_container.HEADER_MAP

        self._create_string_vars()
        self.tasks_window = None
        self.tasks_button = None
        self.polling_job_id = None
        
        # 1. กำหนด Layout หลักของ Frame นี้ก่อน
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 2. สร้าง Header และวางในแถวที่ 0
        if create_default_header:
            self._create_header()

        # ==========================================================
        # 🔥 แก้ไขตรงนี้: สร้าง Tabview สำหรับแยกฟอร์ม และ Daily Report
        # ==========================================================
        self.main_tabview = CTkTabview(self, corner_radius=10, fg_color="transparent")
        self.main_tabview.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")

        # สร้าง 2 แท็บ
        self.tab_form = self.main_tabview.add("📝 สร้าง/แก้ไข Sales Order")
        self.tab_report = self.main_tabview.add("📊 รายงานประจำวัน (Daily Report)")
        self.tab_outstanding = self.main_tabview.add("💸 ยอดค้างชำระ")
        self.tab_so_edit = self.main_tabview.add("✏️ แก้ไข SO")

        self.tab_form.grid_columnconfigure(0, weight=1)
        self.tab_form.grid_rowconfigure(0, weight=1)
        self.tab_report.grid_columnconfigure(0, weight=1)
        self.tab_report.grid_rowconfigure(0, weight=1)
        self.tab_so_edit.grid_columnconfigure(0, weight=1)
        self.tab_so_edit.grid_rowconfigure(0, weight=1)

        # 3. เอา Scrollable Container ย้ายไปใส่ใน self.tab_form (แท็บที่ 1)
        self.scrollable_main_container = CTkScrollableFrame(self.tab_form, fg_color="transparent")
        self.scrollable_main_container.pack(fill="both", expand=True)

        # 4. กำหนด Layout ภายใน Scrollable Container
        self.scrollable_main_container.grid_columnconfigure(0, weight=1, uniform="group1")
        self.scrollable_main_container.grid_columnconfigure(1, weight=1, uniform="group1")
        self.scrollable_main_container.grid_rowconfigure(0, weight=1)

        # 5. สร้าง Frame ซ้าย-ขวา และวางใน Scrollable Container
        self.left_frame = CTkFrame(self.scrollable_main_container, fg_color="transparent")
        self.left_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")

        self.right_frame = CTkFrame(self.scrollable_main_container, fg_color="transparent")
        self.right_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")

        # 6. สร้างฟอร์มทั้งหมดลงใน Frame ซ้าย-ขวา
        self._populate_all_forms()
        
        # 7. โหลดข้อมูลและผูก Event ต่างๆ
        self._load_customer_data()
        self._bind_events()
        
        # 8. เริ่มการทำงานเบื้องหลัง
        self._start_polling()
        self.bind("<Destroy>", self._on_destroy)

        # 9. แจ้งเตือน SO เลื่อนคอมที่ยังไม่รับทราบ (Sale เท่านั้น)
        if self.user_role and self.user_role.lower() not in ('sales manager', 'director', 'hr', 'sale support'):
            self.after(900, self._check_pending_deferrals)
            self.after(1500, self._check_missed_sos)


        # ==========================================================
        # 🔥 9. เรียก DailyReportWidget มาวางใน แท็บที่ 2 (tab_report)
        # ==========================================================
        # (อย่าลืม import DailyReportWidget ไว้ที่ส่วนบนสุดของไฟล์ commission_app.py ด้วยนะครับ)
        current_role = str(getattr(self, 'user_role', '')).lower()
        if current_role in ['sale support', 'sale manager', 'admin', 'director']:
            filter_key = None
        else:
            filter_key = self.sale_key

        self.daily_report_view = DailyReportWidget(
            self.tab_report, 
            app_container=self.app_container,
            sale_key_filter=filter_key  # 🟢 เปลี่ยนมาใช้ตัวแปรที่เช็คสิทธิ์แล้ว
        )
        self.daily_report_view.pack(fill="both", expand=True)

        self.outstanding_view = OutstandingDashboardTab(
            self.tab_outstanding,
            app_container=self.app_container,
            sale_key_filter=filter_key  # 🟢 เปลี่ยนมาใช้ตัวแปรที่เช็คสิทธิ์แล้ว
        )
        self.outstanding_view.pack(fill="both", expand=True)

        # ✅ แท็บแก้ไข SO (Sale สามารถแก้ไขได้บางฟิลด์)
        self.so_edit_view = SOEditTabView(
            self.tab_so_edit,
            app_container=self.app_container,
            sale_key=self.sale_key,
            sale_name=self.sale_name,
        )
        self.so_edit_view.pack(fill="both", expand=True)

    def _open_my_tasks_window(self):
        if self.tasks_window is None or not self.tasks_window.winfo_exists():
            self.tasks_window = SalesTasksWindow(self, app_container=self.app_container, sale_key=self.sale_key)
        else:
            self.tasks_window.focus()

    def _update_tasks_badge(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # ✅ เพิ่ม 'Draft' ในการนับจำนวนงานที่ค้างใน My Tasks
                cursor.execute("""
                    SELECT COUNT(*) FROM commissions 
                    WHERE sale_key = %s 
                    AND status IN ('Draft', 'Original', 'Edited', 'Rejected by SM', 'Rejected by HR', 'Defer Requested') 
                    AND is_active = 1
                """, (self.sale_key,))
                total = cursor.fetchone()[0]
                
            if hasattr(self, 'tasks_button') and self.tasks_button.winfo_exists(): 
                self.tasks_button.configure(text=f"งานของฉัน 🔔 ({total})")
        except Exception as e: 
            print(f"Error update tasks badge: {e}")
        finally:
            # สำคัญที่สุด: คืน Connection กลับเข้า Pool
            if conn: 
                self.app_container.release_connection(conn)

    def _check_pending_deferrals(self):
        """ดึง notifications [DEFER] ที่ยังไม่รับทราบ แล้วแสดง DeferralNoticeDialog"""
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute(
                    "SELECT id, message FROM notifications "
                    "WHERE user_key_to_notify = %s AND is_read = FALSE AND message LIKE '[DEFER]%%' "
                    "ORDER BY id ASC",
                    (self.sale_key,)
                )
                rows = cursor.fetchall()
            self.app_container.release_connection(conn)

            if not rows:
                return

            try:
                if self.winfo_exists():
                    DeferralNoticeDialog(self, self.app_container, [dict(r) for r in rows])
            except Exception as e:
                if "bad window path name" not in str(e):
                    print(f"[pending_deferral_check] {e}")
        except Exception as e:
            if "bad window path name" not in str(e):
                print(f"[pending_deferral_check] {e}")

    def _check_missed_sos(self):
        """แจ้งเตือน SO ที่รอบเดือนผ่านไปแล้วแต่ยังไม่ถูกจ่ายค่าคอม (แสดงครั้งเดียวต่อ session)"""
        try:
            today = datetime.now()
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("""
                    SELECT so_number, customer_name,
                           sales_service_amount, cutting_drilling_fee, other_service_fee,
                           commission_month, commission_year, status
                    FROM commissions
                    WHERE sale_key = %s
                      AND is_active = 1
                      AND status IN ('Forwarded_To_HR', 'HR Verified', 'Defer Requested')
                      AND (
                          commission_year < %s
                          OR (commission_year = %s AND commission_month < %s)
                      )
                    ORDER BY commission_year ASC, commission_month ASC, so_number ASC
                """, (self.sale_key, today.year, today.year, today.month))
                rows = cursor.fetchall()
            self.app_container.release_connection(conn)

            if not rows:
                return

            try:
                if self.winfo_exists():
                    MissedSONoticeDialog(self, [dict(r) for r in rows])
            except Exception as e:
                if "bad window path name" not in str(e):
                    print(f"[check_missed_sos] dialog error: {e}")
        except Exception as e:
            print(f"[check_missed_sos] query error: {e}")

    def _start_polling(self):
        self._update_tasks_badge()
        self.polling_job_id = self.after(30000, self._start_polling)

    def _on_destroy(self, event):
        if hasattr(event, 'widget') and event.widget is self:
            if self.polling_job_id:
                self.after_cancel(self.polling_job_id)

    def _create_header(self):
        self.header_frame = CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        # --- START: แก้ไขส่วนนี้ทั้งหมด ---
        # กำหนดให้คอลัมน์ซ้าย (ชื่อ) ขยายตัว และคอลัมน์ขวา (ปุ่ม) ไม่ขยาย
        self.header_frame.grid_columnconfigure(0, weight=1)
        
        CTkLabel(self.header_frame, text=f"ฝ่ายขาย: {self.sale_name} ({self.sale_key})", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).grid(row=0, column=0, sticky="w")
        
        # สร้าง Frame สำหรับปุ่มต่างๆ ทางขวา และใช้ grid วาง
        button_container = CTkFrame(self.header_frame, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="e")
        
        # ใช้ grid วางปุ่มภายใน button_container
        CTkButton(button_container, text="📊 สรุปรอบคอม",
                  command=lambda: SOSummaryExportDialog(self, self.app_container, self.sale_key, self.sale_name),
                  fg_color="#2563EB", hover_color="#1D4ED8"
                  ).grid(row=0, column=0, padx=(0, 6))

        self.tasks_button = CTkButton(button_container, text="งานของฉัน 🔔 (0)", command=self._open_my_tasks_window)
        self.tasks_button.grid(row=0, column=1, padx=6)

        if self.show_logout_button:
            CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen,
                      fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F",
                      border_width=2, hover_color="#FFEBEE").grid(row=0, column=2, padx=(0, 10))
    
    
    def _refresh_history_if_open(self):
        if self.history_window and self.history_window.winfo_exists():
            if hasattr(self.history_window, '_populate_history_table'):
                self.history_window._populate_history_table()

    def _edit_history_item(self, row_data):
        record_status = row_data.get('status')

        # --- START: แก้ไขเงื่อนไขการตรวจสอบสถานะตรงนี้ ---
        # สถานะที่ "ห้ามแก้ไขเด็ดขาด" คือเมื่อ HR ตรวจสอบแล้ว หรือจ่ายเงินไปแล้ว
        uneditable_statuses = ('Paid', 'HR Verified', 'Cancelled')

        if record_status in uneditable_statuses:
            messagebox.showwarning("ไม่สามารถแก้ไขได้", 
                                 f"รายการนี้มีสถานะ '{record_status}' ซึ่งเป็นขั้นตอนสุดท้ายแล้ว จึงไม่สามารถแก้ไขได้", 
                                 parent=self)
            return
        # --- END ---

        self._clear_form(confirm=False)
        self.editing_record_id = int(row_data.get('id'))
        if self.history_window and self.history_window.winfo_exists():
            self.history_window.destroy()
            self.history_window = None
        if self.tasks_window and self.tasks_window.winfo_exists():
             self.tasks_window.destroy()
             self.tasks_window = None
        self._populate_form_from_data(row_data)
        messagebox.showinfo("โหลดข้อมูลเพื่อแก้ไข", f"โหลด SO Number: {row_data.get('so_number')} สำหรับแก้ไข", parent=self)

    def _on_history_so_select(self, row_data):
        """
        Callback function เมื่อมีการดับเบิลคลิกที่ SO ในหน้าต่างประวัติ
        (เวอร์ชันแก้ไข: เพิ่ม 'Draft' ลงในสถานะที่อนุญาตให้แก้ไข)
        """
        if row_data is None:
            return

        so_status = row_data.get('status')
        so_number = row_data.get('so_number')
        
        # --- จุดที่แก้ไข: เพิ่ม 'Draft' เข้าไปในรายการนี้ ---
        editable_statuses = ['Original', 'Edited', 'Rejected by SM', 'Rejected by HR', 'Deferred by SM', 'Deferred by HR', 'Draft']
        # --------------------------------------------------

        if so_status in editable_statuses:
            if messagebox.askyesno("โหลดข้อมูล", f"คุณต้องการโหลดข้อมูล SO: {so_number} เพื่อแก้ไขหรือไม่?"):
                
                # 1. ล้างข้อมูลในฟอร์มให้เป็นค่าเริ่มต้นก่อน เพื่อป้องกันข้อมูลเก่าค้าง
                self._clear_form(confirm=False)
                
                # 2. (สำคัญที่สุด) ตั้งค่า ID ที่กำลังแก้ไข เพื่อให้โปรแกรมเข้าสู่ "โหมดแก้ไข"
                self.editing_record_id = int(row_data.get('id'))
                
                # 3. โหลดข้อมูลจากแถวที่เลือกมาใส่ในฟอร์ม
                # (ใช้ .to_dict() หากข้อมูลมาเป็น Series จาก Pandas)
                data_to_load = row_data.to_dict() if hasattr(row_data, 'to_dict') else row_data
                self._populate_form_from_data(data_to_load) 
                
                # 4. ปิดหน้าต่างประวัติ (เพื่อให้กลับไปหน้าจอหลัก)
                if self.history_window and self.history_window.winfo_exists():
                    self.history_window.destroy()
                    self.history_window = None
        else:
            messagebox.showinfo(
                "ไม่สามารถแก้ไขได้",
                f"SO: {so_number} อยู่ในสถานะ '{so_status}'\n\n"
                "ข้อมูลได้ถูกส่งต่อไปในกระบวนการแล้ว จึงไม่สามารถแก้ไขได้ในขั้นตอนนี้",
                parent=self.history_window
            )

    def _save_data(self):
        """บันทึกข้อมูล SO และเก็บประวัติการแก้ไข (Audit Trail)"""
        form_data = self._gather_data_from_form()
        is_valid, message = self._validate_form(form_data)
        if not is_valid:
            messagebox.showerror("ข้อมูลไม่ถูกต้อง", message, parent=self)
            return

        # ตรวจสอบวันที่จัดส่งเกิน cutoff — บังคับให้เปลี่ยนรอบคอมก่อนบันทึก
        if self._is_delivery_date_over_cutoff():
            day_str = self.delivery_date_selector.day_var.get()
            month_str = self.delivery_date_selector.month_var.get()
            thai_month_map = {"ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
                              "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
                              "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12}
            month_num = thai_month_map.get(month_str, 0)
            cutoff = 21 if month_num in (2, 12) else 25
            next_month = month_num + 1 if month_num < 12 else 1
            next_month_thai = self.thai_months[next_month - 1]
            messagebox.showwarning(
                "⚠️ วันที่จัดส่งเกินวันตัดรอบ",
                f"วันที่จัดส่ง ({day_str} {month_str}) เกินวันตัดรอบที่ {cutoff} ของเดือนนี้\n\n"
                f"กรุณาเปลี่ยน 'รอบเดือนคอม' เป็น {next_month_thai} ก่อนบันทึก",
                parent=self
            )
            return

        if self.editing_record_id:
            if not messagebox.askyesno("ยืนยัน", "คุณต้องการบันทึกการเปลี่ยนแปลงนี้ใช่หรือไม่?", parent=self):
                return

        # 🟢 [เพิ่มตรงนี้!] ตรวจสอบและบันทึกลูกค้าใหม่ลง Database Master ก่อน
        if form_data.get('customer_type') == 'ลูกค้าใหม่':
            try:
                # 1. สั่งบันทึกลูกค้าใหม่ลงฐานข้อมูล
                self._handle_new_customer(form_data)
                
                # 2. โหลดข้อมูลลูกค้าเข้า Memory ใหม่ เพื่อให้ช่องค้นหาอัปเดตทันที
                self._load_customer_data() 
                
                # 3. ปรับสถานะในฟอร์มกลับไปเป็น "ลูกค้าเก่า" (เพื่อไม่ให้เซฟลูกค้าซ้ำซ้อนถ้าเผลอกดเซฟซ้ำ)
                form_data['customer_type'] = 'ลูกค้าเก่า'
                self.customer_type_var.set("ลูกค้าเก่า")
                self._toggle_customer_fields()
                
            except ValueError as ve:
                messagebox.showerror("รหัสลูกค้าซ้ำ", str(ve), parent=self)
                return
            except Exception as e:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกลูกค้าใหม่ได้: {e}", parent=self)
                return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                if self.editing_record_id:
                    # 1. ยกเลิก Record เดิม (ทำให้เป็น Inactive)
                    cursor.execute("UPDATE commissions SET is_active = 0 WHERE id = %s", (self.editing_record_id,))
                    
                    # 2. ปรับสถานะใบใหม่ (ให้เป็น Edited เสมอ รอเซลส์ไปกด "นำส่ง" เองทีหลัง)
                    form_data['status'] = 'Edited'
                    form_data['original_id'] = self.editing_record_id
                else:
                    form_data['status'] = 'Draft'

                # 3. บันทึกข้อมูล SO ลงฐานข้อมูล
                self._perform_db_insert(form_data)
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลเรียบร้อยแล้ว", parent=self)
            self._clear_form(confirm=False)
            self._update_tasks_badge()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"บันทึกไม่สำเร็จ: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_section_frame(self, parent, title):
        frame = CTkFrame(parent, corner_radius=10, border_width=1)
        frame.grid_columnconfigure(1, weight=1)
        CTkLabel(frame, text=title, font=CTkFont(size=18, weight="bold"), text_color=self.theme["header"]).grid(row=0, column=0, columnspan=3, padx=15, pady=(10, 5), sticky="w")
        return frame

    def _add_form_row(self, parent, label_text, widget, row_index, columnspan=2, padx=(10, 15), pady=4):
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=(15, 10), pady=pady, sticky="w")
        widget.grid(row=row_index, column=1, columnspan=columnspan, padx=padx, pady=pady, sticky="ew")

    def _add_item_row_with_vat(self, parent, label_text, entry_widget, radio_var, row_index, padx=(10, 15), pady=5):
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=15, pady=pady, sticky="w")
        entry_widget.grid(row=row_index, column=1, padx=padx, pady=pady, sticky="ew")

        radio_frame = CTkFrame(parent, fg_color="transparent")
        radio_frame.grid(row=row_index, column=2, padx=padx, pady=pady, sticky="w")
        CTkRadioButton(radio_frame, text="VAT", variable=radio_var, value="VAT").pack(side="left", padx=5)
        CTkRadioButton(radio_frame, text="CASH", variable=radio_var, value="NO VAT").pack(side="left", padx=5)

    def _populate_all_forms(self):
        self._populate_sales_details_form(self.left_frame)
        self._populate_sales_services_frame(self.left_frame)
        self._populate_shipping_frame(self.left_frame)
        self._populate_fees_frame(self.left_frame)

        self._populate_additional_details_frame(self.left_frame)

        self._populate_other_expenses_frame(self.right_frame)
        self._populate_payment_frame(self.right_frame)
        self._populate_so_summary_frame(self.right_frame)
        self._populate_cash_verification_frame(self.right_frame)
        self._populate_action_frame(self.right_frame)

    def _create_string_vars(self):
        now = datetime.now()
        thai_months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_months = thai_months_list
        self.thai_month_map = {name: i+1 for i, name in enumerate(thai_months_list)}

        self.customer_type_var = tk.StringVar(value="ลูกค้าเก่า")
        self.credit_term_var = tk.StringVar(value="เงินสด")
        self.commission_month_var = tk.StringVar(value=self.thai_months[now.month - 1])
        self.commission_year_var = tk.StringVar(value=str(now.year + 543))
        self.payment1_percent_var = tk.StringVar(value="ระบุยอดเอง")
        self.payment2_percent_var = tk.StringVar(value="ระบุยอดเอง")
        self.sales_vat_calc_var = tk.StringVar(value="0.00")
        self.cutting_drilling_vat_calc_var = tk.StringVar(value="0.00")
        self.other_service_vat_calc_var = tk.StringVar(value="0.00")
        self.shipping_vat_calc_var = tk.StringVar(value="0.00")
        self.card_fee_vat_calc_var = tk.StringVar(value="0.00")
        self.relocation_vat_option_var = tk.StringVar(value="VAT")
        self.relocation_vat_calc_var = tk.StringVar(value="0.00")
        self.payment_total_var = tk.StringVar(value="0.00")
        self.so_subtotal_var = tk.StringVar(value="0.00")
        self.so_vat_var = tk.StringVar(value="0.00")
        self.so_grand_total_var = tk.StringVar(value="0.00")
        self.so_vs_payment_result_var = tk.StringVar(value="-")

        self.so_number_var = tk.StringVar(value="SO")

        self.difference_amount_var = tk.StringVar(value="0.00")
        self.balance_due_var = tk.StringVar(value="0.00")
        self.cash_product_input_var = tk.StringVar(value="0.00")
        self.cash_service_total_var = tk.StringVar(value="0.00")
        self.cash_required_total_var = tk.StringVar(value="0.00")
        self.cash_actual_payment_var = tk.StringVar(value="0.00")
        self.cash_verification_result_var = tk.StringVar(value="-")
        self.sales_service_vat_option = tk.StringVar(value="VAT")
        self.cutting_drilling_fee_vat_option = tk.StringVar(value="VAT")
        self.other_service_fee_vat_option = tk.StringVar(value="VAT")
        self.shipping_vat_option_var = tk.StringVar(value="VAT")
        self.credit_card_fee_vat_option_var = tk.StringVar(value="VAT")

        self.payment1_method_var = tk.StringVar(value="ชำระสด")
        self.payment2_method_var = tk.StringVar(value="ชำระสด")
        self.delivery_type_var = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")

        self.delivery_map_var = tk.StringVar(value="")
        self.onsite_contact_name_var = tk.StringVar(value="")
        self.onsite_contact_phone_var = tk.StringVar(value="")
        self.vehicle_type_var = tk.StringVar(value="-")
        self.order_pur_var = tk.StringVar(value="")

        self.special_request_var = tk.StringVar(value="") 
        self.unloading_status_var = tk.StringVar(value="ไม่รวมลง")

    def _force_uppercase_so_number(self, *args):
        current_text = self.so_number_var.get()
        new_text = current_text.upper()
        if new_text != current_text:
            self.so_number_var.set(new_text)

    def _on_delivery_type_change(self, *args):
        selected_option = self.delivery_type_var.get()

    # ตรวจสอบว่าวิดเจ็ตถูกสร้างขึ้นแล้วหรือยังก่อนที่จะใช้งาน
        if hasattr(self, 'date_to_wh_label') and self.date_to_wh_label.winfo_exists() and \
           hasattr(self, 'date_to_wh_selector') and self.date_to_wh_selector.winfo_exists():

         if "คลัง" in selected_option:
        # ถ้ามีคำว่า "คลัง" ในตัวเลือก ให้แสดงช่องวันที่
           self.date_to_wh_label.grid()
           self.date_to_wh_selector.grid()
         else:
        # ถ้าไม่มี ให้ซ่อน
           self.date_to_wh_label.grid_remove()
           self.date_to_wh_selector.grid_remove()

    def _bind_events(self):
        widgets_to_bind_names = [
            "sales_amount_entry", "cutting_drilling_fee_entry", "other_service_fee_entry",
            "shipping_cost_entry", "credit_card_fee_entry", "transfer_fee_entry",
            "wht_fee_entry", "coupon_value_entry", "giveaway_vat_entry", "giveaway_no_vat_entry",
            "brokerage_fee_entry", "payment1_amount_entry", "payment2_amount_entry",
            "cash_product_input_entry", "cash_actual_payment_entry",
            # <<< เพิ่มเติม: เพิ่ม relocation_cost_entry เข้าไปใน list นี้ >>>
            "relocation_cost_entry"
        ]

        # 1. ผูก Event คำนวณตัวเลข (เมื่อพิมพ์ในช่องกรอก)
        for widget_name in widgets_to_bind_names:
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                widget.bind("<KeyRelease>", self._update_final_calculations)

        # 2. ผูก Event คำนวณตัวเลข (เมื่อเปลี่ยน Radio Button)
        for var in [
            self.sales_service_vat_option, self.cutting_drilling_fee_vat_option,
            self.other_service_fee_vat_option, self.shipping_vat_option_var,
            self.credit_card_fee_vat_option_var, self.relocation_vat_option_var
        ]:
            var.trace_add("write", self._update_final_calculations)

        # 3. บังคับ SO Number เป็นตัวพิมพ์ใหญ่เสมอ
        self.so_number_var.trace_add("write", self._force_uppercase_so_number)
        self.so_number_entry.bind("<FocusIn>", lambda e: self.so_number_entry.icursor(tk.END))

        # 4. เรียกฟังก์ชันตรวจสอบการจัดส่งครั้งแรก
        self._on_delivery_type_change()

        # +++ [เพิ่มใหม่] 5. ผูก Event กับ Bill Date เพื่ออัปเดตรอบคอมฯ อัตโนมัติ +++
        try:
            # พยายามเข้าถึง Entry ภายใน DateSelector เพื่อ Bind Event
            # (เมื่อผู้ใช้เลือกวันที่เสร็จ หรือพิมพ์เสร็จแล้วกด Enter/คลิกที่อื่น)
            if hasattr(self.bill_date_selector, 'entry'):
                # ทำงานเมื่อเมาส์คลิกออก (FocusOut) หรือกด Enter
                self.bill_date_selector.entry.bind("<FocusOut>", self._auto_update_commission_period)
                self.bill_date_selector.entry.bind("<Return>", self._auto_update_commission_period)
                # ทำงานทันทีเมื่อพิมพ์ (KeyRelease) - แต่อาจจะรอพิมพ์ครบรูปแบบก่อน
                self.bill_date_selector.entry.bind("<KeyRelease>", self._auto_update_commission_period)
        except AttributeError:
            # กรณีที่ DateSelector ไม่มี attribute 'entry' ให้ข้ามไป (ป้องกัน Error)
            pass

    def _show_history(self):
        try:
            # vvvv แก้ไขบรรทัดนี้ vvvv
            if self.user_role and self.user_role.lower() == 'sale support':
                self.history_window = self.app_container.show_history_window(
                    support_user_key_filter=self.app_container.current_user_key,
                    sale_key_filter=None,
                    edit_callback=self._on_history_so_select
                )
            else:
                self.history_window = self.app_container.show_history_window(
                    sale_key_filter=self.sale_key,
                    support_user_key_filter=None,
                    edit_callback=self._on_history_so_select
                )
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถเปิดหน้าต่างประวัติได้: {e}", parent=self)

    def _populate_form_from_data(self, data):
        def set_entry_value(entry_widget, value):
            if isinstance(entry_widget, (NumericEntry, CTkEntry, AutoCompleteEntry)):
                is_readonly = entry_widget.cget("state") == "readonly"
                if is_readonly: entry_widget.configure(state="normal")
                entry_widget.delete(0, tk.END)
                if pd.notna(value) and value is not None:
                    if isinstance(value, (int, float)):
                        entry_widget.insert(0, f"{value:,.2f}")
                    else:
                        entry_widget.insert(0, str(value))
                if is_readonly: entry_widget.configure(state="readonly")

        def set_date_selector(selector_widget, date_str):
            if pd.notna(date_str) and date_str is not None:
                try:
                    if isinstance(date_str, datetime): dt_obj = date_str
                    elif isinstance(date_str, pd.Timestamp): dt_obj = date_str.to_pydatetime()
                    else: dt_obj = datetime.strptime(str(date_str), '%Y-%m-%d')
                    selector_widget.set_date(dt_obj)
                except (ValueError, TypeError):
                    selector_widget.set_date(None)
            else:
                selector_widget.set_date(None)

        def set_radio_button(radio_var, value, default="VAT"):
            if pd.notna(value) and value is not None:
                radio_var.set(str(value))
            else:
                radio_var.set(default)

        set_date_selector(self.bill_date_selector, data.get('bill_date'))
        month_from_data = data.get('commission_month')
        if pd.notna(month_from_data):
            try:
                month_int = int(utils.convert_to_float(month_from_data))
                if 1 <= month_int <= 12: self.commission_month_var.set(self.thai_months[month_int - 1])
            except (ValueError, TypeError): pass

        year_from_data = data.get('commission_year')
        if pd.notna(year_from_data):
            try:
                year_int = int(utils.convert_to_float(year_from_data))
                self.commission_year_var.set(str(year_int + 543))
            except (ValueError, TypeError): pass

        customer_type = data.get('customer_type', 'ลูกค้าเก่า')
        self.customer_type_var.set(customer_type)
        self._toggle_customer_fields()

        if customer_type == "ลูกค้าเก่า":
            set_entry_value(self.customer_id_entry, data.get('customer_id'))
            self._on_customer_id_selected(data.get('customer_id'))
        else:
            set_entry_value(self.new_customer_id_entry, data.get('customer_id'))
            set_entry_value(self.new_customer_name_entry, data.get('customer_name'))

        self.credit_term_var.set(data.get('credit_term', 'เงินสด'))
        self.so_number_var.set(data.get('so_number', 'SO'))

        set_entry_value(self.sales_amount_entry, data.get('sales_service_amount'))
        set_radio_button(self.sales_service_vat_option, data.get('sales_service_vat_option'))
        set_entry_value(self.cutting_drilling_fee_entry, data.get('cutting_drilling_fee'))
        set_radio_button(self.cutting_drilling_fee_vat_option, data.get('cutting_drilling_fee_vat_option'))
        set_entry_value(self.other_service_fee_entry, data.get('other_service_fee'))
        set_radio_button(self.other_service_fee_vat_option, data.get('other_service_fee_vat_option'))
        set_entry_value(self.shipping_cost_entry, data.get('shipping_cost'))
        set_radio_button(self.shipping_vat_option_var, data.get('shipping_vat_option'))
        set_date_selector(self.delivery_date_selector, data.get('delivery_date'))
        set_entry_value(self.credit_card_fee_entry, data.get('credit_card_fee'))
        set_radio_button(self.credit_card_fee_vat_option_var, data.get('credit_card_fee_vat_option'))
        set_entry_value(self.transfer_fee_entry, data.get('transfer_fee'))
        set_entry_value(self.wht_fee_entry, data.get('wht_3_percent'))
        set_entry_value(self.brokerage_fee_entry, data.get('brokerage_fee'))
        set_entry_value(self.coupon_value_entry, data.get('coupons'))
        set_entry_value(self.giveaway_vat_entry, data.get('giveaway_vat'))        # 🟢 ของใหม่
        set_entry_value(self.giveaway_no_vat_entry, data.get('giveaway_no_vat'))
        set_radio_button(self.relocation_vat_option_var, data.get('relocation_cost_vat_option'))

        # --- [แก้ไข] Payment Logic: จัดการยอดแยกย่อยให้ถูกต้อง ---
        val_p1 = data.get('payment1_amount', 0.0)
        val_p2 = data.get('payment2_amount', 0.0)
        
        # Fallback: ถ้ายอดแยกเป็น 0 แต่ยอดรวม (total) มีค่า (สำหรับข้อมูลเก่า) ให้เอายอดรวมไปใส่ช่อง 1 แทน
        total_payment = data.get('total_payment_amount', 0.0) or 0.0
        if (val_p1 == 0 or val_p1 is None) and (val_p2 == 0 or val_p2 is None) and total_payment > 0:
            val_p1 = total_payment

        set_entry_value(self.payment1_amount_entry, val_p1)
        self.payment1_percent_var.set("ระบุยอดเอง")
        
        set_entry_value(self.payment2_amount_entry, val_p2)
        self.payment2_percent_var.set("ระบุยอดเอง")
        # ---------------------------------------------------

        set_date_selector(self.payment1_date_selector, data.get('payment1_date'))
        set_date_selector(self.payment2_date_selector, data.get('payment2_date'))

        set_entry_value(self.cash_product_input_entry, data.get('cash_product_input'))
        set_entry_value(self.cash_actual_payment_entry, data.get('cash_actual_payment'))

        set_radio_button(self.delivery_type_var, data.get('delivery_type', 'ซัพพลายเออร์จัดส่ง'))
        set_entry_value(self.pickup_location_entry, data.get('pickup_location'))
        set_entry_value(self.relocation_cost_entry, data.get('relocation_cost'))
        set_date_selector(self.date_to_wh_selector, data.get('date_to_warehouse'))
        set_date_selector(self.date_to_customer_selector, data.get('date_to_customer'))
        set_entry_value(self.pickup_rego_entry, data.get('pickup_registration'))

        self.payment1_method_var.set(data.get('payment1_method', 'ไม่เลือก'))
        self.payment2_method_var.set(data.get('payment2_method', 'ไม่เลือก'))

        set_entry_value(self.delivery_map_entry, data.get('delivery_map'))
        set_entry_value(self.onsite_contact_name_entry, data.get('onsite_contact_name'))
        set_entry_value(self.onsite_contact_phone_entry, data.get('onsite_contact_phone'))
        self.vehicle_type_var.set(data.get('vehicle_type') or '-')
        self.order_pur_var.set(data.get('order_pur') or '')
        self.special_request_var.set(data.get('special_request') or '-')          # 🟢 เพิ่มบรรทัดนี้
        self.unloading_status_var.set(data.get('unloading_status', 'ไม่รวมลง'))   # 🟢 เพิ่มบรรทัดนี้

        self._update_final_calculations()

    def _load_customer_data(self):
        try:
            df = pd.read_sql("SELECT customer_name, customer_code, credit_term FROM customers ORDER BY customer_name", self.pg_engine)
            
            self.customer_completion_data = []
            
            for _, row in df.iterrows():
                # 🟢 ดึงข้อมูลและแปลงเป็น String พร้อมตัดช่องว่างซ้าย-ขวาทิ้งให้หมด (ป้องกันปัญหาพิมพ์หาไม่เจอ)
                raw_code = row.get('customer_code')
                raw_name = row.get('customer_name')
                
                code = str(raw_code).strip() if pd.notna(raw_code) else ""
                name = str(raw_name).strip() if pd.notna(raw_name) else ""
                term = str(row.get('credit_term', 'เงินสด')).strip() if pd.notna(row.get('credit_term')) else "เงินสด"
                
                # ถ้าไม่มีทั้งชื่อและรหัส ให้ข้ามไปเลย
                if not code and not name:
                    continue
                    
                display_text = f"{code} - {name}"
                
                self.customer_completion_data.append({
                    "name": name,
                    "code": code,
                    "term": term,
                    "display": display_text
                })
            
            # สร้าง Map สำหรับการอ้างอิงข้อมูล
            self.customer_data_map = {item['display']: item for item in self.customer_completion_data}

            # อัปเดตข้อมูลเข้าไปใน Dropdown/AutoComplete
            if hasattr(self, 'customer_id_entry') and self.customer_id_entry.winfo_exists():
                self.customer_id_entry.update_completion_list(self.customer_completion_data)

        except Exception as e:
            print(f"Error loading customer data: {e}")
            self.customer_completion_data = []
            self.customer_data_map = {}

    def _on_customer_id_selected(self, selection_data):
        """
        รองรับข้อมูล 2 รูปแบบ: Dictionary (ตอนเลือกจากลิสต์) หรือ String (ตอนโหลดข้อมูลประวัติ)
        """
        customer_name = ''
        credit_term = 'เงินสด'
        customer_code = ''

        if isinstance(selection_data, dict):
            # กรณีผู้ใช้คลิกเลือกจากรายการค้นหา
            customer_name = selection_data.get('name', '').strip()
            credit_term = selection_data.get('term', 'เงินสด').strip()
            customer_code = selection_data.get('code', '').strip()

            self.customer_id_entry.delete(0, tk.END)
            self.customer_id_entry.insert(0, customer_code)

        elif isinstance(selection_data, str) and selection_data:
            # กรณีโหลดข้อมูลเก่า ให้ตัดช่องว่างทิ้งก่อนค้นหา
            customer_code_to_find = selection_data.strip()
            
            found_customer = next((item for item in self.customer_completion_data if item.get('code') == customer_code_to_find), None)
            
            if found_customer:
                customer_name = found_customer.get('name', '').strip()
                credit_term = found_customer.get('term', 'เงินสด').strip()

        # อัปเดตช่องชื่อลูกค้าและเครดิตให้สอดคล้องกัน
        self.customer_name_entry.configure(state="normal")
        self.customer_name_entry.delete(0, tk.END)
        self.customer_name_entry.insert(0, customer_name)
        self.customer_name_entry.configure(state="readonly")
        
        self.credit_term_var.set(credit_term)

    def _populate_sales_details_form(self, parent):
        
        frame = self._create_section_frame(parent, "รายละเอียดการขาย")
        frame.pack(fill="x", pady=(0, 10))
        
        # --- Commission Period is now on top ---
        commission_period_outer_frame = CTkFrame(frame, fg_color="transparent")
        commission_period_outer_frame.grid(row=1, column=0, columnspan=3, padx=15, pady=4, sticky="ew") # <-- Changed to row=1
        CTkLabel(commission_period_outer_frame, text="รอบคอมมิชชั่น:", font=CTkFont(size=14)).pack(side="left")

        month_year_frame = CTkFrame(commission_period_outer_frame, fg_color="transparent")
        month_year_frame.pack(side="left")

        self.commission_month_menu = CTkOptionMenu(month_year_frame, variable=self.commission_month_var, values=list(self.thai_month_map.keys()), width=120, **self.dropdown_style)
        self.commission_month_menu.pack(side="left", padx=5)

        current_year_be = datetime.now().year + 543
        year_list = [str(y) for y in range(current_year_be - 2, current_year_be + 5)]
        self.commission_year_menu = CTkOptionMenu(month_year_frame, variable=self.commission_year_var, values=year_list, width=90, **self.dropdown_style)
        self.commission_year_menu.pack(side="left", padx=5)
        
        # --- Bill Date is now below ---
        bill_date_outer_frame = CTkFrame(frame, fg_color="transparent")
        bill_date_outer_frame.grid(row=2, column=0, columnspan=3, padx=15, pady=4, sticky="ew") # <-- Changed to row=2
        CTkLabel(bill_date_outer_frame, text="วันที่เปิด SO:", font=CTkFont(size=14)).pack(side="left")
        self.bill_date_selector = DateSelector(bill_date_outer_frame, dropdown_style=self.dropdown_style)
        self.bill_date_selector.pack(side="left", padx=5)

        # --- The rest of the function remains the same ---
        customer_type_radio_frame = CTkFrame(frame, fg_color="transparent")
        CTkRadioButton(customer_type_radio_frame, text="ลูกค้าเก่า", variable=self.customer_type_var, value="ลูกค้าเก่า", command=self._toggle_customer_fields).pack(side="left", padx=5)
        CTkRadioButton(customer_type_radio_frame, text="ลูกค้าใหม่", variable=self.customer_type_var, value="ลูกค้าใหม่", command=self._toggle_customer_fields).pack(side="left", padx=20)
        self._add_form_row(frame, "ประเภทลูกค้า:", customer_type_radio_frame, 3)

        self.old_customer_frame = CTkFrame(frame, fg_color="transparent")
        self.old_customer_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=0, pady=0)
        self.old_customer_frame.grid_columnconfigure(1, weight=1)

        # <<< START: แก้ไขการเรียกใช้ AutoCompleteEntry ตรงนี้ >>>
        self.customer_id_entry = AutoCompleteEntry(
            master=self.old_customer_frame, 
            completion_list=self.customer_completion_data,
            display_key='display',  # แก้ไข: เพิ่ม display_key ที่จำเป็น
            command=self._on_customer_id_selected, # แก้ไข: เปลี่ยนชื่อ argument เป็น command
            placeholder_text="ค้นหารหัสหรือชื่อลูกค้า..."
        )
        # <<< END: สิ้นสุดการแก้ไข >>>

        self._add_form_row(self.old_customer_frame, "รหัสลูกค้า:", self.customer_id_entry, 0)

        self.customer_name_entry = CTkEntry(self.old_customer_frame, state="readonly")
        self._add_form_row(self.old_customer_frame, "ชื่อลูกค้า:", self.customer_name_entry, 1)

        self.new_customer_frame = CTkFrame(frame, fg_color="transparent")
        self.new_customer_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=0, pady=0)
        self.new_customer_frame.grid_columnconfigure(1, weight=1)

        self.new_customer_id_entry = CTkEntry(self.new_customer_frame, placeholder_text="กำหนดรหัสลูกค้าใหม่")
        self._add_form_row(self.new_customer_frame, "รหัสลูกค้า (ใหม่):", self.new_customer_id_entry, 0)

        self.new_customer_name_entry = CTkEntry(self.new_customer_frame, placeholder_text="กรอกชื่อลูกค้าใหม่")
        self._add_form_row(self.new_customer_frame, "ชื่อลูกค้า (ใหม่):", self.new_customer_name_entry, 1)

        credit_so_frame = CTkFrame(frame, fg_color="transparent")
        credit_so_frame.grid(row=5, column=0, columnspan=3, padx=15, pady=4, sticky="ew")
        credit_so_frame.grid_columnconfigure((1, 3), weight=1)

        CTkLabel(credit_so_frame, text="Credit:", font=CTkFont(size=14)).grid(row=0, column=0, sticky="w")
        self.credit_term_menu = CTkOptionMenu(credit_so_frame, variable=self.credit_term_var, values=["เงินสด", "CR"], **self.dropdown_style)
        self.credit_term_menu.grid(row=0, column=1, sticky="ew", padx=(10, 20))

        CTkLabel(credit_so_frame, text="เลขที่ใบสั่งขาย:", font=CTkFont(size=14)).grid(row=0, column=2, padx=(20, 10), sticky="w")
        self.so_number_entry = CTkEntry(credit_so_frame, textvariable=self.so_number_var)
        self.so_number_entry.grid(row=0, column=3, sticky="ew")

        self.order_pur_entry = CTkEntry(frame, textvariable=self.order_pur_var, placeholder_text="ระบุ Order Pur (บังคับใส่)")
        self._add_form_row(frame, "Order Pur: *", self.order_pur_entry, 6)

        self._toggle_customer_fields()

        self._toggle_customer_fields()

    def _toggle_customer_fields(self):
        if self.customer_type_var.get() == "ลูกค้าเก่า":
            self.new_customer_frame.grid_remove()
            self.old_customer_frame.grid()
            if hasattr(self, 'customer_id_entry'):
                self.customer_id_entry.delete(0, tk.END)
                self.customer_name_entry.configure(state="normal")
                self.customer_name_entry.delete(0, tk.END)
                self.customer_name_entry.configure(state="readonly")
        else:
            self.old_customer_frame.grid_remove()
            self.new_customer_frame.grid()
            self.credit_term_var.set("เงินสด")

    # <<< START: CODE REPLACEMENT >>>
    def _populate_other_expenses_frame(self, parent):
        # <<< แก้ไข: จัด Layout ของ Frame นี้ใหม่ทั้งหมด >>>
        details_container = CTkFrame(parent, fg_color="transparent")
        details_container.pack(fill="x", expand=True, pady=10)
        details_container.grid_columnconfigure(0, weight=1)
        details_container.grid_columnconfigure(1, weight=1)

        # --- คอลัมน์ซ้าย: ส่วนลด/รายการเพิ่มเติม ---
        discounts_frame = self._create_section_frame(details_container, "ส่วนลด/รายการเพิ่มเติม")
        discounts_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        self.brokerage_fee_entry = NumericEntry(discounts_frame)
        self._add_form_row(discounts_frame, "ค่านายหน้า:", self.brokerage_fee_entry, 1)
        
        self.coupon_value_entry = NumericEntry(discounts_frame)
        self._add_form_row(discounts_frame, "คูปอง:", self.coupon_value_entry, 2)
        
        # 🟢 [แก้ไข] เปลี่ยนเป็นของแถม 2 รูปแบบ
        self.giveaway_vat_entry = CTkEntry(discounts_frame)
        self._add_form_row(discounts_frame, "ของแถมใน SO (Vat):", self.giveaway_vat_entry, 3)
        
        self.giveaway_no_vat_entry = CTkEntry(discounts_frame)
        self._add_form_row(discounts_frame, "ของแถมนอก SO (No Vat):", self.giveaway_no_vat_entry, 4)

        # --- คอลัมน์ขวา: Delivery Note ---
        delivery_frame = self._create_section_frame(details_container, "Delivery Note")
        delivery_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        delivery_frame.grid_columnconfigure(1, weight=1) # ทำให้คอลัมน์ที่ 2 ขยายได้

        delivery_options = [
            "ซัพพลายเออร์จัดส่ง", "Aplus Logistic ส่งหน้างาน", "ลูกค้ารับเองที่ซัพ",
            "ลูกค้ารับเองที่คลัง 132", "ย้ายเข้าคลัง Aplus Logistic รอลูกค้ารับที่คลัง",
            "ย้ายเข้าคลัง Aplus Logistic รอ Aplus Logistic จัดส่ง",
            "ย้ายเข้าคลัง Lalamove รอลูกค้ารับที่คลัง 132", "ส่ง Lalamove ให้ลูกค้าหน้างาน",
            "Aplus Logistic+ฝากส่งขนส่ง", "Lalamove +ฝากส่งขนส่ง"
        ]
        self.delivery_type_menu = CTkOptionMenu(delivery_frame, variable=self.delivery_type_var, values=delivery_options, command=self._on_delivery_type_change, **self.dropdown_style)
        self._add_form_row(delivery_frame, "การจัดส่ง:", self.delivery_type_menu, 1)

        self.pickup_location_entry = CTkEntry(delivery_frame, placeholder_text="ใส่ อำเภอ, จังหวัด หรือ Google map link")
        self._add_form_row(delivery_frame, "Location เข้ารับ:", self.pickup_location_entry, 2)

        # --- ส่วนของ "ค่าย้าย" ที่แก้ไขใหม่ ---
        CTkLabel(delivery_frame, text="ค่าย้าย:", font=CTkFont(size=14)).grid(row=3, column=0, padx=15, pady=5, sticky="w")
        relocation_entry_frame = CTkFrame(delivery_frame, fg_color="transparent")
        relocation_entry_frame.grid(row=3, column=1, padx=(10,15), pady=5, sticky="ew")
        self.relocation_cost_entry = NumericEntry(relocation_entry_frame)
        self.relocation_cost_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        
        relocation_radio_frame = CTkFrame(relocation_entry_frame, fg_color="transparent")
        relocation_radio_frame.pack(side="left")
        CTkRadioButton(relocation_radio_frame, text="VAT", variable=self.relocation_vat_option_var, value="VAT").pack(side="left", padx=5)
        CTkRadioButton(relocation_radio_frame, text="CASH", variable=self.relocation_vat_option_var, value="NO VAT").pack(side="left", padx=5)

        # --- ช่องแสดงผล VAT ของ "ค่าย้าย" ---
        self.relocation_vat_display = CTkEntry(delivery_frame, textvariable=self.relocation_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(delivery_frame, "VAT 7% (ค่าย้าย):", self.relocation_vat_display, 4)

        # --- Widget ที่เหลือ ถูกเลื่อนลำดับแถวลงมา ---
        self.date_to_wh_label = CTkLabel(delivery_frame, text="วันที่ย้ายเข้าคลัง:", font=CTkFont(size=14))
        self.date_to_wh_label.grid(row=5, column=0, padx=15, pady=4, sticky="w")
        self.date_to_wh_selector = DateSelector(delivery_frame, dropdown_style=self.dropdown_style)
        self.date_to_wh_selector.grid(row=5, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")

        self.date_to_customer_selector = DateSelector(delivery_frame, dropdown_style=self.dropdown_style)
        self._add_form_row(delivery_frame, "วันที่จัดส่งลูกค้า:", self.date_to_customer_selector, 6)

        self.pickup_rego_entry = CTkEntry(delivery_frame)
        self._add_form_row(delivery_frame, "ทะเบียนเข้ารับ:", self.pickup_rego_entry, 7)
    # <<< END: CODE REPLACEMENT >>>

    # 🟢 [เพิ่มฟังก์ชันใหม่] วาดฟอร์มรายละเอียดเพิ่มเติมหน้างาน
    def _populate_additional_details_frame(self, parent):
        frame = self._create_section_frame(parent, "รายละเอียดเพิ่มเติม (หน้างาน)")
        frame.pack(fill="x", pady=(0, 10))
        frame.grid_columnconfigure(1, weight=1)

        # 1. แผนที่จัดส่ง
        self.delivery_map_entry = CTkEntry(frame, textvariable=self.delivery_map_var)
        self._add_form_row(frame, "แผนที่จัดส่ง:", self.delivery_map_entry, 1)
        CTkLabel(frame, text="* หากมี - (ระบุ Google Map หรือ อำเภอ จังหวัด)", font=CTkFont(size=11, slant="italic"), text_color="gray50").grid(row=2, column=1, sticky="w", padx=15, pady=(0, 5))

        # 2. ชื่อผู้ติดต่อหน้างาน
        self.onsite_contact_name_entry = CTkEntry(frame, textvariable=self.onsite_contact_name_var)
        self._add_form_row(frame, "ชื่อผู้ติดต่อหน้างาน:", self.onsite_contact_name_entry, 3)

        # 3. เบอร์ติดต่อหน้างาน
        self.onsite_contact_phone_entry = CTkEntry(frame, textvariable=self.onsite_contact_phone_var)
        self._add_form_row(frame, "เบอร์ติดต่อหน้างาน:", self.onsite_contact_phone_entry, 4)

        # 4. ประเภทรถ (Dropdown)
        vehicle_options = [
            "-", "กระบะ", "6 ล้อธรรมดา", "6 ล้อเฮียบ", "10 ล้อธรรมดา", "10 ล้อเฮียบ", 
            "รถเทรลเลอร์", "รถเทรลเลอร์-เฮียบ", "lala มอไซ", "lala เก๋ง", 
            "lala กระบะ", "lala กระบะตู้ทึบ","ลูกค้ารับเอง", "ฝากส่งขนส่งเอกชน-ชำระต้นทาง", "ฝากส่งขนส่งเอกชน-เก็บปลายทาง"
        ]
        self.vehicle_type_menu = CTkOptionMenu(frame, variable=self.vehicle_type_var, values=vehicle_options, **self.dropdown_style)
        
        # 🟢 [แก้ไขตรงนี้] เติม * เพื่อให้เซลส์รู้ว่าต้องเลือก
        self._add_form_row(frame, "ประเภทรถ: *", self.vehicle_type_menu, 5)
        # 🟢 [เพิ่มใหม่] 5. เงื่อนไขการลงสินค้า (Radio Button เลือกได้อย่างเดียว)
        unloading_frame = CTkFrame(frame, fg_color="transparent")
        unloading_frame.grid(row=6, column=1, padx=15, pady=4, sticky="w")
        CTkRadioButton(unloading_frame, text="รวมลง", variable=self.unloading_status_var, value="รวมลง").pack(side="left", padx=(0, 15))
        CTkRadioButton(unloading_frame, text="ไม่รวมลง", variable=self.unloading_status_var, value="ไม่รวมลง").pack(side="left")
        self._add_form_row(frame, "เงื่อนไขลงสินค้า:", unloading_frame, 6)

        # 🟢 [เพิ่มใหม่] 6. Special Request
        self.special_request_entry = CTkEntry(frame, textvariable=self.special_request_var)
        self._add_form_row(frame, "Special Request: *", self.special_request_entry, 7)
        CTkLabel(frame, text="* บังคับใส่ (หากไม่มีให้พิมพ์ '-'หรือ 'ไม่มี')", font=CTkFont(size=11, slant="italic"), text_color="#D32F2F").grid(row=8, column=1, sticky="w", padx=15, pady=(0, 5))
        

    def _populate_sales_services_frame(self, parent):
        frame = self._create_section_frame(parent, "ยอดขายและบริการ")
        frame.pack(fill="x", pady=(0,10))
        frame.grid_columnconfigure(1, weight=1)

        CTkLabel(frame, text="รายการ", font=CTkFont(size=14, weight="bold")).grid(row=1, column=0, padx=15)
        CTkLabel(frame, text="ยอดขาย", font=CTkFont(size=14, weight="bold")).grid(row=1, column=1, padx=10)
        CTkLabel(frame, text="ประเภท", font=CTkFont(size=14, weight="bold")).grid(row=1, column=2, padx=10)

        self.sales_amount_entry = NumericEntry(frame)
        self._add_item_row_with_vat(frame, "ยอดขายสินค้า/บริการ:", self.sales_amount_entry, self.sales_service_vat_option, 2)

        note_font = CTkFont(size=12, slant="italic")
        note_label = CTkLabel(frame, text="หมายเหตุ: ไม่รวมค่าส่ง/ค่ารถ :ยอดหลังจากหักส่วนลดแล้ว", font=note_font, text_color="#FF0000")
        note_label.grid(row=3, column=1, columnspan=2, padx=(10, 15), pady=(0, 5), sticky="w")

        self.sales_vat_var_display = CTkEntry(frame, textvariable=self.sales_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "VAT 7% (สินค้า/บริการ):", self.sales_vat_var_display, 4)

        self.cutting_drilling_fee_entry = NumericEntry(frame)
        self._add_item_row_with_vat(frame, "ค่าบริการตัด/เจาะ:", self.cutting_drilling_fee_entry, self.cutting_drilling_fee_vat_option, 5)

        self.cutting_drilling_vat_var_display = CTkEntry(frame, textvariable=self.cutting_drilling_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "VAT 7% (ตัด/เจาะ):", self.cutting_drilling_vat_var_display, 6)

        self.other_service_fee_entry = NumericEntry(frame)
        self._add_item_row_with_vat(frame, "ค่าบริการอื่นๆ:", self.other_service_fee_entry, self.other_service_fee_vat_option, 7)

        self.other_service_vat_var_display = CTkEntry(frame, textvariable=self.other_service_vat_calc_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "VAT 7% (บริการอื่นๆ):", self.other_service_vat_var_display, 8)

    def _populate_shipping_frame(self, parent):
        frame = self._create_section_frame(parent, "ค่าจัดส่ง"); frame.pack(fill="x", pady=(0,10)); frame.grid_columnconfigure(1, weight=1)
        self.shipping_cost_entry = NumericEntry(frame); self._add_item_row_with_vat(frame, "ค่าจัดส่ง:", self.shipping_cost_entry, self.shipping_vat_option_var, 1)
        self.shipping_vat_var_display = CTkEntry(frame, textvariable=self.shipping_vat_calc_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "VAT 7% (ค่าจัดส่ง):", self.shipping_vat_var_display, 2)
        self.delivery_date_selector = DateSelector(frame, dropdown_style=self.dropdown_style); self._add_form_row(frame, "วันที่จัดส่ง:", self.delivery_date_selector, 3, columnspan=2)

        # Warning label — แสดงเมื่อวันที่จัดส่งเกิน cutoff
        self._delivery_cutoff_label = CTkLabel(
            frame, text="", font=CTkFont(size=12, weight="bold"),
            text_color="#DC2626", wraplength=380, justify="left"
        )
        self._delivery_cutoff_label.grid(row=4, column=1, columnspan=2, padx=(10, 15), pady=(0, 6), sticky="w")

        self._cutoff_popup_shown = False  # ป้องกัน popup ซ้ำ
        # Trace เมื่อวันหรือเดือนเปลี่ยน
        self.delivery_date_selector.day_var.trace_add("write", lambda *_: self._check_delivery_date_cutoff())
        self.delivery_date_selector.month_var.trace_add("write", lambda *_: self._check_delivery_date_cutoff())

    def _check_delivery_date_cutoff(self):
        """เช็ควันที่จัดส่ง — ถ้าเกิน cutoff แสดง warning ให้ user เปลี่ยนรอบคอมเอง"""
        try:
            day_str = self.delivery_date_selector.day_var.get()
            month_str = self.delivery_date_selector.month_var.get()
            year_str = self.delivery_date_selector.year_var.get()
            if not day_str or not month_str or not year_str:
                return

            day = int(day_str)
            thai_month_map = {"ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
                              "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
                              "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12}
            month_num = thai_month_map.get(month_str, 0)
            year_be = int(year_str)

            cutoff = 21 if month_num in (2, 12) else 25

            if day > cutoff:
                next_month_num = month_num + 1 if month_num < 12 else 1
                next_year_be = year_be if month_num < 12 else year_be + 1
                next_month_thai = self.thai_months[next_month_num - 1]
                self._delivery_cutoff_label.configure(
                    text=f"⚠️ วันที่จัดส่ง ({day_str}) เกินวันตัดรอบ ({cutoff})\n"
                         f"กรุณาเปลี่ยนรอบคอมเป็น {next_month_thai} {next_year_be} ด้วยตัวเอง"
                )
            else:
                self._delivery_cutoff_label.configure(text="")
        except Exception:
            pass

    def _is_delivery_date_over_cutoff(self):
        """คืนค่า True ถ้าวันที่จัดส่งเกิน cutoff และ commission month ยังไม่ถูกเปลี่ยนเป็นเดือนหน้าที่ถูกต้อง"""
        try:
            day_str = self.delivery_date_selector.day_var.get()
            month_str = self.delivery_date_selector.month_var.get()
            if not day_str or not month_str:
                return False

            day = int(day_str)
            thai_month_map = {"ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
                              "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
                              "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12}
            month_num = thai_month_map.get(month_str, 0)
            cutoff = 21 if month_num in (2, 12) else 25

            if day <= cutoff:
                return False

            # เดือนที่ commission ต้องเป็น = เดือนถัดจากวันจัดส่ง
            required_comm_month = month_num + 1 if month_num < 12 else 1
            comm_month_num = thai_month_map.get(self.commission_month_var.get(), 0)

            # block ถ้า commission month ยังไม่ตรงกับเดือนที่ควรจะเป็น
            return comm_month_num != required_comm_month
        except Exception:
            return False

    def _populate_fees_frame(self, parent):
        frame = self._create_section_frame(parent, "ค่าธรรมเนียม"); frame.pack(fill="x", pady=(0,10)); frame.grid_columnconfigure(1, weight=1)
        self.credit_card_fee_entry = NumericEntry(frame); self._add_item_row_with_vat(frame, "ค่าธรรมเนียมบัตรเครดิต:", self.credit_card_fee_entry, self.credit_card_fee_vat_option_var, 1)
        self.card_fee_vat_var_display = CTkEntry(frame, textvariable=self.card_fee_vat_calc_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "VAT 7% (ค่าธรรมเนียม):", self.card_fee_vat_var_display, 2)
        self.transfer_fee_entry = NumericEntry(frame, placeholder_text="หากมี"); self._add_form_row(frame, "ค่าธรรมเนียมโอน:", self.transfer_fee_entry, 3)
        self.wht_fee_entry = NumericEntry(frame, placeholder_text="หากมี"); self._add_form_row(frame, "ภาษีหัก ณ ที่จ่าย:", self.wht_fee_entry, 4)


    def _populate_payment_frame(self, parent):
        frame = self._create_section_frame(parent, "รายละเอียดการโอนชำระ")
        frame.pack(fill="x", expand=True, pady=10)
        
        payment_options = ["ชำระสด", "โอน KBANK", "โอน TTB - ออมทรัพย์", "โอน TTB - กระแส", "โอน กรรมการ", "ชำระผ่านบัตรเครดิต"]

        # --- Payment 1 Row ---
        payment1_frame = CTkFrame(frame, fg_color="transparent")
        payment1_frame.grid(row=1, column=1, columnspan=2, padx=(10,15), pady=4, sticky="ew")

        self.payment1_percent_menu = CTkOptionMenu(payment1_frame, variable=self.payment1_percent_var, values=["ระบุยอดเอง", "30%", "50%", "100%"], width=120, **self.dropdown_style)
        self.payment1_percent_menu.pack(side="left", padx=(0, 5))

        self.payment1_amount_entry = NumericEntry(payment1_frame, placeholder_text="ระบุยอดโอนตามสลิป")
        self.payment1_amount_entry.pack(side="left", fill="x", expand=True, padx=(0,5))

        self.payment1_method_menu = CTkOptionMenu(payment1_frame, variable=self.payment1_method_var, values=payment_options, width=160, **self.dropdown_style)
        self.payment1_method_menu.pack(side="left", padx=(0,5))

        self.payment1_date_selector = DateSelector(payment1_frame, dropdown_style=self.dropdown_style)
        self.payment1_date_selector.pack(side="left")

        self._add_form_row(frame, "1. มัดจำ/ชำระเต็ม:", payment1_frame, 1)

        # --- Payment 2 Row ---
        payment2_frame = CTkFrame(frame, fg_color="transparent")
        payment2_frame.grid(row=2, column=1, columnspan=2, padx=(10,15), pady=4, sticky="ew")

        self.payment2_percent_menu = CTkOptionMenu(payment2_frame, variable=self.payment2_percent_var, values=["ระบุยอดเอง", "30%", "50%", "70%", "100%"], width=120, **self.dropdown_style)
        self.payment2_percent_menu.pack(side="left", padx=(0, 5))

        self.payment2_amount_entry = NumericEntry(payment2_frame, placeholder_text="ระบุยอดโอนตามสลิป")
        self.payment2_amount_entry.pack(side="left", fill="x", expand=True, padx=(0,5))

        self.payment2_method_menu = CTkOptionMenu(payment2_frame, variable=self.payment2_method_var, values=payment_options, width=160, **self.dropdown_style)
        self.payment2_method_menu.pack(side="left", padx=(0,5))

        self.payment2_date_selector = DateSelector(payment2_frame, dropdown_style=self.dropdown_style)
        self.payment2_date_selector.pack(side="left")

        self._add_form_row(frame, "2. มัดจำ/ชำระเต็ม:", payment2_frame, 2)

        # --- Other rows ---
        self.payment1_percent_menu.configure(command=self._on_payment1_select)
        self.payment2_percent_menu.configure(command=self._on_payment2_select)

        payment_total_output = CTkEntry(frame, textvariable=self.payment_total_var, state="readonly", fg_color="gray85")
        self._add_form_row(frame, "ยอดโอนชำระรวม VAT:", payment_total_output, 3)

        self.balance_due_entry = CTkEntry(frame, textvariable=self.balance_due_var, state="readonly", fg_color="gray85", font=CTkFont(weight="bold"))
        self._add_form_row(frame, "ค้างชำระ:", self.balance_due_entry, 4)

    def _populate_so_summary_frame(self, parent):
        frame = self._create_section_frame(parent, "SO ยอดขายสินค้า/ค่าจัดส่ง รวมภาษีมูลค่าเพิ่ม"); frame.pack(fill="x", pady=(0, 10))
        frame.grid_columnconfigure(1, weight=1)
        self._add_form_row(frame, "รวมยอดขาย SO:", CTkEntry(frame, textvariable=self.so_subtotal_var, state="readonly", fg_color="gray85"), 1, columnspan=1)
        self._add_form_row(frame, "VAT 7%:", CTkEntry(frame, textvariable=self.so_vat_var, state="readonly", fg_color="gray85"), 2, columnspan=1)
        self._add_form_row(frame, "ยอดที่ต้องชำระ:", CTkEntry(frame, textvariable=self.so_grand_total_var, state="readonly", fg_color="gray85", font=CTkFont(weight="bold")), 3, columnspan=1)
        self.so_vs_payment_result_entry = CTkEntry(frame, textvariable=self.so_vs_payment_result_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "ตรวจสอบยอด SO VS ชำระ:", self.so_vs_payment_result_entry, 4, columnspan=1)
        self.difference_amount_entry = CTkEntry(frame, textvariable=self.difference_amount_var, state="readonly", fg_color="gray85"); self._add_form_row(frame, "ผลต่าง:", self.difference_amount_entry, 5, columnspan=1)

        self.balance_due_entry = CTkEntry(frame, textvariable=self.balance_due_var, state="readonly", fg_color="gray85", font=CTkFont(weight="bold"))
        self._add_form_row(frame, "ค้างชำระ:", self.balance_due_entry, 6, columnspan=1)

    def _populate_cash_verification_frame(self, parent):
        frame = self._create_section_frame(parent, "ตรวจสอบยอดชำระเงินสด CASH"); frame.pack(fill="x", pady=0)
        frame.grid_columnconfigure(1, weight=1)

        CTkLabel(frame, text="ยอดค่าสินค้าเงินสด:", font=CTkFont(size=14)).grid(row=1, column=0, padx=15, pady=5, sticky="w")
        self.cash_product_input_entry = NumericEntry(frame, textvariable=self.cash_product_input_var, placeholder_text="0.00")
        self.cash_product_input_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.cash_product_input_entry.bind("<KeyRelease>", self._update_final_calculations)

        CTkLabel(frame, text="ยอดรวมค่าบริการเงินสด:", font=CTkFont(size=14)).grid(row=2, column=0, padx=15, pady=5, sticky="w")
        self.cash_service_total_entry = CTkEntry(frame, textvariable=self.cash_service_total_var, state="readonly", fg_color="gray85")
        self.cash_service_total_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        CTkLabel(frame, text="ยอดที่ต้องชำระเงินสด:", font=CTkFont(size=14)).grid(row=3, column=0, padx=15, pady=5, sticky="w")
        self.cash_required_total_entry = CTkEntry(frame, textvariable=self.cash_required_total_var, state="readonly", fg_color="gray85")
        self.cash_required_total_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        CTkLabel(frame, text="ยอดชำระจริงเงินสด:", font=CTkFont(size=14)).grid(row=4, column=0, padx=15, pady=5, sticky="w")
        self.cash_actual_payment_entry = NumericEntry(frame, textvariable=self.cash_actual_payment_var, placeholder_text="0.00")
        self.cash_actual_payment_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.cash_actual_payment_entry.bind("<KeyRelease>", self._update_final_calculations)

        CTkLabel(frame, text="ตรวจสอบยอดชำระเงินสด:", font=CTkFont(size=14)).grid(row=5, column=0, padx=15, pady=5, sticky="w")
        self.cash_verification_result_entry = CTkEntry(frame, textvariable=self.cash_verification_result_var, state="readonly", fg_color="gray85")
        self.cash_verification_result_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

    def _populate_action_frame(self, parent):
        frame = CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=(10,0))
        frame.grid_rowconfigure((0,1), weight=1)
        frame.grid_columnconfigure((0,1,2), weight=1)
        
        # +++ START: เพิ่ม self. เพื่อเก็บ Reference ของปุ่ม +++
        self.btn_clear = CTkButton(frame, text="ล้างข้อมูล", fg_color="#F97316", command=self._clear_form)
        self.btn_clear.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        self.btn_edit = CTkButton(frame, text="แก้ไขข้อมูลจากประวัติ", fg_color="#EAB308", command=self._show_history)
        self.btn_edit.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        self.save_button = CTkButton(frame, text="บันทึกข้อมูล", command=self._save_data)
        self.save_button.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")
        
        btn_history = CTkButton(frame, text="แสดงประวัติ", command=self._show_history)
        btn_history.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        
        btn_export = CTkButton(frame, text="นำไฟล์ออก EXCEL", fg_color="#1F2937", command=self._export_history_to_excel)
        btn_export.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        
        btn_submit = CTkButton(frame, text="นำส่งข้อมูล...", fg_color="#16A34A", command=self._open_submit_dialog)
        btn_submit.grid(row=1, column=2, padx=5, pady=5, sticky="nsew")
        # +++ END +++
        
        note_text = "หมายเหตุ **\nนำส่งข้อมูลเข้าระบบคำนวณคอมมิชชั่นแล้วไม่สามารถแก้ไขได้"
        note_label = CTkLabel(frame, text=note_text, font=CTkFont(size=13), text_color="#D32F2F", justify="left")
        note_label.grid(row=2, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

    def _open_submit_dialog(self):
        # ฟังก์ชันใหม่สำหรับเปิดหน้าต่างเลือกส่ง SO
        SubmitSODialog(self, self.app_container, self.sale_key, self.sale_name)

    def _update_final_calculations(self, *args):
        # --- 1. รวบรวมข้อมูล (ส่วนนี้ทำงานถูกต้องอยู่แล้ว) ---
        sales = utils.convert_to_float(self.sales_amount_entry.get())
        shipping = utils.convert_to_float(self.shipping_cost_entry.get())
        card_fee = utils.convert_to_float(self.credit_card_fee_entry.get())
        cutting_drilling = utils.convert_to_float(self.cutting_drilling_fee_entry.get())
        other_service = utils.convert_to_float(self.other_service_fee_entry.get())
        relocation = utils.convert_to_float(self.relocation_cost_entry.get())
        wht = utils.convert_to_float(self.wht_fee_entry.get())

        total_vatable_revenue = 0.0
        total_cashable_services_and_fees = 0.0
        items_to_process = [
            (sales, self.sales_service_vat_option.get(), self.sales_vat_calc_var),
            (cutting_drilling, self.cutting_drilling_fee_vat_option.get(), self.cutting_drilling_vat_calc_var),
            (other_service, self.other_service_fee_vat_option.get(), self.other_service_vat_calc_var),
            (shipping, self.shipping_vat_option_var.get(), self.shipping_vat_calc_var),
            (card_fee, self.credit_card_fee_vat_option_var.get(), self.card_fee_vat_calc_var),
            (relocation, self.relocation_vat_option_var.get(), self.relocation_vat_calc_var)
        ]
        for amount, option, var_display in items_to_process:
            item_vat = 0.0
            if option == "VAT":
                total_vatable_revenue += amount
                item_vat = amount * 0.07
            else:
                total_cashable_services_and_fees += amount
            if var_display:
                var_display.set(f"{item_vat:,.2f}")

        final_amount_due = (total_vatable_revenue * 1.07) - wht
        self.so_grand_total_var.set(f"{final_amount_due:,.2f}")

        payment1 = utils.convert_to_float(self.payment1_amount_entry.get())
        payment2 = utils.convert_to_float(self.payment2_amount_entry.get())
        total_payment = payment1 + payment2
        self.payment_total_var.set(f"{total_payment:,.2f}")

        # --- 2. คำนวณส่วนต่าง (Difference) ---
        # แก้ไขสูตรเป็น: ยอดโอน - ยอดที่ต้องชำระ
        # - ถ้าเป็นบวก (+) แปลว่า "โอนเกิน"
        # - ถ้าเป็นลบ (-) แปลว่า "โอนขาด"
        difference = total_payment - final_amount_due
        self.difference_amount_var.set(f"{difference:,.2f}")
        
        # ยอดค้างชำระ (Balance Due) จะเป็นค่าบวกเสมอ
        self.balance_due_var.set(f"{abs(difference):,.2f}")

        def set_check_result(entry, var, diff_val, plus_text, minus_text):
            if not entry or not entry.winfo_exists(): return
            color_map = {"ok": ("#BBF7D0", "#15803D"), "bad": ("#FECACA", "#B91C1C")}
            if abs(diff_val) < 0.01:
                state, text = "ok", "ถูกต้อง"
                entry.configure(fg_color=color_map["ok"][0], text_color=color_map["ok"][1])
            elif diff_val > 0: # โอนเกิน (ค่าเป็นบวก)
                state, text = "ok", f"{plus_text} (+{abs(diff_val):,.2f})"
                entry.configure(fg_color=color_map["ok"][0], text_color=color_map["ok"][1])
            else: # โอนขาด (ค่าเป็นลบ)
                state, text = "bad", f"{minus_text} ({abs(diff_val):,.2f})"
                entry.configure(fg_color=color_map["bad"][0], text_color=color_map["bad"][1])
            var.set(text)

        # --- 3. เรียกใช้ฟังก์ชันแสดงผลด้วย Logic ที่ถูกต้อง ---
        # เมื่อ difference > 0 (บวก) ให้แสดง plus_text ("ยอดโอนเกิน")
        # เมื่อ difference < 0 (ลบ) ให้แสดง minus_text ("ยอดโอนขาด")
        set_check_result(self.so_vs_payment_result_entry, self.so_vs_payment_result_var, difference,
                         plus_text="ยอดโอนเกิน",
                         minus_text="ยอดโอนขาด")

        # --- 4. คำนวณส่วนของเงินสด (Cash) ---
        cash_product_val = utils.convert_to_float(self.cash_product_input_entry.get())
        cash_required_total = cash_product_val + total_cashable_services_and_fees
        self.cash_required_total_var.set(f"{cash_required_total:,.2f}")

        actual_cash_payment = utils.convert_to_float(self.cash_actual_payment_entry.get())
        cash_difference = actual_cash_payment - cash_required_total
        
        set_check_result(self.cash_verification_result_entry, self.cash_verification_result_var, cash_difference,
                         plus_text="เงินสดเกิน",
                         minus_text="เงินสดขาด")

        cash_product_val = utils.convert_to_float(self.cash_product_input_entry.get())
        cash_required_total = cash_product_val + total_cashable_services_and_fees
        self.cash_required_total_var.set(f"{cash_required_total:,.2f}")
        actual_cash_payment = utils.convert_to_float(self.cash_actual_payment_entry.get())
        cash_diff = cash_required_total - actual_cash_payment
        set_check_result(self.cash_verification_result_entry, self.cash_verification_result_var, cash_diff, "เงินสดขาด", "เงินสดเกิน")

    def _gather_data_from_form(self):
        """รวบรวมข้อมูลจากหน้าจอ (แก้ไขชื่อตัวแปรที่สะกดผิด)"""
        is_new_customer = self.customer_type_var.get() == "ลูกค้าใหม่"
        customer_id = self.new_customer_id_entry.get().strip() if is_new_customer else self.customer_id_entry.get().strip()
        customer_name = self.new_customer_name_entry.get().strip() if is_new_customer else self.customer_name_entry.get().strip()

        p1_date = self.payment1_date_selector.get_date()
        p2_date = self.payment2_date_selector.get_date()
        main_payment_date = max(p1_date, p2_date) if p1_date and p2_date else (p1_date or p2_date)

        data = {
            "bill_date": self.bill_date_selector.get_date(),
            "commission_month": self.thai_month_map.get(self.commission_month_var.get()),
            "commission_year": int(self.commission_year_var.get()) - 543 if self.commission_year_var.get().isdigit() else None,
            "customer_type": self.customer_type_var.get(),
            "customer_name": customer_name,
            "customer_id": customer_id,
            "credit_term": self.credit_term_var.get(),
            "so_number": self.so_number_var.get().strip(),
            "sales_service_amount": utils.convert_to_float(self.sales_amount_entry.get()),
            "sales_service_vat_option": self.sales_service_vat_option.get(),
            "cutting_drilling_fee": utils.convert_to_float(self.cutting_drilling_fee_entry.get()),
            "cutting_drilling_fee_vat_option": self.cutting_drilling_fee_vat_option.get(),
            "other_service_fee": utils.convert_to_float(self.other_service_fee_entry.get()),
            "other_service_fee_vat_option": self.other_service_fee_vat_option.get(),
            "shipping_cost": utils.convert_to_float(self.shipping_cost_entry.get()),
            "shipping_vat_option": self.shipping_vat_option_var.get(),
            "delivery_date": self.delivery_date_selector.get_date(),
            "credit_card_fee": utils.convert_to_float(self.credit_card_fee_entry.get()),
            
            # ✅ แก้ไขจุดนี้: เติม _var ให้ถูกต้องตามที่ประกาศไว้
            "credit_card_fee_vat_option": self.credit_card_fee_vat_option_var.get(), 
            
            "transfer_fee": utils.convert_to_float(self.transfer_fee_entry.get()),
            "wht_3_percent": utils.convert_to_float(self.wht_fee_entry.get()),
            "brokerage_fee": utils.convert_to_float(self.brokerage_fee_entry.get()),
            "coupons": utils.convert_to_float(self.coupon_value_entry.get()),
            "giveaway_vat": self.giveaway_vat_entry.get().strip(),
            "giveaway_no_vat": self.giveaway_no_vat_entry.get().strip(),
            "relocation_cost_vat_option": self.relocation_vat_option_var.get(),
            "delivery_type": self.delivery_type_var.get(),
            "pickup_location": self.pickup_location_entry.get().strip(),
            "relocation_cost": utils.convert_to_float(self.relocation_cost_entry.get()),
            "date_to_warehouse": self.date_to_wh_selector.get_date(),
            "date_to_customer": self.date_to_customer_selector.get_date(),
            "pickup_registration": self.pickup_rego_entry.get().strip(),
            "payment1_amount": utils.convert_to_float(self.payment1_amount_entry.get()), 
            "payment2_amount": utils.convert_to_float(self.payment2_amount_entry.get()),
            "total_payment_amount": utils.convert_to_float(self.payment_total_var.get()),
            "payment_date": main_payment_date,
            "payment1_date": p1_date,
            "payment2_date": p2_date,
            "payment1_method": self.payment1_method_var.get(),
            "payment2_method": self.payment2_method_var.get(),
            "cash_product_input": utils.convert_to_float(self.cash_product_input_var.get()),
            "cash_service_total": utils.convert_to_float(self.cash_service_total_var.get()),
            "cash_required_total": utils.convert_to_float(self.cash_required_total_var.get()),
            "cash_actual_payment": utils.convert_to_float(self.cash_actual_payment_var.get()),
            "difference_amount": utils.convert_to_float(self.difference_amount_var.get()),
            "sale_key": self.sale_key,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "is_active": 1,
            "status": "Draft",
            "delivery_map": self.delivery_map_var.get().strip(),
            "onsite_contact_name": self.onsite_contact_name_var.get().strip(),
            "onsite_contact_phone": self.onsite_contact_phone_var.get().strip(),
            "vehicle_type": self.vehicle_type_var.get(),
            "order_pur": self.order_pur_var.get().strip(),
            "special_request": self.special_request_var.get().strip(),   # 🟢 เพิ่มบรรทัดนี้
            "unloading_status": self.unloading_status_var.get(),
            
        }
        return data

    def _validate_form(self, data):
        if not data["so_number"] or data["so_number"] == "SO":
            return False, "กรุณากรอก 'เลขที่ใบสั่งขาย (SO)'"

        if self._is_delivery_date_over_cutoff():
            thai_month_map = {"ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
                              "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
                              "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12}
            month_str = self.delivery_date_selector.month_var.get()
            month_num = thai_month_map.get(month_str, 0)
            cutoff = 21 if month_num in (2, 12) else 25
            next_month = self.thai_months[month_num % 12]
            return False, (f"วันที่จัดส่งเกินวันตัดรอบ ({cutoff})\n"
                           f"กรุณาเปลี่ยนรอบคอมมิชชั่นเป็น '{next_month}' ก่อนบันทึก")

        if not data.get("order_pur"):
            return False, "กรุณากรอกข้อมูลในช่อง 'Order Pur' (ในส่วนรายละเอียดการขาย) ก่อนทำการบันทึก"
        
        if not data.get("special_request"):
            return False, "กรุณากรอก 'Special Request' ในส่วนรายละเอียดเพิ่มเติม (หากไม่มีให้พิมพ์ '-'หรือ 'ไม่มี')"
        
        if not data.get("vehicle_type") or data.get("vehicle_type") == "-":
            return False, "กรุณาเลือก 'ประเภทรถ' ในส่วนรายละเอียดเพิ่มเติม (หน้างาน) ให้ถูกต้อง"

        if data['customer_type'] == "ลูกค้าใหม่":
            if not data["customer_name"] or not data["customer_id"]:
                return False, "สำหรับ 'ลูกค้าใหม่' กรุณากรอก 'ชื่อ' และ 'รหัส' ให้ครบถ้วน"
        else:  # ลูกค้าเก่า
            if not data["customer_id"]:
                return False, "กรุณาเลือก 'รหัสลูกค้า' สำหรับลูกค้าเก่า"

        # <<< START: แก้ไข Logic การตรวจสอบ SO ซ้ำทั้งหมด >>>
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # สร้าง Query และ Parameters พื้นฐาน
                query = "SELECT id FROM commissions WHERE so_number = %s AND is_active = 1"
                params = [data["so_number"]]

                # ถ้าอยู่ในโหมดแก้ไข (self.editing_record_id มีค่า)
                # ให้เพิ่มเงื่อนไขว่า "ไม่ต้องตรวจสอบกับ ID ของตัวเอง"
                if self.editing_record_id:
                    query += " AND id != %s"
                    params.append(self.editing_record_id)

                # ทำการ Query ด้วยเงื่อนไขที่ถูกต้อง
                cursor.execute(query, tuple(params))
                if cursor.fetchone():
                    return False, f"เลขที่ SO '{data['so_number']}' นี้มีอยู่ในระบบแล้ว"
        except Exception as e:
            # กรณีเกิดข้อผิดพลาดในการเชื่อมต่อฐานข้อมูล
            return False, f"เกิดข้อผิดพลาดในการตรวจสอบข้อมูล: {e}"
        finally:
            if conn:
                self.app_container.release_connection(conn)
        # <<< END: สิ้นสุดการแก้ไข >>>

        return True, ""

    def _handle_new_customer(self, data):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("SELECT id FROM customers WHERE customer_code = %s", (data['customer_id'],))
                if cursor.fetchone():
                    raise ValueError(f"รหัสลูกค้า '{data['customer_id']}' นี้มีอยู่แล้วในระบบ\nกรุณาใช้รหัสอื่น หรือเลือกจากเมนูลูกค้าเก่า")

                insert_query = "INSERT INTO customers (customer_code, customer_name, credit_term) VALUES (%s, %s, %s)"
                cursor.execute(insert_query, (data['customer_id'], data['customer_name'], data['credit_term']))
            conn.commit()
            print(f"Added new customer: {data['customer_name']}")
        except Exception as e:
            if conn: conn.rollback()
            raise e
        finally:
            if conn: self.app_container.release_connection(conn)

    def _perform_db_insert(self, data):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'commissions'")
                db_columns = {row[0] for row in cursor.fetchall()}

                filtered_data = {k: v for k, v in data.items() if k in db_columns}

                filtered_data.pop('id', None)

                cols = ', '.join([f'"{k}"' for k in filtered_data.keys()])
                placeholders = ', '.join(['%s'] * len(filtered_data))

                sql = f"INSERT INTO commissions ({cols}) VALUES ({placeholders})"
                cursor.execute(sql, list(filtered_data.values()))
            conn.commit()
        except Exception as e:
            if conn: conn.rollback()
            raise e
        finally:
            if conn: self.app_container.release_connection(conn)

    def _refresh_history_if_open(self):
        if self.history_window and self.history_window.winfo_exists():
            if hasattr(self.history_window, '_populate_history_table'):
                self.history_window._populate_history_table()

    def _validate_so_number(self, *args):
        current_value = self.so_number_var.get()
        new_value = current_value.upper()
        if new_value != current_value:
            self.so_number_var.set(new_value)

    def _auto_update_commission_period(self, *args):
        """
        อัปเดตเดือนและปีของรอบคอมมิชชั่นอัตโนมัติ ตามวันที่เปิดบิล (Bill Date)
        เพื่อให้ Sale รู้ทันทีว่า SO นี้จะถูกคิดในรอบไหน
        """
        try:
            # ดึงวันที่จาก Bill Date Selector
            selected_date = self.bill_date_selector.get_date()
            
            if selected_date:
                # แปลงเป็นเดือนไทย และปี พ.ศ.
                month_idx = selected_date.month - 1
                year_be = selected_date.year + 543
                
                new_month_str = self.thai_months[month_idx]
                new_year_str = str(year_be)
                
                # อัปเดตตัวแปรของ Dropdown
                if self.commission_month_var.get() != new_month_str:
                    self.commission_month_var.set(new_month_str)
                    
                if self.commission_year_var.get() != new_year_str:
                    self.commission_year_var.set(new_year_str)
                    
        except Exception as e:
            print(f"Error auto-updating commission period: {e}")

    def _create_header(self):
        self.header_frame = CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        # --- START: แก้ไขส่วนนี้ทั้งหมด ---
        # กำหนดให้คอลัมน์ซ้าย (ชื่อ) ขยายตัว และคอลัมน์ขวา (ปุ่ม) ไม่ขยาย
        self.header_frame.grid_columnconfigure(0, weight=1)
        
        CTkLabel(self.header_frame, text=f"ฝ่ายขาย: {self.sale_name} ({self.sale_key})", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).grid(row=0, column=0, sticky="w")
        
        # สร้าง Frame สำหรับปุ่มต่างๆ ทางขวา และใช้ grid วาง
        button_container = CTkFrame(self.header_frame, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="e")
        
        # ใช้ grid วางปุ่มภายใน button_container
        CTkButton(button_container, text="📊 Export SO",
                  command=lambda: SOSummaryExportDialog(self, self.app_container, self.sale_key, self.sale_name),
                  fg_color="#2563EB", hover_color="#1D4ED8"
                  ).grid(row=0, column=0, padx=(0, 6))

        self.tasks_button = CTkButton(button_container, text="งานของฉัน 🔔 (0)", command=self._open_my_tasks_window)
        self.tasks_button.grid(row=0, column=1, padx=6)

        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").grid(row=0, column=2, padx=(0, 10))

    def _on_payment1_select(self, selected_value: str):
        self._calculate_payment_from_percentage(self.payment1_percent_var, self.payment1_amount_entry)
    def _on_payment2_select(self, selected_value: str):
        self._calculate_payment_from_percentage(self.payment2_percent_var, self.payment2_amount_entry)

    def _calculate_payment_from_percentage(self, percent_var, amount_entry):
        try:
            selected_option = percent_var.get()
            if selected_option == "ระบุยอดเอง": return

            grand_total = utils.convert_to_float(self.so_grand_total_var.get())
            if grand_total <= 0:
                amount_entry.delete(0, tk.END)
                self._update_final_calculations()
                return

            percent = float(selected_option.replace('%', '')) / 100.0
            calculated_amount = grand_total * percent

            amount_entry.delete(0, tk.END)
            amount_entry.insert(0, f"{calculated_amount:,.2f}")

            if selected_option == "100%":
                if amount_entry == self.payment1_amount_entry:
                    self.payment2_percent_var.set("ระบุยอดเอง")
                    self.payment2_amount_entry.delete(0, tk.END)
                else:
                    self.payment1_percent_var.set("ระบุยอดเอง")
                    self.payment1_amount_entry.delete(0, tk.END)

            self._update_final_calculations()
        except (ValueError, TypeError) as e:
            print(f"Error calculating payment from percentage: {e}")
            self._update_final_calculations()

    def _clear_form(self, confirm=True):
        if confirm and not messagebox.askyesno("ยืนยัน", "คุณต้องการล้างข้อมูลทั้งหมดในฟอร์มใช่หรือไม่?", parent=self):
            return

        widget_attributes = [
            "so_number_entry", "customer_id_entry", "customer_name_entry",
            "new_customer_id_entry", "new_customer_name_entry", "sales_amount_entry",
            "cutting_drilling_fee_entry", "other_service_fee_entry", "shipping_cost_entry",
            "credit_card_fee_entry", "transfer_fee_entry", "wht_fee_entry", "brokerage_fee_entry",
            "coupon_value_entry", "giveaway_vat_entry", "giveaway_no_vat_entry", "payment1_amount_entry",
            "payment2_amount_entry", "cash_product_input_entry", "cash_actual_payment_entry",
            "pickup_location_entry", "relocation_cost_entry", "pickup_rego_entry"
        ]

        for attr_name in widget_attributes:
            if hasattr(self, attr_name):
                widget = getattr(self, attr_name)
                if isinstance(widget, (CTkEntry, NumericEntry, AutoCompleteEntry)):
                    if widget.cget("state") == "readonly":
                        widget.configure(state="normal")
                        widget.delete(0, tk.END)
                        widget.configure(state="readonly")
                    else:
                        widget.delete(0, tk.END)

        today = datetime.now()
        for selector in [self.bill_date_selector, self.delivery_date_selector, 
                         self.payment1_date_selector, self.payment2_date_selector,
                         self.date_to_wh_selector, self.date_to_customer_selector]:
            selector.set_date(today)

        self.so_number_var.set("SO")
        self.customer_type_var.set("ลูกค้าเก่า")
        self.credit_term_var.set("เงินสด")

        self.commission_month_var.set(self.thai_months[today.month - 1])
        self.commission_year_var.set(str(today.year + 543))
        self.payment1_percent_var.set("ระบุยอดเอง")
        self.payment2_percent_var.set("ระบุยอดเอง")

        self.sales_service_vat_option.set("VAT")
        self.cutting_drilling_fee_vat_option.set("VAT")
        self.other_service_fee_vat_option.set("VAT")
        self.shipping_vat_option_var.set("VAT")
        self.credit_card_fee_vat_option_var.set("VAT")
        self.relocation_vat_option_var.set("VAT")

        self.payment1_method_var.set("ไม่เลือก")
        self.payment2_method_var.set("ไม่เลือก")
        self.delivery_type_var.set("ซัพพลายเออร์จัดส่ง")
        self.delivery_map_var.set("")
        self.onsite_contact_name_var.set("")
        self.onsite_contact_phone_var.set("")
        self.vehicle_type_var.set("-")
        self.order_pur_var.set("")
        self.special_request_var.set("")           # 🟢 เพิ่มบรรทัดนี้
        self.unloading_status_var.set("ไม่รวมลง")

        self.editing_record_id = None
        self._toggle_customer_fields()
        self._update_final_calculations()

        if confirm: messagebox.showinfo("สำเร็จ", "ข้อมูลถูกล้างเรียบร้อยแล้ว", parent=self)


    def _export_history_to_excel(self):
        # 1. เปิด DateRangeDialog เพื่อให้ผู้ใช้เลือกช่วงเวลา
        dialog = DateRangeDialog(self)
        self.master.wait_window(dialog)

        start_date_raw = dialog.start_date # รับค่าวันที่แบบ 'YYYY-MM-DD' หรือ None
        end_date_raw = dialog.end_date     # รับค่าวันที่แบบ 'YYYY-MM-DD' หรือ None

        # ถ้าผู้ใช้กดยกเลิกใน dialog
        if not start_date_raw or not end_date_raw:
            messagebox.showinfo("ยกเลิก", "การ Export ถูกยกเลิก", parent=self)
            return

        try:
            # สร้าง datetime objects โดยกำหนดเวลาให้เป็น 00:00:00 สำหรับวันเริ่มต้น
            # และ 23:59:59 สำหรับวันสิ้นสุด เพื่อให้ครอบคลุมทั้งวัน
            start_datetime = datetime.strptime(start_date_raw, '%Y-%m-%d')
            end_datetime = datetime.strptime(end_date_raw, '%Y-%m-%d') + timedelta(hours=23, minutes=59, seconds=59)

        except ValueError as e:
            messagebox.showerror("ข้อผิดพลาดวันที่", f"รูปแบบวันที่ไม่ถูกต้อง: {e}", parent=self)
            return

        if not messagebox.askyesno("ยืนยัน", "คุณต้องการดึงข้อมูลและ Export เป็นไฟล์ Excel หรือไม่?", parent=self):
            return

        try:
            # แก้ไข query ให้รวมเงื่อนไขของ Sale Key และช่วงเวลา
            # commissions.timestamp is TEXT, so direct BETWEEN might not work as expected with TIMESTAMP objects.
            # It's better to cast timestamp to TIMESTAMP or DATE in SQL for comparison with TIMESTAMP objects.
            # Assuming commissions.timestamp is stored as TEXT in 'YYYY-MM-DD HH:MM:SS' format.
            # If it's a real TIMESTAMP WITH TIME ZONE in DB, psycopg2 will handle the datetime objects.
            # If it's TEXT, we pass strings in the format expected by the DB.

            query = """
                SELECT * FROM commissions
                WHERE sale_key = %s AND is_active = 1
                AND timestamp::timestamp BETWEEN %s::timestamp AND %s::timestamp
                ORDER BY timestamp DESC
            """
            # ใช้ params เพื่อป้องกัน SQL Injection
            params = (self.sale_key, start_datetime.strftime('%Y-%m-%d %H:%M:%S'), end_datetime.strftime('%Y-%m-%d %H:%M:%S'))
            df = pd.read_sql_query(query, self.pg_engine, params=params)

            if df.empty:
                messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลสำหรับ Export ในช่วงเวลาที่เลือก", parent=self)
                return

            df_to_export = df.copy()
            df_to_export.rename(columns=self.header_map, inplace=True)

            default_filename = f"commission_history_{self.sale_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="บันทึกไฟล์ประวัติ Commission",
                initialfile=default_filename,
                parent=self
            )

            if save_path:
                df_to_export.to_excel(save_path, index=False)
                messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้วที่:\n{save_path}", parent=self)
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถ Export ไฟล์ได้: {e}", parent=self)


# ============================================================
# ✏️ SO Edit Tab — Sale สามารถแก้ไขบางฟิลด์ของ SO ที่ยื่นแล้ว
# ============================================================

class SOEditTabView(CTkFrame):
    """แท็บแสดงรายการ SO ของ Sale พร้อมปุ่มแก้ไข + SearchBar + Pagination"""

    PAGE_SIZE = 10

    THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                   "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

    def __init__(self, master, app_container, sale_key, sale_name, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app_container = app_container
        self.sale_key = sale_key
        self.sale_name = sale_name

        self._all_rows = []        # ข้อมูลทั้งหมดจาก DB
        self._filtered_rows = []   # หลัง filter ด้วย search
        self._current_page = 0     # 0-indexed
        self._search_after_id = None   # debounce timer

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)   # row 2 = list (ขยายได้)

        # ── Row 0: Title + hint ──────────────────────────────────
        title_bar = CTkFrame(self, fg_color="transparent")
        title_bar.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 4))

        CTkLabel(title_bar, text="✏️ แก้ไข SO ที่ยื่นแล้ว",
                 font=CTkFont(size=16, weight="bold")).pack(side="left")
        CTkLabel(title_bar,
                 text="วันที่จัดส่ง / ค่าจัดส่ง → บันทึกทันที  |  รอบเดือนค่าคอม → รอ SM อนุมัติ",
                 font=CTkFont(size=12), text_color="gray").pack(side="right")

        # ── Row 1: Search bar + Refresh ─────────────────────────
        search_bar = CTkFrame(self, fg_color=("gray92", "gray18"), corner_radius=8)
        search_bar.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))
        search_bar.grid_columnconfigure(1, weight=1)

        CTkLabel(search_bar, text="🔎", font=CTkFont(size=16)).grid(
            row=0, column=0, padx=(12, 4), pady=8)

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._on_search_change)
        self._search_entry = CTkEntry(
            search_bar, textvariable=self._search_var,
            placeholder_text="ค้นหาเลข SO หรือชื่อลูกค้า...",
            height=34, font=CTkFont(size=13))
        self._search_entry.grid(row=0, column=1, padx=(0, 8), pady=8, sticky="ew")

        CTkButton(search_bar, text="✕ ล้าง", width=70, height=34,
                  fg_color="transparent", border_width=1,
                  text_color=("gray20", "gray80"),
                  command=self._clear_search).grid(row=0, column=2, padx=(0, 6), pady=8)

        CTkButton(search_bar, text="⟳ รีเฟรช", width=90, height=34,
                  fg_color="gray", hover_color="#555555",
                  command=self.load_list).grid(row=0, column=3, padx=(0, 12), pady=8)

        # ── Row 2: Scrollable card list ──────────────────────────
        self.list_frame = CTkScrollableFrame(self, fg_color="white", corner_radius=8)
        self.list_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=0)

        # ── Row 3: Pagination bar ────────────────────────────────
        pag_bar = CTkFrame(self, fg_color="transparent")
        pag_bar.grid(row=3, column=0, sticky="ew", padx=10, pady=(4, 10))
        pag_bar.grid_columnconfigure(1, weight=1)

        self._prev_btn = CTkButton(pag_bar, text="◀ ก่อนหน้า", width=110, height=30,
                                   fg_color="#3B82F6", hover_color="#2563EB",
                                   command=self._prev_page)
        self._prev_btn.grid(row=0, column=0, padx=(0, 8))

        self._page_label = CTkLabel(pag_bar, text="", font=CTkFont(size=13))
        self._page_label.grid(row=0, column=1)

        self._next_btn = CTkButton(pag_bar, text="ถัดไป ▶", width=110, height=30,
                                   fg_color="#3B82F6", hover_color="#2563EB",
                                   command=self._next_page)
        self._next_btn.grid(row=0, column=2, padx=(8, 0))

        self.after(200, self.load_list)

    # ── Data loading ──────────────────────────────────────────────
    def load_list(self):
        """โหลดข้อมูลจาก DB ทั้งหมด แล้วแสดงหน้าแรก"""
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
                cur.execute("""
                    SELECT c.id, c.so_number, c.customer_name, c.bill_date,
                           c.delivery_date, c.shipping_cost,
                           c.commission_month, c.commission_year, c.status,
                           (SELECT ser.status FROM so_edit_requests ser
                            WHERE ser.commission_id = c.id AND ser.status = 'pending'
                            ORDER BY ser.id DESC LIMIT 1) AS pending_edit
                    FROM commissions c
                    WHERE c.sale_key = %s AND c.is_active = 1
                      AND c.status NOT IN ('Cancelled')
                    ORDER BY c.bill_date DESC, c.id DESC
                """, (self.sale_key,))
                self._all_rows = [dict(r) for r in cur.fetchall()]
            self.app_container.release_connection(conn)
        except Exception as e:
            if conn:
                try:
                    self.app_container.release_connection(conn)
                except Exception:
                    pass
            for w in self.list_frame.winfo_children():
                w.destroy()
            CTkLabel(self.list_frame, text=f"❌ โหลดข้อมูลไม่ได้: {e}",
                     text_color="red", font=CTkFont(size=13)).pack(pady=30)
            return

        # Reset search + page แล้วแสดงผล
        self._search_var.set("")
        self._current_page = 0
        self._apply_filter_and_render()

    # ── Search (debounced — รอ 250 ms หลังหยุดพิมพ์) ─────────────
    def _on_search_change(self, *_):
        # ยกเลิก timer เดิมถ้ายังค้างอยู่
        if self._search_after_id is not None:
            self.after_cancel(self._search_after_id)
        self._search_after_id = self.after(250, self._do_search)

    def _do_search(self):
        self._search_after_id = None
        self._current_page = 0
        self._apply_filter_and_render()

    def _clear_search(self):
        # ยกเลิก debounce ที่ค้างอยู่ก่อนเคลียร์
        if self._search_after_id is not None:
            self.after_cancel(self._search_after_id)
            self._search_after_id = None
        self._search_var.set("")
        self._current_page = 0
        self._apply_filter_and_render()
        self._search_entry.focus()

    # ── Filter + Render ───────────────────────────────────────────
    def _apply_filter_and_render(self):
        kw = self._search_var.get().strip().lower()
        if kw:
            self._filtered_rows = [
                r for r in self._all_rows
                if kw in (r.get("so_number") or "").lower()
                or kw in (r.get("customer_name") or "").lower()
            ]
        else:
            self._filtered_rows = list(self._all_rows)

        self._render_page()

    def _render_page(self):
        for w in self.list_frame.winfo_children():
            w.destroy()

        total = len(self._filtered_rows)
        total_pages = max(1, -(-total // self.PAGE_SIZE))  # ceil division

        # Clamp page
        self._current_page = max(0, min(self._current_page, total_pages - 1))

        if total == 0:
            msg = "ไม่พบ SO ที่ตรงกับคำค้นหา" if self._search_var.get() else "ยังไม่มี SO ในระบบ"
            CTkLabel(self.list_frame, text=msg,
                     font=CTkFont(size=14), text_color="gray").pack(pady=40)
            self._page_label.configure(text="")
            self._prev_btn.configure(state="disabled")
            self._next_btn.configure(state="disabled")
            return

        start = self._current_page * self.PAGE_SIZE
        end = min(start + self.PAGE_SIZE, total)
        page_rows = self._filtered_rows[start:end]

        for row in page_rows:
            self._render_card(row)

        # Update pagination controls
        self._page_label.configure(
            text=f"หน้า {self._current_page + 1} / {total_pages}   (แสดง {start+1}–{end} จาก {total} รายการ)")
        self._prev_btn.configure(state="normal" if self._current_page > 0 else "disabled")
        self._next_btn.configure(state="normal" if self._current_page < total_pages - 1 else "disabled")

    # ── Pagination controls ───────────────────────────────────────
    def _prev_page(self):
        if self._current_page > 0:
            self._current_page -= 1
            self._render_page()

    def _next_page(self):
        total_pages = max(1, -(-len(self._filtered_rows) // self.PAGE_SIZE))
        if self._current_page < total_pages - 1:
            self._current_page += 1
            self._render_page()

    # ── Card renderer ─────────────────────────────────────────────
    def _render_card(self, row):
        has_pending = row.get("pending_edit") == "pending"

        bg = "#FEF9C3" if has_pending else "#F0F9FF"
        border = "#EAB308" if has_pending else "#BAE6FD"

        card = CTkFrame(self.list_frame, fg_color=bg,
                        border_width=1, border_color=border, corner_radius=8)
        card.pack(fill="x", padx=10, pady=4)
        card.grid_columnconfigure(0, weight=1)

        info = CTkFrame(card, fg_color="transparent")
        info.grid(row=0, column=0, sticky="ew", padx=15, pady=8)

        # Line 1 — SO + customer
        CTkLabel(info,
                 text=f"SO: {row['so_number']}  |  ลูกค้า: {row['customer_name']}",
                 font=CTkFont(size=14, weight="bold")).pack(anchor="w")

        # Line 2 — delivery date + shipping cost
        dd = row.get("delivery_date") or "-"
        sc_raw = row.get("shipping_cost")
        sc = f"{float(sc_raw):,.2f}" if sc_raw is not None else "-"
        CTkLabel(info,
                 text=f"📅 วันที่จัดส่ง: {dd}   🚚 ค่าจัดส่ง: {sc} บาท",
                 font=CTkFont(size=13), text_color="#0369A1").pack(anchor="w", pady=(2, 0))

        # Line 3 — commission period + status
        m, y = row.get("commission_month"), row.get("commission_year")
        try:
            month_str = self.THAI_MONTHS[int(m) - 1] if m else "-"
            year_str = str(int(y) + 543) if y else ""
        except Exception:
            month_str, year_str = str(m), str(y)

        raw_status = row.get("status") or ""
        if raw_status in ("Paid", "HR Verified"):
            status_display = "✅ จ่ายค่าคอมแล้ว"
            status_color   = "#16A34A"
        else:
            status_display = "🕐 ยังไม่จ่ายค่าคอม"
            status_color   = "#D97706"

        pending_badge = "  ⏳ รอ SM อนุมัติรอบเดือน" if has_pending else ""
        CTkLabel(info,
                 text=f"📆 รอบค่าคอม: {month_str} {year_str}   {status_display}{pending_badge}",
                 font=CTkFont(size=13),
                 text_color="#7C3AED" if has_pending else status_color).pack(anchor="w")

        # Edit button
        btn_frame = CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=0, column=1, padx=15, pady=8)
        CTkButton(btn_frame, text="✏️ แก้ไข",
                  fg_color="#3B82F6", hover_color="#2563EB",
                  width=90, height=32,
                  command=lambda r=row: self._open_edit_dialog(r)).pack()

    # ── Open edit dialog ──────────────────────────────────────────
    def _open_edit_dialog(self, row_dict):
        dlg = SOEditDialog(self, self.app_container,
                           row_dict, self.sale_key, self.sale_name)
        self.wait_window(dlg)
        self.load_list()   # reload + reset to page 1 after save


# ============================================================
# ✏️ SO Edit Dialog — กล่องแก้ไข 3 ฟิลด์
# ============================================================

class SOEditDialog(CTkToplevel):
    """กล่องแก้ไข SO: วันที่จัดส่ง / ค่าจัดส่ง (บันทึกทันที)
       และ รอบเดือนค่าคอม (ส่งขออนุมัติ SM)"""

    THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                   "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

    def __init__(self, master, app_container, row, sale_key, sale_name):
        super().__init__(master)
        self.app_container = app_container
        self.row = row
        self.sale_key = sale_key
        self.sale_name = sale_name

        self.title(f"แก้ไข SO: {row.get('so_number', '')}")
        self.geometry("520x460")
        self.resizable(False, False)
        self.grab_set()
        self.focus()
        self._center()

        self._build_ui()

    # ──────────────────────────────────────────────────────────────
    def _center(self):
        self.update_idletasks()
        w, h = 520, 460
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 20, "pady": 8}
        font_lbl = CTkFont(size=14, weight="bold")
        font_val = CTkFont(size=14)

        # ── Title ─────────────────────────────────────
        CTkLabel(self, text=f"✏️ แก้ไข SO: {self.row.get('so_number', '')}",
                 font=CTkFont(size=16, weight="bold"), text_color="#1D4ED8").pack(**pad)

        CTkLabel(self, text=f"ลูกค้า: {self.row.get('customer_name', '')}",
                 font=font_val, text_color="gray").pack(padx=20, pady=(0, 10))

        # ── Section 1: ข้อมูลที่บันทึกได้ทันที ─────────
        sec1 = CTkFrame(self, fg_color="#EFF6FF", corner_radius=8, border_width=1, border_color="#BFDBFE")
        sec1.pack(fill="x", padx=20, pady=(0, 8))

        CTkLabel(sec1, text="📅 วันที่จัดส่ง / ค่าจัดส่ง  (บันทึกทันที)",
                 font=font_lbl, text_color="#1D4ED8").pack(anchor="w", padx=15, pady=(10, 4))

        row_dd = CTkFrame(sec1, fg_color="transparent")
        row_dd.pack(fill="x", padx=15, pady=4)
        CTkLabel(row_dd, text="วันที่จัดส่ง:", font=font_val, width=130, anchor="w").pack(side="left")
        self.delivery_date_entry = DateSelector(row_dd)
        self.delivery_date_entry.pack(side="left")
        # Pre-fill using set_date (expects datetime object)
        dd = self.row.get("delivery_date") or ""
        if dd:
            try:
                from datetime import datetime as _dt
                date_obj = _dt.strptime(str(dd)[:10], "%Y-%m-%d")
                self.delivery_date_entry.set_date(date_obj)
            except Exception:
                pass

        row_sc = CTkFrame(sec1, fg_color="transparent")
        row_sc.pack(fill="x", padx=15, pady=(4, 12))
        CTkLabel(row_sc, text="ค่าจัดส่ง (บาท):", font=font_val, width=130, anchor="w").pack(side="left")
        self.shipping_cost_entry = NumericEntry(row_sc, width=200)
        current_sc = self.row.get("shipping_cost")
        if current_sc is not None:
            self.shipping_cost_entry.insert(0, f"{float(current_sc):,.2f}")
        self.shipping_cost_entry.pack(side="left")

        # ── Section 2: รอบเดือนค่าคอม (ต้องขออนุมัติ SM) ─
        sec2 = CTkFrame(self, fg_color="#FFF7ED", corner_radius=8, border_width=1, border_color="#FED7AA")
        sec2.pack(fill="x", padx=20, pady=(0, 8))

        CTkLabel(sec2, text="📆 รอบเดือนค่าคอม  (ต้องขออนุมัติ SM)",
                 font=font_lbl, text_color="#C2410C").pack(anchor="w", padx=15, pady=(10, 4))

        row_cm = CTkFrame(sec2, fg_color="transparent")
        row_cm.pack(fill="x", padx=15, pady=4)
        CTkLabel(row_cm, text="เดือน:", font=font_val, width=60, anchor="w").pack(side="left")

        self.cm_month_var = tk.StringVar()
        self.cm_year_var = tk.StringVar()

        months_list = [f"{i} - {self.THAI_MONTHS[i-1]}" for i in range(1, 13)]
        self.cm_month_menu = CTkOptionMenu(row_cm, variable=self.cm_month_var,
                                           values=months_list, width=200)
        self.cm_month_menu.pack(side="left", padx=(0, 10))

        current_y = datetime.now().year
        years_list = [str(y) for y in range(current_y - 2, current_y + 3)]
        self.cm_year_menu = CTkOptionMenu(row_cm, variable=self.cm_year_var,
                                          values=years_list, width=100)
        self.cm_year_menu.pack(side="left")

        # Pre-fill commission month/year
        try:
            m_cur = int(self.row.get("commission_month") or datetime.now().month)
            y_cur = int(self.row.get("commission_year") or datetime.now().year)
            self.cm_month_var.set(f"{m_cur} - {self.THAI_MONTHS[m_cur-1]}")
            self.cm_year_var.set(str(y_cur))
        except Exception:
            self.cm_month_var.set(months_list[0])
            self.cm_year_var.set(str(current_y))

        row_reason = CTkFrame(sec2, fg_color="transparent")
        row_reason.pack(fill="x", padx=15, pady=(4, 12))
        CTkLabel(row_reason, text="เหตุผล:", font=font_val, width=60, anchor="w").pack(side="left")
        self.reason_entry = CTkEntry(row_reason, placeholder_text="ระบุเหตุผลที่ต้องการเปลี่ยนรอบเดือนค่าคอม...",
                                     width=340, font=font_val)
        self.reason_entry.pack(side="left")

        # ── Buttons ───────────────────────────────────────
        btn_row = CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(4, 12))

        CTkButton(btn_row, text="💾 บันทึก", fg_color="#16A34A", hover_color="#15803D",
                  font=CTkFont(size=14, weight="bold"), height=38,
                  command=self._save).pack(side="left", padx=(0, 10))

        CTkButton(btn_row, text="ยกเลิก", fg_color="gray", hover_color="#555555",
                  font=CTkFont(size=14), height=38,
                  command=self.destroy).pack(side="left")

    # ──────────────────────────────────────────────────────────────
    def _save(self):
        conn = None
        try:
            # ── Parse delivery date (get_date() returns "YYYY-MM-DD" in CE) ──
            delivery_date_str = self.delivery_date_entry.get_date() or (self.row.get("delivery_date") or None)

            # ── Parse shipping cost ──────────────────────
            try:
                sc_raw = self.shipping_cost_entry.get().replace(",", "").strip()
                shipping_cost_val = float(sc_raw) if sc_raw else 0.0
            except Exception:
                shipping_cost_val = float(self.row.get("shipping_cost") or 0)

            # ── Parse commission month/year ──────────────
            try:
                new_m = int(self.cm_month_var.get().split(" - ")[0])
            except Exception:
                new_m = int(self.row.get("commission_month") or datetime.now().month)
            try:
                new_y = int(self.cm_year_var.get())
            except Exception:
                new_y = int(self.row.get("commission_year") or datetime.now().year)

            old_m = int(self.row.get("commission_month") or 0)
            old_y = int(self.row.get("commission_year") or 0)
            commission_changed = (new_m != old_m) or (new_y != old_y)

            # ── DB operations ────────────────────────────
            conn = self.app_container.get_connection()
            with conn.cursor() as cur:
                # 1. Update delivery_date + shipping_cost ทันที
                cur.execute("""
                    UPDATE commissions
                    SET delivery_date = %s, shipping_cost = %s
                    WHERE id = %s AND sale_key = %s
                """, (delivery_date_str, shipping_cost_val,
                      self.row["id"], self.sale_key))

                # 2. ถ้าเปลี่ยนรอบค่าคอม → INSERT คำขอรออนุมัติ
                if commission_changed:
                    reason_text = self.reason_entry.get().strip() or "-"
                    cur.execute("""
                        INSERT INTO so_edit_requests
                            (commission_id, so_number, sale_key, sale_name,
                             requested_commission_month, requested_commission_year,
                             current_commission_month, current_commission_year,
                             request_reason, status, requested_at)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, 'pending', NOW())
                    """, (self.row["id"],
                          self.row.get("so_number"),
                          self.sale_key,
                          self.sale_name,
                          new_m, new_y,
                          old_m, old_y,
                          reason_text))

            conn.commit()
            self.app_container.release_connection(conn)
            conn = None

            # ── Feedback ─────────────────────────────────
            if commission_changed:
                msg = (f"✅ บันทึกแล้ว\n\n"
                       f"• วันที่จัดส่ง + ค่าจัดส่ง → บันทึกเรียบร้อย\n"
                       f"• รอบเดือนค่าคอม → ส่งคำขออนุมัติไปยัง SM แล้ว ⏳")
            else:
                msg = "✅ บันทึก วันที่จัดส่ง และ ค่าจัดส่ง เรียบร้อยแล้ว"
            messagebox.showinfo("บันทึกสำเร็จ", msg, parent=self)
            self.destroy()

        except Exception as e:
            if conn:
                try:
                    conn.rollback()
                    self.app_container.release_connection(conn)
                except Exception:
                    pass
            messagebox.showerror("ผิดพลาด", f"บันทึกไม่สำเร็จ: {e}", parent=self)
            traceback.print_exc() # สำหรับ Debugging