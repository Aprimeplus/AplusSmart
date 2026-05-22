from matplotlib._api import define_aliases
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from customtkinter import CTkFrame, CTkLabel, CTkButton, CTkEntry, CTkFont, CTkScrollableFrame, CTkTabview, CTkOptionMenu, CTkCheckBox
import pandas as pd
from datetime import datetime
import psycopg2
from custom_widgets import DateSelector
from daily_report_dashboard import DailyDashboard


STATUS_THAI_MAP = {
    # --- ฝั่งเซลล์ (Sales) ---
    'Draft': 'ฉบับร่าง',
    'Edited': 'แก้ไข/บันทึกร่าง',
    'Pending Sale Manager Approval': 'รอ ผจก.ฝ่ายขายอนุมัติ',
    'Rejected by SM': 'ผจก.ขาย ตีกลับ',
    
    # --- ฝั่งจัดซื้อ (PU) ---
    'Pending PU':       'รอจัดซื้อรับงาน',
    'PO In Progress':   'อยู่ระหว่างจัดซื้อ',
    'Pending Approval': 'อยู่ระหว่างจัดซื้อ',
    'Approved':         'อยู่ระหว่างจัดซื้อ',
    'Rejected':         'ถูกตีกลับให้แก้ไข',
    'PO Sent':          'เปิด PO สำเร็จ',
    
    # --- ฝั่งบุคคล/การเงิน (HR/Finance) ---
    'Forwarded_To_HR': 'ส่งต่อให้ HR',
    'HR Verified': 'HR ตรวจสอบแล้ว',
    'Paid': 'จ่ายค่าคอมฯ แล้ว',
    'Defer Requested': 'HR ขอเลื่อนจ่าย',
    'Deferred': 'ถูกเลื่อนการจ่าย',
    
    # --- ยกเลิก (Cancelled) ---
    'Cancelled': 'ยกเลิก',
    'Cancelled by PU': 'ยกเลิกโดยจัดซื้อ'
}

class DailyReportWidget(CTkFrame):
    def __init__(self, master, app_container, sale_key_filter=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.sale_key_filter = sale_key_filter
        self.current_df = None
        
        # --- ข้อมูลสำหรับ Dropdown ---
        self.thai_months = ["ทั้งหมด", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i for i, name in enumerate(self.thai_months) if name != "ทั้งหมด"}
        
        self.sales_list = self._fetch_user_list_by_role(['Sale'])
        self.pu_list = self._fetch_user_list_by_role(['Purchasing', 'Purchasing Manager'])

        # --- 1. สร้าง Tabview ---
        self.tabs = CTkTabview(self, segmented_button_selected_color="#3B82F6")
        self.tabs.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.tab_report = self.tabs.add("📋 รายงานประจำวัน / สรุปยอด")
        self.tab_dashboard = self.tabs.add("📊 Dashboard ยอดขาย")

        # =================================================================
        # ส่วนที่ 2: หน้ารายงาน (Tab: 📋 รายงานประจำวัน)
        # =================================================================
        
        # --- 2.1 Top Bar (Filter Area) ---
        self.filter_frame = CTkFrame(self.tab_report, fg_color="#F8FAFC", border_width=1, border_color="#E2E8F0")
        self.filter_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # แถวที่ 1: วันที่ และ รอบคอม
        row1 = CTkFrame(self.filter_frame, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=5)
        
        self.use_date_var = tk.BooleanVar(value=True)
        CTkCheckBox(row1, text="กรองเฉพาะวันที่:", variable=self.use_date_var, font=CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        self.date_selector = DateSelector(row1)
        self.date_selector.pack(side="left", padx=(0, 20))
        self.date_selector.set_date(datetime.now()) 

        CTkLabel(row1, text="|  รอบคอมมิชชั่น:", font=CTkFont(weight="bold")).pack(side="left", padx=(10, 5))
        self.filter_month_var = tk.StringVar(value="ทั้งหมด")
        CTkOptionMenu(row1, variable=self.filter_month_var, values=self.thai_months, width=110).pack(side="left", padx=5)
        
        current_year = datetime.now().year + 543
        year_options = ["ทั้งหมด"] + [str(y) for y in range(current_year - 2, current_year + 2)]
        self.filter_year_var = tk.StringVar(value="ทั้งหมด")
        CTkOptionMenu(row1, variable=self.filter_year_var, values=year_options, width=90).pack(side="left", padx=5)

        # แถวที่ 2: กรอง Sale / PU และ ปุ่ม Action
        row2 = CTkFrame(self.filter_frame, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=(0, 10))

        CTkLabel(row2, text="เซลส์ (Sale):", font=CTkFont(weight="bold")).pack(side="left", padx=(0, 5))
        self.filter_sale_var = tk.StringVar(value="ทั้งหมด")
        sale_menu = CTkOptionMenu(row2, variable=self.filter_sale_var, values=["ทั้งหมด"] + self.sales_list, width=180)
        sale_menu.pack(side="left", padx=(0, 20))
        
        # 🟢 ล็อกปุ่ม Sale ถ้าระบบบังคับให้ดูได้แค่ตัวเอง (สำหรับสิทธิ์ Sale)
        if self.sale_key_filter:
            my_sale_name = next((s for s in self.sales_list if self.sale_key_filter in s), "ทั้งหมด")
            self.filter_sale_var.set(my_sale_name)
            sale_menu.configure(state="disabled")

        CTkLabel(row2, text="จัดซื้อ (PU):", font=CTkFont(weight="bold")).pack(side="left", padx=(0, 5))
        self.filter_pu_var = tk.StringVar(value="ทั้งหมด")
        CTkOptionMenu(row2, variable=self.filter_pu_var, values=["ทั้งหมด"] + self.pu_list, width=180).pack(side="left", padx=(0, 20))

        # =================================================================
        # 🟢 เพิ่ม Dropdown กรองสถานะ ตรงนี้ครับ (ก่อนถึงปุ่ม Refresh)
        # =================================================================
        CTkLabel(row2, text="สถานะ:", font=CTkFont(weight="bold")).pack(side="left", padx=(0, 5))
        self.filter_status_var = tk.StringVar(value="ทั้งหมด")
        
        # ดึงคำแปลสถานะภาษาไทยแบบไม่ซ้ำกัน มาทำเป็นตัวเลือก
        unique_thai_statuses = sorted(list(set(STATUS_THAI_MAP.values())))
        status_options = ["ทั้งหมด"] + unique_thai_statuses
        
        CTkOptionMenu(row2, variable=self.filter_status_var, values=status_options, width=160).pack(side="left", padx=(0, 20))
        # =================================================================

        self.btn_refresh = CTkButton(row2, text="🔄 ดึงข้อมูล", width=120, fg_color="#3B82F6", hover_color="#2563EB", command=self.load_report_data)
        self.btn_refresh.pack(side="right", padx=5)
        
        self.btn_export = CTkButton(row2, text="📂 Export Excel", width=120, fg_color="#10B981", hover_color="#059669", command=self.export_to_excel)
        self.btn_export.pack(side="right", padx=5)

        # --- 2.2 Table Area (Treeview) ---
        self.table_frame = CTkFrame(self.tab_report, fg_color="white", corner_radius=0)
        self.table_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", background="#E0F2FE", foreground="#0F172A", font=('Prompt', 10, 'bold'), relief="flat")
        style.configure("Treeview", background="white", foreground="#334155", rowheight=35, fieldbackground="white", font=('Arial', 10))
        style.map("Treeview", background=[('selected', '#3B82F6')], foreground=[('selected', 'white')])

        self.tree_scroll_y = ttk.Scrollbar(self.table_frame, orient="vertical")
        self.tree_scroll_y.pack(side="right", fill="y")
        self.tree_scroll_x = ttk.Scrollbar(self.table_frame, orient="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x")
        
        # รวม so_number+po_number และ prepared_by+pu_prepared_by ไว้ใน column เดียว
        self.columns = [
            "so_po", "customer_name", "comm_period", "sales_booking",
            "total_paid", "credit_balance", "payment_status", "delivery_date", "location",
            "services", "sale_pu", "status"
        ]
        self.tree = ttk.Treeview(self.table_frame, columns=self.columns, show="headings",
                                 yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)

        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        headings = {
            "so_po":          ("SO / PO",      200, "w"),
            "customer_name":  ("ชื่อลูกค้า",   160, "w"),
            "comm_period":    ("รอบคอมฯ",       80, "center"),
            "sales_booking":  ("ยอดขาย",        80, "e"),
            "total_paid":     ("ชำระแล้ว",       80, "e"),
            "credit_balance": ("ยอดค้าง",        80, "e"),
            "payment_status": ("ตรวจสอบยอด",   100, "center"),
            "delivery_date":  ("วันที่ส่ง",       80, "center"),
            "location":       ("สถานที่ส่ง",      140, "w"),
            "services":       ("ค่าบริการ",       70, "e"),
            "sale_pu":        ("Sale / PU",       90, "w"),
            "status":         ("สถานะ",           90, "center"),
        }
        
        for col, (text, width, anchor) in headings.items():
            self.tree.heading(col, text=text, command=lambda c=col: self.sort_treeview(c, False))
            self.tree.column(col, width=width, anchor=anchor)
        
        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background="#F8FAFC")
        self.tree.tag_configure('status_missing', foreground="#DC2626") 
        self.tree.tag_configure('status_over', foreground="#2563EB")    
        self.tree.tag_configure('status_ok', foreground="#059669")      
        
        self.tree.pack(fill="both", expand=True)
        
        # --- 2.3 Summary Footer ---
        self.footer_frame = CTkFrame(self.tab_report, height=50, fg_color="white", border_width=1, border_color="#E5E7EB")
        self.footer_frame.pack(fill="x", padx=15, pady=10)
        
        self.lbl_total_so = CTkLabel(self.footer_frame, text="จำนวน SO: 0", font=CTkFont(size=14, weight="bold"), text_color="#64748B")
        self.lbl_total_so.pack(side="left", padx=20)
        
        self.lbl_total_missing = CTkLabel(self.footer_frame, text="ยอดค้างชำระรวม: 0.00", font=CTkFont(size=14, weight="bold"), text_color="#DC2626")
        self.lbl_total_missing.pack(side="right", padx=15)

        self.lbl_total_paid = CTkLabel(self.footer_frame, text="ยอดชำระแล้ว: 0.00", font=CTkFont(size=14, weight="bold"), text_color="#16A34A")
        self.lbl_total_paid.pack(side="right", padx=15)

        self.lbl_total_booking = CTkLabel(self.footer_frame, text="ยอดจองรวม: 0.00", font=CTkFont(size=14, weight="bold"), text_color="#2563EB")
        self.lbl_total_booking.pack(side="right", padx=15)

        # =================================================================
        # ส่วนที่ 3: หน้าแดชบอร์ดกราฟ
        # =================================================================
        self.dashboard_view = DailyDashboard(self.tab_dashboard, self.app_container)
        self.dashboard_view.pack(fill="both", expand=True)

    def _fetch_user_list_by_role(self, roles):
        """ตัวช่วยดึงรายชื่อพนักงานจาก Database ตามสิทธิ์ (Role)"""
        try:
            placeholders = ', '.join(['%s'] * len(roles))
            query = f"SELECT sale_key, sale_name FROM sales_users WHERE role IN ({placeholders}) AND status = 'Active' ORDER BY sale_key"
            df = pd.read_sql(query, self.pg_engine, params=tuple(roles))
            return [f"{row['sale_key']} - {row['sale_name']}" for _, row in df.iterrows()]
        except Exception as e:
            print(f"Error fetching user list: {e}")
            return []

    def load_report_data(self):
        for i in self.tree.get_children(): self.tree.delete(i)
            
        try:
            # 🟢 สร้าง Query แบบ Dynamic ตาม Filter ที่เลือก
            query_base = """
                SELECT 
                    c.so_number, 
                    (SELECT STRING_AGG(po.po_number, ', ') FROM purchase_orders po WHERE po.so_number = c.so_number AND po.status != 'Cancelled') as po_number_list,
                    c.customer_name, 
                    c.commission_month,
                    c.commission_year,
                    c.sales_service_amount, 
                    c.total_payment_amount, 
                    c.difference_amount,
                    c.status,
                    COALESCE(c.date_to_customer, c.delivery_date) as final_delivery_date, 
                    c.pickup_location, 
                    (COALESCE(c.cutting_drilling_fee, 0) + COALESCE(c.other_service_fee, 0)) as service_total,
                    c.sale_key,
                    c.user_key,
                    u.sale_name as sale_name,
                    pu.sale_name as pu_name_comm,
                    (
                        SELECT u2.sale_name
                        FROM purchase_orders po
                        JOIN sales_users u2 ON po.user_key = u2.sale_key
                        WHERE po.so_number = c.so_number
                        LIMIT 1
                    ) as pu_name_po,
                    (
                        SELECT po.user_key
                        FROM purchase_orders po
                        WHERE po.so_number = c.so_number
                        LIMIT 1
                    ) as pu_key_po
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                LEFT JOIN sales_users pu ON c.user_key = pu.sale_key
                WHERE c.is_active = 1
            """
            
            where_clauses = []
            params = []

            # 1. กรองตามวันที่ (ถ้าติ๊กเลือก)
            if self.use_date_var.get() and self.date_selector.get_date():
                date_str = self.date_selector.get_date()
                where_clauses.append("DATE(c.timestamp) = %s")
                params.append(date_str)

            # 2. กรองรอบคอมมิชชั่น
            month_val = self.filter_month_var.get()
            year_val = self.filter_year_var.get()
            if month_val != "ทั้งหมด":
                where_clauses.append("c.commission_month = %s")
                params.append(self.thai_month_map[month_val])
            if year_val != "ทั้งหมด":
                where_clauses.append("c.commission_year = %s")
                params.append(int(year_val) - 543)

            # 3. กรอง Sale
            if self.sale_key_filter:
                where_clauses.append("c.sale_key = %s")
                params.append(self.sale_key_filter)
            else:
                sale_val = self.filter_sale_var.get()
                if sale_val != "ทั้งหมด":
                    sale_key = sale_val.split(" - ")[0]
                    where_clauses.append("c.sale_key = %s")
                    params.append(sale_key)

            # 4. กรอง จัดซื้อ (PU)
            pu_val = self.filter_pu_var.get()
            if pu_val != "ทั้งหมด":
                pu_key = pu_val.split(" - ")[0]
                # หางานที่ PU คนนี้สร้างในหน้า Commission หรือสร้างในหน้า PO
                where_clauses.append("(c.user_key = %s OR EXISTS (SELECT 1 FROM purchase_orders po WHERE po.so_number = c.so_number AND po.user_key = %s))")
                params.extend([pu_key, pu_key])

            # =================================================================
            # 🟢 5. กรอง สถานะ (Status) เพิ่มตรงนี้
            # =================================================================
            status_val = self.filter_status_var.get()
            if status_val != "ทั้งหมด":
                # แปลงสถานะภาษาไทยที่เลือก กลับไปเป็น Database Key (ภาษาอังกฤษ)
                # (ใช้ List Comprehension เผื่อกรณีที่มีหลาย Key แปลเป็นไทยคำเดียวกัน)
                db_keys = [k for k, v in STATUS_THAI_MAP.items() if v == status_val]
                
                if db_keys:
                    where_clauses.append("c.status IN %s")
                    params.append(tuple(db_keys))
            # =================================================================

            # ประกอบร่าง Query
            if where_clauses:
                query_base += " AND " + " AND ".join(where_clauses)
                
            query_base += " ORDER BY c.so_number ASC"
            
            df = pd.read_sql_query(query_base, self.pg_engine, params=tuple(params))
            print(f"[DEBUG] พบข้อมูลทั้งหมด: {len(df)} แถว") 

            if not df.empty:
                # --- Logic รวมชื่อ PU (เก็บไว้ใช้ถ้าจำเป็น) ---
                df['final_pu_name'] = df['pu_name_comm'].fillna(df['pu_name_po'])
                df['final_pu_name'] = df['final_pu_name'].fillna(df['user_key'].apply(lambda x: f"ID:{x}" if pd.notna(x) else "-"))
                df['final_pu_name'] = df['final_pu_name'].fillna("-")

                # --- Logic รวม ID ของ PU (แสดงใน column แทนชื่อเต็ม) ---
                df['final_pu_key'] = df['user_key'].where(df['user_key'].notna() & (df['user_key'] != ""), None)
                df['final_pu_key'] = df['final_pu_key'].fillna(df.get('pu_key_po'))
                df['final_pu_key'] = df['final_pu_key'].fillna("-")
                
                # --- 🟢 Logic สร้างข้อความ "รอบคอม" ---
                def format_comm_period(row):
                    m = row.get('commission_month')
                    y = row.get('commission_year')
                    if pd.notna(m) and pd.notna(y):
                        return f"{int(m):02d}/{int(y)+543}"
                    return "-"
                df['comm_period_display'] = df.apply(format_comm_period, axis=1)
            
            self.current_df = df

            if df.empty:
                messagebox.showinfo("แจ้งเตือน", "ไม่พบข้อมูลที่ตรงกับเงื่อนไขการค้นหา")
                self.update_summary(0, 0, 0, 0)
                return

            sum_booking = 0; sum_paid = 0; sum_missing = 0

            for i, row in df.iterrows():
                booking = float(row.get('sales_service_amount') or 0)
                paid = float(row.get('total_payment_amount') or 0)
                diff_db = float(row.get('difference_amount') or 0)
                
                status_text = ""
                status_tag = ""
                credit_display = 0.0 
                
                if abs(diff_db) < 1.00: 
                    status_text = "✅ ครบถ้วน"
                    status_tag = "status_ok"
                    credit_display = 0.0
                elif diff_db < 0: 
                    missing = abs(diff_db)
                    status_text = f"❌ ขาด {missing:,.2f}"
                    status_tag = "status_missing"
                    sum_missing += missing
                    credit_display = missing 
                else: 
                    over = diff_db
                    status_text = f"⚠️ เกิน {over:,.2f}"
                    status_tag = "status_over"
                    credit_display = 0.0 

                service_fee = float(row.get('service_total') or 0)
                d_date = row.get('final_delivery_date')
                if d_date:
                    try:
                        from datetime import date as _date
                        if hasattr(d_date, 'year'):
                            _d = d_date
                        else:
                            from datetime import datetime as _dt
                            _d = _dt.strptime(str(d_date)[:10], "%Y-%m-%d").date()
                        d_date_str = f"{_d.day:02d}/{_d.month:02d}/{_d.year + 543}"
                    except Exception:
                        d_date_str = str(d_date)
                else:
                    d_date_str = "-"
                
                sum_booking += booking; sum_paid += paid

                # 🟢 2. ดึงสถานะภาษาอังกฤษมาแปลงเป็นภาษาไทยก่อนยัดลงตาราง
                status_en = row['status']
                status_th = STATUS_THAI_MAP.get(status_en, status_en)

                # รวม SO+PO และ Sale+PU ไว้ใน cell เดียว คั่นด้วย " | "
                po_str   = row['po_number_list'] or ""
                sale_str = row.get('sale_key') or "-"
                pu_str   = row.get('final_pu_key') or ""

                so_po_str   = f"{row['so_number']}  |  {po_str}" if po_str and po_str != "-" else row['so_number']
                sale_pu_str = f"{sale_str}  /  {pu_str}" if pu_str and pu_str != "-" else sale_str

                vals = (
                    so_po_str,
                    row['customer_name'],
                    row['comm_period_display'],
                    f"{booking:,.2f}",
                    f"{paid:,.2f}",
                    f"{credit_display:,.2f}",
                    status_text,
                    d_date_str,
                    row['pickup_location'] or "-",
                    f"{service_fee:,.2f}",
                    sale_pu_str,
                    status_th,
                )
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert("", "end", values=vals, tags=(tag, status_tag))

            self.update_summary(len(df), sum_booking, sum_paid, sum_missing)

        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถดึงข้อมูลได้: {e}")
            import traceback
            traceback.print_exc()

    def update_summary(self, count, booking, paid, missing):
        self.lbl_total_so.configure(text=f"จำนวน SO: {count} ใบ")
        self.lbl_total_booking.configure(text=f"ยอดจองรวม: {booking:,.2f}")
        self.lbl_total_paid.configure(text=f"ยอดชำระแล้ว: {paid:,.2f}")
        self.lbl_total_missing.configure(text=f"ยอดค้างชำระรวม: {missing:,.2f}")

    def sort_treeview(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        try: l.sort(key=lambda t: float(t[0].replace(',', '').replace('❌ ขาด ', '').replace('💰 เกิน ', '').replace('✅ ', '')), reverse=reverse)
        except ValueError: l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
            current_tags = self.tree.item(k)['tags']
            status_tag = next((t for t in current_tags if t.startswith('status_')), '')
            stripe_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            new_tags = (stripe_tag, status_tag) if status_tag else (stripe_tag,)
            self.tree.item(k, tags=new_tags)
        self.tree.heading(col, command=lambda: self.sort_treeview(col, not reverse))

    def export_to_excel(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "กรุณาดึงข้อมูลก่อนทำการ Export")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path: return

        try:
            export_df = self.current_df.copy()
            
            export_df['sales_service_amount'] = export_df['sales_service_amount'].fillna(0)
            export_df['total_payment_amount'] = export_df['total_payment_amount'].fillna(0)
            export_df['difference_amount'] = export_df['difference_amount'].fillna(0)
            
            export_df['status'] = export_df['status'].map(lambda x: STATUS_THAI_MAP.get(x, x))
            
            def get_status_text(row):
                d = row['difference_amount']
                if abs(d) < 1: return "ครบถ้วน"
                elif d < 0: return f"ขาด {abs(d):,.2f}"
                else: return f"เกิน {d:,.2f}"
            
            export_df['payment_check_status'] = export_df.apply(get_status_text, axis=1)
            
            def get_credit_balance(row):
                d = row['difference_amount']
                return abs(d) if d < 0 else 0
            
            export_df['remaining_balance'] = export_df.apply(get_credit_balance, axis=1)
            
            rename_map = {
                "so_number": "SO Number",
                "po_number_list": "PO Number",
                "customer_name": "Customer Name",
                "comm_period_display": "รอบคอมมิชชั่น", # 🟢 เพิ่มเข้าไปใน Excel ด้วย
                "sales_service_amount": "Booking Amount",
                "total_payment_amount": "Paid Amount",
                "remaining_balance": "Credit Balance",     
                "payment_check_status": "Payment Status",  
                "status": "System Status",
                "final_delivery_date": "Delivery Date",
                "pickup_location": "Location",
                "service_total": "Service Fee",
                "sale_name": "Sale Prepared By",
                "final_pu_name": "PU Prepared By"          
            }
            
            cols_to_export = [k for k in rename_map.keys() if k in export_df.columns]
            export_df = export_df[cols_to_export].rename(columns=rename_map)
            
            export_df.to_excel(file_path, index=False)
            messagebox.showinfo("สำเร็จ", f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการบันทึกไฟล์: {e}")
            print(e)