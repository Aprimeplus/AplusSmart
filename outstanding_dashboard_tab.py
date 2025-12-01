import customtkinter as ctk
import pandas as pd
from datetime import datetime, timedelta
from tkinter import ttk
import traceback

class OutstandingDashboardTab(ctk.CTkFrame):
    """
    คลาสสำหรับสร้าง Tab Dashboard ติดตามยอดค้างชำระ 
    (เวอร์ชันปรับปรุง: ดึงรายชื่อ Sale ทั้งหมดมาแสดงในตัวกรอง แม้ไม่มีรายการค้าง)
    """
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine

        # --- Fonts ---
        self.header_font = ctk.CTkFont(family="Arial", size=14, weight="bold")
        self.cell_font = ctk.CTkFont(family="Arial", size=12)
        self.kpi_label_font = ctk.CTkFont(family="Arial", size=14)
        self.kpi_value_font = ctk.CTkFont(family="Arial", size=24, weight="bold")

        # --- Main Layout ---
        self.pack(fill="both", expand=True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Header Bar (Refresh Button) ---
        header_bar = ctk.CTkFrame(self, fg_color="transparent")
        header_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(0, 10))
        ctk.CTkButton(header_bar, text="🔄 รีเฟรชข้อมูล", command=self._refresh_data).pack(side="right")

        # --- Main Container for Tabs ---
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=1, column=0, sticky="nsew")
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(0, weight=1)
        
        self.after(50, self._build_ui)

    def _build_ui(self):
        """สร้างข้อมูลและวิดเจ็ตทั้งหมด"""
        self.full_df = self._fetch_outstanding_data()
        self._create_widgets()

    def _refresh_data(self):
        print("Refreshing data from database...")
        for widget in self.main_container.winfo_children():
            widget.destroy()
        self._build_ui()
        print("Refresh complete.")
    
    def _fetch_all_active_sales(self):
        """[เพิ่มใหม่] ดึงรายชื่อ Sale ที่ Active ทั้งหมดจากฐานข้อมูลโดยตรง"""
        try:
            # ดึงเฉพาะ role = 'Sale' และสถานะ Active
            query = "SELECT sale_key FROM sales_users WHERE role = 'Sale' AND status = 'Active' ORDER BY sale_key"
            df = pd.read_sql_query(query, self.pg_engine)
            return df['sale_key'].tolist()
        except Exception as e:
            print(f"Error fetching all active sales: {e}")
            return []

    def _fetch_outstanding_data(self):
        """
        ดึงข้อมูล SO ที่ยังไม่ Paid และคำนวณสถานะ 'Real Diff' (หักนายหน้า)
        """
        print("Fetching all non-paid SO status data...")
        try:
            # Query ข้อมูล
            query = """
                SELECT 
                    c.sale_key AS "พนักงานขาย",
                    c.customer_name AS "ชื่อลูกค้า",
                    c.so_number AS "เลขที่ SO",
                    c.status AS "สถานะระบบ",
                    CASE 
                        WHEN c.credit_term IS NULL OR c.credit_term = 'เงินสด' OR c.credit_term = '0' THEN 'ลูกค้าเงินสด'
                        ELSE 'ลูกค้าเครดิต'
                    END AS "ประเภทการชำระ",
                    
                    -- ยอดเต็ม (ค่าสินค้า + ค่าส่ง + ฯลฯ + VAT - WHT)
                    (COALESCE(c.sales_service_amount, 0) + 
                     COALESCE(c.cutting_drilling_fee, 0) + 
                     COALESCE(c.other_service_fee, 0) +
                     COALESCE(c.shipping_cost, 0) +
                     COALESCE(c.credit_card_fee, 0)) * 1.07 - COALESCE(c.wht_3_percent, 0) as "ยอดเต็ม",
                    
                    COALESCE(c.total_payment_amount, 0) as "ยอดที่ชำระแล้ว",
                    
                    -- ผลต่างดิบ (Paid - GrandTotal)
                    COALESCE(c.difference_amount, 0) as "ผลต่างดิบ",
                    
                    -- ค่านายหน้าดิบ (Brokerage Fee)
                    COALESCE(c.brokerage_fee, 0) as "ค่านายหน้าดิบ"
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE 
                    c.status != 'Paid'
                    AND c.is_active = 1
                    AND c.difference_amount != 0 -- ดึงเฉพาะที่มีผลต่าง
                ORDER BY c.bill_date DESC;
            """
            df = pd.read_sql_query(query, self.pg_engine)

            if df.empty:
                print("No active, non-paid SO found.")
                return pd.DataFrame()

            # --- Smart Logic: คำนวณสถานะโดยหักค่านายหน้า ---
            def calculate_real_status(row):
                raw_diff = row['ผลต่างดิบ']       # ยอดที่โอนเกิน/ขาด (จาก DB)
                broker_raw = row['ค่านายหน้าดิบ'] # ค่านายหน้า (Base)
                
                # คำนวณค่านายหน้า + VAT 7% (ยอดที่ลูกค้ามักจะโอนมาจริง)
                broker_vat = broker_raw * 1.07
                
                # Real Diff = ยอดเกิน - (ค่านายหน้า+VAT)
                # ถ้าผลลัพธ์เป็น 0 แสดงว่าที่เกินมาคือค่านายหน้าพอดี
                real_diff = raw_diff - broker_vat
                
                # กำหนดสถานะ
                status = "ครบถ้วน"
                # ยอมรับความคลาดเคลื่อนทศนิยม ±5 บาท
                if real_diff < -5.0:       
                    status = "ค้างชำระ"         # ยังขาดอยู่ (สีแดง)
                elif real_diff > 5.0:      
                    status = "ชำระเกิน"         # เกินจริง ๆ (สีเขียว)
                else:                      
                    status = "ครบถ้วน"          # พอดี (เพราะหักนายหน้าแล้วลงตัว หรือโอนพอดีแต่แรก)

                return pd.Series([status, real_diff, broker_vat])

            # Apply Logic
            df[['สถานะ', 'real_diff', 'broker_vat_display']] = df.apply(calculate_real_status, axis=1)
            
            # เตรียมข้อมูลสำหรับแสดงผล
            # ยอดคงเหลือ: ถ้าขาด แสดงยอดที่ขาด, ถ้าเกิน แสดงยอดที่เกิน, ถ้าครบถ้วน แสดง 0
            df['ยอดคงเหลือ'] = df['real_diff'].apply(lambda x: abs(x) if abs(x) > 5.0 else 0.0)
            
            # แสดงค่านายหน้า (ยอดรวม VAT)
            df['ค่านายหน้า'] = df['broker_vat_display']

            print(f"Fetched {len(df)} records with outstanding balance.")
            return df

        except Exception as e:
            print(f"!!! DATABASE ERROR fetching SO status data: {e}")
            traceback.print_exc()
            ctk.CTkLabel(self.main_container, text=f"เกิดข้อผิดพลาดในการดึงข้อมูล:\n{e}", text_color="red").pack(expand=True)
            return pd.DataFrame()


    def _create_widgets(self):
        self.tab_view = ctk.CTkTabview(self.main_container)
        self.tab_view.pack(fill="both", expand=True)

        self.cash_tab = self.tab_view.add("ลูกค้าเงินสด (ค้างชำระ)")
        self.credit_tab = self.tab_view.add("ลูกค้าเครดิต (ค้างชำระ)")
        
        # แม้ไม่มีข้อมูล ก็ยังต้องสร้าง Tab เปล่าๆ เพื่อให้แสดงผลได้
        if self.full_df.empty:
            # สร้าง DataFrame เปล่าที่มีโครงสร้าง column ครบถ้วน
            empty_df = pd.DataFrame(columns=['พนักงานขาย', 'ชื่อลูกค้า', 'เลขที่ SO', 'ยอดเต็ม', 'ยอดที่ชำระแล้ว', 'ยอดคงเหลือ', 'สถานะ', 'ค่านายหน้า', 'ประเภทการชำระ'])
            self._create_tab_content(self.cash_tab, empty_df)
            self._create_tab_content(self.credit_tab, empty_df)
            return

        cash_df = self.full_df[self.full_df['ประเภทการชำระ'] == 'ลูกค้าเงินสด']
        credit_df = self.full_df[self.full_df['ประเภทการชำระ'] == 'ลูกค้าเครดิต']
        
        self._create_tab_content(self.cash_tab, cash_df)
        self._create_tab_content(self.credit_tab, credit_df)

    def _create_tab_content(self, parent_tab, data_df):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(2, weight=1)

        filter_bar = ctk.CTkFrame(parent_tab, fg_color="transparent")
        filter_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        
        # --- 1. Filter: พนักงานขาย ---
        ctk.CTkLabel(filter_bar, text="พนักงานขาย:").pack(side="left", padx=(0, 5))
        
        # [แก้ไข] ดึงรายชื่อ Sale จาก Master Data มารวมกับคนที่มีรายการในตาราง
        active_sales = self._fetch_all_active_sales()
        
        # ดึงรายชื่อคนที่มีรายการค้างอยู่ (เผื่อคนที่ Inactive ไปแล้วแต่ยังมีรายการค้าง)
        sales_in_data = []
        if not data_df.empty and 'พนักงานขาย' in data_df.columns:
            sales_in_data = data_df['พนักงานขาย'].unique().tolist()
            
        # รวมและเรียงลำดับ
        all_sales = sorted(list(set(active_sales + sales_in_data)))
        sales_people = ['ทั้งหมด'] + all_sales
        
        sale_var = ctk.StringVar(value="ทั้งหมด")
        
        # --- 2. Filter: สถานะ ---
        ctk.CTkLabel(filter_bar, text="สถานะ:").pack(side="left", padx=(15, 5))
        status_options = ['ทั้งหมด', 'ค้างชำระ', 'ชำระเกิน', 'ครบถ้วน']
        status_var = ctk.StringVar(value="ทั้งหมด")

        # สร้าง OptionMenu
        ctk.CTkOptionMenu(filter_bar, variable=sale_var, values=sales_people,
            command=lambda choice: self._filter_and_update_tab(parent_tab, data_df, choice, status_var.get())
        ).pack(side="left", padx=5)

        ctk.CTkOptionMenu(filter_bar, variable=status_var, values=status_options,
            command=lambda choice: self._filter_and_update_tab(parent_tab, data_df, sale_var.get(), choice)
        ).pack(side="left", padx=5)

        kpi_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        kpi_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        table_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        table_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        
        parent_tab.kpi_frame = kpi_frame
        parent_tab.table_frame = table_frame

        # โหลดข้อมูลครั้งแรก
        self._filter_and_update_tab(parent_tab, data_df, "ทั้งหมด", "ทั้งหมด")

    def _filter_and_update_tab(self, parent_tab, original_df, selected_salesperson, selected_status):
        # 1. กรองตามพนักงานขาย
        if original_df.empty:
             df_filtered = original_df
        elif selected_salesperson != "ทั้งหมด":
            df_filtered = original_df[original_df['พนักงานขาย'] == selected_salesperson]
        else:
            df_filtered = original_df

        # 2. กรองตามสถานะ
        if not df_filtered.empty and selected_status != "ทั้งหมด":
            df_filtered = df_filtered[df_filtered['สถานะ'] == selected_status]

        # --- Clear Widgets ---
        for widget in parent_tab.kpi_frame.winfo_children(): widget.destroy()
        for widget in parent_tab.table_frame.winfo_children(): widget.destroy()
        
        # --- คำนวณ KPI ---
        total_outstanding = 0
        invoice_count = 0

        if not df_filtered.empty:
            # KPI 1: ยอดหนี้คงค้างรวม (นับเฉพาะ 'ค้างชำระ' เท่านั้น)
            outstanding_only = df_filtered[df_filtered['สถานะ'] == 'ค้างชำระ']
            total_outstanding = outstanding_only['ยอดคงเหลือ'].sum() if not outstanding_only.empty else 0
            
            # KPI 2: จำนวนรายการ
            invoice_count = len(df_filtered)

        self._create_kpi_box(parent_tab.kpi_frame, "ยอดหนี้คงค้างสุทธิ (Net Debt)", f"{total_outstanding:,.2f} บาท", "#DC2626").pack(side="left", fill="x", expand=True, padx=5)
        self._create_kpi_box(parent_tab.kpi_frame, "จำนวนรายการ", f"{invoice_count} รายการ", "#F97316").pack(side="left", fill="x", expand=True, padx=5)

        parent_tab.table_frame.grid_columnconfigure(0, weight=1)
        parent_tab.table_frame.grid_rowconfigure(0, weight=1)
        self._create_treeview_table(parent_tab.table_frame, df_filtered)
    
    def _create_kpi_box(self, parent, label, value, value_color):
        frame = ctk.CTkFrame(parent, border_width=1, corner_radius=8)
        lbl_label = ctk.CTkLabel(frame, text=label, font=self.kpi_label_font, text_color="gray")
        lbl_label.pack(pady=(10, 0))
        lbl_value = ctk.CTkLabel(frame, text=value, font=self.kpi_value_font, text_color=value_color)
        lbl_value.pack(pady=(0, 10))
        return frame

    def _create_treeview_table(self, parent, df):
        if df.empty:
            ctk.CTkLabel(parent, text="ไม่พบข้อมูลตามเงื่อนไข", font=("Arial", 18)).pack(expand=True)
            return
            
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=('Arial', 12, 'bold'), padding=5)
        style.configure("Treeview", rowheight=30, font=('Arial', 12))
        style.map("Treeview", background=[('selected', '#3B82F6')])

        # --- คอลัมน์ตาราง ---
        columns = ['สถานะ', 'พนักงานขาย', 'ชื่อลูกค้า', 'เลขที่ SO', 'ยอดเต็ม', 'ยอดที่ชำระแล้ว', 'ยอดคงเหลือ', 'ค่านายหน้า']
        tree = ttk.Treeview(parent, columns=columns, show='headings', style="Treeview")

        # --- กำหนดสี Tag ---
        status_colors = {
            "ค้างชำระ": "#FEE2E2",        # สีแดงอ่อน (หนี้จริง)
            "ชำระเกิน": "#D1FAE5",        # สีเขียวอ่อน (เกินจริง)
            "ครบถ้วน": "#F3F4F6"          # สีเทา (พอดี หรือหักนายหน้าแล้วพอดี)
        }
        for status, color in status_colors.items():
            tree.tag_configure(status, background=color)
            
        # ตั้งค่าหัวตาราง
        for col in columns:
            header_text = col
            if col == 'ค่านายหน้า':
                header_text = 'ค่านายหน้า (+VAT 7%)' 
            
            tree.heading(col, text=header_text, anchor="center", command=lambda _col=col: self._sort_treeview(tree, _col, False))
            
            anchor = "w"; width = 120
            if "ยอด" in col or "ค่า" in col: 
                anchor = "e"
                width = 120
            elif col == 'สถานะ': width = 150
            elif col == 'ชื่อลูกค้า': width = 200
            tree.column(col, anchor=anchor, width=width)

        for _, row in df.iterrows():
            values = list(row[columns])
            values[4] = f"{row['ยอดเต็ม']:,.2f}"
            values[5] = f"{row['ยอดที่ชำระแล้ว']:,.2f}"
            
            # แสดงยอดคงเหลือพร้อมเครื่องหมาย
            balance_val = row['ยอดคงเหลือ']
            status = row['สถานะ']
            
            if status == 'ค้างชำระ':
                values[6] = f"-{balance_val:,.2f}" # ติดลบ (หนี้)
            elif status == 'ชำระเกิน':
                values[6] = f"+{balance_val:,.2f}" # บวก (เกิน)
            else: # ครบถ้วน
                values[6] = "0.00" # แสดงเป็น 0 ให้ชัดเจนว่าเคลียร์แล้ว

            # แสดงค่านายหน้า (ยอดรวม VAT)
            values[7] = f"{row['ค่านายหน้า']:,.2f}"

            tag_to_apply = row.get('สถานะ', '')
            tree.insert("", "end", values=values, tags=(tag_to_apply,))
        
        v_scroll = ctk.CTkScrollbar(parent, command=tree.yview)
        tree.configure(yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
    
    def _sort_treeview(self, tree, col, reverse):
        data_list = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            data_list.sort(key=lambda t: float(str(t[0]).replace(",", "").replace("+", "").replace("-", "")), reverse=reverse)
        except ValueError:
            data_list.sort(key=lambda t: t[0], reverse=reverse)

        for index, (val, k) in enumerate(data_list):
            tree.move(k, '', index)
        tree.heading(col, command=lambda _col=col: self._sort_treeview(tree, _col, not reverse))