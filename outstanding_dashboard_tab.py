import customtkinter as ctk
import pandas as pd
from datetime import datetime, timedelta
from tkinter import ttk
import traceback

class OutstandingDashboardTab(ctk.CTkFrame):
    """
    คลาสสำหรับสร้าง Tab Dashboard ติดตามยอดค้างชำระ (เวอร์ชันปรับปรุงสำหรับเซลส์)
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

    def _fetch_outstanding_data(self):
        """
        ดึงข้อมูล SO ที่ "ยังไม่ถูกจ่ายเงิน" ทั้งหมด
        """
        print("Fetching all non-paid SO status data...")
        try:
            query = """
                SELECT
                    c.sale_key AS "พนักงานขาย",
                    c.customer_name AS "ชื่อลูกค้า",
                    c.so_number AS "เลขที่ SO",
                    c.status AS "สถานะ",
                    CASE
                        WHEN c.credit_term IS NULL OR c.credit_term = 'เงินสด' OR c.credit_term = '0' THEN 'ลูกค้าเงินสด'
                        ELSE 'ลูกค้าเครดิต'
                    END AS "ประเภทการชำระ",
                    (COALESCE(c.total_payment_amount, 0) + COALESCE(c.difference_amount, 0)) AS "ยอดเต็ม",
                    COALESCE(c.total_payment_amount, 0) AS "ยอดที่ชำระแล้ว",
                    COALESCE(c.difference_amount, 0) AS "ยอดคงเหลือ"
                FROM
                    commissions c
                WHERE
                    c.status != 'Paid'
                    AND c.is_active = 1;
            """
            df = pd.read_sql_query(query, self.pg_engine)

            if df.empty:
                print("No active, non-paid SO found.")
                return pd.DataFrame()

            # ✅ ไม่ต้องคำนวณยอดคงเหลือใหม่อีก เพราะ SQL มีอยู่แล้ว
            df_filtered = df[df['ยอดคงเหลือ'] > 0].copy()

            print(f"Fetched {len(df_filtered)} records with outstanding balance.")
            return df_filtered

        except Exception as e:
            print(f"!!! DATABASE ERROR fetching SO status data: {e}")
            ctk.CTkLabel(self.main_container, text=f"เกิดข้อผิดพลาดในการดึงข้อมูล:\n{e}", text_color="red").pack(expand=True)
            return pd.DataFrame()


    def _create_widgets(self):
        self.tab_view = ctk.CTkTabview(self.main_container)
        self.tab_view.pack(fill="both", expand=True)

        self.cash_tab = self.tab_view.add("ลูกค้าเงินสด (ค้างชำระ)")
        self.credit_tab = self.tab_view.add("ลูกค้าเครดิต (ค้างชำระ)")
        
        if self.full_df.empty:
            ctk.CTkLabel(self.cash_tab, text="ไม่พบข้อมูลค้างชำระ").pack(expand=True)
            ctk.CTkLabel(self.credit_tab, text="ไม่พบข้อมูลค้างชำระ").pack(expand=True)
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
        ctk.CTkLabel(filter_bar, text="กรองตามพนักงานขาย:").pack(side="left")
        
        sales_people = ['ทั้งหมด'] + sorted(self.full_df['พนักงานขาย'].unique().tolist())
        filter_var = ctk.StringVar(value="ทั้งหมด")
        
        ctk.CTkOptionMenu(filter_bar, variable=filter_var, values=sales_people,
            command=lambda choice: self._filter_and_update_tab(parent_tab, data_df, choice)
        ).pack(side="left", padx=10)

        kpi_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        kpi_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        table_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        table_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        
        parent_tab.kpi_frame = kpi_frame
        parent_tab.table_frame = table_frame

        self._filter_and_update_tab(parent_tab, data_df, "ทั้งหมด")

    def _filter_and_update_tab(self, parent_tab, original_df, selected_salesperson):
        if selected_salesperson != "ทั้งหมด":
            df_filtered = original_df[original_df['พนักงานขาย'] == selected_salesperson]
        else:
            df_filtered = original_df

        for widget in parent_tab.kpi_frame.winfo_children(): widget.destroy()
        for widget in parent_tab.table_frame.winfo_children(): widget.destroy()
        
        total_outstanding = df_filtered['ยอดคงเหลือ'].sum() if not df_filtered.empty else 0
        invoice_count = len(df_filtered)

        self._create_kpi_box(parent_tab.kpi_frame, "ยอดค้างชำระรวม", f"{total_outstanding:,.2f} บาท", "#DC2626").pack(side="left", fill="x", expand=True, padx=5)
        self._create_kpi_box(parent_tab.kpi_frame, "จำนวน SO ที่ค้างชำระ", f"{invoice_count} รายการ", "#F97316").pack(side="left", fill="x", expand=True, padx=5)

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
            ctk.CTkLabel(parent, text="ไม่พบข้อมูลค้างชำระ", font=("Arial", 18)).pack(expand=True)
            return
            
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=('Arial', 14, 'bold'), padding=10)
        style.configure("Treeview", rowheight=30, font=('Arial', 12))
        style.map("Treeview", background=[('selected', '#3B82F6')])

        columns = ['สถานะ', 'พนักงานขาย', 'ชื่อลูกค้า', 'เลขที่ SO', 'ยอดเต็ม', 'ยอดที่ชำระแล้ว', 'ยอดคงเหลือ']
        tree = ttk.Treeview(parent, columns=columns, show='headings', style="Treeview")

        status_colors = {
            "Original": "#FEFCE8", "Edited": "#FEFCE8", 
            "Pending PU": "#F1F5F9",
            "PO In Progress": "#E0E7FF", "PO Sent": "#DBEAFE", 
            "Approved by SM": "#D1FAE5", "HR Verified": "#A7F3D0",
            "Rejected by SM": "#FEF2F2", "Rejected by HR": "#FECACA",
            "Paid": "#E5E7EB"
        }
        for status, color in status_colors.items():
            tree.tag_configure(status, background=color)
        
        # <<< ไม่จำเป็นต้องใช้ Tag 'Overpaid' อีกต่อไป >>>

        for col in columns:
            tree.heading(col, text=col, anchor="center", command=lambda _col=col: self._sort_treeview(tree, _col, False))
            anchor = "w"; width = 150
            if "ยอด" in col: anchor = "e"
            elif col == 'สถานะ': width = 180
            elif col == 'ชื่อลูกค้า': width = 250
            tree.column(col, anchor=anchor, width=width)

        for _, row in df.iterrows():
            values = list(row[columns])
            values[4] = f"{row['ยอดเต็ม']:,.2f}"
            values[5] = f"{row['ยอดที่ชำระแล้ว']:,.2f}"
            values[6] = f"{row['ยอดคงเหลือ']:,.2f}"
            
            tag_to_apply = row.get('สถานะ', '')
            tree.insert("", "end", values=values, tags=(tag_to_apply,))
        
        v_scroll = ctk.CTkScrollbar(parent, command=tree.yview)
        tree.configure(yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
    
    def _sort_treeview(self, tree, col, reverse):
        data_list = [(tree.set(k, col), k) for k in tree.get_children('')]
        
        try:
            data_list.sort(key=lambda t: float(str(t[0]).replace(",", "")), reverse=reverse)
        except ValueError:
            data_list.sort(key=lambda t: t[0], reverse=reverse)

        for index, (val, k) in enumerate(data_list):
            tree.move(k, '', index)

        tree.heading(col, command=lambda _col=col: self._sort_treeview(tree, _col, not reverse))