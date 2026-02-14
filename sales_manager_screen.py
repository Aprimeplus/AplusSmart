import tkinter as tk
from tkinter import ttk
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                           CTkOptionMenu, CTkRadioButton, CTkTabview)
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import psycopg2.errors
import psycopg2.extras
import traceback
import utils

# --- นำเข้า Class ที่จำเป็น ---
from history_windows import SOPopupWindow
from daily_report_widget import DailyReportWidget

class SalesManagerScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role
        
        self.label_font = CTkFont(size=14, weight="bold")
        self.entry_font = CTkFont(size=14)
        
        # --- เตรียมตัวแปรสำหรับ SOPopupWindow ---
        self.so_popup = None
        self._so_create_string_vars()
        self.sale_theme = self.app_container.THEME.get("sale", {"bg": "white", "primary": "#3B82F6"})
        
        self.pg_engine = self.app_container.pg_engine
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 

        # --- 1. Header ---
        self._create_header()

        # --- 2. TabView ---
        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # สร้าง 2 Tabs
        self.daily_report_tab = self.tab_view.add("📅 รายงานประจำวัน (SO Report)")
        self.master_tab = self.tab_view.add("🛠️ ค้นหาและจัดการ (Master)")

        # สร้างเนื้อหาในแต่ละ Tab
        self._create_daily_report_widget(self.daily_report_tab) 
        self._create_master_tab(self.master_tab)            
        
        # โหลดข้อมูลเริ่มต้น
        self.tab_view.set("📅 รายงานประจำวัน (SO Report)")

    def _create_header(self):
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        CTkLabel(header_frame, text=f"Sale Manager Dashboard: {self.user_name}", font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        button_frame = CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side="right", padx=10)
        
        CTkButton(button_frame, text="🔄 Refresh All", command=self._refresh_all_tabs).pack(side="left", padx=5)
        
        CTkButton(button_frame, text="ออกจากระบบ", command=self.app_container.show_login_screen, 
                  fg_color="transparent", border_color="#D32F2F", 
                  text_color="#D32F2F", border_width=2, 
                  hover_color="#FFEBEE").pack(side="left", padx=5)

    def _refresh_all_tabs(self):
        """รีโหลดข้อมูลทุกแท็บ"""
        if hasattr(self, 'daily_report_widget'):
            self.daily_report_widget.load_report_data()
            if hasattr(self.daily_report_widget, 'dashboard_view'):
                 self.daily_report_widget.dashboard_view._update_chart()
        
        # รีโหลด Master Tab
        self._load_master_data()

    # =========================================================================
    # TAB 1: DAILY REPORT WIDGET
    # =========================================================================
    def _create_daily_report_widget(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.daily_report_widget = DailyReportWidget(parent, self.app_container)
        self.daily_report_widget.pack(fill="both", expand=True)

    # =========================================================================
    # TAB 2: MASTER EDIT & SEARCH (Auto Load + Filters)
    # =========================================================================
    def _create_master_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1) # ให้ Result ขยาย

        # --- 1. Filter Frame ---
        filter_frame = CTkFrame(parent_tab)
        filter_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        
        # ตัวแปร Filter
        current_year = datetime.now().year
        self.mst_year_var = tk.StringVar(value="ทุกปี")
        self.mst_month_var = tk.StringVar(value="ทุกเดือน")
        self.mst_day_var = tk.StringVar(value="ทุกวัน") # [เพิ่ม] วัน
        self.mst_sale_var = tk.StringVar(value="All Sales")
        
        # Options
        years = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 3, -1)]
        months = ["ทุกเดือน"] + ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                                 "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        days = ["ทุกวัน"] + [str(d) for d in range(1, 32)] # [เพิ่ม] 1-31
        sales_list = self._get_sale_list()

        # UI Filters (เรียงแถวเดียวกัน)
        CTkLabel(filter_frame, text="ปี:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_year_var, values=years, width=75, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เดือน:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_month_var, values=months, width=110, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="วัน:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_day_var, values=days, width=70, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="Sale:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_sale_var, values=sales_list, width=120, command=self._load_master_data).pack(side="left", padx=5)

        # Search Box
        self.sm_master_search_entry = CTkEntry(filter_frame, font=self.entry_font, placeholder_text="SO / ลูกค้า...", width=150)
        self.sm_master_search_entry.pack(side="left", padx=(15, 5), fill="x", expand=True)
        self.sm_master_search_entry.bind("<Return>", lambda e: self._load_master_data())
        
        # ปุ่ม
        CTkButton(filter_frame, text="🔍 ค้นหา", width=80, command=self._load_master_data).pack(side="left", padx=5)
        CTkButton(filter_frame, text="↺ รีเซ็ต", width=60, fg_color="gray", command=self._reset_master_filter).pack(side="left", padx=5)

        # --- 2. Results Frame ---
        self.sm_master_results_frame = CTkScrollableFrame(parent_tab, label_text="รายการ SO ล่าสุด (แก้ไข/ตีกลับ)")
        self.sm_master_results_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.sm_master_results_frame.grid_columnconfigure(0, weight=1)

        # Auto Load ทันทีที่เปิดโปรแกรม
        self.after(200, self._load_master_data)

    def _get_sale_list(self):
        """ดึงรายชื่อ Sale"""
        try:
            df = pd.read_sql_query("SELECT sale_key FROM sales_users WHERE role='Sale' ORDER BY sale_key", self.pg_engine)
            return ["All Sales"] + df['sale_key'].tolist()
        except:
            return ["All Sales"]

    def _reset_master_filter(self):
        """ล้างค่าตัวกรอง"""
        self.mst_year_var.set("ทุกปี")
        self.mst_month_var.set("ทุกเดือน")
        self.mst_day_var.set("ทุกวัน")
        self.mst_sale_var.set("All Sales")
        self.sm_master_search_entry.delete(0, "end")
        self._load_master_data()

    def _load_master_data(self, event=None):
        """โหลดข้อมูลตาม Filter + Search"""
        # เคลียร์ข้อมูลเก่า
        for widget in self.sm_master_results_frame.winfo_children(): widget.destroy()
        
        try:
            # Base Query
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status, u.sale_name, c.sales_service_amount
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.is_active = 1
            """
            params = []

            # 1. Filter Year
            if self.mst_year_var.get() != "ทุกปี":
                query += " AND EXTRACT(YEAR FROM c.timestamp::timestamp) = %s"
                params.append(int(self.mst_year_var.get()))

            # 2. Filter Month
            thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                           "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
            if self.mst_month_var.get() != "ทุกเดือน":
                m_idx = thai_months.index(self.mst_month_var.get()) + 1
                query += " AND EXTRACT(MONTH FROM c.timestamp::timestamp) = %s"
                params.append(m_idx)

            # 3. Filter Day
            if self.mst_day_var.get() != "ทุกวัน":
                query += " AND EXTRACT(DAY FROM c.timestamp::timestamp) = %s"
                params.append(int(self.mst_day_var.get()))

            # 4. Filter Sale
            if self.mst_sale_var.get() != "All Sales":
                query += " AND c.sale_key = %s"
                params.append(self.mst_sale_var.get())

            # 5. Text Search (SO / Customer)
            search_txt = self.sm_master_search_entry.get().strip().upper()
            if search_txt:
                term = search_txt.replace("SO", "")
                query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                params.extend([f"%{term}%", f"%{term}%"])

            # Order & Limit (10 รายการล่าสุด ตามที่ขอ)
            query += " ORDER BY c.timestamp DESC LIMIT 20" 

            # Execute
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                CTkLabel(self.sm_master_results_frame, text="ไม่พบข้อมูลตามเงื่อนไข").pack(pady=20)
                return

            # Create Cards
            for _, row in df.iterrows():
                # [แก้ไข] เรียกใช้โดยไม่ต้องส่ง mode
                self._create_so_card(self.sm_master_results_frame, row.to_dict())

        except Exception as e:
            messagebox.showerror("Error", f"Data Load Failed: {e}")
            print(traceback.format_exc())

    # =========================================================================
    # SHARED: SO CARD & ACTIONS
    # =========================================================================
    def _create_so_card(self, parent, so_data):
        """สร้าง Card แสดงข้อมูล SO พร้อมปุ่ม Edit / Reject"""
        so_id = so_data['id']
        so_number = so_data['so_number']
        status = so_data.get('status', 'N/A')
        amount = so_data.get('sales_service_amount', 0)

        # สีพื้นหลังการ์ดตามสถานะ
        status_colors = {
            'PO In Progress': '#E0F2FE', 'Approved': '#DCFCE7', 'Paid': '#D1FAE5',
            'Rejected by SM': '#FEE2E2', 'Cancelled': '#F3F4F6', 'Draft': '#FEF3C7',
            'Pending Sale Manager Approval': '#FEF9C3'
        }
        bg_color = status_colors.get(status, "#FFFFFF")

        card = CTkFrame(parent, border_width=1, corner_radius=8, fg_color=bg_color)
        card.pack(fill="x", padx=5, pady=5)
        card.grid_columnconfigure(0, weight=1)

        # Info Text
        info_text = f"SO: {so_number} | ลูกค้า: {so_data.get('customer_name')} | ยอด: {amount:,.2f} | เซลส์: {so_data.get('sale_name')}\nสถานะ: {status}"
        CTkLabel(card, text=info_text, font=self.entry_font, text_color="black", justify="left").pack(side="left", padx=15, pady=10)

        # Actions Frame
        btn_frame = CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=10, pady=5)

        # 1. ปุ่มแก้ไข (SM แก้ได้ทุกใบ)
        CTkButton(btn_frame, text="🛠️ แก้ไข", width=90, fg_color="#4F46E5", hover_color="#4338CA",
                  command=lambda: self._open_so_editor_for_sm(so_number)).pack(side="left", padx=2)

        # 2. ปุ่มตีกลับ (ตีกลับได้ทุกใบ ยกเว้นที่ Cancelled หรือตีกลับไปแล้ว)
        if status not in ['Cancelled', 'Rejected by SM']:
            CTkButton(btn_frame, text="❌ ตีกลับ", width=90, fg_color="#DC2626", hover_color="#B91C1C",
                      command=lambda: self._reject_so(so_id, so_number)).pack(side="left", padx=2)

    def _reject_so(self, so_id, so_number):
        dialog = CTkInputDialog(text=f"ระบุเหตุผลที่ตีกลับ SO: {so_number}", title="ตีกลับ SO")
        reason = dialog.get_input()
        if not reason or not reason.strip(): return

        try:
            with self.app_container.get_connection() as conn:
                with conn.cursor() as cursor:
                    # อัปเดตสถานะเป็น Rejected by SM
                    cursor.execute(
                        "UPDATE commissions SET status = 'Rejected by SM', rejection_reason = %s WHERE id = %s",
                        (reason.strip(), so_id)
                    )
                    # แจ้งเตือนเจ้าของ SO
                    cursor.execute("SELECT sale_key FROM commissions WHERE id = %s", (so_id,))
                    res = cursor.fetchone()
                    if res:
                        cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) VALUES (%s, %s, FALSE, %s)",
                                       (res[0], f"SO: {so_number} ถูกตีกลับโดย SM: {reason}", so_id))
                conn.commit()
            messagebox.showinfo("สำเร็จ", "ตีกลับเรียบร้อย")
            self._refresh_all_tabs()
        except Exception as e:
            messagebox.showerror("Error", f"Reject Failed: {e}")

    def _open_so_editor_for_sm(self, so_number):
        if self.so_popup is not None and self.so_popup.winfo_exists():
            self.so_popup.focus()
            return
        try:
            so_df = pd.read_sql_query("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", self.pg_engine, params=(so_number,))
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