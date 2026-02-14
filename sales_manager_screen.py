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
# แก้ไขบรรทัด Import ให้เป็นแบบนี้
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                               CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                               CTkOptionMenu, CTkRadioButton, CTkTabview, CTkCheckBox) # ✅ เพิ่ม CTkCheckBox ตรงนี้
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

        # ✅ เพิ่มแท็บ 1: รายการรออนุมัติ (เป็นหน้าแรก)
        self.approval_tab = self.tab_view.add("🗳️ รายการรออนุมัติ (SM Approval)")
        
        # แท็บเดิม
        self.daily_report_tab = self.tab_view.add("📅 รายงานประจำวัน (SO Report)")
        self.master_tab = self.tab_view.add("🛠️ ค้นหาและจัดการ (Master)")

        # สร้างเนื้อหาในแต่ละ Tab
        self._create_approval_tab(self.approval_tab)
        self._create_daily_report_widget(self.daily_report_tab) 
        self._create_master_tab(self.master_tab)            
        
        # ตั้งค่าหน้าแรกที่เปิดขึ้นมา
        self.tab_view.set("🗳️ รายการรออนุมัติ (SM Approval)")

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
        self._load_approval_data() # รีโหลดหน้าอนุมัติ
        
        if hasattr(self, 'daily_report_widget'):
            self.daily_report_widget.load_report_data()
            if hasattr(self.daily_report_widget, 'dashboard_view'):
                 self.daily_report_widget.dashboard_view._update_chart()
        
        self._load_master_data() # รีโหลด Master Tab

    # =========================================================================
    # ✅ NEW TAB: SM APPROVAL (รายการที่รอการอนุมัติ)
    # =========================================================================
    def _create_approval_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        # --- ส่วนตัวกรองและค้นหา ---
        header_frame = CTkFrame(parent, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        
        CTkLabel(header_frame, text="SO ที่รอการตรวจสอบ:", font=self.label_font).pack(side="left", padx=5)

        # ✅ เพิ่มช่อง Search (ดึงสไตล์มาจากหน้า Master)
        self.approval_search_var = tk.StringVar()
        self.approval_search_entry = CTkEntry(
            header_frame, 
            placeholder_text="ค้นหาเลขที่ SO หรือชื่อลูกค้า...", 
            width=300,
            textvariable=self.approval_search_var
        )
        self.approval_search_entry.pack(side="left", padx=20)
        self.approval_search_entry.bind("<Return>", lambda e: self._load_approval_data())

        CTkButton(header_frame, text="🔍 ค้นหา", width=80, command=self._load_approval_data).pack(side="left", padx=5)
        CTkButton(header_frame, text="🔄 รีเฟรช", width=80, fg_color="gray", command=self._load_approval_data).pack(side="left", padx=5)
        
        # --- ส่วนแสดงผล ---
        self.approval_results_frame = CTkScrollableFrame(parent, label_text="รายการรออนุมัติ (Sale -> Manager)")
        self.approval_results_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.approval_results_frame.grid_columnconfigure(0, weight=1)

        self.after(200, self._load_approval_data)

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
                term = search_txt.replace("SO", "") # ตัดคำว่า SO ออกถ้ามี เพื่อความแม่นยำ
                query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                params.extend([f"%{term}%", f"%{term}%"])

            query += " ORDER BY c.timestamp ASC" # เรียงตามลำดับงานที่ส่งมาก่อน
            
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                msg = "ไม่พบรายการที่ตรงกับเงื่อนไขการค้นหา" if search_txt else "ไม่มีรายการรออนุมัติในขณะนี้"
                CTkLabel(self.approval_results_frame, text=msg).pack(pady=30)
                return

            for _, row in df.iterrows():
                self._create_so_card(self.approval_results_frame, row.to_dict(), is_approval_mode=True)
        except Exception as e:
            print(f"Load Approval Error: {e}")

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
        amount = so_data.get('sales_service_amount', 0)

        status_colors = {
            'PO In Progress': '#E0F2FE', 'Approved': '#DCFCE7', 'Paid': '#D1FAE5',
            'Rejected by SM': '#FEE2E2', 'Cancelled': '#F3F4F6', 'Draft': '#FEF3C7',
            'Pending Sale Manager Approval': '#FEF9C3'
        }
        bg_color = status_colors.get(status, "#FFFFFF")

        card = CTkFrame(parent, border_width=1, corner_radius=8, fg_color=bg_color)
        card.pack(fill="x", padx=5, pady=5)
        
        info_text = f"SO: {so_number} | ลูกค้า: {so_data.get('customer_name')} | ยอด: {amount:,.2f} | เซลส์: {so_data.get('sale_name')}\nสถานะ: {status}"
        CTkLabel(card, text=info_text, font=self.entry_font, text_color="black", justify="left").pack(side="left", padx=15, pady=10)

        btn_frame = CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=10, pady=5)

        # ✅ ถ้าอยู่ในโหมดอนุมัติ ให้โชว์ปุ่ม "อนุมัติ" เป็นอันดับแรก
        if is_approval_mode:
            CTkButton(btn_frame, text="✅ อนุมัติ", width=90, fg_color="#16A34A", hover_color="#15803D",
                      command=lambda: self._approve_so(so_id, so_number)).pack(side="left", padx=2)

        # ปุ่มแก้ไข (SM แก้ได้ทุกใบ)
        CTkButton(btn_frame, text="🛠️ แก้ไข", width=90, fg_color="#4F46E5", hover_color="#4338CA",
                  command=lambda: self._open_so_editor_for_sm(so_number)).pack(side="left", padx=2)

        # ปุ่มตีกลับ (ตีกลับได้ทุกใบ ยกเว้นที่ Cancelled หรือตีกลับไปแล้ว)
        if status not in ['Cancelled', 'Rejected by SM']:
            CTkButton(btn_frame, text="❌ ตีกลับ", width=90, fg_color="#DC2626", hover_color="#B91C1C",
                      command=lambda: self._reject_so(so_id, so_number)).pack(side="left", padx=2)

    # ✅ ฟังก์ชันอนุมัติ SO (Workflow: SM -> PU)
    def _approve_so(self, so_id, so_number):
        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการอนุมัติ SO: {so_number} เพื่อส่งต่อให้ฝ่ายจัดซื้อใช่หรือไม่?"):
            return

        conn = None # ✅ เตรียมตัวแปร conn
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # ✅ เปลี่ยนสถานะเป็น 'Pending PU' เพื่อส่งต่อให้ฝ่ายจัดซื้อ
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'Pending PU', 
                        approver_sale_manager_key = %s, 
                        approval_date_sale_manager = CURRENT_TIMESTAMP,
                        claim_timestamp = NULL -- รีเซ็ตเป็น NULL เพื่อให้ PU คนอื่นกดรับงานได้
                    WHERE id = %s
                """, (self.user_key, so_id))
                
                # แจ้งเตือนฝ่ายจัดซื้อ
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Staff' AND status = 'Active'")
                pu_keys = [row[0] for row in cursor.fetchall()]
                
                for pu_key in pu_keys:
                    cursor.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id)
                        VALUES (%s, %s, FALSE, %s)
                    """, (pu_key, f"มี SO ใหม่ ({so_number}) ผ่านการอนุมัติแล้ว รอคุณ Claim งาน", so_id))

            conn.commit() # ✅ ยืนยันการบันทึก
            messagebox.showinfo("สำเร็จ", f"อนุมัติ SO: {so_number} เรียบร้อยแล้ว")
            self._refresh_all_tabs()
            
        except Exception as e:
            if conn: conn.rollback() # ✅ ถ้าพลาดให้ยกเลิก
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
                    cursor.execute("""
                        UPDATE commissions 
                        SET status = 'Rejected by SM', 
                            rejection_reason = %s 
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
                messagebox.showinfo("สำเร็จ", f"ตีกลับ SO: {so_number} เรียบร้อยแล้ว")
                self._refresh_all_tabs()
                
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Error", f"Reject Failed: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)

        # เรียกเปิด Dialog
        SORejectionDialog(self, so_number, save_rejection)

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

class SORejectionDialog(CTkToplevel):
    def __init__(self, master, so_number, on_confirm_callback):
        super().__init__(master)
        self.title(f"ตีกลับ SO: {so_number}")
        self.geometry("450x550")
        self.on_confirm_callback = on_confirm_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True) 

        CTkLabel(self, text=f"ระบุเหตุผลที่ตีกลับ SO: {so_number}", font=CTkFont(size=16, weight="bold")).pack(pady=15)

        self.reasons = [
            "เลขที่ใบสั่งขาย (SO) ไม่ถูกต้อง",
            "ข้อมูลชื่อลูกค้าไม่ถูกต้อง",
            "ค่าจัดส่ง / ค่าขนส่งไม่ถูกต้อง",
            "ยอดโอนชำระไม่ถูกต้อง (ไม่ตรงตามสลิป)",
            "ยอดขายสินค้าหรือค่าบริการไม่ถูกต้อง",
            "วันที่จัดส่งสินค้าไม่ถูกต้อง"
        ]
        
        self.check_vars = []
        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=30)

        for reason in self.reasons:
            var = tk.BooleanVar(value=False)
            # ✅ ตอนนี้จะเรียกใช้ CTkCheckBox ได้แล้วเพราะเรา Import มาแล้ว
            cb = CTkCheckBox(container, text=reason, variable=var, font=CTkFont(size=13))
            cb.pack(anchor="w", pady=5)
            self.check_vars.append((var, reason))

        CTkLabel(self, text="อื่นๆ / ระบุเพิ่มเติม:", font=CTkFont(size=13, weight="bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.other_reason_entry = CTkEntry(self, placeholder_text="พิมพ์เหตุผลเพิ่มเติมที่นี่...", width=380)
        self.other_reason_entry.pack(padx=30, pady=(0, 20))

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=10)
        
        CTkButton(btn_frame, text="ยกเลิก", fg_color="gray", width=100, command=self.destroy).pack(side="left", padx=5)
        CTkButton(btn_frame, text="ตกลง (ตีกลับ)", fg_color="#DC2626", hover_color="#B91C1C", width=100, command=self._on_confirm).pack(side="right", padx=5)

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