# sales_proxy_screen.py (ฉบับแก้ไข Pack/Grid Error)

import tkinter as tk
from tkinter import messagebox
import pandas as pd
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, 
                           CTkButton, CTkToplevel, CTkEntry, CTkScrollableFrame)
import utils
from datetime import datetime

# Import คลาสแม่
from commission_app import CommissionApp

target_roles = ['Sales Support', 'Admin', 'Manager', 'Director']


# ==============================================================================
# 🟢 Class หน้าต่างค้นหา SO สำหรับ Copy Shortnote (สำหรับ Sale Support)
# ==============================================================================
class SOShortnoteSearchDialog(CTkToplevel):
    """หน้าต่างสำหรับค้นหา SO ของเซลส์ทุกคนเพื่อ Copy Shortnote (สำหรับ Sale Support)"""
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        
        self.title("ค้นหาเพื่อคัดลอก Shortnote")
        self.geometry("1150x700") 
        
        self.current_page = 0
        self.rows_per_page = 15
        self.total_pages = 1

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) 

        self.sales_list = self._get_sales_list()
        sale_names = ["ทั้งหมด"] + [s['display'] for s in self.sales_list]

        current_date = datetime.now()
        self.thai_months = ["ทั้งหมด", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        year_list = ["ทั้งหมด"] + [str(y + 543) for y in range(current_date.year - 2, current_date.year + 2)]

        self.search_var = tk.StringVar(value="")
        self.sale_var = tk.StringVar(value="ทั้งหมด")
        self.month_var = tk.StringVar(value=self.thai_months[current_date.month])
        self.year_var = tk.StringVar(value=str(current_date.year + 543))

        filter_frame = CTkFrame(self, fg_color="transparent")
        filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        CTkLabel(filter_frame, text="🔍 ค้นหา (SO/ลูกค้า):").pack(side="left", padx=(5, 2))
        CTkEntry(filter_frame, textvariable=self.search_var, width=150).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เซลส์:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.sale_var, values=sale_names, width=180).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เดือน:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.month_var, values=self.thai_months, width=100).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="ปี:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.year_var, values=year_list, width=80).pack(side="left", padx=5)

        CTkButton(filter_frame, text="ค้นหา", command=self._on_search_clicked, width=80).pack(side="left", padx=15)
        CTkButton(filter_frame, text="ล้างค่า", command=self._clear_filters, width=80, fg_color="gray").pack(side="left", padx=5)

        pagination_frame = CTkFrame(self, fg_color="transparent")
        pagination_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        self.btn_prev = CTkButton(pagination_frame, text="<< หน้าก่อนหน้า", width=100, command=self._prev_page, state="disabled")
        self.btn_prev.pack(side="left", padx=5)

        self.lbl_page_info = CTkLabel(pagination_frame, text="หน้า 1 / 1  (รวม 0 รายการ)", font=CTkFont(weight="bold"))
        self.lbl_page_info.pack(side="left", expand=True)

        self.btn_next = CTkButton(pagination_frame, text="หน้าถัดไป >>", width=100, command=self._next_page, state="disabled")
        self.btn_next.pack(side="right", padx=5)

        self.results_frame = CTkScrollableFrame(self, label_text="ผลการค้นหา (เรียงจากใหม่ไปเก่า)")
        self.results_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)

        self.after(50, self._load_data)
        self.transient(master)
        self.grab_set()

    def _get_sales_list(self):
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = 'Sale' AND status = 'Active'"
            df = pd.read_sql(query, self.app_container.pg_engine)
            res = []
            for _, row in df.iterrows():
                res.append({"key": row['sale_key'], "name": row['sale_name'], "display": f"{row['sale_name']} ({row['sale_key']})"})
            return res
        except:
            return []

    def _on_search_clicked(self):
        self.current_page = 0
        self._load_data()

    def _clear_filters(self):
        current_date = datetime.now()
        self.search_var.set("")
        self.sale_var.set("ทั้งหมด")
        self.month_var.set(self.thai_months[current_date.month])
        self.year_var.set(str(current_date.year + 543))
        
        self.current_page = 0
        self._load_data()

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._load_data()

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._load_data()

    def _load_data(self):
        for widget in self.results_frame.winfo_children():
            widget.destroy()

        search_text = self.search_var.get().strip().lower()
        selected_sale = self.sale_var.get()
        selected_month = self.month_var.get()
        selected_year = self.year_var.get()
        
        thai_months_only = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

        try:
            base_query = """
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.is_active = 1
            """
            params = []

            if search_text:
                base_query += " AND (LOWER(c.so_number) LIKE %s OR LOWER(c.customer_name) LIKE %s)"
                params.extend([f"%{search_text}%", f"%{search_text}%"])

            if selected_sale != "ทั้งหมด":
                sale_key = next((s['key'] for s in self.sales_list if s['display'] == selected_sale), None)
                if sale_key:
                    base_query += " AND c.sale_key = %s"
                    params.append(sale_key)

            if selected_month != "ทั้งหมด":
                base_query += " AND c.commission_month = %s"
                params.append(thai_months_only.index(selected_month) + 1)

            if selected_year != "ทั้งหมด":
                base_query += " AND c.commission_year = %s"
                params.append(int(selected_year) - 543)

            count_query = "SELECT COUNT(c.id) " + base_query
            total_records = pd.read_sql_query(count_query, self.app_container.pg_engine, params=tuple(params)).iloc[0, 0]
            
            self.total_pages = max(1, (total_records + self.rows_per_page - 1) // self.rows_per_page)

            self.lbl_page_info.configure(text=f"หน้า {self.current_page + 1} / {self.total_pages}  (รวม {total_records} รายการ)")
            self.btn_prev.configure(state="normal" if self.current_page > 0 else "disabled")
            self.btn_next.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")

            offset = self.current_page * self.rows_per_page
            data_query = "SELECT c.*, u.sale_name AS owner_sale_name " + base_query + " ORDER BY c.timestamp DESC LIMIT %s OFFSET %s"
            
            data_params = params + [self.rows_per_page, offset]
            df = pd.read_sql_query(data_query, self.app_container.pg_engine, params=tuple(data_params))
            
            if df.empty:
                CTkLabel(self.results_frame, text="ไม่พบข้อมูล").pack(pady=20)
                return

            for _, row in df.iterrows():
                card = CTkFrame(self.results_frame, border_width=1, fg_color="#F0FDF4") 
                card.pack(fill="x", padx=5, pady=3)
                
                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
                
                CTkLabel(info_frame, text=f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}", font=CTkFont(weight="bold", size=14)).pack(anchor="w")
                CTkLabel(info_frame, text=f"เซลส์: {row['owner_sale_name']} | สถานะ: {row['status']}", text_color="gray50").pack(anchor="w")

                btn_frame = CTkFrame(card, fg_color="transparent")
                btn_frame.pack(side="right", padx=10)
                
                CTkButton(
                    btn_frame, text="📋 Copy Shortnote", width=120, height=35,
                    fg_color="#22C55E", hover_color="#16A34A", font=CTkFont(weight="bold"),
                    command=lambda r=row: self._copy_so_shortnote(r.to_dict())
                ).pack(side="left", padx=2)

        except Exception as e:
            messagebox.showerror("Error", f"โหลดข้อมูลล้มเหลว: {e}", parent=self)

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

# ==============================================================================
# 🟢 Class หน้าจอหลักของ Sale Support
# ==============================================================================
class SalesProxyScreen(CommissionApp):
    """
    หน้าจอสำหรับให้ Role อื่น (เช่น Sale Support, HR) ทำงานในนามของ Sale
    """
    def __init__(self, master, app_container, proxy_user_key, proxy_user_name, user_role, role_to_proxy="Sale", show_logout_button=False):
        
        self.proxy_user_key = proxy_user_key
        self.proxy_user_name = proxy_user_name
        self.role_to_proxy = role_to_proxy
        self.show_logout_button = show_logout_button
        self.sale_key_owner = None
        self.user_role = user_role 

        # เรียกใช้ Class แม่ (CommissionApp)
        super().__init__(master=master, 
                         app_container=app_container, 
                         sale_key=proxy_user_key, 
                         sale_name=proxy_user_name,
                         user_role=user_role,
                         show_logout_button=show_logout_button,
                         create_default_header=True)

        self.active_sales_list = self._get_all_active_sales()
        self.after(10, self._setup_proxy_ui)

    def _setup_proxy_ui(self):
        """สร้าง UI พิเศษสำหรับหน้า Proxy"""
        self._create_sale_selection_header()
        self._toggle_main_form(state="disabled")

    def _create_sale_selection_header(self):
        for widget in self.header_frame.winfo_children():
            widget.destroy()

        selection_header = CTkFrame(self.header_frame, fg_color="#FFFBEB", border_width=1, border_color="#FBBF24")
        selection_header.pack(side="left", fill="x", expand=True, pady=(0, 10))
        
        CTkLabel(selection_header, text=f"ผู้ดำเนินการ: {self.proxy_user_name} ({self.proxy_user_key})", font=CTkFont(size=16, weight="bold"), text_color="black").pack(side="left", padx=10, pady=10)
        
        sale_selection_frame = CTkFrame(selection_header, fg_color="transparent")
        sale_selection_frame.pack(side="left", padx=10, pady=10)
        
        CTkLabel(sale_selection_frame, text="ทำงานในนามของ (เซลส์):", font=CTkFont(size=14), text_color="black").pack(side="left")
        
        sale_display_names = [f"{sale['sale_name']} ({sale['sale_key']})" for sale in self.active_sales_list]
        self.selected_sale_key_full = tk.StringVar(value="- กรุณาเลือกเซลส์ -")
        
        self.sale_selector_menu = CTkOptionMenu(
            sale_selection_frame,
            variable=self.selected_sale_key_full,
            values=["- กรุณาเลือกเซลส์ -"] + sale_display_names,
            command=self._on_sale_selected
        )
        self.sale_selector_menu.pack(side="left", padx=10)

    # -----------------------------------------------------------
    #  ★ Override ส่วนสร้างปุ่มด้านล่าง เพื่อพ่วงปุ่ม Support เข้าไปอย่างถูกต้อง ★
    # -----------------------------------------------------------
    def _populate_action_frame(self, parent):
        # ให้ Class แม่ (CommissionApp) สร้างปุ่มบันทึกปกติลงใน parent
        super()._populate_action_frame(parent)

        # 🟢 เพิ่มปุ่มเครื่องมือ Sale Support ต่อท้ายใน parent เดียวกัน
        if self.user_role == 'Sale Support':
            self.support_tools_frame = CTkFrame(parent, fg_color="transparent")
            self.support_tools_frame.pack(fill="x", pady=(10, 0))

            separator = tk.Frame(self.support_tools_frame, height=2, bd=1, relief="sunken")
            separator.pack(fill="x", padx=20, pady=(20, 10))

            tool_label = CTkLabel(self.support_tools_frame, text="เครื่องมือสำหรับ Sale Support:", font=CTkFont(size=14, weight="bold"), text_color="gray50")
            tool_label.pack(anchor="w", padx=20, pady=(0, 5))

            btn_container = CTkFrame(self.support_tools_frame, fg_color="transparent")
            btn_container.pack(fill="x", padx=20, pady=(0, 20))
            btn_container.grid_columnconfigure(0, weight=1)
            btn_container.grid_columnconfigure(1, weight=1)

            reassign_btn = CTkButton(
                btn_container, 
                text="🔄 ย้ายเจ้าของ SO", 
                fg_color="#8B5CF6", hover_color="#7C3AED",
                height=40, font=CTkFont(size=16, weight="bold"),
                command=self._open_reassign_window
            )
            reassign_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

            shortnote_btn = CTkButton(
                btn_container, 
                text="📋 ค้นหา & Copy Shortnote", 
                fg_color="#22C55E", hover_color="#16A34A",
                height=40, font=CTkFont(size=16, weight="bold"),
                command=self._open_shortnote_search_window
            )
            shortnote_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    def _open_reassign_window(self):
        try:
            from history_windows import SOReassignmentDialog
            SOReassignmentDialog(self, self.app_container)
        except ImportError:
            messagebox.showerror("Error", "ไม่พบ Class 'SOReassignmentDialog' ใน history_windows.py")
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")

    def _open_shortnote_search_window(self):
        SOShortnoteSearchDialog(self, self.app_container)

    # --- ฟังก์ชัน Helper ---
    def _get_all_active_sales(self):
        try:
            query = "SELECT sale_key, sale_name FROM sales_users WHERE role = %s AND status = 'Active' ORDER BY sale_name"
            df = pd.read_sql(query, self.pg_engine, params=(self.role_to_proxy,))
            return df.to_dict('records')
        except Exception as e:
            print(f"Database Error: {e}")
            return []

    def _on_sale_selected(self, selected_display_name):
        if "- กรุณาเลือกเซลส์ -" in selected_display_name:
            self.sale_key_owner = None
            self._toggle_main_form(state="disabled")
        else:
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                self.sale_key = self.sale_key_owner
                self.sale_name = selected_sale_data['sale_name']
                self._toggle_main_form(state="normal")
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="disabled")
    
    def _toggle_main_form(self, state):
        if hasattr(self, 'scrollable_main_container') and self.scrollable_main_container.winfo_exists():
            self._recursive_toggle_state(self.scrollable_main_container, state)

    def _recursive_toggle_state(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            try:
                # 🟢 ข้ามเฟรมของ Support Tool ไม่ให้โดน Disable (จะได้กดค้นหาได้แม้ยังไม่เลือกเซลส์)
                if hasattr(self, 'support_tools_frame') and child == self.support_tools_frame:
                    continue
                child.configure(state=state)
            except Exception:
                pass
            if child.winfo_children():
                self._recursive_toggle_state(child, state)