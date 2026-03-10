# sales_proxy_screen.py (ฉบับแก้ไข Layout: รับประกันการแสดงปุ่ม)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd
target_roles = ['Sales Support', 'Admin', 'Manager', 'Director']
# Import คลาสแม่
from commission_app import CommissionApp
import utils
from datetime import datetime

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
        
        # รอให้ UI วาดเสร็จแล้วค่อยสร้าง Header พิเศษด้านบน
        self.after(10, self._setup_proxy_ui)

    def _setup_proxy_ui(self):
        """สร้าง UI พิเศษสำหรับหน้า Proxy"""
        # 1. สร้าง Header สีเหลือง
        self._create_sale_selection_header()
        
        # 2. ปิดการใช้งานฟอร์มหลักก่อน (จนกว่าจะเลือกเซลส์)
        self._toggle_main_form(state="disabled")
        
        # 3. [สำคัญ] บังคับให้สร้างและแสดงปุ่ม Action Frame ทันที
        if hasattr(self, 'action_frame'):
             # ถ้ามี action_frame อยู่แล้ว ให้เคลียร์และสร้างใหม่ เพื่อให้ปุ่มของเราเข้าไปอยู่ด้วย
             for widget in self.action_frame.winfo_children():
                 widget.destroy()
             self._populate_action_frame(self.action_frame)
        else:
             # ถ้ายังไม่มี (ซึ่งแปลกมาก เพราะ CommissionApp ควรสร้างให้) ให้สร้างใหม่
             self.action_frame = CTkFrame(self.scrollable_main_container) # หรือ parent ที่เหมาะสม
             self.action_frame.pack(fill="x", pady=10)
             self._populate_action_frame(self.action_frame)


    def _create_sale_selection_header(self):
        """สร้างแถบ Header สีเหลืองสำหรับเลือกเซลส์"""
        # ล้าง Header เดิมทิ้ง
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
    #  ★ Override ส่วนปุ่มด้านล่าง เพื่อเพิ่มปุ่มพิเศษให้ Sale Support ★
    # -----------------------------------------------------------
    def _populate_action_frame(self, parent):
        # 1. สร้างปุ่มมาตรฐาน (บันทึก, ล้างค่า, นำส่ง) โดยเรียกฟังก์ชันแม่
        # (CommissionApp._populate_action_frame จะสร้างปุ่มมาตรฐานให้)
        super()._populate_action_frame(parent)

        # 2. เพิ่มปุ่มพิเศษต่อท้าย สำหรับ Sale Support เท่านั้น
        if self.user_role == 'Sale Support':
            
            # เส้นคั่น
            separator = tk.Frame(parent, height=2, bd=1, relief="sunken")
            separator.pack(fill="x", padx=20, pady=(20, 10))

            # หัวข้อเครื่องมือ
            tool_label = CTkLabel(parent, text="เครื่องมือสำหรับ Sale Support:", font=CTkFont(size=14, weight="bold"), text_color="gray50")
            tool_label.pack(anchor="w", padx=20, pady=(0, 5))

            # ปุ่มสีม่วง (ย้ายเจ้าของ SO)
            reassign_btn = CTkButton(
                parent, 
                text="🔄 ย้ายเจ้าของ SO (Reassign Owner)", 
                fg_color="#8B5CF6", # สีม่วง
                hover_color="#7C3AED",
                height=40,
                font=CTkFont(size=16, weight="bold"),
                command=self._open_reassign_window
            )
            reassign_btn.pack(fill="x", padx=20, pady=(0, 20))

    def _open_reassign_window(self):
        """เปิดหน้าต่างย้ายเจ้าของ SO"""
        try:
            from history_windows import SOReassignmentDialog
            SOReassignmentDialog(self, self.app_container)
        except ImportError:
            messagebox.showerror("Error", "ไม่พบ Class 'SOReassignmentDialog' ใน history_windows.py")
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")

    # --- ฟังก์ชัน Helper อื่นๆ ---
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
            # หา Sale Key จากชื่อที่เลือก
            selected_sale_data = next((sale for sale in self.active_sales_list if f"{sale['sale_name']} ({sale['sale_key']})" == selected_display_name), None)
            if selected_sale_data:
                self.sale_key_owner = selected_sale_data['sale_key']
                # อัปเดต key ของฟอร์มหลัก
                self.sale_key = self.sale_key_owner
                self.sale_name = selected_sale_data['sale_name']
                # เปิดใช้งานฟอร์ม
                self._toggle_main_form(state="normal")
                
                # --- [เพิ่ม] รีโหลดข้อมูลปุ่ม Action Frame ใหม่ เพื่อให้แน่ใจว่าปุ่มยังอยู่ ---
                if hasattr(self, 'action_frame'):
                     for widget in self.action_frame.winfo_children():
                         widget.destroy()
                     self._populate_action_frame(self.action_frame)
                     
            else:
                self.sale_key_owner = None
                self._toggle_main_form(state="disabled")
    
    def _toggle_main_form(self, state):
        if hasattr(self, 'scrollable_main_container') and self.scrollable_main_container.winfo_exists():
            self._recursive_toggle_state(self.scrollable_main_container, state)

    def _recursive_toggle_state(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            try:
                # ยกเว้น Action Frame ไม่ต้องปิด (เพื่อให้กดปุ่มย้าย SO ได้ตลอด)
                if child == getattr(self, 'action_frame', None):
                    continue
                
                child.configure(state=state)
            except Exception:
                pass
            if child.winfo_children():
                self._recursive_toggle_state(child, state)

class SOShortnoteSearchDialog(CTkToplevel):
    """หน้าต่างสำหรับค้นหา SO ของเซลส์ทุกคนเพื่อ Copy Shortnote (สำหรับ Sale Support)"""
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        
        self.title("ค้นหาเพื่อคัดลอก Shortnote")
        self.geometry("1100x650")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 1. โหลดรายชื่อเซลส์
        self.sales_list = self._get_sales_list()
        sale_names = ["ทั้งหมด"] + [s['display'] for s in self.sales_list]

        current_date = datetime.now()
        thai_months = ["ทั้งหมด", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        year_list = ["ทั้งหมด"] + [str(y + 543) for y in range(current_date.year - 2, current_date.year + 2)]

        self.search_var = tk.StringVar(value="")
        self.sale_var = tk.StringVar(value="ทั้งหมด")
        self.month_var = tk.StringVar(value="ทั้งหมด")
        self.year_var = tk.StringVar(value="ทั้งหมด")

        # 2. สร้าง Filter Frame
        filter_frame = CTkFrame(self, fg_color="transparent")
        filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        CTkLabel(filter_frame, text="🔍 ค้นหา (SO/ลูกค้า):").pack(side="left", padx=(5, 2))
        CTkEntry(filter_frame, textvariable=self.search_var, width=150).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เซลส์:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.sale_var, values=sale_names, width=180).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เดือน:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.month_var, values=thai_months, width=100).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="ปี:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.year_var, values=year_list, width=80).pack(side="left", padx=5)

        CTkButton(filter_frame, text="ค้นหา", command=self._load_data, width=80).pack(side="left", padx=15)
        CTkButton(filter_frame, text="ล้างค่า", command=self._clear_filters, width=80, fg_color="gray").pack(side="left", padx=5)

        # 3. สร้าง Scrollable Frame สำหรับแสดงผล
        self.results_frame = CTkScrollableFrame(self, label_text="ผลการค้นหา (เรียงจากใหม่ไปเก่า)")
        self.results_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

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

    def _clear_filters(self):
        self.search_var.set("")
        self.sale_var.set("ทั้งหมด")
        self.month_var.set("ทั้งหมด")
        self.year_var.set("ทั้งหมด")
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
            # Join ตารางเพื่อเอาชื่อเซลส์มาโชว์ด้วย
            query = """
                SELECT c.*, u.sale_name AS owner_sale_name 
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.is_active = 1
            """
            params = []

            if search_text:
                query += " AND (LOWER(c.so_number) LIKE %s OR LOWER(c.customer_name) LIKE %s)"
                params.extend([f"%{search_text}%", f"%{search_text}%"])

            if selected_sale != "ทั้งหมด":
                sale_key = next((s['key'] for s in self.sales_list if s['display'] == selected_sale), None)
                if sale_key:
                    query += " AND c.sale_key = %s"
                    params.append(sale_key)

            if selected_month != "ทั้งหมด":
                query += " AND c.commission_month = %s"
                params.append(thai_months_only.index(selected_month) + 1)

            if selected_year != "ทั้งหมด":
                query += " AND c.commission_year = %s"
                params.append(int(selected_year) - 543)

            query += " ORDER BY c.timestamp DESC LIMIT 100" # จำกัด 100 รายการเพื่อไม่ให้ช้า

            df = pd.read_sql_query(query, self.app_container.pg_engine, params=tuple(params))
            
            if df.empty:
                CTkLabel(self.results_frame, text="ไม่พบข้อมูล").pack(pady=20)
                return

            for _, row in df.iterrows():
                card = CTkFrame(self.results_frame, border_width=1, fg_color="#F0FDF4") # สีเขียวอ่อนๆ
                card.pack(fill="x", padx=5, pady=3)
                
                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
                
                # บรรทัดแรก: ข้อมูล SO
                CTkLabel(info_frame, text=f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}", font=CTkFont(weight="bold", size=14)).pack(anchor="w")
                # บรรทัดสอง: เจ้าของ และ สถานะ
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
        """ลอจิก Copy Shortnote เหมือนของหน้า Sale เป๊ะๆ"""
        try:
            so_number = so_data.get('so_number', '-')
            pickup_loc = so_data.get('pickup_location') or '-'
            
            def format_money(val):
                v = utils.convert_to_float(val)
                return f"{v:,.0f}" if v > 0 else "-"

            sales_amount = format_money(so_data.get('sales_service_amount'))
            shipping_cost = format_money(so_data.get('shipping_cost'))
            relocation_cost = format_money(so_data.get('relocation_cost'))
            cutting_fee = format_money(so_data.get('cutting_drilling_fee'))
            discount = format_money(so_data.get('coupons'))

            def format_date(date_val):
                if pd.notna(date_val) and date_val:
                    try:
                        if hasattr(date_val, 'strftime'): return date_val.strftime('%d/%m')
                        return pd.to_datetime(date_val).strftime('%d/%m')
                    except: return str(date_val)
                return "-"

            date_to_wh = format_date(so_data.get('date_to_warehouse'))
            date_to_cust = format_date(so_data.get('date_to_customer'))

            delivery_type = so_data.get('delivery_type') or '-'
            order_pur_val = so_data.get('order_pur') or '-'
            rego = so_data.get('pickup_registration') or '-'
            
            delivery_map = so_data.get('delivery_map') or '-'
            contact_name = so_data.get('onsite_contact_name') or '-'
            contact_phone = so_data.get('onsite_contact_phone') or '-'
            vehicle_type = so_data.get('vehicle_type') or '-'
            
            p1_amt = utils.convert_to_float(so_data.get('payment1_amount'))
            total_amt = utils.convert_to_float(so_data.get('total_payment_amount'))
            deposit_text = f"{p1_amt:,.0f}" if p1_amt > 0 else "-"
            full_pay_text = f"{total_amt:,.0f}" if total_amt > 0 else "-"
            remark_text = so_data.get('credit_term', 'เงินสด')

            # 🟢 [สำคัญ] ดึงชื่อเจ้าของ SO ตัวจริงมาโชว์ที่ท้ายข้อความ ไม่ใช่ชื่อ Support
            maker_name = so_data.get('owner_sale_name', 'Unknown')

            shortnote_text = (
                f"เลขที่ {so_number}\n"
                f"ยอดขาย : {sales_amount}\n"
                f"Location เข้ารับ : {pickup_loc}\n"
                f"ค่าส่ง  : {shipping_cost}\n"
                f"ค่าย้าย : {relocation_cost}\n"
                f"ค่าตัด : {cutting_fee}\n"
                f"ส่วนลดโปรโมชั่น : {discount}\n"
                f"วันที่ย้ายสินค้าเข้าคลัง132 : {date_to_wh}\n"
                f"วันที่จัดส่งลูกค้า : {date_to_cust}\n"
                f"Order Pur : {order_pur_val}\n"
                f"ทะเบียนรถ : {rego}\n"
                f"ประเภทรถ : {vehicle_type}\n"
                f"ยอดชำระ  มัดจำ {deposit_text}\n"
                f"ยอดชำระ เต็ม {full_pay_text}\n"
                f"Remark : {remark_text}\n"
                f"อนุมัติโอนยอดค้างส่วนที่เหลือ วันจัดส่งสินค้า ก่อนลงสินค้า\n\n"
                f"แผนที่จัดส่ง : {delivery_map}\n"
                f"ชื่อผู้ติดต่อหน้างาน : {contact_name}\n"
                f"เบอร์ติดต่อหน้างาน : {contact_phone}\n\n"
                f"อ้างอิงจาก Aplus Smart\n"
                f"ผู้จัดทำ: {maker_name}"
            )

            self.clipboard_clear()
            self.clipboard_append(shortnote_text)
            self.update() 

            messagebox.showinfo("คัดลอกสำเร็จ", f"คัดลอก Shortnote ของ {so_number} แล้ว!", parent=self)

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถสร้าง Shortnote ได้: {e}", parent=self)