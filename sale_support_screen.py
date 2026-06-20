# sales_proxy_screen.py (ฉบับแก้ไข Layout: รับประกันการแสดงปุ่ม)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton, CTkToplevel, CTkEntry, CTkScrollableFrame
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

            # ปุ่มสีส้ม (จัดการยอดค้างชำระ)
            payment_btn = CTkButton(
                parent,
                text="💰 จัดการยอดค้างชำระ (ทุก Sale)",
                fg_color="#D97706",
                hover_color="#B45309",
                height=40,
                font=CTkFont(size=16, weight="bold"),
                command=self._open_payment_due_all_window
            )
            payment_btn.pack(fill="x", padx=20, pady=(0, 8))

            # ปุ่มสีม่วง (ย้ายเจ้าของ SO)
            reassign_btn = CTkButton(
                parent,
                text="🔄 ย้ายเจ้าของ SO (Reassign Owner)",
                fg_color="#8B5CF6",
                hover_color="#7C3AED",
                height=40,
                font=CTkFont(size=16, weight="bold"),
                command=self._open_reassign_window
            )
            reassign_btn.pack(fill="x", padx=20, pady=(0, 20))

    def _open_payment_due_all_window(self):
        """เปิดหน้าต่างจัดการยอดค้างชำระของทุก Sale"""
        AllPaymentDueWindow(self, self.app_container)

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

class AllPaymentDueWindow(CTkToplevel):
    """หน้าต่างแสดงรายการยอดค้างชำระของทุก Sale สำหรับ Sale Support"""

    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.title("จัดการยอดค้างชำระ — ทุก Sale")
        self.geometry("900x620")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Filter bar ---
        filter_frame = CTkFrame(self, fg_color="transparent")
        filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))

        CTkLabel(filter_frame, text="🔍 ค้นหา SO / ลูกค้า / เซลส์:").pack(side="left", padx=(5, 4))
        self.search_var = tk.StringVar()
        CTkEntry(filter_frame, textvariable=self.search_var, placeholder_text="พิมพ์เพื่อค้นหา...", width=220).pack(side="left", padx=4)
        CTkButton(filter_frame, text="ค้นหา", width=80, fg_color="#2563EB", command=self._load).pack(side="left", padx=4)
        CTkButton(filter_frame, text="ล้างค่า", width=80, fg_color="gray", command=self._clear).pack(side="left", padx=4)

        # --- Scrollable list ---
        self.list_frame = CTkScrollableFrame(
            self,
            label_text="รายการที่ยอดโอนชำระไม่ครบ (แก้ไขยอดโอนแล้วกดบันทึก)"
        )
        self.list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self._load()
        self.transient(master)
        self.grab_set()

    def _clear(self):
        self.search_var.set("")
        self._load()

    def _load(self):
        for w in self.list_frame.winfo_children():
            w.destroy()

        keyword = self.search_var.get().strip().lower()

        try:
            query = """
                SELECT * FROM commissions
                WHERE ROUND(difference_amount::numeric, 2) < 0
                  AND is_active = 1
                  AND status != 'Paid'
                ORDER BY timestamp DESC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine)
        except Exception as e:
            CTkLabel(self.list_frame, text=f"โหลดข้อมูลไม่สำเร็จ: {e}", text_color="red").pack(pady=20)
            return

        if keyword:
            mask = (
                df['so_number'].str.lower().str.contains(keyword, na=False) |
                df['customer_name'].str.lower().str.contains(keyword, na=False) |
                df['sale_key'].str.lower().str.contains(keyword, na=False)
            )
            df = df[mask]

        if df.empty:
            CTkLabel(self.list_frame, text="ไม่พบรายการยอดค้างชำระ").pack(pady=20)
            return

        for _, row_data in df.iterrows():
            difference = round(row_data.get('difference_amount', 0.0) or 0.0, 2)
            if difference >= 0:
                continue

            card = CTkFrame(self.list_frame, border_width=1, fg_color="#FFFBEB")
            card.pack(fill="x", padx=5, pady=4)
            card.grid_columnconfigure(0, weight=1)

            top = CTkFrame(card, fg_color="transparent")
            top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(6, 0))

            sale_key_display = row_data.get('sale_key', '-')
            info = f"SO: {row_data['so_number']}  |  ลูกค้า: {row_data['customer_name']}  |  เซลส์: {sale_key_display}  |  สถานะ: {row_data['status']}"
            CTkLabel(top, text=info, font=CTkFont(size=13, weight="bold")).pack(side="left")

            CTkButton(
                top, text="แก้ไขยอดชำระ", width=120,
                command=lambda r=row_data.to_dict(): self._open_updater(r)
            ).pack(side="right")

            CTkLabel(
                card,
                text=f"ยอดโอนขาด: {difference:,.2f} บาท",
                text_color="#B45309",
                anchor="w"
            ).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))

    def _open_updater(self, so_data_dict):
        from commission_app import PaymentUpdateWindow
        PaymentUpdateWindow(
            master=self,
            app_container=self.app_container,
            so_data=so_data_dict,
            on_save_callback=self._load
        )


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

            def _fmt_date_be(val):
                """DD/MM/YY (พ.ศ. 2 หลัก) เช่น 08/04/69"""
                try:
                    import pandas as _pd
                    dt = _pd.to_datetime(val, errors='coerce')
                    if _pd.isna(dt): return '-'
                    be2 = (dt.year + 543) % 100
                    return dt.strftime(f'%d/%m/{be2:02d}')
                except: return '-'
            date_to_wh   = _fmt_date_be(so_data.get('date_to_warehouse'))
            date_to_cust = _fmt_date_be(so_data.get('date_to_customer'))

            delivery_type = so_data.get('delivery_type') or '-'
            order_pur_val = so_data.get('order_pur') or '-'
            rego = so_data.get('pickup_registration') or '-'

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
                f"เลขที่ SO: {so_number}\n"
                f"Order Pur : {order_pur_val}\n"
                f"{separator}\n"
                f"ยอดขาย(ก่อนVat) : {sales_amount}\n"
                f"ค่าส่ง(ก่อนVat)  : {shipping_cost}\n"
                f"ค่าย้าย(ก่อนVat) : {relocation_cost}\n"
                f"ค่าตัด(ก่อนVat) : {cutting_fee}\n"
                f"ยอดชำระ(รวม Vat) : {payment_display}\n"
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
                f"การจัดส่ง : {delivery_type}\n"
                f"Location เข้ารับ : {pickup_loc}\n"
                f"ผู้ติดต่อ/เบอร์โทร : {contact_name} {contact_phone}\n"
                f"ประเภทรถ/ทะเบียนรถ : {vehicle_type} ({rego})\n"
                f"เงื่อนไขการลง/เอกสาร : {unloading_stat}\n"
                f"\n"
                f"Special Request : {special_req}\n"
                f"{separator}\n"
                f"ผู้จัดทำ: {maker_name}\n"
                f"อ้างอิงจาก: Aplus Smart"
            )

            self.clipboard_clear()
            self.clipboard_append(shortnote_text)
            self.update() 

            messagebox.showinfo("คัดลอกสำเร็จ", f"คัดลอก Shortnote ของ {so_number} แล้ว!", parent=self)

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถสร้าง Shortnote ได้: {e}", parent=self)
            import traceback
            traceback.print_exc()