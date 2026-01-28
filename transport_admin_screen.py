import tkinter as tk
from tkinter import messagebox, ttk
from customtkinter import (CTkFrame, CTkLabel, CTkEntry, CTkButton, 
                           CTkOptionMenu, CTkFont, CTkTabview)
import pandas as pd
from datetime import datetime

class TransportAdminScreen(CTkFrame):
    def __init__(self, master, app_container, user_key):
        super().__init__(master, corner_radius=0, fg_color="#FFFBEB")
        self.app_container = app_container
        self.user_key = user_key
        self.pg_engine = app_container.pg_engine
        
        # [แก้ไข] ใช้ ID แทน PX เพื่อความแม่นยำในการลบ
        self.current_selected_id = None 
        
        # Pagination
        self.current_page = 1
        self.items_per_page = 20
        self.total_pages = 1

        # Fonts
        self.header_font = CTkFont(size=22, weight="bold", family="TH Sarabun New")
        self.label_font = CTkFont(size=16, weight="bold", family="TH Sarabun New")
        self.normal_font = CTkFont(size=16, family="TH Sarabun New")
        self.result_font = CTkFont(size=16, weight="bold", family="TH Sarabun New")
        self.status_font = CTkFont(size=14, family="TH Sarabun New")

        # --- GRID LAYOUT ---
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # =================================================================
        # 🟢 LEFT SIDE: ENTRY FORM
        # =================================================================
        self.left_frame = CTkFrame(self, fg_color="transparent")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        self.main_card = CTkFrame(self.left_frame, fg_color="white", corner_radius=10)
        self.main_card.pack(fill="both", expand=True, padx=5, pady=5)
        self.main_card.grid_columnconfigure(1, weight=1)

        # Header
        CTkLabel(self.main_card, text="🚚 บันทึกค่าขนส่ง", font=self.header_font, text_color="#B45309").grid(row=0, column=0, columnspan=3, pady=(15, 10))

        # --- Inputs ---
        
        # Row 1: PO
        CTkLabel(self.main_card, text="เลขที่ PO:", font=self.label_font).grid(row=1, column=0, padx=10, pady=(5,0), sticky="e")
        self.po_entry = CTkEntry(self.main_card, placeholder_text="ระบุ PO...", width=180, font=self.normal_font)
        self.po_entry.grid(row=1, column=1, padx=10, pady=(5,0), sticky="w")
        self.po_entry.bind("<FocusOut>", self._check_po_status)
        self.po_entry.bind("<Return>", self._check_po_status)

        # Row 2: Status
        self.po_status_label = CTkLabel(self.main_card, text="", font=self.status_font)
        self.po_status_label.grid(row=2, column=1, padx=10, pady=(0,5), sticky="w")

        # Row 3: Date
        CTkLabel(self.main_card, text="วันที่ขนส่ง:", font=self.label_font).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        date_frame = CTkFrame(self.main_card, fg_color="transparent")
        date_frame.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        
        days = [str(i).zfill(2) for i in range(1, 32)]
        self.day_var = tk.StringVar(value=datetime.now().strftime("%d"))
        self.day_opt = CTkOptionMenu(date_frame, variable=self.day_var, values=days, width=60, font=self.normal_font)
        self.day_opt.pack(side="left", padx=(0, 5))

        months = [str(i).zfill(2) for i in range(1, 13)]
        self.month_var = tk.StringVar(value=datetime.now().strftime("%m"))
        self.month_opt = CTkOptionMenu(date_frame, variable=self.month_var, values=months, width=60, font=self.normal_font)
        self.month_opt.pack(side="left", padx=5)

        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 1, current_year + 2)]
        self.year_var = tk.StringVar(value=str(current_year))
        self.year_opt = CTkOptionMenu(date_frame, variable=self.year_var, values=years, width=70, font=self.normal_font)
        self.year_opt.pack(side="left", padx=5)

        # Row 4: Transport Info
        CTkLabel(self.main_card, text="ขนส่ง/คนขับ:", font=self.label_font).grid(row=4, column=0, padx=10, pady=5, sticky="e")
        self.transporter_entry = CTkEntry(self.main_card, width=220, font=self.normal_font)
        self.transporter_entry.grid(row=4, column=1, columnspan=2, padx=10, pady=5, sticky="w")

        # Row 5: License
        CTkLabel(self.main_card, text="ทะเบียน:", font=self.label_font).grid(row=5, column=0, padx=10, pady=5, sticky="e")
        self.license_plate_entry = CTkEntry(self.main_card, width=180, font=self.normal_font)
        self.license_plate_entry.grid(row=5, column=1, padx=10, pady=5, sticky="w")

        # Row 6: Phone
        CTkLabel(self.main_card, text="เบอร์โทร:", font=self.label_font).grid(row=6, column=0, padx=10, pady=5, sticky="e")
        self.phone_entry = CTkEntry(self.main_card, width=180, font=self.normal_font)
        self.phone_entry.grid(row=6, column=1, padx=10, pady=5, sticky="w")

        # Row 7: Separator
        CTkFrame(self.main_card, height=1, fg_color="#E5E7EB").grid(row=7, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        # Row 8: Cost
        CTkLabel(self.main_card, text="ยอดเงิน:", font=self.label_font, text_color="#B45309").grid(row=8, column=0, padx=10, pady=5, sticky="e")
        self.cost_entry = CTkEntry(self.main_card, placeholder_text="0.00", width=180, font=CTkFont(size=18, weight="bold"))
        self.cost_entry.grid(row=8, column=1, padx=10, pady=5, sticky="w")
        self.cost_entry.bind("<KeyRelease>", self._calculate_totals)

        # Row 9: WHT
        CTkLabel(self.main_card, text="หัก WHT:", font=self.label_font).grid(row=9, column=0, padx=10, pady=5, sticky="e")
        self.wht_option = CTkOptionMenu(self.main_card, values=["ไม่หัก (None)", "1% (ค่าขนส่ง)", "3% (ค่าบริการ)"], command=self._calculate_totals, font=self.normal_font, width=180)
        self.wht_option.grid(row=9, column=1, padx=10, pady=5, sticky="w")

        # Row 10: Result
        self.calculation_label = CTkLabel(self.main_card, text="สุทธิ: 0.00", font=self.result_font, text_color="#16A34A")
        self.calculation_label.grid(row=10, column=1, padx=10, sticky="w")

        # Row 11: Payment Type
        CTkLabel(self.main_card, text="การจ่าย:", font=self.label_font).grid(row=11, column=0, padx=10, pady=5, sticky="e")
        self.payment_type_var = tk.StringVar(value="Credit")
        CTkOptionMenu(self.main_card, variable=self.payment_type_var, values=["Credit (วางบิล)", "Cash (เงินสด)"], font=self.normal_font, width=180).grid(row=11, column=1, padx=10, pady=5, sticky="w")

        # Row 12: Remarks
        CTkLabel(self.main_card, text="หมายเหตุ:", font=self.label_font).grid(row=12, column=0, padx=10, pady=5, sticky="ne")
        self.remark_entry = CTkEntry(self.main_card, width=220, font=self.normal_font)
        self.remark_entry.grid(row=12, column=1, columnspan=2, padx=10, pady=5, sticky="w")

        # Row 13: Buttons
        btn_frame = CTkFrame(self.main_card, fg_color="transparent")
        btn_frame.grid(row=13, column=0, columnspan=3, pady=20)
        
        # ปุ่มบันทึก (สีเขียว)
        CTkButton(btn_frame, text="บันทึก", command=self._save_px, width=100, fg_color="#16A34A", hover_color="#15803D", font=self.label_font).pack(side="left", padx=5)
        
        # ปุ่มลบ (สีแดง)
        self.delete_btn = CTkButton(btn_frame, text="ลบรายการ", command=self._delete_px, width=100, fg_color="#DC2626", hover_color="#B91C1C", font=self.label_font, state="disabled")
        self.delete_btn.pack(side="left", padx=5)
        
        # ปุ่มล้าง (สีเทา)
        CTkButton(btn_frame, text="ล้าง", command=self._clear_form, width=80, fg_color="gray", font=self.label_font).pack(side="left", padx=5)

        # =================================================================
        # 🔵 RIGHT SIDE: TABS
        # =================================================================
        self.right_frame = CTkFrame(self, fg_color="transparent")
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 20), pady=10)
        
        self.tabview = CTkTabview(self.right_frame, width=500, height=600)
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_history = self.tabview.add("📜 ประวัติค่าขนส่ง (History)")
        self.tab_pending = self.tabview.add("⏳ PO ที่ยังไม่มีค่ารถ (Pending)")

        self._setup_history_tab()
        self._setup_pending_tab()

        CTkButton(self, text="ออกจากระบบ", command=self._logout, fg_color="#EF4444", width=80).place(relx=0.98, rely=0.02, anchor="ne")

        self.after(500, self._load_history_data)
        self.after(500, self._load_pending_data)

    # -------------------------------------------------------------------------
    #  TAB 1: History
    # -------------------------------------------------------------------------
    def _setup_history_tab(self):
        search_frame = CTkFrame(self.tab_history, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=5)
        
        self.hist_search = CTkEntry(search_frame, placeholder_text="ค้นหา PX หรือ PO...", width=200)
        self.hist_search.pack(side="left", padx=5)
        self.hist_search.bind("<Return>", lambda e: self._load_history_data())
        CTkButton(search_frame, text="ค้นหา", command=self._load_history_data, width=80).pack(side="left")
        CTkButton(search_frame, text="⟳", command=self._load_history_data, width=40, fg_color="gray").pack(side="left", padx=5)

        self.tree_hist = self._create_treeview(self.tab_history, ["date", "px_no", "po_no", "transporter", "cost", "status"])
        self.tree_hist.bind("<Double-1>", self._on_hist_row_click)

    def _load_history_data(self):
        search_text = self.hist_search.get().strip().upper()
        for item in self.tree_hist.get_children(): self.tree_hist.delete(item)
        try:
            conn = self.app_container.get_connection()
            cursor = conn.cursor()
            # [แก้ไข] เพิ่ม id เข้าไปใน Query (Index 0)
            query = """
                SELECT id, TO_CHAR(transport_date, 'YYYY-MM-DD'), px_number, ref_po_number, 
                       transporter_name, net_amount, status,
                       driver_phone, license_plate, transport_cost, wht_percent, remarks, payment_type
                FROM transport_orders 
                WHERE 1=1
            """
            params = []
            if search_text:
                query += " AND (ref_po_number LIKE %s OR px_number LIKE %s)"
                params.extend([f"%{search_text}%", f"%{search_text}%"])
            query += " ORDER BY id DESC LIMIT 50"
            cursor.execute(query, params)
            rows = cursor.fetchall()
            self.current_hist_data = rows
            for row in rows:
                # row[0] = id, row[1] = date, ...
                cost_fmt = f"{row[5]:,.2f}" if row[5] else "0.00"
                self.tree_hist.insert("", "end", values=(row[1], row[2], row[3], row[4], cost_fmt, row[6]))
        except Exception as e: print(f"Hist Error: {e}")
        finally: self.app_container.release_connection(conn)

    def _on_hist_row_click(self, event):
        sel = self.tree_hist.selection()
        if not sel: return
        idx = self.tree_hist.index(sel)
        if idx < len(self.current_hist_data):
            row = self.current_hist_data[idx]
            
            # 1. เติมข้อมูลลงฟอร์ม (ซึ่งจะไปเคลียร์ตัวแปรเก่าทิ้ง)
            self._fill_form(row[3], row[4], row[7], row[8], row[9], row[10], row[11], row[12], row[1])
            
            # 2. [แก้ไขสำคัญ] ตั้งค่าตัวแปรลบและเปิดปุ่มลบ *หลังจาก* เติมฟอร์มเสร็จแล้ว
            self.current_selected_id = row[0] # ID
            self.current_selected_px = row[2] # PX Number (ไว้โชว์ใน popup)
            self.delete_btn.configure(state="normal") 

    # -------------------------------------------------------------------------
    #  TAB 2: Pending
    # -------------------------------------------------------------------------
    def _setup_pending_tab(self):
        search_frame = CTkFrame(self.tab_pending, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=5)
        self.pending_search = CTkEntry(search_frame, placeholder_text="ค้นหาเลข PO...", width=200)
        self.pending_search.pack(side="left", padx=5)
        self.pending_search.bind("<Return>", lambda e: self._load_pending_data())
        CTkButton(search_frame, text="ค้นหา", command=self._load_pending_data, width=80).pack(side="left")
        CTkButton(search_frame, text="⟳", command=self._load_pending_data, width=40, fg_color="gray").pack(side="left", padx=5)

        self.tree_pending = self._create_treeview(self.tab_pending, ["date", "po_no", "supplier", "amount"])
        self.tree_pending.heading("amount", text="ยอดเงิน PO")
        self.tree_pending.column("amount", width=120, anchor="e")
        self.tree_pending.bind("<Double-1>", self._on_pending_row_click)

    def _load_pending_data(self):
        search_text = self.pending_search.get().strip().upper()
        for item in self.tree_pending.get_children(): self.tree_pending.delete(item)
        try:
            conn = self.app_container.get_connection()
            cursor = conn.cursor()
            query = """
                SELECT TO_CHAR(CAST(timestamp AS TIMESTAMP), 'YYYY-MM-DD'), po_number, supplier_name, grand_total
                FROM purchase_orders po
                WHERE NOT EXISTS (SELECT 1 FROM transport_orders t WHERE t.ref_po_number = po.po_number)
                AND status != 'Cancelled'
            """
            params = []
            if search_text:
                query += " AND po_number LIKE %s"
                params.append(f"%{search_text}%")
            query += " ORDER BY id DESC LIMIT 50"
            cursor.execute(query, params)
            rows = cursor.fetchall()
            for row in rows:
                amt = row[3] if row[3] else 0.0
                self.tree_pending.insert("", "end", values=(row[0], row[1], row[2], f"{amt:,.2f}"))
        except Exception as e: print(f"Pending Error: {e}")
        finally: self.app_container.release_connection(conn)

    def _on_pending_row_click(self, event):
        sel = self.tree_pending.selection()
        if not sel: return
        item = self.tree_pending.item(sel)
        vals = item['values']
        self._clear_form()
        self.po_entry.insert(0, vals[1])
        self._check_po_status()
        self._set_date_today()

    # -------------------------------------------------------------------------
    #  Helpers & Logic
    # -------------------------------------------------------------------------
    def _create_treeview(self, parent, columns):
        style = ttk.Style(); style.theme_use("clam")
        style.configure("Treeview.Heading", font=('TH Sarabun New', 14, 'bold'))
        style.configure("Treeview", font=('TH Sarabun New', 12), rowheight=30)
        tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        headers = {"date": "วันที่", "px_no": "เลขที่ PX", "po_no": "เลขที่ PO", "transporter": "ขนส่ง", "cost": "ยอดสุทธิ", "status": "สถานะ", "supplier": "ซัพพลายเออร์", "amount": "ยอดเงิน PO"}
        for col in columns:
            tree.heading(col, text=headers.get(col, col))
            tree.column(col, width=100, anchor="center" if col not in ["transporter", "supplier"] else "w")
        scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True, padx=(5,0), pady=5)
        scroll.pack(side="right", fill="y", padx=(0,5), pady=5)
        return tree

    def _set_date_to(self, date_str):
        if not date_str: return
        try:
            d = datetime.strptime(date_str, "%Y-%m-%d")
            self.day_var.set(str(d.day).zfill(2))
            self.month_var.set(str(d.month).zfill(2))
            self.year_var.set(str(d.year))
        except: pass

    def _set_date_today(self):
        now = datetime.now()
        self.day_var.set(str(now.day).zfill(2))
        self.month_var.set(str(now.month).zfill(2))
        self.year_var.set(str(now.year))

    def _fill_form(self, po, name, phone, plate, cost, wht, remark, payment, date_val):
        # เรียกเคลียร์ฟอร์ม เพื่อล้างหน้าจอ
        self._clear_form()
        
        self.po_entry.insert(0, po)
        self._set_date_to(date_val)
        self.transporter_entry.insert(0, name or "")
        self.phone_entry.insert(0, phone or "")
        self.license_plate_entry.insert(0, plate or "")
        self.cost_entry.insert(0, f"{cost:.2f}" if cost else "0.00")
        if wht == 1.0: self.wht_option.set("1% (ค่าขนส่ง)")
        elif wht == 3.0: self.wht_option.set("3% (ค่าบริการ)")
        else: self.wht_option.set("ไม่หัก (None)")
        self.remark_entry.insert(0, remark or "")
        self.payment_type_var.set(payment or "Credit")
        self._calculate_totals()
        self._check_po_status()

    def _check_po_status(self, event=None):
        po = self.po_entry.get().strip().upper()
        if not po:
            self.po_status_label.configure(text="", text_color="gray")
            return False
        try:
            query = "SELECT id FROM purchase_orders WHERE po_number = %s LIMIT 1"
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(po,))
            if not df.empty:
                self.po_status_label.configure(text="✅ พบ PO ในระบบ", text_color="#16A34A")
                return True
            else:
                self.po_status_label.configure(text="⚠️ ไม่พบ PO (บันทึกรอได้)", text_color="#D97706")
                return False
        except: return False

    def _calculate_totals(self, event=None):
        try:
            cost_str = self.cost_entry.get().replace(",", "")
            if not cost_str: 
                self.calculation_label.configure(text="สุทธิ: 0.00")
                return 0.0, 0.0, 0.0, 0.0
            cost = float(cost_str)
            opt = self.wht_option.get()
            wht_p = 1.0 if "1%" in opt else 3.0 if "3%" in opt else 0.0
            wht_a = cost * (wht_p / 100)
            net = cost - wht_a
            self.calculation_label.configure(text=f"หัก: {wht_a:,.2f} | สุทธิ: {net:,.2f}")
            return cost, wht_p, wht_a, net
        except: return None

    def _generate_px_number(self, po_number):
        po_clean = po_number.strip().upper()
        if po_clean.startswith("PO"):
            base_px = po_clean.replace("PO", "PX", 1)
        else:
            base_px = f"PX-{po_clean}"
            
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM transport_orders WHERE px_number LIKE %s", (f"{base_px}%",))
                count = cursor.fetchone()[0]
                return base_px if count == 0 else f"{base_px}-{count + 1}"
        except Exception as e:
            print(f"PX Gen Error: {e}")
            return f"{base_px}-{datetime.now().strftime('%M%S')}"
        finally:
            self.app_container.release_connection(conn)

    def _save_px(self):
        po = self.po_entry.get().strip().upper()
        date_val = f"{self.year_var.get()}-{self.month_var.get()}-{self.day_var.get()}"
        calc = self._calculate_totals()
        
        if not calc or not po:
            messagebox.showerror("Error", "ข้อมูลไม่ครบ", parent=self)
            return
        
        px_no = self._generate_px_number(po)
        cost, wht_p, wht_a, net = calc
        status = "Matched" if self._check_po_status() else "Pending Match"
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                sql = """
                    INSERT INTO transport_orders 
                    (px_number, ref_po_number, transport_date, transporter_name, driver_phone, license_plate, 
                     transport_cost, wht_percent, wht_amount, net_amount, remarks,
                     payment_type, status, created_by)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(sql, (
                    px_no, po, date_val, self.transporter_entry.get().strip(),
                    self.phone_entry.get().strip(), self.license_plate_entry.get().strip(),
                    cost, wht_p, wht_a, net, self.remark_entry.get().strip(),
                    self.payment_type_var.get(), status, self.user_key
                ))
            conn.commit()
            messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกข้อมูลเรียบร้อย\n\n📄 เลขที่เอกสาร: {px_no}", parent=self)
            self._clear_form()
            self._load_history_data()
            self._load_pending_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", f"{e}", parent=self)
        finally: self.app_container.release_connection(conn)

    # ---------------------------------------------------------
    #  [แก้ไข] ลบรายการ (Delete by ID)
    # ---------------------------------------------------------
    def _delete_px(self):
        if not self.current_selected_id:
            messagebox.showwarning("เตือน", "กรุณาเลือกรายการจากประวัติก่อนลบ", parent=self)
            return
            
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบรายการ '{self.current_selected_px}' ใช่หรือไม่?", icon='warning', parent=self):
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ลบด้วย ID ปลอดภัยที่สุด
                cursor.execute("DELETE FROM transport_orders WHERE id = %s", (self.current_selected_id,))
            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ลบรายการเรียบร้อยแล้ว", parent=self)
            
            self._clear_form() # ล้างหน้าจอและตัวแปร
            self._load_history_data()
            self._load_pending_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", f"ลบไม่สำเร็จ: {e}", parent=self)
        finally:
            self.app_container.release_connection(conn)

    def _clear_form(self):
        self.po_entry.delete(0, "end"); self.transporter_entry.delete(0, "end")
        self.license_plate_entry.delete(0, "end"); self.phone_entry.delete(0, "end")
        self.cost_entry.delete(0, "end"); self.remark_entry.delete(0, "end")
        self._set_date_today()
        self.wht_option.set("ไม่หัก (None)"); self.calculation_label.configure(text="สุทธิ: 0.00")
        self.po_status_label.configure(text="")
        self.po_entry.focus()
        
        # รีเซ็ตตัวแปรลบ และปิดปุ่ม
        self.current_selected_id = None
        self.current_selected_px = None
        self.delete_btn.configure(state="disabled")

    def _logout(self):
        self.app_container.show_login_screen()