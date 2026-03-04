import tkinter as tk
from tkinter import messagebox
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkButton, CTkOptionMenu
import pandas as pd
import psycopg2.extras
from datetime import datetime

# ติดตั้งด้วย: pip install tksheet
try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False

class CostBenchmarkScreen(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container

        # 🟢 ดึงชื่อ User ที่ Login อยู่ปัจจุบัน
        self.current_user = getattr(self.app_container, 'current_user_key', 'PU_Default')

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.sales_list = []
        self.supplier_list = []
        self.product_list = []
        self.product_sku_map = {} 
        self.supplier_code_map = {} 

        # --- 1. Header & Filters ---
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 10))
        header_frame.grid_columnconfigure(5, weight=1) 

        CTkLabel(header_frame, text=f"📊 ตารางของคุณ: {self.current_user}",
                 font=CTkFont(size=20, weight="bold"), text_color="#1F2937").grid(row=0, column=0, padx=(0, 20))

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        now = datetime.now()
        
        CTkLabel(header_frame, text="รอบบิลเดือน:").grid(row=0, column=1, padx=(0, 5))
        self.month_var = tk.StringVar(value=self.thai_months[now.month - 1])
        CTkOptionMenu(header_frame, variable=self.month_var, values=self.thai_months, width=100, command=self._load_from_db).grid(row=0, column=2, padx=(0, 10))

        CTkLabel(header_frame, text="ปี:").grid(row=0, column=3, padx=(0, 5))
        current_year_th = str(now.year + 543)
        year_list = [str(int(current_year_th) + i) for i in range(-2, 3)] 
        self.year_var = tk.StringVar(value=current_year_th)
        CTkOptionMenu(header_frame, variable=self.year_var, values=year_list, width=80, command=self._load_from_db).grid(row=0, column=4, padx=(0, 15))

        CTkButton(header_frame, text="🔄 โหลดข้อมูล", fg_color="#3B82F6", hover_color="#2563EB", width=90,
                  command=self._load_from_db).grid(row=0, column=5, sticky="w")

        # ปุ่มจัดการ
        btn_frame = CTkFrame(header_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=6, sticky="e")

        CTkButton(btn_frame, text="💾 บันทึกลงฐานข้อมูล", fg_color="#10B981", hover_color="#059669",
                  command=self._save_to_db).pack(side="left", padx=5)
                  
        CTkButton(btn_frame, text="🗑️ ลบบรรทัด", fg_color="#EF4444", hover_color="#DC2626",
                  command=self._delete_selected_rows).pack(side="left", padx=5)
                  
        CTkButton(btn_frame, text="➕ เพิ่มบรรทัดใหม่",
                  command=self._add_new_row).pack(side="left", padx=5)

        self.columns = [
            "วันที่ขอราคา","Order No.", "Sale Order No.", "ชื่อ Sale",
            "PRIORITY", "WIN RATE %", "สถานะ", "QT", "Select",
            "หมวด" ,"ชื่อ Supplier", "แบรนด์", "รายการสินค้า",
            "หมายเหตุ (ความยาว, OD)", "หมายเหตุ", "Product SKU.", "จำนวน", "ต้นทุน/เส้น",
            "น้ำหนัก/เส้น", "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (บาท)",
            "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1", "ส่วนลด 2 (บาท)", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2",
            "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย (ซื้อ)", "ค่าย้าย/เส้น",
            "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)", "Markup Guide (%)", "Markup/กก.",
            "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup", "ค่าส่ง (ขาย)",
            "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น",
            "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat.", "ชื่อ Supplier2",
            "Sup ID.", "คลังสินค้า ต้นทาง", "ปลายทาง", "หมายเหตุ2"
        ]

        self._load_dropdown_data()

        # ================================================================== #
        # แถบสถานะติดตามตัว (Sticky Info Banner)
        # ================================================================== #
        self.banner_frame = CTkFrame(self, fg_color="#FEF3C7", border_width=1, border_color="#F59E0B", corner_radius=8)
        self.banner_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        
        self.current_item_label = CTkLabel(
            self.banner_frame, 
            text="📌 คลิกที่ตารางเพื่อดูว่ากำลังแก้ไขรายการไหนอยู่...", 
            font=CTkFont(size=15, weight="bold"), 
            text_color="#B45309"
        )
        self.current_item_label.pack(pady=8, padx=15, anchor="w")
        # ================================================================== #

        # --- 3. ตาราง ---
        table_frame = tk.Frame(self, bg="white")
        table_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 20))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        if HAS_TKSHEET:
            self._build_tksheet(table_frame)
            self.after(200, self._load_from_db)
        else:
            tk.Label(table_frame, text="⚠️ กรุณาติดตั้ง tksheet", fg="red", bg="white").pack(expand=True)

    def _load_dropdown_data(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT sale_name FROM sales_users WHERE role = 'Sale' AND status = 'Active'")
                self.sales_list = [row[0] for row in cursor.fetchall() if row[0]]

                cursor.execute("SELECT supplier_name, supplier_code FROM suppliers")
                for row in cursor.fetchall():
                    if row[0]:
                        self.supplier_list.append(row[0])
                        self.supplier_code_map[row[0]] = row[1] or ""

                cursor.execute("SELECT product_name, product_code FROM products")
                for row in cursor.fetchall():
                    if row[0]:
                        self.product_list.append(row[0])
                        self.product_sku_map[row[0]] = row[1] or ""
        except Exception as e:
            print(f"Error loading dropdown data: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    # ================================================================== #
    def _build_tksheet(self, parent):
        self.sheet = Sheet(
            parent,
            headers=self.columns,
            data=[[""] * len(self.columns) for _ in range(20)],
            theme="light blue",
            row_height=30,
            header_height=35,
            font=("Tahoma", 11, "normal"),
            header_font=("Tahoma", 11, "bold"),
            show_row_index=True,
            row_index_width=40,
            column_width=120,
            empty_horizontal=0,
            empty_vertical=0,
        )
        self.sheet.grid(row=0, column=0, sticky="nsew")

        self.sheet.enable_bindings((
            "single_select", "row_select", "column_width_resize",
            "arrowkeys", "right_click_popup_menu",
            "rc_select", "copy", "cut", "paste",
            "delete", "undo", "edit_cell",
        ))

        self.sheet.set_options(
            grid_color="#000000", outline_color="#000000", table_bg="white", table_fg="black", 
            table_grid_fg="#000000", header_bg="#D1D5DB", header_fg="#111827", header_grid_fg="#000000",
            header_selected_cells_bg="#9CA3AF", row_index_bg="#F3F4F6", row_index_fg="#111827", 
            row_index_grid_fg="#000000", selected_cells_border_color="#3B82F6", table_selected_cells_border_color="#3B82F6"
        )

        for i, col in enumerate(self.columns):
            if "รายการสินค้า" in col or "หมายเหตุ" in col:
                self.sheet.column_width(i, 250)

        # 🟢 เรียกสร้าง Dropdown แค่ครั้งเดียวตอนเริ่มโปรแกรม
        self._apply_formatting()
        
        self.sheet.bind("<ButtonRelease-1>", self._trigger_banner_update)
        self.sheet.bind("<KeyRelease>", self._trigger_banner_update)
        self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)

    def _trigger_banner_update(self, event=None):
        self.after(50, self._update_current_item_banner)

    def _update_current_item_banner(self):
        try:
            cells = self.sheet.get_selected_cells()
            if not cells:
                return
            
            row = list(cells)[0][0]
            
            # ดึงข้อมูลจากคอลัมน์ต่างๆ
            order_no = str(self.sheet.get_cell_data(row, self.columns.index("Order No.")) or "-").strip()
            supplier = str(self.sheet.get_cell_data(row, self.columns.index("ชื่อ Supplier")) or "-").strip()
            product = str(self.sheet.get_cell_data(row, self.columns.index("รายการสินค้า")) or "-").strip()

            
            # 🟢 [เพิ่มใหม่] ดึงข้อมูลจากช่อง หมายเหตุ (ความยาว, OD)
            remark = str(self.sheet.get_cell_data(row, self.columns.index("หมายเหตุ (ความยาว, OD)")) or "-").strip()
            
            if order_no == "-" and supplier == "-" and product == "-":
                self.current_item_label.configure(text=f"📌 กำลังแก้ไข บรรทัดที่ {row+1} (ช่องว่าง)")
            else:
                # 🟢 [เพิ่มใหม่] นำตัวแปร remark มาต่อท้ายข้อความ
                self.current_item_label.configure(
                    text=f"📌 กำลังแก้ไข บรรทัดที่ {row+1}   👉   Order: {order_no}   |   Supplier: {supplier}   |   รายการสินค้า: {product}   |   หมายเหตุ: {remark}"
                )
        except Exception:
            pass

    def _apply_formatting(self):
        # 1. จัดการ Dropdown
        self.sheet.create_dropdown("all", self.columns.index("ชื่อ Sale"), values=[""] + self.sales_list, state="normal")
        self.sheet.create_dropdown("all", self.columns.index("PRIORITY"), values=["", "HOT", "WARM", "COLD", "ไม่แจ้ง"], state="readonly")
        
        status_opts = ["", "WIN", "STOCK", "LOSE - เซลล์ไม่ทราบสาเหตุ", "LOSE - ลูกค้าได้ราคาถูกกว่า (มีราคาเทียบ)",
                       "LOSE - ลูกค้าได้ราคาถูกกว่า (ไม่มีราคาเทียบ)", "LOSE - ไม่มีกำหนดใช้งานที่แน่นอน เช่น ขอราคาเพื่อเสนอ",
                       "LOSE - ยื่นประมูลงาน (ระบุเดือนในหมายเหตุ)", "LOSE - ลูกค้าเปลี่ยนสเปคการใช้งาน", "LOSE - ลูกค้าใช้เจ้าที่มีเครดิต"]
        self.sheet.create_dropdown("all", self.columns.index("สถานะ"), values=status_opts, state="readonly")
        self.sheet.create_dropdown("all", self.columns.index("Select"), values=["", "✔", "เทียบ", "เทียบเพื่อชุบ"], state="readonly")
        self.sheet.create_dropdown("all", self.columns.index("ชื่อ Supplier"), values=[""] + self.supplier_list, state="normal")
        self.sheet.create_dropdown("all", self.columns.index("รายการสินค้า"), values=[""] + self.product_list, state="normal")

        # 2. จัดการสีและช่อง Auto
        auto_cols_names = [
            "Product SKU.", "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", 
            "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2",
            "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย/เส้น",
            "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)", "Markup/กก.",
            "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup", 
            "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น",
            "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat.",
            "ชื่อ Supplier2", "Sup ID."
        ]
        
        header_styles_map = {
            ("#2563EB", "white"): ["วันที่ขอราคา", "Order No.", "Sale Order No.", "QT"], 
            ("#BAE6FD", "black"): [ 
                "หมายเหตุ (ความยาว, OD)", "หมายเหตุ", "จำนวน", "ต้นทุน/เส้น", "น้ำหนัก/เส้น", "น้ำหนักรวม (Kg.)",
                "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (บาท)", "ส่วนลด 1 (%)", "ส่วนลด 2 (บาท)", "ส่วนลด 2 (%)", 
                "ทุน/เส้น หลังส่วนลด 1", "ทุน/เส้น หลังส่วนลด 2", "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", 
                "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย (ซื้อ)", "ค่าย้าย/เส้น", "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)"
            ],
            ("#FDBA74", "black"): ["ผู้ขอราคา", "ชื่อ Sale", "PRIORITY", "WIN RATE %", "Select", "แบรนด์"],
            ("#6B7280", "white"): ["หมวด", "หมวดหลัก", "หมวดรอง", "หมวดย่อย", "Product SKU.", "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup", "ค่าส่ง (ขาย)", "ค่าส่ง / เส้น"],
            ("#D8B4FE", "black"): ["รายการสินค้า", "ชื่อ Supplier"], 
            ("#FCA5A5", "black"): ["Markup Guide (%)"],
            ("#FDE047", "black"): ["สถานะ", "น้ำหนัก/เส้น 2", "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม + Vat."], 
            ("#86EFAC", "black"): ["ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น", "ราคาขาย รวม", "Vat. รวม"], 
            ("#1F2937", "white"): ["ชื่อ Supplier2", "Sup ID.", "คลังสินค้า ต้นทาง"],
            ("#93C5FD", "black"): ["ปลายทาง", "หมายเหตุ2"]
        }
        
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles_map.items() for c in cols}
        auto_cols_indices = [self.columns.index(c) for c in auto_cols_names if c in self.columns]
        
        self.sheet.readonly_columns(columns=auto_cols_indices, readonly=True)
        
        for i, col in enumerate(self.columns):
            h_bg, h_fg = col_to_style.get(col, ("#E5E7EB", "#111827")) 
            try:
                self.sheet.highlight_cells(row=0, column=i, bg=h_bg, fg=h_fg, canvas="header")
            except Exception: pass
            
            if col in auto_cols_names:
                self.sheet.highlight_columns(columns=[i], bg="#D1D5DB", fg="#111827")

    def _on_sheet_modified(self, event=None):
        try:
            for row_idx in range(self.sheet.get_total_rows()):
                self._auto_calculate_sheet(row_idx)
            self.sheet.redraw()
        except Exception:
            pass

    def _auto_calculate_sheet(self, row_idx):
        def get_val(col_name):
            try:
                val = self.sheet.get_cell_data(row_idx, self.columns.index(col_name))
                return float(str(val).replace(',', '').replace('%', '')) if val else 0.0
            except (ValueError, IndexError): return 0.0
            
        def get_str(col_name):
            try: return str(self.sheet.get_cell_data(row_idx, self.columns.index(col_name)) or "").strip()
            except IndexError: return ""

        def set_val(col_name, val, is_text=False):
            try:
                col_idx = self.columns.index(col_name)
                if not is_text:
                    if val == 0: formatted_val = ""
                    else: formatted_val = f"{val:,.2f}"
                else:
                    formatted_val = val if val and val != "%" else ""
                self.sheet.set_cell_data(row_idx, col_idx, formatted_val, redraw=False)
            except IndexError: pass

        # 1. วันที่ขอราคา
        date_idx = self.columns.index("วันที่ขอราคา")
        row_data = self.sheet.get_row_data(row_idx)
        is_row_active = False
        for i, cell_val in enumerate(row_data):
            if i != date_idx and str(cell_val).strip():
                is_row_active = True
                break
                
        current_date = get_str("วันที่ขอราคา")
        if is_row_active and not current_date:
            now = datetime.now()
            thai_year = (now.year + 543) % 100 
            set_val("วันที่ขอราคา", f"{now.day:02d}/{now.month:02d}/{thai_year}", is_text=True)
        elif not is_row_active and current_date:
            set_val("วันที่ขอราคา", "", is_text=True)

        # 2. Logic ดึง SKU และ Sync Supplier
        product_name = get_str("รายการสินค้า")
        if product_name in self.product_sku_map:
            if get_str("Product SKU.") != self.product_sku_map[product_name]:
                set_val("Product SKU.", self.product_sku_map[product_name], is_text=True)
        elif not product_name:
            set_val("Product SKU.", "", is_text=True)

        supplier_name = get_str("ชื่อ Supplier")
        if supplier_name:
            if get_str("ชื่อ Supplier2") != supplier_name:
                set_val("ชื่อ Supplier2", supplier_name, is_text=True)
            sup_id = self.supplier_code_map.get(supplier_name, "")
            if get_str("Sup ID.") != sup_id:
                set_val("Sup ID.", sup_id, is_text=True)
        else:
            if get_str("ชื่อ Supplier2") != "": set_val("ชื่อ Supplier2", "", is_text=True)
            if get_str("Sup ID.") != "": set_val("Sup ID.", "", is_text=True)

        # 3. ใส่ % อัตโนมัติ
        for col_percent in ["WIN RATE %", "Markup Guide (%)"]:
            val_str = get_str(col_percent)
            if val_str and not val_str.endswith("%"):
                num_str = "".join([c for c in val_str if c.isdigit() or c == '.'])
                if num_str: set_val(col_percent, f"{num_str}%", is_text=True)

        # 4. คำนวณสูตร
        qty = get_val("จำนวน")
        weight_per_unit = get_val("น้ำหนัก/เส้น")
        cost_per_unit = get_val("ต้นทุน/เส้น")
        
        if qty == 0 and weight_per_unit == 0 and cost_per_unit == 0:
            auto_cols_to_clear = [
                "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1", 
                "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2", "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", 
                "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย/เส้น", "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", 
                "ต้นทุนรวม (รวมย้าย)", "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", 
                "ต้นทุนรวม+Markup", "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น", 
                "Vat. / เส้น", "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat."
            ]
            for col in auto_cols_to_clear:
                set_val(col, "", is_text=True)
            return

        total_weight = weight_per_unit * qty
        set_val("น้ำหนักรวม (Kg.)", total_weight)
        
        cost_per_kg = (cost_per_unit / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ทุน/กก.", cost_per_kg)
        
        total_cost = cost_per_unit * qty
        set_val("ทุนรวม", total_cost)

        discount1_baht = get_val("ส่วนลด 1 (บาท)")
        discount1_pct = (discount1_baht / cost_per_unit) if cost_per_unit > 0 else 0
        set_val("ส่วนลด 1 (%)", f"{discount1_pct*100:.2f}%" if discount1_pct > 0 else "", is_text=True)
        
        cost_after_d1 = cost_per_unit - discount1_baht
        set_val("ทุน/เส้น หลังส่วนลด 1", cost_after_d1)

        discount2_baht = get_val("ส่วนลด 2 (บาท)")
        discount2_pct = (discount2_baht / cost_after_d1) if cost_after_d1 > 0 else 0
        set_val("ส่วนลด 2 (%)", f"{discount2_pct*100:.2f}%" if discount2_pct > 0 else "", is_text=True)
        
        cost_after_d2 = cost_after_d1 - discount2_baht
        set_val("ทุน/เส้น หลังส่วนลด 2", cost_after_d2)

        cost_no_move_per_unit = cost_after_d2
        cost_no_move_per_kg = (cost_no_move_per_unit / weight_per_unit) if weight_per_unit > 0 else 0
        cost_no_move_total = cost_no_move_per_unit * qty
        set_val("ต้นทุน/เส้น (ไม่รวมย้าย)", cost_no_move_per_unit)
        set_val("ต้นทุน/กก. (ไม่รวมย้าย)", cost_no_move_per_kg)
        set_val("ต้นทุนรวม (ไม่รวมย้าย)", cost_no_move_total)

        moving_cost = get_val("ค่าย้าย (ซื้อ)")
        moving_cost_per_unit = (moving_cost / qty) if qty > 0 else 0
        set_val("ค่าย้าย/เส้น", moving_cost_per_unit)

        cost_with_move_per_unit = cost_after_d2 + moving_cost_per_unit
        cost_with_move_per_kg = (cost_with_move_per_unit / weight_per_unit) if weight_per_unit > 0 else 0
        cost_with_move_total = cost_with_move_per_unit * qty
        set_val("ต้นทุน/เส้น (รวมย้าย)", cost_with_move_per_unit)
        set_val("ต้นทุน/กก. (รวมย้าย)", cost_with_move_per_kg)
        set_val("ต้นทุนรวม (รวมย้าย)", cost_with_move_total)

        markup_pct = get_val("Markup Guide (%)") / 100.0
        markup_per_kg = cost_with_move_per_kg * markup_pct
        markup_per_unit = cost_with_move_per_unit * markup_pct
        set_val("Markup/กก.", markup_per_kg)
        set_val("Markup/เส้น", markup_per_unit)
        
        cost_markup_per_kg = cost_with_move_per_kg + markup_per_kg
        cost_markup_per_unit = cost_with_move_per_unit + markup_per_unit
        cost_markup_total = cost_markup_per_unit * qty
        set_val("ทุน+Markup/กก.", cost_markup_per_kg)
        set_val("ทุน+Markup/เส้น", cost_markup_per_unit)
        set_val("ต้นทุนรวม+Markup", cost_markup_total)

        shipping_sell = get_val("ค่าส่ง (ขาย)")
        shipping_sell_per_unit = (shipping_sell / qty) if qty > 0 else 0
        set_val("ค่าส่ง / เส้น", shipping_sell_per_unit)
        set_val("น้ำหนัก/เส้น 2", weight_per_unit)

        sell_price_per_unit = cost_markup_per_unit + shipping_sell_per_unit
        sell_price_per_kg = (sell_price_per_unit / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ราคาขาย / เส้น", sell_price_per_unit)
        set_val("ราคาขาย / กก.", sell_price_per_kg)

        vat_per_unit = sell_price_per_unit * 0.07
        sell_price_vat_unit = sell_price_per_unit * 1.07
        set_val("Vat. / เส้น", vat_per_unit)
        set_val("ราคาขาย/เส้น + Vat.", sell_price_vat_unit)

        sell_price_total = sell_price_per_unit * qty
        vat_total = sell_price_total * 0.07
        sell_price_total_vat = sell_price_total * 1.07
        set_val("ราคาขาย รวม", sell_price_total)
        set_val("Vat. รวม", vat_total)
        set_val("ราคาขาย รวม + Vat.", sell_price_total_vat)

    # ------------------------------------------------------------------ #
    def _add_new_row(self):
        if HAS_TKSHEET: self.sheet.insert_row([""] * len(self.columns))

    def _delete_selected_rows(self):
        if not HAS_TKSHEET: return
        selected_rows = self.sheet.get_selected_rows()
        if not selected_rows:
            selected_cells = self.sheet.get_selected_cells()
            if selected_cells:
                selected_rows = list(set(r for r, c in selected_cells))
        
        if not selected_rows:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการลบก่อน", parent=self)
            return

        if messagebox.askyesno("ยืนยัน", f"ต้องการลบข้อมูลจำนวน {len(selected_rows)} บรรทัด ใช่หรือไม่?", parent=self):
            self.sheet.delete_rows(list(selected_rows))
            self.sheet.redraw()

    # 🟢 [จุดสำคัญ] แก้ไขฟังก์ชัน Load ให้หยอดข้อมูลทีละช่อง ป้องกัน Dropdown หาย
    def _load_from_db(self, *args):
        if not HAS_TKSHEET: return
        month_val = self.month_var.get()
        year_val = self.year_var.get()

        conn = self.app_container.get_connection()
        try:
            columns_sql = ", ".join([f'"{col.replace("%", "%%")}"' for col in self.columns])
            
            query = f"SELECT {columns_sql} FROM cost_benchmarks WHERE benchmark_month = %s AND benchmark_year = %s AND created_by = %s ORDER BY id ASC"
            df = pd.read_sql_query(query, conn, params=(month_val, year_val, self.current_user))
            
            self.sheet.extra_bindings("unbind", "<<SheetModified>>")
            total_cols = len(self.columns)
            
            if df.empty:
                data_list = [[""] * total_cols for _ in range(20)]
            else:
                df = df.fillna("")
                data_list = df.values.tolist()
                while len(data_list) < 20:
                    data_list.append([""] * total_cols)

            current_rows = self.sheet.get_total_rows()
            if len(data_list) > current_rows:
                self.sheet.insert_rows(len(data_list) - current_rows)

            # 🟢 หยอดข้อมูลทีละช่อง เพื่อไม่ให้การตั้งค่า Dropdown เดิมโดนล้าง
            for r, row in enumerate(data_list):
                for c, val in enumerate(row):
                    val_str = str(val).strip() if val is not None else ""
                    self.sheet.set_cell_data(r, c, val_str, redraw=False)

            # 🟢 เคลียร์แถวที่เหลือให้ว่าง
            for r in range(len(data_list), self.sheet.get_total_rows()):
                for c in range(total_cols):
                    self.sheet.set_cell_data(r, c, "", redraw=False)

            self.sheet.redraw()
            
        except Exception as e:
            messagebox.showerror("Error", f"โหลดข้อมูลล้มเหลว: {e}", parent=self)
        finally:
            self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
            if conn: self.app_container.release_connection(conn)

    def _save_to_db(self):
        if not HAS_TKSHEET: return
        
        try:
            self.sheet.deselect("all")
        except:
            pass

        raw = self.sheet.get_sheet_data()
        data = []
        
        for row in raw:
            is_active_row = False
            for cell in row:
                cell_text = str(cell).strip()
                if cell_text != "" and cell_text != "%":
                    is_active_row = True
                    break 
            
            if is_active_row:
                clean_row = [str(cell).strip() if cell is not None else "" for cell in row]
                data.append(clean_row)
                
        df = pd.DataFrame(data, columns=self.columns)
        
        if df.empty:
            messagebox.showinfo("แจ้งเตือน", "ไม่มีข้อมูลให้บันทึก", parent=self)
            return

        month_val = self.month_var.get()
        year_val = self.year_var.get()

        df = df.replace(r'^\s*$', None, regex=True)
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("DELETE FROM cost_benchmarks WHERE benchmark_month = %s AND benchmark_year = %s AND created_by = %s", (month_val, year_val, self.current_user))
                
                columns_sql = ", ".join([f'"{col.replace("%", "%%")}"' for col in self.columns]) + ", benchmark_month, benchmark_year, created_by"
                values = [tuple(row) + (month_val, year_val, self.current_user) for row in df.to_numpy()]
                
                insert_query = f"INSERT INTO cost_benchmarks ({columns_sql}) VALUES %s"
                psycopg2.extras.execute_values(cursor, insert_query, values)
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล {len(df)} รายการของคุณ {self.current_user} เรียบร้อยแล้ว!", parent=self)
            
        except Exception as e:
            if conn: conn.rollback()
            import traceback; traceback.print_exc()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)