import tkinter as tk
from tkinter import messagebox
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkButton, CTkOptionMenu, CTkEntry
import pandas as pd
import psycopg2.extras
from datetime import datetime
import re
import json
from tkinter import colorchooser
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

        self.auto_save_job_id = None # 🟢  ตัวแปรจำเวลา Auto-save

        # 🟢 ดึงชื่อ User ที่ Login อยู่ปัจจุบัน
        self.current_user = getattr(self.app_container, 'current_user_key', 'PU_Default')

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        
        self.frozen_col_count = 0
        self.zoom_level = 11 
        self.col_widths_cache = {}
        self.sales_list = []
        self.supplier_list = []
        self.product_list = []
        self.product_sku_map = {} 
        self.supplier_code_map = {} 
        self.product_category_map = {}
        self.hidden_cols_list = []         # จำว่าซ่อนคอลัมน์ไหนไว้
        self.custom_header_colors = {}

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

        # ==========================================
        # 1. สร้าง btn_frame ก่อน
        btn_frame = CTkFrame(header_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=6, sticky="e")

        # 2. ค่อยเอาปุ่มต่างๆ ไปใส่ใน btn_frame (ย้ายปุ่มสีมาไว้ตรงนี้)
        CTkButton(btn_frame, text="🎨 เปลี่ยนสีหัวคอลัมน์", fg_color="#EC4899", hover_color="#DB2777",
                  command=self._change_header_color).pack(side="left", padx=5)

        CTkButton(btn_frame, text="📌 ตรึงคอลัมน์", fg_color="#0891B2", hover_color="#0E7490",
          command=self._freeze_selected_columns).pack(side="left", padx=5)

        CTkButton(btn_frame, text="📌 ยกเลิกตรึง", fg_color="#64748B", hover_color="#475569",
                command=self._unfreeze_columns).pack(side="left", padx=5)
                  
        CTkButton(btn_frame, text="🗑️ ลบบรรทัด", fg_color="#EF4444", hover_color="#DC2626",
                  command=self._delete_selected_rows).pack(side="left", padx=5)

        CTkButton(btn_frame, text="🙈 ซ่อนคอลัมน์", fg_color="#F59E0B", hover_color="#D97706",
                  command=self._hide_selected_columns).pack(side="left", padx=5)
                  
        CTkButton(btn_frame, text="👁️ แสดงคอลัมน์", fg_color="#8B5CF6", hover_color="#7C3AED",
                  command=self._show_all_columns).pack(side="left", padx=5)
                  
        CTkButton(btn_frame, text="➕ เพิ่มบรรทัดใหม่",
                  command=self._add_new_row).pack(side="left", padx=5)

        self.columns = [
            "วันที่ขอราคา","Order No.", "Sale Order No.", "รหัส Sale",
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
        self.current_item_label.pack(pady=8, padx=15, anchor="w") # กลับไปใช้แบบเดิม
        # ================================================================== #

        self.target_formula_cell = None
        
        self.formula_frame = CTkFrame(self, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#CBD5E1")
        self.formula_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10)) 
        
        CTkLabel(self.formula_frame, text=" 𝑓x ", font=CTkFont(family="Arial", size=18, weight="bold", slant="italic"), text_color="#16A34A").pack(side="left", padx=10, pady=5)
        
        self.formula_entry = CTkEntry(self.formula_frame, font=CTkFont(size=14), placeholder_text="คลิกช่องปลายทาง -> คลิกที่นี่ -> พิมพ์ = แล้วใช้เมาส์จิ้มเซลล์ในตารางได้เลย!", border_width=0, fg_color="transparent")
        self.formula_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=5)
        
        self.formula_entry.bind("<FocusIn>", self._on_formula_focus_in)
        self.formula_entry.bind("<Return>", self._apply_formula_from_bar)

        # --- 3. ตาราง ---
        table_frame = tk.Frame(self, bg="white")
        table_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 5))  # ลด pady ด้านล่างลง
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        self.grid_rowconfigure(2, weight=0) 
        self.grid_rowconfigure(3, weight=1) # ให้ตารางขยายได้

        # ================================================================== #
        # 🟢 [เพิ่มใหม่] แถบ Status Bar ด้านล่างสุด (Excel Style)
        # ================================================================== #
        self.bottom_status_frame = CTkFrame(self, fg_color="#E5E7EB", corner_radius=4)
        self.bottom_status_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 10))

        zoom_frame = CTkFrame(self.bottom_status_frame, fg_color="transparent")
        zoom_frame.pack(side="right", padx=5)

        CTkButton(zoom_frame, text="🔍−", width=32, height=24,
                fg_color="#6B7280", hover_color="#4B5563", font=CTkFont(size=12),
                command=lambda: self._zoom(-1)).pack(side="left", padx=2)

        self.zoom_label = CTkLabel(zoom_frame, text="100%",
                font=CTkFont(size=12), text_color="#6B7280", width=40)
        self.zoom_label.pack(side="left")

        CTkButton(zoom_frame, text="🔍+", width=32, height=24,
                fg_color="#6B7280", hover_color="#4B5563", font=CTkFont(size=12),
                command=lambda: self._zoom(1)).pack(side="left", padx=2)
        
        # 🟢 [เพิ่มใหม่] ตัวหนังสือบอกสถานะการเซฟ (อยู่มุมซ้าย)
        self.save_status_label = CTkLabel(
            self.bottom_status_frame,
            text="✅ พร้อมใช้งาน (บันทึกอัตโนมัติ)", 
            font=CTkFont(size=13),
            text_color="gray50"
        )
        self.save_status_label.pack(side="left", padx=20, pady=4)

        self.quick_calc_label = CTkLabel(
            self.bottom_status_frame,
            text="", 
            font=CTkFont(size=14, weight="bold"),
            text_color="#059669" # สีเขียวเข้ม
        )
        self.quick_calc_label.pack(side="right", padx=20, pady=4)

        if HAS_TKSHEET:
            self._build_tksheet(table_frame)
            self.after(200, self._load_from_db)
        else:
            tk.Label(table_frame, text="⚠️ กรุณาติดตั้ง tksheet", fg="red", bg="white").pack(expand=True)

    def _lighten_color(self, hex_color, amount=0.85):
        """ฟังก์ชันแปลงสีที่ User เลือก ให้กลายเป็นสีพาสเทลอ่อนๆ"""
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            
            # ผสมกับสีขาว (255) ตาม % ที่กำหนด
            r = int(r + (255 - r) * amount)
            g = int(g + (255 - g) * amount)
            b = int(b + (255 - b) * amount)
            
            return f'#{r:02x}{g:02x}{b:02x}'
        except:
            return "#F3F4F6"

    def _load_user_settings(self):
        """ดึงความจำจาก Database ตอนเปิดหน้าจอ"""
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT setting_value FROM user_settings WHERE user_name = %s AND setting_key = 'benchmark_table'", (self.current_user,))
                result = cursor.fetchone()
                if result and result[0]:
                    settings = result[0]
                    self.hidden_cols_list = settings.get("hidden_cols", [])
                    self.custom_header_colors = settings.get("header_colors", {}) # 🟢 ดึงสีที่จำไว้
        except Exception as e:
            print(f"Error loading settings: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _freeze_selected_columns(self):
        if not HAS_TKSHEET: return

        real_cols = self._get_real_col_indices()
        if not real_cols:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกคอลัมน์ก่อน", parent=self)
            return

        freeze_up_to = max(real_cols)

        try:
            self.frozen_col_count = freeze_up_to + 1

            frozen = list(range(self.frozen_col_count))
            rest = [c for c in range(len(self.columns)) if c not in frozen]

            self.sheet.display_columns(frozen + rest)
            self.sheet.redraw()

            self.save_status_label.configure(
                text=f"📌 ตรึงถึง: {self.columns[freeze_up_to]} (Fake Freeze)",
                text_color="#0891B2"
            )

        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _unfreeze_columns(self):
        if not HAS_TKSHEET: return
        try:
            self.frozen_col_count = 0
            self.sheet.display_columns("all")
            self.sheet.redraw()

            self.save_status_label.configure(
                text="✅ ยกเลิกตรึงแล้ว",
                text_color="#16A34A"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _save_col_widths(self):
        """จำ width ของทุกคอลัมน์ไว้ใน cache"""
        try:
            for i in range(len(self.columns)):
                w = self.sheet.column_width(i)
                if w and w > 0:
                    self.col_widths_cache[i] = w
        except Exception:
            pass

    def _save_user_settings(self):
        """บันทึกความจำลง Database"""
        settings = {
            "hidden_cols": self.hidden_cols_list,
            "header_colors": self.custom_header_colors # 🟢 เซฟสีที่เลือกไว้
        }
        settings_json = json.dumps(settings)
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO user_settings (user_name, setting_key, setting_value) 
                    VALUES (%s, 'benchmark_table', %s)
                    ON CONFLICT (user_name, setting_key) 
                    DO UPDATE SET setting_value = EXCLUDED.setting_value
                """, (self.current_user, settings_json))
            conn.commit()
        except Exception as e:
            print(f"Error saving settings: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    

    def _load_dropdown_data(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # 🟢 1. [แก้ไขตรงนี้] เปลี่ยนจากดึง sale_name เป็นดึง sale_key (รหัส ID)
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Sale' AND status = 'Active'")
                self.sales_list = [row[0] for row in cursor.fetchall() if row[0]]

                cursor.execute("SELECT supplier_name, supplier_code FROM suppliers")
                for row in cursor.fetchall():
                    if row[0]:
                        self.supplier_list.append(row[0])
                        self.supplier_code_map[row[0]] = row[1] or ""

                cursor.execute("SELECT product_name, product_code, category FROM products")
                for row in cursor.fetchall():
                    if row[0]:
                        self.product_list.append(row[0])
                        self.product_sku_map[row[0]] = row[1] or ""
                        # 🟢 [เพิ่มใหม่] เก็บค่าหมวดหมู่คู่กับชื่อสินค้าไว้ใน Map
                        self.product_category_map[row[0]] = row[2] or ""
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
        self.sheet.bind("<Shift-MouseWheel>", self._lock_horizontal_scroll)
        self.sheet.bind("<Control-MouseWheel>", self._on_ctrl_scroll)
        self.after(500, self._save_col_widths)

        # 🟢 1. ผูก Event เฉพาะของ tksheet (รับแค่ 2 ค่า)
        self.sheet.extra_bindings([
            ("cell_select", self._on_sheet_click_for_formula),
            ("end_edit_cell", self._on_end_edit_combined),      
        ])
        
        # 🟢 2. ผูก Event ของคีย์บอร์ดแยกออกมาต่างหาก (ดักจับตอนกด Enter)
        self.sheet.bind("<Return>", self._on_enter_pressed)
        self.sheet.bind("<KP_Enter>", self._on_enter_pressed)
        
        self.sheet.grid(row=0, column=0, sticky="nsew")

        self.sheet.enable_bindings((
            "single_select", "drag_select", "multi_select", # <--- เพิ่มตรงนี้
            "row_select", "column_select", "column_width_resize",
            "arrowkeys", "right_click_popup_menu",
            "rc_select", "copy", "cut", "paste",
            "delete", "undo", "edit_cell",
        ))

        self.sheet.set_options(
            grid_color="#000000", outline_color="#000000", table_bg="white", table_fg="black", 
            table_grid_fg="#000000", header_bg="#D1D5DB", header_fg="#111827", header_grid_fg="#000000",
            header_selected_cells_bg="#9CA3AF", row_index_bg="#F3F4F6", row_index_fg="#111827", 
            row_index_grid_fg="#000000", selected_cells_border_color="#3B82F6", table_selected_cells_border_color="#3B82F6",
            auto_resize_columns=False,      
            auto_resize_row_index=False,
            dropdown_font=("Tahoma", 11, "normal"), 
        )

        for i, col in enumerate(self.columns):
            if "รายการสินค้า" in col or "หมายเหตุ" in col:
                self.sheet.column_width(i, 250)

        # 🟢 1. โหลดความจำของ User จาก Database ก่อน
        self._load_user_settings()

        # 🟢 2. เรียกสร้าง Dropdown และทาสีตาราง (ทำหลังจากโหลดความจำแล้ว)
        self._apply_formatting()

        # 🟢 3. ซ่อนคอลัมน์ตามที่จำไว้
        if self.hidden_cols_list:
            self.sheet.hide_columns(self.hidden_cols_list)
        
        self.sheet.bind("<ButtonRelease-1>", self._trigger_banner_update)
        self.sheet.bind("<KeyRelease>", self._trigger_banner_update)
        self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)

    def _on_end_edit_combined(self, event=None):
        """รวม end_edit: เลื่อนขวา + lock width"""
        try:
            if event and len(event) >= 2:
                row, col = event[0], event[1]
                self.after(10, lambda: self._move_right(row, col))
        except Exception:
            pass
        # lock width
        if self.col_widths_cache:
            self.after(50, self._restore_col_widths)

    def _lock_horizontal_scroll(self, event):
        """Fake freeze: ล็อคไม่ให้ scroll แนวนอนเมื่อกด Shift"""
        return "break"

    def _lock_column_widths(self, event=None):
        """Restore width กลับหลังจาก dropdown ขยาย"""
        if not self.col_widths_cache:
            return
        try:
            self.after(50, self._restore_col_widths)
        except Exception:
            pass

    def _restore_col_widths(self):
        """คืน width ที่จำไว้ทุกคอลัมน์"""
        try:
            for i, w in self.col_widths_cache.items():
                self.sheet.column_width(i, w)
            self.sheet.redraw()
        except Exception:
            pass

    def _on_ctrl_scroll(self, event):
        """Ctrl+Scroll wheel เพื่อ Zoom"""
        if event.delta > 0:
            self._zoom(1)
        else:
            self._zoom(-1)
        return "break"  # หยุดไม่ให้ scroll ปกติทำงาน

    def _move_right(self, row, col):
        total_cols = self.sheet.get_total_columns()
        total_rows = self.sheet.get_total_rows()
        
        next_col = col + 1
        
        if next_col < total_cols:
            self.sheet.deselect("all")
            self.sheet.select_cell(row, next_col)
            self.sheet.see(row, next_col) # เลื่อนหน้าจอตามไปให้เห็นช่อง
        else:
            # ถ้าถึงช่องขวาสุดแล้ว ให้ปัดลงมาบรรทัดใหม่ เริ่มที่ช่องซ้ายสุด
            if row + 1 < total_rows:
                self.sheet.deselect("all")
                self.sheet.select_cell(row + 1, 0)
                self.sheet.see(row + 1, 0)

    def _on_enter_pressed(self, event=None):
        try:
            curr = self.sheet.get_currently_selected()
            if curr:
                self._move_right(curr[0], curr[1])
            return "break" # สำคัญมาก: ส่งคำสั่งหยุด เพื่อไม่ให้มันเลื่อนลงตามค่า Default
        except Exception:
            pass

    def _on_end_edit(self, event=None):
        try:
            # event ของ tksheet จะส่งมาเป็น (row, col, "end_edit_cell", new_value)
            if event and len(event) >= 2:
                row, col = event[0], event[1]
                # ใช้ .after() หน่วงเวลา 10ms เพื่อทับคำสั่งเลื่อนลงปกติของ tksheet (หลังจากที่มันเซฟค่าเสร็จ)
                self.after(10, lambda: self._move_right(row, col))
        except Exception:
            pass

    def _trigger_banner_update(self, event=None):
        self.after(50, self._update_current_item_banner)

    def _update_current_item_banner(self):
        try:
            cells = self.sheet.get_selected_cells()
            if not cells:
                return
            
            row = list(cells)[0][0]
            
            # --- ส่วนที่ 1: อัปเดตข้อความฝั่งซ้าย (อ้างอิงบรรทัดแรกที่เลือก) ---
            order_no = str(self.sheet.get_cell_data(row, self.columns.index("Order No.")) or "-").strip()
            supplier = str(self.sheet.get_cell_data(row, self.columns.index("ชื่อ Supplier")) or "-").strip()
            product = str(self.sheet.get_cell_data(row, self.columns.index("รายการสินค้า")) or "-").strip()
            remark = str(self.sheet.get_cell_data(row, self.columns.index("หมายเหตุ (ความยาว, OD)")) or "-").strip()
            
            if order_no == "-" and supplier == "-" and product == "-":
                self.current_item_label.configure(text=f"📌 กำลังแก้ไข บรรทัดที่ {row+1} (ช่องว่าง)")
            else:
                self.current_item_label.configure(
                    text=f"📌 กำลังแก้ไข บรรทัดที่ {row+1}  👉  Order: {order_no}  |  Supplier: {supplier}  |  รายการสินค้า: {product}  |  หมายเหตุ: {remark}"
                )

            # --- 🟢 ส่วนที่ 2: [เพิ่มใหม่] ระบบคำนวณผลรวม (Excel-like Quick Calc) ---
            if len(cells) > 1: # ถ้าคลุมมากกว่า 1 ช่อง
                total_sum = 0.0
                num_count = 0
                
                for r, c in cells:
                    val = self.sheet.get_cell_data(r, c)
                    if val is not None and str(val).strip() != "":
                        # ทำความสะอาดตัวเลข (เอาลูกน้ำกับ % ออก)
                        clean_val = str(val).replace(',', '').replace('%', '').strip()
                        try:
                            # ลองแปลงเป็นทศนิยม
                            num = float(clean_val)
                            total_sum += num
                            num_count += 1
                        except ValueError:
                            pass # ถ้าเป็นตัวหนังสือ (แปลงไม่ได้) ให้ข้ามไป ไม่ต้องสนใจ
                
                # ถ้าเจอตัวเลขอย่างน้อย 1 ช่องในที่ลากคลุม
                if num_count > 0:
                    avg = total_sum / num_count
                    self.quick_calc_label.configure(
                        text=f"ผลรวม: {total_sum:,.2f}    |    ค่าเฉลี่ย: {avg:,.2f}   "
                    )
                else:
                    # ถ้าลากคลุมแต่เจอแต่ตัวหนังสือ (ไม่มีตัวเลขเลย)
                    self.quick_calc_label.configure(text=f"จำนวนเซลล์ที่เลือก: {len(cells)}")
            else:
                # ถ้าคลิกแค่ช่องเดียว ไม่ต้องโชว์ผลรวม
                self.quick_calc_label.configure(text="")

        except Exception:
            pass

    def _apply_formatting(self):
        # 1. จัดการ Dropdown
        self.sheet.create_dropdown("all", self.columns.index("รหัส Sale"), values=[""] + self.sales_list, state="normal") # 🟢 เปลี่ยนตรงนี้
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
            ("#FDBA74", "black"): ["ผู้ขอราคา", "รหัส Sale", "PRIORITY", "WIN RATE %", "Select", "แบรนด์"],
            ("#6B7280", "white"): ["หมวด", "หมวดหลัก", "หมวดรอง", "หมวดย่อย", "Product SKU.", "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup", "ค่าส่ง (ขาย)", "ค่าส่ง / เส้น"],
            ("#D8B4FE", "black"): ["รายการสินค้า", "ชื่อ Supplier"], 
            ("#FCA5A5", "black"): ["Markup Guide (%)"],
            ("#FDE047", "black"): ["สถานะ", "น้ำหนัก/เส้น 2", "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม + Vat."], 
            ("#86EFAC", "black"): ["ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น", "ราคาขาย รวม", "Vat. รวม"], 
            ("#1F2937", "white"): ["ชื่อ Supplier2", "Sup ID.", "คลังสินค้า ต้นทาง"],
            ("#93C5FD", "black"): ["ปลายทาง", "หมายเหตุ2"]
        }
        
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles_map.items() for c in cols}
        
        # เอาสีที่ User เลือกมาทับ (ถ้ามี)
        if hasattr(self, 'custom_header_colors'):
            for col_name, bg_color in self.custom_header_colors.items():
                if col_name in self.columns:
                    col_to_style[col_name] = (bg_color, "black") 

        # ล็อคช่องที่คำนวณออโต้
        auto_cols_indices = [self.columns.index(c) for c in auto_cols_names if c in self.columns]
        self.sheet.readonly_columns(columns=auto_cols_indices, readonly=True)
        
        # 3. เริ่มกระบวนการระบายสี 🎨 (แบบวนลูปทีละคอลัมน์)
        total_rows = self.sheet.get_total_rows()
        for i, col in enumerate(self.columns):
            h_bg, h_fg = col_to_style.get(col, ("#E5E7EB", "#111827")) 
            
            if hasattr(self, 'custom_header_colors') and col in self.custom_header_colors:
                b_bg = self._lighten_color(self.custom_header_colors[col], amount=0.85)
                b_fg = "black"
            elif col in auto_cols_names:
                b_bg = "#F3F4F6"
                b_fg = "#111827"
            else:
                b_bg = "white"
                b_fg = "black"

            try:
                # 🟢 ทาสีหัวตาราง
                self.sheet.highlight_cells(row=0, column=i, bg=h_bg, fg=h_fg, canvas="header")
                
                # 🟢 แก้ใหม่: ใช้ highlight_columns แทน highlight_cells สำหรับตาราง
                self.sheet.highlight_columns(
                    columns=[i],
                    bg=b_bg,
                    fg=b_fg,
                    highlight_header=False  # ไม่แตะหัว เพราะทำแยกบรรทัดบนแล้ว
                )
            except Exception:
                # fallback กรณี tksheet เวอร์ชันเก่าไม่รองรับ highlight_header parameter
                try:
                    self.sheet.highlight_cells(row=0, column=i, bg=h_bg, fg=h_fg, canvas="header")
                    for r in range(total_rows):
                        self.sheet.highlight_cells(row=r, column=i, bg=b_bg, fg=b_fg, canvas="table")
                except Exception:
                    pass

    def _on_sheet_modified(self, event=None):
        try:
            # 1. คำนวณสูตรออโต้เหมือนเดิม
            for row_idx in range(self.sheet.get_total_rows()):
                self._auto_calculate_sheet(row_idx)
            self.sheet.redraw()

            # 🟢 2. ระบบ Auto Save (หน่วงเวลา 1.5 วินาที หลังหยุดพิมพ์)
            # ถ้ามีการพิมพ์ใหม่ ให้ยกเลิกคิวการเซฟอันเก่าทิ้งไปก่อน (Debounce)
            if self.auto_save_job_id is not None:
                self.after_cancel(self.auto_save_job_id)

            # เปลี่ยนข้อความให้รู้ว่ากำลังรอจังหวะเซฟ
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

            # ตั้งเวลา: ถ้าไม่พิมพ์อะไรเพิ่มเติมใน 1500 ms (1.5 วิ) ให้เรียกฟังก์ชัน Save แบบไม่โชว์ Popup
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))

        except Exception:
            pass

    def _auto_calculate_sheet(self, row_idx):
        # =================================================================
        # 🟢 ระบบคำนวณสูตร Excel ขั้นสูง (รองรับการอ้างอิง A1, B2)
        # =================================================================
        def col2num(col_str):
            expn = 0
            col_num = 0
            for char in reversed(col_str.upper()):
                col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
                expn += 1
            return col_num - 1

        try:
            row_data = self.sheet.get_row_data(row_idx)
            for c_idx, cell_val in enumerate(row_data):
                val_str = str(cell_val).strip()
                if val_str.startswith('=') and len(val_str) > 1:
                    try:
                        expr = val_str[1:].replace(',', '').upper()
                        # ค้นหาคำที่เป็น A1, B2 ในสมการ
                        cell_refs = set(re.findall(r'[A-Z]+\d+', expr))
                        
                        for ref in cell_refs:
                            match = re.match(r'([A-Z]+)(\d+)', ref)
                            if match:
                                c_str, r_str = match.groups()
                                target_col = col2num(c_str)
                                target_row = int(r_str) - 1
                                
                                ref_val = self.sheet.get_cell_data(target_row, target_col)
                                if not ref_val or str(ref_val).strip() == "":
                                    ref_val = "0"
                                else:
                                    ref_val = str(ref_val).replace(',', '').replace('%', '')
                                    
                                expr = re.sub(rf'\b{ref}\b', str(ref_val), expr)
                        
                        result = eval(expr, {"__builtins__": None}, {})
                        if isinstance(result, (int, float)):
                            self.sheet.set_cell_data(row_idx, c_idx, f"{float(result):.2f}", redraw=False)
                    except Exception:
                        pass
        except Exception:
            pass
        # =================================================================

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
            # อัปเดต SKU
            if get_str("Product SKU.") != self.product_sku_map[product_name]:
                set_val("Product SKU.", self.product_sku_map[product_name], is_text=True)
            
            # 🟢 3. [เพิ่มใหม่] อัปเดต หมวดหมู่ อัตโนมัติ
            mapped_category = self.product_category_map.get(product_name, "")
            if get_str("หมวด") != mapped_category:
                set_val("หมวด", mapped_category, is_text=True)
                
        elif not product_name:
            # ถ้าผู้ใช้ลบชื่อสินค้าออก ให้เคลียร์ SKU และ หมวด ให้เป็นช่องว่างด้วย
            if get_str("Product SKU.") != "": set_val("Product SKU.", "", is_text=True)
            if get_str("หมวด") != "": set_val("หมวด", "", is_text=True)

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

        # 4. คำนวณสูตรหลักของระบบ
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
            self.after(300, self._save_col_widths)
            if conn: self.app_container.release_connection(conn)

    def _save_to_db(self, show_msg=True):
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
            if show_msg: messagebox.showinfo("แจ้งเตือน", "ไม่มีข้อมูลให้บันทึก", parent=self)
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
            
            # 🟢 [เพิ่มใหม่] อัปเดตข้อความมุมซ้ายล่างว่าเซฟสำเร็จตอนกี่โมง
            current_time = datetime.now().strftime("%H:%M:%S")
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text=f"✅ บันทึกล่าสุด: {current_time}", text_color="#16A34A")

            # โชว์ Popup เฉพาะตอนที่ User สั่งกดปุ่ม (ซึ่งตอนนี้ลบไปแล้ว แต่เผื่อไว้)
            if show_msg: 
                messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล {len(df)} รายการของคุณ {self.current_user} เรียบร้อยแล้ว!", parent=self)
            
        except Exception as e:
            if conn: conn.rollback()
            import traceback; traceback.print_exc()
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="❌ บันทึกผิดพลาด กรุณาลองใหม่", text_color="#DC2626")
            if show_msg:
                messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    # =================================================================
    # 🟢 ระบบสมองกลของ FORMULA BAR (ทำงานร่วมกับการคลิกเมาส์)
    # =================================================================
    def _num2col(self, n):
        """แปลงตัวเลขคอลัมน์เป็นตัวอักษร (0 -> A, 1 -> B, ... 26 -> AA)"""
        string = ""
        n += 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def _on_formula_focus_in(self, event):
        """จดจำว่าผู้ใช้เลือกช่องไหนไว้เป็นปลายทาง ก่อนที่จะมาคลิกแถบสูตร"""
        try:
            cells = self.sheet.get_selected_cells()
            if cells:
                self.target_formula_cell = list(cells)[0]
        except: pass

    def _on_sheet_click_for_formula(self, event=None):
        """เมื่อคลิกตาราง ถ้าแถบสูตรมีเครื่องหมาย = อยู่ ให้ดูดตัวเลขจากช่องมาใส่"""
        try:
            current_text = self.formula_entry.get()
            if current_text.startswith("="):
                cells = self.sheet.get_selected_cells()
                if not cells: return
                row, col = list(cells)[0]
                
                # 🟢 1. ดึงข้อมูล "ตัวเลข" จากช่องที่คลิก
                cell_val = self.sheet.get_cell_data(row, col)
                
                # 🟢 2. ทำความสะอาดตัวเลข (เอาลูกน้ำกับ % ออก เพื่อให้พร้อมคำนวณ)
                if not cell_val or str(cell_val).strip() == "":
                    val_to_insert = "0"
                else:
                    val_to_insert = str(cell_val).replace(',', '').replace('%', '').strip()
                    # ตรวจสอบว่าเป็นตัวเลขจริงๆ ไหม ถ้าไปเผลอจิ้มช่องตัวหนังสือให้ใส่เลข 0 แทน
                    try:
                        float(val_to_insert)
                    except ValueError:
                        val_to_insert = "0"
                
                # 🟢 3. นำ "ตัวเลข" ไปต่อท้ายใน Formula Bar
                self.formula_entry.insert(tk.END, val_to_insert)
                
                # 🟢 4. ดึง Focus กลับมาที่แถบสูตร เพื่อให้พิมพ์ + - * / ต่อได้เลย
                self.formula_entry.focus()
                self.formula_entry.icursor(tk.END)
        except Exception:
            pass

    def _apply_formula_from_bar(self, event=None):
        """เมื่อกด Enter ให้นำสูตรกลับไปใส่ในตาราง และคำนวณทันที"""
        if not self.target_formula_cell:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกช่องปลายทางในตารางก่อนเริ่มพิมพ์สูตร")
            return
            
        try:
            row, col = self.target_formula_cell
            formula = self.formula_entry.get()
            
            # 1. นำสูตรไปใส่ในช่องปลายทาง
            self.sheet.set_cell_data(row, col, formula)
            self.formula_entry.delete(0, tk.END)
            
            # 2. เคลียร์ความจำ และเลื่อนช่องที่เลือก (Focus) กลับไปที่ผลลัพธ์
            self.target_formula_cell = None
            self.sheet.select_cell(row, col)
            
            # 3. สั่งให้ตารางคำนวณ (ใช้ฟังก์ชัน _auto_calculate_sheet ที่เราอัปเดตไปก่อนหน้านี้)
            self._auto_calculate_sheet(row)
            
        except Exception as e:
            messagebox.showerror("Error", f"สูตรผิดพลาด: {e}")

    # ------------------------------------------------------------------ #
    # 🟢 ฟังก์ชันสำหรับซ่อน/แสดง คอลัมน์ (Columns แนวตั้ง)
    # ------------------------------------------------------------------ #
    def _get_real_col_indices(self):
        """Helper: แปลง display index → data index"""
        real_cols = set()

        # 🔧 ใน tksheet version นี้:
        # displayed_columns = [] หมายถึง "แสดงทุกคอลัมน์" (ไม่ได้ซ่อนอะไร)
        # displayed_columns = [0,2,3,...] หมายถึง "แสดงเฉพาะคอลัมน์พวกนี้"
        try:
            displayed = self.sheet.displayed_columns
            if not displayed:
                # list ว่าง = แสดงทุกคอลัมน์ = display index == data index
                displayed = list(range(len(self.columns)))
        except Exception:
            displayed = list(range(len(self.columns)))

        # วิธีที่ 1: จาก selected_cells
        try:
            selected_cells = self.sheet.get_selected_cells()
            if selected_cells:
                for r, disp_c in selected_cells:
                    if disp_c < len(displayed):
                        real_cols.add(displayed[disp_c])
        except Exception:
            pass

        # วิธีที่ 2: จาก selected_columns (คลิกหัวคอลัมน์)
        try:
            selected_cols = self.sheet.get_selected_columns()
            if selected_cols:
                for disp_c in selected_cols:
                    if disp_c < len(displayed):
                        real_cols.add(displayed[disp_c])
        except Exception:
            pass

        return real_cols
        
    def _hide_selected_columns(self):
        if not HAS_TKSHEET: return

        real_cols = self._get_real_col_indices()

        if not real_cols:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิก 'ช่องใดๆ' หรือ 'หัวคอลัมน์' ที่ต้องการซ่อนก่อน", parent=self)
            return

        self.hidden_cols_list.extend(real_cols)
        self.hidden_cols_list = list(set(self.hidden_cols_list))

        self.sheet.hide_columns(self.hidden_cols_list)
        self._save_user_settings()
        self.sheet.redraw()

    def _change_header_color(self):
        if not HAS_TKSHEET: return

        real_cols = self._get_real_col_indices()

        if not real_cols:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิก 'ช่องใดๆ' หรือ 'หัวคอลัมน์' ที่ต้องการเปลี่ยนสีก่อน", parent=self)
            return

        color_tuple = colorchooser.askcolor(title="เลือกสีสำหรับหัวคอลัมน์", parent=self)
        if not color_tuple or not color_tuple[1]:
            return

        color_code = color_tuple[1]
        for c_idx in real_cols:
            if c_idx < len(self.columns):
                self.custom_header_colors[self.columns[c_idx]] = color_code

        # ทาสีเฉพาะคอลัมน์ที่เลือก ใช้ data index ตรงๆ ไม่แตะคอลัมน์อื่น
        for c_idx in real_cols:
            if c_idx >= len(self.columns):
                continue
            b_bg = self._lighten_color(color_code, amount=0.85)
            try:
                self.sheet.highlight_cells(row=0, column=c_idx, bg=color_code, fg="black", canvas="header")
                self.sheet.highlight_columns(columns=[c_idx], bg=b_bg, fg="black", highlight_header=False)
            except Exception:
                try:
                    self.sheet.highlight_cells(row=0, column=c_idx, bg=color_code, fg="black", canvas="header")
                except Exception:
                    pass

        self._save_user_settings()
        self.sheet.redraw()

    def _zoom(self, direction):
        if not HAS_TKSHEET: return
        self.zoom_level = max(8, min(20, self.zoom_level + direction))
        self.sheet.set_options(
            font=("Tahoma", self.zoom_level, "normal"),
            header_font=("Tahoma", self.zoom_level, "bold"),
            row_height=self.zoom_level + 20,
            header_height=self.zoom_level + 25,
            auto_resize_columns=False,
        )
        self.sheet.redraw()
        self.after(100, self._save_col_widths)  # ← อัปเดต cache หลัง zoom
        pct = int((self.zoom_level / 11) * 100)
        if hasattr(self, 'zoom_label'):
            self.zoom_label.configure(text=f"{pct}%")
            
    def _show_all_columns(self):
        if not HAS_TKSHEET: return
        
        self.hidden_cols_list = [] # 🟢 ล้างความจำคอลัมน์ซ่อน
        self.sheet.display_columns("all")
        self._save_user_settings() # สั่งจำว่าเลิกซ่อนแล้ว
        self.sheet.redraw()