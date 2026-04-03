import tkinter as tk
from tkinter import messagebox
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkButton, CTkOptionMenu, CTkEntry
import pandas as pd
import psycopg2.extras
from datetime import datetime
import re
import json
from tkinter import colorchooser
from thefuzz import fuzz, process
try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False


class InlineSearchPopup(tk.Toplevel):
    def __init__(self, master, data_list, on_select_callback):
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg="#9CA3AF", padx=1, pady=1)

        self.data_list = data_list
        self.on_select_callback = on_select_callback
        self._destroyed = False

        self.search_var = tk.StringVar()
        self.entry = tk.Entry(
            self,
            textvariable=self.search_var,
            font=("Tahoma", 11),
            relief="flat",
            bd=4,
            highlightthickness=0,
        )
        self.entry.pack(fill="x")
        self.search_var.trace_add("write", self._on_type)

        self.listbox = tk.Listbox(
            self,
            font=("Tahoma", 11),
            selectbackground="#3B82F6",
            selectforeground="white",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            height=8,
        )
        self.listbox.pack(fill="both", expand=True)

        self.listbox.bind("<ButtonRelease-1>", self._on_click)
        self.entry.bind("<Down>", self._on_down)
        self.entry.bind("<Up>", self._on_up)
        self.entry.bind("<Return>", self._on_return)
        self.entry.bind("<KP_Enter>", self._on_return)
        self.entry.bind("<Escape>", self._on_escape)
        self.entry.bind("<FocusOut>", self._on_focus_out)

        self.filter_list("")
        self.entry.focus_set()

    def place_at_mouse(self, ref_widget=None):
        """วาง popup ที่ตำแหน่ง mouse ปัจจุบัน"""
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery() + 15
            self._set_geometry(x, y, 400, ref_widget=ref_widget)
        except Exception as e:
            print(f"place_at_mouse error: {e}")

    def place_near_cell(self, sheet_widget, row, col, row_h=30, header_h=35):
        try:
            master = self.master
            cx = getattr(master, '_last_click_x', 0) or sheet_widget.winfo_pointerx()
            cy = getattr(master, '_last_click_y', 0) or sheet_widget.winfo_pointery()
            self._set_geometry(cx, cy + 15, 400, ref_widget=sheet_widget)
            return  # ← เพิ่มตรงนี้ ไม่ให้ fallback รันทับ
        except Exception as e:
            print(f"place_near_cell error: {e}")

        # Fallback — รันเฉพาะตอน try ด้านบน error เท่านั้น
        try:
            x_pos = sheet_widget.row_index_width if hasattr(sheet_widget, 'row_index_width') else 40
            for c in range(col):
                try:
                    x_pos += sheet_widget.column_width(c)
                except Exception:
                    x_pos += 120
            y_pos = header_h + (row * row_h) + row_h
            rx = sheet_widget.winfo_rootx() + x_pos
            ry = sheet_widget.winfo_rooty() + y_pos
            self._set_geometry(rx, ry, 400, ref_widget=sheet_widget)
        except Exception as e:
            print(f"fallback error: {e}")

    def _set_geometry(self, rx, ry, popup_w, popup_h=240, ref_widget=None):
        """Set geometry โดยกันไม่ให้ออกนอกขอบของ window ที่ popup สังกัดอยู่"""
        try:
            # ✅ ใช้ขอบเขตของ window จริง (toplevel ที่ popup อยู่ใน) แทน screenwidth
            if ref_widget is not None:
                # หา root window (Tk หรือ Toplevel)
                root = ref_widget.winfo_toplevel()
                win_x = root.winfo_rootx()
                win_y = root.winfo_rooty()
                win_w = root.winfo_width()
                win_h = root.winfo_height()
            else:
                # fallback ใช้ screen
                win_x = 0
                win_y = 0
                win_w = self.winfo_screenwidth()
                win_h = self.winfo_screenheight()

            # กันเกินขอบขวาของ window
            if rx + popup_w > win_x + win_w:
                rx = win_x + win_w - popup_w - 5

            # ถ้า popup จะเกินล่าง window → เปิดขึ้นบนแทน
            if ry + popup_h > win_y + win_h:
                ry = ry - popup_h - row_h if hasattr(self, '_last_row_h') else ry - popup_h - 30

            # กันออกนอกขอบซ้าย/บนของ window
            rx = max(rx, win_x)
            ry = max(ry, win_y)

            self.geometry(f"{popup_w}x{popup_h}+{rx}+{ry}")
        except Exception:
            pass

    def filter_list(self, term):
        term = (term or "").lower().strip()
        self.listbox.delete(0, tk.END)
        if not term:
            for item in self.data_list[:100]:
                self.listbox.insert(tk.END, item)
            return
        search_terms = term.split()
        normal_matches = [
            item for item in self.data_list
            if all(t in str(item).lower() for t in search_terms)
        ]
        try:
            fuzzy_results = process.extractBests(
                term, [str(i) for i in self.data_list],
                scorer=fuzz.token_set_ratio, limit=20, score_cutoff=55
            )
            fuzzy_matches = [m[0] for m in fuzzy_results]
        except Exception:
            fuzzy_matches = []
        combined = list(dict.fromkeys(normal_matches + fuzzy_matches))
        for item in combined[:60]:
            self.listbox.insert(tk.END, item)

    def _on_type(self, *args):
        self.filter_list(self.search_var.get())

    def _on_click(self, e=None):
        if not self.listbox.curselection():
            return
        self._commit(self.listbox.get(self.listbox.curselection()))

    def _on_down(self, e=None):
        if self.listbox.size() > 0:
            curr = self.listbox.curselection()
            idx = (curr[0] + 1) if curr else 0
            idx = min(idx, self.listbox.size() - 1)
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
        return "break"

    def _on_up(self, e=None):
        if self.listbox.size() > 0:
            curr = self.listbox.curselection()
            if curr:
                idx = max(curr[0] - 1, 0)
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.see(idx)
        return "break"

    def _on_return(self, e=None):
        if self.listbox.curselection():
            self._commit(self.listbox.get(self.listbox.curselection()))
        elif self.listbox.size() > 0:
            self._commit(self.listbox.get(0))
        return "break"

    def _on_escape(self, e=None):
        self.safe_destroy()

    def _on_focus_out(self, e=None):
        self.after(150, self._check_focus_and_close)

    def _check_focus_and_close(self):
        try:
            if self._destroyed:
                return
            focused = self.focus_get()
            if focused not in (self.entry, self.listbox):
                self.safe_destroy()
        except Exception:
            self.safe_destroy()

    def _commit(self, value):
        self.on_select_callback(value)
        self.safe_destroy()

    def safe_destroy(self):
        if not self._destroyed:
            self._destroyed = True
            try:
                self.destroy()
            except Exception:
                pass


class CostBenchmarkScreen(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container

        self.auto_save_job_id = None
        self.current_user = getattr(self.app_container, 'current_user_key', 'PU_Default')

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.frozen_col_count = 0
        self.zoom_level = 11
        self.col_widths_cache = {}
        self.sales_list = []
        self.supplier_list = []
        self.product_list = []
        self._clipboard_row = []
        self.product_sku_map = {}
        self.supplier_code_map = {}
        self.product_category_map = {}
        self.hidden_cols_list = []
        self.custom_header_colors = {}
        self.sheet_frozen = None
        self._last_yview = -1.0
        self._active_popup = None
        self._last_popup_cell = None
        self._last_click_x = 0
        self._last_click_y = 0
        self._popup_opening = False
        self._header_filter_values = {}
        self._active_filter_popup = None
        self._last_popup_cell = None

        # --- Header & Filters ---
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
        CTkOptionMenu(header_frame, variable=self.month_var, values=self.thai_months, width=100,
                      command=self._load_from_db).grid(row=0, column=2, padx=(0, 10))

        CTkLabel(header_frame, text="ปี:").grid(row=0, column=3, padx=(0, 5))
        current_year_th = str(now.year + 543)
        year_list = [str(int(current_year_th) + i) for i in range(-2, 3)]
        self.year_var = tk.StringVar(value=current_year_th)
        CTkOptionMenu(header_frame, variable=self.year_var, values=year_list, width=80,
                      command=self._load_from_db).grid(row=0, column=4, padx=(0, 15))

        CTkButton(header_frame, text="🔄 โหลดข้อมูล", fg_color="#3B82F6", hover_color="#2563EB", width=90,
                  command=self._load_from_db).grid(row=0, column=5, sticky="w")

        btn_frame = CTkFrame(header_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=6, sticky="e")

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
        CTkButton(btn_frame, text="⮑ แทรกบรรทัด", fg_color="#10B981", hover_color="#059669",
                  command=self._insert_selected_row).pack(side="left", padx=5)

        CTkButton(btn_frame, text="🔽 ล้าง Filter", fg_color="#F59E0B", hover_color="#D97706",
          command=self._clear_all_header_filters).pack(side="left", padx=5)

        self.columns = [
            "วันที่ขอราคา", "Order No.", "Sale Order No.", "รหัส Sale",
            "PRIORITY", "WIN RATE %", "สถานะ", "QT", "Select",
            "หมวด", "ชื่อ Supplier", "แบรนด์", "รายการสินค้า",
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

        # Formula bar
        self.target_formula_cell = None
        self.formula_frame = CTkFrame(self, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#CBD5E1")
        self.formula_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 5))
        CTkLabel(self.formula_frame, text=" 𝑓x ", font=CTkFont(family="Arial", size=18, weight="bold", slant="italic"),
                 text_color="#16A34A").pack(side="left", padx=10, pady=5)
        self.formula_entry = CTkEntry(self.formula_frame, font=CTkFont(size=14),
                                     placeholder_text="คลิกช่องปลายทาง -> คลิกที่นี่ -> พิมพ์ = แล้วใช้เมาส์จิ้มเซลล์ในตารางได้เลย!",
                                     border_width=0, fg_color="transparent")
        self.formula_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=5)
        self.formula_entry.bind("<FocusIn>", self._on_formula_focus_in)
        self.formula_entry.bind("<Return>", self._apply_formula_from_bar)

        # Table frame
        self.table_frame = tk.Frame(self, bg="white")
        self.table_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 5))
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)

        # Status bar
        self.bottom_status_frame = CTkFrame(self, fg_color="#E5E7EB", corner_radius=4)
        self.bottom_status_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 10))

        zoom_frame = CTkFrame(self.bottom_status_frame, fg_color="transparent")
        zoom_frame.pack(side="right", padx=5)
        CTkButton(zoom_frame, text="🔍−", width=32, height=24, fg_color="#6B7280", hover_color="#4B5563",
                  font=CTkFont(size=12), command=lambda: self._zoom(-1)).pack(side="left", padx=2)
        self.zoom_label = CTkLabel(zoom_frame, text="100%", font=CTkFont(size=12), text_color="#6B7280", width=40)
        self.zoom_label.pack(side="left")
        CTkButton(zoom_frame, text="🔍+", width=32, height=24, fg_color="#6B7280", hover_color="#4B5563",
                  font=CTkFont(size=12), command=lambda: self._zoom(1)).pack(side="left", padx=2)

        self.save_status_label = CTkLabel(self.bottom_status_frame,
                                          text="✅ พร้อมใช้งาน (บันทึกอัตโนมัติ)",
                                          font=CTkFont(size=13), text_color="gray50")
        self.save_status_label.pack(side="left", padx=20, pady=4)

        self.quick_calc_label = CTkLabel(self.bottom_status_frame, text="",
                                         font=CTkFont(size=14, weight="bold"), text_color="#059669")
        self.quick_calc_label.pack(side="right", padx=20, pady=4)

        if HAS_TKSHEET:
            self._build_tksheet(self.table_frame)
            self.after(200, self._load_from_db)
        else:
            tk.Label(self.table_frame, text="⚠️ กรุณาติดตั้ง tksheet", fg="red", bg="white").pack(expand=True)

    def _capture_click_pos(self, event=None):
        if event:
            self._last_click_x = event.x_root
            self._last_click_y = event.y_root

    # ================================================================== #
    def _lighten_color(self, hex_color, amount=0.85):
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            r = int(r + (255 - r) * amount)
            g = int(g + (255 - g) * amount)
            b = int(b + (255 - b) * amount)
            return f'#{r:02x}{g:02x}{b:02x}'
        except Exception:
            return "#F3F4F6"

    def _load_user_settings(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    "SELECT setting_value FROM user_settings WHERE user_name = %s AND setting_key = 'benchmark_table'",
                    (self.current_user,))
                result = cursor.fetchone()
                if result and result[0]:
                    settings = result[0]
                    self.hidden_cols_list = settings.get("hidden_cols", [])
                    self.custom_header_colors = settings.get("header_colors", {})
                    # ← เพิ่ม: โหลด col widths และ frozen state
                    saved_widths = settings.get("col_widths", {})
                    if saved_widths:
                        self.col_widths_cache = {int(k): v for k, v in saved_widths.items()}
                    saved_freeze = settings.get("frozen_col_count", 0)
                    if saved_freeze > 0:
                        self.frozen_col_count = saved_freeze
        except Exception as e:
            print(f"Error loading settings: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _save_user_settings(self):
        # เก็บ col widths ปัจจุบัน
        col_widths_to_save = {}
        try:
            if self.sheet_frozen:
                for i in range(self.frozen_col_count):
                    w = self.sheet_frozen.column_width(i)
                    if w: col_widths_to_save[str(i)] = w
                for i in range(self.sheet.get_total_columns()):
                    w = self.sheet.column_width(i)
                    if w: col_widths_to_save[str(self.frozen_col_count + i)] = w
            else:
                for i in range(self.sheet.get_total_columns()):
                    w = self.sheet.column_width(i)
                    if w: col_widths_to_save[str(i)] = w
        except Exception:
            pass

        settings = {
            "hidden_cols": self.hidden_cols_list,
            "header_colors": self.custom_header_colors,
            "col_widths": col_widths_to_save,
            "frozen_col_count": self.frozen_col_count,
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
            data=[[""] * len(self.columns) for _ in range(1000)],
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
        self._load_user_settings()  # ← โหลดก่อน เพื่อรู้ frozen_col_count
        self._rebind_sheet()
        self._rebind_frozen_sheet()
        self._apply_formatting(col_offset=0)
        if self.hidden_cols_list:
            self.sheet.hide_columns(self.hidden_cols_list)
        # ← ถ้ามี frozen ที่บันทึกไว้ ให้ rebuild
        if self.frozen_col_count > 0:
            self.after(100, self._rebuild_frozen_layout)
        self.after(500, self._save_col_widths)
        self.after(5000, self._start_dropdown_refresh_timer)

    # ================================================================== #
    # FREEZE / UNFREEZE
    # ================================================================== #
    def _freeze_selected_columns(self):
        if not HAS_TKSHEET:
            return
        freeze_up_to = None
        try:
            selected_cols = self.sheet.get_selected_columns()
            if selected_cols:
                freeze_up_to = max(selected_cols) + 1
        except Exception:
            pass
        if freeze_up_to is None:
            try:
                selected_cells = self.sheet.get_selected_cells()
                if selected_cells:
                    freeze_up_to = max(c for _, c in selected_cells) + 1
            except Exception:
                pass
        if freeze_up_to is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเซลล์หรือหัวคอลัมน์ที่ต้องการตรึงก่อน", parent=self)
            return
        actual_freeze = self.frozen_col_count + freeze_up_to if self.sheet_frozen is not None else freeze_up_to
        actual_freeze = min(actual_freeze, len(self.columns) - 1)
        self.frozen_col_count = actual_freeze
        self._rebuild_frozen_layout()
        col_name = self.columns[actual_freeze - 1] if actual_freeze <= len(self.columns) else "?"
        self.save_status_label.configure(
            text=f"📌 ตรึง {actual_freeze} คอลัมน์แรก (ถึง: {col_name})",
            text_color="#0891B2"
        )

    def _unfreeze_columns(self):
        if not HAS_TKSHEET:
            return
        self.frozen_col_count = 0
        self._rebuild_frozen_layout()
        self.save_status_label.configure(text="✅ ยกเลิกตรึงแล้ว", text_color="#16A34A")

    def _rebuild_frozen_layout(self):
        try:
            if self.sheet_frozen is not None:
                frozen_data = self.sheet_frozen.get_sheet_data()
                scroll_data = self.sheet.get_sheet_data()
                n_rows = max(len(frozen_data), len(scroll_data))
                current_data = []
                for i in range(n_rows):
                    left = frozen_data[i] if i < len(frozen_data) else [""] * self.frozen_col_count
                    right = scroll_data[i] if i < len(scroll_data) else [""] * (len(self.columns) - self.frozen_col_count)
                    current_data.append(left + right)
            else:
                current_data = self.sheet.get_sheet_data()
        except Exception:
            current_data = [[""] * len(self.columns) for _ in range(20)]

        saved_widths = {}
        try:
            if self.sheet_frozen is not None:
                for i in range(self.frozen_col_count):
                    try:
                        w = self.sheet_frozen.column_width(i)
                        if w and w > 0:
                            saved_widths[i] = w
                    except Exception:
                        pass
                offset = self.frozen_col_count
                for i in range(self.sheet.get_total_columns()):
                    try:
                        w = self.sheet.column_width(i)
                        if w and w > 0:
                            saved_widths[offset + i] = w
                    except Exception:
                        pass
            else:
                for i in range(self.sheet.get_total_columns()):
                    try:
                        w = self.sheet.column_width(i)
                        if w and w > 0:
                            saved_widths[i] = w
                    except Exception:
                        pass
        except Exception:
            pass

        for k, v in self.col_widths_cache.items():
            if k not in saved_widths:
                saved_widths[k] = v

        try:
            self.table_frame.unbind("<Configure>")
        except Exception:
            pass
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        self.sheet_frozen = None
        self.table_frame.grid_rowconfigure(0, weight=1)

        n_freeze = self.frozen_col_count
        frozen_width = 0
        if n_freeze > 0:
            frozen_width = 40
            for i in range(n_freeze):
                frozen_width += saved_widths.get(i, 120)
            frozen_width += 4

        if n_freeze == 0:
            self.table_frame.grid_columnconfigure(0, weight=1, minsize=0)
            self.table_frame.grid_columnconfigure(1, weight=0, minsize=0)
            self.sheet = Sheet(
                self.table_frame,
                headers=self.columns,
                data=current_data,
                theme="light blue",
                row_height=30,
                header_height=35,
                font=("Tahoma", self.zoom_level, "normal"),
                header_font=("Tahoma", self.zoom_level, "bold"),
                show_row_index=True,
                row_index_width=40,
                column_width=120,
                empty_horizontal=0,
                empty_vertical=0,
            )
            self.sheet.grid(row=0, column=0, sticky="nsew")
        else:
            self.table_frame.grid_columnconfigure(0, weight=0, minsize=0)
            self.table_frame.grid_columnconfigure(1, weight=0, minsize=0)
            frozen_cols = self.columns[:n_freeze]
            scroll_cols = self.columns[n_freeze:]
            frozen_data_split = [row[:n_freeze] for row in current_data]
            scroll_data_split = [row[n_freeze:] for row in current_data]
            
            current_row_h = int(30 * (self.zoom_level / 11.0))
            current_header_h = int(35 * (self.zoom_level / 11.0))

            self.sheet_frozen = Sheet(
                self.table_frame,
                headers=frozen_cols,
                data=frozen_data_split,
                theme="light blue",
                row_height=current_row_h,      
                header_height=current_header_h,
                font=("Tahoma", self.zoom_level, "normal"),
                header_font=("Tahoma", self.zoom_level, "bold"),
                show_row_index=True,
                row_index_width=40,
                column_width=120,
                empty_horizontal=0,
                empty_vertical=0,
            )
            try:
                self.sheet_frozen.hide(canvas="x_scrollbar")
                self.sheet_frozen.hide(canvas="y_scrollbar")
                # ← เพิ่ม ซ่อน scrollbar ล่างสุดด้วย
                try:
                    self.sheet_frozen.yscrollbar.grid_remove()
                except Exception:
                    pass
                try:
                    self.sheet_frozen.xscrollbar.grid_remove()
                except Exception:
                    pass
            except Exception:
                pass
            self.sheet_frozen.place(x=0, y=0, width=frozen_width, relheight=1.0)
            self.after(10, lambda: self._hide_frozen_scrollbars())


            self.sheet = Sheet(
                self.table_frame,
                headers=scroll_cols,
                data=scroll_data_split,
                theme="light blue",
                row_height=current_row_h,      
                header_height=current_header_h,
                font=("Tahoma", self.zoom_level, "normal"),
                header_font=("Tahoma", self.zoom_level, "bold"),
                show_row_index=False,
                row_index_width=0,
                column_width=120,
                empty_horizontal=0,
                empty_vertical=0,
            )
            self.sheet.place(x=frozen_width, y=0, relwidth=1.0, relheight=1.0, width=-frozen_width)

            _fw = frozen_width
            def _on_table_resize(event, fw=_fw):
                try:
                    self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                except Exception:
                    pass
            self.table_frame.bind("<Configure>", _on_table_resize)
            self._sync_vertical_scroll()

        self._rebind_sheet()
        self._rebind_frozen_sheet()   
        self._apply_formatting(col_offset=n_freeze)
        if self.sheet_frozen is not None:
            self._apply_formatting_frozen(n_freeze)

        

        def _restore_after_build():
            try:
                if self.sheet_frozen is not None:
                    for i in range(n_freeze):
                        w = saved_widths.get(i, 120)
                        self.sheet_frozen.column_width(i, w)
                    for i in range(len(self.columns) - n_freeze):
                        w = saved_widths.get(n_freeze + i, 120)
                        self.sheet.column_width(i, w)
                    real_fw = 40
                    for i in range(n_freeze):
                        real_fw += saved_widths.get(i, 120)
                    real_fw += 4
                    self.sheet_frozen.place(x=0, y=0, width=real_fw, relheight=1.0)
                    self.sheet.place(x=real_fw, y=0, relwidth=1.0, relheight=1.0, width=-real_fw)

                    # ← เพิ่ม sync row height ให้ตรงกัน
                    try:
                        total_rows = max(self.sheet.get_total_rows(), self.sheet_frozen.get_total_rows())
                        rh = int(30 * (self.zoom_level / 11.0))
                        for r in range(total_rows):
                            self.sheet_frozen.row_height(r, rh)
                            self.sheet.row_height(r, rh)
                    except Exception:
                        pass

                    def _on_table_resize_updated(event, fw=real_fw):
                        try:
                            self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                        except Exception:
                            pass
                    self.table_frame.bind("<Configure>", _on_table_resize_updated)
                    self.after(200, self._hide_frozen_scrollbars)
                    self.after(600, self._hide_frozen_scrollbars)
                else:
                    for i, w in saved_widths.items():
                        try:
                            self.sheet.column_width(i, w)
                        except Exception:
                            pass
                self.sheet.redraw()
                if self.sheet_frozen:
                    self.sheet_frozen.redraw()
            except Exception:
                pass
        self.after(150, _restore_after_build)
    def _sync_vertical_scroll(self):
        if not self.sheet_frozen:
            return
        self._last_yview = -1.0

        def _do_sync():
            try:
                if not self.sheet.winfo_exists() or not self.sheet_frozen.winfo_exists():
                    return
            except Exception:
                return
            try:
                yview = self.sheet.get_yview()
                current_pos = yview[0] if yview else 0.0
                if abs(current_pos - self._last_yview) > 0.0001:
                    self._last_yview = current_pos
                    self.sheet_frozen.yview_moveto(current_pos)
                    try:
                        self.sheet_frozen.MT.yview_moveto(current_pos)
                    except Exception:
                        pass
            except Exception:
                pass
            if self.frozen_col_count > 0:
                self.after(30, _do_sync)
        self.after(50, _do_sync)

        # ← เก็บ sync functions ไว้ใน self เพื่อ unbind/rebind ได้ถูกต้อง
        def _sync_from_main(event=None):
            self.after(20, _sync_to_frozen)

        def _sync_from_frozen(event=None):
            self.after(20, _sync_to_main)

        def _sync_to_frozen():
            try:
                if not self.sheet.winfo_exists() or not self.sheet_frozen.winfo_exists():
                    return
                yview = self.sheet.get_yview()
                if yview:
                    self.sheet_frozen.yview_moveto(yview[0])
                    try:
                        self.sheet_frozen.MT.yview_moveto(yview[0])
                    except Exception:
                        pass
            except Exception:
                pass

        def _sync_to_main():
            try:
                if not self.sheet.winfo_exists() or not self.sheet_frozen.winfo_exists():
                    return
                try:
                    yview = self.sheet_frozen.get_yview()
                except Exception:
                    yview = None
                if yview:
                    self.sheet.yview_moveto(yview[0])
                    try:
                        self.sheet.MT.yview_moveto(yview[0])
                    except Exception:
                        pass
                    try:
                        self.sheet_frozen.MT.yview_moveto(yview[0])
                    except Exception:
                        pass
                    self._last_yview = yview[0]
            except Exception:
                pass

        # ← unbind เก่าก่อน แล้วค่อย bind ใหม่ ป้องกัน event ซ้อนกัน
        for widget in [self.sheet, self.sheet.MT]:
            try:
                widget.unbind("<MouseWheel>")
                widget.unbind("<Button-4>")
                widget.unbind("<Button-5>")
            except Exception:
                pass
            try:
                widget.bind("<MouseWheel>", _sync_from_main)
                widget.bind("<Button-4>",   _sync_from_main)
                widget.bind("<Button-5>",   _sync_from_main)
            except Exception:
                pass

        for widget in [self.sheet_frozen, self.sheet_frozen.MT]:
            try:
                widget.unbind("<MouseWheel>")
                widget.unbind("<Button-4>")
                widget.unbind("<Button-5>")
            except Exception:
                pass
            try:
                widget.bind("<MouseWheel>", _sync_from_frozen)
                widget.bind("<Button-4>",   _sync_from_frozen)
                widget.bind("<Button-5>",   _sync_from_frozen)
            except Exception:
                pass

        # ← bind RI ของ frozen ด้วย
        try:
            self.sheet_frozen.RI.unbind("<MouseWheel>")
            self.sheet_frozen.RI.unbind("<Button-4>")
            self.sheet_frozen.RI.unbind("<Button-5>")
        except Exception:
            pass
        try:
            self.sheet_frozen.RI.bind("<MouseWheel>", _sync_from_frozen)
            self.sheet_frozen.RI.bind("<Button-4>",   _sync_from_frozen)
            self.sheet_frozen.RI.bind("<Button-5>",   _sync_from_frozen)
        except Exception:
            pass

        # ← bind RI ของ main ด้วย (กรณีเมาส์อยู่บน row index ฝั่ง main)
        try:
            self.sheet.RI.unbind("<MouseWheel>")
            self.sheet.RI.unbind("<Button-4>")
            self.sheet.RI.unbind("<Button-5>")
        except Exception:
            pass
        try:
            self.sheet.RI.bind("<MouseWheel>", _sync_from_main)
            self.sheet.RI.bind("<Button-4>",   _sync_from_main)
            self.sheet.RI.bind("<Button-5>",   _sync_from_main)
        except Exception:
            pass

    # ================================================================== #
    # FORMATTING
    # ================================================================== #
    def _get_header_styles_map(self):
        return {
            ("#2563EB", "white"): ["วันที่ขอราคา", "Order No.", "Sale Order No.", "QT"],
            ("#BAE6FD", "black"): [
                "หมายเหตุ (ความยาว, OD)", "หมายเหตุ", "จำนวน", "ต้นทุน/เส้น", "น้ำหนัก/เส้น",
                "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (บาท)", "ส่วนลด 1 (%)",
                "ส่วนลด 2 (บาท)", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 1", "ทุน/เส้น หลังส่วนลด 2",
                "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)",
                "ค่าย้าย (ซื้อ)", "ค่าย้าย/เส้น", "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)",
                "ต้นทุนรวม (รวมย้าย)"
            ],
            ("#FDBA74", "black"): ["ผู้ขอราคา", "รหัส Sale", "PRIORITY", "WIN RATE %", "Select", "แบรนด์"],
            ("#6B7280", "white"): ["หมวด", "หมวดหลัก", "หมวดรอง", "หมวดย่อย", "Product SKU.",
                                   "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น",
                                   "ต้นทุนรวม+Markup", "ค่าส่ง (ขาย)", "ค่าส่ง / เส้น"],
            ("#D8B4FE", "black"): ["รายการสินค้า", "ชื่อ Supplier"],
            ("#FCA5A5", "black"): ["Markup Guide (%)"],
            ("#FDE047", "black"): ["สถานะ", "น้ำหนัก/เส้น 2", "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม + Vat."],
            ("#86EFAC", "black"): ["ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น", "ราคาขาย รวม", "Vat. รวม"],
            ("#1F2937", "white"): ["ชื่อ Supplier2", "Sup ID.", "คลังสินค้า ต้นทาง"],
            ("#93C5FD", "black"): ["ปลายทาง", "หมายเหตุ2"]
        }

    def _get_selected_row_idx(self):
        try:
            curr = self.sheet.get_currently_selected()
            print(f"DEBUG move - get_currently_selected: {curr}")  # ← เพิ่ม
            if curr:
                return int(curr[0])
        except Exception as e:
            print(f"DEBUG move error: {e}")
        
        if self.sheet_frozen:
            try:
                curr = self.sheet_frozen.get_currently_selected()
                print(f"DEBUG move frozen - get_currently_selected: {curr}")  # ← เพิ่ม
                if curr:
                    return int(curr[0])
            except Exception:
                pass
        
        try:
            selected_rows = self.sheet.get_selected_rows()
            print(f"DEBUG move - get_selected_rows: {selected_rows}")  # ← เพิ่ม
            if selected_rows:
                return min(int(r) for r in selected_rows)
        except Exception:
            pass
        
        print("DEBUG move - ไม่พบ selection")  # ← เพิ่ม
        return None

    def _swap_rows(self, idx_a, idx_b):
        """สลับข้อมูล 2 บรรทัดทั้ง main และ frozen"""
        try:
            # main sheet
            row_a = list(self.sheet.get_row_data(idx_a))
            row_b = list(self.sheet.get_row_data(idx_b))
            for c, val in enumerate(row_b):
                self.sheet.set_cell_data(idx_a, c, val, redraw=False)
            for c, val in enumerate(row_a):
                self.sheet.set_cell_data(idx_b, c, val, redraw=False)
        except Exception as e:
            print(f"_swap_rows main error: {e}")

        if self.sheet_frozen:
            try:
                frow_a = list(self.sheet_frozen.get_row_data(idx_a))
                frow_b = list(self.sheet_frozen.get_row_data(idx_b))
                for c, val in enumerate(frow_b):
                    self.sheet_frozen.set_cell_data(idx_a, c, val, redraw=False)
                for c, val in enumerate(frow_a):
                    self.sheet_frozen.set_cell_data(idx_b, c, val, redraw=False)
            except Exception as e:
                print(f"_swap_rows frozen error: {e}")

        self.sheet.redraw()
        if self.sheet_frozen:
            self.sheet_frozen.redraw()

    def _move_row_up(self):
        if not HAS_TKSHEET:
            return
        idx = self._get_selected_row_idx()
        if idx is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการย้ายก่อน", parent=self)
            return
        if idx <= 0:
            return
        self._swap_rows(idx, idx - 1)
        # ย้าย selection ขึ้นด้วย
        self.sheet.select_cell(idx - 1, 0)
        self.sheet.see(idx - 1, 0)
        if self.sheet_frozen:
            self.sheet_frozen.select_cell(idx - 1, 0)
        # debounce save
        if self.auto_save_job_id:
            self.after_cancel(self.auto_save_job_id)
        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

    def _move_row_down(self):
        if not HAS_TKSHEET:
            return
        idx = self._get_selected_row_idx()
        if idx is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการย้ายก่อน", parent=self)
            return
        total_rows = self.sheet.get_total_rows()
        if idx >= total_rows - 1:
            return
        self._swap_rows(idx, idx + 1)
        self.sheet.select_cell(idx + 1, 0)
        self.sheet.see(idx + 1, 0)
        if self.sheet_frozen:
            self.sheet_frozen.select_cell(idx + 1, 0)
        if self.auto_save_job_id:
            self.after_cancel(self.auto_save_job_id)
        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

    def _get_auto_cols(self):
        return [
            "Product SKU.", "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม",
            "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2",
            "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย/เส้น",
            "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)", "Markup/กก.",
            "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup",
            "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น",
            "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat.",
            "ชื่อ Supplier2", "Sup ID."
        ]

    def _apply_formatting(self, col_offset=0):
        auto_cols_names = self._get_auto_cols()
        header_styles_map = self._get_header_styles_map()
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles_map.items() for c in cols}
        if hasattr(self, 'custom_header_colors'):
            for col_name, bg_color in self.custom_header_colors.items():
                if col_name in self.columns:
                    col_to_style[col_name] = (bg_color, "black")

        def get_display_idx(col_name):
            try:
                real_idx = self.columns.index(col_name)
                display_idx = real_idx - col_offset
                return display_idx if display_idx >= 0 else None
            except ValueError:
                return None

        status_opts = [
            "", "WIN", "STOCK",
            "LOSE - เซลล์ไม่ทราบสาเหตุ",
            "LOSE - ลูกค้าได้ราคาถูกกว่า (มีราคาเทียบ)",
            "LOSE - ลูกค้าได้ราคาถูกกว่า (ไม่มีราคาเทียบ)",
            "LOSE - ไม่มีกำหนดใช้งานที่แน่นอน เช่น ขอราคาเพื่อเสนอ",
            "LOSE - ยื่นประมูลงาน (ระบุเดือนในหมายเหตุ)",
            "LOSE - ลูกค้าเปลี่ยนสเปคการใช้งาน",
            "LOSE - ลูกค้าใช้เจ้าที่มีเครดิต"
        ]
        dropdown_defs = [
            ("PRIORITY", ["", "HOT", "WARM", "COLD", "ไม่แจ้ง"], "readonly"),
            ("สถานะ", status_opts, "readonly"),
            ("Select", ["", "✔", "เทียบ", "เทียบเพื่อชุบ"], "readonly"),
        ]
        for col_name, values, state in dropdown_defs:
            idx = get_display_idx(col_name)
            if idx is not None:
                self.sheet.create_dropdown("all", idx, values=values, state=state)

        auto_display_indices = []
        readonly_cols = self._get_auto_cols()
        for c in readonly_cols:
            idx = get_display_idx(c)
            if idx is not None:
                auto_display_indices.append(idx)
        if auto_display_indices:
            self.sheet.readonly_columns(columns=auto_display_indices, readonly=True)

        display_cols = self.columns[col_offset:]
        for di, col in enumerate(display_cols):
            h_bg, h_fg = col_to_style.get(col, ("#E5E7EB", "#111827"))
            if hasattr(self, 'custom_header_colors') and col in self.custom_header_colors:
                b_bg = self._lighten_color(self.custom_header_colors[col], 0.85)
                b_fg = "black"
            elif col in auto_cols_names:
                b_bg = "#F3F4F6"
                b_fg = "#111827"
            else:
                b_bg = "white"
                b_fg = "black"
            try:
                self.sheet.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                self.sheet.highlight_columns(columns=[di], bg=b_bg, fg=b_fg, highlight_header=False)
            except Exception:
                try:
                    self.sheet.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                except Exception:
                    pass
        
        self._setup_header_filters(col_offset)
        self.sheet.set_options(
            grid_color="#000000", outline_color="#000000", table_bg="white", table_fg="black",
            table_grid_fg="#000000", header_bg="#D1D5DB", header_fg="#111827", header_grid_fg="#000000",
            header_selected_cells_bg="#9CA3AF", row_index_bg="#F3F4F6", row_index_fg="#111827",
            row_index_grid_fg="#000000", selected_cells_border_color="#3B82F6",
            table_selected_cells_border_color="#3B82F6",
            auto_resize_columns=False, auto_resize_row_index=False,
            dropdown_font=("Tahoma", 11, "normal"),
        )

    def _apply_formatting_frozen(self, n_freeze):
        if not self.sheet_frozen:
            return
        auto_cols_names = self._get_auto_cols()
        header_styles_map = self._get_header_styles_map()
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles_map.items() for c in cols}
        if hasattr(self, 'custom_header_colors'):
            for col_name, bg_color in self.custom_header_colors.items():
                col_to_style[col_name] = (bg_color, "black")

        frozen_col_names = self.columns[:n_freeze]
        for di, col in enumerate(frozen_col_names):
            h_bg, h_fg = col_to_style.get(col, ("#E5E7EB", "#111827"))
            if hasattr(self, 'custom_header_colors') and col in self.custom_header_colors:
                b_bg = self._lighten_color(self.custom_header_colors[col], 0.85)
                b_fg = "black"
            elif col in auto_cols_names:
                b_bg = "#F3F4F6"
                b_fg = "#111827"
            else:
                b_bg = "white"
                b_fg = "black"
            try:
                self.sheet_frozen.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                self.sheet_frozen.highlight_columns(columns=[di], bg=b_bg, fg=b_fg, highlight_header=False)
            except Exception:
                try:
                    self.sheet_frozen.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                except Exception:
                    pass

        # ← เพิ่ม dropdown สำหรับคอลัมน์ที่อยู่ในฝั่ง frozen
        status_opts = [
            "", "WIN", "STOCK",
            "LOSE - เซลล์ไม่ทราบสาเหตุ",
            "LOSE - ลูกค้าได้ราคาถูกกว่า (มีราคาเทียบ)",
            "LOSE - ลูกค้าได้ราคาถูกกว่า (ไม่มีราคาเทียบ)",
            "LOSE - ไม่มีกำหนดใช้งานที่แน่นอน เช่น ขอราคาเพื่อเสนอ",
            "LOSE - ยื่นประมูลงาน (ระบุเดือนในหมายเหตุ)",
            "LOSE - ลูกค้าเปลี่ยนสเปคการใช้งาน",
            "LOSE - ลูกค้าใช้เจ้าที่มีเครดิต"
        ]
        dropdown_defs = [
            ("PRIORITY", ["", "HOT", "WARM", "COLD", "ไม่แจ้ง"], "readonly"),
            ("สถานะ", status_opts, "readonly"),
            ("Select", ["", "✔", "เทียบ", "เทียบเพื่อชุบ"], "readonly"),
        ]
        for col_name, values, state in dropdown_defs:
            if col_name in frozen_col_names:
                idx = frozen_col_names.index(col_name)
                try:
                    self.sheet_frozen.create_dropdown("all", idx, values=values, state=state)
                except Exception:
                    pass

        # readonly สำหรับ auto cols ที่อยู่ในฝั่ง frozen
        auto_frozen_indices = []
        for col_name in auto_cols_names:
            if col_name in frozen_col_names:
                idx = frozen_col_names.index(col_name)
                auto_frozen_indices.append(idx)
        if auto_frozen_indices:
            try:
                self.sheet_frozen.readonly_columns(columns=auto_frozen_indices, readonly=True)
            except Exception:
                pass

        self.sheet_frozen.set_options(
            grid_color="#000000", outline_color="#000000", table_bg="white", table_fg="black",
            table_grid_fg="#000000", header_bg="#D1D5DB", header_fg="#111827", header_grid_fg="#000000",
            row_index_bg="#F3F4F6", row_index_fg="#111827", row_index_grid_fg="#000000",
            auto_resize_columns=False, auto_resize_row_index=False,
            dropdown_font=("Tahoma", 11, "normal"),
        )

    def _setup_header_filters(self, col_offset=0):
        self._filter_col_names = {"Order No.", "Sale Order No.", "ชื่อ Supplier", "รายการสินค้า"}
        try:
            self.sheet.CH.bind("<ButtonPress-1>", self._on_header_click, add="+")
        except Exception:
            pass

    def _on_header_click(self, event=None):
        try:
            print(f"DEBUG header click x={event.x} y={event.y}")  # ← เพิ่ม
            
            # ลอง identify col หลายวิธี
            col = None
            try:
                col = self.sheet.CH.identify_col(x=event.x, allow_end=False)
            except Exception as e1:
                print(f"identify_col error: {e1}")
                try:
                    col = self.sheet.identify_col(event)
                except Exception as e2:
                    print(f"fallback identify_col error: {e2}")
            
            print(f"DEBUG col={col}")  # ← เพิ่ม
            
            if col is None:
                return
                
            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = int(col) + col_offset
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]
            
            print(f"DEBUG col_name={col_name}, filter_cols={self._filter_col_names}")  # ← เพิ่ม
            
            if col_name not in self._filter_col_names:
                return

            if self._active_filter_popup and not getattr(self._active_filter_popup, '_destroyed', True):
                self._active_filter_popup.safe_destroy()
                self._active_filter_popup = None
                return

            self._show_header_filter_popup(col_name, int(col), event)
        except Exception as e:
            print(f"_on_header_click error: {e}")
            import traceback; traceback.print_exc()

    def _show_header_filter_popup(self, col_name, display_col, event):
        try:
            # ดึงค่าทั้งหมดในคอลัมน์
            all_values = sorted(set(
                str(self.sheet.get_cell_data(r, display_col) or "").strip()
                for r in range(self.sheet.get_total_rows())
                if str(self.sheet.get_cell_data(r, display_col) or "").strip()
            ))
            if not all_values:
                return

            current_filter = self._header_filter_values.get(col_name)

            popup = tk.Toplevel(self)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)
            popup.configure(bg="#D1D5DB", padx=1, pady=1)
            popup._destroyed = False

            def safe_destroy():
                if not popup._destroyed:
                    popup._destroyed = True
                    try: popup.destroy()
                    except: pass
                self._active_filter_popup = None

            popup.safe_destroy = safe_destroy

            inner = tk.Frame(popup, bg="white")
            inner.pack(fill="both", expand=True)

            # Search
            search_var = tk.StringVar()
            search_entry = tk.Entry(inner, textvariable=search_var,
                                font=("Tahoma", 11), relief="flat", bd=6,
                                highlightthickness=1, highlightcolor="#3B82F6",
                                highlightbackground="#D1D5DB")
            search_entry.pack(fill="x", padx=4, pady=4)
            search_entry.insert(0, "🔍 ค้นหา...")
            search_entry.bind("<FocusIn>", lambda e: search_entry.delete(0, "end")
                            if search_entry.get().startswith("🔍") else None)

            # Listbox
            lb_frame = tk.Frame(inner, bg="white")
            lb_frame.pack(fill="both", expand=True, padx=4)

            sb = tk.Scrollbar(lb_frame)
            sb.pack(side="right", fill="y")

            listbox = tk.Listbox(lb_frame, font=("Tahoma", 11),
                                selectbackground="#3B82F6", selectforeground="white",
                                relief="flat", borderwidth=0, highlightthickness=0,
                                height=10, yscrollcommand=sb.set, activestyle="none")
            listbox.pack(side="left", fill="both", expand=True)
            sb.config(command=listbox.yview)

            def fill_list(term=""):
                listbox.delete(0, tk.END)
                listbox.insert(tk.END, "  (ทั้งหมด)")
                for v in all_values:
                    if term.lower() in v.lower():
                        listbox.insert(tk.END, f"  {v}")
                # เลือก current
                if not term:
                    if current_filter:
                        for i in range(listbox.size()):
                            if listbox.get(i).strip() == current_filter:
                                listbox.selection_set(i)
                                listbox.see(i)
                                break
                    else:
                        listbox.selection_set(0)

            fill_list()

            def on_search(*args):
                fill_list(search_var.get())

            search_var.trace_add("write", on_search)

            def on_select(e=None):
                sel = listbox.curselection()
                if not sel:
                    return
                val = listbox.get(sel[0]).strip()
                if val == "(ทั้งหมด)":
                    self._header_filter_values.pop(col_name, None)
                else:
                    self._header_filter_values[col_name] = val
                safe_destroy()
                self._apply_header_filters()

            listbox.bind("<ButtonRelease-1>", on_select)
            listbox.bind("<Return>", on_select)
            search_entry.bind("<Return>", on_select)
            search_entry.bind("<Escape>", lambda e: safe_destroy())

            # ปุ่ม Clear
            btn_bar = tk.Frame(inner, bg="#F9FAFB")
            btn_bar.pack(fill="x")
            tk.Button(btn_bar, text="✕ ล้าง filter", font=("Tahoma", 10),
                    fg="#EF4444", bg="#F9FAFB", relief="flat", bd=0, cursor="hand2",
                    command=lambda: [self._header_filter_values.pop(col_name, None),
                                    safe_destroy(), self._apply_header_filters()]
                    ).pack(side="right", padx=8, pady=4)

            # วางตำแหน่ง
            x = event.x_root
            y = event.y_root + 5
            root = self.winfo_toplevel()
            win_w = root.winfo_rootx() + root.winfo_width()
            if x + 220 > win_w:
                x = win_w - 225
            popup.geometry(f"220x310+{x}+{y}")

            self._active_filter_popup = popup
            search_entry.focus_set()
            popup.bind("<FocusOut>", lambda e: self.after(300, lambda:
                safe_destroy() if not popup._destroyed else None))

        except Exception as e:
            print(f"_show_header_filter_popup error: {e}")

    def _apply_header_filters(self):
        try:
            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            filter_map = {}
            for col_name, val in self._header_filter_values.items():
                try:
                    real_idx = self.columns.index(col_name)
                    display_idx = real_idx - col_offset
                    if display_idx >= 0:
                        filter_map[display_idx] = val
                except ValueError:
                    pass

            for r in range(self.sheet.get_total_rows()):
                if not filter_map:
                    self.sheet.row_height(r, self.zoom_level + 19)
                    continue
                match = all(
                    str(self.sheet.get_cell_data(r, d_idx) or "").strip() == fval
                    for d_idx, fval in filter_map.items()
                )
                self.sheet.row_height(r, self.zoom_level + 19 if match else 0)

            # อัพเดทสีหัวคอลัมน์
            header_styles = self._get_header_styles_map()
            col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles.items() for c in cols}
            for col_name in self._filter_col_names:
                try:
                    real_idx = self.columns.index(col_name)
                    display_idx = real_idx - col_offset
                    if display_idx < 0:
                        continue
                    if col_name in self._header_filter_values:
                        self.sheet.highlight_cells(row=0, column=display_idx,
                                                bg="#F59E0B", fg="white", canvas="header")
                    else:
                        h_bg, h_fg = col_to_style.get(col_name, ("#E5E7EB", "#111827"))
                        self.sheet.highlight_cells(row=0, column=display_idx,
                                                bg=h_bg, fg=h_fg, canvas="header")
                except Exception:
                    pass

            self.sheet.redraw()
        except Exception as e:
            print(f"_apply_header_filters error: {e}")

    def _clear_all_header_filters(self):
        self._header_filter_values = {}
        self._apply_header_filters()

    # ================================================================== #
    def _rebind_sheet(self):
        self.sheet.bind("<Control-MouseWheel>", self._on_ctrl_scroll)
        self.sheet.bind("<Shift-MouseWheel>", self._lock_horizontal_scroll)
        self.sheet.bind("<Return>", self._on_enter_pressed)
        self.sheet.bind("<KP_Enter>", self._on_enter_pressed)
        self.sheet.bind("<Control-c>", lambda e: self.sheet.copy())
        self.sheet.bind("<Control-v>", lambda e: self.sheet.paste())
        self.sheet.bind("<Control-C>", lambda e: self.sheet.copy())
        self.sheet.bind("<Control-V>", lambda e: self.sheet.paste())
        self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)

        # ← Copy/Paste ทั้งบรรทัด ใช้ Ctrl+R copy, Ctrl+V paste
        self.sheet.bind("<Control-r>", self._copy_selected_rows)
        self.sheet.bind("<Control-R>", self._copy_selected_rows)

        try:
            self.sheet.MT.bind("<ButtonPress-1>", self._capture_click_pos, add="+")
            self.sheet.MT.bind("<ButtonPress-1>", self._on_mt_click, add="+")
        except Exception:
            self.sheet.bind("<ButtonPress-1>", self._capture_click_pos, add="+")

        try:
            self.sheet.MT.bind("<Control-c>", lambda e: self.sheet.copy(), add="+")
            self.sheet.MT.bind("<Control-C>", lambda e: self.sheet.copy(), add="+")
            self.sheet.MT.bind("<Control-v>", lambda e: self.sheet.paste(), add="+")
            self.sheet.MT.bind("<Control-V>", lambda e: self.sheet.paste(), add="+")
            self.sheet.MT.bind("<Control-r>", self._copy_selected_rows, add="+")
            self.sheet.MT.bind("<Control-R>", self._copy_selected_rows, add="+")
        except Exception:
            pass

        # ← bind RI (row index canvas) สำหรับ copy/paste ทั้งบรรทัด
        try:
            self.sheet.RI.bind("<Control-c>", self._copy_selected_rows, add="+")
            self.sheet.RI.bind("<Control-C>", self._copy_selected_rows, add="+")
            self.sheet.RI.bind("<Control-v>", self._paste_selected_rows, add="+")
            self.sheet.RI.bind("<Control-V>", self._paste_selected_rows, add="+")
            self.sheet.RI.bind("<Control-r>", self._copy_selected_rows, add="+")
            self.sheet.RI.bind("<Control-R>", self._copy_selected_rows, add="+")
        except Exception:
            pass

        self.sheet.enable_bindings((
            "single_select", "drag_select", "multi_select",
            "row_select", "column_select", "column_width_resize",
            "arrowkeys", "right_click_popup_menu",
            "rc_select", "copy", "cut", "paste",
            "delete", "undo", "edit_cell",
            "row_drag_and_drop",
        ))

        self.sheet.extra_bindings([
            ("cell_select", self._on_cell_select_combined),
            ("end_edit_cell", self._on_end_edit_combined),
            ("column_width_resize", lambda e: self._save_col_widths()),
            ("row_drag_and_drop", lambda e: self._on_row_drag_drop(e)),
        ])

        try:
            self.sheet.MT.bind("<ButtonRelease-1>",
                                lambda e: self.after(50, self._update_quick_calc), add="+")
        except Exception:
            pass


    def _rebind_frozen_sheet(self):
        if not self.sheet_frozen:
            return

        self.sheet_frozen.bind("<Control-MouseWheel>", self._on_ctrl_scroll)
        self.sheet_frozen.bind("<Control-Button-4>", lambda e: self._zoom(1))
        self.sheet_frozen.bind("<Control-Button-5>", lambda e: self._zoom(-1))

        # ← Copy/Paste ทั้งบรรทัด
        self.sheet_frozen.bind("<Control-r>", self._copy_selected_rows)
        self.sheet_frozen.bind("<Control-R>", self._copy_selected_rows)

        try:
            self.sheet_frozen.MT.bind("<Control-MouseWheel>", self._on_ctrl_scroll, add="+")
            self.sheet_frozen.MT.bind("<Control-Button-4>", lambda e: self._zoom(1), add="+")
            self.sheet_frozen.MT.bind("<Control-Button-5>", lambda e: self._zoom(-1), add="+")
            self.sheet_frozen.MT.bind("<Control-c>", lambda e: self.sheet_frozen.copy(), add="+")
            self.sheet_frozen.MT.bind("<Control-C>", lambda e: self.sheet_frozen.copy(), add="+")
            self.sheet_frozen.MT.bind("<Control-v>", lambda e: self.sheet_frozen.paste(), add="+")
            self.sheet_frozen.MT.bind("<Control-V>", lambda e: self.sheet_frozen.paste(), add="+")
            self.sheet_frozen.MT.bind("<Control-r>", self._copy_selected_rows, add="+")
            self.sheet_frozen.MT.bind("<Control-R>", self._copy_selected_rows, add="+")
        except Exception:
            pass

        # ← bind RI ของ frozen sheet ด้วย
        try:
            self.sheet_frozen.RI.bind("<Control-c>", self._copy_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-C>", self._copy_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-v>", self._paste_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-V>", self._paste_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-r>", self._copy_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-R>", self._copy_selected_rows, add="+")
        except Exception:
            pass

        self.sheet_frozen.enable_bindings((
            "single_select", "drag_select", "multi_select",
            "row_select", "column_select", "column_width_resize",
            "arrowkeys", "right_click_popup_menu",
            "rc_select", "copy", "cut", "paste",
            "delete", "undo", "edit_cell",
        ))

        self.sheet_frozen.extra_bindings([
            ("cell_select", lambda e: self._on_cell_select_combined(e, is_frozen=True)),
            ("end_edit_cell", lambda e: self._on_end_edit_combined(e, is_frozen=True)),
            ("column_width_resize", lambda e: self._save_col_widths())
        ])

        self.sheet_frozen.bind("<<SheetModified>>", self._on_sheet_modified)
        self.sheet_frozen.bind("<Control-c>", lambda e: self.sheet_frozen.copy())
        self.sheet_frozen.bind("<Control-v>", lambda e: self.sheet_frozen.paste())
        self.sheet_frozen.bind("<Control-C>", lambda e: self.sheet_frozen.copy())
        self.sheet_frozen.bind("<Control-V>", lambda e: self.sheet_frozen.paste())
        self.sheet_frozen.bind("<Return>", lambda e: self._on_enter_pressed(e, is_frozen=True))
        self.sheet_frozen.bind("<KP_Enter>", lambda e: self._on_enter_pressed(e, is_frozen=True))

        try:
            self.sheet_frozen.MT.bind("<ButtonPress-1>", self._capture_click_pos, add="+")
            self.sheet_frozen.MT.bind("<ButtonPress-1>", self._on_mt_click_frozen, add="+")
        except Exception:
            pass
    
    def _on_row_drag_drop(self, event=None):
        print(f"DEBUG drag drop event: {event}") 
        """หลัง drag row ใน main sheet → sync ข้อมูลไป frozen sheet ด้วย"""
        try:
            # ดึงข้อมูลใหม่จาก main sheet แล้ว rebuild frozen ให้ตรงกัน
            if not self.sheet_frozen:
                # ไม่มี frozen — auto save แล้วจบ
                if self.auto_save_job_id:
                    self.after_cancel(self.auto_save_job_id)
                self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                return

            # มี frozen — ต้อง sync ข้อมูลใหม่ทั้งหมดจาก DB หรือ rebuild
            # วิธีง่ายสุดคือโหลด data จาก main sheet แล้วเขียนลง frozen ใหม่
            main_data = self.sheet.get_sheet_data()
            total_rows = min(len(main_data), self.sheet_frozen.get_total_rows())

            # ← frozen sheet ยังเก็บข้อมูลเดิม ต้อง rebuild ตาม order ใหม่ของ main
            # แต่ frozen ไม่รู้ว่า row ไหนย้ายไปไหน ต้อง full reload จาก DB
            self.after(200, self._reload_frozen_from_main)

            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

        except Exception as e:
            print(f"_on_row_drag_drop error: {e}")

    def _reload_frozen_from_main(self):
        """sync ข้อมูล frozen sheet ให้ตรงกับ main sheet หลัง drag"""
        if not self.sheet_frozen:
            return
        try:
            main_data = self.sheet.get_sheet_data()
            col_offset = self.frozen_col_count
            # ← ถ้า frozen มีข้อมูลน้อยกว่า main ให้เพิ่ม row ก่อน
            frozen_rows = self.sheet_frozen.get_total_rows()
            main_rows = len(main_data)
            if main_rows > frozen_rows:
                self.sheet_frozen.insert_rows(main_rows - frozen_rows)

            # เขียนข้อมูลจาก full_row (main+frozen merged) ลง frozen
            # ปัญหาคือ main sheet ไม่มีข้อมูล frozen cols
            # ต้องใช้ _save_to_db แล้ว _load_from_db แทน
            # แต่นั่น disruptive เกินไป → ใช้วิธี swap frozen rows ตาม order ใหม่

            # วิธีที่ดีที่สุด: rebuild ทั้งหมดจาก current state
            self._rebuild_frozen_layout()
        except Exception as e:
            print(f"_reload_frozen_from_main error: {e}")

    def _on_mt_click_frozen(self, event=None):
        """เหมือน _on_mt_click แต่สำหรับตารางตรึง"""
        try:
            try:
                row = self.sheet_frozen.MT.identify_row(y=event.y, allow_end=False)
                col = self.sheet_frozen.MT.identify_col(x=event.x, allow_end=False)
            except Exception:
                try:
                    row = self.sheet_frozen.MT.get_row_at_y(event.y)
                    col = self.sheet_frozen.MT.get_col_at_x(event.x)
                except Exception:
                    return

            if row is None or col is None:
                return

            row, col = int(row), int(col)
            # frozen ไม่บวก offset เพราะ index เริ่มจาก 0 อยู่แล้ว
            real_col = col
            if real_col >= len(self.columns):
                return

            col_name = self.columns[real_col]
            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }

            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and self._last_popup_cell == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row, _col = row, col

            try:
                current_val = str(self.sheet_frozen.get_cell_data(_row, _col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet_frozen.set_cell_data(_row, _col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet_frozen.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col, is_frozen=True))
                except Exception as ex:
                    print(f"on_select frozen error: {ex}")

            self.after(50, lambda: self._open_popup_delayed(_row, _col, data_list, current_val, on_select, is_frozen=True))

        except Exception as e:
            print(f"_on_mt_click_frozen error: {e}")

    def _refresh_dropdown_data(self):
        """โหลด dropdown data ใหม่จาก DB"""
        self.sales_list = []
        self.supplier_list = []
        self.product_list = []
        self.product_sku_map = {}
        self.supplier_code_map = {}
        self.product_category_map = {}
        self._load_dropdown_data()

    def _on_mt_click(self, event=None):
        try:
            # tksheet เก็บ row/col position ใน MT canvas โดยตรง
            # ใช้ฟังก์ชันของ MT แทน Sheet
            try:
                row = self.sheet.MT.identify_row(y=event.y, allow_end=False)
                col = self.sheet.MT.identify_col(x=event.x, allow_end=False)
            except Exception:
                try:
                    # fallback สำหรับ version เก่ากว่า
                    row = self.sheet.MT.get_row_at_y(event.y)
                    col = self.sheet.MT.get_col_at_x(event.x)
                except Exception:
                    return

            if row is None or col is None:
                return

            row, col = int(row), int(col)
            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = col + col_offset
            if real_col >= len(self.columns):
                return

            col_name = self.columns[real_col]
            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }

            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and self._last_popup_cell == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row, _col = row, col

            try:
                current_val = str(self.sheet.get_cell_data(_row, _col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            self.after(50, lambda: self._open_popup_delayed(_row, _col, data_list, current_val, on_select))

        except Exception as e:
            print(f"_on_mt_click error: {e}")

    def _start_dropdown_refresh_timer(self):
        """Refresh dropdown data ทุก 10 วินาที"""
        self._refresh_dropdown_data()
        self.after(10000, self._start_dropdown_refresh_timer)

    def _on_cell_select_combined(self, event=None, is_frozen=False):
        self._on_sheet_click_for_formula(event)

        try:
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None: return
            row, col = int(row), int(col)

            # 🛠️ ถ้าคลิกที่ตารางตรึง ไม่ต้องบวก Col Offset 
            col_offset = 0 if is_frozen else (self.frozen_col_count if self.sheet_frozen is not None else 0)
            real_col = col + col_offset
            if real_col >= len(self.columns): return
            
            col_name = self.columns[real_col]
            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }

            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            if (self._active_popup is not None and not self._active_popup._destroyed and self._last_popup_cell == (row, col)):
                try: self._active_popup.entry.focus_set()
                except: pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except: pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row, _col = row, col
            
            # 🎯 เลือกเป้าหมายดึงข้อมูลให้ถูกตาราง
            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            try:
                current_val = str(target_sheet.get_cell_data(_row, _col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    target_sheet.set_cell_data(_row, _col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    target_sheet.redraw()
                    if self.auto_save_job_id is not None: self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'): self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col, is_frozen))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            self._last_popup_cell = (_row, _col)
            self.after(80, lambda r=_row, c=_col, dl=data_list, cv=current_val, os=on_select, fz=is_frozen: 
                self._open_popup_delayed(r, c, dl, cv, os, fz))
        except Exception as e:
            print(f"_on_cell_select_combined error: {e}")

    def _open_popup_delayed(self, row, col, data_list, current_val, on_select, is_frozen=False):
        try:
            if (self._active_popup is not None and not self._active_popup._destroyed and self._last_popup_cell == (row, col)):
                try: self._active_popup.entry.focus_set()
                except: pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed: self._active_popup.safe_destroy()
                except: pass
                self._active_popup = None

            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            try: target_sheet.close_text_editor(set_data=False)
            except: pass

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_near_cell(
                sheet_widget=target_sheet, # <-- ใช้ตารางเป้าหมายที่ถูกต้อง
                row=row,
                col=col,
                row_h=self.zoom_level + 19,
                header_h=self.zoom_level + 24,
            )

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (row, col)

        except Exception as e:
            print(f"_open_popup_delayed error: {e}")
            
    def _on_begin_edit_block(self, event=None):
        try:
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = int(col) + col_offset
            if real_col >= len(self.columns):
                return

            col_name = self.columns[real_col]
            if col_name not in {"รายการสินค้า", "ชื่อ Supplier", "รหัส Sale"}:
                return

            # กัน loop — ถ้า popup เปิดอยู่แล้วในเซลล์นี้ ไม่ต้องทำอะไร
            _row, _col = int(row), int(col)
            if getattr(self, '_popup_opening', False):
                return
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (_row, _col)):
                return

            self._popup_opening = True

            def _close_and_popup():
                self._popup_opening = False
                try:
                    self.sheet.close_text_editor(set_data=False)
                except Exception:
                    pass
                try:
                    self.sheet.cancel_edit()
                except Exception:
                    pass
                self._on_cell_select_popup({"row": _row, "column": _col})

            self.after(0, _close_and_popup)

        except Exception as e:
            self._popup_opening = False
            print(f"_on_begin_edit_block error: {e}")

    def _on_cell_select_popup(self, event=None):
        """
        Triggered by tksheet extra_binding 'cell_select' — API มาตรฐาน ทำงานได้ทุก version
        รองรับทั้ง dict, tuple, และ object ที่มี .row / .column attribute
        """
        try:
            # แกะ row, col จาก event ที่ tksheet ส่งมา (format อาจต่างกันตาม version)
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None:
                return
            row, col = int(row), int(col)
            if row < 0 or col < 0:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = col + col_offset
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }

            # ถ้าคลิกคอลัมน์อื่น → ปิด popup ที่เปิดอยู่แล้วออก
            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            # กัน popup เปิดซ้ำในเซลล์เดิม — focus กลับไปที่ entry แทน
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            # ปิด popup เก่าก่อน
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            # ยกเลิก inline editor ของ tksheet ถ้ากำลังจะเปิด
            try:
                self.sheet.close_text_editor(set_data=False)
            except Exception:
                pass

            data_list = popup_cols[col_name]
            _row, _col = row, col

            try:
                current_val = str(self.sheet.get_cell_data(_row, _col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_at_mouse(ref_widget=self.sheet)

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (_row, _col)

        except Exception as e:
            import traceback
            print(f"_on_cell_select_popup error: {e}")
            traceback.print_exc()


    def _on_sheet_mouse_click(self, event=None):
        """ดักจาก mouse click โดยตรง — ใช้ identify_region หาว่าคลิกเซลล์ไหน"""
        try:
            if event is None:
                return

            # ใช้ identify_region หรือ identify_cell ของ tksheet
            try:
                region = self.sheet.identify_region(event)
                if region != "table":
                    return  # คลิก header, scrollbar, ฯลฯ → ไม่ทำอะไร
            except Exception:
                pass

            try:
                row, col = self.sheet.identify_cell(event)
            except Exception:
                # fallback: คำนวณเองจาก pixel
                try:
                    row = self.sheet.identify_row(event)
                    col = self.sheet.identify_col(event)
                except Exception:
                    return

            if row is None or col is None:
                return

            row, col = int(row), int(col)
            if row < 0 or col < 0:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = col + col_offset
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            # กัน popup เปิดซ้ำในเซลล์เดิม
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            # ปิด popup เก่า
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row, _col = row, col

            try:
                current_val = str(self.sheet.get_cell_data(_row, _col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_at_mouse(ref_widget=self.sheet)

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (_row, _col)

        except Exception as e:
            import traceback
            print(f"_on_sheet_mouse_click error: {e}")
            traceback.print_exc()

    def _debug_cell_select(self, event=None):
        print(f"DEBUG cell_select event type: {type(event)}")
        print(f"DEBUG cell_select event value: {repr(event)}")

    def _on_cell_single_click_popup(self, event=None):
        """Single click → เปิด popup ทันที"""
        try:
            # ✅ แก้: รองรับ None value ใน dict
            row, col = None, None

            if isinstance(event, dict):
                r = event.get('row')
                c = event.get('column')
                if r is None or c is None:
                    return  # tksheet ส่ง event ที่ยังไม่มีข้อมูล เช่นตอน deselect
                row, col = int(r), int(c)
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                if event[0] is None or event[1] is None:
                    return
                row, col = int(event[0]), int(event[1])
            else:
                r = getattr(event, 'row', None)
                c = getattr(event, 'column', None)
                if r is None or c is None:
                    return
                row, col = int(r), int(c)

            if row < 0 or col < 0:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = col + col_offset
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                # ถ้าคลิกคอลัมน์อื่น → ปิด popup ที่เปิดอยู่
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            # กัน popup เปิดซ้ำในเซลล์เดิม
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and self._last_popup_cell == (row, col)):
                # popup เปิดอยู่แล้วในเซลล์นี้ — focus ไปที่ entry
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            # ปิด popup เก่า
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]

            try:
                current_val = str(self.sheet.get_cell_data(row, col) or "").strip()
            except Exception:
                current_val = ""

            # capture ค่า row, col ใน closure
            _row, _col = row, col

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_near_cell(
                sheet_widget=self.sheet,
                row=row,
                col=col,
                row_h=self.zoom_level + 19,
                header_h=self.zoom_level + 24,
            )

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (row, col)

        except Exception as e:
            import traceback
            print(f"_on_cell_single_click_popup error: {e}")
            traceback.print_exc()
    # ================================================================== #
    # POPUP SEARCH
    # ================================================================== #
    def _on_begin_edit_for_popup(self, event=None):
        """ดักตอน tksheet กำลังจะเปิด inline editor — ถ้าเป็นคอลัมน์ popup ให้เปิด popup แทน"""
        try:
            # ปิด popup เก่าก่อน
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            # แกะ row, col
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = col + col_offset
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                return  # ปล่อย tksheet ทำงานปกติ

            # ยกเลิก inline editor ของ tksheet
            try:
                self.sheet.close_text_editor(set_data=False)
            except Exception:
                pass
            try:
                self.sheet._reset_text_editor()
            except Exception:
                pass

            data_list = popup_cols[col_name]

            # ดึงค่าเดิมในเซลล์มา pre-fill
            try:
                current_val = str(self.sheet.get_cell_data(row, col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(row, col, value, redraw=True)
                    self._auto_calculate_sheet(row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self.after(50, lambda: self._move_right(row, col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            # สร้าง popup (ไม่มี anchor แล้ว — position จัดการใน place_near_cell)
            popup = InlineSearchPopup(self, data_list, on_select)

            # วางตำแหน่งใต้เซลล์ที่คลิก
            popup.place_near_cell(
                sheet_widget=self.sheet,
                row=row,
                col=col,
                row_h=self.zoom_level + 19,   # ซิงค์กับ row_height ที่ set ไว้
                header_h=self.zoom_level + 24, # ซิงค์กับ header_height ที่ set ไว้
            )

            # pre-fill ค่าเดิม
            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup

        except Exception as e:
            import traceback
            print(f"_on_begin_edit_for_popup error: {e}")
            traceback.print_exc()

    # ================================================================== #
    def _save_col_widths(self):
        try:
            if self.sheet_frozen:
                for i in range(self.frozen_col_count):
                    w = self.sheet_frozen.column_width(i)
                    if w and w > 0:
                        self.col_widths_cache[i] = w
                for i in range(self.sheet.get_total_columns()):
                    w = self.sheet.column_width(i)
                    if w and w > 0:
                        self.col_widths_cache[self.frozen_col_count + i] = w
            else:
                for i in range(self.sheet.get_total_columns()):
                    w = self.sheet.column_width(i)
                    if w and w > 0:
                        self.col_widths_cache[i] = w
            # ← debounce save ลง DB
            if hasattr(self, '_save_settings_job'):
                try: self.after_cancel(self._save_settings_job)
                except: pass
            self._save_settings_job = self.after(2000, self._save_user_settings)
        except Exception:
            pass

    def _update_quick_calc(self, event=None):
        """คำนวณ Sum/Count/Avg จากเซลล์ที่เลือก เหมือน Excel"""
        try:
            selected = self.sheet.get_selected_cells()
            if not selected:
                self.quick_calc_label.configure(text="")
                return

            values = []
            for r, c in selected:
                try:
                    val = self.sheet.get_cell_data(r, c)
                    if val and str(val).strip():
                        num = float(str(val).replace(',', '').replace('%', '').strip())
                        values.append(num)
                except (ValueError, TypeError):
                    pass

            count = len(selected)
            num_count = len(values)

            if num_count == 0:
                self.quick_calc_label.configure(text=f"จำนวน: {count}")
                return

            total = sum(values)
            avg = total / num_count

            def fmt(n):
                if n == int(n):
                    return f"{int(n):,}"
                return f"{n:,.2f}"

            text = f"จำนวน: {count}    ผลรวม: {fmt(total)}    เฉลี่ย: {fmt(avg)}"
            self.quick_calc_label.configure(text=text)
        except Exception:
            pass

    def _restore_col_widths(self):
        try:
            for i, w in self.col_widths_cache.items():
                self.sheet.column_width(i, w)
            self.sheet.redraw()
        except Exception:
            pass

    def _on_end_edit_combined(self, event=None, is_frozen=False):
        try:
            if event and len(event) >= 2:
                row, col = event[0], event[1]
                self.after(10, lambda: self._move_right(row, col, is_frozen))
        except Exception: pass

    def _move_right(self, row, col, is_frozen=False):
        if is_frozen and self.sheet_frozen:
            total_cols_frozen = self.sheet_frozen.get_total_columns()
            next_col = col + 1
            if next_col < total_cols_frozen:
                self.sheet_frozen.deselect("all")
                self.sheet_frozen.select_cell(row, next_col)
                self.sheet_frozen.see(row, next_col)
            else:
                # ➡️ ถ้าสุดขอบตารางตรึง ให้เด้งข้ามไปตารางหลักฝั่งขวา!
                self.sheet_frozen.deselect("all")
                self.sheet.select_cell(row, 0)
                self.sheet.see(row, 0)
                self.sheet.focus_set()
        else:
            total_cols = self.sheet.get_total_columns()
            total_rows = self.sheet.get_total_rows()
            next_col = col + 1
            if next_col < total_cols:
                self.sheet.deselect("all")
                self.sheet.select_cell(row, next_col)
                self.sheet.see(row, next_col)
            else:
                if row + 1 < total_rows:
                    self.sheet.deselect("all")
                    # ↩️ ขึ้นบรรทัดใหม่ กระโดดกลับไปเริ่มที่ตารางตรึง (ถ้ามี)
                    if self.sheet_frozen:
                        self.sheet_frozen.select_cell(row + 1, 0)
                        self.sheet_frozen.see(row + 1, 0)
                        self.sheet_frozen.focus_set()
                    else:
                        self.sheet.select_cell(row + 1, 0)
                        self.sheet.see(row + 1, 0)

    def _on_enter_pressed(self, event=None, is_frozen=False):
        try:
            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            curr = target_sheet.get_currently_selected()
            if curr:
                self._move_right(curr[0], curr[1], is_frozen)
            return "break"
        except Exception: pass

    def _lock_horizontal_scroll(self, event):
        return "break"

    def _on_ctrl_scroll(self, event):
        if event.delta > 0:
            self._zoom(1)
        else:
            self._zoom(-1)
        return "break"

    def _move_right(self, row, col):
        total_cols = self.sheet.get_total_columns()
        total_rows = self.sheet.get_total_rows()
        next_col = col + 1
        if next_col < total_cols:
            self.sheet.deselect("all")
            self.sheet.select_cell(row, next_col)
            self.sheet.see(row, next_col)
        else:
            if row + 1 < total_rows:
                self.sheet.deselect("all")
                self.sheet.select_cell(row + 1, 0)
                self.sheet.see(row + 1, 0)

    def _on_enter_pressed(self, event=None):
        try:
            curr = self.sheet.get_currently_selected()
            if curr:
                self._move_right(curr[0], curr[1])
            return "break"
        except Exception:
            pass

    # ================================================================== #
    def _on_sheet_modified(self, event=None):
        try:
            # ← แก้: คำนวณเฉพาะ row ที่ถูก modify จริงๆ ไม่ใช่ทุก row
            if event and hasattr(event, 'cells'):
                rows_to_calc = set(r for r, c in event.cells)
            elif event and isinstance(event, dict):
                rows_to_calc = {event.get('row')} if event.get('row') is not None else set()
            else:
                # fallback — คำนวณทุก row ที่มีข้อมูล
                rows_to_calc = set()
                for r in range(self.sheet.get_total_rows()):
                    row_data = self.sheet.get_row_data(r)
                    if any(str(v).strip() for v in row_data):
                        rows_to_calc.add(r)
                if self.sheet_frozen:
                    for r in range(self.sheet_frozen.get_total_rows()):
                        row_data = self.sheet_frozen.get_row_data(r)
                        if any(str(v).strip() for v in row_data):
                            rows_to_calc.add(r)

            for row_idx in rows_to_calc:
                if row_idx is not None:
                    self._auto_calculate_sheet(row_idx)
            
            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
                
            if self.auto_save_job_id is not None:
                self.after_cancel(self.auto_save_job_id)
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        except Exception:
            pass


    def _auto_calculate_sheet(self, row_idx):
        col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0

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

        def get_val(col_name):
            try:
                real_idx = self.columns.index(col_name)
                display_idx = real_idx - col_offset
                if display_idx < 0:
                    val = self.sheet_frozen.get_cell_data(row_idx, real_idx) if self.sheet_frozen else None
                else:
                    val = self.sheet.get_cell_data(row_idx, display_idx)
                return float(str(val).replace(',', '').replace('%', '')) if val and str(val).strip() else 0.0
            except (ValueError, IndexError):
                return 0.0

        def get_str(col_name):
            try:
                real_idx = self.columns.index(col_name)
                display_idx = real_idx - col_offset
                if display_idx < 0:
                    # อยู่ในฝั่ง frozen — ใช้ real_idx โดยตรง
                    if self.sheet_frozen:
                        return str(self.sheet_frozen.get_cell_data(row_idx, real_idx) or "").strip()
                    return ""
                return str(self.sheet.get_cell_data(row_idx, display_idx) or "").strip()
            except (IndexError, ValueError):
                return ""

        def set_val(col_name, val, is_text=False):
            try:
                real_idx = self.columns.index(col_name)
                display_idx = real_idx - col_offset
                if not is_text:
                    formatted_val = "" if val == 0 else f"{val:,.2f}"
                else:
                    formatted_val = "" if (val is None or val == "%") else str(val)
                if display_idx < 0:
                    if self.sheet_frozen:
                        self.sheet_frozen.set_cell_data(row_idx, real_idx, formatted_val, redraw=False)
                else:
                    self.sheet.set_cell_data(row_idx, display_idx, formatted_val, redraw=False)
            except IndexError:
                pass

        date_idx = self.columns.index("วันที่ขอราคา")
        row_data = self.sheet.get_row_data(row_idx)
        is_row_active = any(str(cell_val).strip() for i, cell_val in enumerate(row_data) if i != (date_idx - col_offset) and (date_idx - col_offset) >= 0)
        if not is_row_active and self.sheet_frozen:
            frozen_row = self.sheet_frozen.get_row_data(row_idx)
            is_row_active = any(str(v).strip() for i, v in enumerate(frozen_row) if i != date_idx)

        current_date = get_str("วันที่ขอราคา")
        if is_row_active and not current_date:
            now = datetime.now()
            thai_year = (now.year + 543) % 100
            set_val("วันที่ขอราคา", f"{now.day:02d}/{now.month:02d}/{thai_year}", is_text=True)
        elif not is_row_active and current_date:
            set_val("วันที่ขอราคา", "", is_text=True)

        product_name = get_str("รายการสินค้า")
        if product_name in self.product_sku_map:
            if get_str("Product SKU.") != self.product_sku_map[product_name]:
                set_val("Product SKU.", self.product_sku_map[product_name], is_text=True)
            mapped_category = self.product_category_map.get(product_name, "")
            if get_str("หมวด") != mapped_category:
                set_val("หมวด", mapped_category, is_text=True)
        elif not product_name:
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

        for col_percent in ["WIN RATE %", "Markup Guide (%)"]:
            val_str = get_str(col_percent)
            if val_str and not val_str.endswith("%"):
                num_str = "".join([c for c in val_str if c.isdigit() or c == '.'])
                if num_str: set_val(col_percent, f"{num_str}%", is_text=True)

        qty = get_val("จำนวน")
        weight_per_unit = get_val("น้ำหนัก/เส้น")
        cost_per_unit = get_val("ต้นทุน/เส้น")

        if qty == 0 and weight_per_unit == 0 and cost_per_unit == 0:
            auto_cols_to_clear = [
                "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1",
                "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2", "ต้นทุน/กก. (ไม่รวมย้าย)",
                "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย/เส้น",
                "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)",
                "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup",
                "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น",
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

        cost_no_move_per_kg = (cost_after_d2 / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ต้นทุน/เส้น (ไม่รวมย้าย)", cost_after_d2)
        set_val("ต้นทุน/กก. (ไม่รวมย้าย)", cost_no_move_per_kg)
        set_val("ต้นทุนรวม (ไม่รวมย้าย)", cost_after_d2 * qty)

        moving_cost = get_val("ค่าย้าย (ซื้อ)")
        moving_cost_per_unit = (moving_cost / qty) if qty > 0 else 0
        set_val("ค่าย้าย/เส้น", moving_cost_per_unit)

        cost_with_move = cost_after_d2 + moving_cost_per_unit
        cost_with_move_per_kg = (cost_with_move / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ต้นทุน/เส้น (รวมย้าย)", cost_with_move)
        set_val("ต้นทุน/กก. (รวมย้าย)", cost_with_move_per_kg)
        set_val("ต้นทุนรวม (รวมย้าย)", cost_with_move * qty)

        markup_pct = get_val("Markup Guide (%)") / 100.0
        set_val("Markup/กก.", cost_with_move_per_kg * markup_pct)
        set_val("Markup/เส้น", cost_with_move * markup_pct)

        cost_markup_per_kg = cost_with_move_per_kg + cost_with_move_per_kg * markup_pct
        cost_markup_per_unit = cost_with_move + cost_with_move * markup_pct
        set_val("ทุน+Markup/กก.", cost_markup_per_kg)
        set_val("ทุน+Markup/เส้น", cost_markup_per_unit)
        set_val("ต้นทุนรวม+Markup", cost_markup_per_unit * qty)

        shipping_sell = get_val("ค่าส่ง (ขาย)")
        shipping_per_unit = (shipping_sell / qty) if qty > 0 else 0
        set_val("ค่าส่ง / เส้น", shipping_per_unit)
        set_val("น้ำหนัก/เส้น 2", weight_per_unit)

        sell_price = cost_markup_per_unit + shipping_per_unit
        sell_price_per_kg = (sell_price / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ราคาขาย / เส้น", sell_price)
        set_val("ราคาขาย / กก.", sell_price_per_kg)

        vat_per_unit = sell_price * 0.07
        set_val("Vat. / เส้น", vat_per_unit)
        set_val("ราคาขาย/เส้น + Vat.", sell_price * 1.07)

        sell_total = sell_price * qty
        set_val("ราคาขาย รวม", sell_total)
        set_val("Vat. รวม", sell_total * 0.07)
        set_val("ราคาขาย รวม + Vat.", sell_total * 1.07)

    # ================================================================== #
    def _add_new_row(self):
        if not HAS_TKSHEET:
            return
        self.sheet.insert_row([""] * (len(self.columns) - self.frozen_col_count))
        self._apply_formatting(col_offset=self.frozen_col_count)
        total_rows = self.sheet.get_total_rows()
        self.sheet.see(total_rows - 1, 0)
        self.sheet.select_cell(total_rows - 1, 0)
        self.sheet.redraw()
        if self.sheet_frozen is not None:
            self.sheet_frozen.insert_row([""] * self.frozen_col_count)
            self.sheet_frozen.redraw()

    def _insert_selected_row(self):
        if not HAS_TKSHEET:
            return

        insert_idx = None
        
        try:
            curr = self.sheet.get_currently_selected()
            if curr:
                insert_idx = int(curr[0])
        except Exception:
            pass

        if insert_idx is None:
            try:
                selected_rows = self.sheet.get_selected_rows()
                if selected_rows:
                    insert_idx = min(int(r) for r in selected_rows)
            except Exception:
                pass

        if insert_idx is None:
            try:
                selected_cells = self.sheet.get_selected_cells()
                if selected_cells:
                    insert_idx = min(int(r) for r, c in selected_cells)
            except Exception:
                pass

        if insert_idx is None:
            if self.sheet_frozen:
                try:
                    curr = self.sheet_frozen.get_currently_selected()
                    if curr:
                        insert_idx = int(curr[0])
                except Exception:
                    pass

        if insert_idx is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการแทรกก่อน", parent=self)
            return

        empty_main_row = [""] * (len(self.columns) - self.frozen_col_count)
        self.sheet.insert_row(empty_main_row, idx=insert_idx)

        if self.sheet_frozen is not None:
            empty_frozen_row = [""] * self.frozen_col_count
            self.sheet_frozen.insert_row(empty_frozen_row, idx=insert_idx)
            self.sheet_frozen.redraw()

        self._apply_formatting(col_offset=self.frozen_col_count)
        self.sheet.see(insert_idx, 0)
        self.sheet.deselect("all")  # ← เพิ่ม: ล้าง selection ให้ user คลิกเองใหม่
        if self.sheet_frozen:
            self.sheet_frozen.deselect("all")
        self.sheet.redraw()
        if self.sheet_frozen:
            self.after(100, self._sync_vertical_scroll)
    
    def _copy_selected_rows(self, event=None):
        """Copy ทั้ง row รวม frozen cols"""
        try:
            # เก็บ rows ที่ selected จาก main sheet
            selected = self.sheet.get_selected_rows()
            if not selected and self.sheet_frozen:
                selected = set()
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    selected = {int(curr[0])}
            if not selected:
                return
            
            rows_data = []
            for r in sorted(selected):
                frozen_part = []
                main_part = list(self.sheet.get_row_data(r))
                if self.sheet_frozen:
                    frozen_part = list(self.sheet_frozen.get_row_data(r))
                rows_data.append(frozen_part + main_part)
            
            self._clipboard_rows = rows_data
            self.save_status_label.configure(
                text=f"📋 คัดลอก {len(rows_data)} บรรทัดแล้ว (Ctrl+V เพื่อวาง)",
                text_color="#6366F1"
            )
        except Exception as e:
            print(f"_copy_selected_rows error: {e}")

    def _paste_selected_rows(self, event=None):
        """Paste row ที่ copy ไว้ลงที่ row ที่เลือกอยู่"""
        try:
            if not hasattr(self, '_clipboard_rows') or not self._clipboard_rows:
                return
            
            # หา target row
            target_idx = None
            curr = self.sheet.get_currently_selected()
            if curr:
                target_idx = int(curr[0])
            elif self.sheet_frozen:
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    target_idx = int(curr[0])
            
            if target_idx is None:
                messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดปลายทางก่อน Ctrl+V", parent=self)
                return
            
            col_offset = self.frozen_col_count
            for i, row_data in enumerate(self._clipboard_rows):
                r = target_idx + i
                if r >= self.sheet.get_total_rows():
                    break
                # เขียนส่วน frozen
                if self.sheet_frozen and col_offset > 0:
                    for c in range(col_offset):
                        val = row_data[c] if c < len(row_data) else ""
                        self.sheet_frozen.set_cell_data(r, c, val, redraw=False)
                # เขียนส่วน main
                for c in range(self.sheet.get_total_columns()):
                    real_c = c + col_offset
                    val = row_data[real_c] if real_c < len(row_data) else ""
                    self.sheet.set_cell_data(r, c, val, redraw=False)
                self._auto_calculate_sheet(r)
            
            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
            
            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
            
        except Exception as e:
            print(f"_paste_selected_rows error: {e}")

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
            if self.sheet_frozen:
                self.sheet_frozen.delete_rows(list(selected_rows))
            self.sheet.redraw()

    # ================================================================== #
    def _load_from_db(self, *args):
        if not HAS_TKSHEET: return
        month_val = self.month_var.get()
        year_val = self.year_var.get()
        conn = self.app_container.get_connection()
        try:
            columns_sql = ", ".join([f'"{col.replace("%", "%%")}"' for col in self.columns])
            query = f'SELECT {columns_sql} FROM cost_benchmarks WHERE benchmark_month = %s AND benchmark_year = %s AND created_by = %s ORDER BY id ASC'
            df = pd.read_sql_query(query, conn, params=(month_val, year_val, self.current_user))

            self.sheet.bind("<<SheetModified>>", lambda e: None)
            total_cols_main = len(self.columns) - self.frozen_col_count
            col_offset = self.frozen_col_count

            if df.empty:
                data_list = [[""] * len(self.columns) for _ in range(1000)]
            else:
                df = df.fillna("")
                data_list = df.values.tolist()
                while len(data_list) < 1000:
                    data_list.append([""] * len(self.columns))

            # ปรับจำนวนแถวให้พอดี
            current_rows = self.sheet.get_total_rows()
            if len(data_list) > current_rows:
                self.sheet.insert_rows(len(data_list) - current_rows)
                if self.sheet_frozen:
                    self.sheet_frozen.insert_rows(len(data_list) - current_rows)

            # ← เขียนข้อมูลลงเซลล์จริงๆ (ส่วนนี้หายไป)
            for r, row in enumerate(data_list):
                if self.sheet_frozen and col_offset > 0:
                    for c in range(col_offset):
                        val_str = str(row[c]).strip() if c < len(row) and row[c] is not None else ""
                        self.sheet_frozen.set_cell_data(r, c, val_str, redraw=False)
                for c in range(total_cols_main):
                    real_c = c + col_offset
                    val_str = str(row[real_c]).strip() if real_c < len(row) and row[real_c] is not None else ""
                    self.sheet.set_cell_data(r, c, val_str, redraw=False)

            # ล้างแถวส่วนเกิน
            for r in range(len(data_list), self.sheet.get_total_rows()):
                for c in range(total_cols_main):
                    self.sheet.set_cell_data(r, c, "", redraw=False)
                if self.sheet_frozen:
                    for c in range(col_offset):
                        self.sheet_frozen.set_cell_data(r, c, "", redraw=False)

            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()

        except Exception as e:
            messagebox.showerror("Error", f"โหลดข้อมูลล้มเหลว: {e}", parent=self)
        finally:
            self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
            self.after(300, self._save_col_widths)
            if conn: self.app_container.release_connection(conn)
    
    def _hide_frozen_scrollbars(self):
        """ซ่อน scrollbar ทั้งหมดของ frozen sheet และ sync row height"""
        if not self.sheet_frozen:
            return
        try:
            # วิธีที่แรง — resize frozen sheet ให้สูงเต็ม relheight แต่ลบพื้นที่ scrollbar ออก
            # โดยการ configure internal frame
            for attr in ['_y_scrollbar', 'yscrollbar', 'y_scrollbar',
                        '_x_scrollbar', 'xscrollbar', 'x_scrollbar']:
                try:
                    sb = getattr(self.sheet_frozen, attr, None)
                    if sb:
                        sb.grid_remove()
                        sb.place_forget()
                        sb.pack_forget()
                except Exception:
                    pass

            # ซ่อนผ่าน children
            for widget in self.sheet_frozen.winfo_children():
                if widget.winfo_class() == "Scrollbar":
                    widget.grid_remove()
                    try: widget.place_forget()
                    except: pass
                    try: widget.pack_forget()
                    except: pass

            # ใช้ tksheet API
            for canvas_name in ["x_scrollbar", "y_scrollbar", "top_left"]:
                try:
                    self.sheet_frozen.hide(canvas=canvas_name)
                except Exception:
                    pass

        except Exception:
            pass

        # sync row height
        try:
            rh = int(30 * (self.zoom_level / 11.0))
            total_rows = max(
                self.sheet_frozen.get_total_rows(),
                self.sheet.get_total_rows()
            )
            for r in range(total_rows):
                self.sheet_frozen.row_height(r, rh)
                self.sheet.row_height(r, rh)
            self.sheet_frozen.redraw()
            self.sheet.redraw()
        except Exception:
            pass

    def _save_to_db(self, show_msg=True):
        col_offset = self.frozen_col_count
        raw_main = self.sheet.get_sheet_data()
        raw_frozen = self.sheet_frozen.get_sheet_data() if self.sheet_frozen else []

        data = []
        n_rows = max(len(raw_main), len(raw_frozen))
        date_col_idx = self.columns.index("วันที่ขอราคา")  # ← เพิ่มบรรทัดนี้
        
        for i in range(n_rows):
            left = raw_frozen[i] if i < len(raw_frozen) else [""] * col_offset
            right = raw_main[i] if i < len(raw_main) else [""] * (len(self.columns) - col_offset)
            full_row = left + right
            
            # ← แก้ตรงนี้
            is_active = any(
                str(cell).strip() not in ("", "%", "None", "nan")
                for j, cell in enumerate(full_row)
                if j != date_col_idx
            )
            
            if is_active:
                data.append([str(cell).strip() if cell is not None else "" for cell in full_row])
    

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
                cursor.execute(
                    "DELETE FROM cost_benchmarks WHERE benchmark_month = %s AND benchmark_year = %s AND created_by = %s",
                    (month_val, year_val, self.current_user)
                )
                columns_sql = ", ".join([f'"{col.replace("%", "%%")}"' for col in self.columns]) + ", benchmark_month, benchmark_year, created_by"
                values = [tuple(row) + (month_val, year_val, self.current_user) for row in df.to_numpy()]
                insert_query = f"INSERT INTO cost_benchmarks ({columns_sql}) VALUES %s"
                psycopg2.extras.execute_values(cursor, insert_query, values)
            conn.commit()
            current_time = datetime.now().strftime("%H:%M:%S")
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text=f"✅ บันทึกล่าสุด: {current_time}", text_color="#16A34A")
            if show_msg:
                messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล {len(df)} รายการเรียบร้อยแล้ว!", parent=self)
        except Exception as e:
            if conn: conn.rollback()
            import traceback; traceback.print_exc()
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="❌ บันทึกผิดพลาด กรุณาลองใหม่", text_color="#DC2626")
            if show_msg:
                messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    # ================================================================== #
    # FORMULA BAR
    # ================================================================== #
    def _num2col(self, n):
        string = ""
        n += 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def _on_formula_focus_in(self, event):
        try:
            cells = self.sheet.get_selected_cells()
            if cells:
                self.target_formula_cell = list(cells)[0]
        except Exception:
            pass

    def _on_sheet_click_for_formula(self, event=None):
        try:
            current_text = self.formula_entry.get()
            if current_text.startswith("="):
                cells = self.sheet.get_selected_cells()
                if not cells: return
                row, col = list(cells)[0]
                cell_val = self.sheet.get_cell_data(row, col)
                if not cell_val or str(cell_val).strip() == "":
                    val_to_insert = "0"
                else:
                    val_to_insert = str(cell_val).replace(',', '').replace('%', '').strip()
                    try:
                        float(val_to_insert)
                    except ValueError:
                        val_to_insert = "0"
                self.formula_entry.insert(tk.END, val_to_insert)
                self.formula_entry.focus()
                self.formula_entry.icursor(tk.END)
        except Exception:
            pass

    def _apply_formula_from_bar(self, event=None):
        if not self.target_formula_cell:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกช่องปลายทางในตารางก่อนเริ่มพิมพ์สูตร")
            return
        try:
            row, col = self.target_formula_cell
            formula = self.formula_entry.get()
            self.sheet.set_cell_data(row, col, formula)
            self.formula_entry.delete(0, tk.END)
            self.target_formula_cell = None
            self.sheet.select_cell(row, col)
            self._auto_calculate_sheet(row)
        except Exception as e:
            messagebox.showerror("Error", f"สูตรผิดพลาด: {e}")

    # ================================================================== #
    # COLUMN HIDE / SHOW / COLOR
    # ================================================================== #
    def _get_real_col_indices(self):
        real_cols = set()
        col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0

        # 1. เช็คจาก Sheet หลัก
        try:
            selected_cols = self.sheet.get_selected_columns()
            if selected_cols:
                for c in selected_cols:
                    real_cols.add(c + col_offset)
        except Exception:
            pass

        try:
            selected_cells = self.sheet.get_selected_cells()
            if selected_cells:
                for r, c in selected_cells:
                    real_cols.add(c + col_offset)
        except Exception:
            pass

        # 2. เช็คจาก Sheet ที่ตรึงไว้ (ถ้ามี)
        if self.sheet_frozen:
            try:
                frozen_cols = self.sheet_frozen.get_selected_columns()
                if frozen_cols:
                    for c in frozen_cols:
                        real_cols.add(c)
            except Exception:
                pass

            try:
                frozen_cells = self.sheet_frozen.get_selected_cells()
                if frozen_cells:
                    for r, c in frozen_cells:
                        real_cols.add(c)
            except Exception:
                pass

        return real_cols

    def _hide_selected_columns(self):
        if not HAS_TKSHEET: return
        
        real_cols = self._get_real_col_indices()
        if not real_cols:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิก 'ช่องใดๆ' หรือ 'หัวคอลัมน์' ที่ต้องการซ่อนก่อน", parent=self)
            return

        # อัปเดตเก็บประวัติคอลัมน์ที่ซ่อนไว้ (Global Index)
        self.hidden_cols_list.extend(real_cols)
        self.hidden_cols_list = list(set(self.hidden_cols_list))

        # สั่งซ่อนใน Tksheet
        if self.sheet_frozen is not None:
            # แยกซ่อนคอลัมน์ฝั่งซ้าย(ที่ตรึง) และฝั่งขวา(ที่เลื่อนได้)
            frozen_hides = [c for c in self.hidden_cols_list if c < self.frozen_col_count]
            main_hides = [c - self.frozen_col_count for c in self.hidden_cols_list if c >= self.frozen_col_count]
            
            if frozen_hides:
                self.sheet_frozen.hide_columns(frozen_hides)
            if main_hides:
                self.sheet.hide_columns(main_hides)
        else:
            self.sheet.hide_columns(self.hidden_cols_list)

        self._save_user_settings()

        # 🧹 สิ่งสำคัญ: เคลียร์ Selection ทิ้ง ป้องกันการจำค่าไปใช้ครั้งต่อไป ("พาเพื่อนไปด้วย")
        self.sheet.deselect("all")
        if self.sheet_frozen:
            self.sheet_frozen.deselect("all")

        self.sheet.redraw()
        if self.sheet_frozen:
            self.sheet_frozen.redraw()

    def _show_all_columns(self):
        if not HAS_TKSHEET: return
        self.hidden_cols_list = []
        
        self.sheet.display_columns("all")
        if self.sheet_frozen is not None:
            self.sheet_frozen.display_columns("all")
            self.sheet_frozen.redraw()
            
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

        col_offset = self.frozen_col_count
        for c_idx in real_cols:
            if c_idx >= len(self.columns): continue
            b_bg = self._lighten_color(color_code, amount=0.85)
            display_idx = c_idx - col_offset
            try:
                if display_idx >= 0:
                    self.sheet.highlight_cells(row=0, column=display_idx, bg=color_code, fg="black", canvas="header")
                    self.sheet.highlight_columns(columns=[display_idx], bg=b_bg, fg="black", highlight_header=False)
                elif self.sheet_frozen:
                    self.sheet_frozen.highlight_cells(row=0, column=c_idx, bg=color_code, fg="black", canvas="header")
                    self.sheet_frozen.highlight_columns(columns=[c_idx], bg=b_bg, fg="black", highlight_header=False)
            except Exception:
                pass

        self._save_user_settings()
        self.sheet.redraw()
        if self.sheet_frozen:
            self.sheet_frozen.redraw()

    # ================================================================== #
    # ZOOM
    # ================================================================== #
    def _zoom(self, direction):
        if not HAS_TKSHEET: return
        
        old_zoom = self.zoom_level
        self.zoom_level = max(2, min(40, self.zoom_level + direction))

        if old_zoom == self.zoom_level:
            return

        scale_ratio = self.zoom_level / old_zoom
        new_row_height = int(30 * (self.zoom_level / 11.0))
        new_header_height = int(35 * (self.zoom_level / 11.0))

        opts = dict(
            font=("Tahoma", self.zoom_level, "normal"),
            header_font=("Tahoma", self.zoom_level, "bold"),
            row_height=new_row_height,
            header_height=new_header_height,
            auto_resize_columns=False,
        )

        self.sheet.set_options(**opts)
        if self.sheet_frozen:
            self.sheet_frozen.set_options(**opts)

        try:
            for i in range(self.sheet.get_total_columns()):
                current_w = self.sheet.column_width(i)
                if current_w:
                    self.sheet.column_width(i, int(current_w * scale_ratio))

            frozen_width = 40
            if self.sheet_frozen:
                for i in range(self.frozen_col_count):
                    current_w = self.sheet_frozen.column_width(i)
                    if current_w:
                        new_w = int(current_w * scale_ratio)
                        self.sheet_frozen.column_width(i, new_w)
                        frozen_width += new_w
                
                frozen_width += 4
                self.sheet_frozen.place(x=0, y=0, width=frozen_width, relheight=1.0)
                self.sheet.place(x=frozen_width, y=0, relwidth=1.0, relheight=1.0, width=-frozen_width)

                # ← เพิ่ม: update resize handler ให้ใช้ frozen_width ใหม่
                def _on_resize_after_zoom(event, fw=frozen_width):
                    try:
                        self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                    except Exception:
                        pass
                self.table_frame.bind("<Configure>", _on_resize_after_zoom)

        except Exception as e:
            print(f"Zoom resize error: {e}")

        self.sheet.redraw()
        if self.sheet_frozen:
            self.sheet_frozen.redraw()

        # ← เพิ่ม: sync row height ทุก row ให้ตรงกันหลัง zoom
        def _sync_after_zoom():
            try:
                if self.sheet_frozen:
                    rh = int(30 * (self.zoom_level / 11.0))
                    total_rows = max(
                        self.sheet.get_total_rows(),
                        self.sheet_frozen.get_total_rows()
                    )
                    for r in range(total_rows):
                        self.sheet_frozen.row_height(r, rh)
                        self.sheet.row_height(r, rh)
                    self.sheet_frozen.redraw()
                    self.sheet.redraw()
                    # ← ซ่อน scrollbar frozen อีกครั้งหลัง zoom
                    self._hide_frozen_scrollbars()
            except Exception:
                pass

        self.after(100, _sync_after_zoom)
        self.after(100, self._save_col_widths)
        
        pct = int((self.zoom_level / 11) * 100)
        if hasattr(self, 'zoom_label'):
            self.zoom_label.configure(text=f"{pct}%")