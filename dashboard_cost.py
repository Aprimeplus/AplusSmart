import tkinter as tk
from tkinter import messagebox
from customtkinter import (
    CTkFrame, CTkLabel, CTkFont, CTkButton,
    CTkOptionMenu, CTkScrollableFrame, CTkComboBox  # 🟢 เพิ่ม CTkComboBox เข้ามา
)
import pandas as pd
import psycopg2.extras
from datetime import datetime

try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False


# ============================================================
# COLOR PALETTE (Power BI style)
# ============================================================
COLORS = {
    "bg_main":         "#F3F4F6",
    "bg_sidebar":      "#FFFFFF",
    "bg_white":        "#FFFFFF",
    "header_blue":     "#1D4ED8",   # dark-blue table header
    "header_text":     "#FFFFFF",
    "row_alt":         "#F8FAFC",
    "row_normal":      "#FFFFFF",
    "highlight_green": "#D1FAE5",
    "highlight_green_fg": "#065F46",
    "highlight_red":   "#FEE2E2",
    "highlight_red_fg": "#991B1B",
    "kpi_blue":        "#1D4ED8",
    "kpi_green":       "#047857",
    "kpi_red":         "#B91C1C",
    "kpi_purple":      "#6D28D9",
    "border":          "#E5E7EB",
    "text_dark":       "#111827",
    "text_medium":     "#4B5563",
    "text_light":      "#9CA3AF",
    "filter_bg":       "#F9FAFB",
    "filter_border":   "#D1D5DB",
    "btn_blue":        "#2563EB",
    "btn_hover":       "#1D4ED8",
    "sidebar_header":  "#1E3A5F",
}

FONT_FAMILY = "Tahoma"


class DashboardCostScreen(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color=COLORS["bg_main"])
        self.app_container = app_container
        self.current_user = getattr(self.app_container, 'current_user_key', 'PU_Default')

        self.grid_columnconfigure(0, weight=0, minsize=220)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.raw_df = pd.DataFrame()
        self.all_order_nos = [] 
        self.all_sale_order_nos = [] # 🟢 เพิ่มตัวแปรเก็บรายชื่อ Sale Order No. ทั้งหมด

        self._build_sidebar()
        self._build_main_content()

        self.after(200, self._load_data_from_db)

    # =========================================================
    # SIDEBAR
    # =========================================================
    def _build_sidebar(self):
        sidebar = CTkScrollableFrame(self, fg_color=COLORS["bg_sidebar"], corner_radius=0,
                                     border_width=0)
        sidebar.grid(row=0, column=0, sticky="nsew")

        # Sidebar header band
        header_band = CTkFrame(sidebar, fg_color=COLORS["sidebar_header"], corner_radius=0, height=50)
        header_band.pack(fill="x", pady=(0, 10))
        header_band.pack_propagate(False)
        CTkLabel(header_band, text="  ⚙  Filters",
                 font=CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
                 text_color="#FFFFFF").pack(side="left", padx=10, pady=12)

        # 🟢 เพิ่มตัวแปร sale_order_no เข้ามาใน Dictionary
        self.filter_vars = {
            "order_no":      tk.StringVar(value="All"),
            "sale_order_no": tk.StringVar(value="All"), # 🟢 เพิ่ม Sale Order No.
            "pu_user":       tk.StringVar(value="All"),
            "sale_name":     tk.StringVar(value="All"),
            "supplier":      tk.StringVar(value="All"),
            "status":        tk.StringVar(value="All"),
            "priority":      tk.StringVar(value="All"),
            "select":        tk.StringVar(value="All"),
            "year":          tk.StringVar(value="All"),
            "month":         tk.StringVar(value="All"),
        }
        self.filter_menus = {}

        # 🟢 ตั้งค่าป้ายกำกับและตัวแปรผูกมัด (จัด Sale Order No. ให้อยู่ด้านบนคู่กัน)
        filters_config = [
            ("🔍 Order No. (พิมพ์ค้นหา)", "order_no"),
            ("🔍 Sale Order No. (พิมพ์ค้นหา)", "sale_order_no"), # 🟢 เพิ่มเมนู Sale Order No.
            ("ผู้ทำตาราง (PU User)", "pu_user"),
            ("ชื่อ Sale",        "sale_name"),
            ("ชื่อ Supplier",    "supplier"),
            ("สถานะ",            "status"),
            ("PRIORITY",        "priority"),
            ("Select",          "select"),
            ("ปี (Year)",        "year"),
            ("เดือน (Month)",     "month"),
        ]

        for label_text, key in filters_config:
            lbl = CTkLabel(sidebar, text=label_text,
                           font=CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                           text_color=COLORS["text_medium"])
            lbl.pack(anchor="w", padx=12, pady=(8, 1))

            # 🟢 ให้ Order No. และ Sale Order No. ใช้ CTkComboBox เพื่อให้พิมพ์ค้นหาได้
            if key in ["order_no", "sale_order_no"]:
                menu = CTkComboBox(
                    sidebar,
                    variable=self.filter_vars[key],
                    values=["All"],
                    width=190, height=30,
                    fg_color=COLORS["bg_white"],
                    text_color=COLORS["text_dark"],
                    border_color=COLORS["filter_border"],
                    button_color=COLORS["filter_border"],
                    button_hover_color=COLORS["header_blue"],
                    dropdown_fg_color=COLORS["bg_white"],
                    dropdown_text_color=COLORS["text_dark"],
                    font=CTkFont(family=FONT_FAMILY, size=12),
                    command=self._apply_filters
                )
                menu.pack(fill="x", padx=12, pady=(0, 2))
                
                # ผูกคำสั่งแยกกันตามช่อง
                if key == "order_no":
                    menu.bind("<KeyRelease>", self._on_order_search)
                else:
                    menu.bind("<KeyRelease>", self._on_sale_order_search) # 🟢 ผูกคำสั่งค้นหา Sale Order No.
                    
                menu.bind("<Return>", self._apply_filters)
                
            else:
                # 🟢 ตัวอื่นๆ ใช้ CTkOptionMenu (คลิกเลือกแบบเดิม)
                menu = CTkOptionMenu(
                    sidebar,
                    variable=self.filter_vars[key],
                    values=["All"],
                    width=190, height=30,
                    fg_color=COLORS["filter_bg"],
                    text_color=COLORS["text_dark"],
                    button_color=COLORS["filter_border"],
                    button_hover_color=COLORS["header_blue"],
                    dropdown_fg_color=COLORS["bg_white"],
                    dropdown_text_color=COLORS["text_dark"],
                    font=CTkFont(family=FONT_FAMILY, size=12),
                    command=self._apply_filters,
                )
                menu.pack(fill="x", padx=12, pady=(0, 2))
            
            self.filter_menus[key] = menu

        # Divider
        CTkFrame(sidebar, fg_color=COLORS["border"], height=1).pack(fill="x", padx=12, pady=15)

        # Refresh button
        CTkButton(
            sidebar, text="🔄  รีเฟรชฐานข้อมูล",
            fg_color=COLORS["btn_blue"], hover_color=COLORS["btn_hover"],
            font=CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            corner_radius=6, height=36,
            command=self._load_data_from_db
        ).pack(fill="x", padx=12, pady=(0, 20))

    # 🟢 ฟังก์ชันอัจฉริยะ สำหรับค้นหา Order No. ใน Dropdown แบบเรียลไทม์
    def _on_order_search(self, event):
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Return']:
            return 
        typed_text = self.filter_vars["order_no"].get().lower()
        if typed_text == "" or typed_text == "all":
            self.filter_menus["order_no"].configure(values=["All"] + self.all_order_nos)
        else:
            matching_orders = [order for order in self.all_order_nos if typed_text in str(order).lower()]
            if matching_orders:
                self.filter_menus["order_no"].configure(values=matching_orders)
            else:
                self.filter_menus["order_no"].configure(values=["ไม่พบข้อมูล"])

    def _on_sale_order_search(self, event):
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Return']:
            return 
        typed_text = self.filter_vars["sale_order_no"].get().lower()
        if typed_text == "" or typed_text == "all":
            self.filter_menus["sale_order_no"].configure(values=["All"] + self.all_sale_order_nos)
        else:
            matching_orders = [order for order in self.all_sale_order_nos if typed_text in str(order).lower()]
            if matching_orders:
                self.filter_menus["sale_order_no"].configure(values=matching_orders)
            else:
                self.filter_menus["sale_order_no"].configure(values=["ไม่พบข้อมูล"])


    # =========================================================
    # MAIN CONTENT
    # =========================================================
    def _build_main_content(self):
        main_frame = CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 10), pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)

        # ── Header bar ──────────────────────────────────────
        header_frame = CTkFrame(main_frame, fg_color=COLORS["sidebar_header"],
                                corner_radius=8, height=48)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header_frame.grid_propagate(False)
        CTkLabel(header_frame,
                 text="  📊  Cost Benchmark Dashboard",
                 font=CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
                 text_color="#FFFFFF").pack(side="left", padx=14, pady=10)

        # ── KPI Cards ────────────────────────────────────────
        kpi_frame = CTkFrame(main_frame, fg_color="transparent")
        kpi_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for i in range(5):
            kpi_frame.grid_columnconfigure(i, weight=1)

        self.kpi_labels = {}
        kpi_config = [
            ("จำนวน Order รวม",        "total_orders", COLORS["kpi_blue"],   "🗂"),
            ("ปริมาณรวม (เส้น/ชิ้น)",  "total_qty",    COLORS["kpi_green"],  "📦"),
            ("ยอดซื้อรวม (ทุน)",        "total_cost",   COLORS["kpi_red"],    "💰"),
            ("ยอดขายรวม",               "total_sales",  COLORS["kpi_green"],  "📈"),
            ("Markup เฉลี่ย",           "avg_margin",   COLORS["kpi_purple"], "📊"),
        ]

        for i, (title, key, color, icon) in enumerate(kpi_config):
            card = CTkFrame(kpi_frame, fg_color=COLORS["bg_white"],
                            corner_radius=8, border_width=1,
                            border_color=COLORS["border"])
            card.grid(row=0, column=i, padx=4, sticky="nsew")

            accent = CTkFrame(card, fg_color=color, corner_radius=0, height=4)
            accent.pack(fill="x")

            CTkLabel(card, text=f"{icon}  {title}",
                     font=CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
                     text_color=COLORS["text_medium"]).pack(pady=(10, 2))

            val_label = CTkLabel(card, text="–",
                                 font=CTkFont(family=FONT_FAMILY, size=22, weight="bold"),
                                 text_color=color)
            val_label.pack(pady=(0, 12))
            self.kpi_labels[key] = val_label

        # ── Table ────────────────────────────────────────────
        table_outer = CTkFrame(main_frame, fg_color=COLORS["bg_white"],
                               corner_radius=8, border_width=1,
                               border_color=COLORS["border"])
        table_outer.grid(row=2, column=0, sticky="nsew")
        table_outer.grid_columnconfigure(0, weight=1)
        table_outer.grid_rowconfigure(1, weight=1)

        title_bar = CTkFrame(table_outer, fg_color=COLORS["header_blue"],
                             corner_radius=0, height=36)
        title_bar.grid(row=0, column=0, sticky="ew")
        title_bar.grid_propagate(False)
        CTkLabel(title_bar, text="  รายละเอียดข้อมูล Cost Benchmark",
                 font=CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
                 text_color="#FFFFFF").pack(side="left", padx=10, pady=8)

        table_inner = tk.Frame(table_outer, bg=COLORS["bg_white"])
        table_inner.grid(row=1, column=0, sticky="nsew", padx=2, pady=2)

        if HAS_TKSHEET:
            self._build_tksheet(table_inner)
        else:
            tk.Label(table_inner, text="⚠️ กรุณาติดตั้ง tksheet: pip install tksheet",
                     bg=COLORS["bg_white"], fg="red",
                     font=(FONT_FAMILY, 12)).pack(expand=True, pady=40)

    # =========================================================
    # TKSHEET TABLE
    # =========================================================
    def _build_tksheet(self, parent):
        self.display_columns = [
            "Selec\nt",
            "วันที่ขอราคา",
            "PRIORI\nTY",
            "WIN RATE\n%",
            "สถานะ",
            "ชื่อ\nSale",
            "Order No.",
            "Sale Order No.",
            "รายการสินค้า",
            "จำนวน",
            "Max of\nน้ำหนัก\n/เส้น",
            "Average of\nราคาขาย /\nเส้น",
            "ราคาขาย /\nกก.",
            "Sum of\nราคาขาย\nรวม",
            "Average of\nMarkup\nGuide (%)",
            "Sum of\nต้นทุนรวม\n(รวมย้าย)",
            "Sum of\nต้นทุน/เส้น",
            "ชื่อ Supplier",
            "Average of\nส่วนลด 2 (%)",
            "Average of\nส่วนลด 1 (%)",
            "Sum of ต้นทุนรวม\n(ไม่รวมย้าย)",
        ]

        self.col_source = {
            "Selec\nt":                          "Select", 
            "วันที่ขอราคา":                      "วันที่ขอราคา",
            "PRIORI\nTY":                        "PRIORITY",
            "WIN RATE\n%":                        "WIN RATE %",
            "สถานะ":                             "สถานะ",
            "ชื่อ\nSale":                       "ชื่อ Sale",
            "Order No.":                          "Order No.",
            "Sale Order No.":                     "Sale Order No.",
            "รายการสินค้า":                      "รายการสินค้า",
            "จำนวน":                             "จำนวน",
            "Max of\nน้ำหนัก\n/เส้น":           "น้ำหนัก/เส้น",
            "Average of\nราคาขาย /\nเส้น":      "ราคาขาย / เส้น",
            "ราคาขาย /\nกก.":                   "ราคาขาย / กก.",
            "Sum of\nราคาขาย\nรวม":             "ราคาขาย รวม",
            "Average of\nMarkup\nGuide (%)":     "Markup Guide (%)",
            "Sum of\nต้นทุนรวม\n(รวมย้าย)":    "ต้นทุนรวม (รวมย้าย)",
            "Sum of\nต้นทุน/เส้น":              "ต้นทุน/เส้น",
            "ชื่อ Supplier":                     "ชื่อ Supplier",
            "Average of\nส่วนลด 2 (%)":         "ส่วนลด 2 (%)",
            "Average of\nส่วนลด 1 (%)":         "ส่วนลด 1 (%)",
            "Sum of ต้นทุนรวม\n(ไม่รวมย้าย)":  "ต้นทุนรวม (ไม่รวมย้าย)",
        }

        self.money_cols = {
            "Sum of\nราคาขาย\nรวม",
            "Sum of\nต้นทุนรวม\n(รวมย้าย)",
            "Average of\nราคาขาย /\nเส้น",
            "ราคาขาย /\nกก.",
            "Sum of\nต้นทุน/เส้น",
            "Sum of ต้นทุนรวม\n(ไม่รวมย้าย)",
        }
        self.pct_cols = {
            "WIN RATE\n%",
            "Average of\nMarkup\nGuide (%)",
            "Average of\nส่วนลด 2 (%)",
            "Average of\nส่วนลด 1 (%)",
        }
        self.num_cols = {
            "จำนวน",
            "Max of\nน้ำหนัก\n/เส้น",
            "Average of\nราคาขาย /\nเส้น",
        }

        self.sheet = Sheet(
            parent,
            headers=self.display_columns,
            data=[],
            theme="light blue",
            font=(FONT_FAMILY, 10, "normal"),
            header_font=(FONT_FAMILY, 9, "bold"),
            show_row_index=True,
            row_index_width=32,
            empty_horizontal=0,
            empty_vertical=10,
            row_height=26,
            header_height=52,           
        )
        self.sheet.pack(fill="both", expand=True)

        self.sheet.enable_bindings((
            "single_select", "row_select",
            "column_width_resize", "row_height_resize",
            "arrowkeys", "copy",
        ))

        self.sheet.set_options(
            header_bg=COLORS["header_blue"],
            header_fg=COLORS["header_text"],
            header_grid_fg="#3B6FE8",
            selected_rows_border_fg=COLORS["header_blue"],
            selected_rows_bg="#EFF6FF",
            selected_rows_fg=COLORS["text_dark"],
            table_bg=COLORS["bg_white"],
            table_fg=COLORS["text_dark"],
            table_grid_fg=COLORS["border"],
            index_bg="#F1F5F9",
            index_fg=COLORS["text_medium"],
            index_grid_fg=COLORS["border"],
            index_border_fg=COLORS["border"],
        )

        col_widths = {
            "Selec\nt":                          42,
            "วันที่ขอราคา":                      88,
            "PRIORI\nTY":                        62,
            "WIN RATE\n%":                        72,
            "สถานะ":                             52,
            "ชื่อ\nSale":                       52,
            "Order No.":                          88,
            "Sale Order No.":                    100,
            "รายการสินค้า":                     260,
            "จำนวน":                             68,
            "Max of\nน้ำหนัก\n/เส้น":           80,
            "Average of\nราคาขาย /\nเส้น":      95,
            "ราคาขาย /\nกก.":                   80,
            "Sum of\nราคาขาย\nรวม":            110,
            "Average of\nMarkup\nGuide (%)":    100,
            "Sum of\nต้นทุนรวม\n(รวมย้าย)":   120,
            "Sum of\nต้นทุน/เส้น":              90,
            "ชื่อ Supplier":                    160,
            "Average of\nส่วนลด 2 (%)":        105,
            "Average of\nส่วนลด 1 (%)":        105,
            "Sum of ต้นทุนรวม\n(ไม่รวมย้าย)": 140,
        }
        for i, col in enumerate(self.display_columns):
            self.sheet.column_width(i, col_widths.get(col, 100))

    # =========================================================
    # DATA LOADING
    # =========================================================
    def _load_data_from_db(self):
        conn = self.app_container.get_connection()
        try:
            query = "SELECT * FROM cost_benchmarks"
            df = pd.read_sql(query, conn)

            if df.empty:
                self.raw_df = pd.DataFrame()
            else:
                def clean_numeric(series):
                    return (
                        series.astype(str)
                        .str.strip()
                        .str.replace(',', '', regex=False)
                        .str.replace('%', '', regex=False)
                        .str.replace('฿', '', regex=False)
                        .str.replace(' ', '', regex=False)
                        .replace({'': '0', 'None': '0', 'nan': '0', 'NaN': '0', 'none': '0'})
                        .pipe(lambda s: pd.to_numeric(s, errors='coerce'))
                        .fillna(0)
                    )

                numeric_cols = [
                    "จำนวน", "ทุนรวม", "ราคาขาย รวม",
                    "ต้นทุนรวม (รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)",
                    "น้ำหนัก/เส้น", "ราคาขาย / เส้น", "ราคาขาย / กก.",
                    "WIN RATE %", "Markup Guide (%)",
                    "ต้นทุน/เส้น", "ส่วนลด 1 (%)", "ส่วนลด 2 (%)",
                ]
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = clean_numeric(df[col])

                self.raw_df = df

            self._update_filter_dropdowns()
            self._apply_filters()

        except Exception as e:
            messagebox.showerror("Database Error", f"โหลดข้อมูลล้มเหลว:\n{e}", parent=self)
        finally:
            if conn:
                self.app_container.release_connection(conn)

    # =========================================================
    # FILTER DROPDOWNS
    # =========================================================
    def _update_filter_dropdowns(self):
        if self.raw_df.empty:
            for key in self.filter_menus:
                self.filter_menus[key].configure(values=["All"])
                self.filter_vars[key].set("All")
            return

        def uniq(col):
            if col in self.raw_df.columns:
                vals = [str(x) for x in self.raw_df[col].unique()
                        if str(x).strip() not in ("", "None", "nan")]
                return ["All"] + sorted(vals)
            return ["All"]

        # 🟢 ผูกชื่อตัวแปร Filter กับชื่อคอลัมน์ในตาราง Database
        mapping = {
            "order_no":      "Order No.",
            "sale_order_no": "Sale Order No.", # 🟢 เพิ่ม Sale Order No.
            "pu_user":       "created_by",      
            "sale_name":     "ชื่อ Sale",
            "supplier":      "ชื่อ Supplier",
            "status":        "สถานะ",
            "priority":      "PRIORITY",
            "select":        "Select",
            "year":          "benchmark_year",
            "month":         "benchmark_month",
        }
        
        for key, col in mapping.items():
            vals = uniq(col)
            
            # เก็บ Data ทั้งหมดแยกไว้ สำหรับระบบพิมพ์ค้นหา
            if key == "order_no":
                self.all_order_nos = [v for v in vals if v != "All"]
            elif key == "sale_order_no": # 🟢 เก็บของ Sale Order No. ด้วย
                self.all_sale_order_nos = [v for v in vals if v != "All"]
                
            self.filter_menus[key].configure(values=vals)
            
            # ถ้าค่าเดิมที่ตั้งไว้ไม่มีในลิสต์ (และไม่ใช่การค้นหาแบบอิสระ) ให้ปรับกลับเป็น All
            if self.filter_vars[key].get() not in vals and self.filter_vars[key].get() != "All":
                if key == "order_no" and self.filter_vars[key].get() in self.all_order_nos:
                    pass # ปล่อยผ่านถ้าค่าที่พิมพ์อยู่มันมีในลิสต์ Data ดิบ
                elif key == "sale_order_no" and self.filter_vars[key].get() in self.all_sale_order_nos:
                    pass # ปล่อยผ่านสำหรับ Sale order
                else:
                    self.filter_vars[key].set("All")

    # =========================================================
    # APPLY FILTERS
    # =========================================================
    def _apply_filters(self, *args):
        if self.raw_df.empty:
            self._update_kpis(pd.DataFrame())
            self._update_table(pd.DataFrame())
            return

        df = self.raw_df.copy()

        # 🟢 กรองข้อมูลตาม Dropdown ที่เลือก
        col_map = {
            "order_no":      "Order No.",
            "sale_order_no": "Sale Order No.", # 🟢 เพิ่ม Sale Order No.
            "pu_user":       "created_by",
            "sale_name":     "ชื่อ Sale",
            "supplier":      "ชื่อ Supplier",
            "status":        "สถานะ",
            "priority":      "PRIORITY",
            "select":        "Select",
            "year":          "benchmark_year",
            "month":         "benchmark_month",
        }
        
        for key, col in col_map.items():
            val = self.filter_vars[key].get()
            if val != "All" and val != "" and col in df.columns:
                # 🟢 ให้ช่อง Order No. และ Sale Order No. กรองตารางแบบ เจอคำบางส่วนก็แสดงให้เลย
                if key in ["order_no", "sale_order_no"]:
                    df = df[df[col].astype(str).str.contains(val, case=False, na=False)]
                else:
                    df = df[df[col].astype(str) == str(val)]

        self._update_kpis(df)
        self._update_table(df)

    # =========================================================
    # KPI UPDATE
    # =========================================================
    def _update_kpis(self, df):
        if df.empty:
            for key in self.kpi_labels:
                self.kpi_labels[key].configure(text="–")
            return

        df_active = df
        if "รายการสินค้า" in df.columns:
            df_active = df[df["รายการสินค้า"].astype(str).str.strip() != ""]

        total_orders = len(df_active)

        total_qty    = (df_active["จำนวน"].sum()
                        if "จำนวน" in df_active.columns else 0)
        total_cost   = (df_active["ต้นทุนรวม (รวมย้าย)"].sum()
                        if "ต้นทุนรวม (รวมย้าย)" in df_active.columns else 0)
        total_sales  = (df_active["ราคาขาย รวม"].sum()
                        if "ราคาขาย รวม" in df_active.columns else 0)
                        
        # 🟢 [แก้ไข] คำนวณ Markup เฉลี่ย โดยไม่เอาเลข 0 มาคิด
        if "Markup Guide (%)" in df_active.columns:
            margin_series = pd.to_numeric(df_active["Markup Guide (%)"], errors='coerce').dropna()
            margin_non_zero = margin_series[margin_series != 0] # ตัดเลข 0 ทิ้ง
            avg_margin = margin_non_zero.mean() if not margin_non_zero.empty else 0
        else:
            avg_margin = 0

        self.kpi_labels["total_orders"].configure(text=f"{total_orders:,}")
        self.kpi_labels["total_qty"].configure(text=f"{total_qty:,.0f}")
        self.kpi_labels["total_cost"].configure(
            text=f"฿{total_cost/1_000_000:.2f}M" if total_cost >= 1_000_000
            else f"฿{total_cost:,.0f}")
        self.kpi_labels["total_sales"].configure(
            text=f"฿{total_sales/1_000_000:.2f}M" if total_sales >= 1_000_000
            else f"฿{total_sales:,.0f}")
        self.kpi_labels["avg_margin"].configure(
            text=f"{avg_margin:,.2f}%" if pd.notna(avg_margin) else "0.00%")
    # =========================================================
    # TABLE UPDATE
    # =========================================================
    # =========================================================
    # TABLE UPDATE
    # =========================================================
    def _update_table(self, df):
        if not HAS_TKSHEET:
            return

        self.sheet.set_sheet_data([])

        if df.empty:
            self.sheet.redraw()
            return

        col_source = self.col_source
        money_cols = self.money_cols
        pct_cols   = self.pct_cols
        num_cols   = self.num_cols

        table_data = []
        for _, row in df.iterrows():
            row_data = []
            for dcol in self.display_columns:
                src = col_source.get(dcol)

                if src is None or src not in row.index:
                    row_data.append("")
                    continue

                val = row[src]

                is_empty = pd.isna(val) or str(val).strip() in ("", "None", "nan", "NaN", "none")
                if is_empty:
                    row_data.append("")
                    continue

                if dcol in money_cols or dcol in pct_cols or dcol in num_cols:
                    fval = pd.to_numeric(val, errors='coerce')
                    if pd.isna(fval):
                        row_data.append(str(val))
                    elif dcol in money_cols:
                        row_data.append(f"฿{float(fval):,.2f}")
                    elif dcol in pct_cols:
                        row_data.append(f"{float(fval):.2f}%")
                    else:  # num_cols
                        fv = float(fval)
                        row_data.append(f"{fv:,.2f}" if fv != int(fv) else f"{int(fv):,}")
                else:
                    row_data.append(str(val))

            if any(str(x).strip() for x in row_data):
                table_data.append(row_data)

        # ── Total row ──────────────────────────────────────────
        if table_data:
            def cidx(name):
                try: return self.display_columns.index(name)
                except ValueError: return -1

            total_row = [""] * len(self.display_columns)
            total_row[0] = "Total"

            sum_map = {
                "Sum of\nราคาขาย\nรวม":            "ราคาขาย รวม",
                "Sum of\nต้นทุนรวม\n(รวมย้าย)":   "ต้นทุนรวม (รวมย้าย)",
                "Sum of\nต้นทุน/เส้น":             "ต้นทุน/เส้น",
                "Sum of ต้นทุนรวม\n(ไม่รวมย้าย)": "ต้นทุนรวม (ไม่รวมย้าย)",
                "จำนวน":                            "จำนวน",
                "Max of\nน้ำหนัก\n/เส้น":          "น้ำหนัก/เส้น",
            }
            avg_map = {
                "Average of\nMarkup\nGuide (%)":   "Markup Guide (%)",
                "Average of\nส่วนลด 2 (%)":        "ส่วนลด 2 (%)",
                "Average of\nส่วนลด 1 (%)":        "ส่วนลด 1 (%)",
                "Average of\nราคาขาย /\nเส้น":     "ราคาขาย / เส้น",
                "ราคาขาย /\nกก.":                  "ราคาขาย / กก.",
            }

            # 🟢 1. คำนวณผลรวม (Sum)
            for dcol, src in sum_map.items():
                idx = cidx(dcol)
                if idx < 0 or src not in df.columns:
                    continue
                try:
                    series = pd.to_numeric(df[src], errors='coerce').dropna()
                    if series.empty: continue
                    val = series.sum()
                    if val == 0: continue
                    if dcol in money_cols:
                        total_row[idx] = f"฿{float(val):,.2f}"
                    else:
                        fv = float(val)
                        total_row[idx] = f"{fv:,.2f}" if fv != int(fv) else f"{int(fv):,}"
                except Exception:
                    pass

            # 🟢 2. คำนวณค่าเฉลี่ย (Average) แบบไม่เอาเลข 0 มาคิด
            for dcol, src in avg_map.items():
                idx = cidx(dcol)
                if idx < 0 or src not in df.columns:
                    continue
                try:
                    # แปลงข้อมูลเป็นตัวเลขและตัดค่าว่าง (NaN) ทิ้ง
                    series = pd.to_numeric(df[src], errors='coerce').dropna()
                    
                    # กรองเอาเฉพาะตัวเลขที่ "ไม่เท่ากับ 0"
                    series_non_zero = series[series != 0]
                    
                    if series_non_zero.empty: 
                        continue
                        
                    # หาค่าเฉลี่ยจากข้อมูลที่ไม่มีเลข 0 แล้ว
                    m = series_non_zero.mean()
                    
                    if pd.notna(m):
                        if dcol in money_cols:
                            total_row[idx] = f"฿{float(m):,.2f}"
                        else:
                            total_row[idx] = f"{float(m):.2f}%"
                except Exception:
                    pass

            table_data.append(total_row)

        self.sheet.set_sheet_data(table_data)

        # ── Highlight columns ──────────────────────────────────
        green_col = "Sum of\nราคาขาย\nรวม"
        red_col   = "Sum of\nต้นทุนรวม\n(รวมย้าย)"

        try:
            self.sheet.highlight_columns(
                columns=[self.display_columns.index(green_col)],
                bg=COLORS["highlight_green"], fg=COLORS["highlight_green_fg"],
            )
        except ValueError:
            pass
        try:
            self.sheet.highlight_columns(
                columns=[self.display_columns.index(red_col)],
                bg=COLORS["highlight_red"], fg=COLORS["highlight_red_fg"],
            )
        except ValueError:
            pass

        # ── Alternate row shading ──────────────────────────────
        n_data = len(table_data) - 1
        for r in range(n_data):
            if r % 2 == 1:
                self.sheet.highlight_rows(rows=[r], bg="#F8FAFF", fg=COLORS["text_dark"])

        # ── Total row highlight ────────────────────────────────
        self.sheet.highlight_rows(
            rows=[len(table_data) - 1],
            bg="#BFDBFE",
            fg=COLORS["kpi_blue"],
        )

        self.sheet.readonly_columns(
            columns=list(range(len(self.display_columns))),
            readonly=True,
        )
        self.sheet.redraw()