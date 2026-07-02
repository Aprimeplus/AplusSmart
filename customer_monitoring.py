import tkinter as tk
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkEntry,
                           CTkFont, CTkScrollableFrame, CTkComboBox)
import pandas as pd
from tksheet import Sheet
from datetime import datetime

try:
    from tkcalendar import DateEntry as _DateEntry
    class _PatchedDE(_DateEntry):
        def _on_focus_out_cal(self, event):
            try:
                focused = self._top_cal.focus_get()
                if focused and str(focused).startswith(str(self._top_cal)):
                    return
            except Exception:
                pass
            super()._on_focus_out_cal(event)
        def _drop_down(self):
            super()._drop_down()
            if hasattr(self, '_top_cal') and self._top_cal:
                self._top_cal.bind("<FocusOut>", self._on_focus_out_cal)
    _HAS_CAL = True
except ImportError:
    _HAS_CAL = False

_DE_KW = dict(
    background="#1D4ED8", foreground="white", borderwidth=0,
    headersbackground="#1D4ED8", headersforeground="white",
    selectbackground="#2563EB", selectforeground="white",
    normalbackground="#FFFFFF", normalforeground="#111827",
    weekendbackground="#FFFFFF", weekendforeground="#991B1B",
    othermonthforeground="#9CA3AF", othermonthbackground="#FFFFFF",
    othermonthweforeground="#9CA3AF", othermonthwebackground="#FFFFFF",
    font=("Tahoma", 12),
    date_pattern="dd/mm/yyyy", locale="th_TH",
    showweeknumbers=False, showothermonthdays=False,
    year=datetime.now().year,
)

THAI_MONTHS_SHORT = {
    1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.',
    5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.',
    9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'
}

_CANCELLED = ('Cancelled', 'Cancelled by PU')


class CustomerMonitoringWidget(CTkFrame):

    def __init__(self, master, app_container, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_container
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_filter_panel()
        self._build_table_area()
        self._load_filter_options()
        self.bind('<Map>', self._on_map)

    def _on_map(self, _=None):
        self.unbind('<Map>')
        self.after(150, self._load)

    # ── filter panel ─────────────────────────────────────────────────────────

    def _build_filter_panel(self):
        panel = CTkScrollableFrame(
            self, width=175,
            fg_color="#F8FAFC", border_width=1, border_color="#CBD5E1",
            corner_radius=8,
        )
        panel.grid(row=0, column=0, sticky="nsew", padx=(6, 4), pady=6)
        panel.grid_columnconfigure(0, weight=1)

        def lbl(text, row, pady=(6, 0)):
            CTkLabel(panel, text=text, font=CTkFont(size=11),
                     text_color="#475569").grid(row=row, column=0, padx=10,
                                                pady=pady, sticky="w")

        r = 0
        CTkLabel(panel, text="ตัวกรอง",
                 font=CTkFont(size=13, weight="bold"),
                 text_color="#0F172A").grid(row=r, column=0, padx=10,
                                            pady=(10, 8), sticky="w")
        r += 1

        lbl("ค้นหาลูกค้า", r); r += 1
        self._search_var = tk.StringVar()
        CTkEntry(panel, textvariable=self._search_var,
                 placeholder_text="รหัส / ชื่อ", height=28,
                 font=CTkFont(size=11)).grid(
            row=r, column=0, padx=10, pady=(0, 4), sticky="ew"); r += 1

        lbl("รหัสพนักงาน", r); r += 1
        self._sale_var = tk.StringVar(value="ทั้งหมด")
        self._sale_cb = CTkComboBox(
            panel, variable=self._sale_var, values=["ทั้งหมด"],
            height=28, font=CTkFont(size=11), state="readonly")
        self._sale_cb.grid(row=r, column=0, padx=10, pady=(0, 4), sticky="ew"); r += 1

        lbl("ประเภทลูกค้า", r); r += 1
        self._ctype_var = tk.StringVar(value="ทั้งหมด")
        CTkComboBox(
            panel, variable=self._ctype_var,
            values=["ทั้งหมด", "ลูกค้าเก่า", "ลูกค้าใหม่"],
            height=28, font=CTkFont(size=11), state="readonly",
        ).grid(row=r, column=0, padx=10, pady=(0, 4), sticky="ew"); r += 1

        lbl("ปี (พ.ศ.)", r); r += 1
        self._year_var = tk.StringVar(value=str(datetime.now().year + 543))
        self._year_cb = CTkComboBox(
            panel, variable=self._year_var, values=["2568", "2569"],
            height=28, font=CTkFont(size=11), state="readonly")
        self._year_cb.grid(row=r, column=0, padx=10, pady=(0, 4), sticky="ew"); r += 1

        lbl("วันที่เริ่ม", r); r += 1
        self._from_var = tk.StringVar()
        if _HAS_CAL:
            self._cal_from = _PatchedDE(panel, width=13,
                                        textvariable=self._from_var, **_DE_KW)
            self._cal_from.delete(0, "end")
            self._cal_from.grid(row=r, column=0, padx=10, pady=(0, 2), sticky="w")
            self._cal_from.bind("<<DateEntrySelected>>", lambda e: self._load())
        else:
            CTkEntry(panel, textvariable=self._from_var,
                     placeholder_text="dd/mm/yyyy", height=28,
                     font=CTkFont(size=11)).grid(
                row=r, column=0, padx=10, pady=(0, 2), sticky="ew")
        r += 1

        lbl("วันที่สิ้นสุด", r); r += 1
        self._to_var = tk.StringVar()
        if _HAS_CAL:
            self._cal_to = _PatchedDE(panel, width=13,
                                      textvariable=self._to_var, **_DE_KW)
            self._cal_to.delete(0, "end")
            self._cal_to.grid(row=r, column=0, padx=10, pady=(0, 8), sticky="w")
            self._cal_to.bind("<<DateEntrySelected>>", lambda e: self._load())
        else:
            CTkEntry(panel, textvariable=self._to_var,
                     placeholder_text="dd/mm/yyyy", height=28,
                     font=CTkFont(size=11)).grid(
                row=r, column=0, padx=10, pady=(0, 8), sticky="ew")
        r += 1

        CTkButton(panel, text="🔍 ค้นหา", height=30,
                  font=CTkFont(size=12, weight="bold"),
                  fg_color="#1D4ED8", hover_color="#1E40AF",
                  command=self._load).grid(
            row=r, column=0, padx=10, pady=(0, 4), sticky="ew"); r += 1
        CTkButton(panel, text="ล้าง", height=28,
                  font=CTkFont(size=11),
                  fg_color="#94A3B8", hover_color="#64748B",
                  command=self._clear).grid(
            row=r, column=0, padx=10, pady=(0, 10), sticky="ew")

    # ── table area ───────────────────────────────────────────────────────────

    def _build_table_area(self):
        from customtkinter import CTkButton, CTkOptionMenu
        area = CTkFrame(self, fg_color="transparent")
        area.grid(row=0, column=1, sticky="nsew", padx=(4, 6), pady=6)
        area.grid_columnconfigure(0, weight=1)
        area.grid_rowconfigure(1, weight=1)

        hdr = CTkFrame(area, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        CTkLabel(hdr, text="👥  Customer Monitoring — ยอดขายรายลูกค้า",
                 font=CTkFont(size=14, weight="bold"),
                 text_color="#0F172A").pack(side="left")

        # toggle buttons (right side)
        self._btn_chart = CTkButton(
            hdr, text="📊 แผนภูมิ", width=95, height=28,
            font=CTkFont(size=11), fg_color="#6B7280", hover_color="#4B5563",
            command=lambda: self._set_view("chart"))
        self._btn_chart.pack(side="right", padx=(4, 0))
        self._btn_table = CTkButton(
            hdr, text="📋 ตาราง", width=85, height=28,
            font=CTkFont(size=11), fg_color="#1D4ED8", hover_color="#1E40AF",
            command=lambda: self._set_view("table"))
        self._btn_table.pack(side="right", padx=(0, 4))

        self._info_lbl = CTkLabel(hdr, text="",
                                   font=CTkFont(size=11), text_color="#64748B")
        self._info_lbl.pack(side="right", padx=8)

        # ── table sheet frame ─────────────────────────────────────────────
        self._sheet_frame = CTkFrame(area, fg_color="transparent")
        self._sheet_frame.grid(row=1, column=0, sticky="nsew")
        self._sheet_frame.grid_columnconfigure(0, weight=1)
        self._sheet_frame.grid_rowconfigure(0, weight=1)

        self._sheet = Sheet(
            self._sheet_frame,
            theme="light blue",
            header_font=("Tahoma", 11, "bold"),
            data_font=("Tahoma", 11),
            default_column_width=80,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        self._sheet.grid(row=0, column=0, sticky="nsew")
        self._sheet.enable_bindings(
            "column_width_resize",
            "double_click_column_resize",
            "right_click_popup_menu",
            "rc_select",
            "copy",
        )
        self._sort_col_idx: int | None = None
        self._sort_asc: bool = True
        self._pivot_base: "pd.DataFrame | None" = None
        self._months_stored: list = []
        self._year_stored: int = 0
        self._header_press_x: int = 0
        self._sheet.MT.CH.bind("<ButtonPress-1>",   self._on_ch_press,   add=True)
        self._sheet.MT.CH.bind("<ButtonRelease-1>", self._on_ch_release, add=True)

        # ── chart frame (hidden initially) ────────────────────────────────
        self._chart_frame = CTkFrame(area, fg_color="white", corner_radius=8)
        # chart controls bar
        self._chart_ctrl = CTkFrame(self._chart_frame, fg_color="#F1F5F9",
                                     corner_radius=6)
        self._chart_ctrl.pack(fill="x", padx=8, pady=(8, 4))
        CTkLabel(self._chart_ctrl, text="แสดง Top",
                 font=CTkFont(size=11)).pack(side="left", padx=(10, 4))
        self._top_n_var = tk.StringVar(value="20")
        CTkOptionMenu(self._chart_ctrl, variable=self._top_n_var,
                      values=["10", "20", "30", "50"],
                      width=70, height=26, font=CTkFont(size=11),
                      command=lambda _: self._draw_chart()).pack(side="left")
        CTkLabel(self._chart_ctrl, text="ลูกค้า",
                 font=CTkFont(size=11)).pack(side="left", padx=(4, 16))

        self._chart_type_var = tk.StringVar(value="ยอดรวมรายลูกค้า")
        CTkOptionMenu(self._chart_ctrl, variable=self._chart_type_var,
                      values=["ยอดรวมรายลูกค้า", "ยอดรายเดือน (รวม)"],
                      width=170, height=26, font=CTkFont(size=11),
                      command=lambda _: self._draw_chart()).pack(side="left")

        self._chart_canvas_widget = None   # FigureCanvasTkAgg widget
        self._chart_container = CTkFrame(self._chart_frame, fg_color="white")
        self._chart_container.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    # ── data ─────────────────────────────────────────────────────────────────

    def _load_filter_options(self):
        try:
            conn = self.app.get_connection()
            try:
                df_k = pd.read_sql_query(
                    "SELECT DISTINCT sale_key FROM commissions "
                    "WHERE sale_key IS NOT NULL AND is_active=1 ORDER BY sale_key",
                    conn,
                )
                self._sale_cb.configure(
                    values=["ทั้งหมด"] + df_k["sale_key"].tolist())

                df_y = pd.read_sql_query(
                    "SELECT DISTINCT commission_year FROM commissions "
                    "WHERE is_active=1 ORDER BY commission_year",
                    conn,
                )
                years = [str(y + 543) for y in df_y["commission_year"].tolist()]
                if years:
                    self._year_cb.configure(values=years)
                    self._year_var.set(years[-1])
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            print(f"CustomerMonitoring filter_options error: {e}")

    def _clear(self):
        self._search_var.set("")
        self._sale_var.set("ทั้งหมด")
        self._ctype_var.set("ทั้งหมด")
        self._year_var.set(str(datetime.now().year + 543))
        self._from_var.set("")
        self._to_var.set("")
        if _HAS_CAL:
            self._cal_from.delete(0, "end")
            self._cal_to.delete(0, "end")
        self._load()

    def _load(self):
        try:
            year_thai = int(self._year_var.get())
            year = year_thai - 543          # แปลง พ.ศ. → ค.ศ. สำหรับ query DB
        except Exception:
            year_thai = datetime.now().year + 543
            year = datetime.now().year

        search   = self._search_var.get().strip()
        sale_key = self._sale_var.get()
        ctype    = self._ctype_var.get()
        d_from   = self._parse_date(self._from_var.get().strip())
        d_to     = self._parse_date(self._to_var.get().strip())

        ph = ["commission_year = %s", "is_active = 1",
              "status NOT IN ('Cancelled','Cancelled by PU')"]
        params = [year]

        if sale_key != "ทั้งหมด":
            ph.append("sale_key = %s"); params.append(sale_key)
        if ctype != "ทั้งหมด":
            ph.append("customer_type = %s"); params.append(ctype)
        if search:
            ph.append("(customer_id ILIKE %s OR customer_name ILIKE %s)")
            params += [f"%{search}%", f"%{search}%"]
        if d_from:
            ph.append("bill_date >= %s"); params.append(d_from)
        if d_to:
            ph.append("bill_date <= %s"); params.append(d_to)

        sql = (
            "SELECT customer_id, customer_name, sale_key, commission_month,"
            " SUM(sales_service_amount) AS amount"
            f" FROM commissions WHERE {' AND '.join(ph)}"
            " GROUP BY customer_id, customer_name, sale_key, commission_month"
        )

        try:
            conn = self.app.get_connection()
            try:
                df = pd.read_sql_query(sql, conn, params=params)
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            print(f"CustomerMonitoring _load error: {e}")
            return

        self._render(df, year_thai)

    @staticmethod
    def _parse_date(s):
        try:
            return datetime.strptime(s, "%d/%m/%Y").date()
        except Exception:
            return None

    # ── render ───────────────────────────────────────────────────────────────

    # heat map: 5 levels white → dark blue based on value ratio per column
    _HEAT = [
        (0.80, "#1E40AF", "white"),
        (0.60, "#2563EB", "white"),
        (0.40, "#60A5FA", "#0F172A"),
        (0.20, "#BFDBFE", "#0F172A"),
        (0.01, "#EFF6FF", "#0F172A"),
    ]

    # salesperson badge colors: fixed mapping (bg, fg)
    _SALE_COLORS: dict = {
        "VOW-S":    ("#EDE9FE", "#4C1D95"),   # ม่วง
        "VOW-P":    ("#EDE9FE", "#4C1D95"),   # ม่วง (เดียวกัน)
        "ILADA":    ("#DBEAFE", "#1E40AF"),   # ฟ้า
        "BUNNYCEE": ("#FEF9C3", "#713F12"),   # เหลือง
        "PIYAWAN":  ("#FFEDD5", "#9A3412"),   # ส้ม
        "JIRAPORN": ("#FCE7F3", "#9D174D"),   # ชมพู
        "LETHAI":   ("#FCE7F3", "#9D174D"),   # ชมพู
        "WACHIRA":  ("#FCE7F3", "#9D174D"),   # ชมพู
    }
    # fallback palette for unknown keys
    _SALE_PALETTE = [
        ("#D1FAE5", "#065F46"),  # green
        ("#FEE2E2", "#991B1B"),  # red
        ("#CFFAFE", "#164E63"),  # cyan
        ("#F3F4F6", "#111827"),  # gray
    ]

    @classmethod
    def _heat_color(cls, v, col_max):
        if col_max <= 0 or v <= 0:
            return None, None
        ratio = v / col_max
        for threshold, bg, fg in cls._HEAT:
            if ratio >= threshold:
                return bg, fg
        return None, None

    def _render(self, df, year_thai):
        self._sheet.dehighlight_all()

        if df.empty:
            self._sheet.headers(["(ไม่พบข้อมูล)"])
            self._sheet.set_sheet_data([[]])
            self._info_lbl.configure(text="0 ลูกค้า")
            self._sheet.redraw()
            return

        now = datetime.now()
        current_year_thai = now.year + 543
        max_month = now.month if year_thai == current_year_thai else 12
        months = sorted(
            m for m in df["commission_month"].dropna().unique().tolist()
            if int(m) <= max_month
        )

        pivot = df.pivot_table(
            index=["customer_id", "customer_name", "sale_key"],
            columns="commission_month",
            values="amount",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()
        pivot.columns.name = None
        pivot["_total"] = pivot[months].sum(axis=1)
        pivot = pivot.sort_values("_total", ascending=False).reset_index(drop=True)

        # เก็บไว้สำหรับ sort โดยไม่ต้อง query DB ใหม่
        self._pivot_base   = pivot.copy()
        self._months_stored = months
        self._year_stored   = year_thai
        self._sort_col_idx  = None   # reset sort state เมื่อโหลดใหม่
        self._sort_asc      = True

        self._draw_table(pivot, months, year_thai)

    # ── view toggle ──────────────────────────────────────────────────────────

    def _set_view(self, mode: str):
        if mode == "table":
            self._chart_frame.grid_remove()
            self._sheet_frame.grid(row=1, column=0, sticky="nsew")
            self._btn_table.configure(fg_color="#1D4ED8")
            self._btn_chart.configure(fg_color="#6B7280")
        else:
            self._sheet_frame.grid_remove()
            self._chart_frame.grid(row=1, column=0, sticky="nsew")
            self._btn_chart.configure(fg_color="#1D4ED8")
            self._btn_table.configure(fg_color="#6B7280")
            self._draw_chart()

    def _draw_chart(self):
        if self._pivot_base is None:
            return
        try:
            import matplotlib
            import matplotlib.ticker
            matplotlib.rcParams["font.family"] = "Tahoma"
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        except ImportError:
            from customtkinter import CTkLabel as _L
            _L(self._chart_container,
               text="กรุณาติดตั้ง matplotlib ก่อน (pip install matplotlib)").pack()
            return

        # clear previous chart
        if self._chart_canvas_widget:
            self._chart_canvas_widget.get_tk_widget().destroy()
            self._chart_canvas_widget = None

        top_n   = int(self._top_n_var.get())
        ctype   = self._chart_type_var.get()
        pivot   = self._pivot_base
        months  = self._months_stored
        m_names = [THAI_MONTHS_SHORT.get(int(m), str(m)) for m in months]

        _SALE_HEX = {
            "VOW-S": "#7C3AED", "VOW-P": "#7C3AED",
            "ILADA": "#2563EB", "BUNNYCEE": "#CA8A04",
            "PIYAWAN": "#EA580C", "JIRAPORN": "#DB2777",
            "LETHAI": "#DB2777", "WACHIRA": "#DB2777",
        }
        _FALLBACK = ["#059669", "#DC2626", "#0891B2", "#6B7280"]
        _fallback_map: dict = {}
        def _hex(key):
            if key in _SALE_HEX:
                return _SALE_HEX[key]
            if key not in _fallback_map:
                _fallback_map[key] = _FALLBACK[len(_fallback_map) % len(_FALLBACK)]
            return _fallback_map[key]

        def _fmt_val(v):
            """แสดงยอดเป็น K หรือ M"""
            if v >= 1_000:
                return f"{v/1_000:.1f}M"
            return f"{v:,.0f}K"

        if ctype == "ยอดรวมรายลูกค้า":
            df_top  = pivot.nlargest(top_n, "_total")
            labels  = df_top["customer_name"].tolist()
            values  = (df_top["_total"] / 1_000).tolist()   # หน่วย พัน (K)
            colors  = [_hex(k) for k in df_top["sale_key"].tolist()]
            n_bars  = len(labels)

            fig_h = max(5, n_bars * 0.42)
            fig = Figure(figsize=(10, fig_h), facecolor="white")
            ax  = fig.add_subplot(111)

            bars = ax.barh(range(n_bars), values, color=colors, height=0.65)
            ax.set_yticks(range(n_bars))
            ax.set_yticklabels(labels, fontsize=9)
            ax.invert_yaxis()

            # scale axis ที่ P90 เพื่อไม่ให้ outlier บีบ — bar ยาวกว่าจะ clip
            import numpy as np
            p90 = float(np.percentile(values, 90)) if len(values) > 3 else max(values)
            x_lim = max(p90 * 1.35, sorted(values)[-2] * 1.2) if len(values) > 1 else values[0] * 1.2
            ax.set_xlim(0, x_lim)
            ax.set_clip_on(True)

            # value labels — วางในแท่งถ้า bar ยาวพอ ไม่งั้นวางนอก
            for bar, v in zip(bars, values):
                w = bar.get_width()
                label_txt = _fmt_val(v)
                if w > x_lim * 0.15:   # วางในแท่ง
                    ax.text(min(w, x_lim) - x_lim * 0.01,
                            bar.get_y() + bar.get_height() / 2,
                            label_txt, va="center", ha="right",
                            fontsize=8, color="white", fontweight="bold")
                else:                   # วางนอกแท่ง
                    ax.text(w + x_lim * 0.01,
                            bar.get_y() + bar.get_height() / 2,
                            label_txt, va="center", ha="left", fontsize=8)

            ax.set_xlabel("ยอดขาย (พัน บาท)", fontsize=10)
            ax.set_title(
                f"Top {top_n} ลูกค้า — ยอดขายรวม ปี {self._year_stored}",
                fontsize=12, fontweight="bold", loc="left", pad=10)
            ax.xaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
            ax.grid(axis="x", alpha=0.25, linestyle="--")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)

            # margin ซ้ายสำหรับชื่อ Thai
            fig.subplots_adjust(left=0.38, right=0.97, top=0.93, bottom=0.07)

        else:  # ยอดรายเดือน (รวม)
            monthly_totals = [float(pivot[m].sum()) / 1_000 for m in months]
            fig = Figure(figsize=(10, 5), facecolor="white")
            ax  = fig.add_subplot(111)

            bar_colors = ["#2563EB"] * len(months)
            bars2 = ax.bar(m_names, monthly_totals, color=bar_colors, width=0.6)
            y_max = max(monthly_totals) * 1.18 if monthly_totals else 1
            ax.set_ylim(0, y_max)
            for bar, v in zip(bars2, monthly_totals):
                ax.text(bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + y_max * 0.01,
                        _fmt_val(v), ha="center", fontsize=9)
            ax.set_ylabel("ยอดขาย (พัน บาท)", fontsize=10)
            ax.set_title(
                f"ยอดขายรายเดือน — ปี {self._year_stored}",
                fontsize=12, fontweight="bold", loc="left", pad=10)
            ax.yaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
            ax.grid(axis="y", alpha=0.25, linestyle="--")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            fig.subplots_adjust(left=0.12, right=0.97, top=0.91, bottom=0.10)

        canvas = FigureCanvasTkAgg(fig, master=self._chart_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._chart_canvas_widget = canvas

    # ── header sort ──────────────────────────────────────────────────────────

    def _on_ch_press(self, event):
        self._header_press_x = event.x

    def _on_ch_release(self, event):
        if abs(event.x - self._header_press_x) > 6:
            return  # drag (column resize) — ไม่ sort
        import bisect
        x = self._sheet.MT.CH.canvasx(event.x)
        positions = self._sheet.MT.col_positions
        if not positions or len(positions) < 2:
            return
        col = bisect.bisect_right(positions, x) - 1
        col = max(0, min(col, len(positions) - 2))
        self._on_header_click(col)

    def _on_header_click(self, col: int):
        """เรียงข้อมูลตาม column ที่คลิก — สลับ ▲/▼"""
        if self._pivot_base is None:
            return

        months  = self._months_stored
        pivot   = self._pivot_base.copy()
        n_month = len(months)

        # กำหนด column DB ที่จะ sort
        if col == 0:
            sort_key = "customer_id"
            numeric  = False
        elif col == 1:
            sort_key = "customer_name"
            numeric  = False
        elif col == 2:
            sort_key = "sale_key"
            numeric  = False
        elif col == n_month + 3:          # คอลัมน์ "รวม"
            sort_key = "_total"
            numeric  = True
        elif 3 <= col < n_month + 3:      # คอลัมน์เดือน
            sort_key = months[col - 3]
            numeric  = True
        else:
            return

        # toggle: ถ้าคลิก col เดิม → สลับ asc/desc, ถ้า col ใหม่ → desc ก่อน (ตัวเลข) / asc (text)
        if self._sort_col_idx == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col_idx = col
            self._sort_asc = not numeric   # ตัวเลข: เริ่ม desc | text: เริ่ม asc

        pivot = pivot.sort_values(sort_key, ascending=self._sort_asc).reset_index(drop=True)
        self._draw_table(pivot, months, self._year_stored)

    def _draw_table(self, pivot, months, year_thai):
        """วาด sheet จาก pivot ที่ส่งมา (ใช้ได้ทั้งตอน render ครั้งแรกและตอน sort)"""
        self._sheet.dehighlight_all()

        m_names = [THAI_MONTHS_SHORT.get(int(m), str(m)) for m in months]

        # สร้าง header พร้อม sort indicator
        def _hdr(text, col_idx):
            if self._sort_col_idx == col_idx:
                return text + (" ▲" if self._sort_asc else " ▼")
            return text

        base_hdrs   = ["รหัสลูกค้า", "ชื่อลูกค้า", "รหัสพนักงาน"]
        month_hdrs  = [_hdr(mn, 3 + i) for i, mn in enumerate(m_names)]
        total_hdr   = _hdr("รวม", 3 + len(months))
        headers     = [_hdr(b, i) for i, b in enumerate(base_hdrs)] + month_hdrs + [total_hdr]
        self._sheet.headers(headers)

        def _fmt(v):
            v = float(v)
            return f"{v:,.0f}" if v else ""

        rows, num_matrix = [], []
        for _, r in pivot.iterrows():
            row = [r["customer_id"], r["customer_name"], r["sale_key"]]
            num_row = []
            for m in months:
                v = float(r.get(m, 0))
                row.append(_fmt(v))
                num_row.append(v)
            total = float(r["_total"])
            row.append(_fmt(total))
            num_row.append(total)
            rows.append(row)
            num_matrix.append(num_row)

        # summary row (ไม่เปลี่ยนตาม sort)
        base = self._pivot_base if self._pivot_base is not None else pivot
        s_row = ["", "รวมทั้งหมด", ""]
        for m in months:
            s_row.append(_fmt(float(base[m].sum())))
        s_row.append(_fmt(float(base["_total"].sum())))
        rows.append(s_row)

        self._sheet.set_sheet_data(rows)

        widths = [88, 230, 105] + [82] * len(months) + [100]
        self._sheet.set_column_widths(widths)
        self._sheet.align_columns(
            columns=list(range(3, len(headers))), align="right")

        # salesperson badge colors
        _fallback: dict = {}
        def _sale_color(key):
            if key in self._SALE_COLORS:
                return self._SALE_COLORS[key]
            if key not in _fallback:
                _fallback[key] = self._SALE_PALETTE[len(_fallback) % len(self._SALE_PALETTE)]
            return _fallback[key]

        for ri, (_, r) in enumerate(pivot.iterrows()):
            bg, fg = _sale_color(r["sale_key"])
            self._sheet.highlight_cells(row=ri, column=2, bg=bg, fg=fg, redraw=False)

        # heat map per column
        n_data = len(rows) - 1
        n_num  = len(months) + 1
        col_maxes = [
            max((num_matrix[ri][ci] for ri in range(n_data)), default=0)
            for ci in range(n_num)
        ]
        for ri in range(n_data):
            for ci in range(n_num):
                bg, fg = self._heat_color(num_matrix[ri][ci], col_maxes[ci])
                if bg:
                    self._sheet.highlight_cells(
                        row=ri, column=3 + ci, bg=bg, fg=fg, redraw=False)

        self._sheet.highlight_rows(
            rows=[len(rows) - 1], bg="#1E3A5F", fg="white", redraw=False)
        self._sheet.highlight_columns(
            columns=[0, 1, 2], bg="#F8FAFC", fg="#0F172A", redraw=False)

        self._sheet.redraw()
        self._info_lbl.configure(
            text=f"{len(pivot)} ลูกค้า | ปี {year_thai} | {len(months)} เดือน")
