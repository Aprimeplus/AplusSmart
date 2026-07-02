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
        area = CTkFrame(self, fg_color="transparent")
        area.grid(row=0, column=1, sticky="nsew", padx=(4, 6), pady=6)
        area.grid_columnconfigure(0, weight=1)
        area.grid_rowconfigure(1, weight=1)

        hdr = CTkFrame(area, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        CTkLabel(hdr, text="👥  Customer Monitoring — ยอดขายรายลูกค้า",
                 font=CTkFont(size=14, weight="bold"),
                 text_color="#0F172A").pack(side="left")
        self._info_lbl = CTkLabel(hdr, text="",
                                   font=CTkFont(size=11), text_color="#64748B")
        self._info_lbl.pack(side="right", padx=8)

        self._sheet = Sheet(
            area,
            theme="light blue",
            header_font=("Tahoma", 11, "bold"),
            data_font=("Tahoma", 11),
            default_column_width=80,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        self._sheet.grid(row=1, column=0, sticky="nsew")
        self._sheet.enable_bindings(
            "column_width_resize",
            "double_click_column_resize",
            "right_click_popup_menu",
            "rc_select",
            "copy",
        )

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

    # salesperson badge colors — bg / fg pairs (assigned dynamically in order of first appearance)
    _SALE_PALETTE = [
        ("#FEF3C7", "#92400E"),  # amber
        ("#FCE7F3", "#9D174D"),  # pink
        ("#D1FAE5", "#065F46"),  # green
        ("#EDE9FE", "#4C1D95"),  # violet
        ("#FEE2E2", "#991B1B"),  # red
        ("#CFFAFE", "#164E63"),  # cyan
        ("#FEF9C3", "#713F12"),  # yellow
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
        m_names = [THAI_MONTHS_SHORT.get(int(m), str(m)) for m in months]

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

        headers = ["รหัสลูกค้า", "ชื่อลูกค้า", "รหัสพนักงาน"] + m_names + ["รวม"]
        self._sheet.headers(headers)

        def _fmt(v):
            v = float(v)
            return f"{v:,.0f}" if v else ""

        # build rows + keep numeric matrix for heat map
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

        # summary row
        s_row = ["", "รวมทั้งหมด", ""]
        for m in months:
            s_row.append(_fmt(float(pivot[m].sum())))
        s_row.append(_fmt(float(pivot["_total"].sum())))
        rows.append(s_row)

        self._sheet.set_sheet_data(rows)

        # column widths
        widths = [88, 230, 105] + [82] * len(months) + [100]
        self._sheet.set_column_widths(widths)

        # right-align numeric columns
        self._sheet.align_columns(
            columns=list(range(3, len(headers))), align="right")

        # ── salesperson badge colors (col 2) ─────────────────────────────
        sale_color: dict = {}
        for key in pivot["sale_key"].unique():
            if key not in sale_color:
                idx = len(sale_color) % len(self._SALE_PALETTE)
                sale_color[key] = self._SALE_PALETTE[idx]
        for ri, (_, r) in enumerate(pivot.iterrows()):
            key = r["sale_key"]
            if key in sale_color:
                bg, fg = sale_color[key]
                self._sheet.highlight_cells(
                    row=ri, column=2, bg=bg, fg=fg, redraw=False)

        # ── heat map per column ───────────────────────────────────────────
        n_data = len(rows) - 1          # exclude summary row
        n_num  = len(months) + 1        # month cols + total col
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

        # summary row: dark navy
        self._sheet.highlight_rows(
            rows=[len(rows) - 1], bg="#1E3A5F", fg="white", redraw=False)

        # identity columns: subtle tint
        self._sheet.highlight_columns(
            columns=[0, 1, 2], bg="#F8FAFC", fg="#0F172A", redraw=False)

        self._sheet.redraw()
        self._info_lbl.configure(
            text=f"{len(pivot)} ลูกค้า | ปี {year_thai} | {len(months)} เดือน")
