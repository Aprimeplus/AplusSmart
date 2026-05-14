import tkinter as tk
from tkinter import ttk, messagebox
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkFont,
                           CTkScrollableFrame, CTkTabview, CTkOptionMenu)
import pandas as pd
import numpy as np
import traceback
import os
import calendar
from datetime import datetime
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import FuncFormatter
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
import matplotlib.patheffects as pe
import matplotlib

try:
    font_path = os.path.join('resources', 'THSarabunNew.ttf')
    if os.path.exists(font_path):
        from matplotlib.font_manager import fontManager
        fontManager.addfont(font_path)
        matplotlib.rc('font', family='TH Sarabun New')
except Exception:
    pass

from daily_report_widget import DailyReportWidget


# ─────────────────────────────────────────────────────────────────────────────
#  Sale Revenue Report widget (ported from HRScreen, standalone)
# ─────────────────────────────────────────────────────────────────────────────
class SaleRevenueWidget(CTkFrame):
    EXCLUDE_SALE_KEYS = {'s', 'd', 'p', 'mp', 'ms', 'hr', 'sm', 'Sale Center', 'Pimhathai'}
    PERSON_MERGE = {
        'VOW-P': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ภาณุพงศ์'),
        'VOW-S': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ฐรินทร์ญา'),
    }

    def __init__(self, master, app_container, **kwargs):
        super().__init__(master, **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.theme = app_container.THEME.get("hr", {"primary": "#16A34A", "header": "#15803D", "bg": "#F0FDF4"})

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม",
                            "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน",
                            "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}

        self.label_font_bold = CTkFont(size=12, weight="bold", family="Roboto")
        self.header_font = CTkFont(size=14, weight="bold", family="Roboto")

        self.sales_view_mode = 'chart'
        self.custom_target_start = None
        self.custom_target_end = None
        self.sales_target_period_var = tk.StringVar(value="เดือนนี้")
        self._chart_canvas = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_toolbar()
        self._build_chart_area()

    def _build_toolbar(self):
        bar = CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, padx=10, pady=(10, 4), sticky="ew")

        current_year = datetime.now().year
        years = [str(y + 543) for y in range(current_year - 2, current_year + 3)]

        def month_picker(label):
            f = CTkFrame(bar, fg_color="transparent")
            f.pack(side="left", padx=8)
            CTkLabel(f, text=label, font=self.label_font_bold).pack(side="left", padx=(0, 5))
            m_var = tk.StringVar(value=self.thai_months[datetime.now().month - 1])
            CTkOptionMenu(f, variable=m_var, values=self.thai_months, width=110).pack(side="left", padx=2)
            y_var = tk.StringVar(value=str(current_year + 543))
            CTkOptionMenu(f, variable=y_var, values=years, width=80).pack(side="left", padx=2)
            return m_var, y_var

        self.start_m_var, self.start_y_var = month_picker("จากรอบ:")
        self.end_m_var, self.end_y_var = month_picker("ถึงรอบ:")

        CTkButton(bar, text="🔍 ค้นหา", width=100, fg_color=self.theme["primary"],
                  command=self._on_search).pack(side="left", padx=20)

        # Toggle
        toggle = CTkFrame(bar, fg_color="transparent")
        toggle.pack(side="right", padx=(0, 5))

        def _set_view(mode):
            self.sales_view_mode = mode
            if mode == 'chart':
                self._btn_chart.configure(fg_color=self.theme["primary"], text_color="white")
                self._btn_table.configure(fg_color="#E2E8F0", text_color="#475569")
            else:
                self._btn_table.configure(fg_color=self.theme["primary"], text_color="white")
                self._btn_chart.configure(fg_color="#E2E8F0", text_color="#475569")
            self._refresh()

        self._btn_chart = CTkButton(toggle, text="📊 กราฟ", width=90,
                                    fg_color=self.theme["primary"], text_color="white",
                                    corner_radius=6, command=lambda: _set_view('chart'))
        self._btn_chart.pack(side="left", padx=2)

        self._btn_table = CTkButton(toggle, text="📋 ตาราง", width=90,
                                    fg_color="#E2E8F0", text_color="#475569",
                                    corner_radius=6, command=lambda: _set_view('table'))
        self._btn_table.pack(side="left", padx=2)

    def _build_chart_area(self):
        self.chart_frame = CTkFrame(self, border_width=1, corner_radius=10)
        self.chart_frame.grid(row=1, column=0, padx=10, pady=(4, 10), sticky="nsew")
        self.after(100, self._refresh)

    def _show_loading(self):
        for w in self.chart_frame.winfo_children():
            w.destroy()
        lbl = CTkLabel(self.chart_frame, text="กำลังโหลดข้อมูล...",
                       font=CTkFont(size=18, slant="italic"), text_color="gray50")
        lbl.pack(expand=True, pady=20)
        self.update_idletasks()
        return lbl

    def _on_search(self):
        try:
            s_m = self.thai_months.index(self.start_m_var.get()) + 1
            s_y = int(self.start_y_var.get()) - 543
            e_m = self.thai_months.index(self.end_m_var.get()) + 1
            e_y = int(self.end_y_var.get()) - 543

            start = datetime(s_y, s_m, 1)
            last_day = calendar.monthrange(e_y, e_m)[1]
            end = datetime(e_y, e_m, last_day)

            if start > end:
                messagebox.showerror("รอบเดือนไม่ถูกต้อง",
                                     "รอบเริ่มต้นต้องมาก่อนหรือเท่ากับรอบสิ้นสุด",
                                     parent=self)
                return

            self.custom_target_start = start
            self.custom_target_end = end
            self.sales_target_period_var.set("กำหนดช่วงเวลาเอง...")
            self._refresh()
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()

    def _refresh(self):
        loading = self._show_loading()
        try:
            df = self._get_data()
            loading.destroy()
            if self.sales_view_mode == 'table':
                self._draw_table(df)
            else:
                self._draw_chart(df)
        except Exception as e:
            loading.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()

    # ── Data query ────────────────────────────────────────────────────────────
    def _get_data(self):
        today = datetime.now()
        current_year = today.year
        params = []
        clauses = []
        target_multiplier = 1.0
        period = self.sales_target_period_var.get()

        if period == "กำหนดช่วงเวลาเอง..." and self.custom_target_start:
            s = self.custom_target_start
            e = self.custom_target_end
            clauses.append(
                "MAKE_DATE(c.commission_year, c.commission_month, 1) "
                "BETWEEN %s::date AND %s::date"
            )
            params.extend([s.strftime("%Y-%m-%d"), e.strftime("%Y-%m-%d")])
            months_diff = (e.year - s.year) * 12 + (e.month - s.month) + 1
            target_multiplier = max(1, months_diff)

        elif period == "ปีนี้":
            clauses.append("c.commission_year = %s")
            params.append(current_year)
            target_multiplier = 12.0

        elif period in ("Q1", "Q2", "Q3", "Q4"):
            qs = {"Q1": (1,2,3), "Q2": (4,5,6), "Q3": (7,8,9), "Q4": (10,11,12)}
            ms = qs[period]
            clauses.append(f"c.commission_month IN ({','.join(map(str, ms))})")
            clauses.append("c.commission_year = %s")
            params.append(current_year)
            target_multiplier = 3.0

        elif period in self.thai_month_map:
            clauses.append("c.commission_month = %s")
            params.append(self.thai_month_map[period])
            clauses.append("c.commission_year = %s")
            params.append(current_year)

        else:  # fallback / เดือนนี้
            clauses.append("c.commission_month = %s")
            params.append(today.month)
            clauses.append("c.commission_year = %s")
            params.append(current_year)

        date_filter = " AND ".join(clauses)

        query = f"""
            SELECT
                su.sale_name,
                su.sale_key,
                COALESCE(su.sales_target, 0) * %s AS sales_target,
                COALESCE(SUM(c.total_sales), 0) AS total_sales,
                0 AS total_outstanding
            FROM sales_users su
            LEFT JOIN commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '')
                      = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND {date_filter}
            WHERE su.status = 'Active'
            GROUP BY su.sale_name, su.sale_key, su.sales_target, su.role
            HAVING (su.role = 'Sale' OR COALESCE(SUM(c.total_sales), 0) > 0)
            ORDER BY su.sale_name ASC;
        """
        final_params = [target_multiplier] + params
        df = pd.read_sql_query(query, self.pg_engine, params=tuple(final_params))
        df['sales_target'] = df['sales_target'].fillna(0)
        df['total_sales'] = df['total_sales'].fillna(0)
        df['total_outstanding'] = df['total_outstanding'].fillna(0)
        return df

    # ── Shared preprocessing ──────────────────────────────────────────────────
    def _preprocess(self, data_df):
        df2 = data_df[~data_df['sale_key'].isin(self.EXCLUDE_SALE_KEYS)].copy()

        def _group_name(row):
            k = str(row['sale_key']).strip()
            return self.PERSON_MERGE[k][0] if k in self.PERSON_MERGE else row['sale_name']

        def _seg_label(row):
            k = str(row['sale_key']).strip()
            return self.PERSON_MERGE[k][1] if k in self.PERSON_MERGE else row['sale_name']

        df2['_group'] = df2.apply(_group_name, axis=1)
        df2['_seg_label'] = df2.apply(_seg_label, axis=1)

        people_data = []
        for name, grp in df2.groupby('_group', sort=False):
            sub_items = [
                {'sale_key': row['sale_key'], 'label': row['_seg_label'],
                 'sales': float(row['total_sales'])}
                for _, row in grp.iterrows()
            ]
            sub_items.sort(key=lambda s: s['sales'], reverse=True)
            people_data.append({
                'name': name,
                'total_sales': float(grp['total_sales'].sum()),
                'total_outstanding': float(grp['total_outstanding'].sum()),
                'target': float(grp['sales_target'].sum()),
                'sub_items': sub_items,
            })
        people_data.sort(key=lambda p: p['total_sales'], reverse=True)
        return [p for p in people_data if p['total_sales'] > 0 or p['target'] > 0]

    # ── Chart ──────────────────────────────────────────────────────────────────
    def _draw_chart(self, data_df):
        if self._chart_canvas:
            try:
                self._chart_canvas.get_tk_widget().destroy()
            except Exception:
                pass
            self._chart_canvas = None
        for w in self.chart_frame.winfo_children():
            w.destroy()

        if data_df.empty:
            CTkLabel(self.chart_frame, text="ไม่พบข้อมูลพนักงานขาย",
                     font=self.header_font).pack(expand=True)
            return

        people_data = self._preprocess(data_df)
        if not people_data:
            CTkLabel(self.chart_frame, text="ไม่พบข้อมูลพนักงานขาย",
                     font=self.header_font).pack(expand=True)
            return

        n = len(people_data)
        names = [p['name'] for p in people_data]
        targets = [p['target'] for p in people_data]
        sales = [p['total_sales'] for p in people_data]

        ACHIEVE_COLORS = {
            'green':  ('#22C55E', '#86EFAC'),
            'yellow': ('#F59E0B', '#FCD34D'),
            'red':    ('#EF4444', '#FCA5A5'),
            'gray':   ('#94A3B8', '#CBD5E1'),
        }

        def akey(s, t):
            if t <= 0: return 'gray'
            return 'green' if s >= t else ('yellow' if s >= t * 0.7 else 'red')

        pct_labels = [f"{s/t*100:.0f}%" if t > 0 else "N/A" for s, t in zip(sales, targets)]
        max_t = max(targets or [1])
        BG, GRID_C = '#F8FAFC', '#E2E8F0'

        chart_width = max(10, n * 1.6)
        fig = Figure(figsize=(chart_width, 7.2), dpi=100, facecolor=BG)
        ax = fig.add_subplot(111)
        ax.set_facecolor(BG)

        x = np.arange(n)
        width = 0.65

        # Gap bars (grey background)
        for i, p in enumerate(people_data):
            gap = max(0.0, p['target'] - p['total_sales'])
            if gap > 0:
                ax.bar(x[i], gap, width, bottom=p['total_sales'],
                       color='#E2E8F0', zorder=2, linewidth=0)

        # Stacked sales bars
        for i, p in enumerate(people_data):
            colors = ACHIEVE_COLORS[akey(p['total_sales'], p['target'])]
            multi_id = len([s for s in p['sub_items'] if s['sales'] > 0]) > 1
            bottom = 0.0
            for idx, sub in enumerate(p['sub_items']):
                seg_h = sub['sales']
                if seg_h <= 0:
                    continue
                color = colors[min(idx, 1)]
                is_partner = (idx > 0)
                ax.bar(x[i], seg_h, width, bottom=bottom, color=color, zorder=3,
                       linewidth=0.8 if is_partner else 0,
                       edgecolor='white' if is_partner else 'none',
                       hatch='//' if is_partner else None, alpha=0.92)

                mid_y = bottom + seg_h / 2
                seg_top = bottom + seg_h
                t_line = p['target']
                if t_line > 0 and bottom < t_line < seg_top:
                    if (t_line - bottom) >= seg_h * 0.35:
                        mid_y = bottom + (t_line - bottom) / 2
                    else:
                        mid_y = t_line + (seg_top - t_line) / 2

                txt_color = '#1a5c1a' if is_partner else 'white'
                stroke_color = 'white' if txt_color != 'white' else '#00000066'
                fx = [pe.withStroke(linewidth=3, foreground=stroke_color)]

                if multi_id:
                    if seg_h > max_t * 0.10:
                        ax.text(x[i], mid_y, f"{sub['label']}\n{seg_h:,.0f}",
                                ha='center', va='center', fontsize=14, weight='medium',
                                color=txt_color, zorder=7, linespacing=1.4, path_effects=fx)
                    elif seg_h > max_t * 0.04:
                        ax.text(x[i], mid_y, f"{seg_h:,.0f}",
                                ha='center', va='center', fontsize=13, weight='medium',
                                color=txt_color, zorder=7, path_effects=fx)
                else:
                    if seg_h > max_t * 0.06:
                        ax.text(x[i], mid_y, f"{seg_h:,.0f}",
                                ha='center', va='center', fontsize=14, weight='medium',
                                color=txt_color, zorder=7, path_effects=fx)
                bottom += seg_h

        # Target dashed lines + label
        half = width / 2
        for i, t in enumerate(targets):
            if t > 0:
                ax.hlines(t, x[i]-half, x[i]+half, colors='#6366F1',
                          linewidths=2.2, linestyles='--', zorder=5)
                ax.text(x[i], t + max_t * 0.012, f"Target  {t:,.0f}",
                        ha='center', va='bottom', fontsize=9.5, weight='bold',
                        color='#6366F1', zorder=6)

        # % badges
        for i, (p, pct) in enumerate(zip(people_data, pct_labels)):
            s, t = p['total_sales'], p['target']
            if s == 0 and t > 0:
                ax.text(x[i], t * 0.5, "ยังไม่มี\nข้อมูล SO",
                        ha='center', va='center', fontsize=11, weight='bold',
                        color='#94A3B8', zorder=8, style='italic', linespacing=1.4)
            else:
                pct_y = max(s, t) + max_t * 0.16
                ax.text(x[i], pct_y, pct, ha='center', va='bottom',
                        fontsize=16, weight='medium', color='black', zorder=8,
                        path_effects=[pe.withStroke(linewidth=2, foreground='white')])
                if s <= max_t * 0.08:
                    ax.text(x[i], s + max_t * 0.012, f"{s:,.0f}",
                            ha='center', va='bottom', fontsize=12, weight='bold',
                            color='black', zorder=7)

        # Axes decoration
        max_sales = max((p['total_sales'] for p in people_data), default=0)
        ax.set_ylim(0, max(max_sales, max_t) + max_t * 0.55)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:,.0f}'))
        ax.tick_params(axis='y', labelsize=13)
        ax.tick_params(axis='x', pad=8)

        def _fmt(nm, tgt):
            short = nm.replace(' / ', '\n').replace(' ', '\n', 1)
            return f"{short}\nเป้า {tgt:,.0f}" if tgt > 0 else short

        ax.set_xticks(x)
        ax.set_xticklabels([_fmt(nm, tgt) for nm, tgt in zip(names, targets)],
                           rotation=0, ha='center', fontsize=10, weight='medium',
                           color='black', linespacing=1.35)
        ax.set_ylabel('จำนวนเงิน (บาท)', fontsize=12, weight='medium', color='black', labelpad=10)

        team_sales = sum(p['total_sales'] for p in people_data)
        team_target = sum(p['target'] for p in people_data)
        team_pct = (team_sales / team_target * 100) if team_target > 0 else 0
        n_hit = sum(1 for p in people_data if p['target'] > 0 and p['total_sales'] >= p['target'])
        n_miss = sum(1 for p in people_data if p['target'] > 0 and p['total_sales'] < p['target'])
        summary = (f"ทีมรวม  {team_sales:,.0f} / {team_target:,.0f} บาท"
                   f"  ({team_pct:.0f}%)     ถึงเป้า {n_hit} คน  ·  ยังไม่ถึง {n_miss} คน")

        ax.set_title('Sale Revenue Report  (เฉพาะ SO ที่คิดค่าคอมแล้ว)',
                     fontsize=16, weight='semibold', color='#0F172A', loc='left', pad=36)
        ax.text(0, 1.015, summary, transform=ax.transAxes,
                fontsize=10.5, color='#64748B', ha='left', va='bottom', weight='medium')

        ax.yaxis.grid(True, color=GRID_C, linewidth=0.8, zorder=0)
        ax.set_axisbelow(True)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(GRID_C)
        ax.spines['bottom'].set_color(GRID_C)
        ax.tick_params(colors='black')

        has_multi_id = any(len(p['sub_items']) > 1 for p in people_data)
        legend_items = [
            Patch(facecolor='#22C55E', label='≥ 100% เป้า'),
            Patch(facecolor='#F59E0B', label='70–99% เป้า'),
            Patch(facecolor='#EF4444', label='< 70% เป้า'),
            Line2D([0], [0], color='#6366F1', lw=2, linestyle='--', label='เป้าหมาย'),
            Patch(facecolor='#E2E8F0', edgecolor='#CBD5E1', label='ส่วนที่ยังไม่ถึงเป้า'),
        ]
        if has_multi_id:
            legend_items.append(Patch(facecolor='#86EFAC', label='ยอดขายของพาร์ทเนอร์'))

        ax.legend(handles=legend_items, loc='upper right', bbox_to_anchor=(1.0, 1.0),
                  ncol=2, frameon=True, framealpha=0.95, edgecolor='#CBD5E1',
                  prop={'weight': 'bold', 'size': 10},
                  borderpad=0.7, labelspacing=0.4, columnspacing=1.0)

        fig.tight_layout(rect=[0, 0.05, 1, 1])

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._chart_canvas = canvas

    # ── Table ──────────────────────────────────────────────────────────────────
    def _draw_table(self, data_df):
        for w in self.chart_frame.winfo_children():
            w.destroy()

        if data_df.empty:
            CTkLabel(self.chart_frame, text="ไม่พบข้อมูล",
                     font=self.header_font).pack(expand=True)
            return

        people_data = self._preprocess(data_df)
        rows = []
        for p in people_data:
            t, s = p['target'], p['total_sales']
            rows.append({
                'name': p['name'],
                'target': t,
                'sales': s,
                'pct': (s / t * 100) if t > 0 else 0.0,
                'diff': s - t,
            })
        rows.sort(key=lambda r: r['sales'], reverse=True)

        total_t = sum(r['target'] for r in rows)
        total_s = sum(r['sales'] for r in rows)
        total_pct = (total_s / total_t * 100) if total_t > 0 else 0.0
        total_diff = total_s - total_t

        try:
            s_m = self.thai_months.index(self.start_m_var.get()) + 1
            e_m = self.thai_months.index(self.end_m_var.get()) + 1
            s_y = int(self.start_y_var.get())
            e_y = int(self.end_y_var.get())
            if s_m == e_m and s_y == e_y:
                period_label = f"{self.start_m_var.get()} {s_y}"
            else:
                period_label = f"{self.start_m_var.get()} {s_y} – {self.end_m_var.get()} {e_y}"
        except Exception:
            period_label = "รอบที่เลือก"

        outer = CTkScrollableFrame(self.chart_frame, fg_color="white", corner_radius=10)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        CTkLabel(outer, text=f"Sale Revenue Report  —  {period_label}",
                 font=CTkFont(size=15, weight="bold"), text_color="#0F172A"
                 ).pack(anchor="w", padx=16, pady=(12, 8))

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("MgmtTable.Treeview",
                        background="white", foreground="#1E293B",
                        rowheight=34, fieldbackground="white",
                        font=('TH Sarabun New', 13))
        style.configure("MgmtTable.Treeview.Heading",
                        background="#D1FAE5", foreground="#065F46",
                        font=('TH Sarabun New', 13, 'bold'), relief="flat")
        style.map("MgmtTable.Treeview",
                  background=[('selected', '#EDE9FE')],
                  foreground=[('selected', '#1E293B')])

        cols = ('name', 'target', 'sales', 'pct', 'diff')
        tree = ttk.Treeview(outer, columns=cols, show='headings',
                            style="MgmtTable.Treeview", height=len(rows) + 1)
        tree.heading('name',   text='พนักงาน')
        tree.heading('target', text='เป้าหมาย (บาท)')
        tree.heading('sales',  text='ยอดขายจริง (บาท)')
        tree.heading('pct',    text='%')
        tree.heading('diff',   text='ส่วนต่าง (บาท)')
        tree.column('name',   width=200, anchor='w')
        tree.column('target', width=170, anchor='e')
        tree.column('sales',  width=170, anchor='e')
        tree.column('pct',    width=90,  anchor='center')
        tree.column('diff',   width=170, anchor='e')

        tree.tag_configure('odd',      background='#F8FAFC')
        tree.tag_configure('even',     background='white')
        tree.tag_configure('negative', foreground='#DC2626')
        tree.tag_configure('positive', foreground='#16A34A')
        tree.tag_configure('total',    background='#EFF6FF',
                           font=('TH Sarabun New', 13, 'bold'),
                           foreground='#1D4ED8')

        for idx, r in enumerate(rows):
            diff_str = f"+{r['diff']:,.0f}" if r['diff'] >= 0 else f"{r['diff']:,.0f}"
            tag_row  = 'odd' if idx % 2 else 'even'
            tag_diff = 'positive' if r['diff'] >= 0 else 'negative'
            tree.insert('', 'end', tags=(tag_row, tag_diff), values=(
                r['name'],
                f"{r['target']:,.0f}" if r['target'] > 0 else '—',
                f"{r['sales']:,.0f}",
                f"{r['pct']:.1f}%",
                diff_str if r['target'] > 0 else '—',
            ))

        total_diff_str = f"+{total_diff:,.0f}" if total_diff >= 0 else f"{total_diff:,.0f}"
        tree.insert('', 'end', tags=('total',), values=(
            'รวมทีม',
            f"{total_t:,.0f}",
            f"{total_s:,.0f}",
            f"{total_pct:.1f}%",
            total_diff_str,
        ))
        tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))


# ─────────────────────────────────────────────────────────────────────────────
#  Management Report Screen
# ─────────────────────────────────────────────────────────────────────────────
class ManagementReportScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        tabs = CTkTabview(self, segmented_button_selected_color="#1E40AF")
        tabs.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Tab 1 — SO Daily Report
        tab_daily = tabs.add("📋 SO Daily Report")
        tab_daily.grid_columnconfigure(0, weight=1)
        tab_daily.grid_rowconfigure(0, weight=1)
        daily_widget = DailyReportWidget(
            master=tab_daily,
            app_container=app_container,
            fg_color="transparent",
        )
        daily_widget.grid(row=0, column=0, sticky="nsew")

        # Tab 2 — Sale Revenue Report
        tab_revenue = tabs.add("📈 Sale Revenue Report")
        tab_revenue.grid_columnconfigure(0, weight=1)
        tab_revenue.grid_rowconfigure(0, weight=1)
        revenue_widget = SaleRevenueWidget(
            master=tab_revenue,
            app_container=app_container,
            fg_color="transparent",
        )
        revenue_widget.grid(row=0, column=0, sticky="nsew")
