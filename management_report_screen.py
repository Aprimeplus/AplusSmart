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

        # ── resize / draw guards (same pattern as SalesTargetWidget) ──────────
        self._first_map_done  = False
        self._frame_ready_job = None
        self._ready_frame_w   = 0
        self._ready_frame_h   = 0
        self._drawing_chart   = False
        self._last_chart_fw   = 0
        self._last_chart_fh   = 0

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_toolbar()
        self._build_chart_area()

        self.bind("<Map>", self._on_map)

    def _on_map(self, event=None):
        """ทำงานครั้งเดียว — bind Configure บน chart_frame แบบถาวร"""
        if self._first_map_done:
            return
        self._first_map_done = True
        self.unbind('<Map>')
        self.chart_frame.bind('<Configure>', self._on_chart_frame_configure)

    def _on_chart_frame_configure(self, event):
        """debounce resize — suppress ระหว่าง _refresh รัน"""
        if event.width < 200 or event.height < 100:
            return
        if self._drawing_chart:
            return
        fw, fh = event.width, event.height
        if fw == self._last_chart_fw and fh == self._last_chart_fh:
            return
        if self._frame_ready_job:
            self.after_cancel(self._frame_ready_job)
        self._frame_ready_job = self.after(200, lambda: self._on_chart_resize(fw, fh))

    def _on_chart_resize(self, fw, fh):
        """Fired after debounce — initial draw หรือ window-resize redraw"""
        self._frame_ready_job = None
        if self.sales_view_mode not in ('chart', 'monthly'):
            return
        if self._chart_canvas is None:
            self._ready_frame_w = fw
            self._ready_frame_h = fh
            self._refresh()
        else:
            try:
                actual_fw = self.chart_frame.winfo_width()
                actual_fh = self.chart_frame.winfo_height()
                real_fw = actual_fw if actual_fw > 100 else fw
                real_fh = actual_fh if actual_fh > 100 else fh
                self._last_chart_fw = real_fw
                self._last_chart_fh = real_fh
                cw = max(real_fw - 22, 100)
                ch = max(real_fh - 22, 100)
                fig = self._chart_canvas.figure
                fig.set_size_inches(cw / fig.dpi, ch / fig.dpi)
                fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.12)
                self.after(10, self._chart_canvas.draw)
            except Exception:
                traceback.print_exc()

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
            for btn, m in [(self._btn_chart, 'chart'), (self._btn_table, 'table'), (self._btn_monthly, 'monthly')]:
                if m == mode:
                    btn.configure(fg_color=self.theme["primary"], text_color="white")
                else:
                    btn.configure(fg_color="#E2E8F0", text_color="#475569")
            self._refresh()

        self._btn_chart = CTkButton(toggle, text="📊 กราฟ", width=90,
                                    fg_color=self.theme["primary"], text_color="white",
                                    corner_radius=6, command=lambda: _set_view('chart'))
        self._btn_chart.pack(side="left", padx=2)

        self._btn_table = CTkButton(toggle, text="📋 ตาราง", width=90,
                                    fg_color="#E2E8F0", text_color="#475569",
                                    corner_radius=6, command=lambda: _set_view('table'))
        self._btn_table.pack(side="left", padx=2)

        self._btn_monthly = CTkButton(toggle, text="📅 รายเดือน", width=100,
                                      fg_color="#E2E8F0", text_color="#475569",
                                      corner_radius=6, command=lambda: _set_view('monthly'))
        self._btn_monthly.pack(side="left", padx=2)

    def _build_chart_area(self):
        self.chart_frame = CTkFrame(self, border_width=1, corner_radius=10)
        self.chart_frame.grid(row=1, column=0, padx=10, pady=(4, 10), sticky="nsew")

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
        if self._frame_ready_job:
            self.after_cancel(self._frame_ready_job)
            self._frame_ready_job = None
        pre_fw = self.chart_frame.winfo_width()
        pre_fh = self.chart_frame.winfo_height()
        if pre_fw > 100:
            self._ready_frame_w = pre_fw
            self._ready_frame_h = pre_fh
        self._drawing_chart = True
        loading = self._show_loading()
        try:
            df = self._get_data()
            loading.destroy()
            if self.sales_view_mode == 'table':
                self._draw_table(df)
            elif self.sales_view_mode == 'monthly':
                self._draw_monthly_chart(df)
            else:
                self._draw_chart(df)
        except Exception as e:
            loading.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            self._drawing_chart = False
            fw = self.chart_frame.winfo_width()
            fh = self.chart_frame.winfo_height()
            if fw > 100:
                self._last_chart_fw = fw
                self._last_chart_fh = fh

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
                0 AS total_outstanding,
                COALESCE(
                    SUM(cm_agg.total_gp) / NULLIF(SUM(cm_agg.total_sales_amt), 0) * 100,
                    0
                ) AS avg_margin,
                COALESCE(SUM(cm_agg.sales_normal), 0) AS sales_normal,
                COALESCE(SUM(cm_agg.sales_below),  0) AS sales_below
            FROM sales_users su
            LEFT JOIN commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '')
                      = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND {date_filter}
            LEFT JOIN (
                SELECT payout_id,
                       SUM(final_gp)           AS total_gp,
                       SUM(final_sales_amount) AS total_sales_amt,
                       SUM(CASE WHEN final_margin >= 10 THEN final_sales_amount ELSE 0 END) AS sales_normal,
                       SUM(CASE WHEN final_margin <  10 THEN final_sales_amount ELSE 0 END) AS sales_below
                FROM   commissions
                GROUP  BY payout_id
            ) cm_agg ON cm_agg.payout_id = c.id
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
        df['sales_normal'] = df['sales_normal'].fillna(0)
        df['sales_below']  = df['sales_below'].fillna(0)
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
            # avg_margin: weighted avg ตาม total_sales ของแต่ละ sale_key ในกลุ่ม
            total_s = float(grp['total_sales'].sum())
            if 'avg_margin' in grp.columns and total_s > 0:
                wtd_margin = float(
                    (grp['avg_margin'] * grp['total_sales']).sum() / total_s
                )
            else:
                wtd_margin = 0.0
            people_data.append({
                'name': name,
                'total_sales': total_s,
                'total_outstanding': float(grp['total_outstanding'].sum()),
                'target': float(grp['sales_target'].sum()),
                'avg_margin': wtd_margin,
                'sales_normal': float(grp['sales_normal'].sum()) if 'sales_normal' in grp.columns else 0.0,
                'sales_below':  float(grp['sales_below'].sum())  if 'sales_below'  in grp.columns else 0.0,
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

        _fw = getattr(self, '_ready_frame_w', 0) or self.chart_frame.winfo_width()
        _fh = getattr(self, '_ready_frame_h', 0) or self.chart_frame.winfo_height()
        self._ready_frame_w = 0
        self._ready_frame_h = 0
        if _fw <= 10:
            _fw = self.winfo_width() - 20
        if _fh <= 10:
            _fh = self.winfo_height() - 80
        cw = max(_fw - 22, 400)
        ch = max(_fh - 22, 350)
        fig = Figure(figsize=(cw / 100, ch / 100), dpi=100, facecolor=BG)
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

        fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.12)

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        widget = canvas.get_tk_widget()
        widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # อ่านขนาด canvas จริงหลัง layout settle แล้ว resize figure ให้พอดี
        self.update_idletasks()
        _cw = widget.winfo_width()
        _ch = widget.winfo_height()
        if _cw > 50 and _ch > 50:
            fig.set_size_inches(_cw / fig.dpi, _ch / fig.dpi)
            fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.12)

        canvas.draw()
        self._chart_canvas = canvas

    # ── Monthly chart ─────────────────────────────────────────────────────────
    PERSON_COLOR_MAP = {
        'กนกพร':    '#F9ED69',
        'ปิยะวรรณ': '#F97316',
        'ภาณุพงศ์': '#8B5CF6',
        'ฐรินทร์ญา':'#8B5CF6',
        'ไอยลดา':   '#38BDF8',
    }
    COLORS_MONTHLY_FALLBACK = ['#3B82F6','#22C55E','#EC4899','#14B8A6','#EF4444','#84CC16']

    def _get_monthly_data(self):
        params = []
        if self.custom_target_start and self.custom_target_end:
            date_sql = "MAKE_DATE(c.commission_year, c.commission_month, 1) BETWEEN %s::date AND %s::date"
            params.extend([self.custom_target_start.strftime("%Y-%m-%d"),
                           self.custom_target_end.strftime("%Y-%m-%d")])
        else:
            date_sql = "c.commission_year = %s"
            params.append(datetime.now().year)
        query = f"""
            SELECT su.sale_name, su.sale_key,
                   c.commission_month AS month,
                   c.commission_year  AS year,
                   COALESCE(SUM(c.total_sales), 0) AS total_sales
            FROM   sales_users su
            JOIN   commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND {date_sql}
            WHERE  su.status = 'Active' AND su.role = 'Sale'
            GROUP  BY su.sale_name, su.sale_key, c.commission_month, c.commission_year
            ORDER  BY c.commission_year, c.commission_month, su.sale_name
        """
        df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
        df['total_sales'] = df['total_sales'].fillna(0)
        return df

    def _draw_monthly_chart(self, _ignored_df):
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.ticker as mticker
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.transforms import blended_transform_factory

        # อ่านขนาด frame จริงก่อน destroy — เหมือน _draw_chart
        _fw = getattr(self, '_ready_frame_w', 0) or self.chart_frame.winfo_width()
        _fh = getattr(self, '_ready_frame_h', 0) or self.chart_frame.winfo_height()
        self._ready_frame_w = 0
        self._ready_frame_h = 0
        if _fw <= 10: _fw = self.winfo_width() - 20
        if _fh <= 10: _fh = self.winfo_height() - 80

        if self._chart_canvas:
            try: self._chart_canvas.get_tk_widget().destroy()
            except Exception: pass
            self._chart_canvas = None
        for w in self.chart_frame.winfo_children():
            w.destroy()

        data_df = self._get_monthly_data()

        if data_df.empty:
            CTkLabel(self.chart_frame, text="ไม่พบข้อมูล", font=self.header_font).pack(expand=True)
            return

        df = data_df[~data_df['sale_key'].isin(self.EXCLUDE_SALE_KEYS)].copy()
        df['_group'] = df.apply(
            lambda r: self.PERSON_MERGE[r['sale_key']][0] if r['sale_key'] in self.PERSON_MERGE else r['sale_name'],
            axis=1)
        df = df.groupby(['_group', 'year', 'month'], as_index=False)['total_sales'].sum()
        df.rename(columns={'_group': 'sale_name'}, inplace=True)

        sales_list = sorted(df['sale_name'].unique())
        THAI_MONTHS_SHORT = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                             "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
        QTR_NAME = {1:'Q1',2:'Q1',3:'Q1',4:'Q2',5:'Q2',6:'Q2',
                    7:'Q3',8:'Q3',9:'Q3',10:'Q4',11:'Q4',12:'Q4'}
        PERSON_GAP = 1.2

        bar_xs, bar_vals, bar_colors, bar_labels = [], [], [], []
        person_spans, qtr_spans = [], []
        cur_x = 0

        for i, sale in enumerate(sales_list):
            p_df = df[df['sale_name'] == sale]
            p_months = sorted(p_df[['year', 'month']].drop_duplicates().values.tolist())
            if not p_months:
                continue
            color = next((v for k, v in self.PERSON_COLOR_MAP.items() if k in sale),
                         self.COLORS_MONTHLY_FALLBACK[i % len(self.COLORS_MONTHLY_FALLBACK)])
            p_start = cur_x; prev_qtr = None; q_start = cur_x
            for ym in p_months:
                yr, mo = int(ym[0]), int(ym[1])
                row = p_df[(p_df['year'] == yr) & (p_df['month'] == mo)]
                val = float(row['total_sales'].sum()) if not row.empty else 0.0
                bar_xs.append(cur_x); bar_vals.append(val)
                bar_colors.append(color); bar_labels.append(THAI_MONTHS_SHORT[mo - 1])
                qtr = QTR_NAME[mo]
                if prev_qtr is not None and qtr != prev_qtr:
                    qtr_spans.append((prev_qtr, q_start, cur_x - 1)); q_start = cur_x
                prev_qtr = qtr; cur_x += 1
            if prev_qtr is not None:
                qtr_spans.append((prev_qtr, q_start, cur_x - 1))
            person_spans.append((sale, color, p_start, cur_x - 1))
            cur_x += PERSON_GAP

        if not bar_xs:
            CTkLabel(self.chart_frame, text="ไม่พบข้อมูล", font=self.header_font).pack(expand=True)
            return

        dpi = 100
        BOTTOM_MARGIN = 0.28
        fig, ax = plt.subplots(figsize=(max(_fw - 22, 400) / dpi, max(_fh - 22, 300) / dpi), dpi=dpi)
        fig.patch.set_facecolor('#F8FAFC')
        ax.set_facecolor('#F8FAFC')

        max_val = max(bar_vals) if bar_vals else 1
        ax.bar(bar_xs, bar_vals, width=0.75, color=bar_colors, zorder=3,
               edgecolor='white', linewidth=0.4)
        for bx, bv in zip(bar_xs, bar_vals):
            if bv > 0:
                ax.text(bx, bv + max_val * 0.012, f"{bv/1e6:.2f}M",
                        ha='center', va='bottom', fontsize=7, color='#1E293B', fontweight='bold')

        ax.set_xticks(bar_xs)
        ax.set_xticklabels(bar_labels, fontsize=7.5, color='#475569')
        ax.tick_params(axis='x', length=0, pad=2)

        trans = blended_transform_factory(ax.transData, ax.transAxes)
        for idx, (qtr_lbl, qs, qe) in enumerate(qtr_spans):
            mid = (qs + qe) / 2
            ax.text(mid, -0.10, qtr_lbl, transform=trans,
                    ha='center', va='top', fontsize=8, fontweight='bold', color='#334155', clip_on=False)
            ax.plot([qs - 0.4, qe + 0.4], [-0.07, -0.07], transform=trans,
                    color='#94A3B8', lw=0.8, clip_on=False, solid_capstyle='butt')
            if idx > 0:
                prev_end = qtr_spans[idx - 1][2]
                if qs - prev_end < PERSON_GAP:
                    ax.axvline((prev_end + qs) / 2, color='#94A3B8', lw=1.0, ls='--', zorder=1, alpha=0.5)

        for sale, color, ps, pe in person_spans:
            ax.text((ps + pe) / 2, -0.22, sale, transform=trans,
                    ha='center', va='top', fontsize=9, fontweight='bold', color=color, clip_on=False)
            ax.axvline(pe + PERSON_GAP / 2, color='#CBD5E1', lw=1.0, ls='--', zorder=1)

        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"{v/1e6:.1f}M" if v >= 1e6 else f"{v/1e3:.0f}K"))
        ax.set_ylabel("ยอดขายสุทธิ (บาท)", fontsize=9, color='#475569')
        ax.tick_params(axis='y', colors='#475569', labelsize=8)
        ax.spines[['top', 'right']].set_visible(False)
        ax.spines[['left', 'bottom']].set_color('#CBD5E1')
        ax.yaxis.grid(True, color='#E2E8F0', zorder=0)
        ax.set_axisbelow(True)
        ax.set_xlim(min(bar_xs) - 0.6, max(bar_xs) + 0.6)
        ax.set_ylim(0, max_val * 1.18)

        ytd_total = df['total_sales'].sum()
        fig.text(0.5, 0.97, f"{ytd_total:,.2f}",
                 ha='center', va='top', fontsize=20, fontweight='bold', color='#EF4444')
        fig.text(0.5, 0.925, "Sale Rev. YTD  —  ยอดขายสุทธิ แยกตามเดือน (จัดกลุ่มตามพนักงาน)",
                 ha='center', va='top', fontsize=9.5, color='#64748B')
        fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=BOTTOM_MARGIN)

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        widget = canvas.get_tk_widget()
        widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        _resizing = [False]
        def _on_widget_configure(event, _fig=fig, _canvas=canvas, _dpi=dpi, _bm=BOTTOM_MARGIN):
            if _resizing[0] or event.width < 100 or event.height < 100:
                return
            _resizing[0] = True
            try:
                _fig.set_size_inches(event.width / _dpi, event.height / _dpi)
                _fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=_bm)
                _canvas.draw()
            finally:
                _resizing[0] = False

        widget.bind('<Configure>', _on_widget_configure)
        canvas.draw()
        self._chart_canvas = canvas
        plt.close(fig)

    # ── Table ──────────────────────────────────────────────────────────────────
    def _draw_table(self, data_df):
        for w in self.chart_frame.winfo_children():
            w.destroy()

        if data_df.empty:
            CTkLabel(self.chart_frame, text="ไม่พบข้อมูล",
                     font=self.header_font).pack(expand=True)
            return

        people_data = self._preprocess(data_df)
        MARGIN_TARGET = 15.0  # % เป้า margin ตายตัวทุกคน
        rows = []
        for p in people_data:
            t, s = p['target'], p['total_sales']
            rows.append({
                'name': p['name'],
                'target': t,
                'sales': s,
                'pct': (s / t * 100) if t > 0 else 0.0,
                'diff': s - t,
                'margin_target': MARGIN_TARGET,
                'actual_margin': p.get('avg_margin', 0.0),
                'sales_normal': p.get('sales_normal', 0.0),
                'sales_below':  p.get('sales_below', 0.0),
            })
        rows.sort(key=lambda r: r['sales'], reverse=True)

        total_t      = sum(r['target']       for r in rows)
        total_s      = sum(r['sales']        for r in rows)
        total_normal = sum(r['sales_normal'] for r in rows)
        total_below  = sum(r['sales_below']  for r in rows)
        total_pct  = (total_s / total_t * 100) if total_t > 0 else 0.0
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

        cols = ('name', 'margin_target', 'actual_margin', 'target',
                'sales', 'sales_normal', 'sales_below', 'pct', 'diff')
        tree = ttk.Treeview(outer, columns=cols, show='headings',
                            style="MgmtTable.Treeview", height=len(rows) + 1)
        tree.heading('name',          text='พนักงาน')
        tree.heading('margin_target', text='Margin เป้า')
        tree.heading('actual_margin', text='Avg Margin จริง')
        tree.heading('target',        text='เป้าหมาย (บาท)')
        tree.heading('sales',         text='ยอดขายจริง (บาท)')
        tree.heading('sales_normal',  text='ยอดขาย Normal')
        tree.heading('sales_below',   text='ยอดขาย Below T')
        tree.heading('pct',           text='%')
        tree.heading('diff',          text='ส่วนต่าง (บาท)')
        tree.column('name',          width=170, anchor='w')
        tree.column('margin_target', width=100, anchor='center')
        tree.column('actual_margin', width=120, anchor='center')
        tree.column('target',        width=150, anchor='e')
        tree.column('sales',         width=150, anchor='e')
        tree.column('sales_normal',  width=140, anchor='e')
        tree.column('sales_below',   width=140, anchor='e')
        tree.column('pct',           width=75,  anchor='center')
        tree.column('diff',          width=140, anchor='e')

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
            actual_m = r.get('actual_margin', 0.0)
            actual_m_str = f"{actual_m:.2f}%" if actual_m else '—'
            tree.insert('', 'end', tags=(tag_row, tag_diff), values=(
                r['name'],
                f"{r['margin_target']:.0f}%",
                actual_m_str,
                f"{r['target']:,.0f}" if r['target'] > 0 else '—',
                f"{r['sales']:,.0f}",
                f"{r['sales_normal']:,.0f}" if r['sales_normal'] > 0 else '—',
                f"{r['sales_below']:,.0f}"  if r['sales_below']  > 0 else '—',
                f"{r['pct']:.1f}%",
                diff_str if r['target'] > 0 else '—',
            ))

        total_diff_str = f"+{total_diff:,.0f}" if total_diff >= 0 else f"{total_diff:,.0f}"
        # avg margin รวมทีม = weighted avg
        total_wtd_margin = sum(r['actual_margin'] * r['sales'] for r in rows)
        team_avg_margin = (total_wtd_margin / total_s) if total_s > 0 else 0.0
        team_margin_str = f"{team_avg_margin:.2f}%" if team_avg_margin else '—'
        tree.insert('', 'end', tags=('total',), values=(
            'รวมทีม',
            '15%',
            team_margin_str,
            f"{total_t:,.0f}",
            f"{total_s:,.0f}",
            f"{total_normal:,.0f}" if total_normal > 0 else '—',
            f"{total_below:,.0f}"  if total_below  > 0 else '—',
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

