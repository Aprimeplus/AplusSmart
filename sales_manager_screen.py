import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                               CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                               CTkOptionMenu, CTkRadioButton, CTkTabview, CTkCheckBox)
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import psycopg2.errors
import psycopg2.extras
import traceback
import utils
import os
import numpy as np

# เพิ่ม matplotlib สำหรับกราฟ
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import FuncFormatter
import matplotlib
import matplotlib.patheffects as pe
matplotlib.use('TkAgg')

# --- นำเข้า Class ที่จำเป็น ---
from history_windows import SOPopupWindow, DeferralHistoryWindow, ManagerDeferApprovalDialog
from daily_report_widget import DailyReportWidget
from hr_screen import SalesFilterDialog

STATUS_THAI_MAP = {
    'Draft': 'ฉบับร่าง',
    'Edited': 'แก้ไข/บันทึกร่าง',
    'Pending Sale Manager Approval': 'รอ ผจก.ฝ่ายขายอนุมัติ',
    'Rejected by SM': 'ผจก.ขาย ตีกลับ',
    'Pending PU': 'รอฝ่ายจัดซื้อรับงาน',
    'PO In Progress': 'จัดซื้อกำลังดำเนินการ',
    'Pending Approval': 'รออนุมัติ PO',
    'Approved': 'อนุมัติแล้ว',
    'Rejected': 'ถูกตีกลับให้แก้ไข',
    'PO Sent': 'สั่งซื้อ/เปิด PO เรียบร้อย',
    'Forwarded_To_HR': 'ส่งต่อให้ HR',
    'HR Verified': 'HR ตรวจสอบแล้ว',
    'Paid': 'จ่ายค่าคอมฯ แล้ว',
    'Defer Requested': 'HR ขอเลื่อนจ่าย',
    'Deferred': 'ถูกเลื่อนการจ่าย',
    'Cancelled': 'ยกเลิก',
    'Cancelled by PU': 'ยกเลิกโดยจัดซื้อ'
}

# =============================================================================
#  SALES TARGET WIDGET  —  standalone widget แสดง ยอดขาย vs เป้าหมาย
# =============================================================================
class SalesTargetWidget(CTkFrame):
    THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    THAI_MONTH_MAP = {m: i+1 for i, m in enumerate(THAI_MONTHS)}
    EXCLUDE_KEYS   = {'s','d','p','mp','ms','hr','sm','Sale Center','Pimhathai'}
    PERSON_MERGE   = {
        'VOW-P': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ภาณุพงศ์'),
        'VOW-S': ('ภาณุพงศ์ / ฐรินทร์ญา', 'ฐรินทร์ญา'),
    }

    def __init__(self, master, app_container, **kwargs):
        super().__init__(master, **kwargs)
        self.app_container = app_container
        self.pg_engine     = app_container.pg_engine
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # state
        self.sales_view_mode       = 'chart'
        self.selected_sales_filter = None
        self.custom_target_start   = None
        self.custom_target_end     = None
        self.sales_target_chart_canvas = None

        # fonts
        self._font_bold = CTkFont(size=13, weight="bold")
        self._font_hdr  = CTkFont(size=15, weight="bold")

        self._build_ui()
        self.after(300, self._update_dashboard)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        today = datetime.now()
        current_year = today.year
        years = [str(y + 543) for y in range(current_year - 2, current_year + 3)]

        # Filter bar
        fbar = CTkFrame(self, fg_color="transparent")
        fbar.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self._filter_btn = CTkButton(fbar, text="👤 กรองพนักงาน (ทั้งหมด)",
                                     fg_color="#6366F1",
                                     command=self._open_filter_dialog)
        self._filter_btn.pack(side="left", padx=(5, 15))

        def _mk_picker(label):
            f = CTkFrame(fbar, fg_color="transparent")
            f.pack(side="left", padx=8)
            CTkLabel(f, text=label, font=self._font_bold).pack(side="left", padx=(0, 5))
            m_idx = today.month - 1
            mv = tk.StringVar(value=self.THAI_MONTHS[m_idx])
            CTkOptionMenu(f, variable=mv, values=self.THAI_MONTHS, width=110).pack(side="left", padx=2)
            yv = tk.StringVar(value=str(current_year + 543))
            CTkOptionMenu(f, variable=yv, values=years, width=80).pack(side="left", padx=2)
            return mv, yv

        self.start_m_var, self.start_y_var = _mk_picker("จากรอบ:")
        self.end_m_var,   self.end_y_var   = _mk_picker("ถึงรอบ:")

        CTkButton(fbar, text="🔍 ค้นหา", width=100, fg_color="#2563EB",
                  command=self._on_search).pack(side="left", padx=20)

        # Toggle กราฟ/ตาราง
        tgl = CTkFrame(fbar, fg_color="transparent")
        tgl.pack(side="right", padx=(0, 5))

        def _set_view(mode):
            self.sales_view_mode = mode
            btn_chart.configure(fg_color="#2563EB" if mode == 'chart' else "#E2E8F0",
                                text_color="white"   if mode == 'chart' else "#475569")
            btn_table.configure(fg_color="#2563EB" if mode == 'table' else "#E2E8F0",
                                text_color="white"   if mode == 'table' else "#475569")
            self._update_dashboard()

        btn_chart = CTkButton(tgl, text="📊 กราฟ",  width=90,
                              fg_color="#2563EB", text_color="white",
                              corner_radius=6, command=lambda: _set_view('chart'))
        btn_chart.pack(side="left", padx=2)
        btn_table = CTkButton(tgl, text="📋 ตาราง", width=90,
                              fg_color="#E2E8F0", text_color="#475569",
                              corner_radius=6, command=lambda: _set_view('table'))
        btn_table.pack(side="left", padx=2)

        # Chart area
        self._chart_frame = CTkFrame(self, border_width=1, corner_radius=10)
        self._chart_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _show_loading(self):
        for w in self._chart_frame.winfo_children():
            w.destroy()
        lbl = CTkLabel(self._chart_frame, text="กำลังโหลดข้อมูล...",
                       font=CTkFont(size=18, slant="italic"), text_color="gray50")
        lbl.pack(expand=True, pady=20)
        self.update_idletasks()
        return lbl

    # ── Filter dialog ─────────────────────────────────────────────────────────
    def _open_filter_dialog(self):
        try:
            df = pd.read_sql(
                "SELECT sale_key, sale_name FROM sales_users WHERE status='Active' AND role='Sale' ORDER BY sale_name",
                self.pg_engine
            )
            sales_list = list(zip(df['sale_key'], df['sale_name']))
        except Exception:
            sales_list = []
        SalesFilterDialog(self, sales_list, self.selected_sales_filter, self._on_filter_confirmed)

    def _on_filter_confirmed(self, selected_keys):
        self.selected_sales_filter = selected_keys if selected_keys else None
        try:
            total = len(pd.read_sql(
                "SELECT sale_key FROM sales_users WHERE status='Active' AND role='Sale'",
                self.pg_engine
            ))
        except Exception:
            total = 0
        count = len(selected_keys) if selected_keys else total
        if not selected_keys or count >= total:
            self._filter_btn.configure(text="👤 กรองพนักงาน (ทั้งหมด)")
            self.selected_sales_filter = None
        else:
            self._filter_btn.configure(text=f"👤 กรองพนักงาน ({count} คน)")
        self._on_search()

    # ── Search / Update ───────────────────────────────────────────────────────
    def _on_search(self):
        import calendar
        try:
            def _get_my(mv, yv):
                m = self.THAI_MONTHS.index(mv.get()) + 1
                y = int(yv.get()) - 543
                return m, y
            s_m, s_y = _get_my(self.start_m_var, self.start_y_var)
            e_m, e_y = _get_my(self.end_m_var,   self.end_y_var)
            start_date = datetime(s_y, s_m, 1)
            last_day   = calendar.monthrange(e_y, e_m)[1]
            end_date   = datetime(e_y, e_m, last_day)
            if start_date > end_date:
                messagebox.showerror("รอบเดือนไม่ถูกต้อง",
                                     "รอบเริ่มต้นต้องมาก่อนหรือเท่ากับรอบสิ้นสุด", parent=self)
                return
            self.custom_target_start = start_date
            self.custom_target_end   = end_date
            self._update_dashboard()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            traceback.print_exc()

    def _update_dashboard(self):
        loading = self._show_loading()
        try:
            df = self._get_data()
            loading.destroy()
            if self.sales_view_mode == 'table':
                self._create_table(self._chart_frame, df)
            else:
                self._create_chart(self._chart_frame, df)
        except Exception as e:
            loading.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()

    # ── Data fetch ────────────────────────────────────────────────────────────
    def _get_data(self):
        today = datetime.now()
        params = []
        date_clauses = []
        target_mult  = 1.0

        s = self.custom_target_start
        e = self.custom_target_end
        if s and e:
            date_clauses.append(
                "MAKE_DATE(c.commission_year, c.commission_month, 1) BETWEEN %s::date AND %s::date"
            )
            params.extend([s.strftime("%Y-%m-%d"), e.strftime("%Y-%m-%d")])
            months_diff   = (e.year - s.year) * 12 + (e.month - s.month) + 1
            target_mult   = max(1, months_diff)
        else:
            date_clauses.append("c.commission_month = %s")
            params.append(today.month)
            date_clauses.append("c.commission_year = %s")
            params.append(today.year)

        date_sql = " AND ".join(date_clauses)

        sale_filter = ""
        if self.selected_sales_filter:
            ph = ','.join(["REPLACE(LOWER(%s), ' ', '')"] * len(self.selected_sales_filter))
            sale_filter = f" AND REPLACE(LOWER(su.sale_key), ' ', '') IN ({ph})"
            params.extend(self.selected_sales_filter)

        query = f"""
            SELECT su.sale_name, su.sale_key,
                   COALESCE(su.sales_target, 0) * %s AS sales_target,
                   COALESCE(SUM(c.total_sales), 0)   AS total_sales,
                   0                                  AS total_outstanding
            FROM   sales_users su
            LEFT JOIN commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND {date_sql}
            WHERE  su.status = 'Active'
                   {sale_filter}
            GROUP  BY su.sale_name, su.sale_key, su.sales_target, su.role
            HAVING (su.role = 'Sale' OR COALESCE(SUM(c.total_sales), 0) > 0)
            ORDER  BY su.sale_name ASC;
        """
        final_params = tuple([target_mult] + params)
        df = pd.read_sql_query(query, self.pg_engine, params=final_params)
        df['sales_target']      = df['sales_target'].fillna(0)
        df['total_sales']       = df['total_sales'].fillna(0)
        df['total_outstanding'] = df['total_outstanding'].fillna(0)
        return df

    # ── Chart ─────────────────────────────────────────────────────────────────
    def _create_chart(self, parent_frame, data_df):
        if self.sales_target_chart_canvas:
            try: self.sales_target_chart_canvas.get_tk_widget().destroy()
            except Exception: pass
        for w in parent_frame.winfo_children():
            w.destroy()

        EXCLUDE = self.EXCLUDE_KEYS
        MERGE   = self.PERSON_MERGE

        df2 = data_df[~data_df['sale_key'].isin(EXCLUDE)].copy()
        if df2.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูลพนักงานขาย",
                     font=self._font_hdr).pack(expand=True)
            return

        df2['_group']     = df2.apply(lambda r: MERGE[r['sale_key']][0] if r['sale_key'] in MERGE else r['sale_name'], axis=1)
        df2['_seg_label'] = df2.apply(lambda r: MERGE[r['sale_key']][1] if r['sale_key'] in MERGE else r['sale_name'], axis=1)

        people_data = []
        for name, grp in df2.groupby('_group', sort=False):
            subs = [{'sale_key': r['sale_key'], 'label': r['_seg_label'],
                     'sales': float(r['total_sales'])} for _, r in grp.iterrows()]
            subs.sort(key=lambda s: s['sales'], reverse=True)
            people_data.append({
                'name':              name,
                'total_sales':       float(grp['total_sales'].sum()),
                'total_outstanding': float(grp['total_outstanding'].sum()),
                'target':            float(grp['sales_target'].sum()),
                'sub_items':         subs,
            })
        people_data.sort(key=lambda p: p['total_sales'], reverse=True)
        people_data = [p for p in people_data if p['total_sales'] > 0 or p['target'] > 0]

        n       = len(people_data)
        names   = [p['name']        for p in people_data]
        sales   = [p['total_sales'] for p in people_data]
        targets = [p['target']      for p in people_data]

        ACHIEVE_COLORS = {
            'green':  ('#22C55E', '#86EFAC'),
            'yellow': ('#F59E0B', '#FCD34D'),
            'red':    ('#EF4444', '#FCA5A5'),
            'gray':   ('#94A3B8', '#CBD5E1'),
        }
        def achievement_key(s, t):
            if t <= 0: return 'gray'
            return 'green' if s/t >= 1.0 else ('yellow' if s/t >= 0.7 else 'red')

        pct_labels = [f"{s/t*100:.0f}%" if t > 0 else "N/A" for s, t in zip(sales, targets)]

        BG = '#F8FAFC'; GRID_C = '#E2E8F0'
        chart_width = max(10, n * 1.6)
        fig = Figure(figsize=(chart_width, 7.2), dpi=100, facecolor=BG)
        ax  = fig.add_subplot(111)
        ax.set_facecolor(BG)

        x     = np.arange(n)
        width = 0.65
        max_t = max(targets or [1])

        # แท่งพื้นหลัง (ส่วนที่ยังไม่ถึงเป้า)
        for i, p in enumerate(people_data):
            gap = max(0.0, p['target'] - p['total_sales'])
            if gap > 0:
                ax.bar(x[i], gap, width, bottom=p['total_sales'],
                       color='#E2E8F0', zorder=2, linewidth=0)

        # แท่งยอดขาย (stacked)
        for i, p in enumerate(people_data):
            total_s = p['total_sales']; target = p['target']
            akey = achievement_key(total_s, target)
            colors = ACHIEVE_COLORS[akey]
            subs   = p['sub_items']
            multi_id = len([s for s in subs if s['sales'] > 0]) > 1
            bottom = 0.0
            for idx, sub in enumerate(subs):
                seg_h = sub['sales']
                if seg_h <= 0: continue
                color = colors[min(idx, 1)]
                is_partner = (idx > 0)
                ax.bar(x[i], seg_h, width, bottom=bottom, color=color, zorder=3,
                       linewidth=0.8 if is_partner else 0,
                       edgecolor='white' if is_partner else 'none',
                       hatch='//' if is_partner else None, alpha=0.92)
                mid_y = bottom + seg_h / 2
                seg_top = bottom + seg_h; t_line = p['target']
                if t_line > 0 and bottom < t_line < seg_top:
                    mid_y = (bottom + (t_line - bottom)/2) if (t_line - bottom) >= seg_h*0.35 else (t_line + (seg_top - t_line)/2)
                txt_color   = '#1a5c1a' if is_partner else 'white'
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

        # เส้น target
        half = width / 2
        for i, t in enumerate(targets):
            if t > 0:
                ax.hlines(t, x[i]-half, x[i]+half, colors='#6366F1', linewidths=2.2, linestyles='--', zorder=5)
                ax.text(x[i], t + max_t*0.012, f"Target  {t:,.0f}",
                        ha='center', va='bottom', fontsize=9.5, weight='bold', color='#6366F1', zorder=6)

        # % badge
        for i, (p, pct) in enumerate(zip(people_data, pct_labels)):
            s = p['total_sales']; t = p['target']
            if s == 0 and t > 0:
                ax.text(x[i], t*0.5, "ยังไม่มี\nข้อมูล SO",
                        ha='center', va='center', fontsize=11, weight='bold',
                        color='#94A3B8', zorder=8, style='italic', linespacing=1.4)
            else:
                pct_y = max(s, t) + max_t * 0.16
                ax.text(x[i], pct_y, pct, ha='center', va='bottom', fontsize=16,
                        weight='medium', color='black', zorder=8,
                        path_effects=[pe.withStroke(linewidth=2, foreground='white')])
                if s <= max_t * 0.08:
                    ax.text(x[i], s + max_t*0.012, f"{s:,.0f}",
                            ha='center', va='bottom', fontsize=12, weight='bold', color='black', zorder=7)

        # Axes
        max_sales = max((p['total_sales'] for p in people_data), default=0)
        ax.set_ylim(0, max(max_sales, max_t) + max_t * 0.55)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:,.0f}'))
        ax.tick_params(axis='y', labelsize=13); ax.tick_params(axis='x', pad=8)

        def _fmt_name(nm, tgt):
            short = nm.replace(' / ', '\n').replace(' ', '\n', 1)
            return f"{short}\nเป้า {tgt:,.0f}" if tgt > 0 else short

        ax.set_xticks(x)
        ax.set_xticklabels([_fmt_name(nm, tgt) for nm, tgt in zip(names, targets)],
                           rotation=0, ha='center', fontsize=10, weight='medium',
                           color='black', linespacing=1.35)
        ax.set_ylabel('จำนวนเงิน (บาท)', fontsize=12, weight='medium', color='black', labelpad=10)

        team_sales  = sum(p['total_sales'] for p in people_data)
        team_target = sum(p['target']      for p in people_data)
        team_pct    = (team_sales / team_target * 100) if team_target > 0 else 0
        n_hit  = sum(1 for p in people_data if p['target'] > 0 and p['total_sales'] >= p['target'])
        n_miss = sum(1 for p in people_data if p['target'] > 0 and p['total_sales'] < p['target'])
        summary = (f"ทีมรวม  {team_sales:,.0f} / {team_target:,.0f} บาท  ({team_pct:.0f}%)     "
                   f"ถึงเป้า {n_hit} คน  ·  ยังไม่ถึง {n_miss} คน")
        ax.set_title('ยอดขาย vs เป้าหมาย  (เฉพาะ SO ที่คิดค่าคอมแล้ว)',
                     fontsize=16, weight='semibold', color='#0F172A', loc='left', pad=36)
        ax.text(0, 1.015, summary, transform=ax.transAxes,
                fontsize=10.5, color='#64748B', ha='left', va='bottom', weight='medium')

        ax.yaxis.grid(True, color=GRID_C, linewidth=0.8, zorder=0)
        ax.set_axisbelow(True)
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(GRID_C); ax.spines['bottom'].set_color(GRID_C)
        ax.tick_params(colors='black')

        from matplotlib.patches import Patch
        from matplotlib.lines   import Line2D
        has_multi = any(len(p['sub_items']) > 1 for p in people_data)
        legend_items = [
            Patch(facecolor='#22C55E', label='≥ 100% เป้า'),
            Patch(facecolor='#F59E0B', label='70–99% เป้า'),
            Patch(facecolor='#EF4444', label='< 70% เป้า'),
            Line2D([0],[0], color='#6366F1', lw=2, linestyle='--', label='เป้าหมาย'),
            Patch(facecolor='#E2E8F0', edgecolor='#CBD5E1', label='ส่วนที่ยังไม่ถึงเป้า'),
        ]
        if has_multi:
            legend_items.append(Patch(facecolor='#86EFAC', label='ยอดขายของพาร์ทเนอร์'))
        ax.legend(handles=legend_items, loc='upper right', bbox_to_anchor=(1.0, 1.0),
                  ncol=2, frameon=True, framealpha=0.95, edgecolor='#CBD5E1', fontsize=10,
                  prop={'weight': 'bold', 'size': 10}, borderpad=0.7, labelspacing=0.4, columnspacing=1.0)

        fig.tight_layout(rect=[0, 0.05, 1, 1])
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.sales_target_chart_canvas = canvas

    # ── Table ─────────────────────────────────────────────────────────────────
    def _create_table(self, parent_frame, data_df):
        for w in parent_frame.winfo_children():
            w.destroy()
        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูล", font=self._font_hdr).pack(expand=True)
            return

        EXCLUDE = self.EXCLUDE_KEYS; MERGE = self.PERSON_MERGE
        df2 = data_df[~data_df['sale_key'].isin(EXCLUDE)].copy()
        df2['_group'] = df2.apply(
            lambda r: MERGE[r['sale_key']][0] if r['sale_key'] in MERGE else r['sale_name'], axis=1)
        rows = []
        for name, grp in df2.groupby('_group', sort=False):
            t = float(grp['sales_target'].sum()); s = float(grp['total_sales'].sum())
            pct = (s/t*100) if t > 0 else 0.0
            rows.append({'name': name, 'target': t, 'sales': s, 'pct': pct, 'diff': s-t})
        rows.sort(key=lambda r: r['sales'], reverse=True)

        total_t = sum(r['target'] for r in rows); total_s = sum(r['sales'] for r in rows)
        total_pct = (total_s/total_t*100) if total_t > 0 else 0.0

        # Period label
        try:
            s_m = self.THAI_MONTHS.index(self.start_m_var.get()) + 1
            e_m = self.THAI_MONTHS.index(self.end_m_var.get())   + 1
            s_y = int(self.start_y_var.get()); e_y = int(self.end_y_var.get())
            period_label = f"{self.start_m_var.get()} {s_y}" if (s_m == e_m and s_y == e_y) else \
                           f"{self.start_m_var.get()} {s_y} – {self.end_m_var.get()} {e_y}"
        except Exception:
            period_label = "รอบที่เลือก"

        outer = CTkScrollableFrame(parent_frame, fg_color="white", corner_radius=10)
        outer.pack(fill="both", expand=True, padx=10, pady=10)
        CTkLabel(outer, text=f"สรุปยอดขาย vs เป้าหมาย  —  {period_label}",
                 font=CTkFont(size=15, weight="bold"), text_color="#0F172A"
                 ).pack(anchor="w", padx=16, pady=(12, 8))

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("SMSalesTable.Treeview", background="white", foreground="#1E293B",
                        rowheight=34, fieldbackground="white", font=('TH Sarabun New', 13))
        style.configure("SMSalesTable.Treeview.Heading", background="#D1FAE5", foreground="#065F46",
                        font=('TH Sarabun New', 13, 'bold'), relief="flat")
        style.map("SMSalesTable.Treeview", background=[('selected', '#EDE9FE')],
                  foreground=[('selected', '#1E293B')])

        cols = ('name','target','sales','pct','diff')
        tree = ttk.Treeview(outer, columns=cols, show='headings',
                            style="SMSalesTable.Treeview", height=len(rows)+1)
        tree.heading('name',   text='พนักงาน');       tree.column('name',   width=200, anchor='w')
        tree.heading('target', text='เป้าหมาย (บาท)'); tree.column('target', width=170, anchor='e')
        tree.heading('sales',  text='ยอดขายจริง (บาท)'); tree.column('sales', width=170, anchor='e')
        tree.heading('pct',    text='%');              tree.column('pct',    width=90,  anchor='center')
        tree.heading('diff',   text='ส่วนต่าง (บาท)'); tree.column('diff',   width=170, anchor='e')

        tree.tag_configure('odd',      background='#F8FAFC')
        tree.tag_configure('even',     background='white')
        tree.tag_configure('negative', foreground='#DC2626')
        tree.tag_configure('positive', foreground='#16A34A')
        tree.tag_configure('total',    background='#EFF6FF',
                           font=('TH Sarabun New', 13, 'bold'), foreground='#1D4ED8')

        for idx, r in enumerate(rows):
            diff_str = f"+{r['diff']:,.0f}" if r['diff'] >= 0 else f"{r['diff']:,.0f}"
            tree.insert('', 'end', tags=('odd' if idx % 2 else 'even',
                                         'positive' if r['diff'] >= 0 else 'negative'),
                        values=(r['name'],
                                f"{r['target']:,.0f}" if r['target'] > 0 else '—',
                                f"{r['sales']:,.0f}",
                                f"{r['pct']:.1f}%",
                                diff_str if r['target'] > 0 else '—'))

        total_diff = total_s - total_t
        tree.insert('', 'end', tags=('total',),
                    values=('รวมทีม', f"{total_t:,.0f}", f"{total_s:,.0f}",
                            f"{total_pct:.1f}%",
                            f"+{total_diff:,.0f}" if total_diff >= 0 else f"{total_diff:,.0f}"))
        tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))


class ToastNotification(CTkToplevel):
    """หน้าต่างแจ้งเตือนมุมขวาล่าง เด้งโชว์แล้วหายไปเอง (ไม่บล็อกการทำงาน)"""
    def __init__(self, master, title, message, duration=10000, color="#16A34A"):
        super().__init__(master)
        
        # ลบกรอบหน้าต่างออก (ไร้ขอบ)
        self.overrideredirect(True)
        self.attributes("-topmost", True) # ให้อยู่บนสุดเสมอ
        
        # ตั้งค่าสีและ Layout
        self.configure(fg_color=color)
        
        # จัดข้อความ
        CTkLabel(self, text=title, font=CTkFont(size=14, weight="bold"), text_color="white").pack(padx=20, pady=(10, 2), anchor="w")
        CTkLabel(self, text=message, font=CTkFont(size=12), text_color="white", justify="left").pack(padx=20, pady=(0, 10), anchor="w")

        self.update_idletasks()
        width = self.winfo_reqwidth()
        height = self.winfo_reqheight()

        # คำนวณตำแหน่งมุมขวาล่างของหน้าจอ
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # วางที่มุมขวาล่าง (ห่างขอบนิดหน่อย)
        x = screen_width - width - 30
        y = screen_height - height - 60 
        
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # ทำให้คลิกที่แจ้งเตือนแล้วหายไปทันที (เผื่อไม่อยากรอ 10 วิ)
        self.bind("<Button-1>", lambda e: self.destroy())
        for child in self.winfo_children():
            child.bind("<Button-1>", lambda e: self.destroy())

        # ตั้งเวลาทำลายตัวเอง (10 วินาที = 10000 ms)
        self.after(duration, self.destroy)

class SalesManagerScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role
        
        self.label_font = CTkFont(size=14, weight="bold")
        self.entry_font = CTkFont(size=14)

        self.label_font_bold = CTkFont(size=14, weight="bold")
        self.header_font_table = CTkFont(size=16, weight="bold")
        
        # --- เตรียมตัวแปรสำหรับ SOPopupWindow ---
        self.so_popup = None
        self._so_create_string_vars()
        self.sale_theme = self.app_container.THEME.get("sale", {"bg": "white", "primary": "#3B82F6"})
        
        self.pg_engine = self.app_container.pg_engine
        
        # ตัวแปรสำหรับ Filter กราฟ
        self._init_chart_filters()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 

        # --- 1. Header ---
        self._create_header()

        # --- 2. TabView ---
        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # ✅ เพิ่มแท็บ 1: รายการรออนุมัติ (เป็นหน้าแรก)
        self.approval_tab = self.tab_view.add("🗳️ รายการรออนุมัติ (SM Approval)")
        self.defer_approval_tab = self.tab_view.add("⏳ คำขอเลื่อนรอบคอมฯ (Defer)")
        # แท็บเดิม
        self.daily_report_tab = self.tab_view.add("📅 รายงานประจำวัน (SO Report)")
        self.master_tab = self.tab_view.add("🛠️ ค้นหาและจัดการ (Master)")
        self.cancelled_tab = self.tab_view.add("❌ ยกเลิก SO (Cancel)")
        self.target_tab = self.tab_view.add("📊 เป้าการขาย")

        # สร้างเนื้อหาในแต่ละ Tab
        self._create_approval_tab(self.approval_tab)
        self._create_defer_approval_tab(self.defer_approval_tab)
        self._create_daily_report_widget(self.daily_report_tab)
        self._create_master_tab(self.master_tab)
        self._create_cancelled_so_tab(self.cancelled_tab)
        self._create_sales_target_tab(self.target_tab)

        # ตั้งค่าหน้าแรกที่เปิดขึ้นมา
        self.tab_view.set("🗳️ รายการรออนุมัติ (SM Approval)")

        self.known_pending_so_ids = set() # เก็บ ID ของ SO ที่เคยแจ้งเตือนไปแล้ว
        self.reminder_timer_count = 0     # ตัวนับเวลาสำหรับเตือนทุก 10 นาที
        self.noti_job_id = None
        self._sm_noti_dialog_open = False  # ป้องกัน dialog ซ้อนกัน
        
        # เริ่มทำงานระบบ Noti เบื้องหลัง (หน่วงเวลา 3 วินาทีหลังเปิดหน้าจอ)
        self.after(3000, self._start_notification_system)
        
        # ดักจับตอนปิดหน้าจอให้หยุด Noti ด้วย
        self.bind("<Destroy>", self._on_destroy)
    
    def _create_sales_target_tab(self, parent_tab):
        """Embed SalesTargetWidget ใน tab เป้าการขาย"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(0, weight=1)
        widget = SalesTargetWidget(
            master=parent_tab,
            app_container=self.app_container,
            fg_color="transparent"
        )
        widget.grid(row=0, column=0, sticky="nsew")

    def _create_cancelled_so_tab(self, parent_tab):
        # --- Grid Setup: แบ่งหน้าจอเป็น 2 ส่วน (บน-เล็ก / ล่าง-ใหญ่) ---
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(0, weight=0) # ส่วนค้นหา (ความสูงคงที่)
        parent_tab.grid_rowconfigure(1, weight=1) # ส่วนตาราง (ขยายเต็มที่)

        # =========================================================
        #  SECTION 1: Professional Action Bar (แถบเครื่องมือด้านบน)
        # =========================================================
        action_bar = CTkFrame(parent_tab, height=60, fg_color=("gray90", "gray16"), corner_radius=6)
        action_bar.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        
        action_bar.grid_columnconfigure(3, weight=1) 
        
        CTkLabel(action_bar, text="🔎 ระบุเลข SO ที่ต้องการยกเลิก:", font=self.label_font_bold).grid(row=0, column=0, padx=(20, 5), pady=10, sticky="w")
        
        self.cancel_search_entry = CTkEntry(action_bar, placeholder_text="ระบุเลข SO... (เช่น SO6701-001)", width=250, height=34)
        self.cancel_search_entry.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        self.cancel_search_entry.bind("<Return>", lambda e: self._search_so_to_cancel())
        
        CTkButton(action_bar, text="ค้นหา", command=self._search_so_to_cancel, 
                  width=100, height=34, fg_color="#3B82F6", hover_color="#2563EB", font=self.label_font_bold).grid(row=0, column=2, padx=10, pady=10, sticky="w")

        # Inline Result Area
        self.inline_result_frame = CTkFrame(action_bar, fg_color="transparent", height=34)
        self.inline_result_frame.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        CTkButton(action_bar, text="⟳ รีเฟรช", command=self._load_cancelled_so_history, 
                  width=90, height=34, fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).grid(row=0, column=4, padx=20, pady=10, sticky="e")

        # =========================================================
        #  SECTION 2: Full-Width History Table (ตารางเต็มจอ)
        # =========================================================
        table_container = CTkFrame(parent_tab, fg_color="transparent")
        table_container.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        
        header_row = CTkFrame(table_container, fg_color="transparent", height=36)
        header_row.pack(fill="x", pady=(0, 5))
        CTkLabel(header_row, text="📜 ประวัติรายการยกเลิก", font=self.header_font_table, text_color="#EF4444").pack(side="left")

        self._sm_cancel_view_mode = "normal"
        toggle_frame = CTkFrame(header_row, fg_color="transparent")
        toggle_frame.pack(side="right")
        self._sm_toggle_normal_btn = CTkButton(
            toggle_frame, text="ยกเลิก SO", width=110, height=30,
            fg_color="#EF4444", hover_color="#DC2626", text_color="white",
            font=CTkFont(size=12, weight="bold"),
            command=lambda: self._switch_sm_cancel_view("normal"))
        self._sm_toggle_normal_btn.pack(side="left", padx=(0, 4))
        self._sm_toggle_transport_btn = CTkButton(
            toggle_frame, text="🚚 SO ค่าขนส่ง", width=120, height=30,
            fg_color="transparent", border_width=1, border_color="#D97706",
            text_color="#D97706", font=CTkFont(size=12),
            command=lambda: self._switch_sm_cancel_view("transport"))
        self._sm_toggle_transport_btn.pack(side="left")

        self.cancelled_history_frame = CTkFrame(table_container, fg_color="transparent")
        self.cancelled_history_frame.pack(fill="both", expand=True)

        self.after(100, self._load_cancelled_so_history)

    def _switch_sm_cancel_view(self, mode):
        self._sm_cancel_view_mode = mode
        if mode == "normal":
            self._sm_toggle_normal_btn.configure(fg_color="#EF4444", text_color="white")
            self._sm_toggle_transport_btn.configure(fg_color="transparent", text_color="#D97706")
            self._load_cancelled_so_history()
        else:
            self._sm_toggle_transport_btn.configure(fg_color="#D97706", text_color="white")
            self._sm_toggle_normal_btn.configure(fg_color="transparent", text_color="#EF4444")
            self._load_sm_transport_so_history()

    def _load_sm_transport_so_history(self):
        """โหลดประวัติ SO ค่าขนส่งที่ SM อนุมัติยกเลิกแล้ว"""
        for widget in self.cancelled_history_frame.winfo_children():
            widget.destroy()

        # สร้าง table ถ้ายังไม่มี
        conn_chk = self.app_container.get_connection()
        try:
            with conn_chk.cursor() as _cur:
                _cur.execute("""
                    CREATE TABLE IF NOT EXISTS transport_cancel_requests (
                        id SERIAL PRIMARY KEY,
                        so_number VARCHAR(50) NOT NULL,
                        so_id INTEGER,
                        sale_key VARCHAR(50),
                        customer_name VARCHAR(200),
                        original_status VARCHAR(50),
                        requested_by VARCHAR(100),
                        requested_at TIMESTAMP DEFAULT NOW(),
                        status VARCHAR(20) DEFAULT 'pending',
                        reviewed_by VARCHAR(100),
                        reviewed_at TIMESTAMP
                    )
                """)
            conn_chk.commit()
        finally:
            self.app_container.release_connection(conn_chk)

        try:
            query = """
                SELECT requested_at, so_number, sale_key, customer_name,
                       requested_by, reviewed_by
                FROM transport_cancel_requests
                WHERE status = 'approved'
                ORDER BY requested_at DESC LIMIT 100
            """
            df = pd.read_sql_query(query, self.pg_engine)

            columns = ("วันที่", "SO Number", "รหัสพนักงาน", "ชื่อลูกค้า", "ขอโดย HR", "อนุมัติโดย SM")
            tree = ttk.Treeview(self.cancelled_history_frame, columns=columns, show="headings", height=15)

            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview.Heading", font=('Tahoma', 12, 'bold'), background="#FEF2F2")
            style.configure("Treeview", font=('Tahoma', 11), rowheight=30)

            for col in columns:
                tree.heading(col, text=col)
            tree.column("วันที่", width=150, anchor="center")
            tree.column("SO Number", width=150, anchor="center")
            tree.column("รหัสพนักงาน", width=120, anchor="center")
            tree.column("ชื่อลูกค้า", width=280, anchor="w")
            tree.column("ขอโดย HR", width=140, anchor="center")
            tree.column("อนุมัติโดย SM", width=140, anchor="center")

            vsb = ttk.Scrollbar(self.cancelled_history_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vsb.set)
            tree.pack(side="left", fill="both", expand=True)
            vsb.pack(side="right", fill="y")

            if df.empty:
                tree.insert("", "end", values=("-", "ยังไม่มีประวัติ SO ค่าขนส่ง", "-", "-", "-", "-"))
                return

            for _, row in df.iterrows():
                ts = row['requested_at']
                ts_str = str(ts)[:16] if pd.notna(ts) else "-"
                tree.insert("", "end", values=(
                    ts_str,
                    row['so_number'] or "-",
                    row['sale_key'] or "-",
                    row['customer_name'] or "-",
                    row['requested_by'] or "-",
                    row['reviewed_by'] or "-",
                ))

        except Exception as e:
            import traceback
            CTkLabel(self.cancelled_history_frame, text=f"โหลดข้อมูลล้มเหลว: {e}",
                     text_color="red").pack(pady=20)
            traceback.print_exc()

    def _start_notification_system(self):
        """ลูปตรวจสอบ Noti ทุกๆ 1 นาที"""
        self._check_pending_approvals()
        
        # ลูปเรียกตัวเองทุกๆ 60 วินาที (60,000 ms)
        self.noti_job_id = self.after(60000, self._start_notification_system)

    def _check_pending_approvals(self):
        """เช็ค DB ว่ามี SO ใหม่ หรือต้องทวงงานหรือไม่"""
        try:
            query = """
                SELECT id, so_number, customer_name 
                FROM commissions 
                WHERE status = 'Pending Sale Manager Approval' AND is_active = 1
            """
            df = pd.read_sql_query(query, self.pg_engine)
            
            current_pending_ids = set(df['id'].tolist())
            pending_count = len(current_pending_ids)
            
            # 1. เช็คว่ามี SO "ใหม่" เข้ามาหรือไม่ (ID ที่เพิ่งโผล่มาใหม่)
            new_so_ids = current_pending_ids - self.known_pending_so_ids
            
            if new_so_ids:
                # มีงานใหม่เข้า! เด้งแจ้งเตือนแบบ 10 วินาที (สีเขียว/ฟ้า)
                new_count = len(new_so_ids)
                ToastNotification(
                    self.winfo_toplevel(), 
                    title="🔔 มี SO ใหม่รออนุมัติ!", 
                    message=f"มีคำขออนุมัติ SO เข้ามาใหม่จำนวน {new_count} รายการ\nกรุณาตรวจสอบในแท็บ 'รายการรออนุมัติ'",
                    duration=10000, 
                    color="#2563EB" # สีฟ้า
                )
                # อัปเดตรายการที่รู้จักแล้ว และรีเฟรชตารางโชว์งานใหม่
                self.known_pending_so_ids = current_pending_ids
                
                # --- ย้ายบรรทัดนี้ไปไว้ใน after() เพื่อไม่ให้ขัดจังหวะ ---
                self.after(100, self._load_approval_data)
                
            # 2. ระบบทวงงาน (Reminder) ถ้ายังมีงานค้าง
            elif pending_count > 0:
                self.reminder_timer_count += 1
                
                # เช็คทุก 1 นาที ดังนั้น 10 รอบ = 10 นาที
                if self.reminder_timer_count >= 10: 
                    ToastNotification(
                        self.winfo_toplevel(), 
                        title="⚠️ แจ้งเตือนงานค้าง!", 
                        message=f"คุณยังมี SO รอการอนุมัติค้างอยู่ {pending_count} รายการ",
                        duration=10000, 
                        color="#EA580C" # สีส้ม
                    )
                    self.reminder_timer_count = 0 # รีเซ็ตตัวนับ
            else:
                # ไม่มีงานค้างเลย รีเซ็ตเวลา
                self.reminder_timer_count = 0
                self.known_pending_so_ids.clear()

        except Exception as e:
            print(f"Noti System Error: {e}")

        # --- ตรวจสอบ Notifications Table (เช่น HR ยกเลิก SO) ---
        if not self._sm_noti_dialog_open and self.user_key:
            try:
                noti_df = pd.read_sql_query(
                    "SELECT id, message, timestamp FROM notifications "
                    "WHERE user_key_to_notify = %(ukey)s AND is_read = FALSE "
                    "AND message LIKE '[HR\\_CANCEL]%%' "
                    "ORDER BY timestamp ASC LIMIT 1",
                    self.pg_engine,
                    params={"ukey": self.user_key}
                )
                print(f"[SM Noti] user_key={self.user_key}, unread_count={len(noti_df)}")
                if not noti_df.empty:
                    r = noti_df.iloc[0]
                    noti_tuple = (r['id'], r['message'], r.get('timestamp'))
                    self._sm_noti_dialog_open = True
                    dlg = SMNotificationDialog(self.winfo_toplevel(), self.app_container, noti_tuple)
                    dlg.bind("<Destroy>", lambda e: setattr(self, '_sm_noti_dialog_open', False))
            except Exception as e:
                print(f"SM Notification check error: {e}")
                traceback.print_exc()

    def _on_destroy(self, event):
        """หยุด Loop เมื่อปิดหน้าจอ"""
        if hasattr(event, 'widget') and event.widget is self:
            if self.noti_job_id:
                self.after_cancel(self.noti_job_id)

    def _init_chart_filters(self):
        """สร้างตัวแปรสำหรับ Filter กราฟ"""
        now = datetime.now()
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                           "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.chart_month_var = tk.StringVar(value=self.thai_months[now.month - 1])
        self.chart_year_var = tk.StringVar(value=str(now.year + 543))

    def _create_header(self):
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        CTkLabel(header_frame, text=f"Sale Manager Dashboard: {self.user_name}", font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        button_frame = CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side="right", padx=10)
        
        # 🟢 [เพิ่มตรงนี้] ปุ่ม Export ข้อมูลแบบมี Popup
        CTkButton(button_frame, text="📥 Export Data SO", 
                  fg_color="#10B981", hover_color="#059669", font=CTkFont(weight="bold"),
                  command=self._open_sm_export_dialog).pack(side="left", padx=(0, 10))
        
        CTkButton(button_frame, text="🔄 Refresh All", command=self._refresh_all_tabs).pack(side="left", padx=5)
        
        CTkButton(button_frame, text="ออกจากระบบ", command=self.app_container.show_login_screen, 
                  fg_color="transparent", border_color="#D32F2F", 
                  text_color="#D32F2F", border_width=2, 
                  hover_color="#FFEBEE").pack(side="left", padx=5)

    # 🟢 [เพิ่มฟังก์ชันนี้เข้าไปในคลาส SalesManagerScreen ด้วย]
    def _open_sm_export_dialog(self):
        """เรียกเปิดหน้าต่าง Popup สำหรับ Export"""
        SMExportDialog(self, self.app_container)

    def _refresh_all_tabs(self):
        """รีโหลดข้อมูลทุกแท็บ"""
        self._load_approval_data() # รีโหลดหน้าอนุมัติ
        self._update_rejection_chart() # รีโหลดกราฟ
        
        if hasattr(self, 'daily_report_widget'):
            self.daily_report_widget.load_report_data()
            if hasattr(self.daily_report_widget, 'dashboard_view'):
                 self.daily_report_widget.dashboard_view._update_chart()
        
        self._load_master_data() # รีโหลด Master Tab

    # =========================================================================
    # ✅ NEW TAB: SM APPROVAL พร้อม Dashboard กราฟ
    # =========================================================================
    def _create_approval_tab(self, parent):
        """สร้างแท็บรออนุมัติ แบ่งเป็น Dashboard 40% + List 60%"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=3)  # Dashboard ~30%
        parent.grid_rowconfigure(1, weight=7)  # List ~70%

        # ========== ส่วนบน 40%: Dashboard กราฟ ==========
        dashboard_container = CTkFrame(parent, fg_color="#F3F4F6", corner_radius=10)
        dashboard_container.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="nsew")
        dashboard_container.grid_columnconfigure(0, weight=1)
        dashboard_container.grid_rowconfigure(1, weight=1)

        # Header + Filters
        chart_header = CTkFrame(dashboard_container, fg_color="transparent")
        chart_header.grid(row=0, column=0, sticky="ew", padx=15, pady=10)
        
        CTkLabel(chart_header, text="📊 สถิติการตีกลับงานรายบุคคล", 
                font=CTkFont(size=16, weight="bold")).pack(side="left")

        filter_frame = CTkFrame(chart_header, fg_color="transparent")
        filter_frame.pack(side="right")
        
        # ✅ ปุ่ม Export Excel
        CTkButton(filter_frame, 
                 text="📥 Export Excel",
                 width=120,
                 height=32,
                 fg_color="#16A34A",
                 hover_color="#15803D",
                 font=CTkFont(size=12, weight="bold"),
                 command=self._export_rejection_to_excel).pack(side="left", padx=(0, 10))
        
        CTkLabel(filter_frame, text="รอบเดือน:", font=CTkFont(size=12)).pack(side="left", padx=5)
        CTkOptionMenu(filter_frame, variable=self.chart_month_var, values=self.thai_months, 
                      width=110, height=32,
                      command=lambda e: self._update_rejection_chart()).pack(side="left", padx=3)
        
        years = [str(y + 543) for y in range(datetime.now().year - 1, datetime.now().year + 2)]
        CTkOptionMenu(filter_frame, variable=self.chart_year_var, values=years, 
                      width=90, height=32,
                      command=lambda e: self._update_rejection_chart()).pack(side="left", padx=3)

        # พื้นที่กราฟ
        self.chart_area = CTkFrame(dashboard_container, fg_color="white", corner_radius=8)
        self.chart_area.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")

        # ========== ส่วนล่าง 60%: รายการ SO ==========
        list_container = CTkFrame(parent, fg_color="transparent")
        list_container.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew")
        list_container.grid_columnconfigure(0, weight=1)
        list_container.grid_rowconfigure(1, weight=1)

        # Search Bar
        search_frame = CTkFrame(list_container, fg_color="#F3F4F6", corner_radius=8, height=60)
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        search_frame.grid_propagate(False)
        
        CTkLabel(search_frame, text="🗳️ รายการรออนุมัติ", 
                font=CTkFont(size=14, weight="bold")).pack(side="left", padx=20)

        search_right = CTkFrame(search_frame, fg_color="transparent")
        search_right.pack(side="right", padx=20)
        
        self.approval_search_var = tk.StringVar()
        self.approval_search_entry = CTkEntry(
            search_right, 
            placeholder_text="🔍 ค้นหา SO หรือชื่อลูกค้า...", 
            width=300,
            height=36,
            textvariable=self.approval_search_var
        )
        self.approval_search_entry.pack(side="left", padx=(0, 8))
        self.approval_search_entry.bind("<Return>", lambda e: self._load_approval_data())

        CTkButton(search_right, text="ค้นหา", width=80, height=36,
                  command=self._load_approval_data).pack(side="left", padx=3)
        CTkButton(search_right, text="🔄 รีเฟรช", width=80, height=36, fg_color="gray",
                  command=self._load_approval_data).pack(side="left", padx=3)
        
        # รายการ SO
        self.approval_results_frame = CTkScrollableFrame(list_container, 
                                                         fg_color="white",
                                                         corner_radius=8)
        self.approval_results_frame.grid(row=1, column=0, sticky="nsew")
        self.approval_results_frame.grid_columnconfigure(0, weight=1)

        # โหลดข้อมูลครั้งแรก
        self.after(200, lambda: [self._update_rejection_chart(), self._load_approval_data()])

    def _update_rejection_chart(self):
        """วาดกราฟแท่งแนวนอนแสดงจำนวนการตีกลับของแต่ละ Sale (KPI)"""
        for widget in self.chart_area.winfo_children():
            widget.destroy()

        try:
            plt.close('all')
            try:
                plt.rcParams['font.family'] = 'TH Sarabun New'
            except:
                plt.rcParams['font.family'] = 'Tahoma'
            
            month_idx = self.thai_months.index(self.chart_month_var.get()) + 1
            year_val = int(self.chart_year_var.get()) - 543

            # filter ด้วย bill_date (วันที่เปิด SO) แทน commission_month เพื่อให้ตรงกับเดือนจริงที่ถูกตีกลับ
            query = """
                SELECT u.sale_name, SUM(COALESCE(c.sm_reject_count, 0)) as reject_count
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE EXTRACT(MONTH FROM c.bill_date) = %s
                  AND EXTRACT(YEAR FROM c.bill_date) = %s
                  AND c.is_active = 1
                  AND c.sm_reject_count > 0
                GROUP BY u.sale_name
                ORDER BY reject_count ASC
            """

            df = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year_val))
            
            if df.empty:
                empty_frame = CTkFrame(self.chart_area, fg_color="transparent")
                empty_frame.pack(expand=True, pady=30)
                CTkLabel(empty_frame, text="✅", font=CTkFont(size=48)).pack()
                CTkLabel(empty_frame, text="ไม่มีข้อมูลความผิดพลาดในรอบเดือนนี้ (KPI ดีเยี่ยม)",
                        font=CTkFont(size=15, weight="bold"), text_color="#16A34A").pack(pady=(5, 0))
                return

            num_sales = len(df)
            fig_height = max(2.0, min(3.5, num_sales * 0.8))  # จำกัดความสูงไว้ที่ 3.5 นิ้ว
            fig, ax = plt.subplots(figsize=(9.5, fig_height), dpi=80)  # ลด dpi ลง

            bars = ax.barh(df['sale_name'], df['reject_count'], color='#EF4444', height=0.65) # เปลี่ยนสีเป็นแดงสถิติ KPI
            
            ax.set_title(f"KPI สถิติความผิดพลาด (รอบ {self.chart_month_var.get()} {self.chart_year_var.get()})",
                        fontweight='bold', fontsize=16, pad=15)
            ax.set_xlabel("จำนวนครั้งที่ถูกตีกลับสะสม", fontsize=14, fontweight='bold')
            ax.set_ylabel("เซลล์", fontsize=14, fontweight='bold')
            
            ax.tick_params(axis='both', which='major', labelsize=13)
            ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
            
            ax.grid(axis='x', alpha=0.3, linestyle='--', linewidth=1.2)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_linewidth(1.5)
            ax.spines['bottom'].set_linewidth(1.5)

            for bar in bars:
                width = bar.get_width()
                if width > 0:
                    ax.text(width + 0.15, bar.get_y() + bar.get_height()/2,
                           f'{int(width)} ครั้ง', va='center', fontsize=13, fontweight='bold')

            plt.tight_layout(pad=1.5)
            canvas = FigureCanvasTkAgg(fig, master=self.chart_area)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="x", expand=False, padx=10, pady=5)
            
            plt.close(fig)
            
        except Exception as e:
            print(f"Chart Error: {traceback.format_exc()}")

    def _export_rejection_to_excel(self):
        """Export ข้อมูลสถิติ KPI การตีกลับเป็น Excel"""
        try:
            month_idx = self.thai_months.index(self.chart_month_var.get()) + 1
            year_val = int(self.chart_year_var.get()) - 543
            
            query = """
                SELECT
                    u.sale_name      AS sale_name,
                    c.so_number      AS so_number,
                    c.customer_name  AS customer_name,
                    c.sales_service_amount AS sales_amount,
                    c.sm_reject_count      AS reject_count,
                    c.rejection_reason     AS rejection_reason
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE EXTRACT(MONTH FROM c.bill_date) = %s
                  AND EXTRACT(YEAR FROM c.bill_date) = %s
                  AND c.is_active = 1
                  AND c.sm_reject_count > 0
                ORDER BY u.sale_name ASC, c.sm_reject_count DESC
            """

            df_raw = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year_val))

            if df_raw.empty:
                messagebox.showinfo("แจ้งเตือน", f"ไม่มีสถิติการตีกลับในรอบ {self.chart_month_var.get()} {self.chart_year_var.get()}")
                return

            df_detail = df_raw.rename(columns={
                'sale_name':        'ชื่อเซลล์',
                'so_number':        'เลขที่ SO',
                'customer_name':    'ชื่อลูกค้า',
                'sales_amount':     'ยอดขาย (บาท)',
                'reject_count':     'จำนวนครั้งที่ถูกตีกลับ',
                'rejection_reason': 'เหตุผลล่าสุดที่ถูกตีกลับ',
            })

            # สร้างหน้าสรุป KPI รวบยอด พร้อมรวมเหตุผลทั้งหมด
            summary_df = df_detail.groupby('ชื่อเซลล์').agg(
                reject_total=('จำนวนครั้งที่ถูกตีกลับ', 'sum'),
                reasons=('เหตุผลล่าสุดที่ถูกตีกลับ', lambda x: ' / '.join(v for v in x.dropna().unique() if str(v).strip()))
            ).reset_index().rename(columns={
                'reject_total': 'จำนวนครั้งที่ถูกตีกลับ',
                'reasons': 'เหตุผล',
            })
            summary_df = summary_df.sort_values('จำนวนครั้งที่ถูกตีกลับ', ascending=False)
            
            default_filename = f"KPI_รายงานการตีกลับ_{self.chart_month_var.get()}_{self.chart_year_var.get()}.xlsx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_filename,
                title="บันทึกรายงานสถิติ KPI การตีกลับ"
            )
            
            if not file_path: return
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='สรุป KPI รายบุคคล', index=False)
                df_detail.to_excel(writer, sheet_name='รายละเอียด SO ที่ผิดพลาด', index=False)
                
                # จัด Format Excel
                worksheet1 = writer.sheets['สรุป KPI รายบุคคล']
                worksheet1.column_dimensions['A'].width = 30
                worksheet1.column_dimensions['B'].width = 20
                worksheet1.column_dimensions['C'].width = 60
                
                worksheet2 = writer.sheets['รายละเอียด SO ที่ผิดพลาด']
                worksheet2.column_dimensions['A'].width = 25
                worksheet2.column_dimensions['B'].width = 20
                worksheet2.column_dimensions['C'].width = 35
                worksheet2.column_dimensions['D'].width = 18
                worksheet2.column_dimensions['E'].width = 25
                worksheet2.column_dimensions['F'].width = 50
                
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูล KPI เรียบร้อย!\nบันทึกที่: {file_path}")
            if os.name == 'nt': os.startfile(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Export Failed: {str(e)}")
            print(traceback.format_exc())

    def _load_approval_data(self):
        """ดึงรายการรออนุมัติ + คำขอแก้ไขรอบเดือนค่าคอม รวมในแท็บเดียวกัน"""
        for widget in self.approval_results_frame.winfo_children():
            widget.destroy()

        search_txt = self.approval_search_var.get().strip().upper()

        # ── ส่วนที่ 1: SO รออนุมัติปกติ ─────────────────────────────
        try:
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status,
                       u.sale_name, c.sales_service_amount,
                       c.bill_date, c.delivery_date, c.shipping_cost,
                       c.commission_month, c.commission_year
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Pending Sale Manager Approval' AND c.is_active = 1
            """
            params = []
            if search_txt:
                term = search_txt.replace("SO", "")
                query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                params.extend([f"%{term}%", f"%{term}%"])
            query += " ORDER BY c.timestamp ASC"

            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                empty_frame = CTkFrame(self.approval_results_frame, fg_color="transparent")
                empty_frame.pack(pady=30)
                icon = "✅" if not search_txt else "🔍"
                CTkLabel(empty_frame, text=icon, font=CTkFont(size=40)).pack(pady=(0, 6))
                msg = "ไม่มีรายการรออนุมัติในขณะนี้" if not search_txt else "ไม่พบรายการที่ตรงกับคำค้นหา"
                CTkLabel(empty_frame, text=msg,
                         font=CTkFont(size=14), text_color="#6B7280").pack()
            else:
                for _, row in df.iterrows():
                    self._create_so_card(self.approval_results_frame, row.to_dict(), is_approval_mode=True)

        except Exception as e:
            print(f"Load Approval Error: {e}")
            CTkLabel(self.approval_results_frame,
                     text=f"⚠️ เกิดข้อผิดพลาด: {str(e)[:100]}",
                     font=CTkFont(size=12), text_color="#DC2626").pack(pady=20)

        # ── ส่วนที่ 2: คำขอแก้ไขรอบเดือนค่าคอม (Sale Edit Requests) ────
        try:
            edit_query = """
                SELECT ser.id, ser.commission_id, ser.so_number,
                       ser.sale_key, ser.sale_name,
                       ser.current_commission_month, ser.current_commission_year,
                       ser.requested_commission_month, ser.requested_commission_year,
                       ser.request_reason, ser.requested_at
                FROM so_edit_requests ser
                WHERE ser.status = 'pending'
            """
            edit_params = []
            if search_txt:
                term = search_txt.replace("SO", "")
                edit_query += " AND ser.so_number ILIKE %s"
                edit_params.append(f"%{term}%")
            edit_query += " ORDER BY ser.requested_at ASC"

            df_edit = pd.read_sql_query(edit_query, self.pg_engine,
                                        params=tuple(edit_params))

            if not df_edit.empty:
                # Section header
                sep = CTkFrame(self.approval_results_frame,
                               fg_color="#FFF7ED", corner_radius=6,
                               border_width=1, border_color="#FED7AA")
                sep.pack(fill="x", padx=8, pady=(14, 4))
                CTkLabel(sep,
                         text="✏️  คำขอแก้ไขรอบเดือนค่าคอม  (Sale Edit Requests)",
                         font=CTkFont(size=13, weight="bold"),
                         text_color="#C2410C").pack(anchor="w", padx=14, pady=6)

                for _, row in df_edit.iterrows():
                    self._create_so_edit_card(self.approval_results_frame, row)

        except Exception as e:
            print(f"Load SO Edit Requests Error: {e}")

        # ── ส่วนที่ 3: คำขอยกเลิก SO ค่ารถ (Transport Cancel Requests) ─────
        try:
            # สร้าง table ถ้ายังไม่มี
            conn_chk = self.app_container.get_connection()
            try:
                with conn_chk.cursor() as _cur:
                    _cur.execute("""
                        CREATE TABLE IF NOT EXISTS transport_cancel_requests (
                            id SERIAL PRIMARY KEY,
                            so_number VARCHAR(50) NOT NULL,
                            so_id INTEGER,
                            sale_key VARCHAR(50),
                            customer_name VARCHAR(200),
                            original_status VARCHAR(50),
                            requested_by VARCHAR(100),
                            requested_at TIMESTAMP DEFAULT NOW(),
                            status VARCHAR(20) DEFAULT 'pending',
                            reviewed_by VARCHAR(100),
                            reviewed_at TIMESTAMP
                        )
                    """)
                conn_chk.commit()
            finally:
                self.app_container.release_connection(conn_chk)

            transport_query = """
                SELECT t.id, t.so_number, t.so_id, t.sale_key, t.customer_name,
                       t.original_status, t.requested_by, t.requested_at
                FROM transport_cancel_requests t
                WHERE t.status = 'pending'
            """
            transport_params = []
            if search_txt:
                term = search_txt.replace("SO", "")
                transport_query += " AND t.so_number ILIKE %s"
                transport_params.append(f"%{term}%")
            transport_query += " ORDER BY t.requested_at ASC"

            df_transport = pd.read_sql_query(transport_query, self.pg_engine,
                                             params=tuple(transport_params))

            if not df_transport.empty:
                sep_t = CTkFrame(self.approval_results_frame,
                                 fg_color="#FEF2F2", corner_radius=6,
                                 border_width=1, border_color="#FECACA")
                sep_t.pack(fill="x", padx=8, pady=(14, 4))
                CTkLabel(sep_t,
                         text="🚚  คำขอยกเลิก SO ค่ารถ  (Transport Cancel Requests)",
                         font=CTkFont(size=13, weight="bold"),
                         text_color="#B91C1C").pack(anchor="w", padx=14, pady=6)

                for _, row in df_transport.iterrows():
                    self._create_transport_cancel_card(self.approval_results_frame, row)

        except Exception as e:
            print(f"Load Transport Cancel Requests Error: {e}")

    def _create_transport_cancel_card(self, parent, row):
        req_id = row['id']
        so_number = row['so_number']
        so_id = row['so_id']
        customer_name = row.get('customer_name', '-')
        sale_key = row.get('sale_key', '-')
        requested_by = row.get('requested_by', '-')
        original_status = row.get('original_status', '')

        card = CTkFrame(parent, fg_color="#FEE2E2",
                        border_width=1, border_color="#FECACA", corner_radius=10, height=70)
        card.pack(fill="x", padx=8, pady=4)
        card.pack_propagate(False)

        info = CTkFrame(card, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True, padx=15, pady=8)

        CTkLabel(info,
                 text=f"🚚 SO: {so_number}  |  👤 {customer_name}",
                 font=CTkFont(size=13, weight="bold"),
                 anchor="w").pack(anchor="w")
        CTkLabel(info,
                 text=f"เซลล์: {sale_key}  |  ขอโดย HR: {requested_by}  |  เหตุผล: ค่ารถไม่อยู่ในเงื่อนไขคอมมิชชั่น",
                 font=CTkFont(size=12), text_color="#6B7280",
                 anchor="w").pack(anchor="w", pady=(2, 0))

        btn_frame = CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=12, pady=10)

        CTkButton(btn_frame, text="✅ อนุมัติ", width=85, height=32,
                  fg_color="#16A34A", hover_color="#15803D",
                  font=CTkFont(size=12, weight="bold"),
                  command=lambda: self._approve_transport_cancel(req_id, so_number, so_id)).pack(side="left", padx=3)
        CTkButton(btn_frame, text="❌ ปฏิเสธ", width=85, height=32,
                  fg_color="#DC2626", hover_color="#B91C1C",
                  font=CTkFont(size=12, weight="bold"),
                  command=lambda: self._reject_transport_cancel(req_id, so_number, so_id, original_status, sale_key)).pack(side="left", padx=3)

    def _approve_transport_cancel(self, req_id, so_number, so_id):
        if not messagebox.askyesno("ยืนยันการอนุมัติ",
                                   f"อนุมัติยกเลิก SO: {so_number}\n(ค่ารถไม่อยู่ในเงื่อนไขคอมมิชชั่น)\nSO จะถูกยกเลิกทันที", parent=self):
            return
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # อัปเดต transport_cancel_requests
                cursor.execute("""
                    UPDATE transport_cancel_requests
                    SET status = 'approved', reviewed_by = %s, reviewed_at = NOW()
                    WHERE id = %s
                """, (self.user_key, req_id))

                # ยกเลิก SO จริง
                cursor.execute("""
                    UPDATE commissions
                    SET status = 'Cancelled', is_active = 0,
                        rejection_reason = %s
                    WHERE so_number = %s
                """, (f"ยกเลิกโดย SM (ค่ารถ): อนุมัติคำขอจาก HR", so_number))

                # ยกเลิก PO ที่เกี่ยวข้อง
                cursor.execute("""
                    UPDATE purchase_orders
                    SET status = 'Cancelled', approval_status = 'Cancelled'
                    WHERE so_number = %s
                """, (so_number,))

                # แจ้ง SO owner + HR
                cursor.execute("SELECT sale_key FROM commissions WHERE so_number = %s", (so_number,))
                res = cursor.fetchone()
                if res:
                    cursor.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id, timestamp)
                        VALUES (%s, %s, FALSE, %s, NOW())
                    """, (res[0],
                          f"SO: {so_number} ถูกยกเลิกแล้ว (ค่ารถไม่อยู่ในเงื่อนไขคอมมิชชั่น)\nอนุมัติโดย SM",
                          so_id))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"อนุมัติยกเลิก SO: {so_number} เรียบร้อย")
            self._load_approval_data()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"Approve Failed: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _reject_transport_cancel(self, req_id, so_number, so_id, original_status, sale_key):
        if not messagebox.askyesno("ยืนยันการปฏิเสธ",
                                   f"ปฏิเสธคำขอยกเลิก SO: {so_number}\nSO จะกลับไปสถานะเดิม", parent=self):
            return
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("""
                    UPDATE transport_cancel_requests
                    SET status = 'rejected', reviewed_by = %s, reviewed_at = NOW()
                    WHERE id = %s
                """, (self.user_key, req_id))

                # คืนสถานะเดิม
                restore_status = original_status if original_status else 'Pending Sale Manager Approval'
                cursor.execute("""
                    UPDATE commissions SET status = %s WHERE so_number = %s
                """, (restore_status, so_number))

                # แจ้ง SO owner
                cursor.execute("""
                    INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id, timestamp)
                    VALUES (%s, %s, FALSE, %s, NOW())
                """, (sale_key,
                      f"SM ปฏิเสธคำขอยกเลิก SO: {so_number} (ค่ารถ)\nSO กลับสู่สถานะ: {restore_status}",
                      so_id))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ปฏิเสธคำขอยกเลิก SO: {so_number}\nSO กลับสถานะเดิมแล้ว")
            self._load_approval_data()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"Reject Failed: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    # =========================================================================
    # TAB: SO EDIT APPROVAL — อนุมัติ/ปฏิเสธ คำขอเปลี่ยนรอบเดือนค่าคอมจาก Sale
    # =========================================================================

    THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                   "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

    def _create_so_edit_approval_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        header = CTkFrame(parent_tab, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=15, pady=10)
        CTkLabel(header, text="✏️ คำขอแก้ไขรอบเดือนค่าคอมจาก Sale",
                 font=CTkFont(size=16, weight="bold")).pack(side="left")
        CTkButton(header, text="⟳ รีเฟรช", command=self._load_so_edit_requests,
                  width=90, fg_color="gray").pack(side="right")

        self.so_edit_req_frame = CTkScrollableFrame(parent_tab, fg_color="white", corner_radius=8)
        self.so_edit_req_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.after(200, self._load_so_edit_requests)

    def _load_so_edit_requests(self):
        for w in self.so_edit_req_frame.winfo_children():
            w.destroy()

        try:
            query = """
                SELECT ser.id, ser.commission_id, ser.so_number,
                       ser.sale_key, ser.sale_name,
                       ser.current_commission_month, ser.current_commission_year,
                       ser.requested_commission_month, ser.requested_commission_year,
                       ser.request_reason, ser.requested_at
                FROM so_edit_requests ser
                WHERE ser.status = 'pending'
                ORDER BY ser.requested_at DESC
            """
            df = pd.read_sql_query(query, self.pg_engine)
        except Exception as e:
            CTkLabel(self.so_edit_req_frame, text=f"❌ โหลดข้อมูลไม่ได้: {e}",
                     text_color="red", font=CTkFont(size=13)).pack(pady=30)
            return

        if df.empty:
            CTkLabel(self.so_edit_req_frame,
                     text="✅ ไม่มีคำขอแก้ไขรอบเดือนค่าคอมที่รอการอนุมัติ",
                     font=CTkFont(size=14)).pack(pady=40)
            return

        for _, row in df.iterrows():
            self._render_so_edit_card(row)

    def _render_so_edit_card(self, row):
        card = CTkFrame(self.so_edit_req_frame, fg_color="#FFF7ED",
                        border_width=1, border_color="#FED7AA", corner_radius=8)
        card.pack(fill="x", padx=10, pady=5)

        info = CTkFrame(card, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True, padx=15, pady=10)

        CTkLabel(info, text=f"SO: {row['so_number']}  |  Sale: {row['sale_name']} ({row['sale_key']})",
                 font=CTkFont(size=14, weight="bold")).pack(anchor="w")

        try:
            m_cur = int(row["current_commission_month"])
            y_cur = int(row["current_commission_year"]) + 543
            m_req = int(row["requested_commission_month"])
            y_req = int(row["requested_commission_year"]) + 543
            from_str = f"{self.THAI_MONTHS[m_cur-1]} {y_cur}"
            to_str = f"{self.THAI_MONTHS[m_req-1]} {y_req}"
        except Exception:
            from_str, to_str = str(row["current_commission_month"]), str(row["requested_commission_month"])

        CTkLabel(info,
                 text=f"🔄 ขอเปลี่ยนจาก: {from_str}  ➔  เป็น: {to_str}",
                 font=CTkFont(size=13, weight="bold"), text_color="#2563EB").pack(anchor="w", pady=(2, 0))
        CTkLabel(info,
                 text=f"💬 เหตุผล: {row.get('request_reason') or '-'}",
                 font=CTkFont(size=13), text_color="#C2410C").pack(anchor="w")
        try:
            req_at = str(row["requested_at"])[:16]
        except Exception:
            req_at = ""
        CTkLabel(info, text=f"🕐 ขอเมื่อ: {req_at}",
                 font=CTkFont(size=12), text_color="gray").pack(anchor="w")

        btn_f = CTkFrame(card, fg_color="transparent")
        btn_f.pack(side="right", padx=15, pady=10)

        CTkButton(btn_f, text="✅ อนุมัติ",
                  fg_color="#16A34A", hover_color="#15803D", width=100,
                  command=lambda r=row: self._decide_so_edit(r, approve=True)).pack(side="left", padx=5)
        CTkButton(btn_f, text="❌ ปฏิเสธ",
                  fg_color="#DC2626", hover_color="#B91C1C", width=100,
                  command=lambda r=row: self._decide_so_edit(r, approve=False)).pack(side="left", padx=5)

    def _decide_so_edit(self, row, approve):
        action_word = "อนุมัติ" if approve else "ปฏิเสธ"
        try:
            m_req = int(row["requested_commission_month"])
            y_req = int(row["requested_commission_year"]) + 543
            to_str = f"{self.THAI_MONTHS[m_req-1]} {y_req}"
        except Exception:
            to_str = str(row.get("requested_commission_month"))

        confirm_msg = (f"{action_word}การเปลี่ยนรอบค่าคอม\n"
                       f"SO: {row['so_number']}\n"
                       f"รอบใหม่: {to_str}")
        if not messagebox.askyesno(f"{action_word}?", confirm_msg):
            return

        note = ""
        if not approve:
            from tkinter.simpledialog import askstring
            note = askstring("หมายเหตุ", "ระบุเหตุผลที่ปฏิเสธ (ไม่บังคับ):", parent=self) or ""

        conn = None
        try:
            conn = self.app_container.get_connection()
            reviewer = getattr(self.app_container, 'current_user_key', 'Manager')
            with conn.cursor() as cur:
                new_status = "approved" if approve else "rejected"
                cur.execute("""
                    UPDATE so_edit_requests
                    SET status = %s, reviewed_by = %s, reviewed_at = NOW(), review_note = %s
                    WHERE id = %s
                """, (new_status, reviewer, note, int(row["id"])))

                if approve:
                    # อัปเดต commission_month/year ใน commissions ทันที
                    cur.execute("""
                        UPDATE commissions
                        SET commission_month = %s, commission_year = %s
                        WHERE id = %s
                    """, (int(row["requested_commission_month"]),
                          int(row["requested_commission_year"]),
                          int(row["commission_id"])))

                    # แจ้ง Sale ผ่าน notifications
                    try:
                        m_req = int(row["requested_commission_month"])
                        y_req = int(row["requested_commission_year"]) + 543
                        to_str_noti = f"{self.THAI_MONTHS[m_req-1]} {y_req}"
                        msg_noti = (f"[SO_EDIT] ✅ SM อนุมัติเปลี่ยนรอบค่าคอม SO: {row['so_number']}\n"
                                    f"📆 รอบใหม่: {to_str_noti}")
                    except Exception:
                        msg_noti = f"[SO_EDIT] ✅ อนุมัติเปลี่ยนรอบค่าคอม SO: {row['so_number']}"

                    cur.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read)
                        VALUES (%s, %s, FALSE)
                    """, (row["sale_key"], msg_noti))
                else:
                    # แจ้ง Sale ว่าถูกปฏิเสธ
                    msg_noti = (f"[SO_EDIT] ❌ SM ปฏิเสธคำขอเปลี่ยนรอบค่าคอม SO: {row['so_number']}\n"
                                f"💬 เหตุผล: {note or '-'}")
                    cur.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read)
                        VALUES (%s, %s, FALSE)
                    """, (row["sale_key"], msg_noti))

            conn.commit()
            self.app_container.release_connection(conn)
            conn = None
            messagebox.showinfo("สำเร็จ", f"{action_word}เรียบร้อยแล้ว")
            # refresh ทั้ง approval tab (main) และ so_edit tab (ถ้ามี)
            self._load_approval_data()
            if hasattr(self, "so_edit_req_frame"):
                self._load_so_edit_requests()
        except Exception as e:
            if conn:
                try:
                    conn.rollback()
                    self.app_container.release_connection(conn)
                except Exception:
                    pass
            messagebox.showerror("ผิดพลาด", f"ไม่สำเร็จ: {e}")

    # =========================================================================
    # TAB 1: DAILY REPORT WIDGET
    # =========================================================================
    def _create_daily_report_widget(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.daily_report_widget = DailyReportWidget(parent, self.app_container)
        self.daily_report_widget.pack(fill="both", expand=True)

    # =========================================================================
    # TAB 2: MASTER EDIT & SEARCH
    # =========================================================================
    def _create_master_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        filter_frame = CTkFrame(parent_tab)
        filter_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        
        current_year = datetime.now().year
        self.mst_year_var = tk.StringVar(value="ทุกปี")
        self.mst_month_var = tk.StringVar(value="ทุกเดือน")
        self.mst_day_var = tk.StringVar(value="ทุกวัน")
        self.mst_sale_var = tk.StringVar(value="All Sales")
        
        years = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 3, -1)]
        months = ["ทุกเดือน"] + ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                                  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        days = ["ทุกวัน"] + [str(d) for d in range(1, 32)]
        sales_list = self._get_sale_list()

        CTkLabel(filter_frame, text="ปี:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_year_var, values=years, width=75, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="เดือน:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_month_var, values=months, width=110, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="วัน:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_day_var, values=days, width=70, command=self._load_master_data).pack(side="left", padx=5)

        CTkLabel(filter_frame, text="Sale:", font=self.entry_font).pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_frame, variable=self.mst_sale_var, values=sales_list, width=120, command=self._load_master_data).pack(side="left", padx=5)

        self.sm_master_search_entry = CTkEntry(filter_frame, font=self.entry_font, placeholder_text="SO / ลูกค้า...", width=150)
        self.sm_master_search_entry.pack(side="left", padx=(15, 5), fill="x", expand=True)
        self.sm_master_search_entry.bind("<Return>", lambda e: self._load_master_data())
        
        CTkButton(filter_frame, text="🔍 ค้นหา", width=80, command=self._load_master_data).pack(side="left", padx=5)
        CTkButton(filter_frame, text="↺ รีเซ็ต", width=60, fg_color="gray", command=self._reset_master_filter).pack(side="left", padx=5)

        self.sm_master_results_frame = CTkScrollableFrame(parent_tab, label_text="รายการ SO ทั้งหมด (จัดการประวัติ)")
        self.sm_master_results_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.sm_master_results_frame.grid_columnconfigure(0, weight=1)

        self.after(200, self._load_master_data)

    def _get_sale_list(self):
        try:
            df = pd.read_sql_query("SELECT sale_key FROM sales_users WHERE role='Sale' ORDER BY sale_key", self.pg_engine)
            return ["All Sales"] + df['sale_key'].tolist()
        except: return ["All Sales"]

    def _reset_master_filter(self):
        self.mst_year_var.set("ทุกปี")
        self.mst_month_var.set("ทุกเดือน")
        self.mst_day_var.set("ทุกวัน")
        self.mst_sale_var.set("All Sales")
        self.sm_master_search_entry.delete(0, "end")
        self._load_master_data()

    def _load_master_data(self, event=None):
        for widget in self.sm_master_results_frame.winfo_children(): widget.destroy()
        try:
            query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status, u.sale_name, c.sales_service_amount
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.is_active = 1
            """
            params = []
            if self.mst_year_var.get() != "ทุกปี":
                query += " AND EXTRACT(YEAR FROM c.timestamp::timestamp) = %s"
                params.append(int(self.mst_year_var.get()))
            if self.mst_month_var.get() != "ทุกเดือน":
                thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
                params.append(thai_months.index(self.mst_month_var.get()) + 1)
                query += " AND EXTRACT(MONTH FROM c.timestamp::timestamp) = %s"
            if self.mst_day_var.get() != "ทุกวัน":
                query += " AND EXTRACT(DAY FROM c.timestamp::timestamp) = %s"; params.append(int(self.mst_day_var.get()))
            if self.mst_sale_var.get() != "All Sales":
                query += " AND c.sale_key = %s"; params.append(self.mst_sale_var.get())
            
            search_txt = self.sm_master_search_entry.get().strip().upper()
            if search_txt:
                term = search_txt.replace("SO", "")
                query += " AND (c.so_number ILIKE %s OR c.customer_name ILIKE %s)"
                params.extend([f"%{term}%", f"%{term}%"])

            query += " ORDER BY c.timestamp DESC LIMIT 20" 
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                CTkLabel(self.sm_master_results_frame, text="ไม่พบข้อมูล").pack(pady=20)
                return
            for _, row in df.iterrows():
                self._create_so_card(self.sm_master_results_frame, row.to_dict())
        except Exception as e: messagebox.showerror("Error", str(e))

    # =========================================================================
    # SHARED: SO CARD & ACTIONS
    # =========================================================================
    def _create_so_card(self, parent, so_data, is_approval_mode=False):
        """สร้าง Card แสดงข้อมูล SO พร้อมปุ่มต่างๆ"""
        so_id = so_data['id']
        so_number = so_data['so_number']
        status = so_data.get('status', 'N/A')
        amount = so_data.get('sales_service_amount', 0) or 0

        status_colors = {
            'PO In Progress': '#E0F2FE', 'Approved': '#DCFCE7', 'Paid': '#D1FAE5',
            'Rejected by SM': '#FEE2E2', 'Cancelled': '#F3F4F6', 'Draft': '#FEF3C7',
            'Pending Sale Manager Approval': '#FEF9C3'
        }
        bg_color = status_colors.get(status, "#FFFFFF")

        card = CTkFrame(parent, border_width=2, border_color="#E5E7EB",
                       corner_radius=10, fg_color=bg_color, height=70)
        card.pack(fill="x", padx=8, pady=6)
        card.pack_propagate(False)
        
        # ข้อมูล
        info_frame = CTkFrame(card, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, padx=15, pady=10)
        
        line1 = f"📋 SO: {so_number}  |  👤 {so_data.get('customer_name', '-')}"
        CTkLabel(info_frame, text=line1, 
                font=CTkFont(size=13, weight="bold"),
                anchor="w").pack(anchor="w")
        
        # 🟢 แปลงสถานะเป็นภาษาไทยเฉพาะตอนโชว์ข้อความ (สีพื้นหลังจะได้ไม่พัง)
        status_th = STATUS_THAI_MAP.get(status, status)
        
        line2 = f"🎯 เซลล์: {so_data.get('sale_name', '-')}  |  💰 {amount:,.2f} บาท  |  📌 {status_th}"
        CTkLabel(info_frame, text=line2,
                font=CTkFont(size=12),
                text_color="#6B7280",
                anchor="w").pack(anchor="w", pady=(3, 0))

        # ปุ่ม
        btn_frame = CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=12, pady=10)

        if is_approval_mode:
            CTkButton(btn_frame, text="✅ อนุมัติ", width=85, height=32,
                      fg_color="#16A34A", hover_color="#15803D",
                      font=CTkFont(size=12, weight="bold"),
                      command=lambda d=so_data: self._approve_so(d["id"], d["so_number"], d)).pack(side="left", padx=3)

        CTkButton(btn_frame, text="🛠️ แก้ไข", width=80, height=32,
                  fg_color="#4F46E5", hover_color="#4338CA",
                  font=CTkFont(size=12, weight="bold"),
                  command=lambda: self._open_so_editor_for_sm(so_number)).pack(side="left", padx=3)

        if status not in ['Cancelled', 'Rejected by SM']:
            CTkButton(btn_frame, text="❌ ตีกลับ", width=80, height=32,
                      fg_color="#DC2626", hover_color="#B91C1C",
                      font=CTkFont(size=12, weight="bold"),
                      command=lambda: self._reject_so(so_id, so_number)).pack(side="left", padx=3)

    # ✅ ฟังก์ชันอนุมัติ SO — เปิด dialog รายละเอียดก่อนยืนยัน
    def _approve_so(self, so_id, so_number, so_data=None):
        dlg = SMApproveConfirmDialog(self, so_data or {"so_number": so_number, "id": so_id})
        self.wait_window(dlg)
        if not dlg.confirmed:
            return

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("""
                    UPDATE commissions
                    SET status = 'Pending PU',
                        approver_sale_manager_key = %s,
                        approval_date_sale_manager = CURRENT_TIMESTAMP,
                        claim_timestamp = NULL
                    WHERE id = %s
                """, (self.user_key, so_id))

                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Staff' AND status = 'Active'")
                pu_keys = [row[0] for row in cursor.fetchall()]

                for pu_key in pu_keys:
                    cursor.execute("""
                        INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id)
                        VALUES (%s, %s, FALSE, %s)
                    """, (pu_key, f"มี SO ใหม่ ({so_number}) ผ่านการอนุมัติแล้ว รอคุณ Claim งาน", so_id))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"อนุมัติ SO: {so_number} เรียบร้อยแล้ว")
            self._refresh_all_tabs()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"Approve Failed: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    # ── SO Edit card (ยุบรวมในแท็บ Approval) ──────────────────────
    def _create_so_edit_card(self, parent, row):
        _tm = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
               "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

        card = CTkFrame(parent, fg_color="#FFF7ED",
                        border_width=1, border_color="#FED7AA", corner_radius=8)
        card.pack(fill="x", padx=8, pady=4)

        info = CTkFrame(card, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True, padx=15, pady=10)

        CTkLabel(info,
                 text=f"📋 SO: {row['so_number']}  |  👤 Sale: {row['sale_name']} ({row['sale_key']})",
                 font=CTkFont(size=13, weight="bold")).pack(anchor="w")

        try:
            m_c = int(row["current_commission_month"]); y_c = int(row["current_commission_year"]) + 543
            m_r = int(row["requested_commission_month"]); y_r = int(row["requested_commission_year"]) + 543
            from_s = f"{_tm[m_c-1]} {y_c}"; to_s = f"{_tm[m_r-1]} {y_r}"
        except Exception:
            from_s, to_s = str(row["current_commission_month"]), str(row["requested_commission_month"])

        CTkLabel(info, text=f"🔄 ขอเปลี่ยนรอบค่าคอม: {from_s}  ➔  {to_s}",
                 font=CTkFont(size=12, weight="bold"), text_color="#2563EB").pack(anchor="w", pady=(2, 0))
        CTkLabel(info, text=f"💬 เหตุผล: {row.get('request_reason') or '-'}",
                 font=CTkFont(size=12), text_color="#C2410C").pack(anchor="w")

        btn_f = CTkFrame(card, fg_color="transparent")
        btn_f.pack(side="right", padx=12, pady=10)

        CTkButton(btn_f, text="✅ อนุมัติ", fg_color="#16A34A", hover_color="#15803D",
                  width=90, height=32,
                  command=lambda r=row: self._decide_so_edit(r, approve=True)).pack(side="left", padx=4)
        CTkButton(btn_f, text="❌ ปฏิเสธ", fg_color="#DC2626", hover_color="#B91C1C",
                  width=90, height=32,
                  command=lambda r=row: self._decide_so_edit(r, approve=False)).pack(side="left", padx=4)

    def _reject_so(self, so_id, so_number):
        """เปิดหน้าต่างเลือกเหตุผลการตีกลับ"""
        def save_rejection(reason):
            conn = None
            try:
                conn = self.app_container.get_connection()
                with conn.cursor() as cursor:
                    # 🟢 แก้ไขตรงนี้: เพิ่ม sm_reject_count = COALESCE(sm_reject_count, 0) + 1
                    cursor.execute("""
                        UPDATE commissions 
                        SET status = 'Rejected by SM', 
                            rejection_reason = %s,
                            sm_reject_count = COALESCE(sm_reject_count, 0) + 1 
                        WHERE id = %s
                    """, (reason, so_id))
                    
                    cursor.execute("SELECT sale_key FROM commissions WHERE id = %s", (so_id,))
                    res = cursor.fetchone()
                    
                    if res:
                        cursor.execute("""
                            INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) 
                            VALUES (%s, %s, FALSE, %s)
                        """, (res[0], f"SO: {so_number} ถูกตีกลับ: {reason}", so_id))
                
                conn.commit()
                messagebox.showinfo("สำเร็จ", f"ตีกลับ SO: {so_number} เรียบร้อยแล้ว\nระบบบันทึกสถิติความผิดพลาด +1 ครั้ง")
                self._refresh_all_tabs()
                
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Error", f"Reject Failed: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)

        SORejectionDialog(self, so_number, save_rejection)

    def _open_so_editor_for_sm(self, so_number):
        if self.so_popup is not None and self.so_popup.winfo_exists():
            self.so_popup.focus()
            return
        try:
            so_df = pd.read_sql_query("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", 
                                     self.pg_engine, params=(so_number,))
            if so_df.empty:
                messagebox.showerror("Error", "ไม่พบข้อมูล SO")
                return
            
            def _refresh_on_save():
                self._refresh_all_tabs()
            
            self.so_popup = SOPopupWindow(
                master=self,
                app_container=self.app_container,
                sales_data=so_df.iloc[0].to_dict(),
                so_shared_vars=self.so_shared_vars,
                sale_theme=self.sale_theme,
                on_save_callback=_refresh_on_save
            )
        except Exception as e:
            messagebox.showerror("Error", f"Open Editor Failed: {e}")
            print(traceback.format_exc())

    def _so_create_string_vars(self):
        """สร้าง StringVars สำหรับหน้าจอแก้ไข SO"""
        self.so_shared_vars = {}
        now = datetime.now()
        thai_months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        self.so_shared_vars.update({
            'thai_months': thai_months_list,
            'thai_month_map': {name: i + 1 for i, name in enumerate(thai_months_list)},
            'customer_type_var': tk.StringVar(value="ลูกค้าเก่า"),
            'credit_term_var': tk.StringVar(value="เงินสด"),
            'commission_month_var': tk.StringVar(value=thai_months_list[now.month - 1]),
            'commission_year_var': tk.StringVar(value=str(now.year + 543)),
            'payment1_percent_var': tk.StringVar(value="ระบุยอดเอง"),
            'payment2_percent_var': tk.StringVar(value="ระบุยอดเอง"),
            'delivery_type_var': tk.StringVar(value="ซัพพลายเออร์จัดส่ง"),
            'payment_total_var': tk.StringVar(value="0.00"),
            'so_subtotal_var': tk.StringVar(value="0.00"),
            'so_vat_var': tk.StringVar(value="0.00"),
            'so_grand_total_var': tk.StringVar(value="0.00"),
            'so_vs_payment_result_var': tk.StringVar(value="-"),
            'difference_amount_var': tk.StringVar(value="0.00"),
            'balance_due_var': tk.StringVar(value="0.00"),
            'cash_product_input_var': tk.StringVar(value="0.00"),
            'cash_service_total_var': tk.StringVar(value="0.00"),
            'cash_required_total_var': tk.StringVar(value="0.00"),
            'cash_actual_payment_var': tk.StringVar(value="0.00"),
            'cash_verification_result_var': tk.StringVar(value="-"),
            'sales_vat_calc_var': tk.StringVar(value="0.00"),
            'cutting_drilling_vat_calc_var': tk.StringVar(value="0.00"),
            'other_service_vat_calc_var': tk.StringVar(value="0.00"),
            'shipping_vat_calc_var': tk.StringVar(value="0.00"),
            'card_fee_vat_calc_var': tk.StringVar(value="0.00"),
            'relocation_vat_calc_var': tk.StringVar(value="0.00"),
            'sales_service_vat_option': tk.StringVar(value="VAT"),
            'cutting_drilling_fee_vat_option': tk.StringVar(value="VAT"),
            'other_service_fee_vat_option': tk.StringVar(value="VAT"),
            'shipping_vat_option_var': tk.StringVar(value="VAT"),
            'credit_card_fee_vat_option_var': tk.StringVar(value="VAT"),
            'relocation_cost_vat_option': tk.StringVar(value="VAT")
        })

    # =========================================================================
    # ระบบยกเลิก SO (Cancel Logic)
    # =========================================================================
    def _search_so_to_cancel(self):
        """ค้นหา SO และแสดงปุ่มยกเลิกในแถบเดียวกัน"""
        # ล้างผลลัพธ์เดิม
        for widget in self.inline_result_frame.winfo_children():
            widget.destroy()
            
        so_number = self.cancel_search_entry.get().strip().upper()
        if not so_number:
            CTkLabel(self.inline_result_frame, text="⚠️ กรุณาระบุเลข SO", text_color="#DC2626", font=self.label_font_bold).pack(side="left")
            return
            
        try:
            # ค้นหา SO ล่าสุดที่ Active อยู่
            query = "SELECT id, so_number, customer_name, status FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1"
            df = pd.read_sql_query(query, self.pg_engine, params=(so_number,))
            
            if df.empty:
                CTkLabel(self.inline_result_frame, text="❌ ไม่พบข้อมูล SO นี้ในระบบ", text_color="#DC2626", font=self.label_font_bold).pack(side="left")
                return
                
            row = df.iloc[0]
            current_status = row['status']
            
            # เช็คสถานะก่อนว่ายกเลิกได้ไหม
            if current_status == 'Cancelled':
                CTkLabel(self.inline_result_frame, text="⚠️ SO นี้ถูกยกเลิกไปแล้ว", text_color="#D97706", font=self.label_font_bold).pack(side="left")
                return
            if current_status == 'Paid':
                CTkLabel(self.inline_result_frame, text="⚠️ ไม่สามารถยกเลิกได้ (SO นี้ชำระเงินเรียบร้อยแล้ว)", text_color="#D97706", font=self.label_font_bold).pack(side="left")
                return
                
            # ถ้าสามารถยกเลิกได้ แสดงข้อมูล + ปุ่มยกเลิก
            info_text = f"✅ พบ SO: {row['so_number']} | ลูกค้า: {row['customer_name']} | สถานะ: {current_status}"
            CTkLabel(self.inline_result_frame, text=info_text, text_color="#16A34A", font=self.label_font_bold).pack(side="left", padx=(0, 15))
            
            CTkButton(self.inline_result_frame, text="🗑️ ยืนยันยกเลิก SO", fg_color="#DC2626", hover_color="#B91C1C", 
                      width=130, height=30, font=CTkFont(weight="bold"),
                      command=lambda: self._execute_cancel_so(row['id'], row['so_number'])).pack(side="left")
                      
        except Exception as e:
            CTkLabel(self.inline_result_frame, text=f"Error: {str(e)[:50]}", text_color="red").pack(side="left")

    def _execute_cancel_so(self, so_id, so_number):
        """เรียก Popup ระบุเหตุผล และดำเนินการยกเลิก SO"""
        
        def process_cancel(reason):
            # ถามย้ำอีกครั้งเพื่อความชัวร์ ป้องกันกดพลาด
            if not messagebox.askyesno("ยืนยันครั้งสุดท้าย", f"คุณแน่ใจหรือไม่ที่จะยกเลิก SO: {so_number} อย่างถาวร?"):
                return

            so_id_int = int(so_id)  # pandas อ่านค่าเป็น numpy.int64 ต้อง convert ก่อนส่งให้ psycopg2

            conn = None
            try:
                conn = self.app_container.get_connection()
                with conn.cursor() as cursor:
                    # เปลี่ยนสถานะเป็น Cancelled + ปิด active ไม่ให้คิดคอม
                    cursor.execute("""
                        UPDATE commissions
                        SET status = 'Cancelled',
                            is_active = 0,
                            rejection_reason = %s,
                            sm_reject_count = COALESCE(sm_reject_count, 0) + 1
                        WHERE id = %s
                    """, (f"ยกเลิกโดย SM: {reason}", so_id_int))

                    # ยกเลิก PO ที่เกี่ยวข้องทั้งหมด
                    cursor.execute("""
                        UPDATE purchase_orders
                        SET status = 'Cancelled',
                            approval_status = 'Cancelled'
                        WHERE so_number = %s
                    """, (so_number,))

                    # แจ้งเตือนเซลล์เจ้าของ SO ผ่าน Noti
                    cursor.execute("SELECT sale_key FROM commissions WHERE id = %s", (so_id_int,))
                    res = cursor.fetchone()
                    if res:
                        cursor.execute("""
                            INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id)
                            VALUES (%s, %s, FALSE, %s)
                        """, (res[0], f"SO: {so_number} ถูกยกเลิกโดย Manager เหตุผล: {reason}", so_id_int))

                    # Audit Log
                    import json as _json
                    cursor.execute("""
                        INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                        VALUES (%s, %s, %s, %s, %s, NOW())
                    """, ('Cancel SO', 'commissions', so_id_int,
                          getattr(self.app_container, 'current_user_key', 'SM'),
                          _json.dumps({"reason": reason, "cancelled_by": "SM"})))

                conn.commit()
                messagebox.showinfo("สำเร็จ", f"ยกเลิก SO: {so_number} เรียบร้อยแล้ว")
                
                # ล้างหน้าจอและโหลดข้อมูลใหม่
                self.cancel_search_entry.delete(0, tk.END)
                for widget in self.inline_result_frame.winfo_children(): widget.destroy()
                self._load_cancelled_so_history()
                self._refresh_all_tabs() # อัปเดตตารางอื่นๆ ด้วย
                
            except Exception as e:
                if conn: conn.rollback()
                messagebox.showerror("Error", f"ไม่สามารถยกเลิกได้: {e}")
            finally:
                if conn: self.app_container.release_connection(conn)

        # เปิดหน้าต่าง Popup SOCancelDialog ขึ้นมาให้เลือกเหตุผล
        SOCancelDialog(self, so_number, process_cancel)

    def _load_cancelled_so_history(self):
        """โหลดข้อมูลประวัติ SO ที่ถูกยกเลิกลง Treeview"""
        for widget in self.cancelled_history_frame.winfo_children():
            widget.destroy()
            
        try:
            query = """
                SELECT c.timestamp, c.so_number, c.customer_id, c.customer_name,
                       u.sale_name, c.rejection_reason
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Cancelled'
                ORDER BY c.timestamp DESC LIMIT 100
            """
            df = pd.read_sql_query(query, self.pg_engine)

            columns = ("วันที่", "SO Number", "รหัสลูกค้า", "ชื่อลูกค้า", "เซลล์", "เหตุผลที่ยกเลิก")
            tree = ttk.Treeview(self.cancelled_history_frame, columns=columns, show="headings", height=15)

            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview.Heading", font=('Tahoma', 12, 'bold'), background="#F3F4F6")
            style.configure("Treeview", font=('Tahoma', 11), rowheight=30)

            tree.heading("วันที่", text="วันที่ / เวลา")
            tree.heading("SO Number", text="SO Number")
            tree.heading("รหัสลูกค้า", text="รหัสลูกค้า")
            tree.heading("ชื่อลูกค้า", text="ชื่อลูกค้า")
            tree.heading("เซลล์", text="พนักงานขาย")
            tree.heading("เหตุผลที่ยกเลิก", text="เหตุผลที่ยกเลิก")

            tree.column("วันที่", width=150, anchor="center")
            tree.column("SO Number", width=150, anchor="center")
            tree.column("รหัสลูกค้า", width=120, anchor="center")
            tree.column("ชื่อลูกค้า", width=280, anchor="w")
            tree.column("เซลล์", width=150, anchor="center")
            tree.column("เหตุผลที่ยกเลิก", width=380, anchor="w")

            vsb = ttk.Scrollbar(self.cancelled_history_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vsb.set)

            tree.pack(side="left", fill="both", expand=True)
            vsb.pack(side="right", fill="y")

            if df.empty:
                tree.insert("", "end", values=("-", "ไม่มีประวัติการยกเลิก SO", "-", "-", "-", "-"))
                return

            for _, row in df.iterrows():
                ts = row['timestamp']
                ts_str = str(ts)[:16] if pd.notna(ts) else "-"
                reason = row['rejection_reason'] or "ไม่ได้ระบุเหตุผล"
                
                tree.insert("", "end", values=(
                    ts_str,
                    row['so_number'],
                    row['customer_id'] or "-",
                    row['customer_name'] or "-",
                    row['sale_name'] or "-",
                    reason
                ))
                
        except Exception as e:
            CTkLabel(self.cancelled_history_frame, text=f"Error loading history: {e}", text_color="red").pack()

    def _create_defer_approval_tab(self, parent_tab):
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1)

        # Header
        header_frame = CTkFrame(parent_tab, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=10)
        CTkLabel(header_frame, text="⏳ รายการที่ HR/บัญชี ขอเลื่อนรอบจ่ายคอมมิชชั่น", font=CTkFont(size=16, weight="bold")).pack(side="left")
        CTkButton(header_frame, text="⟳ รีเฟรช", command=self._load_defer_requests, width=90, fg_color="gray").pack(side="right")
        CTkButton(header_frame, text="📋 ประวัติการเลื่อน SO",
                  command=lambda: DeferralHistoryWindow(self, self.app_container),
                  fg_color="#2563EB", hover_color="#1D4ED8", width=160).pack(side="right", padx=(0, 8))

        # Scrollable Frame
        self.defer_list_frame = CTkScrollableFrame(parent_tab, fg_color="white", corner_radius=8)
        self.defer_list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.after(200, self._load_defer_requests)

    def _load_defer_requests(self):
        for widget in self.defer_list_frame.winfo_children(): widget.destroy()

        try:
            # 🟢 [แก้ไข] เพิ่มการดึง commission_month และ commission_year มาจาก DB ด้วย
            query = """
                SELECT c.id, c.so_number, c.customer_name, u.sale_name, c.sale_key, c.rejection_reason,
                       c.commission_month, c.commission_year
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status = 'Defer Requested' AND c.is_active = 1
                ORDER BY c.timestamp DESC
            """
            df = pd.read_sql_query(query, self.pg_engine)

            if df.empty:
                CTkLabel(self.defer_list_frame, text="✅ ไม่มีรายการที่ขอเลื่อน SO ในขณะนี้", font=CTkFont(size=14)).pack(pady=40)
                return

            thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                           "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

            for _, row in df.iterrows():
                card = CTkFrame(self.defer_list_frame, fg_color="#FFF7ED", border_width=1, border_color="#FDBA74", corner_radius=8)
                card.pack(fill="x", padx=10, pady=5)

                info_frame = CTkFrame(card, fg_color="transparent")
                info_frame.pack(side="left", fill="both", expand=True, padx=15, pady=10)

                # 🟢 [เพิ่มใหม่] ลอจิกคำนวณเดือนที่ถูกเลื่อน (จากเดือนไหน -> ไปเดือนไหน)
                try:
                    m_current = int(row['commission_month'])
                    y_current = int(row['commission_year']) + 543
                    
                    from_month_str = thai_months[m_current - 1] if 1 <= m_current <= 12 else str(m_current)
                    from_year_str = str(y_current)
                    
                    # คำนวณเดือนถัดไป
                    m_next = m_current + 1
                    y_next = y_current
                    if m_next > 12:
                        m_next = 1
                        y_next += 1
                        
                    to_month_str = thai_months[m_next - 1]
                    to_year_str = str(y_next)
                    
                    defer_period_text = f"🔄 ขอเลื่อนจาก: {from_month_str} {from_year_str}  ➔  ไปเป็น: {to_month_str} {to_year_str}"
                except Exception:
                    defer_period_text = "🔄 ขอเลื่อนไปเดือนถัดไป"

                # แสดงบรรทัดที่ 1: เลข SO + ชื่อลูกค้า
                CTkLabel(info_frame, text=f"SO: {row['so_number']} | ลูกค้า: {row['customer_name']}", font=CTkFont(size=14, weight="bold")).pack(anchor="w")
                
                # แสดงบรรทัดที่ 2: ข้อความบอกรอบเดือน (สีน้ำเงิน) 🟢 [เพิ่มใหม่]
                CTkLabel(info_frame, text=defer_period_text, text_color="#2563EB", font=CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(2, 0))
                
                # แสดงบรรทัดที่ 3: ชื่อเซลล์ + เหตุผล (สีส้ม)
                CTkLabel(info_frame, text=f"👤 เซลล์: {row['sale_name']} | 💬 เหตุผลที่บัญชีขอเลื่อน: {row['rejection_reason']}", text_color="#C2410C", font=CTkFont(size=13)).pack(anchor="w")

                # ส่วนปุ่มกด
                btn_frame = CTkFrame(card, fg_color="transparent")
                btn_frame.pack(side="right", padx=15, pady=10)

                CTkButton(btn_frame, text="✅ อนุมัติให้เลื่อน", fg_color="#16A34A", hover_color="#15803D", width=120,
                          command=lambda r=row: self._action_defer(r, approve=True)).pack(side="left", padx=5)

                CTkButton(btn_frame, text="❌ ไม่อนุมัติ (บังคับจ่าย)", fg_color="#DC2626", hover_color="#B91C1C", width=140,
                          command=lambda r=row: self._action_defer(r, approve=False)).pack(side="left", padx=5)

        except Exception as e:
            print(f"Error loading defer requests: {e}")

    def _action_defer(self, row, approve):
        try:
            cur_m = int(row.get('commission_month', datetime.now().month))
            cur_y = int(row.get('commission_year', datetime.now().year))
        except Exception:
            cur_m, cur_y = datetime.now().month, datetime.now().year

        dialog = ManagerDeferApprovalDialog(self, row['so_number'], cur_m, cur_y, approve=approve)
        self.wait_window(dialog)

        if not dialog.confirmed:
            return

        reason = dialog.reason or ("Manager อนุมัติการเลื่อน" if approve else "Manager ไม่อนุมัติการเลื่อน บังคับจ่ายรอบนี้")

        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                manager_key = getattr(self.app_container, 'current_user_key', 'Manager')

                if approve:
                    new_status = 'Deferred'
                    defer_decision = 'อนุมัติ'
                    t_m, t_y = dialog.target_month, dialog.target_year
                    thai_months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
                    month_label = f"{thai_months[t_m-1]} {t_y+543}"
                    msg_for_sale = (f"[DEFER] ✅ Manager อนุมัติเลื่อนคอม SO: {row['so_number']}\n"
                                    f"📅 รอบคอมที่จะนำกลับมาคิด: {month_label}\n"
                                    f"💬 เหตุผล: {reason}")
                    cursor.execute("""
                        UPDATE commissions
                        SET status = %s, rejection_reason = %s,
                            defer_decision = %s, defer_decision_reason = %s,
                            defer_approved_by = %s,
                            commission_month = %s, commission_year = %s
                        WHERE id = %s
                    """, (new_status, f"Manager Decision: {reason}", defer_decision, reason,
                          manager_key, t_m, t_y, row['id']))
                else:
                    new_status = 'Pending HR Approval'
                    defer_decision = 'ไม่อนุมัติ'
                    msg_for_sale = (f"[DEFER] ❌ Manager ไม่อนุมัติเลื่อนคอม SO: {row['so_number']}\n"
                                    f"บังคับจ่ายรอบปัจจุบัน\n💬 เหตุผล: {reason}")
                    cursor.execute("""
                        UPDATE commissions
                        SET status = %s, rejection_reason = %s,
                            defer_decision = %s, defer_decision_reason = %s,
                            defer_approved_by = %s
                        WHERE id = %s
                    """, (new_status, f"Manager Decision: {reason}", defer_decision, reason,
                          manager_key, row['id']))

                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id) VALUES (%s, %s, FALSE, %s)",
                    (row['sale_key'], msg_for_sale, row['id']))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"บันทึกการตัดสินใจ SO: {row['so_number']} เรียบร้อยแล้ว")
            self._load_defer_requests()
            self._refresh_all_tabs()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", str(e))
        finally:
            if conn: self.app_container.release_connection(conn)

class SOCancelDialog(CTkToplevel):
    def __init__(self, master, so_number, on_confirm_callback):
        super().__init__(master)
        self.title(f"ยกเลิก SO: {so_number}")
        self.geometry("450x550")
        self.on_confirm_callback = on_confirm_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True) 

        CTkLabel(self, text=f"ระบุเหตุผลที่ยกเลิก SO: {so_number}", 
                font=CTkFont(size=16, weight="bold")).pack(pady=15)

        # รายการ Checkbox 
        self.reasons = [
            "1. สินค้าไม่พอ/หาของไม่ได้",
            "2. ลูกค้าเปลี่ยนสเปค/สั่งผิด",
            "3. แจ้งราคาผิดพลาด",
            "4. ราคาปรับขึ้น/ลูกค้าไม่สู้",
            "5. ขนส่งช้า/ไม่ตรงนัด",
            "6. งานด่วน/ส่งไม่ทัน",
            "7. ข้อมูลสเปคคลาดเคลื่อน",
            "8. เปิด SO ซ้ำ/ผิดพลาด"
        ]
        
        self.check_vars = []
        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=30)

        for reason in self.reasons:
            var = tk.BooleanVar(value=False)
            cb = CTkCheckBox(container, text=reason, variable=var, font=CTkFont(size=13))
            cb.pack(anchor="w", pady=5)
            self.check_vars.append((var, reason))

        # ช่องกรอกข้อมูลเพิ่มเติม (อื่นๆ)
        CTkLabel(self, text="อื่นๆ / ระบุเพิ่มเติม:", 
                font=CTkFont(size=13, weight="bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.other_reason_entry = CTkEntry(self, placeholder_text="พิมพ์เหตุผลเพิ่มเติมที่นี่...", width=380)
        self.other_reason_entry.pack(padx=30, pady=(0, 20))

        # ปุ่มกดยกเลิก/ตกลง
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=10)
        
        CTkButton(btn_frame, text="ปิด", fg_color="gray", width=100, 
                 command=self.destroy).pack(side="left", padx=5)
        CTkButton(btn_frame, text="ตกลง (ยกเลิก SO)", fg_color="#DC2626", hover_color="#B91C1C", 
                 width=130, command=self._on_confirm).pack(side="right", padx=5)

    def _on_confirm(self):
        selected_reasons = [text for var, text in self.check_vars if var.get()]
        other_text = self.other_reason_entry.get().strip()
        
        if other_text:
            selected_reasons.append(other_text)

        if not selected_reasons:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกหรือระบุเหตุผลอย่างน้อย 1 ข้อ", parent=self)
            return

        final_reason = ", ".join(selected_reasons)
        self.on_confirm_callback(final_reason)
        self.destroy()

class SORejectionDialog(CTkToplevel):
    def __init__(self, master, so_number, on_confirm_callback):
        super().__init__(master)
        self.title(f"ตีกลับ SO: {so_number}")
        self.geometry("450x550")
        self.on_confirm_callback = on_confirm_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True) 

        CTkLabel(self, text=f"ระบุเหตุผลที่ตีกลับ SO: {so_number}", 
                font=CTkFont(size=16, weight="bold")).pack(pady=15)

        # รายการ Checkbox สำหรับตีกลับ
        self.reasons = [
            "1. เอกสารแนบไม่ครบถ้วน",
            "2. ข้อมูลลูกค้า/ที่อยู่ ไม่ถูกต้อง",
            "3. ยอดเงินไม่ตรงกับใบเสนอราคา",
            "4. เครดิต/เงื่อนไขการชำระเงินไม่ถูกต้อง",
            "5. รายการสินค้า/ราคา ไม่ถูกต้อง",
            "6. ข้อมูลการจัดส่งไม่ชัดเจน",
            "7. รอเอกสารจากทางฝั่งลูกค้า",
            "8. อื่นๆ (ระบุเพิ่มเติมด้านล่าง)"
        ]
        
        self.check_vars = []
        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=30)

        for reason in self.reasons:
            var = tk.BooleanVar(value=False)
            cb = CTkCheckBox(container, text=reason, variable=var, font=CTkFont(size=13))
            cb.pack(anchor="w", pady=5)
            self.check_vars.append((var, reason))

        # ช่องกรอกข้อมูลเพิ่มเติม (อื่นๆ)
        CTkLabel(self, text="อื่นๆ / ระบุเพิ่มเติม:", 
                font=CTkFont(size=13, weight="bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.other_reason_entry = CTkEntry(self, placeholder_text="พิมพ์เหตุผลเพิ่มเติมที่นี่...", width=380)
        self.other_reason_entry.pack(padx=30, pady=(0, 20))

        # ปุ่มกด
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=10)
        
        CTkButton(btn_frame, text="ปิด", fg_color="gray", width=100, 
                 command=self.destroy).pack(side="left", padx=5)
        CTkButton(btn_frame, text="ตกลง (ส่งกลับไปแก้)", fg_color="#DC2626", hover_color="#B91C1C", 
                 width=140, command=self._on_confirm).pack(side="right", padx=5)

        self.transient(master)
        self.grab_set()
        self.focus_force()

    def _on_confirm(self):
        selected_reasons = [text for var, text in self.check_vars if var.get()]
        other_text = self.other_reason_entry.get().strip()
        
        if other_text:
            selected_reasons.append(other_text)

        if not selected_reasons:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกหรือระบุเหตุผลอย่างน้อย 1 ข้อ", parent=self)
            return

        final_reason = ", ".join(selected_reasons)
        self.on_confirm_callback(final_reason)
        self.destroy()

class SMExportDialog(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        
        self.title("📥 Export ข้อมูล SO ตามรอบคอมมิชชั่น")
        self.geometry("450x420") # ขยายความสูงหน้าต่างนิดหน่อยเผื่อที่ให้ข้อความ
        self.grid_columnconfigure(0, weight=1)
        self.attributes("-topmost", True)

        # --- 1. เตรียมข้อมูล Dropdown ---
        now = datetime.now()
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        
        current_year_be = now.year + 543
        self.years = [str(y) for y in range(current_year_be - 2, current_year_be + 2)]
        
        self.sales_list = ["ทั้งหมด (All Sales)"]
        self.sale_mapping = {}
        self._load_sales_users()

        # --- 2. ตัวแปรเก็บค่าที่เลือก ---
        self.selected_year = tk.StringVar(value=str(current_year_be))
        self.selected_month = tk.StringVar(value=self.thai_months[now.month - 1])
        self.selected_sale = tk.StringVar(value="ทั้งหมด (All Sales)")

        # --- 3. สร้าง UI ---
        # หัวข้อหลัก (ปรับ pady ด้านล่างให้ชิดข้อความ Note มากขึ้น)
        CTkLabel(self, text="เลือกเงื่อนไขเพื่อ Export ข้อมูล", font=CTkFont(size=16, weight="bold")).pack(pady=(20, 5))
        
        # 🟢 [เพิ่มตรงนี้] ข้อความ Note อธิบายเพิ่มเติม
        CTkLabel(self, text="* Export Report เพื่อตรวจสอบค่าคอมมิชชั่น ประจำงวด", 
                 font=CTkFont(size=12), text_color="gray50").pack(pady=(0, 15))

        form_frame = CTkFrame(self, fg_color="transparent")
        form_frame.pack(fill="x", padx=40)
        form_frame.grid_columnconfigure(1, weight=1)

        # เปลี่ยนคำให้ชัดเจนขึ้นว่าเป็น "รอบคิดคอมมิชชั่น"
        CTkLabel(form_frame, text="รอบคอมมิชชั่น (ปี พ.ศ.):", font=CTkFont(size=14)).grid(row=0, column=0, sticky="w", pady=10)
        CTkOptionMenu(form_frame, variable=self.selected_year, values=self.years).grid(row=0, column=1, sticky="ew", padx=(10, 0))

        CTkLabel(form_frame, text="รอบคอมมิชชั่น (เดือน):", font=CTkFont(size=14)).grid(row=1, column=0, sticky="w", pady=10)
        CTkOptionMenu(form_frame, variable=self.selected_month, values=self.thai_months).grid(row=1, column=1, sticky="ew", padx=(10, 0))

        # Sale
        CTkLabel(form_frame, text="พนักงานขาย:", font=CTkFont(size=14)).grid(row=2, column=0, sticky="w", pady=10)
        CTkOptionMenu(form_frame, variable=self.selected_sale, values=self.sales_list).grid(row=2, column=1, sticky="ew", padx=(10, 0))

        # ปุ่มกด
        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=40, pady=30)
        
        CTkButton(btn_frame, text="ยกเลิก", fg_color="gray", width=120, command=self.destroy).pack(side="left")
        CTkButton(btn_frame, text="📥 Export Excel", fg_color="#10B981", hover_color="#059669", 
                  width=120, command=self._execute_export).pack(side="right")
    def _load_sales_users(self):
        """ดึงรายชื่อเซลส์ทั้งหมดมาใส่ใน Dropdown"""
        try:
            df = pd.read_sql_query("SELECT sale_key, sale_name FROM sales_users WHERE role = 'Sale' ORDER BY sale_key", self.pg_engine)
            for _, row in df.iterrows():
                display_name = f"[{row['sale_key']}] {row['sale_name']}"
                self.sales_list.append(display_name)
                self.sale_mapping[display_name] = row['sale_key'] # เก็บ mapping ไว้ค้นหา
        except Exception as e:
            print(f"Error loading sales users: {e}")

    def _execute_export(self):
        thai_months_name = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                            "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

        # 1. แปลงค่าจาก Dropdown เป็นค่าสำหรับ Database
        month_num = self.month_map[self.selected_month.get()]
        year_ad = int(self.selected_year.get()) - 543
        sale_selection = self.selected_sale.get()

        # 2. สร้าง SQL Query
        query = """
            SELECT
                c.so_number          AS "เลขที่ SO",
                c.customer_name      AS "ชื่อลูกค้า",
                c.sales_service_amount AS "ยอดขาย (บาท)",
                COALESCE(u.sale_name, c.sale_key) AS "พนักงานขาย",
                c.status             AS "_status",
                c.defer_source_month AS "_defer_src_m",
                c.defer_source_year  AS "_defer_src_y"
            FROM commissions c
            LEFT JOIN sales_users u ON c.sale_key = u.sale_key
            WHERE c.is_active = 1
              AND c.commission_month = %s
              AND c.commission_year = %s
              AND c.status NOT IN ('Cancelled', 'Original')
        """
        params = [month_num, year_ad]

        if sale_selection != "ทั้งหมด (All Sales)":
            target_sale_key = self.sale_mapping[sale_selection]
            query += " AND c.sale_key = %s"
            params.append(target_sale_key)

        query += " ORDER BY c.so_number ASC"

        # 3. ดึงข้อมูล
        try:
            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))

            if df.empty:
                messagebox.showinfo("แจ้งเตือน", "ไม่พบข้อมูล SO ตามเงื่อนไขที่เลือก", parent=self)
                return

            # สร้างคอลัมน์ "สถานะ" และ "หมายเหตุ (เลื่อน)" จากคอลัมน์ internal
            status_map = {
                'PO In Progress':      'รอคิดค่าคอม',
                'Pending PU':          'รอคิดค่าคอม',
                'Pending HR Approval': 'รอ HR ตรวจสอบ',
                'Defer Requested':     'รอ Manager อนุมัติเลื่อน',
                'Deferred':            'รอคิดค่าคอม(เลื่อน)',
                'Approved':            'จ่ายแล้ว',
            }
            df['สถานะ'] = df['_status'].map(status_map).fillna(df['_status'])

            def _defer_note(row):
                m, y = row['_defer_src_m'], row['_defer_src_y']
                if pd.notna(m) and pd.notna(y):
                    try:
                        return f"เลื่อนมาจาก {thai_months_name[int(m)-1]} {int(y)+543}"
                    except Exception:
                        pass
                return ""

            df['หมายเหตุ (เลื่อน)'] = df.apply(_defer_note, axis=1)

            # ลบคอลัมน์ internal ออกก่อน export
            df = df.drop(columns=['_status', '_defer_src_m', '_defer_src_y'])

            # 4. บันทึกไฟล์
            sale_str = "All" if sale_selection == "ทั้งหมด (All Sales)" else target_sale_key
            default_filename = f"SO_Report_{year_ad}_{month_num:02d}_{sale_str}.xlsx"

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_filename,
                title="บันทึกไฟล์ Excel",
                parent=self
            )

            if not file_path: return

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='SO Data', index=False)
                worksheet = writer.sheets['SO Data']
                worksheet.column_dimensions['A'].width = 20  # SO Number
                worksheet.column_dimensions['B'].width = 35  # Customer
                worksheet.column_dimensions['C'].width = 18  # Sales Amount
                worksheet.column_dimensions['D'].width = 30  # Sale Name
                worksheet.column_dimensions['E'].width = 22  # สถานะ
                worksheet.column_dimensions['F'].width = 35  # หมายเหตุ

            self.destroy()
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลเรียบร้อยแล้ว!\n\nจำนวน: {len(df)} รายการ")

            if os.name == 'nt':
                os.startfile(file_path)

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            print(traceback.format_exc())


class SMNotificationDialog(CTkToplevel):
    """Popup แจ้งเตือน Sales Manager กรณีมีการยกเลิก SO หรือแจ้งเตือนอื่นๆ
    ต้องกด 'รับทราบ' เพื่อปิด (ไม่หายไปเอง)"""

    def __init__(self, master, app_container, notification_row):
        """
        notification_row: tuple (id, message, timestamp)
        """
        super().__init__(master)
        self.app_container = app_container
        self.noti_id = notification_row[0]
        self.noti_message = notification_row[1]
        self.noti_time = notification_row[2]

        self.title("📢 แจ้งเตือนจากระบบ")
        self.attributes("-topmost", True)
        self.resizable(False, False)
        self.grab_set()  # บล็อก input จากหน้าต่างอื่น

        self._build_ui()
        self._center_on_master(master)

    def _build_ui(self):
        # Header
        header = CTkFrame(self, fg_color="#DC2626", corner_radius=0)
        header.pack(fill="x")
        CTkLabel(
            header,
            text="⚠️  แจ้งเตือน — กรุณารับทราบ",
            font=CTkFont(size=15, weight="bold"),
            text_color="white"
        ).pack(padx=20, pady=12)

        # Body
        body = CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=16)

        # เวลา
        time_str = ""
        if self.noti_time:
            try:
                if hasattr(self.noti_time, 'strftime'):
                    time_str = self.noti_time.strftime("%d/%m/%Y %H:%M")
                else:
                    time_str = str(self.noti_time)
            except Exception:
                time_str = str(self.noti_time)

        if time_str:
            CTkLabel(
                body,
                text=f"🕐 {time_str}",
                font=CTkFont(size=12),
                text_color=("gray50", "gray60")
            ).pack(anchor="w", pady=(0, 8))

        # ข้อความ
        msg_frame = CTkFrame(body, fg_color=("gray92", "gray20"), corner_radius=8)
        msg_frame.pack(fill="x", pady=(0, 16))
        CTkLabel(
            msg_frame,
            text=self.noti_message,
            font=CTkFont(size=14),
            wraplength=380,
            justify="left"
        ).pack(padx=16, pady=14, anchor="w")

        # ปุ่มรับทราบ
        CTkButton(
            body,
            text="✅  รับทราบ",
            font=CTkFont(size=14, weight="bold"),
            fg_color="#16A34A",
            hover_color="#15803D",
            height=40,
            command=self._acknowledge
        ).pack(fill="x")

    def _center_on_master(self, master):
        w, h = 440, 300
        self.geometry(f"{w}x{h}")
        self.update_idletasks()
        try:
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            mw = master.winfo_width()
            mh = master.winfo_height()
            x = mx + (mw - w) // 2
            y = my + (mh - h) // 2
        except Exception:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = (sw - w) // 2
            y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _acknowledge(self):
        """Mark notification as read in DB then close"""
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute(
                    "UPDATE notifications SET is_read = TRUE WHERE id = %s",
                    (int(self.noti_id),)
                )
            conn.commit()
        except Exception as e:
            print(f"SMNotificationDialog: ไม่สามารถอัปเดต is_read: {e}")
            # ถ้า UPDATE ล้มเหลว ก็ยังปิด dialog ได้
        self.grab_release()


# =============================================================================
# ✅ SM Approval Confirm Dialog — popup รายละเอียดก่อนอนุมัติ SO
# =============================================================================

class SMApproveConfirmDialog(CTkToplevel):
    """Popup รายละเอียดครบก่อนกด อนุมัติ SO"""

    THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

    def __init__(self, master, so_data: dict):
        super().__init__(master)
        self.so_data = so_data
        self.confirmed = False

        self.title("ยืนยันการอนุมัติ SO")
        self.resizable(False, False)
        self.grab_set()
        self.focus()
        self._build()
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _build(self):
        d = self.so_data
        pad = {"padx": 24, "pady": 5}

        # ── Header ──────────────────────────────────────────────
        hdr = CTkFrame(self, fg_color="#1D4ED8", corner_radius=0)
        hdr.pack(fill="x")
        CTkLabel(hdr, text="🗳️  ยืนยันการอนุมัติ SO",
                 font=CTkFont(size=16, weight="bold"),
                 text_color="white").pack(anchor="w", padx=20, pady=12)

        # ── Body ────────────────────────────────────────────────
        body = CTkFrame(self, fg_color="white", corner_radius=0)
        body.pack(fill="both", expand=True)

        def row(label, value, value_color="#111827"):
            f = CTkFrame(body, fg_color="transparent")
            f.pack(fill="x", padx=24, pady=4)
            CTkLabel(f, text=label, font=CTkFont(size=13),
                     text_color="#6B7280", width=150, anchor="w").pack(side="left")
            CTkLabel(f, text=str(value), font=CTkFont(size=13, weight="bold"),
                     text_color=value_color, anchor="w").pack(side="left")

        CTkFrame(body, height=1, fg_color="#E5E7EB").pack(fill="x", padx=24, pady=(14, 8))

        row("เลข SO :", d.get("so_number", "-"))
        row("ลูกค้า :", d.get("customer_name", "-"))
        row("เซลล์ :", f"{d.get('sale_name', '-')}  ({d.get('sale_key', '-')})")

        # ยอดขาย
        try:
            amt = float(d.get("sales_service_amount") or 0)
            row("ยอดขาย :", f"{amt:,.2f}  บาท", "#1D4ED8")
        except Exception:
            row("ยอดขาย :", str(d.get("sales_service_amount", "-")))

        # ค่าจัดส่ง
        try:
            sc = float(d.get("shipping_cost") or 0)
            row("ค่าจัดส่ง :", f"{sc:,.2f}  บาท")
        except Exception:
            pass

        # วันที่บิล
        row("วันที่บิล :", d.get("bill_date", "-"))
        # วันที่จัดส่ง
        row("วันที่จัดส่ง :", d.get("delivery_date", "-"))

        # รอบเดือนค่าคอม
        try:
            m = int(d.get("commission_month") or 0)
            y = int(d.get("commission_year") or 0) + 543
            com_str = f"{self.THAI_MONTHS[m-1]}  {y}"
        except Exception:
            com_str = f"{d.get('commission_month', '-')} / {d.get('commission_year', '-')}"
        row("รอบค่าคอม :", com_str, "#7C3AED")

        CTkFrame(body, height=1, fg_color="#E5E7EB").pack(fill="x", padx=24, pady=(10, 6))

        # คำเตือน
        CTkLabel(body,
                 text="⚠️  เมื่ออนุมัติแล้ว SO จะเข้าสู่ขั้นตอน PU ต่อไป",
                 font=CTkFont(size=12), text_color="#D97706").pack(padx=24, pady=(0, 14))

        # ── Buttons ─────────────────────────────────────────────
        btn_row = CTkFrame(self, fg_color=("gray95", "gray15"), corner_radius=0)
        btn_row.pack(fill="x")

        CTkButton(btn_row, text="✅  ยืนยันอนุมัติ",
                  fg_color="#16A34A", hover_color="#15803D",
                  font=CTkFont(size=14, weight="bold"), height=42, width=160,
                  command=self._confirm).pack(side="left", padx=20, pady=12)

        CTkButton(btn_row, text="ยกเลิก",
                  fg_color="gray", hover_color="#555555",
                  font=CTkFont(size=14), height=42, width=100,
                  command=self.destroy).pack(side="left", padx=(0, 20), pady=12)

    def _confirm(self):
        self.confirmed = True
        self.destroy()
        self.destroy()