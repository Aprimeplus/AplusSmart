# sales_target_widget.py
# SalesTargetWidget — แท็บ "เป้าการขาย" / "Sale Revenue Report"
# ใช้ร่วมกันทั้งหน้า Sales Manager (sales_manager_screen.py) และหน้า
# รายงานผู้บริหาร (management_report_screen.py) — เดิมมี 2 ก้อนโค้ดแยกกันเกือบเหมือนกัน
# ทำให้แก้บัคแล้วลืมแก้อีกจุด (เช่น เป้าบริษัทรายปี, การรวม Inactive Sale) จึงรวมเป็นจุดเดียวที่นี่

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

from hr_screen import SalesFilterDialog
from export_utils import export_sales_target_to_excel
from daily_report_dashboard import TargetSettingsDialog


class SalesTargetWidget(CTkFrame):
    THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    THAI_MONTH_MAP = {m: i+1 for i, m in enumerate(THAI_MONTHS)}
    EXCLUDE_KEYS   = {'s','d','p','mp','ms','hr','sm','Pimhathai','CHARITA-CT'}
    SALE_CENTER_KEY = 'Sale Center'
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

        self._first_map_done  = False
        self._frame_ready_job = None
        self._ready_frame_w   = 0
        self._ready_frame_h   = 0
        self._drawing_chart   = False   # True ระหว่าง _update_dashboard — suppress Configure
        self._last_chart_fw   = 0       # ขนาดล่าสุดที่ process แล้ว — กัน same-size loop
        self._last_chart_fh   = 0
        self._build_ui()
        # bind <Map> เพื่อ draw ครั้งแรกเมื่อ tab ถูกเปิด
        self.bind('<Map>', self._on_first_map)
        self.bind('<Destroy>', self._on_destroy)

    def _on_destroy(self, event=None):
        """กัน after() debounce ที่ค้างคิวอยู่ไปเรียก widget ที่ถูก destroy() ไปแล้วตอนสลับหน้า —
        ไม่งั้น TclError ที่เกิดขึ้นตอนนั้นอาจทำให้แอปค้าง"""
        if event is not None and event.widget is not self:
            return
        if self._frame_ready_job:
            try:
                self.after_cancel(self._frame_ready_job)
            except Exception:
                pass
            self._frame_ready_job = None

    def _on_first_map(self, event=None):
        """ทำงานครั้งเดียวตอน tab ถูกเปิด — bind Configure บน _chart_frame แบบถาวร
        เพื่อรับทั้ง initial draw และทุกครั้งที่ window ถูก resize/maximize"""
        if self._first_map_done:
            return
        self._first_map_done = True
        self.unbind('<Map>')
        # Bind ถาวร — ไม่ unbind หลัง initial draw เพื่อให้รับ resize ได้ตลอด
        self._chart_frame.bind('<Configure>', self._on_chart_frame_configure)

    def _on_chart_frame_configure(self, event):
        """เรียกทุกครั้งที่ _chart_frame เปลี่ยนขนาด (initial layout + window resize)
        debounce 200ms รอให้ layout settle แล้วค่อย redraw"""
        if not self.winfo_exists():
            return
        if event.width < 200 or event.height < 100:
            return
        # Layer 1: suppress ระหว่าง _update_dashboard รัน (ป้องกัน update_idletasks loop)
        if self._drawing_chart:
            return
        fw, fh = event.width, event.height
        # Layer 2: ถ้าขนาดเดิม ไม่ต้อง redraw (ป้องกัน canvas.draw() → Configure → loop)
        if fw == self._last_chart_fw and fh == self._last_chart_fh:
            return
        if self._frame_ready_job:
            self.after_cancel(self._frame_ready_job)
        self._frame_ready_job = self.after(200, lambda: self._on_chart_resize(fw, fh))

    def _on_chart_resize(self, fw, fh):
        """Fired after debounce — initial draw หรือ window-resize redraw"""
        self._frame_ready_job = None
        if not self.winfo_exists():
            return
        if self.sales_view_mode not in ('chart', 'monthly'):
            return  # ไม่ต้อง resize ถ้าอยู่ใน table mode
        if self.sales_target_chart_canvas is None:
            # ยังไม่มี chart — วาด initial draw พร้อม fetch data
            self._ready_frame_w = fw
            self._ready_frame_h = fh
            self._update_dashboard()
        else:
            # มี chart อยู่แล้ว — resize figure โดยตรง ไม่ต้อง fetch data ใหม่
            # ใช้ winfo_width() ณ ตอนนี้แทน fw/fh ที่ capture ไว้
            try:
                actual_fw = self._chart_frame.winfo_width()
                actual_fh = self._chart_frame.winfo_height()
                real_fw = actual_fw if actual_fw > 100 else fw
                real_fh = actual_fh if actual_fh > 100 else fh
                # stamp ก่อน draw — Configure ที่ size นี้หลัง canvas.draw() จะถูก ignore
                self._last_chart_fw = real_fw
                self._last_chart_fh = real_fh
                cw = max(real_fw - 22, 100)
                ch = max(real_fh - 22, 100)
                fig = self.sales_target_chart_canvas.figure
                fig.set_size_inches(cw / fig.dpi, ch / fig.dpi)
                if self.sales_view_mode == 'monthly':
                    fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.22)
                else:
                    fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.12)
                self.after(10, self.sales_target_chart_canvas.draw)
            except Exception:
                traceback.print_exc()

    def _resize_monthly_after_settle(self):
        """Resize monthly chart หลัง layout settle (แก้ปัญหา maximized window)"""
        if self.sales_view_mode != 'monthly' or self.sales_target_chart_canvas is None:
            return
        try:
            fw = self._chart_frame.winfo_width()
            fh = self._chart_frame.winfo_height()
            if fw < 100 or fh < 100:
                return
            if fw == self._last_chart_fw and fh == self._last_chart_fh:
                return
            self._last_chart_fw = fw
            self._last_chart_fh = fh
            fig = self.sales_target_chart_canvas.figure
            dpi = fig.dpi
            fig.set_size_inches(max(fw - 22, 100) / dpi, max(fh - 22, 100) / dpi)
            fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.28)
            self.sales_target_chart_canvas.draw()
        except Exception:
            traceback.print_exc()

    def _force_monthly_resize(self):
        """Force resize monthly chart หลัง layout settle — ไม่มี skip-check"""
        if self.sales_view_mode != 'monthly' or self.sales_target_chart_canvas is None:
            return
        try:
            fw = self._chart_frame.winfo_width()
            fh = self._chart_frame.winfo_height()
            if fw < 100 or fh < 100:
                self.after(200, self._force_monthly_resize)
                return
            fig = self.sales_target_chart_canvas.figure
            dpi = fig.dpi
            cw = max(fw - 22, 100)
            ch = max(fh - 22, 100)
            fig.set_size_inches(cw / dpi, ch / dpi)
            fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.28)
            self.sales_target_chart_canvas.draw()
            self._last_chart_fw = fw
            self._last_chart_fh = fh
            print(f"[force_monthly_resize] resized to {cw}x{ch}")
        except Exception:
            traceback.print_exc()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        today = datetime.now()
        current_year = today.year
        years = [str(y + 543) for y in range(2020, current_year + 3)]

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

        CTkButton(fbar, text="⚙️ ตั้งเป้าหมายรายปี", width=150, fg_color="#F59E0B", hover_color="#D97706",
                  command=self._open_yearly_target_settings).pack(side="left", padx=(0, 10))

        CTkButton(fbar, text="⚙️ ตั้งเส้นคุ้มทุน / เป้า GP", width=190, fg_color="#0EA5E9", hover_color="#0284C7",
                  command=self._open_margin_threshold_settings).pack(side="left", padx=(0, 10))

        # Toggle กราฟ / รายเดือน / ตาราง
        tgl = CTkFrame(fbar, fg_color="transparent")
        tgl.pack(side="right", padx=(0, 5))

        def _set_view(mode):
            self.sales_view_mode = mode
            btn_chart.configure(fg_color="#2563EB" if mode == 'chart'   else "#E2E8F0",
                                text_color="white"   if mode == 'chart'   else "#475569")
            btn_monthly.configure(fg_color="#2563EB" if mode == 'monthly' else "#E2E8F0",
                                  text_color="white"  if mode == 'monthly' else "#475569")
            btn_table.configure(fg_color="#2563EB" if mode == 'table'   else "#E2E8F0",
                                text_color="white"   if mode == 'table'   else "#475569")
            self._update_dashboard()

        btn_chart = CTkButton(tgl, text="📊 กราฟเป้า", width=100,
                              fg_color="#2563EB", text_color="white",
                              corner_radius=6, command=lambda: _set_view('chart'))
        btn_chart.pack(side="left", padx=2)
        btn_monthly = CTkButton(tgl, text="📈 รายเดือน", width=100,
                                fg_color="#E2E8F0", text_color="#475569",
                                corner_radius=6, command=lambda: _set_view('monthly'))
        btn_monthly.pack(side="left", padx=2)
        btn_table = CTkButton(tgl, text="📋 ตาราง", width=90,
                              fg_color="#E2E8F0", text_color="#475569",
                              corner_radius=6, command=lambda: _set_view('table'))
        btn_table.pack(side="left", padx=2)

        CTkButton(tgl, text="📥 Export Excel", width=120,
                  fg_color="#16A34A", hover_color="#15803D", text_color="white",
                  corner_radius=6, command=self._export_sales_target_excel
                  ).pack(side="left", padx=(10, 2))

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

    def _open_yearly_target_settings(self):
        """เปิด dialog ตั้งเป้าหมายรายปี (ตัวเดียวกับหน้า Daily Report) — แก้ตาราง
        sales_yearly_targets ตรงๆ จากหน้านี้ได้เลย ไม่ต้องไปสลับหน้าไปหน้า Daily Report"""
        try:
            year = int(self.end_y_var.get()) - 543
        except Exception:
            year = datetime.now().year
        TargetSettingsDialog(self, self.app_container, year, on_save_callback=self._on_search)

    def _update_dashboard(self):
        # ยกเลิก resize callback ที่ค้างอยู่ก่อน — ป้องกัน stale resize override fresh draw
        if self._frame_ready_job:
            self.after_cancel(self._frame_ready_job)
            self._frame_ready_job = None
        # Capture frame size ก่อน _show_loading() เปลี่ยน layout (ค่าถูกต้องที่สุด ณ จุดนี้)
        # ใช้สำหรับ _create_chart ที่เรียกทีหลัง — ไม่ต้องพึ่ง winfo_width() หลัง update_idletasks
        pre_fw = self._chart_frame.winfo_width()
        pre_fh = self._chart_frame.winfo_height()
        if pre_fw > 100:
            self._ready_frame_w = pre_fw
            self._ready_frame_h = pre_fh
        self._drawing_chart = True   # suppress _on_chart_frame_configure ระหว่าง draw
        loading = self._show_loading()
        try:
            df = self._get_data()
            loading.destroy()
            if self.sales_view_mode == 'table':
                self._create_table(self._chart_frame, df)
            elif self.sales_view_mode == 'monthly':
                monthly_df = self._get_monthly_data()
                self._create_monthly_chart(self._chart_frame, monthly_df)
                # force resize หลัง layout settle — ข้าม skip-check ปกติ
                self.after(400, self._force_monthly_resize)
            else:
                self._create_chart(self._chart_frame, df)
        except Exception as e:
            loading.destroy()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            self._drawing_chart = False
            # stamp ขนาด frame ปัจจุบัน — Configure event ที่ size นี้จะถูก ignore
            fw = self._chart_frame.winfo_width()
            fh = self._chart_frame.winfo_height()
            if fw > 100:
                self._last_chart_fw = fw
                self._last_chart_fh = fh

    # ── Data fetch ────────────────────────────────────────────────────────────
    HISTORY_CUTOFF_YEAR = 2025  # ปีที่ระบบเริ่มมีข้อมูลจริง (ก่อนหน้านี้ใช้ sales_history_monthly)

    def _get_history_totals(self):
        """รวม target/actual ย้อนหลัง (เดือนที่ยังไม่มีข้อมูลจริงในระบบ) จาก sales_history_monthly
        ตามช่วงที่เลือก คืนค่าเป็น dict: {sale_key(lower,no-space): {'target':.., 'actual':..}}
        ตารางนี้เก็บเฉพาะเดือนที่ไม่ทับกับ commission_payout_logs อยู่แล้ว จึงรวมได้ตรง ๆ ไม่ซ้ำ"""
        s = self.custom_target_start
        e = self.custom_target_end
        if s and e:
            start_key = s.year * 12 + s.month
            end_key   = e.year * 12 + e.month
        else:
            today = datetime.now()
            start_key = end_key = today.year * 12 + today.month
        try:
            hdf = pd.read_sql_query(
                """
                SELECT REPLACE(LOWER(sale_key), ' ', '') AS sale_key,
                       COALESCE(SUM(target), 0)       AS hist_target,
                       COALESCE(SUM(actual_sales), 0) AS hist_actual,
                       COALESCE(SUM(CASE WHEN margin_pct IS NOT NULL
                                         THEN margin_pct / 100 * actual_sales ELSE 0 END), 0) AS hist_gp,
                       COALESCE(SUM(CASE WHEN margin_pct IS NOT NULL
                                         THEN actual_sales ELSE 0 END), 0) AS hist_margin_base
                FROM   sales_history_monthly
                WHERE  (year * 12 + month) BETWEEN %s AND %s
                GROUP  BY sale_key
                """,
                self.pg_engine, params=(start_key, end_key)
            )
        except Exception:
            traceback.print_exc()
            return {}
        return {r['sale_key']: {'target': float(r['hist_target']), 'actual': float(r['hist_actual']),
                                 'gp': float(r['hist_gp']), 'margin_base': float(r['hist_margin_base'])}
                for _, r in hdf.iterrows()}

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
            # นับเฉพาะเดือนที่อยู่ในยุคระบบ (>= HISTORY_CUTOFF_YEAR) สำหรับคูณเป้าคงที่
            # เดือนก่อนหน้านั้นใช้ข้อมูลจาก sales_history_monthly แทน (ดู _get_history_totals)
            live_months = 0
            y, m = s.year, s.month
            while (y, m) <= (e.year, e.month):
                if y >= self.HISTORY_CUTOFF_YEAR:
                    live_months += 1
                m += 1
                if m > 12:
                    m = 1
                    y += 1
            target_mult = max(0, live_months)
        else:
            date_clauses.append("c.commission_month = %s")
            params.append(today.month)
            date_clauses.append("c.commission_year = %s")
            params.append(today.year)

        date_sql = " AND ".join(date_clauses)
        date_params = list(params)  # snapshot ก่อนเพิ่ม sale filter

        sale_filter = ""
        if self.selected_sales_filter:
            ph = ','.join(["REPLACE(LOWER(%s), ' ', '')"] * len(self.selected_sales_filter))
            sale_filter = f" AND REPLACE(LOWER(su.sale_key), ' ', '') IN ({ph})"
            params.extend(self.selected_sales_filter)

        query = f"""
            SELECT su.sale_name, su.sale_key, su.status,
                   COALESCE(su.sales_target, 0) * %s AS sales_target,
                   COALESCE(SUM(c.total_sales), 0)   AS total_sales,
                   0                                  AS total_outstanding,
                   COALESCE(SUM(cm_agg.total_gp), 0)       AS total_gp,
                   COALESCE(SUM(cm_agg.total_sales_amt), 0) AS total_sales_amt,
                   COALESCE(SUM(cm_agg.sales_normal), 0) AS sales_normal,
                   COALESCE(SUM(cm_agg.sales_below),  0) AS sales_below
            FROM   sales_users su
            LEFT JOIN commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND {date_sql}
            LEFT JOIN (
                SELECT payout_id,
                       SUM(final_gp)            AS total_gp,
                       SUM(final_sales_amount)  AS total_sales_amt,
                       SUM(CASE WHEN final_margin >= 10 THEN final_sales_amount ELSE 0 END) AS sales_normal,
                       SUM(CASE WHEN final_margin <  10 THEN final_sales_amount ELSE 0 END) AS sales_below
                FROM   commissions
                GROUP  BY payout_id
            ) cm_agg ON cm_agg.payout_id = c.id
            WHERE  1=1
                   {sale_filter}
            GROUP  BY su.sale_name, su.sale_key, su.sales_target, su.role, su.status
            -- แสดงพนักงานที่ยัง Active ทุกคน (แม้ยังไม่มียอด) + พนักงานที่ปิดใช้งานไปแล้ว (Inactive)
            -- เฉพาะกรณีที่มียอดขายจริงในช่วงที่เลือก (กันไม่ให้อดีตพนักงานที่ไม่มียอดโผล่มาเป็นแท่งว่างๆ)
            HAVING (su.status = 'Active' AND su.role = 'Sale')
                   OR COALESCE(SUM(c.total_sales), 0) > 0
            ORDER  BY su.sale_name ASC;
        """
        final_params = tuple([target_mult] + params)
        df = pd.read_sql_query(query, self.pg_engine, params=final_params)
        df['sales_target']      = df['sales_target'].fillna(0)
        df['total_sales']       = df['total_sales'].fillna(0)
        df['total_outstanding'] = df['total_outstanding'].fillna(0)
        df['total_gp']          = df['total_gp'].fillna(0)
        df['total_sales_amt']   = df['total_sales_amt'].fillna(0)
        df['sales_normal']      = df['sales_normal'].fillna(0)
        df['sales_below']       = df['sales_below'].fillna(0)

        # ── รวมข้อมูลย้อนหลัง (ก่อน HISTORY_CUTOFF_YEAR) จาก sales_history_monthly ──
        hist_totals = self._get_history_totals()
        if hist_totals:
            def _norm_key(k):
                return str(k).lower().replace(' ', '')
            df['sales_target'] = df.apply(
                lambda r: r['sales_target'] + hist_totals.get(_norm_key(r['sale_key']), {}).get('target', 0),
                axis=1)
            df['total_sales'] = df.apply(
                lambda r: r['total_sales'] + hist_totals.get(_norm_key(r['sale_key']), {}).get('actual', 0),
                axis=1)
            # margin ย้อนหลัง: แปลง margin_pct เป็น GP โดยประมาณ (margin% * ยอดขายเดือนนั้น) แล้วรวมเป็นตัวตั้ง/ตัวหารเดียวกับ GP จริง
            df['total_gp'] = df.apply(
                lambda r: r['total_gp'] + hist_totals.get(_norm_key(r['sale_key']), {}).get('gp', 0),
                axis=1)
            df['total_sales_amt'] = df.apply(
                lambda r: r['total_sales_amt'] + hist_totals.get(_norm_key(r['sale_key']), {}).get('margin_base', 0),
                axis=1)
        df['avg_margin'] = df.apply(
            lambda r: (r['total_gp'] / r['total_sales_amt'] * 100) if r['total_sales_amt'] > 0 else 0,
            axis=1)

        # ── Sale Center: ดึงจาก commissions โดยตรง (ไม่มีค่าคอม) ─────
        sc_date_filter = date_sql.replace("c.", "sc.")
        sc_query = f"""
            SELECT COALESCE(SUM(sc.sales_service_amount), 0) AS total_sales
            FROM commissions sc
            WHERE sc.sale_key IN ('Sale Center', 'CHARITA-CT')
              AND sc.is_active = 1
              AND sc.status NOT IN ('Cancelled', 'Cancelled by PU')
              AND {sc_date_filter}
        """
        try:
            sc_df = pd.read_sql_query(sc_query, self.pg_engine, params=tuple(date_params))
            sc_total = float(sc_df['total_sales'].iloc[0]) if not sc_df.empty else 0.0
        except Exception:
            sc_total = 0.0

        if sc_total > 0:
            sc_row = pd.DataFrame([{
                'sale_name': 'Sale Center', 'sale_key': 'Sale Center', 'status': 'Active',
                'sales_target': 0.0, 'total_sales': sc_total,
                'total_outstanding': 0.0, 'avg_margin': 0.0,
                'sales_normal': sc_total, 'sales_below': 0.0,
            }])
            df = pd.concat([df, sc_row], ignore_index=True)

        return df

    # ── Inactive sale merge (ตามที่ PM ขอ: ไม่โชว์ชื่อทีละคน รวมเป็นก้อนเดียว) ──────
    def _inactive_bucket_label(self):
        try:
            e = self.custom_target_end
            if e:
                return f"All Inactive Sale {e.year + 543}"
        except Exception:
            pass
        return f"All Inactive Sale {datetime.now().year + 543}"

    # ── เส้นคุ้มทุน / เป้า GP (ตั้งค่าได้ ไม่ hard code) ──────────────────────────
    MARGIN_FLOOR_DEFAULT = 15.0
    MARGIN_GOAL_DEFAULT  = 17.0

    def _get_margin_thresholds(self):
        """คืน (เส้นห้ามต่ำกว่า %, เป้าที่ต้องทำ %) จาก company_settings — ไม่ hard code เพราะทั้งสองค่า
        เปลี่ยนตามยอดขายทั้งปีและอัตราค่าใช้จ่ายของบริษัท (ค่าเริ่มต้น 15 / 17 ถ้ายังไม่เคยตั้ง)"""
        floor_v, goal_v = self.MARGIN_FLOOR_DEFAULT, self.MARGIN_GOAL_DEFAULT
        try:
            df = pd.read_sql_query(
                "SELECT setting_key, setting_value FROM company_settings "
                "WHERE setting_key IN ('sales_margin_floor_pct', 'sales_margin_goal_pct')", self.pg_engine)
            for _, r in df.iterrows():
                v = float(r['setting_value'])
                if r['setting_key'] == 'sales_margin_floor_pct':
                    floor_v = v
                else:
                    goal_v = v
        except Exception:
            traceback.print_exc()
        return floor_v, goal_v

    def _open_margin_threshold_settings(self):
        floor_v, goal_v = self._get_margin_thresholds()
        dlg = CTkToplevel(self)
        dlg.title("ตั้งเส้นคุ้มทุน / เป้า GP")
        dlg.geometry("380x230")
        dlg.transient(self.winfo_toplevel())
        dlg.after(50, dlg.grab_set)
        CTkLabel(dlg, text="เส้นห้ามต่ำกว่า (จุดคุ้มทุน) %", font=CTkFont(size=13)).pack(anchor="w", padx=24, pady=(20, 2))
        e_floor = CTkEntry(dlg, width=120); e_floor.insert(0, f"{floor_v:g}"); e_floor.pack(anchor="w", padx=24)
        CTkLabel(dlg, text="เป้าที่ต้องทำ %", font=CTkFont(size=13)).pack(anchor="w", padx=24, pady=(12, 2))
        e_goal = CTkEntry(dlg, width=120); e_goal.insert(0, f"{goal_v:g}"); e_goal.pack(anchor="w", padx=24)

        def _save():
            try:
                fv = float(e_floor.get().replace(',', '')); gv = float(e_goal.get().replace(',', ''))
                if fv <= 0 or gv < fv:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("แจ้งเตือน", "กรอกตัวเลขให้ถูกต้อง และเป้าต้องไม่ต่ำกว่าเส้นคุ้มทุน", parent=dlg)
                return
            conn = self.app_container.get_connection()
            try:
                with conn.cursor() as cur:
                    for k, v in (('sales_margin_floor_pct', fv), ('sales_margin_goal_pct', gv)):
                        cur.execute(
                            "INSERT INTO company_settings (setting_key, setting_value) VALUES (%s, %s) "
                            "ON CONFLICT (setting_key) DO UPDATE SET setting_value = EXCLUDED.setting_value",
                            (k, str(v)))
                conn.commit()
            except Exception as ex:
                conn.rollback()
                messagebox.showerror("Error", f"บันทึกไม่สำเร็จ: {ex}", parent=dlg)
                return
            finally:
                self.app_container.release_connection(conn)
            dlg.destroy()
            self._on_search()

        CTkButton(dlg, text="บันทึก", width=100, command=_save).pack(anchor="w", padx=24, pady=18)

    def _get_company_annual_target(self, start, end):
        """เป้าบริษัทรวม — ดึงจากตาราง sales_yearly_targets (ที่ PM ตั้งเป็นรายปีไว้แล้ว
        ผ่านหน้า Daily Report) แทนการบวกเป้าของพนักงานแต่ละคนรวมกัน
        เฉลี่ยตามสัดส่วนเดือนที่อยู่ในช่วงที่เลือก (เช่น เลือก 7 เดือน จาก 12 เดือน = เป้าปี/12*7)
        คืนค่า 0.0 ถ้ายังไม่เคยตั้งเป้าปีนั้นไว้ — ให้ผู้เรียก fallback ไปใช้วิธีเดิม (บวกเป้ารายคน)"""
        if not start or not end:
            return 0.0
        try:
            years = list(range(start.year, end.year + 1))
            df = pd.read_sql_query(
                "SELECT year, target_amount FROM sales_yearly_targets WHERE year = ANY(%s)",
                self.pg_engine, params=(years,))
            target_by_year = {int(r['year']): float(r['target_amount'] or 0) for _, r in df.iterrows()}
        except Exception:
            traceback.print_exc()
            return 0.0

        total = 0.0
        y, m = start.year, start.month
        while (y, m) <= (end.year, end.month):
            total += target_by_year.get(y, 0.0) / 12
            m += 1
            if m > 12:
                m = 1
                y += 1
        return total

    def _merge_inactive_group(self, df2, bucket_label=None):
        """df2 ต้องมีคอลัมน์ status + _group (assign ไว้แล้วตาม PERSON_MERGE) —
        เปลี่ยน _group ของทุกคนที่ status != 'Active' (ยกเว้น Sale Center) ให้เป็นก้อนเดียวกัน
        เพื่อไม่ให้โชว์ชื่อพนักงานที่ลาออกไปแล้วทีละคน
        bucket_label: ระบุ label ตายตัวได้ (กันกรณี loop เปลี่ยน custom_target_end ทีละเดือน
        ตอน export แล้วปีคาบเกี่ยว จะได้ไม่ได้ label คนละอันในแต่ละเดือน)"""
        if 'status' not in df2.columns or df2.empty:
            return df2
        is_inactive = (df2['status'] != 'Active') & (df2['sale_key'] != self.SALE_CENTER_KEY)
        if not is_inactive.any():
            return df2
        df2 = df2.copy()
        bucket = bucket_label or self._inactive_bucket_label()
        df2.loc[is_inactive, '_group'] = bucket
        if '_seg_label' in df2.columns:
            df2.loc[is_inactive, '_seg_label'] = df2.loc[is_inactive, 'sale_name']
        return df2

    # ── Chart ─────────────────────────────────────────────────────────────────
    def _create_chart(self, parent_frame, data_df):
        _dpi = 100
        # ใช้ขนาดที่เก็บจาก Configure event (ถูกต้องที่สุด)
        # ถ้าไม่มี (เช่น user กดค้นหา) ให้ fallback ไปที่ winfo_width
        _fw = getattr(self, '_ready_frame_w', 0) or parent_frame.winfo_width()
        _fh = getattr(self, '_ready_frame_h', 0) or parent_frame.winfo_height()
        self._ready_frame_w = 0  # clear หลังใช้
        self._ready_frame_h = 0
        if _fw <= 10:
            _fw = self.winfo_width() - 20
        if _fh <= 10:
            _fh = self.winfo_height() - 80

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
        df2 = self._merge_inactive_group(df2)

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

        SC_COLOR = ('#0EA5E9', '#7DD3FC')
        pct_labels = []
        for p, s, t in zip(people_data, sales, targets):
            is_sc = any(sub['sale_key'] == self.SALE_CENTER_KEY for sub in p['sub_items'])
            if is_sc:
                pct_labels.append("")
            elif t > 0:
                pct_labels.append(f"{s/t*100:.0f}%")
            else:
                pct_labels.append("N/A")

        BG = '#F8FAFC'; GRID_C = '#E2E8F0'
        # สร้าง Figure ด้วยขนาดจริงของ frame (อ่านไว้แล้วตอนต้น)
        _pad  = 24
        fig_w = (_fw - _pad) / _dpi if _fw > 100 else max(10.0, n * 1.6)
        fig_h = (_fh - _pad) / _dpi if _fh > 100 else 7.2
        fig = Figure(figsize=(fig_w, fig_h), dpi=_dpi, facecolor=BG)
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
            is_sc = any(s['sale_key'] == self.SALE_CENTER_KEY for s in p['sub_items'])
            colors = SC_COLOR if is_sc else ACHIEVE_COLORS[achievement_key(total_s, target)]
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
            elif pct:
                pct_y = max(s, t) + max_t * 0.16
                ax.text(x[i], pct_y, pct, ha='center', va='bottom', fontsize=16,
                        weight='medium', color='black', zorder=8,
                        path_effects=[pe.withStroke(linewidth=2, foreground='white')])
                if s <= max_t * 0.08:
                    ax.text(x[i], s + max_t*0.012, f"{s:,.0f}",
                            ha='center', va='bottom', fontsize=12, weight='bold', color='black', zorder=7)
            elif s > 0:
                # Sale Center (ไม่มีเป้า) — แสดงยอดบนหัวแท่ง
                pct_y = s + max_t * 0.16
                ax.text(x[i], pct_y, f"{s:,.0f}",
                        ha='center', va='bottom', fontsize=14,
                        weight='bold', color='#0369A1', zorder=8,
                        path_effects=[pe.withStroke(linewidth=2, foreground='white')])

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
        # 🟢 [ตามที่ PM ขอ] ใช้เป้าบริษัทรายปีที่ตั้งไว้แล้ว (sales_yearly_targets) แทนการบวกเป้ารายคน
        # ถ้ายังไม่เคยตั้งเป้าปีนั้นไว้เลย ค่อย fallback ไปบวกเป้ารายคนแบบเดิม
        company_target = self._get_company_annual_target(self.custom_target_start, self.custom_target_end)
        team_target = company_target if company_target > 0 else sum(p['target'] for p in people_data)
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
        if any(any(s['sale_key'] == self.SALE_CENTER_KEY for s in p['sub_items']) for p in people_data):
            legend_items.append(Patch(facecolor='#0EA5E9', label='ยอดบริษัท (Sale Center)'))
        ax.legend(handles=legend_items, loc='upper right', bbox_to_anchor=(1.0, 1.0),
                  ncol=2, frameon=True, framealpha=0.95, edgecolor='#CBD5E1', fontsize=10,
                  prop={'weight': 'bold', 'size': 10}, borderpad=0.7, labelspacing=0.4, columnspacing=1.0)

        fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.12)

        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        widget = canvas.get_tk_widget()
        widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Pack ก่อน แล้วถาม actual size หลัง layout settle
        # วิธีนี้ guarantee ว่า figure = พื้นที่จริงที่ pack จัดให้ เสมอ ไม่ว่าจะ initial หรือ search
        self.update_idletasks()
        _cw = widget.winfo_width()
        _ch = widget.winfo_height()
        if _cw > 50 and _ch > 50:
            fig.set_size_inches(_cw / fig.dpi, _ch / fig.dpi)
            fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.12)

        canvas.draw()
        self.sales_target_chart_canvas = canvas
        # Resize จัดการโดย _on_chart_frame_configure ที่ bind ไว้บน _chart_frame แบบถาวร

    # ── Monthly Revenue Chart ─────────────────────────────────────────────────
    def _get_monthly_data(self):
        """ดึงข้อมูล total_sales แยกรายเดือน รายคน ตาม filter ที่เลือก"""
        params = []
        date_clauses = []
        s = self.custom_target_start
        e = self.custom_target_end
        if s and e:
            date_clauses.append(
                "MAKE_DATE(c.commission_year, c.commission_month, 1) BETWEEN %s::date AND %s::date"
            )
            params.extend([s.strftime("%Y-%m-%d"), e.strftime("%Y-%m-%d")])
        else:
            from datetime import datetime as _dt
            today = _dt.now()
            date_clauses.append("c.commission_year = %s")
            params.append(today.year)

        date_sql = " AND ".join(date_clauses)
        date_params_only = list(params)  # snapshot ก่อน sale filter
        sale_filter = ""
        if self.selected_sales_filter:
            ph = ','.join(["REPLACE(LOWER(%s), ' ', '')"] * len(self.selected_sales_filter))
            sale_filter = f" AND REPLACE(LOWER(su.sale_key), ' ', '') IN ({ph})"
            params.extend(self.selected_sales_filter)

        query = f"""
            SELECT su.sale_name, su.sale_key,
                   c.commission_month AS month,
                   c.commission_year  AS year,
                   COALESCE(SUM(c.total_sales), 0) AS total_sales
            FROM   sales_users su
            JOIN   commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND {date_sql}
            WHERE  su.status = 'Active'
                   AND su.role = 'Sale'
                   {sale_filter}
            GROUP  BY su.sale_name, su.sale_key, c.commission_month, c.commission_year
            ORDER  BY c.commission_year, c.commission_month, su.sale_name
        """
        df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
        df['total_sales'] = df['total_sales'].fillna(0)

        # ── Sale Center: ดึงรายเดือนจาก commissions โดยตรง ───────────
        sc_date_filter = date_sql.replace("c.", "sc.")
        sc_monthly_query = f"""
            SELECT 'Sale Center' AS sale_name, 'Sale Center' AS sale_key,
                   sc.commission_month AS month, sc.commission_year AS year,
                   COALESCE(SUM(sc.sales_service_amount), 0) AS total_sales
            FROM commissions sc
            WHERE sc.sale_key IN ('Sale Center', 'CHARITA-CT')
              AND sc.is_active = 1
              AND sc.status NOT IN ('Cancelled', 'Cancelled by PU')
              AND {sc_date_filter}
            GROUP BY sc.commission_month, sc.commission_year
            ORDER BY sc.commission_year, sc.commission_month
        """
        try:
            sc_monthly_df = pd.read_sql_query(sc_monthly_query, self.pg_engine,
                                               params=tuple(date_params_only))
            sc_monthly_df['total_sales'] = sc_monthly_df['total_sales'].fillna(0)
            sc_monthly_df = sc_monthly_df[sc_monthly_df['total_sales'] > 0]
            if not sc_monthly_df.empty:
                df = pd.concat([df, sc_monthly_df], ignore_index=True)
        except Exception as e:
            print(f"Sale Center monthly query error: {e}")

        # ── ข้อมูลย้อนหลัง / เดือนที่ยังไม่มีข้อมูลจริง จาก sales_history_monthly ──
        # ตารางนี้เก็บเฉพาะเดือนที่ไม่ทับกับ commission_payout_logs อยู่แล้ว จึงรวมได้ตรง ๆ ไม่ซ้ำ
        if s and e:
            hist_start, hist_end = s.year * 12 + s.month, e.year * 12 + e.month
        else:
            hist_start = hist_end = today.year * 12 + today.month
        try:
            hist_query = f"""
                SELECT su.sale_name, su.sale_key, h.month, h.year,
                       h.actual_sales AS total_sales
                FROM   sales_history_monthly h
                JOIN   sales_users su
                       ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(h.sale_key), ' ', '')
                WHERE  (h.year * 12 + h.month) BETWEEN %s AND %s
                       {sale_filter}
            """
            hist_df = pd.read_sql_query(hist_query, self.pg_engine,
                                         params=tuple([hist_start, hist_end]
                                                      + (self.selected_sales_filter or [])))
            hist_df['total_sales'] = hist_df['total_sales'].fillna(0)
            if not hist_df.empty:
                df = pd.concat([df, hist_df], ignore_index=True)
        except Exception as e:
            print(f"History monthly query error: {e}")

        return df

    def _create_monthly_chart(self, parent_frame, data_df):
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.ticker as mticker
        import matplotlib.patches as mpatches
        import numpy as np
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

        # อ่านขนาดก่อน destroy เหมือน _create_chart
        _fw = getattr(self, '_ready_frame_w', 0) or parent_frame.winfo_width()
        _fh = getattr(self, '_ready_frame_h', 0) or parent_frame.winfo_height()
        self._ready_frame_w = 0
        self._ready_frame_h = 0
        if _fw <= 10: _fw = self.winfo_width() - 20
        if _fh <= 10: _fh = max(400, self.winfo_height() - 200)

        for w in parent_frame.winfo_children():
            w.destroy()

        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูล", font=self._font_hdr).pack(expand=True)
            return

        EXCLUDE = self.EXCLUDE_KEYS
        MERGE   = self.PERSON_MERGE
        data_df = data_df[~data_df['sale_key'].isin(EXCLUDE)].copy()

        data_df['_group'] = data_df.apply(
            lambda r: MERGE[r['sale_key']][0] if r['sale_key'] in MERGE else r['sale_name'], axis=1)
        data_df = data_df.groupby(['_group', 'year', 'month'], as_index=False)['total_sales'].sum()
        data_df.rename(columns={'_group': 'sale_name'}, inplace=True)

        sales_list   = sorted(data_df['sale_name'].unique())
        months_all   = sorted(data_df[['year', 'month']].drop_duplicates().values.tolist())

        THAI_MONTHS_SHORT = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                             "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
        QTR_NAME = {1:'Q1',2:'Q1',3:'Q1',4:'Q2',5:'Q2',6:'Q2',
                    7:'Q3',8:'Q3',9:'Q3',10:'Q4',11:'Q4',12:'Q4'}
        # สีประจำตัว — fallback palette ถ้าไม่มีชื่อใน map
        PERSON_COLOR_MAP = {
            'กนกพร':    '#F9ED69',   # เหลือง
            'ปิยะวรรณ': '#F97316',   # ส้ม
            'ภาณุพงศ์': '#8B5CF6',   # ม่วง
            'ฐรินทร์ญา':'#8B5CF6',   # ม่วง (sub)
            'ไอยลดา':   '#38BDF8',   # ฟ้า
        }
        COLORS_FALLBACK = ['#3B82F6','#22C55E','#EC4899','#14B8A6','#EF4444','#84CC16']

        ytd_total = data_df['total_sales'].sum()
        PERSON_GAP = 1.2  # gap between persons (in x-units)

        # ── Build x-positions: each person gets their own month sequence ─────
        bar_xs       = []   # x for each bar
        bar_vals     = []
        bar_colors   = []
        bar_labels   = []   # month label (ม.ค. etc.)
        person_spans = []   # (sale_name, color, x_start, x_end)
        qtr_spans    = []   # (qtr_label, x_start, x_end)

        cur_x = 0
        for i, sale in enumerate(sales_list):
            p_df = data_df[data_df['sale_name'] == sale]
            p_months = sorted(p_df[['year', 'month']].drop_duplicates().values.tolist())
            if not p_months:
                continue
            color = next((v for k, v in PERSON_COLOR_MAP.items() if k in sale), COLORS_FALLBACK[i % len(COLORS_FALLBACK)])
            p_start = cur_x
            prev_qtr = None
            q_start  = cur_x

            for ym in p_months:
                yr, mo = int(ym[0]), int(ym[1])
                row = p_df[(p_df['year'] == yr) & (p_df['month'] == mo)]
                val = float(row['total_sales'].sum()) if not row.empty else 0.0
                bar_xs.append(cur_x)
                bar_vals.append(val)
                bar_colors.append(color)
                bar_labels.append(THAI_MONTHS_SHORT[mo - 1])

                qtr = QTR_NAME[mo]
                if prev_qtr is not None and qtr != prev_qtr:
                    qtr_spans.append((prev_qtr, q_start, cur_x - 1))
                    q_start = cur_x
                prev_qtr = qtr
                cur_x += 1

            if prev_qtr is not None:
                qtr_spans.append((prev_qtr, q_start, cur_x - 1))
            person_spans.append((sale, color, p_start, cur_x - 1))
            cur_x += PERSON_GAP

        if not bar_xs:
            CTkLabel(parent_frame, text="ไม่พบข้อมูล", font=self._font_hdr).pack(expand=True)
            return

        dpi = 100
        BOTTOM_MARGIN = 0.28
        fig, ax = plt.subplots(figsize=(_fw / dpi, _fh / dpi), dpi=dpi)
        fig.patch.set_facecolor('#F8FAFC')
        ax.set_facecolor('#F8FAFC')

        # ── Draw bars ─────────────────────────────────────────────────────────
        max_val = max(bar_vals) if bar_vals else 1
        bars = ax.bar(bar_xs, bar_vals, width=0.75, color=bar_colors, zorder=3,
                      edgecolor='white', linewidth=0.4)
        for bx, bv in zip(bar_xs, bar_vals):
            if bv > 0:
                ax.text(bx, bv + max_val * 0.012,
                        f"{bv/1e6:.2f}M", ha='center', va='bottom',
                        fontsize=7, color='#1E293B', fontweight='bold')

        # ── X-axis: month labels ──────────────────────────────────────────────
        ax.set_xticks(bar_xs)
        ax.set_xticklabels(bar_labels, fontsize=7.5, color='#475569')
        ax.tick_params(axis='x', length=0, pad=2)

        # ── Qtr + Person labels using blended transform (clip_on=False) ────────
        from matplotlib.transforms import blended_transform_factory
        trans = blended_transform_factory(ax.transData, ax.transAxes)

        for idx, (qtr_lbl, qs, qe) in enumerate(qtr_spans):
            mid = (qs + qe) / 2
            ax.text(mid, -0.10, qtr_lbl, transform=trans,
                    ha='center', va='top', fontsize=8, fontweight='bold',
                    color='#334155', clip_on=False)
            # bracket line
            ax.plot([qs - 0.4, qe + 0.4], [-0.07, -0.07],
                    transform=trans, color='#94A3B8', lw=0.8,
                    clip_on=False, solid_capstyle='butt')
            # Q separator within person
            if idx > 0:
                prev_end = qtr_spans[idx - 1][2]
                if qs - prev_end < PERSON_GAP:
                    sep_x = (prev_end + qs) / 2
                    ax.axvline(sep_x, color='#94A3B8', lw=1.0, ls='--', zorder=1, alpha=0.5)

        for sale, color, ps, pe in person_spans:
            mid = (ps + pe) / 2
            ax.text(mid, -0.22, sale, transform=trans,
                    ha='center', va='top', fontsize=9, fontweight='bold',
                    color=color, clip_on=False)
            # person separator
            sep = pe + PERSON_GAP / 2
            ax.axvline(sep, color='#CBD5E1', lw=1.0, ls='--', zorder=1)

        # ── Y-axis formatting ─────────────────────────────────────────────────
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

        # ── Title + YTD ───────────────────────────────────────────────────────
        fig.text(0.5, 0.97, f"{ytd_total:,.2f}",
                 ha='center', va='top', fontsize=20, fontweight='bold', color='#EF4444')
        fig.text(0.5, 0.925, "Sale Rev. YTD  —  ยอดขายสุทธิ แยกตามเดือน (จัดกลุ่มตามพนักงาน)",
                 ha='center', va='top', fontsize=9.5, color='#64748B')

        fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=BOTTOM_MARGIN)

        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        widget = canvas.get_tk_widget()
        widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Bind Configure บน canvas widget — รับ actual size จาก layout engine โดยตรง
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
        self.sales_target_chart_canvas = canvas
        plt.close(fig)

    # ── Table ─────────────────────────────────────────────────────────────────
    def _create_table(self, parent_frame, data_df):
        for w in parent_frame.winfo_children():
            w.destroy()
        if data_df.empty:
            CTkLabel(parent_frame, text="ไม่พบข้อมูล", font=self._font_hdr).pack(expand=True)
            return

        EXCLUDE = self.EXCLUDE_KEYS; MERGE = self.PERSON_MERGE
        MARGIN_TARGET = 15.0
        df2 = data_df[~data_df['sale_key'].isin(EXCLUDE)].copy()
        df2['_group'] = df2.apply(
            lambda r: MERGE[r['sale_key']][0] if r['sale_key'] in MERGE else r['sale_name'], axis=1)
        df2 = self._merge_inactive_group(df2)
        rows = []
        for name, grp in df2.groupby('_group', sort=False):
            t = float(grp['sales_target'].sum()); s = float(grp['total_sales'].sum())
            pct = (s/t*100) if t > 0 else 0.0
            # avg_margin weighted ตาม total_sales
            if 'avg_margin' in grp.columns and s > 0:
                wtd_margin = float((grp['avg_margin'] * grp['total_sales']).sum() / s)
            else:
                wtd_margin = 0.0
            rows.append({'name': name, 'target': t, 'sales': s, 'pct': pct, 'diff': s-t,
                         'margin_target': MARGIN_TARGET, 'actual_margin': wtd_margin,
                         'sales_normal': float(grp['sales_normal'].sum()) if 'sales_normal' in grp.columns else 0.0,
                         'sales_below':  float(grp['sales_below'].sum())  if 'sales_below'  in grp.columns else 0.0})
        # แยก Sale Center (CT) ออกจาก sale staff
        ct_rows   = [r for r in rows if r['name'] == self.SALE_CENTER_KEY]
        other_rows = [r for r in rows if r['name'] != self.SALE_CENTER_KEY]
        # 🟢 [ตามที่ PM ขอ] แยกก้อน "All Inactive Sale" ออกจากตารางเปรียบเทียบผลงานคนที่ยังอยู่
        inactive_rows = [r for r in other_rows if str(r['name']).startswith('All Inactive Sale')]
        sale_rows = [r for r in other_rows if not str(r['name']).startswith('All Inactive Sale')]
        sale_rows.sort(key=lambda r: r['sales'], reverse=True)
        rows = sale_rows  # rows หลักคือ sale staff ที่ยังอยู่ (Active) เท่านั้น

        # รวมทีม (เฉพาะ Active) เทียบกับเป้ารายคนของคนที่ยังอยู่
        total_t      = sum(r['target'] for r in rows)
        total_s      = sum(r['sales']        for r in rows)
        total_normal = sum(r['sales_normal'] for r in rows)
        total_below  = sum(r['sales_below']  for r in rows)
        total_pct = (total_s/total_t*100) if total_t > 0 else 0.0

        # 🟢 [ตามที่ PM ขอ] รวมทั้งบริษัท ใช้เป้าบริษัทรายปีที่ตั้งไว้ (sales_yearly_targets) ถ้ามี
        # ถ้ายังไม่เคยตั้งเป้าปีนั้นไว้ค่อย fallback ไปบวกเป้ารายคน (Active + Inactive)
        company_target = self._get_company_annual_target(self.custom_target_start, self.custom_target_end)
        company_t = company_target if company_target > 0 else (total_t + sum(r['target'] for r in inactive_rows))

        # Period label
        try:
            s_m = self.THAI_MONTHS.index(self.start_m_var.get()) + 1
            e_m = self.THAI_MONTHS.index(self.end_m_var.get())   + 1
            s_y = int(self.start_y_var.get()); e_y = int(self.end_y_var.get())
            period_label = f"{self.start_m_var.get()} {s_y}" if (s_m == e_m and s_y == e_y) else \
                           f"{self.start_m_var.get()} {s_y} – {self.end_m_var.get()} {e_y}"
        except Exception:
            period_label = "รอบที่เลือก"

        floor_pct, goal_pct = self._get_margin_thresholds()

        def _wtd_margin(rs):
            base = sum(r['sales'] for r in rs if r['actual_margin'] > 0)
            return (sum(r['actual_margin'] * r['sales'] for r in rs if r['actual_margin'] > 0) / base) if base > 0 else 0.0

        def _status(margin):
            # ต้องอ่านออกจากข้อความได้โดยไม่ต้องดูสี (พนักงานบางคนแยกแดง/เขียวไม่ได้) — สีเป็นตัวช่วยเท่านั้น
            if margin <= 0:
                return 'ยังไม่มีข้อมูล', 'st_none'
            if margin < floor_pct:
                return 'ต่ำกว่าจุดคุ้มทุน', 'st_low'
            if margin < goal_pct:
                return 'คุ้มทุนแต่ยังไม่ถึงเป้า', 'st_mid'
            return 'ถึงเป้า', 'st_ok'

        def _gap_baht(margin, sales):
            if margin <= 0 or sales <= 0:
                return '—'
            gap = (margin - goal_pct) / 100 * sales
            return f"+{gap:,.0f}" if gap >= 0 else f"{gap:,.0f}"

        team_margin = _wtd_margin(rows)
        company_margin = _wtd_margin(rows + inactive_rows)

        outer = CTkScrollableFrame(parent_frame, fg_color="white", corner_radius=10)
        outer.pack(fill="both", expand=True, padx=10, pady=10)
        CTkLabel(outer, text=f"สรุปยอดขาย vs เป้าหมาย  —  {period_label}",
                 font=CTkFont(size=15, weight="bold"), text_color="#0F172A"
                 ).pack(anchor="w", padx=16, pady=(12, 2))
        CTkLabel(outer, text="GP สุทธิทุกช่อง หักค่าขนส่ง ค่าตัดเจาะพับ และค่านายหน้าแล้ว",
                 font=CTkFont(size=12), text_color="#64748B").pack(anchor="w", padx=16, pady=(0, 8))

        cards = CTkFrame(outer, fg_color="transparent")
        cards.pack(fill="x", padx=16, pady=(0, 10))
        for i in range(4):
            cards.grid_columnconfigure(i, weight=1)

        def _card(col, title, value, note, title_color):
            c = CTkFrame(cards, fg_color="white", corner_radius=10, border_width=1, border_color="#E2E8F0")
            c.grid(row=0, column=col, sticky="nsew", padx=5)
            CTkLabel(c, text=title, font=CTkFont(size=12, weight="bold"), text_color=title_color).pack(anchor="w", padx=14, pady=(10, 0))
            CTkLabel(c, text=value, font=CTkFont(size=24, weight="bold"), text_color="#0F172A").pack(anchor="w", padx=14)
            CTkLabel(c, text=note, font=CTkFont(size=11), text_color="#64748B").pack(anchor="w", padx=14, pady=(0, 10))

        _card(0, "ห้ามต่ำกว่า", f"{floor_pct:.1f}%", "จุดคุ้มทุน ต่ำกว่านี้คือขาดทุน", "#A3231C")
        _card(1, "เป้าที่ต้องทำ", f"{goal_pct:.1f}%", "ระดับที่ทำให้ทั้งบริษัทถึงเป้ากำไร", "#7A4F00")
        _card(2, "ทำได้จริง เฉพาะทีมที่ยัง active", f"{team_margin:.2f}%" if team_margin else "—",
              "ถ่วงน้ำหนักด้วยยอดขาย ไม่ใช่ค่าเฉลี่ยรายบิล", "#14603A")
        _card(3, "รวมทั้งบริษัท", f"{company_margin:.2f}%" if company_margin else "—",
              "รวมยอดที่ไม่มีเจ้าของแล้ว", "#475569")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("SMSalesTable.Treeview", background="white", foreground="#1E293B",
                        rowheight=34, fieldbackground="white", font=('TH Sarabun New', 13))
        style.configure("SMSalesTable.Treeview.Heading", background="#D1FAE5", foreground="#065F46",
                        font=('TH Sarabun New', 13, 'bold'), relief="flat")
        style.map("SMSalesTable.Treeview", background=[('selected', '#EDE9FE')],
                  foreground=[('selected', '#1E293B')])

        cols = ('name', 'sales', 'target', 'margin', 'floor', 'goal', 'status', 'gap')
        tree = ttk.Treeview(outer, columns=cols, show='headings',
                            style="SMSalesTable.Treeview", height=len(rows) + len(inactive_rows) + 4)
        tree.heading('name',   text='พนักงาน');               tree.column('name',   width=190, anchor='w')
        tree.heading('sales',  text='ยอดขายจริง (บาท)');     tree.column('sales',  width=140, anchor='e')
        tree.heading('target', text='เป้าหมาย (บาท)');       tree.column('target', width=140, anchor='e')
        tree.heading('margin', text='GP สุทธิที่ทำได้'); tree.column('margin', width=140, anchor='center')
        tree.heading('floor',  text='ห้ามต่ำกว่า');           tree.column('floor',  width=90,  anchor='center')
        tree.heading('goal',   text='เป้า');                  tree.column('goal',   width=80,  anchor='center')
        tree.heading('status', text='สถานะ');                 tree.column('status', width=170, anchor='w')
        tree.heading('gap',    text='ส่วนต่างจากเป้า (บาท)'); tree.column('gap',    width=150, anchor='e')

        tree.tag_configure('st_low',  background='#FBE9E7', foreground='#A3231C')
        tree.tag_configure('st_mid',  background='#FBF0DC', foreground='#7A4F00')
        tree.tag_configure('st_ok',   background='#E2F0E8', foreground='#14603A')
        tree.tag_configure('st_none', background='#EDEBE4', foreground='#4A4A50')
        tree.tag_configure('total',       background='#EFF6FF', font=('TH Sarabun New', 13, 'bold'), foreground='#1D4ED8')
        tree.tag_configure('grand_total', background='#DCFCE7', font=('TH Sarabun New', 13, 'bold'), foreground='#166534')

        def _put(label, sales, target, margin, bold_tag=None):
            st_text, st_tag = _status(margin)
            tree.insert('', 'end', tags=((bold_tag,) if bold_tag else (st_tag,)),
                        values=(label,
                                f"{sales:,.0f}" if sales > 0 else '—',
                                f"{target:,.0f}" if target > 0 else '—',
                                f"{margin:.2f}%" if margin > 0 else '—',
                                f"{floor_pct:.1f}%", f"{goal_pct:.1f}%",
                                st_text, _gap_baht(margin, sales)))

        for r in rows:
            _put(r['name'], r['sales'], r['target'], r['actual_margin'])
        _put('รวมทีมที่ยัง active', total_s, total_t, team_margin, bold_tag='total')
        for r in inactive_rows:
            _put(r['name'], r['sales'], r['target'], r['actual_margin'])
        ct_sales = sum(r['sales'] for r in ct_rows)
        if ct_rows:
            _put('Sale Center (CT)', ct_sales, 0.0, 0.0)
        inact_sales = sum(r['sales'] for r in inactive_rows)
        _put('รวมทั้งบริษัท', total_s + inact_sales + ct_sales, company_t, company_margin, bold_tag='grand_total')
        tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    # ── Export Excel (แยกรายเดือน + สรุปรวม) ────────────────────────────────────
    def _export_sales_target_excel(self):
        """Export สรุปยอดขาย vs เป้าหมาย เป็น Excel — แยกคอลัมน์ตามเดือนในช่วงที่เลือก (จากรอบ-ถึงรอบ)
        พร้อมคอลัมน์สรุปรวมทั้งช่วงท้ายตาราง (เป้าหมาย/ยอดขายจริง/%/ส่วนต่าง)
        ตามที่ PM ขอ: ไม่เอา Sale Center (CT) และไม่เอาคอลัมน์ Margin เป้า/Avg Margin จริง/ยอดขาย Normal/Below T
        """
        import calendar

        if not getattr(self, 'custom_target_start', None) or not getattr(self, 'custom_target_end', None):
            messagebox.showwarning("ยังไม่ได้ค้นหา", "กรุณากดค้นหาก่อน Export", parent=self)
            return

        orig_start, orig_end = self.custom_target_start, self.custom_target_end
        EXCLUDE = self.EXCLUDE_KEYS
        MERGE = self.PERSON_MERGE

        months = []
        y, m = orig_start.year, orig_start.month
        while (y, m) <= (orig_end.year, orig_end.month):
            months.append((y, m))
            m += 1
            if m > 12:
                m = 1
                y += 1

        inactive_bucket = f"All Inactive Sale {orig_end.year + 543}"

        def _grouped(df):
            df2 = df[~df['sale_key'].isin(EXCLUDE)].copy()
            df2 = df2[df2['sale_key'] != self.SALE_CENTER_KEY]   # ไม่เอา Sale Center (CT)
            df2['_group'] = df2.apply(
                lambda r: MERGE[r['sale_key']][0] if r['sale_key'] in MERGE else r['sale_name'], axis=1)
            df2 = self._merge_inactive_group(df2, bucket_label=inactive_bucket)
            return df2

        try:
            # ── ยอดขายจริงแยกรายเดือน (loop ทีละเดือน ใช้ query/logic เดิมทุกอย่าง รวม history blending) ──
            per_month_sales = {}   # group name -> {(year, month): sales}
            group_order = []
            for (yy, mm) in months:
                last_day = calendar.monthrange(yy, mm)[1]
                self.custom_target_start = datetime(yy, mm, 1)
                self.custom_target_end = datetime(yy, mm, last_day)
                mdf = _grouped(self._get_data())
                for name, grp in mdf.groupby('_group', sort=False):
                    if name not in per_month_sales:
                        per_month_sales[name] = {}
                        group_order.append(name)
                    per_month_sales[name][(yy, mm)] = float(grp['total_sales'].sum())
        finally:
            self.custom_target_start, self.custom_target_end = orig_start, orig_end

        # ── เป้าหมาย/ยอดขายรวมทั้งช่วง — ใช้ query ช่วงเต็ม เหมือนตารางที่แสดงบนจอ ──
        full_df = _grouped(self._get_data())
        target_by_group = {}
        for name, grp in full_df.groupby('_group', sort=False):
            target_by_group[name] = float(grp['sales_target'].sum())
            if name not in per_month_sales:
                per_month_sales[name] = {}
                group_order.append(name)

        total_by_group = {name: sum(per_month_sales[name].values()) for name in group_order}
        group_order.sort(key=lambda n: total_by_group.get(n, 0.0), reverse=True)

        m_names = [f"{self.THAI_MONTHS[mm - 1]} {yy + 543}" for (yy, mm) in months]

        export_rows = []
        for name in group_order:
            row = {"พนักงาน": name}
            for (yy, mm), label in zip(months, m_names):
                row[label] = per_month_sales[name].get((yy, mm), 0.0)
            target = target_by_group.get(name, 0.0)
            sales = total_by_group.get(name, 0.0)
            row["เป้าหมาย (บาท)"] = target
            row["ยอดขายจริง (บาท)"] = sales
            row["%"] = (sales / target * 100) if target > 0 else 0.0
            row["ส่วนต่าง (บาท)"] = sales - target
            export_rows.append(row)

        # แถวรวมทีม
        summary = {"พนักงาน": "รวมทีม"}
        for (yy, mm), label in zip(months, m_names):
            summary[label] = sum(per_month_sales[n].get((yy, mm), 0.0) for n in group_order)
        # 🟢 [ตามที่ PM ขอ] ใช้เป้าบริษัทรายปีที่ตั้งไว้ (sales_yearly_targets) แทนการบวกเป้ารายคน —
        # ให้ตรงกับตัวเลข "ทีมรวม" บนกราฟเป้าการขาย ถ้ายังไม่เคยตั้งเป้าปีนั้นไว้ค่อย fallback ไปบวกเป้ารายคน
        company_target = self._get_company_annual_target(orig_start, orig_end)
        total_t = company_target if company_target > 0 else sum(target_by_group.get(n, 0.0) for n in group_order)
        total_s = sum(total_by_group.get(n, 0.0) for n in group_order)
        summary["เป้าหมาย (บาท)"] = total_t
        summary["ยอดขายจริง (บาท)"] = total_s
        summary["%"] = (total_s / total_t * 100) if total_t > 0 else 0.0
        summary["ส่วนต่าง (บาท)"] = total_s - total_t
        export_rows.append(summary)

        df_export = pd.DataFrame(export_rows)

        s_lbl = f"{self.THAI_MONTHS[orig_start.month - 1]} {orig_start.year + 543}"
        e_lbl = f"{self.THAI_MONTHS[orig_end.month - 1]} {orig_end.year + 543}"
        period_label = s_lbl if s_lbl == e_lbl else f"{s_lbl} - {e_lbl}"

        export_sales_target_to_excel(self, df_export, period_label)

