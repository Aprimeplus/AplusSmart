# sales_target_widget.py
# SalesTargetWidget — แท็บ "เป้าการขาย" / "Sale Revenue Report"
# ใช้ร่วมกันทั้งหน้า Sales Manager (sales_manager_screen.py) และหน้า
# รายงานผู้บริหาร (management_report_screen.py) — เดิมมี 2 ก้อนโค้ดแยกกันเกือบเหมือนกัน
# ทำให้แก้บัคแล้วลืมแก้อีกจุด (เช่น เป้าบริษัทรายปี, การรวม Inactive Sale) จึงรวมเป็นจุดเดียวที่นี่

import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                               CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                               CTkOptionMenu, CTkRadioButton, CTkTabview, CTkCheckBox, CTkTextbox)
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
from project_screen import _center_and_style_popup
import qt_conversion as qcv


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

        CTkButton(fbar, text="📥 นำเข้า QT", width=110, fg_color="#7C3AED", hover_color="#6D28D9",
                  command=self._open_qt_import).pack(side="left", padx=(0, 10))

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

    def _open_qt_import(self):
        QTImportDialog(self, self.app_container, on_done=self._on_search)

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
            SELECT COALESCE(SUM(sc.sales_service_amount), 0) AS total_sales,
                   -- GP ประมาณการของ CT = ยอดขาย − (ต้นทุน PO × ตัวคูณ 1.03) เฉพาะ SO ที่มี PO แล้ว
                   COALESCE(SUM(CASE WHEN po.cost > 0 THEN sc.sales_service_amount END), 0) AS gp_base,
                   COALESCE(SUM(CASE WHEN po.cost > 0 THEN
                       sc.sales_service_amount - po.cost * (CASE WHEN COALESCE(sc.cost_multiplier, 0) < 1.01
                                                                 THEN 1.03 ELSE sc.cost_multiplier END)
                   END), 0) AS gp_est
            FROM commissions sc
            LEFT JOIN (
                SELECT so_number, COALESCE(SUM(total_cost), 0) AS cost
                FROM purchase_orders
                WHERE status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
                GROUP BY so_number
            ) po ON po.so_number = sc.so_number
            WHERE sc.sale_key IN ('Sale Center', 'CHARITA-CT')
              AND sc.is_active = 1
              AND sc.status NOT IN ('Cancelled', 'Cancelled by PU')
              AND {sc_date_filter}
        """
        try:
            sc_df = pd.read_sql_query(sc_query, self.pg_engine, params=tuple(date_params))
            sc_total = float(sc_df['total_sales'].iloc[0]) if not sc_df.empty else 0.0
            sc_gp_base = float(sc_df['gp_base'].iloc[0]) if not sc_df.empty else 0.0
            sc_gp_est = float(sc_df['gp_est'].iloc[0]) if not sc_df.empty else 0.0
        except Exception:
            sc_total = 0.0
            sc_gp_base = 0.0
            sc_gp_est = 0.0

        if sc_total > 0:
            sc_row = pd.DataFrame([{
                'sale_name': 'Sale Center', 'sale_key': 'Sale Center', 'status': 'Active',
                'sales_target': 0.0, 'total_sales': sc_total,
                'total_outstanding': 0.0,
                'total_gp': sc_gp_est, 'total_sales_amt': sc_gp_base,
                'avg_margin': (sc_gp_est / sc_gp_base * 100) if sc_gp_base > 0 else 0.0,
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
        _center_and_style_popup(dlg, self, 380, 230)   # กลางหน้าต่างหลัก + ไอคอนแอป (เดิมเด้งผิดจอ ไอคอนเป็นของ tk)
        dlg.resizable(False, False)
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
    def _get_heatmap_df(self, year):
        """ยอดขาย + GP รายเดือนต่อคน ของปีที่เลือก (ค.ศ.) สำหรับ Heatmap
        คืน DataFrame: sale_key, sale_name, status, month, sales, gp, gp_base
        (gp/gp_base ใช้คิด GP% แบบถ่วงน้ำหนัก = Σ GP บาท ÷ Σ ยอดขายที่คิด GP)"""
        sale_filter = ""
        params = [int(year)]
        if self.selected_sales_filter:
            ph = ','.join(["REPLACE(LOWER(%s), ' ', '')"] * len(self.selected_sales_filter))
            sale_filter = f" AND REPLACE(LOWER(su.sale_key), ' ', '') IN ({ph})"
            params.extend(self.selected_sales_filter)
        query = f"""
            SELECT su.sale_name, su.sale_key, su.status,
                   c.commission_month AS month,
                   COALESCE(SUM(c.total_sales), 0)       AS sales,
                   COALESCE(SUM(cm.total_gp), 0)         AS gp,
                   COALESCE(SUM(cm.total_sales_amt), 0)  AS gp_base
            FROM   sales_users su
            JOIN   commission_payout_logs c
                   ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(c.sale_key), ' ', '')
                  AND c.commission_year = %s
            LEFT JOIN (
                SELECT payout_id, SUM(final_gp) AS total_gp, SUM(final_sales_amount) AS total_sales_amt
                FROM   commissions GROUP BY payout_id
            ) cm ON cm.payout_id = c.id
            WHERE  1 = 1 {sale_filter}
            GROUP  BY su.sale_name, su.sale_key, su.status, c.commission_month
        """
        df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
        # ข้อมูลย้อนหลัง (ปีก่อนมีระบบจริง) จาก sales_history_monthly
        try:
            hist = pd.read_sql_query(f"""
                SELECT su.sale_name, su.sale_key, su.status, h.month,
                       h.actual_sales AS sales,
                       CASE WHEN h.margin_pct IS NOT NULL THEN h.margin_pct / 100 * h.actual_sales ELSE 0 END AS gp,
                       CASE WHEN h.margin_pct IS NOT NULL THEN h.actual_sales ELSE 0 END AS gp_base
                FROM   sales_history_monthly h
                JOIN   sales_users su
                       ON REPLACE(LOWER(su.sale_key), ' ', '') = REPLACE(LOWER(h.sale_key), ' ', '')
                WHERE  h.year = %s {sale_filter}
            """, self.pg_engine, params=tuple(params))
            if not hist.empty:
                df = pd.concat([df, hist], ignore_index=True)
        except Exception:
            traceback.print_exc()
        for c in ('sales', 'gp', 'gp_base'):
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
        return df

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
        company_margin = _wtd_margin(rows + inactive_rows + ct_rows)

        # ── ค่าเวลา/Pace (ตาม DEV-SPEC-SALEDASH-001) ─────────────────────────────────
        # N = จำนวนเดือนที่เลือก (ปกติ ม.ค.–เดือนสิ้นสุด), T% = N/12, L = เดือนที่เหลือของปี
        SHORT_M = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
        PACE_BAND = 10.0   # ช่วง "เฝ้าระวัง" (percentage point) — ยังไม่มีหน้าตั้งค่า ใช้ค่าเริ่มต้นของ spec
        s_dt, e_dt = self.custom_target_start, self.custom_target_end
        try:
            N = (e_dt.year * 12 + e_dt.month) - (s_dt.year * 12 + s_dt.month) + 1
        except Exception:
            N = 12
        N = max(1, min(N, 12))
        T_PCT = N / 12 * 100
        L = 12 - N
        e_short = SHORT_M[e_dt.month - 1] if e_dt else ""
        s_short = SHORT_M[s_dt.month - 1] if s_dt else ""
        range_short = f"{s_short}–{e_short}"
        rest_short = f"{SHORT_M[e_dt.month % 12]}–ธ.ค." if (L > 0 and e_dt) else "—"

        PACE = {
            'ok':    ("ตามแผน",       "#E3F3EA", "#1E7B4C", "#2E9E64"),
            'watch': ("เฝ้าระวัง",     "#FFF1D6", "#8A5A00", "#E0A100"),
            'low':   ("ต่ำกว่าแผน",    "#FDE7E5", "#B42318", "#D9534F"),
            'none':  ("ยังไม่มอบหมาย", "#EEF0F3", "#5B6573", "#9AA4B2"),
            'nogoal': ("ไม่มีเกณฑ์",   "#EEF0F3", "#5B6573", "#9AA4B2"),
        }
        GPST = {
            'ok':    "#E3F3EA", 'mid': "#FFF1D6", 'low': "#FDE7E5", 'none': "#EEF0F3",
        }
        GPFG = {
            'ok':    "#1E7B4C", 'mid': "#8A5A00", 'low': "#B42318", 'none': "#5B6573",
        }

        def _pace_key(pct_year):
            if pct_year is None:
                return 'none'
            if pct_year >= T_PCT:
                return 'ok'
            if pct_year >= T_PCT - PACE_BAND:
                return 'watch'
            return 'low'

        def _m(v):
            return f"{v / 1e6:,.2f} ล้าน"

        def _num(v):
            return f"{v:,.0f}"

        def _calc(sales, ytd, annual):
            """คืน dict ตัวเลขทั้งหมดของแถวในตาราง A (None = แสดง —)"""
            d = {'sales': sales, 'ytd': ytd, 'annual': annual}
            d['ytd_pct'] = (sales / ytd * 100) if ytd > 0 else None
            d['year_pct'] = (sales / annual * 100) if annual > 0 else None
            if annual > 0:
                d['short'] = max(annual - sales, 0.0)
                d['over'] = sales - annual if sales > annual else 0.0
                d['per_month'] = (d['short'] / L) if L > 0 else None
            else:
                d['short'] = d['over'] = d['per_month'] = None
            return d

        # แถวผู้ขาย (Active) — เป้าทั้งปี = เป้าช่วงที่เลือก ÷ N × 12 (เป้ารายเดือนเท่ากันทุกเดือน)
        def _annual_of(ytd):
            return ytd / N * 12 if ytd > 0 else 0.0

        sale_calc = [(r, _calc(r['sales'], r['target'], _annual_of(r['target']))) for r in rows]
        team_calc = _calc(total_s, total_t, _annual_of(total_t))
        inact_t = sum(r['target'] for r in inactive_rows)
        inact_s = sum(r['sales'] for r in inactive_rows)
        inact_calc = _calc(inact_s, inact_t, _annual_of(inact_t))
        ct_s = sum(r['sales'] for r in ct_rows)
        company_s = total_s + inact_s + ct_s
        company_calc = _calc(company_s, company_t, _annual_of(company_t))

        # ══════════════════════════ UI ══════════════════════════
        outer = CTkScrollableFrame(parent_frame, fg_color="#F4F6F9", corner_radius=10)
        # ยังไม่ pack ตอนนี้: สร้าง widget ทั้งหมดขณะที่ยังไม่แสดงผล แล้วค่อยแสดงทีเดียวตอนท้าย (ลดการคำนวณ layout/วาดซ้ำ)
        # ตอนสร้าง widget จำนวนมาก ตัว scrollbar ของ CTk จะสั่ง update_idletasks ซ้ำทุกครั้งที่ขนาดเปลี่ยน ทำให้หน้าโหลดนานมาก
        # (วัดได้ ~40 วินาที) → ปิดการอัปเดต scrollbar ชั่วคราวระหว่างสร้างหน้า แล้วคืนค่าตอนท้ายฟังก์ชัน
        try:
            _ysc_saved = outer._parent_canvas.cget("yscrollcommand")
            outer._parent_canvas.configure(yscrollcommand=lambda *_a: None)
        except Exception:
            _ysc_saved = None

        head = CTkFrame(outer, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(12, 6))
        left = CTkFrame(head, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        CTkLabel(left, text="APLUS SMART · SALES", font=CTkFont(size=11, weight="bold"),
                 text_color="#1F4E79").pack(anchor="w")
        CTkLabel(left, text=f"สรุปยอดขาย vs เป้าหมาย  —  {period_label}",
                 font=CTkFont(size=20, weight="bold"), text_color="#0F172A").pack(anchor="w")
        CTkLabel(left, text="GP สุทธิทุกช่อง หักค่าขนส่ง ค่าตัดเจาะพับ และค่านายหน้าแล้ว",
                 font=CTkFont(size=12), text_color="#64748B").pack(anchor="w")
        CTkLabel(head, text=f"  เวลาผ่านไป {N} / 12 เดือน ({T_PCT:.0f}%)  ",
                 font=CTkFont(size=12, weight="bold"), text_color="#1F4E79",
                 fg_color="#E3ECF6", corner_radius=8, height=32).pack(side="right", padx=(8, 0))

        def _bar(parent, pct, color, width=150, height=14):
            cv = tk.Canvas(parent, width=width, height=height, highlightthickness=0, bg="white")
            cv.create_rectangle(0, 3, width, height - 3, fill="#E9EDF2", outline="")
            if pct is not None and pct > 0:
                cv.create_rectangle(0, 3, max(2, min(pct, 100) / 100 * width), height - 3, fill=color, outline="")
            mx = T_PCT / 100 * width
            cv.create_line(mx, 0, mx, height, fill="#111827", width=2)
            return cv

        _font_cache = {}

        def _fnt(size, bold=False):
            """ฟอนต์สำหรับ tk.Label / Canvas ใช้รูปแบบ tuple (ครอบครัว, -พิกเซล, น้ำหนัก)
            ห้ามส่ง Font/CTkFont object เข้า Canvas.create_text — ในแอปนี้ Tk จะไม่สนใจ object ฟอนต์ (ขนาดไม่เปลี่ยนตามที่สั่ง)"""
            k_ = (size, bold)
            if k_ not in _font_cache:
                _font_cache[k_] = (CTkFont().cget("family"), -size, "bold" if bold else "normal")
            return _font_cache[k_]

        _cfnt = _fnt

        def _pill(parent, text, bg, fg):
            return CTkLabel(parent, text=text, font=CTkFont(size=12, weight="bold"), text_color=fg,
                            fg_color=bg, corner_radius=12, height=26)

        # กล่องแจ้งเตือน + การ์ดทีม 4 ใบ: สร้างที่ว่างไว้ก่อน แล้วเติมหลังคำนวณ Heatmap (ใช้ข้อมูลรายเดือน)
        alerts_holder = CTkFrame(outer, fg_color="transparent")
        alerts_holder.pack(fill="x")
        cards_holder = CTkFrame(outer, fg_color="transparent")
        cards_holder.pack(fill="x", padx=11, pady=(4, 8))
        for i in range(4):
            cards_holder.grid_columnconfigure(i, weight=1, uniform="kpi4")

        # ── ตัวช่วยสร้างตารางแบบ grid (รองรับแถบสี/ป้ายสถานะ) ──
        def _make_table(title, legend_text, headers, anchors, minw, weights):
            card = CTkFrame(outer, fg_color="white", corner_radius=12, border_width=1, border_color="#E1E6ED")
            card.pack(fill="x", padx=16, pady=(8, 8))
            th = CTkFrame(card, fg_color="transparent")
            th.pack(fill="x", padx=16, pady=(12, 6))
            CTkLabel(th, text=title, font=CTkFont(size=15, weight="bold"), text_color="#0F172A").pack(side="left")
            if legend_text:
                CTkLabel(th, text=legend_text, font=CTkFont(size=11), text_color="#64748B").pack(side="right")
            body = CTkFrame(card, fg_color="transparent")
            body.pack(fill="x", padx=12, pady=(0, 12))
            for ci, (w, mw) in enumerate(zip(weights, minw)):
                body.grid_columnconfigure(ci, weight=w, minsize=mw)
            for ci, h in enumerate(headers):
                hf = tk.Frame(body, bg="#E8F3EC")
                hf.grid(row=0, column=ci, sticky="nsew")
                tk.Label(hf, text=h, font=_fnt(12, True), fg="#1F4E79", bg="#E8F3EC",
                         anchor=anchors[ci], justify="left" if anchors[ci] == "w" else "right",
                         wraplength=max(minw[ci] - 12, 60)).pack(fill="x", padx=10, pady=8, expand=True)
            return body

        def _cell(body, r, c, text, anchor, bg, bold=False, color="#1E293B", sub=None):
            f = tk.Frame(body, bg=bg)
            f.grid(row=r, column=c, sticky="nsew")
            tk.Label(f, text=text, font=_fnt(13, bold), fg=color, bg=bg,
                     anchor=anchor).pack(fill="x", padx=10, pady=(7, 0 if sub else 7), expand=not sub)
            if sub:
                tk.Label(f, text=sub, font=_fnt(10), fg="#B7791F", bg=bg, anchor=anchor
                         ).pack(fill="x", padx=10, pady=(0, 5))
            return f

        def _cell_widget(body, r, c, bg, builder, anchor="w"):
            f = tk.Frame(body, bg=bg)
            f.grid(row=r, column=c, sticky="nsew")
            w = builder(f)
            w.pack(anchor=anchor, padx=10, pady=6, expand=True)
            return f

        # ══ ตาราง A: ยอดขายเทียบเป้า ══
        legend_a = f"เขียว ตามแผน ≥ {T_PCT:.0f}%   ·   เหลือง เฝ้าระวัง {T_PCT - PACE_BAND:.0f}–{T_PCT:.0f}%   ·   แดง ต่ำกว่าแผน < {T_PCT - PACE_BAND:.0f}%"
        hdr_a = ["พนักงาน", "ยอดขายจริง (บาท)", f"เป้า YTD {range_short}", "% ถึงเป้า YTD", "เป้าทั้งปี",
                 "% ของเป้าปี", "ยอดที่ยังขาด", f"ต้องทำ/เดือน {rest_short}", "สถานะ Pace"]
        anc_a = ["w", "e", "e", "e", "e", "w", "e", "e", "w"]
        tb_a = _make_table("A · ยอดขายเทียบเป้า (Sales Achievement)", legend_a, hdr_a, anc_a,
                           [150, 110, 110, 80, 110, 220, 110, 110, 120], [3, 2, 2, 1, 2, 3, 2, 2, 2])

        row_i = [1]

        def _put_a(label, d, kind='person', sub=None):
            r = row_i[0]; row_i[0] += 1
            bg = {"team": "#F1F5FA", "company": "#E6EEF7"}.get(kind, "white")
            bold = kind in ("team", "company")
            nm_color = "#1F4E79" if bold else "#1E293B"
            _cell(tb_a, r, 0, label, "w", bg, bold, nm_color, sub)
            _cell(tb_a, r, 1, _num(d['sales']) if d['sales'] else "—", "e", bg, bold)
            _cell(tb_a, r, 2, _num(d['ytd']) if d['ytd'] > 0 else "—", "e", bg, bold)
            _cell(tb_a, r, 3, f"{d['ytd_pct']:.1f}%" if d['ytd_pct'] is not None else "—", "e", bg, bold)
            _cell(tb_a, r, 4, _num(d['annual']) if d['annual'] > 0 else "—", "e", bg, bold)
            # % ของเป้าปี: แถบ + marker เวลา + ตัวเลข (แสดงค่าจริงแม้เกิน 100%)
            pk_ = 'none' if kind == 'inactive' else ('nogoal' if kind == 'ct' else _pace_key(d['year_pct']))
            yp_ = d['year_pct']

            def _bar_cell(f, yp_=yp_, pk_=pk_):
                box = CTkFrame(f, fg_color="transparent")
                if yp_ is None:
                    CTkLabel(box, text="—", font=CTkFont(size=13), text_color="#64748B").pack(side="left")
                else:
                    _bar(box, yp_, PACE[pk_][3]).pack(side="left")
                    CTkLabel(box, text=f"  {yp_:.1f}%", font=CTkFont(size=13, weight="bold"),
                             text_color=PACE[pk_][2]).pack(side="left")
                return box
            _cell_widget(tb_a, r, 5, bg, _bar_cell)
            if d['short'] is None:
                short_txt, per_txt, short_col = "—", "—", "#1E293B"
            elif d['over'] > 0:
                short_txt, per_txt, short_col = f"เกินเป้า +{_num(d['over'])}", "0", "#1E7B4C"
            else:
                short_txt = _num(d['short'])
                per_txt = _num(d['per_month']) if d['per_month'] is not None else "—"
                short_col = "#1E293B"
            if kind == 'inactive':
                per_txt = "—"
            _cell(tb_a, r, 6, short_txt, "e", bg, bold, short_col)
            _cell(tb_a, r, 7, per_txt, "e", bg, bold)
            lab, pbg, pfg, _ = PACE[pk_]
            _cell_widget(tb_a, r, 8, bg, lambda f, lab=lab, pbg=pbg, pfg=pfg: _pill(f, f"  {lab}  ", pbg, pfg))

        for r_, d_ in sale_calc:
            _put_a(r_['name'], d_)
        _put_a("รวมทีมที่ยัง active", team_calc, 'team')
        for r_ in inactive_rows:
            _put_a(r_['name'], inact_calc, 'inactive', sub="เป้ายังไม่มีผู้รับผิดชอบ")
        if ct_rows:
            _put_a("Sale Center (CT) *", _calc(ct_s, 0.0, 0.0), 'ct')
        _put_a("รวมทั้งบริษัท", company_calc, 'company')

        # ══ ตาราง KPI: ผลงานรายคน (ตามเกณฑ์ KPI ปี 2569 — เป้าหมายหลัก 100%, ผ่านเกณฑ์ ≥ 85%) ══
        # ระดับ 1–5 (ยอดขาย): เซลล์ ≥101 / 90 / 85 / 75 · ทีม(ผู้จัดการ) ≥101 / 95 / 90 / 85 · เทียบที่ทศนิยม 2 ตำแหน่ง ถึงเส้นก็นับ
        # ระดับ GP: > 20 / > 15 / > 13 / > 12 (ต้อง "เกิน" เส้น) · ทุกบทบาทใช้เส้นเดียวกัน
        LV_NAME = {5: "ดีมาก", 4: "ดี", 3: "มาตรฐาน", 2: "ปรับปรุง", 1: "ต่ำกว่าเกณฑ์"}
        LV_BG = {5: "#1E7B4C", 4: "#CFEBDA", 3: "#DCE8F3", 2: "#FFE3A8", 1: "#F6C1BB"}
        LV_FG = {5: "#FFFFFF", 4: "#14603A", 3: "#1F4E79", 2: "#8A5A00", 1: "#B42318"}
        PASS_PCT = 85.0

        def _sales_level(pct, is_team=False):
            p = round(pct, 2)
            cuts = (101.0, 95.0, 90.0, 85.0) if is_team else (101.0, 90.0, 85.0, 75.0)
            if p >= cuts[0]:
                return 5
            if p >= cuts[1]:
                return 4
            if p >= cuts[2]:
                return 3
            if p >= cuts[3]:
                return 2
            return 1

        def _gp_level(gp):
            g = round(gp, 2)
            if g > 20.0:
                return 5
            if g > 15.0:
                return 4
            if g > 13.0:
                return 3
            if g > 12.0:
                return 2
            return 1

        tb_k = _make_table(
            "ผลงานรายคน · " + period_label,
            "เซลล์ใช้เกณฑ์เซลล์ · แถวทีมใช้เกณฑ์ผู้จัดการ",
            ["พนักงาน", "ยอดขาย (บาท)", "เป้า (บาท)", "% ต่อเป้า", "ระดับยอดขาย", "ผล", "GP%", "ระดับ GP", "แจ้งเตือน"],
            ["w", "e", "e", "e", "w", "w", "e", "w", "w"],
            [190, 120, 120, 80, 150, 70, 80, 150, 150], [4, 2, 2, 1, 3, 1, 1, 3, 3])
        row_k = [1]

        def _lv_pill(text, lv):
            return lambda f: _pill(f, f"  {lv} {text}  ", LV_BG[lv], LV_FG[lv])

        def _grey_pill(f):
            return _pill(f, "  ไม่มีเกณฑ์  ", "#EEF0F3", "#5B6573")

        def _put_k(label, sub, sales, target, margin, kind="person"):
            r = row_k[0]; row_k[0] += 1
            bg = {"team": "#E6EEF7", "company": "#EEF1F5"}.get(kind, "white")
            bold = kind in ("team", "company")
            _cell(tb_k, r, 0, label, "w", bg, bold, "#1F4E79" if bold else "#1E293B", sub)
            _cell(tb_k, r, 1, _num(sales) if sales else "—", "e", bg, bold)
            has_target = kind in ("person", "team") and target > 0
            pct = (sales / target * 100) if has_target else None
            _cell(tb_k, r, 2, _num(target) if has_target else "—", "e", bg, bold)
            _cell(tb_k, r, 3, f"{pct:.2f}%" if pct is not None else "—", "e", bg, bold,
                  "#1E7B4C" if (pct is not None and pct >= PASS_PCT) else ("#B42318" if pct is not None else "#1E293B"))
            if pct is not None:
                lv = _sales_level(pct, is_team=(kind == "team"))
                _cell_widget(tb_k, r, 4, bg, _lv_pill(LV_NAME[lv], lv))
                passed = round(pct, 2) >= PASS_PCT
                _cell(tb_k, r, 5, "ผ่าน" if passed else "ไม่ผ่าน", "w", bg, True, "#1E7B4C" if passed else "#B42318")
            else:
                _cell_widget(tb_k, r, 4, bg, _grey_pill)
                _cell(tb_k, r, 5, "—", "w", bg)
            if margin and margin > 0:
                glv = _gp_level(margin)
                _cell(tb_k, r, 6, f"{margin:.2f}%", "e", bg, bold)
                if kind in ("person", "team", "inactive"):
                    _cell_widget(tb_k, r, 7, bg, _lv_pill(LV_NAME[glv], glv))
                else:
                    _cell_widget(tb_k, r, 7, bg, _grey_pill)
            else:
                _cell(tb_k, r, 6, "—", "e", bg, bold)
                _cell_widget(tb_k, r, 7, bg, _grey_pill)
            # แจ้งเตือน (บน dashboard เท่านั้น)
            alert, alert_col = "—", "#1E293B"
            if margin and 0 < margin < 10 and kind in ("person", "team", "inactive"):
                alert, alert_col = "GP วิกฤต < 10%", "#B42318"
            elif pct is not None and round(pct, 2) < PASS_PCT:
                alert, alert_col = ("ทีมไม่ผ่าน 85%" if kind == "team" else "ไม่ผ่านเกณฑ์"), "#B42318"
            elif pct is not None and margin and 0 < margin < 13:
                alert, alert_col = "ยอดดีแต่กำไรต่ำ", "#B42318"
            _cell(tb_k, r, 8, alert, "w", bg, False, alert_col)

        for r_ in rows:
            _put_k(r_['name'], None, r_['sales'], r_['target'], r_['actual_margin'])
        for r_ in inactive_rows:
            _put_k(f"ยอดขาย Inactive {s_dt.year + 543 if s_dt else ''}", "ไม่มีเป้า · นับยอดรวมในทีม",
                   r_['sales'], 0.0, r_['actual_margin'], 'inactive')
        team_sales_all = total_s + inact_s
        team_margin_all = _wtd_margin(rows + inactive_rows)
        _put_k("ทีมขาย", "active + inactive ÷ เป้า active · เกณฑ์ผู้จัดการ", team_sales_all, total_t, team_margin_all, 'team')
        if ct_rows:
            _put_k("Sale Center (CT)", "ไม่อยู่ในเกณฑ์ KPI · GP ประมาณการ", ct_s, 0.0, _wtd_margin(ct_rows), 'ct')
        _put_k("รวมทั้งบริษัท", None, company_s, 0.0, company_margin, 'company')

        # ══ เป้าหมายของแต่ละคน (ยึดเป้า 100%) — ยอดที่ต้องทำต่อเดือนในช่วงที่เหลือ ══
        # ทุกคนเห็น 4 ระดับพร้อมกัน: 100% เป้าทั้งปี / 100% เป้าไตรมาส / Incentive ขั้น 1 (≥101%) / ขั้น 2 (≥115%)
        # ตัวเลข = ยอดที่ต้องทำต่อเดือน · % = ต้องเพิ่มจากยอดเฉลี่ยที่ทำมา (ยอดสะสม ÷ จำนวนเดือน)
        import math
        L_rem = 12 - N
        if L_rem > 0 and e_dt:
            qi = (e_dt.month - 1) // 3 if e_dt.month % 3 != 0 else e_dt.month // 3   # ไตรมาสเป้าหมาย (0-based)
            q_months = [3 * qi + 1, 3 * qi + 2, 3 * qi + 3]
            q_rem = [m_ for m_ in q_months if m_ > e_dt.month]
            m_rem = len(q_rem)
            q_label = f"Q{qi + 1}"
            q_range = f"{SHORT_M[q_months[0] - 1]}–{SHORT_M[q_months[-1] - 1]}"

            def _quarter_sales():
                """ยอดขายของไตรมาสเป้าหมายที่ทำไปแล้ว (กรณีเลือกถึงกลางไตรมาส) แยกตามกลุ่มพนักงาน"""
                qs = datetime(e_dt.year, q_months[0], 1)
                if e_dt < qs:
                    return {}
                saved_rng = (self.custom_target_start, self.custom_target_end)
                try:
                    self.custom_target_start, self.custom_target_end = qs, e_dt
                    qdf = self._get_data()
                finally:
                    self.custom_target_start, self.custom_target_end = saved_rng
                qdf = qdf[~qdf['sale_key'].isin(EXCLUDE)].copy()
                qdf['_group'] = qdf.apply(
                    lambda rr: MERGE[rr['sale_key']][0] if rr['sale_key'] in MERGE else rr['sale_name'], axis=1)
                qdf = self._merge_inactive_group(qdf)
                return qdf.groupby('_group')['total_sales'].sum().to_dict()

            q_sales_by = _quarter_sales()

            def _need(q_target, q_sales, factor):
                """ยอดที่ต้องทำต่อเดือนในช่วงที่เหลือของไตรมาส เพื่อให้ได้ factor × เป้าไตรมาส"""
                return max(q_target * factor - q_sales, 0.0) / m_rem

            def _inc_pill(need, avg):
                if avg <= 0:
                    return None
                inc = (need / avg - 1) * 100
                if inc <= 10:
                    return inc, "#E3F3EA", "#1E7B4C"
                if inc <= 30:
                    return inc, "#FFF1D6", "#8A5A00"
                return inc, "#FDE7E5", "#B42318"

            def _num_pill_cell(need, avg):
                pill = _inc_pill(need, avg)

                def _b(f, need=need, pill=pill):
                    box = CTkFrame(f, fg_color="transparent")
                    CTkLabel(box, text=_num(need), font=CTkFont(size=13), text_color="#1E293B", anchor="e").pack(anchor="e")
                    if pill:
                        inc, pbg, pfg = pill
                        CTkLabel(box, text=f" +{inc:.1f}% " if inc >= 0 else f" {inc:.1f}% ", font=CTkFont(size=10, weight="bold"),
                                 text_color=pfg, fg_color=pbg, corner_radius=9, height=18).pack(anchor="e", pady=(2, 0))
                    return box
                return _b

            goal_range = f"{rest_short}" if rest_short != "—" else ""
            tb_g = _make_table(
                f"เป้าหมายของแต่ละคน · {goal_range} {e_dt.year + 543} (ยึดเป้า 100%)",
                "เขียว เพิ่ม ≤ 10%  ·  เหลือง 10–30% ต้องวางแผน  ·  แดง > 30% ผู้จัดการต้องเข้าช่วย",
                ["พนักงาน", "ทำได้เฉลี่ย/เดือน", "100% เป้าทั้งปี", f"100% เป้า {q_label}",
                 f"Incentive ขั้นที่ 1 ({q_label} ≥ 101%)", f"Incentive ขั้นที่ 2 ({q_label} ≥ 115%)",
                 "GP (ต้อง ≥ 15.00%)", "สิ่งที่ต้องทำ"],
                ["w", "e", "e", "e", "e", "e", "w", "w"],
                [200, 110, 110, 110, 130, 130, 150, 260], [4, 2, 2, 2, 2, 2, 3, 5])
            row_g = [1]
            GP_OK = 15.0

            def _gp_pill_cell(margin):
                ok = margin >= GP_OK
                txt = f"  {margin:.2f}% ✓  " if ok else f"  {margin:.2f}% ✗ ต่ำกว่า 15%  "
                return lambda f: _pill(f, txt, "#E3F3EA" if ok else "#FDE7E5", "#1E7B4C" if ok else "#B42318")

            def _put_g(label, sub, avg, ann_need, q_target, q_sales, margin, action, kind="person"):
                r = row_g[0]; row_g[0] += 1
                bg = "#E6EEF7" if kind == "team" else "white"
                _cell(tb_g, r, 0, label, "w", bg, kind == "team", "#1F4E79" if kind == "team" else "#1E293B", sub)
                _cell(tb_g, r, 1, _num(avg), "e", bg, kind == "team")
                _cell_widget(tb_g, r, 2, bg, _num_pill_cell(ann_need, avg), "e")
                _cell_widget(tb_g, r, 3, bg, _num_pill_cell(_need(q_target, q_sales, 1.00), avg), "e")
                _cell_widget(tb_g, r, 4, bg, _num_pill_cell(_need(q_target, q_sales, 1.01), avg), "e")
                _cell_widget(tb_g, r, 5, bg, _num_pill_cell(_need(q_target, q_sales, 1.15), avg), "e")
                _cell_widget(tb_g, r, 6, bg, _gp_pill_cell(margin))
                _cell(tb_g, r, 7, action, "w", bg, kind == "team")

            def _action(inc_q, margin):
                if inc_q is None:
                    return "—"
                if inc_q <= 10:
                    txt = f"ถึง 100% {q_label} ได้ด้วยแรงเพิ่มเล็กน้อย → ไล่ Incentive ขั้นที่ 1/2"
                elif inc_q <= 30:
                    txt = "วางแผนเพิ่มบิล T3 (100,000–299,999) สัปดาห์ละ 1 ใบ"
                else:
                    txt = "ผู้จัดการเข้าช่วย: ดีล T3–T4 + ไล่ QT ค้าง"
                if margin < GP_OK:
                    txt += " · GP ต้องถึง 15.00% จึงจะได้ Incentive"
                return txt

            for r_, d_ in sale_calc:
                avg_ = r_['sales'] / N
                q_target_ = (d_['annual'] / 12.0) * 3 if d_['annual'] > 0 else 0.0
                q_sales_ = q_sales_by.get(r_['name'], 0.0)
                ann_need_ = max(d_['annual'] - r_['sales'], 0.0) / L_rem if d_['annual'] > 0 else 0.0
                pill_q = _inc_pill(_need(q_target_, q_sales_, 1.00), avg_)
                sub_ = (f"ยอดสะสม {d_['ytd_pct']:.2f}% ของเป้า {range_short}" if d_['ytd_pct'] is not None else None)
                _put_g(r_['name'], sub_, avg_, ann_need_, q_target_, q_sales_, r_['actual_margin'],
                       _action(pill_q[0] if pill_q else None, r_['actual_margin']))

            # แถวทีม: เฉลี่ยใช้ยอดทีม active · ส่วนที่ยังขาดถึงเป้าปีใช้ยอดรวม inactive (ตาม spec)
            team_avg = total_s / N
            team_ann = team_calc['annual']
            team_q_target = (team_ann / 12.0) * 3 if team_ann > 0 else 0.0
            team_q_sales = sum(q_sales_by.get(r_['name'], 0.0) for r_ in rows) + sum(
                q_sales_by.get(r_['name'], 0.0) for r_ in inactive_rows)
            team_ann_need = max(team_ann - team_sales_all, 0.0) / L_rem if team_ann > 0 else 0.0
            tq_need = _need(team_q_target, team_q_sales, 1.00)
            add_m = max(tq_need - team_avg, 0.0)
            bills = math.ceil(add_m / 174000.0) if add_m > 0 else 0
            team_action = (f"100% {q_label}: เพิ่ม ~{add_m / 1e6:.1f} ล้าน/เดือน ≈ บิล T3 เพิ่ม ~{bills} ใบ/เดือนทั้งทีม · "
                           f"100% ทั้งปี: ต้องเพิ่ม T3 + T4 ทุกคน") if add_m > 0 else "ทีมอยู่ในแนวทางถึง 100%"
            team_pct = (team_sales_all / total_t * 100) if total_t > 0 else None
            team_sub = (f"ยอดสะสม {team_pct:.2f}% (รวมยอด inactive) · เป้าปี {team_ann:,.0f}" if team_pct is not None else None)
            _put_g("ทีมขาย", team_sub, team_avg, team_ann_need, team_q_target, team_q_sales, team_margin_all,
                   team_action, 'team')

            if team_ann > 0:
                proj = team_sales_all + L_rem * team_avg
                CTkLabel(outer, text=(f"ผู้จัดการและ CEO เห็นทุกแถว · ถ้าทีมทำตามอัตราเดิม สิ้นปีได้ {proj / team_ann * 100:.2f}% "
                                      f"ของเป้าปี {team_ann / 1e6:.1f} ล้าน (ขาด {max(team_ann - proj, 0) / 1e6:.2f} ล้านจาก 100%)"),
                         font=CTkFont(size=11), text_color="#64748B", anchor="w").pack(anchor="w", padx=20, pady=(0, 6))

        else:
            CTkLabel(outer, text="ตาราง \"เป้าหมายของแต่ละคน\" แสดงเมื่อเลือก \"ถึงรอบ\" ก่อนเดือนธันวาคม (ต้องมีเดือนที่เหลือให้คำนวณยอดที่ต้องทำ)",
                     font=CTkFont(size=11), text_color="#64748B", anchor="w").pack(anchor="w", padx=20, pady=(0, 6))

        # ══ Heatmap ระดับ รายเดือน · รายไตรมาส · รายปี (ทุกช่อง: บน = ยอดขาย, ล่าง = GP) ══
        try:
            hm_year = e_dt.year if e_dt else datetime.now().year
            hdf = self._get_heatmap_df(hm_year)
            hdf = hdf[~hdf['sale_key'].isin(EXCLUDE) & (hdf['sale_key'] != self.SALE_CENTER_KEY)].copy()
            if not hdf.empty:
                hdf['_group'] = hdf.apply(
                    lambda rr: MERGE[rr['sale_key']][0] if rr['sale_key'] in MERGE else rr['sale_name'], axis=1)
                hdf = self._merge_inactive_group(hdf)
            hm = {}                                   # (กลุ่ม, เดือน) → [ยอด, GP, ฐาน GP]
            for _, hr_ in hdf.iterrows():
                cell = hm.setdefault((hr_['_group'], int(hr_['month'])), [0.0, 0.0, 0.0])
                cell[0] += float(hr_['sales']); cell[1] += float(hr_['gp']); cell[2] += float(hr_['gp_base'])
        except Exception:
            traceback.print_exc()
            hm, hm_year = {}, (e_dt.year if e_dt else datetime.now().year)

        today_ = datetime.now()
        if hm_year < today_.year:
            last_closed = 12
        elif hm_year == today_.year:
            last_closed = today_.month - 1            # เดือนปัจจุบันยังไม่ปิดยอด (HR ยังไม่จ่ายค่าคอม) → "รอ"
        else:
            last_closed = 0

        hm_rows = [(r_['name'], (r_['annual'] / 12.0) if r_.get('annual') else (r_['target'] / max(N, 1)), 'person')
                   for r_ in rows]
        inactive_name = inactive_rows[0]['name'] if inactive_rows else None

        def _sum_cells(names, months):
            t = [0.0, 0.0, 0.0]
            for nm in names:
                for mm in months:
                    c_ = hm.get((nm, mm))
                    if c_:
                        t[0] += c_[0]; t[1] += c_[1]; t[2] += c_[2]
            return t

        hm_card = CTkFrame(outer, fg_color="white", corner_radius=12, border_width=1, border_color="#E1E6ED")
        hm_card.pack(fill="x", padx=16, pady=(8, 8))
        CTkLabel(hm_card, text="Heatmap ระดับ รายเดือน · รายไตรมาส · รายปี", font=CTkFont(size=15, weight="bold"),
                 text_color="#0F172A").pack(anchor="w", padx=16, pady=(12, 0))
        CTkLabel(hm_card, text=f"ยอดและ GP รายเดือนจริง ปี {hm_year + 543} เทียบเป้ารายเดือน · เดือนปัจจุบันที่ยังไม่ปิดยอดแสดง \"รอ\" · "
                               "GP รายเดือนของแถวทีม = ทีม active + inactive",
                 font=CTkFont(size=11), text_color="#64748B").pack(anchor="w", padx=16, pady=(0, 6))
        leg = CTkFrame(hm_card, fg_color="transparent")
        leg.pack(anchor="w", padx=16, pady=(0, 6))
        for lv_ in (5, 4, 3, 2, 1):
            CTkLabel(leg, text=f"  {lv_} {LV_NAME[lv_]}  ", font=CTkFont(size=11, weight="bold"), text_color=LV_FG[lv_],
                     fg_color=LV_BG[lv_], corner_radius=6, height=22).pack(side="left", padx=3)
        CTkLabel(leg, text="  ·  กรอบแดง = ยอดดี (≥ 85%) แต่กำไรต่ำ (GP < 13%)  ·  กรอบเข้ม = รวมไตรมาส/ทั้งปี  ·  Incentive = ยอด ≥ 101% และ GP ≥ 15.00%",
                 font=CTkFont(size=11), text_color="#64748B").pack(side="left", padx=6)

        hm_wrap = tk.Frame(hm_card, bg="white")
        hm_wrap.pack(fill="x", padx=12, pady=(0, 10))
        # วาดทั้ง Heatmap บน Canvas ใบเดียว (เดิมใช้ widget ~900 ชิ้น ทำให้หน้าโหลดช้ามาก)
        CW, GAPX, NAME_W = 96, 4, 175
        RH, TOP_H = 78, 62
        f_b = _cfnt(14, True)
        f_s = _cfnt(10)
        f_sb = _cfnt(10, True)
        f_h = _cfnt(12, True)
        Q_BG = "#EAF0F7"

        col_map = []        # (kind, ref)   kind: 'm' เดือน / 'q' ไตรมาส / 'y' ปี
        for qn in range(4):
            for mm in (3 * qn + 1, 3 * qn + 2, 3 * qn + 3):
                col_map.append(('m', mm))
            col_map.append(('q', qn))
        col_map.append(('y', 0))
        n_rows = len(hm_rows) + (1 if inactive_name else 0) + 1
        total_w = NAME_W + len(col_map) * (CW + GAPX) + 6
        total_h = TOP_H + n_rows * RH + 8
        hm_cv = tk.Canvas(hm_wrap, bg="white", highlightthickness=0, height=min(total_h, 600), width=10,
                          scrollregion=(0, 0, total_w, total_h))
        hm_sb = ttk.Scrollbar(hm_wrap, orient="horizontal", command=hm_cv.xview)
        hm_cv.configure(xscrollcommand=hm_sb.set)
        hm_cv.pack(fill="x")
        hm_sb.pack(fill="x")

        def _cx(ci):
            return NAME_W + ci * (CW + GAPX)

        # หัวคอลัมน์
        hm_cv.create_text(4, TOP_H - 12, text="พนักงาน", font=f_h, fill="#1F4E79", anchor="sw")
        for qn in range(4):
            x1 = _cx(qn * 4)
            x2 = _cx(qn * 4 + 3) + CW
            hm_cv.create_text((x1 + x2) / 2, 18, text=f"ไตรมาส {qn + 1}", font=f_h, fill="#8A5A00" if qn == 3 else "#1E293B")
            hm_cv.create_text((x1 + x2) / 2, 36, text=f"{SHORT_M[3 * qn]}–{SHORT_M[3 * qn + 2]}", font=f_s, fill="#64748B")
        yx1 = _cx(16)
        hm_cv.create_text(yx1 + CW / 2, 18, text=f"ทั้งปี {hm_year + 543}", font=f_h, fill="#1F4E79")
        hm_cv.create_text(yx1 + CW / 2, 36, text="สะสม", font=f_s, fill="#64748B")
        for ci, (kind, ref) in enumerate(col_map):
            x = _cx(ci)
            if kind == 'm':
                hm_cv.create_text(x + CW / 2, TOP_H - 10, text=SHORT_M[ref - 1], font=f_h, fill="#334155")
            else:
                hm_cv.create_rectangle(x, 44, x + CW, TOP_H - 2, fill=Q_BG, outline="")
                hm_cv.create_text(x + CW / 2, TOP_H - 12, text=(f"รวม Q{ref + 1}" if kind == 'q' else "ระดับปี"),
                                  font=f_h, fill="#1F4E79")
        # พื้นหลังจาง ๆ ของคอลัมน์รวมไตรมาส/ปี ตลอดแนวแถว
        for ci, (kind, ref) in enumerate(col_map):
            if kind in ('q', 'y'):
                x = _cx(ci)
                hm_cv.create_rectangle(x, TOP_H, x + CW, total_h - 4, fill="#F4F7FB", outline="")

        def _draw_block(x, y, w_, h_, bg, text, fnt, fg):
            hm_cv.create_rectangle(x, y, x + w_, y + h_, fill=bg, outline="")
            hm_cv.create_text(x + w_ / 2, y + h_ / 2, text=text, font=fnt, fill=fg)

        def _draw_cell(x, y, sales, target, gp, gp_base, is_team, closed, kind='m', incentive=True):
            """ช่องเดียวของ Heatmap: บน = ระดับยอดขาย + % ต่อเป้า, ล่าง = GP% + ระดับ GP (+ สถานะ Incentive)"""
            if not closed:
                hm_cv.create_rectangle(x, y, x + CW, y + 40, fill="#F1F5F9", outline="#CBD5E1", dash=(3, 2))
                hm_cv.create_text(x + CW / 2, y + 20, text="รอ", font=f_s, fill="#94A3B8")
                return
            h_total = 0
            if target > 0:
                pct = (sales / target * 100) if sales > 0 else 0.0
                lv = _sales_level(pct, is_team=is_team)
                hm_cv.create_rectangle(x, y, x + CW, y + 38, fill=LV_BG[lv], outline="")
                hm_cv.create_text(x + CW / 2, y + 13, text=str(lv), font=f_b, fill=LV_FG[lv])
                hm_cv.create_text(x + CW / 2, y + 29, text=f"{pct:.1f}%", font=f_s, fill=LV_FG[lv])
            else:
                pct = None
                _draw_block(x, y, CW, 38, "#EEF0F3", _num(sales) if sales else "—", f_sb, "#475569")
            h_total = 38
            gp_pct = (gp / gp_base * 100) if gp_base > 0 else None
            if gp_pct is not None:
                glv = _gp_level(gp_pct)
                _draw_block(x, y + 38, CW, 18, LV_BG[glv], f"GP {gp_pct:.1f}% · {glv}", f_s, LV_FG[glv])
                h_total = 56
                if kind == 'q' and incentive and pct is not None:
                    gp_ok = round(gp_pct, 2) >= 15.0
                    p2 = round(pct, 2)
                    if p2 >= 115.0 and gp_ok:
                        _draw_block(x, y + 56, CW, 16, "#1E7B4C", "ได้ INCENTIVE 2", f_sb, "#FFFFFF"); h_total = 72
                    elif p2 >= 101.0 and gp_ok:
                        _draw_block(x, y + 56, CW, 16, "#1E7B4C", "ได้ INCENTIVE 1", f_sb, "#FFFFFF"); h_total = 72
                    elif p2 >= 101.0:
                        _draw_block(x, y + 56, CW, 16, "#B42318", "ไม่ได้ · GP < 15%", f_sb, "#FFFFFF"); h_total = 72
                    else:
                        _draw_block(x, y + 56, CW, 16, "#FFFFFF", "ไม่ได้ · ยอด < 101%", f_s, "#64748B"); h_total = 72
            else:
                _draw_block(x, y + 38, CW, 18, "#EEF0F3", "GP —", f_s, "#64748B")
                h_total = 56
            if kind in ('q', 'y'):
                hm_cv.create_rectangle(x, y, x + CW, y + h_total, outline="#1F4E79", width=2)
            elif pct is not None and gp_pct is not None and round(pct, 2) >= PASS_PCT and round(gp_pct, 2) < 13.0:
                hm_cv.create_rectangle(x, y, x + CW, y + h_total, outline="#B42318", width=2)

        def _hm_row(row_i, label, sub, names, monthly_target, is_team=False, show_incentive=True, inactive_style=False):
            y = TOP_H + row_i * RH + 4
            if is_team:
                hm_cv.create_rectangle(0, y - 4, NAME_W - 6, y + RH - 6, fill="#F1F5FA", outline="")
            hm_cv.create_text(4, y + 10, text=label, font=f_h, fill="#1F4E79" if is_team else "#1E293B", anchor="w")
            if sub:
                hm_cv.create_text(4, y + 28, text=sub, font=f_s, fill="#B7791F", anchor="w", width=NAME_W - 12)
            for ci, (kind, ref) in enumerate(col_map):
                x = _cx(ci)
                if kind == 'm':
                    months, closed = [ref], ref <= last_closed
                elif kind == 'q':
                    months = [3 * ref + 1, 3 * ref + 2, 3 * ref + 3]
                    closed = months[0] <= last_closed
                else:
                    months = list(range(1, last_closed + 1)) or [1]
                    closed = last_closed > 0
                use_months = [m_ for m_ in months if m_ <= last_closed] if kind in ('q', 'y') else months
                sales_, gp_, gpb_ = _sum_cells(names, use_months)
                target_ = monthly_target * len(use_months)
                if inactive_style:
                    if closed and sales_ > 0:
                        txt = _num(sales_) + (f"\nGP {gp_ / gpb_ * 100:.1f}%" if gpb_ > 0 else "")
                        hm_cv.create_rectangle(x, y, x + CW, y + 38, fill="white", outline="#E1E6ED")
                        hm_cv.create_text(x + CW / 2, y + 19, text=txt, font=f_s, fill="#1E293B", justify="center")
                    elif closed:
                        hm_cv.create_rectangle(x, y, x + CW, y + 38, fill="white", outline="#E1E6ED")
                        hm_cv.create_text(x + CW / 2, y + 19, text="—", font=f_s, fill="#94A3B8")
                    else:
                        hm_cv.create_rectangle(x, y, x + CW, y + 38, fill="#F1F5F9", outline="#CBD5E1", dash=(3, 2))
                        hm_cv.create_text(x + CW / 2, y + 19, text="รอ", font=f_s, fill="#94A3B8")
                    continue
                _draw_cell(x, y, sales_, target_, gp_, gpb_, is_team, closed, kind,
                           incentive=(show_incentive and kind == 'q' and months[-1] <= last_closed))

        ri_ = 0
        for nm, mt, _k in hm_rows:
            _hm_row(ri_, nm, "เกณฑ์เซลล์", [nm], mt)
            ri_ += 1
        if inactive_name:
            _hm_row(ri_, "ยอดขาย Inactive", "ไม่มีเป้า · นับรวมในทีม", [inactive_name], 0.0, inactive_style=True)
            ri_ += 1
        team_names = [nm for nm, _m, _k in hm_rows] + ([inactive_name] if inactive_name else [])
        team_month_target = sum(mt for _n, mt, _k in hm_rows)
        _hm_row(ri_, "ทีมขาย", "เกณฑ์ผู้จัดการ · รวมยอด inactive", team_names, team_month_target,
                is_team=True, show_incentive=False)
        CTkLabel(hm_card, text="", height=4).pack()

        # ══ ส่วนผสมบิล 6 ระดับ (Tier Guest–T5) — ยอดสุทธิก่อน VAT ต่อบิล ══
        try:
            TIERS = [("Guest", "< 10,000", 0.0, 10_000.0), ("T1", "10,000 – 29,999", 10_000.0, 30_000.0),
                     ("T2", "30,000 – 99,999", 30_000.0, 100_000.0), ("T3", "100,000 – 299,999", 100_000.0, 300_000.0),
                     ("T4", "300,000 – 999,999", 300_000.0, 1_000_000.0), ("T5", "≥ 1,000,000", 1_000_000.0, float("inf"))]
            sdt, edt = self.custom_target_start, self.custom_target_end
            bills_df = pd.read_sql_query("""
                SELECT c.so_number, c.sale_key, c.status, COALESCE(c.sales_service_amount, 0) AS amt, pl.project_id
                FROM   commissions c
                LEFT JOIN project_lots pl ON pl.so_number = c.so_number
                WHERE  c.is_active = 1 AND c.bill_date::date BETWEEN %s AND %s
                  AND  COALESCE(c.so_number, '') NOT ILIKE 'TEST%%'
            """, self.pg_engine, params=(sdt.date(), edt.date()))
            bills_df = bills_df[~bills_df['status'].isin(('Draft', 'Cancelled', 'Cancelled by PU', 'Rejected by SM', 'Rejected by HR'))]
            bills_df = bills_df[~bills_df['sale_key'].isin(EXCLUDE) & (bills_df['sale_key'] != self.SALE_CENTER_KEY)]
            # งาน Project ที่แบ่งส่งหลายล็อต รวมทุกล็อตเป็นดีลเดียวก่อนแบ่งระดับ
            bills_df['_deal'] = bills_df.apply(
                lambda rr: f"P{int(rr['project_id'])}" if pd.notna(rr['project_id']) else rr['so_number'], axis=1)
            deals = bills_df.groupby('_deal').agg(amt=('amt', 'sum'), sale_key=('sale_key', 'first')).reset_index()
            deals = deals[deals['amt'] > 0]

            def _tier_of(v):
                for ti_, (_n, _r, lo_, hi_) in enumerate(TIERS):
                    if lo_ <= v < hi_:
                        return ti_
                return len(TIERS) - 1
            deals['_tier'] = deals['amt'].apply(_tier_of)
            n_total = len(deals)
            amt_total = float(deals['amt'].sum())

            # ข้อมูล QT จาก Express (นำเข้าด้วยปุ่ม "นำเข้า QT") → Conversion จริง ถ้ายังไม่เคยนำเข้าแสดง "—"
            conv, qt_info = None, None
            try:
                qdf = pd.read_sql_query(
                    "SELECT qt_number, qt_date, customer_code, seat_code, value_ex_vat FROM qt_header "
                    "WHERE qt_date BETWEEN %s AND %s", self.pg_engine, params=(sdt.date(), edt.date()))
                for c_ in ('value_ex_vat',):
                    qdf[c_] = pd.to_numeric(qdf[c_], errors='coerce').fillna(0.0)
                if not qdf.empty:
                    conv = qcv.compute_conversion(qdf, bills_df.rename(columns={'_deal': 'deal'})[['so_number', 'amt', 'deal']])
                    qt_info = pd.read_sql_query(
                        "SELECT filename, imported_at, date_to FROM qt_import_log ORDER BY id DESC LIMIT 1", self.pg_engine)
            except Exception:
                conv = None

            mix_card = CTkFrame(outer, fg_color="white", corner_radius=12, border_width=1, border_color="#E1E6ED")
            mix_card.pack(fill="x", padx=16, pady=(8, 8))
            ttl = CTkFrame(mix_card, fg_color="transparent")
            ttl.pack(fill="x", padx=16, pady=(12, 0))
            CTkLabel(ttl, text=f"Conversion ตามระดับบิล (ทั้งบริษัท {range_short} {e_dt.year + 543 if e_dt else ''})", font=CTkFont(size=15, weight="bold"),
                     text_color="#0F172A").pack(side="left")
            if conv:
                try:
                    inf = qt_info.iloc[0]
                    note_txt = (f"QT จากไฟล์ Express ({inf['filename']}) นำเข้าล่าสุด {pd.to_datetime(inf['imported_at']):%d/%m/%Y %H:%M} · "
                                "QT รวมจัดระดับตามมูลค่า QT (QT ฉบับแก้ไขนับเป็นดีลเดียว) · ปิดได้ / ยอด SO จัดระดับตามมูลค่า SO ที่จับคู่ด้วยเลขเอกสาร "
                                "(SO ที่ไม่พบ QT ในไฟล์นับเป็นปิดได้) · ยอดสุทธิก่อน VAT ไม่รวมค่ารถ ค่าตัด ค่าเจาะ · งาน Project รวมทุกล็อตเป็นดีลเดียว")
                except Exception:
                    note_txt = "QT จากไฟล์ Express · จับคู่ QT↔SO ด้วยเลขเอกสาร"
            else:
                note_txt = ("ปิดได้ / ยอด SO / % ยอด SO = ข้อมูลจริงจากระบบ APLUS Smart (ยอดสุทธิก่อน VAT ไม่รวมค่ารถ ค่าตัด ค่าเจาะ · นับตามวันที่เอกสาร · "
                            "งาน Project รวมทุกล็อตเป็นดีลเดียว) · QT รวม และ Conversion จะแสดงเมื่อนำเข้าไฟล์ QT จาก Express ด้วยปุ่ม \"📥 นำเข้า QT\" ด้านบน")
            CTkLabel(mix_card, text=note_txt,
                     font=CTkFont(size=11), text_color="#64748B", anchor="w", justify="left", wraplength=1150
                     ).pack(anchor="w", padx=16, pady=(0, 8))
            mix_body = CTkFrame(mix_card, fg_color="transparent")
            mix_body.pack(fill="x", padx=12, pady=(0, 12))
            mix_body.grid_columnconfigure(0, weight=3)
            mix_body.grid_columnconfigure(1, weight=1)

            # ── ตารางซ้าย: Conversion ตามระดับบิล (รูปแบบใหม่ตาม PM) ──
            # ปิดได้ = จำนวน SO จริงในระบบ (ทุก SO ต้องมี QT) · ยอด SO = ยอดสุทธิจาก SO
            # QT รวม / Conversion ต้องมีข้อมูลใบเสนอราคาจาก Express (Phase 2) จึงแสดง "—" ไว้ก่อน ไม่ใส่ตัวเลขสมมติ
            left_t = CTkFrame(mix_body, fg_color="transparent")
            left_t.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            heads_t = ["ระดับ", "ยอด/ใบ", "QT\nรวม", "ปิดได้", "Conversion\nจำนวน", "Conversion\nมูลค่า", "ยอด SO", "% ยอด\nSO"]
            wts_t = [1, 3, 1, 1, 2, 2, 2, 1]
            for ci_, (h_, wt_) in enumerate(zip(heads_t, wts_t)):
                left_t.grid_columnconfigure(ci_, weight=wt_)
                hf_ = CTkFrame(left_t, fg_color="#F3F4F6", corner_radius=0)
                hf_.grid(row=0, column=ci_, sticky="nsew")
                CTkLabel(hf_, text=h_, font=CTkFont(size=12, weight="bold"), text_color="#1E293B", justify="left" if ci_ in (0, 1) else "right",
                         anchor="w" if ci_ in (0, 1) else "e").pack(fill="x", padx=8, pady=8)
            conv_so_total = sum(c_['so_value'] for c_ in conv) if conv else 0.0
            for ti_, (nm_, rg_, _lo, _hi) in enumerate(TIERS):
                sub_ = deals[deals['_tier'] == ti_]
                cnt_, amt_ = len(sub_), float(sub_['amt'].sum())
                p_amt = amt_ / amt_total * 100 if amt_total else 0.0
                bg_ = "#F1F8F4" if nm_ == "T3" else "white"
                bold_ = nm_ == "T3"
                if conv:
                    cv_ = conv[ti_]
                    so_v = cv_['so_value']
                    vals_ = [nm_, rg_, f"{cv_['qt']:,}", f"{cv_['closed']:,}",
                             f"{cv_['closed'] / cv_['qt'] * 100:.1f}%" if cv_['qt'] else "—",
                             f"{so_v / cv_['qt_value'] * 100:.1f}%" if cv_['qt_value'] else "—",
                             f"{so_v / 1e6:.1f} ล้าน", f"{so_v / conv_so_total * 100:.1f}%" if conv_so_total else "—"]
                else:
                    vals_ = [nm_, rg_, "—", f"{cnt_:,}", "—", "—", f"{amt_ / 1e6:.1f} ล้าน", f"{p_amt:.1f}%"]
                for ci_, txt_ in enumerate(vals_):
                    f_ = CTkFrame(left_t, fg_color=bg_, corner_radius=0)
                    f_.grid(row=ti_ + 1, column=ci_, sticky="nsew")
                    CTkLabel(f_, text=txt_, font=CTkFont(size=13, weight="bold" if bold_ else "normal"),
                             text_color="#94A3B8" if txt_ == "—" else "#1E293B",
                             anchor="w" if ci_ in (0, 1) else "e").pack(fill="x", padx=8, pady=8)
            if conv:
                tq_ = sum(c_['qt'] for c_ in conv); tc_ = sum(c_['closed'] for c_ in conv)
                tqv_ = sum(c_['qt_value'] for c_ in conv)
                tot_vals = ["รวม", "", f"{tq_:,}", f"{tc_:,}", f"{tc_ / tq_ * 100:.1f}%" if tq_ else "—",
                            f"{conv_so_total / tqv_ * 100:.1f}%" if tqv_ else "—", f"{conv_so_total / 1e6:.1f} ล้าน", "100%"]
            else:
                tot_vals = ["รวม", "", "—", f"{n_total:,}", "—", "—", f"{amt_total / 1e6:.1f} ล้าน", "100%"]
            for ci_, txt_ in enumerate(tot_vals):
                f_ = CTkFrame(left_t, fg_color="#F1F5FA", corner_radius=0)
                f_.grid(row=7, column=ci_, sticky="nsew")
                CTkLabel(f_, text=txt_, font=CTkFont(size=13, weight="bold"), text_color="#94A3B8" if txt_ == "—" else "#1F4E79",
                         anchor="w" if ci_ in (0, 1) else "e").pack(fill="x", padx=8, pady=8)

            # ── แผงขวา: บิล T3 ต่อเดือน ต่อคน ──
            right_p = CTkFrame(mix_body, fg_color="white", corner_radius=10, border_width=1, border_color="#E1E6ED")
            right_p.grid(row=0, column=1, sticky="nsew")
            CTkLabel(right_p, text="บิล T3 ต่อเดือน ต่อคน", font=CTkFont(size=14, weight="bold"), text_color="#0F172A"
                     ).pack(anchor="w", padx=14, pady=(12, 2))
            t3_target = 52.0 / 12.0
            CTkLabel(right_p, text=f"เป้า: อย่างน้อย 1 บิล/สัปดาห์ (~{t3_target:.1f}/เดือน) · เส้นดำ = เป้า",
                     font=CTkFont(size=11), text_color="#64748B", wraplength=330, justify="left").pack(anchor="w", padx=14, pady=(0, 6))
            seat_of_key = {}
            for r_ in rows:
                seat_of_key[r_['name']] = r_['name']
            # แปลง sale_key → ชื่อที่นั่ง (รวมที่นั่งร่วม) ด้วยข้อมูลพนักงาน Active
            ppl = pd.read_sql_query(
                "SELECT sale_key, sale_name FROM sales_users WHERE role = 'Sale' AND status = 'Active'", self.pg_engine)
            key2seat = {str(p_['sale_key']).lower().replace(' ', ''):
                        (MERGE[p_['sale_key']][0] if p_['sale_key'] in MERGE else p_['sale_name'])
                        for _, p_ in ppl.iterrows() if p_['sale_key'] not in EXCLUDE}
            t3_deals = deals[deals['_tier'] == 3].copy()
            t3_deals['_seat'] = t3_deals['sale_key'].map(lambda k_: key2seat.get(str(k_).lower().replace(' ', '')))
            months_n = max(N, 1)
            t3_rate = []
            for r_ in rows:
                cnt_ = int((t3_deals['_seat'] == r_['name']).sum())
                t3_rate.append((r_['name'], cnt_ / months_n))
            t3_rate.sort(key=lambda t_: -t_[1])
            for nm_, rate_ in t3_rate:
                ratio_ = rate_ / t3_target if t3_target else 0
                col_ = "#1E7B4C" if ratio_ >= 1 else ("#8A5A00" if ratio_ >= 0.7 else "#B42318")
                rw_ = CTkFrame(right_p, fg_color="transparent")
                rw_.pack(fill="x", padx=14, pady=3)
                CTkLabel(rw_, text=nm_, font=CTkFont(size=12), text_color="#1E293B", width=120, anchor="w").pack(side="left")
                cvb_ = tk.Canvas(rw_, width=140, height=12, highlightthickness=0, bg="white")
                cvb_.create_rectangle(0, 2, 140, 10, fill="#E9EDF2", outline="")
                scale_ = 140 / (t3_target * 1.5)
                cvb_.create_rectangle(0, 2, max(2, min(rate_ * scale_, 140)), 10, fill=col_, outline="")
                mxl_ = t3_target * scale_
                cvb_.create_line(mxl_, 0, mxl_, 12, fill="#111827", width=2)
                cvb_.pack(side="left", padx=6)
                CTkLabel(rw_, text=f"{rate_:.1f}", font=CTkFont(size=13, weight="bold"), text_color=col_, width=36, anchor="e"
                         ).pack(side="right")
            hi_amt = float(deals[deals['_tier'] >= 3]['amt'].sum())
            hi_cnt = int((deals['_tier'] >= 3).sum())
            t3_avg = float(deals[deals['_tier'] == 3]['amt'].mean()) if (deals['_tier'] == 3).any() else 0.0
            CTkLabel(right_p, text=(f"บิล ≥ 100,000 = {hi_cnt / n_total * 100 if n_total else 0:.1f}% ของบิล แต่ "
                                    f"{hi_amt / amt_total * 100 if amt_total else 0:.0f}% ของยอด · +1 บิล T3/สัปดาห์/คน ≈ "
                                    f"+{t3_avg * 52 / 12 * max(len(rows), 1) / 1e6:.1f} ล้าน/เดือน"),
                     font=CTkFont(size=11), text_color="#64748B", wraplength=330, justify="left").pack(anchor="w", padx=14, pady=(8, 12))
        except Exception:
            traceback.print_exc()

        # ══ กล่องแจ้งเตือน + การ์ดทีม 4 ใบ (เติมหลังคำนวณ Heatmap เพราะใช้ข้อมูลรายเดือน) ══
        try:
            team_pct_all = (team_sales_all / total_t * 100) if total_t > 0 else None
            mon_closed = last_closed if 1 <= last_closed <= 12 else 0
            mon_txt = SHORT_M[mon_closed - 1] if mon_closed else ""
            active_names = [nm for nm, _m, _k in hm_rows]

            # ---------- การ์ดทีม ----------
            def _team_card(col, title):
                c = CTkFrame(cards_holder, fg_color="white", corner_radius=10, border_width=1, border_color="#E1E6ED")
                c.grid(row=0, column=col, sticky="nsew", padx=5)
                CTkLabel(c, text=title, font=CTkFont(size=11), text_color="#64748B", wraplength=300,
                         justify="left").pack(anchor="w", padx=14, pady=(10, 0))
                return c

            def _card_sub(c, text):
                CTkLabel(c, text=text, font=CTkFont(size=11), text_color="#64748B", wraplength=300, justify="left"
                         ).pack(anchor="w", padx=14, pady=(2, 10))

            c1 = _team_card(0, f"ยอดทีม vs เป้า ({range_short})")
            if team_pct_all is not None:
                lv1 = _sales_level(team_pct_all, is_team=True)
                row1 = CTkFrame(c1, fg_color="transparent")
                row1.pack(anchor="w", padx=14)
                CTkLabel(row1, text=f"{team_pct_all:.2f}%", font=CTkFont(size=26, weight="bold"),
                         text_color="#B42318" if team_pct_all < PASS_PCT else "#1E7B4C").pack(side="left")
                _pill(row1, f"  ระดับ {lv1} · {'ผ่าน' if round(team_pct_all, 2) >= PASS_PCT else 'ไม่ผ่าน'}  ",
                      LV_BG[lv1], LV_FG[lv1]).pack(side="left", padx=10)
                gap100 = max(total_t - team_sales_all, 0.0)
                gap85 = max(total_t * PASS_PCT / 100 - team_sales_all, 0.0)
                _card_sub(c1, f"{team_sales_all:,.0f} ÷ {total_t:,.0f} · ขาดอีก {gap100:,.0f} บาทจะถึง 100%"
                              + (f" (ถึงเกณฑ์ KPI 85% ขาด {gap85:,.0f})" if gap85 > 0 else " (ผ่านเกณฑ์ KPI 85% แล้ว)"))
            else:
                CTkLabel(c1, text="—", font=CTkFont(size=26, weight="bold")).pack(anchor="w", padx=14)
                _card_sub(c1, "ยังไม่มีเป้า")

            c2 = _team_card(1, f"GP ทีม ({range_short})")
            if team_margin_all and team_margin_all > 0:
                lv2 = _gp_level(team_margin_all)
                row2 = CTkFrame(c2, fg_color="transparent")
                row2.pack(anchor="w", padx=14)
                CTkLabel(row2, text=f"{team_margin_all:.2f}%", font=CTkFont(size=26, weight="bold"),
                         text_color="#1E7B4C" if lv2 >= 4 else "#B42318").pack(side="left")
                _pill(row2, f"  ระดับ {lv2} · {LV_NAME[lv2]}  ", LV_BG[lv2], LV_FG[lv2]).pack(side="left", padx=10)
                nxt = {1: 12.0, 2: 13.0, 3: 15.0, 4: 20.0}.get(lv2)
                _card_sub(c2, f"อีก {nxt - team_margin_all + 0.005:.2f} pt ถึงระดับ {lv2 + 1} (> {nxt:.2f}%)"
                          if nxt else "ระดับสูงสุดแล้ว")
            else:
                CTkLabel(c2, text="—", font=CTkFont(size=26, weight="bold")).pack(anchor="w", padx=14)
                _card_sub(c2, "ยังไม่มีข้อมูล GP")

            c3 = _team_card(2, f"ต้องทำต่อเดือน {rest_short} (เป้าปี {team_calc['annual'] / 1e6:,.1f} ล้าน)")
            if L > 0 and team_calc['annual'] > 0:
                need3 = max(team_calc['annual'] - team_sales_all, 0.0) / L
                avg3 = total_s / N
                inc3 = (need3 / avg3 - 1) * 100 if avg3 > 0 else None
                CTkLabel(c3, text=f"{need3 / 1e6:.2f} ล้าน", font=CTkFont(size=26, weight="bold"),
                         text_color="#B42318" if (inc3 or 0) > 30 else "#8A5A00").pack(anchor="w", padx=14)
                _card_sub(c3, f"เพื่อถึง 100% เป้าปี · ทำได้เฉลี่ย {avg3 / 1e6:.2f} ล้าน/เดือน"
                              + (f" ({inc3:+.1f}%)" if inc3 is not None else ""))
            else:
                CTkLabel(c3, text="—", font=CTkFont(size=26, weight="bold")).pack(anchor="w", padx=14)
                _card_sub(c3, "ครบปีแล้ว" if L == 0 else "ยังไม่มีเป้า")

            c4 = _team_card(3, "เงื่อนไข Incentive รายไตรมาส (2 ขั้น)")
            CTkLabel(c4, text="ขั้นที่ 1: ยอด ≥ 101%  ·  ขั้นที่ 2: ยอด ≥ 115%", font=CTkFont(size=14, weight="bold"),
                     text_color="#0F172A", wraplength=300, justify="left").pack(anchor="w", padx=14, pady=(4, 0))
            _card_sub(c4, "ทั้ง 2 ขั้นต้องมี GP ≥ 15.00% · GP ต่ำกว่า 15% = ไม่ได้ · ไม่ยกยอดข้ามไตรมาส")

            # ---------- กล่องแจ้งเตือน ----------
            alerts = []        # (severity 0=แดง 1=เหลือง, หัวข้อ, รายละเอียด)
            for nm_, gp_ in [(r_['name'], r_['actual_margin']) for r_ in rows + inactive_rows]:
                if gp_ and 0 < gp_ < 10:
                    alerts.append((0, "GP วิกฤต (< 10%)", f"{nm_} · GP {gp_:.2f}%"))
            if team_pct_all is not None and team_pct_all < 100:
                alerts.append((0 if team_pct_all < PASS_PCT else 1, f"ทีมห่างเป้า 100% ({range_short})",
                               f"ทำได้ {team_pct_all:.2f}% · ขาดอีก {max(total_t - team_sales_all, 0):,.0f} บาท"))
            if mon_closed:
                ts_, tg_, tgb_ = _sum_cells(active_names, [mon_closed])
                if tgb_ > 0 and tg_ / tgb_ * 100 < 15.0:
                    alerts.append((0, f"GP ทีม {mon_txt} ต่ำกว่าจุดคุ้มทุน",
                                   f"ทีม active {mon_txt} GP {tg_ / tgb_ * 100:.2f}% (ยอด {ts_:,.0f})"))
                for nm_, mt_, _k in hm_rows:
                    s_, g_, gb_ = _sum_cells([nm_], [mon_closed])
                    if mt_ > 0 and gb_ > 0 and s_ / mt_ * 100 >= PASS_PCT and g_ / gb_ * 100 < 13.0:
                        alerts.append((0, "ยอดดีแต่กำไรต่ำ",
                                       f"{nm_} · {mon_txt} ยอด {s_ / mt_ * 100:.1f}% แต่ GP {g_ / gb_ * 100:.2f}%"))
            for r_, d_ in sale_calc:
                if d_['ytd_pct'] is not None and d_['ytd_pct'] < PASS_PCT and 0 < r_['actual_margin'] < 15.0:
                    alerts.append((0, "ไม่ผ่านเกณฑ์ + GP ต่ำกว่า 15%",
                                   f"{r_['name']} · ยอดสะสม {d_['ytd_pct']:.2f}% · GP {r_['actual_margin']:.2f}%"))
            for qn_ in range(4):
                if 3 * qn_ + 3 > last_closed:
                    continue
                for nm_, mt_, _k in hm_rows:
                    s_, g_, gb_ = _sum_cells([nm_], [3 * qn_ + 1, 3 * qn_ + 2, 3 * qn_ + 3])
                    if mt_ > 0 and gb_ > 0 and s_ / (mt_ * 3) * 100 >= 101.0 and g_ / gb_ * 100 < 15.0:
                        alerts.append((0, "ไม่ได้ Incentive เพราะ GP < 15%",
                                       f"{nm_} · Q{qn_ + 1} ยอด {s_ / (mt_ * 3) * 100:.1f}% แต่ GP {g_ / gb_ * 100:.2f}%"))
            alerts.sort(key=lambda a_: a_[0])
            if alerts:
                abox = CTkFrame(alerts_holder, fg_color="white", corner_radius=12, border_width=1, border_color="#F3C6C1")
                abox.pack(fill="x", padx=16, pady=(4, 8))
                CTkLabel(abox, text=f"⚠  แจ้งเตือน ({len(alerts)})", font=CTkFont(size=14, weight="bold"),
                         text_color="#B42318").pack(anchor="w", padx=16, pady=(10, 4))
                agrid = CTkFrame(abox, fg_color="transparent")
                agrid.pack(fill="x", padx=12, pady=(0, 10))
                for i_ in range(4):
                    agrid.grid_columnconfigure(i_, weight=1, uniform="al")
                for k_, (sev_, ttl_, det_) in enumerate(alerts):
                    ac = CTkFrame(agrid, fg_color="#FDE7E5" if sev_ == 0 else "#FFF1D6", corner_radius=8)
                    ac.grid(row=k_ // 4, column=k_ % 4, sticky="nsew", padx=4, pady=4)
                    CTkLabel(ac, text=ttl_, font=CTkFont(size=12, weight="bold"),
                             text_color="#B42318" if sev_ == 0 else "#8A5A00", anchor="w", justify="left",
                             wraplength=330).pack(anchor="w", padx=10, pady=(8, 0))
                    CTkLabel(ac, text=det_, font=CTkFont(size=12), text_color="#1E293B", anchor="w", justify="left",
                             wraplength=330).pack(anchor="w", padx=10, pady=(0, 8))
        except Exception:
            traceback.print_exc()

        # ── Footer: สูตรย่อ + หมายเหตุ ──
        notes = [
            "สูตร: % ถึงเป้า YTD = ยอดขายจริง ÷ เป้า YTD · เป้าทั้งปี = เป้าช่วงที่เลือก ÷ เดือนที่เลือก × 12 (กรณีเป้ารายเดือนเท่ากัน) · "
            "% ของเป้าปี = ยอดขายจริง ÷ เป้าทั้งปี · ยอดที่ยังขาด = เป้าทั้งปี − ยอดขายจริง · ต้องทำ/เดือน = ยอดที่ยังขาด ÷ เดือนที่เหลือ · "
            "สถานะ Pace เทียบ % ของเป้าปีกับ % เวลาที่ผ่านไป",
            "เป้ารวมบริษัทใช้เป้าบริษัทรายปีที่ตั้งไว้ (ไม่เท่ากับผลรวมเป้ารายคน) · เป้าหมายหลัก 100% · ผ่านเกณฑ์ KPI ≥ 85% = ระดับ 3 มาตรฐาน",
        ]
        if ct_rows:
            notes.append("* Sale Center (CT) ไม่มีเป้ายอดขาย · GP เป็นประมาณการ = ยอดขาย − (ต้นทุน PO × 1.03) "
                         "เฉพาะ SO ที่มี PO แล้ว ยังไม่ผ่านการตรวจของ HR")
        for t in notes:
            CTkLabel(outer, text=t, font=CTkFont(size=11), text_color="#64748B", anchor="w", justify="left",
                     wraplength=1100).pack(anchor="w", padx=20, pady=(0, 2))
        CTkLabel(outer, text="", height=8).pack()

        if _ysc_saved is not None:
            try:
                outer._parent_canvas.configure(yscrollcommand=_ysc_saved)
            except Exception:
                pass
        outer.pack(fill="both", expand=True, padx=10, pady=10)

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



class QTImportDialog(CTkToplevel):
    """นำเข้ารายงาน "ใบเสนอราคา แยกตามพนักงานขาย" จาก Express (ไฟล์ CSV ดิบ Windows-874) แล้วประมวลผลให้ทันที
    นำเข้าซ้ำ/ไฟล์ช่วงทับกันได้ (upsert ตามเลข QT) · ไม่เก็บชื่อลูกค้า/ผู้ติดต่อ/เบอร์โทร"""

    def __init__(self, master, app_container, on_done=None):
        super().__init__(master)
        self.app_container = app_container
        self.on_done = on_done
        self.df = None
        self.warnings_ = []
        self.path = None
        self.title("นำเข้า QT จาก Express")
        from project_screen import _center_and_style_popup
        _center_and_style_popup(self, master, 760, 560)
        self.transient(master.winfo_toplevel())
        self.after(50, self.grab_set)

        CTkLabel(self, text="นำเข้าใบเสนอราคา (QT) จาก Express", font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=18, pady=(14, 2))
        CTkLabel(self, text="เลือกไฟล์รายงาน \"ใบเสนอราคา แยกตามพนักงานขาย\" (CSV) ได้เลยโดยไม่ต้องแปลงไฟล์ — ระบบอ่านภาษาไทย (Windows-874) เอง "
                            "แล้วจับคู่กับ SO และคำนวณ Conversion ให้ทันที · ควรดึงรายงานด้วยเงื่อนไขกว้างสุด (ทุกรหัสลูกค้า ทุกสถานะ ทุกแผนก)",
                 font=CTkFont(size=12), text_color="#64748B", wraplength=720, justify="left").pack(anchor="w", padx=18)
        row = CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=10)
        CTkButton(row, text="📂 เลือกไฟล์ CSV", width=150, command=self._choose).pack(side="left")
        self.file_lbl = CTkLabel(row, text="ยังไม่ได้เลือกไฟล์", font=CTkFont(size=12), text_color="#64748B")
        self.file_lbl.pack(side="left", padx=10)
        self.box = CTkTextbox(self, height=300, font=CTkFont(size=12))
        self.box.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        self.box.configure(state="disabled")
        bottom = CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=18, pady=(0, 14))
        CTkButton(bottom, text="ปิด", width=90, fg_color="#94A3B8", hover_color="#64748B", command=self.destroy).pack(side="right")
        self.go_btn = CTkButton(bottom, text="✅ นำเข้าและประมวลผล", width=170, fg_color="#16A34A", hover_color="#15803D",
                                state="disabled", command=self._import)
        self.go_btn.pack(side="right", padx=8)

    def _say(self, text):
        self.box.configure(state="normal")
        self.box.delete("1.0", "end")
        self.box.insert("1.0", text)
        self.box.configure(state="disabled")

    def _choose(self):
        path = filedialog.askopenfilename(parent=self, title="เลือกไฟล์รายงาน QT (CSV)",
                                          filetypes=[("CSV", "*.csv"), ("ทุกไฟล์", "*.*")])
        if not path:
            return
        try:
            df, meta = qcv.parse_express_qt_csv(path)
        except Exception as ex:
            self.df = None
            self.go_btn.configure(state="disabled")
            self._say(f"อ่านไฟล์ไม่สำเร็จ: {ex}")
            return
        self.path = path
        self.file_lbl.configure(text=os.path.basename(path))
        if df.empty:
            self.df = None
            self.go_btn.configure(state="disabled")
            self._say("ไม่พบแถวใบเสนอราคา (QT) ในไฟล์นี้ — ตรวจสอบว่าเป็นรายงาน \"ใบเสนอราคา แยกตามพนักงานขาย\" ของ Express")
            return
        self.df = df
        self.warnings_ = qcv.coverage_warnings(df)
        existing = 0
        try:
            ex_df = pd.read_sql_query("SELECT count(*) AS n FROM qt_header WHERE qt_number = ANY(%s)",
                                      self.app_container.pg_engine, params=(df['qt_number'].tolist(),))
            existing = int(ex_df.iloc[0]['n'])
        except Exception:
            existing = 0           # ยังไม่เคยนำเข้า (ตารางยังไม่ถูกสร้าง)
        lines = [f"ไฟล์: {os.path.basename(path)}",
                 f"ใบเสนอราคาในไฟล์: {meta['rows']:,} ใบ  ·  ช่วงวันที่ {meta.get('date_from')} ถึง {meta.get('date_to')}",
                 f"มูลค่ารวมก่อน VAT: {df['value_ex_vat'].sum():,.0f} บาท  ·  มูลค่า 0: {(df['value_ex_vat'] <= 0).sum():,} ใบ",
                 f"มีในระบบแล้ว (จะอัปเดต): {existing:,} ใบ  ·  ใหม่: {meta['rows'] - existing:,} ใบ", "",
                 "จำนวนใบแยกตามที่นั่ง:"]
        for seat, cnt in df['seat_code'].value_counts().items():
            nm = df.loc[df['seat_code'] == seat, 'seat_name'].iloc[0]
            lines.append(f"   {seat}  {nm}: {cnt:,} ใบ")
        if self.warnings_:
            lines += ["", "⚠ เลข QT ที่อาจหายไป (ไฟล์อาจไม่ครบ — ตรวจเงื่อนไขการดึงรายงาน):"] + [f"   {w_}" for w_ in self.warnings_[:25]]
            if len(self.warnings_) > 25:
                lines.append(f"   ... และอีก {len(self.warnings_) - 25} รายการ")
        self._say("\n".join(lines))
        self.go_btn.configure(state="normal")

    def _import(self):
        if self.df is None:
            return
        if self.warnings_ and not messagebox.askyesno(
                "ไฟล์อาจไม่ครบ", f"พบเลข QT ที่อาจหายไป {len(self.warnings_)} รายการ\nต้องการนำเข้าต่อหรือไม่?", parent=self):
            return
        user = getattr(self.app_container, "current_user_key", None) or "-"
        conn = self.app_container.get_connection()
        try:
            new, upd, _imp = qcv.import_to_db(conn, self.df, os.path.basename(self.path or ""), user, self.warnings_)
        except Exception as ex:
            conn.rollback()
            messagebox.showerror("Error", f"นำเข้าไม่สำเร็จ: {ex}", parent=self)
            return
        finally:
            self.app_container.release_connection(conn)
        messagebox.showinfo("นำเข้าแล้ว", f"นำเข้า QT สำเร็จ — ใหม่ {new:,} ใบ · อัปเดต {upd:,} ใบ\nระบบประมวลผล Conversion ให้แล้ว", parent=self)
        cb = self.on_done
        self.destroy()
        if cb:
            cb()
