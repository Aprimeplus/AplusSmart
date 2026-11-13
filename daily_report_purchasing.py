# daily_report_purchasing.py
# Daily Report ใหม่ — มุมมอง "ฝ่ายจัดซื้อ" (กระดานงาน 3 คอลัมน์ + ภาระงานต่อผู้รับงาน)
# เฟส 1 ตามต้นแบบของ PM (มุมมองฝ่ายขายอยู่ใน daily_report_sales.py; ผู้บริหารเป็นเฟสถัดไป)
# วาดด้วย Canvas (canvas_board.py) ไม่ใช้วิดเจ็ตต่อการ์ด — เดิมแบบวิดเจ็ต 41 การ์ดโหลด ~35 วินาที ตอนนี้ไม่ถึงวินาที
#
# นิยามที่ใช้ (ปรับได้ที่ค่าคงที่ด้านล่าง / ให้ PM ยืนยัน):
#   คอลัมน์ "รอจัดซื้อรับงาน"   = Pending PU
#   คอลัมน์ "อยู่ระหว่างจัดซื้อ" = PO In Progress
#   คอลัมน์ "เปิด PO สำเร็จ"     = PO Sent (เฉพาะที่ยังไม่ถึง/ตรงวันส่งของวันที่รายงาน)
#   D-x = วันส่ง − วันที่รายงาน (D-0 ส่งวันนี้, D+5 เลยกำหนด 5 วัน)
#   งานเปิด (รอรับ/ระหว่างจัดซื้อ) ที่เลยกำหนดเกิน STALE_DAYS วัน ไม่แสดงบนกระดาน (บอกจำนวนไว้ท้ายหน้า)
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import date, datetime
import traceback

import pandas as pd
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkFont, CTkOptionMenu)

from custom_widgets import DateSelector
from canvas_board import BoardCanvas, clean_text

STALE_DAYS = 14                       # งานเปิดที่เลยกำหนดเกินกี่วันถึงจะซ่อน
AUTO_REFRESH_MS = 5 * 60 * 1000       # รีเฟรชอัตโนมัติทุก 5 นาที
UNASSIGNED = "ยังไม่มีผู้รับ"

COLS = [
    ("wait", "รอจัดซื้อรับงาน", "#B45309"),
    ("work", "อยู่ระหว่างจัดซื้อ", "#F0B25B"),
    ("mp", "PO รอ MP อนุมัติ", "#7C3AED"),
    ("done", "เปิด PO สำเร็จ", "#166534"),
]
STATUS_TO_COL = {"Pending PU": "wait", "PO In Progress": "work", "PO Sent": "done"}

SEV = {   # สีตามความใกล้วันส่ง: (แถบซ้าย, พื้นป้าย, ตัวอักษรป้าย)
    "red":    ("#C2410C", "#FEE2E2", "#B91C1C"),
    "orange": ("#EA7B1E", "#FFEDD5", "#C2410C"),
    "gray":   ("#CBD5E1", "#E5E7EB", "#4B5563"),
}

_SQL = """
SELECT c.so_number, c.customer_name, c.status, c.sale_key,
       COALESCE(c.date_to_customer, c.delivery_date) AS delivery,
       COALESCE(NULLIF(c.user_key, ''),
                (SELECT po.user_key FROM purchase_orders po
                  WHERE po.so_number = c.so_number
                    AND po.status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
                  ORDER BY po.id LIMIT 1)) AS pu,
       c.sales_service_amount, c.total_payment_amount, c.difference_amount,
       EXISTS (SELECT 1 FROM purchase_orders pp
                WHERE pp.so_number = c.so_number AND pp.status = 'Pending Approval') AS mp_pending
FROM commissions c
WHERE c.is_active = 1
  AND c.status IN ('Pending PU', 'PO In Progress', 'PO Sent')
  AND COALESCE(c.date_to_customer, c.delivery_date) IS NOT NULL
ORDER BY COALESCE(c.date_to_customer, c.delivery_date), c.so_number
"""


def _short_so(so):
    so = str(so or "")
    return so[6:] if so.startswith("SO") and len(so) > 6 else so


def _fmt_money_short(x):
    x = float(x or 0)
    return f"{x / 1_000_000:.2f} ล้าน" if abs(x) >= 1_000_000 else f"{x:,.0f}"


def _d_label(n):
    return f"D-{n}" if n >= 0 else f"D+{-n}"


def _severity(n):
    return "red" if n <= 0 else ("orange" if n <= 2 else "gray")


def load_board(engine, rep):
    """โหลดงานบนกระดานจัดซื้อ ณ วันที่รายงาน → (DataFrame ที่มีคอลัมน์ col/n/pu/pct/mp_pending, จำนวนงานค้างเก่าที่ซ่อน)
    ใช้ร่วมกับหน้าภาพรวมผู้บริหาร เพื่อให้ตัวเลขสถานะตรงกัน"""
    raw = pd.read_sql_query(_SQL, engine)
    raw["delivery"] = pd.to_datetime(raw["delivery"]).dt.date
    raw["col"] = raw["status"].map(STATUS_TO_COL)
    raw["col"] = raw.apply(lambda r: "mp" if r["mp_pending"] else r["col"], axis=1)   # มี PO ที่รอ MP อนุมัติ → คอลัมน์ "รอ MP อนุมัติ"
    raw["n"] = raw["delivery"].apply(lambda d: (d - rep).days)
    raw["pu"] = raw["pu"].apply(lambda v: v if isinstance(v, str) and v.strip() else None)
    for c in ("sales_service_amount", "total_payment_amount", "difference_amount"):
        raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0.0)
    raw["required"] = raw["total_payment_amount"] - raw["difference_amount"]   # ยอดที่ต้องชำระ = ชำระแล้ว − ส่วนต่าง (ติดลบ = ขาด)
    raw["pct"] = raw.apply(lambda r: (r["total_payment_amount"] / r["required"]) if r["required"] > 0.5 else None, axis=1)

    is_open = raw["col"].isin(["wait", "work", "mp"])
    stale = is_open & (raw["n"] < -STALE_DAYS)
    keep = (~stale) & ((raw["col"] != "done") | (raw["n"] >= 0))   # เปิด PO แล้ว: โชว์เฉพาะที่ยังไม่ผ่านวันส่ง
    df = raw[keep].reset_index(drop=True)
    return df, int(stale.sum())


class PurchasingDailyBoard(CTkFrame):
    def __init__(self, master, app_container, **kwargs):
        super().__init__(master, fg_color="#F3F4F6", **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.df = pd.DataFrame()
        self._stale_hidden = 0
        self._pu_staff_total = 0
        self._auto_date = date.today()          # วันที่ที่ตั้งให้อัตโนมัติ (เลื่อนเป็นวันนี้เองเมื่อเปิดค้างข้ามวัน)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ---------- header ----------
        head = CTkFrame(self, fg_color="white", corner_radius=10)
        head.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        CTkLabel(head, text="กระดานงานจัดซื้อ", font=CTkFont(size=18, weight="bold"),
                 text_color="#111827").pack(side="left", padx=(16, 8), pady=10)
        CTkLabel(head, text="Daily Report · ฝ่ายจัดซื้อ", font=CTkFont(size=12), text_color="#6B7280").pack(side="left")
        CTkButton(head, text="Export Excel", width=110, fg_color="white", text_color="#111827",
                  border_width=1, border_color="#D1D5DB", hover_color="#F3F4F6",
                  command=self.export_excel).pack(side="right", padx=(6, 14))
        CTkButton(head, text="🔄 ดึงข้อมูล", width=110, fg_color="#1D4ED8", hover_color="#1E40AF",
                  command=self.load_data).pack(side="right", padx=6)
        self.pu_var = tk.StringVar(value="ทั้งหมด")
        self.pu_menu = CTkOptionMenu(head, variable=self.pu_var, values=["ทั้งหมด"], width=170,
                                     command=lambda _v: self.render())
        self.pu_menu.pack(side="right", padx=6)
        CTkLabel(head, text="ผู้รับงาน:", text_color="#6B7280").pack(side="right")
        self.date_sel = DateSelector(head)
        self.date_sel.pack(side="right", padx=(0, 14))
        self.date_sel.set_date(date.today())
        CTkLabel(head, text="วันที่รายงาน:", text_color="#6B7280").pack(side="right")

        # ---------- canvas body ----------
        self.board = BoardCanvas(self, bg="#F3F4F6", on_resize=self.render)
        self.board.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 4))

        self.note_lbl = CTkLabel(self, text="", font=CTkFont(size=11), text_color="#6B7280", anchor="w")
        self.note_lbl.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 8))

        self.after(150, self.load_data)
        self._schedule_refresh()

    # ------------------------------------------------------------------ data
    def _report_date(self):
        s = self.date_sel.get_date()
        try:
            return datetime.strptime(s, "%Y-%m-%d").date() if s else date.today()
        except Exception:
            return date.today()

    def load_data(self, silent=False):
        try:
            try:
                staff = pd.read_sql_query(
                    "SELECT count(*) AS n FROM sales_users WHERE role = 'Purchasing Staff' AND status = 'Active'",
                    self.pg_engine)
                self._pu_staff_total = int(staff['n'].iloc[0])
            except Exception:
                self._pu_staff_total = 0

            self.df, self._stale_hidden = load_board(self.pg_engine, self._report_date())
            self.df["hint"] = self.df.apply(self._hint, axis=1)

            pus = sorted({p for p in self.df["pu"].dropna()})
            cur = self.pu_var.get()
            values = ["ทั้งหมด", UNASSIGNED] + pus
            self.pu_menu.configure(values=values)
            if cur not in values:
                self.pu_var.set("ทั้งหมด")
            self.render()
        except Exception as e:
            traceback.print_exc()
            if not silent:            # รีเฟรชอัตโนมัติไม่เด้งหน้าต่างซ้ำๆ ถ้าฐานข้อมูลสะดุด (คงข้อมูลเดิมไว้)
                messagebox.showerror("Error", f"ดึงข้อมูลกระดานงานจัดซื้อไม่สำเร็จ: {e}", parent=self)

    @staticmethod
    def _hint(r):
        n, col = int(r["n"]), r["col"]
        if col == "mp":
            return "เลยกำหนด" if n < 0 else ("ส่งวันนี้" if n == 0 else "รอ MP")
        if col != "done" and n < 0:
            return "เลยกำหนด"
        if n == 0:
            return "จัดรถวันนี้" if col == "done" else "ส่งวันนี้"
        pct = r["pct"]
        if r["required"] > 0.5 and pct is not None:
            if r["total_payment_amount"] <= 0.005 and n <= 2:
                return "ยังไม่ชำระ"
            if 0.005 < pct < 0.995:
                return f"มัดจำ {round(pct * 100)}%"
        return ""

    def _filtered(self):
        df = self.df
        v = self.pu_var.get()
        if v == "ทั้งหมด" or df.empty:
            return df
        if v == UNASSIGNED:
            return df[df["pu"].isna()]
        return df[df["pu"] == v]

    # ---------------------------------------------------------------- render
    def render(self):
        bc = self.board
        S = bc.S
        bc.clear()
        df = self._filtered()
        W = bc.width
        M, G = S(10), S(8)
        y = M

        # ---- การ์ดสรุป 5 ใบ ----
        def part(key):
            return df[df["col"] == key] if not df.empty else df

        stats = []
        for key, title, color in COLS:
            p = part(key)
            stats.append((("รอรับงาน" if key == "wait" else title), len(p),
                          _fmt_money_short(p["sales_service_amount"].sum() if len(p) else 0), color, False))
        if df.empty:
            today_no_po = overdue = df
        else:
            no_po = df["col"].isin(["wait", "work"])                      # ยังไม่มี PO (PO ที่รอ MP อนุมัติ/เปิดแล้ว ไม่นับ)
            today_no_po = df[no_po & (df["n"] == 0)]
            overdue = df[no_po & (df["n"] < 0)]
        sub4 = " · ".join(_short_so(s) for s in list(today_no_po["so_number"])[:4]) if len(today_no_po) else "ไม่มี"
        if len(today_no_po) > 4:
            sub4 += f" +{len(today_no_po) - 4}"
        sub5 = " · ".join(f"{_short_so(r.so_number)} เลย {-int(r.n)} วัน" for r in overdue.head(2).itertuples()) if len(overdue) else "ไม่มี"
        if len(overdue) > 2:
            sub5 += f" +{len(overdue) - 2}"
        stats.append(("ส่งวันนี้ ยังไม่มี PO", len(today_no_po), sub4, "#B91C1C", len(today_no_po) > 0))
        stats.append(("เลยวันส่ง", len(overdue), sub5, "#B91C1C", len(overdue) > 0))

        cw = (W - 2 * M - (len(stats) - 1) * G) / len(stats)
        ch = S(96)
        for i, (title, big, sub, color, alert) in enumerate(stats):
            x = M + i * (cw + G)
            bc.rrect(x, y, x + cw, y + ch, r=S(10), fill="#FEE2E2" if alert else "white",
                     outline="#FCA5A5" if alert else "#E5E7EB")
            bc.text(x + S(14), y + S(10), title, size=10, fill="#B91C1C" if alert else "#6B7280")
            bc.text(x + S(14), y + S(28), str(big), size=22, bold=True, fill=color)
            bc.text(x + S(14), y + S(68), sub, size=9, fill="#6B7280", width=cw - S(28))
        y += ch + G * 1.5

        # ---- แถบภาระงานต่อผู้รับงาน (แนวนอน อยู่เหนือกระดาน) + กระดานเต็มความกว้าง ----
        total_w = W - 2 * M
        y = self._draw_workload(M, y, total_w, self.df if self.pu_var.get() == "ทั้งหมด" else df) + G * 1.5
        cwid = (total_w - (len(COLS) - 1) * G) / len(COLS)
        bottoms = []
        for ci, (key, title, color) in enumerate(COLS):
            bottoms.append(self._draw_column(M + ci * (cwid + G), y, cwid, key, title, color,
                                             df[df["col"] == key] if not df.empty else df))
        bc.set_height(max(bottoms) + M)

        rep = self._report_date()
        txt = f"รายงาน ณ {rep.day:02d}/{rep.month:02d}/{rep.year + 543} · งานทั้งหมด {len(df)} ใบ"
        if self._stale_hidden:
            txt += (f"  |  ไม่แสดงงานเปิดที่เลยกำหนดเกิน {STALE_DAYS} วัน {self._stale_hidden} รายการ "
                    f"(ดูได้ที่แท็บ 'รายงานเดิม')")
        self.note_lbl.configure(text=txt)

    def _draw_column(self, x, y, w, key, title, color, rows):
        bc = self.board
        S = bc.S
        bg = bc.rrect(x, y, x + w, y + 10, r=S(10), fill="#EEF0F4")
        pad = S(10)
        bc.text(x + pad, y + S(10), "■", size=11, fill=color)
        bc.text(x + pad + S(18), y + S(9), title, size=12, bold=True)
        total = rows["sales_service_amount"].sum() if len(rows) else 0
        bc.text(x + w - pad, y + S(10), f"{len(rows)} ใบ · {_fmt_money_short(total)}", size=9, bold=True, fill="#4B5563", anchor="ne")
        cy = y + S(40)
        if rows is None or len(rows) == 0:
            bc.text(x + w / 2, cy + S(14), "— ไม่มีงาน —", size=10, fill="#9CA3AF", anchor="n")
            cy += S(50)
        else:
            for r in rows.sort_values(["n", "so_number"]).itertuples():
                cy += self._draw_job_card(x + S(8), cy, w - S(16), r) + S(8)
        bc.set_rrect(bg, x, y, x + w, cy + S(4), r=S(10))
        return cy + S(4)

    def _draw_job_card(self, x, y, w, r):
        bc = self.board
        S = bc.S
        sev = SEV[_severity(int(r.n))]
        card = bc.rrect(x, y, x + w, y + 10, r=S(8), fill="white", outline="#E5E7EB")
        bar_w = S(5)
        pad = S(10)
        tx = x + bar_w + pad
        right = x + w - pad
        cy = y + S(8)
        _, bb = bc.text(tx, cy, r.so_number, size=11, bold=True)
        bc.badge(right, cy - S(2), _d_label(int(r.n)), sev[1], sev[2], size=9, anchor="ne", padx=7, pady=2)
        cy = bb[3] + S(2)
        _, bb = bc.text(tx, cy, clean_text(r.customer_name) or "-", size=9, fill="#6B7280", width=right - tx)
        cy = bb[3] + S(3)
        bc.text(tx, cy, f"ส่ง {r.delivery.day:02d}/{r.delivery.month:02d} · {r.sale_key or '-'}", size=9, fill="#374151")
        _, bb = bc.text(right, cy, f"{r.sales_service_amount:,.0f}", size=10, bold=True, anchor="ne")
        cy = bb[3] + S(2)
        has_pu = isinstance(r.pu, str) and bool(r.pu)
        bc.text(tx, cy, f"PU: {r.pu if has_pu else UNASSIGNED}", size=9, bold=not has_pu,
                fill="#111827" if has_pu else "#B91C1C")
        if r.hint:
            bc.text(right, cy, r.hint, size=9, fill="#6B7280", anchor="ne")
        h = cy + bc.font(9).metrics("linespace") + S(8) - y
        bc.set_rrect(card, x, y, x + w, y + h, r=S(8))
        bc.rect(x + 1, y + S(6), x + bar_w, y + h - S(6), fill=sev[0])
        return h

    def _draw_workload(self, x, y, w, df):
        """แถบแนวนอน: ภาระงานต่อผู้รับงาน เป็นป้ายชื่อ + จำนวน (ยังไม่มีผู้รับเป็นสีแดง) ขึ้นบรรทัดใหม่เองเมื่อไม่พอ"""
        bc = self.board
        S = bc.S
        bg = bc.rrect(x, y, x + w, y + 10, r=S(10), fill="white", outline="#E5E7EB")
        pad = S(14)
        bc.text(x + pad, y + S(10), "ภาระงานต่อผู้รับงาน", size=11, bold=True)
        if df is None or len(df) == 0:
            bc.text(x + pad, y + S(34), "— ไม่มีงาน —", size=10, fill="#9CA3AF")
            bc.set_rrect(bg, x, y, x + w, y + S(64), r=S(10))
            return y + S(64)
        counts = df["pu"].fillna(UNASSIGNED).value_counts()
        n_pu = int((counts.index != UNASSIGNED).sum())
        if self._pu_staff_total:
            bc.text(x + w - pad, y + S(12), f"ทีมจัดซื้อ {n_pu} จาก {self._pu_staff_total} คน · ใช้เกลี่ยงานและดูคอขวด",
                    size=9, fill="#6B7280", anchor="ne")
        order = ([UNASSIGNED] if UNASSIGNED in counts.index else []) + [k for k in counts.index if k != UNASSIGNED]
        cx, cy = x + pad, y + S(38)
        rowh = bc.font(10, True).metrics("linespace") + S(14)
        for name in order:
            c = int(counts[name])
            red = name == UNASSIGNED
            label = f"{name}  {c} ใบ"
            tw = bc.font(10, True).measure(label) + S(24)
            if cx + tw > x + w - pad:
                cx, cy = x + pad, cy + rowh + S(6)
            bc.badge(cx, cy, label, "#FEE2E2" if red else "#EEF2FF", "#B91C1C" if red else "#1E3A8A", size=10, padx=12, pady=4)
            cx += tw + S(8)
        bottom = cy + rowh + S(10)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(10))
        return bottom

    # ---------------------------------------------------------------- export
    def export_excel(self):
        df = self._filtered()
        if df is None or df.empty:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่มีงานบนกระดานให้ Export", parent=self)
            return
        names = {k: t for k, t, _c in COLS}
        out = pd.DataFrame({
            "เลขที่ SO": df["so_number"], "สถานะงาน": df["col"].map(names), "ลูกค้า": df["customer_name"],
            "เซลส์": df["sale_key"], "ผู้รับงาน (PU)": df["pu"].fillna(UNASSIGNED),
            "วันส่ง": df["delivery"].apply(lambda d: f"{d.day:02d}/{d.month:02d}/{d.year + 543}"),
            "D-x": df["n"].apply(lambda n: _d_label(int(n))), "ยอดขาย (ก่อน VAT)": df["sales_service_amount"],
            "ชำระแล้ว": df["total_payment_amount"], "หมายเหตุ": df["hint"],
        })
        rep = self._report_date()
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"กระดานงานจัดซื้อ_{rep.strftime('%Y%m%d')}.xlsx", parent=self)
        if path:
            try:
                out.to_excel(path, index=False)
                messagebox.showinfo("สำเร็จ", f"บันทึกไฟล์แล้ว:\n{path}", parent=self)
            except Exception as e:
                messagebox.showerror("Error", f"บันทึกไฟล์ไม่สำเร็จ: {e}", parent=self)

    # --------------------------------------------------------------- refresh
    def _schedule_refresh(self):
        try:
            self.after(AUTO_REFRESH_MS, self._auto_refresh)
        except Exception:
            pass

    def _auto_refresh(self):
        try:
            if not self.winfo_exists():
                return
            if self.winfo_ismapped():          # รีเฟรชเฉพาะตอนผู้ใช้กำลังเปิดดูหน้านี้
                if self._report_date() == self._auto_date and date.today() != self._auto_date:
                    self.date_sel.set_date(date.today())      # เปิดค้างข้ามวัน → เลื่อนเป็นวันนี้เอง (ถ้าผู้ใช้ไม่ได้เลือกวันอื่นไว้)
                    self._auto_date = date.today()
                self.load_data(silent=True)
            self._schedule_refresh()
        except Exception:
            pass
