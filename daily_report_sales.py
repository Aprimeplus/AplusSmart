# daily_report_sales.py
# Daily Report ใหม่ — มุมมอง "ฝ่ายขาย" (เฟส 2 ต่อจาก daily_report_purchasing.py) ตามต้นแบบของ PM
#   1) Scoreboard เซลส์ · วันนี้ & รอบเดือน   2) ติดตามเงิน (ค้างชำระ)   3) รายการ SO พร้อมระดับความเสี่ยง
# วาดด้วย Canvas (canvas_board.py) เพื่อให้โหลดเร็ว (วิดเจ็ตต่อเซลล์ช้าเป็นนาทีบนเครื่องจริง)
#
# กติกาที่เป็น "ร่าง" ให้ PM ตรวจ/แก้ อยู่ในค่าคงที่ด้านล่าง (ต้นแบบอ้างเอกสาร DEV-SPEC-RISK-6911-001 ซึ่งไม่มีให้ — เดาจากภาพ):
#   - ขอบเขต SO ในรายการ = เหมือนกระดานจัดซื้อ (Pending PU / PO In Progress / PO Sent ที่ยังไม่ผ่านวันส่ง, ซ่อนงานเปิดที่เลยกำหนดเกิน STALE_DAYS)
#   - ยอดจอง (วันนี้ / MTD) = SO ที่ "บันทึก" ในช่วงวันที่ (timestamp) และผ่านการอนุมัติแล้ว ไม่รวมสถานะใน EXCLUDE_STATUSES
#     (ยกเลิก / ถูกตีกลับ / รอ Sale Manager อนุมัติ / ร่าง) — ไม่ใช่ยอดวัดเป้า KPI ซึ่งหน้า Sale Revenue Report วัดจากใบจ่ายค่าคอม
#   - รอบเดือน: เริ่มวันที่ CYCLE_START_DAY ของเดือน (26 = รอบ 26–25 ; ปีหน้าเดือนต่อเดือนให้ตั้งเป็น 1)
#   - ยอดต่าง: ยอดที่ต้องชำระ เทียบกับ (สินค้า+บริการ+ขนส่ง) × 1.07 หรือ ×1.00 (กรณีไม่คิด VAT) คลาดเกิน 1 บาทถึงติดป้าย
#     (ต้นแบบคำนวณขาดค่าขนส่ง ทำให้ SO ที่มีค่าขนส่งติดป้าย "ต่าง" ผิดทุกใบ — ที่นี่รวมค่าขนส่งแล้ว)
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import date, datetime, timedelta
import traceback

import pandas as pd
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkFont, CTkOptionMenu)

from custom_widgets import DateSelector
from canvas_board import BoardCanvas, clean_text
from daily_report_purchasing import _d_label, _severity, SEV, STALE_DAYS, UNASSIGNED

CYCLE_START_DAY = 26          # วันเริ่มรอบเดือนของยอดสะสม (MTD)
AUTO_REFRESH_MS = 5 * 60 * 1000
VAT = 1.07
TOL = 1.0                      # ผลต่างที่ยอมรับได้ (บาท)
MONEY_CARDS = 4                # จำนวนการ์ด "ติดตามเงิน"
EXCLUDE_STATUSES = ("Cancelled", "Cancelled by PU", "Rejected by SM", "Pending Sale Manager Approval", "Draft", "Original")
_EXCL_SQL = ", ".join("'" + x + "'" for x in EXCLUDE_STATUSES)

# ระดับความเสี่ยงของแต่ละเหตุ: ลำดับในลิสต์ = ลำดับความสำคัญที่ใช้โชว์เป็นเหตุหลัก
RISK_ORDER = ["DL-03", "DL-04", "PY-01", "DQ-01", "DL-05"]
RISK_LEVEL = {"DL-03": "urgent", "DL-04": "urgent", "PY-01": "watch", "DQ-01": "watch", "DL-05": "watch"}

LEVEL_STYLE = {   # (ข้อความ, พื้นป้าย, ตัวอักษร, แถบซ้าย)
    "urgent": ("ด่วน", "#FEE2E2", "#B91C1C", "#C2410C"),
    "watch":  ("เฝ้าระวัง", "#FFEDD5", "#C2410C", "#EA7B1E"),
    "normal": ("ปกติ", "#E5E7EB", "#4B5563", "#CBD5E1"),
}

_SQL_SCOPE = """
SELECT c.so_number, c.customer_name, c.status, c.sale_key,
       COALESCE(c.date_to_customer, c.delivery_date) AS delivery,
       COALESCE(NULLIF(c.user_key, ''),
                (SELECT po.user_key FROM purchase_orders po
                  WHERE po.so_number = c.so_number
                    AND po.status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
                  ORDER BY po.id LIMIT 1)) AS pu,
       COALESCE(c.sales_service_amount, 0)                                       AS goods,
       COALESCE(c.cutting_drilling_fee, 0) + COALESCE(c.other_service_fee, 0)    AS service,
       COALESCE(c.shipping_cost, 0) + COALESCE(c.relocation_cost, 0)             AS ship,
       COALESCE(c.total_payment_amount, 0)                                       AS paid,
       COALESCE(c.difference_amount, 0)                                          AS diff
FROM commissions c
WHERE c.is_active = 1
  AND c.status IN ('Pending PU', 'PO In Progress', 'PO Sent')
  AND COALESCE(c.date_to_customer, c.delivery_date) IS NOT NULL
ORDER BY COALESCE(c.date_to_customer, c.delivery_date), c.so_number
"""

# SO ที่บันทึกในช่วงวันที่ (ไม่รวมยกเลิก/ถูกตีกลับ) — ใช้ทั้ง "วันนี้" และ MTD
_SQL_SOLD = """
SELECT DISTINCT ON (c.so_number) c.so_number, c.sale_key, COALESCE(c.sales_service_amount, 0) AS goods,
       c.timestamp::date AS rec_date
FROM commissions c
WHERE c.is_active = 1
  AND c.status NOT IN (__EXCL__)
  AND c.timestamp::date BETWEEN %(d1)s AND %(d2)s
ORDER BY c.so_number, c.id DESC
""".replace("__EXCL__", _EXCL_SQL)

_SQL_TARGETS = "SELECT sale_key, sale_name, COALESCE(sales_target, 0) AS target FROM sales_users WHERE role = 'Sale' AND status = 'Active'"


def _cycle_start(rep):
    if CYCLE_START_DAY <= 1:
        return rep.replace(day=1)
    if rep.day >= CYCLE_START_DAY:
        return rep.replace(day=CYCLE_START_DAY)
    prev = rep.replace(day=1) - timedelta(days=1)
    return prev.replace(day=min(CYCLE_START_DAY, 28))


def _fmt(x, dec=2):
    return f"{x:,.{dec}f}"


def risk_reasons(r):
    """คืนลิสต์ (code, ข้อความ) ของ SO หนึ่งแถว เรียงตาม RISK_ORDER"""
    out = {}
    n, st = int(r["n"]), r["status"]
    is_open = st in ("Pending PU", "PO In Progress")
    if is_open and n < 0:
        out["DL-03"] = f"เลยวันส่ง {-n} วัน"
    if is_open and n in (0, 1):
        out["DL-04"] = "ส่งวันนี้ ยังไม่มี PO" if n == 0 else "ส่งพรุ่งนี้ ยังไม่มี PO"
    if r["required"] > TOL and r["paid"] <= 0.005 and st in ("PO In Progress", "PO Sent"):
        out["PY-01"] = "ยังไม่ชำระ แต่จัดซื้อดำเนินการแล้ว"
    if r["dq_diff"] is not None and not pd.isna(r["dq_diff"]):
        out["DQ-01"] = f"ยอดต่าง {abs(r['dq_diff']):,.0f} บาท · ตรวจ: {r['sale_key'] or '-'}"
    if st == "Pending PU" and not r["pu"]:
        out["DL-05"] = "ไม่มีผู้รับงาน"
    return [(c, out[c]) for c in RISK_ORDER if c in out]


def load_scope(engine, rep):
    """โหลดขอบเขต SO ของรายงาน (งานเปิด + PO Sent ที่ยังไม่ผ่านวันส่ง) พร้อมคำนวณ D-x, ยอดต่าง, เหตุความเสี่ยง, ระดับ
    ใช้ร่วมกันทั้งกระดานฝ่ายขายและภาพรวมผู้บริหาร เพื่อให้ตัวเลขตรงกันทุกหน้า — คืน (DataFrame, จำนวนงานเปิดค้างเก่าที่ซ่อน)"""
    raw = pd.read_sql_query(_SQL_SCOPE, engine)
    raw["delivery"] = pd.to_datetime(raw["delivery"]).dt.date
    raw["n"] = raw["delivery"].apply(lambda d: (d - rep).days)
    raw["pu"] = raw["pu"].apply(lambda v: v if isinstance(v, str) and v.strip() else None)
    for c in ("goods", "service", "ship", "paid", "diff"):
        raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0.0)
    raw["required"] = raw["paid"] - raw["diff"]                      # ยอดที่ต้องชำระตามระบบ
    raw["outstanding"] = (-raw["diff"]).clip(lower=0)

    base = raw["goods"] + raw["service"] + raw["ship"]
    exp_vat, exp_cash = base * VAT, base
    ok = ((raw["required"] - exp_vat).abs() <= TOL) | ((raw["required"] - exp_cash).abs() <= TOL) | (raw["required"] <= TOL)
    raw["dq_diff"] = (raw["required"] - exp_vat).where(~ok, None)           # เก็บส่วนต่างเฉพาะแถวที่ไม่ตรงสูตร
    raw["dq_diff"] = raw["dq_diff"].apply(lambda v: None if pd.isna(v) else float(v))

    is_open = raw["status"].isin(["Pending PU", "PO In Progress"])
    stale = is_open & (raw["n"] < -STALE_DAYS)
    keep = (~stale) & (is_open | (raw["n"] >= 0))
    df = raw[keep].reset_index(drop=True)
    df["reasons"] = df.apply(risk_reasons, axis=1)
    df["level"] = df["reasons"].apply(
        lambda rs: "urgent" if any(RISK_LEVEL[c] == "urgent" for c, _ in rs) else ("watch" if rs else "normal"))
    return df, int(stale.sum())


class SalesDailyBoard(CTkFrame):
    def __init__(self, master, app_container, sale_key_filter=None, **kwargs):
        super().__init__(master, fg_color="#F3F4F6", **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.locked_sale = sale_key_filter          # Sale ดูได้เฉพาะของตัวเอง
        self.df = pd.DataFrame()
        self.today_df = pd.DataFrame()
        self.mtd_df = pd.DataFrame()
        self.targets = pd.DataFrame()
        self._stale_hidden = 0
        self.chip = "all"
        self.sale_pick = None                      # เซลส์ที่คลิกจาก Scoreboard เพื่อกรองรายการ
        self._auto_date = date.today()             # วันที่ที่ตั้งให้อัตโนมัติ (ใช้เลื่อนเป็นวันนี้เองเมื่อเปิดค้างข้ามวัน)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ---------- header ----------
        head = CTkFrame(self, fg_color="white", corner_radius=10)
        head.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        CTkLabel(head, text="รายงานฝ่ายขาย", font=CTkFont(size=18, weight="bold"), text_color="#111827").pack(
            side="left", padx=(16, 8), pady=10)
        CTkLabel(head, text="Daily Report · ยอดขาย & ติดตามงาน", font=CTkFont(size=12), text_color="#6B7280").pack(side="left")
        CTkButton(head, text="Export Excel", width=110, fg_color="white", text_color="#111827", border_width=1,
                  border_color="#D1D5DB", hover_color="#F3F4F6", command=self.export_excel).pack(side="right", padx=(6, 14))
        CTkButton(head, text="🔄 ดึงข้อมูล", width=110, fg_color="#1D4ED8", hover_color="#1E40AF",
                  command=self.load_data).pack(side="right", padx=6)
        self.sale_var = tk.StringVar(value=self.locked_sale or "ทั้งหมด")
        self.sale_menu = CTkOptionMenu(head, variable=self.sale_var, values=[self.locked_sale or "ทั้งหมด"], width=150,
                                       command=self._on_sale_menu)
        self.sale_menu.pack(side="right", padx=6)
        if self.locked_sale:
            self.sale_menu.configure(state="disabled")
        CTkLabel(head, text="เซลส์:", text_color="#6B7280").pack(side="right")
        self.date_sel = DateSelector(head)
        self.date_sel.pack(side="right", padx=(0, 14))
        self.date_sel.set_date(date.today())
        CTkLabel(head, text="วันที่รายงาน:", text_color="#6B7280").pack(side="right")

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
            rep = self._report_date()
            self.df, self._stale_hidden = load_scope(self.pg_engine, rep)

            sold = pd.read_sql_query(_SQL_SOLD, self.pg_engine, params={"d1": _cycle_start(rep), "d2": rep})
            sold["goods"] = pd.to_numeric(sold["goods"], errors="coerce").fillna(0.0)
            sold["rec_date"] = pd.to_datetime(sold["rec_date"]).dt.date
            self.mtd_df = sold
            self.today_df = sold[sold["rec_date"] == rep]
            self.targets = pd.read_sql_query(_SQL_TARGETS, self.pg_engine)

            sales = sorted({s for s in self.df["sale_key"].dropna()} | set(self.today_df["sale_key"].dropna()))
            if not self.locked_sale:
                cur = self.sale_var.get()
                self.sale_menu.configure(values=["ทั้งหมด"] + sales)
                if cur not in (["ทั้งหมด"] + sales):
                    self.sale_var.set("ทั้งหมด")
            self.render()
        except Exception as e:
            traceback.print_exc()
            if not silent:            # รีเฟรชอัตโนมัติไม่เด้งหน้าต่างซ้ำๆ ถ้าฐานข้อมูลสะดุด (คงข้อมูลเดิมไว้)
                messagebox.showerror("Error", f"ดึงข้อมูลรายงานฝ่ายขายไม่สำเร็จ: {e}", parent=self)

    def _on_sale_menu(self, _v):
        self.sale_pick = None
        self.render()

    def _active_sale(self):
        if self.locked_sale:
            return self.locked_sale
        v = self.sale_var.get()
        return None if v == "ทั้งหมด" else v

    def _scope_df(self):
        s = self._active_sale()
        return self.df if (s is None or self.df.empty) else self.df[self.df["sale_key"] == s]

    # ---------------------------------------------------------------- render
    def render(self):
        bc = self.board
        S = bc.S
        bc.clear()
        W = bc.width
        M = S(10)
        y = M
        y = self._draw_scoreboard(M, y, W - 2 * M) + S(10)
        y = self._draw_money(M, y, W - 2 * M) + S(10)
        y = self._draw_list(M, y, W - 2 * M)
        bc.set_height(y + M)
        rep = self._report_date()
        txt = f"รายงาน ณ {rep.day:02d}/{rep.month:02d}/{rep.year + 543} · รอบสะสมเริ่ม {_cycle_start(rep).strftime('%d/%m')}"
        if self._stale_hidden:
            txt += f"  |  ไม่แสดงงานเปิดที่เลยกำหนดเกิน {STALE_DAYS} วัน {self._stale_hidden} รายการ"
        self.note_lbl.configure(text=txt)

    def _card_bg(self, x, y, w):
        bc = self.board
        return bc.rrect(x, y, x + w, y + 10, r=bc.S(12), fill="white", outline="#E5E7EB")

    def _card_header(self, x, y, w, title, hint=""):
        bc = self.board
        S = bc.S
        bc.text(x + S(16), y + S(12), title, size=13, bold=True)
        if hint:
            bc.text(x + w - S(16), y + S(16), hint, size=9, fill="#6B7280", anchor="ne")
        return y + S(46)

    # ---------- 1) scoreboard ----------
    def _draw_scoreboard(self, x, y, w):
        bc = self.board
        S = bc.S
        bg = self._card_bg(x, y, w)
        cy = self._card_header(x, y, w, "Scoreboard เซลส์ · วันนี้ & รอบเดือน", "เรียงตามยอดวันนี้ · คลิกชื่อเพื่อกรองรายการ SO ด้านล่าง")
        sales = set(self.today_df["sale_key"].dropna()) | set(self.df["sale_key"].dropna())
        if self.locked_sale:
            sales = {self.locked_sale}
        elif self._active_sale():
            sales = {self._active_sale()}
        tgt = dict(zip(self.targets["sale_key"], self.targets["target"])) if not self.targets.empty else {}
        rows = []
        for s in sales:
            t = self.today_df[self.today_df["sale_key"] == s]
            m = self.mtd_df[self.mtd_df["sale_key"] == s]
            d = self.df[self.df["sale_key"] == s]
            cnt, amt = len(t), float(t["goods"].sum())
            rows.append({"sale": s, "cnt": cnt, "amt": amt, "avg": (amt / cnt) if cnt else 0.0,
                         "mtd": float(m["goods"].sum()), "target": float(tgt.get(s, 0) or 0),
                         "outstanding": float(d["outstanding"].sum()),
                         "pending_po": int(d["status"].isin(["Pending PU", "PO In Progress"]).sum())})
        rows.sort(key=lambda r: (-r["amt"], -r["pending_po"], r["sale"]))

        inner_x, inner_w = x + S(16), w - S(32)
        weights = [2, 1, 2, 2, 2, 4, 2, 2]
        tot = float(sum(weights))
        cols, acc = [], 0.0
        for wt in weights:
            cols.append((inner_x + inner_w * acc / tot, inner_x + inner_w * (acc + wt) / tot))
            acc += wt
        heads = ["เซลส์", "SO วันนี้", "ยอดวันนี้", "เฉลี่ย/บิล", "ยอดจอง MTD", "% เป้าเดือน", "ค้างชำระ", "SO ติดจัดซื้อ"]
        for i, h in enumerate(heads):
            if i in (0, 5):
                bc.text(cols[i][0] + S(6), cy, h, size=10, fill="#6B7280")
            else:
                bc.text(cols[i][1] - S(6), cy, h, size=10, fill="#6B7280", anchor="ne")
        cy += S(26)
        bc.line(inner_x, cy - S(4), inner_x + inner_w, cy - S(4))
        if not rows:
            bc.text(x + w / 2, cy + S(6), "— ไม่มีข้อมูล —", size=10, fill="#9CA3AF", anchor="n")
            cy += S(40)
        for ri, r in enumerate(rows):
            sel = (self.sale_pick == r["sale"])
            rh = S(38)
            if sel:
                bc.rrect(inner_x, cy, inner_x + inner_w, cy + rh - S(4), r=S(8), fill="#EFF6FF")
            tag = f"sale_{ri}"
            bc.text(cols[0][0] + S(6), cy + S(8), r["sale"], size=11, bold=True, tags=(tag,))
            bc.tag_click(tag, lambda s=r["sale"]: self._pick_sale(s))
            for ci, val, bold in ((1, str(r["cnt"]), False), (2, _fmt(r["amt"], 0), True),
                                  (3, _fmt(r["avg"], 0), False), (4, _fmt(r["mtd"], 0), False)):
                bc.text(cols[ci][1] - S(6), cy + S(8), val, size=11, bold=bold, anchor="ne")
            # % เป้าเดือน
            bx0, bx1 = cols[5][0] + S(6), cols[5][1] - S(6)
            if r["target"] > 0:
                pct = r["mtd"] / r["target"]
                bc.bar(bx0, cy + S(14), (bx1 - bx0) - S(50), S(8), pct, color="#16A34A" if pct >= 1 else "#1D4ED8")
                bc.text(bx1, cy + S(8), f"{pct * 100:.0f}%", size=10, fill="#374151", anchor="ne")
            else:
                bc.text(bx0, cy + S(9), "ไม่ได้ตั้งเป้า", size=9, fill="#9CA3AF")
            bc.text(cols[6][1] - S(6), cy + S(8), _fmt(r["outstanding"], 0) if r["outstanding"] > 0.5 else "0",
                    size=11, fill="#B45309" if r["outstanding"] > 0.5 else "#6B7280", anchor="ne")
            n = r["pending_po"]
            bc.badge(cols[7][1] - S(6), cy + S(5), n, "#DCFCE7" if n == 0 else ("#FFEDD5" if n <= 2 else "#FEE2E2"),
                     "#166534" if n == 0 else ("#C2410C" if n <= 2 else "#B91C1C"), size=10, anchor="ne", padx=10, pady=2)
            cy += rh
            bc.line(inner_x, cy - S(4), inner_x + inner_w, cy - S(4), fill="#F3F4F6")
        _, bb = bc.text(inner_x, cy + S(2),
                        f"ยอดจอง = SO ที่บันทึกและผ่านการอนุมัติแล้ว (ไม่รวมยกเลิก/ถูกตีกลับ/รออนุมัติ) นับตั้งแต่ {_cycle_start(self._report_date()).strftime('%d/%m')} · "
                        "% เป้าเดือน = ยอดจอง MTD ÷ เป้าของเซลส์ · ไม่ใช่ยอดวัดเป้า KPI (ดูที่ Sale Revenue Report ซึ่งวัดจากใบจ่ายค่าคอม)",
                        size=9, fill="#6B7280", width=inner_w)
        bottom = bb[3] + S(12)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    def _pick_sale(self, s):
        self.sale_pick = None if self.sale_pick == s else s
        self.render()

    # ---------- 2) money tracking ----------
    def _draw_money(self, x, y, w):
        bc = self.board
        S = bc.S
        df = self._scope_df()
        owe = df[df["outstanding"] > TOL] if len(df) else df
        total = float(owe["outstanding"].sum()) if len(owe) else 0.0
        bg = self._card_bg(x, y, w)
        cy = self._card_header(x, y, w, f"ติดตามเงิน · ค้างชำระ {total:,.0f}", "เรียงตามวันส่งที่ใกล้สุด · ต้องเก็บก่อนปล่อยของ")
        if len(owe) == 0:
            bc.text(x + w / 2, cy + S(4), "ไม่มีงานค้างชำระ", size=10, fill="#9CA3AF", anchor="n")
            bottom = cy + S(40)
            bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
            return bottom
        gap = S(10)
        cw = (w - S(32) - gap * (MONEY_CARDS - 1)) / MONEY_CARDS
        ch = S(124)
        for i, r in enumerate(owe.sort_values(["n", "so_number"]).head(MONEY_CARDS).itertuples()):
            cx = x + S(16) + i * (cw + gap)
            unpaid = r.paid <= 0.005
            red = unpaid and r.n <= 2
            bc.rrect(cx, cy, cx + cw, cy + ch, r=S(10), fill="white", outline="#FCA5A5" if red else "#E5E7EB")
            bc.text(cx + S(12), cy + S(10), r.so_number, size=11, bold=True)
            bc.badge(cx + cw - S(10), cy + S(8), _d_label(int(r.n)) + (" · ยังไม่ชำระ" if red else ""),
                     "#FEE2E2" if red else "#E5E7EB", "#B91C1C" if red else "#4B5563", size=9, anchor="ne", padx=7, pady=2)
            bc.text(cx + S(12), cy + S(34), f"{clean_text(r.customer_name) or '-'} · {r.sale_key or '-'}", size=9, fill="#6B7280", width=cw - S(24))
            bc.text(cx + S(12), cy + S(62), "คงเหลือ", size=9, fill="#6B7280")
            bc.text(cx + cw - S(12), cy + S(56), _fmt(r.outstanding), size=14, bold=True, anchor="ne")
            pct = (r.paid / r.required) if r.required > TOL else 0.0
            bc.bar(cx + S(12), cy + S(86), cw - S(24), S(8), pct, color="#166534")
            bc.text(cx + S(12), cy + S(100), f"ชำระแล้ว {pct * 100:.0f}% · ส่ง {r.delivery.day:02d}/{r.delivery.month:02d}", size=9, fill="#6B7280")
        bottom = cy + ch + S(14)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    # ---------- 3) SO list ----------
    def _chips_and_filter(self, df):
        """คืน (df ที่กรองตามชิป, dict ชิป→จำนวน)"""
        me = self.locked_sale or (self.app_container.current_user_key if getattr(
            self.app_container, "current_user_role", "") == "Sale" else None)
        masks = {
            "all": pd.Series(True, index=df.index),
            "urgent": df["level"] == "urgent",
            "watch": df["level"] == "watch",
            "today": df["n"] == 0,
            "overdue": df["reasons"].apply(lambda rs: any(c == "DL-03" for c, _ in rs)),
            "owe": df["outstanding"] > TOL,
            "dq": df["dq_diff"].notna(),
        }
        if me and not self.locked_sale:
            masks["mine"] = df["sale_key"] == me
        counts = {k: int(m.sum()) for k, m in masks.items()}
        return df[masks.get(self.chip, masks["all"])], counts

    def _list_base(self):
        base = self._scope_df()
        if self.sale_pick and len(base):
            base = base[base["sale_key"] == self.sale_pick]
        return base

    def _draw_list(self, x, y, w):
        bc = self.board
        S = bc.S
        base = self._list_base()
        bg = self._card_bg(x, y, w)
        bc.text(x + S(16), y + S(12), f"รายการ SO · {len(base)} ใบ", size=13, bold=True)
        if base.empty:
            bc.text(x + w / 2, y + S(60), "— ไม่มี SO ในขอบเขตรายงาน —", size=10, fill="#9CA3AF", anchor="n")
            bottom = y + S(110)
            bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
            return bottom
        shown, counts = self._chips_and_filter(base)
        chip_defs = [("all", "ทั้งหมด"), ("urgent", "ด่วน"), ("watch", "เฝ้าระวัง"), ("mine", "ของฉัน"),
                     ("today", "ส่งวันนี้"), ("overdue", "เกินกำหนด"), ("owe", "ค้างชำระ"), ("dq", "ยอดต่าง")]
        cx = x + w - S(16)
        for key, label in reversed(chip_defs):               # วาดจากขวาไปซ้าย
            if key not in counts:
                continue
            on = (self.chip == key)
            txt = f"{label} {counts[key]}"
            f = bc.font(10, on)
            cwid = f.measure(txt) + S(26)
            tag = f"chip_{key}"
            bc.rrect(cx - cwid, y + S(8), cx, y + S(8) + S(28), r=S(14), fill="#111827" if on else "white",
                     outline="#111827" if on else "#D1D5DB", tags=(tag,))
            bc.canvas.create_text(cx - cwid / 2, y + S(22), text=txt, font=f, fill="white" if on else "#374151", tags=(tag,))
            bc.tag_click(tag, lambda k=key: self._set_chip(k))
            cx -= cwid + S(6)
        _, bb = bc.text(x + S(16), y + S(44),
                        "เรียงตามความเสี่ยง → วันส่ง → มูลค่า · สีแถบซ้าย = ระดับความเสี่ยง · ใต้ badge คือเหตุหลัก",
                        size=9, fill="#6B7280", width=w - S(32))
        cy = bb[3] + S(10)

        order = {"urgent": 0, "watch": 1, "normal": 2}
        shown = shown.assign(_o=shown["level"].map(order)).sort_values(["_o", "delivery", "goods"], ascending=[True, True, False])

        inner_x, inner_w = x + S(16), w - S(32)
        spec = [("SO", 11, "w"), ("ความเสี่ยง", 19, "w"), ("ลูกค้า", 19, "w"), ("เซลส์", 8, "w"), ("จัดซื้อ (PU)", 9, "w"),
                ("ยอดสินค้า", 9, "e"), ("ค่าบริการ", 8, "e"), ("รวมสุทธิ (VAT)", 13, "e"), ("ชำระ", 9, "w"), ("วันส่ง", 9, "w")]
        tot = float(sum(s[1] for s in spec))
        edge = S(10)       # ที่ให้แถบสีซ้าย
        avail = inner_w - edge
        xs, acc = [], 0.0
        for _, wt, _a in spec:
            xs.append((inner_x + edge + avail * acc / tot, inner_x + edge + avail * (acc + wt) / tot))
            acc += wt
        for i, (h, _wt, a) in enumerate(spec):
            if a == "e":
                bc.text(xs[i][1] - S(6), cy, h, size=10, fill="#6B7280", anchor="ne")
            else:
                bc.text(xs[i][0] + S(6), cy, h, size=10, fill="#6B7280")
        cy += S(24)
        bc.line(inner_x, cy - S(4), inner_x + inner_w, cy - S(4))
        if shown.empty:
            bc.text(x + w / 2, cy + S(10), "— ไม่มีรายการตามตัวกรองนี้ —", size=10, fill="#9CA3AF", anchor="n")
            cy += S(50)
        for r in shown.itertuples():
            lv = LEVEL_STYLE[r.level]
            top = cy
            ty = top + S(8)
            so_size = 11
            while so_size > 8 and bc.font(so_size, True).measure(str(r.so_number)) > (xs[0][1] - xs[0][0]) - S(10):
                so_size -= 1                      # เลข SO ยาวผิดปกติ ย่อให้พอดีคอลัมน์ ไม่ล้นไปทับป้ายความเสี่ยง
            bc.text(xs[0][0] + S(6), ty, r.so_number, size=so_size, bold=True)
            # ความเสี่ยง: ป้าย + เหตุหลัก
            label = lv[0] + (f" · {len(r.reasons)} เหตุ" if len(r.reasons) > 1 else "")
            bw, bh = bc.badge(xs[1][0] + S(6), ty - S(2), label, lv[1], lv[2], size=9, padx=8, pady=2)
            risk_bottom = ty - S(2) + bh
            if r.reasons:
                _, bb = bc.text(xs[1][0] + S(6), risk_bottom + S(2), f"{r.reasons[0][0]} {r.reasons[0][1]}", size=9,
                                fill="#6B7280", width=(xs[1][1] - xs[1][0]) - S(12))
                risk_bottom = bb[3]
            _, bb = bc.text(xs[2][0] + S(6), ty, clean_text(r.customer_name) or "-", size=10, width=(xs[2][1] - xs[2][0]) - S(12))
            cust_bottom = bb[3]
            bc.text(xs[3][0] + S(6), ty, r.sale_key or "-", size=10)
            has_pu = bool(r.pu)
            bc.text(xs[4][0] + S(6), ty, r.pu if has_pu else UNASSIGNED, size=10, fill="#111827" if has_pu else "#B91C1C")
            bc.text(xs[5][1] - S(6), ty, _fmt(r.goods), size=10, anchor="ne")
            bc.text(xs[6][1] - S(6), ty, _fmt(r.service + r.ship) if (r.service + r.ship) else "0.00", size=10, anchor="ne")
            # รวมสุทธิ + ป้ายต่าง
            rx = xs[7][1] - S(6)
            _, bb = bc.text(rx, ty, _fmt(r.required), size=11, bold=True, anchor="ne")
            if r.dq_diff is not None and not pd.isna(r.dq_diff):
                bc.badge(bb[0] - S(6), ty - S(2), f"ต่าง {abs(r.dq_diff):,.0f}", "#FFEDD5", "#C2410C", size=9,
                         anchor="ne", padx=7, pady=2)
            # ชำระ
            if r.required <= TOL or r.outstanding <= TOL:
                pay = ("ครบ", "#DCFCE7", "#166534")
            elif r.paid <= 0.005:
                pay = ("ยังไม่ชำระ", "#FEE2E2", "#B91C1C")
            else:
                pay = (f"มัดจำ {round(r.paid / r.required * 100)}%", "#DBEAFE", "#1D4ED8")
            bc.badge(xs[8][0] + S(6), ty - S(2), pay[0], pay[1], pay[2], size=9, padx=8, pady=2)
            # วันส่ง
            sev = SEV[_severity(int(r.n))]
            _, bb = bc.text(xs[9][0] + S(6), ty, f"{r.delivery.day:02d}/{r.delivery.month:02d}", size=10)
            bc.badge(bb[2] + S(6), ty - S(2), _d_label(int(r.n)), sev[1], sev[2], size=8, padx=6, pady=1)
            row_h = max(risk_bottom, cust_bottom, top + S(40)) - top + S(12)
            bc.rect(inner_x, top + S(2), inner_x + S(4), top + row_h - S(4), fill=lv[3])
            cy = top + row_h
            bc.line(inner_x, cy - S(2), inner_x + inner_w, cy - S(2), fill="#F3F4F6")
        # แถวรวม
        cy += S(4)
        bc.text(xs[0][0] + S(6), cy, "รวม", size=11, bold=True)
        for col, val in ((5, shown["goods"].sum() if len(shown) else 0),
                         (6, (shown["service"] + shown["ship"]).sum() if len(shown) else 0),
                         (7, shown["required"].sum() if len(shown) else 0)):
            bc.text(xs[col][1] - S(6), cy, _fmt(val), size=11, bold=True, anchor="ne")
        cy += S(24)
        bc.text(inner_x + inner_w - S(6), cy,
                f"ชำระแล้ว {shown['paid'].sum() if len(shown) else 0:,.2f} · ค้างชำระ {shown['outstanding'].sum() if len(shown) else 0:,.2f}",
                size=10, bold=True, fill="#166534", anchor="ne")
        cy += S(30)
        bottom = cy + S(10)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    def _set_chip(self, key):
        self.chip = key
        self.render()

    # ---------------------------------------------------------------- export
    def export_excel(self):
        base = self._list_base()
        shown, _ = self._chips_and_filter(base) if len(base) else (base, {})
        if shown.empty:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่มีรายการให้ Export", parent=self)
            return
        lvname = {"urgent": "ด่วน", "watch": "เฝ้าระวัง", "normal": "ปกติ"}
        out = pd.DataFrame({
            "เลขที่ SO": shown["so_number"], "ระดับความเสี่ยง": shown["level"].map(lvname),
            "เหตุ": shown["reasons"].apply(lambda rs: " | ".join(f"{c} {t}" for c, t in rs)),
            "ลูกค้า": shown["customer_name"], "เซลส์": shown["sale_key"], "ผู้รับงาน (PU)": shown["pu"].fillna(UNASSIGNED),
            "ยอดสินค้า": shown["goods"], "ค่าบริการ+ขนส่ง": shown["service"] + shown["ship"],
            "รวมสุทธิ (VAT)": shown["required"], "ชำระแล้ว": shown["paid"], "ค้างชำระ": shown["outstanding"],
            "ยอดต่างจากสูตร": shown["dq_diff"],
            "วันส่ง": shown["delivery"].apply(lambda d: f"{d.day:02d}/{d.month:02d}/{d.year + 543}"),
            "D-x": shown["n"].apply(lambda n: _d_label(int(n))),
        })
        rep = self._report_date()
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                            initialfile=f"รายงานฝ่ายขาย_{rep.strftime('%Y%m%d')}.xlsx", parent=self)
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
            if self.winfo_ismapped():
                if self._report_date() == self._auto_date and date.today() != self._auto_date:
                    self.date_sel.set_date(date.today())      # เปิดค้างข้ามวัน → เลื่อนเป็นวันนี้เอง (ถ้าผู้ใช้ไม่ได้เลือกวันอื่นไว้)
                    self._auto_date = date.today()
                self.load_data(silent=True)
            self._schedule_refresh()
        except Exception:
            pass
