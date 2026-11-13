# daily_report_executive.py
# Daily Report ใหม่ — มุมมอง "ภาพรวมผู้บริหาร" (เฟส 3 ต่อจาก daily_report_purchasing.py และ daily_report_sales.py)
# ใช้ตัวโหลด/กติกาชุดเดียวกับสองกระดานนั้น (load_scope, risk_reasons) เพื่อให้ตัวเลขตรงกันทุกหน้า วาดด้วย Canvas (เร็ว)
#
# นิยาม (ปรับได้ที่ค่าคงที่ด้านล่าง / ให้ PM ยืนยัน):
#   ยอดขายวันนี้ / ยอดสะสมรอบเดือน (MTD) = ยอดจอง (SO ที่บันทึกและอนุมัติแล้ว ก่อน VAT) — ไม่ใช่ยอดวัดเป้า KPI (หน้า Sale Revenue Report วัดจากใบจ่ายค่าคอม)
#   เป้ารวม = COMPANY_TARGET ถ้าตั้งไว้ ไม่งั้นใช้ผลรวม sales_target ของเซลส์ที่ Active · เป้าวัน = เป้ารวม ÷ วันทำงานในรอบ
#   วันทำงาน = WORK_WEEKDAYS (จันทร์–เสาร์) ไม่รวมวันหยุดบริษัท (ตาราง company_holidays) · ยอดที่ "ควรได้" นับตามวันทำงานเช่นเดียวกัน
#   เงินรับวันนี้ = เงินที่บันทึกรับในวันที่รายงานจริง (งวดแรก payment1 + งวดหลัง payment2) ไม่ใช่ยอดชำระสะสมของงานค้าง
#   Blended GM (forecast) = ถ่วงน้ำหนักด้วยยอดขาย: SO ที่ PO ครบใช้ (ขาย−ต้นทุน PO) จริง, SO ที่ยังไม่เปิด PO ครบประมาณด้วย GM ของส่วนที่ยืนยัน (ต้นทุน PO ไม่รวมค่าคอม/ค่าธรรมเนียมอื่น)
#   ยอดต่าง (DQ-01) ใช้สูตรที่รวมค่าขนส่ง (ต้นแบบ PM ลืมค่าขนส่ง ทำให้ติดป้ายผิดทุกใบที่มีค่าขนส่ง)
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import date, datetime, timedelta
import traceback

import pandas as pd
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkFont)

from custom_widgets import DateSelector
from canvas_board import BoardCanvas, clean_text
import daily_report_sales as drs
from daily_report_sales import load_scope, LEVEL_STYLE, _SQL_SOLD, _SQL_TARGETS, _cycle_start, _fmt
from daily_report_purchasing import COLS, UNASSIGNED, _d_label

COMPANY_TARGET = None                    # None = ใช้ผลรวม sales_target ของเซลส์ Active ; ใส่ตัวเลข เช่น 11_000_000 เพื่อกำหนดเอง
WORK_WEEKDAYS = (0, 1, 2, 3, 4, 5)       # จันทร์=0 … เสาร์=5 (ตัดเสาร์ออกถ้าบริษัทไม่ทำงานเสาร์)
TODO_ROWS = 8                            # จำนวนรายการใน "ต้องทำวันนี้"
AUTO_REFRESH_MS = 5 * 60 * 1000

_SQL_CASH = """
SELECT COALESCE(SUM(CASE WHEN payment1_date = %(d)s THEN payment1_amount ELSE 0 END), 0) AS p1,
       COALESCE(SUM(CASE WHEN payment2_date = %(d)s THEN payment2_amount ELSE 0 END), 0) AS p2,
       COUNT(DISTINCT so_number) FILTER (WHERE payment1_date = %(d)s OR payment2_date = %(d)s) AS n
FROM commissions
WHERE is_active = 1 AND status NOT IN ('Cancelled', 'Cancelled by PU')
"""

# ต้นทุนจาก PO (ไม่รวมใบที่ยกเลิก/ถูกตีกลับ) ต่อ SO ที่บันทึกในรอบ — ใช้คำนวณ Gross Margin
_SQL_GM = """
SELECT x.so_number, x.status, x.goods, COALESCE(po.cost, 0) AS cost
FROM (SELECT DISTINCT ON (c.so_number) c.so_number, c.status, COALESCE(c.sales_service_amount, 0) AS goods
      FROM commissions c
      WHERE c.is_active = 1 AND c.status NOT IN (__EXCL__)
        AND c.timestamp::date BETWEEN %(d1)s AND %(d2)s
      ORDER BY c.so_number, c.id DESC) x
LEFT JOIN (SELECT so_number, SUM(total_cost) AS cost FROM purchase_orders
            WHERE status NOT IN ('Cancelled', 'Cancelled by PU', 'Rejected', 'Rejected by SM')
            GROUP BY so_number) po ON po.so_number = x.so_number
""".replace("__EXCL__", drs._EXCL_SQL)
GM_OPEN_STATUSES = ("Pending PU", "PO In Progress")   # ยังเปิด PO ไม่ครบ → ต้นทุนยังไม่นิ่ง ใช้ประมาณการ


def blend_gm(rows):
    """Blended Gross Margin (forecast) ของ SO ที่จองในรอบ. rows = [(สถานะ, ยอดขาย, ต้นทุน PO)]
    ยืนยันแล้ว = SO ที่ PO ครบ (ไม่ใช่ Pending PU/PO In Progress) และมีต้นทุน → ใช้ margin จริง
    ที่เหลือ (ยังไม่เปิด PO/เปิดไม่ครบ) → ประมาณด้วย margin ของส่วนที่ยืนยันแล้ว แล้วถ่วงน้ำหนักด้วยยอดขาย
    คืน dict: blended, confirmed_gm, coverage, sales_total, sales_conf, gp_conf, gp_forecast, n_conf, n_fc"""
    sc = gc = sf = 0.0
    nc = nf = 0
    for st, sales, cost in rows:
        sales, cost = float(sales or 0), float(cost or 0)
        if st not in GM_OPEN_STATUSES and cost > 0 and sales > 0:
            sc += sales; gc += sales - cost; nc += 1
        elif sales > 0:
            sf += sales; nf += 1
    total = sc + sf
    conf_gm = (gc / sc) if sc > 0 else None
    gp_f = sf * conf_gm if conf_gm is not None else 0.0
    blended = ((gc + gp_f) / total) if (total > 0 and conf_gm is not None) else None
    return dict(blended=blended, confirmed_gm=conf_gm, coverage=(sc / total) if total > 0 else 0.0,
                sales_total=total, sales_conf=sc, gp_conf=gc, gp_forecast=gp_f, n_conf=nc, n_fc=nf)


# ----------------------------------------------------------------------------- ฟังก์ชันล้วน (มีเทสต์)
def cycle_bounds(rep):
    """(วันเริ่มรอบ, วันสุดท้ายของรอบ) ที่ครอบวันที่รายงาน — รอบ 26–25 หรือเดือนปฏิทินตาม drs.CYCLE_START_DAY"""
    start = _cycle_start(rep)
    if drs.CYCLE_START_DAY <= 1:
        nxt = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
        return start, nxt - timedelta(days=1)
    y, m = (start.year + 1, 1) if start.month == 12 else (start.year, start.month + 1)
    return start, date(y, m, drs.CYCLE_START_DAY - 1)


def working_days(start, end, holidays=()):
    hol = set(holidays)
    out, d = [], start
    while d <= end:
        if d.weekday() in WORK_WEEKDAYS and d not in hol:
            out.append(d)
        d += timedelta(days=1)
    return out


def daily_target(total_target, work_days):
    return (total_target / len(work_days)) if work_days else 0.0


def expected_to_date(total_target, work_days, rep):
    """ยอดที่ "ควรได้" ถึงวันที่รายงาน นับตามสัดส่วนวันทำงานที่ผ่านมา (รวมวันนี้)"""
    if not work_days:
        return 0.0
    return total_target * sum(1 for d in work_days if d <= rep) / len(work_days)


def concentration(amounts, k=3):
    """(สัดส่วนของ k รายการใหญ่สุด, ผลรวมรายการที่เหลือ, ผลรวมทั้งหมด) — ไม่มียอดคืน (0, 0, 0)"""
    vals = sorted((float(a) for a in amounts), reverse=True)
    total = sum(vals)
    if total <= 0:
        return 0.0, 0.0, 0.0
    top = sum(vals[:k])
    return top / total, total - top, total


# ----------------------------------------------------------------------------- หน้าจอ
class ExecutiveDailyBoard(CTkFrame):
    def __init__(self, master, app_container, **kwargs):
        super().__init__(master, fg_color="#F3F4F6", **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.df = pd.DataFrame()
        self.sold = pd.DataFrame()
        self.today_df = pd.DataFrame()
        self.cash = {"p1": 0.0, "p2": 0.0, "n": 0}
        self.gm = blend_gm([])
        self.total_target = 0.0
        self.work_days = []
        self._stale_hidden = 0
        self._auto_date = date.today()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        head = CTkFrame(self, fg_color="white", corner_radius=10)
        head.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        CTkLabel(head, text="ภาพรวมผู้บริหาร", font=CTkFont(size=18, weight="bold"), text_color="#111827").pack(
            side="left", padx=(16, 8), pady=10)
        CTkLabel(head, text="Daily Report · ยอดขาย & จัดซื้อ", font=CTkFont(size=12), text_color="#6B7280").pack(side="left")
        CTkButton(head, text="Export Excel", width=110, fg_color="white", text_color="#111827", border_width=1,
                  border_color="#D1D5DB", hover_color="#F3F4F6", command=self.export_excel).pack(side="right", padx=(6, 14))
        CTkButton(head, text="🔄 ดึงข้อมูลล่าสุด", width=130, fg_color="#1D4ED8", hover_color="#1E40AF",
                  command=self.load_data).pack(side="right", padx=6)
        self.cycle_lbl = CTkLabel(head, text="", font=CTkFont(size=12), text_color="#374151")
        self.cycle_lbl.pack(side="right", padx=(0, 14))
        CTkLabel(head, text="รอบ:", text_color="#6B7280").pack(side="right")
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
            start, end = cycle_bounds(rep)
            self.df, self._stale_hidden = load_scope(self.pg_engine, rep)

            sold = pd.read_sql_query(_SQL_SOLD, self.pg_engine, params={"d1": start, "d2": min(rep, end)})
            sold["goods"] = pd.to_numeric(sold["goods"], errors="coerce").fillna(0.0)
            sold["rec_date"] = pd.to_datetime(sold["rec_date"]).dt.date
            self.sold = sold
            self.today_df = sold[sold["rec_date"] == rep]

            tg = pd.read_sql_query(_SQL_TARGETS, self.pg_engine)
            self.total_target = float(COMPANY_TARGET) if COMPANY_TARGET else float(tg["target"].sum())
            try:
                hol = pd.read_sql_query("SELECT holiday_date FROM company_holidays WHERE holiday_date BETWEEN %(a)s AND %(b)s",
                                        self.pg_engine, params={"a": start, "b": end})
                holidays = [pd.Timestamp(x).date() for x in hol["holiday_date"]]
            except Exception:
                holidays = []
            self.work_days = working_days(start, end, holidays)

            gm = pd.read_sql_query(_SQL_GM, self.pg_engine, params={"d1": start, "d2": min(rep, end)})
            self.gm = blend_gm(zip(gm["status"], gm["goods"], gm["cost"]))

            c = pd.read_sql_query(_SQL_CASH, self.pg_engine, params={"d": rep})
            self.cash = {"p1": float(c["p1"].iloc[0] or 0), "p2": float(c["p2"].iloc[0] or 0), "n": int(c["n"].iloc[0] or 0)}
            self.cycle_lbl.configure(text=f"{start.day} {_TH[start.month]} – {end.day} {_TH[end.month]} {end.year + 543}")
            self.render()
        except Exception as e:
            traceback.print_exc()
            if not silent:
                messagebox.showerror("Error", f"ดึงข้อมูลภาพรวมผู้บริหารไม่สำเร็จ: {e}", parent=self)

    # ---------------------------------------------------------------- render
    def render(self):
        bc = self.board
        S = bc.S
        bc.clear()
        W = bc.width
        M, G = S(10), S(10)
        rep = self._report_date()
        y = M
        y = self._draw_kpis(M, y, W - 2 * M, G) + G
        y = self._draw_gm(M, y, W - 2 * M) + G
        left_w = (W - 2 * M - G) * 0.64
        right_w = (W - 2 * M - G) - left_w
        rx = M + left_w + G
        yl = self._draw_todo(M, y, left_w)
        yr = self._draw_status(rx, y, right_w)
        yr = self._draw_concentration(rx, yr + G, right_w)
        y = max(yl, yr) + G
        yl = self._draw_chart(M, y, left_w, rep)
        yr = self._draw_by_sale(rx, y, right_w)
        y = max(yl, yr)
        bc.set_height(y + M)
        tgt = f"เป้ารวม {self.total_target / 1e6:.2f} ล้าน" if self.total_target else "ยังไม่ได้ตั้งเป้า"
        txt = (f"รายงาน ณ {rep.day:02d}/{rep.month:02d}/{rep.year + 543} · {tgt} · วันทำงานในรอบ {len(self.work_days)} วัน "
               f"(จันทร์–เสาร์ ไม่รวมวันหยุดบริษัท)")
        if self._stale_hidden:
            txt += f"  |  ไม่แสดงงานเปิดที่เลยกำหนดเกิน {drs.STALE_DAYS} วัน {self._stale_hidden} รายการ"
        self.note_lbl.configure(text=txt)

    def _card(self, x, y, w, h, fill="white", outline="#E5E7EB"):
        return self.board.rrect(x, y, x + w, y + h, r=self.board.S(12), fill=fill, outline=outline)

    # ---- 6 การ์ด KPI ----
    def _draw_kpis(self, x, y, w, gap):
        bc = self.board
        S = bc.S
        rep = self._report_date()
        df = self.df
        n_cards = 6
        cw = (w - gap * (n_cards - 1)) / n_cards
        ch = S(132)

        today_amt, today_n = float(self.today_df["goods"].sum()), len(self.today_df)
        mtd = float(self.sold["goods"].sum())
        dt = daily_target(self.total_target, self.work_days)
        exp = expected_to_date(self.total_target, self.work_days, rep)
        done_days = sum(1 for d in self.work_days if d <= rep)
        owe = df[df["outstanding"] > drs.TOL] if len(df) else df
        n_part = int(((owe["paid"] > 0.005)).sum()) if len(owe) else 0
        n_unpaid = int(((owe["paid"] <= 0.005)).sum()) if len(owe) else 0
        wait = df[df["status"] == "Pending PU"] if len(df) else df
        reasons = df["reasons"] if len(df) else pd.Series([], dtype=object)
        n_dl04 = int(reasons.apply(lambda rs: any(c == "DL-04" for c, _ in rs)).sum()) if len(df) else 0
        n_dl03 = int(reasons.apply(lambda rs: any(c == "DL-03" for c, _ in rs)).sum()) if len(df) else 0
        cash = self.cash["p1"] + self.cash["p2"]

        cards = [
            dict(title="ยอดขายวันนี้ (ก่อน VAT)", big=_fmt(today_amt, 0), color="#111827",
                 bar=(today_amt / dt if dt else None, "#1D4ED8", None),
                 sub=(f"{today_n} SO · {today_amt / dt * 100:.0f}% ของเป้าวัน ≈ {dt:,.0f}" if dt else f"{today_n} SO · ยังไม่ได้ตั้งเป้า")),
            dict(title="ยอดสะสมรอบเดือน (MTD)", big=_fmt(mtd, 0), color="#111827",
                 bar=(mtd / self.total_target if self.total_target else None, "#16A34A" if (self.total_target and mtd >= exp) else "#1D4ED8",
                      (exp / self.total_target) if self.total_target else None),
                 sub=(f"เป้า {self.total_target / 1e6:.2f} ล้าน · วันทำงานที่ {done_days}/{len(self.work_days)} ควรได้ ≈ {exp / 1e6:.2f} ล้าน"
                      if self.total_target else "ยังไม่ได้ตั้งเป้า")),
            dict(title="เงินรับวันนี้ (รวม VAT)", big=_fmt(cash, 0), color="#166534", bar=None,
                 sub=f"{self.cash['n']} SO · งวดแรก {self.cash['p1']:,.0f} · งวดหลัง {self.cash['p2']:,.0f}"),
            dict(title="ค้างชำระ", big=_fmt(float(owe["outstanding"].sum()) if len(owe) else 0.0, 0), color="#B45309", bar=None,
                 sub=f"{len(owe)} SO · มัดจำแล้ว {n_part} · ยังไม่ชำระ {n_unpaid}"),
            dict(title="รอจัดซื้อรับงาน", big=f"{len(wait)} ใบ", color="#B45309", bar=None,
                 sub=f"มูลค่า {float(wait['goods'].sum()) if len(wait) else 0:,.0f}"),
            dict(title="เสี่ยงส่งไม่ทัน", big=f"{n_dl04 + n_dl03} ใบ", color="#B91C1C", bar=None,
                 alert=(n_dl04 + n_dl03) > 0, sub=f"ส่งภายใน D-1 ยังไม่มี PO {n_dl04} · เลยกำหนด {n_dl03}"),
        ]
        for i, c in enumerate(cards):
            cx = x + i * (cw + gap)
            alert = c.get("alert", False)
            self._card(cx, y, cw, ch, fill="#FEE2E2" if alert else "white", outline="#FCA5A5" if alert else "#E5E7EB")
            bc.text(cx + S(14), y + S(10), c["title"], size=9, fill="#B91C1C" if alert else "#6B7280")
            bc.text(cx + S(14), y + S(28), c["big"], size=19, bold=True, fill=c["color"])
            if c["bar"] is not None and c["bar"][0] is not None:
                pct, color, marker = c["bar"]
                bx, bw = cx + S(14), cw - S(28)
                bc.bar(bx, y + S(66), bw, S(7), pct, color=color)
                if marker is not None:
                    mx = bx + bw * max(0.0, min(marker, 1.0))
                    bc.line(mx, y + S(62), mx, y + S(77), fill="#111827", width=2)
            bc.text(cx + S(14), y + S(82), c["sub"], size=8, fill="#B91C1C" if alert else "#6B7280", width=cw - S(28))
        return y + ch

    # ---- Blended Gross Margin (forecast) ----
    def _draw_gm(self, x, y, w):
        bc = self.board
        S = bc.S
        g = self.gm
        h = S(112)
        self._card(x, y, w, h)
        bc.text(x + S(16), y + S(12), "Blended Gross Margin (forecast)", size=13, bold=True)
        bc.text(x + S(16), y + S(34), "SO ที่จองในรอบนี้ · ยืนยันด้วยต้นทุน PO จริง + ประมาณส่วนที่ยังไม่เปิด PO ด้วย margin ที่ยืนยันแล้ว",
                size=8, fill="#6B7280")
        if g["blended"] is None:
            bc.text(x + S(16), y + S(60), "ยังไม่มี SO ที่เปิด PO ครบในรอบนี้ จึงยังประมาณ margin ไม่ได้", size=10, fill="#9CA3AF")
            return y + h
        gp_total = g["gp_conf"] + g["gp_forecast"]
        cols = [
            ("Blended GM", f"{g['blended'] * 100:.1f}%", "#1D4ED8", f"GP รวม ≈ {gp_total:,.0f} จากยอด {g['sales_total']:,.0f}"),
            ("ยืนยันแล้ว (มี PO ครบ)", f"{g['confirmed_gm'] * 100:.1f}%", "#166534", f"{g['n_conf']} SO · GP {g['gp_conf']:,.0f}"),
            ("ประมาณการ (ยังไม่เปิด PO ครบ)", f"≈ {g['gp_forecast']:,.0f}", "#B45309",
             f"{g['n_fc']} SO · ยอด {g['sales_total'] - g['sales_conf']:,.0f}"),
            ("ความน่าเชื่อถือ (สัดส่วนยอดที่ยืนยัน)", f"{g['coverage'] * 100:.0f}%", "#111827", "ยิ่งต่ำ ตัวเลขยิ่งเป็นการประมาณมาก"),
        ]
        cw = (w - S(32)) / len(cols)
        for i, (t, big, color, sub) in enumerate(cols):
            cx = x + S(16) + i * cw
            bc.text(cx, y + S(56), t, size=9, fill="#6B7280")
            bc.text(cx, y + S(70), big, size=17, bold=True, fill=color)
            bc.text(cx, y + S(94), sub, size=8, fill="#6B7280", width=cw - S(12))
        return y + h

    # ---- ต้องทำวันนี้ ----
    def _draw_todo(self, x, y, w):
        bc = self.board
        S = bc.S
        df = self.df
        bg = self._card(x, y, w, 10)
        bc.text(x + S(16), y + S(14), "ต้องทำวันนี้", size=14, bold=True)
        todo = df[df["level"].isin(["urgent", "watch"])] if len(df) else df
        n_u = int((todo["level"] == "urgent").sum()) if len(todo) else 0
        n_w = int((todo["level"] == "watch").sum()) if len(todo) else 0
        dq = df[df["dq_diff"].notna()] if len(df) else df
        cx = x + w - S(16)
        chips = []
        if len(dq):
            chips.append((f"ยอดต่างรอตรวจ {len(dq)} ใบ · {abs(dq['dq_diff']).sum():,.0f}", "#FEF3C7", "#92400E"))
        chips.append((f"เฝ้าระวัง {n_w}", LEVEL_STYLE["watch"][1], LEVEL_STYLE["watch"][2]))
        chips.append((f"ด่วน {n_u}", LEVEL_STYLE["urgent"][1], LEVEL_STYLE["urgent"][2]))
        for txt, bgc, fg in chips:                           # วาดจากขวาไปซ้าย
            wd, _h = bc.badge(cx, y + S(12), txt, bgc, fg, size=9, anchor="ne", padx=10, pady=3)
            cx -= wd + S(6)
        cy = y + S(48)
        bc.line(x + S(16), cy - S(6), x + w - S(16), cy - S(6))
        if todo.empty:
            bc.text(x + w / 2, cy + S(10), "ไม่มีรายการที่ต้องติดตามวันนี้ 🎉", size=10, fill="#9CA3AF", anchor="n")
            cy += S(50)
        order = {"urgent": 0, "watch": 1}
        rows = todo.assign(_o=todo["level"].map(order)).sort_values(["_o", "delivery", "goods"], ascending=[True, True, False]).head(TODO_ROWS) \
            if len(todo) else todo
        for r in rows.itertuples():
            lv = LEVEL_STYLE[r.level]
            top = cy
            bc.badge(x + S(16), top + S(8), lv[0], lv[1], lv[2], size=9, padx=12, pady=3)
            ox = x + S(104)
            _, bb = bc.text(ox, top + S(6), r.so_number, size=10, bold=True)
            _, bb2 = bc.text(ox, bb[3], clean_text(r.customer_name) or "-", size=8, fill="#6B7280", width=w * 0.22)
            reason = (f"{r.reasons[0][1]}" + (f" · {len(r.reasons)} เหตุ" if len(r.reasons) > 1 else "")) if r.reasons else ""
            _, bb3 = bc.text(x + w * 0.42, top + S(10), reason, size=9, width=w * 0.36)
            pu = r.pu if r.pu else UNASSIGNED
            bc.text(x + w - S(16), top + S(10), f"{pu} · {r.sale_key or '-'}", size=9,
                    fill="#B91C1C" if not r.pu else "#6B7280", anchor="ne")
            h = max(bb2[3], bb3[3], top + S(40)) - top + S(8)
            cy = top + h
            bc.line(x + S(16), cy - S(2), x + w - S(16), cy - S(2), fill="#F3F4F6")
        if len(todo) > TODO_ROWS:
            bc.text(x + S(16), cy + S(4), f"และอีก {len(todo) - TODO_ROWS} รายการ — ดูทั้งหมดที่แท็บ 'ภาพรวมฝ่ายขาย' / 'กระดานงานจัดซื้อ'",
                    size=9, fill="#6B7280")
            cy += S(24)
        bottom = cy + S(10)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    # ---- สถานะงานจัดซื้อ ----
    def _draw_status(self, x, y, w):
        bc = self.board
        S = bc.S
        df = self.df
        bg = self._card(x, y, w, 10)
        bc.text(x + S(16), y + S(14), f"สถานะงานจัดซื้อ ({len(df)} SO)", size=13, bold=True)
        cy = y + S(48)
        counts = [int((df["status"] == st).sum()) if len(df) else 0 for st in ("Pending PU", "PO In Progress", "PO Sent")]
        tot = sum(counts)
        bx, bw, bh = x + S(16), w - S(32), S(26)
        if tot:
            acc = 0.0
            for (key, label, color), c in zip(COLS, counts):
                if c:
                    seg = bw * c / tot
                    bc.rect(bx + acc, cy, bx + acc + seg, cy + bh, fill=color)
                    acc += seg
        else:
            bc.rect(bx, cy, bx + bw, cy + bh, fill="#E5E7EB")
        cy += bh + S(16)
        for (key, label, color), c in zip(COLS, counts):
            bc.text(bx, cy, "■", size=10, fill=color)
            bc.text(bx + S(18), cy, label, size=10)
            bc.text(bx + bw, cy, str(c), size=10, bold=True, anchor="ne")
            cy += S(24)
        bottom = cy + S(8)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    # ---- การกระจุกตัวของยอด ----
    def _draw_concentration(self, x, y, w):
        bc = self.board
        S = bc.S
        t = self.today_df
        bg = self._card(x, y, w, 10)
        bc.text(x + S(16), y + S(14), "การกระจุกตัวของยอด", size=13, bold=True)
        share, rest, total = concentration(t["goods"] if len(t) else [], 3)
        cy = y + S(50)
        if total <= 0:
            bc.text(x + w / 2, cy, "วันนี้ยังไม่มียอดขาย", size=10, fill="#9CA3AF", anchor="n")
            bottom = cy + S(40)
        else:
            bc.text(x + S(16), cy, f"{share * 100:.1f}%", size=22, bold=True, fill="#B45309")
            cy += S(44)
            bc.text(x + S(16), cy, f"ของยอดวันนี้มาจาก {min(3, len(t))} SO ใหญ่สุด", size=9, fill="#6B7280")
            cy += S(24)
            for r in t.sort_values("goods", ascending=False).head(3).itertuples():
                bc.text(x + S(16), cy, f"{r.so_number} · {r.sale_key or '-'}", size=9)
                bc.text(x + w - S(16), cy, f"{r.goods:,.0f}", size=9, anchor="ne")
                cy += S(22)
            bc.line(x + S(16), cy, x + w - S(16), cy)
            bc.text(x + S(16), cy + S(8), f"ไม่รวม {max(len(t) - 3, 0)} ใบ ยอดวันนี้ = {rest:,.0f}", size=9, fill="#6B7280")
            bottom = cy + S(36)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    # ---- ยอดขายรายวันตลอดรอบ ----
    def _draw_chart(self, x, y, w, rep):
        bc = self.board
        S = bc.S
        start, end = cycle_bounds(rep)
        bg = self._card(x, y, w, 10)
        bc.text(x + S(16), y + S(14), f"ยอดขายรายวัน · รอบ {start.day} {_TH[start.month]} – {end.day} {_TH[end.month]}", size=13, bold=True)
        bc.text(x + w - S(16), y + S(18), "เส้นประ = เป้าวัน · แท่งสีเทา = ยอดจริงรายวัน · แท่งสีน้ำเงิน = วันที่รายงาน", size=8,
                fill="#6B7280", anchor="ne")
        days = []
        d = start
        while d <= end:
            days.append(d)
            d += timedelta(days=1)
        by_day = self.sold.groupby("rec_date")["goods"].sum().to_dict() if len(self.sold) else {}
        dt = daily_target(self.total_target, self.work_days)
        top_v = max([float(v) for v in by_day.values()] + [dt, 1.0]) * 1.12
        px, pw = x + S(24), w - S(48)
        py, ph = y + S(56), S(150)
        slot = pw / len(days)
        bar_w = max(slot * 0.68, 2)
        for i, day in enumerate(days):
            bx = px + i * slot + (slot - bar_w) / 2
            v = float(by_day.get(day, 0.0))
            if day <= rep:
                hbar = ph * v / top_v
                color = "#1D4ED8" if day == rep else "#D1D5DB"
                if hbar > 0:
                    bc.rect(bx, py + ph - hbar, bx + bar_w, py + ph, fill=color)
                else:
                    bc.rect(bx, py + ph - S(3), bx + bar_w, py + ph, fill="#E5E7EB")
            bc.text(px + i * slot + slot / 2, py + ph + S(6), str(day.day), size=8,
                    fill="#111827" if day == rep else "#6B7280", bold=(day == rep), anchor="n")
        bc.line(px, py + ph, px + pw, py + ph, fill="#CBD5E1")
        if dt > 0:
            ty = py + ph - ph * dt / top_v
            bc.canvas.create_line(px, ty, px + pw, ty, fill="#B45309", width=2, dash=(5, 3))
            bc.text(px + pw, ty - S(4), f"เป้าวัน {dt:,.0f}", size=8, fill="#B45309", anchor="se")
        bottom = py + ph + S(34)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    # ---- ยอดขายรายเซลส์วันนี้ ----
    def _draw_by_sale(self, x, y, w):
        bc = self.board
        S = bc.S
        bg = self._card(x, y, w, 10)
        bc.text(x + S(16), y + S(14), "ยอดขายรายเซลส์วันนี้", size=13, bold=True)
        cy = y + S(50)
        t = self.today_df
        if t.empty:
            bc.text(x + w / 2, cy, "วันนี้ยังไม่มียอดขาย", size=10, fill="#9CA3AF", anchor="n")
            bottom = cy + S(40)
        else:
            g = t.groupby("sale_key")["goods"].sum().sort_values(ascending=False)
            total, mx = float(g.sum()), float(g.max())
            for sale, amt in g.items():
                bc.text(x + S(16), cy, str(sale), size=10, bold=True)
                bc.text(x + w - S(16), cy, f"{amt:,.0f} · {amt / total * 100:.1f}%", size=10, anchor="ne")
                cy += S(22)
                bc.bar(x + S(16), cy, w - S(32), S(8), amt / mx, color="#1D4ED8")
                cy += S(22)
            bottom = cy + S(4)
        bc.set_rrect(bg, x, y, x + w, bottom, r=S(12))
        return bottom

    # ---------------------------------------------------------------- export
    def export_excel(self):
        if self.df is None or self.df.empty:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่มีข้อมูลให้ Export", parent=self)
            return
        rep = self._report_date()
        start, end = cycle_bounds(rep)
        dt = daily_target(self.total_target, self.work_days)
        exp = expected_to_date(self.total_target, self.work_days, rep)
        kpi = pd.DataFrame([
            ("วันที่รายงาน", f"{rep.day:02d}/{rep.month:02d}/{rep.year + 543}"),
            ("รอบ", f"{start.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"),
            ("ยอดขายวันนี้ (ก่อน VAT)", float(self.today_df["goods"].sum())),
            ("เป้าวัน", dt),
            ("ยอดสะสมรอบเดือน (MTD)", float(self.sold["goods"].sum())),
            ("เป้ารวม", self.total_target),
            ("ยอดที่ควรได้ถึงวันนี้ (ตามวันทำงาน)", exp),
            ("เงินรับวันนี้ (รวม VAT)", self.cash["p1"] + self.cash["p2"]),
            ("ค้างชำระ (งานในรายการ)", float(self.df["outstanding"].sum())),
            ("Blended Gross Margin (forecast)", self.gm["blended"]),
            ("GM ยืนยันแล้ว (SO ที่ PO ครบ)", self.gm["confirmed_gm"]),
            ("สัดส่วนยอดที่ยืนยันแล้ว", self.gm["coverage"]),
        ], columns=["รายการ", "ค่า"])
        lvname = {"urgent": "ด่วน", "watch": "เฝ้าระวัง", "normal": "ปกติ"}
        todo = self.df[self.df["level"].isin(["urgent", "watch"])]
        todo_out = pd.DataFrame({
            "เลขที่ SO": todo["so_number"], "ระดับ": todo["level"].map(lvname),
            "เหตุ": todo["reasons"].apply(lambda rs: " | ".join(f"{c} {t}" for c, t in rs)),
            "ลูกค้า": todo["customer_name"].apply(clean_text), "เซลส์": todo["sale_key"],
            "ผู้รับงาน (PU)": todo["pu"].fillna(UNASSIGNED), "ยอดสินค้า": todo["goods"],
            "วันส่ง": todo["delivery"].apply(lambda d: f"{d.day:02d}/{d.month:02d}/{d.year + 543}"),
            "D-x": todo["n"].apply(lambda n: _d_label(int(n))),
        })
        daily = self.sold.groupby("rec_date")["goods"].sum().reset_index().rename(columns={"rec_date": "วันที่", "goods": "ยอดจอง"})
        by_sale = self.today_df.groupby("sale_key")["goods"].sum().reset_index().rename(columns={"sale_key": "เซลส์", "goods": "ยอดวันนี้"})
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                            initialfile=f"ภาพรวมผู้บริหาร_{rep.strftime('%Y%m%d')}.xlsx", parent=self)
        if path:
            try:
                with pd.ExcelWriter(path) as xw:
                    kpi.to_excel(xw, index=False, sheet_name="สรุป")
                    todo_out.to_excel(xw, index=False, sheet_name="ต้องทำวันนี้")
                    daily.to_excel(xw, index=False, sheet_name="ยอดรายวัน")
                    by_sale.to_excel(xw, index=False, sheet_name="รายเซลส์")
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
                    self.date_sel.set_date(date.today())
                    self._auto_date = date.today()
                self.load_data(silent=True)
            self._schedule_refresh()
        except Exception:
            pass


_TH = {1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค."}
