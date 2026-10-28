"""แท็บ "6. แจ้งเตือน/SO ตกค้าง" ในหน้า HR (Director + HR เห็น)

ให้ผู้จ่ายเงิน/กรรมการเห็นต่อเนื่อง (ไม่ใช่ popup ที่อ่านแล้วหาย):
  • เหตุการณ์ที่ SO/PO ถูกแก้หลัง HR Verified/Paid — ใคร แก้อะไร สถานะเปลี่ยนยังไง (ตาราง so_post_verify_events)
  • SO ที่ตกค้างอยู่ — ถูกแก้หลังตรวจแล้วกลับไปรอ HR ตรวจใหม่ และ SO ของเดือนก่อนที่ยังไม่ถูกจ่าย
"""
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
from customtkinter import CTkFrame, CTkLabel, CTkButton, CTkFont, CTkCheckBox, CTkOptionMenu

from post_verify_alerts import FINAL_STATUSES

EVENT_SHORT = {
    'PO_AFTER_VERIFY':        'PO เพิ่ม/แก้ หลัง HR ตรวจ → รอตรวจใหม่',
    'PO_ADDED_AFTER_PAID':    'PO เพิ่ม/แก้ หลังจ่ายแล้ว',
    'PO_EDITED_AFTER_VERIFY': 'แก้ PO หลัง HR ตรวจ → รอตรวจใหม่',
    'PO_EDITED_AFTER_PAID':   'แก้ PO หลังจ่ายแล้ว',
    'SO_EDITED_AFTER_VERIFY': 'แก้ SO หลัง HR ตรวจ → รอตรวจใหม่',
}
VERIFY_TYPES = ('PO_AFTER_VERIFY', 'PO_EDITED_AFTER_VERIFY', 'SO_EDITED_AFTER_VERIFY')
# แสดง SO ตกค้างเฉพาะช่วงล่าสุด — ข้อมูลค้างเก่าตั้งแต่ยุคที่ระบบยังไม่ดี (เช่น ปี 2568) ไม่ใช่ SO ที่ตกหล่นจริงตอนนี้
DEFAULT_LOOKBACK_MONTHS = 3
LOOKBACK_OPTIONS = {"3 เดือนก่อนหน้า": 3, "6 เดือนก่อนหน้า": 6, "ทั้งหมด (รวมข้อมูลเก่า)": None}
THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


class PostVerifyAlertsTab(CTkFrame):

    def __init__(self, master, app_container, user_key=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.user_key = user_key
        self._event_ids = {}      # iid -> event id
        self.only_unacked = tk.BooleanVar(value=True)
        self.only_verified_carry = tk.BooleanVar(value=False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # ── การ์ดสรุป ──
        self.cards = CTkFrame(self, fg_color="transparent")
        self.cards.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        for i in range(3):
            self.cards.grid_columnconfigure(i, weight=1)
        self._card_labels = {}
        for i, (key, title, color) in enumerate([
            ("unacked", "เหตุการณ์ที่ยังไม่รับทราบ", "#B45309"),
            ("stuck", "รอ HR ตรวจใหม่ (ถูกแก้หลังตรวจแล้ว)", "#DC2626"),
            ("carry", "SO ค้างจากเดือนก่อน (ยังไม่จ่าย)", "#2563EB"),
        ]):
            card = CTkFrame(self.cards, fg_color="white", corner_radius=10, border_width=1, border_color="#E2E8F0")
            card.grid(row=0, column=i, sticky="ew", padx=5)
            CTkLabel(card, text=title, font=CTkFont(size=12), text_color="#64748B").pack(anchor="w", padx=14, pady=(10, 0))
            lbl = CTkLabel(card, text="–", font=CTkFont(size=26, weight="bold"), text_color=color)
            lbl.pack(anchor="w", padx=14, pady=(0, 10))
            self._card_labels[key] = lbl

        # ── ตารางเหตุการณ์ ──
        head1 = CTkFrame(self, fg_color="transparent")
        head1.grid(row=1, column=0, sticky="ew", padx=6, pady=(8, 2))
        CTkLabel(head1, text="📝 เหตุการณ์แก้ไข SO/PO หลัง HR ตรวจแล้ว (SO ที่จ่ายแล้วถูกล็อก แก้ไม่ได้)",
                 font=CTkFont(size=14, weight="bold")).pack(side="left")
        CTkButton(head1, text="⟳ รีเฟรช", width=80, fg_color="gray", command=self.load).pack(side="right")
        CTkButton(head1, text="✔ รับทราบทั้งหมดที่แสดง", width=170, fg_color="#16A34A", hover_color="#15803D",
                  command=lambda: self._ack(all_shown=True)).pack(side="right", padx=6)
        CTkButton(head1, text="✔ รับทราบที่เลือก", width=130, fg_color="#2563EB", hover_color="#1D4ED8",
                  command=lambda: self._ack(all_shown=False)).pack(side="right", padx=6)
        CTkCheckBox(head1, text="เฉพาะที่ยังไม่รับทราบ", variable=self.only_unacked,
                    command=self.load).pack(side="right", padx=10)

        ev_cols = ("time", "so", "event", "po", "status", "by", "cur", "ack")
        self.ev_tree = self._make_tree(row=2, columns=ev_cols, headers={
            "time": "เวลา", "so": "SO", "event": "เหตุการณ์", "po": "PO", "status": "สถานะ SO (ก่อน → หลัง)",
            "by": "ผู้ทำรายการ", "cur": "สถานะ SO ตอนนี้", "ack": "รับทราบโดย"},
            widths={"time": 120, "so": 120, "event": 260, "po": 120, "status": 190, "by": 110, "cur": 130, "ack": 130})
        self.ev_tree.tag_configure("stuck", background="#FEF3C7")
        self.ev_tree.tag_configure("done", background="#DCFCE7")
        self.ev_tree.tag_configure("acked", foreground="#94A3B8")
        self.ev_tree.bind("<Double-1>", self._show_event_detail)

        # ── ตาราง SO ตกค้าง ──
        head2 = CTkFrame(self, fg_color="transparent")
        head2.grid(row=3, column=0, sticky="ew", padx=6, pady=(10, 2))
        CTkLabel(head2, text="📌 SO ตกค้าง — ของเดือนก่อนหน้าที่ยังไม่ถูกจ่ายค่าคอม",
                 font=CTkFont(size=14, weight="bold")).pack(side="left")
        CTkCheckBox(head2, text="เฉพาะที่เคยผ่าน HR ตรวจแล้ว (น่าห่วงที่สุด)", variable=self.only_verified_carry,
                    command=self.load).pack(side="right", padx=10)
        self.lookback_var = tk.StringVar(value=list(LOOKBACK_OPTIONS.keys())[0])
        CTkOptionMenu(head2, variable=self.lookback_var, values=list(LOOKBACK_OPTIONS.keys()), width=190,
                      command=lambda _v: self.load()).pack(side="right", padx=6)
        CTkLabel(head2, text="ช่วงเวลา:", font=CTkFont(size=12), text_color="#64748B").pack(side="right")

        carry_cols = ("so", "sale", "customer", "period", "status", "verified")
        self.carry_tree = self._make_tree(row=4, columns=carry_cols, headers={
            "so": "SO", "sale": "เซลส์", "customer": "ลูกค้า", "period": "รอบค่าคอม", "status": "สถานะ",
            "verified": "เคย HR ตรวจแล้ว"},
            widths={"so": 130, "sale": 110, "customer": 320, "period": 100, "status": 140, "verified": 120})
        self.carry_tree.tag_configure("verified", background="#FEE2E2")
        self.after(150, self.load)

    # ---------------------------------------------------------------- UI helper
    def _make_tree(self, row, columns, headers, widths):
        frame = CTkFrame(self, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="nsew", padx=6, pady=(0, 2))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=8, selectmode="extended")
        for c in columns:
            tree.heading(c, text=headers[c])
            tree.column(c, width=widths[c], anchor="w")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        return tree

    @staticmethod
    def _fmt_period(m, y):
        try:
            return f"{THAI_MONTHS[int(m) - 1]} {int(y) + 543}"
        except Exception:
            return "-"

    # ---------------------------------------------------------------- data
    def load(self):
        try:
            self._load_events()
            self._load_carry_over()
        except Exception as e:
            messagebox.showerror("Error", f"โหลดข้อมูลแจ้งเตือนไม่สำเร็จ: {e}", parent=self)

    def _load_events(self):
        df = pd.read_sql_query("""
            SELECT e.id, e.changed_at, e.so_number, e.event_type, e.po_number, e.old_status, e.new_status,
                   e.detail, e.changed_by, e.ack_by, e.ack_at, c.status AS cur_status
            FROM so_post_verify_events e
            LEFT JOIN commissions c ON c.so_number = e.so_number AND c.is_active = 1
            WHERE (e.sale_key NOT IN ('Sale Center', 'CHARITA-CT') AND EXISTS (SELECT 1 FROM sales_users su_ok WHERE su_ok.sale_key = e.sale_key AND su_ok.status = 'Active'))
            ORDER BY e.changed_at DESC
            LIMIT 500
        """, self.pg_engine)
        self._events_df = df

        stuck_sos = set(df.loc[df["event_type"].isin(VERIFY_TYPES) & ~df["cur_status"].isin(FINAL_STATUSES),
                               "so_number"])
        self._card_labels["unacked"].configure(text=f"{int(df['ack_at'].isna().sum()):,}")
        self._card_labels["stuck"].configure(text=f"{len(stuck_sos):,}")

        shown = df[df["ack_at"].isna()] if self.only_unacked.get() else df
        self.ev_tree.delete(*self.ev_tree.get_children())
        self._event_ids = {}
        for _, r in shown.iterrows():
            acked = pd.notna(r["ack_at"])
            is_verify = r["event_type"] in VERIFY_TYPES
            cur = r["cur_status"] if pd.notna(r["cur_status"]) else "-"
            if is_verify and cur not in FINAL_STATUSES:
                tags = ("stuck",)
            elif is_verify and cur in FINAL_STATUSES:
                tags = ("done",)
            else:
                tags = ()
            if acked:
                tags = tags + ("acked",)
            when = r["changed_at"].strftime("%d/%m/%y %H:%M") if pd.notna(r["changed_at"]) else "-"
            iid = self.ev_tree.insert("", "end", tags=tags, values=(
                when, r["so_number"], EVENT_SHORT.get(r["event_type"], r["event_type"]),
                r["po_number"] or "-", f"{r['old_status']} → {r['new_status']}", r["changed_by"] or "-", cur,
                (f"{r['ack_by']} ({r['ack_at'].strftime('%d/%m %H:%M')})" if acked else "ยังไม่รับทราบ")))
            self._event_ids[iid] = int(r["id"])

    def _load_carry_over(self):
        now = datetime.now()
        df = pd.read_sql_query("""
            SELECT c.so_number, c.sale_key, c.customer_name, c.status, c.commission_month, c.commission_year,
                   (c.final_sales_amount IS NOT NULL) AS was_verified
            FROM commissions c
            WHERE c.is_active = 1
              AND c.status IN ('Pending PU', 'PO In Progress', 'PO Sent', 'HR Verified', 'Forwarded_To_HR')
              AND (c.sale_key NOT IN ('Sale Center', 'CHARITA-CT') AND EXISTS (SELECT 1 FROM sales_users su_ok WHERE su_ok.sale_key = c.sale_key AND su_ok.status = 'Active'))
              AND (c.commission_year < %(y)s OR (c.commission_year = %(y)s AND c.commission_month < %(m)s))
              AND (%(lb)s IS NULL OR (c.commission_year * 12 + c.commission_month) >= (%(y)s * 12 + %(m)s - %(lb)s))
            ORDER BY c.commission_year DESC, c.commission_month DESC, c.so_number
        """, self.pg_engine, params={"y": now.year, "m": now.month,
                                     "lb": LOOKBACK_OPTIONS.get(self.lookback_var.get(), DEFAULT_LOOKBACK_MONTHS)})
        self._card_labels["carry"].configure(text=f"{len(df):,}")
        shown = df[df["was_verified"]] if self.only_verified_carry.get() else df
        self.carry_tree.delete(*self.carry_tree.get_children())
        for _, r in shown.iterrows():
            self.carry_tree.insert("", "end", tags=("verified",) if r["was_verified"] else (), values=(
                r["so_number"], r["sale_key"], r["customer_name"] or "-",
                self._fmt_period(r["commission_month"], r["commission_year"]), r["status"],
                "✔ เคยผ่านแล้ว" if r["was_verified"] else "-"))

    # ---------------------------------------------------------------- actions
    def _ack(self, all_shown):
        if all_shown:
            ids = list(self._event_ids.values())
        else:
            ids = [self._event_ids[i] for i in self.ev_tree.selection() if i in self._event_ids]
        if not ids:
            messagebox.showinfo("ไม่มีรายการ", "ไม่มีรายการให้รับทราบ" if all_shown else "กรุณาเลือกรายการก่อน", parent=self)
            return
        if all_shown and not messagebox.askyesno("ยืนยัน", f"รับทราบทั้งหมด {len(ids)} รายการที่แสดงอยู่?", parent=self):
            return
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cur:
                cur.execute("UPDATE so_post_verify_events SET ack_by = %s, ack_at = NOW() "
                            "WHERE id = ANY(%s) AND ack_at IS NULL", (self.user_key, ids))
            conn.commit()
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("Error", f"บันทึกการรับทราบไม่สำเร็จ: {e}", parent=self)
        finally:
            if conn:
                self.app_container.release_connection(conn)
        self.load()

    def _show_event_detail(self, event=None):
        sel = self.ev_tree.selection()
        if not sel or sel[0] not in self._event_ids:
            return
        eid = self._event_ids[sel[0]]
        r = self._events_df[self._events_df["id"] == eid].iloc[0]
        txt = (f"SO: {r['so_number']}\nเหตุการณ์: {EVENT_SHORT.get(r['event_type'], r['event_type'])}\n"
               f"PO: {r['po_number'] or '-'}\nสถานะ SO: {r['old_status']} → {r['new_status']}\n"
               f"ผู้ทำรายการ: {r['changed_by'] or '-'}\nเวลา: {r['changed_at']}\n"
               f"รายละเอียด: {r['detail'] or '-'}\n"
               f"สถานะ SO ตอนนี้: {r['cur_status'] if pd.notna(r['cur_status']) else '-'}")
        messagebox.showinfo("รายละเอียดเหตุการณ์", txt, parent=self)


# ───────────────────────── ใช้ร่วมกับหน้า "คำนวณ & จ่าย" และ popup ยืนยันจ่าย ─────────────────────────
EVENT_PLAIN = {
    'PO_AFTER_VERIFY':        'เพิ่ม/แก้ PO หลัง HR ตรวจ',
    'PO_EDITED_AFTER_VERIFY': 'แก้ PO หลัง HR ตรวจ',
    'SO_EDITED_AFTER_VERIFY': 'เซลส์แก้ SO หลัง HR ตรวจ',
    'PO_ADDED_AFTER_PAID':    'แก้ PO หลังจ่ายแล้ว',
    'PO_EDITED_AFTER_PAID':   'แก้ PO หลังจ่ายแล้ว',
}
STATUS_PLAIN = {
    'Pending PU': 'รอจัดซื้อรับงาน',
    'PO In Progress': 'จัดซื้อกำลังทำ PO',
    'PO Sent': 'รอ HR ตรวจ',
    'Forwarded_To_HR': 'รอ HR ตรวจ',
    'HR Verified': 'HR ตรวจแล้ว รอจ่าย',
    'Paid': 'จ่ายแล้ว',
}


def fetch_post_verify_context(engine, so_list, ids, sale_key, year, month, lookback_months=DEFAULT_LOOKBACK_MONTHS):
    """คืน (events_df, not_included_df) สำหรับกล่อง "ข้อควรตรวจก่อนจ่าย"
    events_df: เหตุการณ์ของ SO ที่อยู่ในรอบจ่ายนี้ (พร้อมยอดขาย)
    not_included_df: SO ของเซลส์คนนี้ที่ไม่ได้อยู่ในรอบจ่ายนี้ (รอ HR ตรวจ/ค้างเดือนก่อน)"""
    events = pd.read_sql_query("""
        SELECT e.id, e.so_number, e.event_type, e.changed_by, e.changed_at, e.detail, e.ack_at,
               (SELECT c.sales_service_amount FROM commissions c
                 WHERE c.so_number = e.so_number AND c.is_active = 1 LIMIT 1) AS sales_amount
        FROM so_post_verify_events e
        WHERE e.so_number = ANY(%(sos)s) AND (e.sale_key NOT IN ('Sale Center', 'CHARITA-CT') AND EXISTS (SELECT 1 FROM sales_users su_ok WHERE su_ok.sale_key = e.sale_key AND su_ok.status = 'Active'))
        ORDER BY e.changed_at DESC
    """, engine, params={"sos": so_list})
    stuck = pd.read_sql_query("""
        SELECT so_number, status, commission_month, commission_year, sales_service_amount AS sales_amount,
               (final_sales_amount IS NOT NULL) AS was_verified
        FROM commissions c
        WHERE c.sale_key = %(sk)s AND c.is_active = 1 AND (c.sale_key NOT IN ('Sale Center', 'CHARITA-CT') AND EXISTS (SELECT 1 FROM sales_users su_ok WHERE su_ok.sale_key = c.sale_key AND su_ok.status = 'Active'))
          AND status IN ('Pending PU', 'PO In Progress', 'PO Sent', 'Forwarded_To_HR')
          AND id <> ALL(%(ids)s)
          AND (commission_year < %(y)s OR (commission_year = %(y)s AND commission_month <= %(m)s))
          AND (commission_year * 12 + commission_month) >= (%(y)s * 12 + %(m)s - %(lb)s)
        ORDER BY (final_sales_amount IS NOT NULL) DESC, commission_year, commission_month, so_number
    """, engine, params={"sk": sale_key, "ids": ids, "y": year, "m": month, "lb": lookback_months})
    return events, stuck


def headline_text(events, stuck, sale_key):
    parts = []
    if not events.empty:
        parts.append(f"{events['so_number'].nunique()} SO ในรอบนี้ถูกแก้หลัง HR ตรวจ")
    if not stuck.empty:
        n_ver = int(stuck['was_verified'].sum())
        t = f"{len(stuck)} SO ของ {sale_key} ไม่ได้อยู่ในรอบจ่ายนี้ (เซลส์จะยังไม่ได้ค่าคอมของ SO เหล่านี้รอบนี้)"
        if n_ver:
            t += f" — {n_ver} ใบเคยผ่าน HR ตรวจแล้ว"
        parts.append(t)
    return " · ".join(parts)


def plain_lines(events, stuck, sale_key, thai_months):
    """ข้อความสั้นแบบภาษาคน ไว้โชว์ใน popup ยืนยันจ่าย"""
    lines = []
    h = headline_text(events, stuck, sale_key)
    if h:
        lines.append(h)
    for _, r in events.head(5).iterrows():
        when = r['changed_at'].strftime('%d/%m %H:%M') if pd.notna(r['changed_at']) else '-'
        det = f" — {r['detail']}" if r['detail'] else ""
        lines.append(f"• {r['so_number']}: {EVENT_PLAIN.get(r['event_type'], r['event_type'])} "
                     f"โดย {r['changed_by'] or '-'} ({when}){det}")
    if len(events) > 5:
        lines.append(f"  …และอีก {len(events) - 5} เหตุการณ์")
    for _, r in stuck.head(5).iterrows():
        per = f"{thai_months[int(r['commission_month']) - 1]} {int(r['commission_year']) + 543}"
        star = " ★เคยผ่าน HR ตรวจแล้ว" if r['was_verified'] else ""
        lines.append(f"• {r['so_number']}: {STATUS_PLAIN.get(r['status'], r['status'])} (รอบ {per}){star}")
    if len(stuck) > 5:
        lines.append(f"  …และอีก {len(stuck) - 5} SO")
    return lines
