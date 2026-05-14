# markup_tiers_screen.py

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from customtkinter import (
    CTkFrame, CTkLabel, CTkFont, CTkButton, CTkEntry, CTkOptionMenu,
)
import psycopg2, psycopg2.extras

CLR = {
    "navy":   "#0C2340",
    "navy2":  "#0F2D50",
    "blue":   "#1A56DB",
    "white":  "#FFFFFF",
    "gray":   "#F9FAFB",
    "border": "#E5E7EB",
    "text":   "#1F2937",
    "sub":    "#6B7280",
    "dim":    "#9CA3AF",
}

TIER_ORDER = ["Guest", "T1", "T2", "T3", "T4", "T5"]
TIER_FG = {
    "Guest": "#1D4ED8",
    "T1":    "#15803D",
    "T2":    "#047857",
    "T3":    "#B45309",
    "T4":    "#6D28D9",
    "T5":    "#BE123C",
}
TIER_BG = {
    "Guest": "#DBEAFE",
    "T1":    "#DCFCE7",
    "T2":    "#D1FAE5",
    "T3":    "#FEF3C7",
    "T4":    "#EDE9FE",
    "T5":    "#FFE4E6",
}

# column id → tier name (columns #4–#9)
_COL_TO_TIER = {f"#{i+4}": t for i, t in enumerate(TIER_ORDER)}

_DB_CFG = dict(host="Server-APrime", dbname="aplus_com_test",
               user="app_user", password="cailfornia123")
_app_container = None
_TTK_STYLED = False


def _apply_style():
    global _TTK_STYLED
    if _TTK_STYLED: return
    _TTK_STYLED = True
    st = ttk.Style()
    try: st.theme_use("default")
    except: pass
    st.configure("MT.Treeview",
                 font=("Tahoma", 11), rowheight=26,
                 background=CLR["white"], fieldbackground=CLR["white"],
                 foreground=CLR["text"], borderwidth=0)
    st.configure("MT.Treeview.Heading",
                 font=("Tahoma", 11, "bold"),
                 background=CLR["navy"], foreground=CLR["white"],
                 relief="flat", borderwidth=0, padding=(6, 5))
    st.map("MT.Treeview",
           background=[("selected", "#BFDBFE")],
           foreground=[("selected", CLR["navy"])])


def _get_conn():
    if _app_container: return _app_container.get_connection(), True
    return psycopg2.connect(**_DB_CFG), False

def _release_conn(conn, pool):
    if pool and _app_container: _app_container.release_connection(conn)
    else: conn.close()

def db_fetch_categories():
    conn, pool = _get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT DISTINCT main_category FROM markup_tiers "
                      "WHERE is_active=TRUE AND main_category<>'' ORDER BY 1")
            return [r[0] for r in c.fetchall()]
    finally: _release_conn(conn, pool)

def db_fetch_pivoted(cat=None, search=""):
    """Returns one dict per product_type with per-tier data nested inside."""
    conn, pool = _get_conn()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as c:
            f, p = ["is_active=TRUE", "main_category<>''", "product_type<>''"], []
            if cat and cat != "ทั้งหมด":
                f.append("main_category=%s"); p.append(cat)
            if search.strip():
                f.append("(product_type ILIKE %s OR sub_category ILIKE %s)")
                p += [f"%{search}%", f"%{search}%"]
            c.execute(
                f"SELECT id,main_category,sub_category,product_type,"
                f"tier_name,markup_pct,amount_range,qty_range "
                f"FROM markup_tiers WHERE {' AND '.join(f)} "
                f"ORDER BY main_category,sub_category,product_type,tier_order", p)
            raw = [dict(r) for r in c.fetchall()]
    finally: _release_conn(conn, pool)

    pivoted: dict[tuple, dict] = {}
    order: list[tuple] = []
    for r in raw:
        key = (r["main_category"], r["sub_category"] or "", r["product_type"])
        if key not in pivoted:
            pivoted[key] = {
                "main_category": r["main_category"],
                "sub_category":  r["sub_category"] or "",
                "product_type":  r["product_type"],
            }
            order.append(key)
        pivoted[key][r["tier_name"]] = {
            "id":  r["id"],
            "pct": r["markup_pct"],
            "amt": r["amount_range"] or "",
            "qty": r["qty_range"]    or "",
        }
    return [pivoted[k] for k in order]

def db_update_markup(row_id, pct, by="MANAGER"):
    conn, pool = _get_conn()
    try:
        with conn.cursor() as c:
            c.execute("UPDATE markup_tiers SET markup_pct=%s,updated_at=NOW(),"
                      "excel_updated_by=%s WHERE id=%s", (pct, by, row_id))
        conn.commit()
    finally: _release_conn(conn, pool)


# ==============================================================================
class MarkupTiersScreen(CTkFrame):

    def __init__(self, master, app_container=None, current_user="MANAGER", **kwargs):
        super().__init__(master, fg_color=CLR["white"], corner_radius=0, **kwargs)
        global _app_container
        if app_container: _app_container = app_container
        self.current_user = current_user
        self._piv_rows: list[dict] = []
        self._piv_map:  dict[str, dict] = {}   # iid → pivoted row
        self._last_col: str = ""               # last clicked tier column

        _apply_style()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_topbar()
        self._build_table()
        self._load_categories()
        self._load_table()

    # ------------------------------------------------------------------
    def _build_topbar(self):
        bar = tk.Frame(self, bg=CLR["navy"])
        bar.grid(row=0, column=0, sticky="ew")

        left = tk.Frame(bar, bg=CLR["navy"])
        left.pack(side="left", padx=(14, 0), pady=5)

        CTkLabel(left, text="Markup Tiers",
                 font=CTkFont(size=13, weight="bold"),
                 text_color="#93C5FD", bg_color=CLR["navy"]).pack(side="left")
        CTkLabel(left, text="A-Prime",
                 font=CTkFont(size=11),
                 text_color="#4B8FCC", bg_color=CLR["navy"]).pack(side="left", padx=(6, 0))

        tk.Frame(left, bg="#1E3A5F", width=1).pack(side="left", padx=12, pady=2, fill="y")

        CTkLabel(left, text="หมวด:", text_color="#94A3B8",
                 font=CTkFont(size=12), bg_color=CLR["navy"]).pack(side="left", padx=(0, 4))
        self._cat_var = tk.StringVar(value="ทั้งหมด")
        self._cat_menu = CTkOptionMenu(
            left, variable=self._cat_var, width=150, height=28,
            fg_color="#1E3A5F", button_color="#2D5A8A",
            button_hover_color="#3A6EA8", text_color="#E2E8F0",
            font=CTkFont(size=12),
            command=lambda _: self._load_table())
        self._cat_menu.pack(side="left", padx=(0, 8))

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self.after(400, self._load_table))
        CTkEntry(left, textvariable=self._search_var, width=200, height=28,
                 fg_color="#1E3A5F", border_color="#2D5A8A",
                 text_color="#E2E8F0", font=CTkFont(size=12),
                 placeholder_text="ค้นหาสินค้า..."
                 ).pack(side="left", padx=(0, 8))

        self._lbl_count = CTkLabel(left, text="", text_color="#94A3B8",
                                   font=CTkFont(size=11), bg_color=CLR["navy"])
        self._lbl_count.pack(side="left")

        # tier legend pills (read-only, no toggle)
        tk.Frame(left, bg="#1E3A5F", width=1).pack(side="left", padx=12, pady=2, fill="y")
        for tier in TIER_ORDER:
            CTkLabel(left, text=f"■ {tier}",
                     text_color=TIER_FG[tier], font=CTkFont(size=11),
                     bg_color=CLR["navy"]).pack(side="left", padx=4)

        right = tk.Frame(bar, bg=CLR["navy"])
        right.pack(side="right", padx=10, pady=5)

        CTkButton(right, text="โหลดใหม่", width=76, height=28,
                  fg_color="#374151", hover_color="#1F2937",
                  font=CTkFont(size=12),
                  command=self._load_table).pack(side="left", padx=(0, 6))
        CTkButton(right, text="แก้ไข Markup %", width=120, height=28,
                  fg_color=CLR["blue"], hover_color="#1A3FA8",
                  font=CTkFont(size=12),
                  command=self._edit_selected).pack(side="left")

    # ------------------------------------------------------------------
    def _build_table(self):
        tf = CTkFrame(self, fg_color=CLR["white"], corner_radius=0,
                      border_width=1, border_color=CLR["border"])
        tf.grid(row=1, column=0, sticky="nsew")
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(1, weight=1)  # row 0 = custom header, row 1 = tree

        # ── Custom colored header ──────────────────────────────────────
        hdr = tk.Frame(tf, bg=CLR["navy"])
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")

        def _hcell(text, width_px, bg, fg, stretch=False, anchor="w"):
            f = tk.Frame(hdr, bg=bg, bd=0)
            if stretch:
                f.pack(side="left", fill="both", expand=True)
            else:
                f.config(width=width_px, height=32)
                f.pack(side="left", fill="y")
                f.pack_propagate(False)
            tk.Label(f, text=text, bg=bg, fg=fg,
                     font=("Tahoma", 11, "bold"),
                     anchor=anchor, padx=8).pack(fill="both", expand=True)
            tk.Frame(hdr, bg="#1E3A5F", width=1).pack(side="left", fill="y")

        _hcell("หมวดหลัก",    110,  CLR["navy"],  "#FFFFFF")
        _hcell("หมวดรอง",     145,  CLR["navy"],  "#FFFFFF")
        _hcell("ประเภทสินค้า", 300, CLR["navy"],  "#FFFFFF", stretch=True)
        for tier in TIER_ORDER:
            _hcell(tier, 68, TIER_BG[tier], TIER_FG[tier], anchor="center")
        # spacer matching the vertical scrollbar width
        tk.Frame(hdr, bg=CLR["navy"], width=17).pack(side="left", fill="y")

        # ── Treeview — no built-in header (replaced by custom above) ──
        cols = ("main", "sub", "product", "guest", "t1", "t2", "t3", "t4", "t5")
        self._tree = ttk.Treeview(tf, columns=cols, show="",
                                  style="MT.Treeview", selectmode="browse")

        for cid, w, anch in [
            ("main",   110, "w"),
            ("sub",    145, "w"),
            ("product", 300, "w"),
            ("guest",   68, "e"),
            ("t1",      68, "e"),
            ("t2",      68, "e"),
            ("t3",      68, "e"),
            ("t4",      68, "e"),
            ("t5",      68, "e"),
        ]:
            self._tree.column(cid, width=w, anchor=anch,
                              stretch=(cid == "product"), minwidth=w)

        # group-level alternating background (changes per sub-category, not per row)
        self._tree.tag_configure("grp_a",   background="#FFFFFF", foreground=CLR["text"],
                                 font=("Tahoma", 11))
        self._tree.tag_configure("grp_b",   background="#EEF3FB", foreground=CLR["text"],
                                 font=("Tahoma", 11))
        self._tree.tag_configure("grp_a_h", background="#FFFFFF", foreground=CLR["text"],
                                 font=("Tahoma", 11, "bold"))
        self._tree.tag_configure("grp_b_h", background="#EEF3FB", foreground=CLR["text"],
                                 font=("Tahoma", 11, "bold"))
        self._tree.tag_configure("edited",  background="#FEF9C3", foreground="#92400E")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")

        self._tree.bind("<Double-1>",        self._on_cell_dbl)
        self._tree.bind("<ButtonRelease-1>", self._on_cell_click)

    # ------------------------------------------------------------------
    def _load_categories(self):
        try: cats = ["ทั้งหมด"] + db_fetch_categories()
        except: cats = ["ทั้งหมด"]
        self._cat_menu.configure(values=cats)

    def _load_table(self):
        self._tree.delete(*self._tree.get_children())
        self._piv_map.clear()
        try:
            self._piv_rows = db_fetch_pivoted(
                self._cat_var.get(), self._search_var.get())
        except Exception as e:
            self._lbl_count.configure(
                text=f"โหลดไม่ได้: {e}", text_color="#EF4444")
            return

        self._lbl_count.configure(
            text=f"{len(self._piv_rows)} รายการ", text_color="#94A3B8")

        prev_main = prev_sub = ""
        grp_idx = -1
        for piv in self._piv_rows:
            main = piv["main_category"]
            sub  = piv["sub_category"]
            is_new_grp = (main, sub) != (prev_main, prev_sub)
            if is_new_grp:
                grp_idx += 1
            base = "grp_a" if grp_idx % 2 == 0 else "grp_b"
            tag  = f"{base}_h" if is_new_grp else base

            def _pct(tier):
                td = piv.get(tier)
                if td and td["pct"] is not None:
                    return f"{td['pct']*100:.1f}%"
                return ""

            iid = str(len(self._piv_map))
            self._tree.insert("", "end", iid=iid, tags=(tag,),
                values=(
                    main if main != prev_main else "",
                    sub  if (sub != prev_sub or main != prev_main) else "",
                    piv["product_type"],
                    _pct("Guest"), _pct("T1"), _pct("T2"),
                    _pct("T3"),    _pct("T4"), _pct("T5"),
                ))
            self._piv_map[iid] = piv
            prev_main, prev_sub = main, sub

    # ------------------------------------------------------------------
    def _on_cell_click(self, event):
        col = self._tree.identify_column(event.x)
        if col in _COL_TO_TIER:
            self._last_col = col

    def _on_cell_dbl(self, event):
        col = self._tree.identify_column(event.x)
        tier = _COL_TO_TIER.get(col)
        if not tier: return
        row = self._tree.identify_row(event.y)
        if row and row in self._piv_map:
            self._edit_tier(row, tier)

    def _edit_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("แก้ไข",
                "เลือกแถวสินค้าก่อน แล้วดับเบิลคลิกที่คอลัมน์ Tier ที่ต้องการแก้ไข",
                parent=self)
            return
        tier = _COL_TO_TIER.get(self._last_col)
        if not tier:
            messagebox.showinfo("แก้ไข",
                "คลิกที่คอลัมน์ Tier ที่ต้องการแก้ไข (Guest, T1–T5) ก่อน",
                parent=self)
            return
        self._edit_tier(sel[0], tier)

    def _edit_tier(self, iid: str, tier: str):
        piv = self._piv_map.get(iid)
        if not piv: return
        td = piv.get(tier)
        if not td:
            messagebox.showinfo("แก้ไข",
                f"ไม่มีข้อมูล {tier} สำหรับสินค้านี้", parent=self)
            return

        cur = td["pct"]
        cur_str = f"{cur*100:.1f}" if cur is not None else "0"
        new_val = simpledialog.askstring(
            "แก้ไข Markup %",
            f"สินค้า: {piv['product_type']}\n"
            f"Tier:   {tier}\n\n"
            f"ปัจจุบัน: {cur_str}%   →   กรอก % ใหม่:",
            initialvalue=cur_str, parent=self)

        if new_val is None: return
        try: val = float(new_val.strip().replace("%", ""))
        except ValueError:
            messagebox.showerror("ข้อผิดพลาด", "กรอกตัวเลขเท่านั้น", parent=self)
            return
        if val > 1: val /= 100

        try: db_update_markup(td["id"], val, self.current_user)
        except Exception as e:
            messagebox.showerror("บันทึกไม่สำเร็จ", str(e), parent=self)
            return

        td["pct"] = val
        vals = list(self._tree.item(iid, "values"))
        col_idx = TIER_ORDER.index(tier)
        vals[3 + col_idx] = f"{val*100:.1f}%"
        self._tree.item(iid, values=vals, tags=("edited",))
        messagebox.showinfo("บันทึกแล้ว",
            f"{piv['product_type']} | {tier} → {val*100:.1f}%", parent=self)


if __name__ == "__main__":
    import customtkinter as ctk
    ctk.set_appearance_mode("light")
    root = ctk.CTk()
    root.title("Markup Tiers")
    root.geometry("1280x720")
    MarkupTiersScreen(root).pack(fill="both", expand=True)
    root.mainloop()
