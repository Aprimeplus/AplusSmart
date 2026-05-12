# markup_tiers_screen.py
# หน้าจอ Manager สำหรับดู / แก้ไข Markup Tiers

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import psycopg2
import psycopg2.extras

# ---- DB helpers (เหมือน super_supplier_list.py) ----
_DB_CFG = dict(host="192.168.1.60", dbname="aplus_com_test",
               user="app_user", password="cailfornia123")
_app_container = None   # ถูก set จากภายนอก

TIER_ORDER = ["Guest", "T1", "T2", "T3", "T4", "T5"]
TIER_COLORS = {
    "Guest": "#78909C",
    "T1":    "#1565C0",
    "T2":    "#2E7D32",
    "T3":    "#F57F17",
    "T4":    "#6A1B9A",
    "T5":    "#B71C1C",
}


def _get_conn():
    if _app_container:
        return _app_container.get_connection(), True
    return psycopg2.connect(**_DB_CFG), False


def _release_conn(conn, use_pool):
    if use_pool and _app_container:
        _app_container.release_connection(conn)
    else:
        conn.close()


# ==============================================================================
#  DB FUNCTIONS
# ==============================================================================
def db_fetch_categories() -> list[str]:
    conn, pool = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT DISTINCT main_category FROM markup_tiers
                WHERE is_active = TRUE
                ORDER BY main_category
            """)
            return [r[0] for r in cur.fetchall()]
    finally:
        _release_conn(conn, pool)


def db_fetch_rows(main_category: str | None = None, search: str = "") -> list[dict]:
    conn, pool = _get_conn()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            filters = ["is_active = TRUE"]
            params = []
            if main_category and main_category != "ทั้งหมด":
                filters.append("main_category = %s")
                params.append(main_category)
            if search.strip():
                filters.append("(product_type ILIKE %s OR sub_category ILIKE %s)")
                params += [f"%{search}%", f"%{search}%"]
            where = " AND ".join(filters)
            cur.execute(f"""
                SELECT id, main_category, sub_category, product_type,
                       tier_name, tier_order, markup_pct, amount_range, qty_range,
                       cost_per_kg, status_note
                FROM markup_tiers
                WHERE {where}
                ORDER BY main_category, sub_category, product_type, tier_order
            """, params)
            return [dict(r) for r in cur.fetchall()]
    finally:
        _release_conn(conn, pool)


def db_update_markup(row_id: int, markup_pct: float, updated_by: str = "MANAGER"):
    conn, pool = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE markup_tiers
                SET markup_pct = %s, updated_at = NOW(), excel_updated_by = %s
                WHERE id = %s
            """, (markup_pct, updated_by, row_id))
        conn.commit()
    finally:
        _release_conn(conn, pool)


# ==============================================================================
#  MAIN SCREEN
# ==============================================================================
class MarkupTiersScreen(ctk.CTkFrame):
    """
    Tab สำหรับ Manager จัดการ Markup Tiers
    ใช้งาน: embed ใน purchasing_manager_screen หรือรันแบบ standalone
    """

    def __init__(self, master, app_container=None, current_user="MANAGER", **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        global _app_container
        if app_container:
            _app_container = app_container
        self.current_user = current_user
        self._pending_edits: dict[int, float] = {}   # {row_id: new_markup_pct}
        self._rows: list[dict] = []

        self._build_ui()
        self._load_categories()
        self._load_table()

    # ------------------------------------------------------------------
    #  BUILD UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---- toolbar ----
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(toolbar, text="📊 ตารางราคา Markup Tiers",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")

        # save button
        self.btn_save = ctk.CTkButton(
            toolbar, text="💾 บันทึกการแก้ไข", width=160,
            fg_color="#1565C0", hover_color="#0D47A1",
            command=self._save_edits)
        self.btn_save.pack(side="right", padx=(8, 0))

        # reload button
        ctk.CTkButton(
            toolbar, text="🔄 โหลดใหม่", width=100,
            fg_color="#37474F", hover_color="#263238",
            command=self._load_table).pack(side="right", padx=(8, 0))

        # ---- filter bar ----
        fbar = ctk.CTkFrame(self, fg_color="transparent")
        fbar.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(fbar, text="หมวด:").pack(side="left", padx=(0, 4))
        self.cat_var = tk.StringVar(value="ทั้งหมด")
        self.cat_menu = ctk.CTkOptionMenu(
            fbar, variable=self.cat_var, width=180,
            command=lambda _: self._load_table())
        self.cat_menu.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(fbar, text="ค้นหา:").pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.after(300, self._load_table))
        ctk.CTkEntry(fbar, textvariable=self.search_var, width=240,
                     placeholder_text="พิมพ์ชื่อสินค้า...").pack(side="left")

        self.lbl_count = ctk.CTkLabel(fbar, text="", text_color="gray")
        self.lbl_count.pack(side="left", padx=12)

        # ---- legend ----
        legend = ctk.CTkFrame(self, fg_color="transparent")
        legend.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(legend, text="Tier:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        for t, c in TIER_COLORS.items():
            ctk.CTkLabel(legend, text=f"● {t}", text_color=c,
                         font=ctk.CTkFont(size=11)).pack(side="left", padx=4)

        # ---- scrollable table area ----
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        # header row
        self._build_header()

    def _build_header(self):
        hdr = ctk.CTkFrame(self.scroll_frame, fg_color="#1E2A38", corner_radius=6)
        hdr.pack(fill="x", pady=(0, 4))
        cols = [
            ("หมวดหลัก",       120, "w"),
            ("หมวดรอง",        140, "w"),
            ("ประเภทสินค้า",   320, "w"),
            ("Tier",            60,  "c"),
            ("Markup %",       100,  "c"),
            ("ช่วงยอด",        200, "w"),
            ("ช่วงจำนวน",     160, "w"),
        ]
        for txt, w, anchor in cols:
            ctk.CTkLabel(hdr, text=txt, width=w, anchor=anchor,
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="white").pack(side="left", padx=6, pady=6)

    # ------------------------------------------------------------------
    #  DATA LOADING
    # ------------------------------------------------------------------
    def _load_categories(self):
        try:
            cats = ["ทั้งหมด"] + db_fetch_categories()
        except Exception as e:
            cats = ["ทั้งหมด"]
            print(f"[markup_tiers] load categories error: {e}")
        self.cat_menu.configure(values=cats)

    def _load_table(self):
        # clear existing rows
        for w in self.scroll_frame.winfo_children():
            if isinstance(w, ctk.CTkFrame) and w.cget("fg_color") != "#1E2A38":
                w.destroy()

        self._pending_edits.clear()
        self._entry_map: dict[int, ctk.CTkEntry] = {}   # {row_id: entry widget}

        try:
            self._rows = db_fetch_rows(
                main_category=self.cat_var.get(),
                search=self.search_var.get()
            )
        except Exception as e:
            ctk.CTkLabel(self.scroll_frame,
                         text=f"❌ โหลดข้อมูลไม่ได้: {e}",
                         text_color="#EF5350").pack(pady=20)
            return

        self.lbl_count.configure(text=f"แสดง {len(self._rows)} รายการ")

        # group rows by product_type for alternating bg
        last_prod = None
        alt = False

        for r in self._rows:
            if r["product_type"] != last_prod:
                last_prod = r["product_type"]
                alt = not alt

            bg = "#1A2332" if alt else "#16202E"
            self._build_row(r, bg)

    def _build_row(self, r: dict, bg: str):
        row_id    = r["id"]
        tier      = r["tier_name"]
        tier_color = TIER_COLORS.get(tier, "white")

        frame = ctk.CTkFrame(self.scroll_frame, fg_color=bg, corner_radius=4)
        frame.pack(fill="x", pady=1)

        # หมวดหลัก
        ctk.CTkLabel(frame, text=r["main_category"] or "", width=120, anchor="w",
                     font=ctk.CTkFont(size=11), text_color="#90CAF9",
                     wraplength=116).pack(side="left", padx=6, pady=4)

        # หมวดรอง
        ctk.CTkLabel(frame, text=r["sub_category"] or "", width=140, anchor="w",
                     font=ctk.CTkFont(size=11), text_color="#B0BEC5",
                     wraplength=136).pack(side="left", padx=4, pady=4)

        # ประเภทสินค้า
        ctk.CTkLabel(frame, text=r["product_type"] or "", width=320, anchor="w",
                     font=ctk.CTkFont(size=11),
                     wraplength=316).pack(side="left", padx=4, pady=4)

        # Tier badge
        ctk.CTkLabel(frame, text=tier, width=60, anchor="center",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=tier_color).pack(side="left", padx=4, pady=4)

        # Markup % — editable entry
        current_pct = r["markup_pct"]
        display_val = f"{current_pct * 100:.1f}%" if current_pct is not None else "-"

        entry = ctk.CTkEntry(frame, width=90, justify="center",
                              font=ctk.CTkFont(size=11),
                              border_color=tier_color)
        entry.insert(0, display_val)
        entry.pack(side="left", padx=4, pady=4)
        entry.bind("<FocusOut>", lambda e, rid=row_id, ent=entry, orig=current_pct:
                   self._on_markup_changed(rid, ent, orig))
        entry.bind("<Return>",   lambda e, rid=row_id, ent=entry, orig=current_pct:
                   self._on_markup_changed(rid, ent, orig))
        self._entry_map[row_id] = entry

        # ช่วงยอด
        ctk.CTkLabel(frame, text=r["amount_range"] or "", width=200, anchor="w",
                     font=ctk.CTkFont(size=11), text_color="#B0BEC5",
                     wraplength=196).pack(side="left", padx=4, pady=4)

        # ช่วงจำนวน
        ctk.CTkLabel(frame, text=r["qty_range"] or "", width=160, anchor="w",
                     font=ctk.CTkFont(size=11), text_color="#B0BEC5",
                     wraplength=156).pack(side="left", padx=4, pady=4)

    # ------------------------------------------------------------------
    #  EDIT LOGIC
    # ------------------------------------------------------------------
    def _on_markup_changed(self, row_id: int, entry: ctk.CTkEntry, original_pct):
        raw = entry.get().strip().replace("%", "").replace(",", "")
        try:
            val = float(raw)
        except ValueError:
            entry.configure(border_color="#EF5350")
            return

        # แปลง: ถ้า > 1 ถือว่าเป็น % (เช่น กรอก 11 → 0.11)
        if val > 1:
            val = val / 100

        orig = original_pct or 0
        if abs(val - orig) < 0.0001:
            # ไม่มีการเปลี่ยนแปลง
            if row_id in self._pending_edits:
                del self._pending_edits[row_id]
            entry.configure(border_color=("#979DA2", "#565B5E"))
            return

        self._pending_edits[row_id] = val
        entry.configure(border_color="#FFA726")   # สีส้ม = มีการแก้ไข

        # อัปเดต display
        entry.delete(0, tk.END)
        entry.insert(0, f"{val * 100:.1f}%")

    def _save_edits(self):
        if not self._pending_edits:
            messagebox.showinfo("บันทึก", "ไม่มีการแก้ไข")
            return

        if not messagebox.askyesno("ยืนยัน",
                                   f"บันทึกการแก้ไข {len(self._pending_edits)} รายการ?"):
            return

        errors = []
        for row_id, new_pct in self._pending_edits.items():
            try:
                db_update_markup(row_id, new_pct, updated_by=self.current_user)
                # เปลี่ยน border กลับเป็นปกติ
                if row_id in self._entry_map:
                    self._entry_map[row_id].configure(border_color=("#979DA2", "#565B5E"))
            except Exception as e:
                errors.append(f"ID {row_id}: {e}")

        self._pending_edits.clear()

        if errors:
            messagebox.showerror("บันทึกไม่สำเร็จบางรายการ", "\n".join(errors))
        else:
            messagebox.showinfo("✅ บันทึกแล้ว", "บันทึก Markup Tiers เรียบร้อย")


# ==============================================================================
#  STANDALONE TEST
# ==============================================================================
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    root.title("Markup Tiers Manager")
    root.geometry("1200x700")
    MarkupTiersScreen(root).pack(fill="both", expand=True)
    root.mainloop()
