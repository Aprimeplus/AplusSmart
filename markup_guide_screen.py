# markup_guide_screen.py
# แท็บ "Markup Guide" — จัดการ/ดูตาราง Markup Guide (T1-T5) ต่อ SKU
# ใช้ร่วมกันได้ทั้งฝั่ง Purchasing Staff (ดูอย่างเดียว) และ Purchasing Manager (แก้ไขได้)
# ตาราง DB: markup_guide_tiers (product_code, tier, markup_percent, price_min/max, weight_min/max)

import tkinter as tk
from tkinter import ttk, messagebox
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton, CTkEntry,
                            CTkOptionMenu, CTkToplevel, CTkScrollableFrame)
import pandas as pd
import psycopg2.extras
from project_screen import _center_and_style_popup


class MarkupGuideTab(CTkFrame):
    """แท็บแสดงรายการสินค้า + จำนวน Tier ที่ตั้งไว้แล้ว — ดับเบิลคลิกเพื่อดู/แก้ไข Markup Guide ของ SKU นั้น
    editable=True เฉพาะ role ที่อนุญาต (ตาม PM: Purchasing Manager) — role อื่นเปิดได้แค่ดู"""

    EDITABLE_ROLES = {"Purchasing Manager"}

    def __init__(self, master, app_container, user_role=None, user_key=None):
        super().__init__(master, corner_radius=0, fg_color="transparent")
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.user_role = user_role
        self.user_key = user_key
        self.editable = user_role in self.EDITABLE_ROLES
        self._debounce_job = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self._build_toolbar()
        self._build_status_bar()
        self._build_table()

        self._load_categories()

    # ------------------------------------------------------------------ UI
    def _build_toolbar(self):
        bar = CTkFrame(self, fg_color=("gray90", "gray16"), corner_radius=8)
        bar.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        bar.grid_columnconfigure(3, weight=1)

        CTkLabel(bar, text="ค้นหาสินค้า (SKU/ชื่อ):", font=CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, padx=(15, 5), pady=10, sticky="w")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._debounce_search())
        CTkEntry(bar, textvariable=self.search_var, width=220,
                 placeholder_text="พิมพ์อย่างน้อย 2 ตัวอักษร...").grid(row=0, column=1, padx=5, pady=10, sticky="w")

        CTkLabel(bar, text="หมวดหมู่:", font=CTkFont(size=13, weight="bold")).grid(
            row=0, column=2, padx=(15, 5), pady=10, sticky="w")
        self.category_var = tk.StringVar(value="ทั้งหมด")
        self.category_menu = CTkOptionMenu(bar, variable=self.category_var, values=["ทั้งหมด"],
                                            command=lambda _: self._load_products())
        self.category_menu.grid(row=0, column=3, padx=5, pady=10, sticky="w")

        CTkButton(bar, text="⟳ รีเฟรช", width=90, height=32, fg_color="transparent", border_width=1,
                  text_color=("gray10", "gray90"), command=self._load_products).grid(
            row=0, column=4, padx=15, pady=10, sticky="e")

        if not self.editable:
            CTkLabel(bar, text="🔒 ดูได้อย่างเดียว (แก้ไขได้เฉพาะผู้จัดการฝ่ายจัดซื้อ)",
                     text_color="#D97706", font=CTkFont(size=12)).grid(
                row=1, column=0, columnspan=5, padx=15, pady=(0, 8), sticky="w")

    def _build_status_bar(self):
        self.status_label = CTkLabel(self, text="พิมพ์คำค้นหา หรือเลือกหมวดหมู่ เพื่อแสดงรายการสินค้า",
                                      font=CTkFont(size=12), text_color="gray50")
        self.status_label.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="w")

    def _build_table(self):
        table_frame = CTkFrame(self, fg_color="transparent")
        table_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # ใช้ show="tree headings" เพื่อให้ SKU (คอลัมน์ #0) กดขยายดูรายละเอียดแต่ละ Tier
        # เป็น child row ได้เลยในตาราง — ไม่ต้องเปิด popup ก็เห็นว่าราคา/น้ำหนักเท่าไหร่ได้กี่ %
        cols = ("product_name", "category", "tiers_set", "price_range", "weight_range", "markup_pct")
        headers = {
            "product_name": "ชื่อสินค้า", "category": "หมวดหมู่", "tiers_set": "Tier ที่ตั้งแล้ว",
            "price_range": "ช่วงราคา", "weight_range": "ช่วงน้ำหนักรวม (กก.)", "markup_pct": "Markup %",
        }
        widths = {"product_name": 320, "category": 130, "tiers_set": 90,
                  "price_range": 170, "weight_range": 190, "markup_pct": 90}

        tree = ttk.Treeview(table_frame, columns=cols, show="tree headings", selectmode="browse")
        tree.heading("#0", text="SKU / Tier", anchor="w")
        tree.column("#0", width=170, anchor="w")
        for c in cols:
            tree.heading(c, text=headers[c], anchor="center")
            tree.column(c, width=widths[c], anchor="center" if c != "product_name" else "w")
        tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree = tree

        tree.tag_configure("has_guide", background="#DCFCE7")
        tree.tag_configure("no_guide", background="white")
        tree.tag_configure("tier_row", background="#EFF6FF", foreground="#1E3A8A")

        tree.bind("<Double-1>", self._on_row_double_click)

    # ------------------------------------------------------------------ data
    def _load_categories(self):
        try:
            df = pd.read_sql_query(
                "SELECT DISTINCT category FROM products WHERE category IS NOT NULL AND category <> '' ORDER BY category",
                self.pg_engine)
            cats = ["ทั้งหมด"] + df["category"].tolist()
            self.category_menu.configure(values=cats)
        except Exception as e:
            print(f"MarkupGuideTab _load_categories error: {e}")

    def _debounce_search(self):
        if self._debounce_job:
            self.after_cancel(self._debounce_job)
        self._debounce_job = self.after(400, self._load_products)

    def _load_products(self):
        search_text = self.search_var.get().strip()
        category = self.category_var.get()

        if not search_text and category == "ทั้งหมด":
            self.status_label.configure(text="พิมพ์คำค้นหา หรือเลือกหมวดหมู่ เพื่อแสดงรายการสินค้า")
            for row in self.tree.get_children():
                self.tree.delete(row)
            return
        if search_text and len(search_text) < 2:
            self.status_label.configure(text="พิมพ์อย่างน้อย 2 ตัวอักษร")
            return

        try:
            query = """
                SELECT p.product_code, p.product_name, p.category,
                       COALESCE(t.tier_count, 0) AS tiers_set
                FROM products p
                LEFT JOIN (
                    SELECT product_code, COUNT(*) AS tier_count
                    FROM markup_guide_tiers
                    GROUP BY product_code
                ) t ON t.product_code = p.product_code
                WHERE 1=1
            """
            params = []
            if search_text:
                query += " AND (p.product_code ILIKE %s OR p.product_name ILIKE %s)"
                params += [f"%{search_text}%", f"%{search_text}%"]
            if category != "ทั้งหมด":
                query += " AND p.category = %s"
                params.append(category)
            query += " ORDER BY p.product_code LIMIT 500"

            df = pd.read_sql_query(query, self.pg_engine, params=tuple(params))
        except Exception as e:
            messagebox.showerror("Database Error", f"โหลดรายการสินค้าไม่สำเร็จ: {e}", parent=self)
            df = pd.DataFrame()

        # ดึงรายละเอียด Tier ของ SKU ที่กำลังจะแสดงทั้งหมดมาทีเดียว (กันยิง query แยกทีละแถว)
        tiers_by_sku = {}
        codes = df["product_code"].tolist() if not df.empty else []
        if codes:
            try:
                tdf = pd.read_sql_query(
                    "SELECT product_code, tier, markup_percent, price_min, price_max, weight_min, weight_max "
                    "FROM markup_guide_tiers WHERE product_code = ANY(%s) ORDER BY product_code, tier",
                    self.pg_engine, params=(codes,))
                for code, group in tdf.groupby("product_code"):
                    tiers_by_sku[code] = group.to_dict("records")
            except Exception as e:
                print(f"MarkupGuideTab load tier details error: {e}")

        for row in self.tree.get_children():
            self.tree.delete(row)
        for _, r in df.iterrows():
            code = r["product_code"]
            tiers_set = int(r["tiers_set"])
            tag = "has_guide" if tiers_set > 0 else "no_guide"
            parent_id = self.tree.insert("", "end", iid=code, text=code, values=(
                r["product_name"] or "-", r["category"] or "-", f"{tiers_set}/5", "", "", "",
            ), tags=(tag,))
            for t in tiers_by_sku.get(code, []):
                price_txt = self._format_range(t["price_min"], t["price_max"])
                weight_txt = self._format_range(t["weight_min"], t["weight_max"])
                self.tree.insert(parent_id, "end", text=f"    T{int(t['tier'])}", values=(
                    "", "", "", price_txt, weight_txt, f"{float(t['markup_percent']):,.2f}%",
                ), tags=("tier_row",))

        note = " (แสดงสูงสุด 500 รายการแรก ค้นหาให้แคบลงถ้าไม่เจอ)" if len(df) == 500 else ""
        self.status_label.configure(text=f"พบ {len(df)} รายการ{note} — กดลูกศรข้าง SKU เพื่อดูรายละเอียด Tier, ดับเบิลคลิก SKU เพื่อแก้ไข")

    @staticmethod
    def _format_range(min_val, max_val):
        """แปลงช่วง min-max เป็นข้อความอ่านง่าย — ไม่จำกัดฝั่งไหนก็โชว์แค่ '≥'/'≤' ฝั่งที่มีค่า"""
        has_min = pd.notna(min_val)
        has_max = pd.notna(max_val)
        if not has_min and not has_max:
            return "ไม่จำกัด"
        if has_min and has_max:
            return f"{min_val:,.0f} - {max_val:,.0f}"
        if has_min:
            return f"≥ {min_val:,.0f}"
        return f"≤ {max_val:,.0f}"

    def _on_row_double_click(self, event):
        item = self.tree.focus()
        if item and self.tree.parent(item):
            item = self.tree.parent(item)   # คลิกที่แถว Tier (child) ให้เปิดแก้ไขของ SKU แม่แทน
        if not item:
            return
        product_code = item  # iid ของแถวแม่ = product_code (ตั้งไว้ตอน insert)
        vals = self.tree.item(item, "values")
        product_name = vals[0]
        MarkupGuideEditDialog(self, self.app_container, product_code, product_name,
                               editable=self.editable, user_key=self.user_key,
                               on_saved=self._load_products)


class MarkupGuideEditDialog(CTkToplevel):
    """ป็อปอัพแก้ไข/ดู Markup Guide 5 Tier ของ 1 SKU"""

    def __init__(self, master, app_container, product_code, product_name, editable=False,
                 user_key=None, on_saved=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.product_code = product_code
        self.editable = editable
        self.user_key = user_key
        self.on_saved = on_saved

        self.title(f"Markup Guide — {product_code}")
        _center_and_style_popup(self, master, 760, 420)
        self.transient(master)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        CTkLabel(self, text=f"{product_code} — {product_name}", font=CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, padx=15, pady=(15, 0), sticky="w")
        CTkLabel(self, text="เงื่อนไข: ต้องเข้าเกณฑ์ราคา และ น้ำหนักรวม พร้อมกัน ถึงจะเข้า Tier นั้น (เว้นว่าง = ไม่จำกัด)",
                 font=CTkFont(size=12), text_color="gray50").grid(row=1, column=0, padx=15, pady=(2, 10), sticky="w")

        table_frame = CTkFrame(self, fg_color="transparent")
        table_frame.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="nsew")

        headers = ["Tier", "Markup %", "ราคาต่ำสุด", "ราคาสูงสุด", "น้ำหนักรวมต่ำสุด (กก.)", "น้ำหนักรวมสูงสุด (กก.)"]
        for c, h in enumerate(headers):
            CTkLabel(table_frame, text=h, font=CTkFont(size=12, weight="bold")).grid(
                row=0, column=c, padx=6, pady=(0, 8), sticky="w")

        self.entries = {}  # tier -> dict(field->entry)
        existing = self._load_existing()

        for i, tier in enumerate(range(1, 6), start=1):
            row_vals = existing.get(tier, {})
            CTkLabel(table_frame, text=f"T{tier}", font=CTkFont(weight="bold")).grid(
                row=i, column=0, padx=6, pady=4, sticky="w")
            fields = {}
            for c, field in enumerate(["markup_percent", "price_min", "price_max", "weight_min", "weight_max"], start=1):
                e = CTkEntry(table_frame, width=110)
                val = row_vals.get(field)
                if val is not None:
                    e.insert(0, f"{val:g}")
                if not self.editable:
                    e.configure(state="disabled")
                e.grid(row=i, column=c, padx=6, pady=4)
                fields[field] = e
            self.entries[tier] = fields

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="e")
        if self.editable:
            CTkButton(btn_frame, text="บันทึก", command=self._save,
                      fg_color="#16A34A", hover_color="#15803D").pack(side="right", padx=(8, 0))
        CTkButton(btn_frame, text="ปิด", fg_color="transparent", border_width=1,
                  text_color=("gray10", "gray90"), command=self.destroy).pack(side="right")

    def _load_existing(self):
        try:
            df = pd.read_sql_query(
                "SELECT tier, markup_percent, price_min, price_max, weight_min, weight_max "
                "FROM markup_guide_tiers WHERE product_code = %s",
                self.pg_engine, params=(self.product_code,))
        except Exception as e:
            print(f"MarkupGuideEditDialog _load_existing error: {e}")
            return {}
        out = {}
        for _, r in df.iterrows():
            out[int(r["tier"])] = {
                "markup_percent": float(r["markup_percent"]) if pd.notna(r["markup_percent"]) else None,
                "price_min": float(r["price_min"]) if pd.notna(r["price_min"]) else None,
                "price_max": float(r["price_max"]) if pd.notna(r["price_max"]) else None,
                "weight_min": float(r["weight_min"]) if pd.notna(r["weight_min"]) else None,
                "weight_max": float(r["weight_max"]) if pd.notna(r["weight_max"]) else None,
            }
        return out

    def _parse_float(self, text):
        text = (text or "").strip().replace(",", "")
        if not text:
            return None
        return float(text)

    def _save(self):
        try:
            rows_to_upsert = []
            tiers_to_clear = []
            for tier, fields in self.entries.items():
                try:
                    markup = self._parse_float(fields["markup_percent"].get())
                    price_min = self._parse_float(fields["price_min"].get())
                    price_max = self._parse_float(fields["price_max"].get())
                    weight_min = self._parse_float(fields["weight_min"].get())
                    weight_max = self._parse_float(fields["weight_max"].get())
                except ValueError:
                    messagebox.showerror("ข้อมูลไม่ถูกต้อง", f"Tier {tier}: กรุณากรอกตัวเลขเท่านั้น", parent=self)
                    return

                if markup is None:
                    # ไม่กรอก markup % = ไม่ตั้ง Tier นี้ (เว้นว่างได้ ไม่บังคับครบ 5)
                    tiers_to_clear.append(tier)
                    continue
                if price_min is not None and price_max is not None and price_min > price_max:
                    messagebox.showerror("ข้อมูลไม่ถูกต้อง", f"Tier {tier}: ราคาต่ำสุดมากกว่าราคาสูงสุด", parent=self)
                    return
                if weight_min is not None and weight_max is not None and weight_min > weight_max:
                    messagebox.showerror("ข้อมูลไม่ถูกต้อง", f"Tier {tier}: น้ำหนักต่ำสุดมากกว่าน้ำหนักสูงสุด", parent=self)
                    return
                rows_to_upsert.append((tier, markup, price_min, price_max, weight_min, weight_max))

            conn = self.app_container.get_connection()
            try:
                with conn.cursor() as cur:
                    if tiers_to_clear:
                        cur.execute(
                            "DELETE FROM markup_guide_tiers WHERE product_code = %s AND tier = ANY(%s)",
                            (self.product_code, tiers_to_clear))
                    for tier, markup, price_min, price_max, weight_min, weight_max in rows_to_upsert:
                        cur.execute("""
                            INSERT INTO markup_guide_tiers
                                (product_code, tier, markup_percent, price_min, price_max, weight_min, weight_max, updated_by, updated_at)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
                            ON CONFLICT (product_code, tier) DO UPDATE SET
                                markup_percent = EXCLUDED.markup_percent,
                                price_min = EXCLUDED.price_min,
                                price_max = EXCLUDED.price_max,
                                weight_min = EXCLUDED.weight_min,
                                weight_max = EXCLUDED.weight_max,
                                updated_by = EXCLUDED.updated_by,
                                updated_at = NOW()
                        """, (self.product_code, tier, markup, price_min, price_max, weight_min, weight_max, self.user_key))
                conn.commit()
            finally:
                self.app_container.release_connection(conn)

            messagebox.showinfo("สำเร็จ", f"บันทึก Markup Guide ของ {self.product_code} เรียบร้อยแล้ว", parent=self)
            if self.on_saved:
                self.on_saved()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Database Error", f"บันทึกไม่สำเร็จ: {e}", parent=self)
