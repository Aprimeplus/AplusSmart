import os
import sys
import tkinter as tk
import customtkinter as ctk
import pandas as pd
from tkinter import ttk, messagebox
from customtkinter import (CTkToplevel, CTkFrame, CTkEntry, CTkLabel, CTkButton, CTkFont,
                           CTkCheckBox, CTkScrollableFrame)
from datetime import datetime
from custom_widgets import AutoCompleteEntry
import business_logic


def _resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def _max_lot_count_for_value(total_project_value):
    """เพดานจำนวน Lot สูงสุดตามมูลค่าโครงการรวม: < 1,000,000 บาท = 4 Lot, >= 1,000,000 บาท = 6 Lot"""
    return 6 if (total_project_value or 0) >= 1_000_000 else 4


def _center_and_style_popup(win, master, w, h):
    """จัด popup ให้อยู่กลางหน้าต่างหลัก + ใส่ icon ของแอป (เหมือน main_app.py)"""
    win.update_idletasks()
    root = master.winfo_toplevel()
    rx = root.winfo_x() + (root.winfo_width() - w) // 2
    ry = root.winfo_y() + (root.winfo_height() - h) // 2
    win.geometry(f"{w}x{h}+{rx}+{ry}")
    try:
        win.after(200, lambda: win.iconbitmap(_resource_path("app_icon.ico")))
    except Exception:
        pass


class _NewProjectDialog(CTkToplevel):
    def __init__(self, master, on_submit, existing=None):
        """existing: ถ้าส่งมา (dict มี id, project_code, project_name, customer_name,
        total_project_value, deposit_pct, deposit_method) จะเข้าโหมด 'แก้ไขโครงการ' — code แก้ไม่ได้ ค่าอื่น prefill ไว้
        on_submit ในโหมดแก้ไขจะถูกเรียกเป็น on_submit(existing['id'], name, customer, value, deposit_pct, deposit_method)
        (ไม่มี code — โหมดสร้างใหม่เรียกแบบเดิม on_submit(code, name, customer, value, deposit_pct, deposit_method))"""
        super().__init__(master)
        self.on_submit = on_submit
        self.existing = existing
        is_edit = existing is not None
        self.title("แก้ไขโครงการ" if is_edit else "สร้างโครงการใหม่")
        _center_and_style_popup(self, master, 420, 690)
        self.grab_set()

        entry_font = CTkFont(size=13)
        pad = dict(padx=16, pady=(10, 0))

        CTkLabel(self, text="รหัสโครงการ (SO แม่ เช่น SO6906TS001-M)", font=entry_font, anchor="w").pack(fill="x", **pad)
        self.code_entry = CTkEntry(self, placeholder_text="SO6906TS001-M")
        self.code_entry.pack(fill="x", padx=16, pady=(2, 0))

        CTkLabel(self, text="ชื่อโครงการ", font=entry_font, anchor="w").pack(fill="x", **pad)
        self.name_entry = CTkEntry(self, placeholder_text="โครงการติดตั้งระบบ ABC")
        self.name_entry.pack(fill="x", padx=16, pady=(2, 0))

        CTkLabel(self, text="ลูกค้า", font=entry_font, anchor="w").pack(fill="x", **pad)
        self.customer_entry = AutoCompleteEntry(
            self, completion_list=self._load_customer_list(master), display_key="display",
            command=self._on_customer_selected, placeholder_text="ค้นหาชื่อลูกค้า...")
        self.customer_entry.pack(fill="x", padx=16, pady=(2, 0))

        CTkLabel(self, text="มูลค่าโครงการรวม (บาท) *", font=entry_font, anchor="w").pack(fill="x", **pad)
        self.value_entry = CTkEntry(self, placeholder_text="0.00")
        self.value_entry.pack(fill="x", padx=16, pady=(2, 0))

        CTkLabel(self, text="จำนวน Lot ทั้งหมด (แผน) *", font=entry_font, anchor="w").pack(fill="x", **pad)
        self.lot_count_entry = CTkEntry(self, placeholder_text="เช่น 4")
        self.lot_count_entry.pack(fill="x", padx=16, pady=(2, 0))
        self.lot_count_hint_label = CTkLabel(
            self, text="", font=CTkFont(size=12), text_color="#475569", anchor="w")
        self.lot_count_hint_label.pack(fill="x", padx=16, pady=(2, 0))

        CTkLabel(self, text="ยอดมัดจำ", font=entry_font, anchor="w").pack(fill="x", **pad)
        deposit_row = CTkFrame(self, fg_color="transparent")
        deposit_row.pack(fill="x", padx=16, pady=(2, 0))
        deposit_row.grid_columnconfigure(0, weight=1)
        self.deposit_entry = CTkEntry(deposit_row, placeholder_text="เช่น 30")
        self.deposit_entry.grid(row=0, column=0, sticky="ew")
        self.deposit_unit_var = tk.StringVar(value="%")
        self.deposit_unit_seg = ctk.CTkSegmentedButton(
            deposit_row, values=["%", "บาท"], variable=self.deposit_unit_var, width=100,
            command=lambda _v: self._update_deposit_preview())
        self.deposit_unit_seg.grid(row=0, column=1, padx=(8, 0))
        CTkLabel(self, text="ถ้าลูกค้าตกลงมัดจำเป็นยอดเงินตรงๆ (ไม่ใช่ %) เลือก \"บาท\" แล้วกรอกจำนวนเงิน "
                             "ระบบจะคำนวณเทียบเป็น % จากมูลค่าโครงการรวมให้เอง",
                 font=CTkFont(size=12), text_color="#475569", anchor="w",
                 wraplength=390, justify="left").pack(fill="x", padx=16, pady=(2, 0))
        self.deposit_preview_label = CTkLabel(
            self, text="", font=CTkFont(size=13, weight="bold"), text_color="#7C3AED", anchor="w")
        self.deposit_preview_label.pack(fill="x", padx=16, pady=(2, 0))

        CTkLabel(self, text="วิธีหักมัดจำ", font=entry_font, anchor="w").pack(fill="x", **pad)
        self.deposit_method_var = tk.StringVar(value="กระจายทุก Lot")
        self.deposit_method_seg = ctk.CTkSegmentedButton(
            self, values=["กระจายทุก Lot", "หักที่ Lot สุดท้าย"], variable=self.deposit_method_var)
        self.deposit_method_seg.pack(fill="x", padx=16, pady=(2, 0))
        CTkLabel(self, text="\"กระจายทุก Lot\" = หักมัดจำเฉลี่ยเป็นสัดส่วนในทุก Lot เท่าๆ กัน\n"
                             "\"หักที่ Lot สุดท้าย\" = Lot อื่นเก็บเงินเต็ม ส่วนมัดจำทั้งก้อนจะถูก Hold ไว้เป็น"
                             "ตัวประกันราคา แล้วหักออกจากยอดชำระของ Lot สุดท้ายเพียง Lot เดียว",
                 font=CTkFont(size=12), text_color="#475569", anchor="w",
                 wraplength=390, justify="left").pack(fill="x", padx=16, pady=(2, 0))

        self.value_entry.bind("<KeyRelease>", self._update_deposit_preview)
        self.deposit_entry.bind("<KeyRelease>", self._update_deposit_preview)
        self.value_entry.bind("<KeyRelease>", self._update_lot_count_hint, add="+")
        self.lot_count_entry.bind("<KeyRelease>", self._update_lot_count_hint)

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=20)
        CTkButton(btn_frame, text="ยกเลิก", fg_color="#94A3B8", hover_color="#64748B",
                  command=self.destroy).pack(side="left", expand=True, fill="x", padx=(0, 6))
        CTkButton(btn_frame, text="บันทึกการแก้ไข" if is_edit else "สร้างโครงการ",
                  command=self._submit).pack(side="left", expand=True, fill="x", padx=(6, 0))

        if is_edit:
            self.code_entry.insert(0, existing.get("project_code") or "")
            self.code_entry.configure(state="disabled")
            self.name_entry.insert(0, existing.get("project_name") or "")
            if existing.get("customer_name"):
                self.customer_entry.insert(0, existing["customer_name"])
            self.value_entry.insert(0, f"{float(existing.get('total_project_value') or 0):.2f}")
            _plc = existing.get("planned_lot_count")
            if _plc is not None and pd.notna(_plc):
                self.lot_count_entry.insert(0, str(int(_plc)))
            dep = existing.get("deposit_pct")
            if dep is not None:
                self.deposit_entry.insert(0, f"{float(dep):.2f}")
            if existing.get("deposit_method") == "last_lot":
                self.deposit_method_var.set("หักที่ Lot สุดท้าย")
            self._update_deposit_preview()
        self._update_lot_count_hint()

    @staticmethod
    def _load_customer_list(master):
        """ดึงรายชื่อลูกค้าจากตาราง customers มาให้ช่อง 'ลูกค้า' ค้นหาแบบ autocomplete —
        ใส่ 'รหัสลูกค้า' เข้าไปในข้อความที่ใช้ค้นหา (display) ด้วย เพื่อให้พิมพ์รหัสลูกค้าแล้วเจอชื่อได้
        (AutoCompleteEntry ค้นหาจากข้อความใน display เท่านั้น) แต่ตอนเลือกแล้วจะตัดรหัสออก
        เหลือแค่ชื่อลูกค้าล้วนๆ ใส่ในช่องจริง — ดู _on_customer_selected"""
        pg_engine = getattr(master, "pg_engine", None)
        if pg_engine is None:
            return []
        try:
            df = pd.read_sql_query(
                "SELECT customer_code, customer_name FROM customers "
                "WHERE customer_name IS NOT NULL AND customer_name != '' "
                "ORDER BY customer_name", pg_engine)
            result = []
            for _, r in df.iterrows():
                code = str(r["customer_code"] or "").strip()
                name = str(r["customer_name"] or "").strip()
                display = f"{code} - {name}" if code else name
                result.append({"name": name, "code": code, "display": display})
            return result
        except Exception as e:
            print(f"Error loading customer list: {e}")
            return []

    def _on_customer_selected(self, selected_object):
        """พอเลือกจาก dropdown แล้ว ตัดรหัสลูกค้าที่ต่อท้ายไว้เพื่อการค้นหาออก เหลือแค่ชื่อลูกค้าล้วนๆ
        ในช่อง — กันไม่ให้ 'รหัส - ชื่อ' หลุดไปบันทึกเป็นชื่อลูกค้าจริงของโครงการ"""
        name = (selected_object or {}).get("name", "")
        self.customer_entry.delete(0, "end")
        self.customer_entry.insert(0, name)

    def _update_lot_count_hint(self, event=None):
        """แสดงเพดานจำนวน Lot สูงสุดสดๆ ตามมูลค่าโครงการที่กรอก + เตือนถ้ากรอกเกินเพดาน"""
        try:
            value = float(self.value_entry.get().strip().replace(",", "") or 0)
        except ValueError:
            value = 0.0
        max_lot = _max_lot_count_for_value(value)
        lot_count_raw = self.lot_count_entry.get().strip()
        base_text = (f"เพดานสูงสุด {max_lot} Lot "
                     f"({'≥' if value >= 1_000_000 else '<'} 1,000,000 บาท)")
        if lot_count_raw:
            try:
                lot_count = int(float(lot_count_raw))
                if lot_count > max_lot:
                    self.lot_count_hint_label.configure(
                        text=f"⚠ {base_text} — กรอกเกินเพดานอยู่ตอนนี้ ({lot_count} Lot)",
                        text_color="#DC2626")
                    return
                elif lot_count <= 0:
                    self.lot_count_hint_label.configure(
                        text=f"⚠ {base_text} — ต้องมากกว่า 0", text_color="#DC2626")
                    return
            except ValueError:
                pass
        self.lot_count_hint_label.configure(text=base_text, text_color="#475569")

    def _update_deposit_preview(self, event=None):
        """โชว์ค่าเทียบหน่วยตรงข้ามสดๆ — พิมพ์ % ก็เห็นว่าเป็นกี่บาท พิมพ์บาทก็เห็นว่าเป็นกี่ %"""
        try:
            value = float(self.value_entry.get().strip().replace(",", "") or 0)
        except ValueError:
            value = 0.0
        try:
            deposit_raw = float(self.deposit_entry.get().strip().replace(",", "").replace("%", "") or 0)
        except ValueError:
            deposit_raw = 0.0
        if value <= 0 or deposit_raw <= 0:
            self.deposit_preview_label.configure(text="")
            return
        if self.deposit_unit_var.get() == "บาท":
            if deposit_raw > value:
                self.deposit_preview_label.configure(
                    text="⚠ ยอดมัดจำห้ามเกินมูลค่าโครงการรวม — กรอกเกินอยู่ตอนนี้",
                    text_color="#DC2626")
                return
            pct = deposit_raw / value * 100
            self.deposit_preview_label.configure(text=f"= {pct:.2f}% ของมูลค่าโครงการรวม",
                                                  text_color="#7C3AED")
        else:
            if deposit_raw > 100:
                self.deposit_preview_label.configure(
                    text="⚠ % มัดจำห้ามเกิน 100%", text_color="#DC2626")
                return
            amount = value * deposit_raw / 100
            self.deposit_preview_label.configure(text=f"= {amount:,.2f} บาท", text_color="#7C3AED")

    def _submit(self):
        code = self.code_entry.get().strip()
        name = self.name_entry.get().strip()
        customer = self.customer_entry.get().strip()
        value_raw = self.value_entry.get().strip().replace(",", "")
        deposit_raw = self.deposit_entry.get().strip().replace(",", "").replace("%", "")
        if not code or not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกรหัสโครงการและชื่อโครงการ", parent=self)
            return
        if not value_raw:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกมูลค่าโครงการรวม", parent=self)
            return
        try:
            value = float(value_raw)
        except ValueError:
            messagebox.showwarning("ข้อมูลผิดพลาด", "มูลค่าโครงการต้องเป็นตัวเลข", parent=self)
            return
        if value <= 0:
            messagebox.showwarning("ข้อมูลผิดพลาด", "มูลค่าโครงการต้องมากกว่า 0", parent=self)
            return
        lot_count_raw = self.lot_count_entry.get().strip()
        if not lot_count_raw:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกจำนวน Lot ทั้งหมด", parent=self)
            return
        try:
            planned_lot_count = int(float(lot_count_raw))
        except ValueError:
            messagebox.showwarning("ข้อมูลผิดพลาด", "จำนวน Lot ต้องเป็นตัวเลข", parent=self)
            return
        max_lot = _max_lot_count_for_value(value)
        if planned_lot_count <= 0:
            messagebox.showwarning("ข้อมูลผิดพลาด", "จำนวน Lot ต้องมากกว่า 0", parent=self)
            return
        if planned_lot_count > max_lot:
            messagebox.showwarning(
                "ข้อมูลผิดพลาด",
                f"มูลค่าโครงการ {value:,.2f} บาท อนุญาตสูงสุด {max_lot} Lot "
                f"แต่กรอกไว้ {planned_lot_count} Lot", parent=self)
            return
        try:
            deposit_raw_num = float(deposit_raw) if deposit_raw else None
        except ValueError:
            messagebox.showwarning("ข้อมูลผิดพลาด", "ยอดมัดจำต้องเป็นตัวเลข", parent=self)
            return
        deposit_unit = self.deposit_unit_var.get()
        if deposit_raw_num is None:
            deposit_pct = None
        elif deposit_unit == "บาท":
            if deposit_raw_num < 0:
                messagebox.showwarning("ข้อมูลผิดพลาด", "ยอดมัดจำต้องไม่ติดลบ", parent=self)
                return
            if deposit_raw_num > value:
                messagebox.showwarning("ข้อมูลผิดพลาด", "ยอดมัดจำ (บาท) ต้องไม่เกินมูลค่าโครงการรวม", parent=self)
                return
            deposit_pct = (deposit_raw_num / value) * 100
        else:
            deposit_pct = deposit_raw_num
        if deposit_pct is not None and not (0 <= deposit_pct <= 100):
            messagebox.showwarning("ข้อมูลผิดพลาด", "ยอดมัดจำต้องอยู่ระหว่าง 0-100%", parent=self)
            return
        deposit_method = "last_lot" if self.deposit_method_var.get() == "หักที่ Lot สุดท้าย" else "spread"
        self.destroy()
        if self.existing is not None:
            self.on_submit(self.existing["id"], name, customer, value, deposit_pct, deposit_method,
                            planned_lot_count)
        else:
            self.on_submit(code, name, customer, value, deposit_pct, deposit_method, planned_lot_count)


class ProjectScreen(ctk.CTkFrame):
    """หน้าจัดการโครงการ (Multi-Lot Project) — Phase 1: โครงสร้างข้อมูล + ติดตามสถานะ"""

    def __init__(self, master, app_container, user_key=None, user_role=None,
                 sale_key_filter=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_container
        self.pg_engine = app_container.pg_engine
        self.user_key = user_key or "system"
        self.user_role = (user_role or "").strip()
        self._locked_sale_key = sale_key_filter   # ถ้า set = Sale mode (เห็นแค่โครงการของตัวเอง)
        self.current_project_id = None

        self.header_font = CTkFont(size=16, weight="bold")
        self.label_font = CTkFont(size=13)
        self.small_font = CTkFont(size=11)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        self.title_label = CTkLabel(header, text="📁 จัดการโครงการ", font=self.header_font)
        self.title_label.pack(side="left")
        self.new_project_btn = CTkButton(header, text="+ สร้างโครงการใหม่", command=self._open_new_project_dialog)
        self.new_project_btn.pack(side="right")
        self.refresh_btn = CTkButton(header, text="🔄 รีเฟรช", width=90, fg_color="#6B7280", hover_color="#4B5563",
                                      command=self._show_list)
        self.refresh_btn.pack(side="right", padx=(0, 8))

        self.body = CTkFrame(self, fg_color="transparent")
        self.body.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_rowconfigure(0, weight=1)

        self._show_list()

    # ---------------------------------------------------------------- list

    _STATUS_BADGE = {
        "Open":   ("เปิดอยู่",   "#DCFCE7", "#166534"),
        "Closed": ("ปิดแล้ว",   "#E2E8F0", "#475569"),
        "Frozen": ("พักไว้",    "#FEF3C7", "#92400E"),
    }

    def _show_list(self):
        self.current_project_id = None
        # โชว์ปุ่ม "รีเฟรช"/"สร้างโครงการใหม่" ที่หัวจอกลับมา (ถูกซ่อนไว้ตอนอยู่หน้ารายละเอียดโครงการ)
        # ต้อง pack ตามลำดับเดิม (new_project ก่อน แล้วค่อย refresh) ไม่งั้นตำแหน่งซ้าย-ขวาจะสลับกัน
        self.new_project_btn.pack(side="right")
        self.refresh_btn.pack(side="right", padx=(0, 8))
        for w in self.body.winfo_children():
            w.destroy()
        # ล้าง row weight ที่ _show_detail() เคยตั้งไว้ (row 2, 3 เป็นต้น) ทิ้งก่อนเสมอ — ไม่งั้น
        # grid ของ self.body จะยังจองพื้นที่ยืดหยุ่นให้แถวว่างๆ เหล่านั้นค้างอยู่ ทำให้การ์ดโครงการ
        # ถูกบีบเหลือแค่เศษเสี้ยวของพื้นที่จริงหลังจากเคยเข้าไปดูหน้ารายละเอียดโครงการมาก่อน
        for i in range(1, 6):
            self.body.grid_rowconfigure(i, weight=0)
        self.body.grid_rowconfigure(0, weight=1)

        try:
            conn = self.app.get_connection()
            try:
                query = """
                    SELECT p.id, p.project_code, p.project_name, p.customer_name,
                           p.total_project_value, p.status,
                           COALESCE(SUM(l.lot_value) FILTER (WHERE l.status != 'Cancelled'), 0) AS lot_sum,
                           COUNT(l.id) FILTER (WHERE l.status != 'Cancelled') AS lot_count,
                           COUNT(l.id) FILTER (WHERE l.kpi_qualified_flag) AS lot_done
                    FROM projects p
                    LEFT JOIN project_lots l ON l.project_id = p.id
                """
                params = None
                if self._locked_sale_key:
                    # Sale เห็นแค่โครงการที่ตัวเองสร้าง หรือมี Lot ที่ผูกกับ SO ของตัวเอง
                    query += """
                    WHERE p.created_by = %(sk)s
                       OR EXISTS (
                            SELECT 1 FROM project_lots pl
                            JOIN commissions c ON c.so_number = pl.so_number AND c.is_active = 1
                            WHERE pl.project_id = p.id AND c.sale_key = %(sk)s
                       )
                    """
                    params = {"sk": self._locked_sale_key}
                query += """
                    GROUP BY p.id
                    ORDER BY p.created_at DESC
                """
                df = pd.read_sql_query(query, conn, params=params)
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", f"โหลดรายการโครงการไม่สำเร็จ: {e}", parent=self)
            return

        scroll = CTkScrollableFrame(self.body, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        if df.empty:
            CTkLabel(scroll, text="ยังไม่มีโครงการ — กด \"+ สร้างโครงการใหม่\" มุมขวาบนเพื่อเริ่มต้น",
                     font=self.label_font, text_color="#94A3B8").grid(row=0, column=0, pady=40)
            return

        for i, (_, r) in enumerate(df.iterrows()):
            self._build_project_card(scroll, i, r)

    def _build_project_card(self, parent, row_idx, r):
        project_id = int(r["id"])
        total_value = float(r["total_project_value"] or 0)
        lot_sum = float(r["lot_sum"] or 0)
        lot_count = int(r["lot_count"] or 0)
        lot_done = int(r["lot_done"] or 0)
        has_target = total_value > 0
        pct = min(100, (lot_sum / total_value * 100)) if has_target else (100 if lot_count else 0)

        card = CTkFrame(parent, fg_color="#FFFFFF", corner_radius=12,
                         border_width=1, border_color="#E2E8F0")
        card.grid(row=row_idx, column=0, sticky="ew", pady=(0, 10), padx=2)
        card.grid_columnconfigure(0, weight=1)

        top = CTkFrame(card, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 0))
        top.grid_columnconfigure(0, weight=1)
        CTkLabel(top, text=r["project_code"], font=CTkFont(size=15, weight="bold"),
                 anchor="w").grid(row=0, column=0, sticky="w")

        badge_text, badge_bg, badge_fg = self._STATUS_BADGE.get(
            r["status"], (r["status"], "#E2E8F0", "#475569"))
        badge = CTkFrame(top, fg_color=badge_bg, corner_radius=999)
        badge.grid(row=0, column=1, sticky="e")
        CTkLabel(badge, text=badge_text, font=CTkFont(size=11, weight="bold"),
                 text_color=badge_fg).pack(padx=10, pady=3)

        subtitle = f"{r['project_name']}"
        if r["customer_name"]:
            subtitle += f"  ·  ลูกค้า {r['customer_name']}"
        subtitle += f"  ·  {lot_count} Lot"
        CTkLabel(card, text=subtitle, font=self.small_font, text_color="#64748B",
                 anchor="w").grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 10))

        bottom = CTkFrame(card, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 8))
        bottom.grid_columnconfigure(1, weight=1)

        value_text = f"{total_value:,.2f} บาท" if has_target else f"รวม Lot {lot_sum:,.2f} บาท"
        CTkLabel(bottom, text=value_text, font=CTkFont(size=14, weight="bold"),
                 text_color="#1D4ED8").grid(row=0, column=0, sticky="w")
        CTkLabel(bottom, text=f"ครบเงื่อนไข {lot_done}/{lot_count} Lot",
                 font=self.small_font, text_color="#0F766E" if lot_done == lot_count and lot_count else "#94A3B8"
                 ).grid(row=0, column=2, sticky="e")

        bar_bg = CTkFrame(card, fg_color="#F1F5F9", height=6, corner_radius=3)
        bar_bg.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 14))
        bar_bg.grid_propagate(False)
        if pct > 0:
            bar_fill_color = "#B45309" if (has_target and lot_sum > total_value) else "#1D4ED8"
            bar_fill = CTkFrame(bar_bg, fg_color=bar_fill_color, height=6, corner_radius=3)
            bar_fill.place(relx=0, rely=0, relwidth=pct / 100, relheight=1)

        self._bind_click_recursive(card, lambda e=None, pid=project_id: self._show_detail(pid))

    def _bind_click_recursive(self, widget, callback):
        if isinstance(widget, CTkFrame):
            widget.configure(cursor="hand2")
        widget.bind("<Button-1>", callback)
        for child in widget.winfo_children():
            self._bind_click_recursive(child, callback)

    def _open_new_project_dialog(self):
        _NewProjectDialog(self, self._create_project)

    def _create_project(self, code, name, customer, value, deposit_pct=None, deposit_method="spread",
                         planned_lot_count=None):
        try:
            conn = self.app.get_connection()
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        INSERT INTO projects (project_code, project_name, customer_name, total_project_value,
                                               deposit_pct, deposit_method, planned_lot_count, created_by)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    """, (code, name, customer, value, deposit_pct, deposit_method, planned_lot_count,
                          self.user_key))
                conn.commit()
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", f"สร้างโครงการไม่สำเร็จ: {e}", parent=self)
            return
        self._show_list()

    def _open_edit_project_dialog(self, project_id, proj):
        """proj: pandas Series แถวโครงการจาก _show_detail (มี id, project_code, project_name,
        customer_name, total_project_value, deposit_pct อยู่แล้ว) — ส่งเข้า dialog เป็น dict"""
        existing = {
            "id": project_id,
            "project_code": proj["project_code"],
            "project_name": proj["project_name"],
            "customer_name": proj["customer_name"],
            "total_project_value": proj["total_project_value"],
            "deposit_pct": proj["deposit_pct"] if "deposit_pct" in proj.index else None,
            "deposit_method": proj["deposit_method"] if "deposit_method" in proj.index else "spread",
            "planned_lot_count": proj["planned_lot_count"] if "planned_lot_count" in proj.index else None,
        }
        _NewProjectDialog(self, self._update_project, existing=existing)

    def _update_project(self, project_id, name, customer, value, deposit_pct=None, deposit_method="spread",
                         planned_lot_count=None):
        try:
            conn = self.app.get_connection()
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        UPDATE projects
                        SET project_name = %s, customer_name = %s, total_project_value = %s, deposit_pct = %s,
                            deposit_method = %s, planned_lot_count = %s
                        WHERE id = %s
                    """, (name, customer, value, deposit_pct, deposit_method, planned_lot_count, project_id))
                conn.commit()
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", f"บันทึกการแก้ไขไม่สำเร็จ: {e}", parent=self)
            return
        self._show_detail(project_id)

    def _toggle_project_deposit_received(self, project_id, received):
        """flag ระดับโครงการสำหรับวิธีมัดจำ 'หักที่ Lot สุดท้าย' — มัดจำเก็บเป็นก้อนเดียวตอนทำสัญญา"""
        try:
            conn = self.app.get_connection()
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        UPDATE projects
                        SET deposit_received_flag = %s,
                            deposit_received_date = CASE WHEN %s THEN COALESCE(deposit_received_date, CURRENT_DATE)
                                                          ELSE NULL END
                        WHERE id = %s
                    """, (received, received, project_id))
                conn.commit()
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", f"อัปเดตสถานะมัดจำไม่สำเร็จ: {e}", parent=self)
            return
        self._show_detail(project_id)

    # -------------------------------------------------------------- detail

    def _sync_payment_collected(self, conn, project_id):
        """เช็ค 'เก็บเงินครบ' อัตโนมัติจาก commissions.difference_amount ของ SO ที่ผูกไว้กับแต่ละ Lot
        (difference_amount = 0 แปลว่ายอดชำระตรงกับยอดเต็มแล้ว) — ไม่ต้องให้คนติ๊กเองอีกต่อไป
        ถ้า SO ที่ผูกไว้ถูกยกเลิก (ทุก record ของ so_number นั้นเป็น Cancelled/is_active=0 หมด)
        ให้ Lot กลายเป็นสถานะ 'Cancelled' ทันที ไม่นับเป็นเก็บเงินครบ/ครบเงื่อนไขอีกต่อไป"""
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE project_lots l
                SET payment_collected_flag = CASE WHEN cs.is_cancelled THEN FALSE
                                                    ELSE (COALESCE(c.difference_amount, 1) = 0) END,
                    payment_collected_date = CASE WHEN NOT cs.is_cancelled
                                                        AND COALESCE(c.difference_amount, 1) = 0
                                                   THEN COALESCE(l.payment_collected_date, CURRENT_DATE)
                                                   ELSE NULL END,
                    kpi_qualified_flag = CASE WHEN cs.is_cancelled THEN FALSE
                                               ELSE (l.delivered_flag AND l.invoice_recorded_flag
                                                     AND COALESCE(c.difference_amount, 1) = 0) END,
                    status = CASE
                        WHEN cs.is_cancelled THEN 'Cancelled'
                        WHEN l.delivered_flag AND l.invoice_recorded_flag
                             AND COALESCE(c.difference_amount, 1) = 0 THEN 'Collected'
                        WHEN l.invoice_recorded_flag THEN 'Invoiced'
                        WHEN l.delivered_flag THEN 'Delivered'
                        ELSE 'Draft'
                    END
                FROM project_lots pl
                LEFT JOIN commissions c ON c.so_number = pl.so_number AND c.is_active = 1
                         AND c.status NOT IN ('Cancelled', 'Cancelled by PU')
                LEFT JOIN LATERAL (
                    SELECT NOT EXISTS (
                        SELECT 1 FROM commissions x
                        WHERE x.so_number = pl.so_number AND x.is_active = 1
                          AND x.status NOT IN ('Cancelled', 'Cancelled by PU')
                    ) AS is_cancelled
                ) cs ON TRUE
                WHERE pl.id = l.id AND pl.project_id = %s AND pl.so_number IS NOT NULL
            """, (project_id,))
        conn.commit()

    def _show_detail(self, project_id):
        self.current_project_id = project_id
        # ซ่อนปุ่ม "รีเฟรช"/"สร้างโครงการใหม่" ที่หัวจอตอนอยู่หน้ารายละเอียด — ไม่งั้นมันไปเบียดซ้อนกับ
        # ปุ่ม "แก้ไขโครงการ" ของหน้านี้ที่อยู่มุมขวาบนเหมือนกัน ดูรก/ทับกัน (แสดงกลับตอน _show_list())
        self.new_project_btn.pack_forget()
        self.refresh_btn.pack_forget()
        for w in self.body.winfo_children():
            w.destroy()

        try:
            conn = self.app.get_connection()
            try:
                proj_df = pd.read_sql_query(
                    "SELECT * FROM projects WHERE id = %s", conn, params=(project_id,))
                self._sync_payment_collected(conn, project_id)
                # grand_total_due = "ยอดที่ต้องชำระ" ของ SO จริงที่ผูกกับ Lot นี้ — คำนวณด้วยสูตรเดียวกับ
                # so_grand_total_var ในฟอร์ม SO ทุกตัว (commission_app.py: _update_final_calculations):
                # รวมเฉพาะรายการที่ติ๊ก VAT (สินค้า/บริการ, ตัด/เจาะ, บริการอื่นๆ, ค่าจัดส่ง, ค่าธรรมเนียมบัตร,
                # ค่าย้าย) คูณ VAT 7% แล้วหัก wht — ดึงสดจาก commissions เสมอ ไม่ได้บันทึกซ้ำไว้ที่ project_lots
                lots_df = pd.read_sql_query("""
                    SELECT pl.*,
                           COALESCE(c.sales_service_amount, 0) AS product_amount,
                           COALESCE(c.shipping_cost, 0)        AS shipping_amount,
                           (
                               (CASE WHEN c.sales_service_vat_option = 'VAT' THEN COALESCE(c.sales_service_amount, 0) ELSE 0 END
                              + CASE WHEN c.cutting_drilling_fee_vat_option = 'VAT' THEN COALESCE(c.cutting_drilling_fee, 0) ELSE 0 END
                              + CASE WHEN c.other_service_fee_vat_option = 'VAT' THEN COALESCE(c.other_service_fee, 0) ELSE 0 END
                              + CASE WHEN c.shipping_vat_option = 'VAT' THEN COALESCE(c.shipping_cost, 0) ELSE 0 END
                              + CASE WHEN c.credit_card_fee_vat_option = 'VAT' THEN COALESCE(c.credit_card_fee, 0) ELSE 0 END
                              + CASE WHEN c.relocation_cost_vat_option = 'VAT' THEN COALESCE(c.relocation_cost, 0) ELSE 0 END
                               ) * 0.07
                           ) AS vat_total,
                           GREATEST(0, (
                               (CASE WHEN c.sales_service_vat_option = 'VAT' THEN COALESCE(c.sales_service_amount, 0) ELSE 0 END
                              + CASE WHEN c.cutting_drilling_fee_vat_option = 'VAT' THEN COALESCE(c.cutting_drilling_fee, 0) ELSE 0 END
                              + CASE WHEN c.other_service_fee_vat_option = 'VAT' THEN COALESCE(c.other_service_fee, 0) ELSE 0 END
                              + CASE WHEN c.shipping_vat_option = 'VAT' THEN COALESCE(c.shipping_cost, 0) ELSE 0 END
                              + CASE WHEN c.credit_card_fee_vat_option = 'VAT' THEN COALESCE(c.credit_card_fee, 0) ELSE 0 END
                              + CASE WHEN c.relocation_cost_vat_option = 'VAT' THEN COALESCE(c.relocation_cost, 0) ELSE 0 END
                               ) * 1.07 - COALESCE(c.wht_3_percent, 0)
                           )) AS grand_total_due
                    FROM project_lots pl
                    LEFT JOIN commissions c ON c.so_number = pl.so_number AND c.is_active = 1
                             AND c.status NOT IN ('Cancelled', 'Cancelled by PU')
                    WHERE pl.project_id = %s ORDER BY pl.lot_number
                """, conn, params=(project_id,))
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", f"โหลดข้อมูลโครงการไม่สำเร็จ: {e}", parent=self)
            return

        if proj_df.empty:
            messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบโครงการนี้แล้ว", parent=self)
            self._show_list()
            return
        proj = proj_df.iloc[0]

        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_rowconfigure(2, weight=1)

        top_row = CTkFrame(self.body, fg_color="transparent")
        top_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top_row.grid_columnconfigure(0, weight=1)

        back_btn = CTkButton(top_row, text="← โครงการทั้งหมด", width=140, fg_color="transparent",
                              text_color="#1D4ED8", hover_color="#EFF4FF", anchor="w",
                              command=self._show_list)
        back_btn.grid(row=0, column=0, sticky="w")

        actions_frame = CTkFrame(top_row, fg_color="transparent")
        actions_frame.grid(row=0, column=1, sticky="e")

        if proj["status"] == "Open":
            edit_btn = CTkButton(actions_frame, text="✏️ แก้ไขโครงการ", width=140,
                                  fg_color="#2563EB", hover_color="#1D4ED8",
                                  command=lambda: self._open_edit_project_dialog(project_id, proj))
            edit_btn.pack(side="left", padx=(0, 8))

        can_close = (self.user_role.lower() in ("director", "hr")) and proj["status"] == "Open"
        if can_close:
            active_lots_check = lots_df[lots_df["status"] != "Cancelled"] if not lots_df.empty else lots_df
            all_collected = (not active_lots_check.empty) and (active_lots_check["status"] == "Collected").all()
            close_btn = CTkButton(actions_frame, text="🔒 ปิดโปรเจกต์ (GP True-Up)", width=200,
                                   fg_color="#B45309" if all_collected else "#94A3B8",
                                   hover_color="#92400E" if all_collected else "#94A3B8",
                                   state="normal" if all_collected else "disabled",
                                   command=lambda: self._open_close_project_dialog(project_id))
            close_btn.pack(side="left")
            if not all_collected:
                CTkLabel(top_row, text="(ต้องให้ทุก Lot \"เก็บเงินครบ\" ก่อนถึงจะปิดโปรเจกต์ได้)",
                          font=self.small_font, text_color="#94A3B8").grid(row=1, column=1, sticky="e", pady=(2, 0))
        elif proj["status"] == "Closed":
            CTkLabel(actions_frame, text=f"🔒 ปิดโปรเจกต์แล้ว — GP จริง {float(proj['final_gp_pct'] or 0):.2f}%",
                      font=CTkFont(size=12, weight="bold"), text_color="#475569").pack(side="left")

        # ฟอนต์ใหญ่ขึ้นเฉพาะแผงข้อมูลนี้ — พื้นที่การ์ดกว้าง แต่ตัวหนังสือเดิมเล็กเกินไปเมื่อเทียบกับพื้นที่ว่าง
        info_title_font = CTkFont(size=20, weight="bold")
        info_sub_font = CTkFont(size=14)
        info_value_font = CTkFont(size=16, weight="bold")

        info = CTkFrame(self.body, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
        info.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        info.grid_columnconfigure(0, weight=1)

        header_row = CTkFrame(info, fg_color="transparent")
        header_row.grid(row=0, column=0, columnspan=4, sticky="ew", padx=18, pady=(16, 4))
        CTkLabel(header_row, text=f"{proj['project_code']} · {proj['project_name']}",
                 font=info_title_font).pack(side="left")
        self._project_status_badge_label = CTkLabel(header_row, text="", font=info_sub_font)
        self._project_status_badge_label.pack(side="right")
        CTkLabel(info, text=f"ลูกค้า: {proj['customer_name'] or '-'}", font=info_sub_font,
                 text_color="#64748B").grid(row=1, column=0, sticky="w", padx=18, pady=(0, 14))

        total_value = float(proj["total_project_value"] or 0)
        active_lots_df = lots_df[lots_df["status"] != "Cancelled"] if not lots_df.empty else lots_df
        lot_sum = float(active_lots_df["lot_value"].sum()) if not active_lots_df.empty else 0.0
        has_target = total_value > 0
        exceeded = has_target and lot_sum > total_value
        pct = min(100, (lot_sum / total_value * 100)) if has_target else 0

        # สถานะโครงการโดยรวม (badge มุมขวาบน) — นับ Lot ที่ "ส่งครบ" แล้วเทียบกับแผน
        planned_for_badge = proj["planned_lot_count"] if "planned_lot_count" in proj.index else None
        planned_for_badge = (int(planned_for_badge) if planned_for_badge is not None and pd.notna(planned_for_badge)
                              else (len(active_lots_df) if not active_lots_df.empty else 0))
        delivered_count = int(active_lots_df["delivered_flag"].sum()) if not active_lots_df.empty else 0
        if proj["status"] != "Open":
            badge_text, badge_color = f"🔒 {proj['status']}", "#475569"
        elif planned_for_badge and delivered_count >= planned_for_badge:
            badge_text, badge_color = "✅ เสร็จสมบูรณ์", "#16A34A"
        elif delivered_count > 0:
            badge_text, badge_color = f"🟢 กำลังดำเนินงาน (ส่งของแล้ว {delivered_count}/{planned_for_badge} Lot)", "#16A34A"
        else:
            badge_text, badge_color = f"⏳ ยังไม่เริ่มส่งของ (0/{planned_for_badge} Lot)", "#94A3B8"
        self._project_status_badge_label.configure(text=badge_text, text_color=badge_color)

        # ยอดมัดจำ — คิดแยกเป็นก้อนต่อ Lot จากยอดที่ต้องชำระจริงของ SO นั้น (grand_total_due ที่คำนวณ
        # มาแล้วเหมือนช่อง "ยอดที่ต้องชำระ" ในฟอร์ม SO — รวม VAT ของทุกรายการที่ติ๊ก VAT หัก wht แล้ว)
        # แล้วรวมยอดขึ้นมา ไม่ได้หารเฉลี่ยจากมูลค่าโครงการรวมลงไป เพื่อไม่ต้องแบ่งใหม่ทุกครั้งที่มี Lot เพิ่ม
        deposit_pct_val = proj["deposit_pct"] if "deposit_pct" in proj.index else None
        deposit_method = (proj["deposit_method"] if "deposit_method" in proj.index else "spread") or "spread"
        planned_lot_count = proj["planned_lot_count"] if "planned_lot_count" in proj.index else None
        planned_lot_count = int(planned_lot_count) if planned_lot_count is not None and pd.notna(planned_lot_count) else None
        grand_total_sum = float(active_lots_df["grand_total_due"].sum()) if not active_lots_df.empty else 0.0
        # deposit_method = 'last_lot': ไล่หักมัดจำย้อนจาก Lot สุดท้ายขึ้นไปหน้าเรื่อยๆ จนกว่ามัดจำจะหมด
        # (ไม่ใช่หักได้แค่ Lot เดียว) — กันกรณี Lot สุดท้ายมูลค่าน้อยกว่ามัดจำทั้งก้อน เช่น มัดจำ 100,000
        # แต่ Lot สุดท้ายมูลค่าแค่ 50,000 ต้องหักส่วนที่เหลือ (50,000) ไหลย้อนไปหัก Lot ก่อนหน้าต่อ
        #
        # ถ้ารู้ planned_lot_count (จำนวน Lot ที่วางแผนไว้ตอนสร้างโครงการ) ให้ยึดเลขนั้นเป็น "Lot สุดท้ายจริง"
        # แทนที่จะเดาจาก Lot ที่มีอยู่ตอนนี้ — กันกรณีสร้าง Lot ไปแล้วแค่บางส่วน (เช่น วางแผน 4 Lot แต่สร้าง
        # ไปแค่ 3 Lot) ซึ่งถ้าเดาจาก Lot ปัจจุบันจะหักมัดจำผิด Lot ไปก่อน (Lot 3 ที่จริงยังไม่ใช่ Lot สุดท้าย)
        # ยอดมัดจำ — คำนวณเฉยๆ ตรงนี้ก่อน (ยังไม่ render) เพราะการ์ดสรุป 4 ใบต้องใช้ total_deposit/
        # proj_dep_received ด้วย — ส่วนกล่องมัดจำ (checkbox/ปุ่ม) ไป render อยู่ใต้การ์ดสรุปแทน (ตาม
        # ตำแหน่งที่ PM ให้ไว้ในแบบ: การ์ดสรุป 4 ใบอยู่บน กล่องมัดจำอยู่ล่าง)
        if deposit_pct_val is not None and pd.notna(deposit_pct_val):
            deposit_ratio = float(deposit_pct_val) / 100
            is_estimate = active_lots_df.empty or grand_total_sum <= 0
            # ยังไม่มี Lot จริงเลย — โชว์ยอดมัดจำประมาณการจาก "มูลค่าโครงการรวม" ที่กรอกไว้ตอนสร้างโครงการ
            # ไปก่อน แทนที่จะโชว์ 0.00 บาท เพราะลูกค้าโอนมัดจำก้อนแรกตั้งแต่ก่อน Lot 1 จะถูกสร้างด้วยซ้ำ
            total_deposit = (total_value * deposit_ratio) if is_estimate else (grand_total_sum * deposit_ratio)
            has_any_manual_dep = False
            if deposit_method == "spread" and not active_lots_df.empty:
                # ยอดมัดจำรวมจริง = ยอดกรอกเอง (ถ้ามี) แทนยอดคำนวณ % สำหรับ Lot นั้นๆ
                actual_total = 0.0
                for _, lr in active_lots_df.iterrows():
                    man_dep = lr.get('manual_deposit_amount')
                    if man_dep is not None and pd.notna(man_dep):
                        actual_total += float(man_dep)
                        has_any_manual_dep = True
                    else:
                        actual_total += float(lr.get('grand_total_due', 0) or 0) * deposit_ratio
                total_deposit = actual_total
                is_estimate = False
            method_text = ("หักที่ Lot สุดท้าย — Hold ตลอดโปรเจกต์" if deposit_method == "last_lot"
                            else "หักกระจายทุก Lot")
            manual_note = " (*มี Lot ที่กรอกยอดมัดจำเองแทนค่าคำนวณ)" if has_any_manual_dep else ""
            estimate_note = " (ประมาณการจากมูลค่าโครงการรวม — ยังไม่มี Lot จริง)" if is_estimate else ""
            proj_dep_received = bool(proj.get("deposit_received_flag") or False)
        else:
            deposit_ratio = None
            total_deposit = 0.0
            proj_dep_received = False

        # ── การ์ดสรุปยอด 4 ใบ (มูลค่ารวม / ชำระแล้ว / คงเหลือค้างชำระ / คิดคอมมิชชั่นสะสม) ──────
        next_row = 3
        paid_lots_df = active_lots_df[active_lots_df["payment_collected_flag"] == True] \
            if not active_lots_df.empty else active_lots_df
        paid_lots_sum = float(paid_lots_df["grand_total_due"].sum()) if not paid_lots_df.empty else 0.0
        commission_sum = float(paid_lots_df["product_amount"].sum()) if not paid_lots_df.empty else 0.0
        total_paid = paid_lots_sum + (total_deposit if proj_dep_received else 0.0)
        remaining_due = max(0.0, total_value - total_paid) if has_target else 0.0
        paid_pct = min(100, (total_paid / total_value * 100)) if has_target and total_value > 0 else 0

        kpi_row = CTkFrame(info, fg_color="transparent")
        kpi_row.grid(row=next_row, column=0, columnspan=4, sticky="ew", padx=12, pady=(6, 14))
        for i in range(4):
            kpi_row.grid_columnconfigure(i, weight=1)

        def _kpi_card(col, title, value_text, value_color, note=None, progress=None, bg_color="white"):
            card = CTkFrame(kpi_row, fg_color=bg_color, corner_radius=10, border_width=1, border_color="#E2E8F0")
            card.grid(row=0, column=col, sticky="new", padx=6)
            CTkLabel(card, text=title, font=CTkFont(size=13), text_color="#64748B").pack(
                anchor="w", padx=16, pady=(12, 2))
            CTkLabel(card, text=value_text, font=CTkFont(size=26, weight="bold"),
                     text_color=value_color).pack(anchor="w", padx=16)
            # จองพื้นที่แถบ progress ไว้เท่ากันทุกการ์ด แม้การ์ดที่ไม่มี progress ก็ตาม — ไม่งั้นการ์ดที่มี
            # progress bar (ยอดชำระแล้ว) จะสูงกว่าการ์ดอื่นๆ เพราะ pack เนื้อหาตามจริงของแต่ละใบ (sticky="new")
            if progress is not None:
                bar = ctk.CTkProgressBar(card, height=8, progress_color=value_color)
                bar.set(progress / 100)
                bar.pack(fill="x", padx=16, pady=(6, 0))
                CTkLabel(card, text=f"{progress:.1f}%", font=CTkFont(size=11),
                         text_color="#94A3B8").pack(anchor="e", padx=16, pady=(0, 12))
            else:
                CTkFrame(card, fg_color="transparent", height=8).pack(fill="x", padx=16, pady=(6, 0))
                # หมายเหตุ (ถ้ามี) มาแทนที่ตำแหน่งเดียวกับเลข % ของการ์ดที่มี progress — ให้อยู่บรรทัด
                # เดียวกันพอดี ไม่ใช่แยกไปอีกบรรทัดข้างล่าง
                CTkLabel(card, text=note or " ", font=CTkFont(size=11), text_color="#94A3B8").pack(
                    anchor="w", padx=16, pady=(0, 12))

        _kpi_card(0, "มูลค่าโครงการรวม (NET ทั้งบิล)", f"{total_value:,.2f} บาท", "#1E293B",
                   note="*รวม ค่าสินค้า + ค่าส่ง + VAT 7%", bg_color="white")
        _kpi_card(1, "ยอดชำระแล้ว", f"{total_paid:,.2f} บาท", "#16A34A", progress=paid_pct,
                   bg_color="#F0FDF4")
        _kpi_card(2, "ยอดคงเหลือค้างชำระ", f"{remaining_due:,.2f} บาท", "#D97706",
                   bg_color="#FFFBEB")
        _kpi_card(3, "ยอดคิดคอมมิชชั่น สะสม", f"{commission_sum:,.2f} บาท", "#2563EB",
                   note="*ตัด VAT & ค่าธรรมเนียมแล้ว", bg_color="#EFF6FF")
        next_row += 1

        # ── กล่องมัดจำ (checkbox "ได้รับมัดจำแล้ว" + ปุ่ม "เงินมัดจำพร้อมใช้งาน") ──────────────
        if deposit_pct_val is not None and pd.notna(deposit_pct_val):
            CTkLabel(info, text=f"ยอดมัดจำ ({float(deposit_pct_val):.0f}% ของยอดที่ต้องชำระ, {method_text}): "
                                 f"{total_deposit:,.2f} บาท{manual_note}{estimate_note}",
                     font=info_value_font, text_color="#7C3AED").grid(
                row=next_row, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 8))
            next_row += 1
            # มติที่ประชุม: ไม่ว่ามัดจำจะเป็นวิธี "กระจายทุก Lot" หรือ "หักที่ Lot สุดท้าย" ก็ตาม ลูกค้าจะโอน
            # มัดจำมาเป็นก้อนเดียวตั้งแต่แรก ก่อน Lot 1 จะส่งของเสมอ (ดูสลิปตัวอย่างจริง SO6908ID011) จึงใช้
            # flag ระดับโครงการตัวเดียวติดตาม "ได้รับมัดจำแล้ว" สำหรับทั้ง 2 วิธี — ไม่ผูกกับ deposit_method
            dep_recv_row = CTkFrame(info, fg_color="transparent")
            dep_recv_row.grid(row=next_row, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 8))
            dep_recv_var = ctk.BooleanVar(value=proj_dep_received)
            CTkCheckBox(dep_recv_row, text="ได้รับมัดจำจากลูกค้าแล้ว (ทั้งก้อน)", variable=dep_recv_var,
                        text_color="#7C3AED",
                        command=lambda: self._toggle_project_deposit_received(
                            project_id, dep_recv_var.get())).pack(side="left")
            if not proj_dep_received:
                CTkLabel(dep_recv_row, text="  ⚠ ต้องได้รับมัดจำก่อนส่งของ Lot แรกเสมอ",
                         font=self.small_font, text_color="#B45309").pack(side="left")
            elif deposit_method == "last_lot":
                CTkButton(dep_recv_row, text="เงินมัดจำพร้อมใช้งาน", width=140, height=24,
                          fg_color="#EDE9FE", hover_color="#DDD6FE", text_color="#6D28D9",
                          font=CTkFont(size=11),
                          command=lambda: messagebox.showinfo(
                              "เงินมัดจำ",
                              f"ได้รับมัดจำแล้ว {total_deposit:,.2f} บาท ({float(deposit_pct_val):.0f}% "
                              "ของยอดทั้งโครงการ)\nสถานะ: Hold ตลอดโปรเจกต์ — ระบบจะนำไปหักคืนให้อัตโนมัติ "
                              "ที่ Lot สุดท้ายตามแผน", parent=self)).pack(side="right", padx=(8, 0))
            next_row += 1

        if exceeded:
            CTkLabel(info, text="⚠ รวม Lot เกินมูลค่าโครงการ", font=info_sub_font,
                     text_color="#B45309").grid(row=next_row, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 14))
        else:
            CTkLabel(info, text="").grid(row=next_row, column=0, pady=(0, 8))

        lots_header = CTkFrame(self.body, fg_color="transparent")
        lots_header.grid(row=2, column=0, sticky="ew")
        lots_header.grid_columnconfigure(0, weight=1)
        CTkLabel(lots_header, text="📋 ตารางรายการ Lot (SO ย่อยสำหรับส่งสินค้า)",
                 font=CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w")
        CTkLabel(lots_header,
                 text="แสดงการถอดค่าสินค้าเพื่อคิดคอมมิชชั่น ควบคู่กับยอดเงินรับชำระจริง — "
                      "ดับเบิลคลิกเพื่ออัปเดตสถานะ · เพิ่ม Lot ใหม่จากหน้า \"สร้าง/แก้ไข Sales Order\"",
                 font=self.small_font, text_color="#94A3B8").grid(row=1, column=0, sticky="w")

        lots_frame = CTkFrame(self.body, fg_color="transparent")
        lots_frame.grid(row=3, column=0, sticky="nsew", pady=(6, 0))
        self.body.grid_rowconfigure(3, weight=1)
        lots_frame.grid_columnconfigure(0, weight=1)
        lots_frame.grid_rowconfigure(0, weight=1)

        columns = ["lot", "name", "product_type", "so", "product_amount", "shipping_amount",
                   "vat_total", "net_total", "pay_status"]
        headers = {
            "lot": "LOT", "name": "ชื่อ LOT / รายละเอียด", "product_type": "ประเภท", "so": "SO ย่อย",
            "product_amount": "มูลค่าสินค้า (ถอด VAT/ค่าส่ง)", "shipping_amount": "ค่าขนส่ง",
            "vat_total": "VAT 7%", "net_total": "ยอดสุทธิ (โอนจริง)", "pay_status": "สถานะชำระเงิน",
        }
        tree = ttk.Treeview(lots_frame, columns=columns, show="headings", selectmode="browse")
        for c in columns:
            tree.heading(c, text=headers[c], anchor="center")
            width = {"lot": 50, "name": 170, "product_type": 100, "so": 140, "product_amount": 170,
                     "shipping_amount": 90, "vat_total": 90, "net_total": 130, "pay_status": 110}[c]
            tree.column(c, width=width, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        vsb2 = ttk.Scrollbar(lots_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb2.set)
        vsb2.grid(row=0, column=1, sticky="ns")
        tree.tag_configure("done", background="#DCFCE7")
        tree.tag_configure("partial", background="#FEF9C3")
        tree.tag_configure("cancelled", background="#F1F5F9", foreground="#94A3B8")

        self._lot_id_map = {}
        for _, r in lots_df.iterrows():
            cancelled = r["status"] == "Cancelled"
            paid = bool(r["payment_collected_flag"])
            partial = (r["delivered_flag"] or r["invoice_recorded_flag"]) and not paid
            if cancelled:
                tag = "cancelled"
                pay_status_text = "ยกเลิก (SO ถูกยกเลิก)"
            else:
                tag = "done" if paid else ("partial" if partial else "")
                pay_status_text = "● ชำระแล้ว" if paid else ("กำลังดำเนินการ" if partial else "ยังไม่เริ่ม")
            product_amount = float(r.get('product_amount', 0) or 0)
            shipping_amount = float(r.get('shipping_amount', 0) or 0)
            vat_total = float(r.get('vat_total', 0) or 0)
            grand_total_due = float(r.get('grand_total_due', 0) or 0)
            row_values = [
                f"L{int(r['lot_number'])}", r["lot_name"] or "-", r.get("product_type") or "-",
                r["so_number"] or "-",
                f"{product_amount:,.2f}", f"{shipping_amount:,.2f}",
                f"{vat_total:,.2f}", f"{grand_total_due:,.2f}",
                pay_status_text,
            ]
            iid = tree.insert("", "end", values=tuple(row_values), tags=(tag,) if tag else ())
            self._lot_id_map[iid] = int(r["id"])

        self._lot_tree = tree
        tree.bind("<Double-1>", self._on_lot_double_click)
        CTkLabel(self.body, text="คอลัมน์ \"ยอดสุทธิ (โอนจริง)\" = ยอดคิดคอมมิชชั่นควบคู่ยอดชำระจริง",
                 font=self.small_font, text_color="#94A3B8").grid(row=4, column=0, sticky="w", pady=(6, 0))

    def _open_close_project_dialog(self, project_id):
        """3b — GP True-Up (POL-KPI-PROJECT-001 หน้า 10): ปิดโปรเจกต์ คำนวณ GP จริงรวมทั้งโปรเจกต์
        จากข้อมูล commissions จริงของทุก Lot แล้วตัดสินว่า Commission Reserve ที่กันไว้ 50% ต่อ Lot
        จะจ่ายคืนกี่ % ตามเกณฑ์ GP>=15% จ่ายเต็ม / 7.5-14.99% จ่ายตามสัดส่วน / <7.5% ริบทั้งหมด"""
        try:
            df = pd.read_sql_query("""
                SELECT c.id, c.so_number, c.sale_key,
                       COALESCE(c.final_sales_amount, c.sales_service_amount, 0) AS final_sales_amount,
                       COALESCE(c.final_cost_amount, 0) AS final_cost_amount,
                       COALESCE(c.cost_multiplier, 1.03) AS cost_multiplier,
                       COALESCE(c.commission_reserve_amount, 0) AS commission_reserve_amount,
                       COALESCE(c.reserve_status, 'Pending') AS reserve_status
                FROM commissions c
                JOIN project_lots pl ON pl.so_number = c.so_number
                WHERE pl.project_id = %s AND c.is_active = 1
                  AND c.status NOT IN ('Cancelled', 'Cancelled by PU')
            """, self.pg_engine, params=(project_id,))
        except Exception as e:
            messagebox.showerror("Database Error", f"โหลดข้อมูลคำนวณ GP ไม่สำเร็จ: {e}", parent=self)
            return

        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบ SO ที่ผูกกับโครงการนี้ ปิดโปรเจกต์ไม่ได้", parent=self)
            return

        total_sales = float(df['final_sales_amount'].sum())
        total_cost = float((df['final_cost_amount'] * df['cost_multiplier']).sum())
        gp_pct = ((total_sales - total_cost) / total_sales * 100) if total_sales > 0 else 0.0
        ratio = business_logic.calculate_reserve_release_ratio(gp_pct)

        pending_df = df[df['reserve_status'] == 'Pending'].copy()
        pending_df['release_amount'] = pending_df['commission_reserve_amount'] * ratio
        pending_df['forfeited_amount'] = pending_df['commission_reserve_amount'] - pending_df['release_amount']
        total_reserve = float(pending_df['commission_reserve_amount'].sum())
        total_release = float(pending_df['release_amount'].sum())
        total_forfeit = float(pending_df['forfeited_amount'].sum())

        win = CTkToplevel(self)
        win.title("ยืนยันปิดโปรเจกต์ — GP True-Up")
        win.grab_set()

        CTkLabel(win, text="🔒 ยืนยันปิดโปรเจกต์ — GP True-Up",
                 font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=18, pady=(16, 6))

        gp_box = CTkFrame(win, fg_color="#F0F4FF", corner_radius=8)
        gp_box.pack(fill="x", padx=18, pady=(0, 10))
        CTkLabel(gp_box, text=f"ยอดขายรวมทั้งโครงการ: {total_sales:,.2f} บาท", font=self.label_font
                 ).pack(anchor="w", padx=14, pady=(10, 0))
        CTkLabel(gp_box, text=f"ต้นทุนรวมทั้งโครงการ: {total_cost:,.2f} บาท", font=self.label_font
                 ).pack(anchor="w", padx=14, pady=(2, 0))
        gp_color = "#16A34A" if gp_pct >= 15 else ("#B45309" if gp_pct >= 7.5 else "#DC2626")
        CTkLabel(gp_box, text=f"GP จริงรวมทั้งโครงการ: {gp_pct:.2f}%",
                 font=CTkFont(size=14, weight="bold"), text_color=gp_color
                 ).pack(anchor="w", padx=14, pady=(4, 10))

        if gp_pct >= 15:
            tier_text = "✅ GP ≥ 15% → จ่าย Commission Reserve คืนเต็มจำนวน (100%)"
        elif gp_pct >= 7.5:
            tier_text = f"⚠ GP 7.50-14.99% → จ่าย Reserve คืนตามสัดส่วน = GP ÷ 15% = {ratio*100:.2f}%"
        else:
            tier_text = "❌ GP < 7.50% (หรือติดลบ) → ริบ Commission Reserve ทั้งหมด"
        CTkLabel(win, text=tier_text, font=CTkFont(size=12, weight="bold"),
                 text_color=gp_color, wraplength=620, justify="left").pack(anchor="w", padx=18, pady=(0, 10))

        if gp_pct < 0:
            CTkLabel(win, text=("⚠️ โปรเจกต์นี้ขาดทุนจริง (GP ติดลบ) — ตามนโยบาย ต้องให้ฝ่ายบริหารพิจารณา "
                                 "อนุมัติการหักคอมฯ ส่วนเกินจากงวดถัดไปด้วยตนเองเป็นกรณีไป ระบบจะ**ไม่หัก"
                                 "อัตโนมัติ** แค่ทำเครื่องหมายโครงการนี้ไว้ให้ฝ่ายบริหารตามงานต่อ"),
                     font=CTkFont(size=11), text_color="#DC2626", wraplength=620,
                     justify="left").pack(anchor="w", padx=18, pady=(0, 10))

        table_frame = CTkFrame(win, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        cols = ["so_number", "sale_key", "commission_reserve_amount", "release_amount", "forfeited_amount"]
        headers = {"so_number": "SO", "sale_key": "เซลส์", "commission_reserve_amount": "Reserve เดิม",
                   "release_amount": "จ่ายคืน", "forfeited_amount": "ริบ"}
        tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=min(8, max(3, len(pending_df))))
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=140, anchor="center" if c in ("so_number", "sale_key") else "e")
        tree.pack(fill="both", expand=True)
        for _, r in pending_df.iterrows():
            tree.insert("", "end", values=(
                r['so_number'], r['sale_key'],
                f"{r['commission_reserve_amount']:,.2f}",
                f"{r['release_amount']:,.2f}",
                f"{r['forfeited_amount']:,.2f}",
            ))
        if pending_df.empty:
            CTkLabel(table_frame, text="(ไม่มี Lot ที่มี Reserve ค้างสถานะ Pending ในโครงการนี้)",
                      text_color="gray").pack(pady=10)

        CTkLabel(win, text=(f"รวม Reserve เดิม {total_reserve:,.2f} บาท  →  "
                             f"จ่ายคืน {total_release:,.2f} บาท / ริบ {total_forfeit:,.2f} บาท"),
                 font=CTkFont(size=12, weight="bold"), text_color="#1F2937").pack(anchor="e", padx=18, pady=(0, 10))

        btn_frame = CTkFrame(win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=18, pady=(0, 16))
        CTkButton(btn_frame, text="ยกเลิก", fg_color="#94A3B8", hover_color="#64748B",
                  command=win.destroy).pack(side="left", expand=True, fill="x", padx=(0, 6))
        CTkButton(btn_frame, text="✅ ยืนยันปิดโปรเจกต์", fg_color="#B45309", hover_color="#92400E",
                  command=lambda: self._commit_close_project(
                      win, project_id, gp_pct, total_sales, total_cost, ratio, pending_df)
                  ).pack(side="left", expand=True, fill="x", padx=(6, 0))

        win.update_idletasks()
        W, H = 720, min(700, 380 + 24 * len(pending_df))
        _center_and_style_popup(win, self, W, H)

    def _commit_close_project(self, win, project_id, gp_pct, total_sales, total_cost, ratio, pending_df):
        try:
            conn = self.app.get_connection()
            try:
                with conn.cursor() as cur:
                    for _, r in pending_df.iterrows():
                        if ratio >= 1.0:
                            new_status = 'Paid'
                        elif ratio > 0:
                            new_status = 'PartiallyPaid'
                        else:
                            new_status = 'Forfeited'
                        cur.execute("""
                            UPDATE commissions
                            SET reserve_status = %s, reserve_decided_at = NOW(), reserve_decided_by = %s
                            WHERE id = %s
                        """, (new_status, self.user_key, int(r['id'])))

                        release_amt = float(r['release_amount'])
                        if release_amt > 0:
                            cur.execute("""
                                INSERT INTO reserve_release_queue
                                    (commission_id, project_id, sale_key, so_number, reserve_amount,
                                     release_ratio, release_amount, forfeited_amount, project_gp_pct,
                                     status, decided_by, decided_at)
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, 'Pending', %s, NOW())
                            """, (int(r['id']), project_id, r['sale_key'], r['so_number'],
                                  float(r['commission_reserve_amount']), float(ratio), release_amt,
                                  float(r['forfeited_amount']), float(gp_pct), self.user_key))

                    close_note = None
                    if gp_pct < 0:
                        close_note = "GP ติดลบ (ขาดทุนจริง) — รอฝ่ายบริหารพิจารณาอนุมัติหักคอมฯ งวดถัดไปเอง"
                    cur.execute("""
                        UPDATE projects
                        SET status = 'Closed', closed_at = NOW(), closed_by = %s,
                            final_gp_pct = %s, final_sales_amount = %s, final_cost_amount = %s,
                            needs_director_approval = %s, close_note = %s
                        WHERE id = %s
                    """, (self.user_key, float(gp_pct), float(total_sales), float(total_cost),
                          gp_pct < 0, close_note, project_id))
                conn.commit()
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", f"ปิดโปรเจกต์ไม่สำเร็จ: {e}", parent=self)
            return

        win.destroy()
        messagebox.showinfo("สำเร็จ", "ปิดโปรเจกต์และคำนวณ GP True-Up เรียบร้อยแล้ว", parent=self)
        self._show_detail(project_id)

    def _on_lot_double_click(self, event=None):
        sel = self._lot_tree.focus()
        if not sel:
            return
        lot_id = self._lot_id_map.get(sel)
        if lot_id is None:
            return
        self._open_lot_status_dialog(lot_id)

    def _open_lot_status_dialog(self, lot_id):
        try:
            conn = self.app.get_connection()
            try:
                df = pd.read_sql_query("""
                    SELECT pl.*, p.deposit_pct AS project_deposit_pct, p.deposit_method AS project_deposit_method,
                           p.deposit_received_flag AS project_deposit_received
                    FROM project_lots pl
                    JOIN projects p ON p.id = pl.project_id
                    WHERE pl.id = %s
                """, conn, params=(lot_id,))
            finally:
                self.app.release_connection(conn)
        except Exception as e:
            messagebox.showerror("Database Error", str(e), parent=self)
            return
        if df.empty:
            return
        lot = df.iloc[0]

        if lot["status"] == "Cancelled":
            messagebox.showwarning(
                "Lot นี้ถูกยกเลิกแล้ว",
                f"SO {lot['so_number']} ที่ผูกกับ Lot นี้ถูกยกเลิกในระบบ\n"
                "ไม่สามารถแก้ไขสถานะได้ — ถ้าต้องการให้ Lot นี้กลับมาใช้งาน "
                "ต้องยกเลิกการยกเลิก SO ก่อน", parent=self)
            return

        # checkbox "รับมัดจำแล้ว" ต่อ Lot มีความหมายเฉพาะวิธี "กระจายทุก Lot" เท่านั้น (แต่ละ Lot มีส่วนมัดจำของตัวเอง)
        # ส่วนวิธี "หักที่ Lot สุดท้าย" มัดจำเก็บเป็นก้อนเดียวตอนทำสัญญา จึงย้ายไป flag ระดับโครงการแทน
        # (ดู self._toggle_project_deposit_received ที่หน้า Project detail)
        has_deposit = (lot["project_deposit_pct"] is not None and pd.notna(lot["project_deposit_pct"])
                       and (lot.get("project_deposit_method") or "spread") == "spread")

        dlg = CTkToplevel(self)
        dlg.title(f"อัปเดตสถานะ Lot {int(lot['lot_number'])} — {lot['lot_name'] or ''}")
        _center_and_style_popup(dlg, self, 360, 320 if has_deposit else 260)
        dlg.grab_set()

        delivered_var = ctk.BooleanVar(value=bool(lot["delivered_flag"]))
        invoiced_var = ctk.BooleanVar(value=bool(lot["invoice_recorded_flag"]))
        deposit_received_var = ctk.BooleanVar(value=bool(lot.get("deposit_received_flag") or False))

        CTkCheckBox(dlg, text="ส่งของครบ 100% (มีใบส่งของที่ลูกค้าเซ็นรับ)",
                    variable=delivered_var).pack(anchor="w", padx=18, pady=(20, 10))
        CTkCheckBox(dlg, text="ออก Invoice / วางบิลแล้ว",
                    variable=invoiced_var).pack(anchor="w", padx=18, pady=10)
        if has_deposit:
            CTkCheckBox(dlg, text="รับมัดจำแล้ว", variable=deposit_received_var,
                        text_color="#7C3AED").pack(anchor="w", padx=18, pady=10)

        paid_now = bool(lot["payment_collected_flag"])
        paid_row = CTkFrame(dlg, fg_color="transparent")
        paid_row.pack(anchor="w", padx=18, pady=10, fill="x")
        CTkLabel(paid_row, text=("✅" if paid_now else "⏳") + " เก็บเงินครบ 100%",
                 text_color="#166534" if paid_now else "#94A3B8").pack(side="left")
        CTkLabel(paid_row, text="(เช็คอัตโนมัติจากยอดชำระ แก้ไขเองไม่ได้)",
                 font=CTkFont(size=11), text_color="#94A3B8").pack(side="left", padx=(6, 0))

        def _save():
            delivered = delivered_var.get()
            invoiced = invoiced_var.get()
            deposit_received = deposit_received_var.get() if has_deposit else bool(lot.get("deposit_received_flag") or False)

            # มติที่ประชุม: ต้องได้รับมัดจำ (ก้อนเดียว ระดับโครงการ) ก่อนถึงจะส่งของ Lot แรกได้ — ไม่ว่าจะเป็น
            # มัดจำวิธีไหนก็ตาม กันไว้ตรงนี้เพื่อไม่ให้เผลอติ๊ก "ส่งของครบ" ของ Lot 1 ก่อนเก็บมัดจำจริง
            if (delivered and int(lot["lot_number"]) == 1
                    and not bool(lot.get("project_deposit_received") or False)):
                messagebox.showwarning(
                    "ยังไม่ได้รับมัดจำ",
                    "โครงการนี้ยังไม่ได้ติ๊ก \"ได้รับมัดจำจากลูกค้าแล้ว\" ที่หน้ารายละเอียดโครงการ\n"
                    "ต้องได้รับมัดจำก่อนถึงจะบันทึกว่า Lot แรกส่งของครบได้", parent=dlg)
                return
            try:
                conn = self.app.get_connection()
                try:
                    with conn.cursor() as cur:
                        cur.execute("""
                            UPDATE project_lots
                            SET delivered_flag = %s,
                                invoice_recorded_flag = %s,
                                deposit_received_flag = %s,
                                deposit_received_date = CASE WHEN %s THEN COALESCE(deposit_received_date, CURRENT_DATE)
                                                              ELSE NULL END
                            WHERE id = %s
                        """, (delivered, invoiced, deposit_received, deposit_received, lot_id))
                    conn.commit()
                    # ให้ _sync_payment_collected คำนวณ kpi_qualified_flag/status ใหม่จากยอดชำระจริง
                    self._sync_payment_collected(conn, self.current_project_id)
                finally:
                    self.app.release_connection(conn)
            except Exception as e:
                messagebox.showerror("Database Error", f"อัปเดตสถานะไม่สำเร็จ: {e}", parent=dlg)
                return
            dlg.destroy()
            self._show_detail(self.current_project_id)

        btn_frame = CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=18, pady=20)
        CTkButton(btn_frame, text="ยกเลิก", fg_color="#94A3B8", hover_color="#64748B",
                  command=dlg.destroy).pack(side="left", expand=True, fill="x", padx=(0, 6))
        CTkButton(btn_frame, text="บันทึก", command=_save).pack(side="left", expand=True, fill="x", padx=(6, 0))
