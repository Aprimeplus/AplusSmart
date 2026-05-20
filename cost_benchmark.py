import tkinter as tk
from tkinter import messagebox
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkButton, CTkOptionMenu, CTkEntry, CTkCheckBox, CTkScrollableFrame, CTkToplevel
import pandas as pd
import psycopg2.extras
from datetime import datetime
import re
import json
from tkinter import colorchooser
from thefuzz import fuzz, process

# ── Windows DPI Awareness ────────────────────────────────────────────────────
# ต้องเรียกก่อน tkinter สร้าง window ใดๆ เพื่อให้ winfo_rooty() และ geometry()
# ใช้ coordinate space เดียวกัน (physical pixels) — ป้องกัน popup ผิดตำแหน่ง
# บน Hi-DPI (Display Scale > 100%)
try:
    import sys, ctypes
    if sys.platform == "win32":
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4 (Windows 10 1703+)
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()   # fallback: system DPI aware
        except Exception:
            pass
# ─────────────────────────────────────────────────────────────────────────────
try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False

# ============================================================
# เพิ่ม class นี้ก่อน class CostBenchmarkScreen
# ============================================================

class AutoFilterManager:
    """จัดการ AutoFilter แบบ Excel สำหรับ tksheet"""

    def __init__(self, parent_screen):
        self.screen = parent_screen
        self._filter_values = {}       # {col_name: set of selected values}
        self._all_values_cache = {}    # {col_name: sorted list of all values}
        self._active_popup = None
        self._filter_arrow_ids = {}



    # ── ดึงค่าทั้งหมดในคอลัมน์ (รวม frozen + main) ──────────────
    def _get_all_values(self, col_name):
        screen = self.screen
        col_offset = screen.frozen_col_count if screen.sheet_frozen else 0
        try:
            real_idx = screen.columns.index(col_name)
        except ValueError:
            return []

        values = set()
        sheet_side, disp = screen._real_to_disp(real_idx)
        if sheet_side is None:
            return []
        tgt = screen.sheet_frozen if sheet_side == 'frozen' else screen.sheet
        for r in range(tgt.get_total_rows()):
            v = str(tgt.get_cell_data(r, disp) or "").strip()
            if v:
                values.add(v)
        return sorted(values)

    # ── เปิด popup filter (เรียกจาก AutoFilterManager.show_filter_popup) ──
    def show_filter_popup(self, col_name, x_root, y_root):
        """เปิด popup filter แบบ checkbox เหมือน Excel"""
        self.close_popup()

        all_vals = self._get_all_values(col_name)
        self._all_values_cache[col_name] = all_vals
        if not all_vals:
            return

        # ค่าที่เลือกอยู่ตอนนี้ (ถ้าไม่มีให้ถือว่าเลือกทั้งหมด)
        current = self._filter_values.get(col_name, None)
        active_set = set(current) if current is not None else set(all_vals)

        import tkinter as tk

        popup = tk.Toplevel(self.screen)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.configure(bg="#D1D5DB", padx=1, pady=1)
        popup._destroyed = False
        self._active_popup = popup

        def safe_destroy():
            if not popup._destroyed:
                popup._destroyed = True
                try:
                    popup.destroy()
                except Exception:
                    pass
            self._active_popup = None

        popup.safe_destroy = safe_destroy

        inner = tk.Frame(popup, bg="white")
        inner.pack(fill="both", expand=True)

        # ── Sort buttons ──────────────────────────────────────────
        sort_bar = tk.Frame(inner, bg="#F9FAFB")
        sort_bar.pack(fill="x", padx=4, pady=(4, 2))

        def sort_asc():
            self._sort_column(col_name, ascending=True)
            safe_destroy()

        def sort_desc():
            self._sort_column(col_name, ascending=False)
            safe_destroy()

        tk.Button(sort_bar, text="↑ Sort A to Z", font=("Tahoma", 10),
                  relief="flat", bg="#F9FAFB", fg="#1F2937", anchor="w",
                  cursor="hand2", command=sort_asc).pack(fill="x", pady=1)
        tk.Button(sort_bar, text="↓ Sort Z to A", font=("Tahoma", 10),
                  relief="flat", bg="#F9FAFB", fg="#1F2937", anchor="w",
                  cursor="hand2", command=sort_desc).pack(fill="x", pady=1)

        tk.Frame(inner, bg="#E5E7EB", height=1).pack(fill="x", padx=4, pady=2)

        # ── Search ────────────────────────────────────────────────
        search_var = tk.StringVar()
        search_entry = tk.Entry(inner, textvariable=search_var, font=("Tahoma", 10),
                                relief="flat", bd=5,
                                highlightthickness=1, highlightcolor="#3B82F6",
                                highlightbackground="#D1D5DB")
        search_entry.pack(fill="x", padx=4, pady=2)
        search_entry.insert(0, "🔍 ค้นหา...")
        search_entry.bind("<FocusIn>", lambda e: search_entry.delete(0, "end")
                          if search_entry.get().startswith("🔍") else None)
        search_entry.bind("<Control-c>",  lambda e: (search_entry.event_generate("<<Copy>>"),  "break")[1])
        search_entry.bind("<Control-C>",  lambda e: (search_entry.event_generate("<<Copy>>"),  "break")[1])
        search_entry.bind("<Control-v>",  lambda e: (search_entry.event_generate("<<Paste>>"), "break")[1])
        search_entry.bind("<Control-V>",  lambda e: (search_entry.event_generate("<<Paste>>"), "break")[1])
        def _se_thai_cb(e):
            if e.keycode == 86: search_entry.event_generate("<<Paste>>"); return "break"
            if e.keycode == 67: search_entry.event_generate("<<Copy>>"); return "break"
        search_entry.bind("<Control-KeyPress>", _se_thai_cb, add="+")

        # ── Checkbox list ─────────────────────────────────────────
        check_frame = tk.Frame(inner, bg="white")
        check_frame.pack(fill="both", expand=True, padx=4)

        canvas_c = tk.Canvas(check_frame, bg="white", highlightthickness=0, height=200)
        sb_c = tk.Scrollbar(check_frame, orient="vertical", command=canvas_c.yview)
        canvas_c.configure(yscrollcommand=sb_c.set)
        sb_c.pack(side="right", fill="y")
        canvas_c.pack(side="left", fill="both", expand=True)

        check_inner = tk.Frame(canvas_c, bg="white")
        canvas_c.create_window((0, 0), window=check_inner, anchor="nw")

        check_vars = {}
        check_widgets = []

        # สร้าง widget ครั้งเดียว แล้วใช้ pack/pack_forget ซ่อน/แสดง (เร็วกว่า destroy มาก)
        var_all = tk.BooleanVar()
        cb_all = tk.Checkbutton(check_inner, text="(Select All)", variable=var_all,
                    font=("Tahoma", 10), bg="white", anchor="w", relief="flat",
                    activebackground="#EFF6FF")
        cb_all.pack(fill="x", pady=1)
        _row_widgets = []
        for v in all_vals:
            var = tk.BooleanVar(value=v in active_set)
            check_vars[v] = var
            cb = tk.Checkbutton(check_inner, text=v, variable=var,
                        font=("Tahoma", 10), bg="white", anchor="w", relief="flat",
                        activebackground="#EFF6FF")
            cb.pack(fill="x", pady=1)
            _row_widgets.append((v, var, cb))

        def build_checklist(term=""):
            check_widgets.clear()
            t = term.lower() if (term and not term.startswith("🔍")) else ""
            filtered = []
            for v, var, cb in _row_widgets:
                if not t or t in v.lower():
                    cb.pack(fill="x", pady=1)
                    var.set(v in active_set)
                    check_widgets.append((v, var, None))
                    filtered.append(v)
                else:
                    cb.pack_forget()
            var_all.set(bool(filtered) and all(v in active_set for v in filtered))
            cb_all.configure(command=lambda f=filtered: toggle_all(var_all.get(), f))
            check_widgets.insert(0, ("__all__", var_all, filtered))
            check_inner.update_idletasks()
            canvas_c.configure(scrollregion=canvas_c.bbox("all"))

        def toggle_all(state, vals):
            for v in vals:
                if v in check_vars:
                    check_vars[v].set(state)
            if state: active_set.update(vals)
            else:
                for v in vals: active_set.discard(v)

        # Debounce 200ms — ไม่ filter จนกว่าจะพิมพ์หยุด
        _sjob = [None]
        def _on_search(*a):
            if _sjob[0]:
                try: popup.after_cancel(_sjob[0])
                except Exception: pass
            _sjob[0] = popup.after(200, lambda: build_checklist(search_var.get()))
        search_var.trace_add("write", _on_search)
        build_checklist()

        # ── OK / Cancel ───────────────────────────────────────────
        btn_bar = tk.Frame(inner, bg="white")
        btn_bar.pack(fill="x", padx=4, pady=4)

        def on_ok():
            new_sel = {v for v, var, _ in check_widgets
                       if v != "__all__" and var.get()}
            if new_sel == set(all_vals):
                self._filter_values.pop(col_name, None)
            else:
                self._filter_values[col_name] = new_sel
            safe_destroy()
            self.apply_filters()

        tk.Button(btn_bar, text="OK", font=("Tahoma", 10), width=8,
                  bg="#3B82F6", fg="white", relief="flat", cursor="hand2",
                  command=on_ok).pack(side="right", padx=(4, 0))
        tk.Button(btn_bar, text="Cancel", font=("Tahoma", 10), width=8,
                  bg="#F3F4F6", fg="#374151", relief="flat", cursor="hand2",
                  command=safe_destroy).pack(side="right")

        # ── Clear filter link (แสดงเฉพาะถ้า filter ทำงานอยู่) ────
        if col_name in self._filter_values:
            def clear_this():
                self._filter_values.pop(col_name, None)
                safe_destroy()
                self.apply_filters()
            tk.Button(inner, text=f'✕ Clear Filter From "{col_name}"',
                      font=("Tahoma", 9), fg="#EF4444", bg="white",
                      relief="flat", cursor="hand2", command=clear_this
                      ).pack(anchor="w", padx=8, pady=(0, 4))

        # ── วางตำแหน่ง popup ──────────────────────────────────────
        # ✅ ใช้ขอบเขตของ window จริง (รองรับ multi-monitor และ DPI scaling)
        # แทนการใช้ winfo_screenwidth/height ซึ่งผิดพลาดกับบาง user
        try:
            root = self.screen.winfo_toplevel()
            win_x  = root.winfo_rootx()
            win_y  = root.winfo_rooty()
            win_w  = root.winfo_width()
            win_h  = root.winfo_height()
        except Exception:
            win_x, win_y = 0, 0
            win_w = self.screen.winfo_screenwidth()
            win_h = self.screen.winfo_screenheight()

        pw, ph = 240, 400
        x = x_root
        y = y_root + 5

        # กันเกินขอบขวาของ window
        if x + pw > win_x + win_w - 5:
            x = win_x + win_w - pw - 5
        # ถ้าเกินล่าง → เปิดขึ้นบนแทน
        if y + ph > win_y + win_h - 5:
            y = y_root - ph - 5
        # กันออกนอกขอบซ้าย/บน
        x = max(x, win_x + 5)
        y = max(y, win_y + 5)

        # withdraw ก่อน set geometry เพื่อป้องกัน popup กระพริบที่ (0,0)
        popup.withdraw()
        popup.geometry(f"{pw}x{ph}+{x}+{y}")

        def _show_and_focus():
            popup.deiconify()
            try:
                search_entry.focus_set()
            except Exception:
                pass

        popup.after(1, _show_and_focus)

        popup.bind("<FocusOut>", lambda e: self.screen.after(
            300, lambda: safe_destroy() if not popup._destroyed and
            self.screen.focus_get() not in popup.winfo_children() else None))

    def close_popup(self):
        if self._active_popup and not getattr(self._active_popup, '_destroyed', True):
            try:
                self._active_popup.safe_destroy()
            except Exception:
                pass
        self._active_popup = None

    # ── Apply filters (ซ่อน row ที่ไม่ตรง) ──────────────────────
    def apply_filters(self):
        screen = self.screen
        col_offset = screen.frozen_col_count if screen.sheet_frozen else 0
        rh = screen.zoom_level + 19

        main_filter_map = {}
        frozen_filter_map = {}

        for col_name, vals in self._filter_values.items():
            side, disp = screen._col_to_disp(col_name)
            if side is None:
                continue
            if side == 'frozen':
                frozen_filter_map[disp] = vals
            else:
                main_filter_map[disp] = vals

        total_rows = screen.sheet.get_total_rows()
        for r in range(total_rows):
            show = True
            for d_idx, vals in main_filter_map.items():
                v = str(screen.sheet.get_cell_data(r, d_idx) or "").strip()
                if isinstance(vals, set):
                    if v not in vals:
                        show = False
                        break
                else:
                    if v != vals:
                        show = False
                        break
            if show and screen.sheet_frozen:
                for f_idx, vals in frozen_filter_map.items():
                    v = str(screen.sheet_frozen.get_cell_data(r, f_idx) or "").strip()
                    if isinstance(vals, set):
                        if v not in vals:
                            show = False
                            break
                    else:
                        if v != vals:
                            show = False
                            break
            screen.sheet.row_height(r, rh if show else 0)
            if screen.sheet_frozen:
                screen.sheet_frozen.row_height(r, rh if show else 0)

        screen.sheet.redraw()
        if screen.sheet_frozen:
            screen.sheet_frozen.redraw()

        self._update_header_indicators()

    # ── Sort column ────────────────────────────────────────────
    def _sort_column(self, col_name, ascending=True):
        screen = self.screen
        col_offset = screen.frozen_col_count if screen.sheet_frozen else 0
        try:
            real_idx = screen.columns.index(col_name)
        except ValueError:
            return

        total_rows = screen.sheet.get_total_rows()
        rows_data = []
        for r in range(total_rows):
            main_row = list(screen.sheet.get_row_data(r))
            frozen_row = list(screen.sheet_frozen.get_row_data(r)) if screen.sheet_frozen else []
            rows_data.append(frozen_row + main_row)

        def sort_key(row):
            v = row[real_idx] if real_idx < len(row) else ""
            v = str(v).strip()
            try:
                return (0, float(v.replace(",", "")))
            except ValueError:
                return (1, v.lower())

        rows_data.sort(key=sort_key, reverse=not ascending)

        for r, row in enumerate(rows_data):
            if screen.sheet_frozen:
                for c in range(col_offset):
                    screen.sheet_frozen.set_cell_data(r, c, row[c] if c < len(row) else "", redraw=False)
            for c in range(screen.sheet.get_total_columns()):
                rc = c + col_offset
                screen.sheet.set_cell_data(r, c, row[rc] if rc < len(row) else "", redraw=False)

        screen.sheet.redraw()
        if screen.sheet_frozen:
            screen.sheet_frozen.redraw()

        screen.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
        if screen.auto_save_job_id:
            screen.after_cancel(screen.auto_save_job_id)
        screen.auto_save_job_id = screen.after(1500, lambda: screen._save_to_db(show_msg=False))

    def clear_all(self):
        self._filter_values.clear()
        self.apply_filters()

    def has_filter(self, col_name):
        return col_name in self._filter_values

    def _update_header_indicators(self):
        """อัพเดทสีหัวคอลัมน์เมื่อมี/ไม่มี filter"""
        screen = self.screen
        col_offset = screen.frozen_col_count if screen.sheet_frozen else 0
        header_styles = screen._get_header_styles_map()
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles.items() for c in cols}

        for col_name in screen.columns:
            try:
                real_idx = screen.columns.index(col_name)
                display_idx = real_idx - col_offset
                if display_idx < 0:
                    # อยู่ใน frozen — update frozen header
                    if screen.sheet_frozen:
                        if self.has_filter(col_name):
                            screen.sheet_frozen.highlight_cells(
                                row=0, column=real_idx, bg="#F59E0B", fg="white", canvas="header")
                        else:
                            h_bg, h_fg = col_to_style.get(col_name, ("#E5E7EB", "#111827"))
                            screen.sheet_frozen.highlight_cells(
                                row=0, column=real_idx, bg=h_bg, fg=h_fg, canvas="header")
                    continue
                if self.has_filter(col_name):
                    screen.sheet.highlight_cells(row=0, column=display_idx,
                                                 bg="#F59E0B", fg="white", canvas="header")
                else:
                    h_bg, h_fg = col_to_style.get(col_name, ("#E5E7EB", "#111827"))
                    screen.sheet.highlight_cells(row=0, column=display_idx,
                                                 bg=h_bg, fg=h_fg, canvas="header")
            except Exception:
                pass

        try:
            screen.sheet.redraw()
            if screen.sheet_frozen:
                screen.sheet_frozen.redraw()
        except Exception:
            pass

class InlineSearchPopup(tk.Toplevel):
    def __init__(self, master, data_list, on_select_callback):
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg="#9CA3AF", padx=1, pady=1)

        self.data_list = data_list
        
        # 🚀 อัปเกรดความเร็ว 1: Precompute ข้อมูลไว้ล่วงหน้าเพื่อไม่ต้องแปลง Type ทุกครั้งที่พิมพ์
        self._data_lower = [str(x).lower() for x in data_list]
        self._data_str = [str(x) for x in data_list] 
        
        self.on_select_callback = on_select_callback
        self._destroyed = False
        self._search_job = None

        self.search_var = tk.StringVar()
        self.entry = tk.Entry(
            self,
            textvariable=self.search_var,
            font=("Tahoma", 11),
            relief="flat",
            bd=4,
            highlightthickness=0,
        )
        self.entry.pack(fill="x")
        self.search_var.trace_add("write", self._on_type)

        list_frame = tk.Frame(self, bg="#9CA3AF")
        list_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(
            list_frame,
            font=("Tahoma", 11),
            selectbackground="#3B82F6",
            selectforeground="white",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            height=8,
        )
        
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.listbox.pack(side="left", fill="both", expand=True)

        self.listbox.bind("<ButtonRelease-1>", self._on_click)
        self.entry.bind("<Down>", self._on_down)
        self.entry.bind("<Up>", self._on_up)
        self.entry.bind("<Return>", self._on_return)
        self.entry.bind("<KP_Enter>", self._on_return)
        self.entry.bind("<Escape>", self._on_escape)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<Control-c>",  lambda e: (self.entry.event_generate("<<Copy>>"),  "break")[1])
        self.entry.bind("<Control-C>",  lambda e: (self.entry.event_generate("<<Copy>>"),  "break")[1])
        self.entry.bind("<Control-v>",  lambda e: (self.entry.event_generate("<<Paste>>"), "break")[1])
        self.entry.bind("<Control-V>",  lambda e: (self.entry.event_generate("<<Paste>>"), "break")[1])
        self.entry.bind("<Control-KeyPress>", self._entry_thai_clipboard, add="+")

        self.filter_list("")
        self.entry.focus_set()

    def place_at_mouse(self, ref_widget=None):
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery() + 15
            self._set_geometry(x, y, 400, ref_widget=ref_widget)
        except Exception as e:
            print(f"place_at_mouse error: {e}")

    def place_near_cell(self, sheet_widget, row, col, row_h=30, header_h=35,
                        anchor_x=None, anchor_y=None, data_col=None):
        try:
            rx = sheet_widget.winfo_pointerx()
            ry = sheet_widget.winfo_pointery() + 8
            self._set_geometry(rx, ry, 400, ref_widget=sheet_widget)
        except Exception as e:
            print(f"place_near_cell error: {e}")

    def _set_geometry(self, rx, ry, popup_w, popup_h=240, ref_widget=None):
        try:
            if ref_widget is not None:
                root = ref_widget.winfo_toplevel()
                win_x = root.winfo_rootx()
                win_y = root.winfo_rooty()
                win_w = root.winfo_width()
                win_h = root.winfo_height()
            else:
                win_x, win_y = 0, 0
                win_w = self.winfo_screenwidth()
                win_h = self.winfo_screenheight()

            if rx + popup_w > win_x + win_w: rx = win_x + win_w - popup_w - 5
            if ry + popup_h > win_y + win_h: ry = ry - popup_h - 30

            rx = max(rx, win_x)
            ry = max(ry, win_y)

            self.withdraw()
            self.geometry(f"{popup_w}x{popup_h}+{rx}+{ry}")
            def _show_and_focus():
                self.deiconify()
                try: self.entry.focus_set()
                except Exception: pass
            self.after(1, _show_and_focus)
        except Exception:
            pass

    def filter_list(self, term):
        term = (term or "").lower().strip()
        self.listbox.delete(0, tk.END)
        
        # 🚀 อัปเกรดความเร็ว 2: ถ้าไม่ได้พิมพ์อะไร จำกัดการโชว์แค่ 150 รายการแรกเพื่อไม่ให้ค้าง
        if not term:
            if self.data_list:
                self.listbox.insert(tk.END, *self.data_list[:150]) 
            return

        search_terms = term.split()

        def _is_thai_char(c): return 'ก' <= c <= '๛'
        def _sort_key(item, term):
            item_lower = str(item).lower()
            starts = item_lower.startswith(term)
            if starts:
                rest = item_lower[len(term):]
                if not rest or _is_thai_char(rest[0]): return (0, len(item))
                else: return (1, len(item))
            else: return (2, len(item))

        # ค้นหาแบบธรรมดา (ใช้ตัวแปรที่ precompute แล้ว)
        normal_matches = []
        for i, item_lower in enumerate(self._data_lower):
            if all(t in item_lower for t in search_terms):
                normal_matches.append(self.data_list[i])
        
        # 🚀 อัปเกรดความเร็ว 3: จำกัดผลลัพธ์ปกติก่อนไป Sort เพื่อลดภาระ CPU
        normal_matches = normal_matches[:150]
        normal_matches.sort(key=lambda item: _sort_key(item, term))
        
        # 🚀 อัปเกรดความเร็ว 4: ค้นหาแบบ Fuzzy เมื่อพิมพ์อย่างน้อย 2 ตัวอักษรขึ้นไปเท่านั้น (พิมพ์ 1 ตัวแล้วทำ Fuzzy มันหาเยอะเกินจนค้าง)
        fuzzy_matches = []
        if len(normal_matches) < 10 and len(term) > 1:
            try:
                # ใช้ _data_str ที่เตรียมไว้แล้ว ไม่ต้องสั่ง str(i) ทุกรอบ
                fuzzy_results = process.extractBests(
                    term, self._data_str,
                    scorer=fuzz.token_set_ratio, limit=20, score_cutoff=55
                )
                normal_set = set(str(x) for x in normal_matches)
                fuzzy_matches = [m[0] for m in fuzzy_results if m[0] not in normal_set]
            except Exception:
                pass
            
        combined = list(dict.fromkeys(normal_matches + fuzzy_matches))
        
        # 🚀 อัปเกรดความเร็ว 5: จำกัดผลลัพธ์สุดท้ายที่จะส่งไปแสดงบนหน้าจอไม่เกิน 150 รายการ
        if combined:
            self.listbox.insert(tk.END, *combined[:150])

    def _on_type(self, *args):
        # Debounce 120ms ลดลงนิดหน่อยเพื่อให้รู้สึกตอบสนองเร็วขึ้น 
        if self._search_job:
            try: self.after_cancel(self._search_job)
            except Exception: pass
        self._search_job = self.after(120, lambda: self.filter_list(self.search_var.get()))

    def _on_click(self, e=None):
        if not self.listbox.curselection(): return
        self._commit(self.listbox.get(self.listbox.curselection()))

    def _on_down(self, e=None):
        if self.listbox.size() > 0:
            curr = self.listbox.curselection()
            idx = (curr[0] + 1) if curr else 0
            idx = min(idx, self.listbox.size() - 1)
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
        return "break"

    def _on_up(self, e=None):
        if self.listbox.size() > 0:
            curr = self.listbox.curselection()
            if curr:
                idx = max(curr[0] - 1, 0)
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.see(idx)
        return "break"

    def _on_return(self, e=None):
        if self.listbox.curselection():
            self._commit(self.listbox.get(self.listbox.curselection()))
        elif self.listbox.size() > 0:
            self._commit(self.listbox.get(0))
        return "break"

    def _entry_thai_clipboard(self, event):
        if event.keycode == 86:
            try: clip = self.entry.clipboard_get()
            except Exception: return "break"
            try: self.entry.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except Exception: pass
            self.entry.insert(tk.INSERT, clip)
            return "break"
        elif event.keycode == 67:
            try:
                text = self.entry.selection_get()
                self.entry.clipboard_clear()
                self.entry.clipboard_append(text)
            except Exception: pass
            return "break"

    def _on_escape(self, e=None):
        self.safe_destroy()

    def _on_focus_out(self, e=None):
        self.after(150, self._check_focus_and_close)

    def _check_focus_and_close(self):
        try:
            if self._destroyed: return
            focused = self.focus_get()
            if focused not in (self.entry, self.listbox):
                self.safe_destroy()
        except Exception:
            self.safe_destroy()

    def _commit(self, value):
        self.on_select_callback(value)
        self.safe_destroy()

    def safe_destroy(self):
        if not self._destroyed:
            self._destroyed = True
            try: self.destroy()
            except Exception: pass


class CostBenchmarkScreen(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container

        self.auto_save_job_id = None
        self.current_user = getattr(self.app_container, 'current_user_key', 'PU_Default')

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.frozen_col_count = 0
        self.zoom_level = 11
        # Default column widths (keyed by real col index) — ใกล้เคียง Excel
        # DB-saved values จะ override ทับได้ปกติใน _load_user_settings
        _col_default_w = [
            75,   # วันที่ขอราคา
            55,   # Order No.
            70,   # Sale Order No.
            55,   # รหัส Sale
            55,   # PRIORITY
            60,   # WIN RATE %
            55,   # สถานะ
            38,   # QT
            50,   # Select
            60,   # หมวด
            110,  # ชื่อ Supplier
            65,   # แบรนด์
            180,  # รายการสินค้า
            75,   # หมายเหตุ (ความยาว, OD)
            75,   # หมายเหตุ
            85,   # Product SKU.
            48,   # จำนวน
            70,   # ต้นทุน/เส้น
            70,   # น้ำหนัก/เส้น
            75,   # น้ำหนักรวม (Kg.)
            60,   # ทุน/กก.
            65,   # ทุนรวม
            70,   # ส่วนลด 1 (บาท)
            65,   # ส่วนลด 1 (%)
            90,   # ทุน/เส้น หลังส่วนลด 1
            70,   # ส่วนลด 2 (บาท)
            65,   # ส่วนลด 2 (%)
            90,   # ทุน/เส้น หลังส่วนลด 2
            95,   # ต้นทุน/กก. (ไม่รวมย้าย)
            95,   # ต้นทุน/เส้น (ไม่รวมย้าย)
            95,   # ต้นทุนรวม (ไม่รวมย้าย)
            70,   # ค่าย้าย (ซื้อ)
            65,   # ค่าย้าย/เส้น
            95,   # ต้นทุน/กก. (รวมย้าย)
            95,   # ต้นทุน/เส้น (รวมย้าย)
            95,   # ต้นทุนรวม (รวมย้าย)
            70,   # Markup Guide (%)
            65,   # Markup/กก.
            65,   # Markup/เส้น
            75,   # ทุน+Markup/กก.
            75,   # ทุน+Markup/เส้น
            90,   # ต้นทุนรวม+Markup
            70,   # ค่าส่ง (ขาย)
            65,   # ค่าส่ง / เส้น
            70,   # น้ำหนัก/เส้น 2
            75,   # ราคาขาย / กก.
            75,   # ราคาขาย / เส้น
            60,   # Vat. / เส้น
            85,   # ราคาขาย/เส้น + Vat.
            70,   # ราคาขาย รวม
            60,   # Vat. รวม
            85,   # ราคาขาย รวม + Vat.
            110,  # ชื่อ Supplier2
            50,   # Sup ID.
            85,   # คลังสินค้า ต้นทาง
            75,   # ปลายทาง
            75,   # หมายเหตุ2
        ]
        self.col_widths_cache = {i: w for i, w in enumerate(_col_default_w)}
        self.sales_list = []
        self.supplier_list = []
        self.product_list = []
        self._clipboard_row = []
        self.product_sku_map = {}
        self.supplier_code_map = {}
        self.product_category_map = {}
        self.hidden_cols_list = []
        self.custom_header_colors = {}
        self.sheet_frozen = None
        self._last_yview = -1.0
        self._active_popup = None
        self._last_popup_cell = None
        self._last_click_x = 0
        self._last_click_y = 0
        self._last_active_sheet = 'main'
        self._popup_opening = False
        self._header_filter_values = {}
        self._active_filter_popup = None
        self._last_popup_cell = None
        self._undo_stack = []
        self._redo_stack = []
        self._max_undo = 50

        # --- Header & Filters ---
        header_container = CTkFrame(self, fg_color="transparent")
        header_container.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))
        header_container.grid_columnconfigure(0, weight=1)

        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        now = datetime.now()

        # สร้างแถบ Scrollable แนวนอนยาวสุดจอ
        from customtkinter import CTkScrollableFrame
        btn_frame = CTkScrollableFrame(header_container, orientation="horizontal", height=45, fg_color="transparent")
        btn_frame.pack(fill="x")

        # 🔹 โซนที่ 1: ชื่อ และ ตัวกรอง
        CTkLabel(btn_frame, text=f"📊 ตารางของคุณ: {self.current_user}",
                 font=CTkFont(size=20, weight="bold"), text_color="#1F2937").pack(side="left", padx=(0, 20))

        CTkLabel(btn_frame, text="รอบบิลเดือน:").pack(side="left", padx=(0, 5))
        self.month_var = tk.StringVar(value=self.thai_months[now.month - 1])
        CTkOptionMenu(btn_frame, variable=self.month_var, values=self.thai_months, width=100,
                      command=self._load_from_db).pack(side="left", padx=(0, 10))

        CTkLabel(btn_frame, text="ปี:").pack(side="left", padx=(0, 5))
        current_year_th = str(now.year + 543)
        year_list = [str(int(current_year_th) + i) for i in range(-2, 3)]
        self.year_var = tk.StringVar(value=current_year_th)
        CTkOptionMenu(btn_frame, variable=self.year_var, values=year_list, width=80,
                      command=self._load_from_db).pack(side="left", padx=(0, 15))

        CTkButton(btn_frame, text="🔄 โหลดข้อมูล", fg_color="#3B82F6", hover_color="#2563EB", width=90,
                  command=self._load_from_db).pack(side="left", padx=(0, 20))

        # 🔹 โซนที่ 2: ปุ่มเครื่องมือ
        CTkButton(btn_frame, text="🎨 สีหัวคอลัมน์", fg_color="#EC4899", hover_color="#DB2777",
                  width=110, height=30, font=CTkFont(size=12),
                  command=self._change_header_color).pack(side="left", padx=2)
        CTkButton(btn_frame, text="📌 ตรึงคอลัมน์", fg_color="#0891B2", hover_color="#0E7490",
                  width=100, height=30, font=CTkFont(size=12),
                  command=self._freeze_selected_columns).pack(side="left", padx=2)
        CTkButton(btn_frame, text="📌 ยกเลิกตรึง", fg_color="#64748B", hover_color="#475569",
                  width=90, height=30, font=CTkFont(size=12),
                  command=self._unfreeze_columns).pack(side="left", padx=2)
        CTkButton(btn_frame, text="🗑️ ลบบรรทัด", fg_color="#EF4444", hover_color="#DC2626",
                  width=90, height=30, font=CTkFont(size=12),
                  command=self._delete_selected_rows).pack(side="left", padx=2)
        CTkButton(btn_frame, text="⚙️ จัดการคอลัมน์", fg_color="#8B5CF6", hover_color="#7C3AED",
                  width=130, height=30, font=CTkFont(size=12),
                  command=self._open_column_visibility_panel).pack(side="left", padx=2)
        CTkButton(btn_frame, text="🔄 รีเฟรชสินค้า", fg_color="#0891B2", hover_color="#0E7490",
                  width=120, height=30, font=CTkFont(size=12),
                  command=self._manual_refresh_dropdown).pack(side="left", padx=2)
        CTkButton(btn_frame, text="➕ เพิ่มบรรทัด",
                  width=90, height=30, font=CTkFont(size=12),
                  command=self._add_new_row).pack(side="left", padx=2)
        CTkButton(btn_frame, text="📋 Short Note",
                  width=100, height=30, font=CTkFont(size=12),
                  fg_color="#059669", hover_color="#047857",
                  command=self._show_short_note_popup).pack(side="left", padx=(8, 2))


        self.columns = [
            "วันที่ขอราคา", "Order No.", "Sale Order No.", "รหัส Sale",
            "PRIORITY", "WIN RATE %", "สถานะ", "QT", "Select",
            "หมวด", "ชื่อ Supplier", "แบรนด์", "รายการสินค้า",
            "หมายเหตุ (ความยาว, OD)", "หมายเหตุ", "Product SKU.", "จำนวน", "ต้นทุน/เส้น",
            "น้ำหนัก/เส้น", "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (บาท)",
            "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1", "ส่วนลด 2 (บาท)", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2",
            "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย (ซื้อ)", "ค่าย้าย/เส้น",
            "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)", "Markup Guide (%)", "Markup/กก.",
            "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup", "ค่าส่ง (ขาย)",
            "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น",
            "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat.", "ชื่อ Supplier2",
            "Sup ID.", "คลังสินค้า ต้นทาง", "ปลายทาง", "หมายเหตุ2"
        ]

        self._load_dropdown_data()

        # Formula bar
        self.target_formula_cell = None
        self.formula_frame = CTkFrame(self, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#CBD5E1")
        self.formula_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 5))
        CTkLabel(self.formula_frame, text=" 𝑓x ", font=CTkFont(family="Arial", size=18, weight="bold", slant="italic"),
                 text_color="#16A34A").pack(side="left", padx=10, pady=5)
        self.formula_entry = CTkEntry(self.formula_frame, font=CTkFont(size=14),
                                     placeholder_text="คลิกช่องปลายทาง -> คลิกที่นี่ -> พิมพ์ = แล้วใช้เมาส์จิ้มเซลล์ในตารางได้เลย!",
                                     border_width=0, fg_color="transparent")
        self.formula_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=5)
        self.formula_entry.bind("<FocusIn>", self._on_formula_focus_in)
        self.formula_entry.bind("<Return>", self._apply_formula_from_bar)

        # Table frame
        self.table_frame = tk.Frame(self, bg="white")
        self.table_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 5))
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)

        # Status bar
        self.bottom_status_frame = CTkFrame(self, fg_color="#E5E7EB", corner_radius=4)
        self.bottom_status_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 10))

        zoom_frame = CTkFrame(self.bottom_status_frame, fg_color="transparent")
        zoom_frame.pack(side="right", padx=5)
        CTkButton(zoom_frame, text="🔍−", width=32, height=24, fg_color="#6B7280", hover_color="#4B5563",
                  font=CTkFont(size=12), command=lambda: self._zoom(-1)).pack(side="left", padx=2)
        self.zoom_label = CTkLabel(zoom_frame, text="100%", font=CTkFont(size=12), text_color="#6B7280", width=40)
        self.zoom_label.pack(side="left")
        CTkButton(zoom_frame, text="🔍+", width=32, height=24, fg_color="#6B7280", hover_color="#4B5563",
                  font=CTkFont(size=12), command=lambda: self._zoom(1)).pack(side="left", padx=2)

        self.save_status_label = CTkLabel(self.bottom_status_frame,
                                          text="✅ พร้อมใช้งาน (บันทึกอัตโนมัติ)",
                                          font=CTkFont(size=13), text_color="gray50")
        self.save_status_label.pack(side="left", padx=20, pady=4)

        self.quick_calc_label = CTkLabel(self.bottom_status_frame, text="",
                                         font=CTkFont(size=14, weight="bold"), text_color="#059669")
        self.quick_calc_label.pack(side="right", padx=20, pady=4)

        if HAS_TKSHEET:
            self._build_tksheet(self.table_frame)
            self.after(200, self._load_from_db)
        else:
            tk.Label(self.table_frame, text="⚠️ กรุณาติดตั้ง tksheet", fg="red", bg="white").pack(expand=True)

    def _capture_click_pos(self, event=None):
        if event:
            self._last_click_x = event.x_root
            self._last_click_y = event.y_root

    # ================================================================== #
    def _lighten_color(self, hex_color, amount=0.85):
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            r = int(r + (255 - r) * amount)
            g = int(g + (255 - g) * amount)
            b = int(b + (255 - b) * amount)
            return f'#{r:02x}{g:02x}{b:02x}'
        except Exception:
            return "#F3F4F6"

    def _load_user_settings(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    "SELECT setting_value FROM user_settings WHERE user_name = %s AND setting_key = 'benchmark_table'",
                    (self.current_user,))
                result = cursor.fetchone()
                if result and result[0]:
                    raw = result[0]
                    # psycopg2 อาจ return เป็น str (TEXT column) หรือ dict (JSONB column)
                    settings = json.loads(raw) if isinstance(raw, str) else raw
                    self.hidden_cols_list = settings.get("hidden_cols", [])
                    self.custom_header_colors = settings.get("header_colors", {})
                    self.cell_highlights_cache = settings.get("cell_highlights", {})
                    
                    # Merge saved widths on top of defaults (ไม่ reset เพื่อ preserve defaults)
                    saved_widths = settings.get("col_widths", {})
                    if saved_widths:
                        for k, v in saved_widths.items():
                            val = int(v) if v else 0
                            # กัน 0/1px ที่เกิดจาก bug เท่านั้น — ไม่รีเซ็ตคอลัมที่ user บีบแคบเอง
                            if val >= 8:
                                self.col_widths_cache[int(k)] = val
                            
                    saved_freeze = settings.get("frozen_col_count", 0)
                    if saved_freeze > 0:
                        self.frozen_col_count = saved_freeze
                        
                    saved_zoom = settings.get("zoom_level", 11)
                    if saved_zoom != 11:
                        self.zoom_level = saved_zoom
                        new_row_height = int(30 * (self.zoom_level / 11.0))
                        new_header_height = int(35 * (self.zoom_level / 11.0))
                        self.sheet.set_options(
                            font=("Tahoma", self.zoom_level, "normal"),
                            header_font=("Tahoma", self.zoom_level, "bold"),
                            row_height=new_row_height,
                            header_height=new_header_height,
                            auto_resize_columns=False,   # ✅ มีอยู่แล้ว — ตรวจให้แน่ใจ
                            auto_resize_row_index=False  # ✅ เพิ่มตรงนี้
                        )
                        if hasattr(self, 'zoom_label'):
                            pct = int((self.zoom_level / 11) * 100)
                            self.zoom_label.configure(text=f"{pct}%")
                            
        except Exception as e:
            print(f"Error loading settings: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _apply_saved_highlights(self):
        """สาดสีไฮไลท์ที่โหลดมาจาก DB คืนลงในตาราง"""
        if not hasattr(self, 'cell_highlights_cache') or not self.cell_highlights_cache:
            return
        
        for key, color in self.cell_highlights_cache.items():
            try:
                side, pos = key.split(':')
                r, c = map(int, pos.split(','))
                # กัน list/tuple จาก tksheet version เก่าที่บันทึกไว้
                if isinstance(color, (list, tuple)):
                    color = color[0] if color else None
                if not isinstance(color, str):
                    continue
                if side == 'main':
                    self.sheet.highlight_cells(row=r, column=c, bg=color, redraw=False)
                elif side == 'frozen' and self.sheet_frozen:
                    self.sheet_frozen.highlight_cells(row=r, column=c, bg=color, redraw=False)
            except Exception: pass
        
        self.sheet.redraw()
        if self.sheet_frozen: self.sheet_frozen.redraw()

    def _save_user_settings(self):
        col_widths_to_save = {str(k): v for k, v in self.col_widths_cache.items() if v and v > 0}

        # 🔹 เพิ่มส่วนนี้: ดึงรหัสสีไฮไลท์จากตาราง (ต้องแปลง tuple เป็น string เพราะ JSON รับ tuple ไม่ได้)
        highlights_to_save = {}
        try:
            def _extract_color(c):
                # tksheet version ใหม่ return เป็น list/tuple [bg, fg, ...]
                if isinstance(c, (list, tuple)):
                    c = c[0] if c else None
                return c if isinstance(c, str) and c.startswith("#") else None

            for (r, c), color in self.sheet.get_highlighted_cells().items():
                col_str = _extract_color(color)
                if col_str:
                    highlights_to_save[f"main:{r},{c}"] = col_str
            if self.sheet_frozen:
                for (r, c), color in self.sheet_frozen.get_highlighted_cells().items():
                    col_str = _extract_color(color)
                    if col_str:
                        highlights_to_save[f"frozen:{r},{c}"] = col_str
        except Exception: pass

        settings = {
            "hidden_cols": self.hidden_cols_list,
            "header_colors": self.custom_header_colors,
            "col_widths": col_widths_to_save,
            "frozen_col_count": self.frozen_col_count,
            "zoom_level": self.zoom_level,
            "cell_highlights": highlights_to_save, # 🔹 เพิ่มเข้าไปใน JSON
        }
        settings_json = json.dumps(settings)
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO user_settings (user_name, setting_key, setting_value)
                    VALUES (%s, 'benchmark_table', %s)
                    ON CONFLICT (user_name, setting_key)
                    DO UPDATE SET setting_value = EXCLUDED.setting_value
                """, (self.current_user, settings_json))
            conn.commit()
        except Exception as e:
            print(f"Error saving settings: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    def _load_dropdown_data(self):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # 1. ดึงข้อมูลรหัส Sale จากฐานข้อมูล
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Sale' AND status = 'Active'")
                raw_sales = [row[0] for row in cursor.fetchall() if row[0]]

                # ==============================================================
                # 📌 จัดการซ่อน เปลี่ยนชื่อ และจัดเรียงหมวดหมู่
                # ==============================================================
                hide_list = {"s", "charita-ct", "vow-p"}
                
                rename_map = {
                    "jiraporn": "JN-IN-JIRAPORN",
                    "sale center": "CT-Sale Center",
                    "piyawan": "AM-IN-PIYAWAN",
                    "ilada": "ID-IN-ILADA",
                    "vow-s": "TG-IN-Tharinya",
                    "bunnycee": "KR-IN-Kanokporn",
                    "lethai": "LT-FL-LETHAI",
                    "wachira": "WA-FL-WACHIRA"
                }

                processed_sales = []
                for sale in raw_sales:
                    sale_lower = sale.strip().lower()
                    
                    if sale_lower in hide_list:
                        continue
                        
                    if sale_lower in rename_map:
                        processed_sales.append(rename_map[sale_lower])
                    else:
                        processed_sales.append(sale.strip())
                
                # ลบตัวซ้ำ (เผื่อกรณีเปลี่ยนชื่อแล้วไปซ้ำ)
                unique_sales = list(dict.fromkeys(processed_sales))

                # แอบเติม Stock เข้าไปในคิวเอาไว้ก่อน
                if "Stock" not in unique_sales:
                    unique_sales.append("Stock")

                # 📌 สร้างกฎการให้คะแนนลำดับใหม่ (อัปเดตตามที่ขอ)
                def get_sort_key(name):
                    name_upper = name.upper()
                    if "-IN-" in name_upper:
                        return 1      # ลำดับ 1: Inbound (มาก่อน)
                    elif "CT-" in name_upper or "-CT" in name_upper:
                        return 2      # ลำดับ 2: Center
                    elif name_upper == "STOCK":
                        return 3      # ลำดับ 3: Stock
                    elif "-FL-" in name_upper:
                        return 4      # ลำดับ 4: Freelance
                    else:
                        return 5      # ลำดับ 5: อื่นๆ (ถ้ามีหลงมา จะอยู่ล่างสุด)

                # สั่งเรียงลำดับ โดยดูจากกลุ่มคะแนนที่ตั้งไว้ และเรียงตามตัวอักษร (A-Z) ภายในกลุ่มนั้นๆ ด้วย
                unique_sales.sort(key=lambda x: (get_sort_key(x), x))
                
                self.sales_list = unique_sales
                # ==============================================================

                # 2. ดึงข้อมูล Supplier
                cursor.execute("SELECT supplier_name, supplier_code FROM suppliers")
                for row in cursor.fetchall():
                    if row[0]:
                        self.supplier_list.append(row[0])
                        self.supplier_code_map[row[0]] = row[1] or ""

                # 3. ดึงข้อมูล Product
                cursor.execute("SELECT product_name, product_code, category FROM products")
                for row in cursor.fetchall():
                    if row[0]:
                        self.product_list.append(row[0])
                        self.product_sku_map[row[0]] = row[1] or ""
                        self.product_category_map[row[0]] = row[2] or ""
        except Exception as e:
            print(f"Error loading dropdown data: {e}")
        finally:
            if conn: self.app_container.release_connection(conn)

    # ================================================================== #
    def _build_tksheet(self, parent):
        self.sheet = Sheet(
            parent,
            headers=self.columns,
            data=[[""] * len(self.columns) for _ in range(1000)],
            theme="light blue",
            row_height=30,
            header_height=35,
            font=("Tahoma", 11, "normal"),
            header_font=("Tahoma", 11, "bold"),
            show_row_index=True,
            row_index_width=40,
            column_width=120,
            empty_horizontal=0,
            empty_vertical=0,
        )
        self.sheet.grid(row=0, column=0, sticky="nsew")
        self._load_user_settings()

        self._rebind_sheet()
        self._rebind_frozen_sheet()
        self._apply_formatting(col_offset=0)
        self.after(100, lambda: self._setup_header_filters(0))

        if self.hidden_cols_list:
            self._apply_hidden_columns()

        if self.frozen_col_count > 0:
            # ถ้ามี frozen ให้ rebuild ก่อน แล้ว restore widths จะถูกเรียกข้างใน
            self.after(100, self._rebuild_frozen_layout)
        else:
            # ไม่มี frozen — restore หลัง widget render เสร็จสมบูรณ์
            self.after(300, self._restore_col_widths)

        self.after(5000, self._start_dropdown_refresh_timer)

    # ================================================================== #
    # FREEZE / UNFREEZE
    # ================================================================== #
    def _freeze_selected_columns(self):
        if not HAS_TKSHEET:
            return
        display_last = None  # 0-based display index of rightmost selected column in main sheet
        try:
            selected_cols = self.sheet.get_selected_columns()
            if selected_cols:
                display_last = max(selected_cols)
        except Exception:
            pass
        if display_last is None:
            try:
                selected_cells = self.sheet.get_selected_cells()
                if selected_cells:
                    display_last = max(c for _, c in selected_cells)
            except Exception:
                pass
        if display_last is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเซลล์หรือหัวคอลัมน์ที่ต้องการตรึงก่อน", parent=self)
            return

        # Convert display index → real column index (accounts for hidden columns)
        col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
        hidden_set = set(getattr(self, "hidden_cols_list", []))
        visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
        if display_last < len(visible_real):
            real_last = visible_real[display_last]
        else:
            real_last = col_offset + display_last  # fallback (no hidden cols)
        actual_freeze = real_last + 1  # freeze UP TO AND INCLUDING this column
        actual_freeze = min(actual_freeze, len(self.columns) - 1)
        self.frozen_col_count = actual_freeze
        
        self._save_user_settings() # <--- ⚠️ บันทึกการตรึงคอลัมน์ลง DB ทันที!
        self._rebuild_frozen_layout()
        
        col_name = self.columns[actual_freeze - 1] if actual_freeze <= len(self.columns) else "?"
        self.save_status_label.configure(
            text=f"📌 ตรึง {actual_freeze} คอลัมน์แรก (ถึง: {col_name})",
            text_color="#0891B2"
        )

    def _unfreeze_columns(self):
        if not HAS_TKSHEET:
            return
        self.frozen_col_count = 0
        self._save_user_settings() # <--- ⚠️ บันทึกการยกเลิกตรึงลง DB ทันที!
        self._rebuild_frozen_layout()
        self.save_status_label.configure(text="✅ ยกเลิกตรึงแล้ว", text_color="#16A34A")

    def _rebuild_frozen_layout(self):
        try:
            if self.sheet_frozen is not None:
                frozen_data = self.sheet_frozen.get_sheet_data()
                scroll_data = self.sheet.get_sheet_data()
                n_rows = max(len(frozen_data), len(scroll_data))
                current_data = []
                for i in range(n_rows):
                    left = frozen_data[i] if i < len(frozen_data) else [""] * self.frozen_col_count
                    right = scroll_data[i] if i < len(scroll_data) else [""] * (len(self.columns) - self.frozen_col_count)
                    current_data.append(left + right)
            else:
                current_data = self.sheet.get_sheet_data()
        except Exception:
            current_data = [[""] * len(self.columns) for _ in range(20)]

        # ── Snapshot current live widths into col_widths_cache before destroying sheets ──
        # ✅ อัปเดต cache เฉพาะถ้า user เคย resize จริง (ค่าใน sheet ≠ 120 default)
        # ไม่ overwrite ค่าที่ load มาจาก DB ด้วยค่า default 120 จาก sheet ที่ยังไม่ถูก restore
        try:
            hidden_set_snap = set(getattr(self, "hidden_cols_list", []))
            DEFAULT_W = 120
            if self.sheet_frozen is not None:
                for i in range(self.frozen_col_count):
                    try:
                        w = self.sheet_frozen.column_width(i)
                        # อัปเดต cache เฉพาะถ้าค่าจาก sheet ไม่ใช่ default (แสดงว่า user resize แล้ว)
                        if w and w >= 8 and w != DEFAULT_W:
                            self.col_widths_cache[i] = w
                    except Exception: pass
                visible_real_snap = [c for c in range(self.frozen_col_count, len(self.columns)) if c not in hidden_set_snap]
                for di in range(self.sheet.get_total_columns()):
                    if di >= len(visible_real_snap): break
                    try:
                        w = self.sheet.column_width(di)
                        if w and w >= 8 and w != DEFAULT_W:
                            self.col_widths_cache[visible_real_snap[di]] = w
                    except Exception: pass
            else:
                visible_real_snap = [c for c in range(len(self.columns)) if c not in hidden_set_snap]
                for di in range(self.sheet.get_total_columns()):
                    if di >= len(visible_real_snap): break
                    try:
                        w = self.sheet.column_width(di)
                        if w and w >= 8 and w != DEFAULT_W:
                            self.col_widths_cache[visible_real_snap[di]] = w
                    except Exception: pass
        except Exception: pass

        # ✅ ใช้ col_widths_cache เป็น source of truth โดยตรง
        # ไม่ดึงจาก sheet เพราะ sheet อาจยังแสดงค่า 120 (default) ถ้า _restore_col_widths ยังไม่ได้ run
        saved_widths = {k: v for k, v in self.col_widths_cache.items() if v and v >= 8}

        try: self.table_frame.unbind("<Configure>")
        except Exception: pass
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        self.sheet_frozen = None
        self.table_frame.grid_rowconfigure(0, weight=1)

        n_freeze = self.frozen_col_count
        frozen_width = 0
        if n_freeze > 0:
            frozen_width = 40
            for i in range(n_freeze):
                w = saved_widths.get(i, 120)
                w = w if w >= 8 else 120 # กัน 0/1px bug
                frozen_width += w
            frozen_width += 4
        self._current_frozen_width = frozen_width  # store ไว้ให้ _on_table_resize อ่านค่าล่าสุดเสมอ

        if n_freeze == 0:
            self.table_frame.grid_columnconfigure(0, weight=1, minsize=0)
            self.table_frame.grid_columnconfigure(1, weight=0, minsize=0)
            
            # --- ให้เพิ่มโค้ดคำนวณความสูง 2 บรรทัดนี้ ---
            current_row_h = int(30 * (self.zoom_level / 11.0))
            current_header_h = int(35 * (self.zoom_level / 11.0))

            self.sheet = Sheet(
                self.table_frame,
                headers=self.columns,
                data=current_data,
                theme="light blue",
                row_height=current_row_h,       # เปลี่ยนจาก 30 เป็น current_row_h
                header_height=current_header_h, # เปลี่ยนจาก 35 เป็น current_header_h
                font=("Tahoma", self.zoom_level, "normal"),
                header_font=("Tahoma", self.zoom_level, "bold"),
                show_row_index=True,
                row_index_width=40,
                column_width=120,
                empty_horizontal=0,
                empty_vertical=0,
            )
            self.sheet.grid(row=0, column=0, sticky="nsew")
        else:
            self.table_frame.grid_columnconfigure(0, weight=0, minsize=0)
            self.table_frame.grid_columnconfigure(1, weight=0, minsize=0)
            frozen_cols = self.columns[:n_freeze]
            scroll_cols = self.columns[n_freeze:]
            frozen_data_split = [row[:n_freeze] for row in current_data]
            scroll_data_split = [row[n_freeze:] for row in current_data]
            
            current_row_h = int(30 * (self.zoom_level / 11.0))
            current_header_h = int(35 * (self.zoom_level / 11.0))

            self.sheet_frozen = Sheet(
                self.table_frame,
                headers=frozen_cols,
                data=frozen_data_split,
                theme="light blue",
                row_height=current_row_h,      
                header_height=current_header_h,
                font=("Tahoma", self.zoom_level, "normal"),
                header_font=("Tahoma", self.zoom_level, "bold"),
                show_row_index=True,
                row_index_width=40,
                column_width=120,
                empty_horizontal=0,
                empty_vertical=0,
            )
            try:
                self.sheet_frozen.hide(canvas="x_scrollbar")
                self.sheet_frozen.hide(canvas="y_scrollbar")
                try: self.sheet_frozen.yscrollbar.grid_remove()
                except Exception: pass
                try: self.sheet_frozen.xscrollbar.grid_remove()
                except Exception: pass
            except Exception: pass
            
            self.sheet_frozen.place(x=0, y=0, width=frozen_width, relheight=1.0)
            self.after(10, lambda: self._hide_frozen_scrollbars())

            self.sheet = Sheet(
                self.table_frame,
                headers=scroll_cols,
                data=scroll_data_split,
                theme="light blue",
                row_height=current_row_h,      
                header_height=current_header_h,
                font=("Tahoma", self.zoom_level, "normal"),
                header_font=("Tahoma", self.zoom_level, "bold"),
                show_row_index=False,
                row_index_width=0,
                column_width=120,
                empty_horizontal=0,
                empty_vertical=0,
            )
            self.sheet.place(x=frozen_width, y=0, relwidth=1.0, relheight=1.0, width=-frozen_width)

            _fw = frozen_width
            def _on_table_resize(event):
                try:
                    fw = getattr(self, '_current_frozen_width', _fw)
                    self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                except Exception: pass
            self.table_frame.bind("<Configure>", _on_table_resize)
            self._sync_vertical_scroll()

        self._rebind_sheet()
        self._rebind_frozen_sheet()
        
        self.after(10, lambda nf=n_freeze: self._apply_formatting(col_offset=nf))
        if self.sheet_frozen is not None:
            self.after(30, lambda nf=n_freeze: self._apply_formatting_frozen(nf))
        self.after(50, lambda: self._setup_header_filters(n_freeze))

        def _restore_after_build():
            def _do_restore():
                self._restore_col_widths()
                if hasattr(self, '_apply_hidden_columns'):
                    self._apply_hidden_columns()

                try:
                    rh = int(30 * (self.zoom_level / 11.0))
                    # เช็คว่ามีตารางตรึงไหม ถ้าไม่มีให้ใช้แถวตารางหลักแทน
                    if self.sheet_frozen:
                        total_rows = max(self.sheet.get_total_rows(), self.sheet_frozen.get_total_rows())
                    else:
                        total_rows = self.sheet.get_total_rows()
                        
                    for r in range(total_rows):
                        if self.sheet_frozen:
                            self.sheet_frozen.row_height(r, rh)
                        self.sheet.row_height(r, rh)
                        
                    if self.sheet_frozen:
                        self.sheet_frozen.redraw()
                    self.sheet.redraw()
                except Exception:
                    pass

                self.after(200, self._hide_frozen_scrollbars)
                self.after(600, self._hide_frozen_scrollbars)

                # ── แก้ช่องขาว: เรียก _zoom จริงๆ เลย zoom in แล้ว zoom out ──
                # ทำใน after() ที่ต่อเนื่องกัน user จะไม่เห็นการเปลี่ยนแปลง
                if self.sheet_frozen:
                    self.after(50, lambda: self._zoom(1))
                    self.after(100, lambda: self._zoom(-1))

            self.after(0, _do_restore)

        self.after(150, _restore_after_build)


        
    def _sync_vertical_scroll(self):
        if not getattr(self, "sheet_frozen", None):
            return

        if getattr(self, '_sync_loop_id', None):
            self.after_cancel(self._sync_loop_id)

        self._last_yview_main = -1.0
        self._last_yview_frozen = -1.0

        def _do_sync():
            if not getattr(self, "sheet", None) or not getattr(self, "sheet_frozen", None):
                return
            try:
                if not self.sheet.winfo_exists() or not self.sheet_frozen.winfo_exists():
                    return
            except Exception:
                return

            try:
                try: y_main = self.sheet.get_yview()[0]
                except Exception: y_main = self.sheet.MT.yview()[0]
                
                try: y_frozen = self.sheet_frozen.get_yview()[0]
                except Exception: y_frozen = self.sheet_frozen.MT.yview()[0]

                if abs(y_main - self._last_yview_main) > 0.0001:
                    self._last_yview_main = y_main
                    self._last_yview_frozen = y_main
                    self.sheet_frozen.yview_moveto(y_main)
                    try: self.sheet_frozen.MT.yview_moveto(y_main)
                    except Exception: pass

                elif abs(y_frozen - self._last_yview_frozen) > 0.0001:
                    self._last_yview_frozen = y_frozen
                    self._last_yview_main = y_frozen
                    self.sheet.yview_moveto(y_frozen)
                    try: self.sheet.MT.yview_moveto(y_frozen)
                    except Exception: pass

            except Exception:
                pass

            if getattr(self, "frozen_col_count", 0) > 0:
                self._sync_loop_id = self.after(20, _do_sync)

        self._sync_loop_id = self.after(50, _do_sync)

        def _sync_to_frozen(event=None):
            try:
                if not self.sheet.winfo_exists() or not self.sheet_frozen.winfo_exists(): return
                pos = self.sheet.get_yview()[0]
                self.sheet_frozen.yview_moveto(pos)
                try: self.sheet_frozen.MT.yview_moveto(pos)
                except Exception: pass
                try: self.sheet_frozen.RI.yview_moveto(pos)
                except Exception: pass
            except Exception: pass

        def _frozen_wheel(event=None):
            try:
                units = -3 if event.delta > 0 else 3
                self.sheet.MT.yview_scroll(units, "units")
                try: self.sheet.redraw()
                except Exception: pass
            except Exception: pass
            self.after_idle(_sync_to_frozen)
            return "break"

        def _frozen_wheel_up(event=None):
            try:
                self.sheet.MT.yview_scroll(-3, "units")
                try: self.sheet.redraw()
                except Exception: pass
            except Exception: pass
            self.after_idle(_sync_to_frozen)
            return "break"

        def _frozen_wheel_down(event=None):
            try:
                self.sheet.MT.yview_scroll(3, "units")
                try: self.sheet.redraw()
                except Exception: pass
            except Exception: pass
            self.after_idle(_sync_to_frozen)
            return "break"

        # bind main → frozen
        for widget in [self.sheet, self.sheet.MT]:
            try:
                widget.bind("<MouseWheel>", lambda e: self.after_idle(_sync_to_frozen), add="+")
                widget.bind("<Button-4>",   lambda e: self.after_idle(_sync_to_frozen), add="+")
                widget.bind("<Button-5>",   lambda e: self.after_idle(_sync_to_frozen), add="+")
            except Exception: pass
        try:
            self.sheet.RI.bind("<MouseWheel>", lambda e: self.after_idle(_sync_to_frozen), add="+")
            self.sheet.RI.bind("<Button-4>",   lambda e: self.after_idle(_sync_to_frozen), add="+")
            self.sheet.RI.bind("<Button-5>",   lambda e: self.after_idle(_sync_to_frozen), add="+")
        except Exception: pass

        # bind frozen → main (ใช้ unbind ก่อน)
        frozen_widgets = [self.sheet_frozen, self.sheet_frozen.MT]
        try: frozen_widgets.append(self.sheet_frozen.RI)
        except Exception: pass
        try: frozen_widgets.append(self.sheet_frozen.CH)
        except Exception: pass

        for w in frozen_widgets:
            if not w: continue
            try:
                w.unbind("<MouseWheel>")
                w.unbind("<Button-4>")
                w.unbind("<Button-5>")
            except Exception: pass
            try:
                w.bind("<MouseWheel>", _frozen_wheel)
                w.bind("<Button-4>",   _frozen_wheel_up)
                w.bind("<Button-5>",   _frozen_wheel_down)
            except Exception: pass
    # ================================================================== #
    # FORMATTING
    # ================================================================== #
    def _get_header_styles_map(self):
        return {
            ("#2563EB", "white"): ["วันที่ขอราคา", "Order No.", "Sale Order No.", "QT"],
            ("#BAE6FD", "black"): [
                "หมายเหตุ (ความยาว, OD)", "หมายเหตุ", "จำนวน", "ต้นทุน/เส้น", "น้ำหนัก/เส้น",
                "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (บาท)", "ส่วนลด 1 (%)",
                "ส่วนลด 2 (บาท)", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 1", "ทุน/เส้น หลังส่วนลด 2",
                "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)",
                "ค่าย้าย (ซื้อ)", "ค่าย้าย/เส้น", "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)",
                "ต้นทุนรวม (รวมย้าย)"
            ],
            ("#FDBA74", "black"): ["ผู้ขอราคา", "รหัส Sale", "PRIORITY", "WIN RATE %", "Select", "แบรนด์"],
            ("#6B7280", "white"): ["หมวด", "หมวดหลัก", "หมวดรอง", "หมวดย่อย", "Product SKU.",
                                   "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น",
                                   "ต้นทุนรวม+Markup", "ค่าส่ง (ขาย)", "ค่าส่ง / เส้น"],
            ("#D8B4FE", "black"): ["รายการสินค้า", "ชื่อ Supplier"],
            ("#FCA5A5", "black"): ["Markup Guide (%)"],
            ("#FDE047", "black"): ["สถานะ", "น้ำหนัก/เส้น 2", "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม + Vat."],
            ("#86EFAC", "black"): ["ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น", "ราคาขาย รวม", "Vat. รวม"],
            ("#1F2937", "white"): ["ชื่อ Supplier2", "Sup ID.", "คลังสินค้า ต้นทาง"],
            ("#93C5FD", "black"): ["ปลายทาง", "หมายเหตุ2"]
        }

    def _get_selected_row_idx(self):
        try:
            curr = self.sheet.get_currently_selected()
            if curr:
                return int(curr[0])
        except Exception:
            pass

        if self.sheet_frozen:
            try:
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    return int(curr[0])
            except Exception:
                pass

        try:
            selected_rows = self.sheet.get_selected_rows()
            if selected_rows:
                return min(int(r) for r in selected_rows)
        except Exception:
            pass

        return None

    def _swap_rows(self, idx_a, idx_b):
        """สลับข้อมูล 2 บรรทัดทั้ง main และ frozen"""
        try:
            # main sheet
            row_a = list(self.sheet.get_row_data(idx_a))
            row_b = list(self.sheet.get_row_data(idx_b))
            for c, val in enumerate(row_b):
                self.sheet.set_cell_data(idx_a, c, val, redraw=False)
            for c, val in enumerate(row_a):
                self.sheet.set_cell_data(idx_b, c, val, redraw=False)
        except Exception as e:
            print(f"_swap_rows main error: {e}")

        if self.sheet_frozen:
            try:
                frow_a = list(self.sheet_frozen.get_row_data(idx_a))
                frow_b = list(self.sheet_frozen.get_row_data(idx_b))
                for c, val in enumerate(frow_b):
                    self.sheet_frozen.set_cell_data(idx_a, c, val, redraw=False)
                for c, val in enumerate(frow_a):
                    self.sheet_frozen.set_cell_data(idx_b, c, val, redraw=False)
            except Exception as e:
                print(f"_swap_rows frozen error: {e}")

        self.sheet.redraw()
        if self.sheet_frozen:
            self.sheet_frozen.redraw()

    def _move_row_up(self):
        if not HAS_TKSHEET:
            return
        idx = self._get_selected_row_idx()
        if idx is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการย้ายก่อน", parent=self)
            return
        if idx <= 0:
            return
        self._swap_rows(idx, idx - 1)
        # ย้าย selection ขึ้นด้วย
        self.sheet.select_cell(idx - 1, 0)
        self.sheet.see(idx - 1, 0)
        if self.sheet_frozen:
            self.sheet_frozen.select_cell(idx - 1, 0)
        # debounce save
        if self.auto_save_job_id:
            self.after_cancel(self.auto_save_job_id)
        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

    def _move_row_down(self):
        if not HAS_TKSHEET:
            return
        idx = self._get_selected_row_idx()
        if idx is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการย้ายก่อน", parent=self)
            return
        total_rows = self.sheet.get_total_rows()
        if idx >= total_rows - 1:
            return
        self._swap_rows(idx, idx + 1)
        self.sheet.select_cell(idx + 1, 0)
        self.sheet.see(idx + 1, 0)
        if self.sheet_frozen:
            self.sheet_frozen.select_cell(idx + 1, 0)
        if self.auto_save_job_id:
            self.after_cancel(self.auto_save_job_id)
        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

    def _get_auto_cols(self):
        return [
            "Product SKU.", "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม",
            "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1", "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2",
            "ต้นทุน/กก. (ไม่รวมย้าย)", "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย/เส้น",
            "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)", "Markup/กก.",
            "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup",
            "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น", "Vat. / เส้น",
            "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat.",
            "ชื่อ Supplier2", "Sup ID."
        ]


    def _real_to_disp(self, real_idx):
        """แปลง real column index -> (sheet_side, data_idx) โดยคำนึงคอลัมน์ซ่อน
        คืน ('frozen', data_idx) หรือ ('main', data_idx) หรือ (None, None) ถ้าซ่อน
        หมายเหตุ: data_idx คือ index ที่ใช้กับ set/get_cell_data ของ tksheet
        (ต่างจาก display_idx ที่ใช้สำหรับ UI เท่านั้น)
        """
        col_offset = self.frozen_col_count if self.sheet_frozen else 0
        hidden_set = set(getattr(self, "hidden_cols_list", []))
        if real_idx in hidden_set:
            return None, None
        if real_idx < col_offset:
            # frozen sheet: data index = real_idx (frozen sheet มี col 0..frozen_col_count-1)
            return 'frozen', real_idx
        else:
            # main sheet: data index = real_idx - col_offset
            return 'main', real_idx - col_offset

    def _col_to_disp(self, col_name):
        """แปลง col_name -> (sheet_side, display_idx) โดยคำนึงคอลัมน์ซ่อน"""
        try:
            return self._real_to_disp(self.columns.index(col_name))
        except ValueError:
            return None, None

    def _sheet_get(self, row_idx, col_name):
        """อ่านค่า cell จาก col_name (รองรับ hidden cols + frozen split)"""
        side, disp = self._col_to_disp(col_name)
        if side == 'frozen':
            return self.sheet_frozen.get_cell_data(row_idx, disp) if self.sheet_frozen else None
        if side == 'main':
            return self.sheet.get_cell_data(row_idx, disp)
        return None

    def _sheet_set(self, row_idx, col_name, value, redraw=False):
        """เขียนค่า cell ไปยัง col_name (รองรับ hidden cols + frozen split)"""
        side, disp = self._col_to_disp(col_name)
        if side == 'frozen' and self.sheet_frozen:
            self.sheet_frozen.set_cell_data(row_idx, disp, value, redraw=redraw)
        elif side == 'main':
            self.sheet.set_cell_data(row_idx, disp, value, redraw=redraw)

    def _apply_formatting(self, col_offset=0):
        auto_cols_names = self._get_auto_cols()
        header_styles_map = self._get_header_styles_map()
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles_map.items() for c in cols}
        if hasattr(self, 'custom_header_colors'):
            for col_name, bg_color in self.custom_header_colors.items():
                if col_name in self.columns:
                    col_to_style[col_name] = (bg_color, "black")

        def get_data_idx(col_name):
            # ใช้ Data Index (ไม่นับ hidden cols) เพื่อให้ readonly_columns ถูกต้องเสมอ
            try:
                real_idx = self.columns.index(col_name)
            except ValueError:
                return None
            if real_idx < col_offset:
                return None
            return real_idx - col_offset

        auto_data_indices = []
        readonly_cols = self._get_auto_cols()
        for c in readonly_cols:
            idx = get_data_idx(c)
            if idx is not None:
                auto_data_indices.append(idx)
        if auto_data_indices:
            self.sheet.readonly_columns(columns=auto_data_indices, readonly=True)
        # ✅ ปลดล็อค Markup Guide (%) และ WIN RATE % เสมอ (ไม่ขึ้นกับ hidden cols)
        for _editable_col in ["Markup Guide (%)", "WIN RATE %"]:
            try:
                _ei = get_data_idx(_editable_col)
                if _ei is not None:
                    self.sheet.readonly_columns(columns=[_ei], readonly=False)
            except Exception:
                pass

        display_cols = self.columns[col_offset:]
        for di, col in enumerate(display_cols):
            h_bg, h_fg = col_to_style.get(col, ("#E5E7EB", "#111827"))
            if hasattr(self, 'custom_header_colors') and col in self.custom_header_colors:
                b_bg = self._lighten_color(self.custom_header_colors[col], 0.85)
                b_fg = "black"
            elif col in auto_cols_names:
                b_bg = "#F3F4F6"
                b_fg = "#111827"
            else:
                b_bg = "white"
                b_fg = "black"
            try:
                self.sheet.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                self.sheet.highlight_columns(columns=[di], bg=b_bg, fg=b_fg, highlight_header=False)
            except Exception:
                try:
                    self.sheet.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                except Exception:
                    pass
        
        self.sheet.set_options(
            grid_color="#000000", outline_color="#000000", table_bg="white", table_fg="black",
            table_grid_fg="#000000", header_bg="#D1D5DB", header_fg="#111827", header_grid_fg="#000000",
            header_selected_cells_bg="#9CA3AF", row_index_bg="#F3F4F6", row_index_fg="#111827",
            row_index_grid_fg="#000000", selected_cells_border_color="#3B82F6",
            table_selected_cells_border_color="#3B82F6",
            auto_resize_columns=False, auto_resize_row_index=False,
            text_editor_autoresize=False,
            dropdown_font=("Tahoma", 11, "normal"),
        )

    def _apply_formatting_frozen(self, n_freeze):
        if not self.sheet_frozen:
            return
        auto_cols_names = self._get_auto_cols()
        header_styles_map = self._get_header_styles_map()
        col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles_map.items() for c in cols}
        if hasattr(self, 'custom_header_colors'):
            for col_name, bg_color in self.custom_header_colors.items():
                col_to_style[col_name] = (bg_color, "black")

        frozen_col_names = self.columns[:n_freeze]
        for di, col in enumerate(frozen_col_names):
            h_bg, h_fg = col_to_style.get(col, ("#E5E7EB", "#111827"))
            if hasattr(self, 'custom_header_colors') and col in self.custom_header_colors:
                b_bg = self._lighten_color(self.custom_header_colors[col], 0.85)
                b_fg = "black"
            elif col in auto_cols_names:
                b_bg = "#F3F4F6"
                b_fg = "#111827"
            else:
                b_bg = "white"
                b_fg = "black"
            try:
                self.sheet_frozen.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                self.sheet_frozen.highlight_columns(columns=[di], bg=b_bg, fg=b_fg, highlight_header=False)
            except Exception:
                try:
                    self.sheet_frozen.highlight_cells(row=0, column=di, bg=h_bg, fg=h_fg, canvas="header")
                except Exception:
                    pass

        # readonly สำหรับ auto cols ที่อยู่ในฝั่ง frozen
        auto_frozen_indices = []
        for col_name in auto_cols_names:
            if col_name in frozen_col_names:
                idx = frozen_col_names.index(col_name)
                auto_frozen_indices.append(idx)
        if auto_frozen_indices:
            try:
                self.sheet_frozen.readonly_columns(columns=auto_frozen_indices, readonly=True)
            except Exception:
                pass

        # ✅ ปลดล็อค Markup Guide (%) และ WIN RATE % ฝั่ง frozen ด้วย
        for _editable_col in ["Markup Guide (%)", "WIN RATE %"]:
            if _editable_col in frozen_col_names:
                try:
                    idx = frozen_col_names.index(_editable_col)
                    self.sheet_frozen.readonly_columns(columns=[idx], readonly=False)
                except Exception:
                    pass

        self.sheet_frozen.set_options(
            grid_color="#000000", outline_color="#000000", table_bg="white", table_fg="black",
            table_grid_fg="#000000", header_bg="#D1D5DB", header_fg="#111827", header_grid_fg="#000000",
            row_index_bg="#F3F4F6", row_index_fg="#111827", row_index_grid_fg="#000000",
            auto_resize_columns=False, auto_resize_row_index=False,
            text_editor_autoresize=False,
            dropdown_font=("Tahoma", 11, "normal"),
        )

    def _setup_header_filters(self, col_offset=0):
        self._filter_col_names = {
            "วันที่ขอราคา", "Order No.", "Sale Order No.", "ชื่อ Supplier", "รายการสินค้า",
            "สถานะ", "PRIORITY", "รหัส Sale", "หมวด", "QT", "Select",
            "ส่วนลด 1 (บาท)", "ส่วนลด 1 (%)", "ส่วนลด 2 (บาท)", "ส่วนลด 2 (%)"
        }
        self._header_col_offset = col_offset

        if not getattr(self.sheet, "_filter_bound", False):
            try:
                self.sheet.CH.bind("<ButtonRelease-1>", 
                    lambda e: self.after(10, lambda ev=e: self._on_header_click(ev)), add="+")
                self.sheet._filter_bound = True
            except Exception: pass

        if self.sheet_frozen:
            if not getattr(self.sheet_frozen, "_filter_bound", False):
                try:
                    self.sheet_frozen.CH.bind("<ButtonRelease-1>", 
                        lambda e: self.after(10, lambda ev=e: self._on_frozen_header_click(ev)), add="+")
                    self.sheet_frozen._filter_bound = True
                except Exception: pass

        # วาดลูกศรหลัง bind เสร็จ
        self.after(100, self._draw_filter_arrows)
        
        # วาดใหม่ทุกครั้งที่ scroll หรือ resize
        try:
            self.sheet.CH.bind("<Configure>", lambda e: self.after(50, self._draw_filter_arrows), add="+")
            self.sheet.bind("<Configure>", lambda e: self.after(50, self._draw_filter_arrows), add="+")
        except Exception: pass

    def _get_col_from_header_click(self, event, sheet_widget, col_offset=0):
        """คำนวณ display_col จากพิกัดที่แท้จริงของ Canvas (แก้บั๊คเลื่อน Scrollbar แนวนอน)"""
        try:
            # 1. ให้ tksheet คำนวณจาก Event โดยตรง (รองรับการ Scroll)
            col = sheet_widget.identify_col(event)
            if col is not None: return int(col)
        except Exception: pass
        
        try:
            # 2. แปลงพิกัดเมาส์ (event.x) เป็นพิกัดตารางจริงๆ (canvasx)
            cx = sheet_widget.CH.canvasx(event.x)
            col = sheet_widget.CH.identify_col(x=cx, allow_end=False)
            if col is not None: return int(col)
        except Exception: pass
        
        try:
            # 3. คำนวณความกว้างคอลัมน์แบบ Manual
            cx = sheet_widget.CH.canvasx(event.x)
            try: row_idx_w = sheet_widget.RI.winfo_width()
            except Exception: row_idx_w = 40 if col_offset == 0 else 0
            
            cx -= row_idx_w
            if cx < 0: return None
            
            accum = 0
            for c in range(sheet_widget.get_total_columns()):
                try: w = sheet_widget.column_width(c)
                except Exception: w = 120
                accum += w
                if cx < accum: return c
        except Exception: pass
        return None

    def _on_header_click(self, event=None):
        try:
            # ดักการลากขยายคอลัมน์ (ปรับเป็น > 3 ป้องกันเมาส์สั่นตอนคลิก)
            if event and hasattr(self, '_header_press_x'):
                if abs(event.x_root - self._header_press_x) > 3:
                    return 

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            display_col = self._get_col_from_header_click(event, self.sheet, col_offset)
            if display_col is None: return

            # Map display index → real column index (accounts for hidden columns)
            hidden_set = set(getattr(self, "hidden_cols_list", []))
            visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            if display_col >= len(visible_real): return
            real_col = visible_real[display_col]
            if real_col >= len(self.columns): return
            col_name = self.columns[real_col]
            
            if col_name not in self._filter_col_names: return

            if self._active_filter_popup and not getattr(self._active_filter_popup, '_destroyed', True):
                self._active_filter_popup.safe_destroy()
                self._active_filter_popup = None
                return
            self._show_header_filter_popup(col_name, display_col, event, is_frozen=False)
        except Exception: pass

    def _on_frozen_header_click(self, event=None):
        try:
            if event and hasattr(self, '_header_press_x'):
                if abs(event.x_root - self._header_press_x) > 3:
                    return

            display_col = self._get_col_from_header_click(event, self.sheet_frozen, 0)
            if display_col is None: return
            if display_col >= len(self.columns): return
            
            col_name = self.columns[display_col]
            if col_name not in self._filter_col_names: return

            if self._active_filter_popup and not getattr(self._active_filter_popup, '_destroyed', True):
                self._active_filter_popup.safe_destroy()
                self._active_filter_popup = None
                return
            self._show_header_filter_popup(col_name, display_col, event, is_frozen=True)
        except Exception: pass

    def _show_header_filter_popup(self, col_name, display_col, event, is_frozen=False):
        try:
            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            
            # ⚠️ อัปเดต: ให้ดึง "ช่องว่าง" มาทำเป็นตัวเลือก "(ว่าง)" ด้วย
            raw_vals = set()
            for r in range(target_sheet.get_total_rows()):
                val = str(target_sheet.get_cell_data(r, display_col) or "").strip()
                raw_vals.add(val if val else "(ว่าง)")
                
            all_values = sorted(list(raw_vals))
            if not all_values:
                all_values = ["(ว่าง)"]

            current_filter = self._header_filter_values.get(col_name)
            popup = tk.Toplevel(self)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)
            popup.configure(bg="#D1D5DB", padx=1, pady=1)
            popup._destroyed = False

            def safe_destroy():
                if not popup._destroyed:
                    popup._destroyed = True
                    try: popup.destroy()
                    except: pass
                self._active_filter_popup = None

            popup.safe_destroy = safe_destroy

            inner = tk.Frame(popup, bg="white")
            inner.pack(fill="both", expand=True)

            # Sort buttons
            sort_bar = tk.Frame(inner, bg="#F9FAFB")
            sort_bar.pack(fill="x", padx=4, pady=(4, 2))
            tk.Button(sort_bar, text="↑ เรียง A → Z", font=("Tahoma", 10),
                    relief="flat", bg="#F9FAFB", cursor="hand2", anchor="w",
                    command=lambda: [self._sort_by_col(col_name, True), safe_destroy()]
                    ).pack(fill="x", pady=1)
            tk.Button(sort_bar, text="↓ เรียง Z → A", font=("Tahoma", 10),
                    relief="flat", bg="#F9FAFB", cursor="hand2", anchor="w",
                    command=lambda: [self._sort_by_col(col_name, False), safe_destroy()]
                    ).pack(fill="x", pady=1)
            tk.Frame(inner, bg="#E5E7EB", height=1).pack(fill="x", padx=4, pady=2)

            # Search
            search_var = tk.StringVar()
            search_entry = tk.Entry(inner, textvariable=search_var,
                                    font=("Tahoma", 10), relief="flat", bd=5,
                                    highlightthickness=1, highlightcolor="#3B82F6")
            search_entry.pack(fill="x", padx=4, pady=2)
            search_entry.insert(0, "🔍 ค้นหา...")
            search_entry.bind("<FocusIn>", lambda e: search_entry.delete(0, "end")
                            if search_entry.get().startswith("🔍") else None)
            def _hf_thai_cb(e, _se=search_entry):
                if e.keycode == 86: _se.event_generate("<<Paste>>"); return "break"
                if e.keycode == 67: _se.event_generate("<<Copy>>"); return "break"
            search_entry.bind("<Control-KeyPress>", _hf_thai_cb, add="+")

            # Checkbox list
            check_frame = tk.Frame(inner, bg="white")
            check_frame.pack(fill="both", expand=True, padx=4)
            canvas_c = tk.Canvas(check_frame, bg="white", highlightthickness=0, height=200)
            sb_c = tk.Scrollbar(check_frame, orient="vertical", command=canvas_c.yview)
            canvas_c.configure(yscrollcommand=sb_c.set)
            sb_c.pack(side="right", fill="y")
            canvas_c.pack(side="left", fill="both", expand=True)
            check_inner = tk.Frame(canvas_c, bg="white")
            canvas_c.create_window((0, 0), window=check_inner, anchor="nw")

            check_vars = {}
            check_widgets = []
            active_set = set(current_filter) if current_filter is not None else set(all_values)

            # สร้าง widget ครั้งเดียวตอนเปิด แล้วใช้ pack/pack_forget ซ่อน/แสดง (เร็วกว่า destroy มาก)
            var_all = tk.BooleanVar()
            cb_all = tk.Checkbutton(check_inner, text="(Select All)", variable=var_all,
                        font=("Tahoma", 10), bg="white", anchor="w", relief="flat",
                        activebackground="#EFF6FF")
            cb_all.pack(fill="x", pady=1)
            _row_widgets = []  # [(value, BooleanVar, Checkbutton)]
            for v in all_values:
                var = tk.BooleanVar(value=v in active_set)
                check_vars[v] = var
                cb = tk.Checkbutton(check_inner, text=v, variable=var,
                            font=("Tahoma", 10), bg="white", anchor="w", relief="flat",
                            activebackground="#EFF6FF")
                cb.pack(fill="x", pady=1)
                _row_widgets.append((v, var, cb))

            def build_checklist(term=""):
                check_widgets.clear()
                t = term.lower() if (term and not term.startswith("🔍")) else ""
                filtered = []
                for v, var, cb in _row_widgets:
                    if not t or t in v.lower():
                        cb.pack(fill="x", pady=1)
                        var.set(v in active_set)
                        check_widgets.append((v, var, None))
                        filtered.append(v)
                    else:
                        cb.pack_forget()
                var_all.set(bool(filtered) and all(v in active_set for v in filtered))
                cb_all.configure(command=lambda f=filtered: toggle_all(var_all.get(), f))
                check_widgets.insert(0, ("__all__", var_all, filtered))
                check_inner.update_idletasks()
                canvas_c.configure(scrollregion=canvas_c.bbox("all"))

            def toggle_all(state, vals):
                for v in vals:
                    if v in check_vars:
                        check_vars[v].set(state)
                if state: active_set.update(vals)
                else:
                    for v in vals: active_set.discard(v)

            # Debounce 200ms — ไม่ filter จนกว่าจะพิมพ์หยุด
            _sjob = [None]
            def _on_search(*a):
                if _sjob[0]:
                    try: popup.after_cancel(_sjob[0])
                    except Exception: pass
                _sjob[0] = popup.after(200, lambda: build_checklist(search_var.get()))
            search_var.trace_add("write", _on_search)
            build_checklist()

            # OK / Cancel
            btn_bar = tk.Frame(inner, bg="white")
            btn_bar.pack(fill="x", padx=4, pady=4)

            def on_ok():
                new_sel = {v for v, var, _ in check_widgets
                        if v != "__all__" and var.get()}
                if new_sel == set(all_values):
                    self._header_filter_values.pop(col_name, None)
                else:
                    self._header_filter_values[col_name] = new_sel
                safe_destroy()
                self._apply_header_filters()

            tk.Button(btn_bar, text="OK", font=("Tahoma", 10),
                    bg="#3B82F6", fg="white", relief="flat",
                    command=on_ok, width=8).pack(side="right", padx=(4, 0))
            tk.Button(btn_bar, text="Cancel", font=("Tahoma", 10),
                    bg="#F3F4F6", relief="flat",
                    command=safe_destroy, width=8).pack(side="right")

            if col_name in self._header_filter_values:
                tk.Button(inner, text="✕ ล้าง filter", font=("Tahoma", 9),
                        fg="#EF4444", bg="white", relief="flat",
                        command=lambda: [self._header_filter_values.pop(col_name, None),
                                        safe_destroy(),
                                        self._apply_header_filters()]
                        ).pack(anchor="w", padx=8, pady=(0, 4))

            # ── วางตำแหน่ง popup ──────────────────────────────────────────────
            pw, ph = 240, 400

            # เมื่อ SetProcessDpiAwarenessContext ถูกเรียกแล้ว winfo_rooty() และ
            # geometry("+x+y") ใช้ physical pixels เหมือนกัน — คำนวณได้ตรง
            try:
                ch = target_sheet.CH
                header_bottom_y = ch.winfo_rooty() + ch.winfo_height()
            except Exception:
                header_bottom_y = event.y_root + 5

            # ขอบเขต window จริงสำหรับ clamp
            try:
                root_win = self.winfo_toplevel()
                win_x = root_win.winfo_rootx()
                win_y = root_win.winfo_rooty()
                win_w = root_win.winfo_width()
                win_h = root_win.winfo_height()
            except Exception:
                win_x, win_y = 0, 0
                win_w = self.winfo_screenwidth()
                win_h = self.winfo_screenheight()

            x = event.x_root
            y = header_bottom_y + 2

            # กันเกินขอบขวาของ window
            if x + pw > win_x + win_w - 5:
                x = win_x + win_w - pw - 5
            # ถ้า popup เกินล่าง window → เปิดขึ้นบน header แทน
            if y + ph > win_y + win_h - 5:
                y = header_bottom_y - ph - 2
            # กันออกนอกขอบซ้าย/บน
            x = max(x, win_x + 5)
            y = max(y, win_y + 5)

            # วาง geometry ขณะที่ popup ถูก withdraw (ซ่อน) ก่อน
            # ป้องกัน popup กระพริบที่ (0,0) แล้ว jump มาตำแหน่งจริง
            # และป้องกัน DPI-scaled window แสดงตำแหน่งผิดบน Windows Hi-DPI
            popup.withdraw()
            popup.geometry(f"{pw}x{ph}+{x}+{y}")

            def _show_and_focus():
                popup.deiconify()
                try:
                    search_entry.focus_set()
                except Exception:
                    pass

            popup.after(1, _show_and_focus)

            self._active_filter_popup = popup
            popup.bind("<FocusOut>", lambda e: self.after(300, lambda:
                    safe_destroy() if not popup._destroyed else None))

        except Exception:
            import traceback; traceback.print_exc()

    def _apply_header_filters(self):
        try:
            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            rh = int(30 * (self.zoom_level / 11.0)) # คำนวณความสูงตาม Zoom ไว้เผื่อกรณี Fallback

            # 1. จัดกลุ่ม Filter ว่าอยู่ฝั่ง Main (ขวา) หรือ Frozen (ซ้าย)
            main_filter_map = {}    
            frozen_filter_map = {}  

            for col_name, val_set in self._header_filter_values.items():
                try:
                    real_idx = self.columns.index(col_name)
                    display_idx = real_idx - col_offset
                    if display_idx < 0:
                        frozen_filter_map[real_idx] = val_set
                    else:
                        main_filter_map[display_idx] = val_set
                except ValueError:
                    pass

            # 2. แสดง (Unhide) แถวทั้งหมดกลับมาก่อน 
            try:
                self.sheet.display_rows("all")
                if self.sheet_frozen:
                    self.sheet_frozen.display_rows("all")
            except Exception:
                pass # ถ้าใช้เวอร์ชันเก่ามากๆ จะข้ามไป

            # 3. ตรวจสอบเงื่อนไขเพื่อหาว่า "แถวไหนบ้างที่ต้องซ่อน"
            if main_filter_map or frozen_filter_map:
                total_rows = self.sheet.get_total_rows()
                rows_to_hide = []

                for r in range(total_rows):
                    match = True

                    # เช็คฝั่ง Main sheet
                    for d_idx, allowed_vals in main_filter_map.items():
                        cell_val = str(self.sheet.get_cell_data(r, d_idx) or "").strip()
                        if cell_val not in allowed_vals:
                            match = False
                            break

                    # เช็คฝั่ง Frozen sheet
                    if match and self.sheet_frozen and frozen_filter_map:
                        for f_idx, allowed_vals in frozen_filter_map.items():
                            cell_val = str(self.sheet_frozen.get_cell_data(r, f_idx) or "").strip()
                            if cell_val not in allowed_vals:
                                match = False
                                break

                    # ถ้าข้อมูลไม่ตรงกับที่กรองไว้ ให้เก็บเข้า List สำหรับซ่อน
                    if not match:
                        rows_to_hide.append(r)

                # 4. สั่งซ่อนแถวทั้งหมดที่หาเจอพร้อมกัน (Tksheet Native)
                if rows_to_hide:
                    try:
                        self.sheet.hide_rows(rows_to_hide)
                        if self.sheet_frozen:
                            self.sheet_frozen.hide_rows(rows_to_hide)
                    except Exception:
                        # Fallback สำหรับเวอร์ชั่นเก่ามากๆ
                        for r in rows_to_hide:
                            self.sheet.row_height(r, 0)
                            if self.sheet_frozen:
                                self.sheet_frozen.row_height(r, 0)

            # 5. อัพเดทสีหัวคอลัมน์ที่เป็นสีส้ม (#F59E0B) เมื่อมี Filter ทำงานอยู่
            header_styles = self._get_header_styles_map()
            col_to_style = {c: (bg, fg) for (bg, fg), cols in header_styles.items() for c in cols}

            for col_name in self._filter_col_names:
                try:
                    real_idx = self.columns.index(col_name)
                    display_idx = real_idx - col_offset

                    is_active = col_name in self._header_filter_values
                    active_bg = "#F59E0B"
                    active_fg = "white"

                    if display_idx < 0:
                        # อยู่ใน frozen
                        if self.sheet_frozen:
                            if is_active:
                                self.sheet_frozen.highlight_cells(
                                    row=0, column=real_idx,
                                    bg=active_bg, fg=active_fg, canvas="header")
                            else:
                                h_bg, h_fg = col_to_style.get(col_name, ("#E5E7EB", "#111827"))
                                self.sheet_frozen.highlight_cells(
                                    row=0, column=real_idx,
                                    bg=h_bg, fg=h_fg, canvas="header")
                    else:
                        if is_active:
                            self.sheet.highlight_cells(
                                row=0, column=display_idx,
                                bg=active_bg, fg=active_fg, canvas="header")
                        else:
                            h_bg, h_fg = col_to_style.get(col_name, ("#E5E7EB", "#111827"))
                            self.sheet.highlight_cells(
                                row=0, column=display_idx,
                                bg=h_bg, fg=h_fg, canvas="header")
                except Exception:
                    pass

            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
                
            # Sync scroll แก้อาการหน้าจอเด้ง
            self.after(50, self._sync_vertical_scroll)
            # อัปเดตลูกศร filter หลัง redraw
            self.after(50, self._draw_filter_arrows)

        except Exception as e:
            print(f"Error in _apply_header_filters: {e}")
            import traceback; traceback.print_exc()

    def _on_header_press(self, event=None):
        if event:
            self._header_press_x = event.x_root
            self._header_press_y = event.y_root
            self._header_press_col_x = event.x

    def _get_col_from_header_click(self, event, sheet_widget, col_offset=0):
        """คำนวณ display_col จากพิกัดที่แท้จริงของ Canvas (แก้บั๊คเลื่อน Scrollbar แนวนอน)"""
        # 1. ลองใช้ฟังก์ชันของ tksheet ดูก่อน
        try:
            col = sheet_widget.identify_col(event)
            if col is not None: return int(col)
        except Exception: pass
        
        # 2. คำนวณแบบ Manual แต่ "ต้องรวมระยะ Scroll แนวนอนด้วย (canvasx)"
        try:
            # ⚠️ สำคัญ: ใช้ canvasx(event.x) เพื่อหาตำแหน่งที่แท้จริงหลังเลื่อนหน้าจอ
            cx = sheet_widget.CH.canvasx(event.x)
            
            # หักระยะความกว้างของ Row Index ออกก่อน (ถ้ามี)
            try: 
                row_idx_w = sheet_widget.RI.winfo_width()
            except Exception: 
                row_idx_w = 40 if col_offset == 0 else 0
            
            cx -= row_idx_w
            if cx < 0: return None
            
            accum = 0
            for c in range(sheet_widget.get_total_columns()):
                try: 
                    w = sheet_widget.column_width(c)
                except Exception: 
                    w = 120
                accum += w
                if cx < accum: 
                    return c
        except Exception as e: 
            print(f"Error finding col: {e}")
            
        return None


    def _sort_by_col(self, col_name, ascending=True):
        try:
            col_offset = self.frozen_col_count if self.sheet_frozen else 0
            real_idx = self.columns.index(col_name)
            total_rows = self.sheet.get_total_rows()
            rows_data = []
            for r in range(total_rows):
                main_row = list(self.sheet.get_row_data(r))
                frozen_row = list(self.sheet_frozen.get_row_data(r)) if self.sheet_frozen else []
                rows_data.append(frozen_row + main_row)

            def sort_key(row):
                v = str(row[real_idx]).strip() if real_idx < len(row) else ""
                try: return (0, float(v.replace(",", "")))
                except ValueError: return (1, v.lower())

            rows_data.sort(key=sort_key, reverse=not ascending)

            for r, row in enumerate(rows_data):
                if self.sheet_frozen:
                    for c in range(col_offset):
                        self.sheet_frozen.set_cell_data(r, c, row[c] if c < len(row) else "", redraw=False)
                for c in range(self.sheet.get_total_columns()):
                    rc = c + col_offset
                    self.sheet.set_cell_data(r, c, row[rc] if rc < len(row) else "", redraw=False)

            self.sheet.redraw()
            if self.sheet_frozen: self.sheet_frozen.redraw()

            if self.auto_save_job_id: self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
        except Exception: pass

    def _clear_all_header_filters(self):
        self._header_filter_values = {}
        self._apply_header_filters()

    def _restore_snapshot(self, frozen_data, main_data):
        col_offset = self.frozen_col_count
        target_rows = max(len(main_data), len(frozen_data))
        current_rows = self.sheet.get_total_rows()
        if target_rows > current_rows:
            self.sheet.insert_rows(target_rows - current_rows)
            if self.sheet_frozen:
                self.sheet_frozen.insert_rows(target_rows - current_rows)
        for r in range(target_rows):
            if self.sheet_frozen and r < len(frozen_data):
                for c, v in enumerate(frozen_data[r]):
                    self.sheet_frozen.set_cell_data(r, c, v, redraw=False)
            if r < len(main_data):
                for c, v in enumerate(main_data[r]):
                    self.sheet.set_cell_data(r, c, v, redraw=False)
        for r in range(target_rows, self.sheet.get_total_rows()):
            for c in range(self.sheet.get_total_columns()):
                self.sheet.set_cell_data(r, c, "", redraw=False)
            if self.sheet_frozen:
                for c in range(col_offset):
                    self.sheet_frozen.set_cell_data(r, c, "", redraw=False)
        self.sheet.redraw()
        if self.sheet_frozen:
            self.sheet_frozen.redraw()
        if self.auto_save_job_id:
            self.after_cancel(self.auto_save_job_id)
        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))

    # ================================================================== #
    def _smart_paste(self, event=None):
        """
        Ctrl+V แบบ smart สำหรับ Cost Benchmark:
        • paste ได้แม้คอลัมน์จะ readonly (เช่น Product SKU.)
        • ข้าม InlineSearchPopup สำหรับ รายการสินค้า / ชื่อ Supplier / รหัส Sale
        • ข้าม auto-calc columns อื่นๆ เช่น น้ำหนักรวม, ทุนรวม ฯลฯ
        • รองรับ multi-row, multi-column paste จาก Excel/ไฟล์งานอื่น
        """
        try:
            clip = self.clipboard_get()
        except Exception:
            return "break"

        if not clip:
            return "break"

        try:
            col_offset = self.frozen_col_count if self.sheet_frozen else 0

            # คอลัมน์ auto-calc ที่ *ห้าม* paste ทับ (คำนวณเองอยู่แล้ว)
            # ยกเว้น "Product SKU." ให้ paste ได้ เพราะ user ต้องการ copy จากไฟล์งานอื่น
            truly_auto = set(self._get_auto_cols())
            truly_auto.discard("Product SKU.")

            # สร้าง map display -> real สำหรับ main sheet
            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            main_visible = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]

            # parse clipboard: แยก row ด้วย \n, แยก column ด้วย \t
            import csv
            import io
            clip_cleaned = clip.strip('\r\n')
            if not clip_cleaned: return "break"
            
            reader = csv.reader(io.StringIO(clip_cleaned), delimiter='\t')
            clip_grid = list(reader)

            modified_rows = set()
            sel_cells = self.sheet.get_selected_cells()

            # 🚀 ฟีเจอร์ใหม่: ถ้าก๊อปปี้มาแค่ 1 ช่อง แต่ระบายเลือกไว้หลายช่อง ให้เติมลงทุกช่อง (Fill Paste)
            if len(clip_grid) == 1 and len(clip_grid[0]) == 1 and len(sel_cells) > 1:
                val_to_paste = clip_grid[0][0].strip()
                for r, c_disp in sel_cells:
                    if c_disp >= len(main_visible): continue
                    real_col = main_visible[c_disp]
                    col_name = self.columns[real_col]
                    if col_name in truly_auto: continue
                    
                    data_col = real_col - col_offset
                    self.sheet.set_cell_data(r, data_col, val_to_paste, redraw=False)
                    modified_rows.add(r)
            else:
                # 📌 การวางข้อมูลแบบปกติ (เริ่มจากช่องบนซ้ายสุด)
                sel = self.sheet.get_currently_selected()
                if not sel: return "break"
                
                start_row = int(sel[0])
                start_col_display = int(sel[1])

                for ri, cells in enumerate(clip_grid):
                    row_idx = start_row + ri
                    if row_idx >= self.sheet.get_total_rows(): break

                    ci = 0
                    disp_offset = 0
                    while ci < len(cells):
                        disp_idx = start_col_display + disp_offset
                        if disp_idx >= len(main_visible): break
                        
                        real_col = main_visible[disp_idx]
                        col_name = self.columns[real_col]
                        if col_name in truly_auto:
                            disp_offset += 1
                            continue
                            
                        data_col = real_col - col_offset
                        self.sheet.set_cell_data(row_idx, data_col, cells[ci].strip(), redraw=False)
                        modified_rows.add(row_idx)
                        ci += 1
                        disp_offset += 1

            # คำนวณและ redraw ทุก row ที่เปลี่ยน
            def _calc_after_paste(rows=sorted(modified_rows)):
                for row_idx in rows:
                    self._auto_calculate_sheet(row_idx)
                self.sheet.redraw()
                if self.sheet_frozen:
                    self.sheet_frozen.redraw()
                self._push_undo()

            self.after(30, _calc_after_paste)

            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

        except Exception as e:
            print(f"_smart_paste error: {e}")

        return "break"


    def _smart_paste_frozen(self, event=None):
        """Ctrl+V สำหรับ frozen sheet — paste แล้ว auto-fill product SKU / supplier code"""
        try:
            clip = self.clipboard_get()
        except Exception:
            return "break"
            
        if not clip:
            return "break"
            
        try:
            truly_auto = set(self._get_auto_cols())
            truly_auto.discard("Product SKU.")

            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            frozen_visible = [c for c in range(self.frozen_col_count) if c not in hidden_set]

            import csv
            import io
            clip_cleaned = clip.strip('\r\n')
            if not clip_cleaned: return "break"
            
            reader = csv.reader(io.StringIO(clip_cleaned), delimiter='\t')
            clip_grid = list(reader)

            modified_rows = set()
            sel_cells = self.sheet_frozen.get_selected_cells()

            # 🚀 ฟีเจอร์ใหม่: ถ้าก๊อปปี้มาแค่ 1 ช่อง แต่ระบายเลือกไว้หลายช่อง ให้เติมลงทุกช่อง (Fill Paste)
            if len(clip_grid) == 1 and len(clip_grid[0]) == 1 and len(sel_cells) > 1:
                val_to_paste = clip_grid[0][0].strip()
                for r, c_disp in sel_cells:
                    if c_disp >= len(frozen_visible): continue
                    real_col = frozen_visible[c_disp]
                    col_name = self.columns[real_col]
                    if col_name in truly_auto: continue
                    
                    data_col = real_col
                    self.sheet_frozen.set_cell_data(r, data_col, val_to_paste, redraw=False)
                    modified_rows.add(r)
            else:
                sel = self.sheet_frozen.get_currently_selected()
                if not sel: return "break"
                
                start_row = int(sel[0])
                start_display_col = int(sel[1])

                for ri, cells in enumerate(clip_grid):
                    row_idx = start_row + ri
                    if row_idx >= self.sheet_frozen.get_total_rows(): break
                    
                    disp_offset = 0
                    ci = 0
                    while ci < len(cells):
                        disp_idx = start_display_col + disp_offset
                        if disp_idx >= len(frozen_visible): break
                        
                        real_col = frozen_visible[disp_idx]
                        col_name = self.columns[real_col]
                        if col_name in truly_auto:
                            disp_offset += 1
                            continue
                            
                        data_col = real_col
                        self.sheet_frozen.set_cell_data(row_idx, data_col, cells[ci].strip(), redraw=False)
                        modified_rows.add(row_idx)
                        ci += 1
                        disp_offset += 1

            def _calc_after_paste_frozen(rows=sorted(modified_rows)):
                for row_idx in rows:
                    self._auto_calculate_sheet(row_idx)
                self.sheet_frozen.redraw()
                self.sheet.redraw()
                self._push_undo()

            self.after(30, _calc_after_paste_frozen)

            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

        except Exception as e:
            print(f"_smart_paste_frozen error: {e}")
            
        return "break"

    # ================================================================== #
    def _rebind_sheet(self):
        self.sheet.bind("<Control-MouseWheel>", self._on_ctrl_scroll)
        self.sheet.bind("<Shift-MouseWheel>", self._lock_horizontal_scroll)
        self.sheet.bind("<Return>", self._on_enter_pressed)
        self.sheet.bind("<KP_Enter>", self._on_enter_pressed)
        
        self.sheet.bind("<Control-c>", self._mt_copy_router)
        self.sheet.bind("<Control-v>", self._smart_paste)   
        self.sheet.bind("<Control-C>", self._mt_copy_router)
        self.sheet.bind("<Control-V>", self._smart_paste)   
        
        self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
        self.sheet.bind("<Control-r>", self._copy_selected_rows)
        self.sheet.bind("<Control-R>", self._copy_selected_rows)

        try:
            self.sheet.MT.bind("<ButtonPress-1>", self._capture_click_pos, add="+")
            self.sheet.MT.bind("<Double-ButtonPress-1>", self._on_mt_click, add="+")
            # KeyPress: เปิด popup เมื่อพิมพ์บน popup columns
            # ใช้ bind โดยตรงแทน begin_edit_cell เพราะ tksheet ไม่ fire begin_edit_cell
            # เมื่อ column ถูกซ่อนด้วย display_columns()
            self.sheet.MT.bind("<KeyPress>", self._on_mt_keypress, add="+")
            
            # 🟢 เพิ่ม 2 บรรทัดนี้: ให้คำนวณ Auto Sum เมื่อ "ปล่อยเมาส์" หรือ "ปล่อยปุ่มคีย์บอร์ด"
            self.sheet.MT.bind("<ButtonRelease-1>", lambda e: self.after(50, self._update_quick_calc), add="+")
            self.sheet.bind("<KeyRelease>", lambda e: self.after(50, self._update_quick_calc), add="+")
        except Exception:
            self.sheet.bind("<ButtonPress-1>", self._capture_click_pos, add="+")

        # 🟢 เพิ่ม Block นี้ด้วย: รองรับกรณีลากคลุมจากเลขบรรทัด (Row Index) หรือ หัวตาราง (Header)
        try:
            self.sheet.RI.bind("<ButtonRelease-1>", lambda e: self.after(50, self._update_quick_calc), add="+")
            self.sheet.CH.bind("<ButtonRelease-1>", lambda e: self.after(50, self._update_quick_calc), add="+")
        except Exception: pass

        self.sheet.enable_bindings((
            "single_select", "drag_select", "multi_select", "row_select", "column_select",
            "column_width_resize", "arrowkeys", "right_click_popup_menu", "rc_select",
            "copy", "cut", "paste", "delete", "edit_cell", "row_drag_and_drop",
        ))

        self.sheet.extra_bindings([
            ("begin_edit_cell", self._on_begin_edit_cell),
            ("end_edit_cell", self._on_end_edit_combined),
            ("cell_select", lambda e: self._on_cell_select_combined(e, is_frozen=False)),
            ("column_width_resize", lambda e: (self._save_col_widths(), self.after(60, self._draw_filter_arrows))),
            ("row_drag_and_drop", lambda e: self._on_row_drag_drop(e)),
            ("paste", lambda e: self._on_paste_event(e, is_frozen=False)),
            ("delete", lambda e: self._on_delete_event(e, is_frozen=False)),
        ])

        # ==========================================================
        # 📌 เพิ่มเมนูคลิกขวา (แทรกบรรทัด / ลบบรรทัด) สำหรับตารางหลัก
        # ==========================================================
        try:
            self.sheet.popup_menu_add_command("⮑ แทรกบรรทัดตรงนี้ (Insert Row)", self._insert_selected_row)
            self.sheet.popup_menu_add_command("🗑️ ลบบรรทัด (Delete Row)", self._delete_selected_rows)
            # เพิ่ม 2 บรรทัดนี้ครับ 👇
            self.sheet.popup_menu_add_command("🎨 ไฮไลท์สีช่อง (Highlight)", self._highlight_selected_cells)
            self.sheet.popup_menu_add_command("✨ ล้างสีไฮไลท์ (Clear Highlight)", self._clear_highlight_selected_cells)
        except Exception as e:
            print(f"Cannot add popup menu to main sheet: {e}")
            
        # CH bind หลัง enable_bindings เท่านั้น
        try:
            self.sheet.CH.bind("<ButtonPress-1>", self._on_header_press, add="+")
        except Exception: pass

        # MT.bind ต้องมาหลัง enable_bindings เพื่อไม่ให้ tksheet overwrite
        # Ctrl+V/C บน main MT ใช้ router เพื่อรองรับกรณี user click frozen sheet แต่ focus ยังอยู่ที่ main MT
        try:
            self.sheet.MT.bind("<Control-c>", self._mt_copy_router)
            self.sheet.MT.bind("<Control-C>", self._mt_copy_router)
            self.sheet.MT.bind("<Control-v>", self._mt_paste_router)
            self.sheet.MT.bind("<Control-V>", self._mt_paste_router)
            self.sheet.MT.bind("<Control-r>", self._copy_selected_rows, add="+")
            self.sheet.MT.bind("<Control-z>", self._undo)
            self.sheet.MT.bind("<Control-Z>", self._undo)
            self.sheet.MT.bind("<Control-y>", self._redo)
            self.sheet.MT.bind("<Control-Y>", self._redo)
            self.sheet.MT.bind("<Control-KeyPress>", self._on_ctrl_keypress_thai, add="+")
        except Exception: pass

        # Safety-net: dialog-level binding เผื่อ focus ไม่อยู่บน sheet ใดเลย
        self.bind("<Control-v>", self._mt_paste_router)
        self.bind("<Control-V>", self._mt_paste_router)
        self.bind("<Control-c>", self._mt_copy_router)
        self.bind("<Control-z>", self._undo)
        self.bind("<Control-Z>", self._undo)
        self.bind("<Control-y>", self._redo)
        self.bind("<Control-Y>", self._redo)
        try:
            self.bind("<Control-KeyPress>", self._on_ctrl_keypress_thai, add="+")
        except Exception: pass

        # ── text_editor_user_bound_keys ──────────────────────────────────────────
        # tksheet applies this dict to tktext EVERY time the text editor opens
        # (main_table.py line ~7505). Plain bind() each time = no accumulation.
        # This is the official hook for customising the in-cell text editor.
        def _editor_thai_cb(e):
            """Thai IME Ctrl+C/V/Z inside the cell text editor."""
            if e.keycode == 86:  # physical V key → Paste
                try:
                    clip = e.widget.clipboard_get()
                except Exception:
                    return "break"
                try: e.widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except Exception: pass
                try: e.widget.insert(tk.INSERT, clip)
                except Exception: pass
                return "break"
            if e.keycode == 67:  # physical C key → Copy
                try:
                    sel_text = e.widget.selection_get()
                    e.widget.clipboard_clear()
                    e.widget.clipboard_append(sel_text)
                except Exception: pass
                return "break"
            if e.keycode == 90:  # 🚀 physical Z key → บังคับออกจากการพิมพ์แล้ว Undo ตาราง
                try: self.sheet.cancel_edit()
                except Exception: pass
                self._undo()
                return "break"

        # edit mode: ลูกศรใน text editor → commit + เลื่อน + เปิด editor ใหม่
        _editor_keys = {
            "<Up>":               lambda e: self._commit_and_move("up",    is_frozen=False),
            "<Down>":             lambda e: self._commit_and_move("down",  is_frozen=False),
            "<Left>":             lambda e: self._commit_and_move("left",  is_frozen=False),
            "<Right>":            lambda e: self._commit_and_move("right", is_frozen=False),
            "<Control-KeyPress>": _editor_thai_cb,
        }
        _editor_keys_frozen = {
            "<Up>":               lambda e: self._commit_and_move("up",    is_frozen=True),
            "<Down>":             lambda e: self._commit_and_move("down",  is_frozen=True),
            "<Left>":             lambda e: self._commit_and_move("left",  is_frozen=True),
            "<Right>":            lambda e: self._commit_and_move("right", is_frozen=True),
            "<Control-KeyPress>": _editor_thai_cb,
        }
        try:
            self.sheet.MT.text_editor_user_bound_keys = _editor_keys
        except Exception: pass
        try:
            if self.sheet_frozen:
                self.sheet_frozen.MT.text_editor_user_bound_keys = _editor_keys_frozen
        except Exception: pass

        # navigation mode: ลูกศรบน MT โดยตรง → เลื่อน + เปิด editor ทันที
        def _make_nav_handler(direction, is_frozen):
            def _handler(event):
                mt = (self.sheet_frozen.MT if is_frozen and self.sheet_frozen else self.sheet.MT)
                if getattr(mt, 'text_editor', None) and getattr(mt.text_editor, 'open', False):
                    return  # edit mode → ปล่อยให้ text_editor_user_bound_keys จัดการ
                self._nav_move_and_edit(direction, is_frozen=is_frozen)
                return "break"
            return _handler

        try:
            for _d in ("up", "down", "left", "right"):
                _key = f"<{_d.capitalize()}>"
                self.sheet.MT.bind(_key, _make_nav_handler(_d, False), add="+")
        except Exception: pass

        # ── Cross-sheet deselect via FocusIn (main side) ─────────────────────
        try:
            self.sheet.MT.bind(
                "<FocusIn>",
                lambda e: self.sheet_frozen.deselect("all") if self.sheet_frozen else None,
                add="+"
            )
        except Exception: pass

    def _mt_paste_router(self, event=None):
        """Ctrl+V router — ตรวจ focus และ _last_active_sheet เพื่อ route paste ไปฝั่งที่ถูกต้อง
        แก้กรณี: user single-click frozen sheet แต่ keyboard focus ยังอยู่ main MT"""
        try:
            focused = self.focus_get()
            focused_class = focused.winfo_class() if focused else 'None'
            # ถ้า focus อยู่บน Entry/Text (formula bar, search, etc.) → อย่าแทรกแซง
            if focused and focused_class in ('Entry', 'Text', 'TEntry'):
                return
            if getattr(self, '_last_active_sheet', 'main') == 'frozen' and self.sheet_frozen:
                return self._smart_paste_frozen()
            else:
                return self._smart_paste()
        except Exception:
            pass

    def _mt_copy_router(self, event=None):
        """Ctrl+C router — copy จาก sheet ที่ user click ล่าสุด"""
        try:
            focused = self.focus_get()
            if focused and focused.winfo_class() in ('Entry', 'Text', 'TEntry'):
                return
            if getattr(self, '_last_active_sheet', 'main') == 'frozen' and self.sheet_frozen:
                self.sheet_frozen.copy()
            else:
                self.sheet.copy()
                
            # ✅ หน่วงเวลา 50ms ให้ tksheet ทำงานเสร็จ แล้วเรียกฟังก์ชันล้างฟันหนู
            self.after(50, self._clean_clipboard_quotes)
        except Exception:
            pass
        return "break"

    def _clean_clipboard_quotes(self):
        """ล้างเครื่องหมาย Double Quote ที่เกิดจากระบบ CSV ตอน Copy"""
        try:
            clip = self.clipboard_get()
            if not clip: return

            import csv
            import io
            # 1. ใช้โมดูล csv ช่วยอ่านข้อความที่มีการครอบ "" (tksheet คั่นคอลัมน์ด้วย Tab)
            reader = csv.reader(io.StringIO(clip), delimiter='\t')
            
            # 2. เอาข้อความที่ถอดฟันหนูส่วนเกินออกแล้ว มารวมกันใหม่
            cleaned_rows = []
            for row in reader:
                cleaned_rows.append('\t'.join(row))

            new_clip = '\n'.join(cleaned_rows)

            # 3. ยัดข้อมูลที่สะอาดแล้วกลับเข้า Clipboard ทับของเดิม
            self.clipboard_clear()
            self.clipboard_append(new_clip)
        except Exception as e:
            print(f"Clean clipboard error: {e}")

    def _on_ctrl_keypress_thai(self, event):
        """Route Thai IME Ctrl+V/C/Z — keycode 86=V, 67=C, 90=Z (reliable across IME versions)."""
        if event.keycode == 86:
            return self._mt_paste_router(event)
        elif event.keycode == 67:
            return self._mt_copy_router(event)
        elif event.keycode == 90:
            return self._undo(event)

    def _on_begin_edit_cell(self, event):
        try:
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
                typed = event.get('value', '') or ''
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)
                typed = getattr(event, 'value', '') or ''

            if row is None or col is None:
                return typed

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            hidden_set = set(getattr(self, 'hidden_cols_list', []))

            # 🛠️ แปลง Display Index เป็น Real Index เพื่อหาชื่อคอลัมน์ให้ถูกต้อง
            visible_main = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            if int(col) >= len(visible_main):
                return typed
            real_col = visible_main[int(col)]
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                return typed

            # data_col สำหรับใช้ดึง/เซ็ตข้อมูล (Data Index ของ Main Sheet คือ real_col - col_offset)
            _data_col_for_popup = real_col - col_offset

            if getattr(self, '_popup_opening', False):
                return None
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (int(row), int(col))):
                if typed and typed not in ('\r', '\n', '\t'):
                    try:
                        self._active_popup.entry.insert(tk.END, typed)
                        self._active_popup.filter_list(self._active_popup.search_var.get())
                    except Exception: pass
                return None

            self._popup_opening = True
            _row, _col = int(row), int(col)
            _typed = str(typed).strip() if typed and typed not in ('\r', '\n', '\t') else ''
            _data_col = _data_col_for_popup
            _data_list = popup_cols[col_name]

            def _open_popup():
                self._popup_opening = False
                try: self.sheet.close_text_editor(set_data=False)
                except Exception: pass
                try: self.sheet.cancel_edit()
                except Exception: pass

                try:
                    current_val = str(self.sheet.get_cell_data(_row, _data_col) or '').strip()
                except Exception:
                    current_val = ''
                prefill = _typed if _typed else current_val

                if (self._active_popup is not None and not self._active_popup._destroyed):
                    try: self._active_popup.safe_destroy()
                    except Exception: pass
                    self._active_popup = None

                def on_select(value):
                    try:
                        self.sheet.set_cell_data(_row, _data_col, value, redraw=True)
                        self._auto_calculate_sheet(_row)
                        self.sheet.redraw()
                        if self.auto_save_job_id is not None:
                            self.after_cancel(self.auto_save_job_id)
                        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                        if hasattr(self, 'save_status_label'):
                            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                        self._last_popup_cell = None
                        self.after(50, lambda: self._move_right(_row, _col))
                    except Exception as ex:
                        print(f"on_select error: {ex}")

                popup = InlineSearchPopup(self, _data_list, on_select)
                popup.place_near_cell(sheet_widget=self.sheet, row=_row, col=_col,
                                      row_h=self.zoom_level + 19, header_h=self.zoom_level + 24,
                                      data_col=_data_col)
                if prefill:
                    popup.entry.delete(0, tk.END)
                    popup.entry.insert(0, prefill)
                    if _typed: popup.entry.icursor(tk.END)
                    else: popup.entry.select_range(0, tk.END)
                    popup.filter_list(prefill)

                self._active_popup = popup
                self._last_popup_cell = (_row, _col)

            self.after(0, _open_popup)
            return None

        except Exception as e:
            self._popup_opening = False
            print(f"_on_begin_edit_cell error: {e}")
            return ""

    def _on_begin_edit_cell_frozen(self, event):
        try:
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
                typed = event.get('value', '') or ''
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)
                typed = getattr(event, 'value', '') or ''

            if row is None or col is None:
                return typed

            hidden_set = set(getattr(self, 'hidden_cols_list', []))

            # 🛠️ แปลง Display Index เป็น Real Index เพื่อหาชื่อคอลัมน์ให้ถูกต้อง
            visible_frozen = [c for c in range(self.frozen_col_count) if c not in hidden_set]
            if int(col) >= len(visible_frozen):
                return typed
            real_col = visible_frozen[int(col)]
            col_name = self.columns[real_col]

            popup_cols_frozen = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols_frozen:
                return typed

            if getattr(self, '_popup_opening', False):
                return None
            _row, _col = int(row), int(col)
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (_row, _col)):
                if typed and typed not in ('\r', '\n', '\t'):
                    try:
                        self._active_popup.entry.insert(tk.END, typed)
                        self._active_popup.filter_list(self._active_popup.search_var.get())
                    except Exception: pass
                return None

            self._popup_opening = True
            _typed = str(typed).strip() if typed and typed not in ('\r', '\n', '\t') else ''
            _data_list = popup_cols_frozen[col_name]
            _data_col = real_col  # 🛠️ Data Index ของ Frozen Sheet จะเท่ากับ real_col ตรงๆ

            def _open_popup_frozen():
                self._popup_opening = False
                try: self.sheet_frozen.close_text_editor(set_data=False)
                except Exception: pass
                try: self.sheet_frozen.cancel_edit()
                except Exception: pass

                try:
                    current_val = str(self.sheet_frozen.get_cell_data(_row, _data_col) or '').strip()
                except Exception:
                    current_val = ''
                prefill = _typed if _typed else current_val

                if (self._active_popup is not None and not self._active_popup._destroyed):
                    try: self._active_popup.safe_destroy()
                    except Exception: pass
                    self._active_popup = None

                def on_select(value):
                    try:
                        self.sheet_frozen.set_cell_data(_row, _data_col, value, redraw=True)
                        self._auto_calculate_sheet(_row)
                        self.sheet_frozen.redraw()
                        if self.auto_save_job_id is not None:
                            self.after_cancel(self.auto_save_job_id)
                        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                        if hasattr(self, 'save_status_label'):
                            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                        self._last_popup_cell = None
                        self.after(50, lambda: self._move_right(_row, _col, is_frozen=True))
                    except Exception as ex:
                        print(f"on_select frozen error: {ex}")

                popup = InlineSearchPopup(self, _data_list, on_select)
                popup.place_near_cell(sheet_widget=self.sheet_frozen, row=_row, col=_col,
                                      row_h=self.zoom_level + 19, header_h=self.zoom_level + 24,
                                      data_col=_data_col)
                if prefill:
                    popup.entry.delete(0, tk.END)
                    popup.entry.insert(0, prefill)
                    if _typed: popup.entry.icursor(tk.END)
                    else: popup.entry.select_range(0, tk.END)
                    popup.filter_list(prefill)

                self._active_popup = popup
                self._last_popup_cell = (_row, _col)

            self.after(0, _open_popup_frozen)
            return None

        except Exception as e:
            self._popup_opening = False
            print(f"_on_begin_edit_cell_frozen error: {e}")
            return ""

    def _open_editor_on_sheet(self, tgt, row, col):
        """เปิด editor บน sheet ที่ระบุ พร้อม fallback F2"""
        try:
            tgt.open_cell_editor(row=row, column=col)
        except Exception:
            try:
                tgt.MT.focus_set()
                tgt.MT.event_generate("<F2>")
            except Exception:
                tgt.MT.focus_set()

    def _nav_move_and_edit(self, direction, is_frozen=False):
        """Navigation mode: เลื่อน selection ตามลูกศร แล้วเปิด editor ทันที
        รองรับการข้ามฝั่ง frozen→main (Right) และ main→frozen (Left)
        - deselect อีกฝั่งเสมอ เพื่อไม่ให้เห็น selection 2 อันพร้อมกัน"""
        try:
            # ── 1. หา selection — fallback ตรวจทั้งสองฝั่ง ────────────────────
            tgt = self.sheet_frozen if is_frozen and self.sheet_frozen else self.sheet
            sel = tgt.get_currently_selected()
            if not sel or len(sel) < 2:
                # focus อาจเพี้ยน ลองอีกฝั่ง
                other = self.sheet if is_frozen else self.sheet_frozen
                if other:
                    sel2 = other.get_currently_selected()
                    if sel2 and len(sel2) >= 2:
                        sel = sel2
                        is_frozen = not is_frozen
                        tgt = other
            if not sel or len(sel) < 2:
                return

            row, col = int(sel[0]), int(sel[1])
            row_count = len(tgt.get_sheet_data())
            col_count = len(tgt.MT.col_positions) - 1  # displayed cols only (col_positions has n+1 entries)

            # ── 2. เคลียร์ selection อีกฝั่งเสมอ ────────────────────────────
            if is_frozen:
                self.sheet.deselect("all")
            elif self.sheet_frozen:
                self.sheet_frozen.deselect("all")

            # ── 3. ตรวจว่าต้องข้ามฝั่ง ───────────────────────────────────────
            cross_to_main   = is_frozen and self.sheet_frozen and direction == "right" and col >= col_count - 1
            cross_to_frozen = (not is_frozen) and self.sheet_frozen and direction == "left" and col <= 0

            if cross_to_main:
                tgt2 = self.sheet
                tgt.deselect("all")
                tgt2.set_currently_selected(row, 0)
                tgt2.see(row, 0)
                self._last_active_sheet = 'main'
                tgt2.MT.focus_set()
                self.after(20, lambda: self._open_editor_on_sheet(tgt2, row, 0))
                return

            if cross_to_frozen:
                tgt2 = self.sheet_frozen
                last_col = len(tgt2.MT.col_positions) - 2  # last displayed col index
                tgt.deselect("all")
                tgt2.set_currently_selected(row, last_col)
                tgt2.see(row, last_col)
                self._last_active_sheet = 'frozen'
                tgt2.MT.focus_set()
                self.after(20, lambda lc=last_col: self._open_editor_on_sheet(tgt2, row, lc))
                return

            # ── 4. เลื่อนปกติในฝั่งเดิม ──────────────────────────────────────
            if   direction == "up":    new_r, new_c = max(0, row - 1), col
            elif direction == "down":  new_r, new_c = min(row_count - 1, row + 1), col
            elif direction == "left":  new_r, new_c = row, max(0, col - 1)
            elif direction == "right": new_r, new_c = row, min(col_count - 1, col + 1)
            else: return

            tgt.set_currently_selected(new_r, new_c)
            tgt.see(new_r, new_c)
            self.after(20, lambda: self._open_editor_on_sheet(tgt, new_r, new_c))
        except Exception as ex:
            print(f"_nav_move_and_edit error: {ex}")

    def _commit_and_move(self, direction, is_frozen=False):
        """Commit cell edit แล้วเลื่อน + เปิด editor ใหม่ทันที (Excel-like)
        รองรับทั้ง main sheet และ frozen sheet รวมถึงการข้ามฝั่ง"""
        try:
            tgt = self.sheet_frozen if is_frozen and self.sheet_frozen else self.sheet
            sel = tgt.get_currently_selected()
            if not sel or len(sel) < 2:
                return "break"
            row, col = int(sel[0]), int(sel[1])
            row_count = len(tgt.get_sheet_data())
            col_count = len(tgt.MT.col_positions) - 1  # displayed cols only

            # ตรวจสอบการข้ามฝั่ง
            cross_to_main   = is_frozen and direction == "right" and col >= col_count - 1
            cross_to_frozen = (not is_frozen) and direction == "left" and col <= 0 and self.sheet_frozen

            if   direction == "up":    new_r, new_c = max(0, row - 1), col
            elif direction == "down":  new_r, new_c = min(row_count - 1, row + 1), col
            elif direction == "left":  new_r, new_c = row, max(0, col - 1)
            elif direction == "right": new_r, new_c = row, min(col_count - 1, col + 1)
            else: return "break"

            self._arrow_nav_in_progress = True
            tgt.MT.focus_set()  # commit via FocusOut

            def _do_move():
                self._arrow_nav_in_progress = False
                if cross_to_main:
                    tgt2 = self.sheet
                    tgt.deselect("all")
                    tgt2.set_currently_selected(row, 0)
                    tgt2.see(row, 0)
                    self._last_active_sheet = 'main'
                    self._open_editor_on_sheet(tgt2, row, 0)
                elif cross_to_frozen:
                    tgt2 = self.sheet_frozen
                    last_col = len(tgt2.MT.col_positions) - 2  # last displayed col index
                    self.sheet.deselect("all")
                    tgt2.set_currently_selected(row, last_col)
                    tgt2.see(row, last_col)
                    self._last_active_sheet = 'frozen'
                    self._open_editor_on_sheet(tgt2, row, last_col)
                else:
                    tgt.set_currently_selected(new_r, new_c)
                    tgt.see(new_r, new_c)
                    self._open_editor_on_sheet(tgt, new_r, new_c)

            self.after(30, _do_move)
        except Exception as ex:
            print(f"_commit_and_move error: {ex}")
        return "break"

    def _draw_filter_arrows(self):
        if not hasattr(self, '_filter_col_names'):
            return
            
        col_offset = self.frozen_col_count if self.sheet_frozen else 0
        
        def _draw_on_sheet(sheet_widget, is_frozen=False):
            try:
                ch = sheet_widget.CH
                ch.update_idletasks()
                ch.delete("filter_arrow")
                
                try:
                    header_h = sheet_widget.MT.header_height
                except:
                    header_h = 35
                
                # Build visible-column list for display-index mapping (accounts for hidden cols)
                hidden_set_draw = set(getattr(self, "hidden_cols_list", []))
                if is_frozen:
                    visible_for_draw = [c for c in range(col_offset) if c not in hidden_set_draw]
                else:
                    visible_for_draw = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set_draw]

                for col_name in self._filter_col_names:
                    try:
                        real_idx = self.columns.index(col_name)
                        if real_idx in hidden_set_draw:
                            continue  # column is hidden, no arrow needed
                        if is_frozen:
                            if real_idx >= col_offset:
                                continue
                            try:
                                display_idx = visible_for_draw.index(real_idx)
                            except ValueError:
                                continue
                        else:
                            if real_idx < col_offset:
                                continue
                            try:
                                display_idx = visible_for_draw.index(real_idx)
                            except ValueError:
                                continue
                        
                        # ← แก้ตรงนี้: col_positions รวม ri_w ไว้แล้ว ไม่ต้องบวกอีก
                        try:
                            x_start = sheet_widget.MT.col_positions[display_idx]
                            x_end = sheet_widget.MT.col_positions[display_idx + 1]
                        except (AttributeError, IndexError):
                            x_start = 0
                            for i in range(display_idx):
                                try: x_start += sheet_widget.column_width(i)
                                except: x_start += 120
                            try: x_end = x_start + sheet_widget.column_width(display_idx)
                            except: x_end = x_start + 120
                        
                        # วาดปุ่ม ▼
                        btn_w, btn_h, margin = 18, 14, 3
                        bx1 = x_end - btn_w - margin
                        by1 = (header_h - btn_h) // 2
                        bx2 = x_end - margin
                        by2 = by1 + btn_h
                        
                        has_filter = col_name in self._header_filter_values
                        bg_color = "#F59E0B" if has_filter else "#9CA3AF"
                        
                        ch.create_rectangle(
                            bx1, by1, bx2, by2,
                            fill=bg_color, outline="",
                            tags=("filter_arrow", f"arrow_{col_name}")
                        )
                        ch.create_text(
                            (bx1 + bx2) // 2, (by1 + by2) // 2,
                            text="▼",
                            font=("Tahoma", 7, "bold"),
                            fill="white",
                            tags=("filter_arrow", f"arrow_{col_name}")
                        )
                    except Exception:
                        pass
            except Exception as e:
                print(f"_draw_on_sheet error: {e}")
        
        _draw_on_sheet(self.sheet, is_frozen=False)
        if self.sheet_frozen:
            _draw_on_sheet(self.sheet_frozen, is_frozen=True)

    def _rebind_frozen_sheet(self):
        if not self.sheet_frozen:
            return

        self.sheet_frozen.bind("<Control-MouseWheel>", self._on_ctrl_scroll)
        self.sheet_frozen.bind("<Control-Button-4>", lambda e: self._zoom(1))
        self.sheet_frozen.bind("<Control-Button-5>", lambda e: self._zoom(-1))
        self.sheet_frozen.bind("<Control-r>", self._copy_selected_rows)
        self.sheet_frozen.bind("<Control-R>", self._copy_selected_rows)

        try:
            self.sheet_frozen.MT.bind("<Control-MouseWheel>", self._on_ctrl_scroll, add="+")
        except Exception: pass

        try:
            self.sheet_frozen.RI.bind("<Control-c>", self._copy_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-v>", self._paste_selected_rows, add="+")
            self.sheet_frozen.RI.bind("<Control-r>", self._copy_selected_rows, add="+")
        except Exception: pass

        self.sheet_frozen.enable_bindings((
            "single_select", "drag_select", "multi_select", "row_select", "column_select",
            "column_width_resize", "right_click_popup_menu", "rc_select",
            "copy", "cut", "paste", "delete", "edit_cell",
        ))
        # หมายเหตุ: ไม่ใส่ "arrowkeys" ใน frozen sheet
        # เพราะ custom handler ใน _make_nav_frozen จัดการเอง
        # ถ้าเปิดทั้งคู่จะได้ selection 2 อันพร้อมกัน

        self.sheet_frozen.extra_bindings([
            ("begin_edit_cell", self._on_begin_edit_cell_frozen),
            ("cell_select", lambda e: self._on_cell_select_combined(e, is_frozen=True)),
            ("end_edit_cell", lambda e: self._on_end_edit_combined(e, is_frozen=True)),
            ("column_width_resize", lambda e: (self._save_col_widths(), self.after(60, self._draw_filter_arrows))),
            ("paste", lambda e: self._on_paste_event(e, is_frozen=True)),
            ("delete", lambda e: self._on_delete_event(e, is_frozen=True)),
        ])

        # ==========================================================
        # 📌 เพิ่มเมนูคลิกขวา (แทรกบรรทัด / ลบบรรทัด) สำหรับตารางที่ตรึง
        # ==========================================================
        try:
            self.sheet_frozen.popup_menu_add_command("⮑ แทรกบรรทัดตรงนี้ (Insert Row)", self._insert_selected_row)
            self.sheet_frozen.popup_menu_add_command("🗑️ ลบบรรทัด (Delete Row)", self._delete_selected_rows)
            # เพิ่ม 2 บรรทัดนี้ครับ 👇
            self.sheet_frozen.popup_menu_add_command("🎨 ไฮไลท์สีช่อง (Highlight)", self._highlight_selected_cells)
            self.sheet_frozen.popup_menu_add_command("✨ ล้างสีไฮไลท์ (Clear Highlight)", self._clear_highlight_selected_cells)
        except Exception as e:
            print(f"Cannot add popup menu to frozen sheet: {e}")

        try:
            self.sheet_frozen.CH.bind("<ButtonPress-1>", self._on_header_press, add="+")
        except Exception: pass
        
        self.sheet_frozen.bind("<<SheetModified>>", self._on_sheet_modified)
        
        self.sheet_frozen.bind("<Control-c>", self._mt_copy_router)
        
        self.sheet_frozen.bind("<Control-v>", self._smart_paste_frozen)
        self.sheet_frozen.bind("<Control-V>", self._smart_paste_frozen)
        self.sheet_frozen.bind("<Return>", lambda e: self._on_enter_pressed(e, is_frozen=True))

        try:
            self.sheet_frozen.MT.bind("<ButtonPress-1>", self._capture_click_pos, add="+")
            # ✅ popup เปิดเฉพาะ double-click — single click แค่ select เซลล์ปกติ
            self.sheet_frozen.MT.bind("<Double-ButtonPress-1>", self._on_mt_click_frozen, add="+")
            
            # 🟢 ปรับปรุงใหม่อีก 2 บรรทัด: ให้ครอบคลุมการลากและการใช้ปุ่ม Shift+ลูกศร
            self.sheet_frozen.MT.bind("<ButtonRelease-1>", lambda e: self.after(50, self._update_quick_calc), add="+")
            self.sheet_frozen.bind("<KeyRelease>", lambda e: self.after(50, self._update_quick_calc), add="+")
        except Exception: pass

        # 🟢 เพิ่ม Block นี้เช่นกัน: เผื่อเราลากคลุมจากเลขบรรทัดฝั่งที่ตรึงไว้
        try:
            self.sheet_frozen.RI.bind("<ButtonRelease-1>", lambda e: self.after(50, self._update_quick_calc), add="+")
            self.sheet_frozen.CH.bind("<ButtonRelease-1>", lambda e: self.after(50, self._update_quick_calc), add="+")
        except Exception: pass

        # MT.bind ต้องมาหลัง enable_bindings — ใช้ router เดียวกับ main sheet เพื่อ consistency
        try:
            self.sheet_frozen.MT.bind("<Control-c>", self._mt_copy_router)
            self.sheet_frozen.MT.bind("<Control-C>", self._mt_copy_router)
            self.sheet_frozen.MT.bind("<Control-v>", self._mt_paste_router)
            self.sheet_frozen.MT.bind("<Control-V>", self._mt_paste_router)
            self.sheet_frozen.MT.bind("<Control-r>", self._copy_selected_rows, add="+")
            self.sheet_frozen.MT.bind("<Control-z>", self._undo)
            self.sheet_frozen.MT.bind("<Control-Z>", self._undo)
            self.sheet_frozen.MT.bind("<Control-y>", self._redo)
            self.sheet_frozen.MT.bind("<Control-Y>", self._redo)
            self.sheet_frozen.MT.bind("<Control-KeyPress>", self._on_ctrl_keypress_thai, add="+")
            self.sheet_frozen.MT.bind("<KeyPress>", self._on_mt_keypress_frozen, add="+")
        except Exception: pass

        # ✅ navigation mode arrow บน frozen MT (รองรับข้ามฝั่ง frozen→main)
        try:
            def _make_nav_frozen(direction):
                def _handler(event):
                    mt = self.sheet_frozen.MT
                    if getattr(mt, 'text_editor', None) and getattr(mt.text_editor, 'open', False):
                        return
                    self._nav_move_and_edit(direction, is_frozen=True)
                    return "break"
                return _handler
            for _d in ("up", "down", "left", "right"):
                _key = f"<{_d.capitalize()}>"
                self.sheet_frozen.MT.bind(_key, _make_nav_frozen(_d), add="+")
        except Exception: pass

        # text_editor_user_bound_keys สำหรับ frozen — ต้องใช้ is_frozen=True เสมอ
        # (ห้าม copy จาก main เพราะ main ใช้ is_frozen=False)
        # text_editor_user_bound_keys สำหรับ frozen — ต้องใช้ is_frozen=True เสมอ
        def _frozen_editor_thai_cb(e):
            if e.keycode == 86:
                try: clip = e.widget.clipboard_get()
                except Exception: return "break"
                try: e.widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except Exception: pass
                try: e.widget.insert(tk.INSERT, clip)
                except Exception: pass
                return "break"
            if e.keycode == 67:
                try:
                    t = e.widget.selection_get()
                    e.widget.clipboard_clear()
                    e.widget.clipboard_append(t)
                except Exception: pass
                return "break"
            if e.keycode == 90:  # 🚀 physical Z key → บังคับออกจากการพิมพ์แล้ว Undo ตาราง
                try: self.sheet_frozen.cancel_edit()
                except Exception: pass
                self._undo()
                return "break"

        # ── Cross-sheet deselect via FocusIn ──────────────────────────────────
        # เมื่อ frozen MT ได้ focus → deselect main เพื่อป้องกัน selection 2 อัน
        try:
            self.sheet_frozen.MT.bind(
                "<FocusIn>",
                lambda e: self.sheet.deselect("all"),
                add="+"
            )
        except Exception: pass

    def _on_sheet_modified(self, event=None):
        try:
            # หน่วงเวลา 50ms ให้ระบบตารางบันทึกค่าลงเซลล์เสร็จก่อน ค่อยจำประวัติ
            self.after(50, self._push_undo)
            if self.auto_save_job_id is not None:
                self.after_cancel(self.auto_save_job_id)
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        except Exception:
            pass

    def _push_undo(self):
        try:
            if not hasattr(self, '_history'):
                self._history = []
                self._history_idx = -1

            main_data = [list(row) for row in self.sheet.get_sheet_data()]
            frozen_data = [list(row) for row in self.sheet_frozen.get_sheet_data()] if getattr(self, 'sheet_frozen', None) else []
            
            # กันซ้ำกับสถานะล่าสุด (ถ้าไม่มีอะไรเปลี่ยน ไม่ต้องจำ)
            if self._history and self._history_idx >= 0:
                if self._history[self._history_idx] == (frozen_data, main_data):
                    return

            # ตัดประวัติอนาคตทิ้ง ถ้าเราย้อนกลับแล้วพิมพ์ใหม่แตกกิ่งไปทางอื่น
            if self._history_idx < len(self._history) - 1:
                self._history = self._history[:self._history_idx + 1]

            self._history.append((frozen_data, main_data))
            
            # จำกัดประวัติไม่เกิน 50 ครั้งเพื่อประหยัด RAM
            if len(self._history) > 50:
                self._history.pop(0)
            else:
                self._history_idx += 1
                
        except Exception as e:
            print(f"_push_undo error: {e}")

    def _undo(self, event=None):
        # 🚀 ป้องกันการ Undo ตาราง ถ้ายูสเซอร์กำลังพิมพ์อยู่ในช่อง (Focus อยู่ที่ Entry/Text)
        try:
            focused = self.focus_get()
            if focused and focused.winfo_class() in ('Entry', 'Text', 'TEntry'):
                return "break"
        except Exception: pass

        if not hasattr(self, '_history') or self._history_idx <= 0:
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⚠️ ย้อนกลับไม่ได้แล้ว", text_color="#F59E0B")
            return "break"
        try:
            self._history_idx -= 1
            frozen_data, main_data = self._history[self._history_idx]
            
            # ถอด Event ออกชั่วคราว ป้องกัน loop พังตอนกำลังกู้คืนข้อมูล
            self.sheet.unbind("<<SheetModified>>")
            if getattr(self, 'sheet_frozen', None):
                self.sheet_frozen.unbind("<<SheetModified>>")
                
            self._restore_snapshot(frozen_data, main_data)
            
            self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
            if getattr(self, 'sheet_frozen', None):
                self.sheet_frozen.bind("<<SheetModified>>", self._on_sheet_modified)

            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text=f"↩ ย้อนกลับ ({self._history_idx} ขั้นตอน)", text_color="#6366F1")
        except Exception as e:
            print(f"_undo error: {e}")
        return "break"

    def _redo(self, event=None):
        # 🚀 ป้องกันการ Redo ตาราง ถ้ายูสเซอร์กำลังพิมพ์อยู่ในช่อง
        try:
            focused = self.focus_get()
            if focused and focused.winfo_class() in ('Entry', 'Text', 'TEntry'):
                return "break"
        except Exception: pass

        if not hasattr(self, '_history') or self._history_idx >= len(self._history) - 1:
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⚠️ ไปข้างหน้าไม่ได้แล้ว", text_color="#F59E0B")
            return "break"
        try:
            self._history_idx += 1
            frozen_data, main_data = self._history[self._history_idx]
            
            self.sheet.unbind("<<SheetModified>>")
            if getattr(self, 'sheet_frozen', None):
                self.sheet_frozen.unbind("<<SheetModified>>")
                
            self._restore_snapshot(frozen_data, main_data)
            
            self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
            if getattr(self, 'sheet_frozen', None):
                self.sheet_frozen.bind("<<SheetModified>>", self._on_sheet_modified)
                
            steps_left = len(self._history) - 1 - self._history_idx
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text=f"↪ ทำซ้ำ ({steps_left} ขั้นตอน)", text_color="#6366F1")
        except Exception as e:
            print(f"_redo error: {e}")
        return "break"

    def _on_row_drag_drop(self, event=None):
        """หลัง drag row ใน main sheet → sync ข้อมูลไป frozen sheet ด้วย"""
        try:
            # ดึงข้อมูลใหม่จาก main sheet แล้ว rebuild frozen ให้ตรงกัน
            if not self.sheet_frozen:
                # ไม่มี frozen — auto save แล้วจบ
                if self.auto_save_job_id:
                    self.after_cancel(self.auto_save_job_id)
                self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                return

            # มี frozen — ต้อง sync ข้อมูลใหม่ทั้งหมดจาก DB หรือ rebuild
            # วิธีง่ายสุดคือโหลด data จาก main sheet แล้วเขียนลง frozen ใหม่
            main_data = self.sheet.get_sheet_data()
            total_rows = min(len(main_data), self.sheet_frozen.get_total_rows())

            # ← frozen sheet ยังเก็บข้อมูลเดิม ต้อง rebuild ตาม order ใหม่ของ main
            # แต่ frozen ไม่รู้ว่า row ไหนย้ายไปไหน ต้อง full reload จาก DB
            self.after(200, self._reload_frozen_from_main)

            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

        except Exception as e:
            print(f"_on_row_drag_drop error: {e}")

    def _reload_frozen_from_main(self):
        """sync ข้อมูล frozen sheet ให้ตรงกับ main sheet หลัง drag"""
        if not self.sheet_frozen:
            return
        try:
            main_data = self.sheet.get_sheet_data()
            col_offset = self.frozen_col_count
            # ← ถ้า frozen มีข้อมูลน้อยกว่า main ให้เพิ่ม row ก่อน
            frozen_rows = self.sheet_frozen.get_total_rows()
            main_rows = len(main_data)
            if main_rows > frozen_rows:
                self.sheet_frozen.insert_rows(main_rows - frozen_rows)

            # เขียนข้อมูลจาก full_row (main+frozen merged) ลง frozen
            # ปัญหาคือ main sheet ไม่มีข้อมูล frozen cols
            # ต้องใช้ _save_to_db แล้ว _load_from_db แทน
            # แต่นั่น disruptive เกินไป → ใช้วิธี swap frozen rows ตาม order ใหม่

            # วิธีที่ดีที่สุด: rebuild ทั้งหมดจาก current state
            self._rebuild_frozen_layout()
        except Exception as e:
            print(f"_reload_frozen_from_main error: {e}")

    def _on_mt_click_frozen(self, event=None):
        try:
            self._last_active_sheet = 'frozen'
            if hasattr(self, 'sheet') and self.sheet:
                self.sheet.deselect("all")
            # delayed focus_set — ต้องรอให้ tksheet's own after() callbacks ทำงานก่อน
            # ไม่งั้น tksheet จะ steal focus กลับไปหลัง immediate focus_set
            self.after(25, lambda: (
                self.sheet_frozen.MT.focus_set()
                if self.sheet_frozen and not (self._active_popup and not self._active_popup._destroyed)
                else None
            ))
            try:
                row = self.sheet_frozen.MT.identify_row(y=event.y, allow_end=False)
                col = self.sheet_frozen.MT.identify_col(x=event.x, allow_end=False)
            except Exception:
                try:
                    row = self.sheet_frozen.MT.get_row_at_y(event.y)
                    col = self.sheet_frozen.MT.get_col_at_x(event.x)
                except Exception:
                    return

            if row is None or col is None:
                return

            row, col = int(row), int(col)

            # Map display col → real col (frozen zone อาจมีคอลัมซ่อนได้)
            hidden_set = set(getattr(self, "hidden_cols_list", []))
            visible_frozen = [c for c in range(self.frozen_col_count) if c not in hidden_set]
            if col >= len(visible_frozen):
                return
            real_col = visible_frozen[col]
            if real_col >= len(self.columns):
                return

            col_name = self.columns[real_col]
            status_opts = [
                "WIN", "STOCK",
                "LOSE - เซลล์ไม่ทราบสาเหตุ",
                "LOSE - ลูกค้าได้ราคาถูกกว่า (มีราคาเทียบ)",
                "LOSE - ลูกค้าได้ราคาถูกกว่า (ไม่มีราคาเทียบ)",
                "LOSE - ไม่มีกำหนดใช้งานที่แน่นอน เช่น ขอราคาเพื่อเสนอ",
                "LOSE - ยื่นประมูลงาน (ระบุเดือนในหมายเหตุ)",
                "LOSE - ลูกค้าเปลี่ยนสเปคการใช้งาน",
                "LOSE - ลูกค้าใช้เจ้าที่มีเครดิต"
            ]
            popup_cols = {
                # ชื่อ Supplier / รายการสินค้า / รหัส Sale จัดการโดย begin_edit_cell แล้ว
                "PRIORITY":       ["HOT", "WARM", "COLD", "HOT-สั่งผลิต", "WARM-สั่งผลิต", "COLD-สั่งผลิต", "ไม่แจ้ง"],
                "สถานะ":          status_opts,
                "Select":         ["✔", "เทียบ", "เทียบเพื่อชุบ", "เทียบเพื่อชุบ ✔"],
            }

            # รายการสินค้า / ชื่อ Supplier / รหัส Sale → เปิด popup โดยตรง
            if col_name in ("รายการสินค้า", "ชื่อ Supplier", "รหัส Sale"):
                if self._active_popup and not self._active_popup._destroyed:
                    if self._last_popup_cell == (row, col):
                        try: self._active_popup.entry.focus_set()
                        except Exception: pass
                        return
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None

                data_map = {
                    "รายการสินค้า": self.product_list,
                    "ชื่อ Supplier":  self.supplier_list,
                    "รหัส Sale":      self.sales_list,
                }
                _data_list = data_map[col_name]
                _data_col = real_col  # frozen sheet ใช้ real_col ตรงๆ
                try:
                    current_val = str(self.sheet_frozen.get_cell_data(row, _data_col) or '').strip()
                except Exception:
                    current_val = ''

                _row, _col = row, col
                def on_select(value, r=_row, dc=_data_col, c=_col):
                    try:
                        self.sheet_frozen.set_cell_data(r, dc, value, redraw=True)
                        self._auto_calculate_sheet(r)
                        self.sheet_frozen.redraw()
                        if self.auto_save_job_id is not None:
                            self.after_cancel(self.auto_save_job_id)
                        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                        if hasattr(self, 'save_status_label'):
                            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                        self._last_popup_cell = None
                        self.after(50, lambda: self._move_right(r, c, is_frozen=True))
                    except Exception as ex:
                        print(f"on_select frozen error: {ex}")

                popup = InlineSearchPopup(self, _data_list, on_select)
                popup.place_near_cell(sheet_widget=self.sheet_frozen, row=_row, col=_col,
                                      row_h=self.zoom_level + 19, header_h=self.zoom_level + 24,
                                      data_col=_data_col)
                if current_val:
                    popup.entry.delete(0, tk.END)
                    popup.entry.insert(0, current_val)
                    popup.entry.select_range(0, tk.END)
                    popup.filter_list(current_val)
                self._active_popup = popup
                self._last_popup_cell = (_row, _col)
                return

            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and self._last_popup_cell == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row = row
            _disp_col = col       # display col — ใช้สำหรับ UI (_move_right, _last_popup_cell)
            _data_col = real_col  # data col — ใช้สำหรับ get/set_cell_data ของ frozen sheet

            try:
                current_val = str(self.sheet_frozen.get_cell_data(_row, _data_col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value, _r=_row, _dc=_data_col, _dd=_disp_col):
                try:
                    self.sheet_frozen.set_cell_data(_r, _dc, value, redraw=True)
                    self._auto_calculate_sheet(_r)
                    self.sheet_frozen.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_r, _dd, is_frozen=True))
                except Exception as ex:
                    print(f"on_select frozen error: {ex}")

            self.after(50, lambda dc=_data_col: self._open_popup_delayed(_row, _disp_col, data_list, current_val, on_select, is_frozen=True, data_col=dc))

        except Exception as e:
            print(f"_on_mt_click_frozen error: {e}")

    def _refresh_dropdown_data(self):
        """โหลด dropdown data ใหม่จาก DB"""
        self.sales_list = []
        self.supplier_list = []
        self.product_list = []
        self.product_sku_map = {}
        self.supplier_code_map = {}
        self.product_category_map = {}
        self._load_dropdown_data()

    def _on_mt_keypress(self, event=None):
        """KeyPress บน main sheet MT — เปิด popup สำหรับ รายการสินค้า / ชื่อ Supplier / รหัส Sale
        ใช้วิธีนี้แทน begin_edit_cell เพราะ tksheet ไม่ fire begin_edit_cell
        เมื่อ column ถูกซ่อนด้วย display_columns()"""
        try:
            if not event or not event.char or ord(event.char) < 32:
                return
            # ตรวจ modifier keys — ถ้ากด Ctrl/Alt ให้ข้ามไป
            if event.state & 0x4 or event.state & 0x8:  # Ctrl or Alt
                return

            # ดู selected cell
            sel = self.sheet.get_currently_selected()
            if not sel or len(sel) < 2:
                return
            row, col = int(sel[0]), int(sel[1])

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            
            if col >= len(visible_real):
                return
            real_col = visible_real[col]
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                return

            # ถ้า popup เปิดอยู่แล้วในเซลล์นี้ ให้ส่ง char ไปที่ entry แทน
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                    self._active_popup.entry.insert(tk.END, event.char)
                    self._active_popup.filter_list(self._active_popup.search_var.get())
                except Exception:
                    pass
                return "break"

            if getattr(self, '_popup_opening', False):
                return

            # ปิด popup เก่า
            if self._active_popup is not None and not self._active_popup._destroyed:
                try: self._active_popup.safe_destroy()
                except Exception: pass
                self._active_popup = None

            self._popup_opening = True
            _row, _col = row, col
            _typed = event.char
            _data_list = popup_cols[col_name]
            # data col = นับ visible cols ก่อน real_col (ถูกต้องแม้มี hidden cols)
            _data_col = real_col - col_offset

            def _open_popup():
                self._popup_opening = False
                try: self.sheet.close_text_editor(set_data=False)
                except Exception: pass
                try: self.sheet.cancel_edit()
                except Exception: pass

                try:
                    current_val = str(self.sheet.get_cell_data(_row, _data_col) or '').strip()
                except Exception:
                    current_val = ''

                if self._active_popup is not None and not self._active_popup._destroyed:
                    try: self._active_popup.safe_destroy()
                    except Exception: pass
                    self._active_popup = None

                def on_select(value, r=_row, dc=_data_col, c=_col):
                    try:
                        self.sheet.set_cell_data(r, dc, value, redraw=True)
                        self._auto_calculate_sheet(r)
                        self.sheet.redraw()
                        if self.auto_save_job_id is not None:
                            self.after_cancel(self.auto_save_job_id)
                        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                        if hasattr(self, 'save_status_label'):
                            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                        self._last_popup_cell = None
                        self.after(50, lambda: self._move_right(r, c))
                    except Exception as ex:
                        print(f"on_select error: {ex}")

                popup = InlineSearchPopup(self, _data_list, on_select)
                popup.place_near_cell(sheet_widget=self.sheet, row=_row, col=_col,
                                      row_h=self.zoom_level + 19, header_h=self.zoom_level + 24,
                                      data_col=_data_col)
                popup.entry.delete(0, tk.END)
                popup.entry.insert(0, _typed)
                popup.entry.icursor(tk.END)
                popup.filter_list(_typed)
                self._active_popup = popup
                self._last_popup_cell = (_row, _col)

            self.after(0, _open_popup)
            return "break"  # บล็อก tksheet ไม่ให้เปิด inline editor

        except Exception as e:
            self._popup_opening = False
            print(f"_on_mt_keypress error: {e}")
    
    def _on_mt_keypress_frozen(self, event=None):
        """ดักจับตอนพิมพ์ในฝั่งที่ตรึงคอลัมน์ไว้"""
        try:
            if not event or not event.char or ord(event.char) < 32:
                return
            if event.state & 0x4 or event.state & 0x8:  # Ctrl or Alt
                return

            sel = self.sheet_frozen.get_currently_selected()
            if not sel or len(sel) < 2:
                return
            row, col = int(sel[0]), int(sel[1])

            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            visible_frozen = [c for c in range(self.frozen_col_count) if c not in hidden_set]
            if col >= len(visible_frozen):
                return
            real_col = visible_frozen[col]
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                return

            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                    self._active_popup.entry.insert(tk.END, event.char)
                    self._active_popup.filter_list(self._active_popup.search_var.get())
                except Exception:
                    pass
                return "break"

            if getattr(self, '_popup_opening', False):
                return

            if self._active_popup is not None and not self._active_popup._destroyed:
                try: self._active_popup.safe_destroy()
                except Exception: pass
                self._active_popup = None

            self._popup_opening = True
            _row, _col = row, col
            _typed = event.char
            _data_list = popup_cols[col_name]
            _data_col = real_col

            def _open_popup():
                self._popup_opening = False
                try: self.sheet_frozen.close_text_editor(set_data=False)
                except Exception: pass
                try: self.sheet_frozen.cancel_edit()
                except Exception: pass

                try:
                    current_val = str(self.sheet_frozen.get_cell_data(_row, _data_col) or '').strip()
                except Exception:
                    current_val = ''

                if self._active_popup is not None and not self._active_popup._destroyed:
                    try: self._active_popup.safe_destroy()
                    except Exception: pass
                    self._active_popup = None

                def on_select(value, r=_row, dc=_data_col, c=_col):
                    try:
                        self.sheet_frozen.set_cell_data(r, dc, value, redraw=True)
                        self._auto_calculate_sheet(r)
                        self.sheet_frozen.redraw()
                        if self.auto_save_job_id is not None:
                            self.after_cancel(self.auto_save_job_id)
                        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                        if hasattr(self, 'save_status_label'):
                            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                        self._last_popup_cell = None
                        self.after(50, lambda: self._move_right(r, c, is_frozen=True))
                    except Exception as ex:
                        print(f"on_select error: {ex}")

                popup = InlineSearchPopup(self, _data_list, on_select)
                popup.place_near_cell(sheet_widget=self.sheet_frozen, row=_row, col=_col,
                                      row_h=self.zoom_level + 19, header_h=self.zoom_level + 24,
                                      data_col=_data_col)
                popup.entry.delete(0, tk.END)
                popup.entry.insert(0, _typed)
                popup.entry.icursor(tk.END)
                popup.filter_list(_typed)
                self._active_popup = popup
                self._last_popup_cell = (_row, _col)

            self.after(0, _open_popup)
            return "break"

        except Exception as e:
            self._popup_opening = False
            print(f"_on_mt_keypress_frozen error: {e}")

    def _on_mt_click(self, event=None):
        try:
            self._last_active_sheet = 'main'
            if hasattr(self, 'sheet_frozen') and self.sheet_frozen:
                self.sheet_frozen.deselect("all")
            self.after(25, lambda: self.sheet.MT.focus_set())
            # ✅ แก้: ใช้ identify จาก MT canvas โดยตรง ไม่ต้องคำนวณ pixel เอง
            try:
                row = self.sheet.MT.identify_row(y=event.y, allow_end=False)
                col = self.sheet.MT.identify_col(x=event.x, allow_end=False)
            except Exception:
                return

            if row is None or col is None:
                return

            row, col = int(row), int(col)
            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0

            # Map display col → real col (accounts for hidden columns)
            hidden_set = set(getattr(self, "hidden_cols_list", []))
            visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            if col >= len(visible_real):
                return
            real_col = visible_real[col]
            if real_col >= len(self.columns):
                return

            col_name = self.columns[real_col]
            status_opts = [
                "WIN", "STOCK",
                "LOSE - เซลล์ไม่ทราบสาเหตุ",
                "LOSE - ลูกค้าได้ราคาถูกกว่า (มีราคาเทียบ)",
                "LOSE - ลูกค้าได้ราคาถูกกว่า (ไม่มีราคาเทียบ)",
                "LOSE - ไม่มีกำหนดใช้งานที่แน่นอน เช่น ขอราคาเพื่อเสนอ",
                "LOSE - ยื่นประมูลงาน (ระบุเดือนในหมายเหตุ)",
                "LOSE - ลูกค้าเปลี่ยนสเปคการใช้งาน",
                "LOSE - ลูกค้าใช้เจ้าที่มีเครดิต"
            ]
            popup_cols = {
                # ชื่อ Supplier / รายการสินค้า / รหัส Sale จัดการโดย begin_edit_cell แล้ว
                "PRIORITY":       ["HOT", "WARM", "COLD", "HOT-สั่งผลิต", "WARM-สั่งผลิต", "COLD-สั่งผลิต", "ไม่แจ้ง"],
                "สถานะ":          status_opts,
                "Select":         ["✔", "เทียบ", "เทียบเพื่อชุบ", "เทียบเพื่อชุบ ✔"],
            }

            # รายการสินค้า / ชื่อ Supplier / รหัส Sale → เปิด popup โดยตรง
            if col_name in ("รายการสินค้า", "ชื่อ Supplier", "รหัส Sale"):
                if self._active_popup and not self._active_popup._destroyed:
                    if self._last_popup_cell == (row, col):
                        try: self._active_popup.entry.focus_set()
                        except Exception: pass
                        return
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None

                data_map = {
                    "รายการสินค้า": self.product_list,
                    "ชื่อ Supplier":  self.supplier_list,
                    "รหัส Sale":      self.sales_list,
                }
                _data_list = data_map[col_name]
                # นับ visible cols ก่อน real_col (รองรับคอลัมน์ซ่อน)
                _data_col = real_col - col_offset
                try:
                    current_val = str(self.sheet.get_cell_data(row, _data_col) or '').strip()
                except Exception:
                    current_val = ''

                _row, _col = row, col
                def on_select(value, r=_row, dc=_data_col, c=_col):
                    try:
                        self.sheet.set_cell_data(r, dc, value, redraw=True)
                        self._auto_calculate_sheet(r)
                        self.sheet.redraw()
                        if self.auto_save_job_id is not None:
                            self.after_cancel(self.auto_save_job_id)
                        self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                        if hasattr(self, 'save_status_label'):
                            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                        self._last_popup_cell = None
                        self.after(50, lambda: self._move_right(r, c))
                    except Exception as ex:
                        print(f"on_select error: {ex}")

                popup = InlineSearchPopup(self, _data_list, on_select)
                popup.place_near_cell(sheet_widget=self.sheet, row=_row, col=_col,
                                      row_h=self.zoom_level + 19, header_h=self.zoom_level + 24,
                                      data_col=_data_col)
                if current_val:
                    popup.entry.delete(0, tk.END)
                    popup.entry.insert(0, current_val)
                    popup.entry.select_range(0, tk.END)
                    popup.filter_list(current_val)
                self._active_popup = popup
                self._last_popup_cell = (_row, _col)
                return

            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and self._last_popup_cell == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row, _col = row, col

            # data_col = local data index ใน scroll sheet (ใช้สำหรับ set/get_cell_data)
            # นับเฉพาะ visible cols ก่อน real_col เพื่อรองรับคอลัมน์ซ่อน
            _data_col = real_col - col_offset

            try:
                current_val = str(self.sheet.get_cell_data(_row, _data_col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _data_col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            self.after(50, lambda dc=_data_col: self._open_popup_delayed(_row, _col, data_list, current_val, on_select, data_col=dc))

        except Exception as e:
            print(f"_on_mt_click error: {e}")

    def _manual_refresh_dropdown(self):
        """กดปุ่ม 🔄 รีเฟรชสินค้า — โหลดรายการสินค้า/Supplier ใหม่จาก DB ทันที"""
        try:
            self._refresh_dropdown_data()
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="✅ รีเฟรชสินค้าแล้ว", text_color="#16A34A")
                self.after(3000, lambda: self.save_status_label.configure(
                    text="", text_color="#6B7280"))
        except Exception as e:
            print(f"manual refresh error: {e}")

    def _start_dropdown_refresh_timer(self):
        """Refresh dropdown data ทุก 10 วินาที — timer วิ่งต่อเสมอแม้ refresh จะ error"""
        try:
            self._refresh_dropdown_data()
        except Exception as e:
            print(f"dropdown refresh error: {e}")
        finally:
            self.after(10000, self._start_dropdown_refresh_timer)

    def _on_cell_select_combined(self, event=None, is_frozen=False):
        # 1. รองรับระบบคลิกเพื่อสร้างสูตร
        self._on_sheet_click_for_formula(event)

        # 🚀 2. บังคับดึง Focus ทันทีที่คลิกเซลล์ (แก้ปัญหา Ctrl+Z, C, V ไม่ติดตอนแรก)
        if is_frozen:
            self._last_active_sheet = 'frozen'
            if hasattr(self, 'sheet_frozen') and self.sheet_frozen:
                self.after(10, lambda: self.sheet_frozen.MT.focus_set())
            if hasattr(self, 'sheet') and self.sheet:
                self.sheet.deselect("all")
        else:
            self._last_active_sheet = 'main'
            if hasattr(self, 'sheet') and self.sheet:
                self.after(10, lambda: self.sheet.MT.focus_set())
            if hasattr(self, 'sheet_frozen') and self.sheet_frozen:
                self.sheet_frozen.deselect("all")

        # 3. ปิด Popup ค้นหาทันทีถ้ามีการคลิกเลือกเซลล์ใหม่
        if self._active_popup is not None and not self._active_popup._destroyed:
            try:
                self._active_popup.safe_destroy()
            except Exception:
                pass
            self._active_popup = None
            self._last_popup_cell = None

        # 🚀 3. ปิด Popup ค้นหาทันทีถ้ามีการคลิกเลือกเซลล์ใหม่
        if self._active_popup is not None and not self._active_popup._destroyed:
            try:
                self._active_popup.safe_destroy()
            except Exception:
                pass
            self._active_popup = None
            self._last_popup_cell = None

    def _open_popup_delayed(self, row, col, data_list, current_val, on_select, is_frozen=False, data_col=None):
        try:
            if (self._active_popup is not None and not self._active_popup._destroyed and self._last_popup_cell == (row, col)):
                try: self._active_popup.entry.focus_set()
                except: pass
                return

            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed: self._active_popup.safe_destroy()
                except: pass
                self._active_popup = None

            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            try: target_sheet.close_text_editor(set_data=False)
            except: pass

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_near_cell(
                sheet_widget=target_sheet,
                row=row,
                col=col,
                row_h=self.zoom_level + 19,
                header_h=self.zoom_level + 24,
                data_col=data_col,
            )

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (row, col)

        except Exception as e:
            print(f"_open_popup_delayed error: {e}")
            
    def _on_begin_edit_block(self, event=None):
        try:
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = int(col) + col_offset
            if real_col >= len(self.columns):
                return

            col_name = self.columns[real_col]
            if col_name not in {"รายการสินค้า", "ชื่อ Supplier", "รหัส Sale"}:
                return

            # กัน loop — ถ้า popup เปิดอยู่แล้วในเซลล์นี้ ไม่ต้องทำอะไร
            _row, _col = int(row), int(col)
            if getattr(self, '_popup_opening', False):
                return
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (_row, _col)):
                return

            self._popup_opening = True

            def _close_and_popup():
                self._popup_opening = False
                try:
                    self.sheet.close_text_editor(set_data=False)
                except Exception:
                    pass
                try:
                    self.sheet.cancel_edit()
                except Exception:
                    pass
                self._on_cell_select_popup({"row": _row, "column": _col})

            self.after(0, _close_and_popup)

        except Exception as e:
            self._popup_opening = False
            print(f"_on_begin_edit_block error: {e}")

    def _on_cell_select_popup(self, event=None):
        """
        Triggered by tksheet extra_binding 'cell_select' — API มาตรฐาน ทำงานได้ทุก version
        รองรับทั้ง dict, tuple, และ object ที่มี .row / .column attribute
        """
        try:
            # แกะ row, col จาก event ที่ tksheet ส่งมา (format อาจต่างกันตาม version)
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None:
                return
            row, col = int(row), int(col)
            if row < 0 or col < 0:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            # แปลง display col → real col (รองรับคอลัมน์ซ่อน)
            visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            if col >= len(visible_real):
                return
            real_col = visible_real[col]
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }

            # ถ้าคลิกคอลัมน์อื่น → ปิด popup ที่เปิดอยู่แล้วออก
            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            # กัน popup เปิดซ้ำในเซลล์เดิม — focus กลับไปที่ entry แทน
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            # ปิด popup เก่าก่อน
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            # ยกเลิก inline editor ของ tksheet ถ้ากำลังจะเปิด
            try:
                self.sheet.close_text_editor(set_data=False)
            except Exception:
                pass

            data_list = popup_cols[col_name]
            _row, _col = row, col
            # นับ visible cols ก่อน real_col เพื่อได้ data index ที่ถูกต้อง
            _data_col = real_col - col_offset

            try:
                current_val = str(self.sheet.get_cell_data(_row, _data_col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _data_col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_at_mouse(ref_widget=self.sheet)

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (_row, _col)

        except Exception as e:
            import traceback
            print(f"_on_cell_select_popup error: {e}")
            traceback.print_exc()


    def _on_sheet_mouse_click(self, event=None):
        """ดักจาก mouse click โดยตรง — ใช้ identify_region หาว่าคลิกเซลล์ไหน"""
        try:
            if event is None:
                return

            # ใช้ identify_region หรือ identify_cell ของ tksheet
            try:
                region = self.sheet.identify_region(event)
                if region != "table":
                    return  # คลิก header, scrollbar, ฯลฯ → ไม่ทำอะไร
            except Exception:
                pass

            try:
                row, col = self.sheet.identify_cell(event)
            except Exception:
                # fallback: คำนวณเองจาก pixel
                try:
                    row = self.sheet.identify_row(event)
                    col = self.sheet.identify_col(event)
                except Exception:
                    return

            if row is None or col is None:
                return

            row, col = int(row), int(col)
            if row < 0 or col < 0:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            # แปลง display col → real col (รองรับคอลัมน์ซ่อน)
            visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            if col >= len(visible_real):
                return
            real_col = visible_real[col]
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            # กัน popup เปิดซ้ำในเซลล์เดิม
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and getattr(self, '_last_popup_cell', None) == (row, col)):
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            # ปิด popup เก่า
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]
            _row, _col = row, col
            # นับ visible cols ก่อน real_col เพื่อได้ data index ที่ถูกต้อง
            _data_col = real_col - col_offset

            try:
                current_val = str(self.sheet.get_cell_data(_row, _data_col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _data_col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_at_mouse(ref_widget=self.sheet)

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (_row, _col)

        except Exception as e:
            import traceback
            print(f"_on_sheet_mouse_click error: {e}")
            traceback.print_exc()

    def _debug_cell_select(self, event=None):
        pass

    def _on_cell_single_click_popup(self, event=None):
        """Single click → เปิด popup ทันที"""
        try:
            # ✅ แก้: รองรับ None value ใน dict
            row, col = None, None

            if isinstance(event, dict):
                r = event.get('row')
                c = event.get('column')
                if r is None or c is None:
                    return  # tksheet ส่ง event ที่ยังไม่มีข้อมูล เช่นตอน deselect
                row, col = int(r), int(c)
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                if event[0] is None or event[1] is None:
                    return
                row, col = int(event[0]), int(event[1])
            else:
                r = getattr(event, 'row', None)
                c = getattr(event, 'column', None)
                if r is None or c is None:
                    return
                row, col = int(r), int(c)

            if row < 0 or col < 0:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            hidden_set = set(getattr(self, 'hidden_cols_list', []))
            # แปลง display col → real col (รองรับคอลัมน์ซ่อน)
            visible_real = [c for c in range(col_offset, len(self.columns)) if c not in hidden_set]
            if col >= len(visible_real):
                return
            real_col = visible_real[col]
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                # ถ้าคลิกคอลัมน์อื่น → ปิด popup ที่เปิดอยู่
                if self._active_popup and not self._active_popup._destroyed:
                    self._active_popup.safe_destroy()
                    self._active_popup = None
                    self._last_popup_cell = None
                return

            # กัน popup เปิดซ้ำในเซลล์เดิม
            if (self._active_popup is not None
                    and not self._active_popup._destroyed
                    and self._last_popup_cell == (row, col)):
                # popup เปิดอยู่แล้วในเซลล์นี้ — focus ไปที่ entry
                try:
                    self._active_popup.entry.focus_set()
                except Exception:
                    pass
                return

            # ปิด popup เก่า
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            data_list = popup_cols[col_name]

            # นับ visible cols ก่อน real_col เพื่อได้ data index ที่ถูกต้อง
            _data_col = real_col - col_offset
            try:
                current_val = str(self.sheet.get_cell_data(row, _data_col) or "").strip()
            except Exception:
                current_val = ""

            # capture ค่า row, col ใน closure
            _row, _col = row, col

            def on_select(value):
                try:
                    self.sheet.set_cell_data(_row, _data_col, value, redraw=True)
                    self._auto_calculate_sheet(_row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self._last_popup_cell = None
                    self.after(50, lambda: self._move_right(_row, _col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            popup = InlineSearchPopup(self, data_list, on_select)
            popup.place_near_cell(
                sheet_widget=self.sheet,
                row=row,
                col=col,
                row_h=self.zoom_level + 19,
                header_h=self.zoom_level + 24,
                data_col=_data_col,
            )

            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup
            self._last_popup_cell = (row, col)

        except Exception as e:
            import traceback
            print(f"_on_cell_single_click_popup error: {e}")
            traceback.print_exc()
    # ================================================================== #
    # POPUP SEARCH
    # ================================================================== #
    def _on_begin_edit_for_popup(self, event=None):
        """ดักตอน tksheet กำลังจะเปิด inline editor — ถ้าเป็นคอลัมน์ popup ให้เปิด popup แทน"""
        try:
            # ปิด popup เก่าก่อน
            if self._active_popup is not None:
                try:
                    if not self._active_popup._destroyed:
                        self._active_popup.safe_destroy()
                except Exception:
                    pass
                self._active_popup = None

            # แกะ row, col
            if isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
            elif isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = event[0], event[1]
            else:
                row = getattr(event, 'row', None)
                col = getattr(event, 'column', None)

            if row is None or col is None:
                return

            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            real_col = int(col) + col_offset
            if real_col >= len(self.columns):
                return
            col_name = self.columns[real_col]

            popup_cols = {
                "รายการสินค้า": self.product_list,
                "ชื่อ Supplier":  self.supplier_list,
                "รหัส Sale":      self.sales_list,
            }
            if col_name not in popup_cols:
                return  # ปล่อย tksheet ทำงานปกติ

            # ยกเลิก inline editor ของ tksheet
            try:
                self.sheet.close_text_editor(set_data=False)
            except Exception:
                pass
            try:
                self.sheet._reset_text_editor()
            except Exception:
                pass

            data_list = popup_cols[col_name]

            data_col = real_col - col_offset
            # ดึงค่าเดิมในเซลล์มา pre-fill
            try:
                current_val = str(self.sheet.get_cell_data(row, data_col) or "").strip()
            except Exception:
                current_val = ""

            def on_select(value):
                try:
                    self.sheet.set_cell_data(row, data_col, value, redraw=True)
                    self._auto_calculate_sheet(row)
                    self.sheet.redraw()
                    if self.auto_save_job_id is not None:
                        self.after_cancel(self.auto_save_job_id)
                    self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
                    if hasattr(self, 'save_status_label'):
                        self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                    self.after(50, lambda: self._move_right(row, col))
                except Exception as ex:
                    print(f"on_select error: {ex}")

            # สร้าง popup (ไม่มี anchor แล้ว — position จัดการใน place_near_cell)
            popup = InlineSearchPopup(self, data_list, on_select)

            # วางตำแหน่งใต้เซลล์ที่คลิก
            popup.place_near_cell(
                sheet_widget=self.sheet,
                row=row,
                col=col,
                row_h=self.zoom_level + 19,
                header_h=self.zoom_level + 24,
                data_col=data_col,
            )

            # pre-fill ค่าเดิม
            if current_val:
                popup.entry.insert(0, current_val)
                popup.entry.select_range(0, tk.END)
                popup.filter_list(current_val)

            self._active_popup = popup

        except Exception as e:
            import traceback
            print(f"_on_begin_edit_for_popup error: {e}")
            traceback.print_exc()

    # ================================================================== #
    def _save_col_widths(self):
        try:
            col_offset = self.frozen_col_count if self.sheet_frozen else 0
            total = len(self.columns)
            hidden_set = set(getattr(self, "hidden_cols_list", []))

            frozen_width = 40
            if self.sheet_frozen:
                # Frozen columns: display idx == real idx (ไม่มี hidden ใน frozen ปกติ)
                for i in range(self.frozen_col_count):
                    try:
                        w = self.sheet_frozen.column_width(i)
                        if w and w > 0:
                            self.col_widths_cache[i] = w
                            frozen_width += w
                    except IndexError: pass

                # Main columns: map display idx → real idx โดยใช้ visible_real list
                # (เพราะถ้ามี hidden columns display idx ≠ real idx)
                visible_real = [c for c in range(col_offset, total) if c not in hidden_set]
                for display_i in range(self.sheet.get_total_columns()):
                    if display_i >= len(visible_real):
                        break
                    real_col = visible_real[display_i]
                    try:
                        w = self.sheet.column_width(display_i)
                        if w and w > 0:
                            self.col_widths_cache[real_col] = w
                    except IndexError: pass

                frozen_width += 4
                if hasattr(self, '_resize_layout_job'):
                    try: self.after_cancel(self._resize_layout_job)
                    except: pass

                def _update_frames(fw=frozen_width):
                    try:
                        self._current_frozen_width = fw
                        self.sheet_frozen.place(x=0, y=0, width=fw, relheight=1.0)
                        self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                        def _on_table_resize(event):
                            try:
                                _fw = getattr(self, '_current_frozen_width', fw)
                                self.sheet.place(x=_fw, y=0, relwidth=1.0, relheight=1.0, width=-_fw)
                            except Exception: pass
                        self.table_frame.bind("<Configure>", _on_table_resize)
                    except Exception: pass

                self._resize_layout_job = self.after(100, _update_frames)
            else:
                # ไม่มี frozen: map display idx → real idx ผ่าน visible_real
                visible_real = [c for c in range(total) if c not in hidden_set]
                for display_i in range(self.sheet.get_total_columns()):
                    if display_i >= len(visible_real):
                        break
                    real_col = visible_real[display_i]
                    try:
                        w = self.sheet.column_width(display_i)
                        if w and w > 0:
                            self.col_widths_cache[real_col] = w
                    except IndexError: pass

            if hasattr(self, '_save_settings_job'):
                try: self.after_cancel(self._save_settings_job)
                except: pass
            self._save_settings_job = self.after(800, self._save_user_settings)
        except Exception as e:
            print(f"_save_col_widths error: {e}")

    def _update_quick_calc(self, event=None):
        """คำนวณ Sum/Count/Avg จากเซลล์ที่เลือก เหมือน Excel — รองรับทั้ง main และ frozen sheet (แก้บั๊กซ่อนคอลัมน์)"""
        try:
            col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0
            # ดึงรายการคอลัมน์ที่ถูกซ่อนไว้
            hidden_set = set(getattr(self, "hidden_cols_list", []))
            values = []
            total_selected = 0

            # สร้างรายการคอลัมน์ที่มองเห็นจริง (เพื่อแปลง display index เป็น data index)
            main_visible = [col for col in range(col_offset, len(self.columns)) if col not in hidden_set]
            frozen_visible = [col for col in range(col_offset) if col not in hidden_set]

            # helper: แปลง display row → data row (รองรับกรณี filter ซ่อนแถว)
            def to_data_row_main(disp_r):
                try:
                    dr = self.sheet.MT.displayed_rows
                    if dr is not None and not isinstance(dr, str):
                        return dr[disp_r]
                except (IndexError, AttributeError, TypeError):
                    pass
                return disp_r

            def to_data_row_frozen(disp_r):
                try:
                    dr = self.sheet_frozen.MT.displayed_rows
                    if dr is not None and not isinstance(dr, str):
                        return dr[disp_r]
                except (IndexError, AttributeError, TypeError):
                    pass
                return disp_r

            # ── ดึงจาก main sheet ──────────────────────────────────────
            try:
                selected = self.sheet.get_selected_cells()
                if selected:
                    total_selected += len(selected)
                    for r, c in selected:
                        try:
                            data_r = to_data_row_main(r)
                            if c < len(main_visible):
                                real_col = main_visible[c]
                                data_col = real_col - col_offset
                                val = self.sheet.get_cell_data(data_r, data_col)
                                if val and str(val).strip():
                                    num = float(str(val).replace(',', '').replace('%', '').strip())
                                    values.append(num)
                        except (ValueError, TypeError, IndexError):
                            pass
            except Exception:
                pass

            # ── ดึงจาก frozen sheet ด้วย ──────────────────────────────
            if self.sheet_frozen:
                try:
                    frozen_selected = self.sheet_frozen.get_selected_cells()
                    if frozen_selected:
                        total_selected += len(frozen_selected)
                        for r, c in frozen_selected:
                            try:
                                data_r = to_data_row_frozen(r)
                                if c < len(frozen_visible):
                                    real_col = frozen_visible[c]
                                    data_col = real_col
                                    val = self.sheet_frozen.get_cell_data(data_r, data_col)
                                    if val and str(val).strip():
                                        num = float(str(val).replace(',', '').replace('%', '').strip())
                                        values.append(num)
                            except (ValueError, TypeError, IndexError):
                                pass
                except Exception:
                    pass

            if total_selected == 0:
                self.quick_calc_label.configure(text="")
                return

            num_count = len(values)

            if num_count == 0:
                self.quick_calc_label.configure(text=f"จำนวน: {total_selected}")
                return

            total = sum(values)
            avg = total / num_count

            def fmt(n):
                if n == int(n):
                    return f"{int(n):,}"
                return f"{n:,.2f}"

            text = f"จำนวน: {total_selected}    ผลรวม: {fmt(total)}    เฉลี่ย: {fmt(avg)}"
            self.quick_calc_label.configure(text=text)
        except Exception as e:
            print(f"Quick calc error: {e}")

    def _restore_col_widths(self):
        try:
            if not self.col_widths_cache:
                return

            # ✅ ปิด auto_resize ก่อน restore ทุกครั้ง
            self.sheet.set_options(auto_resize_columns=False, auto_resize_row_index=False)
            if self.sheet_frozen:
                self.sheet_frozen.set_options(auto_resize_columns=False, auto_resize_row_index=False)

            hidden_set_r = set(getattr(self, "hidden_cols_list", []))
            if self.sheet_frozen is not None:
                frozen_width = 40
                for i in range(self.frozen_col_count):
                    w = self.col_widths_cache.get(i)
                    w = w if w and w >= 8 else 120
                    self.sheet_frozen.column_width(i, w)
                    frozen_width += w

                # Map display index → real col index, skipping hidden cols
                visible_real_r = [c for c in range(self.frozen_col_count, len(self.columns)) if c not in hidden_set_r]
                for di in range(self.sheet.get_total_columns()):
                    if di >= len(visible_real_r): break
                    w = self.col_widths_cache.get(visible_real_r[di])
                    w = w if w and w >= 8 else 120
                    self.sheet.column_width(di, w)

                frozen_width += 4
                try:
                    self._current_frozen_width = frozen_width  # อัปเดตค่าล่าสุดเพื่อให้ _on_table_resize ใช้ค่าถูกต้องเสมอ
                    try: self.table_frame.update_idletasks()
                    except Exception: pass
                    self.sheet_frozen.place(x=0, y=0, width=frozen_width, relheight=1.0)
                    self.sheet.place(x=frozen_width, y=0, relwidth=1.0, relheight=1.0, width=-frozen_width)
                    def _on_table_resize(event):
                        try:
                            fw = getattr(self, '_current_frozen_width', frozen_width)
                            self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                        except Exception: pass
                    self.table_frame.bind("<Configure>", _on_table_resize)
                except Exception: pass

            else:
                # Map display index → real col index, skipping hidden cols
                visible_real_r = [c for c in range(len(self.columns)) if c not in hidden_set_r]
                for di in range(self.sheet.get_total_columns()):
                    if di >= len(visible_real_r): break
                    w = self.col_widths_cache.get(visible_real_r[di])
                    w = w if w and w >= 8 else 120
                    self.sheet.column_width(di, w)
            
            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
                
        except Exception as e:
            print(f"Restore widths error: {e}")

    def _on_end_edit_combined(self, event=None, is_frozen=False):
        try:
            row, col = None, None
            datarn = None  # data row index (filter-safe)

            # ดึงตำแหน่ง row, col จาก event
            if isinstance(event, (tuple, list)) and len(event) >= 2:
                row, col = int(event[0]), int(event[1])
                datarn = row  # tuple format: data index
            elif isinstance(event, dict):
                row = event.get('row')
                col = event.get('column')
                if row is not None:
                    row = int(row)
                if col is not None:
                    col = int(col)
                # event['data'] keys are (datarn, datacn) — actual data indices, filter-safe
                data_dict = event.get('data', {})
                if data_dict:
                    keys = list(data_dict.keys())
                    if keys and isinstance(keys[0], (tuple, list)):
                        datarn = int(keys[0][0])
                if datarn is None:
                    datarn = row

            if row is None or col is None or datarn is None:
                return

            # --- ส่วนที่ 1: คำนวณคณิตศาสตร์ Inline (เช่น =2+2) ---
            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            try:
                # อ่านค่าที่เพิ่งพิมพ์จาก event['value'] (ไม่ต้องอ่านจาก sheet ซึ่งอาจใช้ display index)
                cell_val = str(event.get('value', '') if isinstance(event, dict) else
                               (target_sheet.get_cell_data(datarn, col) or "")).strip()
                if cell_val.startswith("=") and len(cell_val) > 1:
                    math_expr = cell_val[1:].replace(',', '')
                    import re
                    if re.match(r'^[\d\.\+\-\*\/\(\)\s]+$', math_expr):
                        result = eval(math_expr, {"__builtins__": None}, {})
                        if isinstance(result, (int, float)):
                            if result == int(result):
                                formatted_result = f"{int(result)}"
                            else:
                                formatted_result = f"{result:.2f}"
                            # ใช้ datarn (data index) เพื่อเขียนลง sheet ได้ถูก row เสมอ
                            target_sheet.set_cell_data(datarn, col, formatted_result, redraw=False)
            except Exception as e:
                print(f"Inline math error: {e}")

            # 1. คำนวณสูตรในบรรทัดนั้น — ใช้ datarn เสมอ (filter-safe, data index)
            self._auto_calculate_sheet(datarn)

            # 2. รีเฟรชตาราง — ไม่แตะ display_rows เพื่อรักษา filter state ไว้
            self.sheet.redraw()
            if is_frozen and self.sheet_frozen:
                self.sheet_frozen.redraw()

            # 3. สั่ง Auto Save
            if self.auto_save_job_id is not None:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))

            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")

            # 4. เลื่อนเคอร์เซอร์ไปทางขวา (ใช้ display row/col สำหรับ navigation)
            if not getattr(self, '_arrow_nav_in_progress', False):
                self.after(10, lambda: self._move_right(row, col, is_frozen))

        except Exception as e:
            print(f"_on_end_edit_combined error: {e}")

    def _move_right(self, row, col, is_frozen=False):
        try:
            if is_frozen and self.sheet_frozen:
                total_cols_frozen = self.sheet_frozen.get_total_columns()
                next_col = col + 1
                if next_col < total_cols_frozen:
                    self.sheet_frozen.deselect("all")
                    try:
                        self.sheet_frozen.select_cell(row, next_col)
                        self.sheet_frozen.see(row, next_col)
                    except Exception:
                        # ถ้าคอลัมน์ถัดไปโดนซ่อน หรือ Error ให้เด้งข้ามไปตารางฝั่งขวาเลย
                        try:
                            self.sheet.select_cell(row, 0)
                            self.sheet.see(row, 0)
                            self.sheet.focus_set()
                        except Exception: pass
                else:
                    # ➡️ ถ้าสุดขอบตารางตรึง ให้เด้งข้ามไปตารางหลักฝั่งขวา!
                    self.sheet_frozen.deselect("all")
                    try:
                        self.sheet.select_cell(row, 0)
                        self.sheet.see(row, 0)
                        self.sheet.focus_set()
                    except Exception: pass
            else:
                total_cols = self.sheet.get_total_columns()
                total_rows = self.sheet.get_total_rows()
                next_col = col + 1
                if next_col < total_cols:
                    self.sheet.deselect("all")
                    try:
                        self.sheet.select_cell(row, next_col)
                        self.sheet.see(row, next_col)
                    except Exception: pass
                else:
                    if row + 1 < total_rows:
                        self.sheet.deselect("all")
                        # ↩️ ขึ้นบรรทัดใหม่ กระโดดกลับไปเริ่มที่ตารางตรึง (ถ้ามี)
                        if self.sheet_frozen:
                            try:
                                self.sheet_frozen.select_cell(row + 1, 0)
                                self.sheet_frozen.see(row + 1, 0)
                                self.sheet_frozen.focus_set()
                            except Exception: pass
                        else:
                            try:
                                self.sheet.select_cell(row + 1, 0)
                                self.sheet.see(row + 1, 0)
                            except Exception: pass
        except Exception as e:
            print(f"_move_right error: {e}")

    def _on_enter_pressed(self, event=None, is_frozen=False):
        try:
            target_sheet = self.sheet_frozen if is_frozen else self.sheet
            curr = target_sheet.get_currently_selected()
            if curr:
                self._move_right(curr[0], curr[1], is_frozen)
            return "break"
        except Exception: pass

    def _lock_horizontal_scroll(self, event):
        return "break"

    def _on_ctrl_scroll(self, event):
        if event.delta > 0:
            self._zoom(1)
        else:
            self._zoom(-1)
        return "break"


    # ================================================================== #

    def _extract_rows_from_event(self, event, is_frozen=False) -> set:
        """Helper: ดึง row indices ที่ได้รับผลกระทบจาก event payload ของ tksheet"""
        rows = set()
        sheet = self.sheet_frozen if is_frozen else self.sheet

        # 1) อ่านจาก event payload โดยตรง (tksheet ส่งมาเป็น list หรือ dict)
        if event is not None:
            try:
                for item in event:
                    if isinstance(item, (list, tuple)) and len(item) >= 1:
                        rows.add(int(item[0]))
            except TypeError:
                if isinstance(event, dict):
                    for key in ("cells after", "cells", "rows"):
                        val = event.get(key)
                        if val:
                            try:
                                for item in val:
                                    rows.add(int(item[0]) if isinstance(item, (list, tuple)) else int(item))
                            except Exception:
                                pass

        # 2) Fallback — ดึงจาก selection ปัจจุบัน (ครอบคลุมกรณี drag-select)
        try:
            for r, _ in (sheet.get_selected_cells() or []):
                rows.add(int(r))
        except Exception:
            pass
        try:
            for r in (sheet.get_selected_rows() or []):
                rows.add(int(r))
        except Exception:
            pass

        return rows

    def _on_paste_event(self, event=None, is_frozen=False):
        """extra_binding 'paste' — ยิงหลัง right-click paste เสร็จ
        (Ctrl+V ถูก intercept โดย _smart_paste แล้ว — ฟังก์ชันนี้จัดการเฉพาะ RC-paste)"""
        try:
            rows = self._extract_rows_from_event(event, is_frozen)
            if not rows:
                return

            # หน่วง 50ms ให้ tksheet commit ข้อมูลลงเซลล์ก่อนค่อยคำนวณ
            def _do_calc():
                for r in sorted(rows):
                    self._auto_calculate_sheet(r)
                self.sheet.redraw()
                if self.sheet_frozen:
                    self.sheet_frozen.redraw()
                if self.auto_save_job_id:
                    self.after_cancel(self.auto_save_job_id)
                if hasattr(self, 'save_status_label'):
                    self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
                self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))

            self.after(50, _do_calc)
        except Exception as e:
            print(f"_on_paste_event error: {e}")

    def _on_delete_event(self, event=None, is_frozen=False):
        """extra_binding 'delete' — ยิงหลัง user กด Delete/Backspace
        คำนวณทุกบรรทัดที่โดนลบ รองรับ drag-select หลายบรรทัด"""
        try:
            rows = self._extract_rows_from_event(event, is_frozen)
            if not rows:
                return

            for r in sorted(rows):
                self._auto_calculate_sheet(r)

            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()

            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
        except Exception as e:
            print(f"_on_delete_event error: {e}")

    def _auto_calculate_sheet(self, row_idx):
        col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0

        def col2num(col_str):
            expn = 0
            col_num = 0
            for char in reversed(col_str.upper()):
                col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
                expn += 1
            return col_num - 1

        try:
            row_data = self.sheet.get_row_data(row_idx)
            for c_idx, cell_val in enumerate(row_data):
                val_str = str(cell_val).strip()
                if val_str.startswith('=') and len(val_str) > 1:
                    try:
                        expr = val_str[1:].replace(',', '').upper()
                        cell_refs = set(re.findall(r'[A-Z]+\d+', expr))
                        for ref in cell_refs:
                            match = re.match(r'([A-Z]+)(\d+)', ref)
                            if match:
                                c_str, r_str = match.groups()
                                target_col = col2num(c_str)
                                target_row = int(r_str) - 1
                                ref_val = self.sheet.get_cell_data(target_row, target_col)
                                if not ref_val or str(ref_val).strip() == "":
                                    ref_val = "0"
                                else:
                                    ref_val = str(ref_val).replace(',', '').replace('%', '')
                                expr = re.sub(rf'\b{ref}\b', str(ref_val), expr)
                        result = eval(expr, {"__builtins__": None}, {})
                        if isinstance(result, (int, float)):
                            self.sheet.set_cell_data(row_idx, c_idx, f"{float(result):.2f}", redraw=False)
                    except Exception:
                        pass
        except Exception:
            pass

        def get_val(col_name):
            try:
                val = self._sheet_get(row_idx, col_name)
                return float(str(val).replace(',', '').replace('%', '')) if val and str(val).strip() else 0.0
            except (ValueError, IndexError):
                return 0.0

        def get_str(col_name):
            try:
                val = self._sheet_get(row_idx, col_name)
                return str(val or "").strip()
            except (IndexError, ValueError):
                return ""

        def set_val(col_name, val, is_text=False):
            if not is_text:
                formatted_val = "" if val == 0 else f"{val:,.2f}"
            else:
                formatted_val = "" if (val is None or val == "%") else str(val)
            self._sheet_set(row_idx, col_name, formatted_val)

        _, _date_main_disp = self._col_to_disp("วันที่ขอราคา")
        row_data = self.sheet.get_row_data(row_idx)
        is_row_active = any(str(cell_val).strip() for i, cell_val in enumerate(row_data)
                            if _date_main_disp is None or i != _date_main_disp)
        if not is_row_active and self.sheet_frozen:
            _, _date_frz_disp = self._col_to_disp("วันที่ขอราคา")
            frozen_row = self.sheet_frozen.get_row_data(row_idx)
            is_row_active = any(str(v).strip() for i, v in enumerate(frozen_row)
                                if _date_frz_disp is None or i != _date_frz_disp)

        current_date = get_str("วันที่ขอราคา")
        if is_row_active and not current_date:
            now = datetime.now()
            thai_year = (now.year + 543) % 100
            set_val("วันที่ขอราคา", f"{now.day:02d}/{now.month:02d}/{thai_year}", is_text=True)
        elif not is_row_active and current_date:
            set_val("วันที่ขอราคา", "", is_text=True)

        product_name = get_str("รายการสินค้า")
        if product_name in self.product_sku_map:
            if get_str("Product SKU.") != self.product_sku_map[product_name]:
                set_val("Product SKU.", self.product_sku_map[product_name], is_text=True)
            mapped_category = self.product_category_map.get(product_name, "")
            if get_str("หมวด") != mapped_category:
                set_val("หมวด", mapped_category, is_text=True)
        elif not product_name:
            if get_str("Product SKU.") != "": set_val("Product SKU.", "", is_text=True)
            if get_str("หมวด") != "": set_val("หมวด", "", is_text=True)

        supplier_name = get_str("ชื่อ Supplier")
        if supplier_name:
            if get_str("ชื่อ Supplier2") != supplier_name:
                set_val("ชื่อ Supplier2", supplier_name, is_text=True)
            sup_id = self.supplier_code_map.get(supplier_name, "")
            if get_str("Sup ID.") != sup_id:
                set_val("Sup ID.", sup_id, is_text=True)
        else:
            if get_str("ชื่อ Supplier2") != "": set_val("ชื่อ Supplier2", "", is_text=True)
            if get_str("Sup ID.") != "": set_val("Sup ID.", "", is_text=True)

        for col_percent in ["WIN RATE %", "Markup Guide (%)"]:
            val_str = get_str(col_percent)
            if val_str and not val_str.endswith("%"):
                num_str = "".join([c for c in val_str if c.isdigit() or c == '.'])
                if num_str: set_val(col_percent, f"{num_str}%", is_text=True)

        qty = get_val("จำนวน")
        weight_per_unit = get_val("น้ำหนัก/เส้น")
        cost_per_unit = get_val("ต้นทุน/เส้น")

        if qty == 0 and weight_per_unit == 0 and cost_per_unit == 0:
            auto_cols_to_clear = [
                "น้ำหนักรวม (Kg.)", "ทุน/กก.", "ทุนรวม", "ส่วนลด 1 (%)", "ทุน/เส้น หลังส่วนลด 1",
                "ส่วนลด 2 (%)", "ทุน/เส้น หลังส่วนลด 2", "ต้นทุน/กก. (ไม่รวมย้าย)",
                "ต้นทุน/เส้น (ไม่รวมย้าย)", "ต้นทุนรวม (ไม่รวมย้าย)", "ค่าย้าย/เส้น",
                "ต้นทุน/กก. (รวมย้าย)", "ต้นทุน/เส้น (รวมย้าย)", "ต้นทุนรวม (รวมย้าย)",
                "Markup/กก.", "Markup/เส้น", "ทุน+Markup/กก.", "ทุน+Markup/เส้น", "ต้นทุนรวม+Markup",
                "ค่าส่ง / เส้น", "น้ำหนัก/เส้น 2", "ราคาขาย / กก.", "ราคาขาย / เส้น",
                "Vat. / เส้น", "ราคาขาย/เส้น + Vat.", "ราคาขาย รวม", "Vat. รวม", "ราคาขาย รวม + Vat."
            ]
            for col in auto_cols_to_clear:
                set_val(col, "", is_text=True)
            return

        total_weight = weight_per_unit * qty
        set_val("น้ำหนักรวม (Kg.)", total_weight)
        cost_per_kg = (cost_per_unit / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ทุน/กก.", cost_per_kg)
        total_cost = cost_per_unit * qty
        set_val("ทุนรวม", total_cost)

        discount1_baht = get_val("ส่วนลด 1 (บาท)")
        discount1_pct = (discount1_baht / cost_per_unit) if cost_per_unit > 0 else 0
        set_val("ส่วนลด 1 (%)", f"{discount1_pct*100:.2f}%" if discount1_pct > 0 else "", is_text=True)
        cost_after_d1 = cost_per_unit - discount1_baht
        set_val("ทุน/เส้น หลังส่วนลด 1", cost_after_d1)

        discount2_baht = get_val("ส่วนลด 2 (บาท)")
        discount2_pct = (discount2_baht / cost_after_d1) if cost_after_d1 > 0 else 0
        set_val("ส่วนลด 2 (%)", f"{discount2_pct*100:.2f}%" if discount2_pct > 0 else "", is_text=True)
        cost_after_d2 = cost_after_d1 - discount2_baht
        set_val("ทุน/เส้น หลังส่วนลด 2", cost_after_d2)

        cost_no_move_per_kg = (cost_after_d2 / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ต้นทุน/เส้น (ไม่รวมย้าย)", cost_after_d2)
        set_val("ต้นทุน/กก. (ไม่รวมย้าย)", cost_no_move_per_kg)
        set_val("ต้นทุนรวม (ไม่รวมย้าย)", cost_after_d2 * qty)

        moving_cost = get_val("ค่าย้าย (ซื้อ)")
        moving_cost_per_unit = (moving_cost / qty) if qty > 0 else 0
        set_val("ค่าย้าย/เส้น", moving_cost_per_unit)

        cost_with_move = cost_after_d2 + moving_cost_per_unit
        cost_with_move_per_kg = (cost_with_move / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ต้นทุน/เส้น (รวมย้าย)", cost_with_move)
        set_val("ต้นทุน/กก. (รวมย้าย)", cost_with_move_per_kg)
        set_val("ต้นทุนรวม (รวมย้าย)", cost_with_move * qty)

        markup_pct = get_val("Markup Guide (%)") / 100.0
        set_val("Markup/กก.", cost_with_move_per_kg * markup_pct)
        set_val("Markup/เส้น", cost_with_move * markup_pct)

        cost_markup_per_kg = cost_with_move_per_kg + cost_with_move_per_kg * markup_pct
        cost_markup_per_unit = cost_with_move + cost_with_move * markup_pct
        set_val("ทุน+Markup/กก.", cost_markup_per_kg)
        set_val("ทุน+Markup/เส้น", cost_markup_per_unit)
        set_val("ต้นทุนรวม+Markup", cost_markup_per_unit * qty)

        shipping_sell = get_val("ค่าส่ง (ขาย)")
        shipping_per_unit = (shipping_sell / qty) if qty > 0 else 0
        set_val("ค่าส่ง / เส้น", shipping_per_unit)
        set_val("น้ำหนัก/เส้น 2", weight_per_unit)

        sell_price = cost_markup_per_unit + shipping_per_unit
        sell_price_per_kg = (sell_price / weight_per_unit) if weight_per_unit > 0 else 0
        set_val("ราคาขาย / เส้น", sell_price)
        set_val("ราคาขาย / กก.", sell_price_per_kg)

        vat_per_unit = sell_price * 0.07
        set_val("Vat. / เส้น", vat_per_unit)
        set_val("ราคาขาย/เส้น + Vat.", sell_price * 1.07)

        sell_total = sell_price * qty
        set_val("ราคาขาย รวม", sell_total)
        set_val("Vat. รวม", sell_total * 0.07)
        set_val("ราคาขาย รวม + Vat.", sell_total * 1.07)

        if getattr(self, '_header_filter_values', None):
            if hasattr(self, '_fix_render_job'):
                self.after_cancel(self._fix_render_job)
            self._fix_render_job = self.after(150, self._apply_header_filters)

    # ================================================================== #
    def _show_short_note_popup(self):
        """Short Note popup — เลือก SO + filter เฉพาะ row ที่ Select = ✓"""
        from datetime import datetime
        from collections import OrderedDict
        import tkinter as tk
        from customtkinter import (CTkToplevel, CTkLabel, CTkFont, CTkButton,
                                   CTkTextbox, CTkEntry, CTkFrame,
                                   CTkSegmentedButton)

        # ── 1. รวบรวมข้อมูลทั้งหมดจากตาราง ────────────────────────────────
        total_rows = self.sheet.get_total_rows()
        all_rows = []
        so_order = []
        so_seen  = set()

        for r in range(total_rows):
            product = str(self._sheet_get(r, "รายการสินค้า") or "").strip()
            if not product:
                continue
            so = str(self._sheet_get(r, "Sale Order No.") or "").strip()
            if so and so not in so_seen:
                so_seen.add(so)
                so_order.append(so)
            all_rows.append({
                "so":           so,
                "product":      product,
                "note1":        str(self._sheet_get(r, "หมายเหตุ (ความยาว, OD)") or "").strip(),
                "note2":        str(self._sheet_get(r, "หมายเหตุ") or "").strip(),
                "qty":          str(self._sheet_get(r, "จำนวน") or "").strip(),
                "unit_price":   str(self._sheet_get(r, "ราคาขาย / เส้น") or "").strip(),
                "row_total":    str(self._sheet_get(r, "ราคาขาย รวม") or "").strip(),
                "weight_unit":  str(self._sheet_get(r, "น้ำหนัก/เส้น") or "").strip(),
                "weight_total": str(self._sheet_get(r, "น้ำหนักรวม (Kg.)") or "").strip(),
                "dest":         str(self._sheet_get(r, "ปลายทาง") or "").strip(),
                "order_no":     str(self._sheet_get(r, "Order No.") or "").strip(),
                "qt":           str(self._sheet_get(r, "QT") or "").strip(),
                "priority":     str(self._sheet_get(r, "PRIORITY") or "").strip(),
                "win_rate":     str(self._sheet_get(r, "WIN RATE %") or "").strip(),
                "select":       str(self._sheet_get(r, "Select") or "").strip(),
            })

        if not all_rows:
            from tkinter import messagebox
            messagebox.showinfo("Short Note", "ไม่มีรายการสินค้าในตาราง", parent=self)
            return

        # ── 2. formatters ──────────────────────────────────────────────────
        def _fmt_weight(s):
            try:
                v = float(str(s).replace(",", ""))
                return f"นน.{v:g}+-"
            except Exception:
                return f"นน.{s}+-" if s else ""

        def _fmt_price(s):
            try:
                v = float(str(s).replace(",", ""))
                if v == int(v):
                    return f"@{int(v):,}+"
                return f"@{v:,.2f}+"
            except Exception:
                return f"@{s}+" if s else ""

        # ── 3. สร้าง note text ─────────────────────────────────────────────
        def _build_note(chosen_so: str, validity_days: str, yod_status: str) -> str:
            items = [r for r in all_rows
                     if r["so"] == chosen_so and r["select"] in ("✔", "เทียบเพื่อชุบ ✔")]
            if not items:
                return f"⚠️  ไม่มีรายการที่ติ๊ก Select ใน {chosen_so}"

            today = datetime.now().strftime("%d/%m/%Y")
            days  = validity_days.strip() or "1"

            # group by (ปลายทาง, หมายเหตุ) = จุดรับ + สถานะ
            groups = OrderedDict()
            for it in items:
                groups.setdefault((it["dest"], it["note2"]), []).append(it)

            grand_bill   = 0.0
            grand_weight = 0.0
            total_items  = 0

            lines = [
                f"แจ้งราคาขาย  SO: {chosen_so}",
                f"วันที่: {today}",
                "─" * 16,
            ]

            for (dest, status), grp_items in groups.items():
                for it in grp_items:
                    spec     = it["note1"]
                    prod_line = f"{it['product']} {spec}".strip() if spec else it["product"]
                    wt_str    = _fmt_weight(it["weight_unit"])
                    price_str = _fmt_price(it["unit_price"])
                    parts = [prod_line]
                    if it["qty"]:
                        parts.append(f"{it['qty']} เส้น")
                    if wt_str:
                        parts.append(wt_str)
                    if price_str:
                        parts.append(price_str)
                    lines.append("  ".join(parts))
                    total_items += 1
                    try:
                        grand_bill += float(it["row_total"].replace(",", ""))
                    except Exception:
                        pass
                    try:
                        grand_weight += float(it["weight_total"].replace(",", ""))
                    except Exception:
                        pass

            # ── footer เดียวหลังรายการสินค้าทั้งหมด ──────────────────────
            # รวม unique จุดรับ และ หมายเหตุ จากทุก item
            dest_vals = list(dict.fromkeys(
                it["dest"] for it in items if it["dest"]
            ))
            dest_str = ", ".join(dest_vals) if dest_vals else "-"
            note1_vals = list(dict.fromkeys(
                it["note1"] for it in items if it["note1"]
            ))
            note1_str = ", ".join(note1_vals) if note1_vals else ""
            lines.append("─" * 15)
            lines.append("สถานะ :")
            lines.append(f"ยืนราคา {days} วัน")
            lines.append(f"จุดรับ : {dest_str}")
            lines.append(f"หมายเหตุ : {note1_str}")
            lines.append("")

            first     = items[0]
            order_no  = first.get("order_no", "") or "-"
            qt        = first.get("qt", "")
            priority  = first.get("priority", "") or "-"
            win_rate  = (first.get("win_rate", "") or "-").rstrip("%").strip()
            order_pur = f"{order_no} {qt}".strip() if qt else order_no
            lines += [
                "─" * 15,
                f"สถานะยอด      : {yod_status}",
                f"Order Pur     : {order_pur}",
                f"Temp          : {priority}",
                f"WIN           : {win_rate}%",
                f"Order Sale    : {chosen_so}",
            ]
            return "\n".join(lines)

        # ── 4. สร้าง Popup ─────────────────────────────────────────────────
        pop = CTkToplevel(self)
        pop.title("📋 Short Note")
        pop.resizable(True, True)
        pop.grab_set()
        pop.update_idletasks()
        w, h = 700, 740
        root_x = self.winfo_rootx()
        root_y = self.winfo_rooty()
        root_w = self.winfo_width()
        root_h = self.winfo_height()
        cx = root_x + (root_w - w) // 2
        cy = root_y + (root_h - h) // 2
        pop.geometry(f"{w}x{h}+{cx}+{cy}")
        pop.grid_columnconfigure(0, weight=1)
        pop.grid_rowconfigure(3, weight=1)   # txtbox
        pop.grid_rowconfigure(4, minsize=60) # buttons

        # row 0 — Header
        CTkLabel(pop, text="📋 Short Note",
                 font=CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, padx=20, pady=(16, 4), sticky="w")

        # row 1 — SO Selector
        sel_frame = CTkFrame(pop, fg_color="#F1F5F9", corner_radius=8,
                             border_width=1, border_color="#CBD5E1")
        sel_frame.grid(row=1, column=0, padx=20, pady=(0, 6), sticky="ew")
        sel_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(sel_frame, text="Sale Order:",
                 font=CTkFont(size=12), text_color="#374151").grid(
            row=0, column=0, padx=(12, 6), pady=8)

        search_var = tk.StringVar()
        search_entry = CTkEntry(sel_frame, textvariable=search_var,
                                placeholder_text="พิมพ์เพื่อค้นหา SO...",
                                height=32, font=CTkFont(size=13))
        search_entry.grid(row=0, column=1, padx=(0, 8), pady=8, sticky="ew")

        list_frame = tk.Frame(sel_frame, bg="#F1F5F9")
        list_frame.grid(row=1, column=0, columnspan=2, padx=8, pady=(0, 8), sticky="ew")

        sb = tk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(list_frame, font=("Tahoma", 11), height=4,
                             selectbackground="#3B82F6", selectforeground="white",
                             relief="flat", borderwidth=1, highlightthickness=0,
                             yscrollcommand=sb.set, exportselection=False)
        sb.config(command=listbox.yview)
        sb.pack(side="right", fill="y")
        listbox.pack(side="left", fill="x", expand=True)

        # row 2 — Settings (ยืนราคา + สถานะยอด)
        cfg_frame = CTkFrame(pop, fg_color="#F8FAFC", corner_radius=8,
                             border_width=1, border_color="#E2E8F0")
        cfg_frame.grid(row=2, column=0, padx=20, pady=(0, 6), sticky="ew")

        CTkLabel(cfg_frame, text="ยืนราคา:", font=CTkFont(size=12),
                 text_color="#374151").pack(side="left", padx=(12, 4), pady=8)
        validity_var = tk.StringVar(value="1")
        CTkEntry(cfg_frame, textvariable=validity_var, width=48,
                 height=28, font=CTkFont(size=12), justify="center").pack(
            side="left", pady=8)
        CTkLabel(cfg_frame, text="วัน", font=CTkFont(size=12),
                 text_color="#374151").pack(side="left", padx=(4, 24), pady=8)

        CTkLabel(cfg_frame, text="สถานะยอด:", font=CTkFont(size=12),
                 text_color="#374151").pack(side="left", padx=(0, 6), pady=8)
        yod_var = tk.StringVar(value="C")
        CTkSegmentedButton(cfg_frame, values=["C", "เบิกจ่าย"],
                           variable=yod_var, font=CTkFont(size=12),
                           width=160).pack(side="left", pady=8)

        # row 3 — Note display
        txt = CTkTextbox(pop, font=CTkFont(family="Courier New", size=15),
                         wrap="none", corner_radius=8)
        txt.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="nsew")
        txt.insert("0.0", "← เลือก Sale Order ด้านบนเพื่อสร้าง Short Note")
        txt.configure(state="disabled")

        _current_note = {"text": ""}
        _ref = {}

        def _refresh_note(*_):
            sel = listbox.curselection()
            if not sel:
                return
            chosen = listbox.get(sel[0])
            note = _build_note(chosen, validity_var.get(), yod_var.get())
            _current_note["text"] = note
            txt.configure(state="normal")
            txt.delete("0.0", "end")
            txt.insert("0.0", note)
            txt.configure(state="disabled")
            if _ref.get("copy_btn"):
                _ref["copy_btn"].configure(state="normal", fg_color="#059669")

        listbox.bind("<<ListboxSelect>>", _refresh_note)
        validity_var.trace_add("write", _refresh_note)
        yod_var.trace_add("write", _refresh_note)

        def _filter_so(*_):
            term = search_var.get().strip().lower()
            listbox.delete(0, "end")
            for so in so_order:
                if term in so.lower():
                    listbox.insert("end", so)
            if listbox.size() > 0:
                listbox.selection_set(0)
                _refresh_note()

        search_var.trace_add("write", _filter_so)
        _filter_so()

        # row 4 — Bottom buttons
        btn_row = CTkFrame(pop, fg_color="transparent")
        btn_row.grid(row=4, column=0, padx=20, pady=(0, 16), sticky="ew")
        btn_row.grid_columnconfigure(0, weight=1)

        def _copy():
            note = _current_note["text"]
            if not note:
                return
            pop.clipboard_clear()
            pop.clipboard_append(note)
            copy_btn.configure(text="✅ คัดลอกแล้ว!", fg_color="#047857")
            pop.after(2000, lambda: copy_btn.configure(
                text="📋 Copy to Clipboard", fg_color="#059669"))

        copy_btn = CTkButton(btn_row, text="📋 Copy to Clipboard",
                             fg_color="#059669", hover_color="#047857",
                             font=CTkFont(size=13, weight="bold"),
                             height=36, state="disabled", command=_copy)
        copy_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        _ref["copy_btn"] = copy_btn
        CTkButton(btn_row, text="ปิด", fg_color="#6B7280", hover_color="#4B5563",
                  height=36, width=80, command=pop.destroy).grid(
            row=0, column=1, sticky="e")

        search_entry.focus_set()

    def _add_new_row(self):
        if not HAS_TKSHEET:
            return
        self.sheet.insert_row([""] * (len(self.columns) - self.frozen_col_count))
        self._apply_formatting(col_offset=self.frozen_col_count)
        total_rows = self.sheet.get_total_rows()
        self.sheet.see(total_rows - 1, 0)
        self.sheet.select_cell(total_rows - 1, 0)
        self.sheet.redraw()
        if self.sheet_frozen is not None:
            self.sheet_frozen.insert_row([""] * self.frozen_col_count)
            self.sheet_frozen.redraw()

    def _insert_selected_row(self):
        if not HAS_TKSHEET:
            return

        insert_idx = None
        
        try:
            curr = self.sheet.get_currently_selected()
            if curr:
                insert_idx = int(curr[0])
        except Exception:
            pass

        if insert_idx is None:
            try:
                selected_rows = self.sheet.get_selected_rows()
                if selected_rows:
                    insert_idx = min(int(r) for r in selected_rows)
            except Exception:
                pass

        if insert_idx is None:
            try:
                selected_cells = self.sheet.get_selected_cells()
                if selected_cells:
                    insert_idx = min(int(r) for r, c in selected_cells)
            except Exception:
                pass

        if insert_idx is None:
            if self.sheet_frozen:
                try:
                    curr = self.sheet_frozen.get_currently_selected()
                    if curr:
                        insert_idx = int(curr[0])
                except Exception:
                    pass

        if insert_idx is None:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการแทรกก่อน", parent=self)
            return

        empty_main_row = [""] * (len(self.columns) - self.frozen_col_count)
        self.sheet.insert_row(empty_main_row, idx=insert_idx)

        if self.sheet_frozen is not None:
            empty_frozen_row = [""] * self.frozen_col_count
            self.sheet_frozen.insert_row(empty_frozen_row, idx=insert_idx)
            self.sheet_frozen.redraw()

        self._apply_formatting(col_offset=self.frozen_col_count)
        self.sheet.see(insert_idx, 0)
        self.sheet.deselect("all")  # ← เพิ่ม: ล้าง selection ให้ user คลิกเองใหม่
        if self.sheet_frozen:
            self.sheet_frozen.deselect("all")
        self.sheet.redraw()
        if self.sheet_frozen:
            self.after(100, self._sync_vertical_scroll)
    
    def _copy_selected_rows(self, event=None):
        """Copy ทั้ง row รวม frozen cols"""
        try:
            # เก็บ rows ที่ selected จาก main sheet
            selected = self.sheet.get_selected_rows()
            if not selected and self.sheet_frozen:
                selected = set()
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    selected = {int(curr[0])}
            if not selected:
                return
            
            rows_data = []
            for r in sorted(selected):
                frozen_part = []
                main_part = list(self.sheet.get_row_data(r))
                if self.sheet_frozen:
                    frozen_part = list(self.sheet_frozen.get_row_data(r))
                rows_data.append(frozen_part + main_part)
            
            self._clipboard_rows = rows_data
            self.save_status_label.configure(
                text=f"📋 คัดลอก {len(rows_data)} บรรทัดแล้ว (Ctrl+V เพื่อวาง)",
                text_color="#6366F1"
            )
        except Exception as e:
            print(f"_copy_selected_rows error: {e}")

    def _paste_selected_rows(self, event=None):
        """Paste row ที่ copy ไว้ลงที่ row ที่เลือกอยู่"""
        try:
            if not hasattr(self, '_clipboard_rows') or not self._clipboard_rows:
                return
            
            # หา target row
            target_idx = None
            curr = self.sheet.get_currently_selected()
            if curr:
                target_idx = int(curr[0])
            elif self.sheet_frozen:
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    target_idx = int(curr[0])
            
            if target_idx is None:
                messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดปลายทางก่อน Ctrl+V", parent=self)
                return
            
            col_offset = self.frozen_col_count
            for i, row_data in enumerate(self._clipboard_rows):
                r = target_idx + i
                if r >= self.sheet.get_total_rows():
                    break
                # เขียนส่วน frozen
                if self.sheet_frozen and col_offset > 0:
                    for c in range(col_offset):
                        val = row_data[c] if c < len(row_data) else ""
                        self.sheet_frozen.set_cell_data(r, c, val, redraw=False)
                # เขียนส่วน main
                for c in range(self.sheet.get_total_columns()):
                    real_c = c + col_offset
                    val = row_data[real_c] if real_c < len(row_data) else ""
                    self.sheet.set_cell_data(r, c, val, redraw=False)
                self._auto_calculate_sheet(r)
            
            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
            
            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
            
        except Exception as e:
            print(f"_paste_selected_rows error: {e}")

    def _delete_selected_rows(self):
        if not HAS_TKSHEET: return
        
        selected_rows = set()
        
        try:
            rows = self.sheet.get_selected_rows()
            if rows:
                selected_rows.update(rows)
        except Exception:
            pass
        
        try:
            cells = self.sheet.get_selected_cells()
            if cells:
                selected_rows.update(r for r, c in cells)
        except Exception:
            pass
        
        if self.sheet_frozen:
            try:
                rows = self.sheet_frozen.get_selected_rows()
                if rows:
                    selected_rows.update(rows)
            except Exception:
                pass
            
            try:
                cells = self.sheet_frozen.get_selected_cells()
                if cells:
                    selected_rows.update(r for r, c in cells)
            except Exception:
                pass
            
            try:
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    selected_rows.add(int(curr[0]))
            except Exception:
                pass
        
        try:
            curr = self.sheet.get_currently_selected()
            if curr:
                selected_rows.add(int(curr[0]))
        except Exception:
            pass

        if not selected_rows:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกบรรทัดที่ต้องการลบก่อน", parent=self)
            return
        
        if messagebox.askyesno("ยืนยัน", f"ต้องการลบข้อมูลจำนวน {len(selected_rows)} บรรทัด ใช่หรือไม่?", parent=self):
            rows_list = sorted(selected_rows, reverse=True)
            self.sheet.delete_rows(rows_list)
            if self.sheet_frozen:
                self.sheet_frozen.delete_rows(rows_list)
            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
            
            # ← แก้หลัก: re-sync scroll หลังลบ เพราะ row count เปลี่ยน
            # และต้อง re-bind frozen scroll ด้วย เพราะ sync loop อาจหลุด
            if self.sheet_frozen:
                self.after(100, self._sync_vertical_scroll)
            
            if self.auto_save_job_id:
                self.after_cancel(self.auto_save_job_id)
            self.auto_save_job_id = self.after(1500, lambda: self._save_to_db(show_msg=False))
            self.save_status_label.configure(text="⏳ รอการบันทึก...", text_color="#D97706")
    # ================================================================== #

    def _highlight_selected_cells(self):
        """ระบายสีไฮไลท์ช่องที่ถูกเลือก"""
        if not HAS_TKSHEET: return
        
        # เปิดหน้าต่างให้เลือกสี
        color_tuple = colorchooser.askcolor(title="เลือกสีสำหรับไฮไลท์ช่อง", parent=self)
        if not color_tuple or not color_tuple[1]: return
        
        color_code = color_tuple[1]
        
        # tksheet v7: get_selected_cells() คืน DATA index โดยตรง
        # ไม่ต้องแปลงผ่าน visible_main เพราะ highlight_cells ก็รับ data index เช่นกัน
        try:
            cells = self.sheet.get_selected_cells()
            for r, c in cells:
                self.sheet.highlight_cells(row=r, column=c, bg=color_code)
        except Exception as e:
            print(f"Highlight main error: {e}")

        if self.sheet_frozen:
            try:
                f_cells = self.sheet_frozen.get_selected_cells()
                for r, c in f_cells:
                    self.sheet_frozen.highlight_cells(row=r, column=c, bg=color_code)
            except Exception as e:
                print(f"Highlight frozen error: {e}")
            
        self.sheet.redraw()
        if self.sheet_frozen: self.sheet_frozen.redraw()
        
        # 🔹 บันทึกสีลงฐานข้อมูลทันที
        self._save_user_settings()

    def _clear_highlight_selected_cells(self):
        """ล้างสีไฮไลท์ช่องที่ถูกเลือก"""
        if not HAS_TKSHEET: return
        
        try:
            cells = self.sheet.get_selected_cells()
            for r, c in cells:
                self.sheet.dehighlight_cells(row=r, column=c)
        except Exception as e:
            print(f"Clear highlight main error: {e}")

        if self.sheet_frozen:
            try:
                f_cells = self.sheet_frozen.get_selected_cells()
                for r, c in f_cells:
                    self.sheet_frozen.dehighlight_cells(row=r, column=c)
            except Exception as e:
                print(f"Clear highlight frozen error: {e}")
            
        self.sheet.redraw()
        if self.sheet_frozen: self.sheet_frozen.redraw()
        
        # 🔹 บันทึกการล้างสีลงฐานข้อมูลทันที
        self._save_user_settings()

    def _load_from_db(self, *args):
        if not HAS_TKSHEET: return
        month_val = self.month_var.get()
        year_val = self.year_var.get()
        conn = self.app_container.get_connection()
        try:
            columns_sql = ", ".join([f'"{col.replace("%", "%%")}"' for col in self.columns])
            query = f'SELECT {columns_sql} FROM cost_benchmarks WHERE benchmark_month = %s AND benchmark_year = %s AND created_by = %s ORDER BY id ASC'
            df = pd.read_sql_query(query, conn, params=(month_val, year_val, self.current_user))

            self.sheet.bind("<<SheetModified>>", lambda e: None)
            total_cols_main = len(self.columns) - self.frozen_col_count
            col_offset = self.frozen_col_count

            if df.empty:
                data_list = [[""] * len(self.columns) for _ in range(1000)]
            else:
                df = df.fillna("")
                data_list = df.values.tolist()
                while len(data_list) < 1000:
                    data_list.append([""] * len(self.columns))

            # ปรับจำนวนแถวให้พอดี
            current_rows = self.sheet.get_total_rows()
            if len(data_list) > current_rows:
                self.sheet.insert_rows(len(data_list) - current_rows)
                if self.sheet_frozen:
                    self.sheet_frozen.insert_rows(len(data_list) - current_rows)

            # ← เขียนข้อมูลลงเซลล์จริงๆ (ส่วนนี้หายไป)
            for r, row in enumerate(data_list):
                if self.sheet_frozen and col_offset > 0:
                    for c in range(col_offset):
                        val_str = str(row[c]).strip() if c < len(row) and row[c] is not None else ""
                        self.sheet_frozen.set_cell_data(r, c, val_str, redraw=False)
                for c in range(total_cols_main):
                    real_c = c + col_offset
                    val_str = str(row[real_c]).strip() if real_c < len(row) and row[real_c] is not None else ""
                    self.sheet.set_cell_data(r, c, val_str, redraw=False)

            # ล้างแถวส่วนเกิน
            for r in range(len(data_list), self.sheet.get_total_rows()):
                for c in range(total_cols_main):
                    self.sheet.set_cell_data(r, c, "", redraw=False)
                if self.sheet_frozen:
                    for c in range(col_offset):
                        self.sheet_frozen.set_cell_data(r, c, "", redraw=False)

            self.sheet.redraw()
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
            self._apply_saved_highlights()
            self._history = []
            self._history_idx = -1
            self._push_undo()

            self.after(50, lambda: self.sheet.MT.focus_set())

        except Exception as e:
            messagebox.showerror("Error", f"โหลดข้อมูลล้มเหลว: {e}", parent=self)
        finally:
            self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
            if conn: self.app_container.release_connection(conn)
    
    def _hide_frozen_scrollbars(self):
        """ซ่อน scrollbar ทั้งหมดของ frozen sheet และ sync row height"""
        if not self.sheet_frozen:
            return
        try:
            # วิธีที่แรง — resize frozen sheet ให้สูงเต็ม relheight แต่ลบพื้นที่ scrollbar ออก
            # โดยการ configure internal frame
            for attr in ['_y_scrollbar', 'yscrollbar', 'y_scrollbar',
                        '_x_scrollbar', 'xscrollbar', 'x_scrollbar']:
                try:
                    sb = getattr(self.sheet_frozen, attr, None)
                    if sb:
                        sb.grid_remove()
                        sb.place_forget()
                        sb.pack_forget()
                except Exception:
                    pass

            # ซ่อนผ่าน children
            for widget in self.sheet_frozen.winfo_children():
                if widget.winfo_class() == "Scrollbar":
                    widget.grid_remove()
                    try: widget.place_forget()
                    except: pass
                    try: widget.pack_forget()
                    except: pass

            # ใช้ tksheet API
            for canvas_name in ["x_scrollbar", "y_scrollbar", "top_left"]:
                try:
                    self.sheet_frozen.hide(canvas=canvas_name)
                except Exception:
                    pass

        except Exception:
            pass

        # sync row height
        try:
            rh = int(30 * (self.zoom_level / 11.0))
            if self.sheet_frozen:
                total_rows = max(self.sheet_frozen.get_total_rows(), self.sheet.get_total_rows())
            else:
                total_rows = self.sheet.get_total_rows()
                
            for r in range(total_rows):
                if self.sheet_frozen:
                    self.sheet_frozen.row_height(r, rh)
                self.sheet.row_height(r, rh)
                
            if self.sheet_frozen:
                self.sheet_frozen.redraw()
            self.sheet.redraw()
        except Exception:
            pass

    def _save_to_db(self, show_msg=True):
        col_offset = self.frozen_col_count
        raw_main = self.sheet.get_sheet_data()
        raw_frozen = self.sheet_frozen.get_sheet_data() if self.sheet_frozen else []

        data = []
        n_rows = max(len(raw_main), len(raw_frozen))
        date_col_idx = self.columns.index("วันที่ขอราคา")  # ← เพิ่มบรรทัดนี้
        
        for i in range(n_rows):
            left = raw_frozen[i] if i < len(raw_frozen) else [""] * col_offset
            right = raw_main[i] if i < len(raw_main) else [""] * (len(self.columns) - col_offset)
            full_row = left + right
            
            # ← แก้ตรงนี้
            is_active = any(
                str(cell).strip() not in ("", "%", "None", "nan")
                for j, cell in enumerate(full_row)
                if j != date_col_idx
            )
            
            if is_active:
                data.append([str(cell).strip() if cell is not None else "" for cell in full_row])
    

        df = pd.DataFrame(data, columns=self.columns)
        if df.empty:
            if show_msg: messagebox.showinfo("แจ้งเตือน", "ไม่มีข้อมูลให้บันทึก", parent=self)
            return

        month_val = self.month_var.get()
        year_val = self.year_var.get()
        df = df.replace(r'^\s*$', None, regex=True)

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    "DELETE FROM cost_benchmarks WHERE benchmark_month = %s AND benchmark_year = %s AND created_by = %s",
                    (month_val, year_val, self.current_user)
                )
                columns_sql = ", ".join([f'"{col.replace("%", "%%")}"' for col in self.columns]) + ", benchmark_month, benchmark_year, created_by"
                values = [tuple(row) + (month_val, year_val, self.current_user) for row in df.to_numpy()]
                insert_query = f"INSERT INTO cost_benchmarks ({columns_sql}) VALUES %s"
                psycopg2.extras.execute_values(cursor, insert_query, values)
            conn.commit()
            current_time = datetime.now().strftime("%H:%M:%S")
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text=f"✅ บันทึกล่าสุด: {current_time}", text_color="#16A34A")
            if show_msg:
                messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล {len(df)} รายการเรียบร้อยแล้ว!", parent=self)

            # ── Auto-update supplier category จากข้อมูลที่เพิ่งบันทึก ──────────
            self.after(500, self._update_supplier_categories_from_benchmark)

        except Exception as e:
            if conn: conn.rollback()
            import traceback; traceback.print_exc()
            if hasattr(self, 'save_status_label'):
                self.save_status_label.configure(text="❌ บันทึกผิดพลาด กรุณาลองใหม่", text_color="#DC2626")
            if show_msg:
                messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _update_supplier_categories_from_benchmark(self):
        """
        Auto-update suppliers.category โดยดูจาก cost_benchmarks
        หมวดที่ Supplier นั้นๆ ถูกจับคู่บ่อยที่สุด → เซ็ตเป็น category
        เฉพาะแถวที่ยัง category ว่างอยู่เท่านั้น (ไม่เขียนทับที่มีแล้ว)
        """
        # หา column ชื่อ Supplier และ หมวด จาก self.columns
        sup_col  = next((c for c in self.columns if 'Supplier' in c
                         and 'Supplier2' not in c and 'ID' not in c
                         and 'Sup' not in c[:3]), None)
        cat_col  = next((c for c in self.columns if c == 'หมวด'), None)
        if not sup_col or not cat_col:
            return

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ดึง mapping: supplier_name → หมวดที่ใช้บ่อยที่สุด
                cursor.execute(f"""
                    SELECT sup_name, top_cat
                    FROM (
                        SELECT
                            "{sup_col}"  AS sup_name,
                            "{cat_col}"  AS top_cat,
                            COUNT(*)     AS cnt,
                            ROW_NUMBER() OVER (
                                PARTITION BY "{sup_col}"
                                ORDER BY COUNT(*) DESC
                            ) AS rn
                        FROM cost_benchmarks
                        WHERE "{sup_col}" IS NOT NULL AND "{sup_col}" != ''
                          AND "{cat_col}" IS NOT NULL AND "{cat_col}" != ''
                        GROUP BY "{sup_col}", "{cat_col}"
                    ) ranked
                    WHERE rn = 1
                """)
                mappings = cursor.fetchall()

                updated = 0
                for sup_name, category in mappings:
                    cursor.execute("""
                        UPDATE suppliers
                        SET    category = %s
                        WHERE  supplier_name = %s
                          AND  (category IS NULL OR category = '')
                    """, (category, sup_name))
                    updated += cursor.rowcount

            conn.commit()
            if updated > 0:
                print(f"[SSL] Auto-updated category: {updated} suppliers")
        except Exception as e:
            conn.rollback()
            print(f"[SSL] _update_supplier_categories_from_benchmark error: {e}")
        finally:
            self.app_container.release_connection(conn)

    # ================================================================== #
    # FORMULA BAR
    # ================================================================== #
    def _num2col(self, n):
        string = ""
        n += 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def _on_formula_focus_in(self, event):
        try:
            cells = self.sheet.get_selected_cells()
            if cells:
                self.target_formula_cell = list(cells)[0]
        except Exception:
            pass

    def _on_sheet_click_for_formula(self, event=None):
        try:
            current_text = self.formula_entry.get()
            if current_text.startswith("="):
                cells = self.sheet.get_selected_cells()
                if not cells: return
                row, col = list(cells)[0]
                cell_val = self.sheet.get_cell_data(row, col)
                if not cell_val or str(cell_val).strip() == "":
                    val_to_insert = "0"
                else:
                    val_to_insert = str(cell_val).replace(',', '').replace('%', '').strip()
                    try:
                        float(val_to_insert)
                    except ValueError:
                        val_to_insert = "0"
                self.formula_entry.insert(tk.END, val_to_insert)
                self.formula_entry.focus()
                self.formula_entry.icursor(tk.END)
        except Exception:
            pass

    def _apply_formula_from_bar(self, event=None):
        if not self.target_formula_cell:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิกเลือกช่องปลายทางในตารางก่อนเริ่มพิมพ์สูตร")
            return
        try:
            row, col = self.target_formula_cell
            formula = self.formula_entry.get()
            self.sheet.set_cell_data(row, col, formula)
            self.formula_entry.delete(0, tk.END)
            self.target_formula_cell = None
            self.sheet.select_cell(row, col)
            self._auto_calculate_sheet(row)
        except Exception as e:
            messagebox.showerror("Error", f"สูตรผิดพลาด: {e}")

    # ================================================================== #
    # COLUMN HIDE / SHOW / COLOR
    # ================================================================== #
    def _get_real_col_indices(self):
        real_cols = set()
        col_offset = self.frozen_col_count if self.sheet_frozen is not None else 0

        # 1. เช็คจาก Sheet หลักฝั่งขวา
        try:
            sel_cols = self.sheet.get_selected_columns()
            if sel_cols:
                for c in sel_cols: real_cols.add(int(c) + col_offset)
                
            sel_cells = self.sheet.get_selected_cells()
            if sel_cells:
                for r, c in sel_cells: real_cols.add(int(c) + col_offset)
                
            curr = self.sheet.get_currently_selected()
            if curr:
                # รองรับ tksheet ทุกเวอร์ชั่น
                if isinstance(curr[0], int):
                    real_cols.add(int(curr[1]) + col_offset)
                elif isinstance(curr[0], str) and len(curr) >= 2:
                    if curr[0] == "column": real_cols.add(int(curr[1]) + col_offset)
                    elif curr[0] == "cell" and len(curr) >= 3: real_cols.add(int(curr[2]) + col_offset)
        except Exception as e: 
            print(f"Error right cols: {e}")

        # 2. เช็คจาก Sheet ที่ตรึงไว้ฝั่งซ้าย (ถ้ามี)
        if self.sheet_frozen:
            try:
                sel_cols = self.sheet_frozen.get_selected_columns()
                if sel_cols:
                    for c in sel_cols: real_cols.add(int(c))
                    
                sel_cells = self.sheet_frozen.get_selected_cells()
                if sel_cells:
                    for r, c in sel_cells: real_cols.add(int(c))
                    
                curr = self.sheet_frozen.get_currently_selected()
                if curr:
                    if isinstance(curr[0], int):
                        real_cols.add(int(curr[1]))
                    elif isinstance(curr[0], str) and len(curr) >= 2:
                        if curr[0] == "column": real_cols.add(int(curr[1]))
                        elif curr[0] == "cell" and len(curr) >= 3: real_cols.add(int(curr[2]))
            except Exception as e: 
                print(f"Error frozen cols: {e}")

        # กรองเอาเฉพาะ index ที่ถูกต้องจริงๆ
        valid_cols = [int(c) for c in list(real_cols) if 0 <= int(c) < len(self.columns)]
        return valid_cols

    def _open_column_visibility_panel(self):
        """Panel checkbox สำหรับเลือกแสดง/ซ่อน column แบบ list"""
        panel = CTkToplevel(self)
        panel.title("⚙️ จัดการคอลัมน์")
        panel.resizable(False, True)
        panel.grab_set()
        panel.transient(self.winfo_toplevel())
        panel.geometry("320x520")

        CTkLabel(panel, text="เลือกคอลัมน์ที่ต้องการแสดง",
                 font=CTkFont(size=13, weight="bold")).pack(pady=(12, 4), padx=15)
        CTkLabel(panel, text="✓ = แสดง   □ = ซ่อน",
                 font=CTkFont(size=11), text_color="#6B7280").pack(pady=(0, 6))

        scroll = CTkScrollableFrame(panel, width=290, height=370)
        scroll.pack(fill="both", expand=True, padx=10, pady=4)

        hidden_set = set(getattr(self, "hidden_cols_list", []))
        col_vars = {}
        for i, col_name in enumerate(self.columns):
            var = tk.BooleanVar(value=(i not in hidden_set))
            col_vars[i] = var
            CTkCheckBox(scroll, text=col_name, variable=var,
                        font=CTkFont(size=12), height=26).pack(anchor="w", padx=6, pady=1)

        btn_frame = CTkFrame(panel, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(6, 12))

        def apply_visibility():
            self.hidden_cols_list = [i for i, v in col_vars.items() if not v.get()]
            self._apply_hidden_columns()
            self._save_user_settings()
            panel.destroy()

        def reset_all():
            for v in col_vars.values():
                v.set(True)

        CTkButton(btn_frame, text="✅ ตกลง", command=apply_visibility,
                  height=32).pack(side="left", fill="x", expand=True, padx=(0, 4))
        CTkButton(btn_frame, text="🔄 แสดงทั้งหมด", command=reset_all,
                  fg_color="#6B7280", hover_color="#4B5563",
                  height=32).pack(side="left", fill="x", expand=True, padx=(4, 0))

    def _apply_hidden_columns(self):
        """ใช้ display_columns() แทน hide_columns() เพื่อป้องกัน squish bug"""
        if not HAS_TKSHEET: return

        self.hidden_cols_list = [
            c for c in set(getattr(self, "hidden_cols_list", []))
            if 0 <= c < len(self.columns)
        ]
        col_offset = self.frozen_col_count if self.sheet_frozen else 0
        total = len(self.columns)

        # Step 1: reset เป็น "all" ก่อนเสมอ เพื่อป้องกัน index เพี้ยนจาก session ก่อน
        try:
            self.sheet.display_columns("all")
            if self.sheet_frozen:
                self.sheet_frozen.display_columns("all")
        except Exception: pass

        # Step 2: restore column widths ขณะที่ทุก column ยังแสดงอยู่
        self._restore_col_widths()

        if not self.hidden_cols_list:
            try: self.sheet.redraw()
            except Exception: pass
            if self.sheet_frozen:
                try: self.sheet_frozen.redraw()
                except Exception: pass
            return

        # Step 3: apply display_columns() เฉพาะ column ที่ต้องการแสดง
        hidden_set = set(self.hidden_cols_list)

        if self.sheet_frozen:
            frozen_visible = [c for c in range(col_offset) if c not in hidden_set]
            main_visible   = [c - col_offset for c in range(col_offset, total) if c not in hidden_set]
            try:
                self.sheet_frozen.display_columns(
                    columns=frozen_visible,
                    all_columns_displayed=(len(frozen_visible) == col_offset),
                    reset_col_positions=True,
                    redraw=False,
                )
            except Exception as e:
                print(f"frozen display_columns error: {e}")
            try:
                self.sheet.display_columns(
                    columns=main_visible,
                    all_columns_displayed=(len(main_visible) == total - col_offset),
                    reset_col_positions=True,
                    redraw=False,
                )
            except Exception as e:
                print(f"main display_columns error: {e}")
        else:
            visible = [c for c in range(total) if c not in hidden_set]
            try:
                self.sheet.display_columns(
                    columns=visible,
                    all_columns_displayed=(len(visible) == total),
                    reset_col_positions=True,
                    redraw=False,
                )
            except Exception as e:
                print(f"display_columns error: {e}")

        # Step 4: restore widths AFTER display_columns resets col_positions
        # (reset_col_positions=True resets all widths back to 120 — must re-apply cache)
        self._restore_col_widths()

        try: self.sheet.redraw()
        except Exception: pass
        if self.sheet_frozen:
            try: self.sheet_frozen.redraw()
            except Exception: pass

    def _show_all_columns(self):
        """ยังคงไว้เพื่อ backward-compat (เรียกจาก panel "แสดงทั้งหมด")"""
        if not HAS_TKSHEET: return
        self.hidden_cols_list = []
        self._apply_hidden_columns()
        self._save_user_settings()

    def _change_header_color(self):
        if not HAS_TKSHEET: return
        real_cols = self._get_real_col_indices()
        if not real_cols:
            messagebox.showwarning("แจ้งเตือน", "กรุณาคลิก 'ช่องใดๆ' หรือ 'หัวคอลัมน์' ที่ต้องการเปลี่ยนสีก่อน", parent=self)
            return
            
        color_tuple = colorchooser.askcolor(title="เลือกสีสำหรับหัวคอลัมน์", parent=self)
        if not color_tuple or not color_tuple[1]: return
        
        color_code = color_tuple[1]
        col_offset = self.frozen_col_count if self.sheet_frozen else 0
        
        for c_idx in real_cols:
            if c_idx >= len(self.columns): continue
            self.custom_header_colors[self.columns[c_idx]] = color_code
            
            b_bg = self._lighten_color(color_code, amount=0.85)
            display_idx = c_idx - col_offset
            
            if display_idx < 0 and self.sheet_frozen:
                try:
                    self.sheet_frozen.highlight_cells(row=0, column=c_idx, bg=color_code, fg="black", canvas="header")
                    self.sheet_frozen.highlight_columns(columns=[c_idx], bg=b_bg, fg="black", highlight_header=False)
                except Exception: pass
            elif display_idx >= 0:
                try:
                    self.sheet.highlight_cells(row=0, column=display_idx, bg=color_code, fg="black", canvas="header")
                    self.sheet.highlight_columns(columns=[display_idx], bg=b_bg, fg="black", highlight_header=False)
                except Exception: pass

        self._save_user_settings()
        try: self.sheet.redraw()
        except Exception: pass
        if self.sheet_frozen:
            try: self.sheet_frozen.redraw()
            except Exception: pass

    # ================================================================== #
    # ZOOM
    # ================================================================== #
    def _zoom(self, direction):
        if not HAS_TKSHEET: return
        
        old_zoom = self.zoom_level
        self.zoom_level = max(2, min(40, self.zoom_level + direction))

        if old_zoom == self.zoom_level: return

        scale_ratio = self.zoom_level / old_zoom
        new_row_height = int(30 * (self.zoom_level / 11.0))
        new_header_height = int(35 * (self.zoom_level / 11.0))

        opts = dict(
            font=("Tahoma", self.zoom_level, "normal"),
            header_font=("Tahoma", self.zoom_level, "bold"),
            row_height=new_row_height,
            header_height=new_header_height,
            auto_resize_columns=False,
        )

        self.sheet.set_options(**opts)
        if self.sheet_frozen:
            self.sheet_frozen.set_options(**opts)

        try:
            for i in range(self.sheet.get_total_columns()):
                try:
                    current_w = self.sheet.column_width(i)
                    if current_w:
                        self.sheet.column_width(i, int(current_w * scale_ratio))
                except Exception: pass

            frozen_width = 40
            if self.sheet_frozen:
                for i in range(self.frozen_col_count):
                    try:
                        current_w = self.sheet_frozen.column_width(i)
                        if current_w:
                            new_w = int(current_w * scale_ratio)
                            self.sheet_frozen.column_width(i, new_w)
                            frozen_width += new_w
                    except Exception: pass
                
                frozen_width += 4
                self._current_frozen_width = frozen_width
                self.sheet_frozen.place(x=0, y=0, width=frozen_width, relheight=1.0)
                self.sheet.place(x=frozen_width, y=0, relwidth=1.0, relheight=1.0, width=-frozen_width)

                def _on_resize_after_zoom(event):
                    try:
                        fw = getattr(self, '_current_frozen_width', frozen_width)
                        self.sheet.place(x=fw, y=0, relwidth=1.0, relheight=1.0, width=-fw)
                    except Exception: pass
                self.table_frame.bind("<Configure>", _on_resize_after_zoom)
        except Exception as e:
            print(f"Zoom resize error: {e}")

        try: self.sheet.redraw()
        except Exception: pass
        if self.sheet_frozen:
            try: self.sheet_frozen.redraw()
            except Exception: pass

        def _sync_after_zoom():
            try:
                rh = int(30 * (self.zoom_level / 11.0))
                if self.sheet_frozen:
                    total_rows = max(self.sheet.get_total_rows(), self.sheet_frozen.get_total_rows())
                else:
                    total_rows = self.sheet.get_total_rows()
                    
                for r in range(total_rows):
                    if self.sheet_frozen:
                        self.sheet_frozen.row_height(r, rh)
                    self.sheet.row_height(r, rh)
                    
                if self.sheet_frozen:
                    self.sheet_frozen.redraw()
                    self._hide_frozen_scrollbars()
                self.sheet.redraw()
            except Exception: pass

        self.after(100, _sync_after_zoom)
        self.after(100, self._save_col_widths)
        
        pct = int((self.zoom_level / 11) * 100)
        if hasattr(self, 'zoom_label'):
            self.zoom_label.configure(text=f"{pct}%")