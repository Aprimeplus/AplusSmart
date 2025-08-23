# custom_widgets.py (ฉบับสมบูรณ์)

import customtkinter as ctk
from datetime import datetime
import tkinter as tk
import pandas as pd
from thefuzz import process, fuzz

class NumericEntry(ctk.CTkEntry):
    def __init__(self, master, **kwargs):
        self.error_border_color = kwargs.pop("error_border_color", "#D32F2F")
        super().__init__(master, **kwargs)
        self.default_border_color = self.cget("border_color")
        self.bind("<KeyRelease>", self._validate_input)

    def _validate_input(self, event=None):
        current_value = self.get()
        if not current_value:
            self.configure(border_color=self.default_border_color)
            return
        try:
            float(current_value.replace(',', ''))
            self.configure(border_color=self.default_border_color)
        except ValueError:
            self.configure(border_color=self.error_border_color)

class DateSelector(ctk.CTkFrame):
    def __init__(self, master, dropdown_style=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.day_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.year_var = tk.StringVar()
        self.thai_months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
        
        style = dropdown_style if dropdown_style is not None else {}
        
        self.day_menu = ctk.CTkOptionMenu(self, variable=self.day_var, width=65, **style)
        self.month_menu = ctk.CTkOptionMenu(self, variable=self.month_var, values=self.thai_months, width=80, command=lambda _: self._update_days(), **style)
        now = datetime.now()
        current_be_year = now.year + 543
        self.year_menu = ctk.CTkOptionMenu(self, variable=self.year_var, values=[str(y) for y in range(current_be_year - 2, current_be_year + 5)], width=75, command=lambda _: self._update_days(), **style)
        
        self.day_menu.pack(side="left", padx=(0, 2))
        self.month_menu.pack(side="left", padx=2)
        self.year_menu.pack(side="left", padx=(2, 0))
        
        self._set_to_today()
        self.month_var.trace_add("write", self._update_days)
        self.year_var.trace_add("write", self._update_days)

    def _update_days(self, *args):
        try:
            if not self.day_menu.winfo_exists(): return
            thai_month_map = {"ม.ค.": 31, "ก.พ.": 28, "มี.ค.": 31, "เม.ย.": 30, "พ.ค.": 31, "มิ.ย.": 30, "ก.ค.": 31, "ส.ค.": 31, "ก.ย.": 30, "ต.ค.": 31, "พ.ย.": 30, "ธ.ค.": 31}
            year_val = int(self.year_var.get()) - 543
            is_leap = (year_val % 4 == 0 and year_val % 100 != 0) or (year_val % 400 == 0)
            thai_month_map["ก.พ."] = 29 if is_leap else 28
            max_days = thai_month_map.get(self.month_var.get(), 31)
            current_day = self.day_var.get()
            self.day_menu.configure(values=[f"{d:02d}" for d in range(1, max_days + 1)])
            if current_day and int(current_day) > max_days:
                self.day_var.set(f"{max_days:02d}")
            elif not current_day:
                self.day_var.set("01")
        except Exception: return

    def _set_to_today(self):
        now = datetime.now()
        self.day_var.set(f"{now.day:02d}")
        self.month_var.set(self.thai_months[now.month - 1])
        self.year_var.set(str(now.year + 543))
        self._update_days()

    def get_date(self):
        thai_month_map_to_num = {"ม.ค.": "01", "ก.พ.": "02", "มี.ค.": "03", "เม.ย.": "04", "พ.ค.": "05", "มิ.ย.": "06", "ก.ค.": "07", "ส.ค.": "08", "ก.ย.": "09", "ต.ค.": "10", "พ.ย.": "11", "ธ.ค.": "12"}
        try:
            day = self.day_var.get()
            month = thai_month_map_to_num.get(self.month_var.get())
            year = int(self.year_var.get()) - 543
            if not day or not month or not year: return None
            return f"{year}-{month}-{day}"
        except (ValueError, TypeError, KeyError): return None

    def set_date(self, date_obj):
        if date_obj is None or pd.isna(date_obj) or not hasattr(date_obj, 'strftime'): return
        thai_months_rev = {"01": "ม.ค.", "02": "ก.พ.", "03": "มี.ค.", "04": "เม.ย.", "05": "พ.ค.", "06": "มิ.ย.", "07": "ก.ค.", "08": "ส.ค.", "09": "ก.ย.", "10": "ต.ค.", "11": "พ.ย.", "12": "ธ.ค."}
        self.day_var.set(date_obj.strftime("%d"))
        self.month_var.set(thai_months_rev.get(date_obj.strftime("%m"), ""))
        self.year_var.set(str(date_obj.year + 543))

# อยู่ในไฟล์ custom_widgets.py

class AutoCompleteEntry(ctk.CTkEntry):
    def __init__(self, master, completion_list, command_on_select=None, display_key_on_select='name', **kwargs):
        super().__init__(master, **kwargs)
        self.completion_list = completion_list
        self.command_on_select = command_on_select
        self.display_key_on_select = display_key_on_select # <<< แก้ไข: เพิ่มตัวแปรใหม่
        self._suggestions_toplevel = None
        self._suggestion_buttons = []
        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Escape>", lambda e: self._hide_suggestions())

    def _on_key_release(self, event):
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock", "Tab", "Up", "Down", "Left", "Right", "Return", "Escape"):
            return
        
        typed_text = self.get().lower()
        if not typed_text:
            self._hide_suggestions()
            return

        if not self.completion_list or not isinstance(self.completion_list[0], dict):
            return

        # --- START: แก้ไข Logic การค้นหาให้เข้ากันได้ ---
        
        choices_map = {f"{item.get('name', '')} {item.get('code', '')}".lower(): item for item in self.completion_list}
        
        # 1. ดึงผลลัพธ์ทั้งหมด 10 อันดับแรกออกมาก่อน (โดยไม่ใช้ score_cutoff)
        all_results = process.extract(typed_text, choices_map.keys(), scorer=fuzz.partial_ratio, limit=10)
        
        # 2. กรองผลลัพธ์ด้วยตนเอง โดยเอาเฉพาะที่มีคะแนนความคล้ายตั้งแต่ 70% ขึ้นไป
        filtered_results = [result for result in all_results if result[1] >= 70]
        
        # 3. แปลงผลลัพธ์กลับไปเป็นลิสต์ของ Dictionary เดิม
        filtered_list = [choices_map[result[0]] for result in filtered_results]
        
        # --- END: สิ้นสุดการแก้ไข ---

        if not filtered_list:
            self._hide_suggestions()
            return
        
        self._show_suggestions(filtered_list)

    def _show_suggestions(self, suggestions):
        if self._suggestions_toplevel is None or not self._suggestions_toplevel.winfo_exists():
            self._suggestions_toplevel = ctk.CTkToplevel(self)
            self._suggestions_toplevel.overrideredirect(True)
        
        for btn in self._suggestion_buttons:
            btn.destroy()
        self._suggestion_buttons.clear()

        # --- START: เพิ่ม Logic คำนวณความกว้างอัตโนมัติ ---
        # 1. หาข้อความที่ยาวที่สุดในลิสต์ผลลัพธ์
        longest_suggestion_text = ""
        if suggestions:
            longest_suggestion_text = max((item['display'] for item in suggestions), key=len)

        # 2. คำนวณความกว้างเป็นพิกเซลจากข้อความที่ยาวที่สุด
        font = self.cget("font")
        if longest_suggestion_text:
            # เพิ่ม Padding เข้าไปเล็กน้อยเพื่อให้ไม่ชิดขอบเกินไป
            required_width = font.measure(longest_suggestion_text) + 40 
        else:
            required_width = self.winfo_width()

        # 3. ทำให้แน่ใจว่าความกว้างที่คำนวณได้ จะไม่น้อยกว่าความกว้างของช่อง Entry เอง
        min_width = self.winfo_width()
        final_width = max(required_width, min_width)
        # --- END: สิ้นสุด Logic คำนวณความกว้าง ---

        for item in suggestions[:10]:
            display_text = item['display']
            btn = ctk.CTkButton(
                self._suggestions_toplevel, 
                text=display_text, 
                anchor="w", 
                fg_color="white", 
                text_color="black", 
                hover_color="#E5E7EB"
            )
            btn.pack(fill="x", expand=True)
            btn.configure(command=lambda i=item: self._on_suggestion_click(i))
            self._suggestion_buttons.append(btn)
        
        x, y = self.winfo_rootx(), self.winfo_rooty() + self.winfo_height()
        
        # --- ใช้ความกว้างใหม่ที่คำนวณได้ในการกำหนดขนาดหน้าต่าง ---
        self._suggestions_toplevel.geometry(f"{final_width}x{self._suggestions_toplevel.winfo_reqheight()}+{x}+{y}")
        
        self._suggestions_toplevel.lift()
        self._suggestions_toplevel.deiconify()

    def _hide_suggestions(self):
        if self._suggestions_toplevel and self._suggestions_toplevel.winfo_exists():
            self._suggestions_toplevel.withdraw()

   
    def _on_suggestion_click(self, selection_dict):
        self._hide_suggestions()
        self.delete(0, tk.END)
        
        display_value = selection_dict.get(self.display_key_on_select, '')
        self.insert(0, display_value)
        
        if self.command_on_select:
            self.command_on_select(selection_dict) 
        
        self.winfo_toplevel().focus_set()
        
    def _on_focus_out(self, event):
        self.after(200, self._hide_suggestions)