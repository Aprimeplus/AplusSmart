# custom_widgets.py (เวอร์ชันแก้ไขโครงสร้าง Popup)

import customtkinter as ctk
from datetime import datetime
import tkinter as tk
import pandas as pd
from thefuzz import process, fuzz
from tkinter import font as tkFont 

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

# ==============================================================================
# <<< START: โค้ดที่แก้ไขของ AutoCompleteEntry ทั้งคลาส >>>
# ==============================================================================
class AutoCompleteEntry(ctk.CTkEntry):
    def __init__(self, master, completion_list, display_key, **kwargs):
        self.var = tk.StringVar()
        self.command = kwargs.pop('command', None)
        super().__init__(master, textvariable=self.var, **kwargs)
        
        self.completion_list = completion_list
        self.display_key = display_key
        
        self._map_display_to_object = {}
        self._choices = []
        self.update_completion_list(self.completion_list)

        self.popup = None
        self.listbox = None
        
        # ตัวแปรสำหรับติดตาม ID ของ trace เพื่อจัดการการเปิด/ปิด
        # เราจะไม่เพิ่ม trace ในตอนนี้ เพื่อให้สามารถตั้งค่าเริ่มต้นได้โดยไม่แสดง popup
        self._trace_id = None
        
        # ผูก KeyRelease เพื่อเริ่มการค้นหาและเพิ่ม trace เมื่อผู้ใช้เริ่มพิมพ์
        self.bind("<KeyRelease>", self._start_autocomplete)
        
        # บันทึกการผูก Event อื่นๆ เหมือนเดิม
        self.bind("<FocusOut>", self._hide_popup)
        self.bind("<Down>", self._focus_on_listbox)
        self.bind("<Escape>", self._hide_popup)
        self.bind("<Configure>", self._reposition_popup)

    def update_completion_list(self, new_list):
        self.completion_list = new_list
        self._map_display_to_object = {str(item.get(self.display_key, '')): item for item in self.completion_list}
        self._choices = list(self._map_display_to_object.keys())

    def _start_autocomplete(self, event=None):
        """
        เริ่มการ trace เมื่อมีการปล่อยปุ่มใดๆ (ผู้ใช้เริ่มพิมพ์)
        และเรียกใช้ _on_text_change ทันทีเพื่อเริ่มต้นการค้นหา
        """
        # เพิ่ม trace เมื่อ KeyRelease ถูกเรียกครั้งแรก
        if self._trace_id is None:
            self._trace_id = self.var.trace_add("write", self._on_text_change)
        
        # เรียกใช้ _on_text_change ทันทีเพื่อเริ่มต้นการค้นหา
        self._on_text_change()

    def _on_text_change(self, *args, **kwargs):
        current_text = self.var.get()
        
        # <<< START: การแก้ไข: ซ่อน Popup หากไม่มีข้อความหรือข้อความมีแต่ช่องว่าง >>>
        # ใช้ .strip() เพื่อตรวจสอบว่ามีข้อความที่พิมพ์จริงหรือไม่ (ไม่ใช่แค่ space bar)
        if not current_text or not current_text.strip():
            self._hide_popup()
            return
        # <<< END: การแก้ไข >>>
        
        # ใช้ fuzz.partial_ratio ในการค้นหาตามเดิม
        results = process.extract(current_text, self._choices, scorer=fuzz.partial_ratio, limit=10)
        self.matches = [result[0] for result in results if result[1] > 70]

        if self.matches:
            if self.popup is None or not self.popup.winfo_exists():
                self._create_popup()
            
            self.listbox.delete(0, tk.END)
            for item in self.matches:
                self.listbox.insert(tk.END, item)
            self._show_popup()
        else:
            self._hide_popup()


    def _create_popup(self):
        self.popup = tk.Toplevel(self)
        self.popup.overrideredirect(True)
        self.popup.withdraw()

        font_object = self.cget("font")
        font_tuple = (font_object.cget("family"), font_object.cget("size"))
        mode_index = 1 if ctk.get_appearance_mode().lower() == "dark" else 0
        bg_color = self.cget("fg_color")[mode_index]
        text_color = self.cget("text_color")[mode_index]
        border_color = self.cget("border_color")[mode_index]

        self.listbox = tk.Listbox(self.popup, 
                                  font=font_tuple, bg=bg_color, fg=text_color,
                                  selectbackground=border_color, selectforeground=text_color,
                                  highlightthickness=1, highlightcolor=border_color,
                                  borderwidth=0, activestyle="none")
        self.listbox.pack(fill="both", expand=True)
                                  
        self.listbox.bind("<ButtonRelease-1>", self._on_select)
        self.listbox.bind("<Return>", self._on_select)
        self.listbox.bind("<Escape>", self._hide_popup)
    
    def _show_popup(self):
        if not self.popup or not self.winfo_exists() or not hasattr(self, 'matches') or not self.matches:
            self._hide_popup()
            return

        # <<< START: แก้ไข Logic การคำนวณความกว้าง >>>
        # บวกความกว้างเพิ่มเข้าไปอีก 350 pixels จากความกว้างของช่องกรอกข้อมูล
        # วิธีนี้จะทำให้มีพื้นที่เพียงพอสำหรับชื่อที่ยาวมากๆ
        width = self.winfo_width() + 350
        # <<< END >>>

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        
        listbox_height = min(len(self.matches), 7)
        row_height = 25
        height = listbox_height * row_height
        
        self.popup.geometry(f"{width}x{height}+{x}+{y}")
        self.popup.deiconify()
        self.popup.lift()

    def _hide_popup(self, event=None):
        if self.popup:
            self.after(150, lambda: self.popup.withdraw() if self.popup and self.popup.winfo_exists() else None)

    def _reposition_popup(self, event=None):
        if self.popup and self.popup.winfo_viewable():
            self._show_popup()
            
    def _focus_on_listbox(self, event=None):
        if self.listbox and self.popup and self.popup.winfo_viewable():
            self.listbox.focus_set()
            self.listbox.selection_set(0)
            return "break"
            
    def _on_select(self, event=None):
        if not self.listbox or not self.listbox.curselection():
            return "break"
            
        selection_text = self.listbox.get(self.listbox.curselection())
        
        trace_info = self.var.trace_info()
        if trace_info: self.var.trace_vdelete("w", trace_info[0][1])
        
        self.var.set(selection_text)
        
        self.var.trace_add("write", self._on_text_change)

        self._hide_popup()
        self.icursor(tk.END)
        self.focus_set()
        
        if self.command:
            selected_object = self._map_display_to_object.get(selection_text)
            if selected_object:
                self.command(selected_object)
                
        return "break"