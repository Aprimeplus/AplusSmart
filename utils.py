# utils.py (ฉบับสมบูรณ์ จัดระเบียบใหม่)

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import customtkinter as ctk
from datetime import datetime
from customtkinter import CTkFrame, CTkLabel, CTkScrollableFrame, CTkToplevel, CTkFont, CTkEntry, CTkCheckBox, CTkButton

# =========================================================================
# HELPER FUNCTIONS (ฟังก์ชันช่วยเหลือทั่วไป)
# =========================================================================

def convert_to_float(value):
    """แปลงค่าต่างๆ ให้เป็น Float อย่างปลอดภัย"""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        if value is None or value.strip() == '':
            return 0.0
        try:
            return float(value.replace(",", ""))
        except (ValueError, TypeError):
            return 0.0
    return 0.0

def format_date_safe(date_val, format_string='%d/%m/%Y'):
    """
    รับค่าจาก Database (ที่เป็น Date, Timestamp, String หรือ None)
    แปลงเป็น String ตาม format_string อย่างปลอดภัย
    """
    if pd.isna(date_val) or date_val is None or str(date_val).strip() in ('', 'None', 'NaT'):
        return "-"
    
    try:
        dt = pd.to_datetime(date_val, errors='coerce')
        if pd.notna(dt):
            return dt.strftime(format_string)
        return str(date_val)
    except Exception:
        return str(date_val)

# =========================================================================
# THAI DATE HELPERS — ใช้ทั่วระบบเพื่อแสดงปี พ.ศ.
# =========================================================================

def thai_year(dt) -> int:
    """รับ datetime/date แล้วคืนปี พ.ศ."""
    return dt.year + 543

def format_thai_date(dt, show_time: bool = False) -> str:
    """
    แปลง datetime → string รูปแบบไทย DD/MM/YYYY(พ.ศ.)
    เช่น 2026-06-06 → '06/06/2569'
    ถ้า show_time=True → '06/06/2569 10:37'
    """
    if dt is None:
        return "-"
    try:
        dt = pd.to_datetime(dt)
        if pd.isna(dt):
            return "-"
        base = dt.strftime("%d/%m/") + str(dt.year + 543)
        if show_time:
            base += dt.strftime(" %H:%M")
        return base
    except Exception:
        return "-"

def format_thai_date_safe(date_val, show_time: bool = False) -> str:
    """เหมือน format_thai_date แต่รับ string/None/NaT ได้ด้วย"""
    if date_val is None or str(date_val).strip() in ('', 'None', 'NaT'):
        return "-"
    return format_thai_date(date_val, show_time=show_time)

def thai_year_from_str(year_str: str) -> int:
    """รับปี ค.ศ. เป็น string แล้วคืนปี พ.ศ. (เช่น '2026' → 2569)"""
    return int(year_str) + 543

def be_to_ce(be_year: int) -> int:
    """แปลงปี พ.ศ. → ค.ศ. (เช่น 2569 → 2026)"""
    return be_year - 543

def ce_to_be(ce_year: int) -> int:
    """แปลงปี ค.ศ. → พ.ศ. (เช่น 2026 → 2569)"""
    return ce_year + 543

def set_entry_text(entry_widget, text):
    """ใส่ข้อความใน CTkEntry อย่างปลอดภัย รองรับกรณีที่ Widget เป็น Readonly"""
    if not entry_widget or not hasattr(entry_widget, 'winfo_exists') or not entry_widget.winfo_exists():
        return
    
    try:
        current_state = entry_widget.cget("state")
        if current_state == "readonly":
            entry_widget.configure(state="normal")
        
        entry_widget.delete(0, "end")
        entry_widget.insert(0, str(text))
        
        if current_state == "readonly":
            entry_widget.configure(state="readonly")
    except Exception as e:
        print(f"Error in set_entry_text for widget {entry_widget}: {e}")

def center_window(window: tk.Toplevel):
    """จัดตำแหน่งหน้าต่างให้อยู่ตรงกลางของหน้าต่างแม่"""
    window.update_idletasks() 

    window_width = window.winfo_width()
    window_height = window.winfo_height()

    master_x = window.master.winfo_x()
    master_y = window.master.winfo_y()
    master_width = window.master.winfo_width()
    master_height = window.master.winfo_height()

    center_x = int(master_x + (master_width / 2) - (window_width / 2))
    center_y = int(master_y + (master_height / 2) - (window_height / 2))

    window.geometry(f"+{center_x}+{center_y}")


# =========================================================================
# CUSTOM WIDGETS (วิดเจ็ตที่ปรับแต่งพิเศษ)
# =========================================================================

class FormattedNumericEntry(ctk.CTkEntry):
    """ช่องกรอกตัวเลขแบบมีลูกน้ำ (Comma) อัตโนมัติเวลาพิมพ์"""
    def __init__(self, master, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self._variable = tk.StringVar(self)
        self.configure(textvariable=self._variable)
        self._command = command
        self._is_formatting = False
        self._variable.trace_add("write", self._format_and_callback)
        self.bind("<FocusOut>", self._format_on_blur)
        self.bind("<FocusIn>", lambda e: self.after(10, lambda: (self.select_range(0, tk.END), self.icursor(tk.END))))

    def _format_and_callback(self, *args):
        if self._is_formatting: return
        self._is_formatting = True
        try:
            current_value = self._variable.get()
            # ไม่ auto-format ขณะพิมพ์ — แค่ trigger callback
            if self._command: self._command()
        except (ValueError, TypeError): pass
        finally: self._is_formatting = False

    def _format_on_blur(self, event=None):
        """Format เป็น 3 ทศนิยมเมื่อ focus ออก"""
        if self._is_formatting: return
        self._is_formatting = True
        try:
            current_value = self._variable.get()
            numeric_chars = ''.join(filter(lambda c: c.isdigit() or c == '.', current_value.replace(',', '')))
            if numeric_chars and numeric_chars != '.':
                number = float(numeric_chars)
                self._variable.set(f"{number:,.3f}")
        except (ValueError, TypeError): pass
        finally: self._is_formatting = False

    def get_value(self) -> float:
        try: return float(self._variable.get().replace(",", ""))
        except (ValueError, TypeError): return 0.0

    def set(self, value: float):
        try:
            num_value = float(value)
            self._variable.set(f"{num_value:,.3f}")
        except (ValueError, TypeError): self._variable.set("0.000")


# =========================================================================
# DIALOG WINDOWS (หน้าต่าง Popup ต่างๆ)
# =========================================================================

class RejectionReasonDialog(CTkToplevel):
    """หน้าต่างสำหรับเลือกเหตุผลเวลาตีกลับ/ปฏิเสธงาน"""
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title("ระบุเหตุผลที่ปฏิเสธ")
        self.geometry("500x600")
        
        self.reasons_list = [
            "ลงสเปคสินค้าผิด SO", "ลงเสปคสินค้าผิด PO", "ลงราคาต้นทุนผิด PO", 
            "ลงราคาขายผิด SO", "ไม่แยกค่ารถ/ราคาผิด SO", "ไม่แยกค่ารถ/ราคาผิด PO", 
            "รายการต้นทุนไม่ครบ PO", "ค่าตัด/เจาะ ตกหล่น", "ค่าของแถม ตกหล่น"
        ]
        self.checkbox_vars = []
        self._reason_string = None
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        CTkLabel(self, text="กรุณาเลือกเหตุผลที่ปฏิเสธ (เลือกได้มากกว่า 1 ข้อ)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=10)
        
        scroll_frame = CTkScrollableFrame(self)
        scroll_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        
        for reason in self.reasons_list:
            var = tk.StringVar(value="0")
            cb = CTkCheckBox(scroll_frame, text=reason, variable=var, font=CTkFont(size=14))
            cb.pack(pady=5, padx=10, anchor="w")
            self.checkbox_vars.append((var, reason))
            
        other_frame = CTkFrame(self, fg_color="transparent")
        other_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        other_frame.grid_columnconfigure(1, weight=1)
        
        CTkLabel(other_frame, text="อื่นๆ:", font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=(0,5))
        self.other_reason_entry = CTkEntry(other_frame)
        self.other_reason_entry.grid(row=0, column=1, sticky="ew")
        
        button_frame = CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=3, column=0, padx=20, pady=10)
        
        CTkButton(button_frame, text="ยกเลิก", command=self.destroy, fg_color="gray").pack(side="right", padx=5)
        CTkButton(button_frame, text="ตกลง", command=self._on_confirm, fg_color="#DC2626", hover_color="#B91C1C").pack(side="right", padx=5)
        
        self.transient(master)
        self.grab_set()

    def _on_confirm(self):
        selected_reasons = [reason_text for var, reason_text in self.checkbox_vars if var.get() == "1"]
        other_text = self.other_reason_entry.get().strip()
        
        if other_text:
            selected_reasons.append(f"อื่นๆ: {other_text}")
            
        if not selected_reasons:
            messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณาเลือกเหตุผลอย่างน้อย 1 ข้อ", parent=self)
            return
            
        self._reason_string = ", ".join(selected_reasons)
        self.destroy()

# เพิ่มฟังก์ชันแสดง Popup แจ้งเตือนแบบน่ารักๆ ไว้ตรงนี้ด้วย เผื่อเอาไปเรียกใช้ (Optional)
def show_loading_popup(master, message="กำลังโหลดข้อมูล..."):
    """แสดงหน้าต่าง Loading หมุนๆ เล็กๆ ระหว่างดึงข้อมูล (เอาไปใช้ได้ทุกหน้า)"""
    popup = CTkToplevel(master)
    popup.geometry("300x120")
    popup.title("โปรดรอ...")
    center_window(popup)
    
    popup.overrideredirect(True) # ลบขอบ
    popup.attributes("-topmost", True)
    
    frame = CTkFrame(popup, fg_color="#FFFFFF", border_width=2, border_color="#3B82F6")
    frame.pack(fill="both", expand=True)
    
    lbl = CTkLabel(frame, text=message, font=CTkFont(size=14, weight="bold"), text_color="#1E3A8A")
    lbl.pack(expand=True)
    
    popup.label = lbl # เก็บ ref ไว้เผื่อแก้ไขข้อความ
    popup.update()
    
    return popup