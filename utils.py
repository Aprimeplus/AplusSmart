# utils.py (ฉบับสมบูรณ์)

import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import customtkinter as ctk
from datetime import datetime
from customtkinter import CTkFrame, CTkLabel, CTkScrollableFrame, CTkToplevel, CTkFont, CTkEntry, CTkCheckBox, CTkButton
from tkinter import messagebox

# --- ฟังก์ชัน Helper ---
def convert_to_float(value_str):
    if value_str is None or value_str == '':
        return 0.0
    try:
        return float(str(value_str).replace(",", ""))
    except (ValueError, TypeError):
        return 0.0

# --- คลาส Widget พิเศษ ---
class FormattedNumericEntry(ctk.CTkEntry):
    def __init__(self, master, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self._variable = tk.StringVar(self)
        self.configure(textvariable=self._variable)
        self._command = command
        self._is_formatting = False
        self._variable.trace_add("write", self._format_and_callback)

    def _format_and_callback(self, *args):
        if self._is_formatting: return
        self._is_formatting = True
        try:
            current_value = self._variable.get()
            cursor_position = self.index(tk.INSERT)
            numeric_chars = ''.join(filter(lambda char: char.isdigit() or char == '.', current_value))
            if numeric_chars and numeric_chars != '.':
                number = float(numeric_chars)
                formatted_value = f"{number:,.2f}"
                if formatted_value != current_value:
                    self._variable.set(formatted_value)
                    commas_before = current_value[:cursor_position].count(',')
                    commas_after = formatted_value.count(',')
                    new_cursor_pos = cursor_position + (commas_after - commas_before)
                    if new_cursor_pos >= 0: self.icursor(new_cursor_pos)
            if self._command: self._command()
        except (ValueError, TypeError): pass
        finally: self._is_formatting = False

    def get_value(self) -> float:
        try: return float(self._variable.get().replace(",", ""))
        except (ValueError, TypeError): return 0.0

    def set(self, value: float):
        try:
            num_value = float(value)
            self._variable.set(f"{num_value:,.2f}")
        except (ValueError, TypeError): self._variable.set("0.00")

# --- คลาสหน้าต่าง Dialog ที่ใช้งานร่วมกัน ---
class RejectionReasonDialog(CTkToplevel):
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
        
        CTkButton(button_frame, text="ยกเลิก", command=self.destroy).pack(side="right", padx=5)
        CTkButton(button_frame, text="ตกลง", command=self._on_confirm).pack(side="right", padx=5)
        
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
    
def set_entry_text(entry_widget, text):
        """
        Helper function to safely set text in a CTkEntry, handling the readonly state.
        """
        # ตรวจสอบว่า widget ยัง存在และใช้งานได้
        if not entry_widget or not hasattr(entry_widget, 'winfo_exists') or not entry_widget.winfo_exists():
            return
        
        try:
            current_state = entry_widget.cget("state")
            # หากเป็นแบบอ่านอย่างเดียว (readonly) ให้เปิดเป็นปกติก่อน
            if current_state == "readonly":
                entry_widget.configure(state="normal")
            
            # ลบข้อความเก่าและใส่ข้อความใหม่
            entry_widget.delete(0, "end")
            entry_widget.insert(0, str(text))
            
            # ทำให้กลับไปเป็นแบบอ่านอย่างเดียวเหมือนเดิม
            if current_state == "readonly":
                entry_widget.configure(state="readonly")
        except Exception as e:
            # ป้องกันโปรแกรมแครชถ้าเกิดข้อผิดพลาดที่ไม่คาดคิด
            print(f"Error in set_entry_text for widget {entry_widget}: {e}")