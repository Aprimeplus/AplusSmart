# dialog_windows.py

import tkinter as tk
from customtkinter import (CTkToplevel, CTkScrollableFrame, CTkLabel, CTkFont, 
                           CTkFrame, CTkButton, CTkEntry, CTkCheckBox)
from tkinter import messagebox

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