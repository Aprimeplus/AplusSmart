import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox

class CancellationReasonDialog(ctk.CTkToplevel):
    def __init__(self, master, on_confirm_callback):
        super().__init__(master)
        self.title("ระบุสาเหตุการยกเลิก SO")
        self.geometry("400x450")
        self.on_confirm_callback = on_confirm_callback
        
        # สาเหตุมาตรฐาน
        self.reasons = [
            "สินค้าผิดสเปค",
            "Stock ไม่เพียงพอ",
            "การจัดส่งล่าช้าเกินกำหนด",
            "อื่นๆ (ระบุ)"
        ]
        
        self.selected_reason = tk.StringVar(value="")
        
        ctk.CTkLabel(self, text="กรุณาเลือกสาเหตุที่ต้องการยกเลิก:", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(20, 10))
        
        # สร้าง Radio Button (ใช้ Radio เพราะสาเหตุหลักควรมีข้อเดียว)
        # ถ้าอยากได้ Checkbox หลายข้อ บอกได้ครับ แต่ Radio จัดการ Logic "อื่นๆ" ง่ายกว่า
        for reason in self.reasons:
            rb = ctk.CTkRadioButton(
                self, 
                text=reason, 
                variable=self.selected_reason, 
                value=reason,
                command=self._on_radio_select
            )
            rb.pack(anchor="w", padx=40, pady=5)
            
        # ช่องกรอกกรณีเลือก "อื่นๆ"
        self.other_reason_entry = ctk.CTkTextbox(self, height=100, state="disabled")
        self.other_reason_entry.pack(fill="x", padx=40, pady=10)
        
        # ปุ่มยืนยัน
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        ctk.CTkButton(btn_frame, text="ยืนยันการยกเลิก", fg_color="#DC2626", hover_color="#991B1B", command=self._confirm).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="ปิด", fg_color="gray", command=self.destroy).pack(side="left", padx=10)
        
        self.transient(master)
        self.grab_set()

    def _on_radio_select(self):
        """เปิด/ปิดช่องกรอกข้อความตามตัวเลือก"""
        if self.selected_reason.get() == "อื่นๆ (ระบุ)":
            self.other_reason_entry.configure(state="normal")
            self.other_reason_entry.focus_set()
        else:
            self.other_reason_entry.delete("1.0", "end")
            self.other_reason_entry.configure(state="disabled")

    def _confirm(self):
        reason = self.selected_reason.get()
        if not reason:
            messagebox.showwarning("เตือน", "กรุณาเลือกสาเหตุ", parent=self)
            return
            
        final_reason = reason
        if reason == "อื่นๆ (ระบุ)":
            other_text = self.other_reason_entry.get("1.0", "end-1c").strip()
            if not other_text:
                messagebox.showwarning("เตือน", "กรุณาระบุรายละเอียดเพิ่มเติม", parent=self)
                return
            final_reason = f"อื่นๆ: {other_text}"
            
        # ส่งค่ากลับไปที่ Callback
        self.on_confirm_callback(final_reason)
        self.destroy()