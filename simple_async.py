# simple_async.py (เวอร์ชันแก้ไข TclError)

import threading
import customtkinter as ctk
from customtkinter import CTkFont

class SimpleAsyncHelper:
    """
    คลาสสำหรับช่วยรันฟังก์ชันใน background thread เพื่อไม่ให้หน้าจอค้าง
    """
    def __init__(self, master_widget):
        self.master = master_widget

    def run_in_background(self, work_func, on_success, on_error):
        def task_wrapper():
            try:
                result = work_func()
                # เมื่อทำงานเสร็จ ส่งผลลัพธ์กลับไปรันใน main thread ผ่าน .after()
                self.master.after(0, lambda: on_success(result))
            except Exception as e:
                # หากเกิด error ส่ง error กลับไปรันใน main thread
                self.master.after(0, lambda: on_error(e))
        
        thread = threading.Thread(target=task_wrapper)
        thread.daemon = True
        thread.start()

def show_loading_message(parent, text="Loading..."):
    """
    (เวอร์ชันแก้ไข) แสดงข้อความ "กำลังโหลด" โดยใช้ .place() เพื่อไม่ให้ขัดกับ .grid()
    """
    loading_label = ctk.CTkLabel(
        parent,
        text=text,
        font=CTkFont(size=20, weight="bold", slant="italic"),
        fg_color=("gray85", "gray20"),
        corner_radius=10,
        width=300,
        height=100
    )
    # <<< START: จุดที่แก้ไข >>>
    # ใช้ .place() เพื่อวาง widget ไว้กลางหน้าจอและอยู่บนสุดเสมอ
    loading_label.place(relx=0.5, rely=0.5, anchor="center")
    loading_label.lift() # ทำให้ Label อยู่ชั้นบนสุด
    parent.update_idletasks() # บังคับให้ UI อัปเดตทันที
    # <<< END: สิ้นสุดการแก้ไข >>>
    return loading_label

def hide_loading_message(loading_label):
    """
    (เวอร์ชันแก้ไข) ซ่อนข้อความ "กำลังโหลด" โดยใช้ .place_forget()
    """
    if loading_label and loading_label.winfo_exists():
        # <<< แก้ไข: เปลี่ยนจาก .pack_forget() เป็น .place_forget() >>>
        loading_label.place_forget()
