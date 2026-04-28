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
                # แก้ไข: ส่ง result เป็น argument ไปโดยตรง ไม่ใช้ lambda
                self.master.after(0, on_success, result)
            except Exception as e:
                # แก้ไข: ส่ง e เป็น argument ไปโดยตรง (แก้ปัญหา e ถูกลบหลังจบ except)
                self.master.after(0, on_error, e)
        
        thread = threading.Thread(target=task_wrapper)
        thread.daemon = True
        thread.start()

def show_loading_message(parent, text="Loading..."):
    """
    แสดงข้อความ "กำลังโหลด" โดยใช้ .place() เพื่อไม่ให้ขัดกับ .grid()
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
    # ใช้ .place() เพื่อวาง widget ไว้กลางหน้าจอและอยู่บนสุดเสมอ
    loading_label.place(relx=0.5, rely=0.5, anchor="center")
    loading_label.lift() # ทำให้ Label อยู่ชั้นบนสุด
    parent.update_idletasks() # บังคับให้ UI อัปเดตทันที
    return loading_label

def hide_loading_message(loading_label):
    """
    ซ่อนข้อความ "กำลังโหลด" โดยใช้ .place_forget()
    """
    if loading_label and loading_label.winfo_exists():
        loading_label.place_forget()