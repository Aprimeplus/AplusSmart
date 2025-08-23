# login_screen.py (ฉบับแก้ไขล่าสุด)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkEntry, CTkFont, CTkButton, CTkImage
from tkinter import messagebox
import psycopg2
import psycopg2.extras
from PIL import Image
import bcrypt
import os
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class LoginScreen(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="#A6D7F3") # พื้นหลังหลักเป็นสีน้ำเงิน
        self.app_container = app_container
        self.pack(fill="both", expand=True)

        # Frame หลักสำหรับเนื้อหา เปลี่ยนเป็นสีขาว
        self.main_frame = CTkFrame(self, fg_color="white", corner_radius=15)
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center")
        self.main_frame.pack(expand=True, ipadx=80, ipady=70)
        
        try:
            self.user_icon = CTkImage(Image.open(resource_path("user_icon.png")), size=(20, 20))
            self.lock_icon = CTkImage(Image.open(resource_path("lock_icon.png")), size=(20, 20))
        except FileNotFoundError:
            self.user_icon = None
            self.lock_icon = None
            print("Warning: user_icon.png or lock_icon.png not found. Icons will not be displayed.")

        try:
            logo_path = resource_path("company_logo.png")
            pil_image = Image.open(logo_path)
            logo_image = CTkImage(light_image=pil_image, dark_image=pil_image, size=(120, 120))
            logo_label = CTkLabel(self.main_frame, image=logo_image, text="")
            logo_label.pack(pady=(30, 15), padx=60)
        except Exception as e:
            print(f"Warning: Could not load logo: {e}")

        welcome_font = CTkFont(size=28, weight="bold", family="Roboto")
        subtitle_font = CTkFont(size=16, family="Roboto")
        
        # --- START: แก้ไขข้อความต้อนรับเป็นชื่อแอป ---
        CTkLabel(self.main_frame, text="A+ Smart Solution", font=welcome_font, text_color="#2A2D34").pack(pady=(0, 5))
        # --- END: แก้ไขข้อความต้อนรับเป็นชื่อแอป ---
        
        CTkLabel(self.main_frame, text="Login to your account", font=subtitle_font, text_color="#8A91A0").pack(pady=(0, 30))

        username_frame = CTkFrame(self.main_frame, fg_color="#F7F7F9", corner_radius=8)
        username_frame.pack(fill="x", padx=40)
        if self.user_icon:
            CTkLabel(username_frame, image=self.user_icon, text="").pack(side="left", padx=(15, 10))
        
        self.user_key_entry = CTkEntry(username_frame, height=50, border_width=0, fg_color="transparent",
                                       placeholder_text="Username", font=CTkFont(size=14))
        self.user_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 15))

        password_frame = CTkFrame(self.main_frame, fg_color="#F7F7F9", corner_radius=8)
        password_frame.pack(fill="x", padx=40, pady=15)
        if self.lock_icon:
            CTkLabel(password_frame, image=self.lock_icon, text="").pack(side="left", padx=(15, 10))

        self.password_entry = CTkEntry(password_frame, height=50, border_width=0, fg_color="transparent",
                                       placeholder_text="Password", show="*", font=CTkFont(size=14))
        self.password_entry.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.password_entry.bind("<Return>", self.login_event)

        button_font = CTkFont(size=16, weight="bold", family="Roboto")
        self.login_button = CTkButton(self.main_frame, text="Sign in", command=self.login,
                                      height=50, font=button_font, corner_radius=10)
        self.login_button.pack(pady=(20, 30), padx=40, fill="x")

    def login_event(self, event=None):
        self.login()

    def login(self):
        user_key = self.user_key_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not user_key or not password:
            messagebox.showwarning("ข้อมูลว่างเปล่า", "กรุณากรอกรหัสผู้ใช้งานและรหัสผ่าน", parent=self)
            return
        
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT sale_name, role, password_hash FROM sales_users WHERE sale_key = %s AND status = 'Active'", (user_key,))
                result = cursor.fetchone()
                
                if result:
                    user_name, user_role, stored_hash = result['sale_name'], result['role'], result['password_hash']
                    
                    if stored_hash and bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
                        if user_role == 'Sale':
                            self.app_container.show_main_app(sale_key=user_key, sale_name=user_name)
                        elif user_role in ['Purchasing Staff', 'ฝ่ายจัดซื้อ']:
                            self.app_container.show_purchasing_screen(user_key=user_key, user_name=user_name)
                        elif user_role == 'Purchasing Manager':
                            self.app_container.show_purchasing_manager_screen(user_key=user_key, user_name=user_name, user_role=user_role)
                        elif user_role == 'Director':
                            self.app_container.show_director_screen(user_key, user_name, user_role)
                        elif user_role == 'Sales Manager': # <--- เพิ่มเงื่อนไขนี้
                            self.app_container.show_sales_manager_screen(user_key, user_name, user_role) # <--- เรียกฟังก์ชันใหม่
                        elif user_role == 'HR':
                            self.app_container.show_hr_screen(user_key=user_key, user_name=user_name)
                        else:
                            messagebox.showerror("ข้อผิดพลาด", f"ไม่รู้จักประเภทผู้ใช้: {user_role}", parent=self)
                    else:
                        messagebox.showerror("รหัสไม่ถูกต้อง", "รหัสผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง", parent=self)
                        self.password_entry.delete(0, tk.END)
                else:
                    messagebox.showerror("รหัสไม่ถูกต้อง", "รหัสผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง", parent=self)
                    self.user_key_entry.delete(0, tk.END)
                    self.password_entry.delete(0, tk.END)

        except (Exception, psycopg2.Error) as e:
             messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการเชื่อมต่อ: {e}", parent=self)
        finally:
            self.app_container.release_connection(conn)