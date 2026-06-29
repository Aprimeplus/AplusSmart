# login_screen.py (ฉบับสมบูรณ์พร้อมฟังก์ชันดูรหัสผ่าน)

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
    return os.path.join(base_path, relative_path) # แนะนำให้เก็บรูปใน resources

class LoginScreen(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="#EBF5FF")
        self.app_container = app_container
        self.pack(fill="both", expand=True)

        import tkinter as _tk
        canvas = _tk.Canvas(self, highlightthickness=0, bd=0)
        canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.bind("<Configure>", lambda e: self._draw_bg(canvas))
        self._draw_bg(canvas)

        self.main_frame = CTkFrame(self, fg_color="#FFFFFF", corner_radius=20,
                                   border_width=1, border_color="#C8DFF7")
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center")
        self.main_frame.pack(expand=True, ipadx=80, ipady=36)

        try:
            self.user_icon = CTkImage(Image.open(resource_path("user_icon.png")), size=(20, 20))
            self.lock_icon = CTkImage(Image.open(resource_path("lock_icon.png")), size=(20, 20))
            self.eye_open_icon = CTkImage(Image.open(resource_path("eye_open.png")), size=(20, 20))
            self.eye_closed_icon = CTkImage(Image.open(resource_path("eye_closed.png")), size=(20, 20))
        except FileNotFoundError as e:
            self.user_icon, self.lock_icon, self.eye_open_icon, self.eye_closed_icon = None, None, None, None
            print(f"Warning: Icon file not found. {e}. Icons will not be displayed.")

        try:
            logo_path = resource_path("company_logo.png")
            pil_image = Image.open(logo_path)
            logo_image = CTkImage(light_image=pil_image, dark_image=pil_image, size=(165, 165))
            logo_label = CTkLabel(self.main_frame, image=logo_image, text="")
            logo_label.pack(pady=(28, 12), padx=60)
        except Exception as e:
            print(f"Warning: Could not load logo: {e}")

        welcome_font = CTkFont(size=24, weight="bold", family="Tahoma")

        CTkLabel(self.main_frame, text="A+ Smart Solution", font=welcome_font, text_color="#1E3A5C").pack(pady=(0, 4))
        CTkLabel(self.main_frame, text="Sales & Operations Management System",
                 font=CTkFont(size=12, weight="bold", family="Tahoma"), text_color="#4A7AAF").pack(pady=(0, 20))

        username_frame = CTkFrame(self.main_frame, fg_color="#EFF6FF", corner_radius=8,
                                   border_width=2, border_color="#3B82F6")
        username_frame.pack(fill="x", padx=40)
        if self.user_icon:
            CTkLabel(username_frame, image=self.user_icon, text="", fg_color="transparent").pack(side="left", padx=(12, 8))

        self.user_key_entry = CTkEntry(username_frame, height=46, border_width=0, fg_color="#EFF6FF",
                                       placeholder_text="Username", font=CTkFont(size=14),
                                       text_color="#1E3A5C", placeholder_text_color="#9CA3AF")
        self.user_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=2)
        self.user_key_entry.bind("<Return>", lambda event: self.password_entry.focus_set())
        self.user_key_entry.bind("<KP_Enter>", lambda event: self.password_entry.focus_set())
        self.user_key_entry.bind("<FocusIn>",  lambda e: username_frame.configure(border_color="#1D4ED8"))
        self.user_key_entry.bind("<FocusOut>", lambda e: username_frame.configure(border_color="#3B82F6"))

        password_frame = CTkFrame(self.main_frame, fg_color="#EFF6FF", corner_radius=8,
                                   border_width=2, border_color="#3B82F6")
        password_frame.pack(fill="x", padx=40, pady=12)
        if self.lock_icon:
            CTkLabel(password_frame, image=self.lock_icon, text="", fg_color="transparent").pack(side="left", padx=(12, 8))

        self.password_entry = CTkEntry(password_frame, height=46, border_width=0, fg_color="#EFF6FF",
                                       placeholder_text="Password", show="*", font=CTkFont(size=14),
                                       text_color="#1E3A5C", placeholder_text_color="#9CA3AF")
        self.password_entry.pack(side="left", fill="x", expand=True, padx=(0, 5), pady=2)
        self.password_entry.bind("<Return>", self.login_event)
        self.password_entry.bind("<FocusIn>",  lambda e: password_frame.configure(border_color="#1D4ED8"))
        self.password_entry.bind("<FocusOut>", lambda e: password_frame.configure(border_color="#3B82F6"))

        if self.eye_closed_icon:
            self.show_password_button = CTkButton(
                password_frame, text="", image=self.eye_closed_icon,
                width=30, height=30, fg_color="transparent",
                hover_color="#DBEAFE", command=self._toggle_password_visibility
            )
            self.show_password_button.pack(side="right", padx=(0, 10))

        button_font = CTkFont(size=15, weight="bold", family="Tahoma")
        self.login_button = CTkButton(self.main_frame, text="เข้าสู่ระบบ", command=self.login,
                                      height=48, font=button_font, corner_radius=10,
                                      fg_color="#2563EB", hover_color="#1D4ED8", text_color="#FFFFFF")
        self.login_button.pack(pady=(16, 20), padx=40, fill="x")

        CTkLabel(self.main_frame, text="Build 2026.06",
                 font=CTkFont(size=10, family="Tahoma"),
                 text_color="#93C5FD").pack(pady=(0, 10))

    def _draw_bg(self, canvas):
        try:
            canvas.delete("bg")
            w = self.winfo_width() or 800
            h = self.winfo_height() or 600
            steps = 50
            # #DBEAFE (ฟ้าอ่อน) → #EBF5FF (ขาวฟ้า)
            for i in range(steps):
                r = int(0xDB + (0xEB - 0xDB) * i / steps)
                g = int(0xEA + (0xF5 - 0xEA) * i / steps)
                b = int(0xFE + (0xFF - 0xFE) * i / steps)
                color = f"#{r:02x}{g:02x}{b:02x}"
                y0 = h * i // steps
                y1 = h * (i + 1) // steps
                canvas.create_rectangle(0, y0, w, y1, fill=color, outline="", tags="bg")
        except Exception:
            pass

    def _toggle_password_visibility(self):
        """สลับการแสดง/ซ่อนรหัสผ่าน และเปลี่ยนไอคอนปุ่ม"""
        if self.password_entry.cget("show") == "*":
            # ถ้ากำลังซ่อนอยู่ -> ให้แสดงรหัสผ่าน
            self.password_entry.configure(show="")
            # เปลี่ยนไอคอนเป็นรูปตาเปิด
            if self.eye_open_icon:
                self.show_password_button.configure(image=self.eye_open_icon)
        else:
            # ถ้ากำลังแสดงอยู่ -> ให้กลับไปซ่อน
            self.password_entry.configure(show="*")
            # เปลี่ยนไอคอนเป็นรูปตาปิด
            if self.eye_closed_icon:
                self.show_password_button.configure(image=self.eye_closed_icon)

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
                            self.app_container.show_main_app(sale_key=user_key, sale_name=user_name, user_role=user_role)
                        elif user_role in ['Purchasing Staff', 'ฝ่ายจัดซื้อ']:
                            self.app_container.show_purchasing_screen(user_key=user_key, user_name=user_name, user_role=user_role)
                        elif user_role == 'Purchasing Manager':
                            self.app_container.show_purchasing_manager_screen(user_key=user_key, user_name=user_name, user_role=user_role)
                        elif user_role == 'Director':
                            self.app_container.show_director_screen(user_key, user_name, user_role)
                        elif user_role == 'Sales Manager':
                            self.app_container.show_sales_manager_screen(user_key, user_name, user_role)
                        elif user_role == 'Sale Support':
                            self.app_container.show_sale_support_screen(user_key=user_key, user_name=user_name, user_role=user_role)
                        elif user_role == 'HR':
                            self.app_container.show_hr_screen(user_key=user_key, user_name=user_name, user_role=user_role)
                        elif user_role == 'Transport Admin':
                            self.app_container.show_transport_admin_screen(user_key=user_key, user_name=user_name, user_role=user_role)
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