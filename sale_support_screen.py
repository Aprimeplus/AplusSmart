# sale_support_screen.py (ฉบับแก้ไขสมบูรณ์)

import tkinter as tk
from customtkinter import CTkFrame, CTkLabel, CTkFont, CTkOptionMenu, CTkButton
from tkinter import messagebox
import pandas as pd
from sales_proxy_screen import SalesProxyScreen
from commission_app import CommissionApp, SubmitSODialog # <-- เพิ่ม SubmitSODialog

class SaleSupportApp(SalesProxyScreen):
    """
    หน้าจอสำหรับ Sale Support ที่สืบทอดความสามารถมาจาก SalesProxyScreen
    """
    def __init__(self, master, app_container, user_key, user_name, user_role):
        super().__init__(master=master,
                        app_container=app_container,
                        proxy_user_key=user_key,
                        proxy_user_name=user_name,
                        role_to_proxy="Sale",
                        show_logout_button=True) # <--- เพิ่มบรรทัดนี้ครับ