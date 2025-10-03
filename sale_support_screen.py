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
        # เรียกใช้ __init__ ของ SalesProxyScreen และส่งค่าที่จำเป็นไป
        super().__init__(master=master,
                         app_container=app_container,
                         proxy_user_key=user_key,    # รหัสของ Sale Support ที่ Login
                         proxy_user_name=user_name,  # ชื่อของ Sale Support ที่ Login
                         role_to_proxy="Sale")       # ทำงานแทน "Sale"