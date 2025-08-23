# main_app.py (Final Corrected Version)

import matplotlib
import tkinter as tk
from customtkinter import set_appearance_mode, CTk, CTkToplevel, CTkLabel, CTkFont, CTkFrame, CTkImage, CTkProgressBar
import psycopg2
from psycopg2.extras import DictCursor
from psycopg2 import pool
from tkinter import messagebox
from datetime import datetime, timedelta
import threading
import time
import os
import sys
from PIL import Image, ImageTk
import ctypes
from sqlalchemy import create_engine
import pandas as pd
import traceback

# --- START: เพิ่ม Import ที่จำเป็น ---
from so_selection_dialog import SOSelectionDialog
import po_document_generator
# --- END: สิ้นสุดการเพิ่ม Import ---

# We keep this part for type hinting, which helps the code editor but doesn't run
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from login_screen import LoginScreen
    from commission_app import CommissionApp
    from purchasing_screen import PurchasingScreen
    from hr_screen import HRScreen
    from history_windows import PurchaseHistoryWindow, CommissionHistoryWindow, PurchaseDetailWindow, SalesDataViewerWindow
    from edit_commission_window import EditCommissionWindow
    from purchasing_manager_screen import PurchasingManagerScreen
    from sales_manager_screen import SalesManagerScreen
    from director_screen import DirectorScreen


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class NotificationPopup(CTkToplevel):
    def __init__(self, master, title, message, details=""):
        super().__init__(master)
        self.lift()
        self.attributes("-topmost", True)
        self.overrideredirect(True)
        main_frame = CTkFrame(self, corner_radius=10, border_width=2, border_color="#3B82F6")
        main_frame.pack(padx=2, pady=2, fill="both", expand=True)
        CTkLabel(main_frame, text=title, font=CTkFont(size=16, weight="bold")).pack(anchor="w", padx=15, pady=(10, 2))
        CTkLabel(main_frame, text=message, font=CTkFont(size=14), wraplength=380, justify="left").pack(anchor="w", padx=15, pady=(0, 10))
        self.after(7000, self.destroy)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = 400
        height = 120
        x = screen_width - width - 20
        y = screen_height - height - 60
        self.geometry(f"{width}x{height}+{x}+{y}")

class LoadingWindow(CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Loading")
        self.geometry("400x220")
        self.overrideredirect(True)
        self.resizable(False, False)
        master.update_idletasks()
        x = master.winfo_x() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_y() + (master.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        main_frame = CTkFrame(self, fg_color="#FFFFFF", corner_radius=10)
        main_frame.pack(fill="both", expand=True, padx=2, pady=2)
        try:
            logo_path = resource_path("company_logo.png")
            pil_image = Image.open(logo_path)
            logo_image = CTkImage(light_image=pil_image, dark_image=pil_image, size=(80, 80))
            logo_label = CTkLabel(main_frame, image=logo_image, text="")
            logo_label.pack(pady=(20, 10))
        except Exception as e:
            print(f"Could not load logo on loading screen: {e}")
        self.label = CTkLabel(main_frame, text="กำลังโหลดข้อมูล...\nกรุณารอสักครู่", font=CTkFont(size=18, family="Roboto"), text_color="#374151")
        self.label.pack(pady=5)
        self.progressbar = CTkProgressBar(main_frame, mode='indeterminate', height=10)
        self.progressbar.pack(pady=(10, 20), padx=40, fill="x")
        self.progressbar.start()
        self.lift()
        self.grab_set()

    def stop_animation(self):
        self.progressbar.stop()

class AppContainer(CTk):
    def __init__(self):
        super().__init__()
        self.hr_screen = None
        try:
            icon_image = Image.open(resource_path("app_icon.ico"))
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.iconphoto(True, icon_photo)
        except Exception as e:
            print(f"Failed to set iconphoto: {e}")

        myappid = 'mycompany.myapplication.subproduct.version'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        self.title("A+ Smart Solution")
        try:
            self.iconbitmap(resource_path("app_icon.ico"))
        except tk.TclError:
            print("ไม่พบไฟล์ไอคอน app_icon.ico หรือไฟล์ไม่ใช่รูปแบบที่ถูกต้อง")
        self.geometry("1600x900")
        set_appearance_mode("Light")
        self.THEME = {"sale": {"primary": "#3B82F6", "header": "#1D4ED8", "bg": "#EFF6FF", "row": "#EFF6FF"}, "hr": {"primary": "#16A34A", "header": "#15803D", "bg": "#F0FDF4"}, "purchasing": {"primary": "#7C3AED", "header": "#6D28D9", "bg": "#F5F3FF", "row": "#F5F3FF"}}
        self.db_pool = None
        self.HEADER_MAP = { 
                'id': 'ID', 'timestamp': 'เวลาบันทึก', 'status': 'สถานะ', 'is_active': 'Active', 'original_id': 'Original ID', 'so_number': 'เลขที่ SO', 'bill_date': 'วันที่บิล', 'customer_id': 'รหัสลูกค้า', 'customer_name': 'ชื่อลูกค้า', 'customer_type': 'ประเภทลูกค้า', 'credit_term': 'เครดิต', 'sales_service_amount': 'ยอดขาย/บริการ','payment_date': 'วันที่ชำระหลัก', 'total_payment_amount': 'ยอดชำระรวม', 'payment_before_vat': 'ยอดชำระก่อน VAT','payment_no_vat': 'ยอดชำระไม่มี VAT', 'difference_amount': 'ผลต่าง', 'vat_deduction': 'หัก ณ ที่จ่าย (VAT)','no_vat_deduction': 'หัก ณ ที่จ่าย (ไม่มี VAT)', 'shipping_cost': 'ค่าขนส่ง', 'delivery_date': 'วันที่จัดส่ง','separate_shipping_charge': 'ค่ารถเก็บเงินแยก', 'brokerage_fee': 'ค่านายหน้า', 'giveaways': 'ของแถม','coupons': 'คูปอง', 'transfer_fee': 'ค่าธรรมเนียมโอน', 'credit_card_fee': 'ค่าธรรมเนียมบัตร','wht_3_percent': 'ภาษีหัก ณ ที่จ่าย 3%', 'commission_month': 'เดือนคอมมิชชั่น', 'commission_year': 'ปีคอมมิชชั่น','final_commission': 'คอมมิชชั่นสุดท้าย', 'product_vat_7': 'VAT สินค้า 7%', 'shipping_vat_7': 'VAT ขนส่ง 7%', 'sales_uploaded': 'ยอดขาย (Express)', 'margin_db': 'Margin (ระบบ) %', 'margin_uploaded': 'Margin (Express) %', 'cogs_db': 'ต้นทุน (PU)','cost_db': 'ต้นทุน (PU)', 'cost_uploaded': 'ต้นทุน (Express)','cutting_drilling_fee': 'ค่าบริการตัด/เจาะ','cutting_drilling_fee_vat_option': 'ประเภท VAT ตัด/เจาะ', 'other_service_fee': 'ค่าบริการอื่นๆ','other_service_fee_vat_option': 'ประเภท VAT อื่นๆ', 'sales_service_vat_option': 'ประเภท VAT ยอดขาย','shipping_vat_option': 'ประเภท VAT ค่าส่ง', 'credit_card_fee_vat_option': 'ประเภท VAT บัตร','cash_product_input': 'ยอดสินค้าเงินสด', 'cash_service_total': 'ยอดบริการเงินสด','cash_required_total': 'ยอดต้องชำระเงินสด', 'cash_actual_payment': 'ยอดชำระเงินสดจริง','payment1_date': 'วันที่ชำระ1', 'payment1_method': 'วิธีชำระ1', 'payment2_date': 'วันที่ชำระ2','payment2_method': 'วิธีชำระ2', 'delivery_type': 'ประเภทการจัดส่ง', 'pickup_location': 'สถานที่รับ','relocation_cost': 'ค่าย้าย', 'date_to_warehouse': 'วันที่เข้าคลัง', 'date_to_customer': 'วันที่ส่งลูกค้า','pickup_registration': 'ทะเบียนเข้ารับ', 'department': 'แผนก', 'pur_order': 'PUR Order', 'supplier_name': 'ชื่อซัพพลายเออร์', 'po_number': 'เลขที่ PO', 'rr_number': 'เลขที่ RR', 'po_date': 'วันที่สร้าง PO','po_total_payable': 'ยอดชำระ PO', 'po_creator_key': 'ผู้สร้าง PO', 'sale_name': 'พนักงานขาย','commission_plan': 'แผนค่าคอมฯ', 'sales_target': 'ยอดเป้าหมาย', 'status_db': 'สถานะ (DB)','status_file': 'สถานะ (ไฟล์)', 'user_key': 'รหัสผู้ใช้',
                'rejection_reason': 'เหตุผลที่ปฏิเสธ',
                'last_modified_by': 'แก้ไขล่าสุดโดย',
                'total_cost': 'ต้นทุนรวม (ไม่รวม VAT)',
                'total_weight': 'น้ำหนักรวม (กก.)',
                'wht_3_percent_checked': 'หัก ณ ที่จ่าย 3% (Y/N)',
                'wht_3_percent_amount': 'ยอดหัก ณ ที่จ่าย 3%',
                'vat_7_percent_checked': 'VAT 7% (Y/N)',
                'vat_7_percent_amount': 'ยอด VAT 7%',
                'grand_total': 'ยอดรวมสุทธิที่ต้องชำระ',
                'approval_status': 'สถานะการอนุมัติ',
                'po_mode': 'โหมด PO',
                'approver_manager1_key': 'ผู้อนุมัติ 1',
                'approval_date_manager1': 'วันที่อนุมัติ 1',
                'approver_manager2_key': 'ผู้อนุมัติ 2',
                'approval_date_manager2': 'วันที่อนุมัติ 2',
                'approver_hr_key': 'ผู้อนุมัติ (HR)',
                'approval_date_hr': 'วันที่อนุมัติ (HR)',
                'shipping_to_stock_vat_type': 'ประเภท VAT (ส่งเข้าสต็อก)',
                'shipping_to_site_vat_type': 'ประเภท VAT (ส่งเข้าไซต์)',
                'shipping_to_stock_shipper': 'ผู้จัดส่ง (เข้าสต็อก)',
                'shipping_to_site_shipper': 'ผู้จัดส่ง (เข้าไซต์)',
                'shipping_to_stock_date': 'วันที่ส่ง (เข้าสต็อก)',
                'shipping_to_stock_notes': 'หมายเหตุ (เข้าสต็อก)',
                'shipping_to_site_date': 'วันที่ส่ง (เข้าไซต์)',
                'shipping_to_site_notes': 'หมายเหตุ (เข้าไซต์)',
                'approver_manager3_key': 'ผู้อนุมัติ 3',
                'approval_date_manager3': 'วันที่อนุมัติ 3',
                'approver_director_key': 'ผู้อนุมัติ (ผู้บริหาร)',
                'approval_date_director': 'วันที่อนุมัติ (ผู้บริหาร)',
                'shipping_to_stock_cost': 'ค่าขนส่ง (ส่งเข้าสต็อก)',
                'shipping_to_site_cost': 'ค่าขนส่ง (ส่งเข้าไซต์)',
                # เพิ่มคอลัมน์ใหม่จากภาพล่าสุดที่คุณส่งมา
                'hired_truck_cost': 'ค่ารถขนส่ง',
                'shipping_company_name': 'บริษัทขนส่ง',
                'truck_name': 'ชื่อรถ',
                'actual_payment_amount': 'ยอดชำระจริง',
                'actual_payment_date': 'วันที่ชำระจริง',
                'supplier_account_number': 'เลขที่บัญชีซัพพลายเออร์',
                'shipping_type': 'ประเภทการขนส่ง',
                'shipping_notes': 'หมายเหตุการขนส่ง' ,
                'final_sales_amount': 'ยอดขายสุดท้าย',
                'final_cost_amount': 'ต้นทุนสุดท้าย',
                'final_gp': 'กำไรขั้นต้น (GP)',
                'final_margin': 'Margin สุดท้าย (%)',
                'final_commission': 'คอมมิชชั่นสุดท้าย',
            }
        self.pg_engine = None
        self.current_user_key = None
        self.notification_poll_id = None
        try:
            db_params = {"host": "192.168.1.60", "dbname": "aplus_com_test", "user": "app_user", "password": "cailfornia123"}
            self.db_pool = psycopg2.pool.SimpleConnectionPool(1, 10, **db_params)
            self.pg_engine = create_engine(f'postgresql+psycopg2://{db_params["user"]}:{db_params["password"]}@{db_params["host"]}:5432/{db_params["dbname"]}')
            conn = self.get_connection()
            print("Database connection pool created successfully.")
            self.release_connection(conn)
            
            self.stop_background_threads = threading.Event()
            
            submission_thread = threading.Thread(target=self._run_background_task, args=(self._auto_submit_commissions, 3600), daemon=True)
            submission_thread.start()

            po_submission_thread = threading.Thread(target=self._run_background_task, args=(self._auto_submit_overdue_pos, 3600), daemon=True)
            po_submission_thread.start()

        except psycopg2.OperationalError as e:
            messagebox.showerror("Connection Error", f"ไม่สามารถเชื่อมต่อฐานข้อมูล PostgreSQL ได้:\n{e}")
            self.after(100, self.destroy)
            return
        except Exception as e:
            messagebox.showerror("Database Setup Error", f"เกิดข้อผิดพลาดในการตั้งค่าฐานข้อมูล:\n{e}")
            self.after(100, self.destroy)
            return
        self._create_initial_db_tables()
        self.show_login_screen()

    
    def _run_background_task(self, task_function, interval_seconds):
        while not self.stop_background_threads.is_set():
            try:
                task_function()
            except Exception as e:
                print(f"Error in background task {task_function.__name__}: {e}")
            self.stop_background_threads.wait(interval_seconds)
    
    def _auto_submit_overdue_pos(self):
        print("Running auto-submit check for overdue POs...")
        conn = None
        try:
            conn = self.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                twenty_four_hours_ago = datetime.now() - timedelta(hours=24)
                cursor.execute("SELECT id, so_number, user_key FROM commissions WHERE status = 'PO In Progress' AND claim_timestamp < %s", (twenty_four_hours_ago,))
                overdue_sos = cursor.fetchall()
                if not overdue_sos:
                    print("No overdue SOs found.")
                    return
                for so in overdue_sos:
                    so_number = so['so_number']
                    print(f"Found overdue SO: {so_number}. Submitting related draft POs...")
                    cursor.execute("SELECT id, po_number FROM purchase_orders WHERE so_number = %s AND status = 'Draft'", (so_number,))
                    draft_pos = cursor.fetchall()
                    if not draft_pos:
                        print(f"No draft POs found for overdue SO: {so_number}. Skipping.")
                        continue
                    draft_po_ids = [po['id'] for po in draft_pos]
                    psycopg2.extras.execute_values(cursor, "UPDATE purchase_orders SET status = 'Pending Approval', approval_status = 'Pending Mgr 1' WHERE id IN %s", [(draft_po_ids,)])
                    cursor.execute("UPDATE commissions SET status = 'PO Sent' WHERE id = %s", (so['id'],))
                    cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Manager' AND status = 'Active'")
                    manager_keys = [row['sale_key'] for row in cursor.fetchall()]
                    notif_data = []
                    for po in draft_pos:
                        message = f"AUTO-SUBMIT: PO ({po['po_number']}) ถูกส่งอัตโนมัติเนื่องจากเกินกำหนด 24 ชม."
                        for manager_key in manager_keys:
                            notif_data.append((manager_key, message, False, po['id']))
                    if notif_data:
                        psycopg2.extras.execute_values(cursor, "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES %s", notif_data)
                    conn.commit()
                    print(f"Successfully auto-submitted {len(draft_po_ids)} POs for SO: {so_number}.")
        except Exception as e:
            print(f"Error during auto-submission of overdue POs: {e}")
            if conn: conn.rollback()
        finally:
            self.release_connection(conn)

    def _check_for_notifications(self):
        if not self.current_user_key: return
        conn = None
        try:
            conn = self.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT id, message FROM notifications WHERE user_key_to_notify = %s AND is_read = FALSE", (self.current_user_key,))
                new_notifications = cursor.fetchall()
                if new_notifications:
                    for notif in new_notifications:
                        NotificationPopup(self, title="📬 ท่านมีข้อความใหม่", message=notif['message'])
                        cursor.execute("UPDATE notifications SET is_read = TRUE WHERE id = %s", (notif['id'],))
                    conn.commit()
                    if hasattr(self, 'current_screen') and self.current_screen is not None and hasattr(self.current_screen, '_update_tasks_badge'):
                       self.current_screen._update_tasks_badge()
        except Exception as e:
            print(f"Error checking for notifications: {e}")
            if conn: conn.rollback()
        finally:
            if conn: self.release_connection(conn)
        self.notification_poll_id = self.after(15000, self._check_for_notifications)

    def get_connection(self):
        if self.db_pool: return self.db_pool.getconn()
        return None

    def release_connection(self, conn):
        if self.db_pool and conn: self.db_pool.putconn(conn)

    def _auto_submit_commissions(self):
        now = datetime.now()
        conn = None
        try:
            conn = self.get_connection()
            with conn.cursor() as cursor:
                try:
                    inbound_deadline = now.replace(day=27, hour=17, minute=30, second=0, microsecond=0)
                    if now > inbound_deadline:
                        print(f"Inbound deadline passed. Submitting records to 'Pending PU'...")
                        sql_inbound = "UPDATE commissions c SET status = 'Pending PU' FROM sales_users su WHERE c.sale_key = su.sale_key AND c.status = 'Original' AND (su.sale_type IS NULL OR su.sale_type != 'Outbound');"
                        cursor.execute(sql_inbound)
                        if cursor.rowcount > 0:
                            conn.commit()
                            print(f"Auto-submitted {cursor.rowcount} Inbound/Other records to Pending PU.")
                        else:
                            conn.rollback()
                except ValueError:
                    print("Could not create inbound deadline (e.g., Feb has < 27 days).")
                    conn.rollback()
                outbound_deadline = now.replace(day=3, hour=17, minute=30, second=0, microsecond=0)
                if now > outbound_deadline:
                    last_month_date = now - timedelta(days=5)
                    target_month = last_month_date.month
                    target_year = last_month_date.year
                    print(f"Outbound deadline passed. Submitting records for {target_year}-{target_month} to 'Pending PU'...")
                    sql_outbound = "UPDATE commissions c SET status = 'Pending PU' FROM sales_users su WHERE c.sale_key = su.sale_key AND c.status = 'Original' AND su.sale_type = 'Outbound' AND c.commission_month = %s AND c.commission_year = %s;"
                    cursor.execute(sql_outbound, (target_month, target_year))
                    if cursor.rowcount > 0:
                        conn.commit()
                        print(f"Auto-submitted {cursor.rowcount} Outbound records to Pending PU.")
                    else:
                        conn.rollback()
        except Exception as e:
            print(f"Error during auto-submission: {e}")
            if conn: conn.rollback()
        finally:
            self.release_connection(conn)

    def _create_initial_db_tables(self):
        conn = None
        try:
            conn = self.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("CREATE TABLE IF NOT EXISTS sales_users (id SERIAL PRIMARY KEY, sale_key TEXT UNIQUE NOT NULL, sale_name TEXT NOT NULL, password_hash TEXT, role TEXT DEFAULT 'Sale', sales_target REAL DEFAULT 0, status TEXT DEFAULT 'Active', sale_type TEXT)")
                cursor.execute("CREATE TABLE IF NOT EXISTS customers (id SERIAL PRIMARY KEY, customer_code TEXT UNIQUE NOT NULL, customer_name TEXT NOT NULL, credit_term TEXT)")
                cursor.execute("CREATE TABLE IF NOT EXISTS commissions (id SERIAL PRIMARY KEY, bill_date TEXT, customer_id TEXT, customer_name TEXT, so_number TEXT, sales_service_amount REAL, payment_date TEXT, shipping_cost REAL, delivery_date TEXT, total_payment_amount REAL, vat_deduction REAL, no_vat_deduction REAL, brokerage_fee REAL, giveaways REAL, coupons REAL, transfer_fee REAL, credit_card_fee REAL, wht_3_percent REAL, product_vat_7 REAL, shipping_vat_7 REAL, difference_amount REAL, sale_key TEXT, timestamp TEXT, status TEXT, is_active INTEGER, original_id INTEGER, payment_before_vat REAL DEFAULT 0, payment_no_vat REAL DEFAULT 0, separate_shipping_charge REAL DEFAULT 0, customer_type TEXT, credit_term TEXT, commission_month INTEGER, commission_year INTEGER, rejection_reason TEXT, claim_timestamp TIMESTAMP)")
                cursor.execute("CREATE TABLE IF NOT EXISTS audit_log (id SERIAL PRIMARY KEY, timestamp TEXT, action TEXT, table_name TEXT, record_id INTEGER, user_info TEXT, old_value TEXT, new_value TEXT, changes TEXT)")
                cursor.execute("CREATE TABLE IF NOT EXISTS suppliers (id SERIAL PRIMARY KEY, supplier_code TEXT UNIQUE NOT NULL, supplier_name TEXT NOT NULL, credit_term TEXT)")
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS purchase_orders (
                        id SERIAL PRIMARY KEY, so_number TEXT, po_number TEXT, rr_number TEXT,
                        supplier_name TEXT, user_key TEXT, timestamp TEXT, status TEXT DEFAULT 'Draft', 
                        rejection_reason TEXT, last_modified_by TEXT, credit_term TEXT, 
                        total_cost REAL, total_weight REAL, wht_3_percent_checked BOOLEAN, 
                        wht_3_percent_amount REAL, vat_7_percent_checked BOOLEAN, vat_7_percent_amount REAL, 
                        grand_total REAL, form_data_json TEXT, approval_status TEXT DEFAULT 'Draft',
                        approver_manager1_key TEXT, approval_date_manager1 TIMESTAMP,
                        approver_manager2_key TEXT, approval_date_manager2 TIMESTAMP,
                        approver_director_key TEXT, approval_date_director TIMESTAMP
                    )
                """)
                cursor.execute("CREATE TABLE IF NOT EXISTS purchase_order_items (id SERIAL PRIMARY KEY, purchase_order_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE, product_name TEXT, status TEXT, quantity REAL, weight_per_unit REAL, unit_price REAL, total_weight REAL, total_price REAL)")
                cursor.execute("CREATE TABLE IF NOT EXISTS purchase_order_payments (id SERIAL PRIMARY KEY, purchase_order_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE, payment_type TEXT, amount REAL, payment_date DATE)")
                cursor.execute("CREATE TABLE IF NOT EXISTS notifications (id SERIAL PRIMARY KEY, user_key_to_notify TEXT NOT NULL, message TEXT NOT NULL, related_po_id INTEGER, is_read BOOLEAN DEFAULT FALSE, timestamp TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP)")
            conn.commit()
        except Exception as e:
            messagebox.showerror("Database Setup Error", f"ไม่สามารถสร้างตารางเริ่มต้นได้: {e}")
            if conn: conn.rollback()
        finally:
            self.release_connection(conn)

    def show_screen(self, screen_class, **kwargs):
        for widget in self.winfo_children():
            widget.destroy()
        loading_win = LoadingWindow(self)
        self.update_idletasks()
        self.current_user_key = kwargs.get('user_key')
        if self.notification_poll_id:
           self.after_cancel(self.notification_poll_id)
        self.after(5000, self._check_for_notifications)
        def create_new_screen_and_close_loading():
            from hr_screen import HRScreen
            screen = screen_class(self, **kwargs)
            screen.pack(fill="both", expand=True)
            if isinstance(screen, HRScreen):
             self.hr_screen = screen
            else:
             self.hr_screen = None
            loading_win.stop_animation()
            loading_win.destroy()
        self.after(200, create_new_screen_and_close_loading)

    def show_login_screen(self): 
        from login_screen import LoginScreen
        self.show_screen(LoginScreen, app_container=self)

    def show_main_app(self, sale_key, sale_name): 
        from commission_app import CommissionApp
        self.show_screen(CommissionApp, sale_key=sale_key, sale_name=sale_name, app_container=self, show_logout_button=True)

    def show_hr_screen(self, user_key, user_name):
        from hr_screen import HRScreen
        self.show_screen(HRScreen, app_container=self, user_key=user_key, user_name=user_name)

    def show_purchasing_screen(self, user_key, user_name): 
        from purchasing_screen import PurchasingScreen
        self.show_screen(PurchasingScreen, user_key=user_key, user_name=user_name)

    def show_purchasing_manager_screen(self, user_key, user_name, user_role):
        from purchasing_manager_screen import PurchasingManagerScreen
        self.show_screen(PurchasingManagerScreen, app_container=self, user_key=user_key, user_name=user_name, user_role=user_role)

    def show_director_screen(self, user_key, user_name, user_role):
        from director_screen import DirectorScreen
        self.show_screen(DirectorScreen, app_container=self, user_key=user_key, user_name=user_name, user_role=user_role)

    def show_sales_manager_screen(self, user_key, user_name, user_role):
        from sales_manager_screen import SalesManagerScreen
        self.show_screen(SalesManagerScreen, app_container=self, user_key=user_key, user_name=user_name, user_role=user_role)

    def show_history_window(self, sale_key_filter=None, edit_callback=None):
        from history_windows import CommissionHistoryWindow, PurchaseHistoryWindow
        if sale_key_filter: 
            win = CommissionHistoryWindow(master=self, app_container=self, sale_key_filter=sale_key_filter, on_row_double_click=edit_callback)
            return win
        else: 
            win = PurchaseHistoryWindow(master=self, app_container=self)
            return win

    def show_purchase_detail_window(self, purchase_id, approve_callback=None, reject_callback=None): # <<< แก้ไข: เพิ่ม callbacks
     from history_windows import PurchaseDetailWindow
     PurchaseDetailWindow(
        master=self, 
        app_container=self, 
        purchase_id=purchase_id,
        approve_callback=approve_callback, # <<< เพิ่ม: ส่ง callback ไปยังหน้าต่าง
        reject_callback=reject_callback   # <<< เพิ่ม: ส่ง callback ไปยังหน้าต่าง
    )

    def show_edit_commission_window(self, data, refresh_callback, user_role=None):
        from edit_commission_window import EditCommissionWindow
        EditCommissionWindow(parent=self, app_container=self, data=data, refresh_callback=refresh_callback, user_role=user_role)

    def show_hr_verification_window(self, system_data, excel_data, po_data, refresh_callback=None):
        from hr_windows import HRVerificationWindow
        # สร้างหน้าต่างใหม่โดยมี master เป็น AppContainer (self)
        win = HRVerificationWindow(
            master=self, 
            app_container=self,
            system_data=system_data,
            excel_data=excel_data,
            po_data=po_data,
            refresh_callback=refresh_callback
        )

    def show_sales_data_viewer(self, so_number):
        from hr_windows import SalesDataViewerWindow
        SalesDataViewerWindow(master=self, app_container=self, so_number=so_number)
    
    # --- START: เพิ่มฟังก์ชันใหม่ 2 ฟังก์ชันนี้เข้ามาในคลาส ---
    def open_so_print_dialog(self):
        """เปิดหน้าต่างสำหรับเลือก SO เพื่อพิมพ์"""
        dialog = SOSelectionDialog(
            master=self,
            pg_engine=self.pg_engine,
            print_callback=self.generate_multi_po_document_for_so
        )

    def generate_multi_po_document_for_so(self, so_number):
        """รวบรวมข้อมูลและสั่งสร้าง PDF สำหรับทุก PO ใน SO ที่เลือก"""
        try:
            # 1. ดึงข้อมูล SO Header (สำหรับคอลัมน์ซ้ายของ PDF)
            so_header_df = pd.read_sql(
                "SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1",
                self.pg_engine, params=(so_number,)
            )
            if so_header_df.empty:
                messagebox.showerror("ผิดพลาด", f"ไม่พบข้อมูล SO Header สำหรับ {so_number}")
                return
            so_header_data = so_header_df.iloc[0].to_dict()

            # 2. ดึง ID ของ PO ทั้งหมดที่อนุมัติแล้วใน SO นี้
            po_ids_df = pd.read_sql(
                "SELECT id FROM purchase_orders WHERE so_number = %s AND status = 'Approved' ORDER BY id",
                self.pg_engine, params=(so_number,)
            )
            if po_ids_df.empty:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบใบสั่งซื้อ (PO) ที่อนุมัติแล้วสำหรับ SO: {so_number}")
                return
            
            # 3. วนลูปเพื่อดึงข้อมูลของแต่ละ PO (สำหรับคอลัมน์ขวาของ PDF)
            all_po_data_for_pdf = []
            for po_id in po_ids_df['id']:
                po_header_df = pd.read_sql("SELECT * FROM purchase_orders WHERE id = %s", self.pg_engine, params=(int(po_id),))
                po_items_df = pd.read_sql("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s", self.pg_engine, params=(int(po_id),))
                
                all_po_data_for_pdf.append({
                    'header': po_header_df.iloc[0].to_dict(),
                    'items': po_items_df.to_dict('records')
                })
            
            # 4. เรียกใช้ฟังก์ชันสร้าง PDF เวอร์ชันใหม่
            po_document_generator.generate_multi_po_pdf(so_header_data, all_po_data_for_pdf)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดระหว่างเตรียมข้อมูล: {e}")
            traceback.print_exc()
    # --- END: สิ้นสุดฟังก์ชันที่เพิ่มเข้ามา ---

    def generate_single_po_document(self, po_id):
        """
        ฟังก์ชันตัวกลางสำหรับพิมพ์ PO ใบเดียวโดยใช้ระบบสร้าง PDF ตัวใหม่
        """
        try:
            # 1. ค้นหา SO Number จาก PO ID ที่ได้รับมา
            so_number_df = pd.read_sql("SELECT so_number FROM purchase_orders WHERE id = %s", self.pg_engine, params=(int(po_id),))
            if so_number_df.empty:
                messagebox.showerror("ผิดพลาด", f"ไม่พบ SO Number สำหรับ PO ID {po_id}")
                return
            so_number = so_number_df.iloc[0]['so_number']

            # 2. ดึงข้อมูล SO Header
            so_header_df = pd.read_sql("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", self.pg_engine, params=(so_number,))
            if so_header_df.empty:
                messagebox.showerror("ผิดพลาด", f"ไม่พบข้อมูล SO Header สำหรับ {so_number}")
                return
            so_header_data = so_header_df.iloc[0].to_dict()

            # 3. ดึงข้อมูลของ PO ใบนี้โดยเฉพาะ
            po_header_df = pd.read_sql("SELECT * FROM purchase_orders WHERE id = %s", self.pg_engine, params=(int(po_id),))
            po_items_df = pd.read_sql("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s", self.pg_engine, params=(int(po_id),))

            # 4. สร้าง List ที่มีข้อมูลแค่ PO ใบเดียว
            single_po_data_list = [{
                'header': po_header_df.iloc[0].to_dict(),
                'items': po_items_df.to_dict('records')
            }]

            # 5. เรียกใช้ฟังก์ชันสร้าง PDF ตัวใหม่ แต่ส่งข้อมูลไปแค่ชุดเดียว
            po_document_generator.generate_multi_po_pdf(so_header_data, single_po_data_list)

        except Exception as e:
            messagebox.showerror("ผิดพลาดในการพิมพ์ PO", f"เกิดข้อผิดพลาด: {e}")
            traceback.print_exc()
    
    def on_closing(self):
        self.stop_background_threads.set()
        if self.db_pool: self.db_pool.closeall()
        print("Database connection pool closed.")
        if self.pg_engine: self.pg_engine.dispose()
        self.destroy()

if __name__ == "__main__":
    app = AppContainer()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()