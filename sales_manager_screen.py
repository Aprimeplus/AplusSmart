# sales_manager_screen.py (ฉบับแก้ไขตามคำขอ)

import tkinter as tk
from tkinter import ttk
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkEntry,
                           CTkOptionMenu, CTkRadioButton)
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import psycopg2.errors
import psycopg2.extras
import traceback
import utils

# --- นำเข้า Class ที่จำเป็น ---
from history_windows import SOPopupWindow
from custom_widgets import NumericEntry, DateSelector

class SalesManagerScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role
        
        self.label_font = CTkFont(size=14, weight="bold")
        self.entry_font = CTkFont(size=14)
        
        # --- START: เพิ่มตัวแปรและ Theme ที่จำเป็นสำหรับ SOPopupWindow ---
        self.so_popup = None
        self.so_form_widgets = {}
        self._so_create_string_vars() # สร้าง StringVars ที่จำเป็น
        
        # กำหนด Theme (อาจปรับแก้ได้ตาม Theme หลักของแอป)
        self.sale_theme = {
            "primary": "#3B82F6",
            "header": "#1E40AF",
            "bg": "#EFF6FF"
        }
        self.dropdown_style = {
            "fg_color": "white",
            "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }
        # --- END: เพิ่มตัวแปร ---

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._create_header()
        self._create_task_list_view()
        
        self.after(100, self._load_pending_so_tasks)

    # --- START: เพิ่มฟังก์ชันที่คัดลอกมาจาก purchasing_screen.py ---
    def _so_create_string_vars(self):
        """สร้าง StringVars และตัวแปรที่ใช้ร่วมกันสำหรับฟอร์ม SO"""
        now = datetime.now()
        self.so_form_widgets['thai_months'] = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.so_form_widgets['thai_month_map'] = {name: i+1 for i, name in enumerate(self.so_form_widgets['thai_months'])}
        
        self.so_form_widgets['customer_type_var'] = tk.StringVar(value="ลูกค้าเก่า")
        self.so_form_widgets['credit_term_var'] = tk.StringVar(value="เงินสด")
        self.so_form_widgets['commission_month_var'] = tk.StringVar(value=self.so_form_widgets['thai_months'][now.month - 1])
        self.so_form_widgets['commission_year_var'] = tk.StringVar(value=str(now.year + 543))
        self.so_form_widgets['payment1_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_form_widgets['payment2_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_form_widgets['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['payment_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_subtotal_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vat_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_form_widgets['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_product_input_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_service_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_actual_payment_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_verification_result_var'] = tk.StringVar(value="-")

        self.so_form_widgets['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")

    def _open_so_editor_popup(self, so_data):
        """เปิดหน้าต่าง SOPopupWindow สำหรับแก้ไขข้อมูล"""
        if self.so_popup and self.so_popup.winfo_exists():
            self.so_popup.focus()
            return
        # ส่ง self (SalesManagerScreen) เป็น master ไปให้ SOPopupWindow
        self.so_popup = SOPopupWindow(self, sales_data=so_data, so_shared_vars=self.so_form_widgets, sale_theme=self.sale_theme)

    def _save_so_changes_from_popup(self, so_id, so_shared_vars_data, current_popup_widgets_ref):
        """บันทึกการแก้ไขข้อมูล SO จาก Popup (ฟังก์ชันนี้จำเป็นสำหรับ SOPopupWindow)"""
        updated_data = {}
        # Mapping ระหว่าง key ของ widget กับชื่อคอลัมน์ใน DB
        db_key_map = {
            'bill_date_selector': 'bill_date', 'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id',
            'credit_term_entry': 'credit_term', 'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost', 'delivery_date_selector': 'delivery_date',
            'credit_card_fee_entry': 'credit_card_fee', 'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons', 'giveaway_value_entry': 'giveaways',
            'payment_date_selector': 'payment_date', 'cash_product_input_entry': 'cash_product_input', 'cash_actual_payment_entry': 'cash_actual_payment',
            'sales_service_vat_option': 'sales_service_vat_option', 'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option': 'other_service_fee_vat_option', 'shipping_vat_option_var': 'shipping_vat_option',
            'credit_card_fee_vat_option_var': 'credit_card_fee_vat_option',
            'delivery_type_var': 'delivery_type', 'pickup_location_entry': 'pickup_location', 'relocation_cost_entry': 'relocation_cost',
            'date_to_wh_selector': 'date_to_warehouse', 'date_to_customer_selector': 'date_to_customer', 'pickup_rego_entry': 'pickup_registration'
        }

        def _safe_get_float(entry_widget):
            if entry_widget and hasattr(entry_widget, 'winfo_exists') and entry_widget.winfo_exists():
                try: return utils.convert_to_float(entry_widget.get())
                except (ValueError, tk.TclError): return 0.0
            return 0.0

        for widget_key, db_col_name in db_key_map.items():
            value = None
            if widget_key in current_popup_widgets_ref:
                widget_instance = current_popup_widgets_ref[widget_key]
                if isinstance(widget_instance, NumericEntry): value = _safe_get_float(widget_instance)
                elif isinstance(widget_instance, DateSelector): value = widget_instance.get_date() if widget_instance.winfo_exists() else None
                elif isinstance(widget_instance, (CTkEntry)): value = widget_instance.get().strip() or None if widget_instance.winfo_exists() else None
            elif widget_key in so_shared_vars_data and isinstance(so_shared_vars_data[widget_key], tk.StringVar): value = so_shared_vars_data[widget_key].get()
            if value is not None: updated_data[db_col_name] = value
        
        p1 = _safe_get_float(current_popup_widgets_ref.get('payment1_amount_entry'))
        p2 = _safe_get_float(current_popup_widgets_ref.get('payment2_amount_entry'))
        updated_data['total_payment_amount'] = p1 + p2
        updated_data['so_grand_total'] = utils.convert_to_float(so_shared_vars_data['so_grand_total_var'].get())
        updated_data['difference_amount'] = utils.convert_to_float(so_shared_vars_data['difference_amount_var'].get())
        
        # เพิ่มสถานะว่าถูกแก้ไขโดย SM
        updated_data['status'] = 'Corrected_By_SM'

        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'commissions'")
                db_columns = {row[0] for row in cursor.fetchall()}
            
            final_data_to_save = {k: v for k, v in updated_data.items() if k in db_columns}
            set_clauses = [f'"{k}" = %s' for k in final_data_to_save.keys()]
            params = list(final_data_to_save.values()) + [so_id]
            sql_update = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
            
            with conn.cursor() as cursor: cursor.execute(sql_update, tuple(params))
            conn.commit()
            
            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล SO Number: {updated_data.get('so_number', so_id)} เรียบร้อยแล้ว", parent=self)
            self._load_pending_so_tasks() # โหลดข้อมูลใหม่หลังบันทึกสำเร็จ
            
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล SO จาก Pop-up:\n{e}\n{traceback.format_exc()}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    # --- END: สิ้นสุดส่วนที่คัดลอกมา ---

    def _create_header(self):
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        
        title_text = f"หน้าจอ Sales Manager: {self.user_name}"
        CTkLabel(header_frame, text=title_text, font=CTkFont(size=22, weight="bold")).pack(side="left")
        
        button_container = CTkFrame(header_frame, fg_color="transparent")
        button_container.pack(side="right")
        
        CTkButton(button_container, text="Refresh", command=self._load_pending_so_tasks).pack(side="left", padx=10)
        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").pack(side="left")

    def _create_task_list_view(self):
        self.main_frame = CTkScrollableFrame(self, label_text="รายการ SO ที่รอการตรวจสอบ (จากฝ่ายจัดซื้อ)")
        self.main_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

    def _load_pending_so_tasks(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        try:
            query = """
                SELECT c.*, u.sale_name
                FROM commissions c
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.status IN ('Pending Sale Manager Approval', 'Corrected_By_SM') AND c.is_active = 1
                ORDER BY c.timestamp ASC
            """
            df = pd.read_sql_query(query, self.app_container.pg_engine)
            
            if df.empty:
                CTkLabel(self.main_frame, text="ไม่มีรายการที่รอการอนุมัติ").pack(pady=20)
                return

            for _, row in df.iterrows():
                self._create_so_card(self.main_frame, row)

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูลได้: {e}", parent=self)

    def _create_so_card(self, parent, so_data):
        card_color = "#FEFCE8" if so_data.get('status') == 'Corrected_By_SM' else "#FFFFFF"
        card = CTkFrame(parent, border_width=1, fg_color=card_color)
        card.pack(fill="x", padx=10, pady=5)
        
        so_id = so_data['id']
        
        info_frame = CTkFrame(card, fg_color="transparent")
        info_frame.pack(fill="x", padx=15, pady=10)
        
        sales_amount_val = so_data.get('sales_service_amount', 0.0)
        sales_amount_str = f"{sales_amount_val:,.2f}" if isinstance(sales_amount_val, (int, float)) else "N/A"
        
        so_info = f"SO: {so_data['so_number']}  |  ลูกค้า: {so_data['customer_name']}  |  ยอดขาย: {sales_amount_str} บาท"
        CTkLabel(info_frame, text=so_info, font=CTkFont(size=16)).pack(anchor="w")

        status_text = f"สถานะ: {so_data.get('status')}"
        details = f"ส่งโดย: {so_data['sale_name']}  |  วันที่ส่ง: {pd.to_datetime(so_data['timestamp']).strftime('%Y-%m-%d %H:%M')} | {status_text}"
        CTkLabel(info_frame, text=details, font=CTkFont(size=12), text_color="gray").pack(anchor="w")

        action_frame = CTkFrame(card, fg_color="transparent")
        action_frame.pack(fill="x", padx=15, pady=(0, 10))

        CTkButton(action_frame, text="ปฏิเสธ", command=lambda: self._reject_so(so_id), fg_color="#DC2626", hover_color="#B91C1C").pack(side="right", padx=5)
        CTkButton(action_frame, text="อนุมัติส่งให้ HR", command=lambda: self._approve_so(so_id), fg_color="#16A34A", hover_color="#15803D").pack(side="right", padx=5)
        
        # --- START: แก้ไขปุ่ม "ดูรายละเอียด SO" และลบปุ่ม "ตรวจสอบ SO" ---
        CTkButton(action_frame, text="ดู/แก้ไขรายละเอียด SO", 
                  command=lambda data=so_data.to_dict(): self._open_so_editor_popup(data)).pack(side="right", padx=5)
        # --- END: แก้ไขปุ่ม ---

    def _approve_so(self, so_id):
        if not messagebox.askyesno("ยืนยัน", "คุณต้องการอนุมัติ SO รายการนี้เพื่อส่งต่อไปยังฝ่าย HR ใช่หรือไม่?", parent=self):
            return
        
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'Forwarded_To_HR', approver_sale_manager_key = %s, approval_date_sale_manager = %s
                    WHERE id = %s
                """, (self.user_key, datetime.now(), so_id))
                
                cursor.execute("SELECT so_number, sale_key FROM commissions WHERE id = %s", (so_id,))
                so_info = cursor.fetchone()
                
                cursor.execute("SELECT sale_name FROM sales_users WHERE sale_key = %s", (so_info['sale_key'],))
                sale_info = cursor.fetchone()

                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'HR' AND status = 'Active'")
                hr_keys = [row['sale_key'] for row in cursor.fetchall()]
                
                message = f"SO: {so_info['so_number']} (จากฝ่ายขาย: {sale_info['sale_name']}) ได้รับการอนุมัติแล้ว รอการตรวจสอบจาก HR"
                for hr_key in hr_keys:
                    cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)",
                        (hr_key, message, so_id)
                    )

            conn.commit()
            messagebox.showinfo("สำเร็จ", "อนุมัติ SO และส่งต่อไปยังฝ่าย HR เรียบร้อยแล้ว", parent=self)
            self._load_pending_so_tasks()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _reject_so(self, so_id):
        dialog = CTkInputDialog(text="กรุณาระบุเหตุผลที่ปฏิเสธ (SO จะถูกส่งกลับไปให้ฝ่ายขายแก้ไข):", title="ปฏิเสธ SO")
        reason = dialog.get_input()

        if reason is None or not reason.strip():
             return

        conn = None
        try:
             conn = self.app_container.get_connection()
             with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("UPDATE commissions SET status = 'Rejected by SM', rejection_reason = %s WHERE id = %s", (reason.strip(), so_id,))
                
                cursor.execute("SELECT so_number, sale_key FROM commissions WHERE id = %s", (so_id,))
                so_info = cursor.fetchone()
            
                message = f"SO ของคุณ ({so_info['so_number']}) ถูกปฏิเสธโดย Sales Manager\nเหตุผล: {reason}\nกรุณาแก้ไขและนำส่งข้อมูลอีกครั้ง"
                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)",
                    (so_info['sale_key'], message, so_id)
                )
 
             conn.commit()
             messagebox.showinfo("สำเร็จ", "ปฏิเสธ SO และส่งกลับไปให้ฝ่ายขายเรียบร้อยแล้ว", parent=self)
             self._load_pending_so_tasks()
        except Exception as e:
          if conn: conn.rollback()
          messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
          traceback.print_exc()
        finally:
          if conn: self.app_container.release_connection(conn)