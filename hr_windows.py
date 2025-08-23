import tkinter as tk
from tkinter import ttk, filedialog
from customtkinter import (CTkToplevel, CTkTextbox, CTkScrollableFrame, CTkLabel, CTkFont, CTkFrame, CTkButton, CTkEntry, CTkRadioButton, CTkOptionMenu, CTkInputDialog)
from tkinter import messagebox
import json
import customtkinter as ctk 
import pandas as pd
from datetime import datetime
import traceback
from custom_widgets import NumericEntry, DateSelector, AutoCompleteEntry
import utils

import psycopg2.errors
import psycopg2.extras
import numpy as np

from sqlalchemy import create_engine
from history_windows import SOPopupWindow


class SOPopupWindow(CTkToplevel):
    def __init__(self, master, sales_data, so_shared_vars, sale_theme):
        super().__init__(master)
        self.master = master
        self.sales_data = sales_data
        self.so_shared_vars = so_shared_vars
        self.sale_theme = sale_theme
        
        self.popup_widgets = {}
        self.trace_ids_for_so_calc = []

        self.title(f"ข้อมูล Sales Order (SO: {sales_data.get('so_number', 'N/A')})")
        self.geometry("700x800")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_frame = CTkFrame(self)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        self.so_data_form_frame = CTkScrollableFrame(main_frame, corner_radius=10, label_text="ข้อมูล Sales Order (แก้ไขได้)")
        self.so_data_form_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self._create_so_data_form_content(self.so_data_form_frame)
        self._so_bind_events()
        self.after(100, lambda: self._populate_so_form(self.sales_data))

        button_frame = CTkFrame(main_frame, fg_color="transparent")
        button_frame.grid(row=1, column=0, pady=(5, 15))

        save_button = CTkButton(button_frame, text="บันทึกและปิด", command=self._save_and_close, fg_color="#16A34A", hover_color="#15803D")
        save_button.pack(side="left", padx=10)

        close_button = CTkButton(button_frame, text="ยกเลิก", command=self._on_popup_close, fg_color="gray")
        close_button.pack(side="left", padx=10)

        self.protocol("WM_DELETE_WINDOW", self._on_popup_close)
        self.transient(master)
        self.grab_set()
    
    def _save_and_close(self):
        self._save_so_changes()
        self._on_popup_close()

    def _on_popup_close(self):
        for var, trace_id in self.trace_ids_for_so_calc:
            try:
                var.trace_vdelete("write", trace_id)
            except tk.TclError as e:
                print(f"คำเตือน: ไม่สามารถยกเลิก trace '{trace_id}' ได้: {e}")
        self.trace_ids_for_so_calc = []
        self.destroy()

    def _create_so_section_frame(self, parent, title):
        frame = CTkFrame(parent, corner_radius=10, border_width=1, border_color=self.sale_theme['primary'])
        frame.pack(fill="x", pady=(10, 5), padx=5)
        frame.grid_columnconfigure(1, weight=1)
        CTkLabel(frame, text=title, font=CTkFont(size=18, weight="bold"), text_color=self.sale_theme["header"]).grid(row=0, column=0, columnspan=3, padx=15, pady=(10, 5), sticky="w")
        return frame
            
    def _add_form_row(self, parent, label_text, widget, key, row_index):
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=(15, 10), pady=4, sticky="w")
        widget.grid(row=row_index, column=1, columnspan=2, padx=(10, 15), pady=4, sticky="ew")
        self.popup_widgets[key] = widget

    def _add_item_row_with_vat(self, parent, label_text, entry_key, radio_key, row_index):
        entry_widget = NumericEntry(parent)
        radio_var = self.so_shared_vars[radio_key]
        CTkLabel(parent, text=label_text, font=CTkFont(size=14)).grid(row=row_index, column=0, padx=15, pady=5, sticky="w")
        entry_widget.grid(row=row_index, column=1, padx=(10, 15), pady=5, sticky="ew")
        radio_frame = CTkFrame(parent, fg_color="transparent")
        radio_frame.grid(row=row_index, column=2, padx=(10, 15), pady=5, sticky="w")
        CTkRadioButton(radio_frame, text="VAT", variable=radio_var, value="VAT").pack(side="left", padx=5)
        CTkRadioButton(radio_frame, text="NO VAT", variable=radio_var, value="NO VAT").pack(side="left", padx=5)
        self.popup_widgets[entry_key] = entry_widget

    def _create_so_data_form_content(self, parent_frame):
        # Section 1: Sales Details
        f1 = self._create_so_section_frame(parent_frame, "รายละเอียดการขาย")
        self._add_form_row(f1, "วันที่เปิด SO:", DateSelector(f1, dropdown_style=self.master.dropdown_style), 'bill_date_selector', 1)
        self._add_form_row(f1, "ชื่อลูกค้า:", CTkEntry(f1), 'customer_name_entry', 2)
        self._add_form_row(f1, "รหัสลูกค้า:", CTkEntry(f1), 'customer_id_entry', 3)
        self._add_form_row(f1, "Credit Term:", CTkEntry(f1), 'credit_term_entry', 4)

        # Section 2: Sales and Services
        f2 = self._create_so_section_frame(parent_frame, "ยอดขายและบริการ")
        self._add_item_row_with_vat(f2, "ยอดขายสินค้า/บริการ:", 'sales_amount_entry', 'sales_service_vat_option', 1)
        self._add_item_row_with_vat(f2, "ค่าบริการตัด/เจาะ:", 'cutting_drilling_fee_entry', 'cutting_drilling_fee_vat_option', 2)
        self._add_item_row_with_vat(f2, "ค่าบริการอื่นๆ:", 'other_service_fee_entry', 'other_service_fee_vat_option', 3)
        
        # Section 3: Shipping Cost
        f3 = self._create_so_section_frame(parent_frame, "ค่าจัดส่ง")
        self._add_item_row_with_vat(f3, "ค่าจัดส่ง:", 'shipping_cost_entry', 'shipping_vat_option_var', 1)
        self._add_form_row(f3, "วันที่จัดส่ง:", DateSelector(f3, dropdown_style=self.master.dropdown_style), 'delivery_date_selector', 2)

        # Section 6: Payment Details (ย้ายมาไว้ตรงนี้เพื่อให้เห็นข้อมูลการชำระเงินครบถ้วน)
        f6 = self._create_so_section_frame(parent_frame, "รายละเอียดการโอนชำระ")
        self._add_form_row(f6, "ยอดโอนชำระ 1:", NumericEntry(f6), 'payment1_amount_entry', 1)
        self._add_form_row(f6, "ยอดโอนชำระ 2:", NumericEntry(f6), 'payment2_amount_entry', 2)
        self._add_form_row(f6, "วันที่ชำระ:", DateSelector(f6, dropdown_style=self.master.dropdown_style), 'payment_date_selector', 3)
        self._add_form_row(f6, "ชำระจริงก่อน VAT:", NumericEntry(f6), 'payment_before_vat_entry', 4)
        self._add_form_row(f6, "ชำระจริง NO VAT:", NumericEntry(f6), 'payment_no_vat_entry', 5)

        # Section 7: SO Summary
        f7 = self._create_so_section_frame(parent_frame, "SO สรุปยอดรวม VAT")
        self._add_form_row(f7, "ยอดรวมที่ต้องชำระ:", CTkLabel(f7, textvariable=self.so_shared_vars['so_grand_total_var']), 'grand_total_display', 1)
        self._add_form_row(f7, "ผลต่าง (โอน vs ยอดรวม):", CTkLabel(f7, textvariable=self.so_shared_vars['difference_amount_var']), 'difference_display', 3)

        # Section 8: Cash Verification
        f8 = self._create_so_section_frame(parent_frame, "ตรวจสอบยอดชำระเงินสด")
        self._add_form_row(f8, "ยอดค่าสินค้าเงินสด:", NumericEntry(f8), 'cash_product_input_entry', 1)
        self._add_form_row(f8, "ยอดที่ต้องชำระเงินสด:", CTkLabel(f8, textvariable=self.so_shared_vars['cash_required_total_var']), 'cash_required_display', 2)
        self._add_form_row(f8, "ยอดชำระจริงเงินสด:", NumericEntry(f8), 'cash_actual_payment_entry', 3)
        self._add_form_row(f8, "ตรวจสอบยอดเงินสด:", CTkLabel(f8, textvariable=self.so_shared_vars['cash_verification_result_var']), 'cash_check_display', 4)
        
        # --- Single Save Button at the very end ---
        save_button = CTkButton(parent_frame, text="บันทึกข้อมูล SO", command=self._save_so_changes, fg_color="#16A34A", hover_color="#15803D", font=CTkFont(size=16, weight="bold"))
        save_button.pack(fill="x", padx=10, pady=20)

    def _so_bind_events(self):
        self.trace_ids_for_so_calc = []
        widgets_to_bind_keys = [
            "sales_amount_entry", "cutting_drilling_fee_entry", "other_service_fee_entry",
            "shipping_cost_entry", "credit_card_fee_entry", "transfer_fee_entry",
            "wht_fee_entry", "coupon_value_entry", "giveaway_value_entry",
            "brokerage_fee_entry", "payment1_amount_entry", "payment2_amount_entry",
            "cash_product_input_entry", "cash_actual_payment_entry"
        ]
        for key in widgets_to_bind_keys:
            if key in self.popup_widgets and isinstance(self.popup_widgets[key], (CTkEntry, NumericEntry)):
                self.popup_widgets[key].bind("<KeyRelease>", self._so_update_final_calculations)
            
        radio_vars_keys = [
            'sales_service_vat_option', 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option', 'shipping_vat_option_var',
            'credit_card_fee_vat_option_var'
        ]
        for key in radio_vars_keys:
            if key in self.so_shared_vars and isinstance(self.so_shared_vars[key], tk.StringVar):
                # ใช้ 'w' แทน 'write'
                trace_id = self.so_shared_vars[key].trace_add("write", self._so_update_final_calculations)
                self.trace_ids_for_so_calc.append((self.so_shared_vars[key], trace_id))

    def _so_update_final_calculations(self, *args):
        if not self.winfo_exists(): return

        w_vars = self.so_shared_vars
        w_widgets = self.popup_widgets

        def get_float_from_entry(entry_key):
            entry_widget = w_widgets.get(entry_key)
            if entry_widget:
                if hasattr(entry_widget, '_entry') and entry_widget._entry.winfo_exists():
                    try: return utils.convert_to_float(entry_widget.get())
                    except (tk.TclError, ValueError): return 0.0
                else: return 0.0
            return 0.0

        sales = get_float_from_entry('sales_amount_entry')
        shipping = get_float_from_entry('shipping_cost_entry')
        card_fee = get_float_from_entry('credit_card_fee_entry')
        cutting_drilling = get_float_from_entry('cutting_drilling_fee_entry')
        other_service = get_float_from_entry('other_service_fee_entry')
        
        total_vatable_subtotal, total_cashable_services_and_fees = 0.0, 0.0
        
        items_to_process = [(sales, w_vars['sales_service_vat_option'].get()), (cutting_drilling, w_vars['cutting_drilling_fee_vat_option'].get()), (other_service, w_vars['other_service_fee_vat_option'].get()), (shipping, w_vars['shipping_vat_option_var'].get()), (card_fee, w_vars['credit_card_fee_vat_option_var'].get())]
        for amount, option in items_to_process:
            if option == "VAT": total_vatable_subtotal += amount
            else: total_cashable_services_and_fees += amount
            
        so_grand_total = total_vatable_subtotal * 1.07
        w_vars['so_grand_total_var'].set(f"{so_grand_total:,.2f}")

        payment1 = get_float_from_entry('payment1_amount_entry')
        payment2 = get_float_from_entry('payment2_amount_entry')
        so_vs_payment_diff = (payment1 + payment2) - so_grand_total
        w_vars['difference_amount_var'].set(f"{so_vs_payment_diff:,.2f}")

        def set_check_result(label_widget_key, var, diff_val, plus_text, minus_text):
            label_widget_ref = w_widgets.get(label_widget_key)
            color_map = {"-": ("gray85", "black"), "ok": ("#BBF7D0", "#15803D"), "bad": ("#FECACA", "#B91C1C")}
            if abs(diff_val) < 0.01: state, text = "ok", "ถูกต้อง"
            elif diff_val > 0: state, text = "ok", f"{plus_text} (+{abs(diff_val):,.2f})"
            else: state, text = "bad", f"{minus_text} (-{abs(diff_val):,.2f})"
            if label_widget_ref and label_widget_ref.winfo_exists():
                var.set(text)
                if isinstance(label_widget_ref, CTkLabel): label_widget_ref.configure(fg_color=color_map[state][0], text_color=color_map[state][1], text=text)
                elif isinstance(label_widget_ref, CTkEntry):
                    current_state = label_widget_ref.cget("state")
                    if current_state == "readonly": label_widget_ref.configure(state="normal")
                    label_widget_ref.delete(0, "end"); label_widget_ref.insert(0, text)
                    label_widget_ref.configure(fg_color=color_map[state][0], text_color=color_map[state][1])
                    if current_state == "readonly": label_widget_ref.configure(state="readonly")

        set_check_result('so_check_display', w_vars.get('so_vs_payment_result_var'), so_vs_payment_diff, "ยอดโอนเกิน", "ยอดโอนขาด")

        cash_product_val = get_float_from_entry('cash_product_input_entry')
        cash_required_total = cash_product_val + total_cashable_services_and_fees
        w_vars['cash_required_total_var'].set(f"{cash_required_total:,.2f}")
        
        actual_cash_payment = get_float_from_entry('cash_actual_payment_entry')
        cash_diff = actual_cash_payment - cash_required_total
        
        set_check_result('cash_check_display', w_vars.get('cash_verification_result_var'), cash_diff, "เงินสดเกิน", "เงินสดขาด")

    def _populate_so_form(self, data):
        if not self.winfo_exists(): return

        def set_val(widget_or_var, value):
            if not widget_or_var: return
            
            if not (hasattr(widget_or_var, 'winfo_exists') and widget_or_var.winfo_exists()) and not isinstance(widget_or_var, tk.StringVar):
                return

            if isinstance(widget_or_var, (CTkEntry, NumericEntry, AutoCompleteEntry)):
                state = widget_or_var.cget("state")
                widget_or_var.configure(state="normal")
                widget_or_var.delete(0, "end")
                if pd.notna(value):
                    widget_or_var.insert(0, f"{value:,.2f}" if isinstance(value, (float, int)) else str(value))
                widget_or_var.configure(state=state)
            elif isinstance(widget_or_var, DateSelector):
                dt = pd.to_datetime(value, errors='coerce')
                widget_or_var.set_date(dt.to_pydatetime() if pd.notna(dt) else None)
            elif isinstance(widget_or_var, tk.StringVar):
                widget_or_var.set(str(value) if pd.notna(value) and value else "")
            elif isinstance(widget_or_var, CTkLabel):
                widget_or_var.configure(text=f"{value:,.2f}" if isinstance(value, (float, int)) else str(value) if value is not None and value != "" else "")
            elif isinstance(widget_or_var, CTkOptionMenu):
                widget_or_var.set(str(value) if pd.notna(value) and value else widget_or_var.cget("values")[0])

        key_map = {
            'bill_date': 'bill_date_selector', 'customer_name': 'customer_name_entry', 'customer_id': 'customer_id_entry',
            'credit_term': 'credit_term_entry', 'sales_service_amount': 'sales_amount_entry', 'cutting_drilling_fee': 'cutting_drilling_fee_entry',
            'other_service_fee': 'other_service_fee_entry', 'shipping_cost': 'shipping_cost_entry', 'delivery_date': 'delivery_date_selector',
            'credit_card_fee': 'credit_card_fee_entry', 'transfer_fee': 'transfer_fee_entry', 'wht_3_percent': 'wht_fee_entry',
            'brokerage_fee': 'brokerage_fee_entry', 'coupons': 'coupon_value_entry', 'giveaways': 'giveaway_value_entry',
            'payment_date': 'payment_date_selector', 'cash_product_input': 'cash_product_input_entry', 'cash_actual_payment': 'cash_actual_payment_entry',
            'sales_service_vat_option': 'sales_service_vat_option', 'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
            'other_service_fee_vat_option': 'other_service_fee_vat_option', 'shipping_vat_option': 'shipping_vat_option_var',
            'credit_card_fee_vat_option': 'credit_card_fee_vat_option_var', 'so_grand_total': 'so_grand_total_var',
            'so_vs_payment_result': 'so_vs_payment_result_var', 'difference_amount': 'difference_amount_var',
            'cash_required_total': 'cash_required_total_var', 'cash_verification_result': 'cash_verification_result_var',
            'delivery_type': 'delivery_type_var', 'pickup_location': 'pickup_location_entry',
            'relocation_cost': 'relocation_cost_entry', 'date_to_warehouse': 'date_to_wh_selector',
            'date_to_customer': 'date_to_customer_selector', 'pickup_registration': 'pickup_rego_entry'
        }
        
        for key, widget in self.popup_widgets.items():
            if isinstance(widget, (CTkEntry, NumericEntry, AutoCompleteEntry)): set_val(widget, "")
            elif isinstance(widget, DateSelector): set_val(widget, None)
            elif isinstance(widget, CTkLabel): widget.configure(text="")
        
        for key, var in self.so_shared_vars.items():
            if isinstance(var, tk.StringVar): var.set("")
        
        if data is not None:
            for db_key, w_key in key_map.items():
                widget_or_var = self.so_shared_vars.get(w_key) or self.popup_widgets.get(w_key)
                if widget_or_var:
                    set_val(widget_or_var, data.get(db_key))
            
            payment1_entry = self.popup_widgets.get('payment1_amount_entry')
            if payment1_entry: set_val(payment1_entry, data.get('total_payment_amount'))
        
        self.update_idletasks()
        self._so_update_final_calculations()

    def _save_so_changes(self):
        if self.sales_data is None: 
            messagebox.showerror("ข้อผิดพลาด", "ไม่มีข้อมูล SO ให้บันทึก", parent=self)
            return
        self.master._save_so_changes_from_popup(self.sales_data.get('id'), self.so_shared_vars, self.popup_widgets)

class SalesDataViewerWindow(CTkToplevel):
    def __init__(self, master, app_container, so_number):
        super().__init__(master)
        self.app_container = app_container
        self.so_number = so_number
        self.so_data = None

        self.sale_theme = self.app_container.THEME.get("sale", {"primary": "#3B82F6", "header": "#1E40AF"})
        self.title(f"รายละเอียดข้อมูล SO: {self.so_number}")
        self.geometry("700x800")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkLabel(header_frame, text=f"ข้อมูลสำหรับ SO Number: {self.so_number}", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(header_frame, text="ปิด", command=self.destroy, width=80).pack(side="right")

        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.load_data()
        self.create_widgets()
        
        self.transient(master)
        self.grab_set()

    def load_data(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1", (self.so_number,))
                self.so_data = cursor.fetchone()
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล SO ได้: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_section_frame(self, parent, title):
        frame = ctk.CTkFrame(parent, corner_radius=10, border_width=1, border_color=self.sale_theme['primary'])
        frame.pack(fill="x", pady=(10, 5), padx=5)
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=18, weight="bold"), text_color=self.sale_theme["header"]).grid(
            row=0, column=0, columnspan=2, padx=15, pady=(10, 5), sticky="w")
        return frame

    def _add_display_row(self, parent, row_index, label_text, value):
        if value is None or pd.isna(value):
            value_text = "-"
        elif isinstance(value, (int, float, np.floating)):
            value_text = f"{value:,.2f}"
        elif isinstance(value, (datetime, pd.Timestamp)):
            value_text = value.strftime('%Y-%m-%d')
        else:
            value_text = str(value)
            
        ctk.CTkLabel(parent, text=label_text, font=ctk.CTkFont(size=14)).grid(
            row=row_index, column=0, padx=(15, 10), pady=4, sticky="w")
        ctk.CTkLabel(parent, text=value_text, font=ctk.CTkFont(size=14), wraplength=400, justify="left").grid(
            row=row_index, column=1, padx=(10, 15), pady=4, sticky="ew")

    def create_widgets(self):
        if not self.so_data:
            ctk.CTkLabel(self.main_frame, text="ไม่พบข้อมูล").pack(pady=20)
            return
        
        header_map = self.app_container.HEADER_MAP

        # --- Section 1: รายละเอียดการขาย ---
        f1 = self._create_section_frame(self.main_frame, "รายละเอียดการขาย")
        self._add_display_row(f1, 1, header_map.get('so_number', 'เลขที่ SO'), self.so_data.get('so_number'))
        self._add_display_row(f1, 2, header_map.get('bill_date', 'วันที่เปิด SO'), self.so_data.get('bill_date'))
        self._add_display_row(f1, 3, header_map.get('customer_id', 'รหัสลูกค้า'), self.so_data.get('customer_id'))
        self._add_display_row(f1, 4, header_map.get('customer_name', 'ชื่อลูกค้า'), self.so_data.get('customer_name'))
        self._add_display_row(f1, 5, header_map.get('credit_term', 'เครดิต'), self.so_data.get('credit_term'))
        self._add_display_row(f1, 6, header_map.get('commission_month', 'เดือนคอมมิชชั่น'), self.so_data.get('commission_month'))
        self._add_display_row(f1, 7, header_map.get('commission_year', 'ปีคอมมิชชั่น'), self.so_data.get('commission_year') + 543 if self.so_data.get('commission_year') else None)

        # --- Section 2: ยอดขายและบริการ ---
        f2 = self._create_section_frame(self.main_frame, "ยอดขายและบริการ")
        self._add_display_row(f2, 1, f"{header_map.get('sales_service_amount', 'ยอดขาย/บริการ')} ({self.so_data.get('sales_service_vat_option')})", self.so_data.get('sales_service_amount'))
        self._add_display_row(f2, 2, f"{header_map.get('cutting_drilling_fee', 'ค่าบริการตัด/เจาะ')} ({self.so_data.get('cutting_drilling_fee_vat_option')})", self.so_data.get('cutting_drilling_fee'))
        self._add_display_row(f2, 3, f"{header_map.get('other_service_fee', 'ค่าบริการอื่นๆ')} ({self.so_data.get('other_service_fee_vat_option')})", self.so_data.get('other_service_fee'))

        # --- Section 3: ค่าจัดส่ง ---
        f3 = self._create_section_frame(self.main_frame, "ค่าจัดส่ง")
        self._add_display_row(f3, 1, f"{header_map.get('shipping_cost', 'ค่าขนส่ง')} ({self.so_data.get('shipping_vat_option')})", self.so_data.get('shipping_cost'))
        self._add_display_row(f3, 2, header_map.get('delivery_date', 'วันที่จัดส่ง'), self.so_data.get('delivery_date'))
        
        # +++ START: เพิ่ม Section ที่ขาดไป +++
        
        # --- Section 4: Delivery Note ---
        f4 = self._create_section_frame(self.main_frame, "Delivery Note")
        self._add_display_row(f4, 1, header_map.get('delivery_type', 'การจัดส่ง'), self.so_data.get('delivery_type'))
        self._add_display_row(f4, 2, header_map.get('pickup_location', 'Location เข้ารับ'), self.so_data.get('pickup_location'))
        self._add_display_row(f4, 3, header_map.get('relocation_cost', 'ค่าย้าย'), self.so_data.get('relocation_cost'))
        self._add_display_row(f4, 4, header_map.get('date_to_warehouse', 'วันที่เข้าคลัง'), self.so_data.get('date_to_warehouse'))
        self._add_display_row(f4, 5, header_map.get('date_to_customer', 'วันที่ส่งลูกค้า'), self.so_data.get('date_to_customer'))
        self._add_display_row(f4, 6, header_map.get('pickup_registration', 'ทะเบียนเข้ารับ'), self.so_data.get('pickup_registration'))

        # --- Section 5: ค่าธรรมเนียมและส่วนลด ---
        f5 = self._create_section_frame(self.main_frame, "ค่าธรรมเนียมและส่วนลด")
        self._add_display_row(f5, 1, f"{header_map.get('credit_card_fee', 'ค่าธรรมเนียมบัตร')} ({self.so_data.get('credit_card_fee_vat_option')})", self.so_data.get('credit_card_fee'))
        self._add_display_row(f5, 2, header_map.get('transfer_fee', 'ค่าธรรมเนียมโอน'), self.so_data.get('transfer_fee'))
        self._add_display_row(f5, 3, header_map.get('wht_3_percent', 'หัก ณ ที่จ่าย 3%'), self.so_data.get('wht_3_percent'))
        self._add_display_row(f5, 4, header_map.get('brokerage_fee', 'ค่านายหน้า'), self.so_data.get('brokerage_fee'))
        self._add_display_row(f5, 5, header_map.get('giveaways', 'ของแถม'), self.so_data.get('giveaways'))
        self._add_display_row(f5, 6, header_map.get('coupons', 'คูปอง'), self.so_data.get('coupons'))

        # --- Section 6: รายละเอียดการชำระเงิน ---
        f6 = self._create_section_frame(self.main_frame, "รายละเอียดการชำระเงิน")
        self._add_display_row(f6, 1, header_map.get('total_payment_amount', 'ยอดชำระรวม'), self.so_data.get('total_payment_amount'))
        self._add_display_row(f6, 2, header_map.get('payment_date', 'วันที่ชำระ'), self.so_data.get('payment_date'))
        self._add_display_row(f6, 3, header_map.get('payment_before_vat', 'ชำระก่อน VAT'), self.so_data.get('payment_before_vat'))
        self._add_display_row(f6, 4, header_map.get('payment_no_vat', 'ชำระ NV'), self.so_data.get('payment_no_vat'))


    def load_and_populate_data(self):
        self._show_loading()
        try:
            query = "SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1"
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.so_number,))
            if df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูลสำหรับ SO Number: {self.so_number}", parent=self)
                self.after(100, self.destroy)
                return
            
            data = df.iloc[0]

            for key, (parent, label, row, col_name) in self.fields.items():
                entry_widget = getattr(self, f"{key}_entry")
                value = data.get(col_name)
                display_text = ""
                if pd.notna(value):
                    if isinstance(value, (int, float)):
                        display_text = f"{value:,.2f}"
                    else:
                        display_text = str(value)

                entry_widget.configure(state="normal")
                entry_widget.insert(0, display_text)
                entry_widget.configure(state="readonly")
        
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล SO ได้: {e}", parent=self)
            self.after(100, self.destroy)
        finally:
            self._hide_loading()
# +++ END: เพิ่ม Class ใหม่ +++

# hr_windows.py (คลาส HRVerificationWindow ที่แก้ไขสมบูรณ์)

class HRVerificationWindow(CTkToplevel):

    def __init__(self, master, app_container, system_data, excel_data, po_data, refresh_callback=None):
        super().__init__(master)
        self.master = master
        self.app_container = app_container
        self.system_data = system_data
        self.excel_data = excel_data
        self.po_data = po_data
        self.refresh_callback = refresh_callback
        self.so_number = self.system_data.get('so_number', 'N/A')
        self.record_id = self.system_data.get('id')
        self.cost_multiplier_var = tk.StringVar(value="1.03") # สร้างตัวแปรพร้อมค่าเริ่มต้น

        # --- ตัวแปรสำหรับเก็บค่าที่ User Override และค่าที่คำนวณแล้ว ---
        self.cost_overrides = {} # เก็บค่า cost ที่ HR อาจแก้ไขเอง
        self.calculated_values = {} # เก็บผลรวมที่คำนวณล่าสุด
        self.final_sale_source = tk.StringVar(value="system")
        self.final_cost_source = tk.StringVar(value="system")
        self.final_sale_source.trace_add("write", self._update_selection_display)
        self.final_cost_source.trace_add("write", self._update_selection_display)

        self._so_create_string_vars() # สร้าง StringVars สำหรับหน้าต่างแก้ไข SO

        self.title(f"สรุปข้อมูล SO: {self.so_number}")
        self.geometry("950x750")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- UI Layout ใหม่ทั้งหมด ---
        self._create_new_ui_layout()

        initial_multiplier = self.system_data.get('cost_multiplier')
        if initial_multiplier and f"{initial_multiplier:.2f}" in ["1.01", "1.02", "1.03", "1.04", "1.05"]:
            self.cost_multiplier_var.set(f"{initial_multiplier:.2f}")
        else:
            self.cost_multiplier_var.set("1.03") # ถ้าไม่มีข้อมูล ให้ใช้ค่าเริ่มต้น

        # --- โหลดและคำนวณข้อมูล ---
        self.after(50, self._update_all_calculations_and_ui)
        
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.transient(master)
        self.grab_set()

    def _create_new_ui_layout(self):
        """สร้าง UI Layout ใหม่ทั้งหมดสำหรับหน้าต่างนี้"""
        # --- Header ---
        header_frame = CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 10))
        CTkLabel(header_frame, text=f"SO Number: {self.so_number}", font=CTkFont(size=20, weight="bold")).pack(side="left")
        
        detail_button_frame = CTkFrame(header_frame, fg_color="transparent")
        detail_button_frame.pack(side="right")
        CTkButton(detail_button_frame, text="ดูข้อมูล SO", command=self._view_so_data).pack(side="left", padx=(0, 5))
        CTkButton(detail_button_frame, text="✏️ แก้ไขข้อมูล SO", command=self._open_so_editor_popup).pack(side="left", padx=(5, 0))

        # --- Main Scrollable Frame ---
        scroll_frame = CTkScrollableFrame(self, fg_color="#F0F2F5") # สีพื้นหลังอ่อนๆ
        scroll_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        scroll_frame.grid_columnconfigure(0, weight=1)
        scroll_frame.grid_columnconfigure(1, weight=1)

        # ### UI ใหม่สำหรับแสดงผลสรุป ###
        sales_card = CTkFrame(scroll_frame, corner_radius=10)
        sales_card.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self._create_summary_card(sales_card, "ยอดขายรวมสุดท้าย (Final Sales)", "sales")

        cost_card = CTkFrame(scroll_frame, corner_radius=10)
        cost_card.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self._create_summary_card(cost_card, "ยอดต้นทุนรวมสุดท้าย (Final Cost)", "cost")
        
        # ### UI ใหม่สำหรับแสดงรายการ PO ###
        self.po_container_frame = CTkFrame(scroll_frame, fg_color="transparent")
        self.po_container_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
        self.po_container_frame.grid_columnconfigure(0, weight=1)
        CTkLabel(self.po_container_frame, text="ใบสั่งซื้อ (PO) ที่เกี่ยวข้อง", font=CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 5))

        # --- Action Buttons Frame (ด้านล่างสุด) ---
        action_frame = CTkFrame(self, fg_color="transparent")
        action_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(10, 15))
        action_frame.grid_columnconfigure((0,1,2), weight=1)

        CTkButton(action_frame, text="ตีกลับให้ฝ่ายจัดซื้อ (Reject)", height=40, fg_color="#F97316", hover_color="#EA580C", command=self._reject_to_purchasing).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        CTkButton(action_frame, text="บันทึกการแก้ไข", height=40, fg_color="#3B82F6", hover_color="#2563EB", command=self._save_intermediate_changes).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        CTkButton(action_frame, text="ยืนยันข้อมูลถูกต้อง (Verify Data)", height=40, fg_color="#16A34A", hover_color="#15803D", command=self._verify_and_save_data).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
    
    def _populate_po_cards(self):
        """สร้างการ์ดสำหรับ PO แต่ละใบ"""
        for widget in self.po_container_frame.winfo_children():
            # ไม่ลบ Label หัวข้อ
            if isinstance(widget, CTkLabel) and "ใบสั่งซื้อ" in widget.cget("text"):
                continue
            widget.destroy()

        if self.po_data.empty:
            CTkLabel(self.po_container_frame, text="ไม่พบข้อมูล PO ที่เกี่ยวข้อง").pack(pady=10)
            return

        for index, row in self.po_data.iterrows():
            po_card = CTkFrame(self.po_container_frame, border_width=1, corner_radius=8)
            po_card.pack(fill="x", expand=True, padx=0, pady=4)
            po_card.grid_columnconfigure(0, weight=1)

            info_text = f"PO: {row['po_number']}  |  Supplier: {row['supplier_name']}  |  ยอดรวม: {row.get('total_cost', 0):,.2f} บาท"
            CTkLabel(po_card, text=info_text).grid(row=0, column=0, sticky="w", padx=10, pady=5)
            
            status_color = "#16A34A" if row['status'] == 'Approved' else 'gray'
            CTkLabel(po_card, text=f"สถานะ: {row['status']}", text_color=status_color).grid(row=1, column=0, sticky="w", padx=10, pady=(0,5))

            detail_button = CTkButton(
                po_card, text="ดูรายละเอียด", width=120,
                command=lambda po_id=row['id']: self.app_container.show_purchase_detail_window(int(po_id))
            )
            detail_button.grid(row=0, rowspan=2, column=1, padx=10, pady=5)

    def _create_summary_card(self, parent, title, card_type):
        """Helper function สำหรับสร้างการ์ดสรุป"""
        parent.grid_columnconfigure(0, weight=1)
        CTkLabel(parent, text=title, font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(10, 0), sticky="w")
        
        value_label = CTkLabel(parent, text="0.00", font=CTkFont(size=32, weight="bold"))
        value_label.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="w")

        source_label = CTkLabel(parent, text="Source: System", font=CTkFont(size=12), text_color="gray50")
        source_label.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="w")

        if card_type == "sales":
            self.final_sales_label = value_label
            self.final_sales_source_label = source_label
        else: # cost
            self.final_cost_label = value_label
            self.final_cost_source_label = source_label 

            # --- START: เพิ่มโค้ดส่วนนี้เข้าไป ---
            CTkLabel(parent, text="ตัวคูณต้นทุน (Cost Multiplier):", font=CTkFont(size=12)).grid(row=3, column=0, padx=(15, 5), pady=(10, 0), sticky="w")
            multiplier_options = ["1.01", "1.02", "1.03", "1.04", "1.05"]

            # self.cost_multiplier_var ถูกสร้างใน __init__ แล้ว
            self.cost_multiplier_menu = CTkOptionMenu(parent, variable=self.cost_multiplier_var, values=multiplier_options)
            self.cost_multiplier_menu.grid(row=4, column=0, padx=15, pady=(0, 10), sticky="w")

    def _so_create_string_vars(self):
        """สร้าง StringVars ที่จำเป็นสำหรับ SOPopupWindow"""
        self.so_form_widgets = {}
        self.sale_theme = self.app_container.THEME["sale"]
        self.dropdown_style = {
            "fg_color": "white", "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }
        self.so_form_widgets['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        self.so_form_widgets['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_form_widgets['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_verification_result_var'] = tk.StringVar(value="-")

    def _view_so_data(self):
        if self.system_data and self.system_data.get('so_number'):
            self.app_container.show_sales_data_viewer(self.system_data['so_number'])
        else:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบ SO Number สำหรับแสดงข้อมูล", parent=self)

    def _view_po_data(self):
        if not self.po_data.empty:
            # ดึงค่า ID จากแถวแรก
            po_id_value = self.po_data.iloc[0].get('id')

            # ตรวจสอบว่ามีค่า ID จริงๆ
            if po_id_value is not None:
                # <<< START: แก้ไขจุดนี้ >>>
                # แปลงค่า numpy.int64 ให้เป็น int ปกติก่อนส่งไปใช้งาน
                first_po_id = int(po_id_value)
                self.app_container.show_purchase_detail_window(first_po_id)
                # <<< END: สิ้นสุดการแก้ไข >>>
            else:
                messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบ ID ของ PO ในข้อมูลที่แสดง", parent=self)
        else:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูล PO ที่เกี่ยวข้อง", parent=self)

    def _edit_data(self):
        if self.system_data and self.system_data.get('id_db'):
            record_id_to_load = self.system_data.get('id_db')
            # ต้องสร้าง DataFrame ชั่วคราวเพื่อให้เข้ากับ Format ที่ EditCommissionWindow คาดหวัง
            data_for_edit = pd.DataFrame([self.system_data])
            row_to_edit = data_for_edit.iloc[0]
            
            # ต้องมี refresh_callback ที่ส่งมาจาก hr_screen
            self.app_container.show_edit_commission_window(row_to_edit, self.refresh_callback, user_role="HR")
            self.destroy() # ปิดหน้าต่างปัจจุบันหลังเปิดหน้าแก้ไข
        else:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่สามารถหาข้อมูลสำหรับแก้ไขได้", parent=self)

    def _create_revenue_table(self):
        revenue_keys = ['sales_service_amount', 'shipping_cost', 'cutting_drilling_fee', 'other_service_fee', 'credit_card_fee']
        revenue_headers = self.app_container.HEADER_MAP
        
        # Header
        CTkLabel(self.revenue_table_frame, text="หัวข้อ", font=CTkFont(weight="bold")).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        CTkLabel(self.revenue_table_frame, text="ข้อมูลในระบบ", font=CTkFont(weight="bold")).grid(row=1, column=1, padx=5, pady=2, sticky="e")
        CTkLabel(self.revenue_table_frame, text="ข้อมูลจาก Express", font=CTkFont(weight="bold")).grid(row=1, column=2, padx=5, pady=2, sticky="e")

        # Data Rows
        for i, key in enumerate(revenue_keys):
            header = revenue_headers.get(key, key)
            system_val = self.system_data.get(key, 0)
            excel_val = self.excel_data.get(key, 'N/A')
            
            CTkLabel(self.revenue_table_frame, text=header).grid(row=i+2, column=0, padx=5, pady=2, sticky="w")
            CTkLabel(self.revenue_table_frame, text=f"{system_val:,.2f}").grid(row=i+2, column=1, padx=5, pady=2, sticky="e")
            CTkLabel(self.revenue_table_frame, text=f"{excel_val if isinstance(excel_val, str) else f'{excel_val:,.2f}'}").grid(row=i+2, column=2, padx=5, pady=2, sticky="e")

    def _create_cost_table(self):
        cost_keys = ['final_cost_amount', 'giveaways', 'brokerage_fee', 'wht_3_percent', 'transfer_fee']
        cost_headers = self.app_container.HEADER_MAP

        # Header
        CTkLabel(self.cost_table_frame, text="หัวข้อ", font=CTkFont(weight="bold")).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        CTkLabel(self.cost_table_frame, text="ข้อมูลในระบบ", font=CTkFont(weight="bold")).grid(row=1, column=1, padx=5, pady=2, sticky="e")
        CTkLabel(self.cost_table_frame, text="ข้อมูลจาก Express", font=CTkFont(weight="bold")).grid(row=1, column=2, padx=5, pady=2, sticky="e")

        # Data Rows
        for i, key in enumerate(cost_keys):
            header = cost_headers.get(key, key)
            system_val = self.system_data.get(key) # ดึงค่ามาก่อน
            excel_val = self.excel_data.get(key, 0)

            # --- START: เพิ่มการตรวจสอบค่า None ---
            system_text = f"{system_val:,.2f}" if system_val is not None else "0.00"
            # --- END: สิ้นสุดการแก้ไข ---
            
            CTkLabel(self.cost_table_frame, text=header).grid(row=i+2, column=0, padx=5, pady=2, sticky="w")
            CTkLabel(self.cost_table_frame, text=system_text).grid(row=i+2, column=1, padx=5, pady=2, sticky="e")
            
            val_for_comparison = system_val if system_val is not None else 0
            color = "red" if val_for_comparison < excel_val else "green"
            CTkLabel(self.cost_table_frame, text=f"{excel_val:,.2f}", text_color=color).grid(row=i+2, column=2, padx=5, pady=2, sticky="e")
    
    def _create_section_frame(self, parent, title, col):
        frame = CTkFrame(parent, corner_radius=10, border_width=1)
        frame.grid(row=0, column=col, sticky="nsew", padx=5, pady=5)
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        label = CTkLabel(frame, text=title, font=CTkFont(size=14, weight="bold"))
        label.grid(row=0, column=0, padx=10, pady=(5, 10), sticky="w")
        
        return frame
    
    def _populate_data(self):
        """
        คำนวณและแสดงผลข้อมูลสรุป (ยอดขาย, ต้นทุน, และข้อมูล PO)
        ใน Label ที่เกี่ยวข้อง (เวอร์ชันแก้ไข)
        """
        # --- 1. คำนวณยอดขายรวมสุดท้าย (Final Sales) ---
        revenue_cols = [
            'sales_service_amount', 'shipping_cost', 'cutting_drilling_fee', 
            'other_service_fee', 'credit_card_fee'
        ]
        # ใช้ค่าจาก system_data ที่ได้รับมาในการคำนวณ
        final_sales = sum(float(self.system_data.get(col, 0) or 0) for col in revenue_cols)

        # --- 2. คำนวณยอดต้นทุนรวมสุดท้าย (Final Cost) ---
        # ดึงค่าจาก PO ที่เกี่ยวข้อง
        po_product_cost = self.po_data['total_cost'].sum()
        po_shipping_cost = self.po_data['shipping_to_stock_cost'].sum() + self.po_data['shipping_to_site_cost'].sum()

        # ดึงค่าใช้จ่ายอื่นๆ จาก SO
        brokerage_cost = float(self.system_data.get('brokerage_fee', 0) or 0)
        transfer_cost = float(self.system_data.get('transfer_fee', 0) or 0)
        giveaways_cost = float(self.system_data.get('giveaways', 0) or 0)
        
        # รวมเป็นต้นทุนสุดท้าย
        final_cost = po_product_cost + brokerage_cost + transfer_cost + giveaways_cost

        # --- 3. อัปเดต Label ที่แสดงผล ---
        self.final_sales_label.configure(text=f"{final_sales:,.2f} บาท")
        self.final_cost_label.configure(text=f"{final_cost:,.2f} บาท")

        # --- 4. แสดงข้อมูล PO ที่เกี่ยวข้อง ---
        for widget in self.po_container_frame.winfo_children():
            widget.destroy()

        if self.po_data.empty:
            CTkLabel(self.po_container_frame, text="ไม่พบข้อมูล PO ที่เกี่ยวข้อง").pack(pady=10)
        else:
            for index, row in self.po_data.iterrows():
                # สร้าง Frame สำหรับ PO แต่ละใบ
                po_card = CTkFrame(self.po_container_frame, fg_color="transparent")
                po_card.pack(fill="x", expand=True, padx=5, pady=2)
                po_card.grid_columnconfigure(0, weight=1)

                # สร้าง Label แสดงข้อมูล
                info_text = f"PO: {row['po_number']}, สถานะ: {row['status']}, ยอดรวม: {row.get('total_cost', 0):,.2f} บาท"
                CTkLabel(po_card, text=info_text).grid(row=0, column=0, sticky="w")

                # สร้างปุ่ม "ดูรายละเอียด" สำหรับ PO ใบนี้โดยเฉพาะ
                detail_button = CTkButton(
                    po_card, 
                    text="ดูรายละเอียด", 
                    width=100,
                    # ใช้ lambda เพื่อส่ง po_id ที่ถูกต้องของแถวนี้ไป
                    command=lambda po_id=row['id']: self.app_container.show_purchase_detail_window(int(po_id))
                )
                detail_button.grid(row=0, column=1, padx=5)

    
    
    def _on_verify(self):
        conn = None
        try:
            # --- Determine Final Values ---
            if self.sales_choice_var.get() == 'system':
                # <<< จุดแก้ไข: เปลี่ยนจากการดึงค่าบนหน้าจอ มาเป็นค่าจากข้อมูลดิบ 'system_data'
                final_sale = utils.convert_to_float(self.system_data.get('sales_service_amount', 0))
            else: # 'excel'
                final_sale = utils.convert_to_float(self.excel_data.get('sales_uploaded', 0))

            if self.cost_choice_var.get() == 'system':
                final_cost = utils.convert_to_float(self.system_data.get('cost_db', 0))
            else: # 'excel'
                final_cost = utils.convert_to_float(self.excel_data.get('cost_uploaded', 0))

            final_gp = final_sale - final_cost
            final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
            
            # --- Database Operation ---
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                update_query = """
                    UPDATE commissions SET
                        final_sales_amount = %s,
                        final_cost_amount = %s,
                        final_gp = %s,
                        final_margin = %s,
                        status = %s
                    WHERE id = %s
                """
                cursor.execute(update_query, (final_sale, final_cost, final_gp, final_margin, 'HR Verified', self.record_id))
            conn.commit()

            messagebox.showinfo("สำเร็จ", "ยืนยันและบันทึกข้อมูลเรียบร้อยแล้ว", parent=self.master)

            if self.refresh_callback:
                self.refresh_callback()
            self.destroy()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_close(self):
        if self.refresh_callback:
            self.refresh_callback()
        self.destroy()
    
    def _update_all_calculations_and_ui(self):
        """
        อัปเดตข้อมูลที่คำนวณได้ทั้งหมด และรีเฟรช UI ที่เกี่ยวข้อง
        (เวอร์ชันแก้ไขสำหรับ UI ใหม่)
        """
        self._recalculate_summaries()
        self._populate_po_cards() # รีเฟรชรายการ PO
        self._update_selection_display() # อัปเดตการ์ดสรุปยอดขายและต้นทุน

    # อยู่ในไฟล์ hr_windows.py ภายในคลาส HRVerificationWindow

    def _update_selection_display(self, *args):
        # อัปเดตการ์ดสรุปยอดขาย
        final_sales_val = self.calculated_values.get('total_sale_system', 0.0)
        source_sales = self.final_sale_source.get()
        if source_sales == "express":
            final_sales_val = self.calculated_values.get('total_sale_express', 0.0)
        self.final_sales_label.configure(text=f"{final_sales_val:,.2f} บาท")
        self.final_sales_source_label.configure(text=f"ที่มา: {'System' if source_sales == 'system' else 'Express'}")

        # อัปเดตการ์ดสรุปต้นทุน
        final_cost_val = self.calculated_values.get('total_cost_system', 0.0)
        source_cost = self.final_cost_source.get()
        if source_cost == "express":
            final_cost_val = self.calculated_values.get('total_cost_express', 0.0)
        self.final_cost_label.configure(text=f"{final_cost_val:,.2f} บาท")
        self.final_cost_source_label.configure(text=f"ที่มา: {'System' if source_cost == 'system' else 'Express'}")


    def _so_create_string_vars(self):
        """สร้าง StringVars ที่จำเป็นสำหรับ SOPopupWindow"""
        self.so_form_widgets = {}
        self.sale_theme = self.app_container.THEME["sale"]
        self.dropdown_style = {
            "fg_color": "white", "text_color": "black",
            "button_color": self.sale_theme.get("primary", "#3B82F6"),
            "button_hover_color": "#2563EB"
        }
        self.so_form_widgets['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
        self.so_form_widgets['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_form_widgets['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_form_widgets['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_form_widgets['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_form_widgets['cash_verification_result_var'] = tk.StringVar(value="-")

    def _save_so_changes_from_popup(self, so_id, so_shared_vars_data, current_popup_widgets_ref):
      """
      Callback function ที่ SOPopupWindow จะเรียกใช้เพื่ออัปเดตข้อมูล
      ฟังก์ชันนี้จะอัปเดตข้อมูลในหน่วยความจำ (Memory) เท่านั้น ยังไม่บันทึกลง DB
    """
    # <<< START: แก้ไขให้ครอบคลุมทุก Widget Type >>>

    # 1. สร้าง Map ที่สมบูรณ์ขึ้น
      key_map = {
        # Numeric Entries
        'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
        'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost',
        'relocation_cost_entry': 'relocation_cost', 'credit_card_fee_entry': 'credit_card_fee',
        'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
        'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons',
        'giveaway_value_entry': 'giveaways', 'cash_product_input_entry': 'cash_product_input',
        'cash_actual_payment_entry': 'cash_actual_payment', 'payment1_amount_entry': 'payment1_amount', # เพิ่ม payment เข้ามา
        'payment2_amount_entry': 'payment2_amount', # เพิ่ม payment เข้ามา

        'payment_before_vat_entry': 'payment_before_vat', 
        'payment_no_vat_entry': 'payment_no_vat',

        # Date Selectors
        'bill_date_selector': 'bill_date', 'delivery_date_selector': 'delivery_date',
        'payment_date_selector': 'payment_date', 'date_to_wh_selector': 'date_to_warehouse',
        'date_to_customer_selector': 'date_to_customer',
        

        # Text Entries
        'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id',
        'credit_term_entry': 'credit_term', 'pickup_location_entry': 'pickup_location',
        'pickup_rego_entry': 'pickup_registration',

        # Option Menus / Radio Buttons (ใช้ StringVar)
        'delivery_type_menu': 'delivery_type', 'sales_service_vat_option': 'sales_service_vat_option',
        'cutting_drilling_fee_vat_option': 'cutting_drilling_fee_vat_option',
        'other_service_fee_vat_option': 'other_service_fee_vat_option',
        'shipping_vat_option_var': 'shipping_vat_option',
        'credit_card_fee_vat_option_var': 'credit_card_fee_vat_option'
    }

    # 2. วนลูปเพื่อดึงค่าจาก Widget แต่ละประเภท
      for widget_key, data_key in key_map.items():
          value = None
          if widget_key in current_popup_widgets_ref: # ถ้าเป็น Widget ที่สร้างใน Popup
              widget = current_popup_widgets_ref[widget_key]
              if not (widget and widget.winfo_exists()): continue

              if isinstance(widget, (NumericEntry, CTkEntry)):
                value = widget.get()
              elif isinstance(widget, DateSelector):
                value = widget.get_date()

          elif widget_key in so_shared_vars_data: # ถ้าเป็น StringVar ที่ใช้ร่วมกัน
              value = so_shared_vars_data[widget_key].get()

        # 3. อัปเดตค่าลงใน self.system_data
          if value is not None:
            # แปลงค่าตัวเลขก่อนเก็บ
              if isinstance(value, str) and data_key not in ['customer_name', 'customer_id', 'credit_term', 'pickup_location', 'pickup_registration', 'delivery_type'] and 'vat_option' not in data_key:
                 self.system_data[data_key] = utils.convert_to_float(value)
              else:
                   self.system_data[data_key] = value

    # 4. คำนวณยอดชำระรวมแยกต่างหาก
      p1 = utils.convert_to_float(current_popup_widgets_ref.get('payment1_amount_entry').get())
      p2 = utils.convert_to_float(current_popup_widgets_ref.get('payment2_amount_entry').get())
      self.system_data['total_payment_amount'] = p1 + p2

    # 5. Refresh หน้าจอหลัก
      self._recalculate_summaries()
      self._refresh_sales_comparison_table()
      self._update_all_calculations_and_ui()

      messagebox.showinfo("อัปเดตข้อมูล", "ข้อมูล SO ถูกอัปเดตในหน้าต่างนี้แล้ว\nกรุณากด 'บันทึกการแก้ไข' เพื่อบันทึกข้อมูลลงฐานข้อมูล", parent=self)
    
    def _open_so_editor_popup(self):
        """เปิดหน้าต่าง SOPopupWindow สำหรับให้ HR แก้ไขข้อมูล SO โดยละเอียด"""
        SOPopupWindow(
            master=self, 
            sales_data=self.system_data, 
            so_shared_vars=self.so_form_widgets, 
            sale_theme=self.sale_theme
        )
    
    def _save_so_changes_from_popup(self, so_id, so_shared_vars_data, current_popup_widgets_ref):
        """
        Callback ที่ถูกเรียกจาก SOPopupWindow เมื่อมีการบันทึก
        ฟังก์ชันนี้จะอัปเดตข้อมูลในหน่วยความจำ (Memory) ของหน้าต่างปัจจุบัน
        """
        key_map = {
            'sales_amount_entry': 'sales_service_amount', 'cutting_drilling_fee_entry': 'cutting_drilling_fee',
            'other_service_fee_entry': 'other_service_fee', 'shipping_cost_entry': 'shipping_cost',
            'relocation_cost_entry': 'relocation_cost', 'credit_card_fee_entry': 'credit_card_fee',
            'transfer_fee_entry': 'transfer_fee', 'wht_fee_entry': 'wht_3_percent',
            'brokerage_fee_entry': 'brokerage_fee', 'coupon_value_entry': 'coupons',
            'giveaway_value_entry': 'giveaways', 'cash_product_input_entry': 'cash_product_input',
            'cash_actual_payment_entry': 'cash_actual_payment', 'bill_date_selector': 'bill_date',
            'delivery_date_selector': 'delivery_date', 'payment_date_selector': 'payment_date',
            'date_to_wh_selector': 'date_to_warehouse', 'date_to_customer_selector': 'date_to_customer',
            'customer_name_entry': 'customer_name', 'customer_id_entry': 'customer_id',
            'credit_term_entry': 'credit_term', 'pickup_location_entry': 'pickup_location',
            'pickup_rego_entry': 'pickup_registration'
        }

        # วนลูปเพื่อดึงค่าจาก Widget แต่ละประเภท
        for widget_key, data_key in key_map.items():
            value = None
            if widget_key in current_popup_widgets_ref:
                widget = current_popup_widgets_ref[widget_key]
                if not (widget and widget.winfo_exists()): continue

                if isinstance(widget, (NumericEntry, CTkEntry)):
                    value = widget.get()
                elif isinstance(widget, DateSelector):
                    value = widget.get_date()
            
            if value is not None:
                # แปลงค่าตัวเลขก่อนเก็บ
                if isinstance(value, str) and 'amount' in data_key or 'cost' in data_key or 'fee' in data_key:
                    self.system_data[data_key] = utils.convert_to_float(value)
                else:
                    self.system_data[data_key] = value

        # คำนวณยอดชำระรวมแยกต่างหาก
        p1 = utils.convert_to_float(current_popup_widgets_ref.get('payment1_amount_entry').get())
        p2 = utils.convert_to_float(current_popup_widgets_ref.get('payment2_amount_entry').get())
        self.system_data['total_payment_amount'] = p1 + p2

        # สั่งให้หน้าต่างหลักคำนวณและ Refresh UI ใหม่ทั้งหมด
        self._update_all_calculations_and_ui()

        messagebox.showinfo("อัปเดตข้อมูล", 
                            "ข้อมูล SO ถูกอัปเดตในหน้าต่างนี้แล้ว\n"
                            "กรุณากด 'บันทึกการแก้ไข' เพื่อยืนยันการเปลี่ยนแปลงลงฐานข้อมูล", 
                            parent=self)

    def _on_cell_double_click(self, event):
        """จัดการ Event เมื่อมีการดับเบิลคลิกที่เซลล์ในตาราง"""
        tree = event.widget
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column_id = tree.identify_column(event.x)
        column_index = int(column_id.replace('#', '')) - 1
        
        # อนุญาตให้แก้ไขเฉพาะคอลัมน์ "ข้อมูลในระบบ" และ "ข้อมูลจาก Express"
        if column_index not in [1, 2]:
            return

        item_id = tree.focus()
        column_box = tree.bbox(item_id, column_id)

        if not column_box:
            return # หยุดการทำงานถ้าหาตำแหน่งเซลล์ไม่เจอ
        
        # --- START: แก้ไขจุดนี้ ---
        # ดึงค่าตำแหน่งและขนาดออกมาจาก bbox ก่อน
        x, y, width, height = column_box
        
        # นำ width และ height มาใส่ตอนสร้าง CTkEntry
        entry_edit = ctk.CTkEntry(tree, width=width, height=height, corner_radius=0)
        
        # ส่วน .place() จะใช้แค่กำหนดตำแหน่ง x, y
        entry_edit.place(x=x, y=y)
        
        current_value = tree.item(item_id, "values")[column_index]
        entry_edit.insert(0, current_value)
        entry_edit.select_range(0, 'end')
        entry_edit.focus_set()

        def on_edit_finished(event, save=True):
            if save:
                new_value_str = entry_edit.get()
                tree.set(item_id, column_id, new_value_str)
                self._save_cell_edit(tree, item_id, column_index, new_value_str)
            entry_edit.destroy()

        entry_edit.bind("<Return>", on_edit_finished)
        entry_edit.bind("<KP_Enter>", on_edit_finished)
        entry_edit.bind("<FocusOut>", on_edit_finished)
        entry_edit.bind("<Escape>", lambda e: on_edit_finished(e, save=False))

    def _save_cell_edit(self, tree, item_id, column_index, new_value_str):
        """บันทึกค่าที่แก้ไขและคำนวณสรุปใหม่"""
        try:
            new_value = float(str(new_value_str).replace(",", ""))
        except (ValueError, TypeError):
            messagebox.showwarning("ข้อมูลผิดพลาด", "กรุณาใส่ค่าเป็นตัวเลขเท่านั้น", parent=self)
            self._recalculate_summaries() # Refresh to original state
            return

        # หาว่าเป็นตาราง Sales หรือ Cost
        is_sales_tree = (tree == self.sales_tree)
        field_name = self.sales_tree.item(item_id, "values")[0] if is_sales_tree else self.costing_tree.item(item_id, "values")[0]
        field_name = field_name.replace(" *", "").strip()

        # Map ชื่อที่แสดงผลกลับไปเป็นชื่อ Key ใน Dictionary
        field_to_key_map = {
            'รายได้ค่าสินค้า/บริการ': ('sales_service_amount', 'sales_service_amount'),
            'รายได้ค่าการจัดส่ง': ('shipping_cost', 'shipping_cost_uploaded'),
            'ต้นทุนค่าสินค้า/บริการ': (None, 'cost_uploaded'),
            'ต้นทุนค่าจัดส่ง': (None, 'shipping_cost_uploaded'),
            'ต้นทุนค่าย้าย': (None, 'relocation_cost_uploaded'),
            'ต้นทุนค่านายหน้า': ('brokerage_fee', 'brokerage_fee_uploaded'),
            'ต้นทุนค่าธรรมเนียมโอน': ('transfer_fee', 'transfer_fee_uploaded'),
        }
        
        keys = field_to_key_map.get(field_name)
        if not keys:
            return

        # อัปเดตข้อมูลใน Memory
# อัปเดตข้อมูลใน Memory
        if column_index == 1:  # System Data
           if is_sales_tree:
              target_dict = self.system_data
              key = keys[0]
              if key:
                 target_dict[key] = new_value
           else:  # ถ้าเป็นตาราง Costing
            self.cost_overrides[field_name] = new_value

        elif column_index == 2:  # Express Data
             target_dict = self.excel_data
             key = keys[1]
             if key:
                target_dict[key] = new_value

        
        # คำนวณทุกอย่างใหม่
        self._recalculate_summaries()
        self._refresh_sales_comparison_table()
        self._refresh_costing_comparison_table()
        self._update_all_calculations_and_ui()
        self._update_all_calculations_and_ui()
# อยู่ในไฟล์ hr_windows.py ภายในคลาส HRVerificationWindow

    def _recalculate_summaries(self):
        """คำนวณค่าสรุปทั้งหมดและกำหนดค่าเริ่มต้นที่ถูกต้อง"""
        # --- ส่วนคำนวณตัวเลขทั้งหมด ---
        total_sale_system = (
            float(self.system_data.get('sales_service_amount', 0) or 0) +
            float(self.system_data.get('cutting_drilling_fee', 0) or 0) +
            float(self.system_data.get('other_service_fee', 0) or 0) -
            float(self.system_data.get('coupons', 0) or 0) # <--- หักคูปองออก
        )
        po_shipping = self.po_data['shipping_to_stock_cost'].sum() + self.po_data['shipping_to_site_cost'].sum()
        po_product_cost = self.po_data['total_cost'].sum()
        po_relocation = self.po_data['relocation_cost'].sum() if 'relocation_cost' in self.po_data.columns else 0
        cost_product_sys = float(self.cost_overrides.get('ต้นทุนค่าสินค้า/บริการ', po_product_cost))
        cost_shipping_sys = float(self.cost_overrides.get('ต้นทุนค่าจัดส่ง', po_shipping))
        cost_relocation_sys = float(self.cost_overrides.get('ต้นทุนค่าย้าย', po_relocation))
        cost_brokerage_sys = float(self.cost_overrides.get('ต้นทุนค่านายหน้า', (self.system_data.get('brokerage_fee', 0) or 0)))
        cost_transfer_sys = float(self.cost_overrides.get('ต้นทุนค่าธรรมเนียมโอน', (self.system_data.get('transfer_fee', 0) or 0)))
        total_cost_system = cost_product_sys + cost_shipping_sys + cost_relocation_sys + cost_brokerage_sys + cost_transfer_sys
        total_sale_express = float(self.excel_data.get('sales_uploaded', 0) or 0) + \
                           float(self.excel_data.get('shipping_cost_uploaded', 0) or 0)
        total_cost_express = float(self.excel_data.get('cost_uploaded', 0) or 0) + \
                           float(self.excel_data.get('shipping_cost_uploaded', 0) or 0) + \
                           float(self.excel_data.get('relocation_cost_uploaded', 0) or 0) + \
                           float(self.excel_data.get('brokerage_fee_uploaded', 0) or 0) + \
                           float(self.excel_data.get('transfer_fee_uploaded', 0) or 0)
        self.calculated_values = {
            'total_sale_system': total_sale_system,
            'total_sale_express': total_sale_express,
            'total_cost_system': total_cost_system,
            'total_cost_express': total_cost_express
        }

        # --- Logic การเลือกค่าเริ่มต้นที่ถูกต้องและสมบูรณ์ ---
        saved_sale_source = self.system_data.get('hr_sale_source')
        if saved_sale_source not in ['system', 'express']:
            self.final_sale_source.set("system")
        saved_cost_source = self.system_data.get('hr_cost_source')
        if saved_cost_source not in ['system', 'express']:
            self.final_cost_source.set("system")
    # --- END: สิ้นสุดโค้ดเวอร์ชันสมบูรณ์ ---
            
    def _create_so_info_section(self):
        frame = CTkFrame(self, fg_color="#F0F0F0")
        frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        font_bold=CTkFont(size=14, weight="bold"); font_normal=CTkFont(size=14)
        
        so_info_frame = CTkFrame(frame, fg_color="transparent")
        so_info_frame.pack(side="left", padx=10, pady=5)
        
        CTkLabel(so_info_frame, text="SO Number:", font=font_bold).grid(row=0, column=0, sticky="w")
        CTkLabel(so_info_frame, text=self.system_data.get('so_number', 'N/A'), font=font_normal).grid(row=0, column=1, padx=5, sticky="w")
        
        sale_name = self.system_data.get('sale_name', self.system_data.get('sale_key', 'N/A'))
        CTkLabel(so_info_frame, text="Sale Name:", font=font_bold).grid(row=0, column=2, padx=10, sticky="w")
        CTkLabel(so_info_frame, text=sale_name, font=font_normal).grid(row=0, column=3, padx=5, sticky="w")
        
        CTkLabel(so_info_frame, text="Customer:", font=font_bold).grid(row=1, column=0, sticky="w")
        CTkLabel(so_info_frame, text=self.system_data.get('customer_name', 'N/A'), font=font_normal, wraplength=400).grid(row=1, column=1, columnspan=3, sticky="w")
        
        CTkButton(frame, text="ดู/แก้ไขข้อมูล SO", 
          command=self._open_so_editor_popup
        ).pack(side="left", padx=20)

    def _create_main_paned_window(self):
        paned_window = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=8, bg="#D1D5DB")
        paned_window.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        left_pane = CTkFrame(paned_window, fg_color="transparent")
        left_pane.grid_rowconfigure(0, weight=1); left_pane.grid_columnconfigure(0, weight=1)
        paned_window.add(left_pane, width=700)
        self._create_sales_info_column(left_pane)
        right_pane = CTkFrame(paned_window, fg_color="transparent")
        right_pane.grid_rowconfigure(0, weight=1); right_pane.grid_columnconfigure(0, weight=1)
        paned_window.add(right_pane, width=700)
        self._create_costing_info_column(right_pane)

    def _create_final_summary_section(self):
        """สร้าง Widget ในส่วนสรุป (จะถูกเรียกแค่ครั้งเดียว)"""
        frame = CTkFrame(self, border_width=1)
        frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        frame.grid_columnconfigure((1, 3), weight=1)

        CTkLabel(frame, text="สรุปและเลือกข้อมูลเพื่อคำนวณ Margin/Commission", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=4, padx=10, pady=10)

        # --- ส่วนเลือกยอดขาย ---
        sales_frame = CTkFrame(frame, fg_color="transparent")
        sales_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        sales_frame.grid_columnconfigure(1, weight=1)
        
        CTkLabel(sales_frame, text="ยอดขายรวมทั้งหมด:", font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w")
        
        # +++ START: แก้ไขโดยการเพิ่ม self. เข้าไปข้างหน้าชื่อตัวแปร +++
        self.sales_system_radio = CTkRadioButton(sales_frame, text="", variable=self.final_sale_source, value="system")
        self.sales_system_radio.grid(row=1, column=0, columnspan=2, sticky="w", padx=20)
        
        self.sales_express_radio = CTkRadioButton(sales_frame, text="", variable=self.final_sale_source, value="express")
        self.sales_express_radio.grid(row=2, column=0, columnspan=2, sticky="w", padx=20)
        # +++ END: สิ้นสุดการแก้ไข +++

        # --- ส่วนเลือกยอดต้นทุน ---
        cost_frame = CTkFrame(frame, fg_color="transparent")
        cost_frame.grid(row=1, column=2, columnspan=2, padx=10, pady=5, sticky="ew")
        cost_frame.grid_columnconfigure(1, weight=1)

        CTkLabel(cost_frame, text="ยอดต้นทุนรวมทั้งหมด:", font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w")
        
        # +++ START: แก้ไขโดยการเพิ่ม self. เข้าไปข้างหน้าชื่อตัวแปร +++
        self.cost_system_radio = CTkRadioButton(cost_frame, text="", variable=self.final_cost_source, value="system")
        self.cost_system_radio.grid(row=1, column=0, columnspan=2, sticky="w", padx=20)
        
        self.cost_express_radio = CTkRadioButton(cost_frame, text="", variable=self.final_cost_source, value="express")
        self.cost_express_radio.grid(row=2, column=0, columnspan=2, sticky="w", padx=20)
        # +++ END: สิ้นสุดการแก้ไข +++

    def _create_sales_info_column(self, parent):
        frame = CTkFrame(parent); frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(1, weight=1); frame.grid_columnconfigure(0, weight=1)
        CTkLabel(frame, text="เปรียบเทียบข้อมูลขาย (ในระบบ vs. Express)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=5)
        self.sales_tree = ttk.Treeview(frame, columns=("Field", "System", "Express"), show="headings")
        self.sales_tree.grid(row=1, column=0, sticky="nsew")
        self.sales_tree.heading("Field", text="ฟิลด์ข้อมูล"); self.sales_tree.heading("System", text="ข้อมูลในระบบ"); self.sales_tree.heading("Express", text="ข้อมูลจาก Express")
        self.sales_tree.column("Field", width=180); self.sales_tree.column("System", width=180, anchor="e"); self.sales_tree.column("Express", width=180, anchor="e")
        self.sales_tree.tag_configure('mismatch', background='#FEE2E2'); self.sales_tree.tag_configure('match', background='#F0FDF4')
        self._refresh_sales_comparison_table()
        self.sales_tree.bind("<Double-1>", self._on_cell_double_click)

    def _create_costing_info_column(self, parent):
        frame = CTkFrame(parent); frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(1, weight=1); frame.grid_columnconfigure(0, weight=1)
        CTkLabel(frame, text="เปรียบเทียบข้อมูลทุน (ในระบบ vs. Express)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=5)
        self.costing_tree = ttk.Treeview(frame, columns=("Field", "System", "Express"), show="headings")
        self.costing_tree.grid(row=1, column=0, sticky="nsew")
        self.costing_tree.heading("Field", text="ฟิลด์ข้อมูล"); self.costing_tree.heading("System", text="ข้อมูลในระบบ (PU)"); self.costing_tree.heading("Express", text="ข้อมูลจาก Express")
        self.costing_tree.column("Field", width=180); self.costing_tree.column("System", width=180, anchor="e"); self.costing_tree.column("Express", width=180, anchor="e")
        self.costing_tree.tag_configure('mismatch', background='#FEE2E2'); self.costing_tree.tag_configure('match', background='#F0FDF4')
        self._refresh_costing_comparison_table()
        self.costing_tree.bind("<Double-1>", self._on_cell_double_click)
        self.costing_tree.bind("<Double-1>", self._on_cell_double_click)

    def _refresh_sales_comparison_table(self):
        for item in self.sales_tree.get_children(): self.sales_tree.delete(item)
        sales_fields_map = {
            'รายได้ค่าสินค้า/บริการ': ('sales_service_amount', 'sales_uploaded'),
            'รายได้ค่าการจัดส่ง': ('shipping_cost', 'shipping_cost_uploaded'),
        }
        for display_name, (sys_key, exp_key) in sales_fields_map.items():
            sys_val = self.system_data.get(sys_key)
            exp_val = self.excel_data.get(exp_key)
            sys_str = f"{sys_val:,.2f}" if isinstance(sys_val, (int, float)) else str(sys_val or 'N/A')
            exp_str = f"{exp_val:,.2f}" if isinstance(exp_val, (int, float)) else str(exp_val or 'N/A')
            tag = 'match' if sys_str == exp_str else 'mismatch'
            self.sales_tree.insert("", "end", values=(display_name, sys_str, exp_str), tags=(tag,))

    def _refresh_costing_comparison_table(self):
        for item in self.costing_tree.get_children(): self.costing_tree.delete(item)
        total_po_shipping_cost = self.po_data['shipping_to_stock_cost'].sum() + self.po_data['shipping_to_site_cost'].sum()
        total_po_product_cost = self.po_data['total_cost'].sum() - total_po_shipping_cost
        total_po_relocation_cost = self.po_data['relocation_cost'].sum() if 'relocation_cost' in self.po_data.columns else 0
        brokerage_cost = self.system_data.get('brokerage_fee', 0) or 0
        transfer_fee_cost = self.system_data.get('transfer_fee', 0) or 0
        cost_fields_map = {
            'ต้นทุนค่าสินค้า/บริการ': (total_po_product_cost, 'cost_uploaded'),
            'ต้นทุนค่าจัดส่ง': (total_po_shipping_cost, 'shipping_cost_uploaded'),
            'ต้นทุนค่าย้าย': (total_po_relocation_cost, 'relocation_cost_uploaded'),
            'ต้นทุนค่านายหน้า': (brokerage_cost, 'brokerage_fee_uploaded'),
            'ต้นทุนค่าธรรมเนียมโอน': (transfer_fee_cost, 'transfer_fee_uploaded'),
        }
        for display_name, (original_sys_val, exp_key) in cost_fields_map.items():
           # ตรวจสอบว่ามีค่าที่ถูกแก้ไขเก็บไว้หรือไม่
           sys_val = self.cost_overrides.get(display_name, original_sys_val)

           display_name_with_star = f"{display_name} *" if display_name in self.cost_overrides else display_name

           exp_val = self.excel_data.get(exp_key, 0) or 0
           tag = 'match' if f"{sys_val:,.2f}" == f"{exp_val:,.2f}" else 'mismatch'
           self.costing_tree.insert("", "end", values=(display_name_with_star, f"{sys_val:,.2f}", f"{exp_val:,.2f}"), tags=(tag,))

    def _create_po_summary_table(self):
        main_frame = CTkScrollableFrame(self, label_text="ใบสั่งซื้อ (PO) ที่เกี่ยวข้อง")
        main_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        if self.po_data.empty: CTkLabel(main_frame, text="ไม่พบข้อมูล Purchase Order ที่เกี่ยวข้อง").pack(pady=20); return
        for _, po_row in self.po_data.iterrows():
            po_id = po_row['id']; card = CTkFrame(main_frame, border_width=1); card.pack(fill="x", padx=5, pady=5)
            header_frame = CTkFrame(card, fg_color="#F3F4F6", corner_radius=6); header_frame.pack(fill="x", padx=3, pady=3); header_frame.grid_columnconfigure(1, weight=1)
            toggle_label = CTkLabel(header_frame, text="▶", font=CTkFont(size=14), cursor="hand2"); toggle_label.grid(row=0, column=0, padx=(5,2))
            po_text = f"PO: {po_row['po_number']}  |  ซัพพลายเออร์: {po_row['supplier_name']}  |  ยอดรวมต้นทุน: {po_row['total_cost']:,.2f}"; info_label = CTkLabel(header_frame, text=po_text, font=CTkFont(size=14)); info_label.grid(row=0, column=1, padx=5, pady=8, sticky="w")
            status_color = "#16A34A" if po_row['status'] == 'Approved' else 'gray'; status_label = CTkLabel(header_frame, text=f"สถานะ: {po_row['status']}", font=CTkFont(size=12, weight="bold"), text_color=status_color); status_label.grid(row=0, column=2, padx=10, sticky="e")
            detail_button = CTkButton(header_frame, text="ดูรายละเอียด PO", width=120, command=lambda p_id=po_id: self.app_container.show_purchase_detail_window(p_id)); detail_button.grid(row=0, column=3, padx=10)
            detail_frame = CTkFrame(card, fg_color="#FAFAFA")
            toggle_widgets = [toggle_label, info_label] 
            for widget in toggle_widgets:
                widget.bind("<Button-1>", lambda e, df=detail_frame, pid=po_id, tl=toggle_label: self._toggle_po_items(df, pid, tl))
    
    def _toggle_po_items(self, detail_frame, po_id, toggle_label):
        if detail_frame.winfo_viewable(): detail_frame.pack_forget(); toggle_label.configure(text="▶")
        else:
            detail_frame.pack(fill="both", expand=True, padx=15, pady=(0,5)); toggle_label.configure(text="▼")
            if not detail_frame.winfo_children():
                try:
                    query = "SELECT product_name, quantity, unit_price, total_price FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id"
                    items_df = pd.read_sql_query(query, self.app_container.pg_engine, params=(po_id,))
                    if items_df.empty: CTkLabel(detail_frame, text="ไม่พบรายการสินค้าใน PO นี้").pack(pady=5); return
                    tree_container = CTkFrame(detail_frame, fg_color="transparent"); tree_container.pack(fill="both", expand=True, pady=5); tree_container.grid_columnconfigure(0, weight=1); tree_container.grid_rowconfigure(0, weight=1)
                    tree = ttk.Treeview(tree_container, columns=("name", "qty", "price", "total"), show="headings"); tree.grid(row=0, column=0, sticky="nsew")
                    scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview); scrollbar.grid(row=0, column=1, sticky="ns"); tree.configure(yscrollcommand=scrollbar.set)
                    tree.heading("name", text="ชื่อสินค้า"); tree.heading("qty", text="จำนวน"); tree.heading("price", text="ราคา/หน่วย"); tree.heading("total", text="ราคารวม")
                    tree.column("name", width=300, anchor="w"); tree.column("qty", width=80, anchor="e"); tree.column("price", width=100, anchor="e"); tree.column("total", width=100, anchor="e")
                    for _, item_row in items_df.iterrows(): tree.insert("", "end", values=(item_row['product_name'], f"{item_row['quantity']:,.2f}", f"{item_row['unit_price']:,.2f}", f"{item_row['total_price']:,.2f}"))
                except Exception as e: CTkLabel(detail_frame, text=f"Error loading items: {e}", text_color="red").pack(pady=5)
    
    def _create_action_buttons(self):
        frame = CTkFrame(self, fg_color="transparent")
        frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew") # <<< แก้ไข: sticky="ew"
        frame.grid_columnconfigure((0,1,2), weight=1) # <<< เพิ่ม: ทำให้ปุ่มขยายเท่ากัน

        CTkButton(frame, text="ตีกลับให้ฝ่ายจัดซื้อ (Reject to PU)", fg_color="#F97316", hover_color="#EA580C", command=self._reject_to_purchasing).grid(row=0, column=0, padx=5, pady=5, sticky="ew")

    # <<< เพิ่ม: ปุ่มบันทึกการแก้ไข >>>
        CTkButton(frame, text="💾 บันทึกการแก้ไข", fg_color="#3B82F6", hover_color="#2563EB", command=self._save_intermediate_changes).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        CTkButton(frame, text="ยืนยันข้อมูลถูกต้อง (Verify Data)", fg_color="#16A34A", hover_color="#15803D", command=self._verify_and_save_data).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
    # hr_windows.py (ฟังก์ชัน _reject_to_purchasing ที่แก้ไขแล้ว)

    def _reject_to_purchasing(self):
        """จัดการ Logic การตีกลับ SO ไปยังฝ่ายจัดซื้อ"""
        # 1. แสดงหน้าต่างให้กรอกเหตุผล
        dialog = CTkInputDialog(text="กรุณาระบุเหตุผลที่ตีกลับ (จะถูกส่งไปให้ฝ่ายจัดซื้อ):", title="ตีกลับ SO")
        reason = dialog.get_input()
        
        if not reason or not reason.strip():
            return # ผู้ใช้กดยกเลิก

        so_number = self.system_data.get('so_number')
        so_id = self.system_data.get('id')

        # 2. เริ่มกระบวนการบันทึกข้อมูลลง Database
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                # 3. อัปเดตสถานะ SO และบันทึกเหตุผล (ถูกต้องแล้ว)
                cursor.execute(
                    "UPDATE commissions SET status = 'Rejected by HR', rejection_reason = %s WHERE id = %s",
                    (reason.strip(), so_id)
                )

                # +++ START: เพิ่มส่วนที่ขาดหายไป +++
                # 4. อัปเดตสถานะของ PO ที่เกี่ยวข้องทั้งหมดให้เป็น 'Rejected'
                rejection_note = f"ตีกลับโดย HR: {reason.strip()}"
                cursor.execute(
                    "UPDATE purchase_orders SET status = 'Rejected', approval_status = 'Rejected', rejection_reason = %s WHERE so_number = %s",
                    (rejection_note, so_number)
                )
                # +++ END: สิ้นสุดส่วนที่เพิ่ม +++

                # 5. ค้นหา User ฝ่ายจัดซื้อทั้งหมดที่เกี่ยวข้องกับ SO นี้ (ถูกต้องแล้ว)
                cursor.execute(
                    "SELECT DISTINCT user_key FROM purchase_orders WHERE so_number = %s",
                    (so_number,)
                )
                pu_keys = [row[0] for row in cursor.fetchall()]

                # 6. สร้าง Notification แจ้งเตือนฝ่ายจัดซื้อ (ถูกต้องแล้ว)
                message = f"SO: {so_number} ถูกตีกลับโดย HR\nเหตุผล: {reason.strip()}"
                for key in pu_keys:
                    cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, is_read) VALUES (%s, %s, FALSE)",
                        (key, message)
                    )
            
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ตีกลับ SO ไปยังฝ่ายจัดซื้อเรียบร้อยแล้ว", parent=self.master)
            
            # 7. Refresh หน้าจอหลักและปิดหน้าต่างนี้ (ถูกต้องแล้ว)
            if self.app_container.hr_screen:
                self.app_container.hr_screen._refresh_comparison_view()
            self._on_close()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาด: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)
    
    def _save_intermediate_changes(self):
      """บันทึกข้อมูล SO และ PU overrides ลง DB โดยไม่เปลี่ยนสถานะ"""
      if not messagebox.askyesno("ยืนยัน", "คุณต้องการบันทึกการเปลี่ยนแปลงทั้งหมดลงฐานข้อมูลใช่หรือไม่?", parent=self):
          return

      so_id = self.system_data.get('id')
      conn = None
      try:
          conn = self.app_container.get_connection()
          with conn.cursor() as cursor:
              columns_to_update = [
                  "sales_service_amount", "shipping_cost", "relocation_cost", "brokerage_fee",
                  "transfer_fee", "cutting_drilling_fee", "other_service_fee", "credit_card_fee",
                  "wht_3_percent", "coupons", "giveaways", "total_payment_amount",
                  "cash_product_input", "cash_actual_payment", "bill_date", "delivery_date",
                  "payment_date", "date_to_warehouse", "date_to_customer", "customer_name",
                  "customer_id", "credit_term", "delivery_type", "pickup_location", "pickup_registration",
                  "payment_before_vat", 
                  "payment_no_vat"      
              ]
              set_clauses = [f"{col} = %s" for col in columns_to_update]
              params = [self.system_data.get(col) for col in columns_to_update]

              cost_overrides_json = json.dumps(self.cost_overrides)
              set_clauses.append("hr_cost_overrides = %s")
              params.append(cost_overrides_json)

              set_clauses.append("hr_sale_source = %s")
              params.append(self.final_sale_source.get())
              set_clauses.append("hr_cost_source = %s")
              params.append(self.final_cost_source.get())

              set_clauses.append("cost_multiplier = %s")
              params.append(float(self.cost_multiplier_var.get()))

              update_query = f"UPDATE commissions SET {', '.join(set_clauses)} WHERE id = %s"
              params.append(so_id)

              # --- START: เพิ่มโค้ดสำหรับ Debug ---
              print("================ DEBUG SAVE DATA ================")
              print(f"--- 1. กำลังจะบันทึก cost_overrides: {self.cost_overrides}")
              print(f"--- 2. คำสั่ง SQL: {cursor.mogrify(update_query, tuple(params)).decode('utf-8')}")

              cursor.execute(update_query, tuple(params))

              print(f"--- 3. จำนวนแถวที่ถูกอัปเดต: {cursor.rowcount}")
              print("=================================================")
            # --- END: สิ้นสุดโค้ดสำหรับ Debug ---

          conn.commit()
          messagebox.showinfo("สำเร็จ", "บันทึกการแก้ไขข้อมูลเรียบร้อยแล้ว", parent=self)
      except Exception as e:
          if conn: conn.rollback()
          messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึก: {e}", parent=self)
          traceback.print_exc()
      finally:
          if conn: self.app_container.release_connection(conn)

    def _verify_and_save_data(self):
        # --- ดึงค่าจากตัวแปรกลางที่คำนวณไว้แล้ว ---
        final_sale_source = self.final_sale_source.get()
        final_cost_source = self.final_cost_source.get()

        total_sale_system = self.calculated_values.get('total_sale_system', 0.0)
        total_sale_express = self.calculated_values.get('total_sale_express', 0.0)
        total_cost_system = self.calculated_values.get('total_cost_system', 0.0)
        total_cost_express = self.calculated_values.get('total_cost_express', 0.0)

        final_sale = float(total_sale_system if final_sale_source == "system" else total_sale_express)
        final_cost = float(total_cost_system if final_cost_source == "system" else total_cost_express)
        
        so_id = self.system_data.get('id')
        final_gp = final_sale - final_cost
        final_margin = (final_gp / final_sale) * 100 if final_sale != 0 else 0
        
        msg = (f"คุณต้องการยืนยันข้อมูลสำหรับ SO นี้ใช่หรือไม่?\n\n"
               f"ยอดขายสุดท้ายที่เลือก: {final_sale:,.2f} บาท\n"
               f"ยอดต้นทุนสุดท้ายที่เลือก: {final_cost:,.2f} บาท\n"
               # ... (ข้อความที่เหลือ) ...
              )
        if not messagebox.askyesno("ยืนยันข้อมูล", msg, parent=self):
            return
            
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                update_query = """
                    UPDATE commissions 
                    SET 
                        status = 'HR Verified', 
                        final_sales_amount = %s, 
                        final_cost_amount = %s,
                        final_gp = %s,
                        final_margin = %s,
                        hr_sale_source = %s,
                        hr_cost_source = %s
                    WHERE id = %s
                """
                params = (final_sale, final_cost, final_gp, final_margin, 
                          self.final_sale_source.get(), self.final_cost_source.get(), so_id)
                cursor.execute(update_query, params)
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ยืนยันและบันทึกข้อมูลเรียบร้อยแล้ว", parent=self.master)
            self._on_close()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

class PayoutDetailWindow(CTkToplevel):
    """
    หน้าต่างสำหรับแสดงรายละเอียดของ SO ทั้งหมดที่อยู่ในรอบการจ่ายค่าคอมมิชชั่นครั้งนั้นๆ
    """
    def __init__(self, master, app_container, payout_id):
        super().__init__(master)
        self.app_container = app_container
        self.payout_id = payout_id
        
        self.title(f"รายละเอียดการจ่ายค่าคอม ID: {payout_id}")
        self.geometry("900x600")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- ส่วนแสดงข้อมูลสรุป ---
        self.summary_frame = CTkFrame(self, corner_radius=10)
        self.summary_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # --- ส่วนแสดงข้อความ ---
        CTkLabel(self, text="รายการ SO ทั้งหมดในรอบการจ่ายเงินนี้:", font=CTkFont(size=14, weight="bold")).grid(row=1, column=0, padx=10, pady=(5,0), sticky="w")

        # --- ส่วนแสดงตาราง ---
        self.tree_frame = CTkFrame(self)
        self.tree_frame.grid(row=2, column=0, padx=10, pady=(0,10), sticky="nsew")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)

        self.after(50, self._load_and_display_details)
        
        self.transient(master)
        self.grab_set()

    def _load_and_display_details(self):
        """โหลดข้อมูลจาก DB และสร้าง UI"""
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                cursor.execute("SELECT *, notes FROM commission_payout_logs WHERE id = %s", (self.payout_id,))
                log_data = cursor.fetchone()

                if not log_data:
                    messagebox.showerror("ไม่พบข้อมูล", "ไม่พบประวัติการจ่ายเงินสำหรับ ID นี้", parent=self)
                    self.destroy()
                    return

                summary_list = log_data['summary_json'] 
                net_commission = next((item['value'] for item in summary_list if 'หลังหัก ณ ที่จ่าย' in item['description']), 0.0)
                
                info_text = (
                    f"พนักงานขาย: {log_data['sale_key']} | แผน: {log_data['plan_name']} | "
                    f"วันที่จ่าย: {pd.to_datetime(log_data['timestamp']).strftime('%d/%m/%Y %H:%M')} | "
                    f"ยอดสุทธิ: {net_commission:,.2f} บาท"
                )
                CTkLabel(self.summary_frame, text=info_text, font=CTkFont(size=14)).pack(pady=10)

                payout_notes = log_data.get('notes')
                if payout_notes and payout_notes.strip():
                    notes_frame = CTkFrame(self.summary_frame, fg_color="#fefce8", border_width=1, border_color="#facc15")
                    notes_frame.pack(fill="x", padx=10, pady=(0, 10))
                    notes_label = CTkLabel(notes_frame, text=f"หมายเหตุ: {payout_notes}", wraplength=800, justify="left", text_color="#ca8a04")
                    notes_label.pack(pady=5, padx=5)

                so_numbers = tuple(log_data['so_numbers_json'])
                
                if not so_numbers:
                    CTkLabel(self.tree_frame, text="ไม่พบรายการ SO ในรอบการจ่ายเงินนี้").pack(pady=20)
                    return
                    
                # --- START: แก้ไข Query ตรงนี้ ---
                # เพิ่มเงื่อนไข AND is_active = 1 เพื่อกรองเอาเฉพาะข้อมูลล่าสุด
                placeholders = ', '.join(['%s'] * len(so_numbers))
                query = f"""
                    SELECT so_number, final_sales_amount, final_margin
                    FROM commissions 
                    WHERE so_number IN ({placeholders}) AND is_active = 1
                """
                # --- END: สิ้นสุดการแก้ไข ---
                df = pd.read_sql_query(query, self.app_container.pg_engine, params=so_numbers)

                df['status'] = df['final_margin'].apply(lambda x: 'Normal' if pd.notna(x) and x >= 10.0 else 'Below Tier')
                
                self._create_detail_table(df)

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดรายละเอียด: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)
            
    def _create_detail_table(self, df):
        """สร้าง Treeview สำหรับแสดงรายละเอียด SO"""
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Detail.Treeview.Heading", font=CTkFont(size=12, weight="bold"))
        style.configure("Detail.Treeview", rowheight=28, font=CTkFont(size=12))

        columns = ['SO Number', 'สถานะ (คำนวณ)', 'ยอดขายสุดท้าย', 'Margin สุดท้าย (%)']
        tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", style="Detail.Treeview")
        
        tree.heading('SO Number', text='SO Number')
        tree.heading('สถานะ (คำนวณ)', text='สถานะ (คำนวณ)')
        tree.heading('ยอดขายสุดท้าย', text='ยอดขายสุดท้าย (บาท)')
        tree.heading('Margin สุดท้าย (%)', text='Margin สุดท้าย (%)')
        
        tree.column('SO Number', width=150, anchor='w')
        tree.column('สถานะ (คำนวณ)', width=120, anchor='center')
        tree.column('ยอดขายสุดท้าย', width=150, anchor='e')
        tree.column('Margin สุดท้าย (%)', width=150, anchor='e')
        
        tree.tag_configure('normal_row', background='#F0FDF4')
        tree.tag_configure('below_row', background='#FEFCE8')

        for _, row in df.iterrows():
            tag = 'normal_row' if row['status'] == 'Normal' else 'below_row'
            values = (
                row['so_number'],
                row['status'],
                f"{row['final_sales_amount']:,.2f}",
                f"{row['final_margin']:,.2f}%"
            )
            tree.insert("", "end", values=values, tags=(tag,))

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

class ComparisonDetailViewer(CTkToplevel):
    def __init__(self, master, detail_df):
        super().__init__(master)
        self.title("รายละเอียดผลการเปรียบเทียบ")
        self.geometry("1100x500")
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_frame = CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # ใช้ฟังก์ชันสร้างตารางสวยงามที่คัดลอกมา
        self.theme = master.app_container.THEME["hr"]
        self.header_font_table = master.header_font_table
        self.entry_font = CTkFont(size=14)
        self._create_styled_dataframe_table(main_frame, detail_df)

        self.transient(master)
        self.grab_set()

    def _create_styled_dataframe_table(self, parent, df, label_text="", on_row_click=None, status_colors=None, status_column=None):
        for widget in parent.winfo_children():
            widget.destroy()
        if df is None or df.empty:
            CTkLabel(parent, text=f"ไม่พบข้อมูลสำหรับ '{label_text}'").pack(pady=20)
            return
        
        # --- START: โค้ดส่วนที่แก้ไข ---
        # 1. สร้าง Dictionary สำหรับแปลงชื่อหัวข้อ
        header_map = {
            'so_number': 'SO Number',
            'sales_service_amount': 'ยอดขาย (ระบบ)',
            'sales_uploaded': 'ยอดขาย (Express)',
            'cost_db': 'ต้นทุน (ระบบ)',
            'cost_uploaded': 'ต้นทุน (Express)',
            'margin_db': 'Margin (ระบบ) %',
            'margin_uploaded': 'Margin (Express) %',
            'ผลต่างยอดขาย': 'ผลต่างยอดขาย',
            'ผลต่างต้นทุน': 'ผลต่างต้นทุน',
            'สถานะ': 'สถานะ'
        }
        
        # 2. เปลี่ยนชื่อคอลัมน์ใน DataFrame ก่อนนำไปแสดงผล
        df_display = df.rename(columns=header_map)
        # --- END: สิ้นสุดโค้ดส่วนที่แก้ไข ---

        container = CTkFrame(parent, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        if label_text:
            CTkLabel(container, text=label_text, font=self.header_font_table).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        tree_frame = CTkFrame(container, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = df_display.columns.tolist()
        
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Custom.Treeview.Heading", 
                        font=self.header_font_table, 
                        background=self.theme["primary"],
                        foreground="white",
                        relief="flat")
        style.map("Custom.Treeview.Heading",
                background=[('active', self.theme["header"])])
        
        style.configure("Custom.Treeview", 
                        rowheight=28, 
                        font=self.entry_font
                        )
        
        style.map("Custom.Treeview",
                background=[('selected', self.theme["primary"])],
                foreground=[('selected', "white")])

        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Custom.Treeview", height=15)
        tree.grid(row=0, column=0, sticky="nsew")

        # กำหนดสีตามสถานะที่อาจมี
        status_colors_map = {
            "ผ่านเกณฑ์": "#BBF7D0", "ยอดขายต่ำกว่า Express": "#FECACA",
            "ต้นทุนต่ำกว่า Express": "#FEF08A", "มีในไฟล์, ไม่มีในระบบ": "#FECACA",
            "มีในระบบ, ไม่มีในไฟล์": "#FEF08A", "ข้อมูลไม่ตรงกัน": "#FED7AA",
            "กำไรดี": "#BBF7D0", "กำไรน้อย": "#FEF08A",
            "ขาดทุน": "#FECACA", "ยืนยันแล้ว (รอผล)": "#E5E7EB"
        }
        for status, color in status_colors_map.items():
            tree.tag_configure(status, background=color)

        for col_id in columns:
            header_text = col_id
            tree.heading(col_id, text=header_text)
            width = 150
            if "ยอดขาย" in col_id or "ต้นทุน" in col_id:
                tree.column(col_id, width=120, anchor='e')
            elif "Margin" in col_id:
                tree.column(col_id, width=100, anchor='e')
            else:
                tree.column(col_id, width=width, anchor='w')

        for index, row in df_display.iterrows():
            # ดึงค่า 'สถานะ' จาก df_display เพื่อใช้กำหนดสี
            status_value = row.get('สถานะ', '')
            tags = [status_value] if status_value in status_colors_map else []
            
            values = []
            for col_name in columns:
                value = row[col_name]
                if pd.notna(value):
                    if isinstance(value, (float, np.floating)): values.append(f"{value:,.2f}")
                    else: values.append(str(value))
                else:
                    values.append("")
            
            # ใช้ so_number จาก DataFrame ต้นฉบับ (df) เพื่อเป็น ID ที่ไม่ซ้ำกัน
            iid_value = df.iloc[index]['SO Number']
            tree.insert("", "end", values=values, tags=tuple(tags), iid=str(iid_value))
        
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        if on_row_click: 
            tree.bind("<Double-1>", lambda e: on_row_click(e, tree, df))

class ComparisonHistoryWindow(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.master_screen = master
        self.all_logs_df = pd.DataFrame()
        self.detail_df = pd.DataFrame()
        self.current_view = 'summary'
        
        self.theme = self.app_container.THEME["hr"]
        self.header_font_table = ctk.CTkFont(size=14, weight="bold")
        self.entry_font = ctk.CTkFont(size=14)

        # --- START: เพิ่มตัวแปรสำหรับฟิลเตอร์เดือน/ปี ---
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        self.month_var = tk.StringVar(value="ทุกเดือน")
        self.year_var = tk.StringVar(value="ทุกปี")
        # --- END ---

        self.title("ประวัติการเปรียบเทียบข้อมูล")
        self.geometry("1200x700")
        
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- ส่วนควบคุมและฟิลเตอร์ ---
        self.control_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.back_button = ctk.CTkButton(self.control_frame, text="◀ กลับไปหน้าสรุป", command=self._show_summary_view)
        self.back_button.pack(side="left", padx=5)
        self.back_button.pack_forget() 

        # --- START: เพิ่ม Dropdown สำหรับเดือนและปี ---
        month_options = ["ทุกเดือน"] + self.thai_months
        self.month_menu = ctk.CTkOptionMenu(self.control_frame, variable=self.month_var, values=month_options)
        self.month_menu.pack(side="left", padx=(10, 5))

        current_year = datetime.now().year
        year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]
        self.year_menu = ctk.CTkOptionMenu(self.control_frame, variable=self.year_var, values=year_options)
        self.year_menu.pack(side="left", padx=5)
        # --- END ---

        self.search_entry = ctk.CTkEntry(self.control_frame, placeholder_text="ค้นหา...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(10, 5))
        self.search_entry.bind("<KeyRelease>", lambda e: self._apply_filter())
        
        ctk.CTkButton(self.control_frame, text="ค้นหา", width=80, command=self._apply_filter).pack(side="left")

        # --- Label แสดงสถานะ ---
        status_bar_frame = ctk.CTkFrame(self, fg_color="transparent")
        status_bar_frame.grid(row=1, column=0, padx=10, sticky="ew")

        self.title_label = ctk.CTkLabel(status_bar_frame, text="ภาพรวมการเปรียบเทียบ (ดับเบิลคลิกเพื่อดูรายละเอียด)", font=ctk.CTkFont(size=16, weight="bold"))
        self.title_label.pack(side="left")

        self.count_label = ctk.CTkLabel(status_bar_frame, text="กำลังโหลด...", text_color="gray")
        self.count_label.pack(side="right")

        # --- Frame สำหรับตาราง ---
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.grid(row=2, column=0, padx=10, pady=(0,10), sticky="nsew")
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        self.after(50, self._load_and_display_history)
        self.transient(master)
        self.grab_set()

    def _show_summary_view(self):
        """สลับกลับไปแสดงผลหน้าสรุป Log"""
        self.current_view = 'summary'
        self.title_label.configure(text="ภาพรวมการเปรียบเทียบ (ดับเบิลคลิกเพื่อดูรายละเอียด)")
        self.search_entry.configure(placeholder_text="ค้นหาไฟล์หรือผู้จัดทำ...")
        self.back_button.pack_forget()
        self.month_menu.pack(side="left", padx=(10, 5))
        self.year_menu.pack(side="left", padx=5)
        self._apply_filter() # <--- แก้ไขให้เรียกใช้ _apply_filter แทน

    def _show_log_details(self, log_id):
        """สลับไปแสดงผลหน้ารายละเอียดของ Log ที่เลือก"""
        self.current_view = 'detail'
        log_row = self.all_logs_df[self.all_logs_df['ID'] == log_id].iloc[0]
        timestamp = pd.to_datetime(log_row['เวลาที่ทำรายการ']).strftime('%Y-%m-%d %H:%M')
        
        self.title_label.configure(text=f"รายละเอียด Log ID: {log_id} (เวลา: {timestamp}) - ดับเบิลคลิก SO เพื่อตรวจสอบ")
        self.search_entry.configure(placeholder_text="ค้นหา SO...")
        # แสดงปุ่ม Back และซ่อนฟิลเตอร์เดือน/ปี
        self.back_button.pack(side="left", padx=5)
        self.month_menu.pack_forget()
        self.year_menu.pack_forget()

        details_list = log_row['detail_json_hidden']
        if details_list and isinstance(details_list, list):
            self.detail_df = pd.DataFrame(details_list)
        else:
            self.detail_df = pd.DataFrame()
        self._populate_treeview(self.detail_df)

    def _load_and_display_history(self):
        """โหลดข้อมูล Log ทั้งหมด และแปลง JSON ให้เป็น DataFrame ที่ใช้งานได้"""
        try:
            query = "SELECT id, timestamp, hr_user_key, salesperson_filter, source_info, summary_json, detail_json FROM comparison_logs ORDER BY timestamp DESC LIMIT 200"
            logs_df = pd.read_sql_query(query, self.app_container.pg_engine)

            if logs_df.empty:
                self.all_logs_df = pd.DataFrame()
                self._populate_treeview()
                return

            def safe_json_normalize(series):
                processed_data = []
                for item in series:
                    if isinstance(item, str):
                        try: processed_data.append(json.loads(item))
                        except json.JSONDecodeError: processed_data.append({})
                    elif isinstance(item, dict):
                        processed_data.append(item)
                    else:
                        processed_data.append({})
                return pd.json_normalize(processed_data)

            summary_df = safe_json_normalize(logs_df['summary_json'])
            logs_df = pd.concat([logs_df.drop(columns=['summary_json']), summary_df], axis=1)
            
            self.all_logs_df = logs_df.rename(columns={
                'id': 'ID', 'timestamp': 'เวลาที่ทำรายการ', 'hr_user_key': 'ทำโดย (HR)',
                'salesperson_filter': 'ข้อมูลของเซลส์', 'source_info': 'ไฟล์/แหล่งข้อมูล',
                'total_records': 'ยอดรวม', 'matched_records': 'ตรงกัน', 'diff_records': 'แตกต่าง',
                'detail_json': 'detail_json_hidden'
            })
            self._show_summary_view()

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถโหลดประวัติได้: {e}", parent=self)
            traceback.print_exc()

    def _apply_filter(self, *args):
        """กรองข้อมูลใน DataFrame ตามเงื่อนไขที่เลือกและอัปเดตตาราง"""
        if self.all_logs_df.empty:
            self._populate_treeview(self.all_logs_df)
            return
            
        df = self.all_logs_df.copy()

        # --- START: เพิ่ม Logic การกรองด้วยเดือนและปี ---
        df['เวลาที่ทำรายการ'] = pd.to_datetime(df['เวลาที่ทำรายการ'])
        
        selected_month_str = self.month_var.get()
        if selected_month_str != "ทุกเดือน":
            month_num = self.thai_month_map[selected_month_str]
            df = df[df['เวลาที่ทำรายการ'].dt.month == month_num]

        selected_year_str = self.year_var.get()
        if selected_year_str != "ทุกปี":
            year_num = int(selected_year_str)
            df = df[df['เวลาที่ทำรายการ'].dt.year == year_num]
        # --- END ---
            
        search_term = self.search_entry.get().strip().lower()
        if search_term:
             if self.current_view == 'summary':
                df = df[
                    df['ทำโดย (HR)'].str.lower().str.contains(search_term, na=False) |
                    df['ไฟล์/แหล่งข้อมูล'].str.lower().str.contains(search_term, na=False)
                ]
             else: # detail view
                if not self.detail_df.empty:
                    df = self.detail_df[self.detail_df['SO Number'].str.lower().str.contains(search_term, na=False)]

        self._populate_treeview(df)

    def _populate_treeview(self, df):
        # ... (โค้ดส่วนนี้เหมือนเดิม ไม่มีการเปลี่ยนแปลง) ...
        for widget in self.tree_frame.winfo_children():
            widget.destroy()

        self.count_label.configure(text=f"พบ {len(df)} รายการ")
        if df.empty:
            ctk.CTkLabel(self.tree_frame, text="ไม่พบข้อมูล").pack(pady=20)
            return

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("History.Treeview.Heading", font=self.header_font_table, background=self.theme['primary'], foreground="white", relief="flat")
        style.configure("History.Treeview", rowheight=28, font=self.entry_font, fieldbackground="#FFFFFF")
        style.map("History.Treeview", background=[('selected', self.theme["header"])])
        
        tree = ttk.Treeview(self.tree_frame, show="headings", style="History.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        if self.current_view == 'summary':
            columns = ['ID', 'เวลาที่ทำรายการ', 'ทำโดย (HR)', 'ข้อมูลของเซลส์', 'ไฟล์/แหล่งข้อมูล', 'ยอดรวม', 'ตรงกัน', 'แตกต่าง']
            tree['columns'] = columns
            for col in columns:
                width = 80 if col in ['ID', 'ยอดรวม', 'ตรงกัน', 'แตกต่าง'] else 180
                tree.heading(col, text=col)
                tree.column(col, width=width, anchor='w')
            
            tree.tag_configure('matched_ok', background='#F0FDF4')
            tree.tag_configure('mismatched', background='#FEFCE8')

            for _, row in df.iterrows():
                diff_count = row.get('แตกต่าง', 0)
                tag = 'mismatched' if pd.to_numeric(diff_count, errors='coerce') > 0 else 'matched_ok'
                timestamp = pd.to_datetime(row['เวลาที่ทำรายการ']).strftime('%Y-%m-%d %H:%M')
                values = (row['ID'], timestamp, row['ทำโดย (HR)'], row['ข้อมูลของเซลส์'], row['ไฟล์/แหล่งข้อมูล'], 
                          row.get('total_records', 0), row.get('matched_records', 0), diff_count)
                tree.insert("", "end", values=values, iid=row['ID'], tags=(tag,))
        else: # detail view
            columns = ['SO Number', 'ยอดขาย (ระบบ)', 'ยอดขาย (Express)', 'ต้นทุน (ระบบ)', 'ต้นทุน (Express)', 'สถานะ']
            tree['columns'] = columns
            status_colors_map = {
                "ผ่านเกณฑ์": "#F0FDF4", "ยอดขายต่ำกว่า Express": "#FEF2F2",
                "ต้นทุนต่ำกว่า Express": "#FEFCE8", "มีในไฟล์, ไม่มีในระบบ": "#FEF2F2",
                "มีในระบบ, ไม่มีในไฟล์": "#FEFCE8", "ข้อมูลไม่ตรงกัน": "#FFF7ED",
            }
            for status, color in status_colors_map.items():
                tree.tag_configure(status, background=color)

            for col in columns:
                anchor = 'e' if 'ยอด' in col or 'ต้นทุน' in col else 'w'
                width = 150
                tree.heading(col, text=col)
                tree.column(col, anchor=anchor, width=width)
            
            for _, row in df.iterrows():
                tag = row.get('สถานะ', '')
                tags_tuple = (tag,) if tag else ()
                values = (
                    row.get('SO Number'), f"{row.get('ยอดขาย (ระบบ)', 0):,.2f}", f"{row.get('ยอดขาย (Express)', 0):,.2f}",
                    f"{row.get('ต้นทุน (ระบบ)', 0):,.2f}", f"{row.get('ต้นทุน (Express)', 0):,.2f}", row.get('สถานะ')
                )
                unique_iid = f"{row.get('log_id')}-{row.get('SO Number')}"
                tree.insert("", "end", values=values, tags=tags_tuple, iid=unique_iid)
        
        tree.bind("<Double-1>", self._on_row_double_click)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vsb.set)

    def _on_row_double_click(self, event):
        tree = event.widget
        selected_iid = tree.focus()
        if not selected_iid: return

        if self.current_view == 'summary':
            log_id = int(tree.item(selected_iid, "values")[0])
            self._show_log_details(log_id)
        else: # detail view
            so_number = tree.item(selected_iid, "values")[0]
            self.master_screen._open_verification_window(so_number)
# hr_windows.py (เพิ่มคลาสใหม่นี้ต่อท้ายไฟล์)

class PayoutCalculationViewer(CTkToplevel):
    def __init__(self, master, app_container, payout_id):
        super().__init__(master)
        self.app_container = app_container
        self.payout_id = payout_id
        self.so_numbers = []

        # เตรียม Theme และ Font สำหรับฟังก์ชันสร้างตาราง
        self.theme = self.app_container.THEME["hr"]
        self.header_font_table = CTkFont(size=14, weight="bold")
        self.entry_font = CTkFont(size=14)

        self.title(f"รายละเอียดการคำนวณค่าคอม (Payout ID: {payout_id})")
        self.geometry("800x600")
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        top_frame = CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.summary_label = CTkLabel(top_frame, text="สรุปการคำนวณ", font=self.header_font_table)
        self.summary_label.pack(side="left")

        # เพิ่มปุ่มสำหรับกดดูรายชื่อ SO
        CTkButton(top_frame, text="ดูรายชื่อ SO ที่เกี่ยวข้อง", command=self._open_so_list_viewer).pack(side="right")
        
        self.table_container = CTkFrame(self)
        self.table_container.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

        self.after(50, self._load_and_display_data)
        
        self.transient(master)
        self.grab_set()

    def _open_so_list_viewer(self):
        # เปิดหน้าต่างแสดงรายชื่อ SO (PayoutDetailWindow เดิม)
        PayoutDetailWindow(master=self, app_container=self.app_container, payout_id=self.payout_id)

    def _load_and_display_data(self):
        try:
            query = "SELECT summary_json, so_numbers_json, sale_key, plan_name, timestamp FROM commission_payout_logs WHERE id = %s"
            log_data = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.payout_id,)).iloc[0]

            summary_df = pd.DataFrame(log_data['summary_json'])
            self.so_numbers = log_data['so_numbers_json']
            
            # อัปเดต Label ด้านบน
            info_text = f"สรุปการคำนวณสำหรับ: {log_data['sale_key']} (แผน: {log_data['plan_name']}) - {pd.to_datetime(log_data['timestamp']).strftime('%d/%m/%Y')}"
            self.summary_label.configure(text=info_text)

            # เรียกใช้ฟังก์ชันสร้างตารางสรุป
            self._create_commission_summary_table(summary_df, self.table_container)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถโหลดรายละเอียดการคำนวณได้: {e}", parent=self)
            self.destroy()

    # **สำคัญมาก:** คัดลอกฟังก์ชันนี้มาจาก hr_screen.py
    def _create_commission_summary_table(self, summary_df, container=None):
        if container is None:
            container = self
            
        for widget in container.winfo_children(): widget.destroy()

        if summary_df is None or summary_df.empty:
            CTkLabel(container, text="ไม่พบข้อมูลสำหรับสร้างสรุป").pack(pady=20)
            return

        tree_frame = CTkFrame(container, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Summary.Treeview.Heading", font=self.header_font_table, background="#3B82F6", foreground="white")
        style.configure("Summary.Treeview", rowheight=30, font=self.entry_font)
        style.map("Summary.Treeview", background=[('selected', "#DBEAFE")])

        columns_to_show = list(summary_df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns_to_show, show="headings", style="Summary.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        header_map = {'description': 'รายการสรุป', 'value': 'ยอดรวม (บาท)'}
        for col_id in columns_to_show:
            header_text = header_map.get(col_id, col_id)
            anchor = 'e' if col_id == 'value' else 'w'
            width = 400 if col_id == 'description' else 200
            tree.heading(col_id, text=header_text)
            tree.column(col_id, width=width, anchor=anchor)

        tree.tag_configure('summary_row', font=self.header_font_table, background="#F3F4F6")
        tree.tag_configure('final_row', font=self.header_font_table, background="#D1FAE5")

        for _, row in summary_df.iterrows():
            values_tuple = []
            for col in columns_to_show:
                val = row[col]
                if isinstance(val, (int, float)):
                    values_tuple.append(f"{val:,.2f}")
                else:
                    values_tuple.append(val)
            
            desc = row['description']
            tags = ()
            if any(s in desc for s in ["สรุป", "รวม", "ขั้นต้น"]): tags = ('summary_row',)
            if "หลังหัก" in desc: tags = ('final_row',)
            
            tree.insert("", "end", values=tuple(values_tuple), tags=tags)