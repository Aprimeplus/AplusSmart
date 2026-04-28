import tkinter as tk
from tkinter import messagebox, ttk
from customtkinter import (CTkFrame, CTkLabel, CTkEntry, CTkButton, 
                           CTkOptionMenu, CTkFont, CTkTabview, CTkRadioButton, 
                           CTkScrollableFrame)
import pandas as pd
from datetime import datetime
# หากไม่มี custom_widgets ให้ใช้ standard entry แทน หรือคอมเมนต์ออกถ้า error
try:
    from custom_widgets import NumericEntry, DateSelector
except ImportError:
    # Fallback classes กรณีไม่มีไฟล์ custom_widgets
    class NumericEntry(CTkEntry): pass
    class DateSelector(CTkFrame):
        def __init__(self, master, dropdown_style=None, **kwargs):
            super().__init__(master, **kwargs)
            self.d = tk.StringVar(); self.m = tk.StringVar(); self.y = tk.StringVar()
        def get_date(self): return datetime.now().strftime("%Y-%m-%d")
        def grid(self, **kwargs): super().grid(**kwargs)

class TransportAdminScreen(CTkFrame):
    def __init__(self, master, app_container, user_key):
        super().__init__(master, corner_radius=0, fg_color="#FFFBEB")
        self.app_container = app_container
        self.user_key = user_key
        self.pg_engine = app_container.pg_engine
        
        self.current_selected_id = None 
        self.current_selected_px = None
        
        # Fonts
        self.header_font = CTkFont(size=22, weight="bold", family="TH Sarabun New")
        self.section_font = CTkFont(size=18, weight="bold", family="TH Sarabun New")
        self.label_font = CTkFont(size=16, weight="bold", family="TH Sarabun New")
        self.normal_font = CTkFont(size=16, family="TH Sarabun New")
        self.result_font = CTkFont(size=14, weight="bold", family="TH Sarabun New")
        self.status_font = CTkFont(size=14, family="TH Sarabun New")

        self.grid_columnconfigure(0, weight=0) 
        self.grid_columnconfigure(1, weight=1) 
        self.grid_rowconfigure(0, weight=1)

        # =================================================================
        # 🟢 LEFT SIDE: SCROLLABLE FORM
        # =================================================================
        self.left_frame = CTkFrame(self, fg_color="transparent")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.left_frame.grid_rowconfigure(1, weight=1)
        
        self._init_header_section(self.left_frame)
        
        self.form_scroll = CTkScrollableFrame(self.left_frame, fg_color="white", corner_radius=10, label_text="แบบฟอร์มบันทึกค่าขนส่ง")
        self.form_scroll.grid(row=1, column=0, sticky="nsew", pady=5)
        self.form_scroll.grid_columnconfigure(0, weight=1)
        
        self._init_stock_section(self.form_scroll) # สีแดง
        self._init_site_section(self.form_scroll)  # สีน้ำเงิน
        
        self._init_footer_buttons(self.left_frame)

        # =================================================================
        # 🔵 RIGHT SIDE: TABS
        # =================================================================
        self.right_frame = CTkFrame(self, fg_color="transparent")
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=10)
        
        self.tabview = CTkTabview(self.right_frame, width=500)
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_history = self.tabview.add("📜 ประวัติค่าขนส่ง (History)")
        self.tab_pending = self.tabview.add("⏳ PO ที่ยังไม่มีค่ารถ (Pending)")

        self._setup_history_tab()
        self._setup_pending_tab()

        CTkButton(self, text="ออกจากระบบ", command=self._logout, fg_color="#EF4444", width=80).place(relx=0.98, rely=0.02, anchor="ne")

        self.after(500, self._load_history_data)
        self.after(500, self._load_pending_data)

    # -------------------------------------------------------------------------
    #  UI CONSTRUCTION METHODS
    # -------------------------------------------------------------------------
    def _init_header_section(self, parent):
        header_card = CTkFrame(parent, fg_color="white", corner_radius=10)
        header_card.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        CTkLabel(header_card, text="🚚 บันทึกค่าขนส่ง (Admin)", font=self.header_font, text_color="#B45309").pack(pady=5)
        
        row_frame = CTkFrame(header_card, fg_color="transparent")
        row_frame.pack(fill="x", padx=10, pady=5)
        
        CTkLabel(row_frame, text="เลขที่ PO:", font=self.label_font).pack(side="left", padx=5)
        self.po_entry = CTkEntry(row_frame, placeholder_text="ระบุ PO...", width=160, font=self.normal_font)
        self.po_entry.pack(side="left", padx=5)
        self.po_entry.bind("<FocusOut>", self._check_po_status)
        self.po_entry.bind("<Return>", self._check_po_status)
        
        self.po_status_label = CTkLabel(row_frame, text="", font=self.status_font)
        self.po_status_label.pack(side="left", padx=5)

    def _init_stock_section(self, parent):
        self.stock_frame = CTkFrame(parent, fg_color="#FEF2F2", border_color="#EF4444", border_width=1)
        self.stock_frame.pack(fill="x", padx=5, pady=5, ipadx=5, ipady=5)
        self.stock_frame.grid_columnconfigure(1, weight=1)
        
        CTkLabel(self.stock_frame, text="1. ค่าจัดส่งเข้าสต๊อก (ค่าย้าย)", font=self.section_font, text_color="#B91C1C").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        self.stock_date_vars = self._create_date_selector(self.stock_frame, row=1, label="วันที่:")
        self.stock_driver = self._create_entry_row(self.stock_frame, 2, "คนขับ/ขนส่ง:")
        self.stock_plate = self._create_entry_row(self.stock_frame, 3, "ทะเบียน:")
        
        CTkLabel(self.stock_frame, text="ยอดเงิน:", font=self.label_font).grid(row=4, column=0, sticky="e", padx=5)
        money_frame = CTkFrame(self.stock_frame, fg_color="transparent")
        money_frame.grid(row=4, column=1, sticky="w", padx=5)
        
        self.stock_cost = CTkEntry(money_frame, width=100, placeholder_text="0.00")
        self.stock_cost.pack(side="left")
        self.stock_cost.bind("<KeyRelease>", self._calculate_totals)
        
        self.stock_vat_var = tk.StringVar(value="No")
        CTkRadioButton(money_frame, text="ไม่มี VAT", variable=self.stock_vat_var, value="No", command=self._calculate_totals).pack(side="left", padx=5)
        CTkRadioButton(money_frame, text="มี VAT 7%", variable=self.stock_vat_var, value="Yes", command=self._calculate_totals).pack(side="left")
        
        CTkLabel(self.stock_frame, text="หัก WHT:", font=self.label_font).grid(row=5, column=0, sticky="e", padx=5)
        self.stock_wht_opt = CTkOptionMenu(self.stock_frame, values=["ไม่หัก (None)", "1% (ค่าขนส่ง)", "3% (ค่าบริการ)"], command=self._calculate_totals)
        self.stock_wht_opt.grid(row=5, column=1, sticky="w", padx=5, pady=2)
        
        self.stock_summary_lbl = CTkLabel(self.stock_frame, text="สุทธิ: 0.00", font=self.result_font, text_color="#B91C1C")
        self.stock_summary_lbl.grid(row=6, column=1, sticky="w", padx=5)
        
        self.stock_remark = self._create_entry_row(self.stock_frame, 7, "หมายเหตุ:")

    def _init_site_section(self, parent):
        self.site_frame = CTkFrame(parent, fg_color="#EFF6FF", border_color="#3B82F6", border_width=1)
        self.site_frame.pack(fill="x", padx=5, pady=10, ipadx=5, ipady=5)
        self.site_frame.grid_columnconfigure(1, weight=1)
        
        CTkLabel(self.site_frame, text="2. ค่าจัดส่งเข้าไซต์ (ค่ารถ)", font=self.section_font, text_color="#1D4ED8").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        self.site_date_vars = self._create_date_selector(self.site_frame, row=1, label="วันที่:")
        self.site_driver = self._create_entry_row(self.site_frame, 2, "คนขับ/ขนส่ง:")
        self.site_plate = self._create_entry_row(self.site_frame, 3, "ทะเบียน:")
        
        CTkLabel(self.site_frame, text="ยอดเงิน:", font=self.label_font).grid(row=4, column=0, sticky="e", padx=5)
        money_frame = CTkFrame(self.site_frame, fg_color="transparent")
        money_frame.grid(row=4, column=1, sticky="w", padx=5)
        
        self.site_cost = CTkEntry(money_frame, width=100, placeholder_text="0.00")
        self.site_cost.pack(side="left")
        self.site_cost.bind("<KeyRelease>", self._calculate_totals)
        
        self.site_vat_var = tk.StringVar(value="No")
        CTkRadioButton(money_frame, text="ไม่มี VAT", variable=self.site_vat_var, value="No", command=self._calculate_totals).pack(side="left", padx=5)
        CTkRadioButton(money_frame, text="มี VAT 7%", variable=self.site_vat_var, value="Yes", command=self._calculate_totals).pack(side="left")
        
        CTkLabel(self.site_frame, text="หัก WHT:", font=self.label_font).grid(row=5, column=0, sticky="e", padx=5)
        self.site_wht_opt = CTkOptionMenu(self.site_frame, values=["ไม่หัก (None)", "1% (ค่าขนส่ง)", "3% (ค่าบริการ)"], command=self._calculate_totals)
        self.site_wht_opt.grid(row=5, column=1, sticky="w", padx=5, pady=2)
        
        self.site_summary_lbl = CTkLabel(self.site_frame, text="สุทธิ: 0.00", font=self.result_font, text_color="#1D4ED8")
        self.site_summary_lbl.grid(row=6, column=1, sticky="w", padx=5)
        
        self.site_remark = self._create_entry_row(self.site_frame, 7, "หมายเหตุ:")

    def _init_footer_buttons(self, parent):
        btn_frame = CTkFrame(parent, fg_color="white")
        btn_frame.grid(row=2, column=0, sticky="ew", pady=5)
        btn_frame.grid_columnconfigure((0,1,2), weight=1)
        
        CTkButton(btn_frame, text="บันทึกรายการ", command=self._save_px, fg_color="#16A34A", font=self.label_font).grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        
        self.delete_btn = CTkButton(btn_frame, text="ลบรายการ", command=self._delete_px, fg_color="#DC2626", state="disabled", font=self.label_font)
        self.delete_btn.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        
        CTkButton(btn_frame, text="ล้างฟอร์ม", command=self._clear_form, fg_color="gray", font=self.label_font).grid(row=0, column=2, padx=5, pady=10, sticky="ew")

    # -------------------------------------------------------------------------
    #  HELPER WIDGETS
    # -------------------------------------------------------------------------
    def _create_entry_row(self, parent, row, label):
        CTkLabel(parent, text=label, font=self.normal_font).grid(row=row, column=0, sticky="e", padx=5, pady=2)
        entry = CTkEntry(parent, font=self.normal_font)
        entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        return entry

    def _create_date_selector(self, parent, row, label):
        CTkLabel(parent, text=label, font=self.normal_font).grid(row=row, column=0, sticky="e", padx=5, pady=2)
        frame = CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=1, sticky="w", padx=5, pady=2)
        
        now = datetime.now()
        days = [str(i).zfill(2) for i in range(1, 32)]
        d_var = tk.StringVar(value=str(now.day).zfill(2))
        m_var = tk.StringVar(value=str(now.month).zfill(2))
        y_var = tk.StringVar(value=str(now.year))
        
        CTkOptionMenu(frame, variable=d_var, values=days, width=60).pack(side="left")
        CTkOptionMenu(frame, variable=m_var, values=[str(i).zfill(2) for i in range(1, 13)], width=60).pack(side="left", padx=2)
        CTkOptionMenu(frame, variable=y_var, values=[str(y) for y in range(now.year-1, now.year+2)], width=70).pack(side="left")
        
        return (d_var, m_var, y_var)

    def _get_date_str(self, date_vars):
        try:
            d, m, y = int(date_vars[0].get()), int(date_vars[1].get()), int(date_vars[2].get())
            # [🔥 แก้ไข] ตรวจสอบวันที่ว่าเป็นวันที่มีอยู่จริงหรือไม่
            datetime(y, m, d) 
            return f"{y}-{m:02d}-{d:02d}"
        except ValueError:
            return None # วันที่ผิด
        except Exception:
            return datetime.now().strftime("%Y-%m-%d")

    def _set_date_vars(self, date_vars, date_str):
        if not date_str: return
        try:
            dt = datetime.strptime(str(date_str), "%Y-%m-%d")
            date_vars[0].set(str(dt.day).zfill(2))
            date_vars[1].set(str(dt.month).zfill(2))
            date_vars[2].set(str(dt.year))
        except: pass

    # -------------------------------------------------------------------------
    #  LOGIC METHODS
    # -------------------------------------------------------------------------
    def _check_po_status(self, event=None):
        po = self.po_entry.get().strip().upper()
        if not po:
            self.po_status_label.configure(text="", text_color="gray")
            return False
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # [🔥 แก้ไข] Query จะทำงานได้ต่อเมื่อมีคอลัมน์ใหม่แล้ว (จากการรัน SQL ด้านบน)
                try:
                    cursor.execute("""
                        SELECT id, 
                               shipping_to_stock_driver, shipping_to_stock_plate, shipping_to_stock_cost,
                               shipping_to_site_driver, shipping_to_site_plate, shipping_to_site_cost
                        FROM purchase_orders WHERE po_number = %s LIMIT 1
                    """, (po,))
                    row = cursor.fetchone()
                    
                    if row:
                        self.po_status_label.configure(text="✅ พบ PO", text_color="#16A34A")
                        
                        # Auto-Fill Logic (เติมถ้าช่องยังว่าง)
                        if not self.stock_driver.get(): self.stock_driver.insert(0, row[1] or "")
                        if not self.stock_plate.get(): self.stock_plate.insert(0, row[2] or "")
                        
                        if not self.site_driver.get(): self.site_driver.insert(0, row[4] or "")
                        if not self.site_plate.get(): self.site_plate.insert(0, row[5] or "")
                        
                        self._calculate_totals()
                        return True
                    else:
                        self.po_status_label.configure(text="⚠️ ไม่พบ", text_color="#D97706")
                        return False
                except Exception as db_err:
                    print(f"DB Error checking PO columns: {db_err}")
                    return False
        except Exception as e:
            print(f"Error checking PO: {e}")
            return False
        finally:
            self.app_container.release_connection(conn)

    def _calculate_totals(self, event=None):
        # 1. Stock Calc
        try:
            s_cost = float(self.stock_cost.get().replace(",","") or 0)
            s_vat = s_cost * 0.07 if self.stock_vat_var.get() == "Yes" else 0
            s_wht_opt = self.stock_wht_opt.get()
            s_wht_p = 1.0 if "1%" in s_wht_opt else (3.0 if "3%" in s_wht_opt else 0.0)
            s_wht = s_cost * (s_wht_p / 100)
            s_net = s_cost + s_vat - s_wht
            self.stock_summary_lbl.configure(text=f"V:{s_vat:.2f} | W:{s_wht:.2f} | สุทธิ: {s_net:,.2f}")
        except: self.stock_summary_lbl.configure(text="Error")

        # 2. Site Calc
        try:
            t_cost = float(self.site_cost.get().replace(",","") or 0)
            t_vat = t_cost * 0.07 if self.site_vat_var.get() == "Yes" else 0
            t_wht_opt = self.site_wht_opt.get()
            t_wht_p = 1.0 if "1%" in t_wht_opt else (3.0 if "3%" in t_wht_opt else 0.0)
            t_wht = t_cost * (t_wht_p / 100)
            t_net = t_cost + t_vat - t_wht
            self.site_summary_lbl.configure(text=f"V:{t_vat:.2f} | W:{t_wht:.2f} | สุทธิ: {t_net:,.2f}")
        except: self.site_summary_lbl.configure(text="Error")

    # [🔥 แก้ไข] รับ cursor เข้ามาเพื่อใช้ transaction เดียวกันในการนับ
    def _generate_px_number(self, po_number, cursor=None):
        po_clean = po_number.strip().upper()
        base = po_clean.replace("PO", "PX", 1) if po_clean.startswith("PO") else f"PX-{po_clean}"
        
        # [🔥 แก้ไข] Logic การ Gen ID ให้ไม่ชนกันเมื่อ save พร้อมกันหลายรายการ
        # ให้ใช้ logic ว่าถ้าใน Transaction นี้มี PX นี้อยู่กี่อันแล้ว + DB มีกี่อัน
        
        try:
            query = "SELECT COUNT(*) FROM transport_orders WHERE px_number LIKE %s"
            param = (f"{base}%",)
            
            count = 0
            if cursor:
                cursor.execute(query, param)
                count = cursor.fetchone()[0]
            else:
                # Fallback (ไม่ควรเข้าเคสนี้บ่อย)
                conn = self.app_container.get_connection()
                with conn.cursor() as cur:
                    cur.execute(query, param)
                    count = cur.fetchone()[0]
                self.app_container.release_connection(conn)
            
            # รันเลขต่อท้ายเสมอ (-1, -2, ...) เพื่อป้องกันการชนกับ Base และกันสับสน
            return f"{base}-{count + 1}"

        except Exception as e:
            print(f"Gen PX Error: {e}")
            return f"{base}-{datetime.now().strftime('%M%S')}"

    def _save_single_record(self, cursor, po, t_type, date_vars, driver_entry, plate_entry, cost_entry, vat_var, wht_opt, remark_entry):
        try:
            cost = float(cost_entry.get().replace(",", "") or 0)
        except: cost = 0
        
        if cost <= 0: return False # ไม่บันทึกถ้ายอดเป็น 0

        date_val = self._get_date_str(date_vars)
        if date_val is None:
            messagebox.showerror("วันที่ผิดพลาด", f"วันที่ระบุไม่ถูกต้อง (เช่น 31 ก.พ.) กรุณาตรวจสอบวันที่ของรายการ '{t_type}'", parent=self)
            raise ValueError("Invalid Date")

        vat_amt = cost * 0.07 if vat_var.get() == "Yes" else 0
        wht_p = 1.0 if "1%" in wht_opt.get() else (3.0 if "3%" in wht_opt.get() else 0.0)
        wht_amt = cost * (wht_p / 100)
        net = cost + vat_amt - wht_amt
        
        # [🔥 แก้ไข] ส่ง cursor ไปด้วย เพื่อให้นับรวมรายการที่เพิ่ง insert ใน transaction นี้
        px_no = self._generate_px_number(po, cursor) 
        
        sql = """
            INSERT INTO transport_orders 
            (px_number, ref_po_number, transport_date, transporter_name, license_plate, 
             transport_cost, vat_amount, wht_percent, wht_amount, net_amount, remarks,
             payment_type, status, created_by, transport_type)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(sql, (
            px_no, po, date_val, driver_entry.get().strip(), plate_entry.get().strip(),
            cost, vat_amt, wht_p, wht_amt, net, remark_entry.get().strip(),
            "Credit", "Matched", self.user_key, t_type
        ))
        return True

    def _save_px(self):
        po = self.po_entry.get().strip().upper()
        if not po:
            messagebox.showwarning("เตือน", "กรุณาระบุเลข PO", parent=self)
            return

        conn = self.app_container.get_connection()
        try:
            saved_count = 0
            with conn.cursor() as cursor:
                # 1. ลองบันทึกส่วน Stock
                if self._save_single_record(cursor, po, "Stock", self.stock_date_vars, self.stock_driver, self.stock_plate, self.stock_cost, self.stock_vat_var, self.stock_wht_opt, self.stock_remark):
                    saved_count += 1
                
                # 2. ลองบันทึกส่วน Site
                if self._save_single_record(cursor, po, "Site", self.site_date_vars, self.site_driver, self.site_plate, self.site_cost, self.site_vat_var, self.site_wht_opt, self.site_remark):
                    saved_count += 1
            
            if saved_count > 0:
                conn.commit()
                messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลเรียบร้อยจำนวน {saved_count} รายการ", parent=self)
                self._clear_form()
                self._load_history_data()
                self._load_pending_data()
            else:
                messagebox.showwarning("ไม่ได้บันทึก", "กรุณาระบุยอดเงินอย่างน้อย 1 รายการ (Stock หรือ Site)", parent=self)

        except ValueError:
            pass # Error วันที่แจ้งเตือนไปแล้ว
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Error", f"บันทึกไม่สำเร็จ: {e}", parent=self)
        finally:
            self.app_container.release_connection(conn)

    def _delete_px(self):
        if not self.current_selected_id: return
        if not messagebox.askyesno("ยืนยัน", f"ต้องการลบรายการ {self.current_selected_px} หรือไม่?"): return
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("DELETE FROM transport_orders WHERE id = %s", (self.current_selected_id,))
            conn.commit()
            messagebox.showinfo("สำเร็จ", "ลบรายการเรียบร้อย", parent=self)
            self._clear_form()
            self._load_history_data()
            self._load_pending_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", f"{e}")
        finally:
            self.app_container.release_connection(conn)

    def _clear_form(self):
        self.po_entry.delete(0, "end")
        self.po_status_label.configure(text="")
        
        # Clear Stock
        self.stock_driver.delete(0, "end"); self.stock_plate.delete(0, "end")
        self.stock_cost.delete(0, "end"); self.stock_remark.delete(0, "end")
        self.stock_vat_var.set("No"); self.stock_wht_opt.set("ไม่หัก (None)")
        self.stock_summary_lbl.configure(text="สุทธิ: 0.00")
        
        # Clear Site
        self.site_driver.delete(0, "end"); self.site_plate.delete(0, "end")
        self.site_cost.delete(0, "end"); self.site_remark.delete(0, "end")
        self.site_vat_var.set("No"); self.site_wht_opt.set("ไม่หัก (None)")
        self.site_summary_lbl.configure(text="สุทธิ: 0.00")
        
        self.current_selected_id = None
        self.delete_btn.configure(state="disabled")

    # -------------------------------------------------------------------------
    #  TAB LOADER METHODS (เหมือนเดิมแต่ปรับ Query นิดหน่อย)
    # -------------------------------------------------------------------------
    def _create_treeview(self, parent, columns):
        style = ttk.Style(); style.theme_use("clam")
        style.configure("Treeview.Heading", font=('TH Sarabun New', 14, 'bold'))
        style.configure("Treeview", font=('TH Sarabun New', 12), rowheight=30)
        tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        
        headers = {
            "date": "วันที่", "type": "ประเภท", "px_no": "เลขที่ PX", "po_no": "เลขที่ PO", 
            "transporter": "ขนส่ง", "cost": "ยอดสุทธิ", "status": "สถานะ", 
            "supplier": "ซัพพลายเออร์", "amount": "ยอดเงิน PO"
        }
        for col in columns:
            tree.heading(col, text=headers.get(col, col))
            w = 80 if col == "type" else 100
            tree.column(col, width=w, anchor="center" if col not in ["transporter", "supplier"] else "w")
            
        scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True, padx=(5,0), pady=5)
        scroll.pack(side="right", fill="y", padx=(0,5), pady=5)
        return tree

    def _load_history_data(self):
        search_text = self.hist_search.get().strip().upper() if hasattr(self, 'hist_search') else ""
        for item in self.tree_hist.get_children(): self.tree_hist.delete(item)
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                query = """
                    SELECT id, TO_CHAR(transport_date, 'YYYY-MM-DD'), px_number, ref_po_number, 
                           transporter_name, net_amount, status,
                           license_plate, transport_cost, wht_percent, remarks, payment_type,
                           transport_type, vat_amount
                    FROM transport_orders WHERE 1=1 
                """
                params = []
                if search_text:
                    query += " AND (ref_po_number LIKE %s OR px_number LIKE %s)"
                    params.extend([f"%{search_text}%", f"%{search_text}%"])
                query += " ORDER BY id DESC LIMIT 50"
                cursor.execute(query, params)
                self.current_hist_data = cursor.fetchall()
                
                for row in self.current_hist_data:
                    cost_fmt = f"{row[5]:,.2f}" if row[5] else "0.00"
                    t_type = row[12] if row[12] else "-"
                    self.tree_hist.insert("", "end", values=(row[1], t_type, row[2], row[3], row[4], cost_fmt, row[6]))
        except Exception as e: print(e)
        finally: self.app_container.release_connection(conn)

    def _on_hist_row_click(self, event):
        sel = self.tree_hist.selection()
        if not sel: return
        idx = self.tree_hist.index(sel)
        if idx < len(self.current_hist_data):
            row = self.current_hist_data[idx]
            # row: 0=id, 1=date, 2=px, 3=po, 4=driver, 5=net, 6=status, 7=plate, 8=cost, 9=wht, 10=remark, 11=pay, 12=type, 13=vat
            self._clear_form()
            
            # Populate Only Specific Section based on Type
            self.po_entry.insert(0, row[3])
            t_type = row[12]
            
            target_driver = self.stock_driver if t_type == "Stock" else self.site_driver
            target_plate = self.stock_plate if t_type == "Stock" else self.site_plate
            target_cost = self.stock_cost if t_type == "Stock" else self.site_cost
            target_vat = self.stock_vat_var if t_type == "Stock" else self.site_vat_var
            target_wht = self.stock_wht_opt if t_type == "Stock" else self.site_wht_opt
            target_remark = self.stock_remark if t_type == "Stock" else self.site_remark
            target_date = self.stock_date_vars if t_type == "Stock" else self.site_date_vars
            
            target_driver.insert(0, row[4] or "")
            target_plate.insert(0, row[7] or "")
            target_cost.insert(0, f"{row[8]:.2f}" if row[8] else "0.00")
            target_remark.insert(0, row[10] or "")
            
            if row[13] and float(row[13]) > 0: target_vat.set("Yes")
            else: target_vat.set("No")
            
            if row[9] == 1.0: target_wht.set("1% (ค่าขนส่ง)")
            elif row[9] == 3.0: target_wht.set("3% (ค่าบริการ)")
            else: target_wht.set("ไม่หัก (None)")
            
            self._set_date_vars(target_date, row[1])
            self._calculate_totals()
            
            self.current_selected_id = row[0]
            self.current_selected_px = row[2]
            self.delete_btn.configure(state="normal")

    def _setup_history_tab(self):
        search_frame = CTkFrame(self.tab_history, fg_color="transparent"); search_frame.pack(fill="x", padx=10, pady=5)
        self.hist_search = CTkEntry(search_frame, placeholder_text="ค้นหา PX หรือ PO...", width=200); self.hist_search.pack(side="left", padx=5)
        self.hist_search.bind("<Return>", lambda e: self._load_history_data())
        CTkButton(search_frame, text="ค้นหา", command=self._load_history_data, width=80).pack(side="left")
        
        self.tree_hist = self._create_treeview(self.tab_history, ["date", "type", "px_no", "po_no", "transporter", "cost", "status"])
        self.tree_hist.bind("<Double-1>", self._on_hist_row_click)

    def _setup_pending_tab(self):
        search_frame = CTkFrame(self.tab_pending, fg_color="transparent"); search_frame.pack(fill="x", padx=10, pady=5)
        self.pending_search = CTkEntry(search_frame, placeholder_text="ค้นหาเลข PO...", width=200); self.pending_search.pack(side="left", padx=5)
        self.pending_search.bind("<Return>", lambda e: self._load_pending_data())
        CTkButton(search_frame, text="ค้นหา", command=self._load_pending_data, width=80).pack(side="left")
        
        self.tree_pending = self._create_treeview(self.tab_pending, ["date", "po_no", "supplier", "amount"])
        self.tree_pending.bind("<Double-1>", self._on_pending_row_click)

    def _load_pending_data(self):
        search_text = self.pending_search.get().strip().upper()
        for item in self.tree_pending.get_children(): self.tree_pending.delete(item)
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                query = "SELECT TO_CHAR(CAST(timestamp AS TIMESTAMP), 'YYYY-MM-DD'), po_number, supplier_name, grand_total FROM purchase_orders po WHERE status != 'Cancelled'"
                if search_text: query += f" AND po_number LIKE '%{search_text}%'"
                query += " ORDER BY id DESC LIMIT 50"
                cursor.execute(query)
                for row in cursor.fetchall():
                    self.tree_pending.insert("", "end", values=(row[0], row[1], row[2], f"{row[3]:,.2f}"))
        except: pass
        finally: self.app_container.release_connection(conn)

    def _on_pending_row_click(self, event):
        sel = self.tree_pending.selection()
        if not sel: return
        self._clear_form()
        vals = self.tree_pending.item(sel)['values']
        self.po_entry.insert(0, vals[1])
        self._check_po_status()

    def _logout(self):
        self.app_container.show_login_screen()