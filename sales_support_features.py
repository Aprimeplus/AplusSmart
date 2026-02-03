import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import pandas as pd
import psycopg2.extras
from datetime import datetime
import utils
from custom_widgets import NumericEntry

class SalesSupportOutstandingManager(ctk.CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        
        # --- UI Layout ---
        self._setup_layout()
        self._load_data()

    def _setup_layout(self):
        # 1. Header & Filter Area
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(header_frame, text="ติดตามและบันทึกยอดค้างชำระ (Sales Support)", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")

        # Filter Frame
        filter_frame = ctk.CTkFrame(self)
        filter_frame.pack(fill="x", padx=10, pady=(0, 10))

        # Dropdown เลือก Sales (ดูของคนอื่นได้)
        ctk.CTkLabel(filter_frame, text="เลือก Sales:").pack(side="left", padx=10, pady=10)
        self.sale_filter_var = tk.StringVar(value="All")
        self.sale_combo = ctk.CTkOptionMenu(filter_frame, variable=self.sale_filter_var, 
                                            values=["All"], command=self._filter_data, width=150)
        self.sale_combo.pack(side="left", padx=5)

        # ช่องค้นหา
        ctk.CTkLabel(filter_frame, text="ค้นหา (SO/ลูกค้า):").pack(side="left", padx=10)
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(filter_frame, textvariable=self.search_var, width=200)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", lambda e: self._filter_data())
        
        ctk.CTkButton(filter_frame, text="🔍 ค้นหา", width=80, command=self._filter_data).pack(side="left", padx=5)
        ctk.CTkButton(filter_frame, text="🔄 รีเฟรช", width=80, fg_color="gray", command=self._load_data).pack(side="left", padx=5)

        # 2. Table Area
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        columns = ["ID", "SO Number", "Sales", "ลูกค้า", "ยอดเต็ม (VAT)", "จ่ายแล้ว", "คงเหลือ (ค้าง)", "Due Date", "สถานะ"]
        
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", height=20)
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Scrollbar
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        # Config Columns
        col_widths = [0, 120, 100, 200, 120, 120, 120, 100, 100] # ID hidden
        for i, col in enumerate(columns):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths[i], anchor="center" if i != 3 else "w") # ชื่อลูกค้าชิดซ้าย
            
        self.tree.column("ID", width=0, stretch=False) # Hide ID

        # Style Configuration
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Sarabun', 11, 'bold'))
        style.configure("Treeview", rowheight=30, font=('Sarabun', 10))
        
        # Bind Double Click to Update
        self.tree.bind("<Double-1>", self._open_payment_update_popup)

        # 3. Footer / Action
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(footer_frame, text="* ดับเบิลคลิกที่รายการเพื่อ 'หยอด' ยอดชำระเพิ่ม", text_color="gray").pack(side="left")

    def _load_data(self):
        # 1. โหลดรายชื่อ Sales มาใส่ Dropdown ก่อน
        try:
            sales_df = pd.read_sql("SELECT DISTINCT sale_key FROM sales_users ORDER BY sale_key", self.pg_engine)
            sales_list = ["All"] + sales_df['sale_key'].tolist()
            self.sale_combo.configure(values=sales_list)
        except: pass

        # 2. โหลดข้อมูล Outstanding
        self._filter_data()

    def _filter_data(self, *args):
        # Clear Tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        selected_sale = self.sale_filter_var.get()
        search_txt = self.search_var.get().strip().lower()

        # Query: ดึงเฉพาะที่ยังจ่ายไม่ครบ (difference_amount > 1 บาท เผื่อเศษทศนิยม)
        # และยังไม่ถูก Cancel
        sql = """
            SELECT 
                id, so_number, sale_key, customer_name,
                -- คำนวณ Grand Total (รวม VAT) เพื่อแสดงผล
                (sales_service_amount + cutting_drilling_fee + other_service_fee + shipping_cost + 
                 COALESCE(relocation_cost,0) + COALESCE(credit_card_fee,0) - COALESCE(coupons,0)) * (CASE WHEN sales_service_vat_option = 'VAT' THEN 1.07 ELSE 1.0 END) as grand_total_est,
                 
                total_payment_amount,
                difference_amount,
                payment_date, -- หรือ Due Date ถ้ามี
                status
            FROM commissions
            WHERE difference_amount > 1 
              AND status != 'Cancelled'
        """
        
        params = []
        if selected_sale != "All":
            sql += " AND sale_key = %s"
            params.append(selected_sale)
            
        if search_txt:
            sql += " AND (LOWER(so_number) LIKE %s OR LOWER(customer_name) LIKE %s)"
            params.append(f"%{search_txt}%")
            params.append(f"%{search_txt}%")
            
        sql += " ORDER BY difference_amount DESC" # เอาที่ค้างเยอะๆ ขึ้นก่อน

        try:
            df = pd.read_sql_query(sql, self.pg_engine, params=tuple(params))
            
            for _, row in df.iterrows():
                # จัดการสี (ค้างเยอะ = แดง, ค้างน้อย = ส้ม)
                diff = row['difference_amount']
                tag = 'high_debt' if diff > 10000 else 'normal_debt'
                
                # Format วันที่
                date_str = "-"
                if pd.notna(row['payment_date']):
                    date_str = str(row['payment_date'])[:10]

                vals = (
                    row['id'],
                    row['so_number'],
                    row['sale_key'],
                    row['customer_name'],
                    f"{row['grand_total_est']:,.2f}",
                    f"{row['total_payment_amount']:,.2f}",
                    f"{row['difference_amount']:,.2f}", # ยอดที่ Sale Support สนใจ
                    date_str,
                    row['status']
                )
                self.tree.insert("", "end", values=vals, tags=(tag,))
                
            self.tree.tag_configure('high_debt', foreground='#DC2626') # แดงเข้ม
            self.tree.tag_configure('normal_debt', foreground='#D97706') # ส้ม

        except Exception as e:
            print(f"Error loading outstanding: {e}")
            messagebox.showerror("Error", f"โหลดข้อมูลไม่สำเร็จ: {e}")

    def _open_payment_update_popup(self, event):
        """หน้าต่าง Popup สำหรับหยอดยอดเงิน"""
        selected_item = self.tree.focus()
        if not selected_item: return
        
        values = self.tree.item(selected_item, "values")
        comm_id = values[0]
        so_number = values[1]
        current_paid = float(values[5].replace(",",""))
        current_diff = float(values[6].replace(",",""))
        
        # --- Create Popup ---
        popup = ctk.CTkToplevel(self)
        popup.title(f"อัปเดตการชำระเงิน: {so_number}")
        popup.geometry("400x450")
        popup.transient(self)
        popup.grab_set()
        
        # Info
        ctk.CTkLabel(popup, text=f"SO Number: {so_number}", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(20,5))
        ctk.CTkLabel(popup, text=f"ลูกค้า: {values[3]}", font=ctk.CTkFont(size=14)).pack(pady=5)
        
        info_frame = ctk.CTkFrame(popup, fg_color="#F3F4F6")
        info_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(info_frame, text=f"ยอดค้างชำระปัจจุบัน:", text_color="#DC2626").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(info_frame, text=f"{current_diff:,.2f}", font=ctk.CTkFont(weight="bold"), text_color="#DC2626").grid(row=0, column=1, padx=10, pady=5, sticky="e")
        
        ctk.CTkLabel(info_frame, text=f"จ่ายแล้วสะสม:", text_color="gray").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(info_frame, text=f"{current_paid:,.2f}", text_color="gray").grid(row=1, column=1, padx=10, pady=5, sticky="e")

        # Input Area
        ctk.CTkLabel(popup, text="ระบุยอดที่ลูกค้าโอนมาเพิ่ม (Top-up):", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 5))
        
        topup_entry = NumericEntry(popup, placeholder_text="ระบุจำนวนเงิน...")
        topup_entry.pack(pady=5, padx=20, fill="x")
        
        # Checkbox: จ่ายครบแล้ว?
        is_fully_paid = ctk.CTkCheckBox(popup, text="เคลียร์ยอดทั้งหมด (จ่ายครบแล้ว)")
        is_fully_paid.pack(pady=10)
        
        def on_full_paid_toggle():
            if is_fully_paid.get():
                topup_entry.delete(0, "end")
                topup_entry.insert(0, f"{current_diff:.2f}") # ใส่ยอดค้างทั้งหมดให้อัตโนมัติ
        
        is_fully_paid.configure(command=on_full_paid_toggle)

        # Save Action
        def save_payment():
            try:
                top_up_amount = utils.convert_to_float(topup_entry.get())
                if top_up_amount <= 0:
                    messagebox.showwarning("เตือน", "กรุณาระบุยอดเงินที่มากกว่า 0", parent=popup)
                    return
                
                # Logic: อัปเดตยอดรวม
                new_total_paid = current_paid + top_up_amount
                new_diff = current_diff - top_up_amount
                
                if new_diff < -1: # ยอมให้เกินได้นิดหน่อยเรื่องทศนิยม
                    if not messagebox.askyesno("ยอดเกิน", f"ยอดโอนรวม ({new_total_paid:,.2f}) มากกว่ายอดเต็ม\nต้องการบันทึกหรือไม่?"):
                        return

                conn = self.app_container.get_connection()
                with conn.cursor() as cursor:
                    # 1. Update Commissions
                    cursor.execute("""
                        UPDATE commissions 
                        SET total_payment_amount = %s,
                            difference_amount = %s,
                            payment_date = %s
                        WHERE id = %s
                    """, (new_total_paid, new_diff, datetime.now(), comm_id))
                    
                    # 2. Audit Log (สำคัญมาก! ต้องรู้ว่า Sales Support คนไหนแก้)
                    log_msg = f"Support Top-up: {top_up_amount:,.2f} | New Paid: {new_total_paid:,.2f} | New Diff: {new_diff:,.2f}"
                    cursor.execute("""
                        INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, ('Support Update Payment', 'commissions', comm_id, 
                          self.app_container.current_user_key, log_msg, datetime.now()))
                
                conn.commit()
                messagebox.showinfo("สำเร็จ", "บันทึกยอดชำระเพิ่มเติมเรียบร้อยแล้ว", parent=popup)
                popup.destroy()
                self._load_data() # Refresh ตารางหลัก
                
            except Exception as e:
                messagebox.showerror("Error", f"บันทึกไม่สำเร็จ: {e}", parent=popup)
            finally:
                if conn: self.app_container.release_connection(conn)

        ctk.CTkButton(popup, text="💾 บันทึกยอด", fg_color="#16A34A", hover_color="#15803D", 
                      height=40, command=save_payment).pack(pady=20, padx=20, fill="x")