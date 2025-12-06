import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from customtkinter import CTkFrame, CTkLabel, CTkButton, CTkEntry, CTkFont, CTkScrollableFrame
import pandas as pd
from datetime import datetime
import psycopg2
from custom_widgets import DateSelector

class DailyReportWidget(CTkFrame):
    def __init__(self, master, app_container, **kwargs):
        super().__init__(master, **kwargs)
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        self.current_df = None
        
        # --- 1. Top Bar (Filter) ---
        self.top_frame = CTkFrame(self, fg_color="transparent")
        self.top_frame.pack(fill="x", padx=15, pady=(15, 10))
        
        CTkLabel(self.top_frame, text="รายงานประจำวัน:", font=CTkFont(size=16, weight="bold"), text_color="#374151").pack(side="left", padx=(0, 10))
        
        self.date_selector = DateSelector(self.top_frame)
        self.date_selector.pack(side="left")
        self.date_selector.set_date(datetime.now()) 
        
        self.btn_refresh = CTkButton(self.top_frame, text="🔄 ดึงข้อมูล", width=100, fg_color="#3B82F6", hover_color="#2563EB", command=self.load_report_data)
        self.btn_refresh.pack(side="left", padx=10)
        
        self.btn_export = CTkButton(self.top_frame, text="📂 Export Excel", width=100, fg_color="#10B981", hover_color="#059669", command=self.export_to_excel)
        self.btn_export.pack(side="left", padx=10)

        # --- 2. Table Area (Treeview) ---
        self.table_frame = CTkFrame(self, fg_color="white", corner_radius=0)
        self.table_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", background="#E0F2FE", foreground="#0F172A", font=('Prompt', 10, 'bold'), relief="flat")
        style.configure("Treeview", background="white", foreground="#334155", rowheight=35, fieldbackground="white", font=('Arial', 10))
        style.map("Treeview", background=[('selected', '#3B82F6')], foreground=[('selected', 'white')])

        self.tree_scroll_y = ttk.Scrollbar(self.table_frame, orient="vertical")
        self.tree_scroll_y.pack(side="right", fill="y")
        self.tree_scroll_x = ttk.Scrollbar(self.table_frame, orient="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x")
        
        self.columns = [
            "so_number", "po_number", "customer_name", "sales_booking", 
            "total_paid", "credit_balance", "payment_status", "delivery_date", "location",  # ย้าย status มาใกล้ๆ ยอดเงิน
            "services", "prepared_by", "pu_prepared_by", "status"
        ]
        self.tree = ttk.Treeview(self.table_frame, columns=self.columns, show="headings", 
                                 yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)
        
        headings = {
            "so_number": ("SO Number", 90, "center"),
            "po_number": ("PO Number", 90, "center"),
            "customer_name": ("ชื่อลูกค้า", 160, "w"),
            "sales_booking": ("ยอดจอง", 80, "e"),
            "total_paid": ("ชำระแล้ว", 80, "e"),
            "credit_balance": ("ยอดค้าง", 80, "e"),
            "payment_status": ("ตรวจสอบยอด", 100, "center"),
            "delivery_date": ("วันที่ส่ง", 80, "center"),
            "location": ("สถานที่ส่ง", 140, "w"),
            "services": ("ค่าบริการ", 70, "e"),
            "prepared_by": ("Sale", 70, "center"),
            "pu_prepared_by": ("PU", 70, "center"),
            "status": ("สถานะ", 80, "center")
        }
        
        for col, (text, width, anchor) in headings.items():
            self.tree.heading(col, text=text, command=lambda c=col: self.sort_treeview(c, False))
            self.tree.column(col, width=width, anchor=anchor)
        
        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background="#F8FAFC")
        self.tree.tag_configure('status_missing', foreground="#DC2626") 
        self.tree.tag_configure('status_over', foreground="#2563EB")    
        self.tree.tag_configure('status_ok', foreground="#059669")      
        
        self.tree.pack(fill="both", expand=True)
        
        # --- 3. Summary Footer ---
        self.footer_frame = CTkFrame(self, height=50, fg_color="white", border_width=1, border_color="#E5E7EB")
        self.footer_frame.pack(fill="x", padx=15, pady=10)
        
        self.lbl_total_so = CTkLabel(self.footer_frame, text="จำนวน SO: 0", font=CTkFont(size=14, weight="bold"), text_color="#64748B")
        self.lbl_total_so.pack(side="left", padx=20)
        
        self.lbl_total_missing = CTkLabel(self.footer_frame, text="ยอดค้างชำระรวม: 0.00", font=CTkFont(size=14, weight="bold"), text_color="#DC2626")
        self.lbl_total_missing.pack(side="right", padx=15)

        self.lbl_total_paid = CTkLabel(self.footer_frame, text="ยอดชำระแล้ว: 0.00", font=CTkFont(size=14, weight="bold"), text_color="#16A34A")
        self.lbl_total_paid.pack(side="right", padx=15)

        self.lbl_total_booking = CTkLabel(self.footer_frame, text="ยอดจองรวม: 0.00", font=CTkFont(size=14, weight="bold"), text_color="#2563EB")
        self.lbl_total_booking.pack(side="right", padx=15)

    def load_report_data(self):
        selected_date = self.date_selector.get_date()
        if not selected_date: return

        date_str = selected_date
        print(f"\n{'='*20} START DEBUG: {date_str} {'='*20}") 
        
        for i in self.tree.get_children(): self.tree.delete(i)
            
        try:
            query = """
                SELECT 
                    c.so_number, 
                    (SELECT STRING_AGG(po.po_number, ', ') FROM purchase_orders po WHERE po.so_number = c.so_number) as po_number_list,
                    c.customer_name, 
                    c.sales_service_amount, 
                    c.total_payment_amount, 
                    c.difference_amount,
                    c.status,
                    COALESCE(c.date_to_customer, c.delivery_date) as final_delivery_date, 
                    c.pickup_location, 
                    (COALESCE(c.cutting_drilling_fee, 0) + COALESCE(c.other_service_fee, 0)) as service_total,
                    c.sale_key,
                    c.user_key,
                    u.sale_name as sale_name,
                    pu.sale_name as pu_name_comm,
                    (
                        SELECT u2.sale_name 
                        FROM purchase_orders po 
                        JOIN sales_users u2 ON po.user_key = u2.sale_key 
                        WHERE po.so_number = c.so_number 
                        LIMIT 1
                    ) as pu_name_po
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                LEFT JOIN sales_users pu ON c.user_key = pu.sale_key
                WHERE date(c.timestamp) = %s AND c.is_active = 1
                ORDER BY c.so_number ASC
            """
            
            df = pd.read_sql_query(query, self.pg_engine, params=(date_str,))
            print(f"[DEBUG] พบข้อมูลทั้งหมด: {len(df)} แถว") 

            if not df.empty:
                # --- Logic รวมชื่อ PU ---
                df['final_pu_name'] = df['pu_name_comm'].fillna(df['pu_name_po'])
                df['final_pu_name'] = df['final_pu_name'].fillna(df['user_key'].apply(lambda x: f"ID:{x}" if pd.notna(x) else "-"))
                df['final_pu_name'] = df['final_pu_name'].fillna("-")
            
            self.current_df = df

            if df.empty:
                messagebox.showinfo("แจ้งเตือน", f"ไม่พบข้อมูล SO ของวันที่ {date_str}")
                self.update_summary(0, 0, 0, 0)
                return

            sum_booking = 0; sum_paid = 0; sum_missing = 0

            for i, row in df.iterrows():
                booking = float(row.get('sales_service_amount') or 0)
                paid = float(row.get('total_payment_amount') or 0)
                
                diff_db = float(row.get('difference_amount') or 0)
                
                status_text = ""
                status_tag = ""
                credit_display = 0.0 # ตัวแปรนี้คือพระเอกของเรา
                
                if abs(diff_db) < 1.00: 
                    status_text = "✅ ครบถ้วน"
                    status_tag = "status_ok"
                    credit_display = 0.0
                elif diff_db < 0: 
                    missing = abs(diff_db)
                    status_text = f"❌ ขาด {missing:,.2f}"
                    status_tag = "status_missing"
                    sum_missing += missing
                    credit_display = missing 
                else: 
                    over = diff_db
                    status_text = f"⚠️ เกิน {over:,.2f}"
                    status_tag = "status_over"
                    credit_display = 0.0 

                service_fee = float(row.get('service_total') or 0)
                d_date = row.get('final_delivery_date'); d_date_str = str(d_date) if d_date else "-"
                
                sum_booking += booking; sum_paid += paid

                pu_show = row['final_pu_name']

                vals = (
                    row['so_number'],
                    row['po_number_list'] or "-",
                    row['customer_name'],
                    f"{booking:,.2f}",
                    f"{paid:,.2f}",
                    f"{credit_display:,.2f}", # <--- แก้ไขจุดนี้: เปลี่ยนจาก credit เป็น credit_display
                    status_text,
                    d_date_str,
                    row['pickup_location'] or "-",
                    f"{service_fee:,.2f}",
                    row.get('sale_name') or row.get('sale_key'),
                    pu_show, 
                    row['status']
                )
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert("", "end", values=vals, tags=(tag, status_tag))

            self.update_summary(len(df), sum_booking, sum_paid, sum_missing)
            print(f"{'='*20} END DEBUG {'='*20}\n")

        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถดึงข้อมูลได้: {e}")
            print(f"!!! CRITICAL ERROR !!! : {e}")
            import traceback
            traceback.print_exc()

    def update_summary(self, count, booking, paid, missing):
        self.lbl_total_so.configure(text=f"จำนวน SO: {count} ใบ")
        self.lbl_total_booking.configure(text=f"ยอดจองรวม: {booking:,.2f}")
        self.lbl_total_paid.configure(text=f"ยอดชำระแล้ว: {paid:,.2f}")
        self.lbl_total_missing.configure(text=f"ยอดค้างชำระรวม: {missing:,.2f}")

    def sort_treeview(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        try: l.sort(key=lambda t: float(t[0].replace(',', '').replace('❌ ขาด ', '').replace('💰 เกิน ', '').replace('✅ ', '')), reverse=reverse)
        except ValueError: l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
            current_tags = self.tree.item(k)['tags']
            status_tag = next((t for t in current_tags if t.startswith('status_')), '')
            stripe_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            new_tags = (stripe_tag, status_tag) if status_tag else (stripe_tag,)
            self.tree.item(k, tags=new_tags)
        self.tree.heading(col, command=lambda: self.sort_treeview(col, not reverse))

    def export_to_excel(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "กรุณาดึงข้อมูลก่อนทำการ Export")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path: return

        try:
            # สร้าง DataFrame ใหม่เพื่อไม่ให้กระทบอันเก่า
            export_df = self.current_df.copy()
            
            # แปลงข้อมูลตัวเลขให้พร้อมคำนวณ (จัดการ NaN)
            export_df['sales_service_amount'] = export_df['sales_service_amount'].fillna(0)
            export_df['total_payment_amount'] = export_df['total_payment_amount'].fillna(0)
            export_df['difference_amount'] = export_df['difference_amount'].fillna(0)
            
            # 1. สร้างคอลัมน์สถานะตรวจสอบยอด (Payment Status)
            def get_status_text(row):
                d = row['difference_amount']
                if abs(d) < 1: return "ครบถ้วน"
                elif d < 0: return f"ขาด {abs(d):,.2f}"
                else: return f"เกิน {d:,.2f}"
            
            export_df['payment_check_status'] = export_df.apply(get_status_text, axis=1)
            
            # 2. สร้างคอลัมน์ยอดค้างชำระ (Credit Balance)
            # โชว์เฉพาะยอดที่ขาด (ถ้าเกิน ให้เป็น 0)
            def get_credit_balance(row):
                d = row['difference_amount']
                return abs(d) if d < 0 else 0
            
            export_df['remaining_balance'] = export_df.apply(get_credit_balance, axis=1)
            
            # 3. กำหนดชื่อคอลัมน์ใน Excel (Mapping)
            rename_map = {
                "so_number": "SO Number",
                "po_number_list": "PO Number",
                "customer_name": "Customer Name",
                "sales_service_amount": "Booking Amount",
                "total_payment_amount": "Paid Amount",
                "remaining_balance": "Credit Balance",     # ยอดค้าง
                "payment_check_status": "Payment Status",  # สถานะตรวจสอบ
                "status": "System Status",
                "final_delivery_date": "Delivery Date",
                "pickup_location": "Location",
                "service_total": "Service Fee",
                "sale_name": "Sale Prepared By",
                "final_pu_name": "PU Prepared By"          # ผู้จัดทำ PU
            }
            
            # เลือกเฉพาะคอลัมน์ที่มีใน Map และเรียงลำดับ
            # (ตรวจสอบก่อนว่าคอลัมน์มีอยู่จริง เพื่อป้องกัน Error)
            cols_to_export = [k for k in rename_map.keys() if k in export_df.columns]
            export_df = export_df[cols_to_export].rename(columns=rename_map)
            
            export_df.to_excel(file_path, index=False)
            messagebox.showinfo("สำเร็จ", f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการบันทึกไฟล์: {e}")
            print(e)