# so_selection_dialog.py

import tkinter as tk
from tkinter import ttk, messagebox
from customtkinter import CTkToplevel, CTkFrame, CTkEntry
import pandas as pd

class SOSelectionDialog(CTkToplevel):
    def __init__(self, master, pg_engine, print_callback):
        super().__init__(master)
        self.pg_engine = pg_engine
        self.print_callback = print_callback
        self.all_sos_df = None

        self.title("เลือก Sales Order (SO) เพื่อพิมพ์ใบปะหน้า PO ทั้งหมด")
        self.geometry("800x600")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        search_frame = CTkFrame(self)
        search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.search_entry = CTkEntry(search_frame, placeholder_text="ค้นหาจากเลขที่ SO, ชื่อลูกค้า...")
        self.search_entry.pack(fill="x", expand=True, padx=5, pady=5)
        self.search_entry.bind("<KeyRelease>", self._filter_so_list)

        self.tree_frame = CTkFrame(self)
        self.tree_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self._create_treeview()
        self._load_all_sos()

        self.transient(master)
        self.grab_set()

    def _create_treeview(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Roboto', 14, 'bold'))
        style.configure("Treeview", rowheight=28, font=('Roboto', 12))
        
        columns = ['SO Number', 'Customer Name', 'Sale Name', 'Date']
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show='tree headings')
        
        for col in columns:
            self.tree.heading(col, text=col)
            width = 250 if col == 'Customer Name' else 150
            self.tree.column(col, width=width, anchor='w')

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self._on_so_select)

    def _load_all_sos(self):
        try:
            query = """
                SELECT DISTINCT ON (c.so_number)
                    c.so_number,
                    c.customer_name,
                    u.sale_name,
                    c.bill_date
                FROM commissions c
                JOIN purchase_orders po ON c.so_number = po.so_number
                JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE po.status = 'Approved' AND c.is_active = 1
                ORDER BY c.so_number, c.bill_date DESC
            """
            self.all_sos_df = pd.read_sql_query(query, self.pg_engine)
# --- เพิ่มบรรทัดนี้เพื่อแปลงวันที่ ---
            self.all_sos_df['bill_date'] = pd.to_datetime(self.all_sos_df['bill_date'])
            self._populate_treeview(self.all_sos_df)
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดรายการ SO ได้: {e}", parent=self)

    def _populate_treeview(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df.empty:
            return

        # สร้าง Dictionary สำหรับแปลงเดือนเป็นภาษาไทย
        thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        # จัดกลุ่มข้อมูลตาม "ปี-เดือน"
        df['year_month_group'] = df['bill_date'].dt.strftime('%Y-%m')
        grouped = df.groupby('year_month_group')

        # วนลูปตามกลุ่มเดือน (เรียงจากล่าสุดไปเก่าสุด)
        for name, group in sorted(grouped, key=lambda x: x[0], reverse=True):
            year, month = map(int, name.split('-'))
            
            # สร้างชื่อกลุ่มภาษาไทย เช่น "สิงหาคม 2568"
            group_display_name = f"{thai_months[month - 1]} {year + 543}"
            
            # 1. เพิ่มแถวหลักของเดือนเข้าไปในตาราง (เป็น Parent)
            month_iid = self.tree.insert('', 'end', text=group_display_name, open=True) # open=True คือให้กางออกไว้ก่อน
            
            # 2. วนลูปเพื่อเพิ่ม SO ของเดือนนั้นๆ เป็นรายการย่อย (เป็น Child)
            for _, row in group.sort_values(by='bill_date', ascending=False).iterrows():
                dt_str = row['bill_date'].strftime('%Y-%m-%d')
                values = (row['so_number'], row['customer_name'], row['sale_name'], dt_str)
                # ใช้ month_iid เป็น parent เพื่อให้เป็นรายการย่อย
                self.tree.insert(month_iid, 'end', values=values)

    def _filter_so_list(self, event=None):
        search_term = self.search_entry.get().lower()
        if not search_term:
            self._populate_treeview(self.all_sos_df)
            return
        filtered_df = self.all_sos_df[
            self.all_sos_df['so_number'].str.lower().str.contains(search_term, na=False) |
            self.all_sos_df['customer_name'].str.lower().str.contains(search_term, na=False)
        ]
        self._populate_treeview(filtered_df)

    def _on_so_select(self, event=None):
        selected_item = self.tree.focus()
        if not selected_item: return

        # --- START: เพิ่มโค้ดตรวจสอบ ---
        # 1. ดึงข้อมูล values ของแถวที่เลือก
        item_values = self.tree.item(selected_item)['values']

        # 2. ตรวจสอบว่าแถวที่คลิกมีข้อมูล values หรือไม่
        # (แถวที่เป็นชื่อเดือนจะไม่มีข้อมูลในส่วนนี้)
        if not item_values or not item_values[0]:
            return # ถ้าไม่มีข้อมูล ให้หยุดการทำงานทันที

        # 3. ถ้ามีข้อมูล ให้ทำงานต่อไปตามปกติ
        so_number = item_values[0]
        # --- END: สิ้นสุดโค้ดตรวจสอบ ---

        if messagebox.askyesno("ยืนยัน", f"คุณต้องการพิมพ์ใบปะหน้าสำหรับ PO ทั้งหมดใน SO: {so_number} ใช่หรือไม่?", parent=self):
            self.destroy()
            self.print_callback(so_number)