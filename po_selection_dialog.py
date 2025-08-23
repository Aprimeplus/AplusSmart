# po_selection_dialog.py

import tkinter as tk
from tkinter import ttk, messagebox
from customtkinter import CTkToplevel, CTkFrame, CTkEntry, CTkButton
import pandas as pd

class POSelectionDialog(CTkToplevel):
    def __init__(self, master, pg_engine, print_callback):
        super().__init__(master)
        self.pg_engine = pg_engine
        self.print_callback = print_callback
        self.all_pos_df = None

        self.title("เลือกใบสั่งซื้อ (PO) ที่อนุมัติแล้วเพื่อพิมพ์")
        self.geometry("800x600")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Search Bar ---
        search_frame = CTkFrame(self)
        search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.search_entry = CTkEntry(search_frame, placeholder_text="ค้นหาจากเลขที่ PO, SO, หรือชื่อซัพพลายเออร์...")
        self.search_entry.pack(fill="x", expand=True, padx=5, pady=5)
        self.search_entry.bind("<KeyRelease>", self._filter_po_list)

        # --- Treeview for PO List ---
        self.tree_frame = CTkFrame(self)
        self.tree_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self._create_treeview()

        self._load_all_pos()

        self.transient(master)
        self.grab_set()

    def _create_treeview(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Roboto', 14, 'bold'))
        style.configure("Treeview", rowheight=28, font=('Roboto', 12))
        
        columns = ['ID', 'PO Number', 'SO Number', 'Supplier', 'Date']
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show='headings')
        
        for col in columns:
            self.tree.heading(col, text=col)
            width = 60 if col == 'ID' else 250 if col == 'Supplier' else 150
            self.tree.column(col, width=width, anchor='w')

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self._on_po_select)

    def _load_all_pos(self):
        try:
            query = """
                SELECT id, po_number, so_number, supplier_name, timestamp 
                FROM purchase_orders 
                WHERE status = 'Approved'
                ORDER BY timestamp DESC
            """
            self.all_pos_df = pd.read_sql_query(query, self.pg_engine)
            self._populate_treeview(self.all_pos_df)
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดรายการ PO ได้: {e}", parent=self)

    def _populate_treeview(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for _, row in df.iterrows():
            dt = pd.to_datetime(row['timestamp']).strftime('%Y-%m-%d %H:%M')
            values = (row['id'], row['po_number'], row['so_number'], row['supplier_name'], dt)
            self.tree.insert('', 'end', values=values)

    def _filter_po_list(self, event=None):
        search_term = self.search_entry.get().lower()
        if not search_term:
            self._populate_treeview(self.all_pos_df)
            return

        filtered_df = self.all_pos_df[
            self.all_pos_df['po_number'].str.lower().str.contains(search_term, na=False) |
            self.all_pos_df['so_number'].str.lower().str.contains(search_term, na=False) |
            self.all_pos_df['supplier_name'].str.lower().str.contains(search_term, na=False)
        ]
        self._populate_treeview(filtered_df)

    def _on_po_select(self, event=None):
        selected_item = self.tree.focus()
        if not selected_item:
            return
        
        item_values = self.tree.item(selected_item)['values']
        po_id = item_values[0]
        
        if messagebox.askyesno("ยืนยัน", f"คุณต้องการพิมพ์ใบปะหน้าสำหรับ PO ID: {po_id} ใช่หรือไม่?", parent=self):
            self.destroy()
            self.print_callback(po_id)