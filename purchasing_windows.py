import customtkinter as ctk
from tkinter import messagebox, StringVar, TclError
import psycopg2
import pandas as pd
from datetime import datetime

class PurchaseOrderWindow(ctk.CTkToplevel):
    def __init__(self, master, app_container, po_id=None, on_close_callback=None):
        super().__init__(master)
        self.app_container = app_container
        self.pg_engine = self.app_container.pg_engine
        self.po_id = po_id
        self.on_close_callback = on_close_callback
        self.item_widgets = []

        self.title(f"ใบสั่งซื้อ (PO) - {'แก้ไข' if self.po_id else 'สร้างใหม่'}")
        self.geometry("900x700")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # --- ตัวแปรสำหรับคำนวณอัตโนมัติ ---
        self.total_po_shipping_cost = 0.0
        self.shipping_cost_var = StringVar(value="0.00")

        # --- Main Frame ---
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        # --- PO Details Frame ---
        self.po_details_frame = ctk.CTkFrame(main_frame)
        self.po_details_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.po_details_frame.grid_columnconfigure(1, weight=1)

        # Supplier, PO Number, SO Number
        ctk.CTkLabel(self.po_details_frame, text="Supplier:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_supplier = ctk.CTkEntry(self.po_details_frame)
        self.entry_supplier.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self.po_details_frame, text="PO Number:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_po_number = ctk.CTkEntry(self.po_details_frame)
        self.entry_po_number.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self.po_details_frame, text="SO Number:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_so_number = ctk.CTkEntry(self.po_details_frame)
        self.entry_so_number.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        # Total Cost (Read-only)
        ctk.CTkLabel(self.po_details_frame, text="Total Cost (excl. Shipping):").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.total_cost_var = StringVar(value="0.00")
        ctk.CTkLabel(self.po_details_frame, textvariable=self.total_cost_var, font=("Arial", 12, "bold")).grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # --- START: เปลี่ยนช่องกรอกค่าขนส่งเป็นป้ายแสดงผล ---
        ctk.CTkLabel(self.po_details_frame, text="ค่าขนส่ง PO (ก่อน VAT):").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.shipping_value_label = ctk.CTkLabel(self.po_details_frame, textvariable=self.shipping_cost_var, font=("Arial", 12, "bold"))
        self.shipping_value_label.grid(row=4, column=1, padx=10, pady=5, sticky="w")
        # --- END ---

        # --- Items Frame ---
        items_outer_frame = ctk.CTkFrame(main_frame)
        items_outer_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        items_outer_frame.grid_columnconfigure(0, weight=1)
        items_outer_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(items_outer_frame, text="รายการสินค้า (PO Items)", font=("Arial", 14, "bold")).grid(row=0, column=0, pady=(5,10))

        self.items_canvas = ctk.CTkCanvas(items_outer_frame, highlightthickness=0)
        self.items_scrollbar = ctk.CTkScrollbar(items_outer_frame, orientation="vertical", command=self.items_canvas.yview)
        self.scrollable_items_frame = ctk.CTkFrame(self.items_canvas)

        self.scrollable_items_frame.bind("<Configure>", lambda e: self.items_canvas.configure(scrollregion=self.items_canvas.bbox("all")))
        self.items_canvas.create_window((0, 0), window=self.scrollable_items_frame, anchor="nw")
        self.items_canvas.configure(yscrollcommand=self.items_scrollbar.set)

        self.items_canvas.grid(row=1, column=0, sticky="nsew")
        self.items_scrollbar.grid(row=1, column=1, sticky="ns")

        # Header
        header_frame = ctk.CTkFrame(self.scrollable_items_frame)
        header_frame.pack(fill="x", expand=True, pady=2)
        ctk.CTkLabel(header_frame, text="Description", width=250).pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(header_frame, text="Qty", width=50).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Cost/Unit", width=100).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Shipping to Stock", width=100).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Shipping to Site", width=100).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="", width=45).pack(side="left", padx=5) # For delete button

        # --- Buttons Frame ---
        buttons_frame = ctk.CTkFrame(main_frame)
        buttons_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        buttons_frame.grid_columnconfigure(0, weight=1) # Center button

        self.add_item_button = ctk.CTkButton(buttons_frame, text="เพิ่มรายการ", command=self._add_item_row)
        self.add_item_button.grid(row=0, column=0, padx=10, pady=10)

        self.submit_button = ctk.CTkButton(self, text="Submit PO", command=self._submit_po)
        self.submit_button.pack(pady=10)

        if self.po_id:
            self._load_po_data()

    def _load_po_data(self):
        try:
            query_po = "SELECT * FROM purchase_orders WHERE id = %s"
            po_df = pd.read_sql(query_po, self.pg_engine, params=(self.po_id,))
            if po_df.empty:
                messagebox.showerror("Error", "PO not found.", parent=self)
                return
            po_data = po_df.iloc[0].to_dict()

            self.entry_supplier.insert(0, po_data.get('supplier_name', ''))
            self.entry_po_number.insert(0, po_data.get('po_number', ''))
            self.entry_so_number.insert(0, po_data.get('so_number', ''))

            query_items = "SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id"
            items_df = pd.read_sql(query_items, self.pg_engine, params=(self.po_id,))
            for index, row in items_df.iterrows():
                self._add_item_row(row.to_dict())
            
            self._update_totals() # Calculate totals after loading all items

        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to load PO data: {e}", parent=self)

    def _add_item_row(self, item=None):
        item_frame = ctk.CTkFrame(self.scrollable_items_frame)
        item_frame.pack(fill="x", expand=True, pady=2, padx=2)

        desc_var = StringVar(value=item.get('description', '') if item else "")
        entry_desc = ctk.CTkEntry(item_frame, textvariable=desc_var)
        entry_desc.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        qty_var = StringVar(value=str(item.get('quantity', 1)) if item else "1")
        entry_qty = ctk.CTkEntry(item_frame, textvariable=qty_var, width=50)
        entry_qty.pack(side="left", padx=5, pady=5)

        cost_var = StringVar(value=str(item.get('cost_per_unit', '')) if item else "")
        entry_cost = ctk.CTkEntry(item_frame, textvariable=cost_var, width=100)
        entry_cost.pack(side="left", padx=5, pady=5)
        
        shipping_stock_var = StringVar(value=str(item.get('shipping_to_stock_cost', '')) if item else "")
        entry_shipping_stock = ctk.CTkEntry(item_frame, textvariable=shipping_stock_var, width=100)
        entry_shipping_stock.pack(side="left", padx=5, pady=5)

        shipping_site_var = StringVar(value=str(item.get('shipping_to_site_cost', '')) if item else "")
        entry_shipping_site = ctk.CTkEntry(item_frame, textvariable=shipping_site_var, width=100)
        entry_shipping_site.pack(side="left", padx=5, pady=5)

        # --- START: สั่งให้คำนวณใหม่ทุกครั้งที่มีการเปลี่ยนแปลง ---
        qty_var.trace_add("write", self._update_totals)
        cost_var.trace_add("write", self._update_totals)
        shipping_stock_var.trace_add("write", self._update_totals)
        shipping_site_var.trace_add("write", self._update_totals)
        # --- END ---

        delete_button = ctk.CTkButton(item_frame, text="X", width=20, command=lambda f=item_frame: self._delete_item_row(f))
        delete_button.pack(side="left", padx=5, pady=5)

        self.item_widgets.append({
            'frame': item_frame, 'desc_var': desc_var, 'qty_var': qty_var, 'cost_var': cost_var,
            'shipping_stock_var': shipping_stock_var, 'shipping_site_var': shipping_site_var,
            'id': item.get('id', None) if item else None
        })
        self._update_totals()

    def _delete_item_row(self, frame_to_delete):
        self.item_widgets = [row for row in self.item_widgets if row['frame'] != frame_to_delete]
        frame_to_delete.destroy()
        self._update_totals()

    def _update_totals(self, *args):
        """คำนวณผลรวมค่าสินค้าและค่าขนส่งจากทุกรายการ และอัปเดตหน้าจอ"""
        total_cost = 0.0
        total_shipping = 0.0
        for item_row in self.item_widgets:
            try:
                qty = float(item_row['qty_var'].get() or 0)
                cost_per_unit = float(item_row['cost_var'].get() or 0)
                stock_shipping = float(item_row['shipping_stock_var'].get() or 0)
                site_shipping = float(item_row['shipping_site_var'].get() or 0)
                
                total_cost += qty * cost_per_unit
                total_shipping += stock_shipping + site_shipping
            except (ValueError, TclError):
                continue # Skip if a field is empty or invalid
        
        self.total_cost = total_cost
        self.total_po_shipping_cost = total_shipping
        
        self.total_cost_var.set(f"{total_cost:,.2f}")
        self.shipping_cost_var.set(f"{total_shipping:,.2f}")

    def _submit_po(self):
        supplier_name = self.entry_supplier.get().strip()
        po_number = self.entry_po_number.get().strip()
        so_number = self.entry_so_number.get().strip()

        if not supplier_name or not po_number:
            messagebox.showwarning("Input Error", "Supplier Name and PO Number are required.", parent=self)
            return

        # Use the calculated total cost and shipping cost
        total_cost = self.total_cost
        po_shipping_cost = self.total_po_shipping_cost
        
        items_data = []
        for row in self.item_widgets:
            desc = row['desc_var'].get().strip()
            if not desc: continue # Skip empty rows
            items_data.append({
                'id': row['id'],
                'description': desc,
                'quantity': float(row['qty_var'].get() or 0),
                'cost_per_unit': float(row['cost_var'].get() or 0),
                'shipping_to_stock_cost': float(row['shipping_stock_var'].get() or 0),
                'shipping_to_site_cost': float(row['shipping_site_var'].get() or 0),
            })
        
        if not items_data:
            messagebox.showwarning("Input Error", "At least one item is required.", parent=self)
            return
            
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                if self.po_id: # UPDATE logic
                    update_po_query = """
                        UPDATE purchase_orders SET 
                        supplier_name=%s, po_number=%s, so_number=%s, total_cost=%s,
                        shipping_to_stock_cost=%s, shipping_to_site_cost=%s,
                        last_updated_by=%s, last_updated_at=NOW()
                        WHERE id=%s;
                    """
                    # For simplicity in this example, I'm summing all shipping into shipping_to_site_cost
                    # You might have separate columns or different logic
                    cursor.execute(update_po_query, (
                        supplier_name, po_number, so_number, total_cost,
                        0, po_shipping_cost, # Saving total shipping into shipping_to_site_cost
                        self.app_container.current_user_id, self.po_id
                    ))
                else: # INSERT logic
                    insert_po_query = """
                        INSERT INTO purchase_orders (supplier_name, po_number, so_number, total_cost, 
                        shipping_to_stock_cost, shipping_to_site_cost, status, created_by)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s) RETURNING id;
                    """
                    cursor.execute(insert_po_query, (
                        supplier_name, po_number, so_number, total_cost,
                        0, po_shipping_cost, # Saving total shipping into shipping_to_site_cost
                        'Pending Approval', self.app_container.current_user_id
                    ))
                    self.po_id = cursor.fetchone()[0]

                # --- Handle Items (Delete, Update, Insert) ---
                existing_item_ids = {item['id'] for item in items_data if item['id']}
                cursor.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = %s AND id NOT IN %s",
                               (self.po_id, tuple(existing_item_ids) if existing_item_ids else (None,)))

                for item in items_data:
                    if item['id']: # Update existing item
                        update_item_query = """
                            UPDATE purchase_order_items SET
                            description=%s, quantity=%s, cost_per_unit=%s,
                            shipping_to_stock_cost=%s, shipping_to_site_cost=%s
                            WHERE id=%s;
                        """
                        cursor.execute(update_item_query, (
                            item['description'], item['quantity'], item['cost_per_unit'],
                            item['shipping_to_stock_cost'], item['shipping_to_site_cost'], item['id']
                        ))
                    else: # Insert new item
                        insert_item_query = """
                            INSERT INTO purchase_order_items (purchase_order_id, description, quantity, 
                            cost_per_unit, shipping_to_stock_cost, shipping_to_site_cost)
                            VALUES (%s, %s, %s, %s, %s, %s);
                        """
                        cursor.execute(insert_item_query, (
                            self.po_id, item['description'], item['quantity'], item['cost_per_unit'],
                            item['shipping_to_stock_cost'], item['shipping_to_site_cost']
                        ))

            conn.commit()
            messagebox.showinfo("Success", "Purchase Order saved successfully.", parent=self)
            self._on_close()

        except psycopg2.Error as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"Failed to save PO: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _on_close(self):
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()