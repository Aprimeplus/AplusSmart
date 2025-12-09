# purchasing_manager_screen.py (ฉบับปรับปรุง เพิ่มแท็บ Master Edit)

import tkinter as tk
from tkinter import ttk
from customtkinter import (CTkFrame, CTkLabel, CTkFont, CTkButton,
                           CTkScrollableFrame, CTkInputDialog, CTkToplevel, CTkCheckBox, CTkEntry,
                           CTkOptionMenu, CTkTabview) # <-- เพิ่ม CTkTabview
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import psycopg2.errors
import psycopg2.extras
import json
import traceback
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
from sqlalchemy import create_engine
from export_utils import export_approved_pos_to_excel
from pdf_utils import export_approved_pos_to_pdf
from po_selection_dialog import POSelectionDialog
from hr_windows import SOPopupWindow
from history_windows import PurchaseDetailWindow, PurchaseHistoryWindow
from purchasing_screen import PurchasingScreen # <-- Import หน้าจอของ PU เข้ามา
from reject_history import RejectionHistoryWindow


class RejectionReasonDialog(CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.master = master; self.title("ระบุเหตุผลที่ปฏิเสธ"); self.geometry("500x600")
        self.reasons_list = ["ลงสเปคสินค้าผิด SO", "ลงเสปคสินค้าผิด PO", "ลงราคาต้นทุนผิด PO", "ลงราคาขายผิด SO", "ไม่แยกค่ารถ/ราคาผิด SO", "ไม่แยกค่ารถ/ราคาผิด PO", "รายการต้นทุนไม่ครบ PO", "ค่าตัด/เจาะ ตกหล่น", "ค่าของแถม ตกหล่น"]
        self.checkbox_vars = []; self._reason_string = None
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        CTkLabel(self, text="กรุณาเลือกเหตุผลที่ปฏิเสธ (เลือกได้มากกว่า 1 ข้อ)", font=CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=10)
        scroll_frame = CTkScrollableFrame(self); scroll_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        for reason in self.reasons_list:
            var = tk.StringVar(value="0"); cb = CTkCheckBox(scroll_frame, text=reason, variable=var, font=CTkFont(size=14)); cb.pack(pady=5, padx=10, anchor="w"); self.checkbox_vars.append((var, reason))
        other_frame = CTkFrame(self, fg_color="transparent"); other_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew"); other_frame.grid_columnconfigure(1, weight=1)
        CTkLabel(other_frame, text="อื่นๆ:", font=CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=(0,5)); self.other_reason_entry = CTkEntry(other_frame); self.other_reason_entry.grid(row=0, column=1, sticky="ew")
        button_frame = CTkFrame(self, fg_color="transparent"); button_frame.grid(row=3, column=0, padx=20, pady=10)
        CTkButton(button_frame, text="ยกเลิก", command=self.destroy).pack(side="right", padx=5); CTkButton(button_frame, text="ตกลง", command=self._on_confirm).pack(side="right", padx=5)
        self.transient(master); self.grab_set()

    def _on_confirm(self):
        selected_reasons = [reason_text for var, reason_text in self.checkbox_vars if var.get() == "1"]
        other_text = self.other_reason_entry.get().strip()
        if other_text: selected_reasons.append(f"อื่นๆ: {other_text}")
        if not selected_reasons: messagebox.showwarning("ข้อมูลไม่ครบถ้วน", "กรุณาเลือกเหตุผลอย่างน้อย 1 ข้อ", parent=self); return
        self._reason_string = ", ".join(selected_reasons); self.destroy()

class ReopenPOWindow(CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.app_container = master.app_container
        self.user_key = master.user_key
        self.title("ดึงงาน PO กลับมาแก้ไข (Re-open PO)")
        self.geometry("900x600")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.all_pos_df = None

        search_frame = CTkFrame(self, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        self.search_entry = CTkEntry(search_frame, placeholder_text="ค้นหาจากเลขที่ PO, SO, หรือชื่อซัพพลายเออร์...")
        self.search_entry.pack(fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self._filter_po_list)

        self.main_frame = CTkScrollableFrame(self, label_text="รายการ PO ที่สามารถดึงกลับมาแก้ไขได้")
        self.main_frame.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        self._load_reopenable_pos()
        self.transient(master)
        self.grab_set()

    def _load_reopenable_pos(self):
        conn = None
        try:
            conn = self.app_container.get_connection()
            query = """
                SELECT
                    po.id AS po_id,
                    po.po_number,
                    po.so_number,
                    po.supplier_name,
                    po.user_key AS po_creator_key,
                    po.timestamp
                FROM
                    purchase_orders po
                JOIN
                    commissions c ON po.so_number = c.so_number
                WHERE
                    c.status = 'Pending Sale Manager Approval'
                    AND (po.approver_manager1_key = %s OR po.approver_manager2_key = %s)
                    AND po.status = 'Approved'
                ORDER BY
                    po.timestamp DESC;
            """
            self.all_pos_df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.user_key, self.user_key))
            self._populate_po_list(self.all_pos_df)
        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลด PO: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

    def _populate_po_list(self, df_to_show):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        if df_to_show.empty:
            CTkLabel(self.main_frame, text="ไม่พบข้อมูล PO ที่ตรงกับเงื่อนไข", text_color="gray50").pack(pady=20)
            return

        for _, po_data in df_to_show.iterrows():
            card = CTkFrame(self.main_frame, border_width=1)
            card.pack(fill="x", padx=10, pady=5)
            
            info_text = f"PO: {po_data['po_number']} | SO: {po_data['so_number']} | Supplier: {po_data['supplier_name']}"
            
            CTkLabel(card, text=info_text).pack(side="left", padx=15, pady=10)

            # <<< START: เพิ่ม Frame สำหรับจัดวางปุ่ม และเพิ่มปุ่ม "แก้ไข" >>>
            action_frame = CTkFrame(card, fg_color="transparent")
            action_frame.pack(side="right", padx=15, pady=10)

            # --- ปุ่มใหม่สำหรับ "แก้ไข" โดยตรง ---
            edit_button = CTkButton(
                action_frame,
                text="แก้ไข",
                fg_color="#3B82F6", # สีน้ำเงิน
                hover_color="#2563EB",
                command=lambda p_id=po_data['po_id']: self.app_container.show_purchase_detail_window(
                    purchase_id=int(p_id),
                    on_save_callback=self._load_reopenable_pos # สั่งให้หน้านี้ Refresh ตัวเองหลังบันทึก
                )
            )
            edit_button.pack(side="left", padx=(0, 5))

            # --- ปุ่มเดิมสำหรับ "ดึง PO กลับ" ---
            reopen_button = CTkButton(
                action_frame, 
                text="ดึง PO กลับ", # ปรับข้อความให้สั้นลง
                fg_color="#F97316", 
                hover_color="#EA580C", 
                command=lambda data=po_data.to_dict(): self._reopen_po(data))
            reopen_button.pack(side="left")

    def _filter_po_list(self, event=None):
        search_term = self.search_entry.get().lower().strip()
        if not search_term:
            filtered_df = self.all_pos_df
        else:
            if self.all_pos_df is not None:
                filtered_df = self.all_pos_df[
                    self.all_pos_df['po_number'].str.lower().str.contains(search_term, na=False) |
                    self.all_pos_df['so_number'].str.lower().str.contains(search_term, na=False) |
                    self.all_pos_df['supplier_name'].str.lower().str.contains(search_term, na=False)
                ]
            else:
                filtered_df = pd.DataFrame()
        self._populate_po_list(filtered_df)

    def _reopen_po(self, po_data):
        po_id = po_data['po_id']
        po_number = po_data['po_number']
        so_number = po_data['so_number']
        po_creator_key = po_data['po_creator_key']

        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการดึงงาน PO: {po_number} กลับมาใช่หรือไม่?\nPO ใบนี้จะถูกเปลี่ยนเป็น 'Draft' และส่งกลับไปให้ฝ่ายจัดซื้อ ({po_creator_key}) แก้ไข", icon="warning", parent=self): 
            return
        
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                cursor.execute("""
                    UPDATE purchase_orders 
                    SET status = 'Draft', approval_status = 'Draft',
                        approver_manager1_key = NULL, approval_date_manager1 = NULL,
                        approver_manager2_key = NULL, approval_date_manager2 = NULL
                    WHERE id = %s
                """, (po_id,))
                
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'PO In Progress' 
                    WHERE so_number = %s
                """, (so_number,))
                
                message_to_pu = f"PO: {po_number} ถูกดึงกลับมาเพื่อแก้ไขโดย Manager ({self.user_key})"
                cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)", 
                               (po_creator_key, message_to_pu, po_id))
                
                cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Sales Manager' AND status = 'Active'")
                manager_keys = [row[0] for row in cursor.fetchall()]
                message_to_sm = f"SO: {so_number} ถูกดึงกลับไปให้ฝ่ายจัดซื้อแก้ไขโดย PU Manager"
                for key in manager_keys:
                    cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read) VALUES (%s, %s, FALSE)", (key, message_to_sm))

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ดึงงาน PO: {po_number} กลับมาเรียบร้อยแล้ว", parent=self)
            self.master._load_pending_pos() 
            self._load_reopenable_pos()
        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการ Re-open PO: {e}", parent=self)
        finally:
            if conn: self.app_container.release_connection(conn)

class SOPendingDetailWindow(CTkToplevel):
    def __init__(self, master, so_number):
        super().__init__(master)
        self.app_container = master.app_container; self.so_number = so_number; self.df = None
        self.title(f"สรุปรายการสินค้าทั้งหมดสำหรับ SO: {self.so_number} (ดับเบิลคลิกเพื่อดู PO ต้นทาง)"); self.geometry("1200x600")
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        header_frame = CTkFrame(self, fg_color="transparent"); header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(10,5))
        CTkLabel(header_frame, text=f"รายการสินค้าทั้งหมดของ SO: {self.so_number} (ดับเบิลคลิกเพื่อดู PO ต้นทาง)", font=CTkFont(size=16, weight="bold")).pack(side="left")
        CTkButton(header_frame, text="Refresh", command=self._load_and_display_table, width=100).pack(side="right")
        self.tree_frame = CTkFrame(self); self.tree_frame.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")
        self.tree_frame.grid_columnconfigure(0, weight=1); self.tree_frame.grid_rowconfigure(0, weight=1)
        self._load_and_display_table(); self.transient(master); self.grab_set()

    def _load_and_display_table(self):
        for widget in self.tree_frame.winfo_children(): widget.destroy()
        try:
            query = "SELECT po.id as po_id, item.id as item_id, po.po_number, po.supplier_name, item.product_name, item.quantity, item.unit_price, item.total_price FROM purchase_orders po JOIN purchase_order_items item ON po.id = item.purchase_order_id WHERE po.so_number = %s AND po.approval_status IN ('Pending Mgr 1', 'Pending Mgr 2', 'Pending Director') ORDER BY po.id, item.id;"
            self.df = pd.read_sql_query(query, self.app_container.pg_engine, params=(self.so_number,))
            if self.df.empty:
                CTkLabel(self.tree_frame, text="ไม่พบรายการสินค้าที่รออนุมัติสำหรับ SO นี้", text_color="gray50").pack(pady=20)
                self.after(1500, self.destroy); return
            self._create_table_view()
        except Exception as e: messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการโหลดรายละเอียด: {e}", parent=self)
            
    def _create_table_view(self):
        style = ttk.Style(self); style.theme_use("default"); style.configure("Treeview.Heading", font=('Roboto', 14, 'bold')); style.configure("Treeview", rowheight=28, font=('Roboto', 12))
        columns = ['PO Number', 'Supplier', 'Product Name', 'Quantity', 'Unit Price', 'Total Price']
        tree = ttk.Treeview(self.tree_frame, columns=columns, show='headings')
        for col in columns:
            width = 300 if col == 'Product Name' else 200 if col == 'Supplier' else 150 if col == 'PO Number' else 120
            anchor = 'e' if 'Price' in col or 'Quantity' in col else 'w'
            tree.heading(col, text=col); tree.column(col, width=width, anchor=anchor)
        for _, row in self.df.iterrows():
            values = (row['po_number'], row['supplier_name'], row['product_name'], f"{row['quantity']:,.2f}", f"{row['unit_price']:,.2f}", f"{row['total_price']:,.2f}")
            unique_iid = f"{row['po_id']}-{row['item_id']}"
            tree.insert('', 'end', values=values, iid=unique_iid)
        v_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=v_scroll.set)
        tree.grid(row=0, column=0, sticky="nsew"); v_scroll.grid(row=0, column=1, sticky="ns")
        tree.bind("<Double-1>", self._on_po_double_click)

    def _on_po_double_click(self, event):
        iid_to_view = event.widget.focus()
        if not iid_to_view: return
        po_id_to_view = str(iid_to_view).split('-')[0]
        self.app_container.show_purchase_detail_window(po_id_to_view)

class PurchasingManagerScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None):
        super().__init__(master, corner_radius=0, fg_color=app_container.THEME["purchasing"]["bg"])
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role
        self.theme = self.app_container.THEME["purchasing"]
        self.header_font = CTkFont(size=16, weight="bold")
        self.dropdown_style = {
            "fg_color": "white", "text_color": "black",
            "button_color": self.theme.get("primary", "#3B82F6"),
            "button_hover_color": self.theme.get("header", "#2563EB")
        }
        self.pg_engine = self.app_container.pg_engine
        self.all_pending_df = pd.DataFrame()
        self.so_cards = {}
        self.po_cards = {}
        
        self.current_page = 0
        self.rows_per_page = 15
        self._debounce_job = None
        
        self.rejection_chart_canvas, self.polling_job_id, self.so_detail_window, self.reopen_window = None, None, None, None
        self._so_create_string_vars()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._create_header()

        self.tab_view = CTkTabview(self, corner_radius=10, border_width=1)
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.manager_view_tab = self.tab_view.add("ภาพรวมและอนุมัติ (Manager View)")
        self.staff_view_tab = self.tab_view.add("สร้าง/แก้ไข PO (Staff View)")
        
        # --- START: เพิ่มแท็บ Master Edit ---
        self.master_edit_tab = self.tab_view.add("ค้นหา & ตีกลับ (Master Edit)")
        self.master_edit_tab.grid_columnconfigure(0, weight=1)
        self.master_edit_tab.grid_rowconfigure(1, weight=1)
        # --- END ---

        self.manager_view_tab.grid_columnconfigure(0, weight=1)
        self.manager_view_tab.grid_rowconfigure(1, weight=1)

        self._create_dashboard_view(parent_tab=self.manager_view_tab)
        self._create_pending_list_view(parent_tab=self.manager_view_tab)
        
        # --- START: เพิ่มการเรียกใช้ฟังก์ชันสร้างแท็บใหม่ ---
        self._create_master_edit_tab(self.master_edit_tab)
        # --- END ---
        
        # +++ START: แก้ไขการสร้าง PurchasingScreen ตรงนี้ +++
        self.purchasing_staff_screen = PurchasingScreen(
            master=self.staff_view_tab,
            app_container=self.app_container, # <-- ส่ง app_container เข้าไปโดยตรง
            user_key=self.user_key,
            user_name=self.user_name,
            user_role=self.user_role
        )
        self.purchasing_staff_screen.pack(fill="both", expand=True)
        # +++ END +++

        self._load_data()
        self._start_polling()
        self.bind("<Destroy>", self._on_destroy)
    
    # --- START: เพิ่ม 6 ฟังก์ชันใหม่สำหรับแท็บ Master Edit ---
    
    def _create_master_edit_tab(self, parent_tab):
        """สร้าง UI สำหรับแท็บ Master Edit (คล้ายของ HR)"""
        parent_tab.grid_columnconfigure(0, weight=1)
        parent_tab.grid_rowconfigure(1, weight=1) # ให้แถวที่ 1 (ScrollFrame) ขยาย

        # --- Frame สำหรับฟิลเตอร์และการค้นหา ---
        search_frame = CTkFrame(parent_tab)
        search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)
        
        self.mp_master_search_entry = CTkEntry(search_frame, font=self.header_font, placeholder_text="กรอก SO หรือ PO ที่ต้องการค้นหา...")
        self.mp_master_search_entry.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="ew")
        
        self.mp_master_search_entry.bind("<Return>", lambda event: self._mp_master_search())
        self.mp_master_search_entry.bind("<KP_Enter>", lambda event: self._mp_master_search())
        
        search_button = CTkButton(search_frame, text="ค้นหา", command=self._mp_master_search, width=100)
        search_button.grid(row=0, column=1, padx=5, pady=10)
        
        clear_button = CTkButton(search_frame, text="ล้างค่า", command=lambda: self.mp_master_search_entry.delete(0, 'end'), fg_color="gray", width=80)
        clear_button.grid(row=0, column=2, padx=5, pady=10)

        # --- Frame สำหรับแสดงผลลัพธ์ ---
        self.mp_master_results_frame = CTkScrollableFrame(parent_tab, label_text="ผลการค้นหา")
        self.mp_master_results_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.mp_master_results_frame.grid_columnconfigure(0, weight=1)

    def _open_rejection_history(self):
        try:
            RejectionHistoryWindow(master=self, app_container=self.app_container)
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถเปิดประวัติการตีกลับได้: {e}", parent=self)

    def _mp_master_search(self):
        """ค้นหา SO/PO ทั้งหมดจาก Keyword"""
        for widget in self.mp_master_results_frame.winfo_children():
            widget.destroy()
            
        keyword = self.mp_master_search_entry.get().strip().upper()
        if not keyword:
            return

        search_term = keyword
        if search_term.startswith("SO"): search_term = search_term[2:]
        elif search_term.startswith("PO"): search_term = search_term[2:]
        
        try:
            # ค้นหา SO (ดึงข้อมูลเซลส์และสถานะมาด้วย)
            so_query = """
                SELECT c.id, c.so_number, c.customer_name, c.sale_key, c.status, u.sale_name
                FROM commissions c
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
                WHERE c.so_number ILIKE %s AND c.is_active = 1
            """
            so_df = pd.read_sql_query(so_query, self.pg_engine, params=(f"%{search_term}%",))
            
            # ค้นหา PO (ดึงข้อมูลผู้สร้างและสถานะมาด้วย)
            po_query = """
                SELECT p.id, p.so_number, p.po_number, p.supplier_name, p.status, p.user_key
                FROM purchase_orders p 
                WHERE p.po_number ILIKE %s
            """
            po_df = pd.read_sql_query(po_query, self.pg_engine, params=(f"%{search_term}%",))

            if so_df.empty and po_df.empty:
                CTkLabel(self.mp_master_results_frame, text=f"ไม่พบข้อมูลสำหรับ '{keyword}'").pack(pady=20)
                return

            if not so_df.empty:
                CTkLabel(self.mp_master_results_frame, text="ผลการค้นหา: Sales Orders (SO)", font=self.header_font).pack(anchor="w", padx=10, pady=(10,0))
                for _, row in so_df.iterrows():
                    self._create_mp_so_card(self.mp_master_results_frame, row.to_dict())

            if not po_df.empty:
                CTkLabel(self.mp_master_results_frame, text="ผลการค้นหา: Purchase Orders (PO)", font=self.header_font).pack(anchor="w", padx=10, pady=(10,0))
                for _, row in po_df.iterrows():
                    self._create_mp_po_card(self.mp_master_results_frame, row.to_dict())

        except Exception as e:
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการค้นหา: {e}", parent=self)

    def _create_mp_so_card(self, parent, so_data):
        """สร้าง Card แสดงผลลัพธ์ SO (เวอร์ชันปรับปรุง: เพิ่มปุ่มตีกลับ SO)"""
        so_number = so_data['so_number']
        so_status = so_data.get('status', 'N/A')
        
        status_colors = {
            'Paid': '#D1FAE5', 'HR Verified': '#A7F3D0', 
            'PO Sent': '#67E8F9', 'Pending Sale Manager Approval': '#FDE047',
            'PO In Progress': '#FEF3C7', 'Draft': '#E5E7EB',
            'Rejected': '#FECACA', 'Cancelled': '#FECACA'
        }
        card_color = status_colors.get(so_status, "#F9FAFB")

        so_card = CTkFrame(parent, border_width=1, fg_color=card_color)
        so_card.pack(fill="x", padx=10, pady=5)
        
        info_text = f"SO: {so_number} | ลูกค้า: {so_data.get('customer_name','N/A')} | สถานะ: {so_status}"
        CTkLabel(so_card, text=info_text).pack(side="left", padx=10, pady=5)
        
        # --- START: เพิ่ม Frame สำหรับจัดวางปุ่ม ---
        action_frame = CTkFrame(so_card, fg_color="transparent")
        action_frame.pack(side="right", padx=10, pady=5)
        
        # ปุ่มสำหรับ SO "แก้ไข" (เหมือนเดิม)
        CTkButton(
            action_frame, 
            text="แก้ไข SO", 
            width=100, 
            command=lambda s=so_number: self._open_so_editor_for_mp(s)
        ).pack(side="left", padx=5)
        
        # --- START: เพิ่มปุ่ม "ตีกลับ SO ทั้งหมด" (ปุ่มใหม่) ---
        CTkButton(
            action_frame, 
            text="ตีกลับ SO ทั้งหมด", 
            width=140, 
            fg_color="#D32F2F", hover_color="#B71C1C", # สีแดง (อันตราย)
            command=lambda s_num=so_number: self._mp_master_revert_so(s_num)
        ).pack(side="left", padx=5)

    def _create_mp_po_card(self, parent, po_data):
        """สร้าง Card แสดงผลลัพธ์ PO"""
        po_id = int(po_data['id'])
        po_number = po_data['po_number']
        so_number = po_data['so_number']
        po_status = po_data.get('status', 'N/A')
        po_creator = po_data.get('user_key', 'N/A')
        
        status_colors = {
            'Approved': '#D1FAE5', 'Pending Approval': '#FEF3C7',
            'Draft': '#E5E7EB', 'Rejected': '#FECACA', 'Cancelled': '#FECACA'
        }
        card_color = status_colors.get(po_status, "#F9FAFB")

        po_card = CTkFrame(parent, border_width=1, fg_color=card_color)
        po_card.pack(fill="x", padx=10, pady=5)
        
        info_text = f"PO: {po_number} | SO: {so_number} | Supplier: {po_data.get('supplier_name','N/A')} | สถานะ: {po_status}"
        CTkLabel(po_card, text=info_text).pack(side="left", padx=10, pady=5)
        
        action_frame = CTkFrame(po_card, fg_color="transparent")
        action_frame.pack(side="right", padx=10, pady=5)

        # ปุ่ม "แก้ไข" (เปิดหน้าต่างแก้ไข PO)
        CTkButton(
            action_frame, 
            text="แก้ไข PO", 
            width=100, 
            command=lambda pid=po_id: self._view_details(pid)
        ).pack(side="left", padx=5)

        # ปุ่ม "ตีกลับ" (Revert)
        CTkButton(
            action_frame, 
            text="ตีกลับ (Revert)", 
            width=120, 
            fg_color="#F97316", hover_color="#EA580C",
            command=lambda pid=po_id, s_num=so_number, creator=po_creator: self._mp_master_revert_po(pid, s_num, creator)
        ).pack(side="left", padx=5)

    def _mp_master_revert_po(self, po_id, so_number, po_creator_key):
        """
        ฟังก์ชันสำหรับ "ตีกลับ" PO ที่เลือกให้กลับไปเป็น Draft
        """
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. ตรวจสอบสถานะ SO ก่อน
                cursor.execute("SELECT status, sale_key FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", (so_number,))
                so_record = cursor.fetchone()          
                so_status = so_record['status'] if so_record else 'N/A'
                so_sale_key = so_record['sale_key'] if so_record and 'sale_key' in so_record else None # <-- เพิ่มบรรทัดนี้
                print(f"--- DEBUG REVERT PO ---")
                print(f"SO: {so_number} | Status ที่ดึงได้: {so_status} | Sale Key ที่ดึงได้: {so_sale_key}")
                # ^^^^ สิ้นสุดส่วนที่เพิ่ม ^^^^
                                
                
                
                # 2. ตั้งค่าคำเตือน
                warning_message = ""
                if so_status in ('Paid', 'HR Verified'):
                    warning_message = (
                        f"\n\n!! คำเตือน !!\n"
                        f"SO ({so_number}) นี้ อยู่ในสถานะ '{so_status}' แล้ว\n"
                        "การตีกลับ PO อาจทำให้ข้อมูลค่าคอมมิชชั่นที่คำนวณไปแล้ว 'ผิดพลาด'\n"
                        "กรุณาประสานงานกับ HR ก่อนดำเนินการ"
                    )
                
                if not messagebox.askyesno(
                    "ยืนยันการตีกลับ (Revert)",
                    f"คุณต้องการตีกลับ PO ID: {po_id} ใช่หรือไม่?\n"
                    f"PO ใบนี้จะถูกเปลี่ยนเป็น 'Draft' และส่งกลับไปให้ฝ่ายจัดซื้อ ({po_creator_key}) แก้ไข"
                    f"{warning_message}",
                    icon="warning",
                    parent=self
                ):
                    return

                # 3. ดำเนินการตีกลับ (Revert)
                # 3.1 อัปเดต PO
                cursor.execute("""
                    UPDATE purchase_orders 
                    SET 
                        status = 'Draft', 
                        approval_status = 'Draft',
                        approver_manager1_key = NULL, approval_date_manager1 = NULL,
                        approver_manager2_key = NULL, approval_date_manager2 = NULL,
                        approver_director_key = NULL, approval_date_director = NULL,
                        last_modified_by = %s,
                        rejection_reason = %s
                    WHERE id = %s
                """, (self.user_key, f"Reverted by MP ({self.user_key}) at {datetime.now()}", po_id))

                # 3.2 อัปเดต SO (ถ้า SO ยังไม่ถูกจ่ายเงิน)
                if so_status not in ('Paid', 'HR Verified'):
                    cursor.execute("""
                        UPDATE commissions 
                        SET status = 'Draft'
                        WHERE so_number = %s AND is_active = 1
                    """, (so_number,))
                
                # 3.3 ส่ง Notification
                message_to_pu = f"PO ID: {po_id} ถูกตีกลับ (Revert) โดย Manager ({self.user_key}) กรุณาแก้ไขและส่งอนุมัติใหม่"
                cursor.execute(
                    "INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)", 
                    (po_creator_key, message_to_pu, po_id)
                )
                
                if so_sale_key:
                    message_to_sale = f"SO: {so_number} (จาก PO ID: {po_id}) ถูกตีกลับโดย Manager ({self.user_key}) กรุณาตรวจสอบและแก้ไข SO"
                    cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, is_read) VALUES (%s, %s, FALSE)", 
                        (so_sale_key, message_to_sale)
                    )
                 
                # 3.4 บันทึก Audit Log
                log_details = {'reverted_by': self.user_key, 'original_so_status': so_status}
                cursor.execute(
                    "INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp) VALUES (%s, %s, %s, %s, %s, %s)", 
                    ('PO Reverted', 'purchase_orders', po_id, self.user_key, json.dumps(log_details), datetime.now())
                )

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ตีกลับ PO ID: {po_id} เป็น 'Draft' เรียบร้อยแล้ว", parent=self)
            
            # Refresh หน้าจอค้นหา
            self._mp_master_search()
            
            # Refresh หน้าจอหลัก (เผื่อรายการนั้นค้างอยู่ที่หน้าแรก)
            self._load_data()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการ Revert PO: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _mp_master_revert_so(self, so_number):
        """
        (ฟังก์ชันแก้ไข: ตีกลับเฉพาะ SO)
        สำหรับ "ตีกลับ" SO กลับไปเป็น 'Draft' เพื่อให้เซลส์แก้ไข
        **โดยจะไม่กระทบสถานะของ PO ที่มีอยู่แล้ว**
        **และแจ้งเตือนทั้ง Sale หลัก และ Sale Support (ถ้ามี)**
        """
        conn = None
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                
                # 1. ค้นหาข้อมูล SO และ Support User Key
                # (เพิ่มการดึง support_user_key เพื่อนำมาแจ้งเตือน)
                cursor.execute("""
                    SELECT id, status, sale_key, support_user_key 
                    FROM commissions 
                    WHERE so_number = %s AND is_active = 1 
                    LIMIT 1
                """, (so_number,))
                so_record = cursor.fetchone()
                
                if not so_record:
                    messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบ SO: {so_number} ในระบบ", parent=self)
                    return

                so_id = so_record['id']
                so_status = so_record['status']
                so_sale_key = so_record['sale_key']
                support_key = so_record['support_user_key'] # ดึงรหัส Sale Support

                # 2. สร้างคำเตือน
                warning_message = ""
                if so_status in ('Paid', 'HR Verified'):
                    warning_message = (
                        f"\n\n!! คำเตือนรุนแรง !!\n"
                        f"SO ({so_number}) นี้ อยู่ในสถานะ '{so_status}' แล้ว\n"
                        "การตีกลับอาจส่งผลกระทบต่อข้อมูลการเงินและค่าคอมมิชชั่น\n"
                        "**กรุณาประสานงานกับ HR และบัญชีหากดำเนินการต่อ**"
                    )
                
                # ข้อความยืนยัน
                if not messagebox.askyesno(
                    "ยืนยันการตีกลับ SO (เฉพาะ SO)",
                    f"คุณต้องการตีกลับ SO: {so_number} ใช่หรือไม่?\n\n"
                    f"1. SO จะถูกเปลี่ยนสถานะเป็น 'Draft' (ส่งคืนเซลส์/Support)\n"
                    f"2. ใบสั่งซื้อ (PO) ที่เกี่ยวข้องจะ **ไม่ถูกแก้ไข** และคงสถานะเดิม\n"
                    f"{warning_message}",
                    icon="warning",
                    parent=self
                ):
                    return

                # 3. ดำเนินการตีกลับ (Revert) เฉพาะ SO
                
                # อัปเดต SO ให้เป็น Draft (เพื่อให้เซลส์/Support แก้ไขได้)
                cursor.execute("""
                    UPDATE commissions 
                    SET status = 'Draft', rejection_reason = %s
                    WHERE so_number = %s AND is_active = 1
                """, (f"Reverted by MP ({self.user_key}) on {datetime.now().strftime('%Y-%m-%d')}", so_number))
                
                # 4. ส่ง Notification
                msg_content = f"SO: {so_number} ถูกตีกลับโดย Manager ({self.user_key}) เพื่อให้แก้ไขข้อมูล (PO ไม่ถูกยกเลิก)"

                # 4.1 แจ้งเตือน Sale เจ้าของเคส
                if so_sale_key:
                    cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, is_read) VALUES (%s, %s, FALSE)", 
                        (so_sale_key, msg_content)
                    )
                
                # 4.2 [เพิ่ม] แจ้งเตือน Sale Support (ถ้ามีคนทำให้)
                if support_key and support_key != so_sale_key:
                    msg_support = f"งาน SO: {so_number} ที่คุณดูแล ถูกตีกลับโดย Manager ({self.user_key}) กรุณาตรวจสอบ"
                    cursor.execute(
                        "INSERT INTO notifications (user_key_to_notify, message, is_read) VALUES (%s, %s, FALSE)", 
                        (support_key, msg_support)
                    )

                # 5. บันทึก Audit Log
                log_details = {
                    'reverted_by': self.user_key, 
                    'action': 'Revert SO Only (Keep POs)', 
                    'original_so_status': so_status,
                    'notified_support': support_key
                }
                cursor.execute(
                    "INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp) VALUES (%s, %s, %s, %s, %s, %s)", 
                    ('SO Reverted', 'commissions', so_id, self.user_key, json.dumps(log_details, default=str), datetime.now())
                )

            conn.commit()
            messagebox.showinfo("สำเร็จ", f"ตีกลับ SO: {so_number} เป็น 'Draft' เรียบร้อยแล้ว\n(PO ทั้งหมดยังคงสถานะเดิม)", parent=self)
            
            # Refresh หน้าจอ
            self._mp_master_search()
            self._load_data()

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการ Revert SO: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)
    
    # --- END: สิ้นสุด 6 ฟังก์ชันใหม่ ---
    
    
    def _clear_search(self):
        """ล้างข้อความในช่องค้นหาและโหลดข้อมูลทั้งหมดใหม่"""
        self.search_entry.delete(0, 'end')
        self._filter_pending_list()

    def _next_page(self):
        """เลื่อนไปหน้าถัดไป"""
        if self.filtered_df is None: return
        total_pages = (len(self.filtered_df.groupby('so_number')) + self.rows_per_page - 1) // self.rows_per_page
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._populate_pending_list(self.filtered_df)

    def _prev_page(self):
        """ย้อนกลับไปหน้าก่อนหน้า"""
        if self.current_page > 0:
            self.current_page -= 1
            self._populate_pending_list(self.filtered_df)

    def _load_data(self):
        # แก้ไขให้เรียก _load_pending_pos โดยตรงเมื่อ Refresh
        self._update_manager_dashboard()
        self._load_pending_pos()

    def _so_create_string_vars(self):
        """(เวอร์ชันแก้ไข) สร้าง StringVars ทั้งหมดที่ SOPopupWindow ต้องการ"""
        self.so_shared_vars = {}
        now = datetime.now()
        thai_months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        # --- START: เพิ่มโค้ดทั้งหมดนี้เข้าไป ---
        self.so_shared_vars['thai_months'] = thai_months_list
        self.so_shared_vars['thai_month_map'] = {name: i + 1 for i, name in enumerate(thai_months_list)}
        self.so_shared_vars['customer_type_var'] = tk.StringVar(value="ลูกค้าเก่า")
        self.so_shared_vars['credit_term_var'] = tk.StringVar(value="เงินสด")
        self.so_shared_vars['commission_month_var'] = tk.StringVar(value=thai_months_list[now.month - 1])
        self.so_shared_vars['commission_year_var'] = tk.StringVar(value=str(now.year + 543))
        self.so_shared_vars['payment1_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_shared_vars['payment2_percent_var'] = tk.StringVar(value="ระบุยอดเอง")
        self.so_shared_vars['payment_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_subtotal_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_vat_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_grand_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['so_vs_payment_result_var'] = tk.StringVar(value="-")
        self.so_shared_vars['difference_amount_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['balance_due_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_product_input_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_service_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_required_total_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_actual_payment_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cash_verification_result_var'] = tk.StringVar(value="-")
        
        # ตัวแปรสำหรับคำนวณ VAT ย่อย (ที่เป็นต้นเหตุของปัญหา)
        self.so_shared_vars['sales_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['cutting_drilling_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['other_service_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['shipping_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['card_fee_vat_calc_var'] = tk.StringVar(value="0.00")
        self.so_shared_vars['relocation_vat_calc_var'] = tk.StringVar(value="0.00")

        # ตัวแปรสำหรับตัวเลือก VAT/CASH
        self.so_shared_vars['sales_service_vat_option'] = tk.StringVar(value="VAT")
        self.so_shared_vars['cutting_drilling_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_shared_vars['other_service_fee_vat_option'] = tk.StringVar(value="VAT")
        self.so_shared_vars['shipping_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_shared_vars['credit_card_fee_vat_option_var'] = tk.StringVar(value="VAT")
        self.so_shared_vars['relocation_cost_vat_option'] = tk.StringVar(value="VAT")

        # ตัวแปรอื่นๆ
        self.so_shared_vars['delivery_type_var'] = tk.StringVar(value="ซัพพลายเออร์จัดส่ง")
    
    def _open_so_editor_for_mp(self, so_number):
        """เปิดหน้าต่างแก้ไข SO สำหรับ MP"""
        try:
            # ค้นหา SO ID จาก SO Number ก่อน
            so_df = pd.read_sql_query("SELECT * FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1", self.pg_engine, params=(so_number,))
            if so_df.empty:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบข้อมูล SO: {so_number}", parent=self)
                return
            
            # เปิดหน้าต่างแก้ไข SO
            SOPopupWindow(
                master=self, 
                app_container=self.app_container, 
                sales_data=so_df.iloc[0].to_dict(), 
                so_shared_vars=self.so_shared_vars, 
                sale_theme=self.app_container.THEME["sale"],
                on_save_callback=self._load_data # Refresh หน้าจอหลักหลังบันทึก
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างแก้ไข SO ได้: {e}", parent=self)
            traceback.print_exc()
    
    def _prepare_data_for_pdf(self, po_header_data, po_items_data, po_payments_data):
        """
        ฟังก์ชัน Helper สำหรับแปลงข้อมูล PO ให้อยู่ในรูปแบบที่ PDF Generator ต้องการ
        """
        # เริ่มต้นด้วยข้อมูล header ของ PO ทั้งหมด
        pdf_data = po_header_data.copy()

        # --- 1. แปลงข้อมูลการชำระเงิน (Payments) ---
        # ตั้งค่าเริ่มต้นให้เป็น 0.00 ทั้งหมด
        payment_keys = ['deposit_amount', 'balance_due_po', 'full_payment_amount', 'cn_refund_amount']
        for key in payment_keys:
            pdf_data[key] = 0.0
        
        # วนลูปเพื่อดึงข้อมูลจาก payments_data มาใส่ใน key ที่ถูกต้อง
        for payment in po_payments_data:
            amount = payment.get('amount', 0.0)
            payment_date = pd.to_datetime(payment.get('payment_date')).strftime('%Y-%m-%d') if pd.notna(payment.get('payment_date')) else ''
            
            if payment['payment_type'] == 'Payment 1':
                pdf_data['deposit_amount'] += amount
                pdf_data['deposit_date'] = payment_date
            elif payment['payment_type'] == 'Payment 2':
                pdf_data['deposit_amount'] += amount # ยอดมัดจำคือผลรวมของ Payment 1 และ 2
                pdf_data['deposit_date'] = payment_date # ใช้วันที่ของรายการล่าสุด
            elif payment['payment_type'] == 'Full Payment':
                pdf_data['full_payment_amount'] = amount
                pdf_data['full_payment_date'] = payment_date
            elif payment['payment_type'] == 'CN Refund':
                pdf_data['cn_refund_amount'] = amount
                pdf_data['cn_refund_date'] = payment_date
        
        # คำนวณยอดค้างชำระ (Balance Due)
        grand_total = pdf_data.get('grand_total', 0.0) or 0.0
        total_paid = pdf_data['deposit_amount'] + pdf_data['full_payment_amount']
        pdf_data['balance_due_po'] = grand_total - total_paid
        pdf_data['net_payable_po'] = grand_total - (pdf_data.get('wht_3_percent_po', 0.0) or 0.0)

        # --- 2. แปลงข้อมูลค่าจัดส่ง (Shipping) ---
        pdf_data['shipping_cost_1'] = pdf_data.get('shipping_to_stock_cost', 0.0)
        pdf_data['shipping_vat_type_1'] = pdf_data.get('shipping_to_stock_vat_type', 'CASH')
        pdf_data['shipper_1'] = pdf_data.get('shipping_to_stock_shipper', '')
        
        pdf_data['shipping_cost_2'] = pdf_data.get('shipping_to_site_cost', 0.0)
        pdf_data['shipping_vat_type_2'] = pdf_data.get('shipping_to_site_vat_type', 'CASH')
        pdf_data['shipper_2'] = pdf_data.get('shipping_to_site_shipper', '')
        
        # --- 3. แปลงข้อมูลผู้อนุมัติ (Approvers) ---
        # ชื่อผู้อนุมัติถูกดึงมาด้วยชื่อที่ถูกต้องแล้ว (approver_1, approver_2, approver_3)
        # แต่เราอาจจะต้องใส่ค่าว่างไว้เผื่อกรณีที่ยังไม่มีข้อมูล
        pdf_data['creator_user'] = pdf_data.get('user_name', '')
        pdf_data['approver_1'] = pdf_data.get('approver_1', '')
        pdf_data['approver_2'] = pdf_data.get('approver_2', '')
        pdf_data['approver_3'] = pdf_data.get('approver_3', '')

        return pdf_data

    def _print_selected_po_from_manager(self, po_id):
        """
        ฟังก์ชันสำหรับดึงข้อมูล PO และ SO ที่สมบูรณ์เพื่อส่งไปสร้าง PDF
        (ใช้ Logic เดียวกับ purchasing_screen.py)
        """
        try:
            # นี่คือ Query ที่ถูกต้องและสมบูรณ์ที่สุด
            query = """
                SELECT
                    po.*, c.*,
                    u_po.sale_name AS user_name,
                    u_so.sale_name AS sale_name,
                    m1.sale_name AS approver_1,
                    m2.sale_name AS approver_2,
                    d.sale_name AS approver_3
                FROM purchase_orders po
                LEFT JOIN commissions c ON po.so_number = c.so_number AND c.is_active = 1
                LEFT JOIN sales_users u_po ON po.user_key = u_po.sale_key
                LEFT JOIN sales_users u_so ON c.sale_key = u_so.sale_key
                LEFT JOIN sales_users m1 ON po.approver_manager1_key = m1.sale_key
                LEFT JOIN sales_users m2 ON po.approver_manager2_key = m2.sale_key
                LEFT JOIN sales_users d ON po.approver_director_key = d.sale_key
                WHERE po.id = %s LIMIT 1;
            """
            po_df = pd.read_sql_query(query, self.pg_engine, params=(po_id,))
            if po_df.empty:
                messagebox.showerror("Error", "ไม่พบข้อมูล PO ที่เลือก", parent=self)
                return

            header_data = po_df.iloc[0].to_dict()
            items_df = pd.read_sql_query("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
            payments_df = pd.read_sql_query("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))

            all_po_data = [{
                "header": header_data,
                "items": items_df.to_dict('records'),
                "payments": payments_df.to_dict('records')
            }]
            
            # ส่งข้อมูลที่ครบถ้วนไปให้ฟังก์ชันสร้าง PDF
            from po_document_generator import generate_multi_po_pdf
            generate_multi_po_pdf(so_header_data=header_data, all_po_data=all_po_data)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการเตรียมข้อมูลเพื่อพิมพ์: {e}", parent=self)
            traceback.print_exc()

    def _open_po_print_dialog(self):
        """
        (แก้ไข) เปิดหน้าต่างสำหรับเลือก SO เพื่อพิมพ์ PO ทั้งหมดที่เกี่ยวข้อง
        """
        try:
            from po_selection_dialog import SOSelectionPrintDialog # Import ซ้ำเพื่อความแน่นอน
            dialog = SOSelectionPrintDialog(
                master=self, 
                pg_engine=self.app_container.pg_engine, 
                print_callback=self._print_all_pos_for_so # <-- ส่ง callback ไปยังฟังก์ชันใหม่
            )
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถเปิดหน้าต่างเลือก SO ได้: {e}", parent=self)
            traceback.print_exc()
    
    def _print_all_pos_for_so(self, so_number):
        """
        (เวอร์ชันแก้ไข) รวบรวมข้อมูลทั้งหมดของ SO และ PO ที่เกี่ยวข้องเพื่อส่งไปสร้าง PDF
        """
        try:
            # 1. ดึงข้อมูล Header ของ SO (เหมือนเดิม)
            so_query = """
                SELECT c.*, u_so.sale_name AS sale_name
                FROM commissions c
                LEFT JOIN sales_users u_so ON c.sale_key = u_so.sale_key
                WHERE c.so_number = %s AND c.is_active = 1 LIMIT 1;
            """
            so_header_df = pd.read_sql_query(so_query, self.pg_engine, params=(so_number,))
            if so_header_df.empty:
                messagebox.showerror("Error", f"ไม่พบข้อมูล SO: {so_number}", parent=self)
                return
            so_header_data = so_header_df.iloc[0].to_dict()

            # 2. ดึงข้อมูล PO ทั้งหมดที่เกี่ยวข้อง (เหมือนเดิม)
            po_query = """
                SELECT po.*, u_po.sale_name AS user_name,
                    m1.sale_name AS approver_1, m2.sale_name AS approver_2, d.sale_name AS approver_3
                FROM purchase_orders po
                LEFT JOIN sales_users u_po ON po.user_key = u_po.sale_key
                LEFT JOIN sales_users m1 ON po.approver_manager1_key = m1.sale_key
                LEFT JOIN sales_users m2 ON po.approver_manager2_key = m2.sale_key
                LEFT JOIN sales_users d ON po.approver_director_key = d.sale_key
                WHERE po.so_number = %s AND po.status = 'Approved';
            """
            all_po_df = pd.read_sql_query(po_query, self.pg_engine, params=(so_number,))
            
            # 3. เตรียมข้อมูลสำหรับส่งไปสร้าง PDF (*** ส่วนที่แก้ไข ***)
            all_po_data_list = []
            for _, po_row in all_po_df.iterrows():
                po_id = po_row['id']
                items_df = pd.read_sql("SELECT * FROM purchase_order_items WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
                payments_df = pd.read_sql("SELECT * FROM purchase_order_payments WHERE purchase_order_id = %s ORDER BY id", self.pg_engine, params=(po_id,))
                
                # *** เรียกใช้ฟังก์ชันแปลงข้อมูลตรงนี้ ***
                prepared_header = self._prepare_data_for_pdf(
                    po_row.to_dict(), 
                    items_df.to_dict('records'), 
                    payments_df.to_dict('records')
                )
                
                all_po_data_list.append({
                    "header": prepared_header, # <-- ใช้ข้อมูลที่ผ่านการแปลงแล้ว
                    "items": items_df.to_dict('records'),
                    "payments": payments_df.to_dict('records') # ส่งไปเผื่อ แต่ปัจจุบันไม่ได้ใช้
                })
            
            # 4. เรียกใช้ฟังก์ชันสร้าง PDF (เหมือนเดิม)
            from po_document_generator import generate_multi_po_pdf
            generate_multi_po_pdf(so_header_data=so_header_data, all_po_data=all_po_data_list)

        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการเตรียมข้อมูลเพื่อพิมพ์: {e}", parent=self)
            traceback.print_exc()
   
    def _open_approved_po_history(self):
        """เปิดหน้าต่างประวัติ PO ที่อนุมัติแล้วโดยตรง"""
        try:
            PurchaseHistoryWindow(master=self, app_container=self.app_container)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดหน้าต่างประวัติได้: {e}", parent=self)    

    def _create_header(self):
        header_frame = CTkFrame(self, fg_color="transparent"); header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(10,0))
        CTkLabel(header_frame, text=f"หน้าจอหัวหน้าฝ่ายจัดซื้อ: {self.user_name}", font=CTkFont(size=22, weight="bold"), text_color=self.theme["header"]).pack(side="left")
        button_container = CTkFrame(header_frame, fg_color="transparent"); button_container.pack(side="right")
        self.approve_all_button = CTkButton(button_container, 
                                            text="อนุมัติทุกรายการที่ค้างอยู่ (0)", 
                                            command=self._approve_all_pending_pos,
                                            fg_color="#84CC16", # สีเขียวมะนาว
                                            hover_color="#65A30D",
                                            state="disabled") # เริ่มต้นที่ปิดใช้งาน
        self.approve_all_button.pack(side="left", padx=10)
        CTkButton(button_container, text="📄 พิมพ์ใบสั่งซื้อ PO", command=self._open_po_print_dialog, fg_color="#7C3AED", hover_color="#6D28D9").pack(side="left", padx=10)
        CTkButton(button_container, text="ดึงงาน PO กลับมาแก้ไข", command=self._open_reopen_po_window, fg_color="#F97316", hover_color="#EA580C").pack(side="left", padx=10)
        CTkButton(button_container, text="ประวัติการตีกลับ", command=self._open_rejection_history, fg_color="#EF4444", hover_color="#B91C1C").pack(side="left", padx=10)
        CTkButton(button_container, text="ดูประวัติ PO ที่อนุมัติแล้ว", command=self._open_approved_po_history).pack(side="left", padx=(0, 10))
        CTkButton(button_container, text="Refresh All", command=self._load_data).pack(side="left", padx=10)
        CTkButton(button_container, text="Export PDF (PO อนุมัติ)", command=lambda: export_approved_pos_to_pdf(self, self.pg_engine), fg_color="#c026d3", hover_color="#a21caf").pack(side="left", padx=5)
        CTkButton(button_container, text="Export Excel (PO อนุมัติ)", command=lambda: export_approved_pos_to_excel(self, self.pg_engine), fg_color="#107C41", hover_color="#0B532B").pack(side="left", padx=5)   
        CTkButton(button_container, text="ออกจากระบบ", command=self.app_container.show_login_screen, fg_color="transparent", border_color="#D32F2F", text_color="#D32F2F", border_width=2, hover_color="#FFEBEE").pack(side="left")
       
    def _open_so_detail_window(self, so_number):
        if self.so_detail_window is None or not self.so_detail_window.winfo_exists():
            self.so_detail_window = SOPendingDetailWindow(self, so_number)
        else:
            self.so_detail_window.focus()
            
    def _open_reopen_po_window(self):
        if self.reopen_window is None or not self.reopen_window.winfo_exists():
         self.reopen_window = ReopenPOWindow(self) 
        else:
         self.reopen_window.focus()

    def _create_so_group_card(self, row_data):
        """(เวอร์ชันแก้ไข) สร้างการ์ด SO พร้อมแสดงชื่อเจ้าของ"""
        so_number = row_data['so_number']
        po_count = row_data['po_count']
        so_owner = row_data.get('so_owner', 'Unknown') # ดึงชื่อเจ้าของ

        card = CTkFrame(self.main_frame, border_width=1, corner_radius=10, fg_color="#F9FAFB")
        header = CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=10)
        header.grid_columnconfigure(0, weight=1)

        # แก้ไข: เพิ่มชื่อเจ้าของในวงเล็บ
        display_text = f"SO: {so_number} (เจ้าของ: {so_owner}) | มี {po_count} POs รออนุมัติ"
        
        CTkLabel(header, text=display_text, font=self.header_font).grid(row=0, column=0, sticky="w")
        
        action_frame = CTkFrame(header, fg_color="transparent")
        action_frame.grid(row=0, column=1, sticky="e")
        
        detail_frame = CTkFrame(card, fg_color="transparent")
        
        approve_all_button = CTkButton(action_frame, 
                                    text=f"อนุมัติทั้งหมด ({po_count})", 
                                    command=lambda s=so_number: self._approve_all_for_so(s),
                                    fg_color="#16A34A",
                                    hover_color="#15803D")
        approve_all_button.pack(side="right", padx=(10,0))

        edit_so_button = CTkButton(action_frame, 
                                text="ดู/แก้ไขข้อมูล SO", 
                                command=lambda s=so_number: self._open_so_editor_for_mp(s),
                                fg_color="#3B82F6",
                                hover_color="#2563EB")
        edit_so_button.pack(side="right", padx=5)
        
        CTkButton(action_frame, text="ดูสรุปรายการสินค้า", command=lambda s=so_number: self._open_so_detail_window(s)).pack(side="right", padx=5)
        CTkButton(action_frame, text="แสดง/ซ่อน PO ย่อย", width=150, command=lambda s=so_number, df=detail_frame: self._toggle_po_details(s, df)).pack(side="right", padx=5)
        
        return card
    
    def _approve_po(self, po_id, confirm=True):
        """(เวอร์ชันแก้ไข) อนุมัติ PO ในขั้นตอนเดียว"""
        
        if confirm and not messagebox.askyesno("ยืนยันการอนุมัติ", f"คุณต้องการอนุมัติ PO ID: {po_id} ใช่หรือไม่?", parent=self):
            return
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # ดึงข้อมูล SO Number และสถานะปัจจุบัน
                cursor.execute("SELECT so_number, approval_status FROM purchase_orders WHERE id = %s", (po_id,))
                po = cursor.fetchone()
                if not po:
                    messagebox.showerror("ผิดพลาด", "ไม่พบ PO ที่ต้องการอนุมัติ", parent=self)
                    return
                
                so_number = po['so_number']
                current_status = po['approval_status']

                # --- START: แก้ไข Logic การอนุมัติให้ง่ายขึ้น ---
                # ตรวจสอบว่า PO อยู่ในสถานะที่อนุมัติได้หรือไม่
                allowed_statuses = ['Pending Mgr 1']
                if self.user_role == 'Director':
                    allowed_statuses.append('Pending Director')

                if current_status not in allowed_statuses:
                    messagebox.showinfo("ข้อมูลล่าสุด", f"PO นี้ไม่อยู่ในสถานะที่รออนุมัติจากคุณ (สถานะปัจจุบัน: {current_status})", parent=self)
                    self._load_data()
                    return

                # อัปเดตสถานะเป็น Approved ทันที
                set_clauses = "status = %s, approval_status = %s, last_modified_by = %s"
                params = ["Approved", "Approved", self.user_key]

                if self.user_role == 'Director':
                    set_clauses += ", approver_director_key = %s, approval_date_director = %s"
                    params.extend([self.user_key, datetime.now()])
                else: # Purchasing Manager
                    set_clauses += ", approver_manager1_key = %s, approval_date_manager1 = %s"
                    params.extend([self.user_key, datetime.now()])
                
                params.append(po_id)
                sql_query = f"UPDATE purchase_orders SET {set_clauses} WHERE id = %s"
                cursor.execute(sql_query, tuple(params))
                # --- END ---

            conn.commit()
            
            if confirm:
                messagebox.showinfo("สำเร็จ", "อนุมัติรายการเรียบร้อยแล้ว", parent=self)
            
            self._update_ui_after_action(po_id, so_number)
            
            # <<< START: แก้ไขชื่อฟังก์ชันที่เรียกผิดตรงนี้ >>>
            # ตรวจสอบว่า SO นี้ควรส่งต่อไปให้ Sale Manager หรือยัง
            if so_number:
                self._check_and_forward_so_to_sale_manager(so_number) 
            # <<< END >>>

        except Exception as e:
            if conn: conn.rollback()
            if confirm: messagebox.showerror("ผิดพลาด", f"ไม่สามารถอนุมัติได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)
              

    def _check_and_forward_so_to_sale_manager(self, so_number):
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                # ตรวจสอบจำนวน PO ทั้งหมดและที่อนุมัติแล้ว (เหมือนเดิม)
                cursor.execute("""
                    SELECT COUNT(id) FROM purchase_orders 
                    WHERE so_number = %s AND status NOT IN ('Draft', 'Rejected')
                """, (so_number,))
                total_pos = cursor.fetchone()[0]

                cursor.execute("""
                    SELECT COUNT(id) FROM purchase_orders 
                    WHERE so_number = %s AND status = 'Approved'
                """, (so_number,))
                approved_pos = cursor.fetchone()[0]

                # เงื่อนไข: ถ้า PO ทุกใบของ SO นี้ถูกอนุมัติครบแล้ว
                if total_pos > 0 and total_pos == approved_pos:
                    print(f"All POs for SO {so_number} are approved. Forwarding to HR.")
                    
                    # <<< START: จุดแก้ไขสำคัญ >>>
                    # เปลี่ยนสถานะ SO เป็น 'PO Sent' เพื่อส่งไปให้ HR
                    new_so_status = 'PO Sent'
                    cursor.execute("""
                        UPDATE commissions SET status = %s 
                        WHERE so_number = %s AND is_active = 1
                    """, (new_so_status, so_number))
                    # <<< END: สิ้นสุดการแก้ไข >>>

                    # สร้าง Notification แจ้งเตือนฝ่าย HR
                    cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'HR' AND status = 'Active'")
                    hr_keys = [row[0] for row in cursor.fetchall()]
                    
                    message = f"SO: {so_number} มี PO ที่อนุมัติครบแล้ว รอการตรวจสอบจากท่าน"
                    for hr_key in hr_keys:
                        cursor.execute("""
                            INSERT INTO notifications (user_key_to_notify, message, is_read) 
                            VALUES (%s, %s, FALSE)
                        """, (hr_key, message))
                    
                    conn.commit()
                else:
                    print(f"SO {so_number} still has pending POs ({approved_pos}/{total_pos} approved). Waiting for completion.")
                    
        except Exception as e:
            print(f"Error in _check_and_forward_so_to_hr: {e}") 
            if conn: conn.rollback()
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _create_approval_notification(self, cursor, po_id, next_status):
        cursor.execute("SELECT po_number, approver_manager1_key, approver_manager2_key FROM purchase_orders WHERE id = %s", (po_id,))
        po_info = cursor.fetchone()
        if not po_info: return
        po_number = po_info['po_number']
        
        message = ""
        user_keys_to_notify = []

        if next_status == 'Approved':
            cursor.execute("SELECT user_key FROM purchase_orders WHERE id = %s", (po_id,))
            user_keys_to_notify = [row['user_key'] for row in cursor.fetchall()]
            message = f"PO ของคุณ ({po_number}) ได้รับการอนุมัติครบถ้วนแล้ว"
        
        elif next_status == 'Pending Mgr 2':
            cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Purchasing Manager' AND status = 'Active' AND sale_key != %s", (self.user_key,))
            user_keys_to_notify = [row['sale_key'] for row in cursor.fetchall()]
            message = f"PO ({po_number}) รอการอนุมัติจากผู้จัดการคนที่ 2"

        elif next_status == 'Pending Director':
            cursor.execute("SELECT sale_key FROM sales_users WHERE role = 'Director' AND status = 'Active'")
            user_keys_to_notify = [row['sale_key'] for row in cursor.fetchall()]
            message = f"PO ({po_number}) ยอดสูง รอการอนุมัติจากท่าน"

        for user_key in user_keys_to_notify:
            cursor.execute("INSERT INTO notifications (user_key_to_notify, message, related_po_id) VALUES (%s, %s, %s)", (user_key, message, po_id))

    def _reject_po(self, po_id):
        dialog = RejectionReasonDialog(self)
        self.wait_window(dialog)
        reason = getattr(dialog, '_reason_string', None)
        if reason is None:
            return

        conn = self.app_container.get_connection()
        try:
            # <<< START: จุดที่แก้ไข >>>
            so_number_to_update = "" # สร้างตัวแปรเปล่าไว้ก่อน
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                
                # 1. ดึงข้อมูล SO Number มาเก็บไว้ก่อนที่จะทำ Action
                cursor.execute("SELECT user_key, po_number, so_number FROM purchase_orders WHERE id = %s", (po_id,))
                po_info = cursor.fetchone()
                if not po_info:
                    messagebox.showerror("ผิดพลาด", "ไม่พบ PO ที่ต้องการปฏิเสธ", parent=self)
                    return
                
                po_creator_key, po_number, so_number_to_update = po_info['user_key'], po_info['po_number'], po_info['so_number']

                # 2. อัปเดตสถานะ PO ให้เป็น 'Rejected' และบันทึกเหตุผล (เหมือนเดิม)
                cursor.execute("""
                    UPDATE purchase_orders 
                    SET status = 'Rejected', approval_status = 'Rejected', rejection_reason = %s, last_modified_by = %s
                    WHERE id = %s
                """, (reason.strip(), self.user_key, po_id))

                # 3. สร้าง Notification (เหมือนเดิม)
                message_to_pu = f"PO: {po_number} ของคุณถูกปฏิเสธโดย Manager\nเหตุผล: {reason.strip()}"
                cursor.execute("INSERT INTO notifications (user_key_to_notify, message, is_read, related_po_id) VALUES (%s, %s, FALSE, %s)", 
                               (po_creator_key, message_to_pu, po_id))

                # 4. บันทึก Audit Log (เหมือนเดิม)
                log_details = {'rejected_by': self.user_key, 'reason': reason.strip()}
                cursor.execute("INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp) VALUES (%s, %s, %s, %s, %s, %s)", 
                               ('PO Rejected', 'purchase_orders', po_id, po_creator_key, json.dumps(log_details), datetime.now()))

            conn.commit()
            messagebox.showinfo("สำเร็จ", "ปฏิเสธ PO และส่งกลับให้ฝ่ายจัดซื้อเรียบร้อยแล้ว", parent=self)
            
            # 5. เรียกใช้ฟังก์ชันอัปเดต UI โดยใช้ตัวแปรที่เก็บไว้
            self._update_ui_after_action(po_id, so_number_to_update)
            # <<< END: สิ้นสุดการแก้ไข >>>

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("Database Error", f"เกิดข้อผิดพลาดในการปฏิเสธ PO: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _update_ui_after_action(self, po_id, so_number):
        """
        (เวอร์ชันแก้ไข) อัปเดต UI อย่างชาญฉลาดหลังจากอนุมัติหรือปฏิเสธ PO
        """
        # 1. หา PO card ที่เพิ่งกระทำไปแล้วลบออกจากหน้าจอ
        if po_id in self.po_cards:
            po_card_widget = self.po_cards.pop(po_id)
            if po_card_widget.winfo_exists():
                po_card_widget.destroy()

        # 2. อัปเดต DataFrame หลัก โดยการลบแถวของ PO ที่เพิ่งกระทำออกไป
        self.all_pending_df = self.all_pending_df[self.all_pending_df['id'] != po_id].copy()

        # 3. อัปเดตการ์ด SO และปุ่มสรุปทั้งหมด
        self._update_all_counts(so_number_to_update=so_number)

    def _update_all_counts(self, so_number_to_update=None):
        """
        (เวอร์ชันแก้ไข) อัปเดต SO Card ที่เกี่ยวข้อง และปุ่มสรุปด้านบน
        """
        if so_number_to_update:
            so_card_widget = self.so_cards.get(so_number_to_update)
            if so_card_widget and so_card_widget.winfo_exists():
                # นับจำนวน PO ที่เหลืออยู่ของ SO นี้จาก DataFrame
                remaining_count = len(self.all_pending_df[self.all_pending_df['so_number'] == so_number_to_update])
                
                if remaining_count > 0:
                    # ยังมี PO เหลืออยู่ ให้อัปเดตข้อความ
                    label = so_card_widget.winfo_children()[0].winfo_children()[0]
                    label.configure(text=f"SO: {so_number_to_update} (มี {remaining_count} POs รออนุมัติ)")
                    
                    # อัปเดตปุ่ม "อนุมัติทั้งหมด" ของ SO card นั้นๆ
                    action_frame = so_card_widget.winfo_children()[0].winfo_children()[1]
                    approve_all_button = action_frame.winfo_children()[0]
                    approve_all_button.configure(text=f"อนุมัติทั้งหมด ({remaining_count})")
                else:
                    # ไม่มี PO เหลือแล้ว ลบ SO Card ทิ้ง
                    so_card_widget.destroy()
                    self.so_cards.pop(so_number_to_update, None)
        
        # อัปเดตปุ่มใหญ่ด้านบนสุดเสมอ
        total_pending_count = len(self.all_pending_df)
        if hasattr(self, 'approve_all_button') and self.approve_all_button.winfo_exists():
            self.approve_all_button.configure(text=f"อนุมัติทุกรายการที่ค้างอยู่ ({total_pending_count})")
            self.approve_all_button.configure(state="normal" if total_pending_count > 0 else "disabled")

    def _on_destroy(self, event):
        if hasattr(event, 'widget') and event.widget is self: self._stop_polling()
        
    def _start_polling(self):
        self._stop_polling()
        self.polling_job_id = self.after(300000, self._perform_polling)
        
    def _stop_polling(self):
        if self.polling_job_id: self.after_cancel(self.polling_job_id); self.polling_job_id = None
        
    def _perform_polling(self):
        self._load_pending_pos()
        self.polling_job_id = self.after(300000, self._perform_polling) #<-- เปลี่ยนเป็น 300000 (5 นาที)
        
    def _create_dashboard_view(self, parent_tab):
        dashboard_frame = CTkFrame(parent_tab, corner_radius=10, border_width=1) # <-- ใช้ parent_tab
        dashboard_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # --- Frame สำหรับฟิลเตอร์ ---
        filter_container = CTkFrame(dashboard_frame, fg_color="transparent")
        filter_container.pack(fill="x", padx=10, pady=5, anchor="nw")

        # --- เตรียมข้อมูลสำหรับ Dropdown ---
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i + 1 for i, name in enumerate(self.thai_months)}
        current_year = datetime.now().year
        
        # --- Dropdown เลือกเดือน ---
        self.month_var = tk.StringVar(value="ทุกเดือน")
        month_options = ["ทุกเดือน"] + self.thai_months
        CTkLabel(filter_container, text="เดือน:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_container, variable=self.month_var, values=month_options).pack(side="left", padx=(0, 10))

        # --- Dropdown เลือกปี ---
        self.year_var = tk.StringVar(value=str(current_year))
        year_options = [str(y) for y in range(current_year, current_year - 5, -1)]
        CTkLabel(filter_container, text="ปี:").pack(side="left", padx=(10, 2))
        CTkOptionMenu(filter_container, variable=self.year_var, values=year_options).pack(side="left", padx=(0, 10))

        # --- ปุ่มสำหรับกดค้นหา ---
        CTkButton(filter_container, text="แสดงผล", command=self._update_manager_dashboard).pack(side="left", padx=10)

        # --- Frame สำหรับแสดงกราฟ ---
        self.rejection_chart_frame = CTkFrame(dashboard_frame)
        self.rejection_chart_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def _update_manager_dashboard(self):
        for widget in self.rejection_chart_frame.winfo_children():
            widget.destroy()
        loading_label = CTkLabel(self.rejection_chart_frame, text="กำลังโหลดข้อมูล Dashboard...", font=CTkFont(size=18, slant="italic"), text_color="gray50")
        loading_label.pack(expand=True, pady=20)
        self.update_idletasks()
        
        try:
            # --- ดึงค่าจาก Dropdown ---
            selected_year_str = self.year_var.get()
            selected_month_str = self.month_var.get()

            # --- แปลงค่าเป็นตัวเลขสำหรับส่งเข้า Query ---
            year_to_query = int(selected_year_str)
            month_to_query = self.thai_month_map.get(selected_month_str, None) # ถ้าเลือก "ทุกเดือน" จะได้ None

            # --- เรียกใช้ฟังก์ชันดึงข้อมูลพร้อมกับส่งค่าที่เลือก ---
            rejection_data = self._get_rejection_summary(year=year_to_query, month=month_to_query)
            
            # --- สร้างกราฟจากข้อมูลที่ได้มา ---
            self._create_rejection_bar_chart(self.rejection_chart_frame, rejection_data)

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการอัปเดต Dashboard: {e}", parent=self)
        finally:
            if loading_label.winfo_exists():
                loading_label.destroy()
                
    def _get_rejection_summary(self, year, month):
        try:
            # Query ใหม่เพื่อนับจาก audit_log
            sql_where_clause = "WHERE log.action = 'PO Rejected' AND log.table_name = 'purchase_orders'"
            params = []

            if year:
                sql_where_clause += " AND EXTRACT(YEAR FROM log.timestamp::timestamp) = %s"
                params.append(year)
            
            if month:
                sql_where_clause += " AND EXTRACT(MONTH FROM log.timestamp::timestamp) = %s"
                params.append(month)

            # เราจะ JOIN กับ sales_users เพื่อเอาชื่อของ PU ที่เป็นคนทำ PO (เก็บไว้ใน user_info)
            query = f"""
                SELECT 
                    su.sale_name, 
                    COUNT(log.id) as rejection_count 
                FROM audit_log log
                JOIN sales_users su ON log.user_info = su.sale_key 
                {sql_where_clause}
                GROUP BY su.sale_name 
                ORDER BY rejection_count DESC
            """
            return pd.read_sql_query(query, self.pg_engine, params=tuple(params))

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถดึงข้อมูลสรุปการตีกลับได้: {e}", parent=self)
            return pd.DataFrame(columns=['sale_name', 'rejection_count'])
        
    def _create_rejection_bar_chart(self, parent_frame, data_df):
        if hasattr(self, 'rejection_chart_canvas') and self.rejection_chart_canvas: self.rejection_chart_canvas.get_tk_widget().destroy()
        for widget in parent_frame.winfo_children(): widget.destroy()
        if data_df.empty: CTkLabel(parent_frame, text="ไม่พบข้อมูลการตีกลับ", font=self.header_font).pack(expand=True, pady=20); return
        fig = Figure(figsize=(8, 4), dpi=100, facecolor=self.theme["bg"]); ax = fig.add_subplot(111); ax.set_facecolor(self.theme["bg"])
        colors = ['#e76f51', '#f4a261', '#e9c46a', '#2a9d8f', '#264653']; bar_colors = [colors[i % len(colors)] for i in range(len(data_df))]
        bars = ax.barh(data_df['sale_name'], data_df['rejection_count'], color=bar_colors); ax.invert_yaxis(); font_name = 'Tahoma'
        ax.set_xlabel('จำนวนครั้งที่ถูกตีกลับ', fontname=font_name, fontsize=12); ax.set_title('สรุปสถิติการตีกลับงาน (Rejected POs)', fontname=font_name, fontsize=16, weight="bold")
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.tick_params(axis='y', labelsize=12, labelfontfamily=font_name)
        ax.xaxis.set_major_locator(MaxNLocator(integer=True)); ax.set_xlim(left=0)
        for bar in bars:
            width = bar.get_width(); ax.text(width + 0.1, bar.get_y() + bar.get_height()/2, f'{int(width)}', va='center', fontname=font_name)
        fig.tight_layout(pad=2); canvas = FigureCanvasTkAgg(fig, master=parent_frame); canvas.draw(); canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10); self.rejection_chart_canvas = canvas
        
    def _create_pending_list_view(self, parent_tab):
        container = CTkFrame(parent_tab) # <-- ใช้ parent_tab
        container.grid(row=1, column=0, padx=10, pady=10, sticky="nsew") # <-- แก้ไข row เป็น 1
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(2, weight=1)

        # --- ส่วนของช่องค้นหา ---
        search_frame = CTkFrame(container, fg_color="transparent")
        search_frame.grid(row=0, column=0, padx=10, pady=(5, 5), sticky="ew")
        
        self.search_entry = CTkEntry(search_frame, placeholder_text="🔍 ค้นหาจาก SO, PO, หรือชื่อซัพพลายเออร์...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # --- เพิ่ม 2 บรรทัดนี้ ---
        self.search_entry.bind("<Return>", self._filter_pending_list)
        self.search_entry.bind("<KP_Enter>", self._filter_pending_list) # สำหรับ Enter บน Numpad
        
        search_button = CTkButton(search_frame, text="ค้นหา", width=100, command=self._filter_pending_list)
        search_button.pack(side="left", padx=(0, 5))
        
        clear_button = CTkButton(search_frame, text="ล้างค่า", width=100, fg_color="gray", command=self._clear_search)
        clear_button.pack(side="left")

        # --- ส่วนของ Pagination ---
        pagination_frame = CTkFrame(container, fg_color="transparent")
        pagination_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        self.prev_button = CTkButton(pagination_frame, text="<< หน้าก่อนหน้า", command=self._prev_page, state="disabled")
        self.prev_button.pack(side="left")

        self.page_label = CTkLabel(pagination_frame, text="Page 1 / 1")
        self.page_label.pack(side="left", expand=True)

        self.next_button = CTkButton(pagination_frame, text="หน้าถัดไป >>", command=self._next_page, state="disabled")
        self.next_button.pack(side="right")
        
        # --- ส่วนของ Scrollable Frame ---
        self.main_frame = CTkScrollableFrame(container, label_text="รายการที่รอการอนุมัติ (Grouped by SO)")
        self.main_frame.grid(row=2, column=0, padx=0, pady=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        
    def _load_pending_pos(self):
        """(เวอร์ชันแก้ไข) โหลดข้อมูล PO พร้อมชื่อเจ้าของ SO"""
        try:
            # Base Query: เพิ่มการ JOIN ไปหา sales_users เพื่อเอาชื่อเจ้าของ SO
            base_query = """
                SELECT 
                    po.id, po.timestamp, po.user_key, po.so_number, po.po_number, 
                    po.supplier_name, po.grand_total, po.approval_status, po.approver_manager1_key,
                    u.sale_name AS so_owner
                FROM purchase_orders po
                LEFT JOIN commissions c ON po.so_number = c.so_number AND c.is_active = 1
                LEFT JOIN sales_users u ON c.sale_key = u.sale_key
            """
            
            where_clause = ""
            
            if self.user_role == 'Purchasing Manager':
                where_clause = "WHERE po.status = 'Pending Approval' AND po.approval_status = 'Pending Mgr 1'"
            elif self.user_role == 'Director':
                where_clause = "WHERE po.status = 'Pending Approval' AND po.approval_status = 'Pending Director'"
            
            if where_clause:
                final_query = f"{base_query} {where_clause} ORDER BY po.timestamp ASC"
                self.all_pending_df = pd.read_sql_query(final_query, self.app_container.pg_engine)
            else:
                self.all_pending_df = pd.DataFrame()

            self._filter_pending_list()

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล PO ที่รออนุมัติได้: {e}", parent=self)
            self.all_pending_df = pd.DataFrame()
            self._populate_pending_list(self.all_pending_df)    
            
    def _populate_pending_list(self, df_to_show):
        """(เวอร์ชันแก้ไข) แสดงรายการโดยจัดกลุ่มตาม SO และแสดงชื่อเจ้าของ"""
        for widget in self.main_frame.winfo_children(): widget.destroy()
        self.so_cards.clear(); self.po_cards.clear()

        if df_to_show.empty:
            CTkLabel(self.main_frame, text="ไม่พบรายการที่รอการอนุมัติ").pack(pady=20)
            self.approve_all_button.configure(text="อนุมัติทุกรายการที่ค้างอยู่ (0)", state="disabled")
            self.page_label.configure(text="Page 1 / 1")
            self.prev_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            return
        
        # จัดการค่า NaN ใน so_owner ก่อน Group
        df_display = df_to_show.copy() 

        if 'so_owner' in df_display.columns:
            df_display['so_owner'] = df_display['so_owner'].fillna('Unknown')
        else:
            df_display['so_owner'] = 'Unknown'
        
        # --- Logic การแบ่งหน้า (Group โดย SO และ Owner) ---
        # แก้ไข: Group โดยทั้ง so_number และ so_owner เพื่อให้ดึงค่า so_owner มาใช้ได้
        grouped_so = df_display.groupby(['so_number', 'so_owner'], sort=False).size().reset_index(name='po_count')
        
        total_groups = len(grouped_so)
        total_pages = (total_groups + self.rows_per_page - 1) // self.rows_per_page
        start_index = self.current_page * self.rows_per_page
        end_index = start_index + self.rows_per_page
        
        # --- แสดงผลเฉพาะ SO ในหน้าปัจจุบัน ---
        for _, group_row in grouped_so.iloc[start_index:end_index].iterrows():
            # ส่งข้อมูลทั้งแถว (มี so_number, so_owner, po_count) ไปให้ฟังก์ชันสร้างการ์ด
            so_card = self._create_so_group_card(group_row)
            so_card.pack(fill="x", padx=10, pady=(10, 5))
            self.so_cards[group_row['so_number']] = so_card

        # --- อัปเดต UI (Pagination และปุ่มอนุมัติทั้งหมด) ---
        total_pending_count = len(df_to_show)
        self.approve_all_button.configure(text=f"อนุมัติทุกรายการที่ค้างอยู่ ({total_pending_count})", state="normal" if total_pending_count > 0 else "disabled")
        self.page_label.configure(text=f"Page {self.current_page + 1} / {max(1, total_pages)}")
        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.configure(state="normal" if self.current_page < total_pages - 1 else "disabled")
              
    def _filter_pending_list(self, event=None):
        """(เวอร์ชันแก้ไข) กรองข้อมูล, รีเซ็ตหน้าเป็น 0, และวาด UI ใหม่"""
        search_term = self.search_entry.get().lower().strip()
        
        if self.all_pending_df is None:
            self.filtered_df = pd.DataFrame()
        elif not search_term:
            self.filtered_df = self.all_pending_df
        else:
            self.filtered_df = self.all_pending_df[
                self.all_pending_df['so_number'].str.lower().str.contains(search_term, na=False) |
                self.all_pending_df['po_number'].str.lower().str.contains(search_term, na=False) |
                self.all_pending_df['supplier_name'].str.lower().str.contains(search_term, na=False)
            ]
        
        # เมื่อมีการค้นหา ให้กลับไปที่หน้าแรกเสมอ
        self.current_page = 0
        self._populate_pending_list(self.filtered_df)

    def _toggle_po_details(self, so_number, detail_frame):
        # (ฟังก์ชันนี้แก้ไขเล็กน้อยเพื่อใช้ DataFrame ที่เก็บไว้)
        if detail_frame.winfo_viewable():
           detail_frame.pack_forget()
           return
    
        detail_frame.pack(fill="x", padx=10, pady=(0, 10))
        if not detail_frame.winfo_children():
            try:
                # กรอง PO เฉพาะของ SO นี้จาก DataFrame หลัก
                df_po = self.all_pending_df[self.all_pending_df['so_number'] == so_number]
                
                for _, row in df_po.iterrows():
                    po_id = row['id']
                    po_card = self._create_po_card_widget(detail_frame, row)
                    po_card.pack(fill="x", padx=20, pady=(2,5))
                    # เก็บ PO Card ไว้ใน Dictionary โดยใช้ po_id เป็น key
                    self.po_cards[po_id] = po_card
            except Exception as e:
                CTkLabel(detail_frame, text=f"Error loading PO details: {e}").pack()
            
    def _create_po_card_widget(self, parent, row_data, from_detail_window=False):
        card = CTkFrame(parent, border_width=1, corner_radius=10)
        card.grid_columnconfigure(0, weight=3)
        card.grid_columnconfigure(1, weight=1)
        
        info_frame = CTkFrame(card, fg_color="transparent")
        info_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        status_color = "#FB923C" if row_data['approval_status'] == 'Pending Mgr 1' else "#FACC15" if row_data['approval_status'] == 'Pending Mgr 2' else "#A855F7"
        
        # <<< START: แก้ไขจุดนี้ >>>
        # ดึงค่า grand_total มาใช้แสดงผลแทน total_cost
        grand_total = row_data.get('grand_total', 0) or 0
        
        CTkLabel(info_frame, text=f"PO: {row_data['po_number']}", font=self.header_font).pack(anchor="w")
        CTkLabel(info_frame, text=f"Supplier: {row_data['supplier_name']} | ยอดรวม: {grand_total:,.2f} บาท").pack(anchor="w")
        # <<< END >>>
        
        CTkLabel(info_frame, text=f"Status: {row_data['approval_status']}", text_color=status_color, font=CTkFont(weight="bold")).pack(anchor="w")
        CTkLabel(info_frame, text=f"Submitted by: {row_data['user_key']} at {pd.to_datetime(row_data['timestamp']).strftime('%Y-%m-%d %H:%M')}").pack(anchor="w")
        
        action_frame = CTkFrame(card, fg_color="transparent")
        action_frame.grid(row=0, column=1, padx=10, pady=10, sticky="e")
        
        approve_cmd = lambda d=row_data['id']: self._approve_po(d)
        reject_cmd = lambda d=row_data['id']: self._reject_po(d)
        
        CTkButton(action_frame, text="ดูรายละเอียด", width=120, command=lambda d=row_data['id']: self._view_details(d)).pack(fill="x", pady=2)
        CTkButton(action_frame, text="อนุมัติ", width=120, fg_color="#16A34A", hover_color="#15803D", command=approve_cmd).pack(fill="x", pady=2)
        CTkButton(action_frame, text="ปฏิเสธ", width=120, fg_color="#DC2626", hover_color="#B91C1C", command=reject_cmd).pack(fill="x", pady=2)
        
        return card
        
    # (ในไฟล์ purchasing_manager_screen.py)
# ให้นำฟังก์ชันนี้ไปวางทับของเดิม

    # (ในไฟล์ purchasing_manager_screen.py)
# ให้นำฟังก์ชันนี้ไปวางทับของเดิม

    # purchasing_manager_screen.py

    def _view_details(self, po_id):
     try: 
        # แก้ไขโดยการส่ง on_save_callback ไปแทน เพื่อให้หน้าจอ refresh ตัวเองหลัง HR แก้ไขข้อมูล
        self.app_container.show_purchase_detail_window(
            purchase_id=po_id,
            on_save_callback=self._load_data # <--- แก้ไขชื่อ parameter เป็น on_save_callback
        )
     except Exception as e: 
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถเปิดดูรายละเอียดได้: {e}", parent=self)

    def _create_notification(self, cursor, po_id, action_type, reason=""):
        try:
            cursor.execute("SELECT user_key, po_number FROM purchase_orders WHERE id = %s", (po_id,))
            po_info = cursor.fetchone()
            if po_info:
                user_to_notify, po_number = po_info[0], po_info[1]
                message = f"PO ของคุณ ({po_number}) ได้รับการอนุมัติแล้ว" if action_type == 'Approved' else f"PO ของคุณ ({po_number}) ถูกปฏิเสธ\nเหตุผล: {reason}"
                cursor.execute("INSERT INTO notifications (user_key_to_notify, message, related_po_id) VALUES (%s, %s, %s)", (user_to_notify, message, po_id))
        except Exception as e: print(f"Error creating notification: {e}")
        
    def _approve_all_for_so(self, so_number):
        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการอนุมัติ PO ทุกใบสำหรับ SO: {so_number} ใช่หรือไม่?", parent=self):
            return
        
        conn = None
        approved_count = 0
        failed_pos = []
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. ดึง ID ของ PO ทั้งหมดที่รออนุมัติสำหรับ SO นี้
                # และสำหรับ Role ปัจจุบัน
                status_to_fetch = ()
                if self.user_role == 'Purchasing Manager':
                    status_to_fetch = ('Pending Mgr 1', 'Pending Mgr 2')
                elif self.user_role == 'Director':
                    status_to_fetch = ('Pending Director',)

                if not status_to_fetch:
                    messagebox.showerror("ผิดพลาด", "Role ของคุณไม่สามารถอนุมัติได้", parent=self)
                    return

                cursor.execute(
                    "SELECT id, approval_status, approver_manager1_key FROM purchase_orders WHERE so_number = %s AND approval_status IN %s",
                    (so_number, status_to_fetch)
                )
                po_list = cursor.fetchall()

                if not po_list:
                    messagebox.showinfo("ข้อมูลล่าสุด", "ไม่พบ PO ที่รอการอนุมัติสำหรับ SO นี้", parent=self)
                    return

                # 2. วนลูปเพื่ออนุมัติทีละใบ
                for po in po_list:
                    po_id = po['id']
                    # ตรวจสอบเงื่อนไขการอนุมัติซ้ำ (สำหรับ Mgr)
                    if self.user_role == 'Purchasing Manager' and po['approval_status'] == 'Pending Mgr 2' and po['approver_manager1_key'] == self.user_key:
                        failed_pos.append(po_id)
                        continue # ข้ามไปทำใบถัดไป

                    # สร้างการอนุมัติ (โค้ดส่วนนี้จะเหมือนใน _approve_po)
                    try:
                        self._approve_po(po_id, confirm=False) # ส่ง confirm=False เพื่อไม่ให้ถามซ้ำ
                        approved_count += 1
                    except Exception as e:
                        print(f"Failed to approve PO ID {po_id}: {e}")
                        failed_pos.append(po_id)
                
            # 3. แสดงผลสรุป
            success_message = f"อนุมัติ PO สำหรับ SO: {so_number} สำเร็จ {approved_count} รายการ"
            if failed_pos:
                success_message += f"\n\nเกิดข้อผิดพลาด {len(failed_pos)} รายการ (ID: {', '.join(map(str, failed_pos))})"
                messagebox.showwarning("อนุมัติสำเร็จบางส่วน", success_message, parent=self)
            else:
                messagebox.showinfo("สำเร็จ", success_message, parent=self)

            self._load_pending_pos() # Refresh หน้าจอหลัก

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถอนุมัติทั้งหมดได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)

    def _approve_all_pending_pos(self):
        # ดึงจำนวน PO ที่ค้างอยู่จากปุ่ม
        current_pending_text = self.approve_all_button.cget("text")
        # ใช้ regular expression เพื่อดึงตัวเลขออกจาก string เช่น "อนุมัติ (15)" -> "15"
        import re
        match = re.search(r'\((\d+)\)', current_pending_text)
        if not match: return
        
        pending_count = int(match.group(1))

        if not messagebox.askyesno("ยืนยัน", f"คุณต้องการอนุมัติ PO ที่ค้างอยู่ทั้งหมด {pending_count} รายการใช่หรือไม่?\n(ระบบจะอนุมัติเฉพาะรายการที่คุณมีสิทธิ์)", icon="question", parent=self):
            return

        conn = None
        approved_count = 0
        failed_pos = []
        try:
            conn = self.app_container.get_connection()
            with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cursor:
                # 1. ดึง ID ของ PO ทั้งหมดที่รออนุมัติสำหรับ Role ปัจจุบัน
                status_to_fetch = ()
                if self.user_role == 'Purchasing Manager':
                    status_to_fetch = ('Pending Mgr 1', 'Pending Mgr 2')
                elif self.user_role == 'Director':
                    status_to_fetch = ('Pending Director',)

                if not status_to_fetch:
                    messagebox.showerror("ผิดพลาด", "Role ของคุณไม่สามารถอนุมัติได้", parent=self)
                    return

                cursor.execute(
                    "SELECT id, approval_status, approver_manager1_key FROM purchase_orders WHERE approval_status IN %s",
                    (status_to_fetch,)
                )
                all_pending_pos = cursor.fetchall()
                
                # 2. วนลูปเพื่ออนุมัติทีละใบ
                for po in all_pending_pos:
                    po_id = po['id']
                    # ตรวจสอบเงื่อนไขการอนุมัติซ้ำ (สำหรับ Mgr)
                    if self.user_role == 'Purchasing Manager' and po['approval_status'] == 'Pending Mgr 2' and po['approver_manager1_key'] == self.user_key:
                        failed_pos.append(po_id)
                        continue

                    try:
                        self._approve_po(po_id, confirm=False)
                        approved_count += 1
                    except Exception as e:
                        print(f"Failed to bulk approve PO ID {po_id}: {e}")
                        failed_pos.append(po_id)
            
            # 3. แสดงผลสรุป
            success_message = f"อนุมัติ PO สำเร็จ {approved_count} รายการ"
            if failed_pos:
                success_message += f"\n\nเกิดข้อผิดพลาด/ข้ามการอนุมัติ {len(failed_pos)} รายการ (ID: {', '.join(map(str, failed_pos))})"
                messagebox.showwarning("อนุมัติสำเร็จบางส่วน", success_message, parent=self)
            else:
                messagebox.showinfo("สำเร็จ", success_message, parent=self)

            self._load_pending_pos() # Refresh หน้าจอหลัก

        except Exception as e:
            if conn: conn.rollback()
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถอนุมัติทั้งหมดได้: {e}", parent=self)
            traceback.print_exc()
        finally:
            if conn: self.app_container.release_connection(conn)
        
    def _load_data(self):
        self._update_manager_dashboard()
        self._load_pending_pos()
        
    def _check_and_complete_so(self, so_number):
        if not so_number:
            return
        
        conn = self.app_container.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    SELECT COUNT(id) FROM purchase_orders 
                    WHERE so_number = %s AND status NOT IN ('Draft', 'Rejected')
                """, (so_number,))
                total_pos = cursor.fetchone()[0]

                cursor.execute("""
                    SELECT COUNT(id) FROM purchase_orders 
                    WHERE so_number = %s AND status = 'Approved'
                """, (so_number,))
                approved_pos = cursor.fetchone()[0]

                if total_pos > 0 and total_pos == approved_pos:
                    print(f"All POs for SO {so_number} are approved. Setting SO status to 'PO Complete'.")
                    cursor.execute("""
                        UPDATE commissions SET status = 'PO Complete' 
                        WHERE so_number = %s
                    """, (so_number,))
                    conn.commit()
                else:
                    print(f"SO {so_number} still has pending POs ({approved_pos}/{total_pos} approved).")
                    
        except Exception as e:
            print(f"Error in _check_and_complete_so: {e}")
            if conn: conn.rollback()
        finally:
            if conn: self.app_container.release_connection(conn)