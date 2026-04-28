import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from customtkinter import (CTkToplevel, CTkFrame, CTkLabel, CTkButton, 
                           CTkOptionMenu, CTkEntry, CTkFont)
import pandas as pd
from datetime import datetime
import json

class RejectionHistoryWindow(CTkToplevel):
    def __init__(self, master, app_container):
        super().__init__(master)
        self.app_container = app_container
        self.title("ประวัติการตีกลับงาน (PO Rejection History)")
        self.geometry("1100x700")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- ส่วนหัวและตัวกรอง (Header & Filters) ---
        filter_frame = CTkFrame(self, fg_color="transparent")
        filter_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=15)

        CTkLabel(filter_frame, text="ประวัติการตีกลับ", 
                 font=CTkFont(size=20, weight="bold")).pack(side="left", padx=(0, 20))

        # ตัวแปรสำหรับ Filter
        self.thai_months = ["ทุกเดือน", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        self.thai_month_map = {name: i for i, name in enumerate(self.thai_months) if i > 0}
        
        current_year = datetime.now().year
        self.year_options = ["ทุกปี"] + [str(y) for y in range(current_year, current_year - 5, -1)]

        # Dropdown เลือกเดือน
        self.month_var = tk.StringVar(value="ทุกเดือน")
        CTkLabel(filter_frame, text="เดือน:").pack(side="left", padx=(10, 5))
        CTkOptionMenu(filter_frame, variable=self.month_var, values=self.thai_months, width=110).pack(side="left", padx=5)

        # Dropdown เลือกปี
        self.year_var = tk.StringVar(value="ทุกปี")
        CTkLabel(filter_frame, text="ปี:").pack(side="left", padx=(10, 5))
        CTkOptionMenu(filter_frame, variable=self.year_var, values=self.year_options, width=90).pack(side="left", padx=5)

        # ช่องค้นหา Search Bar
        self.search_entry = CTkEntry(filter_frame, placeholder_text="ระบุเลข PO...", width=160)
        self.search_entry.pack(side="left", padx=(15, 5))
        self.search_entry.bind("<Return>", lambda event: self._load_data())

        # ปุ่มค้นหา
        CTkButton(filter_frame, text="ค้นหา", command=self._load_data, width=100,
                  fg_color="#3B82F6", hover_color="#2563EB").pack(side="left", padx=5)

        # [NEW] ปุ่ม Export Excel
        CTkButton(filter_frame, text="Export Excel", command=self._export_to_excel, width=120,
                  fg_color="#107C41", hover_color="#0B532B").pack(side="left", padx=5)

        # --- ส่วนแสดงข้อมูล (Table Area) ---
        table_container = CTkFrame(self)
        table_container.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        table_container.grid_columnconfigure(0, weight=1)
        table_container.grid_rowconfigure(0, weight=1)

        self._create_treeview(table_container)
        
        self.after(100, self._load_data)

    def _create_treeview(self, parent):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Tahoma', 12, 'bold'))
        style.configure("Treeview", rowheight=30, font=('Tahoma', 11))

        columns = ("timestamp", "po_number", "returned_to", "rejected_by", "reason")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("timestamp", text="วัน-เวลา ที่ตีกลับ")
        self.tree.heading("po_number", text="เลขที่ PO")
        self.tree.heading("returned_to", text="ตีกลับให้ (ผู้สร้าง)")
        self.tree.heading("rejected_by", text="ผู้ตีกลับ (Manager)")
        self.tree.heading("reason", text="สาเหตุ / หมายเหตุ")

        self.tree.column("timestamp", width=150, anchor="center")
        self.tree.column("po_number", width=120, anchor="center")
        self.tree.column("returned_to", width=150, anchor="w")
        self.tree.column("rejected_by", width=150, anchor="w")
        self.tree.column("reason", width=400, anchor="w")

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

    def _get_query_and_params(self):
        """Helper Function: สร้าง SQL Query และ Params ตาม Filter ที่เลือก"""
        selected_year = self.year_var.get()
        selected_month_str = self.month_var.get()
        search_text = self.search_entry.get().strip()
        
        where_clauses = ["log.action = 'PO Rejected'", "log.table_name = 'purchase_orders'"]
        params = []

        if selected_year != "ทุกปี":
            where_clauses.append("EXTRACT(YEAR FROM log.timestamp::timestamp) = %s")
            params.append(int(selected_year))
        
        if selected_month_str != "ทุกเดือน":
            month_num = self.thai_month_map[selected_month_str]
            where_clauses.append("EXTRACT(MONTH FROM log.timestamp::timestamp) = %s")
            params.append(month_num)

        if search_text:
            where_clauses.append("po.po_number ILIKE %s")
            params.append(f"%{search_text}%")

        where_str = " AND ".join(where_clauses)

        query = f"""
            SELECT 
                log.timestamp,
                po.po_number,
                u_creator.sale_name AS returned_to_name,
                log.changes
            FROM audit_log log
            LEFT JOIN purchase_orders po ON log.record_id = po.id
            LEFT JOIN sales_users u_creator ON log.user_info = u_creator.sale_key
            WHERE {where_str}
            ORDER BY log.timestamp DESC
        """
        return query, params

    def _parse_changes_column(self, df):
        """Helper Function: แปลงข้อมูล JSON ในคอลัมน์ changes"""
        reasons = []
        rejected_bys = []
        
        for _, row in df.iterrows():
            changes_json = row['changes']
            reason = "-"
            rejected_by = "-"
            
            if isinstance(changes_json, str):
                try:
                    changes_data = json.loads(changes_json)
                    reason = changes_data.get('reason', '-')
                    rejected_by = changes_data.get('rejected_by', '-')
                except: pass
            elif isinstance(changes_json, dict):
                    reason = changes_json.get('reason', '-')
                    rejected_by = changes_json.get('rejected_by', '-')
            
            reasons.append(reason)
            rejected_bys.append(rejected_by)
            
        df['reason'] = reasons
        df['rejected_by'] = rejected_bys
        # ลบคอลัมน์ changes เดิมที่ไม่ใช้แล้วออก
        df.drop(columns=['changes'], inplace=True)
        return df

    def _load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            query, params = self._get_query_and_params()
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=tuple(params))
            
            if df.empty:
                return

            df = self._parse_changes_column(df)

            for _, row in df.iterrows():
                ts = pd.to_datetime(row['timestamp']).strftime('%d/%m/%Y %H:%M')
                self.tree.insert("", "end", values=(
                    ts,
                    row['po_number'] if row['po_number'] else "Deleted/Unknown",
                    row['returned_to_name'] if row['returned_to_name'] else "Unknown",
                    row['rejected_by'],
                    row['reason']
                ))

        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถโหลดข้อมูล: {e}", parent=self)

    def _export_to_excel(self):
        try:
            # 1. ดึงข้อมูลด้วย Logic เดียวกับที่แสดงบนหน้าจอ
            query, params = self._get_query_and_params()
            df = pd.read_sql_query(query, self.app_container.pg_engine, params=tuple(params))

            if df.empty:
                messagebox.showinfo("ไม่พบข้อมูล", "ไม่มีข้อมูลสำหรับ Export ตามเงื่อนไขที่เลือก", parent=self)
                return

            # 2. จัดการข้อมูลให้สวยงาม
            df = self._parse_changes_column(df)
            
            # จัดรูปแบบวันที่
            df['timestamp'] = pd.to_datetime(df['timestamp']).dt.strftime('%d/%m/%Y %H:%M:%S')
            
            # เปลี่ยนชื่อคอลัมน์ภาษาไทย
            df.rename(columns={
                'timestamp': 'วัน-เวลา ที่ตีกลับ',
                'po_number': 'เลขที่ PO',
                'returned_to_name': 'ตีกลับให้ (ผู้สร้าง)',
                'rejected_by': 'ผู้ตีกลับ (Manager)',
                'reason': 'สาเหตุ / หมายเหตุ'
            }, inplace=True)

            # 3. เปิด Dialog ให้เลือกที่เซฟไฟล์
            current_time_str = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"Rejection_History_{current_time_str}.xlsx"
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=filename,
                title="บันทึกไฟล์ Excel"
            )

            if not file_path:
                return

            # 4. บันทึกลง Excel
            df.to_excel(file_path, index=False)
            messagebox.showinfo("สำเร็จ", f"ส่งออกข้อมูลเรียบร้อยแล้วที่:\n{file_path}", parent=self)

        except Exception as e:
            messagebox.showerror("Export Error", f"เกิดข้อผิดพลาดในการส่งออกไฟล์: {e}", parent=self)