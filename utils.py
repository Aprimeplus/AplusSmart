import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
from datetime import datetime
from customtkinter import CTkFrame, CTkLabel
from customtkinter import CTkToplevel, CTkLabel, CTkFont

def convert_to_float(value_str):
    """
    แปลง String ที่อาจมี comma ให้เป็น float อย่างปลอดภัย
    """
    if value_str is None or value_str == '':
        return 0.0
    try:
        return float(str(value_str).replace(",", ""))
    except (ValueError, TypeError):
        return 0.0
    

def _create_styled_dataframe_table(self, parent, df, label_text="", on_row_click=None, status_colors=None, status_column=None):
        for widget in parent.winfo_children():
            widget.destroy()
        if df is None or df.empty:
            CTkLabel(parent, text=f"ไม่พบข้อมูลสำหรับ '{label_text}'").pack(pady=20)
            return
        
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

        columns = df.columns.tolist()
        
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

        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Custom.Treeview")
        tree.grid(row=0, column=0, sticky="nsew")

        # กำหนดสีพื้นฐานสำหรับสลับแถว
        tree.tag_configure('oddrow', background='#FFFFFF')
        tree.tag_configure('evenrow', background='#F9FAFB') 

        # กำหนดสีตามเงื่อนไขพิเศษทั้งหมด
        if status_colors:
            for tag_name, color in status_colors.items():
                tree.tag_configure(tag_name, background=color)
        
        tree.tag_configure('hit_target_good', background='#D1FAE5')
        tree.tag_configure('hit_target_ok', background='#FEF9C3')
        tree.tag_configure('hit_target_bad', background='#FEE2E2')

        for col_id in columns:
            header_text = col_id
            tree.heading(col_id, text=header_text)
            width = 150 
            if any(s in col_id for s in ['พนักงานขาย']): width = 120
            elif any(s in col_id for s in ['ยอด', 'Margin', 'Incentive', 'Hit Target']): width = 110; tree.column(col_id, anchor='e')
            elif any(s in col_id for s in ['วันที่จ่าย']): width = 160
            elif any(s in col_id for s in ['หมายเหตุ']): width = 250
            elif any(s in col_id for s in ['ID']): width = 50; tree.column(col_id, anchor='center')
            tree.column(col_id, width=width)

        for index, row in df.iterrows():
            # --- START: แก้ไข Logic การกำหนด Tag ใหม่ทั้งหมด ---
            # 1. เริ่มต้นด้วย Tag สีสลับแถวเป็นพื้นฐาน
            final_tags = ['evenrow' if index % 2 == 0 else 'oddrow']

            # 2. ตรวจสอบเงื่อนไขพิเศษ และถ้าตรง ให้ใช้ Tag นั้น (มันจะแสดงผลทับสีพื้นฐาน)
            # สำหรับตาราง User (สีตามแผนก)
            if status_colors and status_column and status_column in df.columns:
                status_val = str(row.get(status_column, ''))
                if status_val in status_colors:
                    final_tags = [status_val] # << ลบของเก่าทิ้ง แล้วใช้สีตามแผนกแทน

            # สำหรับตารางประวัติการจ่ายเงิน (สีตาม Hit Target)
            if 'Hit Target (%)' in row:
                try:
                    hit_target_val = float(row['Hit Target (%)'])
                    if hit_target_val >= 100:
                        final_tags = ['hit_target_good']
                    elif hit_target_val >= 80:
                        final_tags = ['hit_target_ok']
                    else:
                        final_tags = ['hit_target_bad']
                except (ValueError, TypeError):
                    pass # ถ้ามี Error ให้ใช้สีสลับแถวพื้นฐานต่อไป
            
            # --- END: สิ้นสุดการแก้ไข Logic ---
                
            values = []
            for col_name in columns:
                value = row[col_name]
                if pd.notna(value):
                    if isinstance(value, (float, np.floating)): values.append(f"{value:,.2f}")
                    elif isinstance(value, (datetime, pd.Timestamp)): values.append(value.strftime('%Y-%m-%d %H:%M:%S'))
                    else: values.append(str(value))
                else:
                    values.append("")

            row_id_for_iid = row.get('ID', row.get('User Key', str(index)))
            tree.insert("", "end", values=values, tags=tuple(final_tags), iid=str(row_id_for_iid))
        
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.grid(row=0, column=1, sticky='ns'); h_scroll.grid(row=1, column=0, sticky='ew')
        
        if on_row_click: 
            tree.bind("<Double-1>", lambda e: on_row_click(e, tree, df))

def show_loading_popup(master, text="กำลังโหลด..."):
    """
    แสดงหน้าต่าง Pop-up 'กำลังโหลด' แบบง่ายๆ ตรงกลางหน้าจอ
    และคืนค่า object ของหน้าต่างนั้นกลับไป
    """
    popup = CTkToplevel(master)
    popup.title("โปรดรอ")
    popup.geometry("300x100")
    popup.resizable(False, False)

    # จัดให้ Pop-up อยู่ตรงกลางของหน้าต่างหลัก
    master.update_idletasks()
    x = master.winfo_x() + (master.winfo_width() - popup.winfo_width()) // 2
    y = master.winfo_y() + (master.winfo_height() - popup.winfo_height()) // 2
    popup.geometry(f"+{x}+{y}")

    # --- START: จุดที่แก้ไข ---
    # 1. สร้าง Label และเก็บไว้ในตัวแปรชื่อ popup.label
    popup.label = CTkLabel(popup, text=text, font=CTkFont(size=16))
    # 2. นำตัวแปรนั้นมา pack
    popup.label.pack(expand=True)
    # --- END: สิ้นสุดจุดที่แก้ไข ---

    popup.transient(master)
    popup.grab_set()
    master.update_idletasks()

    return popup