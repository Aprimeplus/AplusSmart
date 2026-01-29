import tkinter as tk
from tkinter import messagebox
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkOptionMenu, 
                           CTkFont, CTkToplevel, CTkEntry)
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
import calendar
import psycopg2

# =============================================================================
#  ส่วนที่ 1: Dialog ตั้งค่าเป้าหมาย (Popup)
#  (ใส่ไว้ในไฟล์นี้เลย จะได้ไม่ต้องแยกหลายไฟล์เกินไป)
# =============================================================================
class TargetSettingsDialog(CTkToplevel):
    def __init__(self, master, app_container, year, on_save_callback):
        super().__init__(master)
        self.app_container = app_container
        self.target_year = year
        self.on_save_callback = on_save_callback
        
        self.title(f"ตั้งค่าเป้าหมายปี {year}")
        self.geometry("350x300")
        self.grab_set() 
        
        CTkLabel(self, text=f"ตั้งเป้าหมายปี {year}", font=("Arial", 18, "bold")).pack(pady=15)
        
        # ช่องกรอกเป้าหมาย
        CTkLabel(self, text="เป้าหมายยอดขายทั้งปี (บาท):").pack(pady=(5,0))
        self.target_entry = CTkEntry(self, width=200)
        self.target_entry.pack(pady=5)
        
        # ปุ่มบันทึก
        CTkButton(self, text="บันทึก", command=self._save_target, fg_color="#16A34A").pack(pady=20)
        self._load_current_settings()

    def _load_current_settings(self):
        try:
            conn = self.app_container.get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT target_amount FROM sales_yearly_targets WHERE year = %s", (self.target_year,))
            row = cursor.fetchone()
            if row: self.target_entry.insert(0, f"{row[0]:.2f}")
            else: self.target_entry.insert(0, "120000000")
        except Exception as e: print(f"Error: {e}")
        finally: self.app_container.release_connection(conn)

    def _save_target(self):
        try:
            target = float(self.target_entry.get().replace(",", ""))
            conn = self.app_container.get_connection()
            with conn.cursor() as cursor:
                sql = """
                    INSERT INTO sales_yearly_targets (year, target_amount, updated_at)
                    VALUES (%s, %s, NOW())
                    ON CONFLICT (year) DO UPDATE SET target_amount = EXCLUDED.target_amount, updated_at = NOW();
                """
                cursor.execute(sql, (self.target_year, target))
            conn.commit()
            messagebox.showinfo("สำเร็จ", "บันทึกเป้าหมายแล้ว", parent=self)
            if self.on_save_callback: self.on_save_callback()
            self.destroy()
        except Exception as e: messagebox.showerror("Error", f"{e}", parent=self)

# =============================================================================
#  ส่วนที่ 2: ตัวแสดงผล Dashboard (Class หลัก)
# =============================================================================
class DailyDashboard(CTkFrame):
    def __init__(self, master, app_container):
        super().__init__(master, fg_color="transparent")
        self.app_container = app_container
        self.pg_engine = app_container.pg_engine
        
        # --- 1. Control Bar ---
        ctrl_frame = CTkFrame(self)
        ctrl_frame.pack(fill="x", padx=10, pady=10)
        
        self.thai_months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        
        self.dash_month_var = tk.StringVar(value=self.thai_months[datetime.now().month - 1])
        CTkOptionMenu(ctrl_frame, variable=self.dash_month_var, values=self.thai_months, width=120, command=self._update_chart).pack(side="left", padx=5)
        
        current_year = datetime.now().year
        self.dash_year_var = tk.StringVar(value=str(current_year))
        years = [str(y) for y in range(current_year - 2, current_year + 3)]
        CTkOptionMenu(ctrl_frame, variable=self.dash_year_var, values=years, width=80, command=self._update_chart).pack(side="left", padx=5)
        
        CTkButton(ctrl_frame, text="⟳ รีเฟรช", width=80, command=self._update_chart).pack(side="left", padx=10)
        CTkButton(ctrl_frame, text="⚙️ ตั้งค่าเป้าหมาย", fg_color="#F59E0B", command=self._open_settings).pack(side="right", padx=10)

        # --- 2. พื้นที่ Dashboard หลัก ---
        self.chart_frame = CTkFrame(self, fg_color="white")
        self.chart_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # กราฟแท่งหลัก ขยายเต็มพื้นที่
        self.main_chart_area = CTkFrame(self.chart_frame, fg_color="white")
        self.main_chart_area.place(relx=0, rely=0, relwidth=1, relheight=1)
        
        # --- 3. แผงข้อมูลมุมซ้ายบน (Info Panel) ---
        self.info_panel = CTkFrame(
            self.main_chart_area, 
            fg_color="white", 
            corner_radius=10, 
            border_width=1, 
            border_color="#E5E7EB",
            width=220,    
            height=280
        )
        self.info_panel.place(x=85, y=145) # วางตำแหน่งทับบนกราฟ
        
        # พื้นที่สำหรับ Gauge (เว้นระยะด้านบนให้ Legend ของกราฟ)
        self.gauge_container = CTkFrame(self.info_panel, fg_color="transparent")
        self.gauge_container.pack(fill="both", expand=True, padx=5, pady=(45, 5))

        self.canvas = None
        self.gauge_canvas = None
        self.after(500, self._update_chart)

    def _open_settings(self):
        year = int(self.dash_year_var.get())
        TargetSettingsDialog(self, self.app_container, year, on_save_callback=self._update_chart)

    def _update_chart(self, event=None):
        month_idx = self.thai_months.index(self.dash_month_var.get()) + 1
        year = int(self.dash_year_var.get())
        _, num_days = calendar.monthrange(year, month_idx)
        
        # 1. ดึงยอดขายรายวัน
        try:
            query = """
                SELECT EXTRACT(DAY FROM bill_date) as day, SUM(sales_service_amount) as amount
                FROM commissions
                WHERE EXTRACT(MONTH FROM bill_date) = %s 
                  AND EXTRACT(YEAR FROM bill_date) = %s
                  AND is_active = 1
                GROUP BY day ORDER BY day
            """
            sales_df = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year))
        except Exception as e:
            print(f"Error: {e}"); return

        # 2. คำนวณวันและเป้าหมาย (ตัดวันอาทิตย์ออก)
        days_list = []
        for d in range(1, num_days + 1):
            if datetime(year, month_idx, d).weekday() != 6: # 6 คือวันอาทิตย์
                days_list.append(d)
        
        all_days = pd.DataFrame({'day': days_list})
        
        # คำนวณ Weighted Target (จ-ศ x1.5, ส x1.0)
        def get_day_weight(d):
            wd = datetime(year, month_idx, d).weekday()
            return 1.5 if wd < 5 else 1.0

        all_days['weight'] = all_days['day'].apply(get_day_weight)
        
        # ดึงเป้าหมายปี
        target_annual = 120000000
        try:
            t_query = "SELECT target_amount FROM sales_yearly_targets WHERE year = %s"
            targets = pd.read_sql_query(t_query, self.pg_engine, params=(year,))
            if not targets.empty: target_annual = targets.iloc[0]['target_amount']
        except: pass

        unit_value = (target_annual / 12) / all_days['weight'].sum()
        all_days['daily_target'] = all_days['weight'] * unit_value
        
        # 3. รวมข้อมูลและคำนวณสะสม
        merged = pd.merge(all_days, sales_df, on='day', how='left').fillna(0)
        merged['cumulative_sales'] = merged['amount'].cumsum()
        merged['cumulative_target'] = merged['daily_target'].cumsum()
        
        total_actual = merged['amount'].sum()
        total_target_month = merged['daily_target'].sum()

        # หยุดวาดวันในอนาคตที่ยังไม่มี SO
        last_sale_day = merged[merged['amount'] > 0]['day'].max()
        if pd.isna(last_sale_day): last_sale_day = 0
        merged.loc[merged['day'] > last_sale_day, 'cumulative_sales'] = None

        # 4. วาดกราฟ
        self._draw_matplotlib_chart(merged, year, self.dash_month_var.get())
        self.info_panel.lift()
        self._draw_gauge_chart(self.gauge_container, total_actual, total_target_month)

    def _draw_matplotlib_chart(self, df, year, month_name):
        if self.canvas: self.canvas.get_tk_widget().destroy()
        
        plt.rcParams['font.family'] = 'Tahoma' 
        fig, ax = plt.subplots(figsize=(14, 7), dpi=100)
        
        # --- วาดกราฟแท่งยอดขายจริง (Actual) ---
        ax.bar(df['day'], df['cumulative_sales'], color='#F59E0B', alpha=0.8, 
               label='ยอดขายสะสม (Actual)', width=0.6)
        
        # --- วาดเส้นเป้าหมาย (Target Line) ---
        ax.plot(df['day'], df['cumulative_target'], color='#DC2626', marker='o', 
                markersize=4, linewidth=2, label='เป้าหมายสะสม (Weighted Target)')

        # จัดการแกน X (ไม่มีวันอาทิตย์)
        thai_days = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]
        month_idx = self.thai_months.index(self.dash_month_var.get()) + 1
        x_labels = [f"{int(d)}\n{thai_days[datetime(year, month_idx, int(d)).weekday()]}" for d in df['day']]
        
        ax.set_xticks(df['day'])
        ax.set_xticklabels(x_labels, fontsize=7)
        
        # แสดงตัวเลขล้านบนหัวแท่ง
        for i, val in enumerate(df['cumulative_sales']):
            if pd.notna(val) and val > 0:
                ax.text(df['day'][i], val, f'{val/1000000:.1f}M', ha='center', va='bottom', 
                        fontsize=7, color='#B45309', weight='bold')

        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000000:.1f}M'))
        ax.set_title(f"Dashboard ยอดขายสะสมรายวัน - {month_name} {year}", fontsize=15, weight='bold', pad=20)
        ax.legend(loc='upper left', fontsize=9, frameon=False)
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        
        fig.tight_layout()
        self.canvas = FigureCanvasTkAgg(fig, master=self.main_chart_area)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _draw_gauge_chart(self, parent_frame, actual, target):
        # ล้าง Gauge เก่าออกเสมอ
        if self.gauge_canvas: 
            self.gauge_canvas.get_tk_widget().destroy()

        # สร้างมาตรวัดความสำเร็จ (Gauge) แบบครึ่งวงกลม
        fig, ax = plt.subplots(figsize=(2.5, 2), dpi=90)
        fig.patch.set_alpha(0) # พื้นหลังโปร่งใสเพื่อให้เข้ากับแผงสีขาว

        pct = (actual / target * 100) if target > 0 else 0
        gap = max(0, target - actual)
        
        # ข้อมูล [ยอดขายที่ทำได้, ยอดที่ยังขาด, ฐานล่างที่ซ่อนไว้]
        values = [min(actual, target), gap, target] 
        colors = ['#F59E0B', '#3B82F6', '#F1F5F9'] 

        ax.pie(values, colors=colors, startangle=180, counterclock=True, 
               wedgeprops={'width': 0.35, 'edgecolor': 'w', 'linewidth': 1})

        # แสดงตัวเลข % ตรงกลางวงแหวน
        ax.text(0, 0.1, f'{pct:.1f}%', ha='center', va='center', fontsize=16, weight='bold', color='#1E293B')
        
        # แสดงยอดเงิน Act และ Gap ด้านล่างมาตรวัด
        ax.text(0, -0.25, f'Act: {actual/1000000:.2f}M', ha='center', fontsize=9, color='#F59E0B', weight='bold')
        ax.text(0, -0.45, f'Gap: {gap/1000000:.2f}M', ha='center', fontsize=8, color='#64748B')

        ax.axis('equal') 
        plt.tight_layout()
        
        # นำ Gauge Canvas ไปใส่ใน parent_frame (gauge_container)
        self.gauge_canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        self.gauge_canvas.draw()
        self.gauge_canvas.get_tk_widget().pack(fill="both", expand=True)