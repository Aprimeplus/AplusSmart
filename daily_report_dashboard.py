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
        # 1. ดึงข้อมูลพื้นฐาน
        month_idx = self.thai_months.index(self.dash_month_var.get()) + 1
        year = int(self.dash_year_var.get())
        _, num_days = calendar.monthrange(year, month_idx)
        
        # 2. ดึงยอดขายจาก Database
        try:
            query = """
                SELECT EXTRACT(DAY FROM bill_date) as day, SUM(sales_service_amount) as amount
                FROM commissions
                WHERE EXTRACT(MONTH FROM bill_date) = %s 
                  AND EXTRACT(YEAR FROM bill_date) = %s
                  AND is_active = 1
                GROUP BY day
                ORDER BY day
            """
            sales_df = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year))
        except Exception as e:
            print(f"Sales Data Error: {e}")
            return

        # 3. ดึงเป้าหมายปี
        target_annual = 0
        try:
            t_query = "SELECT target_amount FROM sales_yearly_targets WHERE year = %s"
            targets = pd.read_sql_query(t_query, self.pg_engine, params=(year,))
            target_annual = targets.iloc[0]['target_amount'] if not targets.empty else 120000000
        except Exception: pass

        # 4. สร้าง DataFrame (ต้องสร้างก่อนใช้งาน apply)
        all_days = pd.DataFrame({'day': range(1, num_days + 1)})
        
        # 5. คำนวณ Weighted Target
        def get_day_weight(d):
            dt = datetime(year, month_idx, d)
            weekday = dt.weekday()
            if weekday < 5: return 1.5   # จ-ศ
            if weekday == 5: return 1.0  # ส
            return 0                     # อา

        all_days['weight'] = all_days['day'].apply(get_day_weight)
        target_monthly = target_annual / 12
        total_month_weights = all_days['weight'].sum()
        
        if total_month_weights > 0:
            unit_value = target_monthly / total_month_weights
            all_days['daily_target'] = all_days['weight'] * unit_value
        else:
            all_days['daily_target'] = 0

        # 6. รวมข้อมูลและคำนวณยอดสะสม
        merged = pd.merge(all_days, sales_df, on='day', how='left').fillna(0).infer_objects(copy=False)
        merged['cumulative_sales'] = merged['amount'].cumsum()
        merged['cumulative_target'] = merged['daily_target'].cumsum()
        
        # 7. สรุปค่าสำหรับ Gauge
        total_actual = merged['amount'].sum()
        total_target_month = merged['daily_target'].sum()
        
        # 8. ตัดยอดสะสมวันที่ยังไม่มีการขายจริง (แก้ปัญหายอดพุ่งทะลุ)
        last_sale_day = merged[merged['amount'] > 0]['day'].max()
        if pd.isna(last_sale_day): last_sale_day = 0
        merged.loc[merged['day'] > last_sale_day, 'cumulative_sales'] = None

        # 9. สั่งวาดกราฟทั้ง 2 ส่วน
        # วาดกราฟแท่งหลักก่อน
        self._draw_matplotlib_chart(merged, year, self.dash_month_var.get(), 0)
        
        # ยกแผง info_panel ขึ้นมาด้านหน้าสุดเพื่อให้ Gauge แสดงผล
        self.info_panel.lift()
        
        # วาด Gauge ลงใน container
        self._draw_gauge_chart(self.gauge_container, total_actual, total_target_month)

    def _draw_matplotlib_chart(self, df, year, month_name, daily_target):
        # ล้างกราฟเก่าออกก่อนวาดใหม่
        if self.canvas: 
            self.canvas.get_tk_widget().destroy()
        
        # ตั้งค่า Font ให้รองรับภาษาไทย
        plt.rcParams['font.family'] = 'Tahoma' 
        # ปรับขนาด figsize ให้เหมาะสมกับหน้าจอ Full Screen
        fig, ax = plt.subplots(figsize=(14, 7), dpi=100)
        
        # --- 1. เตรียมข้อมูลสำหรับ Stacked Bar ---
        days = df['day']
        actual = df['cumulative_sales']
        target = df['cumulative_target']
        
        # ใช้ .fillna(0) สำหรับการคำนวณพื้นที่แท่ง เพื่อให้วันที่ไม่มีข้อมูล (None) แสดงเป็นแท่งเป้าหมายว่างๆ
        actual_for_calc = actual.fillna(0)
        
        # ส่วนต่างของเป้าหมายที่ยังเหลืออยู่ (เพื่อให้แท่งรวมสูงเท่ากับเป้าหมายเสมอ)
        remaining_target = (target - actual_for_calc).clip(lower=0)
        
        # ส่วนที่ยอดขายเกินเป้าหมาย (Over Target)
        over_target = (actual_for_calc - target).clip(lower=0)
        
        # ยอดขายจริงในส่วนที่ไม่เกินเส้นเป้าหมาย
        actual_in_target = actual_for_calc.where(actual_for_calc <= target, target)

        width = 0.6 # ความกว้างของแท่งกราฟ

        # --- 2. วาดกราฟแท่งซ้อน (Stacked) ---
        # ชั้นที่ 1: ยอดขายสะสมจริง (สีส้ม) - วาดเฉพาะวันที่มีข้อมูลจริง (ไม่เป็น None)
        ax.bar(days, actual_in_target.where(actual.notna()), width, 
               label='ยอดขายสะสม (Actual)', color='#F59E0B', alpha=0.9)
        
        # ชั้นที่ 2: เป้าหมายส่วนที่เหลือ (สีฟ้าใส)
        # หากวันนั้น Actual เป็น None แท่งนี้จะสูงเท่ากับ Target พอดี (เป็นแท่งว่างๆ รอการเติม)
        ax.bar(days, remaining_target, width, bottom=actual_in_target.where(actual.notna(), 0), 
               label='เป้าหมายที่รอการเติม (Target)', color='#3B82F6', alpha=0.2, 
               edgecolor='#3B82F6', linestyle='--', linewidth=0.5)
        
        # ชั้นที่ 3: ส่วนที่ทำยอดเกินเป้าหมาย (สีส้มเข้ม)
        ax.bar(days, over_target.where(actual.notna()), width, bottom=target, 
               label='ยอดขายส่วนเกินเป้าหมาย', color='#D97706', alpha=1.0)

        # --- 3. จัดการแกน X (วันที่และชื่อวันภาษาไทย) ---
        thai_day_names = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]
        month_idx = self.thai_months.index(self.dash_month_var.get()) + 1
        
        # สร้าง Label เช่น "1\nพฤ."
        x_labels = []
        for d in days:
            weekday_idx = datetime(year, month_idx, int(d)).weekday()
            x_labels.append(f"{int(d)}\n{thai_day_names[weekday_idx]}")
        
        ax.set_xticks(days)
        ax.set_xticklabels(x_labels, fontsize=7)
        
        # ไฮไลท์ชื่อวันเสาร์-อาทิตย์ให้เป็นสีแดง
        for i, label in enumerate(ax.get_xticklabels()):
            weekday_idx = datetime(year, month_idx, int(df['day'][i])).weekday()
            if weekday_idx >= 5: # 5=Sat, 6=Sun
                label.set_color('#DC2626')
                label.set_weight('bold')

        # --- 4. แสดงตัวเลขยอดขายบนหัวแท่ง (เอียง 45 องศา) ---
        for i, val in enumerate(actual):
            # แสดงตัวเลขเฉพาะวันที่มีการขายจริง และยอดมากกว่า 0
            if pd.notna(val) and val > 0 and df.get('amount', pd.Series([0]*len(df)))[i] > 0:
                ax.text(days[i], val, f'{val/1000000:.1f}M', ha='center', va='bottom', 
                        fontsize=7, color='#B45309', weight='bold', rotation=45)

        # --- 5. ตกแต่งแกน Y และสัญลักษณ์ (Legend) ---
        # ฟอร์แมตแกน Y ให้แสดงเป็นหน่วยล้าน (M)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, pos: f'{x/1000000:.1f}M'))
        
        ax.set_title(f"สัดส่วนยอดขายสะสมเทียบเป้าหมาย (Weighted Stacked) - {month_name} {year}", 
                     fontsize=15, weight='bold', pad=25)
        
        # สร้าง Legend แบบกำหนดเอง (จะไปปรากฏอยู่เหนือ Gauge ในแผงข้อมูลฝั่งซ้าย)
        from matplotlib.lines import Line2D
        custom_legend = [
            Line2D([0], [0], color='#F59E0B', lw=8, label='ยอดขายสะสม (Actual)'),
            Line2D([0], [0], color='#3B82F6', alpha=0.2, lw=8, label='เป้าหมายที่รอการเติม (Target)'),
            Line2D([0], [0], color='#D97706', lw=8, label='ทำยอดได้เกินเป้า (Over Target)')
        ]
        # วาง Legend ไว้มุมซ้ายบนให้สัมพันธ์กับตำแหน่งของ info_panel
        ax.legend(handles=custom_legend, loc='upper left', fontsize=9, frameon=False)
        
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        
        # ปรับขอบกราฟให้พอดี
        fig.tight_layout()
        
        # นำกราฟไปแสดงผลในพื้นที่หลัก (main_chart_area)
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