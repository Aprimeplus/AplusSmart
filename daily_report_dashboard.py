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
        
        # --- 1. Control Bar (ด้านบนสุด) ---
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

        # --- 2. พื้นที่ Dashboard (แบ่งซ้าย-ขวา) ---
        self.display_container = CTkFrame(self, fg_color="white")
        self.display_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # ฝั่งซ้าย: กราฟแท่งรายวัน (ขยายเต็มที่)
        self.left_graph_area = CTkFrame(self.display_container, fg_color="white")
        self.left_graph_area.pack(side="left", fill="both", expand=True)
        
        self.main_chart_canvas_area = CTkFrame(self.left_graph_area, fg_color="white")
        self.main_chart_canvas_area.place(relx=0, rely=0, relwidth=1, relheight=1)

        # ฝั่งขวา: แถบสรุปเป้าหมาย (Side Panel)
        self.right_summary_panel = CTkFrame(self.display_container, fg_color="#F8FAFC", width=260)
        self.right_summary_panel.pack(side="right", fill="y", padx=(5, 0))
        self.right_summary_panel.pack_propagate(False) 
        
        # ส่วนแสดงผล Monthly
        CTkLabel(self.right_summary_panel, text="ยอดขายสะสมรายเดือน", font=("Arial", 14, "bold")).pack(pady=(30, 0))
        # แก้ไข Error: กำหนด height ในนี้แทน
        self.monthly_circle_area = CTkFrame(self.right_summary_panel, fg_color="transparent", height=220)
        self.monthly_circle_area.pack(fill="x", padx=10)

        # เส้นคั่น
        CTkFrame(self.right_summary_panel, height=2, fg_color="#E2E8F0").pack(fill="x", padx=30, pady=20)

        # ส่วนแสดงผล Yearly
        CTkLabel(self.right_summary_panel, text="ยอดขายสะสมรายปี", font=("Arial", 14, "bold")).pack(pady=(0, 0))
        self.yearly_circle_area = CTkFrame(self.right_summary_panel, fg_color="transparent", height=220)
        self.yearly_circle_area.pack(fill="x", padx=10)

        self.canvas = None
        self.monthly_canvas = None
        self.yearly_canvas = None
        self.after(500, self._update_chart)

    def _open_settings(self):
        year = int(self.dash_year_var.get())
        TargetSettingsDialog(self, self.app_container, year, on_save_callback=self._update_chart)

    def _update_chart(self, event=None):
        # 1. เตรียมข้อมูลเดือนและปีที่เลือก
        month_name = self.dash_month_var.get()
        month_idx = self.thai_months.index(month_name) + 1
        year = int(self.dash_year_var.get())
        _, num_days = calendar.monthrange(year, month_idx)
        
        # 2. ดึงยอดขายรายวันและยอดรวมสะสมทั้งปีจากฐานข้อมูล
        try:
            # ยอดรายวันเฉพาะเดือนที่เลือก
            query = """
                SELECT EXTRACT(DAY FROM bill_date) as day, SUM(sales_service_amount) as amount
                FROM commissions
                WHERE EXTRACT(MONTH FROM bill_date) = %s 
                  AND EXTRACT(YEAR FROM bill_date) = %s 
                  AND is_active = 1
                GROUP BY day ORDER BY day
            """
            sales_df = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year))
            
            # ยอดรวมทั้งปีสะสมจนถึงปัจจุบัน
            y_query = """
                SELECT SUM(sales_service_amount) 
                FROM commissions 
                WHERE EXTRACT(YEAR FROM bill_date) = %s 
                  AND is_active = 1
            """
            y_res = pd.read_sql_query(y_query, self.pg_engine, params=(year,))
            total_actual_year = y_res.iloc[0, 0] or 0
        except Exception as e:
            print(f"Database Error: {e}")
            return

        # 3. สร้างรายการวันที่โดย "ไม่รวมวันอาทิตย์"
        days_list = [d for d in range(1, num_days + 1) if datetime(year, month_idx, d).weekday() != 6]
        all_days = pd.DataFrame({'day': days_list})
        
        # 4. คำนวณเป้าหมายสะสมแบบถ่วงน้ำหนัก (Weighted Target)
        def get_day_weight(d):
            wd = datetime(year, month_idx, d).weekday()
            return 1.5 if wd < 5 else 1.0 # จันทร์-ศุกร์ น้ำหนัก 1.5, เสาร์ น้ำหนัก 1.0

        all_days['weight'] = all_days['day'].apply(get_day_weight)
        
        # ดึงเป้าหมายปีจากฐานข้อมูล (ถ้าไม่มีให้ใช้ค่าเริ่มต้น)
        target_annual = 120000000
        try:
            t_query = "SELECT target_amount FROM sales_yearly_targets WHERE year = %s"
            t_df = pd.read_sql_query(t_query, self.pg_engine, params=(year,))
            if not t_df.empty: target_annual = t_df.iloc[0, 0]
        except: pass

        # กระจายเป้าหมายรายเดือนลงในแต่ละวันตามน้ำหนัก
        unit_value = (target_annual / 12) / all_days['weight'].sum()
        all_days['daily_target'] = all_days['weight'] * unit_value

        # 5. รวมข้อมูลยอดขายจริงเข้ากับรายการวันและคำนวณยอดสะสม
        merged = pd.merge(all_days, sales_df, on='day', how='left').fillna(0).infer_objects(copy=False)
        
        # คำนวณยอดสะสมปัจจุบัน และยอดสะสมก่อนหน้าเพื่อวาดแท่ง 2 สี
        merged['cumulative_sales'] = merged['amount'].cumsum()
        merged['previous_sales'] = merged['cumulative_sales'] - merged['amount']
        
        # คำนวณเส้นเป้าหมายสะสม (เส้นสีแดง)
        merged['cumulative_target'] = merged['daily_target'].cumsum()
        
        # สรุปค่าสำหรับวงกลม Progress ทางขวา
        total_actual_month = merged['amount'].sum()
        total_target_month = merged['daily_target'].sum()

        # 6. Logic จัดการวันในอนาคต (ไม่ให้ยอดสะสมแสดงในวันที่ยังไม่มีการขายจริง)
        last_sale_day = merged[merged['amount'] > 0]['day'].max()
        if pd.isna(last_sale_day): last_sale_day = 0
        
        # เก็บค่า copy ไว้ใช้แสดงผลตัวเลขรวม ก่อนจะสั่งให้วันในอนาคตเป็น None เพื่อหยุดวาดกราฟ
        merged.loc[merged['day'] > last_sale_day, 'cumulative_sales'] = None

        # 7. สั่งวาดกราฟลงพื้นที่แสดงผล
        # วาดกราฟแท่ง 2 สี + เส้นแดงฝั่งซ้าย
        self._draw_matplotlib_chart(merged, year, month_name)
        
        # วาดวงกลมสรุปรายเดือนและรายปีฝั่งขวา
        self._draw_circle_progress(self.monthly_circle_area, total_actual_month, total_target_month, "monthly_canvas")
        self._draw_circle_progress(self.yearly_circle_area, total_actual_year, target_annual, "yearly_canvas")

    def _draw_matplotlib_chart(self, df, year, month_name):
        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        plt.rcParams['font.family'] = 'Tahoma'
        fig, ax = plt.subplots(figsize=(15, 8), dpi=100)

        days = df['day']
        prev_sales = df['previous_sales']
        daily_amount = df['amount']
        total_sales = df['cumulative_sales']
        target = df['cumulative_target']

        bar_width = 0.45

        # -----------------------------
        # 1. Bars
        # -----------------------------
        ax.bar(days, prev_sales, width=bar_width,
            color='#F59E0B', alpha=0.7, label='ยอดสะสมเดิม')

        ax.bar(days, daily_amount, bottom=prev_sales, width=bar_width,
            color='#0EA5E9', alpha=0.9, label='ยอดขายเพิ่มวันนี้')

        # -----------------------------
        # 2. Target line
        # -----------------------------
        ax.plot(days, target,
                color='#DC2626', linewidth=2,
                marker='o', markersize=4,
                label='เป้าหมายสะสม')

        # -----------------------------
        # 3. LABEL CONTROL (หัวใจของการแก้ซ้อน)
        # -----------------------------
        MIN_SHOW = 0.10e6                 # ไม่โชว์ < 0.10M
        OFFSET_LEVELS = [0.12e6, 0.32e6]  # สลับสูง-ต่ำ

        for i, day in enumerate(days):
            if (
                pd.notna(total_sales[i]) and
                daily_amount[i] >= MIN_SHOW
            ):
                offset = OFFSET_LEVELS[i % 2]  # <<< stagger

                ax.text(
                    day,
                    total_sales[i] + offset,
                    f'+{daily_amount[i] / 1e6:.2f}M',
                    ha='center',
                    va='bottom',
                    fontsize=9,
                    weight='bold',
                    color='#0369A1',
                    bbox=dict(
                        facecolor='white',
                        edgecolor='none',
                        alpha=0.85,
                        pad=2
                    )
                )

        # -----------------------------
        # 4. Cumulative at LAST DAY ONLY
        # -----------------------------
        last_idx = df[pd.notna(total_sales)].index.max()
        if pd.notna(last_idx):
            last_day = days[last_idx]
            last_total = total_sales[last_idx]

            ax.text(
                last_day,
                last_total + 0.6e6,
                f'ยอดสะสม {last_total/1e6:.1f}M',
                ha='center',
                va='bottom',
                fontsize=10,
                weight='bold',
                color='#1E293B'
            )

        # -----------------------------
        # 5. Axis
        # -----------------------------
        thai_days = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]
        month_idx = self.thai_months.index(month_name) + 1

        ax.set_xticks(days)
        ax.set_xticklabels(
            [f"{int(d)}\n{thai_days[datetime(year, month_idx, int(d)).weekday()]}" for d in days],
            fontsize=8
        )

        ax.yaxis.set_major_formatter(
            plt.FuncFormatter(lambda x, p: f'{x/1e6:.1f}M')
        )

        # <<< เผื่อพื้นที่ด้านบนเพิ่ม เพื่อรองรับ label
        ymax = max(target.max(skipna=True), total_sales.max(skipna=True))
        ax.set_ylim(0, ymax * 1.22)

        # -----------------------------
        # 6. Styling
        # -----------------------------
        ax.set_title(
            f"วิเคราะห์ยอดขายและการเติบโตรายวัน - {month_name} {year}",
            fontsize=16, weight='bold', pad=20
        )

        ax.legend(loc='upper left', fontsize=9, frameon=False, ncol=3)
        ax.grid(axis='y', linestyle='--', alpha=0.3)

        plt.subplots_adjust(top=0.88, bottom=0.12, left=0.08, right=0.95)

        self.canvas = FigureCanvasTkAgg(fig, master=self.main_chart_canvas_area)
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

    def _draw_circle_progress(self, parent_frame, actual, target, title, canvas_attr):
        # ล้าง Canvas เก่า
        old_canvas = getattr(self, canvas_attr)
        if old_canvas: old_canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots(figsize=(3, 3), dpi=80)
        fig.patch.set_alpha(0)
        
        pct = min(100, (actual / target * 100)) if target > 0 else 0
        
        # วาดวงกลม 2 ส่วน (สีฟ้าคือส่วนที่ทำได้, สีเทาคือส่วนที่เหลือ)
        ax.pie([pct, 100-pct], colors=['#0EA5E9', '#E2E8F0'], startangle=90, counterclock=False, 
               wedgeprops={'width': 0.3, 'edgecolor': 'w'})
        
        # ใส่ตัวเลขตรงกลาง
        ax.text(0, 0, f'{pct:.1f}%\n{actual/1000000:.1f}M', ha='center', va='center', 
                fontsize=12, weight='bold', color='#0369A1')
        
        ax.axis('equal')
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        setattr(self, canvas_attr, canvas)

    def _draw_circle_progress(self, parent_frame, actual, target, canvas_attr):
        old_canvas = getattr(self, canvas_attr)
        if old_canvas: old_canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots(figsize=(2.8, 2.8), dpi=85)
        fig.patch.set_alpha(0)
        
        pct = (actual / target * 100) if target > 0 else 0
        
        # วาด Donut Chart
        ax.pie([min(pct, 100), max(0, 100-pct)], colors=['#F59E0B', '#F1F5F9'], 
               startangle=90, counterclock=False, wedgeprops={'width': 0.3})
        
        ax.text(0, 0, f'{pct:.1f}%\n{actual/1000000:.1f}M', ha='center', va='center', 
                fontsize=11, weight='bold', color='#1E293B')
        
        ax.axis('equal')
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        setattr(self, canvas_attr, canvas)