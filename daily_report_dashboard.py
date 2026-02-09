import tkinter as tk
from tkinter import messagebox
from customtkinter import (CTkFrame, CTkLabel, CTkButton, CTkOptionMenu, 
                           CTkFont, CTkToplevel, CTkEntry)
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MultipleLocator
from datetime import datetime
import calendar
import psycopg2

# =============================================================================
#  ส่วนที่ 1: Dialog ตั้งค่าเป้าหมาย (Popup)
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
#  ส่วนที่ 2: ตัวแสดงผล Dashboard (Class หลัก) - ปรับปรุงแล้ว
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
        
        # 2. ดึงยอดขายจากฐานข้อมูล
        try:
            # 2.1 ยอดรายวัน
            query = """
                SELECT EXTRACT(DAY FROM bill_date) as day, SUM(sales_service_amount) as amount
                FROM commissions
                WHERE EXTRACT(MONTH FROM bill_date) = %s 
                  AND EXTRACT(YEAR FROM bill_date) = %s 
                  AND is_active = 1
                GROUP BY day ORDER BY day
            """
            sales_df = pd.read_sql_query(query, self.pg_engine, params=(month_idx, year))
            
            # 2.2 ยอดสะสมเดือนก่อนหน้า (YTD Offset)
            prev_months_query = """
                SELECT SUM(sales_service_amount) 
                FROM commissions 
                WHERE EXTRACT(YEAR FROM bill_date) = %s 
                  AND EXTRACT(MONTH FROM bill_date) < %s
                  AND is_active = 1
            """
            prev_res = pd.read_sql_query(prev_months_query, self.pg_engine, params=(year, month_idx))
            start_sales_offset = prev_res.iloc[0, 0] or 0.0 

            # 2.3 ยอดรวมทั้งปี
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

        # 3. สร้างรายการวันที่
        days_list = [d for d in range(1, num_days + 1) if datetime(year, month_idx, d).weekday() != 6]
        all_days = pd.DataFrame({'day': days_list})
        
        # 4. คำนวณเป้าหมาย
        def get_day_weight(d):
            wd = datetime(year, month_idx, d).weekday()
            return 1.5 if wd < 5 else 1.0 

        all_days['weight'] = all_days['day'].apply(get_day_weight)
        
        target_annual = 120000000
        try:
            t_query = "SELECT target_amount FROM sales_yearly_targets WHERE year = %s"
            t_df = pd.read_sql_query(t_query, self.pg_engine, params=(year,))
            if not t_df.empty: target_annual = t_df.iloc[0, 0]
        except: pass

        target_per_month = target_annual / 12
        start_target_offset = target_per_month * (month_idx - 1)
        unit_value = target_per_month / all_days['weight'].sum()
        all_days['daily_target'] = all_days['weight'] * unit_value

        # 5. รวมข้อมูล
        merged = pd.merge(all_days, sales_df, on='day', how='left').fillna(0).infer_objects(copy=False)
        
        # คำนวณยอดสะสม
        merged['cumulative_sales'] = merged['amount'].cumsum() + start_sales_offset
        merged['previous_sales'] = merged['cumulative_sales'] - merged['amount']
        merged['cumulative_target'] = merged['daily_target'].cumsum() + start_target_offset
        
        total_actual_month = merged['amount'].sum()
        total_target_month = merged['daily_target'].sum()

        # ==============================================================================
        # 6. [🔥 แก้ไขจุดนี้] ตัดกราฟอนาคตทิ้งให้หมด (ทั้งยอดรวม และยอดเดิม)
        # ==============================================================================
        current_date = datetime.now()
        is_current_month = (year == current_date.year and month_idx == current_date.month)
        
        # คอลัมน์ที่ต้องลบค่าทิ้งถ้าเป็นอนาคต
        cols_to_hide = ['cumulative_sales', 'previous_sales', 'amount']

        if is_current_month:
            # ถ้าเป็นเดือนปัจจุบัน: ซ่อนตั้งแต่วัน "พรุ่งนี้" เป็นต้นไป
            today_day = current_date.day
            merged.loc[merged['day'] > today_day, cols_to_hide] = None
        
        elif year > current_date.year or (year == current_date.year and month_idx > current_date.month):
             # ถ้าเป็นเดือนในอนาคต: ซ่อนทั้งหมด
             merged[cols_to_hide] = None
        
        # (หมายเหตุ: เส้นเป้าหมาย cumulative_target ไม่ถูกลบ จะแสดงยาวไปจนจบเดือนเหมือนเดิม ถูกต้องแล้ว)

        # 7. สั่งวาดกราฟ
        self._draw_matplotlib_chart(merged, year, month_name)
        self._draw_circle_progress(self.monthly_circle_area, total_actual_month, total_target_month, "monthly_canvas")
        self._draw_circle_progress(self.yearly_circle_area, total_actual_year, target_annual, "yearly_canvas")
        
    def _draw_matplotlib_chart(self, df, year, month_name):
        """
        🔥 ปรับปรุงแล้ว: แก้ปัญหา Label ซ้อนกัน + แกน X แน่นเกิน
        """
        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        plt.rcParams['font.family'] = 'Tahoma'
        fig, ax = plt.subplots(figsize=(15, 8), dpi=100)

        days = df['day']
        prev_sales = df['previous_sales']
        daily_amount = df['amount']
        total_sales = df['cumulative_sales']
        target = df['cumulative_target']

        bar_width = 0.6  # เพิ่มจาก 0.45 เพื่อให้แท่งหนาขึ้น

        # -----------------------------
        # 1. Stacked Bars (2 สี)
        # -----------------------------
        ax.bar(days, prev_sales, width=bar_width,
            color='#F59E0B', alpha=0.7, label='ยอดสะสมเดิม')

        ax.bar(days, daily_amount, bottom=prev_sales, width=bar_width,
            color='#0EA5E9', alpha=0.9, label='ยอดขายเพิ่มวันนี้')

        # -----------------------------
        # 2. Target line
        # -----------------------------
        ax.plot(days, target,
                color='#DC2626', linewidth=2.5,
                marker='o', markersize=5,
                label='เป้าหมายสะสม', zorder=5)

        # -----------------------------
        # 3. 🔥 SMART LABEL CONTROL (แก้ปัญหาซ้อนกัน)
        # -----------------------------
        MIN_SHOW = 0.15e6  # เพิ่มจาก 0.10M → แสดงน้อยลง
        LABEL_STEP = 3     # แสดงทุกๆ 3 วัน แทนที่จะทุกวัน
        
        # คำนวณ Max สำหรับกำหนด Offset แบบ Dynamic
        max_cumulative = total_sales.max(skipna=True)
        offset_base = max_cumulative * 0.015  # 1.5% ของ max
        OFFSET_LEVELS = [offset_base * 1.2, offset_base * 2.8]  # สลับสูง-ต่ำ

        for i, day in enumerate(days):
            # เงื่อนไข: แสดงเฉพาะวันที่
            # 1. มียอดขาย >= 0.15M
            # 2. เป็นทุกๆ 3 วัน (หรือวันสุดท้าย)
            # 3. ไม่ใช่วันในอนาคต
            if (
                pd.notna(total_sales[i]) and
                daily_amount[i] >= MIN_SHOW and
                (i % LABEL_STEP == 0 or i == len(days) - 1)
            ):
                offset = OFFSET_LEVELS[i % 2]  # สลับความสูง

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
                        edgecolor='#E2E8F0',
                        alpha=0.9,
                        pad=3,
                        boxstyle='round,pad=0.4'
                    )
                )

        # -----------------------------
        # 4. Cumulative Total (ท้ายเดือน)
        # -----------------------------
        last_idx = df[pd.notna(total_sales)].index.max()
        if pd.notna(last_idx):
            last_day = days[last_idx]
            last_total = total_sales[last_idx]

            ax.text(
                last_day,
                last_total + max_cumulative * 0.08,  # 8% ของ max
                f'ยอดสะสม {last_total/1e6:.1f}M',
                ha='center',
                va='bottom',
                fontsize=11,
                weight='bold',
                color='#1E293B',
                bbox=dict(
                    facecolor='#FEF3C7',
                    edgecolor='#F59E0B',
                    alpha=0.9,
                    pad=4,
                    boxstyle='round,pad=0.5'
                )
            )

        # -----------------------------
        # 5. 🔥 X-Axis Improvement (แก้ปัญหาแน่นเกิน)
        # -----------------------------
        thai_days = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]
        month_idx = self.thai_months.index(month_name) + 1

        # แสดง Tick ทุกๆ 3 วัน + วันแรก + วันสุดท้าย
        tick_days = [days.iloc[0]]  # วันแรก
        tick_days.extend([d for i, d in enumerate(days) if (i % 3 == 0 and i > 0)])
        if days.iloc[-1] not in tick_days:
            tick_days.append(days.iloc[-1])  # วันสุดท้าย

        ax.set_xticks(tick_days)
        ax.set_xticklabels(
            [f"{int(d)}\n{thai_days[datetime(year, month_idx, int(d)).weekday()]}" 
             for d in tick_days],
            fontsize=9
        )

        # Grid แนวตั้งเบาๆ ที่จุด Tick
        ax.set_xticks(days, minor=True)
        ax.grid(axis='x', which='minor', linestyle=':', alpha=0.15)

        # -----------------------------
        # 6. Y-Axis Format
        # -----------------------------
        ax.yaxis.set_major_formatter(
            plt.FuncFormatter(lambda x, p: f'{x/1e6:.1f}M')
        )

        # เผื่อพื้นที่ด้านบนเพิ่มเพื่อรองรับ Label
        ymax = max(target.max(skipna=True), total_sales.max(skipna=True))
        ax.set_ylim(0, ymax * 1.25)  # เพิ่มจาก 1.22 → 1.25

        # -----------------------------
        # 7. Styling
        # -----------------------------
        ax.set_title(
            f"วิเคราะห์ยอดขายและการเติบโตรายวัน - {month_name} {year}",
            fontsize=16, weight='bold', pad=20
        )

        ax.legend(loc='upper left', fontsize=10, frameon=True, 
                 framealpha=0.95, edgecolor='#E2E8F0', ncol=3)
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        
        # เพิ่มสีพื้นหลังอ่อนๆ
        ax.set_facecolor('#FAFAFA')
        fig.patch.set_facecolor('white')

        plt.subplots_adjust(top=0.88, bottom=0.12, left=0.08, right=0.95)

        self.canvas = FigureCanvasTkAgg(fig, master=self.main_chart_canvas_area)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _draw_circle_progress(self, parent_frame, actual, target, canvas_attr):
        """
        🔥 แก้ไขแล้ว: รวมฟังก์ชันซ้ำให้เหลือ 1 ฟังก์ชันเดียว
        """
        old_canvas = getattr(self, canvas_attr, None)
        if old_canvas: 
            old_canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots(figsize=(2.8, 2.8), dpi=85)
        fig.patch.set_alpha(0)
        
        pct = (actual / target * 100) if target > 0 else 0
        
        # เลือกสีตามเปอร์เซนต์
        if pct >= 100:
            color = '#10B981'  # เขียว - ผ่านเป้า
        elif pct >= 80:
            color = '#F59E0B'  # ส้ม - ใกล้เป้า
        else:
            color = '#EF4444'  # แดง - ต่ำกว่าเป้า
        
        # วาด Donut Chart
        ax.pie([min(pct, 100), max(0, 100-pct)], 
               colors=[color, '#F1F5F9'], 
               startangle=90, 
               counterclock=False, 
               wedgeprops={'width': 0.35, 'edgecolor': 'white', 'linewidth': 2})
        
        # ตัวเลขตรงกลาง
        ax.text(0, 0.15, f'{pct:.1f}%', 
                ha='center', va='center', 
                fontsize=14, weight='bold', color='#1E293B')
        
        ax.text(0, -0.25, f'{actual/1000000:.1f}M', 
                ha='center', va='center', 
                fontsize=11, color='#64748B')
        
        ax.axis('equal')
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        setattr(self, canvas_attr, canvas)