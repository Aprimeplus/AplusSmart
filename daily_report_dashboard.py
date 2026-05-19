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

        # ✅ pack ฝั่งขวาก่อนเสมอ เพื่อให้ tkinter จองพื้นที่ไว้ก่อน
        # ฝั่งขวา: แถบสรุปเป้าหมาย (Side Panel)
        self.right_summary_panel = CTkFrame(self.display_container, fg_color="#F8FAFC", width=260)
        self.right_summary_panel.pack(side="right", fill="y", padx=(5, 0))
        self.right_summary_panel.pack_propagate(False)

        # ฝั่งซ้าย: กราฟแท่งรายวัน (ขยายเต็มที่ หลังจาก right_summary_panel จองพื้นที่แล้ว)
        self.left_graph_area = CTkFrame(self.display_container, fg_color="white")
        self.left_graph_area.pack(side="left", fill="both", expand=True)
        
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
        self._map_after_id = None
        self.after(500, self._update_chart)

        # ✅ วาดกราฟใหม่ทุกครั้งที่ Tab นี้ถูกเปิดขึ้นมา (ได้ขนาดจริง + ข้อมูลล่าสุด)
        self.bind("<Map>", self._on_map)

    def _on_map(self, event=None):
        """เรียกวาดกราฟใหม่เมื่อ tab ถูกแสดง — debounce 200ms"""
        if self._map_after_id:
            self.after_cancel(self._map_after_id)
        self._map_after_id = self.after(200, self._do_map_update)

    def _do_map_update(self):
        self._map_after_id = None
        self._update_chart()

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
        🎨 ปรับปรุง V.8 - Smart Flip Tooltip
        ✅ แก้ปัญหา Tooltip ตกขอบขวา (Auto Flip ซ้าย/ขวา ตามตำแหน่งเมาส์)
        """
        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        plt.rcParams['font.family'] = ['Tahoma', 'Segoe UI Emoji']

        # ✅ คำนวณขนาดกราฟจากพื้นที่จริงที่มีอยู่ ไม่ใช้ค่าตายตัว
        self.update_idletasks()
        w_px = max(self.left_graph_area.winfo_width(), 800)
        h_px = max(self.left_graph_area.winfo_height(), 400)
        fig, ax = plt.subplots(figsize=(w_px / 100, h_px / 100), dpi=100)

        days = df['day']
        prev_sales = df['previous_sales']
        daily_amount = df['amount']
        total_sales = df['cumulative_sales']
        target = df['cumulative_target']

        bar_width = 0.6
        total_bars_count = len(days) # นับจำนวนวันทั้งหมดเพื่อใช้คำนวณตำแหน่ง

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 📊 1. STACKED BAR CHART
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        rects1 = ax.bar(days, prev_sales, width=bar_width,
                        color='#F59E0B', alpha=0.75, 
                        label='ยอดสะสมเดิม', zorder=2,
                        edgecolor='#D97706', linewidth=0.5)

        rects2 = ax.bar(days, daily_amount, bottom=prev_sales, width=bar_width,
                        color='#0EA5E9', alpha=0.9, 
                        label='ยอดขายเพิ่มวันนี้', zorder=2,
                        edgecolor='#0284C7', linewidth=0.5)

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 🎯 2. TARGET LINE
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        ax.plot(days, target,
                color='#DC2626', linewidth=2.5,
                marker='o', markersize=6, markerfacecolor='#FEE2E2',
                markeredgewidth=2, markeredgecolor='#DC2626',
                label='เป้าหมายสะสม', zorder=5)

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 💬 3. TOOLTIP (SMART POSITIONING)
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        annot = ax.annotate("", xy=(0,0), xytext=(25, 25),
                            textcoords="offset points",
                            bbox=dict(
                                boxstyle="round,pad=0.8",
                                fc="#F8FAFC",
                                ec="#D1D5DB",
                                alpha=0.98,
                                linewidth=1.2
                            ),
                            arrowprops=dict(
                                arrowstyle="->",
                                fc="#64748B", 
                                ec="#64748B",
                                lw=1.5,
                                alpha=0.8,
                                connectionstyle="arc3,rad=0.2"
                            ),
                            fontsize=10.5, 
                            color='#1E293B',
                            fontfamily='Tahoma',
                            fontweight='normal',
                            zorder=25)
        annot.set_visible(False)

        def update_annot(bar, idx):
            x = bar.get_x() + bar.get_width() / 2
            y = bar.get_y() + bar.get_height()
            annot.xy = (x, y)
            
            # --- 🔥 SMART FLIP LOGIC (แก้ปัญหาตกขอบ) ---
            # ถ้าตำแหน่งแท่งกราฟ เกิน 60% ของกราฟทั้งหมด ให้ปัด Tooltip ไปทางซ้าย
            if idx > (total_bars_count * 0.6):
                # ปัดซ้าย (xytext ติดลบ) + ปรับโค้งลูกศรกลับด้าน
                annot.set_position((-20, 25))  
                annot.arrow_patch.set_connectionstyle("arc3,rad=-0.2") # rad ติดลบเพื่อกลับด้านโค้ง
                # ปรับตำแหน่ง Text alignment ให้ชิดขวาของกล่อง (optional)
                annot.set_ha('right') 
            else:
                # ปัดขวา (ค่าปกติ)
                annot.set_position((25, 25))
                annot.arrow_patch.set_connectionstyle("arc3,rad=0.2")
                annot.set_ha('left')

            # --- Data Retrieval ---
            day_val = days.iloc[idx]
            daily_val = daily_amount.iloc[idx]
            total_val = total_sales.iloc[idx]
            target_val = target.iloc[idx] if idx < len(target) else 0
            
            progress = (total_val / target_val * 100) if target_val > 0 else 0
            gap = total_val - target_val
            
            # --- Smart Format ---
            def smart_format(val):
                if abs(val) >= 1e6: return f"{val/1e6:,.2f}M"
                elif abs(val) >= 1e3: return f"{val/1e3:,.1f}K"
                else: return f"{val:,.0f}"
            
            # --- Status Logic ---
            if progress >= 100:
                status_icon, status_text, border_color = "✅", "บรรลุเป้า", "#10B981"
            elif progress >= 90:
                status_icon, status_text, border_color = "🟢", "ใกล้เป้า", "#84CC16"
            elif progress >= 75:
                status_icon, status_text, border_color = "🟡", "พอใช้", "#F59E0B"
            else:
                status_icon, status_text, border_color = "🔴", "ห่างเป้า", "#EF4444"
            
            # --- Visual Elements ---
            bar_len = 12
            filled = min(int(progress / 100 * bar_len), bar_len)
            p_bar = "█" * filled + "░" * (bar_len - filled)
            
            thai_days_full = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
            m_idx = self.thai_months.index(month_name) + 1
            d_name = thai_days_full[datetime(year, m_idx, int(day_val)).weekday()]
            
            if gap > 0: gap_icon, gap_lbl, gap_txt = "📈", "เกินเป้า", f"+{smart_format(gap)}"
            elif gap < 0: gap_icon, gap_lbl, gap_txt = "📉", "ห่างเป้า", smart_format(gap)
            else: gap_icon, gap_lbl, gap_txt = "🎯", "ตรงเป้า", "0"
            
            # --- Text Content ---
            sep_line = "─" * 22 
            text = (
                f"📅  {int(day_val)} {month_name} ({d_name})\n"
                f"{sep_line}\n"
                f"💰  วันนี้       {smart_format(daily_val):>9}\n"
                f"📦  สะสม       {smart_format(total_val):>9}\n"
                f"🎯  เป้าหมาย     {smart_format(target_val):>9}\n\n"
                f"{status_icon}  {status_text}  ({progress:.1f}%)\n"
                f"{p_bar}\n"
                f"{gap_icon}  {gap_lbl}: {gap_txt}"
            )
            
            annot.set_text(text)
            annot.get_bbox_patch().set_edgecolor(border_color)

        def hover(event):
            vis = annot.get_visible()
            if event.inaxes == ax:
                for idx, bar in enumerate(rects2):
                    if bar.contains(event)[0]:
                        update_annot(bar, idx)
                        annot.set_visible(True)
                        event.canvas.draw_idle()
                        return
                for idx, bar in enumerate(rects1):
                    if bar.contains(event)[0]:
                        update_annot(bar, idx)
                        annot.set_visible(True)
                        event.canvas.draw_idle()
                        return
            
            if vis:
                annot.set_visible(False)
                event.canvas.draw_idle()

        fig.canvas.mpl_connect("motion_notify_event", hover)

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 🏷️ 4. STATIC LABELS
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        MIN_SHOW = 0.15e6
        LABEL_STEP = 3
        
        max_cumulative = total_sales.max(skipna=True) if not total_sales.empty else 0
        offset_base = max_cumulative * 0.015
        OFFSET_LEVELS = [offset_base * 1.2, offset_base * 2.8]

        for i, day in enumerate(days):
            if (pd.notna(total_sales[i]) and 
                daily_amount[i] >= MIN_SHOW and 
                (i % LABEL_STEP == 0 or i == len(days) - 1)):
                
                offset = OFFSET_LEVELS[i % 2]
                if daily_amount[i] >= 1e6: label_text = f'+{daily_amount[i] / 1e6:.2f}M'
                else: label_text = f'+{daily_amount[i] / 1e3:.0f}K'
                
                ax.text(day, total_sales[i] + offset, label_text,
                    ha='center', va='bottom', fontsize=9, weight='bold', color='#0369A1', zorder=10,
                    bbox=dict(facecolor='white', edgecolor='#BAE6FD', alpha=0.9, pad=2.5, boxstyle='round,pad=0.35', linewidth=1.2))

        last_idx = df[pd.notna(total_sales)].index.max()
        if pd.notna(last_idx):
            last_day = days[last_idx]
            last_total = total_sales[last_idx]
            last_target = target[last_idx] if last_idx < len(target) else 0
            achievement = (last_total / last_target * 100) if last_target > 0 else 0
            
            if achievement >= 100: badge_bg, badge_border, badge_icon = '#D1FAE5', '#10B981', '✅'
            elif achievement >= 90: badge_bg, badge_border, badge_icon = '#FEF3C7', '#F59E0B', '🟡'
            else: badge_bg, badge_border, badge_icon = '#FEE2E2', '#EF4444', '⚠️'
            
            ax.text(last_day, last_total + max_cumulative * 0.09,
                f'{badge_icon} ยอดสะสม {last_total/1e6:.2f}M ({achievement:.1f}%)',
                ha='center', va='bottom', fontsize=11, weight='bold', color='#1E293B', zorder=10,
                bbox=dict(facecolor=badge_bg, edgecolor=badge_border, alpha=0.95, pad=5, boxstyle='round,pad=0.6', linewidth=1.5))

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 📐 5. AXES & STYLING
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        thai_days_short = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]
        month_idx_int = self.thai_months.index(month_name) + 1
        
        tick_days = [days.iloc[0]]
        tick_days.extend([d for i, d in enumerate(days) if (i % 3 == 0 and i > 0)])
        if days.iloc[-1] not in tick_days: tick_days.append(days.iloc[-1])

        ax.set_xticks(tick_days)
        ax.set_xticklabels([f"{int(d)}\n{thai_days_short[datetime(year, month_idx_int, int(d)).weekday()]}" for d in tick_days], fontsize=9, color='#374151')
        ax.set_xticks(days, minor=True)
        ax.grid(axis='x', which='minor', linestyle=':', alpha=0.2, color='#D1D5DB')
        
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.1f}M'))
        max_target = target.max(skipna=True) if not target.empty else 0
        max_total = total_sales.max(skipna=True) if not total_sales.empty else 0
        ymax = max(max_target, max_total)
        if ymax > 0: ax.set_ylim(0, ymax * 1.28)

        ax.set_title(f"📊 วิเคราะห์ยอดขายและการเติบโตรายวัน - {month_name} {year}", fontsize=17, weight='bold', pad=25, color='#1F2937')
        ax.legend(loc='upper left', fontsize=10.5, frameon=True, framealpha=0.97, edgecolor='#D1D5DB', fancybox=True, shadow=False, ncol=3, columnspacing=1.5, labelspacing=0.8)
        ax.grid(axis='y', linestyle='--', alpha=0.35, color='#D1D5DB', linewidth=0.8)
        ax.set_facecolor('#F9FAFB')
        fig.patch.set_facecolor('white')
        
        for spine in ax.spines.values():
            spine.set_edgecolor('#E5E7EB')
            spine.set_linewidth(1)
        
        plt.subplots_adjust(top=0.88, bottom=0.12, left=0.08, right=0.95)

        self.canvas = FigureCanvasTkAgg(fig, master=self.left_graph_area)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _draw_circle_progress(self, parent_frame, actual, target, canvas_attr):
        """
        🔥 แก้ไขแล้ว: รวมฟังก์ชันซ้ำให้เหลือ 1 ฟังก์ชันเดียว
        """
        old_canvas = getattr(self, canvas_attr, None)
        if old_canvas: 
            old_canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots(figsize=(2.5, 2.5), dpi=85)
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