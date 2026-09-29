import os
from customtkinter import CTkFrame, CTkTabview
import matplotlib

try:
    font_path = os.path.join('resources', 'THSarabunNew.ttf')
    if os.path.exists(font_path):
        from matplotlib.font_manager import fontManager
        fontManager.addfont(font_path)
        matplotlib.rc('font', family='TH Sarabun New')
except Exception:
    pass

from daily_report_widget import DailyReportWidget
# Sale Revenue Report ใช้ widget เดียวกับหน้า "เป้าการขาย" ของ Sales Manager แล้ว
# (เดิมเป็นโค้ดคนละก้อนที่ port มาแยกกัน ทำให้แก้บัคแล้วลืมแก้อีกจุด — รวมเป็นไฟล์กลางที่ sales_target_widget.py)
from sales_target_widget import SalesTargetWidget


# ─────────────────────────────────────────────────────────────────────────────
#  Management Report Screen
# ─────────────────────────────────────────────────────────────────────────────
class ManagementReportScreen(CTkFrame):
    def __init__(self, master, app_container, user_key=None, user_name=None, user_role=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app_container = app_container
        self.user_key = user_key
        self.user_name = user_name
        self.user_role = user_role

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        tabs = CTkTabview(self, segmented_button_selected_color="#1E40AF")
        tabs.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Tab 1 — SO Daily Report
        tab_daily = tabs.add("📋 SO Daily Report")
        tab_daily.grid_columnconfigure(0, weight=1)
        tab_daily.grid_rowconfigure(0, weight=1)
        daily_widget = DailyReportWidget(
            master=tab_daily,
            app_container=app_container,
            fg_color="transparent",
        )
        daily_widget.grid(row=0, column=0, sticky="nsew")

        # Tab 2 — Sale Revenue Report
        tab_revenue = tabs.add("📈 Sale Revenue Report")
        tab_revenue.grid_columnconfigure(0, weight=1)
        tab_revenue.grid_rowconfigure(0, weight=1)
        revenue_widget = SalesTargetWidget(
            master=tab_revenue,
            app_container=app_container,
            fg_color="transparent",
        )
        revenue_widget.grid(row=0, column=0, sticky="nsew")
