"""
สคริปต์ทดสอบหน้าจอ project_screen.py แบบเดี่ยว ๆ (standalone)
ไม่ผ่านหน้า login และไม่แตะการทำงานของ main_app.py เลย — ใช้แค่ทดสอบ UI/DB ของ Phase 1
รัน: python test_project_screen.py
"""
import customtkinter as ctk
from psycopg2 import pool
import sqlalchemy
from project_screen import ProjectScreen

DB_PARAMS = dict(host="Server-APrime", dbname="aplus_com_test", user="app_user", password="cailfornia123")


class _FakeAppContainer:
    def __init__(self):
        self.db_pool = pool.SimpleConnectionPool(1, 5, **DB_PARAMS)
        self.pg_engine = sqlalchemy.create_engine(
            f"postgresql+psycopg2://{DB_PARAMS['user']}:{DB_PARAMS['password']}@{DB_PARAMS['host']}:5432/{DB_PARAMS['dbname']}?client_encoding=utf8"
        )

    def get_connection(self):
        return self.db_pool.getconn()

    def release_connection(self, conn):
        if conn:
            self.db_pool.putconn(conn)


if __name__ == "__main__":
    ctk.set_appearance_mode("light")
    root = ctk.CTk()
    root.title("ทดสอบ: จัดการโครงการ (Phase 1)")
    root.geometry("1100x650")

    app_container = _FakeAppContainer()
    screen = ProjectScreen(root, app_container, user_key="TEST_USER")
    screen.pack(fill="both", expand=True)

    root.mainloop()
