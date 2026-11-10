from main_app import AppContainer


def migrate_db():
    print("Starting migration: sla_copy_log (log ทุกครั้งที่ copy Short Note สำหรับตรวจสอบงานค้าง SLA)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS sla_copy_log (
                    id          SERIAL PRIMARY KEY,
                    so_number   TEXT NOT NULL,
                    user_key    TEXT,
                    method      TEXT,
                    result      TEXT,
                    attempts    INTEGER,
                    error_msg   TEXT,
                    copied_at   TIMESTAMP,
                    logged_at   TIMESTAMP NOT NULL DEFAULT NOW()
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sla_copy_log_so ON sla_copy_log (so_number)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sla_copy_log_logged ON sla_copy_log (logged_at)")
        conn.commit()
        print("Migration completed.")
    except Exception as e:
        conn.rollback()
        print(f"Migration failed: {e}")
    finally:
        app.release_connection(conn)
        app.destroy()


if __name__ == "__main__":
    migrate_db()
