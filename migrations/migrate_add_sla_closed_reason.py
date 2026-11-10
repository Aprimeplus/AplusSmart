from main_app import AppContainer


def migrate_db():
    print("Starting migration: sla_benchmark.closed_reason (ปิดงานค้างโดยไม่นับ SLA)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("ALTER TABLE sla_benchmark ADD COLUMN IF NOT EXISTS closed_reason TEXT")
            cursor.execute("ALTER TABLE sla_benchmark ADD COLUMN IF NOT EXISTS closed_at TIMESTAMP")
        conn.commit()
        print("Done.")
    except Exception as e:
        conn.rollback()
        print(f"Migration failed: {e}")
    finally:
        app.release_connection(conn)


if __name__ == "__main__":
    migrate_db()
