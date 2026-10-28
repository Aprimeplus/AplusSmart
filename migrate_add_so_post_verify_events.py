from main_app import AppContainer


def migrate_db():
    print("Starting migration: so_post_verify_events (เหตุการณ์แก้ SO/PO หลัง HR Verified/Paid สำหรับแจ้ง Director/HR)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS so_post_verify_events (
                    id            SERIAL PRIMARY KEY,
                    so_number     TEXT NOT NULL,
                    event_type    TEXT NOT NULL,
                    po_number     TEXT,
                    old_status    TEXT,
                    new_status    TEXT,
                    detail        TEXT,
                    changed_by    TEXT,
                    changed_at    TIMESTAMP NOT NULL DEFAULT NOW(),
                    sale_key      TEXT,
                    customer_name TEXT,
                    ack_by        TEXT,
                    ack_at        TIMESTAMP
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_pv_events_so ON so_post_verify_events (so_number)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_pv_events_ack ON so_post_verify_events (ack_at)")
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
