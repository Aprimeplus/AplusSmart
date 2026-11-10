from main_app import AppContainer


def migrate_db():
    print("Starting migration: so_change_requests (Sale ขอแก้ไข SO ที่ส่งไปแล้ว → SM อนุมัติ 2 ด่าน)...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            # ตารางนี้แยกจาก so_edit_requests เดิม (ของเดิมใช้ขอเปลี่ยนรอบเดือนค่าคอมอย่างเดียว)
            # status ของคำขอ: Requested (รอ SM อนุมัติสิทธิ์แก้) → Approved (เซลส์แก้ได้) → Submitted (รอ SM อนุมัติผล)
            #   → Completed | Rejected (SM ไม่อนุมัติสิทธิ์) | Reverted (SM ไม่อนุมัติผล/หมดอายุ ย้อนค่าเดิม)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS so_change_requests (
                    id                  SERIAL PRIMARY KEY,
                    so_number           TEXT NOT NULL,
                    original_commission_id INTEGER NOT NULL,
                    new_commission_id   INTEGER,
                    sale_key            TEXT NOT NULL,
                    original_status     TEXT NOT NULL,
                    reason              TEXT NOT NULL,
                    status              TEXT NOT NULL DEFAULT 'Requested',
                    affects_po          BOOLEAN NOT NULL DEFAULT FALSE,
                    requested_at        TIMESTAMP NOT NULL DEFAULT NOW(),
                    approved_by         TEXT,
                    approved_at         TIMESTAMP,
                    expires_at          TIMESTAMP,
                    submitted_at        TIMESTAMP,
                    result_by           TEXT,
                    result_at           TIMESTAMP,
                    reject_reason       TEXT
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_so_chg_req_so ON so_change_requests (so_number)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_so_chg_req_status ON so_change_requests (status)")
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
