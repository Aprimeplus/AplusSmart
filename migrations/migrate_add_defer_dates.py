import psycopg2
from main_app import AppContainer

def migrate_db():
    print("Starting migration: Adding defer date/remarks/risk columns to commissions table...")
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS expected_delivery_date DATE DEFAULT NULL")
            print("- Added 'expected_delivery_date' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS expected_payment_date DATE DEFAULT NULL")
            print("- Added 'expected_payment_date' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS defer_remarks TEXT DEFAULT NULL")
            print("- Added 'defer_remarks' column")
            cursor.execute("ALTER TABLE commissions ADD COLUMN IF NOT EXISTS is_collection_risk BOOLEAN DEFAULT FALSE")
            print("- Added 'is_collection_risk' column")
        conn.commit()
        print("Migration completed successfully!")
    except Exception as e:
        conn.rollback()
        print(f"Migration failed: {e}")
    finally:
        app.release_connection(conn)
        app.destroy()

if __name__ == "__main__":
    migrate_db()
