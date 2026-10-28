"""SO ที่จ่ายค่าคอมมิชชั่นแล้ว (Paid) = จบแล้ว ล็อก SO/PO ทั้งหมด ดูได้อย่างเดียว ไม่มีใครแก้ได้

ฟังก์ชันตรวจสถานะ + ข้อความกลาง ใช้เสียบที่จุดบันทึก/ตีกลับ/แก้ไขของ SO และ PO ทุกจุด
"""
from tkinter import messagebox

LOCK_MSG = ("SO นี้จ่ายค่าคอมมิชชั่นแล้ว (Paid) — ถือว่าจบแล้ว\n"
            "ข้อมูล SO และ PO ถูกล็อก ดูได้อย่างเดียว แก้ไขไม่ได้")


def _fetch_one(app_container, sql, params):
    conn = app_container.get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchone()
    finally:
        app_container.release_connection(conn)


def is_so_paid(app_container, so_number):
    if not so_number:
        return False
    return bool(_fetch_one(app_container,
        "SELECT 1 FROM commissions WHERE so_number = %s AND is_active = 1 AND status = 'Paid' LIMIT 1", (so_number,)))


def is_po_paid(app_container, po_id):
    if po_id is None:
        return False
    return bool(_fetch_one(app_container, """
        SELECT 1 FROM purchase_orders p
        JOIN commissions c ON c.so_number = p.so_number AND c.is_active = 1
        WHERE p.id = %s AND c.status = 'Paid' LIMIT 1""", (int(po_id),)))


def is_commission_paid(app_container, commission_id):
    if commission_id is None:
        return False
    return bool(_fetch_one(app_container,
        "SELECT 1 FROM commissions WHERE id = %s AND status = 'Paid' LIMIT 1", (int(commission_id),)))


# แบบใช้ cursor ที่เปิดอยู่แล้ว (ภายใน transaction/thread ที่แสดง messagebox ไม่ได้)
def cur_is_so_paid(cursor, so_number):
    if not so_number:
        return False
    cursor.execute("SELECT 1 FROM commissions WHERE so_number = %s AND is_active = 1 AND status = 'Paid' LIMIT 1",
                   (so_number,))
    return cursor.fetchone() is not None


def cur_is_po_paid(cursor, po_id):
    cursor.execute("""
        SELECT 1 FROM purchase_orders p
        JOIN commissions c ON c.so_number = p.so_number AND c.is_active = 1
        WHERE p.id = %s AND c.status = 'Paid' LIMIT 1""", (int(po_id),))
    return cursor.fetchone() is not None


def warn_locked(parent=None):
    messagebox.showwarning("ล็อก — ดูได้อย่างเดียว", LOCK_MSG, parent=parent)
