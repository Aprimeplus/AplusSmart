"""SO/PO ที่ถูกแก้หลัง HR ตรวจ (HR Verified) หรือจ่ายแล้ว (Paid)

1) sync_so_status_after_po_save — กฎเปลี่ยนสถานะ SO ตอนจัดซื้อบันทึก PO (แทนการสั่ง UPDATE ตรงๆ เดิมที่ดึงสถานะถอยหลังมั่ว)
     • Paid            → ไม่แตะสถานะ (บันทึกเหตุการณ์ + แจ้ง ถ้าเป็น PO จริง ไม่ใช่ร่าง)
     • HR Verified     → PO จริง: กลับไป 'PO Sent' ให้ HR ตรวจใหม่ + บันทึกเหตุการณ์ + แจ้ง / ร่าง: ไม่แตะ
     • สถานะอื่นในคิว  → PO จริง: 'PO Sent' / ร่าง: 'PO In Progress' (ยกเว้น PO Sent อยู่แล้ว ไม่ถอยหลัง)
     • สถานะที่ถูกพัก (ขอแก้ไข SO) / ยกเลิก → ไม่แตะ
2) record_event — บันทึกลง so_post_verify_events และส่ง notification ให้ Director + HR ทุกคนที่ Active
   (ผู้จ่ายเงิน/กรรมการต้องรู้ว่า ใคร แก้อะไร ยังไง และ SO ไหนกลับไปรอตรวจแล้วอาจตกหล่น)
"""

PROTECTED_STATUSES = ('Cancelled', 'Cancelled by PU', 'Pending Transport Cancel',
                      'Edit Requested', 'Edit Approved', 'Edit Review')

EVENT_LABELS = {
    'PO_AFTER_VERIFY':        'มีการบันทึก/แก้ไข PO หลัง HR ตรวจแล้ว → SO กลับไปรอ HR ตรวจใหม่',
    'PO_ADDED_AFTER_PAID':    'มีการบันทึก/แก้ไข PO หลังจ่ายค่าคอมแล้ว (ไม่เปลี่ยนสถานะ ควรตรวจต้นทุนที่เปลี่ยน)',
    'PO_EDITED_AFTER_VERIFY': 'PO ถูกแก้ไขหลัง HR ตรวจแล้ว → SO กลับไปรอ HR ตรวจใหม่',
    'PO_EDITED_AFTER_PAID':   'PO ถูกแก้ไขหลังจ่ายค่าคอมแล้ว (ไม่เปลี่ยนสถานะ ควรตรวจต้นทุนที่เปลี่ยน)',
    'SO_EDITED_AFTER_VERIFY': 'SO ถูกแก้ไขหลัง HR ตรวจแล้ว → SO กลับไปรอ HR ตรวจใหม่',
}

# สถานะที่ถือว่า "ยังรอ HR ตรวจ/จ่าย" — ใช้แยกเหตุการณ์ที่ตกค้างอยู่
FINAL_STATUSES = ('HR Verified', 'Paid')

# SO ที่ไม่ต้องแจ้ง/ไม่ต้องแสดงในหน้าตกค้าง: งาน CT เป็นของบริษัท ไม่คิดค่าคอม และ sale ที่ Inactive แล้ว
NON_COMMISSION_SALE_KEYS = ('Sale Center', 'CHARITA-CT')


def is_eligible_sale(cursor, sale_key):
    if not sale_key or sale_key in NON_COMMISSION_SALE_KEYS:
        return False
    cursor.execute("SELECT status FROM sales_users WHERE sale_key = %s", (sale_key,))
    row = cursor.fetchone()
    return bool(row and row[0] == 'Active')


def record_event(cursor, so_number, event_type, by_user, old_status, new_status, po_number=None, detail=''):
    """บันทึกเหตุการณ์ + แจ้งเตือน Director/HR (ทำภายใน transaction ของผู้เรียก ผู้เรียก commit เอง)
    ครอบด้วย savepoint: ถ้าส่วนแจ้งเตือนพัง จะย้อนเฉพาะส่วนนี้ ไม่ทำให้การบันทึก PO/SO ของผู้เรียกพังตาม"""
    cursor.execute("SAVEPOINT pv_alert")
    try:
        _record_event_inner(cursor, so_number, event_type, by_user, old_status, new_status, po_number, detail)
        cursor.execute("RELEASE SAVEPOINT pv_alert")
    except Exception as e:
        cursor.execute("ROLLBACK TO SAVEPOINT pv_alert")
        print(f"[post_verify_alerts] record_event error (SO {so_number}): {e}")


def _record_event_inner(cursor, so_number, event_type, by_user, old_status, new_status, po_number=None, detail=''):
    """บันทึกเหตุการณ์ + แจ้งเตือน Director/HR (ทำภายใน transaction ของผู้เรียก ผู้เรียก commit เอง)"""
    cursor.execute("""
        SELECT sale_key, customer_name, commission_month, commission_year
        FROM commissions WHERE so_number = %s AND is_active = 1 LIMIT 1
    """, (so_number,))
    row = cursor.fetchone()
    sale_key, customer, cm, cy = (row if row else (None, None, None, None))
    if not is_eligible_sale(cursor, sale_key):
        return          # งาน CT (ของบริษัท ไม่คิดค่าคอม) หรือ sale ที่ Inactive → ไม่บันทึก ไม่แจ้งเตือน
    cursor.execute("""
        INSERT INTO so_post_verify_events
            (so_number, event_type, po_number, old_status, new_status, detail, changed_by, sale_key, customer_name)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (so_number, event_type, po_number, old_status, new_status, detail, by_user, sale_key, customer))

    label = EVENT_LABELS.get(event_type, event_type)
    msg = (f"⚠ แจ้งเตือน: SO {so_number}\n{label}\n"
           f"โดย: {by_user or '-'}"
           + (f"\nPO: {po_number}" if po_number else "")
           + f"\nสถานะ SO: {old_status} → {new_status}"
           + (f"\n{detail}" if detail else "")
           + (f"\nรอบค่าคอม: {cm}/{cy}" if cm else "")
           + "\n(ดูที่หน้า HR แท็บ \"6. แจ้งเตือน/SO ตกค้าง\")")
    cursor.execute("SELECT sale_key FROM sales_users WHERE role IN ('Director', 'HR') AND status = 'Active'")
    for (uk,) in cursor.fetchall():
        cursor.execute("""
            INSERT INTO notifications (user_key_to_notify, message, is_read)
            VALUES (%s, %s, FALSE)
        """, (uk, msg))


def sync_so_status_after_po_save(cursor, so_number, po_number, is_draft, by_user, detail=''):
    """กฎใหม่ (ทางเลือก ข.) แทน UPDATE commissions SET status='PO Sent'/'PO In Progress' แบบไม่ดูสถานะเดิม
    คืนค่าสถานะ SO หลังจัดการ (หรือ None ถ้าไม่พบ SO)"""
    if not so_number:
        return None
    cursor.execute("SELECT id, status FROM commissions WHERE so_number = %s AND is_active = 1 FOR UPDATE", (so_number,))
    row = cursor.fetchone()
    if not row:
        return None
    cid, st = row

    if st in PROTECTED_STATUSES:
        return st

    if st == 'Paid':
        if not is_draft:
            record_event(cursor, so_number, 'PO_ADDED_AFTER_PAID', by_user, 'Paid', 'Paid', po_number, detail)
        return st

    if st == 'HR Verified':
        if is_draft:
            return st
        cursor.execute("UPDATE commissions SET status = 'PO Sent' WHERE id = %s", (cid,))
        record_event(cursor, so_number, 'PO_AFTER_VERIFY', by_user, 'HR Verified', 'PO Sent', po_number, detail)
        return 'PO Sent'

    if is_draft:
        if st == 'PO Sent':
            return st                      # ไม่ถอยหลังจากที่ส่งให้ HR แล้ว
        if st != 'PO In Progress':
            cursor.execute("UPDATE commissions SET status = 'PO In Progress' WHERE id = %s", (cid,))
        return 'PO In Progress'
    if st != 'PO Sent':
        cursor.execute("UPDATE commissions SET status = 'PO Sent' WHERE id = %s", (cid,))
    return 'PO Sent'


def record_po_edit_if_post_verify(cursor, so_number, po_number, by_user, detail=''):
    """เรียกเมื่อ PO ที่มีอยู่แล้วถูกแก้ไข: ถ้า SO อยู่ HR Verified → กลับ PO Sent ให้ HR ตรวจใหม่ + แจ้ง
    ถ้า Paid → แจ้งอย่างเดียว ไม่เปลี่ยนสถานะ สถานะอื่นไม่ทำอะไร คืนสถานะ SO หลังจัดการ"""
    if not so_number:
        return None
    cursor.execute("SELECT id, status FROM commissions WHERE so_number = %s AND is_active = 1 FOR UPDATE", (so_number,))
    row = cursor.fetchone()
    if not row:
        return None
    cid, st = row
    if st == 'HR Verified':
        cursor.execute("UPDATE commissions SET status = 'PO Sent' WHERE id = %s", (cid,))
        record_event(cursor, so_number, 'PO_EDITED_AFTER_VERIFY', by_user, 'HR Verified', 'PO Sent', po_number, detail)
        return 'PO Sent'
    if st == 'Paid':
        record_event(cursor, so_number, 'PO_EDITED_AFTER_PAID', by_user, 'Paid', 'Paid', po_number, detail)
    return st


def record_po_edit_by_id(cursor, po_id, by_user, detail=''):
    """เหมือน record_po_edit_if_post_verify แต่รับรหัส PO (ค้นเลข SO / เลข PO ให้เอง) — ไว้เสียบในหน้าต่างแก้ PO"""
    cursor.execute("SELECT so_number, po_number FROM purchase_orders WHERE id = %s", (po_id,))
    row = cursor.fetchone()
    if not row:
        return None
    return record_po_edit_if_post_verify(cursor, row[0], row[1], by_user, detail)
