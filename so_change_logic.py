"""ตรรกะกลางของ flow "Sale ขอแก้ไข SO ที่ส่งไปแล้ว" (ตาราง so_change_requests)

สถานะคำขอ: Requested → Approved → Submitted → Completed | Rejected | Reverted
สถานะ SO ระหว่างทาง: Edit Requested / Edit Approved / Edit Review (ถูกพัก PU/HR ไม่เห็นในคิว)
"""

LOCK_STATUSES = ('Edit Requested', 'Edit Approved', 'Edit Review')


def revert_change_request(cursor, req_id, result_by, reason):
    """ย้อน SO กลับเป็นค่าเดิมก่อนขอแก้ไข (ไม่อนุมัติผล / สิทธิ์หมดอายุ)
    ต้องเรียกภายใน transaction ของผู้เรียก (ผู้เรียกเป็นคน commit)
    คืนค่า (so_number, sale_key) หรือ None ถ้าคำขอนี้จบไปแล้ว"""
    cursor.execute("""
        SELECT so_number, original_commission_id, new_commission_id, sale_key, original_status, status
        FROM so_change_requests WHERE id = %s FOR UPDATE
    """, (req_id,))
    req = cursor.fetchone()
    if not req or req[5] not in ('Approved', 'Submitted'):
        return None
    so_number, orig_id, new_id, sale_key, original_status = req[0], req[1], req[2], req[3], req[4]

    if new_id is not None:
        # แถวใหม่ที่เซลส์แก้ไว้: ปิดทิ้ง (เก็บไว้เป็นประวัติ ไม่ลบ)
        cursor.execute("UPDATE commissions SET is_active = 0 WHERE id = %s", (new_id,))
    cursor.execute("""
        UPDATE commissions SET is_active = 1, status = %s WHERE id = %s
    """, (original_status, orig_id))
    cursor.execute("""
        UPDATE so_change_requests
        SET status = 'Reverted', result_by = %s, result_at = NOW(), reject_reason = %s
        WHERE id = %s
    """, (result_by, reason, req_id))
    cursor.execute("""
        INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id)
        VALUES (%s, %s, FALSE, %s)
    """, (sale_key, f"SO {so_number} ถูกย้อนกลับเป็นค่าเดิม: {reason}", orig_id))
    cursor.execute("""
        INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
        VALUES ('SO Change Reverted', 'commissions', %s, %s, %s, NOW())
    """, (orig_id, result_by, f"request #{req_id}: {reason}"))
    return so_number, sale_key


# ── เปรียบเทียบก่อน/หลัง และอนุมัติผล (ด่าน 2) ─────────────────────────────

# คอลัมน์ที่ระบบอื่นเติมไว้ในแถวเดิม (ฟอร์ม SO ไม่ได้เก็บ) — ต้องคัดลอกไปแถวใหม่ตอนอนุมัติ
# ไม่งั้นจัดซื้อที่รับงานอยู่จะหายจากงานตัวเอง ฯลฯ  (ใส่เฉพาะเมื่อแถวใหม่ยังว่าง)
CARRY_OVER_IF_EMPTY = (
    'user_key', 'claim_timestamp', 'support_user_key',
    'approver_sale_manager_key', 'approval_date_sale_manager',
    'hr_cost_overrides', 'hr_sale_source', 'hr_cost_source', 'is_collection_risk',
    'defer_type', 'defer_decision', 'defer_decision_reason', 'defer_requested_by', 'defer_approved_by',
    'expected_delivery_date', 'expected_payment_date', 'defer_remarks',
    'defer_source_month', 'defer_source_year',
    'commission_now_amount', 'commission_reserve_amount', 'reserve_status',
    'reserve_payout_id', 'reserve_decided_at', 'reserve_decided_by',
)
# ฟอร์มใส่ค่าเริ่มต้นให้เสมอ (ไม่ว่าง) ต้องบังคับใช้ค่าจากแถวเดิม
CARRY_OVER_ALWAYS = ('cost_multiplier', 'sm_reject_count')
# ไม่คัดลอก: final_* คือค่าสุทธิที่ HR คำนวณตอนตรวจ ยอดเปลี่ยนแล้วค่าเดิมจะผิด ให้ HR ตรวจใหม่

# คอลัมน์ที่ไม่นำมาเทียบก่อน/หลัง (ระบบจัดการเอง ไม่ใช่สิ่งที่เซลส์แก้)
DIFF_IGNORE = set(CARRY_OVER_IF_EMPTY) | set(CARRY_OVER_ALWAYS) | {
    'id', 'timestamp', 'status', 'is_active', 'original_id', 'sale_key', 'rejection_reason', 'payout_id',
    'final_sales_amount', 'final_cost_amount', 'final_gp', 'final_margin', 'final_commission', 'profit',
}

FIELD_LABELS = {
    'bill_date': 'วันที่เปิดบิล', 'customer_id': 'รหัสลูกค้า', 'customer_name': 'ชื่อลูกค้า',
    'customer_type': 'ประเภทลูกค้า', 'credit_term': 'เครดิตเทอม', 'so_number': 'เลข SO',
    'sales_service_amount': 'ยอดขายสินค้า/บริการ', 'sales_service_vat_option': 'ยอดขาย VAT/เงินสด',
    'shipping_cost': 'ค่าจัดส่ง', 'shipping_vat_option': 'ค่าจัดส่ง VAT/เงินสด',
    'cutting_drilling_fee': 'ค่าตัด/เจาะ', 'cutting_drilling_fee_vat_option': 'ค่าตัด/เจาะ VAT/เงินสด',
    'other_service_fee': 'บริการอื่นๆ', 'other_service_fee_vat_option': 'บริการอื่นๆ VAT/เงินสด',
    'relocation_cost': 'ค่าย้าย', 'relocation_cost_vat_option': 'ค่าย้าย VAT/เงินสด',
    'credit_card_fee': 'ค่าธรรมเนียมบัตร', 'credit_card_fee_vat_option': 'ค่าธรรมเนียมบัตร VAT/เงินสด',
    'transfer_fee': 'ค่าธรรมเนียมโอน', 'coupons': 'คูปอง', 'giveaways': 'ของแถม', 'wht_3_percent': 'หัก ณ ที่จ่าย 3%',
    'payment1_amount': 'ยอดชำระงวด 1', 'payment1_date': 'วันที่ชำระงวด 1', 'payment1_method': 'วิธีชำระงวด 1',
    'payment2_amount': 'ยอดชำระงวด 2', 'payment2_date': 'วันที่ชำระงวด 2', 'payment2_method': 'วิธีชำระงวด 2',
    'total_payment_amount': 'ยอดชำระรวม', 'difference_amount': 'ส่วนต่างยอดชำระ', 'payment_date': 'วันที่ชำระ',
    'balance_due': 'ยอดค้างชำระ', 'cash_required_total': 'ยอดเงินสดที่ต้องรับ',
    'commission_month': 'เดือนค่าคอม', 'commission_year': 'ปีค่าคอม',
    'delivery_date': 'วันที่จัดส่ง', 'delivery_type': 'รูปแบบจัดส่ง', 'pickup_location': 'จุดรับสินค้า',
    'vehicle_type': 'ประเภทรถ', 'delivery_map': 'แผนที่จัดส่ง', 'onsite_contact_name': 'ผู้ติดต่อหน้างาน',
    'onsite_contact_phone': 'เบอร์ผู้ติดต่อ', 'special_request': 'คำขอพิเศษ', 'unloading_status': 'การลงสินค้า',
    'order_pur': 'ประเภทสั่งซื้อ', 'date_to_warehouse': 'วันที่เข้าคลัง', 'date_to_customer': 'วันที่ส่งลูกค้า',
    'subject': 'หัวข้อ',
}


def _norm(v):
    """ทำให้ค่าเทียบกันได้: ว่าง/None/NaN เท่ากัน ตัวเลขเทียบแบบปัดเศษ วันที่เป็นข้อความ"""
    try:
        import pandas as pd
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return None
    except Exception:
        if v is None:
            return None
    if isinstance(v, str):
        v = v.strip()
        return v or None
    if hasattr(v, 'isoformat'):
        return v.isoformat(timespec='seconds') if hasattr(v, 'hour') else v.isoformat()
    try:
        return round(float(v), 2)
    except Exception:
        return str(v)


def compute_so_diff(old, new):
    """คืนรายการ (ชื่อช่อง, ค่าก่อน, ค่าหลัง) เฉพาะช่องที่เซลส์แก้ (old/new เป็น dict ของแถว commissions)"""
    out = []
    for col in new.keys():
        if col in DIFF_IGNORE or col not in old:
            continue
        a, b = _norm(old[col]), _norm(new[col])
        if a == b:
            continue
        out.append((FIELD_LABELS.get(col, col), '' if a is None else a, '' if b is None else b))
    return out


def complete_change_request(cursor, req_id, approved_by):
    """SM อนุมัติผลการแก้ไข: แถวใหม่กลายเป็น SO จริง คัดลอกข้อมูลที่ระบบเติมไว้จากแถวเดิม
    แล้วกลับสถานะเดิม (HR Verified กลับไปรอ HR ตรวจใหม่ที่ PO Sent) ผู้เรียกเป็นคน commit
    คืนค่า (so_number, สถานะที่กลับไป) หรือ None ถ้าคำขอนี้จบไปแล้ว"""
    cursor.execute("""
        SELECT so_number, original_commission_id, new_commission_id, sale_key, original_status, status
        FROM so_change_requests WHERE id = %s FOR UPDATE
    """, (req_id,))
    req = cursor.fetchone()
    if not req or req[5] != 'Submitted' or req[2] is None:
        return None
    so_number, orig_id, new_id, sale_key, original_status = req[0], req[1], req[2], req[3], req[4]

    cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'commissions'")
    existing_cols = {r[0] for r in cursor.fetchall()}

    sets = []
    for col in CARRY_OVER_ALWAYS:
        if col in existing_cols:
            sets.append(f'"{col}" = o."{col}"')
    for col in CARRY_OVER_IF_EMPTY:
        if col in existing_cols:
            sets.append(f'"{col}" = COALESCE(n."{col}", o."{col}")')
    restore_status = 'PO Sent' if original_status == 'HR Verified' else original_status
    cursor.execute(
        f"""UPDATE commissions n SET {', '.join(sets)}, status = %s, is_active = 1
            FROM commissions o WHERE n.id = %s AND o.id = %s""",
        (restore_status, new_id, orig_id))
    cursor.execute("UPDATE commissions SET is_active = 0 WHERE id = %s", (orig_id,))
    cursor.execute("""
        UPDATE so_change_requests SET status = 'Completed', result_by = %s, result_at = NOW() WHERE id = %s
    """, (approved_by, req_id))
    note = "ต้องให้ HR ตรวจใหม่" if original_status == 'HR Verified' else "ทำงานต่อจากจุดเดิม"
    cursor.execute("""
        INSERT INTO notifications (user_key_to_notify, message, is_read, related_so_id)
        VALUES (%s, %s, FALSE, %s)
    """, (sale_key, f"SM อนุมัติผลการแก้ไข SO {so_number} แล้ว ({note})", new_id))
    cursor.execute("""
        INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
        VALUES ('SO Change Completed', 'commissions', %s, %s, %s, NOW())
    """, (new_id, approved_by, f"request #{req_id}: กลับสถานะ {restore_status}"))
    return so_number, restore_status


def expire_overdue_requests(conn):
    """ย้อน SO กลับเป็นค่าเดิมสำหรับคำขอที่ SM อนุมัติสิทธิ์แล้วแต่เซลส์ไม่แก้/ไม่ส่งภายในกำหนด (expires_at)
    commit ทีละคำขอ เพื่อให้คำขอหนึ่งพังแล้วไม่ลากตัวอื่นไปด้วย คืนจำนวนที่ย้อนสำเร็จ
    เรียกจากหน้าเซลส์/SM ตามรอบ polling (แอปนี้ไม่มีตัวรันงานเบื้องหลังฝั่งเซิร์ฟเวอร์)"""
    done = 0
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM so_change_requests WHERE status = 'Approved' AND expires_at < NOW() ORDER BY id")
            ids = [r[0] for r in cur.fetchall()]
        conn.rollback()
    except Exception:
        conn.rollback()
        return 0
    for rid in ids:
        try:
            with conn.cursor() as cur:
                res = revert_change_request(cur, rid, 'system', 'สิทธิ์แก้ไขหมดอายุ (เกิน 7 วัน)')
            conn.commit()
            if res:
                done += 1
        except Exception as e:
            conn.rollback()
            print(f"expire_overdue_requests #{rid} error: {e}")
    return done
