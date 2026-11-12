"""เทสต์กติกาของ Daily Report ใหม่ (ฝ่ายขาย/ฝ่ายจัดซื้อ) — ฟังก์ชันล้วน ไม่ต้องเปิดหน้าจอหรือเชื่อมฐานข้อมูล"""
from datetime import date

import pandas as pd

import daily_report_sales as drs
from daily_report_purchasing import _d_label, _severity, _short_so, PurchasingDailyBoard
from canvas_board import clean_text


def _row(**kw):
    base = dict(n=5, status="PO Sent", required=1000.0, paid=1000.0, dq_diff=None, pu="AP093", sale_key="ILADA")
    base.update(kw)
    return pd.Series(base)


# ------------------------------------------------------------------ รอบเดือน (MTD)
def test_cycle_start_26_to_25():
    assert drs._cycle_start(date(2026, 11, 12)) == date(2026, 10, 26)   # ก่อนวันที่ 26 → เริ่ม 26 เดือนก่อน
    assert drs._cycle_start(date(2026, 11, 26)) == date(2026, 11, 26)   # วันที่ 26 เป็นวันแรกของรอบใหม่
    assert drs._cycle_start(date(2026, 1, 5)) == date(2025, 12, 26)     # ข้ามปี


def test_cycle_start_calendar_month(monkeypatch):
    monkeypatch.setattr(drs, "CYCLE_START_DAY", 1)                       # ปีหน้าเดือนต่อเดือน
    assert drs._cycle_start(date(2027, 3, 18)) == date(2027, 3, 1)


# ------------------------------------------------------------------ ความเสี่ยง
def test_overdue_open_job_is_urgent_dl03():
    rs = drs.risk_reasons(_row(status="PO In Progress", n=-5))
    assert rs[0][0] == "DL-03" and "5 วัน" in rs[0][1]


def test_delivery_today_without_po_dl04():
    rs = dict(drs.risk_reasons(_row(status="Pending PU", n=0, pu="AP093")))
    assert "DL-04" in rs and "วันนี้" in rs["DL-04"]
    rs = dict(drs.risk_reasons(_row(status="Pending PU", n=1, pu="AP093")))
    assert "พรุ่งนี้" in rs["DL-04"]


def test_po_sent_overdue_is_not_flagged():
    assert drs.risk_reasons(_row(status="PO Sent", n=-3)) == []         # เปิด PO แล้ว ไม่ใช่งานค้างจัดซื้อ


def test_unpaid_but_purchasing_started_py01():
    rs = dict(drs.risk_reasons(_row(status="PO In Progress", n=5, paid=0.0, required=500.0)))
    assert "PY-01" in rs
    rs = dict(drs.risk_reasons(_row(status="Pending PU", n=5, paid=0.0, required=500.0, pu="AP093")))
    assert "PY-01" not in rs                                            # ยังไม่เริ่มจัดซื้อ ไม่ใช่ PY-01


def test_no_pu_only_when_pending_dl05():
    assert "DL-05" in dict(drs.risk_reasons(_row(status="Pending PU", n=9, pu=None)))
    assert "DL-05" not in dict(drs.risk_reasons(_row(status="PO In Progress", n=9, pu=None)))


def test_reason_order_and_main_reason():
    rs = drs.risk_reasons(_row(status="Pending PU", n=-2, pu=None, dq_diff=120.0, required=500, paid=0.0))
    codes = [c for c, _ in rs]
    assert codes[0] == "DL-03"                                           # เหตุหลักตามลำดับความสำคัญ
    assert codes == [c for c in drs.RISK_ORDER if c in codes]


def test_dq_diff_nan_is_ignored():
    assert "DQ-01" not in dict(drs.risk_reasons(_row(dq_diff=float("nan"))))


# ------------------------------------------------------------------ ยอดต่าง (สูตรรวมค่าขนส่ง)
def _dq(goods, service, ship, required):
    """ทำซ้ำสูตรใน load_data: คืนส่วนต่างถ้าไม่ตรง ×1.07 และไม่ตรง ×1.00 มากกว่า TOL"""
    base = goods + service + ship
    ok = abs(required - base * drs.VAT) <= drs.TOL or abs(required - base) <= drs.TOL or required <= drs.TOL
    return None if ok else required - base * drs.VAT


def test_shipping_included_no_false_flag():
    # AM036 จริง: ขาย 6,988 + ขนส่ง 2,000 → รวม 9,617.16 (ต้นแบบ PM ติดป้าย "ต่าง 2,140" ผิดเพราะลืมค่าขนส่ง)
    assert _dq(6988.0, 0.0, 2000.0, 9617.16) is None


def test_cash_without_vat_ok_and_real_mismatch_flagged():
    assert _dq(1000.0, 0.0, 0.0, 1000.0) is None                         # ไม่คิด VAT
    assert _dq(1000.0, 0.0, 0.0, 1200.0) is not None                     # ยอดผิดจริงต้องติดป้าย


# ------------------------------------------------------------------ กระดานจัดซื้อ
def test_hint_logic():
    h = PurchasingDailyBoard._hint
    base = dict(col="work", n=3, required=1000.0, total_payment_amount=0.0, pct=None)
    assert h(pd.Series({**base, "n": -1})) == "เลยกำหนด"
    assert h(pd.Series({**base, "n": 0})) == "ส่งวันนี้"
    assert h(pd.Series({**base, "n": 0, "col": "done"})) == "จัดรถวันนี้"
    assert h(pd.Series({**base, "n": 2, "pct": 0.0})) == "ยังไม่ชำระ"
    assert h(pd.Series({**base, "n": 9, "pct": 0.5, "total_payment_amount": 500.0})) == "มัดจำ 50%"
    assert h(pd.Series({**base, "n": 9, "pct": 1.0, "total_payment_amount": 1000.0})) == ""


def test_labels_and_helpers():
    assert _d_label(0) == "D-0" and _d_label(5) == "D-5" and _d_label(-3) == "D+3"
    assert _severity(0) == "red" and _severity(-2) == "red" and _severity(2) == "orange" and _severity(7) == "gray"
    assert _short_so("SO6911AM019") == "AM019" and _short_so("X") == "X"


# ------------------------------------------------------------------ ตัวอักษรเสีย
def test_clean_text_removes_private_use_but_keeps_thai():
    assert clean_text("เอ็นจิเนียริ\U001000c9ง") == "เอ็นจิเนียริง"
    assert clean_text("เอ็นจิเนียริ่ง") == "เอ็นจิเนียริ่ง"
    assert clean_text(None) is None
