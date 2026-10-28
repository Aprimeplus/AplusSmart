"""QT (ใบเสนอราคา) จาก Express → นำเข้า + จับคู่กับ SO + คำนวณ Conversion ตามระดับบิล (spec DEV-SPEC-SALEDASH-001 หัวข้อ 16)

ไฟล์ต้นทาง: รายงาน "ใบเสนอราคา แยกตามพนักงานขาย" (CSV, Windows-874) ไม่ต้องแปลงไฟล์ก่อน นำเข้าไฟล์ดิบได้เลย
เก็บเฉพาะข้อมูลที่ใช้วิเคราะห์ (เลข QT, วันที่, รหัสลูกค้า, ที่นั่ง, มูลค่า) — ไม่เก็บชื่อลูกค้า/ผู้ติดต่อ/เบอร์โทร
"""
import csv
import io
import re
from datetime import date

import pandas as pd

TIERS = [("Guest", 0.0, 10_000.0), ("T1", 10_000.0, 30_000.0), ("T2", 30_000.0, 100_000.0),
         ("T3", 100_000.0, 300_000.0), ("T4", 300_000.0, 1_000_000.0), ("T5", 1_000_000.0, float("inf"))]
NON_COMMISSION_SEATS = ("CT001",)          # Sale Center (CT) ไม่อยู่ในเกณฑ์ KPI

_KEY_RE = re.compile(r"^(?:QT|SO)?(\d{4}[A-Z]{2}\d{3})")


def doc_key(number):
    """ตัด QT/SO และตัวอักษรท้าย/ล็อต ให้เหลือเลขหลักที่ใช้จับคู่: QT6908KR049A / SO6908KR049-L1 → 6908KR049"""
    m = _KEY_RE.match(str(number or "").strip().upper())
    return m.group(1) if m else None


def tier_index(value):
    for i, (_n, lo, hi) in enumerate(TIERS):
        if lo <= value < hi:
            return i
    return len(TIERS) - 1


def _f(v):
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _g(r, i):
    return r[i].strip() if len(r) > i else ""


def parse_express_qt_csv(path):
    """อ่านไฟล์รายงาน QT ของ Express (Windows-874) → (DataFrame, meta)
    คอลัมน์: qt_number, qt_date, customer_code, seat_code, seat_name, vat_type, discount, value_ex_vat, vat, total, sent_flag, status"""
    raw = open(path, "rb").read()
    try:
        text = raw.decode("cp874")
    except UnicodeDecodeError:
        text = raw.decode("utf-8-sig")        # เผื่อเป็นไฟล์ที่แปลงเป็น UTF-8 แล้ว
    rows = list(csv.reader(io.StringIO(text)))
    seat_code = seat_name = None
    out = []
    for r in rows:
        if _g(r, 1) and not _g(r, 4) and re.match(r"^[A-Z][A-Z0-9-]*\d{3}$", _g(r, 1)) and _g(r, 2):
            seat_code = _g(r, 1)
            seat_name = re.sub(r"[\d\s\-]+$", "", _g(r, 2)).strip()    # ตัดเบอร์โทร (รวมขีด) ท้ายชื่อ
            continue
        num = _g(r, 4)
        if not re.match(r"^QT\w+", num):
            continue
        d = None
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", _g(r, 5))
        if m:
            dd, mm, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                d = date(yy - 543 if yy > 2400 else yy, mm, dd)
            except ValueError:
                d = None
        out.append({"qt_number": num, "qt_date": d, "customer_code": _g(r, 6), "seat_code": seat_code,
                    "seat_name": seat_name, "vat_type": _g(r, 9), "discount": _f(_g(r, 10)),
                    "value_ex_vat": _f(_g(r, 11)), "vat": _f(_g(r, 12)), "total": _f(_g(r, 13)),
                    "sent_flag": _g(r, 14), "status": _g(r, 15)})
    df = pd.DataFrame(out)
    meta = {"rows": len(df)}
    if not df.empty:
        meta["date_from"] = df["qt_date"].dropna().min()
        meta["date_to"] = df["qt_date"].dropna().max()
        meta["dup_numbers"] = int(df["qt_number"].duplicated().sum())
    return df, meta


def coverage_warnings(df):
    """ตรวจความครบของเลข QT รูปแบบ QT<ปีเดือน><ที่นั่ง><ลำดับ> (เช่น QT6901AM001): เลขลำดับที่หายไป = ไฟล์อาจไม่ครบ"""
    warn = []
    pat = re.compile(r"^QT(\d{4})([A-Z]{2})(\d{3})")
    seen = {}
    for n in df["qt_number"]:
        m = pat.match(n)
        if m:
            seen.setdefault((m.group(1), m.group(2)), set()).add(int(m.group(3)))
    for (ym, seat), seqs in sorted(seen.items()):
        mx = max(seqs)
        missing = mx - len(seqs)
        if mx >= 20 and missing / mx > 0.10:
            warn.append(f"QT{ym}{seat}: มีเลขถึง {mx} แต่พบ {len(seqs)} ใบ (หายไป {missing})")
    return warn


MIGRATE_SQL = [
    """
    CREATE TABLE IF NOT EXISTS qt_header (
        qt_number     TEXT PRIMARY KEY,
        qt_date       DATE,
        customer_code TEXT,
        seat_code     TEXT,
        seat_name     TEXT,
        vat_type      TEXT,
        discount      NUMERIC(14,2),
        value_ex_vat  NUMERIC(14,2),
        vat           NUMERIC(14,2),
        total         NUMERIC(14,2),
        sent_flag     TEXT,
        status        TEXT,
        import_id     INTEGER,
        imported_at   TIMESTAMP DEFAULT NOW()
    )""",
    """
    CREATE TABLE IF NOT EXISTS qt_import_log (
        id           SERIAL PRIMARY KEY,
        filename     TEXT,
        imported_by  TEXT,
        imported_at  TIMESTAMP DEFAULT NOW(),
        rows_total   INTEGER,
        rows_new     INTEGER,
        rows_updated INTEGER,
        date_from    DATE,
        date_to      DATE,
        warnings     TEXT
    )""",
    "CREATE INDEX IF NOT EXISTS idx_qt_date ON qt_header(qt_date)",
    "CREATE INDEX IF NOT EXISTS idx_qt_seat ON qt_header(seat_code)",
]


def ensure_tables(conn):
    with conn.cursor() as cur:
        for sql in MIGRATE_SQL:
            cur.execute(sql)
    conn.commit()


def import_to_db(conn, df, filename, user, warnings=None):
    """upsert ตามเลข QT (นำเข้าไฟล์ซ้ำ/ไฟล์ช่วงทับกันได้ ไม่เกิดข้อมูลซ้ำ) คืน (new, updated, import_id)"""
    import psycopg2.extras
    ensure_tables(conn)
    with conn.cursor() as cur:
        cur.execute("SELECT qt_number FROM qt_header WHERE qt_number = ANY(%s)", (df["qt_number"].tolist(),))
        existing = {r[0] for r in cur.fetchall()}
        cur.execute(
            "INSERT INTO qt_import_log (filename, imported_by, rows_total, date_from, date_to, warnings) "
            "VALUES (%s, %s, %s, %s, %s, %s) RETURNING id",
            (filename, user, len(df), df["qt_date"].dropna().min(), df["qt_date"].dropna().max(),
             "; ".join(warnings or []) or None))
        imp_id = cur.fetchone()[0]
        rows = [(r.qt_number, r.qt_date, r.customer_code, r.seat_code, r.seat_name, r.vat_type, r.discount,
                 r.value_ex_vat, r.vat, r.total, r.sent_flag, r.status, imp_id) for r in df.itertuples()]
        psycopg2.extras.execute_values(cur, """
            INSERT INTO qt_header (qt_number, qt_date, customer_code, seat_code, seat_name, vat_type, discount,
                                   value_ex_vat, vat, total, sent_flag, status, import_id)
            VALUES %s
            ON CONFLICT (qt_number) DO UPDATE SET
                qt_date = EXCLUDED.qt_date, customer_code = EXCLUDED.customer_code, seat_code = EXCLUDED.seat_code,
                seat_name = EXCLUDED.seat_name, vat_type = EXCLUDED.vat_type, discount = EXCLUDED.discount,
                value_ex_vat = EXCLUDED.value_ex_vat, vat = EXCLUDED.vat, total = EXCLUDED.total,
                sent_flag = EXCLUDED.sent_flag, status = EXCLUDED.status, import_id = EXCLUDED.import_id,
                imported_at = NOW()
        """, rows)
        new = len(set(df["qt_number"]) - existing)
        upd = len(set(df["qt_number"]) & existing)
        cur.execute("UPDATE qt_import_log SET rows_new = %s, rows_updated = %s WHERE id = %s", (new, upd, imp_id))
    conn.commit()
    return new, upd, imp_id


def compute_conversion(qt_df, so_df):
    """Conversion ตามระดับบิล
    qt_df: qt_number, qt_date, customer_code, seat_code, value_ex_vat
    so_df: so_number, amt (ยอดสุทธิ SO), deal (รหัสดีล: งาน Project รวมล็อต = project, อื่นๆ = เลข SO)
    กติกา (spec 16.2):
      • QT มูลค่า 0 ไม่นับ · QT ฉบับแก้ไข (เลขเดียวกันต่อท้ายตัวอักษร) นับเป็นดีลเดียว (ใช้มูลค่าสูงสุด)
      • จับคู่ QT↔SO ด้วยเลขเอกสาร (ตัด QT/SO และตัวอักษรท้าย/ล็อต) · QT รวมจัดระดับตามมูลค่า QT, ปิดได้/ยอด SO จัดระดับตามมูลค่า SO
      • SO ที่ไม่พบ QT ในไฟล์ นับเป็น QT ที่ปิดได้ (ใช้ยอด SO) · ยอด SO ของ QT ที่ปิดได้ = Σ SO ที่จับคู่ (รวมทุกล็อต)
    คืน list 6 แถว: dict(qt, closed, qt_value, so_value) ตามระดับ"""
    res = [{"qt": 0, "closed": 0, "qt_value": 0.0, "so_value": 0.0} for _ in TIERS]
    # ยอด SO ต่อคีย์เอกสาร (รวมล็อต/เวอร์ชัน)
    so_df = so_df.copy()
    so_df["_k"] = so_df["so_number"].map(doc_key)
    so_by_key = so_df.dropna(subset=["_k"]).groupby("_k")["amt"].sum().to_dict()
    so_deal_amt = so_df.groupby("deal")["amt"].sum().to_dict()
    so_deal_keys = so_df.dropna(subset=["_k"]).groupby("deal")["_k"].agg(set).to_dict()

    q = qt_df[(qt_df["value_ex_vat"] > 0) & (~qt_df["seat_code"].isin(NON_COMMISSION_SEATS))].copy()
    q["_k"] = q["qt_number"].map(doc_key)
    # QT ฉบับแก้ไข = เลขเดียวกันต่อท้ายด้วยตัวอักษร (QT6901ID070 / QT6901ID070A) → นับเป็นดีลเดียว ใช้มูลค่าสูงสุด
    # เลขที่อ่านรูปแบบไม่ได้ (รูปแบบเก่า) นับเป็นรายใบ
    q["_gk"] = q["_k"].fillna(q["qt_number"])
    deals = q.groupby("_gk").agg(value=("value_ex_vat", "max"), keys=("_k", lambda s: {x for x in s if x})).reset_index()
    matched_deals = set()
    for d in deals.itertuples():
        # QT รวม / มูลค่า QT: จัดระดับตามมูลค่า QT
        ti = tier_index(d.value)
        res[ti]["qt"] += 1
        res[ti]["qt_value"] += float(d.value)
        for k in d.keys:
            if k in so_by_key:
                for dl, ks in so_deal_keys.items():
                    if k in ks:
                        matched_deals.add(dl)
    # ปิดได้ / ยอด SO: จัดระดับตามมูลค่าของ SO (ดีล) ที่จับคู่ได้ — ตามรูปแบบตาราง Conversion ของ PM
    for dl in matched_deals:
        amt = so_deal_amt.get(dl, 0.0)
        if amt <= 0:
            continue
        ti = tier_index(amt)
        res[ti]["closed"] += 1
        res[ti]["so_value"] += float(amt)
    # SO ที่ไม่พบ QT ในไฟล์ → นับเป็น QT ที่ปิดได้ (QT ตกหล่นจากรายงาน) ใช้ยอด SO
    for dl, amt in so_deal_amt.items():
        if dl in matched_deals or amt <= 0:
            continue
        ti = tier_index(amt)
        res[ti]["qt"] += 1
        res[ti]["closed"] += 1
        res[ti]["qt_value"] += float(amt)
        res[ti]["so_value"] += float(amt)
    return res
