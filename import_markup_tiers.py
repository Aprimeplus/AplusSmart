# import_markup_tiers.py
# อ่าน Markup Guide.xlsx แล้ว import ข้อมูลเข้า DB
# รัน: python import_markup_tiers.py
# ถ้าต้องการ re-import ใหม่ทั้งหมด: python import_markup_tiers.py --reset

import sys
import pandas as pd
import psycopg2
import psycopg2.extras
from datetime import datetime

_DB_CFG = dict(host="Server-APrime", dbname="aplus_com_test",
               user="app_user", password="cailfornia123")

EXCEL_PATH = r"C:\Users\Nitro V15\Downloads\Markup Guide.xlsx"

# mapping: (col_markup_pct, col_amount_range, col_qty_range)
TIER_COLS = {
    "Guest": (8,  10, 11),
    "T1":    (13, 15, 16),
    "T2":    (18, 20, 21),
    "T3":    (23, 25, 26),
    "T4":    (28, 30, 31),
    "T5":    (33, 35, 36),
}
TIER_ORDER = {"Guest": 0, "T1": 1, "T2": 2, "T3": 3, "T4": 4, "T5": 5}


def _safe_float(val):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        return float(val)
    except Exception:
        return None


def _safe_str(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    s = str(val).strip()
    return s if s and s != "NaT" and s != "nan" else None


def _safe_date(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    try:
        if isinstance(val, datetime):
            if val.year < 1900 or val.year > 2100:
                return None
            return val.date()
        if isinstance(val, str):
            val = val.strip()
            if not val:
                return None
        ts = pd.to_datetime(val)
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def parse_excel(path: str) -> list[dict]:
    df = pd.read_excel(path, header=None, sheet_name="Sheet1")
    records = []

    last_main_cat = ""
    last_sub_cat = ""

    for row_idx in range(2, len(df)):  # เริ่ม row 2 (skip header rows 0,1)
        row = df.iloc[row_idx]

        # ดึงค่าพื้นฐาน
        main_cat   = _safe_str(row.iloc[0]) or last_main_cat
        status_note = _safe_str(row.iloc[1])
        sub_cat    = _safe_str(row.iloc[2]) or last_sub_cat
        prod_type  = _safe_str(row.iloc[3])
        champ      = _safe_str(row.iloc[4])
        upd_date   = _safe_date(row.iloc[5])
        cost_kg    = _safe_float(row.iloc[6])

        # ข้ามแถวที่ไม่มีชื่อ product_type
        if not prod_type:
            # อัปเดต category carry-forward ถ้ามี
            if _safe_str(row.iloc[0]):
                last_main_cat = _safe_str(row.iloc[0])
            if _safe_str(row.iloc[2]):
                last_sub_cat = _safe_str(row.iloc[2])
            continue

        # อัปเดต carry-forward
        last_main_cat = main_cat
        last_sub_cat  = sub_cat

        # วน Tier
        for tier_name, (col_pct, col_amt, col_qty) in TIER_COLS.items():
            markup_pct  = _safe_float(row.iloc[col_pct]) if col_pct < len(row) else None
            amount_range = _safe_str(row.iloc[col_amt]) if col_amt < len(row) else None
            qty_range    = _safe_str(row.iloc[col_qty]) if col_qty < len(row) else None

            # ข้าม tier ที่ไม่มี markup_pct เลย
            if markup_pct is None:
                continue

            records.append({
                "main_category":      main_cat,
                "sub_category":       sub_cat,
                "product_type":       prod_type,
                "product_champ":      champ,
                "cost_per_kg":        cost_kg,
                "tier_name":          tier_name,
                "tier_order":         TIER_ORDER[tier_name],
                "markup_pct":         markup_pct,
                "amount_range":       amount_range,
                "qty_range":          qty_range,
                "status_note":        status_note,
                "excel_updated_date": upd_date,
                "excel_updated_by":   None,
            })

    return records


def import_to_db(records: list[dict], reset: bool = False):
    conn = psycopg2.connect(**_DB_CFG)
    conn.autocommit = False
    cur = conn.cursor()

    if reset:
        print("  [RESET] Deleting existing rows...")
        cur.execute("DELETE FROM markup_tiers")
        print(f"  Deleted. Re-inserting {len(records)} records...")

    INSERT_SQL = """
        INSERT INTO markup_tiers
            (main_category, sub_category, product_type, product_champ,
             cost_per_kg, tier_name, tier_order, markup_pct,
             amount_range, qty_range, status_note,
             excel_updated_date, excel_updated_by)
        VALUES
            (%(main_category)s, %(sub_category)s, %(product_type)s, %(product_champ)s,
             %(cost_per_kg)s, %(tier_name)s, %(tier_order)s, %(markup_pct)s,
             %(amount_range)s, %(qty_range)s, %(status_note)s,
             %(excel_updated_date)s, %(excel_updated_by)s)
        ON CONFLICT DO NOTHING
    """
    psycopg2.extras.execute_batch(cur, INSERT_SQL, records, page_size=200)
    conn.commit()
    print(f"  OK: Inserted {len(records)} records into markup_tiers.")
    cur.close()
    conn.close()


if __name__ == "__main__":
    reset_flag = "--reset" in sys.argv
    print(f"Reading Excel: {EXCEL_PATH}")
    recs = parse_excel(EXCEL_PATH)
    print(f"  Parsed {len(recs)} tier records from {len(set(r['product_type'] for r in recs))} product types.")
    print("Connecting to DB...")
    import_to_db(recs, reset=reset_flag)
    print("Done!")
