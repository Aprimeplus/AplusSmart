"""กู้ Order No. ของตาราง cost_benchmarks จากไฟล์ Excel/CSV ที่ export ไว้ "ก่อน" ข้อมูลเสีย
(เช่น Export Excel จากหน้า Dashboard/Cost Benchmark ที่เปิดค้างไว้โดยยังไม่รีเฟรช)

วิธีใช้:
    python restore_order_no_from_export.py ไฟล์.xlsx --user KOY --bad "1022-07 K Rev2"          # ดูผลก่อน (dry-run ไม่แก้ DB)
    python restore_order_no_from_export.py ไฟล์.xlsx --user KOY --bad "1022-07 K Rev2" --apply  # แก้จริง

จับคู่ด้วย (Sale Order No., รายการสินค้า, ชื่อ Supplier, จำนวน) — แก้เฉพาะแถวที่ค่า Order No. ปัจจุบัน = ค่าที่เสีย (--bad)
และแถวนั้นจับคู่ได้ "แถวเดียว" เท่านั้น (กำกวม/ไม่พบ = ไม่แตะ แล้วรายงานให้ตรวจเอง) ก่อนแก้จะสำรองแถวที่เสียไว้ในตารางสำรอง
"""
import argparse
import re
import sys

import pandas as pd

COLS = {
    "so": ["Sale Order No.", "Sale Order No", "sale_order_no"],
    "item": ["รายการสินค้า", "รายการ", "item"],
    "sup": ["ชื่อ Supplier", "Supplier", "supplier"],
    "qty": ["จำนวน", "qty"],
    "ord": ["Order No.", "Order No", "order_no"],
}


def _pick(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None


def _norm(v):
    s = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
    s = re.sub(r"\s+", " ", s).strip()
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--user", default="KOY")
    ap.add_argument("--bad", required=True, help="ค่า Order No. ที่เสีย เช่น '1022-07 K Rev2'")
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    src = pd.read_csv(a.file, dtype=str, encoding="utf-8-sig") if a.file.lower().endswith(".csv") else pd.read_excel(a.file, dtype=str)
    c = {k: _pick(src, v) for k, v in COLS.items()}
    missing = [k for k, v in c.items() if v is None]
    if missing:
        print("ไฟล์ไม่มีคอลัมน์ที่ต้องใช้:", missing, "| คอลัมน์ในไฟล์:", list(src.columns)[:25])
        return
    src["_k"] = list(zip(src[c["so"]].map(_norm), src[c["item"]].map(_norm), src[c["sup"]].map(_norm), src[c["qty"]].map(_norm)))
    src["_o"] = src[c["ord"]].map(_norm)
    src = src[(src["_o"] != "") & (src["_o"] != a.bad)]
    by_key = src.groupby("_k")["_o"].agg(lambda s: sorted(set(s)))

    from main_app import AppContainer
    app = AppContainer()
    conn = app.get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute('''SELECT id, "Sale Order No.", "รายการสินค้า", "ชื่อ Supplier", "จำนวน"
                           FROM cost_benchmarks WHERE created_by = %s AND "Order No." = %s ORDER BY id''', (a.user, a.bad))
            bad_rows = cur.fetchall()
        updates, ambiguous, notfound = [], [], []
        for rid, so, item, sup, qty in bad_rows:
            key = (_norm(so), _norm(item), _norm(sup), _norm(qty))
            cands = by_key.get(key)
            if cands is None:
                notfound.append(rid)
            elif len(cands) == 1:
                updates.append((cands[0], rid))
            else:
                ambiguous.append((rid, cands))
        print(f"แถวที่เสียใน DB: {len(bad_rows)} | จับคู่ได้ชัดเจน: {len(updates)} | กำกวม: {len(ambiguous)} | ไม่พบในไฟล์: {len(notfound)}")
        by_order = pd.Series([u[0] for u in updates]).value_counts()
        print("เลข Order ที่จะใส่กลับ (จำนวนแถว):", dict(by_order.head(15)), "..." if len(by_order) > 15 else "")
        if not a.apply:
            print("\n(dry-run — ยังไม่แก้ DB ใส่ --apply เมื่อพร้อม)")
            return
        if not updates:
            print("ไม่มีแถวให้แก้")
            return
        tbl = f"cost_benchmarks_backup_{a.user.lower()}_orderno_fix"
        with conn.cursor() as cur:
            cur.execute(f'CREATE TABLE IF NOT EXISTS {tbl} AS SELECT * FROM cost_benchmarks WHERE false')
            cur.execute(f'INSERT INTO {tbl} SELECT * FROM cost_benchmarks WHERE id = ANY(%s)', ([u[1] for u in updates],))
            import psycopg2.extras
            psycopg2.extras.execute_batch(cur, 'UPDATE cost_benchmarks SET "Order No." = %s WHERE id = %s AND "Order No." = %s',
                                          [(o, rid, a.bad) for o, rid in updates])
        conn.commit()
        print(f"แก้แล้ว {len(updates)} แถว (สำรองแถวที่เสียไว้ที่ตาราง {tbl})")
    finally:
        app.release_connection(conn)


if __name__ == "__main__":
    main()
