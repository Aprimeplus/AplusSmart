# super_supplier_list.py
# Aplus Smart — Super Supplier List Tab (Phase 1)
# ─────────────────────────────────────────────────────────────────────────────
# วิธีใช้งาน (ใน PurchasingManagerScreen.__init__):
#
#   from super_supplier_list import SuperSupplierTab
#
#   self.ssl_tab = self.tab_view.add("Super Supplier List")
#   self.ssl_tab.grid_columnconfigure(0, weight=1)
#   self.ssl_tab.grid_rowconfigure(0, weight=1)
#   SuperSupplierTab(master=self.ssl_tab, app_container=self.app_container)\
#       .grid(row=0, column=0, sticky="nsew")
# ─────────────────────────────────────────────────────────────────────────────

import tkinter as tk
from tkinter import ttk, messagebox
from customtkinter import (
    CTkFrame, CTkLabel, CTkFont, CTkButton,
    CTkScrollableFrame, CTkEntry, CTkOptionMenu,
    CTkToplevel, CTkCheckBox, CTkTextbox,
)
import pandas as pd
import json
from datetime import datetime

# =============================================================================
#  PALETTE
# =============================================================================
CLR = {
    "navy":     "#0C2340",
    "blue":     "#1A56DB",
    "blue_lt":  "#EBF5FF",
    "teal":     "#0694A2",
    "teal_lt":  "#E6FFFA",
    "green":    "#057A55",
    "green_lt": "#DEF7EC",
    "amber":    "#B45309",
    "amber_lt": "#FEF3C7",
    "red":      "#C81E1E",
    "red_lt":   "#FDE8E8",
    "gray":     "#6B7280",
    "gray_lt":  "#F3F4F6",
    "white":    "#FFFFFF",
    "border":   "#E5E7EB",
}

TIER_STYLE = {
    "Tier 1":    {"bg": CLR["green_lt"],  "fg": CLR["green"],  "dot": "#10B981"},
    "Tier 2":    {"bg": CLR["blue_lt"],   "fg": CLR["blue"],   "dot": "#3B82F6"},
    "Tier 3":    {"bg": "#FFF7ED",        "fg": "#C2410C",     "dot": "#F97316"},
    "SN":        {"bg": CLR["amber_lt"],  "fg": CLR["amber"],  "dot": "#F59E0B"},
    "Blacklist": {"bg": CLR["red_lt"],    "fg": CLR["red"],    "dot": "#EF4444"},
}

WIN_LOSS_REASONS = [
    "ราคาแพงกว่าคู่แข่ง", "ไม่มีรถจัดส่ง", "สินค้าไม่ครบ",
    "ตอบช้าเกินกำหนด", "สเปกไม่ตรง", "เครดิตเทอมไม่เพียงพอ", "อื่นๆ",
]

# =============================================================================
#  CONSTANTS  (filter options, thresholds, in-memory state)
# =============================================================================
MOCK_SUPPLIERS = []  # replaced by DB


# In-app notifications (demo)
MOCK_NOTIFICATIONS: list = []   # {"id", "msg", "type": "high"|"medium"|"low", "read": bool}


SOURCE_TAG_STYLE = {
    "Legacy": {"bg": "#EDE9FE", "fg": "#5B21B6"},
    "Manual": {"bg": "#E0F2FE", "fg": "#0369A1"},
    "System": {"bg": "#F0FDF4", "fg": "#166534"},
}

DEMOTE_THRESHOLD  = 30   # Win% ต่ำกว่านี้ → แสดง Demote suggestion
PROMOTE_THRESHOLD = 60   # Tier 2 ที่ Score สูงกว่านี้ → แสดง Promote alert

# Auto-Tier thresholds (doc: Super Supplier List Requirement 15.5.2569)
AUTO_TIER_T1        = 85   # score >= 85 AND quality >= KNOCKOUT_QUALITY → Tier 1
AUTO_TIER_T2        = 70   # score 70-84 → Tier 2
KNOCKOUT_QUALITY    = 15   # quality_score ต้องไม่ต่ำกว่านี้จึงจะขึ้น Tier 1 ได้
WIN_GRACE_PERIOD_MO = 3    # เดือน grace period สำหรับ Supplier ใหม่
WIN_GRACE_DEFAULT   = 50   # win_pct default (= 10/20) ระหว่าง grace period

# Quality event-based scoring (0-100 scale; doc values ×5)
QUALITY_START           = 100  # เริ่มต้น 20/20
QUALITY_CLAIM_DELTA     = -50  # เคลม/return → -10/20
QUALITY_RECOVER_GREAT   =  40  # แก้ปัญหาเยี่ยม → +8/20
QUALITY_RECOVER_NORMAL  =  25  # แก้ปัญหาปกติ  → +5/20
QUALITY_RECOVER_BAD     = -25  # ทิ้งงาน       → -5/20

MOCK_TIERS       = ["ทุก Tier", "Tier 1", "Tier 2", "Tier 3", "SN", "Blacklist"]
MOCK_AVAIL       = ["ทุกสถานะ", "พร้อม", "สต็อกต่ำ", "ปิดชั่วคราว"]
MOCK_SOURCE_TAGS = ["ทุก Source", "Legacy", "Manual", "System"]
MOCK_CREDIT_OPT  = ["ทุกเครดิต", "มีเครดิต (>0 วัน)", "เงินสด"]

SN_AGING_DAYS = 30   # SN ที่ค้างเกินกี่วันถือว่า aging

# =============================================================================
#  DATABASE LAYER  — แทนที่ MOCK_DATA ทั้งหมด
# =============================================================================
_DB_CFG = dict(host="Server-APrime", dbname="aplus_com_test",
               user="app_user", password="cailfornia123")

# _app_container เก็บ reference ที่ SuperSupplierTab ส่งมาให้
_app_container = None


def _get_conn():
    """ดึง connection จาก app_container ถ้ามี มิฉะนั้น connect ตรง"""
    if _app_container:
        return _app_container.get_connection(), True   # (conn, use_pool)
    import psycopg2
    return psycopg2.connect(**_DB_CFG), False


def _release_conn(conn, use_pool: bool):
    if use_pool and _app_container:
        _app_container.release_connection(conn)
    else:
        conn.close()


def _to_int_safe(val) -> int:
    """แปลงค่าเป็น int อย่างปลอดภัย รองรับ 'เงินสด', '30 วัน', None, '' ฯลฯ"""
    try:
        cleaned = str(val or "0").replace("วัน", "").strip()
        return int(float(cleaned))
    except Exception:
        return 0


def _auto_suggest_zone(coverage_area: str = "", supplier_name: str = "") -> str:
    """
    คาดเดา dispatch_zone อัตโนมัติจาก coverage_area หรือชื่อ Supplier
    ระบบคิดให้เลย — ไม่ต้องกรอกมือ (9 โซนตามที่ PM กำหนด)
    """
    text = (coverage_area + " " + supplier_name).lower()

    # โซน 2: ตะวันออก (เช็คก่อนกรุงเทพ)
    if any(k in text for k in ["ระยอง", "ชลบุรี", "ตะวันออก", "eastern",
                                "บ้านฉาง", "พัทยา", "มาบตาพุด", "อมตะ"]):
        return "ตะวันออก (ชลบุรี / ระยอง / สมุทรปราการ)"

    # โซน 4: พระราม 2 / พุทธมณฑล / สมุทรสาคร
    if any(k in text for k in ["พุทธมณฑล", "พระราม 2", "สมุทรสาคร",
                                "สมุทรสงคราม", "กาญจนบุรี"]):
        return "พระราม 2 / พุทธมณฑล / สมุทรสาคร"

    # โซน 3: วังน้อย / อยุธยา / สระบุรี
    if any(k in text for k in ["วังน้อย", "อยุธยา", "สระบุรี", "ลพบุรี",
                                "สิงห์บุรี", "อ่างทอง", "นครสวรรค์"]):
        return "วังน้อย / อยุธยา / สระบุรี"

    # โซน 8: ภาคตะวันตก
    if any(k in text for k in ["ราชบุรี", "เพชรบุรี", "ประจวบ", "ภาคตะวันตก",
                                "สุพรรณบุรี", "นครปฐม"]):
        return "ภาคตะวันตก"

    # โซน 5: ภาคเหนือ
    if any(k in text for k in ["เชียงใหม่", "เชียงราย", "ลำปาง", "ลำพูน",
                                "พะเยา", "แพร่", "น่าน", "แม่ฮ่องสอน", "ตาก",
                                "พิษณุโลก", "เพชรบูรณ์", "อุตรดิตถ์", "ภาคเหนือ"]):
        return "ภาคเหนือ"

    # โซน 6: ภาคอีสาน
    if any(k in text for k in ["ขอนแก่น", "อุดร", "นครราชสีมา", "โคราช",
                                "อุบล", "สุรินทร์", "บุรีรัมย์", "ร้อยเอ็ด",
                                "มุกดาหาร", "สกลนคร", "อีสาน", "ภาคอีสาน"]):
        return "ภาคอีสาน"

    # โซน 7: ภาคใต้
    if any(k in text for k in ["สุราษฎร์", "ภูเก็ต", "หาดใหญ่", "สงขลา",
                                "นครศรีธรรมราช", "กระบี่", "ตรัง", "พัทลุง",
                                "ยะลา", "ปัตตานี", "นราธิวาส", "ภาคใต้"]):
        return "ภาคใต้"

    # โซน 1: กทม. / ปริมณฑล (เช็คทีหลังสุด)
    if any(k in text for k in ["กรุงเทพ", "บางกอก", "bangkok", "ปริมณฑล",
                                "สมุทรปราการ", "นนทบุรี", "ปทุม", "รังสิต",
                                "บางนา", "ลาดกระบัง", "มีนบุรี",
                                "ทั่วประเทศ", "national"]):
        return "กทม. / ปริมณฑล"

    return "— ยังไม่ระบุ —"


def _row_to_sup(row: dict) -> dict:
    """แปลง DB row → dict ที่โค้ดเดิมใช้งาน (keys เหมือน MOCK_SUPPLIERS)"""
    win_loss_raw = row.get("win_loss_log") or "[]"
    try:
        wl = json.loads(win_loss_raw) if isinstance(win_loss_raw, str) else win_loss_raw
    except Exception:
        wl = []
    return {
        "id":            row["id"],
        "supplier_id":   row.get("supplier_code") or str(row["id"]),
        "name":          row.get("supplier_name") or "",
        "category":      row.get("category")      or "",
        "tier":          row.get("tier")           or "Tier 2",
        "is_locked":     bool(row.get("is_locked", False)),
        "source_tag":    row.get("source_tag")     or "Manual",
        "contact":       row.get("contact_name")   or "",
        "phone":         row.get("phone_number")   or "",
        "line_id":       row.get("line_id")        or "",
        "email":         row.get("email")          or "",
        "coverage_area": row.get("coverage_area")  or "",
        "availability":  row.get("availability")   or "พร้อม",
        "reopen_date":   row.get("reopen_date")    or "",
        "sn_created":    row.get("sn_created")     or "",
        "win_pct":        _to_int_safe(row.get("win_pct")),
        "sla_score":      _to_int_safe(row.get("sla_score")),
        "price_score":    _to_int_safe(row.get("price_score")),
        "service_score":  _to_int_safe(row.get("service_score")),
        "quality_score":  _to_int_safe(row.get("quality_score")),
        "credit_days":    _to_int_safe(row.get("credit_term")),
        "credit_term_label": row.get("credit_term_label") or "สด",
        "business_type": row.get("business_type")  or "",
        "standard_focus": row.get("standard_focus") or "",
        "note":          row.get("note")            or "",
        "win_loss_log":  wl,
        # ── Zoning Phase 1 ──────────────────────────────────────────────
        "dispatch_zone":    row.get("dispatch_zone") or _auto_suggest_zone(
                                row.get("coverage_area",""), row.get("supplier_name","")),
        "service_area":     row.get("service_area")     or "National",
        "logistics_assets": row.get("logistics_assets") or "",
        "wh_zone":          row.get("wh_zone")          or "",
        "wh_coordinates":   row.get("wh_coordinates")   or "",
    }


def db_get_all_suppliers() -> list:
    """ดึง Supplier ทั้งหมดจาก DB → list of dict"""
    conn, use_pool = _get_conn()
    try:
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("""
                SELECT id, supplier_name, supplier_code, contact_name, phone_number,
                       credit_term, bank_account_type, created_by, created_at,
                       tier, category, is_locked, source_tag,
                       line_id, email, coverage_area, availability,
                       reopen_date, sn_created, win_pct, sla_score,
                       price_score, service_score, quality_score,
                       note, blacklist_reason, win_loss_log,
                       dispatch_zone, service_area, logistics_assets,
                       business_type, standard_focus, credit_term_label,
                       wh_zone, wh_coordinates
                FROM suppliers
                ORDER BY id DESC
            """)
            rows = cur.fetchall()
            return [_row_to_sup(dict(r)) for r in rows]
    except Exception as e:
        print(f"[SSL] db_get_all_suppliers error: {e}")
        return []
    finally:
        _release_conn(conn, use_pool)


def db_get_categories() -> list:
    """ดึง category ที่มีใน DB จริงๆ เท่านั้น ไม่ hardcode"""
    conn, use_pool = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT DISTINCT category
                FROM   suppliers
                WHERE  category IS NOT NULL AND category != ''
                ORDER  BY category
            """)
            cats = [r[0] for r in cur.fetchall()]
        return ["ทุกหมวด"] + cats if cats else ["ทุกหมวด"]
    except Exception as e:
        print(f"[SSL] db_get_categories error: {e}")
        return ["ทุกหมวด"]
    finally:
        _release_conn(conn, use_pool)


def db_save_supplier(sup: dict, action: str, user: str):
    """
    บันทึก / อัปเดต supplier ลง DB
    action: 'add' | 'edit' | 'tier' | 'convert' | 'blacklist'
    """
    conn, use_pool = _get_conn()
    try:
        wl_json = json.dumps(sup.get("win_loss_log", []), ensure_ascii=False)
        with conn.cursor() as cur:
            if action == 'add':
                cur.execute("""
                    INSERT INTO suppliers
                        (supplier_name, supplier_code, contact_name, phone_number,
                         credit_term, tier, category, source_tag,
                         line_id, email, coverage_area, availability,
                         sn_created, win_pct, sla_score, price_score,
                         service_score, quality_score,
                         note, win_loss_log, created_by, created_at,
                         dispatch_zone, service_area, logistics_assets,
                         business_type, standard_focus, credit_term_label,
                         wh_zone, wh_coordinates)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW(),%s,%s,%s,%s,%s,%s,%s,%s)
                    RETURNING id
                """, (
                    sup["name"], sup["supplier_id"], sup.get("contact",""),
                    sup.get("phone",""), sup.get("credit_days", 0),
                    "SN", sup.get("category",""), "Manual",
                    sup.get("line_id",""), sup.get("email",""),
                    sup.get("coverage_area",""), "พร้อม",
                    sup.get("sn_created",""), 0, 0, 0,
                    sup.get("service_score", 0), sup.get("quality_score", 0),
                    sup.get("note",""), "[]", user,
                    sup.get("dispatch_zone",""), sup.get("service_area","National"),
                    sup.get("logistics_assets",""),
                    sup.get("business_type",""), sup.get("standard_focus",""),
                    sup.get("credit_term_label","สด"),
                    sup.get("wh_zone",""), sup.get("wh_coordinates","")
                ))
                new_id = cur.fetchone()[0]
                sup["id"] = new_id
            else:
                cur.execute("""
                    UPDATE suppliers SET
                        supplier_code  = %s,
                        supplier_name  = %s, contact_name  = %s, phone_number = %s,
                        credit_term    = %s, category      = %s, tier         = %s,
                        is_locked      = %s, source_tag    = %s, line_id      = %s,
                        email          = %s, coverage_area = %s, availability = %s,
                        reopen_date    = %s, win_pct       = %s, sla_score    = %s,
                        price_score    = %s, service_score = %s,
                        note           = %s, win_loss_log  = %s,
                        blacklist_reason = %s,
                        dispatch_zone    = %s, service_area  = %s, logistics_assets = %s,
                        business_type    = %s, standard_focus = %s, credit_term_label = %s,
                        wh_zone          = %s, wh_coordinates  = %s
                    WHERE id = %s
                """, (
                    sup.get("supplier_id", ""),
                    sup["name"], sup.get("contact",""), sup.get("phone",""),
                    sup.get("credit_days", 0), sup.get("category",""), sup.get("tier","Tier 2"),
                    sup.get("is_locked", False), sup.get("source_tag","Manual"),
                    sup.get("line_id",""), sup.get("email",""),
                    sup.get("coverage_area",""), sup.get("availability","พร้อม"),
                    sup.get("reopen_date",""), sup.get("win_pct", 0),
                    sup.get("sla_score", 0), sup.get("price_score", 0),
                    sup.get("service_score", 0),
                    sup.get("note",""), wl_json,
                    sup.get("blacklist_reason",""),
                    sup.get("dispatch_zone",""), sup.get("service_area","National"),
                    sup.get("logistics_assets",""),
                    sup.get("business_type",""), sup.get("standard_focus",""),
                    sup.get("credit_term_label","สด"),
                    sup.get("wh_zone",""), sup.get("wh_coordinates",""),
                    sup["id"]
                ))

            # Audit log
            cur.execute("""
                INSERT INTO audit_log (action, table_name, record_id, user_info, changes, timestamp)
                VALUES (%s, %s, %s, %s, %s, NOW())
            """, (
                f"SSL:{action}", "suppliers", sup.get("id"),
                user,
                json.dumps({"supplier_id": sup.get("supplier_id"), "tier": sup.get("tier")},
                           ensure_ascii=False)
            ))

        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"[SSL] db_save_supplier error: {e}")
        return False
    finally:
        _release_conn(conn, use_pool)


def db_check_sw_duplicate(sw_code: str, exclude_id: int = 0) -> bool:
    """True = ซ้ำ, False = ไม่ซ้ำ"""
    conn, use_pool = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT 1 FROM suppliers WHERE supplier_code = %s AND id != %s LIMIT 1",
                (sw_code, exclude_id)
            )
            return cur.fetchone() is not None
    except Exception:
        return False
    finally:
        _release_conn(conn, use_pool)


def db_next_sn_code(user: str) -> str:
    """สร้างรหัส SN ถัดไป: SN[YY]-[NNNN]-[USER]
    ใช้ MAX ของเลข sequence แทน COUNT เพื่อป้องกัน duplicate เมื่อมีการลบ record"""
    conn, use_pool = _get_conn()
    try:
        yy = datetime.now().strftime("%y")
        with conn.cursor() as cur:
            # หาเลข sequence สูงสุดของปีนี้
            cur.execute(
                """SELECT COALESCE(
                       MAX(CAST(SPLIT_PART(supplier_code, '-', 2) AS INTEGER)), 0
                   )
                   FROM suppliers
                   WHERE supplier_code LIKE %s
                     AND supplier_code ~ '^SN[0-9]{2}-[0-9]{4}-'""",
                (f"SN{yy}-%",)
            )
            max_n = cur.fetchone()[0] or 0
            # หา sequence ถัดไปที่ไม่ซ้ำ
            n = max_n + 1
            for _ in range(100):  # loop สูงสุด 100 ครั้งกันไม่รู้จบ
                candidate = f"SN{yy}-{n:04d}-{user}"
                cur.execute(
                    "SELECT 1 FROM suppliers WHERE supplier_code = %s", (candidate,)
                )
                if not cur.fetchone():
                    break
                n += 1
        return f"SN{yy}-{n:04d}-{user}"
    except Exception as e:
        print(f"[SSL] db_next_sn_code error: {e}")
        import random
        return f"SN{datetime.now().strftime('%y')}-{random.randint(1000, 9999):04d}-{user}"
    finally:
        _release_conn(conn, use_pool)


def db_get_audit_log() -> list:
    conn, use_pool = _get_conn()
    try:
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("""
                SELECT timestamp, action, user_info, changes
                FROM audit_log
                WHERE action LIKE 'SSL:%'
                ORDER BY timestamp DESC
                LIMIT 200
            """)
            rows = cur.fetchall()
            return [{
                "timestamp": str(r["timestamp"])[:16],
                "action":    r["action"].replace("SSL:", ""),
                "user":      r["user_info"],
                "detail":    r["changes"] or "",
            } for r in rows]
    except Exception as e:
        print(f"[SSL] db_get_audit_log error: {e}")
        return []
    finally:
        _release_conn(conn, use_pool)


# =============================================================================
#  QUARTERLY SNAPSHOT  DB FUNCTIONS
# =============================================================================
def db_get_quarterly_snapshots(cat: str) -> list:
    """
    ดึง Quarterly Snapshot จาก DB สำหรับหมวดที่กำหนด
    คืน list ของ {"quarter": "Q1/2025", "top5": [(name, score), ...]}
    """
    conn, use_pool = _get_conn()
    try:
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("""
                SELECT quarter_label, snapshot_data
                FROM   supplier_quarterly_snapshots
                WHERE  category = %s
                ORDER  BY quarter_year ASC, quarter_num ASC
                LIMIT  4
            """, (cat,))
            rows = cur.fetchall()
            if rows:
                result = []
                for r in rows:
                    try:
                        data = json.loads(r["snapshot_data"]) if isinstance(r["snapshot_data"], str) \
                               else r["snapshot_data"]
                        result.append({
                            "quarter": r["quarter_label"],
                            "top5":    [tuple(x) for x in data.get("top5", [])]
                        })
                    except Exception:
                        pass
                if result:
                    return result
    except Exception as e:
        print(f"[SSL] db_get_quarterly_snapshots error: {e}")
    finally:
        _release_conn(conn, use_pool)
    # Fallback → สร้าง snapshot จาก live data ปัจจุบัน
    return _build_live_snapshot(cat)


def _calc_supplier_scores_from_benchmark(cat: str = None) -> dict:
    """
    คำนวณ win_pct จาก cost_benchmarks จริง
    คืน dict: {supplier_name: {"win_pct": int, "price_score": int, "name": str, "category": str, "tier": str}}
    """
    conn, use_pool = _get_conn()
    try:
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            # ดึง column names จาก cost_benchmarks
            cur.execute("""
                SELECT column_name FROM information_schema.columns
                WHERE table_name = 'cost_benchmarks' ORDER BY ordinal_position
            """)
            cols = [r[0] for r in cur.fetchall()]

            # หา column ชื่อ Supplier, สถานะ, หมวด, ต้นทุน/เส้น
            sup_col   = next((c for c in cols if 'Supplier' in c
                              and 'Supplier2' not in c and 'ID' not in c
                              and not c.startswith('Sup')), None)
            stat_col  = next((c for c in cols if c == 'สถานะ'), None)
            cat_col   = next((c for c in cols if c == 'หมวด'), None)
            cost_col  = next((c for c in cols if c == 'ต้นทุน/เส้น'), None)

            if not sup_col or not stat_col:
                return {}

            # WHERE clause สำหรับ filter category
            cat_filter = f'AND "{cat_col}" = %s' if cat_col and cat else ''
            params = (cat,) if cat_col and cat else ()

            # คำนวณ win_pct และ avg cost ต่อ Supplier
            # pre-compute to avoid backslash-in-f-string-expression (Python < 3.12)
            if cost_col:
                cost_avg_expr = f"""AVG(NULLIF(NULLIF(REGEXP_REPLACE("{cost_col}", '[^0-9.]', '', 'g'), '')::numeric, 0))"""
            else:
                cost_avg_expr = "0"
            query = f"""
                SELECT
                    "{sup_col}"   AS sup_name,
                    {f'"{cat_col}" AS category,' if cat_col else "'' AS category,"}
                    COUNT(*)      AS total,
                    SUM(CASE WHEN "{stat_col}" = 'WIN' THEN 1 ELSE 0 END) AS wins,
                    {cost_avg_expr} AS avg_cost
                FROM cost_benchmarks
                WHERE "{sup_col}" IS NOT NULL AND "{sup_col}" != ''
                  AND "{stat_col}" IS NOT NULL AND "{stat_col}" != ''
                  {cat_filter}
                GROUP BY "{sup_col}" {f', "{cat_col}"' if cat_col else ''}
                HAVING COUNT(*) >= 1
            """
            cur.execute(query, params)
            rows = cur.fetchall()

            if not rows:
                return {}

            # หา min/max cost สำหรับ normalize เป็น price_score
            costs = [float(r["avg_cost"] or 0) for r in rows if r["avg_cost"]]
            max_cost = max(costs) if costs else 1
            min_cost = min(costs) if costs else 0

            result = {}
            for r in rows:
                sup_name = r["sup_name"]
                total    = int(r["total"] or 1)
                wins     = int(r["wins"] or 0)
                win_pct  = round(wins / total * 100) if total > 0 else 0
                avg_cost = float(r["avg_cost"] or 0)

                # price_score: ยิ่งราคาถูกยิ่งสูง (100 = ถูกสุด, 0 = แพงสุด)
                if max_cost > min_cost and avg_cost > 0:
                    price_score = round(100 - ((avg_cost - min_cost) / (max_cost - min_cost)) * 100)
                else:
                    price_score = 50  # default ถ้าไม่มีข้อมูลเปรียบเทียบ

                result[sup_name] = {
                    "name":        sup_name,
                    "category":    r["category"] or "",
                    "win_pct":     win_pct,
                    "price_score": max(0, min(100, price_score)),
                    "sla_score":   0,   # ยังไม่มีข้อมูล SLA จาก benchmark
                    "total":       total,
                    "wins":        wins,
                }
            return result
    except Exception as e:
        print(f"[SSL] _calc_supplier_scores_from_benchmark error: {e}")
        return {}
    finally:
        _release_conn(conn, use_pool)


def _build_live_snapshot(cat: str) -> list:
    """สร้าง snapshot จาก live data — ดึง win_pct จาก cost_benchmarks จริง"""
    # ดึง scores จาก cost_benchmarks
    bench_scores = _calc_supplier_scores_from_benchmark(cat)

    # ดึง tier จาก suppliers
    all_sups = db_get_all_suppliers()
    tier_map = {s["name"]: s["tier"] for s in all_sups}
    cat_map  = {s["name"]: s["category"] for s in all_sups}

    if not bench_scores:
        return []

    # รวมข้อมูล: ใช้ win_pct จาก benchmark + tier จาก suppliers
    import pandas as pd
    rows = []
    for sup_name, scores in bench_scores.items():
        tier     = tier_map.get(sup_name, "Tier 2")
        category = cat_map.get(sup_name) or scores.get("category", "")
        if cat and category != cat:
            continue
        if tier not in ("Tier 1", "Tier 2", "Tier 3"):
            continue
        win_pct     = scores["win_pct"]
        price_score = scores["price_score"]
        # ราคา 20% + สต็อก(Win%) 20% — service/sla/quality ยังไม่มีใน snapshot → เฉลี่ย 2 ตัวที่มี
        score = round((price_score + win_pct) / 2)
        rows.append({"name": sup_name, "score": score,
                     "win_pct": win_pct, "tier": tier})

    if not rows:
        return []

    df = pd.DataFrame(rows).sort_values("score", ascending=False).head(5)
    now   = datetime.now()
    q_num = (now.month - 1) // 3 + 1
    top5  = [(row["name"], int(row["score"])) for _, row in df.iterrows()]
    while len(top5) < 5:
        top5.append(("", 0))
    return [{"quarter": f"Q{q_num}/{now.year}", "top5": top5}]


def db_save_quarterly_snapshot():
    """
    บันทึก Top 5 ปัจจุบันลง DB เป็น Snapshot รายไตรมาส
    ปลอดภัย: ON CONFLICT → UPDATE (ไม่ duplicate)
    """
    conn, use_pool = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS supplier_quarterly_snapshots (
                    id             SERIAL PRIMARY KEY,
                    category       TEXT NOT NULL,
                    quarter_label  TEXT NOT NULL,
                    quarter_year   INT  NOT NULL,
                    quarter_num    INT  NOT NULL,
                    snapshot_data  JSONB,
                    created_at     TIMESTAMP DEFAULT NOW(),
                    UNIQUE(category, quarter_label)
                )
            """)
        conn.commit()

        import pandas as pd
        all_sups = db_get_all_suppliers()
        df_all = pd.DataFrame(all_sups) if all_sups else pd.DataFrame(
            columns=["name","category","tier","win_pct","sla_score","price_score",
                     "service_score","quality_score"])
        if df_all.empty:
            return 0

        now     = datetime.now()
        q_num   = (now.month - 1) // 3 + 1
        q_label = f"Q{q_num}/{now.year}"
        saved   = 0

        # ดึง scores จาก cost_benchmarks จริง (ทุก category)
        _all_sups_snap = db_get_all_suppliers()
        bench_scores = _calc_supplier_scores_from_benchmark(cat=None)
        tier_map = {s["name"]: s["tier"] for s in _all_sups_snap}
        cat_map_snap = {s["name"]: s["category"] for s in _all_sups_snap}

        # จัดกลุ่มตาม category
        import pandas as pd
        rows_all = []
        for sup_name, scores in bench_scores.items():
            tier     = tier_map.get(sup_name, "Tier 2")
            # ใช้ category จาก suppliers table ก่อน (ป้องกัน category เก่าจาก benchmark)
            category = cat_map_snap.get(sup_name) or scores.get("category", "")
            if not category or tier not in ("Tier 1", "Tier 2"):
                continue
            win_pct     = scores["win_pct"]
            price_score = scores["price_score"]
            score = round((price_score + win_pct) / 2)
            rows_all.append({"name": sup_name, "category": category,
                             "score": score, "tier": tier})

        if not rows_all:
            # fallback ถ้า benchmark ไม่มีข้อมูล ใช้ suppliers โดยตรง
            df_all["score"] = df_all.apply(calc_score, axis=1)
            rows_all = df_all[df_all["tier"].isin(["Tier 1","Tier 2"])][
                ["name","category","score"]].to_dict("records")

        df_rows = pd.DataFrame(rows_all)
        cats = [c for c in df_rows["category"].dropna().unique() if c]

        with conn.cursor() as cur:
            for cat in cats:
                cat_df = df_rows[df_rows["category"] == cat].sort_values(
                          "score", ascending=False).head(5)
                top5 = [(row["name"], int(row["score"])) for _, row in cat_df.iterrows()]
                while len(top5) < 5:
                    top5.append(("", 0))
                snapshot = json.dumps({"top5": top5}, ensure_ascii=False)
                cur.execute("""
                    INSERT INTO supplier_quarterly_snapshots
                        (category, quarter_label, quarter_year, quarter_num, snapshot_data)
                    VALUES (%s, %s, %s, %s, %s)
                    ON CONFLICT (category, quarter_label)
                    DO UPDATE SET snapshot_data = EXCLUDED.snapshot_data,
                                  created_at    = NOW()
                """, (cat, q_label, now.year, q_num, snapshot))
                saved += 1
        conn.commit()
        print(f"[SSL] Snapshot saved: {saved} categories → {q_label}")
        return saved
    except Exception as e:
        conn.rollback()
        print(f"[SSL] db_save_quarterly_snapshot error: {e}")
        return 0
    finally:
        _release_conn(conn, use_pool)


# =============================================================================
#  WEIGHTED SCORE  ราคา 20% + สต็อก(Win%) 20% + บริการ 20% + SLA 20% + คุณภาพ 20%
# =============================================================================
def calc_score(sup) -> int:
    if isinstance(sup, dict):
        p  = sup.get("price_score",   0)
        w  = sup.get("win_pct",       0)
        sv = sup.get("service_score", 0)
        s  = sup.get("sla_score",     0)
        q  = sup.get("quality_score", 0)
    else:
        p  = sup["price_score"]
        w  = sup["win_pct"]
        sv = sup.get("service_score", 0)
        s  = sup["sla_score"]
        q  = sup.get("quality_score", 0)
    return round((p + w + sv + s + q) / 5)


def calc_auto_tier(score: int, quality_score: int,
                   is_locked: bool, current_tier: str) -> str:
    """
    คำนวณ Tier อัตโนมัติตาม score + Knockout Rule
    - Blacklist / SN → ไม่เปลี่ยน (ต้องกด manual)
    - is_locked      → ล็อค Tier 1 ถาวร
    - score >= 85 AND quality >= 15 → Tier 1
    - score >= 70                   → Tier 2
    - score <  70                   → Tier 3
    """
    if current_tier in ("Blacklist", "SN"):
        return current_tier
    if is_locked:
        return "Tier 1"
    if score >= AUTO_TIER_T1 and quality_score >= KNOCKOUT_QUALITY:
        return "Tier 1"
    elif score >= AUTO_TIER_T2:
        return "Tier 2"
    else:
        return "Tier 3"


def _calc_all_sla_scores() -> dict:
    """
    คำนวณ SLA score ทุก Supplier พร้อมกันใน 1 query (GROUP BY)
    คืน dict: {supplier_name: sla_score (0-100)}
    """
    conn, use_pool = _get_conn()
    try:
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("""
                SELECT
                    supplier_name,
                    COUNT(*) AS total,
                    SUM(CASE WHEN actual_received_date IS NOT NULL
                              AND actual_received_date <= expected_delivery_date
                             THEN 1 ELSE 0 END) AS on_time
                FROM purchase_orders
                WHERE expected_delivery_date IS NOT NULL
                  AND supplier_name IS NOT NULL AND supplier_name != ''
                GROUP BY supplier_name
            """)
            rows = cur.fetchall()
            result = {}
            for r in rows:
                total   = int(r["total"]   or 0)
                on_time = int(r["on_time"] or 0)
                if total > 0:
                    result[r["supplier_name"]] = round(on_time / total * 100)
            return result
    except Exception as e:
        print(f"[SSL] _calc_all_sla_scores error: {e}")
        return {}
    finally:
        _release_conn(conn, use_pool)


# =============================================================================
#  QUALITY EVENT-BASED SCORING
# =============================================================================

def _ensure_quality_events_table(cur):
    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplier_quality_events (
            id              SERIAL PRIMARY KEY,
            supplier_name   TEXT    NOT NULL,
            event_type      TEXT    NOT NULL,
            delta           INT     NOT NULL,
            reason          TEXT,
            recovery_label  TEXT,
            resolved        BOOLEAN DEFAULT FALSE,
            parent_event_id INT,
            created_by      TEXT,
            created_at      TIMESTAMP DEFAULT NOW()
        )
    """)


def db_add_quality_event(supplier_name: str, reason: str, user: str) -> int:
    """เพิ่ม claim event: หัก quality_score และบันทึก history, คืน event_id"""
    import psycopg2 as _pg2
    conn = _pg2.connect(**_DB_CFG)
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO supplier_quality_events
                    (supplier_name, event_type, delta, reason, resolved, created_by)
                VALUES (%s, 'claim', %s, %s, FALSE, %s)
                RETURNING id
            """, (supplier_name, QUALITY_CLAIM_DELTA, reason, user))
            event_id = cur.fetchone()[0]
            cur.execute("""
                UPDATE suppliers
                SET quality_score = GREATEST(0, LEAST(100, quality_score + %s))
                WHERE supplier_name = %s
            """, (QUALITY_CLAIM_DELTA, supplier_name))
            print(f"[SSL] quality UPDATE rowcount={cur.rowcount} supplier={supplier_name!r}")
            conn.commit()
            return event_id
    except Exception as e:
        conn.rollback()
        print(f"[SSL] db_add_quality_event error: {e}")
        return -1
    finally:
        conn.close()


def db_resolve_quality_event(event_id: int, recovery_label: str,
                              supplier_name: str, user: str) -> bool:
    """Resolve claim event ด้วย recovery type และอัปเดต quality_score"""
    delta_map = {
        "great":  QUALITY_RECOVER_GREAT,
        "normal": QUALITY_RECOVER_NORMAL,
        "bad":    QUALITY_RECOVER_BAD,
    }
    delta = delta_map.get(recovery_label, 0)
    import psycopg2 as _pg2
    conn = _pg2.connect(**_DB_CFG)
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE supplier_quality_events
                SET resolved = TRUE
                WHERE id = %s AND resolved = FALSE
            """, (event_id,))
            cur.execute("""
                INSERT INTO supplier_quality_events
                    (supplier_name, event_type, delta, recovery_label,
                     resolved, parent_event_id, created_by)
                VALUES (%s, 'recovery', %s, %s, TRUE, %s, %s)
            """, (supplier_name, delta, recovery_label, event_id, user))
            cur.execute("""
                UPDATE suppliers
                SET quality_score = GREATEST(0, LEAST(100, quality_score + %s))
                WHERE supplier_name = %s
            """, (delta, supplier_name))
            print(f"[SSL] resolve UPDATE rowcount={cur.rowcount} supplier={supplier_name!r} delta={delta}")
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        print(f"[SSL] db_resolve_quality_event error: {e}")
        return False
    finally:
        conn.close()


def db_get_quality_events(supplier_name: str) -> list:
    """ดึง quality event history ของ supplier (ล่าสุด 10 รายการ)"""
    conn, use_pool = _get_conn()
    try:
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("""
                SELECT id, event_type, delta, reason, recovery_label,
                       resolved, created_by,
                       TO_CHAR(created_at, 'DD/MM/YY HH24:MI') AS ts
                FROM supplier_quality_events
                WHERE supplier_name = %s
                ORDER BY created_at DESC
                LIMIT 10
            """, (supplier_name,))
            return [dict(r) for r in cur.fetchall()]
    except Exception as e:
        print(f"[SSL] db_get_quality_events error: {e}")
        return []
    finally:
        _release_conn(conn, use_pool)


# =============================================================================
#  QUERY
# =============================================================================
def get_suppliers_df(cat="ทุกหมวด", tier="ทุก Tier", avail="ทุกสถานะ",
                     search="", source="ทุก Source", credit="ทุกเครดิต"):
    """ดึงข้อมูลจาก DB จริง แล้ว filter ใน Python เหมือนเดิม
    win_pct และ price_score ดึงจาก cost_benchmarks จริง (ไม่ใช่ 0 ใน suppliers)
    """
    all_sups = db_get_all_suppliers()
    df = pd.DataFrame(all_sups) if all_sups else pd.DataFrame(columns=[
        "id","supplier_id","name","category","tier","is_locked","source_tag",
        "contact","phone","line_id","email","coverage_area","availability",
        "reopen_date","sn_created","win_pct","sla_score","price_score",
        "service_score","quality_score","credit_days","note","win_loss_log"
    ])
    if df.empty:
        df["score"] = []
        return df

    # ── Merge กับ benchmark scores (vectorized — ไม่ใช้ lambda closure) ───────
    bench = _calc_supplier_scores_from_benchmark(cat=None)
    if bench:
        bench_df = pd.DataFrame([
            {"name": k, "_bwp": v["win_pct"], "_bps": v["price_score"],
             "_btotal": v.get("total", 99)}
            for k, v in bench.items()
        ])
        df = df.merge(bench_df, on="name", how="left")
        df["win_pct"]     = df["_bwp"].fillna(df["win_pct"]).astype(int)
        df["price_score"] = df["_bps"].fillna(df["price_score"]).astype(int)
        df["_bench_total"] = df["_btotal"].fillna(99).astype(int)
        df = df.drop(columns=["_bwp", "_bps", "_btotal"])
    else:
        df["_bench_total"] = 99

    # ── Grace Period: Supplier ใหม่ (< 3 เดือน หรือ quote < 5 ครั้ง) ─────────
    from datetime import date as _date
    today = _date.today()
    def _apply_grace(row):
        if row.get("tier") not in ("Tier 1", "Tier 2", "Tier 3"):
            return int(row.get("win_pct", 0))
        created_str = str(row.get("sn_created") or "")[:10]
        if created_str:
            try:
                created = datetime.strptime(created_str, "%Y-%m-%d").date()
                months_old = (today.year - created.year) * 12 + (today.month - created.month)
                if months_old < WIN_GRACE_PERIOD_MO and int(row.get("_bench_total", 99)) < 5:
                    return WIN_GRACE_DEFAULT
            except ValueError:
                pass
        return int(row.get("win_pct", 0))
    df["win_pct"] = df.apply(_apply_grace, axis=1)

    # ── Merge SLA จาก purchase_orders (1 batch query แทน N queries) ──────────
    sla_map = _calc_all_sla_scores()
    if sla_map:
        df["sla_score"] = df.apply(
            lambda row: sla_map.get(row["name"]) or int(row.get("sla_score") or 0),
            axis=1
        )

    # ── Filters (ยกเว้น tier — ต้องคำนวณ auto-tier ก่อน) ──────────────────────
    if cat    != "ทุกหมวด":    df = df[df["category"]    == cat]
    if avail  != "ทุกสถานะ":   df = df[df["availability"] == avail]
    if source != "ทุก Source": df = df[df["source_tag"]   == source]
    if credit == "มีเครดิต (>0 วัน)":
        df = df[df["credit_days"] > 0]
    elif credit == "เงินสด":
        df = df[df["credit_days"] == 0]
    if search:
        kw = search.lower()
        df = df[
            df["name"].str.lower().str.contains(kw, na=False) |
            df["supplier_id"].str.lower().str.contains(kw, na=False) |
            df["contact"].str.lower().str.contains(kw, na=False) |
            df["coverage_area"].str.lower().str.contains(kw, na=False)
        ]
    df = df.copy()
    # ── Score: 5 เกณฑ์ × 20% ─────────────────────────────────────────────────
    for col in ("service_score", "quality_score"):
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    df["score"] = (
        df["price_score"].astype(float)   +
        df["win_pct"].astype(float)        +
        df["service_score"].astype(float)  +
        df["sla_score"].astype(float)      +
        df["quality_score"].astype(float)
    ).div(5).round().astype(int)

    # ── Auto-Tier: คำนวณ tier จาก score + Knockout Rule ──────────────────────
    if "_bench_total" in df.columns:
        df = df.drop(columns=["_bench_total"])
    _scores   = df["score"].to_numpy()
    _quality  = df["quality_score"].fillna(0).astype(int).to_numpy()
    _locked   = df["is_locked"].fillna(False).astype(bool).to_numpy()
    _tiers    = df["tier"].fillna("Tier 2").astype(str).to_numpy()
    df["tier"] = [
        calc_auto_tier(int(s), int(q), bool(lk), str(t))
        for s, q, lk, t in zip(_scores, _quality, _locked, _tiers)
    ]

    # ── Tier filter หลัง auto-tier เพื่อให้ตรงกับ tier จริง ──────────────────
    if tier != "ทุก Tier":
        df = df[df["tier"] == tier]

    return df.reset_index(drop=True)


def get_aging_sns(threshold_days=SN_AGING_DAYS):
    """คืนรายชื่อ SN ที่ยังไม่ Convert เกิน threshold_days — ดึงจาก DB"""
    today  = datetime.now().date()
    result = []
    for s in db_get_all_suppliers():
        if s["tier"] != "SN":
            continue
        created_str = s.get("sn_created", "")
        if not created_str:
            continue
        try:
            created = datetime.strptime(created_str[:10], "%Y-%m-%d").date()
            days    = (today - created).days
            if days >= threshold_days:
                result.append({**s, "_aging_days": days})
        except ValueError:
            pass
    return result

# =============================================================================
#  FUZZY MATCH
# =============================================================================
def fuzzy_candidates(name: str, threshold=35):
    """Fuzzy search ชื่อ Supplier จาก DB"""
    name_l = name.lower()
    if not name_l:
        return []
    results = []
    for s in db_get_all_suppliers():
        sn = s["name"].lower()
        if name_l in sn or sn in name_l:
            score = 90
        else:
            common = sum(c in sn for c in name_l)
            score  = int(common / max(len(name_l), 1) * 100)
        if score >= threshold:
            results.append((score, s))
    results.sort(key=lambda x: -x[0])
    return [s for _, s in results[:5]]

# =============================================================================
#  TTK STYLE
# =============================================================================
_TTK_STYLED = False

def _place_popup(popup, width: int, height: int):
    """
    วาง popup กึ่งกลาง parent window บนจอเดียวกัน
    รองรับ multi-monitor: ใช้พิกัด parent จริง ไม่ใช้ winfo_screenwidth (=จอหลักเสมอ)
    """
    popup.update_idletasks()

    import tkinter as _tk
    parent = popup.master
    while parent and not isinstance(parent, (_tk.Tk, _tk.Toplevel)):
        try:
            parent = parent.master
        except Exception:
            break

    if parent is None:
        popup.geometry(f"{width}x{height}")
        return

    parent.update_idletasks()

    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    pw = parent.winfo_width()  if parent.winfo_width()  > 10 else parent.winfo_reqwidth()
    ph = parent.winfo_height() if parent.winfo_height() > 10 else parent.winfo_reqheight()

    # กึ่งกลาง popup บน parent
    cx = px + (pw - width)  // 2
    cy = py + (ph - height) // 2

    # Clamp ให้อยู่ใน boundary ของจอที่ parent อยู่
    # ไม่ใช้ winfo_screenwidth เพราะมันคืนค่าจอหลักเสมอ ทำให้วิ่งไปจออื่น
    margin = 8
    cx = max(px + margin, min(cx, px + pw - width  - margin))
    cy = max(py + margin, min(cy, py + ph - height - margin))

    popup.geometry(f"{width}x{height}+{cx}+{cy}")

def _apply_ttk_style():
    global _TTK_STYLED
    if _TTK_STYLED:
        return
    _TTK_STYLED = True
    st = ttk.Style()
    try:
        st.theme_use("default")
    except Exception:
        pass
    st.configure("SSL.Treeview",
                 font=("Tahoma", 12), rowheight=32,
                 background=CLR["white"], fieldbackground=CLR["white"],
                 foreground="#1F2937", borderwidth=0)
    st.configure("SSL.Treeview.Heading",
                 font=("Tahoma", 13, "bold"),
                 background=CLR["navy"], foreground=CLR["white"],
                 relief="flat", borderwidth=0, padding=(6, 6))
    st.map("SSL.Treeview",
           background=[("selected", "#BFDBFE")],
           foreground=[("selected", CLR["navy"])])

# =============================================================================
#  CONVERT SN → SW  POPUP
# =============================================================================
class ConvertToSWPopup(CTkToplevel):

    def __init__(self, master, supplier: dict, on_success=None, current_user="USER_DEMO"):
        super().__init__(master)
        self.sup          = dict(supplier)
        self.on_success   = on_success
        self.current_user = current_user
        self.title("Convert SN → SW Official ID")
        _place_popup(self, 500, 460)
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        CTkLabel(hdr, text="Convert SN → SW Official ID",
                 font=F(size=15, weight="bold"),
                 text_color=CLR["white"]).pack(padx=20, pady=12, anchor="w")

        # SN info
        info = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=8)
        info.grid(row=1, column=0, padx=20, pady=(16, 0), sticky="ew")
        info.grid_columnconfigure(1, weight=1)
        for r, (lbl, val) in enumerate([
            ("ชื่อ",    supplier["name"]),
            ("รหัส SN", supplier["supplier_id"]),
            ("หมวด",    supplier["category"]),
            ("Win %",   f"{supplier['win_pct']}%"),
        ]):
            CTkLabel(info, text=lbl + ":", text_color=CLR["gray"],
                     font=F(size=12)).grid(row=r, column=0, padx=(12, 8), pady=3, sticky="w")
            CTkLabel(info, text=val, font=F(size=12, weight="bold")).grid(
                row=r, column=1, sticky="w", pady=3)

        CTkLabel(self, text="Express Official ID  (SW รหัสทางการ)",
                 font=F(size=13, weight="bold")).grid(
            row=2, column=0, padx=20, pady=(16, 4), sticky="w")
        self._sw_entry = CTkEntry(self, placeholder_text="เช่น SW690123",
                                  font=F(size=14), height=38)
        self._sw_entry.grid(row=3, column=0, padx=20, sticky="ew")
        self._sw_entry.bind("<KeyRelease>", self._validate)

        self._val_lbl = CTkLabel(self, text="", font=F(size=12), anchor="w")
        self._val_lbl.grid(row=4, column=0, padx=20, pady=(4, 0), sticky="w")

        CTkLabel(self, text="หมายเหตุ (Audit Note)",
                 font=F(size=13, weight="bold")).grid(
            row=5, column=0, padx=20, pady=(14, 4), sticky="w")
        self._note = CTkTextbox(self, height=65, font=F(size=12))
        self._note.grid(row=6, column=0, padx=20, sticky="ew")

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=7, column=0, sticky="ew")
        CTkButton(bf, text="ยกเลิก", fg_color="gray50", hover_color="gray40",
                  width=90, command=self.destroy).pack(side="right", padx=12, pady=10)
        self._ok_btn = CTkButton(bf, text="ยืนยัน Convert",
                                 fg_color=CLR["green"], hover_color="#065F46",
                                 width=140, state="disabled",
                                 command=self._do_convert)
        self._ok_btn.pack(side="right", pady=10)

        self.transient(master)
        self.grab_set()

    def _validate(self, _=None):
        import re
        val = self._sw_entry.get().strip()
        if not val:
            self._val_lbl.configure(text="", text_color=CLR["gray"])
            self._ok_btn.configure(state="disabled")
            return
        if not re.match(r'^SW\d{6}$', val):
            self._val_lbl.configure(
                text="✗  รูปแบบไม่ถูกต้อง — ต้องเป็น SW + 6 หลัก เช่น SW690123",
                text_color=CLR["red"])
            self._ok_btn.configure(state="disabled")
            return
        if db_check_sw_duplicate(val, exclude_id=self.sup.get("id", 0)):
            self._val_lbl.configure(
                text="✗  รหัสนี้มีอยู่ในระบบแล้ว",
                text_color=CLR["red"])
            self._ok_btn.configure(state="disabled")
            return
        self._val_lbl.configure(
            text="✓  รูปแบบถูกต้อง และไม่ซ้ำในระบบ",
            text_color=CLR["green"])
        self._ok_btn.configure(state="normal")

    def _do_convert(self):
        new_id = self._sw_entry.get().strip()
        note   = self._note.get("1.0", "end").strip()
        old_id = self.sup["supplier_id"]
        # อัปเดต dict แล้วบันทึก DB
        updated = dict(self.sup)
        updated["supplier_id"] = new_id
        updated["tier"]        = "Tier 2"
        updated["note"]        = (updated.get("note","") + f" | Converted from {old_id}").strip(" |")
        ok = db_save_supplier(updated, action="convert", user=self.current_user)
        if ok:
            messagebox.showinfo("Convert สำเร็จ",
                                f"รหัส {old_id} → {new_id}\n"
                                f"Tier ปรับเป็น Tier 2 อัตโนมัติ\n"
                                f"หมายเหตุ: {note or '-'}", parent=self)
        else:
            messagebox.showerror("ผิดพลาด", "บันทึกลงฐานข้อมูลไม่สำเร็จ", parent=self)
        if self.on_success:
            self.on_success()
        self.destroy()

# =============================================================================
#  SUPPLIER DETAIL POPUP
# =============================================================================
class SupplierDetailPopup(CTkToplevel):

    def __init__(self, master, supplier: dict, on_save=None, current_user="USER_DEMO"):
        super().__init__(master)
        self.sup          = dict(supplier)
        self.on_save      = on_save
        self.current_user = current_user
        self.title(f"Supplier Profile — {supplier['name']}")
        _place_popup(self, 620, 680)
        self.resizable(False, True)
        self.grid_columnconfigure(0, weight=1)
        F  = CTkFont
        ts = TIER_STYLE.get(supplier.get("tier", "SN"),
                             {"bg": CLR["gray_lt"], "fg": CLR["gray"]})

        # Header
        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text=supplier["name"],
                 font=F(size=17, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=20, pady=(14, 2), sticky="w")
        CTkLabel(hdr, text=supplier["supplier_id"],
                 font=F(size=12), text_color="#93C5FD").grid(
            row=1, column=0, padx=20, pady=(0, 12), sticky="w")

        # Tier badge — column 1, ไม่ใช้ rowspan แล้ว
        lock_prefix = "🔒 " if supplier.get("is_locked") else ""
        CTkLabel(hdr, text=f"  {lock_prefix}{supplier.get('tier','SN')}  ",
                 fg_color=ts["bg"], text_color=ts["fg"],
                 corner_radius=6, font=F(size=12, weight="bold")).grid(
            row=0, column=1, padx=(0, 8), pady=(14, 4), sticky="e")

        # source_tag badge — column 1 row 1 (ใต้ Tier badge)
        src     = supplier.get("source_tag", "")
        src_sty = SOURCE_TAG_STYLE.get(src, {"bg": CLR["gray_lt"], "fg": CLR["gray"]})
        if src:
            CTkLabel(hdr, text=f"  {src}  ",
                     fg_color=src_sty["bg"], text_color=src_sty["fg"],
                     corner_radius=6, font=F(size=11)).grid(
                row=1, column=1, padx=(0, 20), pady=(0, 12), sticky="e")

        # Scrollable body
        body = CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=0)
        self.grid_rowconfigure(1, weight=1)
        body.grid_columnconfigure(0, weight=1)

        def sec_label(row, title):
            f = CTkFrame(body, fg_color="transparent")
            f.grid(row=row, column=0, sticky="ew", padx=20, pady=(14, 4))
            CTkLabel(f, text=title, font=F(size=12, weight="bold"),
                     text_color=CLR["blue"]).pack(anchor="w")
            CTkFrame(f, height=1, fg_color=CLR["border"]).pack(fill="x", pady=(2, 0))

        def make_grid(row):
            f = CTkFrame(body, fg_color="transparent")
            f.grid(row=row, column=0, sticky="ew", padx=20)
            f.grid_columnconfigure(1, weight=1)
            return f

        # ── ติดต่อ ────────────────────────────────────────────────────────────
        sec_label(0, "ข้อมูลการติดต่อ")
        cf = make_grid(1)
        self._fields = {}
        for ri, (key, lbl, ph) in enumerate([
            ("name",          "ชื่อบริษัท",      ""),
            ("supplier_id",   "รหัสซัพ",          ""),
            ("contact",       "ผู้ติดต่อ",         ""),
            ("phone",         "เบอร์ติดต่อ",     "0XX-XXX-XXXX"),
            ("line_id",       "Line ID",          "@"),
            ("email",         "Email",            "example@domain.com"),
            ("coverage_area", "พื้นที่จัดส่ง",   "เช่น กรุงเทพ, ชลบุรี"),
            ("wh_zone",       "โซนที่ตั้ง",       "เช่น กรุงเทพตะวันออก, ชลบุรี"),
        ]):
            CTkLabel(cf, text=lbl + ":", text_color=CLR["gray"],
                     font=F(size=12), width=110, anchor="w").grid(
                row=ri, column=0, sticky="w", pady=5, padx=(0, 10))
            e = CTkEntry(cf, font=F(size=13), height=34,
                         placeholder_text=ph)
            e.insert(0, str(supplier.get(key, "")))
            e.grid(row=ri, column=1, sticky="ew", pady=5)
            self._fields[key] = e

        # ── พิกัด Warehouse (row ต่อจาก wh_zone) ────────────────────────────
        coord_row = 8  # ต่อจาก wh_zone ที่ row=7 (name + supplier_id + 6 fields)
        CTkLabel(cf, text="พิกัด Warehouse:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=coord_row, column=0, sticky="w", pady=5, padx=(0, 10))

        coord_inner = CTkFrame(cf, fg_color="transparent")
        coord_inner.grid(row=coord_row, column=1, sticky="ew", pady=5)
        coord_inner.grid_columnconfigure(0, weight=1)

        coord_e = CTkEntry(coord_inner, font=F(size=13), height=34,
                           placeholder_text="Google Maps URL หรือ lat, lng (เช่น 13.7563, 100.5018)")
        coord_e.insert(0, str(supplier.get("wh_coordinates", "")))
        coord_e.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._fields["wh_coordinates"] = coord_e

        def _open_map():
            import webbrowser
            val = coord_e.get().strip()
            if not val:
                return
            if val.startswith("http"):
                webbrowser.open(val)
            else:
                # สมมติเป็น "lat, lng"
                query = val.replace(" ", "")
                webbrowser.open(f"https://maps.google.com/?q={query}")

        CTkButton(coord_inner, text="📍 เปิดแผนที่", width=100,
                  font=F(size=12), fg_color=CLR["teal"], hover_color="#047481",
                  height=34, command=_open_map).grid(row=0, column=1)

        # ── สถานะ & Tier ──────────────────────────────────────────────────────
        sec_label(2, "สถานะและ Tier")
        sf = make_grid(3)

        CTkLabel(sf, text="สถานะสต็อก:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=0, column=0, sticky="w", pady=5, padx=(0, 10))
        self._avail_var = tk.StringVar(value=supplier.get("availability", "พร้อม"))
        CTkOptionMenu(sf, variable=self._avail_var,
                      values=["พร้อม", "สต็อกต่ำ", "ปิดชั่วคราว"],
                      font=F(size=13),
                      command=self._on_avail_change).grid(row=0, column=1, sticky="w", pady=5)

        # reopen_date — แสดงเมื่อ ปิดชั่วคราว
        self._reopen_row = CTkFrame(sf, fg_color="transparent")
        self._reopen_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0,4))
        self._reopen_row.grid_columnconfigure(1, weight=1)
        CTkLabel(self._reopen_row, text="วันเปิดใหม่:", text_color=CLR["amber"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10))
        self._reopen_entry = CTkEntry(self._reopen_row, font=F(size=13), height=30,
                                      placeholder_text="เช่น 2025-05-01 (YYYY-MM-DD)")
        self._reopen_entry.grid(row=0, column=1, sticky="ew")
        if supplier.get("reopen_date"):
            self._reopen_entry.insert(0, supplier["reopen_date"])
        # toggle visibility
        is_closed = supplier.get("availability") == "ปิดชั่วคราว"
        if not is_closed:
            self._reopen_row.grid_remove()

        row_next = 1  # ปรับ row index ถัดไปสำหรับ Tier
        CTkLabel(sf, text="Tier:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=2, column=0, sticky="w", pady=5, padx=(0, 10))
        self._tier_var = tk.StringVar(value=supplier.get("tier", "SN"))
        CTkOptionMenu(sf, variable=self._tier_var,
                      values=["Tier 1", "Tier 2", "Tier 3", "SN", "Blacklist"],
                      font=F(size=13)).grid(row=2, column=1, sticky="w", pady=5)

        lock_row = CTkFrame(sf, fg_color="transparent")
        lock_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=5)
        self._lock_var = tk.BooleanVar(value=supplier.get("is_locked", False))
        CTkCheckBox(lock_row, text="Lock Tier 1 ถาวร (Manager Override — ป้องกัน Auto-Demote)",
                    variable=self._lock_var, font=F(size=12)).pack(side="left")

        # ── Weighted Score ─────────────────────────────────────────────────────
        sec_label(4, "Weighted Score  (แต่ละเกณฑ์ 20%  รวม 5 ด้าน)")
        sc_wrap = CTkFrame(body, fg_color=CLR["gray_lt"], corner_radius=8)
        sc_wrap.grid(row=5, column=0, sticky="ew", padx=20, pady=4)
        sc_wrap.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        # แถว auto-calculated (ราคา, Win%, SLA)
        auto_items = [
            ("ราคา\n(20%)",       supplier.get("price_score", 0), CLR["blue"]),
            ("สต็อก/Win\n(20%)",  supplier.get("win_pct",     0), CLR["teal"]),
            ("SLA\n(20%)",        supplier.get("sla_score",   0), "#7C3AED"),
        ]
        for col, (lbl, val, color) in enumerate(auto_items):
            t = CTkFrame(sc_wrap, fg_color=CLR["white"], corner_radius=8,
                         border_width=1, border_color=CLR["border"])
            t.grid(row=0, column=col, padx=5, pady=(8, 2), sticky="ew")
            CTkLabel(t, text=str(val), font=F(size=20, weight="bold"),
                     text_color=color).pack(pady=(6, 0))
            CTkLabel(t, text=lbl, font=F(size=10), text_color=CLR["gray"],
                     justify="center").pack(pady=(0, 4))
            CTkLabel(t, text="(อัตโนมัติ)", font=F(size=9),
                     text_color=CLR["gray"]).pack(pady=(0, 6))

        # คะแนนรวม (colspan 2 ด้านขวา ใช้ 2 cols เดียวกัน)
        total = calc_score(supplier)
        tr = CTkFrame(sc_wrap, fg_color=CLR["navy"], corner_radius=8)
        tr.grid(row=0, column=3, columnspan=2, padx=5, pady=(8, 2), sticky="ew")
        CTkLabel(tr, text=str(total), font=F(size=28, weight="bold"),
                 text_color=CLR["white"]).pack(pady=(8, 0))
        CTkLabel(tr, text="คะแนนรวม\n(เต็ม 100)", font=F(size=10),
                 text_color="#93C5FD", justify="center").pack(pady=(0, 8))

        # แถว manual input — บริการ + คุณภาพ
        manual_row = CTkFrame(sc_wrap, fg_color="transparent")
        manual_row.grid(row=1, column=0, columnspan=5, sticky="ew", padx=6, pady=(0, 8))
        manual_row.grid_columnconfigure((1, 4), weight=1)

        CTkLabel(manual_row, text="บริการ (20%):", font=F(size=12),
                 text_color=CLR["gray"]).grid(row=0, column=0, padx=(4, 6), pady=4)
        self._service_score_var = tk.StringVar(
            value=str(supplier.get("service_score", 0)))
        CTkEntry(manual_row, textvariable=self._service_score_var,
                 width=70, font=F(size=13), justify="center").grid(
            row=0, column=1, sticky="w", pady=4)
        CTkLabel(manual_row, text="/ 100", font=F(size=11),
                 text_color=CLR["gray"]).grid(row=0, column=2, padx=(4, 20), pady=4)

        _q_raw  = supplier.get("quality_score")
        q_score = int(_q_raw) if _q_raw is not None else QUALITY_START
        CTkLabel(manual_row, text="คุณภาพ (20%):", font=F(size=12),
                 text_color=CLR["gray"]).grid(row=0, column=3, padx=(4, 6), pady=4)
        self._quality_score_lbl = CTkLabel(
            manual_row, text=str(q_score), font=F(size=14, weight="bold"),
            text_color=CLR["green"] if q_score >= 75 else (CLR["amber"] if q_score >= 50 else CLR["red"]))
        self._quality_score_lbl.grid(row=0, column=4, sticky="w", pady=4)
        CTkLabel(manual_row, text="/ 100  (อัตโนมัติ)", font=F(size=11),
                 text_color=CLR["gray"]).grid(row=0, column=5, padx=(4, 4), pady=4)

        # ── Quality Events (row 6-7: ทันทีหลัง Weighted Score row5) ──────────
        sec_label(6, "คุณภาพสินค้า — Event History")
        qe_wrap = CTkFrame(body, fg_color=CLR["white"], corner_radius=8,
                           border_width=1, border_color=CLR["border"])
        qe_wrap.grid(row=7, column=0, sticky="ew", padx=20, pady=(0, 6))
        qe_wrap.grid_columnconfigure(0, weight=1)

        self._qe_frame = qe_wrap
        self._supplier_for_qe = supplier

        def _reload_quality_score():
            """อ่าน quality_score จาก DB จริง แล้วอัปเดต label"""
            import psycopg2 as _pg2
            conn = _pg2.connect(**_DB_CFG)
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT quality_score FROM suppliers WHERE supplier_name = %s",
                                (self._supplier_for_qe["name"],))
                    row = cur.fetchone()
                    q = int(row[0] or 0) if row else 0
                    self._quality_score_lbl.configure(
                        text=str(q),
                        text_color=CLR["green"] if q >= 75 else (
                            CLR["amber"] if q >= 50 else CLR["red"]))
            except Exception as e:
                print(f"[SSL] _reload_quality_score error: {e}")
            finally:
                conn.close()

        def _refresh_qe():
            for w in self._qe_frame.winfo_children():
                w.destroy()
            events = db_get_quality_events(self._supplier_for_qe["name"])
            if not events:
                CTkLabel(self._qe_frame,
                         text="ยังไม่มีเหตุการณ์  (คะแนนเต็ม 100/100)",
                         text_color=CLR["gray"], font=F(size=12)).pack(
                    padx=14, pady=8, anchor="w")
            else:
                for ev in events:
                    row_f = CTkFrame(self._qe_frame, fg_color="transparent")
                    row_f.pack(fill="x", padx=10, pady=3)
                    is_claim    = ev["event_type"] == "claim"
                    is_resolved = ev["resolved"]
                    dot_color   = CLR["red"] if is_claim else CLR["green"]
                    dot_txt     = "● เคลม" if is_claim else "  ↩ recovery"
                    CTkLabel(row_f, text=dot_txt, font=F(size=11, weight="bold"),
                             text_color=dot_color, width=80).pack(side="left")
                    detail = ev.get("reason") or ev.get("recovery_label") or ""
                    delta_txt = f"{ev['delta']:+d}"
                    CTkLabel(row_f, text=f"{ev['ts']}  {delta_txt}  {detail}",
                             font=F(size=11), text_color=CLR["gray"]).pack(
                        side="left", padx=(6, 0))
                    if is_claim and not is_resolved:
                        def _open_recovery(eid=ev["id"]):
                            _show_recovery_popup(eid)
                        CTkButton(row_f, text="แก้ปัญหาแล้ว",
                                  fg_color=CLR["blue"], hover_color="#1e40af",
                                  height=24, width=100, font=F(size=11),
                                  command=_open_recovery).pack(
                            side="right", padx=(6, 0))

            # ปุ่มแจ้งปัญหา
            btn_row = CTkFrame(self._qe_frame, fg_color="transparent")
            btn_row.pack(fill="x", padx=10, pady=(4, 8))
            CTkButton(btn_row, text="+ แจ้งปัญหา / เคลม",
                      fg_color=CLR["red"], hover_color="#991b1b",
                      height=28, font=F(size=12),
                      command=_open_claim_popup).pack(side="left")

        def _open_claim_popup():
            pop = CTkToplevel(self)
            pop.title("แจ้งปัญหา / เคลม")
            _place_popup(pop, 400, 200)
            pop.resizable(False, False)
            pop.transient(self)
            pop.grab_set()
            pop.grid_columnconfigure(0, weight=1)
            CTkLabel(pop, text="เหตุผล / รายละเอียดปัญหา:",
                     font=F(size=12), text_color=CLR["gray"]).grid(
                row=0, column=0, padx=20, pady=(16, 4), sticky="w")
            reason_e = CTkEntry(pop, font=F(size=13), height=34,
                                placeholder_text="เช่น สเปกไม่ตรง, ของแตก, ส่งผิดรุ่น...")
            reason_e.grid(row=1, column=0, padx=20, sticky="ew")
            def _confirm_claim():
                reason = reason_e.get().strip() or "ไม่ระบุเหตุผล"
                eid = db_add_quality_event(
                    self._supplier_for_qe["name"], reason, self.current_user)
                if eid > 0:
                    _reload_quality_score()
                    _refresh_qe()
                    if self.on_save:
                        self.on_save(self.sup)
                pop.destroy()
            CTkButton(pop, text="ยืนยัน หัก -50 คะแนน",
                      fg_color=CLR["red"], hover_color="#991b1b",
                      command=_confirm_claim).grid(
                row=2, column=0, padx=20, pady=14, sticky="e")

        def _show_recovery_popup(event_id: int):
            pop = CTkToplevel(self)
            pop.title("ประเมินการแก้ปัญหา")
            _place_popup(pop, 400, 230)
            pop.resizable(False, False)
            pop.transient(self)
            pop.grab_set()
            pop.grid_columnconfigure(0, weight=1)
            CTkLabel(pop, text="Supplier แก้ปัญหาอย่างไร?",
                     font=F(size=13, weight="bold")).grid(
                row=0, column=0, padx=20, pady=(16, 12))
            options = [
                ("แก้ปัญหาเยี่ยม / รวดเร็ว",   "great",  CLR["green"],  f"+{QUALITY_RECOVER_GREAT}"),
                ("แก้ปัญหาปกติ / ใช้เวลา",      "normal", CLR["blue"],   f"+{QUALITY_RECOVER_NORMAL}"),
                ("ทิ้งงาน / บ่ายเบี่ยง",        "bad",    CLR["red"],    f"{QUALITY_RECOVER_BAD}"),
            ]
            for ri, (label, rtype, color, delta_txt) in enumerate(options, 1):
                def _pick(rt=rtype):
                    ok = db_resolve_quality_event(
                        event_id, rt,
                        self._supplier_for_qe["name"], self.current_user)
                    if ok:
                        _reload_quality_score()
                        _refresh_qe()
                        if self.on_save:
                            self.on_save(self.sup)
                    pop.destroy()
                CTkButton(pop, text=f"{label}  ({delta_txt})",
                          fg_color=color, hover_color="#374151",
                          height=30, font=F(size=12),
                          command=_pick).grid(
                    row=ri, column=0, padx=20, pady=4, sticky="ew")

        _refresh_qe()
        _reload_quality_score()   # sync label กับ DB จริงทันทีที่เปิด popup

        # ── Win-Loss Log ───────────────────────────────────────────────────────
        sec_label(8, "Win-Loss Log")
        wl_f = make_grid(9)
        logs = supplier.get("win_loss_log", [])
        if logs:
            for li, entry in enumerate(logs[-3:]):
                dot = CLR["green"] if entry["result"] == "Win" else CLR["red"]
                CTkLabel(wl_f, text="●", text_color=dot,
                         font=F(size=13)).grid(row=li, column=0, padx=(0, 8), sticky="w")
                txt = f"{entry['date']}  {entry['result']}"
                if entry.get("reason"):
                    txt += f" — {entry['reason']}"
                CTkLabel(wl_f, text=txt, font=F(size=12)).grid(
                    row=li, column=1, sticky="w", pady=2)
        else:
            CTkLabel(wl_f, text="ยังไม่มีประวัติ", text_color=CLR["gray"],
                     font=F(size=12)).grid(row=0, column=0, columnspan=2, sticky="w")

        # เพิ่ม log
        add_f = CTkFrame(body, fg_color=CLR["gray_lt"], corner_radius=8)
        add_f.grid(row=10, column=0, sticky="ew", padx=20, pady=(4, 0))
        add_f.grid_columnconfigure(1, weight=1)
        CTkLabel(add_f, text="บันทึกผล:", font=F(size=12),
                 text_color=CLR["gray"]).grid(row=0, column=0, padx=(10, 6), pady=8)
        self._wl_touched = False
        self._wl_result  = tk.StringVar(value="Win")
        self._wl_reason  = tk.StringVar(value=WIN_LOSS_REASONS[0])

        def _on_result_change(val):
            self._wl_touched = True
            # ซ่อน/แสดง reason dropdown ตาม result
            if val == "Loss":
                self._reason_opt.grid(row=0, column=2, padx=6, pady=8)
            else:
                self._wl_reason.set(WIN_LOSS_REASONS[0])
                self._reason_opt.grid_remove()

        CTkOptionMenu(add_f, variable=self._wl_result, values=["Win", "Loss"],
                      width=80, font=F(size=12),
                      command=_on_result_change).grid(row=0, column=1, sticky="w", pady=8)
        self._reason_opt = CTkOptionMenu(add_f, variable=self._wl_reason,
                                          values=WIN_LOSS_REASONS, width=200,
                                          font=F(size=12))
        # ซ่อน reason ตอนแรก (default = Win)
        self._reason_opt.grid_remove()

        # ── Note ──────────────────────────────────────────────────────────────
        # ── Tier Transition History ────────────────────────────────────────────
        # ดึง log ของ Supplier นี้จาก MOCK_AUDIT_LOG
        sup_id   = supplier.get("supplier_id", "")
        _audit   = db_get_audit_log()
        tier_logs = [
            l for l in _audit
            if sup_id in l.get("detail", "") and
               any(kw in l.get("action", "") for kw in
                   ("tier", "blacklist", "convert", "edit"))
        ]
        if tier_logs:
            sec_label(11, "ประวัติการเปลี่ยน Tier")
            th_f = CTkFrame(body, fg_color=CLR["white"], corner_radius=8,
                            border_width=1, border_color=CLR["border"])
            th_f.grid(row=12, column=0, sticky="ew", padx=20, pady=(0, 4))
            for li, entry in enumerate(reversed(tier_logs[-5:])):
                rf2 = CTkFrame(th_f, fg_color="transparent")
                rf2.pack(fill="x", padx=12, pady=3)
                action = entry.get("action", "")
                dot_c  = (CLR["green"]  if "Tier 1" in action else
                          CLR["blue"]   if "Tier 2" in action else
                          CLR["red"]    if "Blacklist" in action else
                          CLR["amber"]  if "SN→SW" in action else CLR["gray"])
                CTkLabel(rf2, text="●", text_color=dot_c,
                         font=F(size=12), width=14).pack(side="left")
                detail_txt = entry.get("detail", "")
                remark_part = ""
                if "Remark:" in detail_txt:
                    remark_part = "  (" + detail_txt.split("Remark:")[-1].strip() + ")"
                txt = f"{entry.get('timestamp','')}  {action}  โดย {entry.get('user','')}{remark_part}"
                CTkLabel(rf2, text=txt, font=F(size=11),
                         text_color=CLR["gray"]).pack(side="left", padx=(4, 0))

        # ── Zoning Phase 1 ────────────────────────────────────────────────────
        # Section header
        _z_hdr = CTkFrame(body, fg_color="transparent")
        _z_hdr.grid(row=13, column=0, sticky="ew", padx=20, pady=(14, 4))
        CTkLabel(_z_hdr, text="Zoning & การขนส่ง",
                 font=F(size=12, weight="bold"),
                 text_color=CLR["blue"]).pack(anchor="w")
        CTkFrame(_z_hdr, height=1, fg_color=CLR["border"]).pack(fill="x", pady=(2, 0))

        DISPATCH_ZONES = [
            "— ยังไม่ระบุ —",
            "1. กทม. / ปริมณฑล",
            "2. โซนตะวันออก (ชลบุรี / ระยอง / สมุทรปราการ)",
            "3. โซนวังน้อย / อยุธยา / สระบุรี",
            "4. โซนพระราม 2 / พุทธมณฑล / สมุทรสาคร",
            "5. ภาคเหนือ",
            "6. ภาคอีสาน",
            "7. ภาคใต้",
            "8. ภาคตะวันตก",
            "9. อื่นๆ",
        ]
        SERVICE_AREAS  = ["National (ทั่วประเทศ)", "Regional (เฉพาะภาค)", "Local (เฉพาะจังหวัด)"]
        LOGISTICS_OPTS = ["กระบะ", "6 ล้อ", "10 ล้อ", "เทรลเลอร์"]

        zf = CTkFrame(body, fg_color="transparent")
        zf.grid(row=14, column=0, sticky="ew", padx=20)
        zf.grid_columnconfigure(1, weight=1)

        # โซนที่ตั้ง
        CTkLabel(zf, text="โซนที่ตั้ง:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=0, column=0, sticky="w", pady=5, padx=(0, 10))
        self._zone_var = tk.StringVar(value=supplier.get("dispatch_zone") or "— ยังไม่ระบุ —")
        CTkOptionMenu(zf, variable=self._zone_var, values=DISPATCH_ZONES,
                      font=F(size=13), width=220).grid(row=0, column=1, sticky="w", pady=5)

        # ขอบเขตการส่ง
        CTkLabel(zf, text="ขอบเขตส่ง:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=1, column=0, sticky="w", pady=5, padx=(0, 10))
        self._service_var = tk.StringVar(value=supplier.get("service_area") or SERVICE_AREAS[0])
        CTkOptionMenu(zf, variable=self._service_var, values=SERVICE_AREAS,
                      font=F(size=13), width=220).grid(row=1, column=1, sticky="w", pady=5)

        # ประเภทรถขนส่ง (multi-select checkboxes)
        CTkLabel(zf, text="ประเภทรถ:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=2, column=0, sticky="nw", pady=8, padx=(0, 10))
        truck_frame = CTkFrame(zf, fg_color="transparent")
        truck_frame.grid(row=2, column=1, sticky="w", pady=5)
        saved_trucks = supplier.get("logistics_assets", "") or ""
        saved_set    = {t.strip() for t in saved_trucks.split(",") if t.strip()}
        self._truck_vars = {}
        for i, opt in enumerate(LOGISTICS_OPTS):
            var = tk.BooleanVar(value=(opt in saved_set))
            self._truck_vars[opt] = var
            CTkCheckBox(truck_frame, text=opt, variable=var,
                        font=F(size=12)).grid(row=0, column=i, padx=(0, 12))

        # ── ประเภทธุรกิจ ──────────────────────────────────────────────────────
        CTkLabel(zf, text="ประเภทธุรกิจ:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=3, column=0, sticky="w", pady=5, padx=(0, 10))
        self._biz_var = tk.StringVar(value=supplier.get("business_type") or "— ยังไม่ระบุ —")
        CTkOptionMenu(zf, variable=self._biz_var,
                      values=["— ยังไม่ระบุ —", "โรงงานผลิต / ผู้นำเข้า", "ตัวแทนจำหน่าย / ร้านค้าใหญ่", "ร้านค้าทั่วไป", "Modern Trade"],
                      font=F(size=13), width=220).grid(row=3, column=1, sticky="w", pady=5)

        # ── เครดิต ────────────────────────────────────────────────────────────
        CTkLabel(zf, text="เครดิต:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=4, column=0, sticky="w", pady=5, padx=(0, 10))
        self._credit_lbl_var = tk.StringVar(value=supplier.get("credit_term_label") or "สด")
        CTkOptionMenu(zf, variable=self._credit_lbl_var,
                      values=["สด", "เครดิต 2D", "เครดิต 3D", "เครดิต 7D",
                               "เครดิต 15D", "เครดิต 30D", "เครดิต 45D", "เครดิต 60D"],
                      font=F(size=13), width=220).grid(row=4, column=1, sticky="w", pady=5)

        # ── มาตรฐานสินค้า ─────────────────────────────────────────────────────
        CTkLabel(zf, text="มาตรฐาน:", text_color=CLR["gray"],
                 font=F(size=12), width=110, anchor="w").grid(
            row=5, column=0, sticky="w", pady=5, padx=(0, 10))
        self._std_var = tk.StringVar(value=supplier.get("standard_focus") or "— ยังไม่ระบุ —")
        CTkOptionMenu(zf, variable=self._std_var,
                      values=["— ยังไม่ระบุ —", "เกรดราชการ / มอก.", "เกรดทั่วไป", "ทั้งสองประเภท"],
                      font=F(size=13), width=220).grid(row=5, column=1, sticky="w", pady=5)

        sec_label(15, "จุดแข็ง / หมายเหตุ")
        self._note_e = CTkEntry(body, font=F(size=13), height=36,
                                placeholder_text="เช่น ให้เครดิต 60 วัน, ส่งด่วน, ISO ผ่าน...")
        self._note_e.grid(row=16, column=0, sticky="ew", padx=20, pady=4)
        if supplier.get("note"):
            self._note_e.insert(0, supplier["note"])

        # SN → SW Conversion section removed

        # Bottom bar
        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ยกเลิก", fg_color="gray50", hover_color="gray40",
                  width=90, command=self.destroy).pack(side="right", padx=12, pady=10)
        CTkButton(bf, text="บันทึก", fg_color=CLR["blue"], hover_color="#1e40af",
                  width=120, command=self._save).pack(side="right", pady=10)

        self.transient(master)
        self.grab_set()

    def _open_convert(self):
        ConvertToSWPopup(self, self.sup, on_success=self._after_convert,
                         current_user=self.current_user)

    def _on_avail_change(self, val):
        if val == "ปิดชั่วคราว":
            self._reopen_row.grid()
        else:
            self._reopen_row.grid_remove()
            self._reopen_entry.delete(0, "end")

    def _after_convert(self):
        if self.on_save:
            self.on_save(self.sup)
        self.destroy()

    def _save(self):
        # ตรวจสอบชื่อบริษัทก่อน save
        new_name = self._fields.get("name")
        new_name = new_name.get().strip() if new_name else ""
        if not new_name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาระบุชื่อบริษัท", parent=self)
            return

        # ตรวจสอบรหัสซัพก่อน save
        new_code = self._fields.get("supplier_id")
        new_code = new_code.get().strip() if new_code else ""
        old_code = self.sup.get("supplier_id", "")
        if not new_code:
            messagebox.showwarning("ข้อมูลไม่ครบ", "รหัสซัพไม่สามารถว่างได้", parent=self)
            return
        if new_code != old_code:
            # เช็คซ้ำกับรหัสอื่น
            conn_chk, use_pool_chk = _get_conn()
            try:
                with conn_chk.cursor() as _cur:
                    _cur.execute(
                        "SELECT 1 FROM suppliers WHERE supplier_code = %s AND id != %s",
                        (new_code, self.sup.get("id", 0))
                    )
                    if _cur.fetchone():
                        messagebox.showerror("รหัสซ้ำ", f"รหัส '{new_code}' มีอยู่แล้วในระบบ", parent=self)
                        return
            finally:
                _release_conn(conn_chk, use_pool_chk)

        for key, widget in self._fields.items():
            self.sup[key] = widget.get().strip()

        def _clamp_score(var_name):
            try:
                v = int(getattr(self, var_name).get())
                return max(0, min(100, v))
            except Exception:
                return 0

        self.sup.update({
            "availability": self._avail_var.get(),
            "reopen_date":  self._reopen_entry.get().strip(),
            "tier":         self._tier_var.get(),
            "is_locked":    self._lock_var.get(),
            "note":         self._note_e.get().strip(),
            # ── Manual scores ────────────────────────────────────────────
            "service_score": _clamp_score("_service_score_var"),
            # quality_score ไม่ save จาก manual — อัปเดตผ่าน event เท่านั้น
            # ── Zoning ──────────────────────────────────────────────────
            "dispatch_zone":    self._zone_var.get() if hasattr(self, "_zone_var") else "",
            "service_area":     self._service_var.get() if hasattr(self, "_service_var") else "National",
            "logistics_assets": ",".join(
                opt for opt, var in self._truck_vars.items() if var.get()
            ) if hasattr(self, "_truck_vars") else "",
            "business_type":    self._biz_var.get() if hasattr(self, "_biz_var") else "",
            "standard_focus":   self._std_var.get() if hasattr(self, "_std_var") else "",
            "credit_term_label": self._credit_lbl_var.get() if hasattr(self, "_credit_lbl_var") else "สด",
            "wh_zone":           self._wh_zone_e.get().strip() if hasattr(self, "_wh_zone_e") else "",
            "wh_coordinates":    self._wh_coord_e.get().strip() if hasattr(self, "_wh_coord_e") else "",
        })
        result = self._wl_result.get()
        reason = self._wl_reason.get() if result == "Loss" else ""
        if hasattr(self, "_wl_touched") and self._wl_touched:
            self.sup.setdefault("win_loss_log", []).append({
                "date":   datetime.now().strftime("%Y-%m"),
                "result": result,
                "reason": reason,
            })
        ok = db_save_supplier(self.sup, action="edit", user=self.current_user)
        if ok:
            messagebox.showinfo("บันทึกสำเร็จ",
                                f"อัปเดตข้อมูล '{self.sup['name']}' เรียบร้อยแล้ว",
                                parent=self)
        else:
            messagebox.showerror("ผิดพลาด", "บันทึกลงฐานข้อมูลไม่สำเร็จ", parent=self)
        if self.on_save:
            self.on_save(self.sup)
        self.destroy()

# =============================================================================
#  ADD NEW SN POPUP
# =============================================================================
class AddSupplierPopup(CTkToplevel):

    def __init__(self, master, on_success=None, current_user="USER_DEMO"):
        super().__init__(master)
        self.on_success   = on_success
        self.current_user = current_user
        self.title("เพิ่ม Supplier ใหม่ (SN)")
        _place_popup(self, 640, 800)
        self.resizable(False, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._debounce_job = None
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        CTkLabel(hdr, text="เพิ่ม Supplier ใหม่ (รหัสชั่วคราว SN)",
                 font=F(size=15, weight="bold"),
                 text_color=CLR["white"]).pack(padx=20, pady=12, anchor="w")

        # ── Scrollable body ───────────────────────────────────────────────────
        scroll_body = CTkScrollableFrame(self, fg_color="transparent")
        scroll_body.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        scroll_body.grid_columnconfigure(0, weight=1)

        form = CTkFrame(scroll_body, fg_color="transparent")
        form.grid(row=0, column=0, padx=20, pady=12, sticky="ew")
        form.grid_columnconfigure(1, weight=1)

        fields_cfg = [
            ("supplier_id",   "รหัสซัพพลายเออร์", "เว้นว่างเพื่อสร้าง SN อัตโนมัติ", False), # <--- เพิ่มบรรทัดนี้
            ("name",          "ชื่อบริษัท *",   "พิมพ์ชื่อ Supplier...",  True),
            ("category",      "หมวดสินค้า",     None,                      False),
            ("contact",       "ผู้ติดต่อ",       "",                       False),
            ("phone",         "เบอร์โทร *",      "0XX-XXX-XXXX",          True),
            ("line_id",       "Line ID",          "@",                      False),
            ("email",         "Email",            "example@domain.com",    False),
            ("coverage_area", "พื้นที่จัดส่ง",   "เช่น กรุงเทพ, ชลบุรี", False),
        ]
        self._inputs = {}
        fuzzy_row_idx = None
        # ดึง category จาก DB จริง
        _live_cats = db_get_categories()[1:]  # ข้าม "ทุกหมวด"
        if not _live_cats:
            _live_cats = ["(ยังไม่มีหมวด)"]
        for ri, (key, lbl, ph, _) in enumerate(fields_cfg):
            CTkLabel(form, text=lbl + ":", text_color=CLR["gray"],
                     font=F(size=12), width=110, anchor="w").grid(
                row=ri * 2, column=0, sticky="w", pady=(6, 0), padx=(0, 10))
            if key == "category":
                cat_row = CTkFrame(form, fg_color="transparent")
                cat_row.grid(row=ri * 2, column=1, sticky="w", pady=(6, 0))
                self._cat_var = tk.StringVar(value=_live_cats[0])
                w = CTkOptionMenu(cat_row, variable=self._cat_var,
                                  values=_live_cats, font=F(size=13))
                w.pack(side="left")
                CTkButton(cat_row, text="➕ ขอเพิ่มหมวดใหม่", width=140, height=28,
                          font=F(size=12), fg_color=CLR["gray"], hover_color=CLR["blue"],
                          command=self._show_add_category_sop).pack(side="left", padx=(8, 0))
                self._inputs[key] = self._cat_var
            else:
                e = CTkEntry(form, font=F(size=13), height=34,
                             placeholder_text=ph or "")
                e.grid(row=ri * 2, column=1, sticky="ew", pady=(6, 0))
                self._inputs[key] = e
                if key == "name":
                    e.bind("<KeyRelease>", self._on_name_key)
                    fuzzy_row_idx = ri * 2 + 1

        # Fuzzy warning (ซ่อนไว้ก่อน)
        self._fuzzy_f = CTkFrame(form, fg_color=CLR["amber_lt"],
                                  corner_radius=6, border_width=1,
                                  border_color="#F59E0B")
        self._fuzzy_visible = False

        # ── Zoning fields ─────────────────────────────────────────────────────
        DISPATCH_ZONES = [
            "— ยังไม่ระบุ —",
            "1. กทม. / ปริมณฑล",
            "2. โซนตะวันออก (ชลบุรี / ระยอง / สมุทรปราการ)",
            "3. โซนวังน้อย / อยุธยา / สระบุรี",
            "4. โซนพระราม 2 / พุทธมณฑล / สมุทรสาคร",
            "5. ภาคเหนือ",
            "6. ภาคอีสาน",
            "7. ภาคใต้",
            "8. ภาคตะวันตก",
            "9. อื่นๆ",
        ]
        SERVICE_AREAS  = ["National (ทั่วประเทศ)", "Regional (เฉพาะภาค)", "Local (เฉพาะจังหวัด)"]
        LOGISTICS_OPTS = ["กระบะ", "6 ล้อ", "10 ล้อ", "เทรลเลอร์"]

        zone_section = CTkFrame(scroll_body, fg_color=CLR["gray_lt"], corner_radius=8,
                                border_width=1, border_color=CLR["border"])
        zone_section.grid(row=1, column=0, padx=20, pady=(4, 4), sticky="ew")
        zone_section.grid_columnconfigure(1, weight=1)

        CTkLabel(zone_section, text="Zoning & การขนส่ง",
                 font=F(size=12, weight="bold"),
                 text_color=CLR["blue"]).grid(
            row=0, column=0, columnspan=2, padx=12, pady=(8, 4), sticky="w")

        CTkLabel(zone_section, text="โซนที่ตั้ง:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=1, column=0, padx=(12, 8), pady=4, sticky="w")
        self._zone_var = tk.StringVar(value="— ยังไม่ระบุ —")
        CTkOptionMenu(zone_section, variable=self._zone_var, values=DISPATCH_ZONES,
                      font=F(size=12), width=200).grid(row=1, column=1, sticky="w", pady=4, padx=(0, 12))

        CTkLabel(zone_section, text="ขอบเขตส่ง:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=2, column=0, padx=(12, 8), pady=4, sticky="w")
        self._service_var = tk.StringVar(value=SERVICE_AREAS[0])
        CTkOptionMenu(zone_section, variable=self._service_var, values=SERVICE_AREAS,
                      font=F(size=12), width=200).grid(row=2, column=1, sticky="w", pady=4, padx=(0, 12))

        CTkLabel(zone_section, text="ประเภทรถ:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=3, column=0, padx=(12, 8), pady=(4, 8), sticky="nw")
        truck_frame = CTkFrame(zone_section, fg_color="transparent")
        truck_frame.grid(row=3, column=1, sticky="w", pady=(4, 8))
        self._truck_vars = {}
        for i, opt in enumerate(LOGISTICS_OPTS):
            var = tk.BooleanVar(value=False)
            self._truck_vars[opt] = var
            CTkCheckBox(truck_frame, text=opt, variable=var,
                        font=F(size=12)).grid(row=0, column=i, padx=(0, 10))

        # ประเภทธุรกิจ
        CTkLabel(zone_section, text="ประเภทธุรกิจ:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=4, column=0, padx=(12, 8), pady=4, sticky="w")
        self._biz_var = tk.StringVar(value="— ยังไม่ระบุ —")
        CTkOptionMenu(zone_section, variable=self._biz_var,
                      values=["— ยังไม่ระบุ —", "โรงงานผลิต / ผู้นำเข้า", "ตัวแทนจำหน่าย / ร้านค้าใหญ่", "ร้านค้าทั่วไป", "Modern Trade"],
                      font=F(size=12), width=200).grid(row=4, column=1, sticky="w", pady=4, padx=(0, 12))

        # เครดิต
        CTkLabel(zone_section, text="เครดิต:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=5, column=0, padx=(12, 8), pady=4, sticky="w")
        self._credit_lbl_var = tk.StringVar(value="สด")
        CTkOptionMenu(zone_section, variable=self._credit_lbl_var,
                      values=["สด", "เครดิต 2D", "เครดิต 3D", "เครดิต 7D",
                               "เครดิต 15D", "เครดิต 30D", "เครดิต 45D", "เครดิต 60D"],
                      font=F(size=12), width=200).grid(row=5, column=1, sticky="w", pady=4, padx=(0, 12))

        # มาตรฐานสินค้า
        CTkLabel(zone_section, text="มาตรฐาน:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=6, column=0, padx=(12, 8), pady=(4, 4), sticky="w")
        self._std_var = tk.StringVar(value="— ยังไม่ระบุ —")
        CTkOptionMenu(zone_section, variable=self._std_var,
                      values=["— ยังไม่ระบุ —", "เกรดราชการ / มอก.", "เกรดทั่วไป", "ทั้งสองประเภท"],
                      font=F(size=12), width=200).grid(row=6, column=1, sticky="w", pady=(4, 4), padx=(0, 12))

        # ── WH Location ─────────────────────────────────────────────────────
        CTkLabel(zone_section, text="โซนที่ตั้ง WH:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=7, column=0, padx=(12, 8), pady=(4, 4), sticky="w")
        self._wh_zone_e = CTkEntry(zone_section, font=F(size=13), height=32,
                                    placeholder_text="เช่น กรุงเทพตะวันออก, ชลบุรี")
        self._wh_zone_e.grid(row=7, column=1, sticky="ew", pady=(4, 4), padx=(0, 12))

        CTkLabel(zone_section, text="พิกัด WH:", text_color=CLR["gray"],
                 font=F(size=12), width=90, anchor="w").grid(
            row=8, column=0, padx=(12, 8), pady=(4, 10), sticky="w")
        _coord_row = CTkFrame(zone_section, fg_color="transparent")
        _coord_row.grid(row=8, column=1, sticky="ew", pady=(4, 10), padx=(0, 12))
        _coord_row.grid_columnconfigure(0, weight=1)
        self._wh_coord_e = CTkEntry(_coord_row, font=F(size=13), height=32,
                                     placeholder_text="https://maps.google.com/... หรือ lat, lng")
        self._wh_coord_e.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        def _open_map_sn():
            import webbrowser
            val = self._wh_coord_e.get().strip()
            if not val:
                return
            if val.startswith("http"):
                webbrowser.open(val)
            else:
                # รองรับ "lat, lng" เช่น 13.6428, 100.3898
                query = val.replace(" ", "")
                webbrowser.open(f"https://maps.google.com/?q={query}")
        CTkButton(_coord_row, text="📍 เปิดแผนที่", width=100, height=32,
                  fg_color="#0EA5E9", hover_color="#0284C7",
                  font=F(size=11),
                  command=_open_map_sn).grid(row=0, column=1)

        # SN preview — อยู่ใน scroll_body
        self._sn_lbl = CTkLabel(scroll_body, text="", font=F(size=12, weight="bold"),
                                 text_color=CLR["blue"])
        self._sn_lbl.grid(row=2, column=0, padx=20, pady=(4, 8), sticky="w")
        self._refresh_sn_preview()

        # Bottom bar — ติดขอบล่างหน้าต่างเสมอ
        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ยกเลิก", fg_color="gray50", hover_color="gray40",
                  width=90, command=self.destroy).pack(side="right", padx=12, pady=10)
        CTkButton(bf, text="บันทึกและสร้างรหัส SN",
                  fg_color=CLR["navy"], hover_color=CLR["blue"], width=200,
                  command=self._save).pack(side="right", pady=10)

        self.transient(master)
        self.grab_set()

    def _show_add_category_sop(self):
        """แจ้งขั้นตอน SOP การขอเพิ่มหมวดสินค้าใหม่ (ไม่มีการอนุมัติในระบบ — ทำนอกระบบตาม SOP)"""
        win = CTkToplevel(self)
        win.title("ขอเพิ่มหมวดสินค้าใหม่")
        win.transient(self)
        _place_popup(win, 460, 420)
        win.grab_set()

        CTkLabel(win, text="ขั้นตอนการขอเพิ่มหมวดสินค้าใหม่",
                 font=CTkFont(size=15, weight="bold"),
                 text_color=CLR["navy"]).pack(padx=20, pady=(16, 4), anchor="w")
        CTkLabel(win, text="กรุณาติดต่อ ผู้จัดการฝ่ายขาย / ผช.ผู้จัดการฝ่ายขาย",
                 font=CTkFont(size=13, weight="bold"),
                 text_color=CLR["blue"]).pack(padx=20, pady=(0, 10), anchor="w")

        steps = [
            "1. ผู้จัดการฝ่ายขาย / ผช.ผู้จัดการฝ่ายขาย พิจารณาตามเห็นควร "
            "เช่น เพื่อขอเพิ่ม code หมวดใหม่ หรือพิจารณาแล้วให้ใช้หมวดเดิมที่สอดคล้องกัน",
            "2. ประสานงานปิ้น + บัญชี เพิ่มหมวดใน Express จริง",
            "3. หากต้องเพิ่มหมวดใหม่ ทาง ผจก. จะอนุมัติและแจ้งทาง DEV เพื่อเพิ่มหมวดใน "
            "A+ Smart เป็นลายลักษณ์อักษร ผ่าน LINE กลุ่ม 19.2",
            "4. ส่วนนี้ไม่ใช่ required field มาอัพเดตย้อนหลังได้ และ ผจก./ผช.ผจก. "
            "หรือเจ้าหน้าที่ทำเองย้อนหลังได้",
        ]
        body = CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        for s in steps:
            CTkLabel(body, text=s, font=CTkFont(size=12), justify="left",
                     wraplength=410, anchor="w").pack(fill="x", pady=(0, 10), anchor="w")

        CTkButton(win, text="รับทราบ", width=120, fg_color=CLR["navy"],
                  hover_color=CLR["blue"], command=win.destroy).pack(pady=(0, 16))

    def _refresh_sn_preview(self):
        sn_code = db_next_sn_code(self.current_user)
        self._sn_lbl.configure(text=f"หากไม่ระบุรหัส ระบบจะสร้าง:  {sn_code}")
        return sn_code

    def _on_name_key(self, _=None):
        if self._debounce_job:
            self.after_cancel(self._debounce_job)
        self._debounce_job = self.after(400, self._show_fuzzy)

    def _show_fuzzy(self):
        name = self._inputs["name"].get().strip()
        for w in self._fuzzy_f.winfo_children():
            w.destroy()
        if len(name) < 2:
            self._fuzzy_f.grid_forget()
            return
        candidates = fuzzy_candidates(name)
        if not candidates:
            self._fuzzy_f.grid_forget()
            return
        # ใส่ fuzzy frame ต่อจาก entry ของ name (row=1 ของ form grid)
        self._fuzzy_f.grid(row=1, column=0, columnspan=2, sticky="ew",
                           padx=0, pady=(2, 4),
                           in_=self._inputs["name"].master)
        CTkLabel(self._fuzzy_f,
                 text="⚠  พบรายชื่อที่ใกล้เคียง — กรุณาตรวจสอบก่อนบันทึก",
                 font=CTkFont(size=11, weight="bold"),
                 text_color=CLR["amber"]).pack(anchor="w", padx=10, pady=(6, 2))
        for c in candidates:
            CTkLabel(self._fuzzy_f,
                     text=f"  •  {c['supplier_id']}  {c['name']}  ({c['tier']})",
                     font=CTkFont(size=11), text_color=CLR["amber"]).pack(anchor="w", padx=10)
        CTkFrame(self._fuzzy_f, height=4, fg_color="transparent").pack()

    def _save(self):
        name  = self._inputs["name"].get().strip()
        phone = self._inputs["phone"].get().strip()
        if not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาระบุชื่อบริษัท", parent=self)
            return
        if not phone:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาระบุเบอร์โทร", parent=self)
            return
        # ใช้รหัสที่ user กรอกเอง ถ้าเว้นว่างถึงจะ auto-generate
        manual_code = self._inputs.get("supplier_id")
        manual_code = manual_code.get().strip() if manual_code else ""
        sn_code = manual_code if manual_code else db_next_sn_code(self.current_user)
        new_sup = {
            "id": 0,  # จะได้จาก DB หลัง INSERT
            "supplier_id": sn_code,
            "name":          name,
            "category":      self._cat_var.get(),
            "tier":          "SN", "is_locked": False,
            "source_tag":    "Manual",
            "contact":       self._inputs["contact"].get().strip(),
            "phone":         phone,
            "line_id":       self._inputs["line_id"].get().strip(),
            "email":         self._inputs["email"].get().strip(),
            "coverage_area": self._inputs["coverage_area"].get().strip(),
            "availability":  "พร้อม",
            "sn_created":    datetime.now().strftime("%Y-%m-%d"),
            "win_pct": 0, "sla_score": 0, "price_score": 0,
            "service_score": 0, "quality_score": 0,
            "credit_days": 0, "note": "", "win_loss_log": [],
            # ── Zoning ──────────────────────────────────────────────────────
            "dispatch_zone":    self._zone_var.get() if hasattr(self, "_zone_var") else "",
            "service_area":     self._service_var.get() if hasattr(self, "_service_var") else "National",
            "logistics_assets": ",".join(
                opt for opt, var in self._truck_vars.items() if var.get()
            ) if hasattr(self, "_truck_vars") else "",
            "business_type":    self._biz_var.get() if hasattr(self, "_biz_var") else "",
            "standard_focus":   self._std_var.get() if hasattr(self, "_std_var") else "",
            "credit_term_label": self._credit_lbl_var.get() if hasattr(self, "_credit_lbl_var") else "สด",
            "wh_zone":           self._wh_zone_e.get().strip() if hasattr(self, "_wh_zone_e") else "",
            "wh_coordinates":    self._wh_coord_e.get().strip() if hasattr(self, "_wh_coord_e") else "",
        }
        ok = db_save_supplier(new_sup, action="add", user=self.current_user)
        if ok:
            messagebox.showinfo("สร้างรหัสสำเร็จ",
                                f"รหัส Supplier ใหม่ของคุณคือ\n\n{sn_code}\n\n"
                                f"บันทึกลงฐานข้อมูลเรียบร้อยแล้ว", parent=self)
            if self.on_success:
                self.on_success()
            self.destroy()
        else:
            messagebox.showerror("ผิดพลาด", "บันทึกลงฐานข้อมูลไม่สำเร็จ\nกรุณาลองใหม่อีกครั้ง", parent=self)

# =============================================================================
#  D3 — BLACKLIST  POPUP  (บังคับระบุเหตุผล)
# =============================================================================
class BlacklistReasonPopup(CTkToplevel):
    """Popup บังคับใส่เหตุผลก่อน Flag Blacklist"""

    BL_REASONS = [
        "ส่งของไม่ตรงสเปก", "โกงราคาภายหลัง", "ไม่ส่งของตามสัญญา",
        "ปัญหาคุณภาพซ้ำซาก", "ทำให้งานหน้าไซต์เสียหาย", "อื่นๆ",
    ]

    def __init__(self, master, supplier: dict, on_confirm=None):
        super().__init__(master)
        self.sup        = supplier
        self.on_confirm = on_confirm
        self.title("Flag Blacklist — ระบุเหตุผล")
        _place_popup(self, 460, 380)
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["red"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        CTkLabel(hdr, text=f"Flag Blacklist: {supplier['name']}",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).pack(padx=16, pady=10, anchor="w")

        CTkLabel(self, text="เลือกเหตุผล (บังคับ):",
                 font=F(size=13, weight="bold")).grid(
            row=1, column=0, padx=20, pady=(14, 6), sticky="w")

        self._reason_var = tk.StringVar(value=self.BL_REASONS[0])
        for i, r in enumerate(self.BL_REASONS):
            CTkButton(self, text=r, fg_color="transparent",
                      text_color=CLR["red"], hover_color=CLR["red_lt"],
                      border_width=1, border_color=CLR["border"],
                      font=F(size=12), height=28, anchor="w",
                      command=lambda rv=r: self._reason_var.set(rv)).grid(
                row=2 + i, column=0, sticky="ew", padx=20, pady=2)

        CTkLabel(self, text="หมายเหตุเพิ่มเติม:",
                 font=F(size=12), text_color=CLR["gray"]).grid(
            row=9, column=0, padx=20, pady=(10, 2), sticky="w")
        self._note = CTkEntry(self, font=F(size=12), height=32)
        self._note.grid(row=10, column=0, sticky="ew", padx=20)

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=11, column=0, sticky="ew", pady=(12, 0))
        CTkButton(bf, text="ยกเลิก", fg_color="gray50", hover_color="gray40",
                  width=90, command=self.destroy).pack(side="right", padx=12, pady=10)
        CTkButton(bf, text="ยืนยัน Flag Blacklist",
                  fg_color=CLR["red"], hover_color="#991B1B",
                  width=180, command=self._confirm).pack(side="right", pady=10)

        self.transient(master)
        self.grab_set()

    def _confirm(self):
        reason = self._reason_var.get()
        note   = self._note.get().strip()
        full   = reason + (f" — {note}" if note else "")
        if self.on_confirm:
            self.on_confirm(full)
        self.destroy()

# =============================================================================
#  C6 — AUDIT TRAIL LOG  POPUP
# =============================================================================
class AuditLogPopup(CTkToplevel):
    """หน้าดู Audit Trail Log ทั้งหมด"""

    def __init__(self, master):
        super().__init__(master)
        self.title("Audit Trail Log")
        _place_popup(self, 760, 500)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        CTkLabel(hdr, text="Audit Trail Log — ประวัติการเปลี่ยนแปลงทั้งหมด",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).pack(padx=16, pady=10, anchor="w")

        tf = CTkFrame(self, fg_color=CLR["white"], corner_radius=0)
        tf.grid(row=1, column=0, sticky="nsew", padx=0)
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ["timestamp", "action", "user", "detail"]
        col_cfg = {
            "timestamp": ("เวลา",   140, "center"),
            "action":    ("Action", 160, "w"),
            "user":      ("ผู้ดำเนินการ", 110, "center"),
            "detail":    ("รายละเอียด",   330, "w"),
        }
        tree = ttk.Treeview(tf, columns=cols, show="headings",
                            style="SSL.Treeview")
        for col in cols:
            lbl, w, anchor = col_cfg[col]
            tree.heading(col, text=lbl)
            tree.column(col, width=w, anchor=anchor)

        vsb = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        logs = db_get_audit_log()
        if not logs:
            CTkLabel(tf, text="ยังไม่มีประวัติ", text_color=CLR["gray"],
                     font=F(size=13)).grid(row=0, column=0)
        for entry in logs:
            tree.insert("", "end", values=(
                entry.get("timestamp", ""),
                entry.get("action", ""),
                entry.get("user", ""),
                entry.get("detail", ""),
            ))

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        self.transient(master)
        self.grab_set()


# =============================================================================
#  B2 — TOP 5 VIEW  (แสดง Top 5 ต่อหมวด พร้อม Alerts)
# =============================================================================
class Top5View(CTkToplevel):
    """หน้า Top 5 Supplier ต่อหมวด พร้อม B3 Alert และ B4 Demote Suggestion"""

    def __init__(self, master):
        super().__init__(master)
        self.title("Top 5 Super Supplier — จัดอันดับตาม Weighted Score")
        _place_popup(self, 900, 620)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text="Top 5 Super Supplier ต่อหมวดสินค้า",
                 font=F(size=15, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=10, sticky="w")
        CTkLabel(hdr, text="Score = ราคา 20% + สต็อก 20% + บริการ 20% + SLA 20% + คุณภาพ 20%",
                 font=F(size=11), text_color="#93C5FD").grid(
            row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        body = CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=0)
        body.grid_columnconfigure(0, weight=1)

        cats   = db_get_categories()[1:]   # ข้าม "ทุกหมวด"

        # ── ดึง scores จาก cost_benchmarks จริง ──────────────────────────────
        bench_scores = _calc_supplier_scores_from_benchmark(cat=None)
        _sups        = db_get_all_suppliers()
        tier_map     = {s["name"]: s["tier"] for s in _sups}
        cat_map      = {s["name"]: s["category"] for s in _sups}
        is_locked_map = {s["name"]: s.get("is_locked", False) for s in _sups}
        avail_map    = {s["name"]: s.get("availability", "พร้อม") for s in _sups}

        # สร้าง df_all จาก benchmark scores + tier จาก suppliers
        rows_bench = []
        for sup_name, sc in bench_scores.items():
            tier = tier_map.get(sup_name, "Tier 2")
            # ใช้ category จาก suppliers table ก่อน ถ้าไม่มีค่อยใช้จาก benchmark
            cat  = cat_map.get(sup_name) or sc.get("category", "")
            win_pct     = sc["win_pct"]
            price_score = sc["price_score"]
            score = round(price_score * 0.60 + win_pct * 0.40)
            rows_bench.append({
                "name": sup_name, "category": cat, "tier": tier,
                "score": score, "win_pct": win_pct, "price_score": price_score,
                "is_locked": is_locked_map.get(sup_name, False),
                "availability": avail_map.get(sup_name, "พร้อม"),
            })

        # เพิ่ม Supplier ที่อยู่ใน DB แต่ยังไม่เคยอยู่ใน benchmark
        bench_names = set(bench_scores.keys())
        for s in _sups:
            if s["name"] not in bench_names and s.get("category"):
                rows_bench.append({
                    "name": s["name"], "category": s["category"], "tier": s["tier"],
                    "score": calc_score(s), "win_pct": s["win_pct"],
                    "price_score": s["price_score"],
                    "is_locked": s.get("is_locked", False),
                    "availability": s.get("availability", "พร้อม"),
                })

        df_all = pd.DataFrame(rows_bench) if rows_bench else pd.DataFrame(
            columns=["name","category","tier","score","win_pct","price_score","is_locked","availability"])

        # ── B3: หา Tier 2 ที่ Score แซง Tier 1 ──────────────────────────────
        alerts = []
        for cat in cats:
            cat_df = df_all[df_all["category"] == cat].copy()
            t1 = cat_df[cat_df["tier"] == "Tier 1"]["score"].max() if len(cat_df[cat_df["tier"] == "Tier 1"]) else 0
            t2_top = cat_df[cat_df["tier"] == "Tier 2"].sort_values("score", ascending=False).head(1)
            if not t2_top.empty and t2_top.iloc[0]["score"] > t1:
                alerts.append((cat, t2_top.iloc[0]["name"], int(t2_top.iloc[0]["score"]), int(t1)))

        if alerts:
            alert_f = CTkFrame(body, fg_color="#FEF3C7",
                               corner_radius=8, border_width=1,
                               border_color="#F59E0B")
            alert_f.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))
            CTkLabel(alert_f,
                     text="⚠  System Alert — Tier 2 Score แซง Tier 1 แล้ว! พิจารณา Promote",
                     font=F(size=12, weight="bold"),
                     text_color=CLR["amber"]).pack(anchor="w", padx=12, pady=(8, 4))
            for cat, name, t2s, t1s in alerts:
                CTkLabel(alert_f,
                         text=f"  •  [{cat}]  {name}  Score={t2s}  >  Tier 1 Max Score={t1s}",
                         font=F(size=11), text_color=CLR["amber"]).pack(anchor="w", padx=12)
            CTkFrame(alert_f, height=4, fg_color="transparent").pack()

        row_offset = 1
        for ci, cat in enumerate(cats):
            cat_df = df_all[df_all["category"] == cat].copy()
            active = cat_df[
                cat_df["tier"].isin(["Tier 1", "Tier 2"]) &
                (cat_df["score"] > 0)
            ].sort_values("score", ascending=False).head(5)

            sec = CTkFrame(body, fg_color=CLR["white"],
                           corner_radius=10, border_width=1,
                           border_color=CLR["border"])
            sec.grid(row=row_offset + ci, column=0, sticky="ew",
                     padx=16, pady=(8, 0))
            sec.grid_columnconfigure(0, weight=1)

            th = CTkFrame(sec, fg_color=CLR["navy"], corner_radius=0)
            th.grid(row=0, column=0, sticky="ew")
            CTkLabel(th, text=f"  {cat}",
                     font=F(size=13, weight="bold"),
                     text_color=CLR["white"]).pack(side="left", padx=10, pady=8)

            if active.empty:
                CTkLabel(sec, text="ยังไม่มีข้อมูล", text_color=CLR["gray"],
                         font=F(size=12)).grid(row=1, column=0, padx=16, pady=10)
                continue

            for rank, (_, row) in enumerate(active.iterrows(), 1):
                tier  = row["tier"]
                ts    = TIER_STYLE.get(tier, {"bg": CLR["gray_lt"], "fg": CLR["gray"]})
                score = int(row["score"])
                avail = row["availability"]

                # B4: Demote suggestion
                demote_warn = (tier == "Tier 1"
                               and row["win_pct"] < DEMOTE_THRESHOLD
                               and not row.get("is_locked"))

                rf = CTkFrame(sec, fg_color=CLR["red_lt"] if demote_warn else "transparent",
                              corner_radius=0)
                rf.grid(row=rank, column=0, sticky="ew", padx=8, pady=2)
                rf.grid_columnconfigure(2, weight=1)

                # rank badge
                rank_color = ["#F59E0B","#9CA3AF","#D97706","#6B7280","#6B7280"][rank-1]
                CTkLabel(rf, text=f" #{rank} ",
                         fg_color=rank_color, text_color=CLR["white"],
                         corner_radius=4, font=F(size=12, weight="bold"),
                         width=36).grid(row=0, column=0, padx=(8, 6), pady=6)

                # tier badge
                CTkLabel(rf, text=f" {tier} ",
                         fg_color=ts["bg"], text_color=ts["fg"],
                         corner_radius=4, font=F(size=11, weight="bold")).grid(
                    row=0, column=1, padx=(0, 8), pady=6)

                # name
                lock_icon = "🔒 " if row.get("is_locked") else ""
                name_lbl  = lock_icon + row["name"]
                CTkLabel(rf, text=name_lbl,
                         font=F(size=13, weight="bold")).grid(
                    row=0, column=2, sticky="w", pady=6)

                # score bar
                bar_w = max(4, int(score * 1.8))
                bar_f = CTkFrame(rf, height=10, fg_color=CLR["border"], corner_radius=5,
                                 width=190)
                bar_f.grid(row=0, column=3, padx=(8, 0), pady=6)
                bar_f.grid_propagate(False)
                bar_color = CLR["green"] if score >= 60 else CLR["amber"] if score >= 40 else CLR["red"]
                CTkFrame(bar_f, height=10, width=min(bar_w, 190),
                         fg_color=bar_color, corner_radius=5).place(x=0, y=0, relheight=1)

                CTkLabel(rf, text=f"Score: {score}",
                         font=F(size=12, weight="bold"),
                         text_color=CLR["navy"]).grid(row=0, column=4, padx=(8, 0), pady=6)

                # availability
                avail_icon = {"พร้อม": "✓", "สต็อกต่ำ": "⚠", "ปิดชั่วคราว": "✗"}
                avail_color = {
                    "พร้อม": CLR["green"], "สต็อกต่ำ": CLR["amber"],
                    "ปิดชั่วคราว": CLR["red"],
                }
                CTkLabel(rf, text=avail_icon.get(avail, "-"),
                         text_color=avail_color.get(avail, CLR["gray"]),
                         font=F(size=13, weight="bold")).grid(
                    row=0, column=5, padx=(8, 0), pady=6)

                # B4: demote suggestion badge
                if demote_warn:
                    CTkLabel(rf,
                             text=f" ⚠ Win {int(row['win_pct'])}% ต่ำกว่า {DEMOTE_THRESHOLD}% — พิจารณา Demote ",
                             fg_color=CLR["red_lt"], text_color=CLR["red"],
                             corner_radius=4, font=F(size=10, weight="bold")).grid(
                        row=0, column=6, padx=(6, 8), pady=6)

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        self.transient(master)


# =============================================================================
#  D4 — QUARTERLY SNAPSHOT  POPUP
# =============================================================================
class QuarterlySnapshotPopup(CTkToplevel):
    """เปรียบเทียบ Top 5 ย้อนหลัง 3 ไตรมาส"""

    def __init__(self, master):
        super().__init__(master)
        self.title("Quarterly Snapshot — เปรียบเทียบ Top 5 รายไตรมาส")
        _place_popup(self, 860, 560)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text="Quarterly Snapshot — ประวัติ Top 5 ต่อหมวด",
                 font=F(size=15, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=(10, 2), sticky="w")
        CTkLabel(hdr, text="ลูกศร ▲▼ แสดงเทรนด์จากไตรมาสก่อนหน้า",
                 font=F(size=11), text_color="#93C5FD").grid(
            row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        # Category selector
        sel_f = CTkFrame(hdr, fg_color="transparent")
        sel_f.grid(row=0, column=1, rowspan=2, padx=16, pady=10)
        CTkLabel(sel_f, text="หมวด:", font=F(size=12),
                 text_color="#93C5FD").pack(side="left", padx=(0, 6))
        _live_cats = db_get_categories()[1:]  # ข้าม "ทุกหมวด"
        _default_cat = _live_cats[0] if _live_cats else ""
        self._cat_var = tk.StringVar(value=_default_cat)
        CTkOptionMenu(sel_f, variable=self._cat_var,
                      values=_live_cats if _live_cats else ["(ยังไม่มีหมวด)"],
                      fg_color=CLR["white"], text_color=CLR["navy"],
                      button_color="#3B82F6", button_hover_color="#1A56DB",
                      font=F(size=12), width=130,
                      command=lambda _: self._render()).pack(side="left")

        self._body = CTkScrollableFrame(self, fg_color="transparent")
        self._body.grid(row=1, column=0, sticky="nsew")
        self._body.grid_columnconfigure(0, weight=1)

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        # ── ปุ่มบันทึก Snapshot ไตรมาสนี้ ──────────────────────────────
        self._save_lbl = CTkLabel(bf, text="", font=F(size=11),
                                  text_color=CLR["green"])
        self._save_lbl.pack(side="left", padx=12)
        CTkButton(bf, text="💾 บันทึก Snapshot ไตรมาสนี้",
                  fg_color=CLR["teal"], hover_color="#0F6E56",
                  width=200, height=30, font=F(size=12),
                  command=self._save_snapshot).pack(side="left", padx=(0, 8), pady=8)

        self._render()
        self.transient(master)

    def _save_snapshot(self):
        saved = db_save_quarterly_snapshot()
        if saved > 0:
            self._save_lbl.configure(
                text=f"✅ บันทึกแล้ว {saved} หมวด",
                text_color=CLR["green"])
            self._render()   # refresh ข้อมูล
        else:
            self._save_lbl.configure(
                text="⚠ ไม่มีข้อมูลให้บันทึก",
                text_color=CLR["amber"])

    def _render(self):
        for w in self._body.winfo_children():
            w.destroy()
        F   = CTkFont
        cat = self._cat_var.get()
        # ── ดึงจาก DB แทน MOCK_SNAPSHOTS ───────────────────────────────
        snaps = db_get_quarterly_snapshots(cat)
        if not snaps:
            CTkLabel(self._body, text="ยังไม่มีข้อมูล Snapshot\nกด 'บันทึก Snapshot ไตรมาสนี้' เพื่อเริ่มเก็บข้อมูล",
                     text_color=CLR["gray"], font=F(size=13),
                     justify="center").pack(pady=20)
            return

        # Header row (quarters)
        hf = CTkFrame(self._body, fg_color="transparent")
        hf.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 0))
        CTkLabel(hf, text="อันดับ", font=F(size=12, weight="bold"),
                 width=60, anchor="center").grid(row=0, column=0, padx=(0, 6))
        for qi, snap in enumerate(snaps):
            CTkLabel(hf, text=snap["quarter"],
                     font=F(size=13, weight="bold"), text_color=CLR["navy"],
                     width=220, anchor="center",
                     fg_color=CLR["blue_lt"], corner_radius=6).grid(
                row=0, column=qi + 1, padx=6)

        # Data rows
        for rank in range(5):
            rf = CTkFrame(self._body, fg_color=CLR["white"],
                          corner_radius=6, border_width=1,
                          border_color=CLR["border"])
            rf.grid(row=rank + 1, column=0, sticky="ew", padx=16, pady=3)
            rf.grid_columnconfigure((1, 2, 3), weight=1)

            rank_colors = ["#F59E0B","#9CA3AF","#D97706","#6B7280","#6B7280"]
            CTkLabel(rf, text=f"#{rank+1}",
                     fg_color=rank_colors[rank], text_color=CLR["white"],
                     corner_radius=4, font=F(size=13, weight="bold"),
                     width=48, height=32).grid(row=0, column=0, padx=8, pady=8)

            for qi, snap in enumerate(snaps):
                name, score = snap["top5"][rank] if rank < len(snap["top5"]) else ("", 0)
                if not name:
                    CTkLabel(rf, text="—", text_color=CLR["gray"],
                             font=F(size=12), width=220, anchor="center").grid(
                        row=0, column=qi + 1, padx=6, pady=8)
                    continue

                # trend arrow vs previous quarter
                trend = ""
                trend_color = CLR["gray"]
                if qi > 0:
                    prev_snaps = snaps[qi - 1]["top5"]
                    prev_names = [p[0] for p in prev_snaps]
                    prev_rank  = prev_names.index(name) + 1 if name in prev_names else 6
                    if rank + 1 < prev_rank:
                        trend, trend_color = " ▲", CLR["green"]
                    elif rank + 1 > prev_rank:
                        trend, trend_color = " ▼", CLR["red"]
                    else:
                        trend, trend_color = " —", CLR["gray"]

                cell = CTkFrame(rf, fg_color="transparent")
                cell.grid(row=0, column=qi + 1, padx=6, pady=8, sticky="ew")
                CTkLabel(cell, text=name + trend,
                         font=F(size=12), text_color=trend_color if trend else CLR["navy"],
                         anchor="center").pack()
                CTkLabel(cell, text=f"Win {score}%",
                         font=F(size=11), text_color=CLR["gray"],
                         anchor="center").pack()


# =============================================================================
#  WIN-LOSS ANALYSIS  POPUP
# =============================================================================
class WinLossAnalysisPopup(CTkToplevel):
    """สรุป Win/Loss Analysis ต่อ Supplier พร้อม bar chart breakdown เหตุผล"""

    def __init__(self, master, supplier: dict):
        super().__init__(master)
        self.sup = supplier
        self.title(f"Win-Loss Analysis — {supplier['name']}")
        _place_popup(self, 640, 520)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text=f"Win-Loss Analysis: {supplier['name']}",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=(10, 2), sticky="w")
        CTkLabel(hdr, text=supplier["supplier_id"] + f"  |  {supplier['category']}",
                 font=F(size=11), text_color="#93C5FD").grid(
            row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        body = CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)

        logs = supplier.get("win_loss_log", [])
        total = len(logs)
        wins  = sum(1 for l in logs if l["result"] == "Win")
        losses= total - wins
        win_rate = round(wins / total * 100) if total else 0

        # ── Summary tiles ───────────────────────────────────────────────────
        tile_f = CTkFrame(body, fg_color="transparent")
        tile_f.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 8))
        tile_f.grid_columnconfigure((0,1,2,3), weight=1)
        for col, (lbl, val, color, bg) in enumerate([
            ("ครั้งทั้งหมด",  str(total),       CLR["navy"],  "#F0F4FF"),
            ("Win",           str(wins),         CLR["green"], CLR["green_lt"]),
            ("Loss",          str(losses),       CLR["red"],   CLR["red_lt"]),
            ("Win Rate",      f"{win_rate}%",    CLR["teal"],  CLR["teal_lt"]),
        ]):
            t = CTkFrame(tile_f, fg_color=bg, corner_radius=8,
                         border_width=1, border_color=CLR["border"])
            t.grid(row=0, column=col, padx=6, pady=0, sticky="ew")
            CTkLabel(t, text=val, font=F(size=24, weight="bold"),
                     text_color=color).pack(pady=(10, 0))
            CTkLabel(t, text=lbl, font=F(size=11), text_color=CLR["gray"]).pack(pady=(0, 10))

        # ── Win Rate bar ─────────────────────────────────────────────────────
        bar_wrap = CTkFrame(body, fg_color="transparent")
        bar_wrap.grid(row=1, column=0, sticky="ew", padx=16, pady=(4, 12))
        bar_wrap.grid_columnconfigure(1, weight=1)
        CTkLabel(bar_wrap, text="Win Rate:", font=F(size=12),
                 text_color=CLR["gray"], width=70).grid(row=0, column=0, sticky="w")
        outer = CTkFrame(bar_wrap, height=16, fg_color=CLR["border"],
                         corner_radius=8)
        outer.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        outer.grid_propagate(False)
        bar_color = CLR["green"] if win_rate >= 60 else CLR["amber"] if win_rate >= 30 else CLR["red"]
        if win_rate > 0:
            inner = CTkFrame(outer, height=16,
                             fg_color=bar_color, corner_radius=8)
            inner.place(relx=0, rely=0, relwidth=win_rate/100, relheight=1)
        CTkLabel(bar_wrap, text=f"{win_rate}%", font=F(size=12, weight="bold"),
                 text_color=bar_color, width=40).grid(row=0, column=2, sticky="e")

        # ── Loss reason breakdown ─────────────────────────────────────────────
        if losses > 0:
            from collections import Counter
            reasons = [l.get("reason","") or "ไม่ระบุ"
                       for l in logs if l["result"] == "Loss"]
            counted = Counter(reasons).most_common()

            CTkLabel(body, text="เหตุผลที่แพ้ (Loss Reason Breakdown)",
                     font=F(size=12, weight="bold"),
                     text_color=CLR["blue"]).grid(
                row=2, column=0, padx=16, pady=(4, 6), sticky="w")

            rb_f = CTkFrame(body, fg_color=CLR["white"], corner_radius=8,
                            border_width=1, border_color=CLR["border"])
            rb_f.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 8))
            rb_f.grid_columnconfigure(1, weight=1)

            for ri, (reason, cnt) in enumerate(counted):
                pct = round(cnt / losses * 100)
                CTkLabel(rb_f, text=reason, font=F(size=12),
                         width=200, anchor="w").grid(
                    row=ri, column=0, padx=(12, 8), pady=5, sticky="w")
                bar_out = CTkFrame(rb_f, height=12, fg_color=CLR["border"],
                                   corner_radius=6)
                bar_out.grid(row=ri, column=1, sticky="ew", padx=(0, 8), pady=5)
                bar_out.grid_propagate(False)
                CTkFrame(bar_out, height=12, fg_color=CLR["red"],
                         corner_radius=6,
                         width=max(8, int(pct * 2))).place(x=0, y=0, relheight=1)
                CTkLabel(rb_f, text=f"{cnt} ครั้ง ({pct}%)",
                         font=F(size=11), text_color=CLR["gray"],
                         width=80, anchor="e").grid(
                    row=ri, column=2, padx=(0, 12), pady=5, sticky="e")

        # ── Log table ──────────────────────────────────────────────────────
        CTkLabel(body, text="ประวัติทั้งหมด",
                 font=F(size=12, weight="bold"),
                 text_color=CLR["blue"]).grid(
            row=4, column=0, padx=16, pady=(8, 6), sticky="w")

        log_f = CTkFrame(body, fg_color=CLR["white"], corner_radius=8,
                         border_width=1, border_color=CLR["border"])
        log_f.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 12))

        if not logs:
            CTkLabel(log_f, text="ยังไม่มีประวัติ", text_color=CLR["gray"],
                     font=F(size=12)).pack(padx=16, pady=12)
        else:
            for li, entry in enumerate(reversed(logs)):
                rf  = CTkFrame(log_f, fg_color="transparent")
                rf.pack(fill="x", padx=12, pady=3)
                dot_c = CLR["green"] if entry["result"] == "Win" else CLR["red"]
                CTkLabel(rf, text="●", text_color=dot_c,
                         font=F(size=13), width=16).pack(side="left")
                txt = f"{entry['date']}  {entry['result']}"
                if entry.get("reason"):
                    txt += f"  —  {entry['reason']}"
                CTkLabel(rf, text=txt, font=F(size=12)).pack(side="left", padx=(6, 0))

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)
        self.transient(master)
        self.grab_set()


# =============================================================================
#  EXPORT  POPUP
# =============================================================================
class ExportPopup(CTkToplevel):
    """Export รายชื่อ Supplier ที่กำลังแสดงเป็น CSV หรือ Excel"""

    def __init__(self, master, df: "pd.DataFrame"):
        super().__init__(master)
        self.df = df
        self.title("Export Supplier List")
        _place_popup(self, 420, 280)
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        CTkLabel(hdr, text="Export รายชื่อ Supplier",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).pack(padx=16, pady=10, anchor="w")

        info = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=8)
        info.grid(row=1, column=0, padx=20, pady=(14, 0), sticky="ew")
        CTkLabel(info,
                 text=f"รายการที่จะ Export: {len(df)} รายการ  ({', '.join(df['category'].unique()[:3])}...)",
                 font=F(size=12), text_color=CLR["gray"]).pack(padx=12, pady=8)

        CTkLabel(self, text="เลือกรูปแบบไฟล์:",
                 font=F(size=13, weight="bold")).grid(
            row=2, column=0, padx=20, pady=(12, 6), sticky="w")

        btn_f = CTkFrame(self, fg_color="transparent")
        btn_f.grid(row=3, column=0, padx=20, sticky="ew")
        btn_f.grid_columnconfigure((0,1), weight=1)

        CTkButton(btn_f, text="📄 Export CSV",
                  fg_color=CLR["teal"], hover_color="#0F6E56", height=44,
                  font=F(size=13, weight="bold"),
                  command=self._export_csv).grid(row=0, column=0, padx=(0, 6), pady=4, sticky="ew")
        CTkButton(btn_f, text="📊 Export Excel (.xlsx)",
                  fg_color=CLR["green"], hover_color="#065F46", height=44,
                  font=F(size=13, weight="bold"),
                  command=self._export_excel).grid(row=0, column=1, padx=(6, 0), pady=4, sticky="ew")

        self._status_lbl = CTkLabel(self, text="", font=F(size=12),
                                    text_color=CLR["green"])
        self._status_lbl.grid(row=4, column=0, padx=20, pady=(8, 0), sticky="w")

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        self.transient(master)
        self.grab_set()

    def _get_export_df(self):
        cols = ["supplier_id", "name", "category", "tier", "source_tag",
                "availability", "contact", "phone", "line_id", "email",
                "coverage_area", "score", "win_pct", "sla_score",
                "price_score", "service_score", "quality_score",
                "credit_days", "note"]
        export_cols = [c for c in cols if c in self.df.columns]
        return self.df[export_cols].copy()

    def _export_csv(self):
        try:
            import tkinter.filedialog as fd
            path = fd.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                initialfile=f"super_supplier_{datetime.now().strftime('%Y%m%d')}.csv",
                parent=self)
            if not path:
                return
            self._get_export_df().to_csv(path, index=False, encoding="utf-8-sig")
            self._status_lbl.configure(text=f"✓ บันทึกสำเร็จ: {path.split('/')[-1]}")
        except Exception as e:
            self._status_lbl.configure(text=f"✗ Error: {e}", text_color=CLR["red"])

    def _export_excel(self):
        try:
            import tkinter.filedialog as fd
            path = fd.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"super_supplier_{datetime.now().strftime('%Y%m%d')}.xlsx",
                parent=self)
            if not path:
                return
            self._get_export_df().to_excel(path, index=False)
            self._status_lbl.configure(text=f"✓ บันทึกสำเร็จ: {path.split('/')[-1]}")
        except ImportError:
            self._status_lbl.configure(
                text="✗ ต้องติดตั้ง openpyxl ก่อน: pip install openpyxl",
                text_color=CLR["red"])
        except Exception as e:
            self._status_lbl.configure(text=f"✗ Error: {e}", text_color=CLR["red"])


# =============================================================================
#  SN AGING ALERT  POPUP
# =============================================================================
class SNAgingPopup(CTkToplevel):
    """แสดงรายชื่อ SN ที่ยังไม่ Convert เกิน SN_AGING_DAYS วัน"""

    def __init__(self, master):
        super().__init__(master)
        self.title(f"SN Aging Alert — ยังไม่ Convert เกิน {SN_AGING_DAYS} วัน")
        _place_popup(self, 680, 420)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        aging = get_aging_sns(SN_AGING_DAYS)

        hdr = CTkFrame(self,
                       fg_color=CLR["red"] if aging else CLR["green"],
                       corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        if aging:
            CTkLabel(hdr,
                     text=f"⚠  พบ {len(aging)} รายการ SN ที่ค้างนานเกิน {SN_AGING_DAYS} วัน — กรุณา Convert โดยเร็ว",
                     font=F(size=13, weight="bold"),
                     text_color=CLR["white"]).grid(
                row=0, column=0, padx=16, pady=10, sticky="w")
        else:
            CTkLabel(hdr,
                     text=f"✓  ไม่มี SN ที่ค้างนานเกิน {SN_AGING_DAYS} วัน",
                     font=F(size=13, weight="bold"),
                     text_color=CLR["white"]).grid(
                row=0, column=0, padx=16, pady=10, sticky="w")

        if aging:
            tf = CTkFrame(self, fg_color=CLR["white"], corner_radius=0)
            tf.grid(row=1, column=0, sticky="nsew")
            tf.grid_columnconfigure(0, weight=1)
            tf.grid_rowconfigure(0, weight=1)

            cols = ["supplier_id", "name", "category", "contact", "sn_created", "aging"]
            col_cfg = {
                "supplier_id": ("รหัส SN",     130, "center"),
                "name":        ("ชื่อ",         180, "w"),
                "category":    ("หมวด",          90, "center"),
                "contact":     ("ผู้ติดต่อ",    110, "w"),
                "sn_created":  ("วันที่เพิ่ม",  110, "center"),
                "aging":       ("ค้างมา (วัน)", 100, "center"),
            }
            tree = ttk.Treeview(tf, columns=cols, show="headings",
                                style="SSL.Treeview")
            for col in cols:
                lbl, w, anchor = col_cfg[col]
                tree.heading(col, text=lbl)
                tree.column(col, width=w, anchor=anchor)

            tree.tag_configure("critical", background="#FEE2E2")
            tree.tag_configure("warning",  background="#FEF3C7")

            for s in aging:
                days = s["_aging_days"]
                tag  = "critical" if days >= 60 else "warning"
                tree.insert("", "end",
                            values=(s["supplier_id"], s["name"], s["category"],
                                    s["contact"], s.get("sn_created",""),
                                    f"{days} วัน"),
                            tags=(tag,))

            vsb = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
        else:
            CTkLabel(self, text="ทุก SN ถูก Convert ทันเวลา",
                     text_color=CLR["green"], font=F(size=14)).grid(
                row=1, column=0, pady=30)

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        self.transient(master)
        self.grab_set()


# =============================================================================
#  NOTIFICATION  HELPER
# =============================================================================
def _push_noti(msg: str, ntype: str = "medium"):
    """เพิ่ม notification เข้า MOCK_NOTIFICATIONS"""
    MOCK_NOTIFICATIONS.append({
        "id":   len(MOCK_NOTIFICATIONS) + 1,
        "msg":  msg,
        "type": ntype,   # high / medium / low
        "read": False,
        "ts":   datetime.now().strftime("%Y-%m-%d %H:%M"),
    })


# =============================================================================
#  NOTIFICATION CENTER  POPUP
# =============================================================================
class NotificationCenterPopup(CTkToplevel):
    """หน้า Notification Center ทั้งหมด"""

    TYPE_COLOR = {"high": CLR["red"], "medium": CLR["amber"], "low": CLR["blue"]}
    TYPE_BG    = {"high": CLR["red_lt"], "medium": CLR["amber_lt"], "low": CLR["blue_lt"]}
    TYPE_LABEL = {"high": "🔴 สำคัญมาก", "medium": "🟡 ปานกลาง", "low": "🔵 ทั่วไป"}

    def __init__(self, master, on_read=None):
        super().__init__(master)
        self.on_read = on_read
        self.title("Notification Center")
        _place_popup(self, 560, 500)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text="🔔  Notification Center",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=10, sticky="w")
        CTkButton(hdr, text="อ่านทั้งหมด", width=90, height=26,
                  fg_color="#3B82F6", hover_color="#1A56DB",
                  font=F(size=11),
                  command=self._mark_all_read).grid(row=0, column=1, padx=12, pady=10)

        body = CTkScrollableFrame(self, fg_color=CLR["gray_lt"])
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        self._body = body

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        self._render()
        self.transient(master)
        self.grab_set()

    def _render(self):
        for w in self._body.winfo_children():
            w.destroy()
        F    = CTkFont
        logs = list(reversed(MOCK_NOTIFICATIONS))
        if not logs:
            CTkLabel(self._body, text="ไม่มีการแจ้งเตือน",
                     text_color=CLR["gray"], font=F(size=13)).pack(pady=30)
            return
        for entry in logs:
            is_read = entry.get("read", False)
            ntype   = entry.get("type", "low")
            bg      = CLR["white"] if is_read else self.TYPE_BG.get(ntype, CLR["blue_lt"])
            card    = CTkFrame(self._body, fg_color=bg, corner_radius=8,
                               border_width=1, border_color=CLR["border"])
            card.pack(fill="x", padx=10, pady=4)
            card.grid_columnconfigure(1, weight=1)

            # dot
            dot_c = self.TYPE_COLOR.get(ntype, CLR["blue"])
            CTkLabel(card, text="●", text_color=dot_c if not is_read else CLR["gray"],
                     font=F(size=11), width=14).grid(row=0, column=0, padx=(10, 4), pady=10)

            # message
            CTkLabel(card, text=entry.get("msg", ""),
                     font=F(size=12),
                     text_color=CLR["gray"] if is_read else CLR["navy"],
                     wraplength=380, justify="left", anchor="w").grid(
                row=0, column=1, sticky="w", pady=10)

            # type + ts
            meta = f"{self.TYPE_LABEL.get(ntype,'')}  •  {entry.get('ts','')}"
            CTkLabel(card, text=meta, font=F(size=10),
                     text_color=CLR["gray"]).grid(
                row=1, column=1, sticky="w", padx=(0, 8), pady=(0, 8))

            # mark read button
            if not is_read:
                CTkButton(card, text="✓", width=28, height=24,
                          fg_color=dot_c, hover_color="#333",
                          font=F(size=11),
                          command=lambda e=entry: self._mark_read(e)).grid(
                    row=0, column=2, padx=8, pady=10)

    def _mark_read(self, entry):
        entry["read"] = True
        self._render()
        if self.on_read:
            self.on_read()

    def _mark_all_read(self):
        for n in MOCK_NOTIFICATIONS:
            n["read"] = True
        self._render()
        if self.on_read:
            self.on_read()


# =============================================================================
#  BULK IMPORT  POPUP
# =============================================================================
class BulkImportPopup(CTkToplevel):
    """นำเข้า Supplier จาก CSV / Excel ทีเดียวหลายรายการ"""

    REQUIRED_COLS = {"name", "category", "phone"}
    OPTIONAL_COLS = {
        "contact", "line_id", "email", "coverage_area",
        "win_pct", "sla_score", "price_score", "service_score", "quality_score",
        "credit_days", "tier", "source_tag", "note", "availability",
    }

    def __init__(self, master, on_success=None, current_user="USER_DEMO"):
        super().__init__(master)
        self.on_success   = on_success
        self.current_user = current_user
        self.title("Bulk Import Supplier จาก CSV / Excel")
        _place_popup(self, 660, 560)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._df_preview = None
        F = CTkFont

        # Header
        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text="📥  Bulk Import — นำเข้า Supplier จากไฟล์",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=10, sticky="w")

        # Instructions
        inst = CTkFrame(self, fg_color=CLR["blue_lt"], corner_radius=8)
        inst.grid(row=1, column=0, padx=16, pady=(12, 0), sticky="ew")
        CTkLabel(inst,
                 text="คอลัมน์บังคับ: name, category, phone\n"
                      "คอลัมน์เสริม: contact, line_id, email, coverage_area, tier, source_tag,\n"
                      "              win_pct, sla_score, price_score, service_score, quality_score,\n"
                      "              credit_days, note",
                 font=F(size=11), text_color=CLR["blue"],
                 justify="left").pack(padx=12, pady=8, anchor="w")

        # Preview frame
        preview_f = CTkFrame(self, fg_color=CLR["white"], corner_radius=0,
                             border_width=1, border_color=CLR["border"])
        preview_f.grid(row=2, column=0, sticky="nsew", padx=0)
        preview_f.grid_columnconfigure(0, weight=1)
        preview_f.grid_rowconfigure(0, weight=1)
        self._preview_f = preview_f

        CTkLabel(preview_f, text="กดปุ่ม 'เลือกไฟล์' เพื่อเริ่มต้น",
                 text_color=CLR["gray"], font=F(size=13)).pack(expand=True)

        # Bottom bar
        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=3, column=0, sticky="ew")
        CTkButton(bf, text="ยกเลิก", fg_color="gray50", hover_color="gray40",
                  width=90, command=self.destroy).pack(side="right", padx=12, pady=10)
        self._import_btn = CTkButton(
            bf, text="✓ นำเข้าทั้งหมด",
            fg_color=CLR["green"], hover_color="#065F46",
            width=140, state="disabled",
            command=self._do_import)
        self._import_btn.pack(side="right", pady=10)
        CTkButton(bf, text="📂 เลือกไฟล์ (CSV/Excel)",
                  fg_color=CLR["blue"], hover_color="#1e40af",
                  width=180, command=self._pick_file).pack(side="left", padx=12, pady=10)

        # Download template button
        CTkButton(bf, text="⬇ Template CSV",
                  fg_color=CLR["teal"], hover_color="#0F6E56",
                  width=120, command=self._download_template).pack(side="left", pady=10)

        self._status_lbl = CTkLabel(bf, text="", font=F(size=12),
                                    text_color=CLR["green"])
        self._status_lbl.pack(side="left", padx=10)

        self.transient(master)
        self.grab_set()

    def _download_template(self):
        """บันทึก template CSV ให้ user ดาวน์โหลด"""
        import tkinter.filedialog as fd
        path = fd.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="supplier_import_template.csv",
            parent=self)
        if not path:
            return
        import csv
        header = ["name", "category", "phone", "contact", "line_id", "email",
                  "coverage_area", "tier", "source_tag", "win_pct",
                  "sla_score", "price_score", "service_score", "quality_score",
                  "credit_days", "note"]
        sample = ["ตัวอย่าง บริษัท จำกัด", "เหล็กเส้น", "081-000-0000",
                  "คุณตัวอย่าง", "@example", "ex@company.com",
                  "กรุงเทพ", "Tier 2", "Manual", "50", "80", "75", "0", "0", "30", ""]
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerow(sample)
        self._status_lbl.configure(text=f"✓ บันทึก template แล้ว")

    def _pick_file(self):
        import tkinter.filedialog as fd
        path = fd.askopenfilename(
            filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls")],
            parent=self)
        if not path:
            return
        try:
            if path.endswith(".csv"):
                df = pd.read_csv(path, encoding="utf-8-sig")
            else:
                df = pd.read_excel(path)
            df.columns = [c.strip().lower() for c in df.columns]
            missing = self.REQUIRED_COLS - set(df.columns)
            if missing:
                messagebox.showerror(
                    "คอลัมน์ไม่ครบ",
                    f"ไม่พบคอลัมน์บังคับ: {', '.join(missing)}\n"
                    f"กรุณาดาวน์โหลด template และกรอกข้อมูลใหม่",
                    parent=self)
                return
            # ลบแถวที่ name ว่าง
            df = df[df["name"].notna() & (df["name"].astype(str).str.strip() != "")]
            self._df_preview = df
            self._show_preview(df)
            self._import_btn.configure(state="normal")
            self._status_lbl.configure(
                text=f"พบ {len(df)} รายการ — ตรวจสอบแล้วกด 'นำเข้าทั้งหมด'",
                text_color=CLR["blue"])
        except Exception as e:
            messagebox.showerror("อ่านไฟล์ไม่ได้", str(e), parent=self)

    def _show_preview(self, df):
        for w in self._preview_f.winfo_children():
            w.destroy()
        self._preview_f.grid_columnconfigure(0, weight=1)
        self._preview_f.grid_rowconfigure(0, weight=1)

        show_cols = [c for c in ["name", "category", "tier", "phone", "source_tag", "note"]
                     if c in df.columns]
        tree = ttk.Treeview(self._preview_f, columns=show_cols,
                            show="headings", style="SSL.Treeview")
        col_w = {"name": 180, "category": 100, "tier": 80,
                 "phone": 110, "source_tag": 80, "note": 140}
        for col in show_cols:
            tree.heading(col, text=col)
            tree.column(col, width=col_w.get(col, 100), anchor="w")

        # highlight duplicates
        existing_names = {s["name"].lower() for s in db_get_all_suppliers()}
        tree.tag_configure("dup", background="#FEF3C7")

        for _, row in df.head(50).iterrows():
            tag = ("dup",) if str(row.get("name","")).lower() in existing_names else ()
            tree.insert("", "end",
                        values=tuple(str(row.get(c, "")) for c in show_cols),
                        tags=tag)

        vsb = ttk.Scrollbar(self._preview_f, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(self._preview_f, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        if len(df) > 50:
            CTkLabel(self._preview_f,
                     text=f"แสดง 50 แถวแรก (ทั้งหมด {len(df)} แถว)",
                     text_color=CLR["gray"],
                     font=CTkFont(size=11)).grid(row=2, column=0, pady=4)

    def _do_import(self):
        df = self._df_preview
        if df is None or df.empty:
            return
        skipped = 0
        added   = 0

        def _safe(row, col, default=""):
            v = row.get(col, default)
            return str(v).strip() if pd.notna(v) else default

        for _, row in df.iterrows():
            name = _safe(row, "name")
            if not name:
                skipped += 1
                continue

            tier       = _safe(row, "tier", "SN")
            if tier not in ("Tier 1","Tier 2","SN","Blacklist"):
                tier = "SN"
            source_tag = _safe(row, "source_tag", "Manual")
            if source_tag not in ("Legacy","Manual","System"):
                source_tag = "Manual"

            def _num(col, default=0):
                try:
                    return int(float(row.get(col, default) or default))
                except Exception:
                    return default

            sn_code = db_next_sn_code(self.current_user)
            new_sup = {
                "id":           0,
                "supplier_id":  sn_code if tier == "SN" else _safe(row, "supplier_id", sn_code),
                "name":         name,
                "category":     _safe(row, "category", "เหล็กเส้น"),
                "tier":         tier,
                "is_locked":    False,
                "source_tag":   source_tag,
                "contact":      _safe(row, "contact"),
                "phone":        _safe(row, "phone"),
                "line_id":      _safe(row, "line_id"),
                "email":        _safe(row, "email"),
                "coverage_area":_safe(row, "coverage_area"),
                "availability": _safe(row, "availability", "พร้อม"),
                "win_pct":       _num("win_pct"),
                "sla_score":     _num("sla_score"),
                "price_score":   _num("price_score"),
                "service_score": _num("service_score"),
                "quality_score": _num("quality_score"),
                "credit_days":   _num("credit_days"),
                "note":          _safe(row, "note"),
                "win_loss_log":  [],
                "sn_created":    datetime.now().strftime("%Y-%m-%d"),
            }
            ok = db_save_supplier(new_sup, action="add", user=self.current_user)
            if ok:
                added += 1
                _push_noti(f"Bulk Import: เพิ่ม '{name}' ({sn_code}) โดย {self.current_user}", "low")
            else:
                skipped += 1

        msg = f"นำเข้าสำเร็จ {added} รายการ"
        if skipped:
            msg += f"  (ข้าม {skipped} แถวที่ไม่มีชื่อ)"
        messagebox.showinfo("Bulk Import สำเร็จ", msg, parent=self)
        self._import_btn.configure(state="disabled")
        self._status_lbl.configure(text=f"✓ เพิ่มแล้ว {added} รายการ")
        if self.on_success:
            self.on_success()


# =============================================================================
#  SUGGESTED SUPPLIER  POPUP  (สำหรับ Cost Benchmarking)
# =============================================================================
class SuggestedSupplierPopup(CTkToplevel):
    """Pop Top 5 Super Supplier แนะนำ เมื่อ PU เปิดหน้าสอบราคา"""

    def __init__(self, master, category: str = "ทุกหมวด", on_select=None):
        super().__init__(master)
        self.on_select = on_select
        self.title(f"Suggested Supplier — {category}")
        _place_popup(self, 720, 480)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr,
                 text=f"⭐  Top 5 Suggested Supplier — {category}",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=(10,2), sticky="w")
        CTkLabel(hdr,
                 text="คลิก 'เลือก' เพื่อใช้ข้อมูล Supplier นี้ในใบสอบราคา",
                 font=F(size=11), text_color="#93C5FD").grid(
            row=1, column=0, padx=16, pady=(0,10), sticky="w")

        # Category selector
        sel_f = CTkFrame(hdr, fg_color="transparent")
        sel_f.grid(row=0, column=1, rowspan=2, padx=16, pady=10)
        CTkLabel(sel_f, text="หมวด:", font=F(size=11),
                 text_color="#93C5FD").pack(side="left", padx=(0,4))
        self._cat_var = tk.StringVar(value=category)
        _live_cats_all = db_get_categories()  # รวม "ทุกหมวด"
        CTkOptionMenu(sel_f, variable=self._cat_var,
                      values=_live_cats_all if _live_cats_all else ["ทุกหมวด"],
                      fg_color=CLR["white"], text_color=CLR["navy"],
                      button_color="#3B82F6", button_hover_color="#1A56DB",
                      font=F(size=12), width=130,
                      command=lambda _: self._render()).pack(side="left")

        body = CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        self._body = body

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)
        CTkLabel(bf,
                 text="🔒 = Lock Tier 1  •  [L]=Legacy  [M]=Manual  [S]=System",
                 font=F(size=11), text_color=CLR["gray"]).pack(
            side="left", padx=14, pady=8)

        self._render()
        self.transient(master)
        self.grab_set()

    def _render(self):
        for w in self._body.winfo_children():
            w.destroy()
        F   = CTkFont
        cat = self._cat_var.get()
        df  = get_suppliers_df(cat=cat if cat != "ทุกหมวด" else "ทุกหมวด",
                               tier="ทุก Tier")
        df  = df[df["tier"].isin(["Tier 1", "Tier 2"])].sort_values(
                  "score", ascending=False).head(5)

        if df.empty:
            CTkLabel(self._body, text="ไม่พบ Supplier ในหมวดนี้",
                     text_color=CLR["gray"], font=F(size=13)).pack(pady=30)
            return

        rank_colors = ["#F59E0B","#9CA3AF","#D97706","#6B7280","#6B7280"]

        for rank, (_, row) in enumerate(df.iterrows(), 1):
            tier  = row["tier"]
            ts    = TIER_STYLE.get(tier, {"bg":CLR["gray_lt"],"fg":CLR["gray"]})
            score = int(row["score"])
            avail = row["availability"]
            avail_color = {"พร้อม": CLR["green"], "สต็อกต่ำ": CLR["amber"],
                           "ปิดชั่วคราว": CLR["red"]}

            card = CTkFrame(self._body, fg_color=CLR["white"],
                            corner_radius=8, border_width=1,
                            border_color=CLR["border"])
            card.pack(fill="x", padx=16, pady=5)
            card.grid_columnconfigure(2, weight=1)

            # rank badge
            CTkLabel(card, text=f" #{rank} ",
                     fg_color=rank_colors[rank-1], text_color=CLR["white"],
                     corner_radius=4, font=F(size=12, weight="bold"),
                     width=36).grid(row=0, column=0, padx=(10,6), pady=10, rowspan=2)

            # tier badge
            CTkLabel(card, text=f" {tier} ",
                     fg_color=ts["bg"], text_color=ts["fg"],
                     corner_radius=4, font=F(size=11, weight="bold")).grid(
                row=0, column=1, padx=(0,8), pady=10, rowspan=2, sticky="w")

            # name + contact
            lock = "🔒 " if row.get("is_locked") else ""
            CTkLabel(card, text=lock + row["name"],
                     font=F(size=13, weight="bold")).grid(
                row=0, column=2, sticky="w", pady=(10,2))
            CTkLabel(card,
                     text=f"{row['contact']}  •  {row['phone']}  •  เครดิต {int(row['credit_days'])} วัน",
                     font=F(size=11), text_color=CLR["gray"]).grid(
                row=1, column=2, sticky="w", pady=(0,10))

            # score
            CTkLabel(card, text=f"Score\n{score}",
                     font=F(size=12, weight="bold"),
                     text_color=CLR["navy"], justify="center").grid(
                row=0, column=3, rowspan=2, padx=8, pady=10)

            # availability
            avail_icon = {"พร้อม":"✓ พร้อม","สต็อกต่ำ":"⚠ ต่ำ","ปิดชั่วคราว":"✗ ปิด"}
            CTkLabel(card,
                     text=avail_icon.get(avail, avail),
                     font=F(size=11, weight="bold"),
                     text_color=avail_color.get(avail, CLR["gray"])).grid(
                row=0, column=4, rowspan=2, padx=(0,8), pady=10)

            # select button
            if self.on_select:
                CTkButton(card, text="เลือก", width=70, height=30,
                          fg_color=CLR["green"], hover_color="#065F46",
                          font=F(size=12),
                          command=lambda s=dict(row): self._select(s)).grid(
                    row=0, column=5, rowspan=2, padx=(0,10), pady=10)

    def _select(self, sup_dict):
        if self.on_select:
            self.on_select(sup_dict)
        self.destroy()


# =============================================================================
#  RANKING TIMELINE  POPUP
# =============================================================================
class RankingTimelinePopup(CTkToplevel):
    """History Timeline — เส้นการขึ้น-ลงอันดับ Supplier รายไตรมาส"""

    def __init__(self, master):
        super().__init__(master)
        self.title("Ranking Timeline — ประวัติการเปลี่ยนอันดับ")
        _place_popup(self, 860, 580)
        self.resizable(True, True)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        F = CTkFont

        hdr = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        CTkLabel(hdr, text="📈  Ranking Timeline — อันดับ Supplier รายไตรมาส",
                 font=F(size=14, weight="bold"),
                 text_color=CLR["white"]).grid(row=0, column=0, padx=16, pady=(10,2), sticky="w")
        CTkLabel(hdr, text="▲ = ขึ้นอันดับ  ▼ = ลงอันดับ  — = คงที่  ✦ = เข้าใหม่",
                 font=F(size=11), text_color="#93C5FD").grid(
            row=1, column=0, padx=16, pady=(0,10), sticky="w")

        # Category selector
        sel_f = CTkFrame(hdr, fg_color="transparent")
        sel_f.grid(row=0, column=1, rowspan=2, padx=16, pady=10)
        CTkLabel(sel_f, text="หมวด:", font=F(size=11),
                 text_color="#93C5FD").pack(side="left", padx=(0,4))
        _live_cats = db_get_categories()[1:]  # ข้าม "ทุกหมวด"
        _default_cat = _live_cats[0] if _live_cats else ""
        self._cat_var = tk.StringVar(value=_default_cat)
        CTkOptionMenu(sel_f, variable=self._cat_var,
                      values=_live_cats if _live_cats else ["(ยังไม่มีหมวด)"],
                      fg_color=CLR["white"], text_color=CLR["navy"],
                      button_color="#3B82F6", button_hover_color="#1A56DB",
                      font=F(size=12), width=130,
                      command=lambda _: self._render()).pack(side="left")

        body = CTkScrollableFrame(self, fg_color=CLR["gray_lt"])
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        self._body = body

        bf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        bf.grid(row=2, column=0, sticky="ew")
        CTkButton(bf, text="ปิด", width=90, fg_color="gray50",
                  hover_color="gray40", command=self.destroy).pack(
            side="right", padx=12, pady=8)

        self._render()
        self.transient(master)

    def _render(self):
        for w in self._body.winfo_children():
            w.destroy()
        F    = CTkFont
        cat  = self._cat_var.get()
        # ── ดึงจาก DB แทน MOCK_SNAPSHOTS ───────────────────────────────
        snaps = db_get_quarterly_snapshots(cat)
        if not snaps:
            CTkLabel(self._body, text="ยังไม่มีข้อมูล\nกด 'บันทึก Snapshot ไตรมาสนี้' ใน Quarterly Snapshot ก่อน",
                     text_color=CLR["gray"], font=F(size=13),
                     justify="center").pack(pady=30)
            return

        quarters = [s["quarter"] for s in snaps]
        n_q      = len(quarters)

        # Build {name: [rank_q1, rank_q2, ...]}  None = ไม่อยู่ใน top5 ไตรมาสนั้น
        all_names = []
        for snap in snaps:
            for name, _ in snap["top5"]:
                if name and name not in all_names:
                    all_names.append(name)

        rank_map: dict = {}
        score_map: dict = {}
        for qi, snap in enumerate(snaps):
            name_list = [n for n, _ in snap["top5"]]
            sc_list   = {n: sc for n, sc in snap["top5"]}
            for name in all_names:
                if name not in rank_map:
                    rank_map[name]  = [None] * n_q
                    score_map[name] = [None] * n_q
                if name in name_list:
                    rank_map[name][qi]  = name_list.index(name) + 1
                    score_map[name][qi] = sc_list[name]

        # Header row
        header_f = CTkFrame(self._body, fg_color="transparent")
        header_f.pack(fill="x", padx=16, pady=(12,4))
        col_w = 160
        CTkLabel(header_f, text="Supplier",
                 font=F(size=12, weight="bold"), width=190, anchor="w").pack(side="left")
        for q in quarters:
            CTkLabel(header_f, text=q,
                     font=F(size=12, weight="bold"), text_color=CLR["navy"],
                     width=col_w, anchor="center",
                     fg_color=CLR["blue_lt"], corner_radius=6).pack(side="left", padx=4)

        # Data rows
        rank_colors = ["#F59E0B","#9CA3AF","#D97706","#6B7280","#6B7280","#9CA3AF"]

        for name in all_names:
            ranks  = rank_map[name]
            scores = score_map[name]
            if all(r is None for r in ranks):
                continue

            row_f = CTkFrame(self._body, fg_color=CLR["white"],
                             corner_radius=6, border_width=1,
                             border_color=CLR["border"])
            row_f.pack(fill="x", padx=16, pady=3)

            CTkLabel(row_f, text=name,
                     font=F(size=12, weight="bold"),
                     width=190, anchor="w").pack(side="left", padx=(10,0), pady=8)

            for qi, (r, sc) in enumerate(zip(ranks, scores)):
                cell_f = CTkFrame(row_f, fg_color="transparent", width=col_w)
                cell_f.pack(side="left", padx=4, pady=4)
                cell_f.pack_propagate(False)

                if r is None:
                    CTkLabel(cell_f, text="—",
                             text_color=CLR["gray"], font=F(size=12),
                             anchor="center").pack(expand=True)
                    continue

                # trend vs previous quarter
                if qi > 0:
                    prev_r = ranks[qi-1]
                    if prev_r is None:
                        trend, tc = " ✦", "#7C3AED"
                    elif r < prev_r:
                        trend, tc = f" ▲{prev_r-r}", CLR["green"]
                    elif r > prev_r:
                        trend, tc = f" ▼{r-prev_r}", CLR["red"]
                    else:
                        trend, tc = " —", CLR["gray"]
                else:
                    trend, tc = "", CLR["gray"]

                rank_c = rank_colors[r-1] if r <= len(rank_colors) else CLR["gray"]

                inner = CTkFrame(cell_f, fg_color=CLR["gray_lt"], corner_radius=6)
                inner.pack(expand=True, fill="both", padx=4, pady=2)

                CTkLabel(inner, text=f"#{r}",
                         fg_color=rank_c, text_color=CLR["white"],
                         corner_radius=4, font=F(size=11, weight="bold"),
                         width=28, height=22).pack(side="left", padx=(6,4), pady=4)

                right_f = CTkFrame(inner, fg_color="transparent")
                right_f.pack(side="left", expand=True)
                CTkLabel(right_f, text=f"Win {sc}%",
                         font=F(size=11), anchor="w").pack(anchor="w")
                CTkLabel(right_f, text=trend,
                         font=F(size=10, weight="bold"),
                         text_color=tc, anchor="w").pack(anchor="w")


# =============================================================================
#  MAIN TAB
# =============================================================================
class SuperSupplierTab(CTkFrame):

    def __init__(self, master, app_container=None, current_user="USER_DEMO"):
        super().__init__(master, corner_radius=0, fg_color="transparent")
        self.app_container = app_container
        self.current_user  = current_user
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._debounce_job  = None
        self._refresh_token = 0   # ป้องกัน race condition เมื่อ refresh ซ้อนกัน
        # ผูก app_container กับ DB layer
        global _app_container
        _app_container = app_container
        _apply_ttk_style()
        self._build_toolbar()
        self._build_kpi_bar()
        self._build_table()
        self._refresh_table()
        self._poll_notifications()   # เริ่ม polling bell badge

    # ── Toolbar ───────────────────────────────────────────────────────────────
    def _build_toolbar(self):
        bar = CTkFrame(self, fg_color=CLR["navy"], corner_radius=0)
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_columnconfigure(8, weight=1)
        F  = CTkFont
        kw = dict(fg_color=CLR["white"], text_color=CLR["navy"],
                  button_color="#3B82F6", button_hover_color="#1A56DB",
                  font=F(size=12))

        # ── Row 0: Category / Tier / Availability / Search ───────────────────
        CTkLabel(bar, text="หมวด:", font=F(size=12),
                 text_color="#93C5FD").grid(row=0, column=0, padx=(14, 4), pady=(8,2))
        self._cat_var = tk.StringVar(value="ทุกหมวด")
        _live_cats = db_get_categories()
        CTkOptionMenu(bar, variable=self._cat_var, values=_live_cats,
                      width=130, command=lambda _: self._apply_filter(),
                      **kw).grid(row=0, column=1, padx=(0, 8), pady=(8,2))

        CTkLabel(bar, text="Tier:", font=F(size=12),
                 text_color="#93C5FD").grid(row=0, column=2, padx=(0, 4), pady=(8,2))
        self._tier_var = tk.StringVar(value="ทุก Tier")
        CTkOptionMenu(bar, variable=self._tier_var, values=MOCK_TIERS,
                      width=110, command=lambda _: self._apply_filter(),
                      **kw).grid(row=0, column=3, padx=(0, 8), pady=(8,2))

        CTkLabel(bar, text="สถานะ:", font=F(size=12),
                 text_color="#93C5FD").grid(row=0, column=4, padx=(0, 4), pady=(8,2))
        self._avail_var = tk.StringVar(value="ทุกสถานะ")
        CTkOptionMenu(bar, variable=self._avail_var, values=MOCK_AVAIL,
                      width=130, command=lambda _: self._apply_filter(),
                      **kw).grid(row=0, column=5, padx=(0, 8), pady=(8,2))

        self._search_e = CTkEntry(bar, placeholder_text="🔍  ค้นหาชื่อ / รหัส / พื้นที่...",
                                  font=F(size=12), width=220, height=30,
                                  fg_color=CLR["white"], text_color=CLR["navy"])
        self._search_e.grid(row=0, column=6, padx=(0, 4), pady=(8,2))
        self._search_e.bind("<Return>",     lambda _: self._apply_filter())
        self._search_e.bind("<KP_Enter>",   lambda _: self._apply_filter())
        self._search_e.bind("<KeyRelease>", lambda _: self._debounce())

        CTkButton(bar, text="ค้นหา", width=70, height=28, font=F(size=12),
                  fg_color="#3B82F6", hover_color="#1A56DB",
                  command=self._apply_filter).grid(row=0, column=7, padx=(0, 4), pady=(8,2))
        CTkButton(bar, text="ล้าง", width=56, height=28, font=F(size=12),
                  fg_color="gray40", hover_color="gray30",
                  command=self._clear_filter).grid(row=0, column=8, padx=(0, 14), pady=(8,2), sticky="w")

        # ── Row 1: Source Tag / Credit filter ────────────────────────────────
        CTkLabel(bar, text="Source:", font=F(size=11),
                 text_color="#93C5FD").grid(row=1, column=0, padx=(14, 4), pady=(2,6))
        self._source_var = tk.StringVar(value="ทุก Source")
        CTkOptionMenu(bar, variable=self._source_var, values=MOCK_SOURCE_TAGS,
                      width=130, command=lambda _: self._apply_filter(),
                      **kw).grid(row=1, column=1, padx=(0, 8), pady=(2,6))

        CTkLabel(bar, text="เครดิต:", font=F(size=11),
                 text_color="#93C5FD").grid(row=1, column=2, padx=(0, 4), pady=(2,6))
        self._credit_var = tk.StringVar(value="ทุกเครดิต")
        CTkOptionMenu(bar, variable=self._credit_var, values=MOCK_CREDIT_OPT,
                      width=170, command=lambda _: self._apply_filter(),
                      **kw).grid(row=1, column=3, padx=(0, 8), pady=(2,6))

        # 🔔 Notification bell (row 1, right)
        self._bell_btn = CTkButton(
            bar, text="🔔  0", width=68, height=26,
            fg_color="#3B82F6", hover_color="#1A56DB",
            font=F(size=12, weight="bold"),
            command=self._open_notifications)
        self._bell_btn.grid(row=1, column=7, padx=(0,4), pady=(2,6), sticky="e")

    # ── KPI Bar ───────────────────────────────────────────────────────────────
    def _build_kpi_bar(self):
        kf = CTkFrame(self, fg_color=CLR["gray_lt"], corner_radius=0)
        kf.grid(row=1, column=0, sticky="ew")
        self._kpi_vals = {}
        specs = [
            ("total",       "Supplier ทั้งหมด",   CLR["navy"],  "#F0F4FF"),
            ("tier1",       "Tier 1",              CLR["green"], CLR["green_lt"]),
            ("tier2",       "Tier 2",              CLR["blue"],  CLR["blue_lt"]),
            ("sn",          "SN (รอ Convert)",     CLR["amber"], CLR["amber_lt"]),
            ("blacklist",   "Blacklist",            CLR["red"],   CLR["red_lt"]),
            ("conv_rate",   "Conversion Rate",      CLR["teal"],  CLR["teal_lt"]),
        ]
        for col, (key, title, fg, bg) in enumerate(specs):
            t = CTkFrame(kf, fg_color=bg, corner_radius=8,
                         border_width=1, border_color=CLR["border"])
            t.grid(row=0, column=col, padx=8, pady=8, ipadx=12, ipady=4)
            lbl = CTkLabel(t, text="-", font=CTkFont(size=22, weight="bold"),
                           text_color=fg)
            lbl.pack(pady=(8, 0))
            CTkLabel(t, text=title, font=CTkFont(size=11),
                     text_color=CLR["gray"]).pack(pady=(0, 8))
            self._kpi_vals[key] = lbl

    # ── Table ─────────────────────────────────────────────────────────────────
    def _build_table(self):
        tf = CTkFrame(self, fg_color=CLR["white"], corner_radius=0,
                      border_width=1, border_color=CLR["border"])
        tf.grid(row=2, column=0, sticky="nsew")
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ["supplier_id", "name", "category", "tier",
                "availability", "contact", "phone",
                "score", "credit_days", "note",
                "wh_zone", "wh_coordinates"]
        col_cfg = {
            "supplier_id":    ("รหัส",           140, "center"),
            "name":           ("ชื่อ Supplier",   220, "w"),
            "category":       ("หมวดสินค้า",      100, "center"),
            "tier":           ("Tier",             80, "center"),
            "availability":   ("สถานะ",           100, "center"),
            "contact":        ("ผู้ติดต่อ",        120, "w"),
            "phone":          ("เบอร์ติดต่อ",      120, "center"),
            "score":          ("Score",             70, "center"),
            "credit_days":    ("เครดิต",            80, "center"),
            "note":           ("จุดแข็ง",          200, "w"),
            "wh_zone":        ("โซนที่ตั้ง",       130, "w"),
            "wh_coordinates": ("พิกัด Warehouse",  200, "w"),
        }
        self._tree = ttk.Treeview(tf, columns=cols, show="headings",
                                  style="SSL.Treeview")
        for col in cols:
            lbl, w, anchor = col_cfg[col]
            self._tree.heading(col, text=lbl,
                               command=lambda c=col: self._sort_by(c))
            self._tree.column(col, width=w, anchor=anchor,
                              stretch=(col == "wh_coordinates"))

        self._tree.tag_configure("Tier 1",    background="#F0FDF4")
        self._tree.tag_configure("Tier 2",    background="#EFF6FF")
        self._tree.tag_configure("Tier 3",    background="#FFF7ED")
        self._tree.tag_configure("SN",        background="#FFFBEB")
        self._tree.tag_configure("Blacklist", background="#FFF1F2")
        self._tree.tag_configure("locked",    background="#E0F2FE")
        self._tree.tag_configure("closed",    background="#FEF9C3")

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self._tree.bind("<Double-1>", lambda _: self._open_detail())
        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<Motion>",   self._on_tree_motion)
        self._tree.bind("<Leave>",    self._on_tree_leave)

        # Tooltip window (shared)
        self._tooltip = tk.Toplevel(self)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.withdraw()
        self._tooltip_lbl = tk.Label(
            self._tooltip, text="", justify="left",
            background="#1F2937", foreground="white",
            font=("Tahoma", 11), padx=10, pady=6,
            relief="flat", borderwidth=0)
        self._tooltip_lbl.pack()
        self._tooltip_col  = None   # คอลัมน์ที่ hover อยู่
        self._tooltip_iid  = None   # row iid ที่ hover อยู่
        self._df           = pd.DataFrame()  # cache ของ df ที่แสดงในตาราง

        ctx = tk.Menu(self, tearoff=0)
        ctx.add_command(label="ดู / แก้ไข Profile",  command=self._open_detail)
        ctx.add_separator()
        ctx.add_command(label="Promote → Tier 1",    command=lambda: self._quick_tier("Tier 1"))
        ctx.add_command(label="Demote → Tier 2",     command=lambda: self._quick_tier("Tier 2"))
        ctx.add_command(label="Demote → Tier 3",     command=lambda: self._quick_tier("Tier 3"))
        ctx.add_command(label="Convert SN → SW",     command=self._quick_convert)
        ctx.add_separator()
        ctx.add_command(label="Flag Blacklist",       command=lambda: self._quick_tier("Blacklist"))
        self._ctx = ctx

        # Bottom bar
        bot = CTkFrame(self, fg_color=CLR["white"], corner_radius=0,
                       border_width=1, border_color=CLR["border"])
        bot.grid(row=3, column=0, sticky="ew")

        # ── 1. ปุ่มหลัก (Primary Actions) ให้โชว์ตลอด ──
        CTkButton(bot, text="+ เพิ่ม Supplier ใหม่",
                  fg_color=CLR["navy"], hover_color=CLR["blue"],
                  font=CTkFont(size=13, weight="bold"), height=34,
                  command=self._open_add).pack(side="left", padx=12, pady=8)
                  
        CTkButton(bot, text="แก้ไขที่เลือก",
                  fg_color=CLR["blue"], hover_color="#1e40af", height=34,
                  command=self._open_detail).pack(side="left", padx=(0, 6), pady=8)
                  
        CTkButton(bot, text="Convert SN → SW",
                  fg_color=CLR["amber"], hover_color="#92400E",
                  text_color=CLR["white"], height=34,
                  command=self._quick_convert).pack(side="left", padx=(0, 6), pady=8)
                  
        CTkButton(bot, text="🏆 Top 5 ต่อหมวด",
                  fg_color=CLR["teal"], hover_color="#0F6E56", height=34,
                  command=lambda: Top5View(self)).pack(side="left", padx=(0, 6), pady=8)
                  
        CTkButton(bot, text="⭐ แนะนำ Supplier",
                  fg_color="#D97706", hover_color="#B45309", height=34,
                  command=self._open_suggested).pack(side="left", padx=(0, 6), pady=8)

        # ── 2. สร้าง Custom Dropdown แบบ Modern (แก้ไขใหม่) ──
        # ── 2. สร้าง Custom Dropdown แบบ Modern (แก้ไขใหม่ล่าสุด) ──
        def _show_modern_more_menu():
            popup = tk.Toplevel(self)
            popup.wm_overrideredirect(True)
            popup.configure(bg="white")
            
            # [Fix Bug 1]: แปะฟังก์ชันดัมมี่ เพื่อป้องกัน CustomTkinter แจ้ง Error Scaling
            popup.block_update_dimensions_event = lambda: None
            popup.unblock_update_dimensions_event = lambda: None
            popup.set_scaling = lambda *args, **kwargs: None
            
            # ซ่อนหน้าต่างตอนกำลังโหลดเพื่อป้องกันการกระพริบ
            popup.attributes("-alpha", 0.0)
            
            # [Fix Bug 2]: ล็อกขนาดเป๊ะๆ ป้องกันพื้นที่เหลือด้านล่าง (ปุ่ม 32x7=224 + เส้นคั่น 13 + ขอบ 8 = 245)
            menu_width = 200
            menu_height = 245 
            
            menu_frame = CTkFrame(popup, fg_color="white", corner_radius=0, 
                                  border_width=1, border_color="#D1D5DB")
            menu_frame.pack(fill="both", expand=True)

            CTkFrame(menu_frame, height=4, fg_color="white").pack(fill="x")
            
            def add_menu_btn(text, cmd, icon="", text_color="#1F2937", hover_color="#F3F4F6"):
                btn = CTkButton(menu_frame, text=f"  {icon}   {text}", anchor="w",
                                fg_color="transparent", text_color=text_color, hover_color=hover_color,
                                font=CTkFont(size=13), height=32, corner_radius=0,
                                command=lambda: [popup.destroy(), cmd()])
                btn.pack(fill="x", pady=0)
                return btn

            # เพิ่มรายการเมนู
            add_menu_btn("Quarterly Snapshot", lambda: QuarterlySnapshotPopup(self), icon="📊")
            add_menu_btn("Audit Log",          lambda: AuditLogPopup(self),          icon="📋")
            add_menu_btn("SN Aging",           lambda: SNAgingPopup(self),           icon="⏰")
            add_menu_btn("Win-Loss",           self._open_win_loss,                  icon="📈")
            add_menu_btn("Timeline",           lambda: RankingTimelinePopup(self),   icon="📅")
            
            separator = CTkFrame(menu_frame, height=1, fg_color="#E5E7EB")
            separator.pack(fill="x", pady=6, padx=12)
            
            add_menu_btn("Export",             self._open_export,                    icon="📤")
            add_menu_btn("Bulk Import",        self._open_bulk_import,               icon="📥")

            popup.update_idletasks()
            
            # หาพิกัดเริ่มต้น
            x = btn_more.winfo_rootx()
            y = btn_more.winfo_rooty() + btn_more.winfo_height() + 2
            
            # เช็คขอบจอด้านล่าง หากล้นให้เด้งขึ้นบน
            screen_height = btn_more.winfo_screenheight()
            if (y + menu_height) > (screen_height - 40):
                y = btn_more.winfo_rooty() - menu_height - 2
                
            popup.geometry(f"{menu_width}x{menu_height}+{x}+{y}")
            
            # โชว์หน้าต่างกลับมาตามปกติ
            popup.attributes("-alpha", 1.0)
            
            # หายไปเมื่อคลิกที่อื่น
            popup.bind("<FocusOut>", lambda e: popup.destroy())
            popup.focus_force()

        # ปุ่มกดเพื่อเรียก Custom Dropdown
        btn_more = CTkButton(bot, text="เครื่องมือเพิ่มเติม ▾",
                             fg_color="gray85", hover_color="gray75", text_color="black",
                             height=34, font=CTkFont(size=12, weight="bold"),
                             command=_show_modern_more_menu)
        btn_more.pack(side="left", padx=(6, 6), pady=8)

        # ── 3. Label และ ปุ่ม Refresh (ชิดขวา) ──
        self._row_lbl = CTkLabel(bot, text="", font=CTkFont(size=12), text_color=CLR["gray"])
        self._row_lbl.pack(side="right", padx=12)
        
        CTkButton(bot, text="Refresh", width=80, height=34,
                  fg_color="gray50", hover_color="gray40",
                  command=self._refresh_table).pack(side="right", padx=(0, 8), pady=8)

    # ── Data ──────────────────────────────────────────────────────────────────
    def _apply_filter(self):
        self._refresh_table()

    def _clear_filter(self):
        self._cat_var.set("ทุกหมวด")
        self._tier_var.set("ทุก Tier")
        self._avail_var.set("ทุกสถานะ")
        self._source_var.set("ทุก Source")
        self._credit_var.set("ทุกเครดิต")
        self._search_e.delete(0, "end")
        self._refresh_table()

    def _debounce(self):
        if self._debounce_job:
            self.after_cancel(self._debounce_job)
        self._debounce_job = self.after(350, self._refresh_table)

    def _refresh_table(self):
        import threading

        # เพิ่ม token เพื่อยกเลิก result ที่ค้างจาก request เก่า
        self._refresh_token += 1
        token = self._refresh_token

        self._row_lbl.configure(text="กำลังโหลด...")

        # snapshot filter values ก่อน thread เริ่ม (ป้องกัน StringVar race)
        cat    = self._cat_var.get()
        tier   = self._tier_var.get()
        avail  = self._avail_var.get()
        search = self._search_e.get().strip()
        source = self._source_var.get()
        credit = self._credit_var.get()

        def _load():
            df       = get_suppliers_df(cat, tier, avail, search, source, credit)
            all_sups = db_get_all_suppliers()
            aging    = get_aging_sns(SN_AGING_DAYS)
            audit    = db_get_audit_log()
            self.after(0, lambda: _apply(df, all_sups, aging, audit))

        def _apply(df, all_sups, aging, audit):
            if token != self._refresh_token:
                return  # มี refresh ใหม่กว่าแล้ว ยกเลิก

            for item in self._tree.get_children():
                self._tree.delete(item)

            self._df = df

            avail_map = {"พร้อม": "✓ พร้อม", "สต็อกต่ำ": "⚠ ต่ำ", "ปิดชั่วคราว": "✗ ปิด"}
            src_badge = {"Legacy": "[L]", "Manual": "[M]", "System": "[S]"}

            for _, row in df.iterrows():
                t     = row["tier"]
                av    = row["availability"]
                lock  = "\U0001f512 " if row.get("is_locked") else ""
                src   = src_badge.get(row.get("source_tag",""), "")
                cr    = f"{int(row['credit_days'])} วัน" if row["credit_days"] else "เงินสด"
                av_d  = avail_map.get(av, av)
                if av == "ปิดชั่วคราว" and row.get("reopen_date"):
                    av_d += f" ({row['reopen_date']})"
                coords_raw = row.get("wh_coordinates", "") or ""
                coords_d   = "\U0001f4cd ดูแผนที่" if coords_raw.startswith("http") else coords_raw
                values = (
                    row["supplier_id"],
                    lock + row["name"] + (f"  {src}" if src else ""),
                    row["category"], t, av_d,
                    row["contact"], row.get("phone", "") or "",
                    row["score"], cr, row["note"],
                    row.get("wh_zone", "") or "", coords_d,
                )
                tag = "locked" if row.get("is_locked") else \
                      ("closed" if av == "ปิดชั่วคราว" else t)
                self._tree.insert("", "end", iid=str(int(row["id"])),
                                  values=values, tags=(tag,))

            aging_txt = f"  ⚠ SN aging: {len(aging)}" if aging else ""
            self._row_lbl.configure(text=f"แสดง {len(df)} รายการ{aging_txt}")

            all_df   = pd.DataFrame(all_sups) if all_sups else pd.DataFrame(
                columns=["id","supplier_id","tier"])
            total_sn  = len(all_df[all_df["supplier_id"].str.startswith("SN")]) if not all_df.empty else 0
            converted = sum(1 for l in audit if l.get("action") == "convert")
            conv_pct  = f"{round(converted / max(total_sn + converted, 1) * 100)}%" \
                        if (total_sn + converted) else "0%"

            self._kpi_vals["total"].configure(    text=str(len(all_df)))
            self._kpi_vals["tier1"].configure(    text=str(len(all_df[all_df["tier"] == "Tier 1"])) if not all_df.empty else "0")
            self._kpi_vals["tier2"].configure(    text=str(len(all_df[all_df["tier"] == "Tier 2"])) if not all_df.empty else "0")
            self._kpi_vals["sn"].configure(       text=str(len(all_df[all_df["tier"] == "SN"]))      if not all_df.empty else "0")
            self._kpi_vals["blacklist"].configure(text=str(len(all_df[all_df["tier"] == "Blacklist"])) if not all_df.empty else "0")
            self._kpi_vals["conv_rate"].configure(text=conv_pct)

        threading.Thread(target=_load, daemon=True).start()

    _sort_state = {}

    def _sort_by(self, col):
        asc = not self._sort_state.get(col, True)
        self._sort_state[col] = asc
        df = get_suppliers_df(self._cat_var.get(), self._tier_var.get(),
                              self._avail_var.get(), self._search_e.get().strip(),
                              self._source_var.get(), self._credit_var.get())
        if col in df.columns:
            df = df.sort_values(col, ascending=asc).reset_index(drop=True)
        self._df = df  # cache for tooltip lookup
        avail_map = {"พร้อม": "✓ พร้อม", "สต็อกต่ำ": "⚠ ต่ำ", "ปิดชั่วคราว": "✗ ปิด"}
        src_badge = {"Legacy": "[L]", "Manual": "[M]", "System": "[S]"}
        for item in self._tree.get_children():
            self._tree.delete(item)
        for _, row in df.iterrows():
            tier  = row["tier"]
            avail = row["availability"]
            lock  = "🔒 " if row.get("is_locked") else ""
            src   = src_badge.get(row.get("source_tag",""), "")
            cr    = f"{int(row['credit_days'])} วัน" if row["credit_days"] else "เงินสด"
            avail_display = avail_map.get(avail, avail)
            if avail == "ปิดชั่วคราว" and row.get("reopen_date"):
                avail_display += f" ({row['reopen_date']})"
            # พิกัด: ถ้าเป็น URL ยาว → ย่อแสดงเป็น "📍 ดูแผนที่" แทน
            coords_raw = row.get("wh_coordinates", "")
            coords_display = "📍 ดูแผนที่" if coords_raw.startswith("http") else coords_raw
            values = (row["supplier_id"], lock + row["name"] + (f"  {src}" if src else ""),
                      row["category"], tier, avail_display, row["contact"],
                      row.get("phone", ""), row["score"],
                      cr, row["note"], row.get("wh_zone", ""), coords_display)
            tag = "locked" if row.get("is_locked") else \
                  ("closed" if avail == "ปิดชั่วคราว" else tier)
            self._tree.insert("", "end", iid=str(int(row["id"])),
                              values=values, tags=(tag,))

    def _open_win_loss(self):
        sup = self._get_selected()
        if sup:
            WinLossAnalysisPopup(self, sup)

    def _open_export(self):
        df = get_suppliers_df(
            self._cat_var.get(), self._tier_var.get(),
            self._avail_var.get(), self._search_e.get().strip(),
            self._source_var.get(), self._credit_var.get())
        ExportPopup(self, df)

    def _open_bulk_import(self):
        BulkImportPopup(self, on_success=self._refresh_table,
                        current_user=self.current_user)

    def _open_suggested(self):
        cat = self._cat_var.get()
        SuggestedSupplierPopup(self, category=cat,
                               on_select=self._on_suggested_select)

    def _on_suggested_select(self, sup_dict):
        """Callback เมื่อ user เลือก Supplier จาก SuggestedSupplierPopup"""
        name = sup_dict.get("name", "")
        phone = sup_dict.get("phone", "")
        from tkinter import messagebox as mb
        mb.showinfo("เลือก Supplier แล้ว",
                    f"เลือก: {name}\nโทร: {phone}\n\n"
                    f"(Phase 2: ระบบจะส่งข้อมูลนี้ไปที่ฟอร์มสอบราคาโดยอัตโนมัติ)",
                    parent=self)

    def _open_notifications(self):
        NotificationCenterPopup(self, on_read=self._update_bell)

    def _update_bell(self):
        unread = sum(1 for n in MOCK_NOTIFICATIONS if not n.get("read", False))
        color  = CLR["red"] if unread else "#3B82F6"
        self._bell_btn.configure(
            text=f"🔔  {unread}" if unread else "🔔",
            fg_color=color,
            hover_color="#991B1B" if unread else "#1A56DB")

    def _poll_notifications(self):
        """อัปเดต bell badge ทุก 5 วินาที"""
        try:
            self._update_bell()
            self.after(5000, self._poll_notifications)
        except Exception:
            pass

    # ── Interactions ──────────────────────────────────────────────────────────
    def _on_tree_motion(self, event):
        """แสดง tooltip Score breakdown เมื่อ hover คอลัมน์ Score"""
        region = self._tree.identify_region(event.x, event.y)
        if region != "cell":
            self._tooltip.withdraw()
            return

        col_id = self._tree.identify_column(event.x)
        iid    = self._tree.identify_row(event.y)

        # เช็คว่า hover คอลัมน์ score — ใช้ column list เดียวกับที่ตารางสร้าง
        _tree_cols = ["supplier_id", "name", "category", "tier", "availability",
                      "contact", "phone", "score", "credit_days",
                      "note", "wh_zone", "wh_coordinates"]
        try:
            col_idx  = int(col_id.replace("#", "")) - 1
            col_name = _tree_cols[col_idx] if col_idx < len(_tree_cols) else ""
        except Exception:
            col_name = ""

        if col_name != "score" or not iid:
            self._tooltip.withdraw()
            self._tooltip_iid = None
            return

        # ถ้า row เดิมอยู่แล้วไม่ต้อง re-render
        if iid == self._tooltip_iid:
            self._tooltip.wm_geometry(f"+{event.x_root+12}+{event.y_root-10}")
            return

        self._tooltip_iid = iid
        try:
            row_id = int(iid)
            # ใช้ df ที่ merge กับ benchmark แล้ว (เหมือนที่ตารางแสดง)
            df = self._df
            if df.empty or "id" not in df.columns:
                self._tooltip.withdraw()
                return
            rows = df[df["id"] == row_id]
            if rows.empty:
                self._tooltip.withdraw()
                return
            sup = rows.iloc[0]
            p  = int(sup.get("price_score",   0) or 0)
            w  = int(sup.get("win_pct",       0) or 0)
            sv = int(sup.get("service_score", 0) or 0)
            s  = int(sup.get("sla_score",     0) or 0)
            q  = int(sup.get("quality_score", 0) or 0)
            total = int(sup.get("score", calc_score(sup.to_dict())))
            text = (f"Score Breakdown\n"
                    f"─────────────────\n"
                    f"ราคา    (×20%)  {p:>3}  →  {round(p*0.20):>3}\n"
                    f"สต็อก   (×20%)  {w:>3}  →  {round(w*0.20):>3}\n"
                    f"บริการ  (×20%)  {sv:>3}  →  {round(sv*0.20):>3}\n"
                    f"SLA     (×20%)  {s:>3}  →  {round(s*0.20):>3}\n"
                    f"คุณภาพ  (×20%)  {q:>3}  →  {round(q*0.20):>3}\n"
                    f"─────────────────\n"
                    f"รวม              {total:>3}")
            self._tooltip_lbl.configure(text=text)
            self._tooltip.wm_geometry(f"+{event.x_root+12}+{event.y_root-10}")
            self._tooltip.deiconify()
            self._tooltip.lift()
        except Exception:
            self._tooltip.withdraw()

    def _on_tree_leave(self, _=None):
        self._tooltip.withdraw()
        self._tooltip_iid = None

    def _get_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("ไม่ได้เลือก", "กรุณาเลือก Supplier ก่อน", parent=self)
            return None
        row_id = int(sel[0])
        _all = db_get_all_suppliers()
        hits = [s for s in _all if s["id"] == row_id]
        return hits[0] if hits else None

    def _on_right_click(self, e):
        row = self._tree.identify_row(e.y)
        if row:
            self._tree.selection_set(row)
            self._ctx.tk_popup(e.x_root, e.y_root)

    def _open_detail(self):
        sup = self._get_selected()
        if sup:
            SupplierDetailPopup(self, sup,
                                on_save=lambda _: self._refresh_table(),
                                current_user=self.current_user)

    def _open_add(self):
        AddSupplierPopup(self, on_success=self._refresh_table,
                         current_user=self.current_user)

    def _quick_tier(self, new_tier: str):
        sup = self._get_selected()
        if not sup:
            return
        if sup.get("is_locked") and new_tier != "Tier 1":
            messagebox.showwarning("ล็อกอยู่",
                                   "Supplier นี้ถูก Lock Tier 1 โดย Manager\n"
                                   "กรุณาปลด Lock ก่อนทำการ Demote", parent=self)
            return

        # D3: Blacklist ต้องใส่เหตุผลก่อนเสมอ
        if new_tier == "Blacklist":
            def _do_blacklist(reason):
                updated = dict(sup)
                updated["tier"] = "Blacklist"
                updated["blacklist_reason"] = reason
                updated["note"] = reason
                ok = db_save_supplier(updated, action="blacklist", user=self.current_user)
                if ok:
                    _push_noti(f"Flag Blacklist: '{sup['name']}' โดย {self.current_user} — {reason}", "high")
                else:
                    messagebox.showerror("ผิดพลาด", "บันทึกไม่สำเร็จ", parent=self)
                self._refresh_table()
                self._update_bell()
            BlacklistReasonPopup(self, sup, on_confirm=_do_blacklist)
            return

        # B6: ขอ Remark สำหรับ Promote/Demote
        remark_popup = CTkToplevel(self)
        remark_popup.title("ระบุเหตุผล Manual Adjust Tier")
        _place_popup(remark_popup, 420, 200)
        remark_popup.resizable(False, False)
        remark_popup.grid_columnconfigure(0, weight=1)
        remark_popup.transient(self)
        remark_popup.grab_set()
        F = CTkFont
        CTkLabel(remark_popup,
                 text=f"เปลี่ยน Tier ของ '{sup['name']}' → {new_tier}",
                 font=F(size=13, weight="bold")).grid(
            row=0, column=0, padx=20, pady=(16, 6), sticky="w")
        CTkLabel(remark_popup, text="เหตุผล / Remark (ไม่บังคับ):",
                 font=F(size=12), text_color=CLR["gray"]).grid(
            row=1, column=0, padx=20, pady=(0, 4), sticky="w")
        remark_entry = CTkEntry(remark_popup, font=F(size=13), height=34,
                                placeholder_text="เช่น ราคาเริ่มหลุดเกณฑ์, ข่าวสารวงใน...")
        remark_entry.grid(row=2, column=0, padx=20, sticky="ew")

        def _confirm_tier():
            remark = remark_entry.get().strip()
            updated = dict(sup)
            updated["tier"] = new_tier
            if remark:
                updated["note"] = (updated.get("note","") + f" | {remark}").strip(" |")
            ok = db_save_supplier(updated, action="tier", user=self.current_user)
            if ok:
                ntype = "medium" if new_tier in ("Tier 1","Tier 2") else "high"
                _push_noti(f"เปลี่ยน Tier: '{sup['name']}' → {new_tier} โดย {self.current_user}", ntype)
            else:
                messagebox.showerror("ผิดพลาด", "บันทึกไม่สำเร็จ", parent=self)
            remark_popup.destroy()
            self._refresh_table()
            self._update_bell()

        bf = CTkFrame(remark_popup, fg_color="transparent")
        bf.grid(row=3, column=0, padx=20, pady=14, sticky="e")
        CTkButton(bf, text="ยกเลิก", fg_color="gray50", hover_color="gray40",
                  width=80, command=remark_popup.destroy).pack(side="right", padx=(8, 0))
        CTkButton(bf, text="ยืนยัน", fg_color=CLR["blue"], hover_color="#1e40af",
                  width=100, command=_confirm_tier).pack(side="right")

    def _quick_convert(self):
        sup = self._get_selected()
        if not sup:
            return
        if sup.get("tier") != "SN":
            messagebox.showinfo("ไม่ใช่ SN",
                                "สามารถ Convert ได้เฉพาะ Supplier ที่มี Tier เป็น SN เท่านั้น",
                                parent=self)
            return
        ConvertToSWPopup(self, sup, on_success=self._refresh_table,
                         current_user=self.current_user)