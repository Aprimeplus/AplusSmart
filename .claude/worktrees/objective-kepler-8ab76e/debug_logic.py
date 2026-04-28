import pandas as pd
import numpy as np

# ==============================================================================
# 1. จำลองฟังก์ชัน Logic จากไฟล์ business_logic.py (ที่เราเพิ่งแก้ล่าสุด)
# ==============================================================================
def prepare_and_calculate_profit(df):
    # (Copy Logic ล่าสุดมาวางตรงนี้เพื่อให้มั่นใจว่าทดสอบ Logic เดียวกันเป๊ะๆ)
    
    # 1. Fill NA
    cols_to_fix = [
        'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee', 
        'difference_amount', 'payment_before_vat', 'payment_no_vat', 
        'shipping_cost', 'relocation_cost', 'so_cutting_rev', 'so_service_rev',      
        'po_shipping_site_cost', 'po_shipping_stock_cost', 'po_cutting_cost', 'po_service_cost'
    ]
    for col in cols_to_fix:
        if col not in df.columns: df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    # 2. Multiplier
    if 'cost_multiplier' in df.columns:
         df['cost_multiplier'] = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
    else:
         df['cost_multiplier'] = 1.03

    # =========================================
    # [🔥 CORE LOGIC ที่เราต้องการตรวจสอบ] 
    # =========================================
    
    # 1. ค่ารถ Site (Matching): จ่ายจริง - เก็บลูกค้า (ถ้าจ่ายเกิน = หัก)
    deduct_site_shipping = (df['po_shipping_site_cost'] - df['shipping_cost']).clip(lower=0)
    
    # 2. ค่าย้าย Stock (Direct Deduction): 
    # EXP-0174 เป็นสินค้าปกติ รายรับรวมใน Sales แล้ว
    # ดังนั้นต้นทุนส่วนนี้ (po_shipping_stock_cost) หักออกเต็มจำนวน
    deduct_stock_shipping = df['po_shipping_stock_cost']
    
    # 3. ค่าตัด/เจาะ (Matching):
    deduct_cutting = (df['po_cutting_cost'] - df['so_cutting_rev']).clip(lower=0)
    
    # 4. บริการอื่น (Matching):
    deduct_service = (df['po_service_cost'] - df['so_service_rev']).clip(lower=0)

    # --- ต้นทุนสินค้าหลัก ---
    # หักต้นทุนพิเศษออกจาก "ต้นทุนรวม (Grid PO)" ก่อนคูณ 1.03
    # (สมมติว่าใน Grid PO รวมค่าตัดเจาะมาแล้ว แต่ค่าขนส่งแยกอยู่ช่องล่าง)
    adjusted_po_cost = df['final_cost_amount'] - df['po_cutting_cost'] - df['po_service_cost']
    adjusted_po_cost = adjusted_po_cost.clip(lower=0)

    # คำนวณต้นทุนรวม Margin
    main_cost_calculated = (adjusted_po_cost * df['cost_multiplier'])
    
    # --- กำไรสุทธิ ---
    df['profit'] = (df['sales_service_amount'] - main_cost_calculated) \
                   + df['difference_amount'] \
                   - deduct_site_shipping \
                   - deduct_stock_shipping \
                   - deduct_cutting \
                   - deduct_service
    
    # เก็บค่าต่างๆ ไว้ตรวจสอบ (เพื่อการ Debug)
    df['_debug_deduct_site'] = deduct_site_shipping
    df['_debug_deduct_stock'] = deduct_stock_shipping
    df['_debug_deduct_cutting'] = deduct_cutting
    df['_debug_main_cost_calc'] = main_cost_calculated

    return df

# ==============================================================================
# 2. สร้างข้อมูลทดสอบ (Mock Data)
# ==============================================================================
print("\n" + "="*80)
print("🛠️  START DEBUGGING COMMISSION LOGIC")
print("="*80)

# สร้าง DataFrame จำลองเหตุการณ์ต่างๆ
data = [
    # Case 1: ปกติ (ขาย 10,000 ทุน 5,000)
    {
        "case": "1. เคสปกติ",
        "so_number": "SO-001",
        "sales_service_amount": 10000, 
        "final_cost_amount": 5000,   # ทุนสินค้า
        "cost_multiplier": 1.03
    },
    # Case 2: มีค่ารถเข้า Site (EXP-0006) -> ลูกค้าจ่าย 1000, เราจ่ายจริง 1200 (ขาดทุนค่ารถ 200)
    {
        "case": "2. ขาดทุนค่ารถ Site",
        "so_number": "SO-002",
        "sales_service_amount": 10000,
        "final_cost_amount": 5000,
        "shipping_cost": 1000,           # SO Revenue (เก็บลูกค้า)
        "po_shipping_site_cost": 1200,   # PO Cost (จ่ายจริง)
    },
    # Case 3: ค่าย้าย (EXP-0174) -> เป็น Normal Product
    # - ยอดขาย 20,000 (รวมค่าย้ายแล้ว)
    # - มีค่ารถขนย้ายจริง 500 บาท (ใส่ในช่องค่าส่งเข้า Stock)
    {
        "case": "3. ค่าย้าย (Normal Product)",
        "so_number": "SO-003",
        "sales_service_amount": 20000, # รวมรายได้ค่าย้ายแล้ว
        "relocation_cost": 2000,       # (field นี้มีค่า แต่ไม่ถูกใช้ Matching)
        "final_cost_amount": 10000,
        "po_shipping_stock_cost": 500, # จ่ายค่ารถย้ายของจริง
    },
    # Case 4: ค่าตัด/เจาะ -> เก็บลูกค้า 500, จ่ายโรงงาน 500 (เท่าทุน)
    {
        "case": "4. ค่าตัด/เจาะ (เท่าทุน)",
        "so_number": "SO-004",
        "sales_service_amount": 5000,
        "final_cost_amount": 3000,     # ในนี้รวมค่าตัด 500 แล้ว
        "so_cutting_rev": 500,
        "po_cutting_cost": 500
    },
    # Case 5: ค่าตัด/เจาะ -> เก็บลูกค้า 500, จ่ายโรงงาน 800 (ขาดทุน 300)
    {
        "case": "5. ค่าตัด/เจาะ (ขาดทุน)",
        "so_number": "SO-005",
        "sales_service_amount": 5000,
        "final_cost_amount": 3300,     # ในนี้รวมค่าตัด 800 แล้ว
        "so_cutting_rev": 500,
        "po_cutting_cost": 800
    }
]

df_test = pd.DataFrame(data)

# ==============================================================================
# 3. รันการคำนวณและแสดงผล (Report)
# ==============================================================================
df_result = prepare_and_calculate_profit(df_test)

for _, row in df_result.iterrows():
    print(f"\n📌 {row['case']} (SO: {row['so_number']})")
    print("-" * 50)
    
    # ข้อมูลดิบ
    print(f"   💰 ยอดขาย (Sales): {row['sales_service_amount']:,.2f}")
    print(f"   📦 ต้นทุนสินค้า (Cost): {row['final_cost_amount']:,.2f}")
    print(f"   ✖️ Multiplier: {row['cost_multiplier']}")
    
    # 1. ตรวจสอบการคำนวณต้นทุนหลัก
    # สูตร: (Cost - Cutting - Service) * 1.03
    base_cost = row['final_cost_amount'] - row['po_cutting_cost'] - row['po_service_cost']
    print(f"   🧮 ฐานต้นทุน (หัก Special Cost): {base_cost:,.2f}")
    print(f"   📉 ต้นทุนรวม Margin (x1.03): {row['_debug_main_cost_calc']:,.2f}")
    
    # 2. ตรวจสอบ Matching & Deductions
    if row['po_shipping_site_cost'] > 0 or row['shipping_cost'] > 0:
        loss = row['_debug_deduct_site']
        print(f"   🚚 ค่ารถ Site: รับ {row['shipping_cost']:,.0f} | จ่าย {row['po_shipping_site_cost']:,.0f} | -> หักกำไร: {loss:,.2f}")
    
    if row['po_shipping_stock_cost'] > 0:
        loss = row['_debug_deduct_stock']
        print(f"   🏭 ค่าย้าย (Stock): จ่ายจริง {row['po_shipping_stock_cost']:,.0f} | -> หักกำไรทันที: {loss:,.2f}")
        
    if row['po_cutting_cost'] > 0:
        loss = row['_debug_deduct_cutting']
        print(f"   ✂️ ค่าตัด/เจาะ: รับ {row['so_cutting_rev']:,.0f} | จ่าย {row['po_cutting_cost']:,.0f} | -> หักกำไรส่วนเกิน: {loss:,.2f}")
        
    # 3. สรุปกำไร
    expected_profit = row['sales_service_amount'] - row['_debug_main_cost_calc'] - row['_debug_deduct_site'] - row['_debug_deduct_stock'] - row['_debug_deduct_cutting']
    print("-" * 50)
    print(f"   ✅ กำไรสุทธิ (Profit): {row['profit']:,.2f}")
    
    # Double Check
    if abs(row['profit'] - expected_profit) < 0.01:
        print("      [CHECK: OK] สูตรคำนวณถูกต้อง")
    else:
        print(f"      [CHECK: FAIL] คำนวณไม่ตรง! (Expected: {expected_profit})")

print("\n" + "="*80)

def prepare_and_calculate_profit(df):
    # 1. Fill NA
    cols_to_fix = [
        'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee', 
        'difference_amount', 'payment_before_vat', 'payment_no_vat', 
        'shipping_cost', 'relocation_cost', 'so_cutting_rev', 'so_service_rev',      
        'po_shipping_site_cost', 'po_shipping_stock_cost', 'po_cutting_cost', 'po_service_cost'
    ]
    for col in cols_to_fix:
        if col not in df.columns: df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    # 2. Multiplier
    if 'cost_multiplier' in df.columns:
         df['cost_multiplier'] = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
    else:
         df['cost_multiplier'] = 1.03

    # --- CORE LOGIC ---
    # 1. ค่ารถ Site (Matching)
    deduct_site_shipping = (df['po_shipping_site_cost'] - df['shipping_cost']).clip(lower=0)
    
    # 2. ค่าย้าย Stock (Direct Deduction - เพราะรายรับรวมในยอดขายแล้ว)
    deduct_stock_shipping = df['po_shipping_stock_cost']
    
    # 3. ค่าตัด/เจาะ (Matching)
    deduct_cutting = (df['po_cutting_cost'] - df['so_cutting_rev']).clip(lower=0)
    
    # 4. บริการอื่น (Matching)
    deduct_service = (df['po_service_cost'] - df['so_service_rev']).clip(lower=0)

    # --- ต้นทุนสินค้าหลัก ---
    adjusted_po_cost = df['final_cost_amount'] - df['po_cutting_cost'] - df['po_service_cost']
    adjusted_po_cost = adjusted_po_cost.clip(lower=0)

    # คำนวณต้นทุนรวม Margin
    main_cost_calculated = (adjusted_po_cost * df['cost_multiplier'])
    
    # --- กำไรสุทธิ ---
    df['profit'] = (df['sales_service_amount'] - main_cost_calculated) \
                   + df['difference_amount'] \
                   - deduct_site_shipping \
                   - deduct_stock_shipping \
                   - deduct_cutting \
                   - deduct_service
    return df

# ==============================================================================
# 2. ข้อมูลทดสอบ (Test Data)
# ==============================================================================
data = [
    # Case 3: ค่าย้าย (พระเอกของเรา)
    {
        "case": "3. ค่าย้าย (Normal Product)",
        "so_number": "SO-003",
        "sales_service_amount": 20000, # รวมรายได้ค่าย้ายแล้ว
        "relocation_cost": 2000,       # (field นี้มีค่า แต่ไม่ถูกใช้ Matching)
        "final_cost_amount": 10000,
        "po_shipping_stock_cost": 500, # จ่ายค่ารถย้ายของจริง (ต้องถูกหักออก)
    }
]

df_test = pd.DataFrame(data)
df_result = prepare_and_calculate_profit(df_test)

for _, row in df_result.iterrows():
    print(f"\n📌 {row['case']}")
    print("-" * 30)
    print(f"💰 ยอดขาย: {row['sales_service_amount']:,.2f}")
    print(f"📉 ต้นทุน (x1.03): {(10000*1.03):,.2f}")
    print(f"🚚 หักค่าย้าย: {row['po_shipping_stock_cost']:,.2f}")
    print(f"✅ กำไรสุทธิ: {row['profit']:,.2f}")
    print("-" * 30)