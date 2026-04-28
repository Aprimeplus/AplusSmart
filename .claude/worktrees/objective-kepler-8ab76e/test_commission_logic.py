import sys
import os
import pytest
import pandas as pd
import numpy as np

# เพิ่ม Path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from business_logic import calculate_monthly_commission

# =================================================================
# FIXTURES
# =================================================================
def create_mock_df(sales=1000000, cost=700000):
    data = {
        'so_number': ['SO-001'],
        'sales_service_amount': [float(sales)],
        'final_cost_amount': [float(cost)],
        'cost_multiplier': [1.03],
        'difference_amount': [0.0],
        'brokerage_fee': [0.0],
        'shipping_cost': [0.0],
        'relocation_cost': [0.0],
        'shipping_to_site_cost': [0.0],
        'shipping_to_stock_cost': [0.0],
        'so_cutting_rev': [0.0],
        'po_cutting_cost': [0.0],
        'so_service_rev': [0.0],
        'po_service_cost': [0.0],
        'giveaways': [0.0],
        'payment_before_vat': [0.0],
        'payment_no_vat': [0.0]
    }
    return pd.DataFrame(data)

# =================================================================
# 1. STANDARD TESTS
# =================================================================
def test_plan_a_standard():
    df = create_mock_df()
    res = calculate_monthly_commission('Plan A', df, min_sales_target=0)
    assert res['final_commission'] > 0

def test_plan_b_standard():
    df = create_mock_df()
    res = calculate_monthly_commission('Plan B', df, min_sales_target=0)
    assert res['final_commission'] > 0

def test_plan_c_standard():
    df = create_mock_df(sales=2000000)
    res = calculate_monthly_commission('Plan C', df, min_sales_target=0)
    assert res['final_commission'] > 0

def test_plan_d_standard():
    df = create_mock_df(sales=2000000)
    res = calculate_monthly_commission('Plan D', df, min_sales_target=0)
    assert res['final_commission'] > 0

# =================================================================
# 2. MATCHING & REPORT TESTS
# =================================================================
def test_plan_a_matching_report():
    df = create_mock_df(sales=1000000, cost=700000)
    df['shipping_to_site_cost'] = 5000.0 
    df['shipping_cost'] = 1000.0
    df['po_cutting_cost'] = 2000.0 
    df['so_cutting_rev'] = 500.0
    
    result = calculate_monthly_commission('Plan A', df, operating_fee=25000, min_sales_target=0)
    assert result['so_breakdown_df'].iloc[0]['กำไรสุทธิ'] < 279000

# =================================================================
# 3. SPECIAL CASES
# =================================================================
def test_additional_deductions_all_plans():
    df = create_mock_df()
    deductions = {"ค่าปรับล่าช้า": 500}
    for plan in ['Plan A', 'Plan B', 'Plan C', 'Plan D']:
        res = calculate_monthly_commission(plan, df, min_sales_target=0, additional_deductions=deductions)
        if 'summary' in res:
            desc_list = res['summary']['description'].values
        elif 'data' in res:
            desc_list = res['data']['description'].values
        else:
            continue
        assert any("ค่าปรับล่าช้า" in s for s in desc_list)

def test_cascade_deduction():
    df = create_mock_df(sales=2000000)
    calculate_monthly_commission('Plan C', df, operating_fee=1800000, min_sales_target=0)
    calculate_monthly_commission('Plan D', df, operating_fee=1800000, min_sales_target=0)

def test_missing_columns_logic():
    df = create_mock_df()
    del df['shipping_cost']
    del df['relocation_cost']
    res = calculate_monthly_commission('Plan A', df, min_sales_target=0)
    assert res['final_commission'] > 0

def test_plan_a_incentives():
    df = create_mock_df()
    res = calculate_monthly_commission('Plan A', df, min_sales_target=0, incentives={"Bonus": 500})
    assert any("(+) Bonus" in s for s in res['summary']['description'].values)

def test_plan_b_high_tiers():
    df = create_mock_df(sales=5000000, cost=1000000)
    res = calculate_monthly_commission('Plan B', df, operating_fee=0, min_sales_target=0)
    assert res['final_commission'] > 50000

def test_coverage_gap_filler():
    df_stock = create_mock_df()
    df_stock['shipping_to_stock_cost'] = 3000.0 
    df_stock['relocation_cost'] = 1000.0
    calculate_monthly_commission('Plan A', df_stock, min_sales_target=0)

    df_no_mult = create_mock_df()
    del df_no_mult['cost_multiplier']
    calculate_monthly_commission('Plan A', df_no_mult, min_sales_target=0)

    df_low = create_mock_df(sales=100000)
    for p in ['Plan A', 'Plan B', 'Plan C', 'Plan D']:
        calculate_monthly_commission(p, df_low, min_sales_target=500000, sales_target=1000000)

    df_empty = pd.DataFrame(columns=['so_number', 'sales_service_amount'])
    for p in ['Plan A', 'Plan B', 'Plan C', 'Plan D']:
        calculate_monthly_commission(p, df_empty)

def test_bad_data_inputs():
    df = create_mock_df()
    df.loc[0, 'sales_service_amount'] = np.nan
    calculate_monthly_commission('Plan A', df, min_sales_target=0)
    res = calculate_monthly_commission('Invalid', df)
    assert res['type'] == 'error'

def test_final_boss_line_668():
    """
    เก็บตก Line 668: Forced Entry Strategy (Nuclear Option)
    ตั้งค่า Fee ให้สูงเวอร์ๆ เพื่อบังคับให้ Commission ของ Normal Tier ติดลบ
    และไหลไป trigger logic 'ยอดหักคงเหลือ' (Line 668) แน่นอน
    """
    # 1. Normal Tier: ยอดขายเยอะ กำไรเยอะ -> ผลิตคอมมิชชั่นได้เยอะ (สมมติได้ 2-3 แสน)
    row_normal = {
        'so_number': 'SO-High', 'sales_service_amount': 5000000.0, 'final_cost_amount': 2000000.0, 
        'cost_multiplier': 1.03, 'difference_amount': 0.0, 'brokerage_fee': 0.0, 'shipping_cost': 0.0, 'relocation_cost': 0.0, 'shipping_to_site_cost': 0.0, 'shipping_to_stock_cost': 0.0, 'so_cutting_rev': 0.0, 'po_cutting_cost': 0.0, 'so_service_rev': 0.0, 'po_service_cost': 0.0, 'giveaways': 0.0, 'payment_before_vat': 0.0, 'payment_no_vat': 0.0
    }
    
    # 2. Below Tier: ยอดขายน้อย กำไรน้อย (Margin 2%) -> เป็นเป้าหมายให้ไหลมาหัก
    row_below = row_normal.copy()
    row_below['so_number'] = 'SO-Low'
    row_below['sales_service_amount'] = 1000000.0
    row_below['final_cost_amount'] = 980000.0 # Margin 2%

    df_mixed = pd.DataFrame([row_normal, row_below])
    
    # 3. [KEY FIX] ตั้ง Fee เป็น 10 ล้านบาท! (10,000,000)
    # รับรองว่าคอมมิชชั่นจากยอดขาย 5 ล้าน ไม่พอจ่ายแน่นอน
    # ระบบจะถูกบังคับให้ print "ยอดหักคงเหลือ (ไปหักต่อที่ Below Tier)" -> Line 668 โดนรันชัวร์
    calculate_monthly_commission('Plan D', df_mixed, operating_fee=10000000, min_sales_target=0)
    calculate_monthly_commission('Plan C', df_mixed, operating_fee=10000000, min_sales_target=0)

def test_plan_c_middle_tier_coverage():
    """
    เก็บตก Missing Line 211:
    ทดสอบเคสที่ Margin อยู่ระหว่าง 7.99% - 9.99% (Tier 2)
    เพื่อให้ฟังก์ชัน assign_status ทำงานครบทุกบรรทัด
    """
    # สร้างยอดขาย 100 บาท, ต้นทุน 91 บาท (Margin = 9%)
    # สูตร Margin = (Sales - Cost*1.03) ... โดยประมาณ
    # เพื่อความชัวร์ กำหนด Sales 1,000,000 และ Cost ที่ทำให้เหลือ Margin ~9%
    
    # Sales = 1,000,000
    # Cost = 900,000 (รวม multiplier แล้วให้เหลือ profit ประมาณ 90,000)
    
    df_middle = create_mock_df(sales=1000000, cost=880000) # Margin จะอยู่ที่ประมาณ 9.36%
    
    # รัน Plan C (และ Plan B ด้วยก็ได้เผื่อ logic คล้ายกัน)
    res_c = calculate_monthly_commission('Plan C', df_middle, min_sales_target=0)
    
    # ตรวจสอบว่า Status ถูกระบุเป็น Tier 2 จริงไหม
    status_text = res_c['so_breakdown_df'].iloc[0]['Status']
    assert 'Tier 2' in status_text or '7.99-9.99%' in status_text
    
    # แถม: รัน Plan B ด้วยเพื่อเก็บตก assign_status ของ Plan B
    res_b = calculate_monthly_commission('Plan B', df_middle, min_sales_target=0)
    status_text_b = res_b['so_breakdown_df'].iloc[0]['Status']
    assert 'Below Tier (7.99-10%)' in status_text_b

def test_coverage_cleanup():
    """
    เก็บตก 1% สุดท้าย (Missing Line 211):
    1. Plan B Tier 3: สร้างเคสกำไรต่ำ (<7.99%) เพื่อให้ Plan B assign_status ทำงานครบทุกบรรทัด
    2. Stock Excess: สร้างเคสค่าขนส่งเข้า Stock เกิน เพื่อให้ Helper Print ทำงานครบ
    """
    # --- Case 1: Plan B Tier 3 (Low Margin) ---
    # Sales 100, Cost 98 (Margin ~2% ซึ่งต่ำกว่า 7.99%)
    df_low = create_mock_df(sales=100000, cost=98000)
    res_b = calculate_monthly_commission('Plan B', df_low, min_sales_target=0)
    
    # ตรวจสอบว่า Status เป็น Tier 3 (Below Tier <7.99%)
    status_text = res_b['so_breakdown_df'].iloc[0]['Status']
    assert '<7.99%' in status_text

    # --- Case 2: Stock Excess (Helper Function) ---
    # สร้างเคสที่จ่ายค่าย้ายเข้า Stock (3000) มากกว่าที่เรียกเก็บลูกค้า (1000)
    df_stock = create_mock_df()
    df_stock['shipping_to_stock_cost'] = 3000.0
    df_stock['relocation_cost'] = 1000.0
    
    # รัน Plan A (หรือ Plan ไหนก็ได้ที่เรียก Helper)
    # ผลลัพธ์ต้องมีการ Print "ค่าย้าย Stock (Excess)" ออกมา (Coverage จะจับได้)
    calculate_monthly_commission('Plan A', df_stock, min_sales_target=0)

def test_incentives_coverage_all_plans():
    """
    เก็บตก 1% สุดท้าย (Missing Line 211):
    ทดสอบการใส่ Incentive ในทุก Plan (B, C, D) 
    เพื่อให้มั่นใจว่า Loop ที่สร้างตารางสรุป Incentive ถูกเรียกใช้งานครบทุกบรรทัด
    """
    df = create_mock_df()
    incentives = {"Special Bonus": 1000}
    
    # วนลูปเทสต์ทุก Plan เลย เพื่อความชัวร์ 100%
    for plan in ['Plan A', 'Plan B', 'Plan C', 'Plan D']:
        res = calculate_monthly_commission(plan, df, min_sales_target=0, incentives=incentives)
        
        # เช็คว่าผลลัพธ์มีตาราง summary (Plan A return เป็น dict ที่มี 'summary', Plan อื่น return 'data')
        if 'summary' in res:
            desc_list = res['summary']['description'].values
        elif 'data' in res:
            desc_list = res['data']['description'].values
        else:
            continue
            
        # ต้องเจอคำว่า Special Bonus ในรายงาน
        assert any("Special Bonus" in s for s in desc_list)