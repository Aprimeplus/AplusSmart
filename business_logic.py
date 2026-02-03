import pandas as pd
import numpy as np

def calculate_monthly_commission(plan_name, comm_df, sales_target=0, operating_fee=None, 
                                 additional_deductions=None, incentives=None, 
                                 min_sales_target=500000):
    """
    Calculates the monthly commission based on the specified plan.
    (Updated: Supports Reconciliation for Special Services - Cutting/Drilling & Services)
    """
    total_brokerage_fee = 0.0

    # --- เตรียมตัวแปรสำหรับสรุปผลสุดท้าย (Incentive/Deduction) ---
    if incentives is None: incentives = {}
    total_incentives = sum(incentives.values())
    
    if additional_deductions is None: additional_deductions = {}
    total_additional_deductions = sum(additional_deductions.values())

    # ==================================================================================
    # HELPER: ฟังก์ชันสำหรับเตรียมข้อมูลและคำนวณกำไรสุทธิ (ใช้ร่วมกันทุก Plan)
    # ==================================================================================
    def prepare_and_calculate_profit(df):
        # 1. Fill NA ให้เป็น 0 ให้หมด เพื่อป้องกัน Error
        cols_to_fix = [
            'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee', 
            'difference_amount', 'payment_before_vat', 'payment_no_vat', 
            'shipping_cost', 'total_po_shipping_cost', 'coupon_fee',
            # คอลัมน์ใหม่ที่ส่งมาจาก HRScreen
            'so_cutting_rev', 'so_service_rev', 'po_cutting_cost', 'po_service_cost'
        ]
        
        for col in cols_to_fix:
            if col not in df.columns: 
                df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # 2. จัดการ Multiplier (ตัวคูณต้นทุน)
        if 'cost_multiplier' in df.columns and not df['cost_multiplier'].isnull().all():
             df['cost_multiplier'] = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
             # ป้องกันค่า 0 หรือ 1 ที่ผิดพลาด ให้บังคับเป็น 1.03 (ถ้าต่ำกว่า 1.01)
             df['cost_multiplier'] = df['cost_multiplier'].apply(lambda x: 1.03 if x < 1.01 else x)
        else:
             df['cost_multiplier'] = 1.03

        # 3. คำนวณ Shipping Deduction (ขาดทุนค่าส่ง)
        # สูตร: ถ้า PO (จ่ายจริง) > SO (เก็บลูกค้า) ให้หักส่วนต่าง
        net_shipping_deduction = (df['total_po_shipping_cost'] - df['shipping_cost']).clip(lower=0)

        # =========================================================
        # [🔥 NEW] Reconciliation for Special Services (กระทบยอด)
        # =========================================================
        
        # A. กลุ่มค่าตัด/เจาะ (Cutting/Drilling)
        # ถ้า ต้นทุน(PO) > รายรับ(SO) -> ถือเป็นขาดทุน ต้องนำมาหัก
        deduct_cutting = (df['po_cutting_cost'] - df['so_cutting_rev']).clip(lower=0)
        
        # B. กลุ่มค่าบริการอื่นๆ (Other Services)
        # ถ้า ต้นทุน(PO) > รายรับ(SO) -> ถือเป็นขาดทุน ต้องนำมาหัก
        deduct_service = (df['po_service_cost'] - df['so_service_rev']).clip(lower=0)

        # C. ปรับปรุงต้นทุน (Exclusion)
        # ต้องหักต้นทุนค่าบริการเหล่านี้ ออกจาก "ต้นทุนรวม" ก่อนนำไปคูณ 1.03
        # เพราะค่าบริการพวกนี้เราคิดแบบ Net (หักลบกลบหนี้) ไม่ได้คิด Margin ปกติ
        adjusted_po_cost = df['final_cost_amount'] - df['po_cutting_cost'] - df['po_service_cost']
        
        # ป้องกันไม่ให้ต้นทุนติดลบ (กรณีข้อมูลผิดพลาด)
        adjusted_po_cost = adjusted_po_cost.clip(lower=0)

        # D. คำนวณ Profit (สูตรใหม่)
        # Profit = ยอดขายรวม - (ต้นทุนสินค้า x Multiplier) + ส่วนต่างโอน - ขาดทุนค่าส่ง - ขาดทุนค่าตัด - ขาดทุนค่าบริการ
        
        main_cost_calculated = (adjusted_po_cost * df['cost_multiplier'])
        
        # สูตรคำนวณกำไรสุทธิ
        df['profit'] = (df['sales_service_amount'] - main_cost_calculated) + df['difference_amount'] \
                       - net_shipping_deduction - deduct_cutting - deduct_service
        
        # E. คำนวณ Margin %
        df['margin'] = (df['profit'] / df['sales_service_amount'].replace(0, np.nan)) * 100
        df['margin'] = df['margin'].fillna(0)

        if df['so_cutting_rev'].sum() > 0 or df['po_cutting_cost'].sum() > 0 or df['po_service_cost'].sum() > 0:
            print("\n" + "="*40)
            print(f"🤖 SYSTEM THINKING: ตรวจสอบค่าตัด/เจาะ (Reconciliation)")
            print("="*40)
            
            # วนลูปดูทีละใบ (แสดงเฉพาะใบที่มีค่าบริการ)
            temp_df = df[ (df['so_cutting_rev'] > 0) | (df['po_cutting_cost'] > 0) | (df['po_service_cost'] > 0) ]
            
            for index, row in temp_df.iterrows():
                print(f"📄 SO: {row['so_number']}")
                print(f"   1. ต้นทุนรวมดิบ (Final Cost): {row['final_cost_amount']:,.2f}")
                
                print(f"   2. แยกค่าตัด/บริการออก (Exclude EXP):")
                print(f"      - ตัดออก: {row['po_cutting_cost']:,.2f} (ค่าตัด)")
                print(f"      - ตัดออก: {row['po_service_cost']:,.2f} (ค่าบริการ)")
                print(f"      = ต้นทุนสินค้าจริง (ที่จะคูณ 1.03): {adjusted_po_cost[index]:,.2f}")
                
                print(f"   3. กระทบยอด (Reconcile EXP):")
                print(f"      [ค่าตัด] จ่าย {row['po_cutting_cost']:,.2f} vs รับ {row['so_cutting_rev']:,.2f} -> หักเพิ่ม: {deduct_cutting[index]:,.2f}")
                print(f"      [บริการ] จ่าย {row['po_service_cost']:,.2f} vs รับ {row['so_service_rev']:,.2f} -> หักเพิ่ม: {deduct_service[index]:,.2f}")
                
                print(f"   4. สรุปกำไร (Profit): {row['profit']:,.2f}")
                print("-" * 40)
            print("="*40 + "\n")

        # =========================================================

        return df

    # ==================================================================================
    # PLAN A
    # ==================================================================================
    if plan_name == 'Plan A':
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper Function)
        comm_df = prepare_and_calculate_profit(comm_df)
        
        # 2. คำนวณคอมมิชชั่น
        total_sales = comm_df['sales_service_amount'].sum()
        
        initial_commission, calculated_commission, commission_normal, commission_below = 0.0, 0.0, 0.0, 0.0
        OPERATING_FEE = 25000.00 if operating_fee is None else operating_fee
        
        if total_sales >= min_sales_target:
            NORMAL_RATE = 0.35
            BELOW_T_RATE = 0.175

            # แยกกลุ่มตาม Margin
            normal_df = comm_df[comm_df['margin'] >= 10]
            below_df = comm_df[comm_df['margin'] < 10]

            # คำนวณยอด
            val_normal_sales = normal_df['sales_service_amount'].sum()
            val_below_sales = below_df['sales_service_amount'].sum()
            
            # คำนวณคอมมิชชั่นจาก Profit
            commission_normal = normal_df['profit'].sum() * NORMAL_RATE
            commission_below = below_df['profit'].sum() * BELOW_T_RATE
            initial_commission = commission_normal + commission_below
            
            # หัก Brokerage Fee
            brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
            total_brokerage_fee = brokerage.sum()
            
            # คำนวณยอดสุทธิ (ห้ามติดลบ)
            calculated_commission = max(0, initial_commission - total_brokerage_fee - OPERATING_FEE)

            # Assign ค่ากลับไปใน DataFrame เพื่อทำ Breakdown
            conditions = [comm_df['margin'] >= 10, comm_df['margin'] < 10]
            choices_commission = [comm_df['profit'] * NORMAL_RATE, comm_df['profit'] * BELOW_T_RATE]
            comm_df['commission_amount'] = np.select(conditions, choices_commission, default=0)
        else:
            # ยอดขายไม่ถึงเป้าขั้นต่ำ
            val_normal_sales = 0.0
            val_below_sales = 0.0
            comm_df['commission_amount'] = 0.0

        # --- 3. สรุปยอดเงิน (Financials) ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 4. สร้างตาราง Debug/Breakdown ---
        num_so = len(comm_df['so_number'].unique())
        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"
        
        # Breakdown DF
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost', # เพิ่ม column ต้นทุนพิเศษให้ดูง่าย
            'profit', 'margin', 'commission_amount', 'cost_multiplier'
        ]].copy()
        
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        so_breakdown_df.rename(columns={
            'so_number': 'SO Number', 
            'sales_service_amount': 'ยอดขาย', 
            'final_cost_amount': 'ต้นทุนรวม', 
            'po_cutting_cost': 'ทุนค่าตัด',
            'po_service_cost': 'ทุนบริการ',
            'profit': 'กำไร', 
            'margin': 'Margin (%)',
            'commission_amount': 'ค่าคอมที่ได้รับ'
        }, inplace=True)
        
        # Debug DF
        debug_details = []
        debug_details.append({'รายการ': '## Report I: Sale Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'เป้ายอดขาย (Target KPI)', 'ค่า': sales_target})
        debug_details.append({'รายการ': 'จำนวนบิล (SO)', 'ค่า': f"{num_so} บิล"})
        debug_details.append({'รายการ': 'สรุปยอดขาย (Sales Base)', 'ค่า': total_sales})
        debug_details.append({'รายการ': 'KPI Monthly SALE TARGET', 'ค่า': f"{hit_target_status} ({hit_target_percent:.2f}%)"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## 2. การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_sales >= min_sales_target:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่นจากกำไรปกติ (Normal Margin @ 35%)', 'ค่า': commission_normal})
            debug_details.append({'รายการ': 'คอมมิชชั่นจากกำไรนอกเงื่อนไข (Below Margin @ 17.5%)', 'ค่า': commission_below})
            debug_details.append({'รายการ': '  = ยอดรวมคอมมิชชั่นก่อนหักค่าดำเนินการ', 'ค่า': initial_commission})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ', 'ค่า': OPERATING_FEE})
            debug_details.append({'รายการ': '  (-) หักค่านายหน้า (Brokerage Fee)', 'ค่า': total_brokerage_fee})
        else:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ไม่ผ่าน ❌'})
        
        debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})

        # Debug Report: Tax
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## 3. สรุปยอดจ่ายสุทธิ (Final Payout) ##', 'ค่า': ''})
        if total_incentives > 0:
             debug_details.append({'รายการ': '(+) รวม Incentive เพิ่มเติม', 'ค่า': total_incentives})
        if total_additional_deductions > 0:
             debug_details.append({'รายการ': '(-) หักค่าใช้จ่ายอื่นๆ', 'ค่า': -total_additional_deductions})
        debug_details.append({'รายการ': 'ยอดคอมมิชชั่นก่อนหักภาษี (Pre-tax Base)', 'ค่า': pre_tax_commission})
        debug_details.append({'รายการ': f'(-) หัก ณ ที่จ่าย 3% ({pre_tax_commission:,.2f} x 3%)', 'ค่า': -withholding_tax})
        debug_details.append({'รายการ': '=== ยอดโอนสุทธิ (Net Commission) ===', 'ค่า': net_commission})
        
        debug_df = pd.DataFrame(debug_details)

        # Summary DataFrame
        summary_desc = ["ยอดขาย Normal", "ยอดขาย Below Tier", "(-) หัก ค่าดำเนินการ/นายหน้า", "ยอดรวมค่าคอมที่คำนวณได้"]
        summary_val = [val_normal_sales, val_below_sales, OPERATING_FEE + total_brokerage_fee, calculated_commission]
        
        for k, v in incentives.items():
            summary_desc.append(f"(+) {k}")
            summary_val.append(v)
            
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)")
        summary_val.append(gross_commission)
        
        for k, v in additional_deductions.items():
            summary_desc.append(f"(-) {k}")
            summary_val.append(v)

        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])

        return {
            'type': 'summary_plan_a',
            'summary': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'final_commission_pre_deductions': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df,
            'details': so_breakdown_df
        }

    # ==================================================================================
    # PLAN B
    # ==================================================================================
    elif plan_name == 'Plan B':
        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper)
        comm_df = prepare_and_calculate_profit(comm_df)

        if 'po_number' not in comm_df.columns: comm_df['po_number'] = comm_df['so_number']
        
        # Group by PO (รวมยอด Profit ที่คำนวณมาแล้ว)
        agg_rules = {
            'sales_service_amount': 'sum', 
            'giveaways': 'sum', 'coupon_fee': 'sum', 'brokerage_fee': 'sum', 
            'difference_amount': 'sum', 
            'shipping_cost': 'first', 'total_po_shipping_cost': 'first', 
            'final_cost_amount': 'sum', 'cost_multiplier': 'first', 
            'so_number': lambda x: ', '.join(sorted(set(x))),
            'profit': 'sum', # รวม Profit ที่คำนวณถูกต้องแล้ว
            # รวมยอดพิเศษด้วย (เผื่อแสดงผล)
            'so_cutting_rev': 'sum', 'so_service_rev': 'sum', 
            'po_cutting_cost': 'sum', 'po_service_cost': 'sum'
        }
        po_grouped_df = comm_df.groupby('po_number').agg(agg_rules).reset_index()

        # คำนวณ Margin ใหม่จากยอดรวม
        po_grouped_df['margin'] = (po_grouped_df['profit'] / po_grouped_df['sales_service_amount'].replace(0, np.nan)) * 100    
        po_grouped_df['margin'] = po_grouped_df['margin'].fillna(0)

        # --- 3. คำนวณค่าคอมมิชชั่น (ตาม Tier) ---
        standard_margin_df = po_grouped_df[po_grouped_df['margin'] >= 10]
        below_tier1_df = po_grouped_df[(po_grouped_df['margin'] >= 7.99) & (po_grouped_df['margin'] < 10)]
        below_tier2_df = po_grouped_df[po_grouped_df['margin'] < 7.99]
        
        total_standard_sales = standard_margin_df['sales_service_amount'].sum()
        total_below_tier1_sales = below_tier1_df['sales_service_amount'].sum()
        total_below_tier2_sales = below_tier2_df['sales_service_amount'].sum()
        total_monthly_sales = total_standard_sales + total_below_tier1_sales + total_below_tier2_sales
        
        commission_below_t1 = total_below_tier1_sales * 0.0063
        commission_below_t2 = total_below_tier2_sales * 0.0050
        below_tier_commission = commission_below_t1 + commission_below_t2

        total_brokerage_fee = po_grouped_df['brokerage_fee'].sum()
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        commission_base = total_standard_sales - total_brokerage_fee - OPERATING_FEE
        
        t1, t2, t3, tier_commission, calculated_commission = 0, 0, 0, 0, 0
        hit_target_percent = (total_monthly_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        if total_monthly_sales >= min_sales_target:
            remaining_base = commission_base if commission_base > 0 else 0
            amount_in_t1 = min(remaining_base, 1000000)
            t1 = amount_in_t1 * 0.0125
            remaining_base -= amount_in_t1
            if remaining_base > 0:
                amount_in_t2 = min(remaining_base, 1000000)
                t2 = amount_in_t2 * 0.0175
                remaining_base -= amount_in_t2
            if remaining_base > 0:
                amount_in_t3 = remaining_base
                t3 = amount_in_t3 * 0.0225
            tier_commission = t1 + t2 + t3
            calculated_commission = tier_commission + below_tier_commission

        # --- 4. คำนวณภาษี ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 5. สร้าง Debug Report ---
        debug_details = []
        debug_details.append({'รายการ': '## 1. สรุปยอดขาย (Sale Summary) ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'เป้ายอดขาย (Target KPI)', 'ค่า': sales_target})
        debug_details.append({'รายการ': 'ยอดขายรวมทั้งหมด (Total Monthly Sales)', 'ค่า': total_monthly_sales})
        debug_details.append({'รายการ': '   - ยอดขาย Normal (Margin >= 10%)', 'ค่า': total_standard_sales})
        debug_details.append({'รายการ': '   - ยอดขาย Below Tier 1 (7.99-9.99%)', 'ค่า': total_below_tier1_sales})
        debug_details.append({'รายการ': '   - ยอดขาย Below Tier 2 (< 7.99%)', 'ค่า': total_below_tier2_sales})
        debug_details.append({'รายการ': 'KPI Achievement (%)', 'ค่า': f"{hit_target_percent:.2f}% ({hit_target_status})"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## 2. การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_monthly_sales >= min_sales_target:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ฐานคำนวณคอมฯ Normal (หลังหักค่าใช้จ่าย)', 'ค่า': commission_base})
            debug_details.append({'รายการ': '   - Tier 1 (1.25%)', 'ค่า': t1})
            debug_details.append({'รายการ': '   - Tier 2 (1.75%)', 'ค่า': t2})
            debug_details.append({'รายการ': '   - Tier 3 (2.25%)', 'ค่า': t3})
            debug_details.append({'รายการ': 'รวมคอมฯ Normal', 'ค่า': tier_commission})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่น Below Tier', 'ค่า': ''})
            debug_details.append({'รายการ': '   - Below Tier 1 (0.63%)', 'ค่า': commission_below_t1})
            debug_details.append({'รายการ': '   - Below Tier 2 (0.50%)', 'ค่า': commission_below_t2})
            debug_details.append({'รายการ': 'รวมคอมฯ Below Tier', 'ค่า': below_tier_commission})
        else:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ไม่ผ่าน ❌'})
        
        debug_details.append({'รายการ': '=== ยอดรวมคอมมิชชั่นที่คำนวณได้ ===', 'ค่า': calculated_commission})

        # Debug Report: Tax
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## 3. สรุปยอดจ่ายสุทธิ (Final Payout) ##', 'ค่า': ''})
        if total_incentives > 0: debug_details.append({'รายการ': '(+) รวม Incentive เพิ่มเติม', 'ค่า': total_incentives})
        if total_additional_deductions > 0: debug_details.append({'รายการ': '(-) หักค่าใช้จ่ายอื่นๆ', 'ค่า': -total_additional_deductions})
        debug_details.append({'รายการ': 'ยอดคอมมิชชั่นก่อนหักภาษี (Pre-tax Base)', 'ค่า': pre_tax_commission})
        debug_details.append({'รายการ': f'(-) หัก ณ ที่จ่าย 3% ({pre_tax_commission:,.2f} x 3%)', 'ค่า': -withholding_tax})
        debug_details.append({'รายการ': '=== ยอดโอนสุทธิ (Net Commission) ===', 'ค่า': net_commission})
        
        summary_desc = ["ยอดขายรวม", "ยอดขาย Normal", "ยอดขาย Below T1", "ยอดขาย Below T2", "(-) หัก ค่าดำเนินการ", "ฐานคำนวณคอมฯ Normal", "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"]
        summary_val = [total_monthly_sales, total_standard_sales, total_below_tier1_sales, total_below_tier2_sales, OPERATING_FEE, commission_base, calculated_commission]
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"]); summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        so_breakdown_df = po_grouped_df[[
            'po_number', 'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        def assign_b_tier_status(margin):
            if margin >= 10: return 'Normal (>=10%)'
            if margin >= 7.99: return 'Below Tier (7.99-10%)'
            return 'Below Tier (<7.99%)'
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_b_tier_status)

        return {
            'type': 'summary_other', 'data': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'debug_df': pd.DataFrame(debug_details),
            'so_breakdown_df': so_breakdown_df
        }

    # ==================================================================================
    # PLAN C (Logic คล้าย Plan B)
    # ==================================================================================
    elif plan_name == 'Plan C':
        if comm_df.empty: return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper)
        comm_df = prepare_and_calculate_profit(comm_df)

        # --- 2. คำนวณค่าคอมมิชชั่น ---
        # Plan C ดูเหมือนจะไม่ได้ Group PO จากโค้ดเดิม (ใช้ comm_df โดยตรง)
        tier1 = comm_df[comm_df['margin'] >= 10]['sales_service_amount'].sum()
        tier2 = comm_df[(comm_df['margin'] >= 7.99) & (comm_df['margin'] < 10)]['sales_service_amount'].sum()
        tier3 = comm_df[comm_df['margin'] < 7.99]['sales_service_amount'].sum()
        total_sales = tier1 + tier2 + tier3
        
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        total_brokerage = comm_df['brokerage_fee'].sum()
        total_deduction = OPERATING_FEE + total_brokerage

        comm_t1, comm_t2, comm_t3, calculated_commission = 0,0,0,0
        base_t1, base_t2, base_t3 = 0,0,0
        
        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        if total_sales >= min_sales_target:
            base_t1 = max(0, tier1 - total_deduction)
            rem1 = max(0, total_deduction - tier1)
            comm_t1 = round(base_t1 * 0.01, 2)
            
            base_t2 = max(0, tier2 - rem1)
            rem2 = max(0, rem1 - tier2)
            comm_t2 = round(base_t2 * 0.0063, 2)
            
            base_t3 = max(0, tier3 - rem2)
            comm_t3 = round(base_t3 * 0.005, 2)
            calculated_commission = comm_t1 + comm_t2 + comm_t3

        # --- 4. คำนวณภาษี ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 5. Debug Report ---
        debug_details = []
        debug_details.append({'รายการ': '## 1. สรุปยอดขาย ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Total Sales', 'ค่า': total_sales})
        debug_details.append({'รายการ': '   - Tier 1 (>=10%)', 'ค่า': tier1})
        debug_details.append({'รายการ': '   - Tier 2 (7.99-9.99%)', 'ค่า': tier2})
        debug_details.append({'รายการ': '   - Tier 3 (<7.99%)', 'ค่า': tier3})
        debug_details.append({'รายการ': 'KPI Achievement', 'ค่า': f"{hit_target_percent:.2f}% ({hit_target_status})"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## 2. การคำนวณคอมมิชชั่น ##', 'ค่า': ''})
        if total_sales >= min_sales_target:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'รวมยอดหัก (ค่าดำเนินการ+นายหน้า)', 'ค่า': total_deduction})
            debug_details.append({'รายการ': 'ฐานคำนวณ Tier 1', 'ค่า': base_t1})
            debug_details.append({'รายการ': '   * คอมฯ T1 (1.0%)', 'ค่า': comm_t1})
            debug_details.append({'รายการ': 'ฐานคำนวณ Tier 2', 'ค่า': base_t2})
            debug_details.append({'รายการ': '   * คอมฯ T2 (0.63%)', 'ค่า': comm_t2})
            debug_details.append({'รายการ': 'ฐานคำนวณ Tier 3', 'ค่า': base_t3})
            debug_details.append({'รายการ': '   * คอมฯ T3 (0.50%)', 'ค่า': comm_t3})
        else:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ไม่ผ่าน ❌'})

        debug_details.append({'รายการ': '=== ยอดรวมคอมมิชชั่นที่คำนวณได้ ===', 'ค่า': calculated_commission})

        # Debug Report: Tax
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## 3. สรุปยอดจ่ายสุทธิ (Final Payout) ##', 'ค่า': ''})
        if total_incentives > 0: debug_details.append({'รายการ': '(+) รวม Incentive เพิ่มเติม', 'ค่า': total_incentives})
        if total_additional_deductions > 0: debug_details.append({'รายการ': '(-) หักค่าใช้จ่ายอื่นๆ', 'ค่า': -total_additional_deductions})
        debug_details.append({'รายการ': 'ยอดคอมมิชชั่นก่อนหักภาษี (Pre-tax Base)', 'ค่า': pre_tax_commission})
        debug_details.append({'รายการ': f'(-) หัก ณ ที่จ่าย 3% ({pre_tax_commission:,.2f} x 3%)', 'ค่า': -withholding_tax})
        debug_details.append({'รายการ': '=== ยอดโอนสุทธิ (Net Commission) ===', 'ค่า': net_commission})

        summary_desc = ["ยอดขาย Tier 1", "ยอดขาย Tier 2", "ยอดขาย Tier 3", "ยอดรวมคอมที่คำนวณได้"]
        summary_val = [tier1, tier2, tier3, calculated_commission]
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"]); summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        return {
            'type': 'summary_other', 'data': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'debug_df': pd.DataFrame(debug_details),
            'so_breakdown_df': comm_df[[
                'so_number', 'sales_service_amount', 'profit', 'margin', 'cost_multiplier', 'final_cost_amount',
                'po_cutting_cost', 'po_service_cost'
            ]]
        }

    # ==================================================================================
    # PLAN D (Logic คล้าย Plan B)
    # ==================================================================================
    elif plan_name == 'Plan D':

        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper)
        comm_df = prepare_and_calculate_profit(comm_df)
        
        # --- 2. แบ่งกลุ่มและรวมยอดขาย ---
        normal_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]
        total_normal_sales = normal_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_sales = total_normal_sales + total_below_sales
        
        # --- 3. คำนวณค่าคอมมิชชั่น ---
        normal_commission, below_tier_commission, calculated_commission = 0.0, 0.0, 0.0
        OPERATING_FEE = 750000.00 if operating_fee is None else operating_fee
        
        commission_base_normal = 0.0
        commission_base_below = 0.0

        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        if total_sales >= min_sales_target:
            total_brokerage_fee = comm_df['brokerage_fee'].sum()
            total_deduction_to_cascade = OPERATING_FEE + total_brokerage_fee

            commission_base_normal = max(0, total_normal_sales - total_deduction_to_cascade)
            normal_commission = commission_base_normal * 0.007

            remaining_deduction = max(0, total_deduction_to_cascade - total_normal_sales)
            commission_base_below = max(0, total_below_sales - remaining_deduction)
            below_tier_commission = commission_base_below * 0.003
            
            calculated_commission = normal_commission + below_tier_commission
        
        # --- 4. คำนวณภาษี ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 5. สร้าง debug_df ---
        debug_details = []
        debug_details.append({'รายการ': '## Report I: Sale Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน (Sales Base)', 'ค่า': total_sales})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_sales >= min_sales_target:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่นปกติ (Normal)', 'ค่า': normal_commission})
            debug_details.append({'รายการ': 'คอมมิชชั่นนอกเงื่อนไข (Below Tier)', 'ค่า': below_tier_commission})
        else:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ไม่ผ่าน ❌'})
        
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})

        # Debug Report: Tax
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## 3. สรุปยอดจ่ายสุทธิ (Final Payout) ##', 'ค่า': ''})
        if total_incentives > 0: debug_details.append({'รายการ': '(+) รวม Incentive เพิ่มเติม', 'ค่า': total_incentives})
        if total_additional_deductions > 0: debug_details.append({'รายการ': '(-) หักค่าใช้จ่ายอื่นๆ', 'ค่า': -total_additional_deductions})
        debug_details.append({'รายการ': 'ยอดคอมมิชชั่นก่อนหักภาษี (Pre-tax Base)', 'ค่า': pre_tax_commission})
        debug_details.append({'รายการ': f'(-) หัก ณ ที่จ่าย 3% ({pre_tax_commission:,.2f} x 3%)', 'ค่า': -withholding_tax})
        debug_details.append({'รายการ': '=== ยอดโอนสุทธิ (Net Commission) ===', 'ค่า': net_commission})

        debug_df = pd.DataFrame(debug_details)
        
        # 6. สร้าง so_breakdown_df
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        so_breakdown_df.rename(columns={'so_number': 'SO Number', 'sales_service_amount': 'ยอดขาย', 'final_cost_amount': 'ต้นทุน', 'profit': 'กำไร', 'margin': 'Margin (%)'}, inplace=True)
        
        summary_desc = ["ยอดขาย Normal", "ยอดขาย Below Tier", "(-) หัก ค่าดำเนินการ/นายหน้า", "ฐานคอม Normal", "ยอดรวมค่าคอมที่คำนวณได้"]
        summary_val = [total_normal_sales, total_below_sales, OPERATING_FEE + total_brokerage_fee, commission_base_normal, calculated_commission]
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"]); summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
        }
        
    else:   
         return {'type': 'error', 'message': f'ไม่พบ Plan ที่ชื่อว่า {plan_name}'}