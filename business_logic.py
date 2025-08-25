import pandas as pd
import numpy as np

def calculate_monthly_commission(plan_name, comm_df, additional_deductions=None, incentives=None):
    """
    Calculates the monthly commission based on the specified plan.
    (ฉบับแก้ไขสมบูรณ์ตาม "เฉลย" Excel)

    Args:
        plan_name (str): The name of the commission plan (e.g., 'Plan A').
        comm_df (pd.DataFrame): DataFrame containing the commission data for the month.
        additional_deductions (dict, optional): Additional deductions for the month. Defaults to None.
        incentives (dict, optional): Incentives for the month. Defaults to None.

    Returns:
        dict: A dictionary containing the summary DataFrame, detailed DataFrame, and total commission.
    """
    
    if plan_name == 'Plan A':
    
        # 1. ดึงข้อมูลทุกส่วนที่จำเป็น
        sales_raw = pd.to_numeric(comm_df.get('sales_service_amount', 0), errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df.get('final_cost_amount', 0), errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df.get('giveaways', 0), errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df.get('brokerage_fee', 0), errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df.get('difference_amount', 0), errors='coerce').fillna(0)
        
        # ดึงข้อมูล Z5, Y5, Q5
        payment_before_vat = pd.to_numeric(comm_df.get('payment_before_vat', 0), errors='coerce').fillna(0) # Z5
        payment_no_vat = pd.to_numeric(comm_df.get('payment_no_vat', 0), errors='coerce').fillna(0)          # Y5
        so_shipping_cost = pd.to_numeric(comm_df.get('shipping_cost', 0), errors='coerce').fillna(0)        # Q5
        
        # --- START: แก้ไขสูตรคำนวณกำไรเป็นเวอร์ชัน Final V2 ---
        # 2. คำนวณแต่ละส่วนของสูตร
        
        # <<< START: โค้ดที่แก้ไขข้อผิดพลาด ValueError >>>
        multiplier = comm_df.get('cost_multiplier') if 'cost_multiplier' in comm_df else 1.03
        if isinstance(multiplier, pd.Series):
            multiplier = multiplier.fillna(1.03)
        # <<< END: โค้ดที่แก้ไขข้อผิดพลาด ValueError >>>
        
        # ตรวจสอบก่อนว่ามีคอลัมน์ coupon_fee หรือไม่ ถ้าไม่มีให้สร้างขึ้นมาใหม่เป็น 0
        if 'coupon_fee' not in comm_df.columns:
            comm_df['coupon_fee'] = 0
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)

        other_deductions = giveaways + brokerage + coupon_fee - difference_amount
        
        # คำนวณส่วนของค่ารถตามสูตรใหม่
        net_shipping_adjustment = (payment_before_vat - payment_no_vat) - so_shipping_cost
        comm_df['profit'] = (sales_raw - (po_cost * multiplier)) - other_deductions - net_shipping_adjustment
        # --- END: สิ้นสุดการแก้ไขสูตร ---

        # 4. คำนวณ Margin และ ค่าคอม
        comm_df['margin'] = (comm_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)
        
        NORMAL_RATE = 0.35
        BELOW_T_RATE = 0.175
        conditions = [ comm_df['margin'] >= 10, comm_df['margin'] < 10 ]
        choices_commission = [ comm_df['profit'] * NORMAL_RATE, comm_df['profit'] * BELOW_T_RATE ]
        comm_df['commission_amount'] = np.select(conditions, choices_commission, default=0)

        # 5. สร้างตาราง Debug ที่สะท้อนสูตร Final V2
        comm_df['profit_formula_str'] = comm_df.apply(
            lambda row: f"(...) - (({row['payment_before_vat']:,.2f} - {row['payment_no_vat']:,.2f}) - {row['shipping_cost']:,.2f}) = {row['profit']:,.2f}",
            axis=1
        )
        
        details_df = comm_df[['so_number', 'profit_formula_str', 'profit', 'margin']].copy()
        details_df.rename(columns={
            'so_number': 'SO Number',
            'profit_formula_str': 'สูตรกำไร: (...) - ((Z5 - Y5) - Q5) = กำไร',
            'profit': 'กำไร',
            'margin': 'Margin (%)'
        }, inplace=True)
        # --- END: สิ้นสุดการแก้ไข ---

        # 8. สรุปยอด (เหมือนเดิม)
        OPERATING_FEE = 25000.00
        initial_commission = comm_df['commission_amount'].sum()
        calculated_commission = max(0, initial_commission - OPERATING_FEE)
        commission_normal = comm_df[comm_df['margin'] >= 10]['commission_amount'].sum()
        commission_below = comm_df[comm_df['margin'] < 10]['commission_amount'].sum()
        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax
        summary_desc = ["ยอดคอมมิชชั่นปกติ (Normal)", "ยอดคอมมิชชั่นนอกเงื่อนไข (Below Tier)", "ยอดรวมค่าคอมมิชชั่น", "(-) หัก ค่าดำเนินการ", "ยอดคอมมิชชั่นที่คำนวณได้"]
        summary_val = [commission_normal, commission_below, initial_commission, OPERATING_FEE, calculated_commission]
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก: {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_df = pd.DataFrame({'description': summary_desc, 'value': summary_val})

        return {
            'type': 'summary_plan_a', 
            'summary': summary_df, 
            'details': details_df,
            'final_commission': calculated_commission
        }
        
    elif plan_name == 'Plan B':
        
        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # Step 1: คำนวณค่าพื้นฐานและแยกประเภท Margin
        sales_raw = pd.to_numeric(comm_df.get('sales_service_amount', 0), errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df.get('final_cost_amount', 0), errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df.get('giveaways', 0), errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df.get('brokerage_fee', 0), errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df.get('difference_amount', 0), errors='coerce').fillna(0)
        payment_before_vat = pd.to_numeric(comm_df.get('payment_before_vat', 0), errors='coerce').fillna(0)
        payment_no_vat = pd.to_numeric(comm_df.get('payment_no_vat', 0), errors='coerce').fillna(0)
        so_shipping_cost = pd.to_numeric(comm_df.get('shipping_cost', 0), errors='coerce').fillna(0)
        
        multiplier = comm_df.get('cost_multiplier') if 'cost_multiplier' in comm_df else 1.03
        if isinstance(multiplier, pd.Series):
            multiplier = multiplier.fillna(1.03)

        if 'coupon_fee' not in comm_df.columns:
            comm_df['coupon_fee'] = 0
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)

        other_deductions = giveaways + brokerage + coupon_fee - difference_amount
        net_shipping_adjustment = (payment_before_vat - payment_no_vat) - so_shipping_cost
        
        comm_df['profit'] = (sales_raw - (po_cost * multiplier)) - other_deductions - net_shipping_adjustment
        comm_df['margin'] = (comm_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)
        
        standard_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]

        # Step 2: คำนวณยอดขายและกำไรแยกตามประเภท
        total_standard_sales = standard_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_monthly_sales = total_standard_sales + total_below_sales
        total_profit_below = below_margin_df['profit'].sum() 
        
        # --- START: โค้ด Plan B ฉบับสมบูรณ์ตาม Logic ล่าสุด ---
        operating_fee = 100000.00
        below_tier_commission = total_profit_below * 0.005
        commission_base = total_standard_sales - operating_fee

        t1, t2, t3 = 0, 0, 0
        tier_commission = 0
        calculated_commission = 0

        # 1. เช็คยอดขายดิบ "รวม" (Normal + Below Tier) ขั้นต่ำ 500,000
        if total_monthly_sales >= 500000:
            # 2. ถ้าผ่าน ให้คำนวณค่าคอมแบบขั้นบันได
            remaining_base = commission_base if commission_base > 0 else 0

            # T1: 1.25% on the first 1,000,000 of the base
            amount_in_t1 = min(remaining_base, 1000000)
            t1 = amount_in_t1 * 0.0125
            remaining_base -= amount_in_t1

            # T2: 1.75% on the next 1,000,000 of the base
            if remaining_base > 0:
                amount_in_t2 = min(remaining_base, 1000000)
                t2 = amount_in_t2 * 0.0175
                remaining_base -= amount_in_t2
            
            # T3: 2.25% on the rest of the base
            if remaining_base > 0:
                t3 = remaining_base * 0.0225

            tier_commission = t1 + t2 + t3
            calculated_commission = tier_commission + below_tier_commission
        else:
            # 3. ถ้าไม่ผ่านเงื่อนไข 500k ค่าคอมทั้งหมดเป็น 0
            below_tier_commission = 0 # กำหนดให้เป็น 0 เพื่อการแสดงผล
            calculated_commission = 0
        # --- END: สิ้นสุดการปรับปรุงโค้ด ---

        # Step 3: สรุปยอดสุดท้าย
        if incentives is None:
            incentives = {}
        total_incentives = sum(incentives.values())
        gross_commission = calculated_commission + total_incentives
        
        if additional_deductions is None:
            additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        
        pre_tax_commission = gross_commission - total_additional_deductions

        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax
        
        summary_desc = [
            "ยอดขายรวม (สำหรับเช็คเงื่อนไข)",
            "ยอดขายปกติ (สำหรับคำนวณฐานคอม)",
            "(-) หัก ค่าดำเนินการ",
            "ฐานสำหรับคำนวณคอมมิชชั่น",
            "คอมมิชชั่น T1 (ฐานคอม 0 - 1M @ 1.25%)",
            "คอมมิชชั่น T2 (ฐานคอม 1M - 2M @ 1.75%)",
            "คอมมิชชั่น T3 (ฐานคอม > 2M @ 2.25%)",
            "คอมมิชชั่นนอกเงื่อนไข (Below Tier)",
            "ยอดคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_monthly_sales,
            total_standard_sales,
            operating_fee,
            commission_base if commission_base > 0 else 0,
            t1, t2, t3,
            below_tier_commission,
            calculated_commission
        ]
        
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission
        }

    elif plan_name == 'Plan C':
        
        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # Step 1: คำนวณ Margin เพื่อใช้แบ่งกลุ่ม
        sales_raw = pd.to_numeric(comm_df.get('sales_service_amount', 0), errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df.get('final_cost_amount', 0), errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df.get('giveaways', 0), errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df.get('brokerage_fee', 0), errors='coerce').fillna(0)
        if 'coupon_fee' not in comm_df.columns:
            comm_df['coupon_fee'] = 0
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df.get('difference_amount', 0), errors='coerce').fillna(0)
        payment_before_vat = pd.to_numeric(comm_df.get('payment_before_vat', 0), errors='coerce').fillna(0)
        payment_no_vat = pd.to_numeric(comm_df.get('payment_no_vat', 0), errors='coerce').fillna(0)
        so_shipping_cost = pd.to_numeric(comm_df.get('shipping_cost', 0), errors='coerce').fillna(0)
        
        # <<< START: โค้ดที่แก้ไขข้อผิดพลาด ValueError >>>
        multiplier = comm_df.get('cost_multiplier') if 'cost_multiplier' in comm_df else 1.03
        if isinstance(multiplier, pd.Series):
            multiplier = multiplier.fillna(1.03)
        # <<< END: โค้ดที่แก้ไขข้อผิดพลาด ValueError >>>

        if 'coupon_fee' not in comm_df.columns:
            comm_df['coupon_fee'] = 0
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)

        other_deductions = giveaways + brokerage + coupon_fee - difference_amount
        net_shipping_adjustment = (payment_before_vat - payment_no_vat) - so_shipping_cost
        
        comm_df['profit'] = (sales_raw - (po_cost * multiplier)) - other_deductions - net_shipping_adjustment
        comm_df['margin'] = (comm_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)
        standard_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]

        # Step 2: รวมยอดขายดิบของแต่ละกลุ่ม
        total_normal_sales = standard_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_sales = total_normal_sales + total_below_sales
        
        normal_commission = 0.0
        below_tier_commission = 0.0
        calculated_commission = 0.0
        
        # Step 3: ตรวจสอบยอดขายขั้นต่ำ 500,000 บาท ก่อน
        if total_sales >= 500000:
            operating_fee = 100000.00
            
            # Step 4: หักค่าดำเนินการตามลำดับความสำคัญ
            adjusted_normal_base = max(0, total_normal_sales - operating_fee)
            remaining_fee = max(0, operating_fee - total_normal_sales)
            adjusted_below_base = max(0, total_below_sales - remaining_fee)
            
            # Step 5: คำนวณค่าคอมจากฐานที่ปรับปรุงแล้ว
            normal_commission = adjusted_normal_base * 0.01
            below_tier_commission = adjusted_below_base * 0.005
            
            calculated_commission = normal_commission + below_tier_commission
        
        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        gross_commission = calculated_commission + total_incentives
        
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())

        pre_tax_commission = gross_commission - total_additional_deductions

        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        summary_desc = [
            "ยอดขาย Normal (ก่อนหัก)", "ยอดขาย Below T (ก่อนหัก)", "ยอดขายรวม",
            "เงื่อนไขขั้นต่ำ (500,000)", "คอมมิชชั่น Normal (1%)",
            "คอมมิชชั่น Below T (0.5%)", "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_normal_sales, total_below_sales, total_sales,
            "ผ่าน" if total_sales >= 500000 else "ไม่ผ่าน",
            normal_commission, below_tier_commission, calculated_commission
        ]

        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก: {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission
        }
    
    elif plan_name == 'Plan D':
        
        print("\n" + "="*20 + " DEBUG: Plan D " + "="*20)
        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        sales_raw = pd.to_numeric(comm_df.get('sales_service_amount', 0), errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df.get('final_cost_amount', 0), errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df.get('giveaways', 0), errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df.get('brokerage_fee', 0), errors='coerce').fillna(0)
        if 'coupon_fee' not in comm_df.columns:
            comm_df['coupon_fee'] = 0
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df.get('difference_amount', 0), errors='coerce').fillna(0)
        payment_before_vat = pd.to_numeric(comm_df.get('payment_before_vat', 0), errors='coerce').fillna(0)
        payment_no_vat = pd.to_numeric(comm_df.get('payment_no_vat', 0), errors='coerce').fillna(0)
        so_shipping_cost = pd.to_numeric(comm_df.get('shipping_cost', 0), errors='coerce').fillna(0)
        
        # <<< START: โค้ดที่แก้ไขข้อผิดพลาด ValueError >>>
        multiplier = comm_df.get('cost_multiplier') if 'cost_multiplier' in comm_df else 1.03
        if isinstance(multiplier, pd.Series):
            multiplier = multiplier.fillna(1.03)
        # <<< END: โค้ดที่แก้ไขข้อผิดพลาด ValueError >>>

        if 'coupon_fee' not in comm_df.columns:
            comm_df['coupon_fee'] = 0
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)

        other_deductions = giveaways + brokerage + coupon_fee - difference_amount
        net_shipping_adjustment = (payment_before_vat - payment_no_vat) - so_shipping_cost
        
        comm_df['profit'] = (sales_raw - (po_cost * multiplier)) - other_deductions - net_shipping_adjustment
        comm_df['margin'] = (comm_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)
        standard_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]
        
        # --- NEW DEBUG START: Show individual SOs ---
        print("\n" + "-"*15 + " Breakdown by SO " + "-"*15)
        print("\n--- SOs with Normal Margin (>= 10%) ---")
        if not standard_margin_df.empty:
            print(standard_margin_df[['so_number', 'sales_service_amount', 'margin']].to_string())
        else:
            print("No SOs with Normal Margin.")
        print("\n--- SOs with Below Margin (< 10%) ---")
        if not below_margin_df.empty:
            print(below_margin_df[['so_number', 'sales_service_amount', 'margin']].to_string())
        else:
            print("No SOs with Below Margin.")
        print("-" * 50)
        # --- NEW DEBUG END ---

        total_normal_sales = standard_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        print(f"\n[DEBUG] Total Normal Sales (Margin >= 10%): {total_normal_sales:,.2f}")
        print(f"[DEBUG] Total Below Sales (Margin < 10%): {total_below_sales:,.2f}")
        print("-"*50)
        operating_fee = 750000.00
        g63_base = total_normal_sales - operating_fee
        f68_base = total_below_sales + g63_base if g63_base < 0 else total_below_sales
        base_normal = max(0, g63_base)
        base_below = max(0, f68_base)
        amount_in_tier1 = min(base_normal, 750000)
        comm_tier1 = amount_in_tier1 * 0.007
        comm_tier2 = 0.0
        if base_normal > 750000:
            amount_in_tier2 = min(base_normal, 1500000) - 750000
            comm_tier2 = amount_in_tier2 * 0.01 
        comm_tier3 = 0.0
        if base_normal > 1500000:
            amount_in_tier3 = base_normal - 1500000
            comm_tier3 = amount_in_tier3 * 0.01
        comm_below_tier = base_below * 0.003
        
        print(f"[DEBUG] G63 Base (Normal - Op Fee): {g63_base:,.2f}")
        print(f"[DEBUG] F68 Base (Below Adjusted): {f68_base:,.2f}")
        print(f"[DEBUG] Commission Tier 1 (Normal): {comm_tier1:,.2f}")
        print(f"[DEBUG] Commission Tier 2 (Normal): {comm_tier2:,.2f}")
        print(f"[DEBUG] Commission Tier 3 (Normal): {comm_tier3:,.2f}")
        print(f"[DEBUG] Commission Below Tier: {comm_below_tier:,.2f}")
        
        calculated_commission = comm_tier1 + comm_tier2 + comm_tier3 + comm_below_tier
        
        print(f"[DEBUG] Final Calculated Commission: {calculated_commission:,.2f}")
        print("="*50 + "\n")

        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        gross_commission = calculated_commission + total_incentives
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax
        summary_desc = ["ยอดขาย Normal", "ยอดขาย Below T", "(-) ค่าดำเนินการ", "ฐานคอม Normal ", "ฐานคอม Below T ", "คอมมิชชั่น Normal (0 - 750,000 บาท)", "คอมมิชชั่น Normal (750,001 - 1,500,000 บาท)", "คอมมิชชั่น Normal (1,500,001 บาท ขึ้นไป)", "คอมมิชชั่น Below Tier", "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"]
        summary_val = [total_normal_sales, total_below_sales, operating_fee, g63_base, f68_base, comm_tier1, comm_tier2, comm_tier3, comm_below_tier, calculated_commission]
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission
        }
        
    else:
         return {'type': 'error', 'message': f'ไม่พบ Plan ที่ชื่อว่า {plan_name}'}