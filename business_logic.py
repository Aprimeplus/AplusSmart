import pandas as pd
import numpy as np

def calculate_monthly_commission(plan_name, comm_df, sales_target=0, additional_deductions=None, incentives=None):
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
        
    # business_logic.py

    elif plan_name == 'Plan B':
    
        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # --- ขั้นตอนที่ 1: ตรวจสอบและเตรียมข้อมูล (เหมือนเดิม) ---
        if 'po_number' not in comm_df.columns:
            print("WARNING: 'po_number' column not found. Grouping by 'so_number'. This may lead to inaccurate calculations.")
            comm_df['po_number'] = comm_df['so_number']
        if 'so_number' in comm_df.columns:
            comm_df = comm_df.drop_duplicates(subset=['so_number'])
        for col in ['coupon_fee', 'giveaways', 'brokerage_fee', 'difference_amount', 'payment_no_vat']:
            if col not in comm_df.columns: comm_df[col] = 0
        numeric_cols = [
            'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee',
            'difference_amount', 'payment_before_vat', 'payment_no_vat', 'shipping_cost', 'coupon_fee'
        ]
        for col in numeric_cols:
            comm_df[col] = pd.to_numeric(comm_df.get(col), errors='coerce').fillna(0)
        comm_df['cost_multiplier'] = pd.to_numeric(comm_df.get('cost_multiplier'), errors='coerce').fillna(1.03)

        # --- ขั้นตอนที่ 2: รวมข้อมูลตาม PO (Aggregation) (เหมือนเดิม) ---
        agg_rules = {
            'sales_service_amount': 'sum', 'giveaways': 'sum', 'coupon_fee': 'sum',
            'brokerage_fee': 'sum', 'difference_amount': 'sum', 'payment_before_vat': 'first',
            'payment_no_vat': 'first', 'shipping_cost': 'first', 'final_cost_amount': 'first',
            'cost_multiplier': 'first', 'so_number': lambda x: ', '.join(sorted(set(x)))
        }
        po_grouped_df = comm_df.groupby('po_number').agg(agg_rules).reset_index()

        # --- ขั้นตอนที่ 3: คำนวณ Profit และ Margin (เหมือนเดิม) ---
        sales_raw = po_grouped_df['sales_service_amount']
        po_cost = po_grouped_df['final_cost_amount']
        multiplier = po_grouped_df['cost_multiplier']
        main_revenue_minus_cost = sales_raw - (po_cost * multiplier)
        other_costs = po_grouped_df['giveaways'] + po_grouped_df['coupon_fee'] + po_grouped_df['brokerage_fee'] + po_grouped_df['difference_amount']
        net_shipping_adjustment = (po_grouped_df['payment_before_vat'] - po_grouped_df['payment_no_vat']) - po_grouped_df['shipping_cost']
        po_grouped_df['profit'] = main_revenue_minus_cost - other_costs - net_shipping_adjustment
        po_grouped_df['margin'] = (po_grouped_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        po_grouped_df['margin'] = po_grouped_df['margin'].fillna(0)

        # --- ขั้นตอนที่ 4: คำนวณค่าคอมมิชชั่น (เหมือนเดิม) ---
        standard_margin_df = po_grouped_df[po_grouped_df['margin'] >= 10]
        below_margin_df = po_grouped_df[po_grouped_df['margin'] < 10]
        total_standard_sales = standard_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_monthly_sales = total_standard_sales + total_below_sales
        operating_fee = 100000.00
        below_tier_commission = total_below_sales * 0.005
        commission_base = total_standard_sales - operating_fee
        
        t1, t2, t3 = 0, 0, 0; amount_in_t1, amount_in_t2, amount_in_t3 = 0,0,0
        tier_commission = 0; calculated_commission = 0

        if total_monthly_sales >= 500000:
            remaining_base = commission_base if commission_base > 0 else 0
            amount_in_t1 = min(remaining_base, 1000000); t1 = amount_in_t1 * 0.0125; remaining_base -= amount_in_t1
            if remaining_base > 0: amount_in_t2 = min(remaining_base, 1000000); t2 = amount_in_t2 * 0.0175; remaining_base -= amount_in_t2
            if remaining_base > 0: amount_in_t3 = remaining_base; t3 = amount_in_t3 * 0.0225
            tier_commission = t1 + t2 + t3
            calculated_commission = tier_commission + below_tier_commission
        else:
            below_tier_commission = 0; calculated_commission = 0

        # <<< START: สร้าง DataFrame สำหรับ Debug (ฉบับปรับปรุง) >>>
        debug_details = []
        
        # --- เพิ่มข้อมูลสรุปส่วน Report I และ II ---
        num_po = len(po_grouped_df['po_number'].unique())
        hit_target_percent = (total_monthly_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        cost_c1 = po_grouped_df['final_cost_amount'].sum()
        cost_c2 = (po_grouped_df['final_cost_amount'] * po_grouped_df['cost_multiplier']).sum()
        cost_c3 = (po_grouped_df['giveaways'] + po_grouped_df['coupon_fee'] + po_grouped_df['brokerage_fee']).sum()
        cost_c4_diff = (po_grouped_df['payment_before_vat'] - po_grouped_df['payment_no_vat']).sum()
        cost_c4_deduct = po_grouped_df['shipping_cost'].sum()
        total_cost = cost_c2 + cost_c3 - po_grouped_df['difference_amount'].sum() + net_shipping_adjustment.sum()

        debug_details.append({'รายการ': '## Report I: Sale Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'Sale Target KPI', 'ค่า': sales_target})
        debug_details.append({'รายการ': 'สรุปการขาย PO รายเดือน', 'ค่า': f"{num_po} บิล"})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน', 'ค่า': total_monthly_sales})
        debug_details.append({'รายการ': '  - ยอดขายปกติ (Standard Margin)', 'ค่า': total_standard_sales})
        debug_details.append({'รายการ': '  - ยอดขายนอกเงื่อนไข (Below Margin)', 'ค่า': total_below_sales})
        debug_details.append({'รายการ': 'KPI Monthly SALE TARGET', 'ค่า': f"{hit_target_status} ({hit_target_percent:.2f}%)"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})

        debug_details.append({'รายการ': '## Report II: Cost Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้า', 'ค่า': cost_c1})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้าบวกค่าใช้จ่ายบริหารจัดการ', 'ค่า': cost_c2})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าใช้จ่ายการตลาด', 'ค่า': cost_c3})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าบริการขนส่ง (ส่วนต่าง - ติดลบ)', 'ค่า': f"{cost_c4_diff:,.2f} - {cost_c4_deduct:,.2f}"})
        debug_details.append({'รายการ': 'ต้นทุนรวม', 'ค่า': total_cost})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        # --- ส่วนที่เหลือของการสร้าง debug_details ---
        debug_details.append({'รายการ': '## การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_monthly_sales >= 500000:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ฐานคอมฯ (ยอดขายปกติ)', 'ค่า': total_standard_sales})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ', 'ค่า': operating_fee})
            debug_details.append({'รายการ': '  = ฐานสำหรับคำนวณคอมฯ แบบขั้นบันได', 'ค่า': commission_base})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่น T1 (ส่วนแรก 1,000,000 @ 1.25%)', 'ค่า': t1})
            debug_details.append({'รายการ': f'  (จากฐาน: {amount_in_t1:,.2f})', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่น T2 (ส่วนถัดไป 1,000,000 @ 1.75%)', 'ค่า': t2})
            debug_details.append({'รายการ': f'  (จากฐาน: {amount_in_t2:,.2f})', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่น T3 (ส่วนที่เหลือ @ 2.25%)', 'ค่า': t3})
            debug_details.append({'รายการ': f'  (จากฐาน: {amount_in_t3:,.2f})', 'ค่า': ''})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
        else:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ไม่ผ่าน ❌'})

        debug_details.append({'รายการ': '## สรุปค่าคอมมิชชั่น ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'คอมมิชชั่นแบบขั้นบันได (T1+T2+T3)', 'ค่า': tier_commission})
        debug_details.append({'รายการ': 'คอมมิชชั่นนอกเงื่อนไข (Below Tier @ 0.5%)', 'ค่า': below_tier_commission})
        debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})
        # <<< END >>>

        # --- ส่วนสรุปผลและ return ---
        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        gross_commission = calculated_commission + total_incentives
        
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())

        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        summary_desc = [
            "ยอดขายรวม (สำหรับเช็คเงื่อนไข)", "ยอดขายปกติ (สำหรับคำนวณฐานคอม)", "(-) หัก ค่าดำเนินการ",
            "ฐานสำหรับคำนวณคอมมิชชั่น", "คอมมิชชั่น T1 (ฐานคอม 0 - 1M @ 1.25%)",
            "คอมมิชชั่น T2 (ฐานคอม 1M - 2M @ 1.75%)", "คอมมิชชั่น T3 (ฐานคอม > 2M @ 2.25%)",
            "คอมมิชชั่นนอกเงื่อนไข (Below Tier)", "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_monthly_sales, total_standard_sales, operating_fee,
            commission_base if commission_base > 0 else 0, t1, t2, t3,
            below_tier_commission, calculated_commission
        ]
        
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        
        so_breakdown_df = po_grouped_df[['po_number', 'so_number', 'sales_service_amount', 'final_cost_amount', 'profit', 'margin']].copy()
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        so_breakdown_df.rename(columns={
            'po_number': 'PO Number', 'so_number': 'SO Number (Grouped)',
            'sales_service_amount': 'ยอดขาย', 'final_cost_amount': 'ต้นทุน',
            'profit': 'กำไร', 'margin': 'Margin (%)'
        }, inplace=True)
        
        debug_df = pd.DataFrame(debug_details)
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
        }

    elif plan_name == 'Plan C':
        
        if comm_df.empty:
            # ... (ส่วนนี้เหมือนเดิม) ...
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # --- Step 1: คำนวณ Margin (เหมือนเดิม) ---
        # ... (ส่วนนี้เหมือนเดิมทุกประการ) ...
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
        multiplier = comm_df.get('cost_multiplier', 1.03)
        if isinstance(multiplier, pd.Series):
            multiplier = multiplier.fillna(1.03)
        other_deductions = giveaways + brokerage + coupon_fee - difference_amount
        net_shipping_adjustment = (payment_before_vat - payment_no_vat) - so_shipping_cost
        comm_df['profit'] = (sales_raw - (po_cost * multiplier)) - other_deductions - net_shipping_adjustment
        comm_df['margin'] = (comm_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)
        tier1_df = comm_df[comm_df['margin'] >= 10]
        tier2_df = comm_df[(comm_df['margin'] >= 7.99) & (comm_df['margin'] < 10)]
        tier3_df = comm_df[comm_df['margin'] < 7.99]
        total_sales_t1 = tier1_df['sales_service_amount'].sum()
        total_sales_t2 = tier2_df['sales_service_amount'].sum()
        total_sales_t3 = tier3_df['sales_service_amount'].sum()
        total_sales = total_sales_t1 + total_sales_t2 + total_sales_t3
        
        commission_t1, commission_t2, commission_t3 = 0.0, 0.0, 0.0
        calculated_commission = 0.0
        operating_fee = 100000.00
        base_t1, base_t2, base_t3 = 0.0, 0.0, 0.0

        # <<< START: ปรับโครงสร้างข้อมูล Debug ใหม่ทั้งหมด >>>
        debug_details = []
        
        num_so = len(comm_df['so_number'].unique())
        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        cost_c1 = po_cost.sum()
        cost_c2 = (po_cost * multiplier).sum()
        cost_c3 = (giveaways + coupon_fee + brokerage).sum()
        cost_c4_diff = (payment_before_vat - payment_no_vat).sum()
        cost_c4_deduct = so_shipping_cost.sum()
        total_cost = cost_c2 + cost_c3 - difference_amount.sum() + net_shipping_adjustment.sum()

        debug_details.append({'รายการ': '## Report I: Sale Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'Sale Target KPI', 'ค่า': sales_target})
        debug_details.append({'รายการ': 'สรุปการขาย SO รายเดือน', 'ค่า': f"{num_so} บิล"})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน', 'ค่า': total_sales})
        debug_details.append({'รายการ': '  - ยอดขายนอกเงื่อนไข (Below Margin)', 'ค่า': total_sales_t2 + total_sales_t3})
        debug_details.append({'รายการ': '  - ยอดขายปกติ (Standard Margin)', 'ค่า': total_sales_t1})
        debug_details.append({'รายการ': 'KPI Monthly SALE TARGET', 'ค่า': f"{hit_target_status} ({hit_target_percent:.2f}%)"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})

        debug_details.append({'รายการ': '## Report II: Cost Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้า', 'ค่า': cost_c1})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้าบวกค่าใช้จ่ายบริหารจัดการ', 'ค่า': cost_c2})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าใช้จ่ายการตลาด', 'ค่า': cost_c3})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าบริการขนส่ง (ส่วนต่าง - ติดลบ)', 'ค่า': f"{cost_c4_diff:,.2f} - {cost_c4_deduct:,.2f}"})
        debug_details.append({'รายการ': 'ต้นทุนรวม', 'ค่า': total_cost})
        debug_details.append({'รายการ': '---', 'ค่า': ''})

        if total_sales >= 500000:
            base_t1 = max(0, total_sales_t1 - operating_fee)
            remaining_fee_after_t1 = max(0, operating_fee - total_sales_t1)
            commission_t1 = base_t1 * 0.01
            base_t2 = max(0, total_sales_t2 - remaining_fee_after_t1)
            remaining_fee_after_t2 = max(0, remaining_fee_after_t1 - total_sales_t2)
            commission_t2 = base_t2 * 0.0063
            base_t3 = max(0, total_sales_t3 - remaining_fee_after_t2)
            commission_t3 = base_t3 * 0.005
            calculated_commission = commission_t1 + commission_t2 + commission_t3

            debug_details.append({'รายการ': '## การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})

            debug_details.append({'รายการ': 'ยอดขาย Tier 1 (>=10%)', 'ค่า': total_sales_t1})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ', 'ค่า': min(total_sales_t1, operating_fee)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': base_t1})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (1.00%)', 'ค่า': commission_t1})
            debug_details.append({'รายการ': '---', 'ค่า': ''})

            debug_details.append({'รายการ': 'ยอดขาย Tier 2 (7.99-10%)', 'ค่า': total_sales_t2})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ (ส่วนที่เหลือ)', 'ค่า': min(total_sales_t2, remaining_fee_after_t1)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': base_t2})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (0.63%)', 'ค่า': commission_t2})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            
            debug_details.append({'รายการ': 'ยอดขาย Tier 3 (<7.99%)', 'ค่า': total_sales_t3})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ (ส่วนที่เหลือ)', 'ค่า': min(total_sales_t3, remaining_fee_after_t2)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': base_t3})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (0.50%)', 'ค่า': commission_t3})
            debug_details.append({'รายการ': '---', 'ค่า': ''})

            debug_details.append({'รายการ': '## สรุปค่าคอมมิชชั่น ##', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่น Tier 1', 'ค่า': commission_t1})
            debug_details.append({'รายการ': 'คอมมิชชั่น Tier 2', 'ค่า': commission_t2})
            debug_details.append({'รายการ': 'คอมมิชชั่น Tier 3', 'ค่า': commission_t3})
            debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})
        else:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ไม่ผ่าน ❌'})
        # <<< END >>>

        # ... (ส่วนสรุปผล summary_data เหมือนเดิม) ...
        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        gross_commission = calculated_commission + total_incentives
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        summary_desc = [
            "ยอดขาย Tier 1 (>=10%)", "ยอดขาย Tier 2 (7.99-10%)", "ยอดขาย Tier 3 (<7.99%)", "ยอดขายรวม",
            "เงื่อนไขขั้นต่ำ (500,000)", "คอมฯ T1 (1.00%)", "คอมฯ T2 (0.63%)", "คอมฯ T3 (0.50%)",
            "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_sales_t1, total_sales_t2, total_sales_t3, total_sales,
            "ผ่าน" if total_sales >= 500000 else "ไม่ผ่าน",
            commission_t1, commission_t2, commission_t3,
            calculated_commission
        ]
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก: {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        
        so_breakdown_df = comm_df[['so_number', 'sales_service_amount', 'shipping_cost', 'final_cost_amount', 'profit', 'margin']].copy()
        def assign_tier_status(margin):
            if margin >= 10: return "Normal (>=10%)"
            if margin >= 7.99: return "Below Tier (7.99-10%)"
            return "Below Tier (<7.99%)"
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_tier_status)
        so_breakdown_df.rename(columns={
            'so_number': 'SO Number', 'sales_service_amount': 'ยอดขาย', 'shipping_cost': 'ค่าส่ง',
            'final_cost_amount': 'ต้นทุน', 'profit': 'กำไร', 'margin': 'Margin (%)'
        }, inplace=True)
        
        debug_df = pd.DataFrame(debug_details)
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
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