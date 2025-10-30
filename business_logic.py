import pandas as pd
import numpy as np

def calculate_monthly_commission(plan_name, comm_df, sales_target=0, operating_fee=None, additional_deductions=None, incentives=None):
    """
    Calculates the monthly commission based on the specified plan.
    (ฉบับแก้ไขสมบูรณ์ตาม Logic ล่าสุด)
    """
    total_brokerage_fee = 0.0

    if plan_name == 'Plan A':

        # --- 1. เตรียมข้อมูล ---
        required_cols = [
            'total_revenue', 'sales_service_amount', 'final_cost_amount', 'giveaways', 
            'brokerage_fee', 'difference_amount', 'payment_before_vat', 'payment_no_vat', 
            'shipping_cost', 'coupon_fee'
        ]
        for col in required_cols:
            if col not in comm_df.columns:
                comm_df[col] = 0

        total_revenue = pd.to_numeric(comm_df['total_revenue'], errors='coerce').fillna(0)
        sales_raw = pd.to_numeric(comm_df['sales_service_amount'], errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df['final_cost_amount'], errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df['giveaways'], errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df['difference_amount'], errors='coerce').fillna(0)
        payment_before_vat = pd.to_numeric(comm_df['payment_before_vat'], errors='coerce').fillna(0)
        payment_no_vat = pd.to_numeric(comm_df['payment_no_vat'], errors='coerce').fillna(0)
        so_shipping_cost = pd.to_numeric(comm_df['shipping_cost'], errors='coerce').fillna(0)
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)

        if 'cost_multiplier' in comm_df.columns and not comm_df['cost_multiplier'].isnull().all():
            multiplier = pd.to_numeric(comm_df['cost_multiplier'], errors='coerce').fillna(1.03)
        else:
            multiplier = 1.03
                
        other_deductions = giveaways + brokerage + coupon_fee - difference_amount
        net_shipping_adjustment = (payment_before_vat - payment_no_vat) - so_shipping_cost

        # --- 2. คำนวณ Profit/Margin ---
        comm_df['profit'] = (total_revenue - (po_cost * multiplier)) + difference_amount
        comm_df['margin'] = (comm_df['profit'] / sales_raw.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)

        comm_df['commission_amount'] = 0.0
        
        # --- 3. คำนวณค่าคอมมิชชั่น (ใช้ Logic ใหม่ตามที่คุณอธิบาย) ---
        total_sales = total_revenue.sum()
        initial_commission, calculated_commission, commission_normal, commission_below = 0.0, 0.0, 0.0, 0.0
        OPERATING_FEE = 25000.00 if operating_fee is None else operating_fee
        
        if total_sales >= 500000:
            NORMAL_RATE = 0.35
            BELOW_T_RATE = 0.175

            # --- แยก SO ตาม Tier ---
            normal_df = comm_df[comm_df['margin'] >= 10]
            below_df = comm_df[comm_df['margin'] < 10]

            # --- รวม 'กำไร' ของแต่ละ Tier ก่อน ---
            total_normal_profit = normal_df['profit'].sum()
            total_below_profit = below_df['profit'].sum()

            print(f"\n>>> DEBUG: Total Normal Profit (Plan A): {total_normal_profit:,.2f}\n")

            # --- คำนวณคอมมิชชั่นของแต่ละส่วน (A และ B) ---
            commission_normal = total_normal_profit * NORMAL_RATE
            commission_below = total_below_profit * BELOW_T_RATE

            # --- รวมเป็นคอมมิชชั่นก่อนหัก (C) ---
            initial_commission = commission_normal + commission_below
            total_brokerage_fee = brokerage.sum()
            print(f">>> DEBUG: Total Brokerage Fee to Deduct (Plan A): {total_brokerage_fee:,.2f}")


            # --- หักค่าดำเนินการ ได้เป็นค่าคอมสุดท้าย ---
            calculated_commission = max(0, initial_commission - total_brokerage_fee - OPERATING_FEE)

            # --- [Optional] คำนวณค่าคอมราย SO (สำหรับแสดงผลใน Breakdown) ---
            conditions = [comm_df['margin'] >= 10, comm_df['margin'] < 10]
            choices_commission = [comm_df['profit'] * NORMAL_RATE, comm_df['profit'] * BELOW_T_RATE]
            comm_df['commission_amount'] = np.select(conditions, choices_commission, default=0)
        
        print("\n" + "---" * 20)
        print("### DEBUG: Plan A - Per-SO Calculation Breakdown ###")
        # สร้างคอลัมน์ชั่วคราวเพื่อแสดงผล
        comm_df['debug_adj_cost'] = po_cost * multiplier
        comm_df['debug_other_deduct'] = other_deductions
        for index, row in comm_df.iterrows():
            print(f"\n -> SO: {row['so_number']}")
            print(f"    - Total Revenue : {row['total_revenue']:,.2f}")
            print(f"    - Adjusted Cost : {row['debug_adj_cost']:,.2f}")
            print(f"    - Other Deduct  : {row['debug_other_deduct']:,.2f}")
            print(f"    - Profit        : {row['profit']:,.2f}")
            print(f"    - Margin        : {row['margin']:.2f}%")
            print(f"    - Commission Amt: {row['commission_amount']:,.2f}")
        print("---" * 20 + "\n")

        # --- 4. สร้าง DataFrame สำหรับ Debug และรายละเอียด SO ---
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
        debug_details.append({'รายการ': 'KPI Monthly SALE TARGET', 'ค่า': f"{hit_target_status} ({hit_target_percent:.2f}%)"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## Report II: Cost Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้า', 'ค่า': cost_c1})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้าบวกค่าใช้จ่ายบริหารจัดการ', 'ค่า': cost_c2})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าใช้จ่ายการตลาด', 'ค่า': cost_c3})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าบริการขนส่ง (ส่วนต่าง - ติดลบ)', 'ค่า': f"{cost_c4_diff:,.2f} - {cost_c4_deduct:,.2f}"})
        debug_details.append({'รายการ': 'ต้นทุนรวม', 'ค่า': total_cost})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_sales >= 500000:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่นจากกำไรปกติ (Normal Margin @ 35%)', 'ค่า': commission_normal})
            debug_details.append({'รายการ': 'คอมมิชชั่นจากกำไรนอกเงื่อนไข (Below Margin @ 17.5%)', 'ค่า': commission_below})
            debug_details.append({'รายการ': '  = ยอดรวมคอมมิชชั่นก่อนหักค่าดำเนินการ', 'ค่า': initial_commission})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ', 'ค่า': OPERATING_FEE})
        else:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ไม่ผ่าน ❌'})
        
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})
        debug_df = pd.DataFrame(debug_details)

        # --- 5. สร้าง so_breakdown_df (รายละเอียดตาม SO) ---
        so_breakdown_df = comm_df[['so_number', 'sales_service_amount', 'shipping_cost', 'final_cost_amount', 'profit', 'margin', 'commission_amount']].copy()
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        so_breakdown_df.rename(columns={
            'so_number': 'SO Number', 'sales_service_amount': 'ยอดขาย',
            'shipping_cost': 'ค่าส่ง', 'final_cost_amount': 'ต้นทุน',
            'profit': 'กำไร', 'margin': 'Margin (%)',
            'commission_amount': 'ค่าคอมที่ได้รับ'
        }, inplace=True)

        # --- 6. สรุปผลสุดท้าย ---
        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax
        
        summary_desc = ["ยอดคอมฯ ปกติ (Normal)", "ยอดคอมฯ นอกเงื่อนไข (Below Tier)", "ยอดรวมค่าคอมฯ", "(-) หัก ค่าดำเนินการ", "ยอดคอมมิชชั่นที่คำนวณได้"]
        summary_val = [commission_normal, commission_below, initial_commission, OPERATING_FEE, calculated_commission]
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมฯ ขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก: {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมฯ ก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_df = pd.DataFrame({'description': summary_desc, 'value': summary_val})

        return {
            'type': 'summary_other',
            'data': summary_df, 
            'final_commission': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
        }

    elif plan_name == 'Plan B':
    
        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # --- ขั้นตอนที่ 1 และ 2: เตรียมข้อมูลและรวมข้อมูลตาม PO ---
        if 'po_number' not in comm_df.columns:
            print("WARNING: 'po_number' column not found. Grouping by 'so_number'. This may lead to inaccurate calculations.")
            comm_df['po_number'] = comm_df['so_number']
        if 'so_number' in comm_df.columns:
            comm_df = comm_df.drop_duplicates(subset=['so_number'])
                
        for col in ['coupon_fee', 'giveaways', 'brokerage_fee', 'difference_amount', 'payment_no_vat']:
            if col not in comm_df.columns: comm_df[col] = 0
        numeric_cols = [
            'total_revenue', 'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee',
            'difference_amount', 'payment_before_vat', 'payment_no_vat', 'shipping_cost', 'coupon_fee'
        ]
        for col in numeric_cols:
            comm_df[col] = pd.to_numeric(comm_df.get(col), errors='coerce').fillna(0)
        comm_df['cost_multiplier'] = pd.to_numeric(comm_df.get('cost_multiplier'), errors='coerce').fillna(1.03)

        agg_rules = {
            'total_revenue': 'sum', 'sales_service_amount': 'sum', 'giveaways': 'sum', 
            'coupon_fee': 'sum', 'brokerage_fee': 'sum', 'difference_amount': 'sum', 
            'payment_before_vat': 'first', 'payment_no_vat': 'first', 'shipping_cost': 'first', 
            'final_cost_amount': 'first', 'cost_multiplier': 'first', 
            'so_number': lambda x: ', '.join(sorted(set(x)))
        }
        po_grouped_df = comm_df.groupby('po_number').agg(agg_rules).reset_index()

        # --- ขั้นตอนที่ 3: คำนวณ Profit และ Margin ---
        main_profit = po_grouped_df['total_revenue'] - (po_grouped_df['final_cost_amount'] * po_grouped_df['cost_multiplier'])
        difference_adjustment = po_grouped_df['difference_amount']
        po_grouped_df['profit'] = main_profit + difference_adjustment
        po_grouped_df['margin'] = (po_grouped_df['profit'] / po_grouped_df['sales_service_amount'].replace(0, np.nan)) * 100    
        po_grouped_df['margin'] = po_grouped_df['margin'].fillna(0)

        # --- ขั้นตอนที่ 4: คำนวณค่าคอมมิชชั่น (ใช้ Logic Tier ใหม่) ---
        
        # แบ่งกลุ่ม PO ตาม Margin 3 ระดับ
        standard_margin_df = po_grouped_df[po_grouped_df['margin'] >= 10]
        below_tier1_df = po_grouped_df[(po_grouped_df['margin'] >= 7.99) & (po_grouped_df['margin'] < 10)]
        below_tier2_df = po_grouped_df[po_grouped_df['margin'] < 7.99]
        
        # รวมยอดขายของแต่ละกลุ่ม
        total_standard_sales = standard_margin_df['sales_service_amount'].sum()

        print(f"\n>>> DEBUG: Total Standard Sales (Plan B): {total_standard_sales:,.2f}\n")

        total_below_tier1_sales = below_tier1_df['sales_service_amount'].sum()
        total_below_tier2_sales = below_tier2_df['sales_service_amount'].sum()

        # รวมยอดขายทั้งหมดเพื่อเช็คเงื่อนไข
        total_monthly_sales = total_standard_sales + total_below_tier1_sales + total_below_tier2_sales
        
        # คำนวณคอมมิชชั่นของกลุ่ม Below Tier ทั้ง 2 ระดับ
        commission_below_t1 = total_below_tier1_sales * 0.0063  # 0.63%
        commission_below_t2 = total_below_tier2_sales * 0.0050  # 0.50%
        below_tier_commission = commission_below_t1 + commission_below_t2 # รวมเป็นคอมฯ Below Tier ทั้งหมด

        total_brokerage_fee = po_grouped_df['brokerage_fee'].sum()
        print(f">>> DEBUG: Total Brokerage Fee to Deduct (Plan B): {total_brokerage_fee:,.2f}")

        # คำนวณคอมมิชชั่นของกลุ่ม Standard Tier
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        commission_base = total_standard_sales - total_brokerage_fee - OPERATING_FEE
        
        t1, t2, t3 = 0, 0, 0
        amount_in_t1, amount_in_t2, amount_in_t3 = 0,0,0
        tier_commission = 0
        calculated_commission = 0

        if total_monthly_sales >= 500000:
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
            # รวมคอมมิชชั่นจากทุกส่วน
            calculated_commission = tier_commission + below_tier_commission
        else:
            below_tier_commission = 0
            calculated_commission = 0

        # --- สร้าง DataFrame สำหรับ Debug ---
        debug_details = []
        num_po = len(po_grouped_df['po_number'].unique())
        hit_target_percent = (total_monthly_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        cost_c1 = po_grouped_df['final_cost_amount'].sum()
        cost_c2 = (po_grouped_df['final_cost_amount'] * po_grouped_df['cost_multiplier']).sum()
        cost_c3 = (po_grouped_df['giveaways'] + po_grouped_df['coupon_fee'] + po_grouped_df['brokerage_fee']).sum()
        cost_c4_diff = (po_grouped_df['payment_before_vat'] - po_grouped_df['payment_no_vat']).sum()
        cost_c4_deduct = po_grouped_df['shipping_cost'].sum()
        
        shipping_profit_or_loss = cost_c4_diff - cost_c4_deduct
        shipping_adjustment_for_cost_report = min(shipping_profit_or_loss, 0)
        
        total_cost = cost_c2 + cost_c3 - po_grouped_df['difference_amount'].sum() - shipping_adjustment_for_cost_report

        debug_details.append({'รายการ': '## Report I: Sale Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'Sale Target KPI', 'ค่า': sales_target})
        debug_details.append({'รายการ': 'สรุปการขาย PO รายเดือน', 'ค่า': f"{num_po} บิล"})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน', 'ค่า': total_monthly_sales})
        debug_details.append({'รายการ': '  - ยอดขายปกติ (Standard Margin)', 'ค่า': total_standard_sales})
        debug_details.append({'รายการ': '  - ยอดขายนอกเงื่อนไข (Below Tier 1: 7.99-10%)', 'ค่า': total_below_tier1_sales})
        debug_details.append({'รายการ': '  - ยอดขายนอกเงื่อนไข (Below Tier 2: <7.99%)', 'ค่า': total_below_tier2_sales})
        debug_details.append({'รายการ': 'KPI Monthly SALE TARGET', 'ค่า': f"{hit_target_percent:.2f}%)"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})

        debug_details.append({'รายการ': '## Report II: Cost Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้า', 'ค่า': cost_c1})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้าบวกค่าใช้จ่ายบริหารจัดการ', 'ค่า': cost_c2})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าใช้จ่ายการตลาด', 'ค่า': cost_c3})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าบริการขนส่ง (ส่วนต่าง - ติดลบ)', 'ค่า': f"{cost_c4_diff:,.2f} - {cost_c4_deduct:,.2f}"})
        debug_details.append({'รายการ': 'ต้นทุนรวม', 'ค่า': total_cost})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
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
        debug_details.append({'รายการ': 'คอมมิชชั่นแบบขั้นบันได (Standard)', 'ค่า': tier_commission})
        debug_details.append({'รายการ': 'คอมมิชชั่นนอกเงื่อนไข (Below Tier 1 @ 0.63%)', 'ค่า': commission_below_t1})
        debug_details.append({'รายการ': 'คอมมิชชั่นนอกเงื่อนไข (Below Tier 2 @ 0.50%)', 'ค่า': commission_below_t2})
        debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})
        
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
            "คอมมิชชั่นนอกเงื่อนไข (Below Tier 1 @ 0.63%)",
            "คอมมิชชั่นนอกเงื่อนไข (Below Tier 2 @ 0.50%)",
            "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_monthly_sales, total_standard_sales, operating_fee,
            commission_base if commission_base > 0 else 0, t1, t2, t3,
            commission_below_t1, commission_below_t2,
            calculated_commission
        ]
        
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"]); summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        
        so_breakdown_df = po_grouped_df[['po_number', 'so_number', 'sales_service_amount', 'final_cost_amount', 'profit', 'margin']].copy()
        def assign_b_tier_status(margin):
            if margin >= 10: return 'Normal (>=10%)'
            if margin >= 7.99: return 'Below Tier (7.99-10%)' # <--- ลบเลข 1 ออก
            return 'Below Tier (<7.99%)'                 # <--- ลบเลข 2 ออก
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_b_tier_status)
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
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        # --- [แก้ไข] 1. เตรียมข้อมูล ---
        required_cols = [
            'total_revenue', 'sales_service_amount', 'final_cost_amount', 'giveaways', 
            'brokerage_fee', 'difference_amount', 'coupon_fee'
        ]
        for col in required_cols:
            if col not in comm_df.columns:
                comm_df[col] = 0

        total_revenue = pd.to_numeric(comm_df['total_revenue'], errors='coerce').fillna(0)
        sales_service_amount = pd.to_numeric(comm_df['sales_service_amount'], errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df['final_cost_amount'], errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df['giveaways'], errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df['difference_amount'], errors='coerce').fillna(0)
        
        if 'cost_multiplier' in comm_df.columns and not comm_df['cost_multiplier'].isnull().all():
            multiplier = pd.to_numeric(comm_df['cost_multiplier'], errors='coerce').fillna(1.03)
        else:
            multiplier = 1.03
            
        other_deductions = giveaways + brokerage + coupon_fee - difference_amount

        # --- 2. คำนวณกำไรและ Margin ---
        comm_df['profit'] = (total_revenue - (po_cost * multiplier)) + difference_amount
        comm_df['margin'] = (comm_df['profit'] / sales_service_amount.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)

        print("\n" + "---" * 20)
        print("### DEBUG: SO Breakdown for Total Raw Sales ###")
        for index, row in comm_df.iterrows():
            so_num = row['so_number']
            raw_sales = row['sales_service_amount']
            print(f" -> SO: {so_num}, Raw Sales: {raw_sales:,.2f}")
        print("---" * 20 + "\n")
        
        # --- 3. แบ่ง SO ตาม Tier ---
        tier1_df = comm_df[comm_df['margin'] >= 10]
        tier2_df = comm_df[(comm_df['margin'] >= 7.99) & (comm_df['margin'] < 10)]
        tier3_df = comm_df[comm_df['margin'] < 7.99]
        
        # --- 4. คำนวณฐานค่าคอม (จากยอดขายดิบ) ---
        total_sales_t1 = tier1_df['sales_service_amount'].sum()
        total_sales_t2 = tier2_df['sales_service_amount'].sum()
        total_sales_t3 = tier3_df['sales_service_amount'].sum()
        total_sales = total_sales_t1 + total_sales_t2 + total_sales_t3
        
        print(f"\n>>> DEBUGGING TOTAL SALES (Plan C): {total_sales:,.2f}\n")

        # (โค้ดส่วนที่เหลือทั้งหมดในการคำนวณ Commission, สร้าง Debug, และ Return)
        commission_t1, commission_t2, commission_t3 = 0.0, 0.0, 0.0
        calculated_commission = 0.0
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        base_t1, base_t2, base_t3 = 0.0, 0.0, 0.0
        debug_details = []
        num_so = len(comm_df['so_number'].unique())
        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"
        cost_c1 = po_cost.sum()
        cost_c2 = (po_cost * multiplier).sum()
        cost_c3 = (giveaways + coupon_fee + brokerage).sum()
        total_cost = cost_c2 + cost_c3 - difference_amount.sum()
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
        debug_details.append({'รายการ': 'ต้นทุนรวม', 'ค่า': total_cost})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        if total_sales >= 500000:
            # 1. รวมค่านายหน้าทั้งหมดของเดือน
            total_brokerage_fee = brokerage.sum()
            print(f">>> DEBUG: Total Brokerage Fee to Deduct (Plan C): {total_brokerage_fee:,.2f}")

            # 2. รวมยอดหักทั้งหมด (ค่าดำเนินการ + ค่านายหน้า)
            total_deduction_to_cascade = OPERATING_FEE + total_brokerage_fee

            # 3. ใช้วิธีหักลดหลั่นเหมือนเดิม แต่ใช้ยอดรวมใหม่
            base_t1 = max(0, total_sales_t1 - total_deduction_to_cascade)
            remaining_deduction_after_t1 = max(0, total_deduction_to_cascade - total_sales_t1)
            commission_t1 = round(base_t1 * 0.01, 2)
            
            base_t2 = max(0, total_sales_t2 - remaining_deduction_after_t1)
            remaining_deduction_after_t2 = max(0, remaining_deduction_after_t1 - total_sales_t2)
            commission_t2 = round(base_t2 * 0.0063, 2)
            
            base_t3 = max(0, total_sales_t3 - remaining_deduction_after_t2)
            commission_t3 = round(base_t3 * 0.005, 2)
            
            calculated_commission = commission_t1 + commission_t2 + commission_t3
            debug_details.append({'รายการ': '## การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (500,000)', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ยอดขาย Tier 1 (>=10%)', 'ค่า': total_sales_t1})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ', 'ค่า': min(total_sales_t1, total_deduction_to_cascade)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': base_t1})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (1.00%)', 'ค่า': commission_t1})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ยอดขาย Tier 2 (7.99-10%)', 'ค่า': total_sales_t2})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ (ส่วนที่เหลือ)', 'ค่า': min(total_sales_t2, remaining_deduction_after_t1)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': base_t2})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (0.63%)', 'ค่า': commission_t2})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ยอดขาย Tier 3 (<7.99%)', 'ค่า': total_sales_t3})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ (ส่วนที่เหลือ)', 'ค่า': min(total_sales_t3, remaining_deduction_after_t2)})
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

        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values()); gross_commission = calculated_commission + total_incentives
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        pre_tax_commission = gross_commission - total_additional_deductions 
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax
        final_calculated_commission = commission_t1 + commission_t2 + commission_t3
        summary_desc = [
            "ยอดขาย Tier 1 (>=10%)", "ยอดขาย Tier 2 (7.99-10%)", "ยอดขาย Tier 3 (<7.99%)", "ยอดขายรวม",
            "เงื่อนไขขั้นต่ำ (500,000)", "คอมฯ T1 (1.00%)", "คอมฯ T2 (0.63%)", "คอมฯ T3 (0.50%)",
            "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_sales_t1, total_sales_t2, total_sales_t3, total_sales,
            "ผ่าน" if total_sales >= 500000 else "ไม่ผ่าน",
            commission_t1, commission_t2, commission_t3,
            final_calculated_commission # <<< แก้ไข: ใช้ตัวแปรใหม่ที่นี่
        ]
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก: {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"]); summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        so_breakdown_df = comm_df[['so_number', 'sales_service_amount', 'shipping_cost', 'final_cost_amount', 'profit', 'margin']].copy()
        def assign_tier_status(margin):
            if margin >= 10: return "Normal (>=10%)"
            if margin >= 7.99: return "Below Tier (7.99-10%)"
            return "Below Tier (<7.99%)"
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_tier_status)
        so_breakdown_df.rename(columns={'so_number': 'SO Number', 'sales_service_amount': 'ยอดขาย', 'shipping_cost': 'ค่าส่ง', 'final_cost_amount': 'ต้นทุน', 'profit': 'กำไร', 'margin': 'Margin (%)'}, inplace=True)
        debug_df = pd.DataFrame(debug_details)
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission, # <<< แก้ไข: ใช้ตัวแปรใหม่ที่นี่ด้วย
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
        }
    
    elif plan_name == 'Plan D':

        if comm_df.empty:
            summary_data = {'description': ["ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"], 'value': [0.0]}
            return {'type': 'summary_other', 'data': pd.DataFrame(summary_data)}
        
        required_cols = [
            'total_revenue', 'sales_service_amount', 'final_cost_amount', 'giveaways', 
            'brokerage_fee', 'coupon_fee', 'difference_amount', 'payment_before_vat', 
            'payment_no_vat', 'shipping_cost', 'total_po_shipping_cost'
        ]
        for col in required_cols:
            if col not in comm_df.columns:
                comm_df[col] = 0
        
        # --- 1. ดึงข้อมูล (เวอร์ชันแก้ไข) ---
        total_revenue = pd.to_numeric(comm_df['total_revenue'], errors='coerce').fillna(0)
        sales_service_amount = pd.to_numeric(comm_df['sales_service_amount'], errors='coerce').fillna(0)
        po_cost = pd.to_numeric(comm_df['final_cost_amount'], errors='coerce').fillna(0)
        giveaways = pd.to_numeric(comm_df['giveaways'], errors='coerce').fillna(0)
        brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
        coupon_fee = pd.to_numeric(comm_df['coupon_fee'], errors='coerce').fillna(0)
        difference_amount = pd.to_numeric(comm_df['difference_amount'], errors='coerce').fillna(0)
        payment_before_vat = pd.to_numeric(comm_df['payment_before_vat'], errors='coerce').fillna(0)
        payment_no_vat = pd.to_numeric(comm_df['payment_no_vat'], errors='coerce').fillna(0)
        so_shipping_cost = pd.to_numeric(comm_df['shipping_cost'], errors='coerce').fillna(0)
        po_shipping_cost = pd.to_numeric(comm_df.get('total_po_shipping_cost', 0), errors='coerce').fillna(0)
        multiplier = pd.to_numeric(comm_df.get('cost_multiplier'), errors='coerce').fillna(1.03)

        other_deductions = giveaways + brokerage + coupon_fee
        net_shipping_deduction = (po_shipping_cost - so_shipping_cost).clip(lower=0)
        
        # --- 2. คำนวณ Profit/Margin (เวอร์ชันแก้ไข) ---
        comm_df['profit'] = (total_revenue - (po_cost * multiplier)) + difference_amount - net_shipping_deduction
        comm_df['margin'] = (comm_df['profit'] / sales_service_amount.replace(0, np.nan)) * 100
        comm_df['margin'] = comm_df['margin'].fillna(0)
        
        # --- 3. แบ่งกลุ่มและรวมยอดขาย (Logic เดิมของ Plan D) ---
        normal_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]
        total_normal_sales = normal_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_sales = total_normal_sales + total_below_sales
        
        # --- 4. คำนวณค่าคอมมิชชั่น ---
        normal_commission, below_tier_commission, calculated_commission = 0.0, 0.0, 0.0
        OPERATING_FEE = 750000.00 if operating_fee is None else operating_fee
        
        # +++ START: แก้ไข Logic การคำนวณทั้งหมด +++
        commission_base_normal = 0.0
        commission_base_below = 0.0

        if total_sales >= 750000:
            # 1. รวมค่านายหน้าทั้งหมดของเดือน
            total_brokerage_fee = brokerage.sum()
            print(f">>> DEBUG: Total Brokerage Fee to Deduct (Plan D): {total_brokerage_fee:,.2f}")

            # 2. รวมยอดหักทั้งหมดที่ต้องนำไปลดหลั่น
            total_deduction_to_cascade = OPERATING_FEE + total_brokerage_fee

            # 3. หักจาก Normal Tier ก่อน
            commission_base_normal = max(0, total_normal_sales - total_deduction_to_cascade)
            normal_commission = commission_base_normal * 0.007

            # 4. คำนวณยอดหักที่เหลือเพื่อนำไปหัก Below Tier ต่อ
            remaining_deduction = max(0, total_deduction_to_cascade - total_normal_sales)
            
            # 5. หักออกจาก Below Tier
            commission_base_below = max(0, total_below_sales - remaining_deduction)
            below_tier_commission = commission_base_below * 0.003
            
            # 6. รวมค่าคอมสุดท้าย
            calculated_commission = normal_commission + below_tier_commission
        # +++ END +++

        # --- 5. สร้าง debug_df (ขั้นตอนการคำนวณ) ---
        debug_details = []
        num_so = len(comm_df['so_number'].unique())
        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"
        cost_c1 = po_cost.sum()
        cost_c2 = (po_cost * multiplier).sum()
        cost_c3 = (giveaways + coupon_fee + brokerage).sum()
        cost_c4_diff = (payment_before_vat - payment_no_vat).sum()
        cost_c4_deduct = so_shipping_cost.sum()
        total_cost = cost_c2 + cost_c3 - difference_amount.sum()

        debug_details.append({'รายการ': '## Report I: Sale Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'Sale Target KPI', 'ค่า': sales_target})
        debug_details.append({'รายการ': 'สรุปการขาย SO รายเดือน', 'ค่า': f"{num_so} บิล"})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน', 'ค่า': total_sales})
        debug_details.append({'รายการ': 'KPI Monthly SALE TARGET', 'ค่า': f"{hit_target_status} ({hit_target_percent:.2f}%)"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## Report II: Cost Summary ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้า', 'ค่า': cost_c1})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าสินค้าบวกค่าใช้จ่ายบริหารจัดการ', 'ค่า': cost_c2})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าใช้จ่ายการตลาด', 'ค่า': cost_c3})
        debug_details.append({'รายการ': 'สรุปต้นทุน: ค่าบริการขนส่ง (ขาดทุน)', 'ค่า': net_shipping_deduction.sum()})
        debug_details.append({'รายการ': 'ต้นทุนรวม', 'ค่า': total_cost})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_sales >= 750000:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (750,000)', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            
            # +++ START: อัปเดต Debug Report ให้สอดคล้องกับ Logic ใหม่ +++
            debug_details.append({'รายการ': 'คอมมิชชั่นนอกเงื่อนไข (Below Tier)', 'ค่า': ''})
            debug_details.append({'รายการ': '  ยอดขายดิบ', 'ค่า': total_below_sales})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ/นายหน้า (ส่วนที่เหลือ)', 'ค่า': min(total_below_sales, remaining_deduction)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': commission_base_below})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (0.30%)', 'ค่า': below_tier_commission})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'คอมมิชชั่นปกติ (Normal)', 'ค่า': ''})
            debug_details.append({'รายการ': '  ยอดขายดิบ', 'ค่า': total_normal_sales})
            debug_details.append({'รายการ': '  (-) หักค่าดำเนินการ/นายหน้า', 'ค่า': min(total_normal_sales, total_deduction_to_cascade)})
            debug_details.append({'รายการ': '  = ฐานคำนวณ', 'ค่า': commission_base_normal})
            debug_details.append({'รายการ': '  * คอมมิชชั่น (0.70%)', 'ค่า': normal_commission})
            # +++ END +++
        else:
            debug_details.append({'รายการ': 'เงื่อนไขยอดขายขั้นต่ำ (750,000)', 'ค่า': 'ไม่ผ่าน ❌'})
        
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## สรุปค่าคอมมิชชั่น ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'คอมมิชชั่นปกติ (Normal)', 'ค่า': normal_commission})
        debug_details.append({'รายการ': 'คอมมิชชั่นนอกเงื่อนไข (Below Tier)', 'ค่า': below_tier_commission})
        debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': calculated_commission})
        debug_df = pd.DataFrame(debug_details)
        
        # 6. สร้าง so_breakdown_df
        so_breakdown_df = comm_df[['so_number', 'sales_service_amount', 'final_cost_amount', 'profit', 'margin']].copy()
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        so_breakdown_df.rename(columns={
            'so_number': 'SO Number', 'sales_service_amount': 'ยอดขาย',
            'final_cost_amount': 'ต้นทุน', 'profit': 'กำไร',
            'margin': 'Margin (%)'
        }, inplace=True)
        
        # 7. สรุปผลสุดท้าย
        if incentives is None: incentives = {}
        total_incentives = sum(incentives.values())
        gross_commission = calculated_commission + total_incentives
        if additional_deductions is None: additional_deductions = {}
        total_additional_deductions = sum(additional_deductions.values())
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax
        
        summary_desc = [
            "ยอดขาย Normal (>=10%)", "ยอดขาย Below T (<10%)", "(-) ค่าดำเนินการ/นายหน้า", 
            "ฐานคอม Normal", "ฐานคอม Below Tier", "คอมมิชชั่น Normal (0.70%)", 
            "คอมมิชชั่น Below Tier (0.30%)", "ยอดรวมค่าคอมมิชชั่นที่คำนวณได้"
        ]
        summary_val = [
            total_normal_sales, total_below_sales, OPERATING_FEE + total_brokerage_fee,
            commission_base_normal, commission_base_below,
            normal_commission, below_tier_commission, calculated_commission
        ]
        
        for key, value in incentives.items(): summary_desc.append(f"(+) Incentive: {key}"); summary_val.append(value)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        for key, value in additional_deductions.items(): summary_desc.append(f"(-) หัก {key}"); summary_val.append(value)
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"]); summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        summary_data = {'description': summary_desc, 'value': summary_val}
        
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame(summary_data),
            'final_commission': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
        }
        
    else:   
         return {'type': 'error', 'message': f'ไม่พบ Plan ที่ชื่อว่า {plan_name}'}