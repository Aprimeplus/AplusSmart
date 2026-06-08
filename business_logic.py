import pandas as pd
import numpy as np

def calculate_monthly_commission(plan_name, comm_df, sales_target=0, operating_fee=None, 
                                 additional_deductions=None, incentives=None, 
                                 min_sales_target=500000):
    """
    Calculates the monthly commission based on the specified plan.
    (Updated: Correct Mapping for Relocation Cost from Delivery Note)
    """
    total_brokerage_fee = 0.0

    # --- เตรียมตัวแปรสำหรับสรุปผลสุดท้าย (Incentive/Deduction) ---
    if incentives is None: incentives = {}
    total_incentives = sum(incentives.values())
    
    if additional_deductions is None: additional_deductions = {}
    total_additional_deductions = sum(additional_deductions.values())

    def prepare_and_calculate_profit(df):
        # 1. Fill NA ให้ครบทุกฟิลด์
        cols_to_fix = [
            'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee',
            'difference_amount', 'payment_before_vat', 'payment_no_vat',
            'shipping_cost', 'relocation_cost', 'so_cutting_rev', 'so_service_rev',
            'shipping_to_site_cost', 'shipping_to_stock_cost', 'po_cutting_cost', 'po_service_cost',
            # ค่าตัด/เจาะ reconciliation columns
            'cutting_drilling_fee',
            'po_cutting_vat_cost', 'po_cutting_cash_cost', 'po_cutting_item_cost',
        ]
        
        for col in cols_to_fix:
            if col not in df.columns: df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # 2. จัดการ Multiplier
        if 'cost_multiplier' in df.columns and not df['cost_multiplier'].isnull().all():
             df['cost_multiplier'] = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
             df['cost_multiplier'] = df['cost_multiplier'].apply(lambda x: 1.03 if x < 1.01 else x)
        else:
             df['cost_multiplier'] = 1.03

        # =========================================================
        # [🔥 MATCHING LOGIC] 
        # =========================================================
        
        # --- A. คำนวณส่วนต่างค่าขนส่ง (Excess Cost) ---
        total_shipping_revenue = df['shipping_cost'] + df['relocation_cost']
        total_shipping_cost = df['shipping_to_site_cost'] + df['shipping_to_stock_cost']
        
        # Logic: ถ้าขาดทุน (ต้นทุน > รายรับ) ให้เก็บยอดขาดทุนไว้หักกำไร
        deduct_shipping_total = (total_shipping_cost - total_shipping_revenue).clip(lower=0)
        
        # บันทึกค่าลง DataFrame
        df['excess_stock_shipping'] = deduct_shipping_total
        df['excess_site_shipping'] = 0.0

        # --- A2. ค่าตัด/เจาะ Reconciliation ────────────────────────────────
        # Sale เลือก VAT → match กับ PO cutting (VAT type) + EXP items ในรายการสินค้า
        # Sale เลือก CASH → match กับ PO cutting (CASH type)
        # ถ้า SO < PO → ส่วนต่างหักกำไร (เหมือน logic ค่ารถ)
        cutting_vat_option = df.get('cutting_drilling_fee_vat_option', pd.Series([''] * len(df), index=df.index))
        so_cutting = df['cutting_drilling_fee']

        # VAT case: SO cutting vs (PO cutting VAT + EXP items)
        po_cutting_vat_total = df['po_cutting_vat_cost'] + df['po_cutting_item_cost']
        so_cutting_vat = so_cutting.where(cutting_vat_option == 'VAT', 0.0)
        excess_cutting_vat = (po_cutting_vat_total - so_cutting_vat).clip(lower=0)

        # CASH case: SO cutting vs PO cutting CASH
        so_cutting_cash = so_cutting.where(cutting_vat_option == 'CASH', 0.0)
        excess_cutting_cash = (df['po_cutting_cash_cost'] - so_cutting_cash).clip(lower=0)

        df['excess_cutting'] = excess_cutting_vat + excess_cutting_cash

        # --- B. คำนวณต้นทุนสินค้า (Main Cost Calculation) ---
        # [ตาม HR Excel]: ทุกต้นทุนรวม (รวมค่าตัด/ค่าบริการ) × 1.03 ทั้งหมด
        total_cost_calculated = df['final_cost_amount'] * df['cost_multiplier']

        # --- C. คำนวณกำไรสุทธิ (Final Profit) ---
        df['profit'] = (df['sales_service_amount'] - total_cost_calculated) \
                        + df['difference_amount'] \
                        - deduct_shipping_total \
                        - df['excess_cutting']
        
        # คำนวณ Margin %
        df['margin'] = (df['profit'] / df['sales_service_amount'].replace(0, np.nan)) * 100
        df['margin'] = df['margin'].fillna(0)

        return df

    # ==================================================================================
    # PLAN A
    # ==================================================================================
    if plan_name == 'Plan A':
        comm_df = prepare_and_calculate_profit(comm_df)
        total_sales = comm_df['sales_service_amount'].sum()
        
        initial_commission = 0.0
        calculated_commission = 0.0
        commission_normal = 0.0
        commission_below = 0.0
        total_brokerage_fee = 0.0
        
        OPERATING_FEE = 25000.00 if operating_fee is None else operating_fee
        
        if total_sales >= min_sales_target:
            NORMAL_RATE = 0.35   
            BELOW_T_RATE = 0.175 

            normal_df = comm_df[comm_df['margin'] >= 10]
            below_df = comm_df[comm_df['margin'] < 10]

            val_normal_sales = normal_df['sales_service_amount'].sum()
            val_below_sales = below_df['sales_service_amount'].sum()
            
            commission_normal = normal_df['profit'].sum() * NORMAL_RATE
            commission_below = below_df['profit'].sum() * BELOW_T_RATE
            initial_commission = commission_normal + commission_below
            
            brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
            total_brokerage_fee = brokerage.sum()
            
            calculated_commission = max(0, initial_commission - total_brokerage_fee - OPERATING_FEE)

            conditions = [comm_df['margin'] >= 10, comm_df['margin'] < 10]
            choices_commission = [comm_df['profit'] * NORMAL_RATE, comm_df['profit'] * BELOW_T_RATE]
            comm_df['commission_amount'] = np.select(conditions, choices_commission, default=0)
        else:
            val_normal_sales = 0.0
            val_below_sales = 0.0
            comm_df['commission_amount'] = 0.0
            initial_commission = 0.0
            calculated_commission = 0.0

        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        kpi_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0.0

        debug_info = []
        debug_info.append({"รายการ": "## 1. ข้อมูลภาพรวม (Plan A)", "ค่า": ""})
        debug_info.append({"รายการ": "ยอดขายรวม (ทั้งหมด)", "ค่า": f"{total_sales:,.2f}"})
        
        debug_info.append({"รายการ": "เป้าหมายการขาย (KPI Target)", "ค่า": f"{sales_target:,.2f}"})
        if sales_target > 0:
            debug_info.append({"รายการ": "KPI Achievement", "ค่า": f"{kpi_percent:,.2f}%"})
        else:
            debug_info.append({"รายการ": "KPI Achievement", "ค่า": "(ไม่มีเป้า KPI)"})

        debug_info.append({"รายการ": "---", "ค่า": "---"})
        
        debug_info.append({"รายการ": "ยอดขายขั้นต่ำ (เงื่อนไขรับค่าคอม)", "ค่า": f"{min_sales_target:,.2f}"})
        
        if total_sales >= min_sales_target:
            debug_info.append({"รายการ": "สถานะยอดขายขั้นต่ำ", "ค่า": "✅ ผ่านเกณฑ์"})
            
            debug_info.append({"รายการ": "---", "ค่า": "---"})
            debug_info.append({"รายการ": "## 2. การคำนวณจากกำไร (Profit-based)", "ค่า": ""})
            debug_info.append({"รายการ": "ยอดขาย Normal (Margin >= 10%)", "ค่า": f"{val_normal_sales:,.2f}"})
            debug_info.append({"รายการ": "> ได้ค่าคอมฯ (35% ของกำไร)", "ค่า": f"{commission_normal:,.2f}"})
            debug_info.append({"รายการ": "ยอดขาย Below Tier (Margin < 10%)", "ค่า": f"{val_below_sales:,.2f}"})
            debug_info.append({"รายการ": "> ได้ค่าคอมฯ (17.5% ของกำไร)", "ค่า": f"{commission_below:,.2f}"})
            debug_info.append({"รายการ": "รวมยอดคอมมิชชั่นขั้นต้น (ก่อนหักค่าใช้จ่าย)", "ค่า": f"{initial_commission:,.2f}"})

            debug_info.append({"รายการ": "---", "ค่า": "---"})
            debug_info.append({"รายการ": "## 3. หักค่าใช้จ่ายและค่าดำเนินการ", "ค่า": ""})
            debug_info.append({"รายการ": "(-) หัก ค่านายหน้า (Brokerage Fee)", "ค่า": f"{total_brokerage_fee:,.2f}"})
            debug_info.append({"รายการ": "(-) หัก ค่าดำเนินการ (Operating Fee)", "ค่า": f"{OPERATING_FEE:,.2f}"})
            debug_info.append({"รายการ": "ยอดรวมค่าคอมที่คำนวณได้ (ไม่ติดลบ)", "ค่า": f"{calculated_commission:,.2f}"})
        else:
            debug_info.append({"รายการ": "สถานะยอดขายขั้นต่ำ", "ค่า": "❌ ไม่ผ่านเกณฑ์ (ยอดขายไม่ถึงขั้นต่ำ)"})
            debug_info.append({"รายการ": "รวมคอมมิชชั่น", "ค่า": "0.00"})

        debug_info.append({"รายการ": "---", "ค่า": "---"})
        debug_info.append({"รายการ": "## 4. สรุปยอดสุทธิ", "ค่า": ""})
        debug_info.append({"รายการ": "คอมมิชชั่นที่คำนวณได้รวม", "ค่า": f"{calculated_commission:,.2f}"})
        
        for k, v in incentives.items():
            debug_info.append({"รายการ": f"(+) {k}", "ค่า": f"{v:,.2f}"})
            
        debug_info.append({"รายการ": "ยอดคอมมิชชั่นขั้นต้น (Gross)", "ค่า": f"{gross_commission:,.2f}"})
        
        for k, v in additional_deductions.items():
            debug_info.append({"รายการ": f"(-) {k}", "ค่า": f"{v:,.2f}"})
            
        debug_info.append({"รายการ": "ยอดก่อนหักภาษี", "ค่า": f"{pre_tax_commission:,.2f}"})
        debug_info.append({"รายการ": "(-) ภาษีหัก ณ ที่จ่าย (3%)", "ค่า": f"{withholding_tax:,.2f}"})
        debug_info.append({"รายการ": "ยอดโอนสุทธิ (Net)", "ค่า": f"{net_commission:,.2f}"})
        
        debug_df_output = pd.DataFrame(debug_info)

        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'shipping_to_stock_cost', 'shipping_to_site_cost', 
            'profit', 'margin', 'commission_amount', 'cost_multiplier'
        ]].copy()
        
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        
        # 🔥 แก้บัก SO Number ตรงนี้: จัดเรียงก่อนเปลี่ยนชื่อ!
        so_breakdown_df.sort_values(by='so_number', ascending=True, inplace=True)
        so_breakdown_df.reset_index(drop=True, inplace=True)

        so_breakdown_df.rename(columns={
            'so_number': 'SO Number',
            'sales_service_amount': 'ยอดขาย', 
            'final_cost_amount': 'ต้นทุนสินค้า', 
            'shipping_to_stock_cost': 'ค่าย้าย(Stock)',
            'shipping_to_site_cost': 'ค่ารถ(Site)',
            'profit': 'กำไรสุทธิ', 
            'margin': 'Margin (%)', 
            'commission_amount': 'ค่าคอมฯ (ก่อนหัก)'
        }, inplace=True)
        
        summary_desc = [
            "ยอดขาย Normal", "ยอดขาย Below Tier", "รวมยอดคอมมิชชั่นจาก Profit",
            "(-) หัก ค่านายหน้า (Brokerage)", "(-) หัก ค่าดำเนินการ (Operating Fee)",
            "ยอดรวมค่าคอมที่คำนวณได้"
        ]
        summary_val = [
            val_normal_sales, val_below_sales, initial_commission,
            total_brokerage_fee, OPERATING_FEE, calculated_commission
        ]
        
        for k, v in incentives.items(): summary_desc.append(f"(+) {k}"); summary_val.append(v)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross)"); summary_val.append(gross_commission)
        for k, v in additional_deductions.items(): summary_desc.append(f"(-) {k}"); summary_val.append(v)
        
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        return {
            'type': 'summary_plan_a',
            'summary': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'so_breakdown_df': so_breakdown_df,
            'debug_df': debug_df_output 
        }

    # ==================================================================================
    # PLAN B
    # ==================================================================================
    elif plan_name == 'Plan B':
        if comm_df.empty: 
            return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        comm_df = prepare_and_calculate_profit(comm_df)

        if 'po_number' not in comm_df.columns: 
            comm_df['po_number'] = comm_df['so_number']
        
        agg_rules = {
            'sales_service_amount': 'sum',
            'giveaways': 'sum',
            'brokerage_fee': 'sum',
            'difference_amount': 'sum',
            'final_cost_amount': 'sum',
            'cost_multiplier': 'first',
            'so_number': lambda x: ', '.join(sorted(set(str(v) for v in x))),
            'profit': 'sum',
            'po_cutting_cost': 'sum',
            'po_service_cost': 'sum',
            'shipping_to_stock_cost': 'sum',
            'shipping_to_site_cost': 'sum',
            'relocation_cost': 'sum',
            'shipping_cost': 'sum',
            'excess_site_shipping': 'sum',
            'excess_stock_shipping': 'sum',
            'excess_cutting': 'sum',          # ค่าตัด/เจาะ excess
        }
        
        valid_agg_rules = {k: v for k, v in agg_rules.items() if k in comm_df.columns}
        po_grouped_df = comm_df.groupby('po_number').agg(valid_agg_rules).reset_index()

        po_grouped_df['margin'] = (po_grouped_df['profit'] / po_grouped_df['sales_service_amount'].replace(0, np.nan)) * 100    
        po_grouped_df['margin'] = po_grouped_df['margin'].fillna(0)

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
        commission_base = max(0, total_standard_sales - total_brokerage_fee - OPERATING_FEE)
        
        t1, t2, t3, tier_commission, calculated_commission = 0, 0, 0, 0, 0

        if total_monthly_sales >= min_sales_target:
            remaining_base = commission_base
            
            amount_in_t1 = min(remaining_base, 1000000)
            t1 = amount_in_t1 * 0.0125
            remaining_base -= amount_in_t1
            
            if remaining_base > 0:
                amount_in_t2 = min(remaining_base, 1000000)
                t2 = amount_in_t2 * 0.0175
                remaining_base -= amount_in_t2
            
            if remaining_base > 0:
                t3 = remaining_base * 0.0225
            
            tier_commission = t1 + t2 + t3
            calculated_commission = tier_commission + below_tier_commission
        else:
            calculated_commission = 0.0

        commission_base_normal = commission_base

        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        kpi_percent = (total_monthly_sales / sales_target * 100) if sales_target > 0 else 0.0

        debug_info = []
        debug_info.append({"รายการ": "## 1. ข้อมูลภาพรวม (Plan B)", "ค่า": ""})
        debug_info.append({"รายการ": "ยอดขายรวม (ทั้งหมด)", "ค่า": f"{total_monthly_sales:,.2f}"})
        
        debug_info.append({"รายการ": "เป้าหมายการขาย (KPI Target)", "ค่า": f"{sales_target:,.2f}"})
        if sales_target > 0:
            debug_info.append({"รายการ": "KPI Achievement", "ค่า": f"{kpi_percent:,.2f}%"})
        else:
            debug_info.append({"รายการ": "KPI Achievement", "ค่า": "(ไม่มีเป้า KPI)"})

        debug_info.append({"รายการ": "---", "ค่า": "---"})
        
        debug_info.append({"รายการ": "ยอดขายขั้นต่ำ (เงื่อนไขรับค่าคอม)", "ค่า": f"{min_sales_target:,.2f}"})
        
        if total_monthly_sales >= min_sales_target:
            debug_info.append({"รายการ": "สถานะยอดขายขั้นต่ำ", "ค่า": "✅ ผ่านเกณฑ์"})
            
            debug_info.append({"รายการ": "---", "ค่า": "---"})
            debug_info.append({"รายการ": "## 2. การคำนวณ Normal Tier (Margin >= 10%)", "ค่า": ""})
            debug_info.append({"รายการ": "ยอดขาย Normal", "ค่า": f"{total_standard_sales:,.2f}"})
            debug_info.append({"รายการ": "(-) ค่านายหน้า", "ค่า": f"{total_brokerage_fee:,.2f}"})
            debug_info.append({"รายการ": "(-) ค่าดำเนินการ (Operating Fee)", "ค่า": f"{OPERATING_FEE:,.2f}"})
            debug_info.append({"รายการ": "ฐานการคำนวณ Normal", "ค่า": f"{commission_base:,.2f}"})
            debug_info.append({"รายการ": "> Tier 1 (0-1M) 1.25%", "ค่า": f"{t1:,.2f}"})
            debug_info.append({"รายการ": "> Tier 2 (1-2M) 1.75%", "ค่า": f"{t2:,.2f}"})
            debug_info.append({"รายการ": "> Tier 3 (>2M) 2.25%", "ค่า": f"{t3:,.2f}"})
            debug_info.append({"รายการ": "รวมคอมมิชชั่น Normal", "ค่า": f"{tier_commission:,.2f}"})

            debug_info.append({"รายการ": "---", "ค่า": "---"})
            debug_info.append({"รายการ": "## 3. การคำนวณ Below Tier", "ค่า": ""})
            debug_info.append({"รายการ": "ยอดขาย Tier 2 (Margin 7.99-9.99%)", "ค่า": f"{total_below_tier1_sales:,.2f}"})
            debug_info.append({"รายการ": "ยอดขาย Tier 3 (Margin < 7.99%)", "ค่า": f"{total_below_tier2_sales:,.2f}"})
            debug_info.append({"รายการ": "> ค่าคอม Tier 2 (0.63%)", "ค่า": f"{commission_below_t1:,.2f}"})
            debug_info.append({"รายการ": "> ค่าคอม Tier 3 (0.50%)", "ค่า": f"{commission_below_t2:,.2f}"})
            debug_info.append({"รายการ": "รวมคอมมิชชั่น Below Tier", "ค่า": f"{below_tier_commission:,.2f}"})

        else:
            debug_info.append({"รายการ": "สถานะยอดขายขั้นต่ำ", "ค่า": "❌ ไม่ผ่านเกณฑ์ (ยอดขายไม่ถึงขั้นต่ำ)"})
            debug_info.append({"รายการ": "รวมคอมมิชชั่น", "ค่า": "0.00"})

        debug_info.append({"รายการ": "---", "ค่า": "---"})
        debug_info.append({"รายการ": "## 4. สรุปยอดสุทธิ", "ค่า": ""})
        debug_info.append({"รายการ": "คอมมิชชั่นที่คำนวณได้รวม", "ค่า": f"{calculated_commission:,.2f}"})
        
        for k, v in incentives.items():
            debug_info.append({"รายการ": f"(+) {k}", "ค่า": f"{v:,.2f}"})
            
        debug_info.append({"รายการ": "ยอดคอมมิชชั่นขั้นต้น (Gross)", "ค่า": f"{gross_commission:,.2f}"})
        
        for k, v in additional_deductions.items():
            debug_info.append({"รายการ": f"(-) {k}", "ค่า": f"{v:,.2f}"})
            
        debug_info.append({"รายการ": "ยอดก่อนหักภาษี", "ค่า": f"{pre_tax_commission:,.2f}"})
        debug_info.append({"รายการ": "(-) ภาษีหัก ณ ที่จ่าย (3%)", "ค่า": f"{withholding_tax:,.2f}"})
        debug_info.append({"รายการ": "ยอดโอนสุทธิ (Net)", "ค่า": f"{net_commission:,.2f}"})
        
        debug_df_output = pd.DataFrame(debug_info)

        summary_desc = [
            "ยอดขายรวม", "ยอดขาย Normal (Margin >= 10%)", "(-) หัก ค่าดำเนินการ (Operating Fee)", 
            "ฐานคำนวณคอมฯ Normal", "ยอดรวมค่าคอมที่คำนวณได้"
        ]
        summary_val = [
            total_monthly_sales, total_standard_sales, OPERATING_FEE, 
            commission_base_normal, calculated_commission
        ]
        
        for k, v in incentives.items(): summary_desc.append(f"(+) {k}"); summary_val.append(v)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross)"); summary_val.append(gross_commission)
        for k, v in additional_deductions.items(): summary_desc.append(f"(-) {k}"); summary_val.append(v)
        
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])

        so_breakdown_df = po_grouped_df[[
            'po_number', 'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'shipping_to_stock_cost', 'shipping_to_site_cost',
            'excess_stock_shipping', 'excess_site_shipping',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        def assign_b_tier_status(margin):
            if margin >= 10: return 'Normal (>=10%)'
            if margin >= 7.99: return 'Below Tier (7.99-10%)'
            return 'Below Tier (<7.99%)'
            
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_b_tier_status)
        
        # 🔥 แก้บัก SO Number ตรงนี้: จัดเรียงก่อนเปลี่ยนชื่อ!
        so_breakdown_df.sort_values(by='so_number', ascending=True, inplace=True)
        so_breakdown_df.reset_index(drop=True, inplace=True)

        so_breakdown_df.rename(columns={
            'sales_service_amount': 'ยอดขายรวม',
            'final_cost_amount': 'ต้นทุนรวม',
            'po_cutting_cost': 'ทุนค่าตัด',
            'po_service_cost': 'ทุนบริการ',
            'shipping_to_stock_cost': 'ทุนเข้าStock',
            'shipping_to_site_cost': 'ทุนเข้าSite',
            'excess_stock_shipping': 'หักส่วนต่างStock',
            'excess_site_shipping': 'หักส่วนต่างSite',
            'profit': 'กำไรสุทธิ',
            'margin': 'Margin (%)'
        }, inplace=True)
        
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'so_breakdown_df': so_breakdown_df,
            'debug_df': debug_df_output 
        }

    # ==================================================================================
    # PLAN C (Logic คล้าย Plan B)
    # ==================================================================================
    elif plan_name == 'Plan C':
        if comm_df.empty: 
            return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        comm_df = prepare_and_calculate_profit(comm_df)

        tier1_df = comm_df[comm_df['margin'] >= 10]
        tier1_sales = tier1_df['sales_service_amount'].sum()
        
        tier2_df = comm_df[(comm_df['margin'] >= 7.99) & (comm_df['margin'] < 10)]
        tier2_sales = tier2_df['sales_service_amount'].sum()
        
        tier3_df = comm_df[comm_df['margin'] < 7.99]
        tier3_sales = tier3_df['sales_service_amount'].sum()
        
        total_sales = tier1_sales + tier2_sales + tier3_sales
        
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        total_brokerage = comm_df['brokerage_fee'].sum()
        total_deduction = OPERATING_FEE + total_brokerage

        comm_t1, comm_t2, comm_t3, calculated_commission = 0, 0, 0, 0
        base_t1, base_t2, base_t3 = 0, 0, 0 
        
        if total_sales >= min_sales_target:
            base_t1 = max(0, tier1_sales - total_deduction)
            remaining_deduction = max(0, total_deduction - tier1_sales) 
            comm_t1 = round(base_t1 * 0.01, 2) 
            
            base_t2 = max(0, tier2_sales - remaining_deduction)
            remaining_deduction = max(0, remaining_deduction - tier2_sales) 
            comm_t2 = round(base_t2 * 0.0063, 2) 
            
            base_t3 = max(0, tier3_sales - remaining_deduction)
            comm_t3 = round(base_t3 * 0.005, 2) 
            
            calculated_commission = comm_t1 + comm_t2 + comm_t3
        else:
            calculated_commission = 0.0

        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        kpi_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0.0

        debug_details = []
        debug_details.append({"รายการ": "## 1. ข้อมูลภาพรวม (Plan C)", "ค่า": ""})
        debug_details.append({"รายการ": "ยอดขายรวม (ทั้งหมด)", "ค่า": f"{total_sales:,.2f}"})
        
        debug_details.append({"รายการ": "เป้าหมายการขาย (KPI Target)", "ค่า": f"{sales_target:,.2f}"})
        if sales_target > 0:
            debug_details.append({"รายการ": "KPI Achievement", "ค่า": f"{kpi_percent:,.2f}%"})
        else:
            debug_details.append({"รายการ": "KPI Achievement", "ค่า": "(ไม่มีเป้า KPI)"})

        debug_details.append({"รายการ": "---", "ค่า": "---"})
        
        debug_details.append({"รายการ": "ยอดขายขั้นต่ำ (เงื่อนไขรับค่าคอม)", "ค่า": f"{min_sales_target:,.2f}"})
        
        if total_sales >= min_sales_target:
            debug_details.append({"รายการ": "สถานะยอดขายขั้นต่ำ", "ค่า": "✅ ผ่านเกณฑ์"})
            
            debug_details.append({"รายการ": "---", "ค่า": "---"})
            debug_details.append({"รายการ": "## 2. การคำนวณคอมมิชชั่น (Waterfall Logic)", "ค่า": ""})
            debug_details.append({"รายการ": "ยอดหักรวม (ค่าดำเนินการ + นายหน้า)", "ค่า": f"{total_deduction:,.2f}"})
            
            debug_details.append({"รายการ": "ยอดขาย Tier 1 (Margin >= 10%)", "ค่า": f"{tier1_sales:,.2f}"})
            debug_details.append({"รายการ": "ฐานคำนวณ Tier 1 (หลังหัก)", "ค่า": f"{base_t1:,.2f}"})
            debug_details.append({"รายการ": "> คอมฯ Tier 1 (1.0%)", "ค่า": f"{comm_t1:,.2f}"})
            
            debug_details.append({"รายการ": "ยอดขาย Tier 2 (Margin 7.99-9.99%)", "ค่า": f"{tier2_sales:,.2f}"})
            debug_details.append({"รายการ": "ฐานคำนวณ Tier 2 (หลังหัก)", "ค่า": f"{base_t2:,.2f}"})
            debug_details.append({"รายการ": "> คอมฯ Tier 2 (0.63%)", "ค่า": f"{comm_t2:,.2f}"})
            
            debug_details.append({"รายการ": "ยอดขาย Tier 3 (Margin < 7.99%)", "ค่า": f"{tier3_sales:,.2f}"})
            debug_details.append({"รายการ": "ฐานคำนวณ Tier 3 (หลังหัก)", "ค่า": f"{base_t3:,.2f}"})
            debug_details.append({"รายการ": "> คอมฯ Tier 3 (0.50%)", "ค่า": f"{comm_t3:,.2f}"})

            debug_details.append({"รายการ": "ยอดรวมค่าคอมที่คำนวณได้", "ค่า": f"{calculated_commission:,.2f}"})
        else:
            debug_details.append({"รายการ": "สถานะยอดขายขั้นต่ำ", "ค่า": "❌ ไม่ผ่านเกณฑ์ (ยอดขายไม่ถึงขั้นต่ำ)"})
            debug_details.append({"รายการ": "รวมคอมมิชชั่น", "ค่า": "0.00"})

        debug_details.append({"รายการ": "---", "ค่า": "---"})
        debug_details.append({"รายการ": "## 3. สรุปยอดสุทธิ", "ค่า": ""})
        debug_details.append({"รายการ": "คอมมิชชั่นที่คำนวณได้รวม", "ค่า": f"{calculated_commission:,.2f}"})
        
        for k, v in incentives.items():
            debug_details.append({"รายการ": f"(+) {k}", "ค่า": f"{v:,.2f}"})
            
        debug_details.append({"รายการ": "ยอดคอมมิชชั่นขั้นต้น (Gross)", "ค่า": f"{gross_commission:,.2f}"})
        
        for k, v in additional_deductions.items():
            debug_details.append({"รายการ": f"(-) {k}", "ค่า": f"{v:,.2f}"})
            
        debug_details.append({"รายการ": "ยอดก่อนหักภาษี", "ค่า": f"{pre_tax_commission:,.2f}"})
        debug_details.append({"รายการ": "(-) ภาษีหัก ณ ที่จ่าย (3%)", "ค่า": f"{withholding_tax:,.2f}"})
        debug_details.append({"รายการ": "ยอดโอนสุทธิ (Net)", "ค่า": f"{net_commission:,.2f}"})

        summary_desc = ["ยอดขาย Tier 1", "ยอดขาย Tier 2", "ยอดขาย Tier 3", "ยอดรวมคอมที่คำนวณได้"]
        summary_val = [tier1_sales, tier2_sales, tier3_sales, calculated_commission]
        
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        
        for k, v in additional_deductions.items():
            summary_desc.append(f"(-) หัก {k}")
            summary_val.append(v)

        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'shipping_to_stock_cost', 'shipping_to_site_cost',
            'excess_stock_shipping', 'excess_site_shipping',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        def assign_c_tier_status(margin):
            if margin >= 10: return 'Tier 1 (>=10%)'
            if margin >= 7.99: return 'Tier 2 (7.99-9.99%)'
            return 'Tier 3 (<7.99%)'
            
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_c_tier_status)
        
        # 🔥 แก้บัก SO Number ตรงนี้: จัดเรียงก่อนเปลี่ยนชื่อ!
        so_breakdown_df.sort_values(by='so_number', ascending=True, inplace=True)
        so_breakdown_df.reset_index(drop=True, inplace=True)

        so_breakdown_df.rename(columns={
            'so_number': 'SO Number', 
            'sales_service_amount': 'ยอดขาย', 
            'final_cost_amount': 'ต้นทุนรวม', 
            'po_cutting_cost': 'ทุนตัด',
            'po_service_cost': 'ทุนบริการ',
            'shipping_to_stock_cost': 'ทุนเข้าStock',
            'shipping_to_site_cost': 'ทุนเข้าSite',
            'excess_stock_shipping': 'หักส่วนต่างStock',
            'excess_site_shipping': 'หักส่วนต่างSite',
            'profit': 'กำไร', 
            'margin': 'Margin (%)'
        }, inplace=True)
        
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'debug_df': pd.DataFrame(debug_details),
            'so_breakdown_df': so_breakdown_df
        }

    # ==================================================================================
    # PLAN D (Logic คล้าย Plan C แต่ Fee/Rate ต่างกัน)
    # ==================================================================================
    elif plan_name == 'Plan D':
        if comm_df.empty: 
            return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        comm_df = prepare_and_calculate_profit(comm_df)
        
        normal_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]
        
        total_normal_sales = normal_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_sales = total_normal_sales + total_below_sales
        
        normal_commission = 0.0
        below_tier_commission = 0.0
        calculated_commission = 0.0
        
        OPERATING_FEE = 750000.00 if operating_fee is None else operating_fee
        commission_base_normal = 0.0
        commission_base_below = 0.0

        if total_sales >= min_sales_target:
            brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
            total_brokerage_fee = brokerage.sum()
            
            total_deduction_to_cascade = OPERATING_FEE + total_brokerage_fee

            commission_base_normal = max(0, total_normal_sales - total_deduction_to_cascade)
            normal_commission = commission_base_normal * 0.007 

            remaining_deduction = max(0, total_deduction_to_cascade - total_normal_sales)
            
            commission_base_below = max(0, total_below_sales - remaining_deduction)
            below_tier_commission = commission_base_below * 0.003 
            
            calculated_commission = normal_commission + below_tier_commission
        else:
             calculated_commission = 0.0
        
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        kpi_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0.0

        debug_details = []
        debug_details.append({'รายการ': '## 1. สรุปยอดขาย (Sale Summary) ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน (Sales Base)', 'ค่า': f"{total_sales:,.2f}"})
        
        debug_details.append({"รายการ": "เป้าหมายการขาย (KPI Target)", "ค่า": f"{sales_target:,.2f}"})
        if sales_target > 0:
            debug_details.append({"รายการ": "KPI Achievement", "ค่า": f"{kpi_percent:,.2f}%"})
        else:
            debug_details.append({"รายการ": "KPI Achievement", "ค่า": "(ไม่มีเป้า KPI)"})

        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({"รายการ": "ยอดขายขั้นต่ำ (เงื่อนไขรับค่าคอม)", "ค่า": f"{min_sales_target:,.2f}"})
        
        if total_sales >= min_sales_target:
            debug_details.append({'รายการ': 'สถานะยอดขายขั้นต่ำ', 'ค่า': '✅ ผ่านเกณฑ์'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': '## 2. การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
            debug_details.append({'รายการ': 'ยอดหักรวม (ค่าดำเนินการ + นายหน้า)', 'ค่า': f"{total_deduction_to_cascade:,.2f}"})
            
            debug_details.append({'รายการ': 'ฐานคำนวณ Normal (หลังหัก)', 'ค่า': f"{commission_base_normal:,.2f}"})
            debug_details.append({'รายการ': '   * คอมฯ Normal (0.7%)', 'ค่า': f"{normal_commission:,.2f}"})
            
            if remaining_deduction > 0:
                debug_details.append({'รายการ': 'ยอดหักคงเหลือ (ไปหักต่อที่ Below Tier)', 'ค่า': f"{remaining_deduction:,.2f}"})
            
            debug_details.append({'รายการ': 'ฐานคำนวณ Below Tier (หลังหัก)', 'ค่า': f"{commission_base_below:,.2f}"})
            debug_details.append({'รายการ': '   * คอมฯ Below Tier (0.3%)', 'ค่า': f"{below_tier_commission:,.2f}"})
            
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ยอดรวมคอมมิชชั่นที่คำนวณได้', 'ค่า': f"{calculated_commission:,.2f}"})
        else:
            debug_details.append({'รายการ': 'สถานะยอดขายขั้นต่ำ', 'ค่า': '❌ ไม่ผ่านเกณฑ์ (ยอดขายไม่ถึงขั้นต่ำ)'})
            debug_details.append({'รายการ': 'รวมคอมมิชชั่น', 'ค่า': '0.00'})
        
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        debug_details.append({'รายการ': '## 3. สรุปยอดจ่ายสุทธิ (Final Payout) ##', 'ค่า': ''})
        if total_incentives > 0: debug_details.append({'รายการ': '(+) รวม Incentive เพิ่มเติม', 'ค่า': f"{total_incentives:,.2f}"})
        if total_additional_deductions > 0: debug_details.append({'รายการ': '(-) หักค่าใช้จ่ายอื่นๆ', 'ค่า': f"{-total_additional_deductions:,.2f}"})
        debug_details.append({'รายการ': 'ยอดคอมมิชชั่นก่อนหักภาษี (Pre-tax Base)', 'ค่า': f"{pre_tax_commission:,.2f}"})
        debug_details.append({'รายการ': f'(-) หัก ณ ที่จ่าย 3%', 'ค่า': f"{-withholding_tax:,.2f}"})
        debug_details.append({'รายการ': '=== ยอดโอนสุทธิ (Net Commission) ===', 'ค่า': f"{net_commission:,.2f}"})

        debug_df = pd.DataFrame(debug_details)
        
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'shipping_to_stock_cost', 'shipping_to_site_cost',
            'excess_stock_shipping', 'excess_site_shipping',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
        
        # 🔥 แก้บัก SO Number ตรงนี้: จัดเรียงก่อนเปลี่ยนชื่อ!
        so_breakdown_df.sort_values(by='so_number', ascending=True, inplace=True)
        so_breakdown_df.reset_index(drop=True, inplace=True)

        so_breakdown_df.rename(columns={
            'so_number': 'SO Number', 
            'sales_service_amount': 'ยอดขาย', 
            'final_cost_amount': 'ต้นทุนรวม', 
            'po_cutting_cost': 'ทุนตัด',
            'po_service_cost': 'ทุนบริการ',
            'shipping_to_stock_cost': 'ทุนเข้าStock',
            'shipping_to_site_cost': 'ทุนเข้าSite',
            'excess_stock_shipping': 'หักส่วนต่างStock',
            'excess_site_shipping': 'หักส่วนต่างSite',
            'profit': 'กำไร', 
            'margin': 'Margin (%)'
        }, inplace=True)
        
        summary_desc = ["ยอดขาย Normal", "ยอดขาย Below Tier", "(-) หัก ค่าดำเนินการ/นายหน้า", "ฐานคอม Normal", "ยอดรวมค่าคอมที่คำนวณได้"]
        summary_val = [total_normal_sales, total_below_sales, total_deduction_to_cascade if total_sales >= min_sales_target else 0, commission_base_normal, calculated_commission]
        
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        
        for k, v in additional_deductions.items():
            summary_desc.append(f"(-) หัก {k}")
            summary_val.append(v)

        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        return {
            'type': 'summary_other', 
            'data': pd.DataFrame({'description': summary_desc, 'value': summary_val}),
            'final_commission': calculated_commission,
            'debug_df': debug_df,
            'so_breakdown_df': so_breakdown_df
        }
    
    else:   
         return {'type': 'error', 'message': f'ไม่พบ Plan ที่ชื่อว่า {plan_name}'}