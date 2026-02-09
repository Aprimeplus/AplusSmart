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

    # ==================================================================================
    # HELPER: ฟังก์ชันสำหรับเตรียมข้อมูลและคำนวณกำไรสุทธิ (Core Logic)
    # ==================================================================================
   
    def prepare_and_calculate_profit(df):
        # 1. Fill NA ให้ครบทุกฟิลด์ที่ต้องใช้
        cols_to_fix = [
            'sales_service_amount', 'final_cost_amount', 'giveaways', 'brokerage_fee', 
            'difference_amount', 'payment_before_vat', 'payment_no_vat', 
            
            # --- [SO Revenue Fields] ---
            'shipping_cost',       # รายได้ค่าขนส่ง (Site)
            'relocation_cost',     # รายได้ค่าย้าย (Stock) จาก Delivery Note
            'so_cutting_rev',      # รายได้ค่าตัด (Revenue)
            'so_service_rev',      # รายได้บริการอื่นๆ (Revenue)
            
            # --- [PO Cost Fields] (แก้ชื่อให้ตรงกับ DB) ---
            'shipping_to_site_cost',   # <--- แก้จาก po_shipping_site_cost
            'shipping_to_stock_cost',  # <--- แก้จาก po_shipping_stock_cost
            'po_cutting_cost',        # PO: ค่าตัด
            'po_service_cost'         # PO: ค่าบริการ
        ]
        
        for col in cols_to_fix:
            if col not in df.columns: 
                df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # 2. จัดการ Multiplier
        if 'cost_multiplier' in df.columns and not df['cost_multiplier'].isnull().all():
             df['cost_multiplier'] = pd.to_numeric(df['cost_multiplier'], errors='coerce').fillna(1.03)
             df['cost_multiplier'] = df['cost_multiplier'].apply(lambda x: 1.03 if x < 1.01 else x)
        else:
             df['cost_multiplier'] = 1.03

        # =========================================================
        # [🔥 MATCHING LOGIC - ปรับปรุงตามตาราง] 
        # =========================================================
        
        # 1. คู่ค่าส่งเข้าไซต์ (Site): จับคู่ชนกัน
        # รายรับ (shipping_cost) vs รายจ่าย (shipping_to_site_cost)
        deduct_site_shipping = (df['shipping_to_site_cost'] - df['shipping_cost']).clip(lower=0)
        
        # 2. คู่ค่าย้ายเข้าสต็อก (Stock): จับคู่ชนกัน
        # รายรับ (relocation_cost) vs รายจ่าย (shipping_to_stock_cost)
        deduct_stock_shipping = (df['shipping_to_stock_cost'] - df['relocation_cost']).clip(lower=0)
        
        # 3. คู่ค่าตัด/เจาะ: จับคู่ชนกัน
        deduct_cutting = (df['po_cutting_cost'] - df['so_cutting_rev']).clip(lower=0)
        
        # 4. คู่ค่าบริการอื่นๆ: จับคู่ชนกัน
        deduct_service = (df['po_service_cost'] - df['so_service_rev']).clip(lower=0)

        # --- คำนวณต้นทุนสินค้าหลัก (Main Cost Calculation) ---
        # ต้องหักต้นทุนพิเศษออกจาก "ต้นทุนรวมใน PO" ก่อนนำไปคูณ 1.03
        adjusted_po_cost = df['final_cost_amount'] - df['po_cutting_cost'] - df['po_service_cost']
        adjusted_po_cost = adjusted_po_cost.clip(lower=0)

        # คำนวณต้นทุนรวม Margin (Cost * 1.03)
        main_cost_calculated = (adjusted_po_cost * df['cost_multiplier'])
        
        # --- คำนวณกำไรสุทธิ (Final Profit) ---
        # Profit = ยอดขายสินค้า - ต้นทุนสินค้า(x1.03) + ผลต่างโอน - (ส่วนต่างค่าขนส่ง/ค่าบริการต่างๆ)
        df['profit'] = (df['sales_service_amount'] - main_cost_calculated) \
                       + df['difference_amount'] \
                       - deduct_site_shipping \
                       - deduct_stock_shipping \
                       - deduct_cutting \
                       - deduct_service
        
        # คำนวณ Margin %
        df['margin'] = (df['profit'] / df['sales_service_amount'].replace(0, np.nan)) * 100
        df['margin'] = df['margin'].fillna(0)

        # [🔥 เพิ่ม] บันทึกค่า Excess ลง DataFrame เพื่อให้ Plan อื่นเรียกใช้ทำ Report ได้
        df['excess_site_shipping'] = deduct_site_shipping
        df['excess_stock_shipping'] = deduct_stock_shipping

        # --- Debug Print (แสดงผลเมื่อมีการหักลบยอด) ---
        check_sum = deduct_site_shipping.sum() + deduct_stock_shipping.sum() + deduct_cutting.sum() + deduct_service.sum()
        if check_sum > 0:
            print("\n" + "="*60)
            print(f"🤖 SYSTEM MATCHING REPORT: รายการที่มีการหักลบต้นทุนส่วนเกิน")
            print("="*60)
            temp_df = df[ (deduct_site_shipping > 0) | (deduct_stock_shipping > 0) | (deduct_cutting > 0) | (deduct_service > 0)]
            
            for index, row in temp_df.iterrows():
                print(f"📄 SO: {row.get('so_number', 'N/A')}")
                if deduct_site_shipping[index] > 0:
                    print(f"   🚚 ค่ารถ Site (Excess): จ่าย {row['shipping_to_site_cost']:,.2f} - รับ {row['shipping_cost']:,.2f} = หักกำไร {deduct_site_shipping[index]:,.2f}")
                if deduct_stock_shipping[index] > 0:
                    print(f"   🏭 ค่าย้าย Stock (Excess): จ่าย {row['shipping_to_stock_cost']:,.2f} - รับ {row['relocation_cost']:,.2f} = หักกำไร {deduct_stock_shipping[index]:,.2f}")
                if deduct_cutting[index] > 0:
                    print(f"   ✂️ ค่าตัด (Excess): จ่าย {row['po_cutting_cost']:,.2f} - รับ {row['so_cutting_rev']:,.2f} = หักกำไร {deduct_cutting[index]:,.2f}")
                print("-" * 60)
            print("\n")

        return df

    # ==================================================================================
    # PLAN A
    # ==================================================================================
    if plan_name == 'Plan A':
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper Function ด้านบน)
        # [🔥 สำคัญ] ขั้นตอนนี้กำไรจะถูกหักลบด้วยค่าขนส่งส่วนเกินแล้ว
        comm_df = prepare_and_calculate_profit(comm_df)
        
        # 2. คำนวณยอดขายรวม
        total_sales = comm_df['sales_service_amount'].sum()
        
        initial_commission = 0.0
        calculated_commission = 0.0
        
        # ค่าดำเนินการ (Default 25,000)
        OPERATING_FEE = 25000.00 if operating_fee is None else operating_fee
        
        # ตรวจสอบเงื่อนไขยอดขายขั้นต่ำ
        if total_sales >= min_sales_target:
            NORMAL_RATE = 0.35   # 35% สำหรับ Normal Margin
            BELOW_T_RATE = 0.175 # 17.5% สำหรับ Below Margin

            # แยกกลุ่มข้อมูลตาม Margin
            normal_df = comm_df[comm_df['margin'] >= 10]
            below_df = comm_df[comm_df['margin'] < 10]

            # คำนวณยอดขายแต่ละกลุ่ม (เพื่อแสดงผล)
            val_normal_sales = normal_df['sales_service_amount'].sum()
            val_below_sales = below_df['sales_service_amount'].sum()
            
            # คำนวณค่าคอมมิชชั่นจาก Profit (ตามสูตรที่ให้มา)
            commission_normal = normal_df['profit'].sum() * NORMAL_RATE
            commission_below = below_df['profit'].sum() * BELOW_T_RATE
            initial_commission = commission_normal + commission_below
            
            # รวมค่านายหน้าทั้งหมด (Brokerage Fee)
            brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
            total_brokerage_fee = brokerage.sum()
            
            # คำนวณยอดสุทธิ: (ค่าคอมรวม - ค่านายหน้า - ค่าดำเนินการ)
            # ห้ามติดลบ (max 0)
            calculated_commission = max(0, initial_commission - total_brokerage_fee - OPERATING_FEE)

            # Assign ค่ากลับไปใน DataFrame เพื่อทำตารางแจกแจงรายละเอียด (Breakdown)
            conditions = [comm_df['margin'] >= 10, comm_df['margin'] < 10]
            choices_commission = [comm_df['profit'] * NORMAL_RATE, comm_df['profit'] * BELOW_T_RATE]
            comm_df['commission_amount'] = np.select(conditions, choices_commission, default=0)
        
        else:
            # ยอดขายไม่ถึงเป้าขั้นต่ำ -> ไม่ได้ค่าคอม
            val_normal_sales = 0.0
            val_below_sales = 0.0
            comm_df['commission_amount'] = 0.0
            initial_commission = 0.0
            calculated_commission = 0.0

        # --- 3. สรุปยอดเงินสุดท้าย (Financials) ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 4. สร้างตาราง Breakdown ---
        # [🔥 เพิ่ม] คอลัมน์ Shipping เพื่อให้เห็นรายละเอียด
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'shipping_to_stock_cost', 'shipping_to_site_cost', # เพิ่มตรงนี้
            'profit', 'margin', 'commission_amount', 'cost_multiplier'
        ]].copy()
        
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
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
        
        # --- 5. สร้างตาราง Summary ---
        summary_desc = [
            "ยอดขาย Normal", 
            "ยอดขาย Below Tier", 
            "รวมยอดคอมมิชชั่นจาก Profit",
            "(-) หัก ค่านายหน้า (Brokerage)",
            "(-) หัก ค่าดำเนินการ (Operating Fee)",
            "ยอดรวมค่าคอมที่คำนวณได้"
        ]
        summary_val = [
            val_normal_sales, 
            val_below_sales, 
            initial_commission,
            total_brokerage_fee,
            OPERATING_FEE,
            calculated_commission
        ]
        
        # เพิ่ม Incentive/Deduction ใน Summary
        for k, v in incentives.items(): 
            summary_desc.append(f"(+) {k}")
            summary_val.append(v)
            
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross)")
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
            'so_breakdown_df': so_breakdown_df,
            'debug_df': [] 
        }

    # ==================================================================================
    # PLAN B
    # ==================================================================================
    elif plan_name == 'Plan B':
        if comm_df.empty: 
            return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        # 1. เตรียมข้อมูล & คำนวณ Profit รายบรรทัด (เรียกใช้ Helper Function)
        # [🔥 สำคัญ] ขั้นตอนนี้กำไรสุทธิ (Profit) จะถูกคำนวณโดยหัก "ส่วนต่างค่าขนส่ง (Excess)" เรียบร้อยแล้ว
        # ดังนั้น เมื่อเรารวมยอด Profit ในขั้นตอนต่อไป เราจะได้กำไรที่ถูกต้องตามจริง
        comm_df = prepare_and_calculate_profit(comm_df)

        if 'po_number' not in comm_df.columns: 
            comm_df['po_number'] = comm_df['so_number']
        
        # 2. รวมยอดตามใบสั่งซื้อ (Group by PO)
        # เราต้องรวม Profit ที่คำนวณมาแล้ว เพื่อให้ได้กำไรสุทธิจริงของ PO นั้นๆ (ซึ่งอาจประกอบด้วยหลาย SO)
        agg_rules = {
            'sales_service_amount': 'sum', 
            'giveaways': 'sum', 
            'brokerage_fee': 'sum', 
            'difference_amount': 'sum', 
            'final_cost_amount': 'sum', 
            'cost_multiplier': 'first', 
            'so_number': lambda x: ', '.join(sorted(set(str(v) for v in x))),
            'profit': 'sum', # รวม Profit (ที่หักลบทุกอย่างมาแล้ว)
            
            # รวมยอดพิเศษเพื่อแสดงใน Breakdown (Option)
            'po_cutting_cost': 'sum', 
            'po_service_cost': 'sum',
            'shipping_to_stock_cost': 'sum', # รวมต้นทุนค่าย้าย (Stock)
            'shipping_to_site_cost': 'sum',  # รวมต้นทุนค่ารถ (Site)
            'relocation_cost': 'sum',        # รวมรายรับค่าย้าย
            'shipping_cost': 'sum',          # รวมรายรับค่ารถ
            
            # เก็บค่า Excess ไว้ดูเล่นใน Breakdown (ถ้าต้องการ)
            'excess_site_shipping': 'sum',
            'excess_stock_shipping': 'sum'
        }
        
        # ต้องระวัง Error กรณีคอลัมน์ไม่มีจริง (ถึงแม้ prepare_and_calculate_profit จะสร้างให้แล้วก็ตาม)
        # กรองเอาเฉพาะคอลัมน์ที่มีอยู่จริงใน comm_df มาใส่ใน agg_rules
        valid_agg_rules = {k: v for k, v in agg_rules.items() if k in comm_df.columns}
        
        po_grouped_df = comm_df.groupby('po_number').agg(valid_agg_rules).reset_index()

        # 3. คำนวณ Margin ใหม่จากยอดรวมของ PO
        # Margin = (กำไรสุทธิรวม / ยอดขายรวม) * 100
        po_grouped_df['margin'] = (po_grouped_df['profit'] / po_grouped_df['sales_service_amount'].replace(0, np.nan)) * 100    
        po_grouped_df['margin'] = po_grouped_df['margin'].fillna(0)

        # 4. แบ่ง Tier ตาม Margin ของ PO
        # Tier 1 (Normal): Margin >= 10%
        standard_margin_df = po_grouped_df[po_grouped_df['margin'] >= 10]
        
        # Tier 2 (Below 1): Margin 7.99% - 9.99%
        below_tier1_df = po_grouped_df[(po_grouped_df['margin'] >= 7.99) & (po_grouped_df['margin'] < 10)]
        
        # Tier 3 (Below 2): Margin < 7.99%
        below_tier2_df = po_grouped_df[po_grouped_df['margin'] < 7.99]
        
        # รวมยอดขายแต่ละกลุ่ม
        total_standard_sales = standard_margin_df['sales_service_amount'].sum()
        total_below_tier1_sales = below_tier1_df['sales_service_amount'].sum()
        total_below_tier2_sales = below_tier2_df['sales_service_amount'].sum()
        total_monthly_sales = total_standard_sales + total_below_tier1_sales + total_below_tier2_sales
        
        # คำนวณ Below Tier Commission (คิดจากยอดขาย)
        # Tier 2: 0.63%
        commission_below_t1 = total_below_tier1_sales * 0.0063
        # Tier 3: 0.50%
        commission_below_t2 = total_below_tier2_sales * 0.0050
        
        below_tier_commission = commission_below_t1 + commission_below_t2

        # เตรียมฐานคำนวณ Normal Tier (Waterfall Logic)
        total_brokerage_fee = po_grouped_df['brokerage_fee'].sum()
        
        # ค่าดำเนินการ (Default 100,000 สำหรับ Plan B)
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        
        # ฐานคอม = ยอดขาย Normal - ค่านายหน้า - ค่าดำเนินการ
        commission_base = max(0, total_standard_sales - total_brokerage_fee - OPERATING_FEE)
        
        t1, t2, t3, tier_commission, calculated_commission = 0, 0, 0, 0, 0

        # ตรวจสอบเป้าขายรวม (Sales Target)
        if total_monthly_sales >= min_sales_target:
            remaining_base = commission_base
            
            # Step 1: 0 - 1,000,000 (1.25%)
            amount_in_t1 = min(remaining_base, 1000000)
            t1 = amount_in_t1 * 0.0125
            remaining_base -= amount_in_t1
            
            # Step 2: 1,000,001 - 2,000,000 (1.75%)
            if remaining_base > 0:
                amount_in_t2 = min(remaining_base, 1000000)
                t2 = amount_in_t2 * 0.0175
                remaining_base -= amount_in_t2
            
            # Step 3: > 2,000,000 (2.25%)
            if remaining_base > 0:
                t3 = remaining_base * 0.0225
            
            tier_commission = t1 + t2 + t3
            calculated_commission = tier_commission + below_tier_commission
        else:
            # ไม่ผ่านเป้าขั้นต่ำ -> ได้ 0
            calculated_commission = 0.0

        commission_base_normal = commission_base # For reporting

        # 5. สรุปยอดเงินสุดท้าย (Financials)
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # 6. สร้าง Report
        summary_desc = [
            "ยอดขายรวม", 
            "ยอดขาย Normal (Margin >= 10%)", 
            "(-) หัก ค่าดำเนินการ (Operating Fee)", 
            "ฐานคำนวณคอมฯ Normal", 
            "ยอดรวมค่าคอมที่คำนวณได้"
        ]
        summary_val = [
            total_monthly_sales, 
            total_standard_sales, 
            OPERATING_FEE, 
            commission_base_normal, 
            calculated_commission
        ]
        
        # เพิ่ม Incentive/Deduction
        for k, v in incentives.items(): summary_desc.append(f"(+) {k}"); summary_val.append(v)
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross)"); summary_val.append(gross_commission)
        for k, v in additional_deductions.items(): summary_desc.append(f"(-) {k}"); summary_val.append(v)
        
        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])

        # สร้างตาราง Breakdown (แสดงรายละเอียดตาม PO) - ปรับปรุงคอลัมน์ให้ครบถ้วน
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
            'debug_df': []
        }

    # ==================================================================================
    # PLAN C (Logic คล้าย Plan B)
    # ==================================================================================
    elif plan_name == 'Plan C':
        if comm_df.empty: 
            return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper Function ที่แก้ไขแล้ว)
        # [🔥 สำคัญ] ขั้นตอนนี้กำไรสุทธิ (Profit) จะถูกคำนวณโดยหัก "ส่วนต่างค่าขนส่ง (Excess)" เรียบร้อยแล้ว
        # และ EXP-0174 (ค่าย้าย) จะถูกจัดการใน Logic ของ Matching Stock Shipping
        comm_df = prepare_and_calculate_profit(comm_df)

        # --- 2. แบ่งกลุ่มตาม Margin และรวมยอดขาย ---
        # Tier 1 (Normal): Margin >= 10%
        tier1_df = comm_df[comm_df['margin'] >= 10]
        tier1_sales = tier1_df['sales_service_amount'].sum()
        
        # Tier 2: Margin 7.99% - 9.99%
        tier2_df = comm_df[(comm_df['margin'] >= 7.99) & (comm_df['margin'] < 10)]
        tier2_sales = tier2_df['sales_service_amount'].sum()
        
        # Tier 3: Margin < 7.99%
        tier3_df = comm_df[comm_df['margin'] < 7.99]
        tier3_sales = tier3_df['sales_service_amount'].sum()
        
        total_sales = tier1_sales + tier2_sales + tier3_sales
        
        # --- 3. คำนวณค่าคอมมิชชั่น ---
        # ค่าดำเนินการ (Default 100,000 สำหรับ Plan C)
        OPERATING_FEE = 100000.00 if operating_fee is None else operating_fee
        
        # ค่านายหน้า
        total_brokerage = comm_df['brokerage_fee'].sum()
        
        # ยอดหักรวม (Deduction Pool)
        total_deduction = OPERATING_FEE + total_brokerage

        comm_t1, comm_t2, comm_t3, calculated_commission = 0, 0, 0, 0
        base_t1, base_t2, base_t3 = 0, 0, 0 # ฐานคำนวณแต่ละ Tier
        
        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        # ตรวจสอบเป้าขั้นต่ำ
        if total_sales >= min_sales_target:
            # Waterfall Logic: หักค่าใช้จ่ายจาก Tier สูงไล่ลงมา Tier ต่ำ
            
            # Step 1: หักจาก Tier 1 ก่อน
            base_t1 = max(0, tier1_sales - total_deduction)
            remaining_deduction = max(0, total_deduction - tier1_sales) # ยอดหักที่เหลือไปต่อ Tier 2
            comm_t1 = round(base_t1 * 0.01, 2) # Rate Tier 1 = 1.0%
            
            # Step 2: หักจาก Tier 2 (ถ้ายอดหักยังเหลือ)
            base_t2 = max(0, tier2_sales - remaining_deduction)
            remaining_deduction = max(0, remaining_deduction - tier2_sales) # ยอดหักที่เหลือไปต่อ Tier 3
            comm_t2 = round(base_t2 * 0.0063, 2) # Rate Tier 2 = 0.63%
            
            # Step 3: หักจาก Tier 3 (ถ้ายอดหักยังเหลือ)
            base_t3 = max(0, tier3_sales - remaining_deduction)
            comm_t3 = round(base_t3 * 0.005, 2) # Rate Tier 3 = 0.5%
            
            calculated_commission = comm_t1 + comm_t2 + comm_t3
        else:
            # ไม่ผ่านเป้าขั้นต่ำ
            calculated_commission = 0.0

        # --- 4. สรุปยอดเงิน (Financials) ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 5. สร้าง Debug Report (สำหรับตรวจสอบที่มาตัวเลข) ---
        debug_details = []
        debug_details.append({'รายการ': '## 1. สรุปยอดขาย (Sale Summary) ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'ยอดขายรวม (Total Sales)', 'ค่า': total_sales})
        debug_details.append({'รายการ': '   - Tier 1 (Margin >= 10%)', 'ค่า': tier1_sales})
        debug_details.append({'รายการ': '   - Tier 2 (Margin 7.99-9.99%)', 'ค่า': tier2_sales})
        debug_details.append({'รายการ': '   - Tier 3 (Margin < 7.99%)', 'ค่า': tier3_sales})
        debug_details.append({'รายการ': 'KPI Achievement', 'ค่า': f"{hit_target_percent:.2f}% ({hit_target_status})"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## 2. การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_sales >= min_sales_target:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'รวมยอดหัก (ค่าดำเนินการ+นายหน้า)', 'ค่า': total_deduction})
            debug_details.append({'รายการ': 'ฐานคำนวณ Tier 1 (หลังหัก)', 'ค่า': base_t1})
            debug_details.append({'รายการ': '   * คอมฯ T1 (1.0%)', 'ค่า': comm_t1})
            debug_details.append({'รายการ': 'ฐานคำนวณ Tier 2 (หลังหัก)', 'ค่า': base_t2})
            debug_details.append({'รายการ': '   * คอมฯ T2 (0.63%)', 'ค่า': comm_t2})
            debug_details.append({'รายการ': 'ฐานคำนวณ Tier 3 (หลังหัก)', 'ค่า': base_t3})
            debug_details.append({'รายการ': '   * คอมฯ T3 (0.50%)', 'ค่า': comm_t3})
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

        # --- 6. สร้างตาราง Summary & Breakdown ---
        summary_desc = ["ยอดขาย Tier 1", "ยอดขาย Tier 2", "ยอดขาย Tier 3", "ยอดรวมคอมที่คำนวณได้"]
        summary_val = [tier1_sales, tier2_sales, tier3_sales, calculated_commission]
        
        # เพิ่มส่วนเสริมใน Summary
        summary_desc.append("ยอดคอมมิชชั่นขั้นต้น (Gross Commission)"); summary_val.append(gross_commission)
        
        for k, v in additional_deductions.items():
            summary_desc.append(f"(-) หัก {k}")
            summary_val.append(v)

        summary_desc.extend(["ยอดคอมมิชชั่นก่อนหักภาษี", "(-) หัก ณ ที่จ่าย 3%", "ยอดสรุปคอมหลังหัก ณ ที่จ่าย"])
        summary_val.extend([pre_tax_commission, withholding_tax, net_commission])
        
        # สร้าง Breakdown Table - เพิ่มคอลัมน์ Shipping Breakdown
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'shipping_to_stock_cost', 'shipping_to_site_cost',
            'excess_stock_shipping', 'excess_site_shipping',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        # กำหนด Status ให้แต่ละ SO ตาม Margin
        def assign_c_tier_status(margin):
            if margin >= 10: return 'Tier 1 (>=10%)'
            if margin >= 7.99: return 'Tier 2 (7.99-9.99%)'
            return 'Tier 3 (<7.99%)'
            
        so_breakdown_df['Status'] = so_breakdown_df['margin'].apply(assign_c_tier_status)
        
        so_breakdown_df.rename(columns={
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
    # PLAN D (Logic คล้าย Plan B)
    # ==================================================================================
    # ==================================================================================
    # PLAN D (Logic คล้าย Plan C แต่ Fee/Rate ต่างกัน)
    # ==================================================================================
    elif plan_name == 'Plan D':
        if comm_df.empty: 
            return {'type': 'summary_other', 'data': pd.DataFrame({'description': ["ยอดรวม"], 'value': [0.0]})}
        
        # 1. เตรียมข้อมูล & คำนวณ Profit (ใช้ Helper Function ที่แก้ไขแล้ว)
        # [🔥 สำคัญ] ขั้นตอนนี้กำไรสุทธิ (Profit) จะถูกคำนวณโดยหัก "ส่วนต่างค่าขนส่ง (Excess)" เรียบร้อยแล้ว
        # และ EXP-0174 (ค่าย้าย) จะถูกจัดการใน Logic ของ Matching Stock Shipping
        comm_df = prepare_and_calculate_profit(comm_df)
        
        # --- 2. แบ่งกลุ่มและรวมยอดขาย ---
        normal_margin_df = comm_df[comm_df['margin'] >= 10]
        below_margin_df = comm_df[comm_df['margin'] < 10]
        
        total_normal_sales = normal_margin_df['sales_service_amount'].sum()
        total_below_sales = below_margin_df['sales_service_amount'].sum()
        total_sales = total_normal_sales + total_below_sales
        
        # --- 3. คำนวณค่าคอมมิชชั่น ---
        normal_commission = 0.0
        below_tier_commission = 0.0
        calculated_commission = 0.0
        
        # ค่าดำเนินการสำหรับ Plan D (Default 750,000)
        OPERATING_FEE = 750000.00 if operating_fee is None else operating_fee
        
        commission_base_normal = 0.0
        commission_base_below = 0.0

        hit_target_percent = (total_sales / sales_target * 100) if sales_target > 0 else 0
        hit_target_status = "TARGET" if hit_target_percent >= 100 else "UNDER TARGET"

        # ตรวจสอบเป้าขั้นต่ำ
        if total_sales >= min_sales_target:
            # รวมค่านายหน้า
            brokerage = pd.to_numeric(comm_df['brokerage_fee'], errors='coerce').fillna(0)
            total_brokerage_fee = brokerage.sum()
            
            # ยอดหักรวม (ค่าดำเนินการ + ค่านายหน้า)
            total_deduction_to_cascade = OPERATING_FEE + total_brokerage_fee

            # Step 1: หักจากยอดขาย Normal ก่อน
            # ฐานคอม Normal = ยอดขาย Normal - ยอดหักรวม
            commission_base_normal = max(0, total_normal_sales - total_deduction_to_cascade)
            normal_commission = commission_base_normal * 0.007 # Rate Normal = 0.7%

            # Step 2: ถ้ายอดขาย Normal ไม่พอหัก ให้เอายอดที่เหลือไปหักจาก Below Tier
            remaining_deduction = max(0, total_deduction_to_cascade - total_normal_sales)
            
            # ฐานคอม Below Tier = ยอดขาย Below - ยอดหักที่เหลือ
            commission_base_below = max(0, total_below_sales - remaining_deduction)
            below_tier_commission = commission_base_below * 0.003 # Rate Below = 0.3%
            
            calculated_commission = normal_commission + below_tier_commission
        else:
             # ไม่ผ่านเป้าขั้นต่ำ
             calculated_commission = 0.0
        
        # --- 4. สรุปยอดเงิน (Financials) ---
        gross_commission = calculated_commission + total_incentives
        pre_tax_commission = gross_commission - total_additional_deductions
        withholding_tax = pre_tax_commission * 0.03
        net_commission = pre_tax_commission - withholding_tax

        # --- 5. สร้าง Debug Report ---
        debug_details = []
        debug_details.append({'รายการ': '## 1. สรุปยอดขาย (Sale Summary) ##', 'ค่า': ''})
        debug_details.append({'รายการ': 'Commission Plan', 'ค่า': plan_name})
        debug_details.append({'รายการ': 'สรุปยอดขายประจำเดือน (Sales Base)', 'ค่า': total_sales})
        debug_details.append({'รายการ': 'KPI Achievement', 'ค่า': f"{hit_target_percent:.2f}% ({hit_target_status})"})
        debug_details.append({'รายการ': '---', 'ค่า': ''})
        
        debug_details.append({'รายการ': '## 2. การคำนวณคอมมิชชั่น (Commission Calculation) ##', 'ค่า': ''})
        if total_sales >= min_sales_target:
            debug_details.append({'รายการ': f'เงื่อนไขยอดขายขั้นต่ำ ({min_sales_target:,.0f})', 'ค่า': 'ผ่าน ✅'})
            debug_details.append({'รายการ': '---', 'ค่า': ''})
            debug_details.append({'รายการ': 'ยอดหักรวม (ค่าดำเนินการ + นายหน้า)', 'ค่า': total_deduction_to_cascade})
            
            debug_details.append({'รายการ': 'ฐานคำนวณ Normal (หลังหัก)', 'ค่า': commission_base_normal})
            debug_details.append({'รายการ': '   * คอมฯ Normal (0.7%)', 'ค่า': normal_commission})
            
            if remaining_deduction > 0:
                debug_details.append({'รายการ': 'ยอดหักคงเหลือ (ไปหักต่อที่ Below Tier)', 'ค่า': remaining_deduction})
            
            debug_details.append({'รายการ': 'ฐานคำนวณ Below Tier (หลังหัก)', 'ค่า': commission_base_below})
            debug_details.append({'รายการ': '   * คอมฯ Below Tier (0.3%)', 'ค่า': below_tier_commission})
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
        
        # --- 6. สร้างตาราง Breakdown ---
        # [🔥 แก้ไข] เพิ่มคอลัมน์ Shipping Breakdown
        so_breakdown_df = comm_df[[
            'so_number', 'sales_service_amount', 'final_cost_amount', 
            'po_cutting_cost', 'po_service_cost',
            'shipping_to_stock_cost', 'shipping_to_site_cost',
            'excess_stock_shipping', 'excess_site_shipping',
            'profit', 'margin', 'cost_multiplier'
        ]].copy()
        
        so_breakdown_df['Status'] = np.where(so_breakdown_df['margin'] >= 10, 'Normal (>=10%)', 'Below Tier (<10%)')
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
    
    # กรณีไม่ใช่ Plan A-D หรือ Plan อื่นๆ ที่ยังไม่รองรับ
    else:   
         return {'type': 'error', 'message': f'ไม่พบ Plan ที่ชื่อว่า {plan_name}'}