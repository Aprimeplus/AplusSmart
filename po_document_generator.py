import os
import sys
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox
import utils
# ReportLab Core
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, 
    Spacer, Table, TableStyle, Image, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4

# Font & Metrics
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def register_thai_fonts():
    """ ลงทะเบียนฟอนต์ไทยและสร้าง Mapping เพื่อให้ใช้งานตัวหนา <b> ได้ถูกต้อง """
    try:
        # ระบุ Path ให้ถูกต้อง (รองรับทั้งตอน Dev และตอน Build เป็น .exe)
        font_path = resource_path("resources/THSarabunNew.ttf")
        font_bold_path = resource_path("resources/THSarabunNew Bold.ttf")
        
        # ตรวจสอบว่าไฟล์มีอยู่จริงก่อนลงทะเบียน
        if not os.path.exists(font_path): font_path = "THSarabunNew.ttf"
        if not os.path.exists(font_bold_path): font_bold_path = "THSarabunNew Bold.ttf"

        pdfmetrics.registerFont(TTFont('THSarabunNew', font_path))
        pdfmetrics.registerFont(TTFont('THSarabunNew-Bold', font_bold_path))
        
        # สำคัญ: ต้อง Map ให้ระบบรู้จักความสัมพันธ์ของฟอนต์
        registerFontFamily('THSarabunNew', normal='THSarabunNew', bold='THSarabunNew-Bold')
    except Exception as e:
        print(f"Font Registration Warning: {e}")

def _build_left_column(header_data, styles, P, PB, format_num, width):
    """
    สร้างคอลัมน์ซ้าย (SELL AUDITOR)
    แก้ไขล่าสุด: 
    1. ตัดแถว 'เรื่อง (Subject)' ออก
    2. Logic Checkbox ยอดเงิน: ถ้าไม่มียอดเงิน ให้แสดง [ ] (ว่าง)
    3. Logic Checkbox ธนาคาร: ติ๊กถูกตาม Payment Method ที่เลือก
    """
    story = []
    
    def safe_add_style(styles, style):
        if style.name not in styles.byName: styles.add(style)

    safe_add_style(styles, ParagraphStyle(name='Small_TH', fontName='THSarabunNew', fontSize=10, leading=12))
    safe_add_style(styles, ParagraphStyle(name='Tiny_TH', fontName='THSarabunNew', fontSize=9, leading=11))
    safe_add_style(styles, ParagraphStyle(name='Small_Center_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=1))
    safe_add_style(styles, ParagraphStyle(name='Small_Right_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=2))

    def PS(text, style='Small_TH'): return Paragraph(str(text) if text is not None else '', styles[style])
    def PSafe(text, style='Small_TH', max_length=100):
        text_str = str(text) if text is not None else ''
        if len(text_str) > max_length: text_str = text_str[:max_length-3] + "..."
        return Paragraph(text_str, styles[style])
    
    def fmt_date(d):
        if not d: return ""
        try:
            if isinstance(d, str): 
                d_obj = datetime.strptime(d[:10], "%Y-%m-%d")
                return d_obj.strftime("%d/%m/%Y")
            elif hasattr(d, 'strftime'):
                return d.strftime("%d/%m/%Y")
            return ""
        except:
            return ""

    c1, c3, c4 = 2.2*cm, 1.0*cm, 2.0*cm 
    c2 = width - (c1 + c3 + c4)

    # --- ส่วน Header ---
    combined_header_data = [
        [PB('ขาย', 'Small_TH'), PS('SELL AUDITOR', 'Small_Center_TH'), PB('ผู้ตรวจ..............', 'Small_TH'), None],
        [PB('SO NUMBER', 'Small_TH'), PSafe(header_data.get('so_number', ''), 'Small_TH'), PB('แผนก', 'Small_TH'), PS(header_data.get('department', ''))],
        [PB('Sale Name', 'Small_TH'), PSafe(header_data.get('sale_name', ''), 'Small_TH'), PB('วันที่', 'Small_TH'), PS(str(header_data.get('bill_date', '')))],
        [PB('Customer Name', 'Small_TH'), PSafe(header_data.get('customer_name', ''), 'Small_TH', 80), None, None]
    ]
    
    header_table = Table(combined_header_data, colWidths=[c1, c2, c3, c4]) 
    header_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('SPAN', (2,0), (3,0)), 
        ('SPAN', (1,3), (-1,3)), 
        ('BACKGROUND', (0,1), (0,3), colors.lemonchiffon), 
        ('BACKGROUND', (2,1), (2,2), colors.lemonchiffon)
    ]))
    story.append(header_table)

    title_style = TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,-1), colors.lemonchiffon), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')])
    story.append(Table([[PB('SELLING RECORD', style='Small_Center_TH')]], colWidths=[width], rowHeights=[0.5*cm], style=title_style))

    # ==============================================================================
    # [Logic] คำนวณยอดเงินและ Checkbox
    # ==============================================================================
    box_checked = "[ / ]"
    box_unchecked = "[   ]"

    # 1. ยอดขายสินค้า
    sales_before_vat = utils.convert_to_float(header_data.get('sales_service_amount', 0))
    sales_vat_opt = str(header_data.get('sales_service_vat_option', 'CASH')).upper()
    is_sales_vat = (sales_vat_opt == 'VAT')
    sales_vat = sales_before_vat * 0.07 if is_sales_vat else 0.0
    
    if sales_before_vat > 0:
        chk_sales_yes = box_checked if is_sales_vat else box_unchecked
        chk_sales_no  = box_unchecked if is_sales_vat else box_checked
    else:
        chk_sales_yes = box_unchecked
        chk_sales_no  = box_unchecked

    # 2. ค่าตัดเหล็ก
    cutting_fee = utils.convert_to_float(header_data.get('cutting_drilling_fee', 0))
    cutting_opt = str(header_data.get('cutting_drilling_fee_vat_option', 'CASH')).upper()
    is_cut_vat = (cutting_opt == 'VAT')
    
    if cutting_fee > 0:
        chk_cut_yes = box_checked if is_cut_vat else box_unchecked
        chk_cut_no  = box_unchecked if is_cut_vat else box_checked
    else:
        chk_cut_yes = box_unchecked
        chk_cut_no  = box_unchecked

    # 3. ค่าจัดส่ง
    shipping_fee = utils.convert_to_float(header_data.get('shipping_cost', 0))
    shipping_opt = str(header_data.get('shipping_vat_option', 'CASH')).upper()
    is_ship_vat = (shipping_opt == 'VAT')
    shipping_vat = shipping_fee * 0.07 if is_ship_vat else 0.0
    
    # 4. อื่นๆ
    wht_3 = utils.convert_to_float(header_data.get('wht_3_percent', 0))
    transfer_fee = utils.convert_to_float(header_data.get('transfer_fee', 0))
    coupons = utils.convert_to_float(header_data.get('coupons', 0))
    marketing_fee = utils.convert_to_float(header_data.get('marketing_fee', 0))
    grand_total = utils.convert_to_float(header_data.get('so_grand_total', 0))

    selling_data = [
        [PB('ยอดขายสินค้าก่อน VAT', 'Small_TH'), PS(format_num(sales_before_vat), 'Small_Right_TH'), PS(f'{chk_sales_yes} Vat 7%', 'Tiny_TH'), PS(f'{chk_sales_no} ไม่เอาVat', 'Tiny_TH')],
        [PB('ค่าตัดเหล็ก', 'Small_TH'), PS(format_num(cutting_fee), 'Small_Right_TH'), PS(f'{chk_cut_yes} Vat 7%', 'Tiny_TH'), PS(f'{chk_cut_no} ไม่เอาVat', 'Tiny_TH')],
        [PB('Vat 7% ค่าสินค้า', 'Small_TH'), PS(format_num(sales_vat), 'Small_Right_TH'), PB('ภาษีถูกหัก ณ ที่จ่าย', 'Small_TH'), PS(format_num(wht_3), 'Small_Right_TH')],
        [PB('ค่าจัดส่งก่อน Vat', 'Small_TH'), PS(format_num(shipping_fee), 'Small_Right_TH'), PB('ค่าธรรมเนียมโอน', 'Small_TH'), PS(format_num(transfer_fee), 'Small_Right_TH')],
        [PB('Vat 7% ค่าจัดส่ง', 'Small_TH'), PS(format_num(shipping_vat), 'Small_Right_TH'), PB('ส่วนลด', 'Small_TH'), PS(format_num(coupons), 'Small_Right_TH')],
        [PB('ยอดขายรวมทั้งสิ้น', 'Small_TH'), PS(format_num(grand_total), 'Small_Right_TH'), PB('ค่าการตลาด', 'Small_TH'), PS(format_num(marketing_fee), 'Small_Right_TH')],
    ]
    selling_table = Table(selling_data, colWidths=[width*0.29, width*0.16, width*0.29, width*0.26], rowHeights=[0.5*cm]*6)
    selling_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3)]))
    story.append(selling_table)
    
    story.append(Table([[PB('ยอดชำระค่าสินค้า/บริการ', style='Small_Center_TH')]], colWidths=[width], rowHeights=[0.5*cm], style=title_style))
    
    # ==============================================================================
    # [Logic ใหม่] Checkbox ธนาคาร: เช็คจาก Payment Method ที่เลือก
    # ==============================================================================
    p1_method = str(header_data.get('payment1_method', ''))
    p2_method = str(header_data.get('payment2_method', ''))
    all_methods = f"{p1_method} {p2_method}" # รวมข้อความเพื่อเช็ครวดเดียว

    chk_kbank = box_unchecked
    chk_ttb_sav = box_unchecked
    chk_ttb_cur = box_unchecked
    chk_comm = box_unchecked

    if 'KBANK' in all_methods: chk_kbank = box_checked
    if 'ออมทรัพย์' in all_methods: chk_ttb_sav = box_checked
    if 'กระแส' in all_methods: chk_ttb_cur = box_checked
    if 'กรรมการ' in all_methods: chk_comm = box_checked

    bank_data = [[PB('บัญชีที่ลูกค้าโอน', 'Small_TH'), PS(f'{chk_kbank} ธ.กสิกรไทย', 'Small_TH'), PS(f'{chk_ttb_sav} ธ.ทหารไทย(ออมทรัพย์)', 'Small_TH')], 
                 [None, PS(f'{chk_ttb_cur} ธ.ทหารไทย(กระแสรายวัน)', 'Small_TH'), PS(f'{chk_comm} บัญชีกรรมการ', 'Small_TH')]]
    bank_table = Table(bank_data, colWidths=[width*0.29, width*0.35, width*0.36], rowHeights=[0.5*cm]*2)
    bank_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('SPAN', (0,0), (0,1))]))
    story.append(bank_table)

    # --- ส่วน Payment ---
    payment1 = utils.convert_to_float(header_data.get('payment1_amount', 0))
    payment2 = utils.convert_to_float(header_data.get('payment2_amount', 0))
    
    p1_date_str = fmt_date(header_data.get('payment1_date'))
    p2_date_str = fmt_date(header_data.get('payment2_date'))
    
    if payment1 == 0 and payment2 == 0:
        total_pay_in_db = utils.convert_to_float(header_data.get('total_payment_amount', 0))
        if total_pay_in_db > 0:
            payment1 = total_pay_in_db
            if not p1_date_str:
                p1_date_str = fmt_date(header_data.get('payment_date'))

    total_deposit = payment1 + payment2
    date_display_1 = p1_date_str if (payment1 > 0 and p1_date_str) else '....../....../......'
    date_display_2 = p2_date_str if (payment2 > 0 and p2_date_str) else '....../....../......'
    
    chk_pay_done = box_checked
    chk_pay_wait = box_unchecked

    # Logic การแสดง Checkbox มัดจำ: ถ้ามียอด > 0 ถึงจะติ๊ก
    chk_p1 = box_checked if payment1 > 0 else box_unchecked
    chk_p2 = box_checked if payment2 > 0 else box_unchecked
    chk_done_p1 = chk_pay_done if payment1 > 0 else box_unchecked # ถ้าไม่มีเงิน ก็ไม่ต้องติ๊กชำระแล้ว
    chk_done_p2 = chk_pay_done if payment2 > 0 else box_unchecked

    combined_payment_data = [
        [PS(f'{chk_p1} มัดจำ 1'), PS(format_num(payment1), 'Small_Right_TH'), PS(f'{chk_done_p1} ชำระเงินแล้ว'), PS(date_display_1, 'Small_Center_TH')],
        [PS(f'{chk_p2} มัดจำ 2'), PS(format_num(payment2), 'Small_Right_TH'), PS(f'{chk_done_p2} ชำระเงินแล้ว'), PS(date_display_2, 'Small_Center_TH')],
        [PB('รวมมัดจำ', 'Small_TH'), PS(format_num(total_deposit), 'Small_Right_TH'), None, None],
        [PB('ยอดค้างชำระ', 'Small_TH'), PS(format_num(header_data.get('balance_due',0)), 'Small_Right_TH'), PS(f'{box_unchecked} ชำระเงินแล้ว'), PS('....../....../......', 'Small_Center_TH')],
        [PB('ยอดชำระรวม VAT', 'Small_TH'), PS(format_num(header_data.get('total_payment_amount',0)), 'Small_Right_TH'), PS(f'{box_unchecked} ชำระเงินแล้ว'), PS('....../....../......', 'Small_Center_TH')],
        [PB('เลขที่ใบกำกับภาษี', 'Small_TH'), PS('', 'Small_TH'), PB('วันที่ออกเอกสาร', 'Small_TH'), PS(str(header_data.get('bill_date', '')), 'Small_TH')],
        [PB('Remark*', 'Small_TH'), PSafe(header_data.get('remark', ''), 'Small_TH', 80), None, None]
    ]
    
    payment_final_table = Table(combined_payment_data, colWidths=[width*0.29, width*0.19, width*0.25, width*0.27], 
                                rowHeights=[0.5*cm]*6 + [1.0*cm])
    
    payment_final_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), 
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
        ('LEFTPADDING', (0,0), (-1,-1), 3), 
        ('RIGHTPADDING', (0,0), (-1,-1), 3), 
        ('VALIGN', (0,-1), (-1,-1), 'TOP'), 
        ('SPAN', (1,-1), (-1,-1)),
        ('SPAN', (2, 2), (3, 2)), 
        ('BACKGROUND', (0, 2), (-1, 2), colors.whitesmoke) 
    ]))
    story.append(payment_final_table)
    return story

def _build_right_column(header_data, items_data, payments_data, styles, P, PB, format_num, width):
    """
    สร้างคอลัมน์ขวา (Cost Auditor) - เพิ่มค่าบริการตัด/เจาะ และประเภทบัญชี
    """
    story = []
    
    # ... (ส่วน setup style เหมือนเดิม) ...
    def safe_add_style(styles, style):
        if style.name not in styles.byName: styles.add(style)
    
    safe_add_style(styles, ParagraphStyle(name='Small_TH', fontName='THSarabunNew', fontSize=10, leading=12))
    safe_add_style(styles, ParagraphStyle(name='Small_Center_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=1))
    safe_add_style(styles, ParagraphStyle(name='Small_Right_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=2))
    safe_add_style(styles, ParagraphStyle(name='Small_Wrapped_TH', fontName='THSarabunNew', fontSize=9, leading=11, wordWrap='CJK'))
    safe_add_style(styles, ParagraphStyle(name='Header_Bold_TH', fontName='THSarabunNew-Bold', fontSize=11, leading=13, alignment=1))
    safe_add_style(styles, ParagraphStyle(name='Product_Name_TH', fontName='THSarabunNew', fontSize=9, leading=11, wordWrap='CJK'))
    safe_add_style(styles, ParagraphStyle(name='Tiny_Center_TH', fontName='THSarabunNew', fontSize=8, leading=10, alignment=1))

    def make_para(text, style='Small_TH'):
        return Paragraph(str(text) if text is not None and str(text) != 'nan' else '', styles[style])

    def fmt_date(d):
        if not d or str(d).lower() == 'nan': return ""
        try:
            if isinstance(d, str): 
                d = d.split(' ')[0]
                d_obj = datetime.strptime(d, "%Y-%m-%d")
                return d_obj.strftime("%d/%m/%Y")
            elif hasattr(d, 'strftime'):
                return d.strftime("%d/%m/%Y")
            return str(d)
        except:
            return str(d)

    # --- 1. ตารางส่วนหัว ---
    header_widths = [1.8*cm, 1.8*cm, 1.5*cm, 2.0*cm, 1.5*cm, 1.5*cm]
    header_scale = width / sum(header_widths)
    HEADER_COL_WIDTHS = [w * header_scale for w in header_widths]
    
    header_data_grid = [
        [PB('ลำดับ', 'Small_TH'), None, make_para('COST AUDITOR', 'Header_Bold_TH'), None, PB('แผนก', 'Small_TH'), make_para(header_data.get('department', ''))],
        [PB('ชื่อ', 'Small_TH'), None, make_para(header_data.get('user_name', '')), None, PB('RR Number', 'Small_TH'), make_para(header_data.get('rr_number', ''))],
        [PB('PO NUMBER', 'Small_TH'), None, make_para(header_data.get('po_number', '')), None, PB('Credit Term', 'Small_TH'), make_para(header_data.get('credit_term', ''))],
        [PB('Supplier Name', 'Small_TH'), None, make_para(header_data.get('supplier_name', ''), 'Small_Wrapped_TH'), None, PB('REMARK', 'Small_TH'), make_para(header_data.get('remark', ''), 'Small_Wrapped_TH')]
    ]
    
    header_table = Table(header_data_grid, colWidths=HEADER_COL_WIDTHS)
    header_table.setStyle(TableStyle([ 
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
        ('LEFTPADDING', (0,0), (-1,-1), 2), 
        ('SPAN', (0,0), (1,0)), ('SPAN', (2,0), (3,0)), 
        ('SPAN', (0,1), (1,1)), ('SPAN', (2,1), (3,1)), 
        ('SPAN', (0,2), (1,2)), ('SPAN', (2,2), (3,2)),
        ('SPAN', (0,3), (1,3)), ('SPAN', (2,3), (3,3)),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#DDEBF7")), 
        ('BACKGROUND', (0,1), (1,1), colors.HexColor("#DDEBF7")), 
        ('BACKGROUND', (4,1), (4,1), colors.HexColor("#DDEBF7")), 
        ('BACKGROUND', (0,2), (1,2), colors.HexColor("#DDEBF7")),
        ('BACKGROUND', (4,2), (4,2), colors.HexColor("#DDEBF7")),
        ('BACKGROUND', (0,3), (1,3), colors.HexColor("#DDEBF7")), 
        ('BACKGROUND', (4,3), (4,3), colors.HexColor("#DDEBF7")),
    ]))
    story.append(header_table)

    # --- 2. ตารางสินค้า ---
    item_widths = [1.0*cm, 1.5*cm, 4.3*cm, 1.2*cm, 2.0*cm, 2.0*cm]
    item_scale = width / sum(item_widths)
    ITEM_COL_WIDTHS = [w * item_scale for w in item_widths]
    item_rows = []
    
    recalc_total_cost = 0.0
    for i, item in enumerate(items_data, 1):
        try: total_price = float(item.get('total_price', 0) or 0)
        except: total_price = 0.0
        recalc_total_cost += total_price
        
        item_rows.append([
            make_para(str(i), 'Small_Center_TH'), 
            make_para(item.get('status', ''), 'Small_Center_TH'), 
            make_para(item.get('product_name', ''), 'Product_Name_TH'), 
            make_para(f"{item.get('quantity', 0):.2f}", 'Small_Right_TH'), 
            make_para(format_num(item.get('unit_price', 0)), 'Small_Right_TH'), 
            make_para(format_num(total_price), 'Small_Right_TH'),
        ])
    
    while len(item_rows) < 5: item_rows.append([''] * 6)
    item_row_heights = [0.6*cm, 0.6*cm] + [None] * len(item_rows)
    
    full_item_rows = [[make_para("PURCHASED RECORD", 'Header_Bold_TH')], 
                      [make_para("ลำดับ", 'Tiny_Center_TH'), make_para("สถานะ", 'Small_Center_TH'), make_para("ชื่อสินค้า", 'Small_Center_TH'), make_para("จำนวน", 'Tiny_Center_TH'), make_para("ราคา", 'Small_Center_TH'), make_para("รวม", 'Small_Center_TH')]] + item_rows
    item_table = Table(full_item_rows, colWidths=ITEM_COL_WIDTHS, rowHeights=item_row_heights)
    item_table.setStyle(TableStyle([ ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('SPAN', (0,0), (-1,0)), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#DDEBF7")), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('ALIGN', (0,1), (-1,1), 'CENTER'), ]))
    story.append(item_table)
    
    # --- 3. ส่วนการเงินและการคำนวณใหม่ ---
    deposit_amount = 0.0; full_payment_amount = 0.0; cn_refund_amount = 0.0; latest_deposit_date = None
    full_payment_date = None; cn_refund_date = None
    
    display_bank_name = ""
    display_account_number = ""
    # [🔥 เพิ่ม] ตัวแปรสำหรับประเภทบัญชี
    display_account_type = ""

    for payment in payments_data:
        p_type = payment.get('payment_type'); amount = payment.get('amount', 0); p_date = payment.get('payment_date')
        if p_type in ["Payment 1", "Payment 2"]: 
            deposit_amount += amount
            if p_date: latest_deposit_date = p_date
        elif p_type == "Full Payment": full_payment_amount = amount; full_payment_date = p_date
        elif p_type == "CN Refund": cn_refund_amount = amount; cn_refund_date = p_date
        
        # ดึงข้อมูลธนาคารจากประวัติการจ่าย
        if not display_bank_name and payment.get('bank_name'):
            display_bank_name = payment.get('bank_name')
            display_account_number = payment.get('bank_account_number')
            # [🔥 เพิ่ม] ถ้าใน Payment มีประเภทบัญชี ให้ใช้เลย
            if payment.get('bank_account_type'):
                display_account_type = payment.get('bank_account_type')

    # ถ้าไม่มีข้อมูลใน Payment (ยังไม่จ่าย) ให้ดึงจาก Supplier Master (ที่เตรียมไว้ใน header_data)
    if not display_bank_name:
        display_bank_name = header_data.get('supplier_bank_name', '')
        display_account_number = header_data.get('supplier_account_number', '')
        
    # ถ้าประเภทบัญชียังว่างอยู่ ให้ดึงจาก Supplier Master หรือใช้ค่า Default
    if not display_account_type:
        display_account_type = header_data.get('supplier_account_type', '') or header_data.get('bank_account_type', '')

    payment_widths = [2.0*cm, 2.2*cm, 1.8*cm, 3.8*cm] 
    payment_scale = width / sum(payment_widths)
    UNIFIED_COL_WIDTHS = [w * payment_scale for w in payment_widths]

    # [คำนวณใหม่]
    # VAT
    is_vat_checked = header_data.get('vat_7_percent_checked')
    if isinstance(is_vat_checked, str): is_vat_checked = is_vat_checked.lower() in ['true', '1', 't', 'y', 'yes']
    elif isinstance(is_vat_checked, int): is_vat_checked = (is_vat_checked == 1)
    recalc_product_vat = recalc_total_cost * 0.07 if is_vat_checked else 0.0

    # Shipping Costs & VAT
    shipping_stock_cost = float(header_data.get('shipping_to_stock_cost', 0) or 0)
    shipping_stock_vat = shipping_stock_cost * 0.07 if header_data.get('shipping_to_stock_vat_type') == 'VAT' else 0.0
    shipping_site_cost = float(header_data.get('shipping_to_site_cost', 0) or 0)
    shipping_site_vat = shipping_site_cost * 0.07 if header_data.get('shipping_to_site_vat_type') == 'VAT' else 0.0
    total_shipping_cost = shipping_stock_cost + shipping_site_cost

    # Cutting/Drilling Costs & VAT
    cutting_cost = float(header_data.get('cutting_cost', 0) or 0)
    cutting_vat = cutting_cost * 0.07 if header_data.get('cutting_vat_type') == 'VAT' else 0.0
    cutting_wht_amount = float(header_data.get('cutting_wht_amount', 0) or 0)
    cutting_remark = header_data.get('cutting_remark', '')
    cutting_wht_type = header_data.get('cutting_wht_type', 'No')
    if cutting_wht_type == "No": cutting_wht_type = "ไม่มีหัก"

    # Total VAT & Grand Total
    recalc_total_vat = recalc_product_vat + shipping_stock_vat + shipping_site_vat + cutting_vat    
    recalc_grand_total = recalc_total_cost + cutting_cost + recalc_total_vat
    
    # Balance Due
    balance_due = recalc_grand_total - (deposit_amount + full_payment_amount)

    unified_payment_data = []
    
    # Top (3 rows) - [🔥 แก้ไข] แสดงประเภทบัญชีที่ถูกต้อง
    payment_data_top = [
        [PB('เลขที่บัญชี', 'Small_TH'), make_para(display_account_number), PB('รวมต้นทุน', 'Small_TH'), make_para(format_num(recalc_total_cost), 'Small_Right_TH')], 
        [PB('ธนาคาร', 'Small_TH'), make_para(display_bank_name), PB('Vat 7%', 'Small_TH'), make_para(format_num(recalc_total_vat), 'Small_Right_TH')], 
        
        # [🔥 แก้ไขจุดนี้] ใช้ตัวแปร display_account_type แทนการ get ใหม่
        [PB('ประเภท', 'Small_TH'), make_para(display_account_type, 'Small_TH'), PB('รวมทั้งสิ้น', 'Small_TH'), make_para(format_num(recalc_grand_total), 'Small_Right_TH')]
    ]
    unified_payment_data.extend(payment_data_top)
    
    # Mid (4 rows)
    payment_data_mid = [
        [PB('มัดจำ', 'Small_TH'), make_para(format_num(deposit_amount), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para(fmt_date(latest_deposit_date), 'Small_Center_TH')], 
        [PB('ยอดค้าง', 'Small_TH'), make_para(format_num(balance_due), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para('', 'Small_Center_TH')], 
        [PB('ชำระเต็ม', 'Small_TH'), make_para(format_num(full_payment_amount), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para(fmt_date(full_payment_date), 'Small_Center_TH')], 
        [PB('CN/คืนส่วนต่าง', 'Small_TH'), make_para(format_num(cn_refund_amount), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para(fmt_date(cn_refund_date), 'Small_Center_TH')]
    ]
    unified_payment_data.extend(payment_data_mid)

    # Shipping (4 rows)
    stock_wht_type = header_data.get('shipping_to_stock_wht_type', 'ไม่มีหัก')
    stock_wht_1 = shipping_stock_cost * 0.01 if stock_wht_type == '1%' else 0
    site_wht_type = header_data.get('shipping_to_site_wht_type', 'ไม่มีหัก')
    site_wht_1 = shipping_site_cost * 0.01 if site_wht_type == '1%' else 0
    total_wht_1 = stock_wht_1 + site_wht_1

    stock_wht_3 = shipping_stock_cost * 0.03 if stock_wht_type == '3%' else 0
    site_wht_3 = shipping_site_cost * 0.03 if site_wht_type == '3%' else 0
    total_wht_3 = stock_wht_3 + site_wht_3

    shipper_display = header_data.get('shipping_to_stock_shipper', '') or header_data.get('shipping_to_site_shipper', '')
    truck_info = header_data.get('truck_name', '')
    if header_data.get('license_plate'): truck_info += f" ({header_data.get('license_plate')})"
    shipping_date = fmt_date(header_data.get('shipping_date') or header_data.get('date_to_warehouse')) or "..../../.."

    shipping_data = [
        [PB('ค่าจัดส่งรับจ้าง', 'Small_TH'), make_para(format_num(total_shipping_cost), 'Small_Right_TH'), PB('วันที่จัดส่ง', 'Small_TH'), make_para(shipping_date, 'Small_Center_TH')], 
        [PB('ชื่อบริษัทจัดส่ง', 'Small_TH'), make_para(shipper_display, 'Small_Wrapped_TH'), PB('รอบส่ง', 'Small_TH'), make_para(header_data.get('shipping_round', ''), 'Small_TH')], 
        [PB('ประเภทรถ/ทะเบียน', 'Small_TH'), make_para(truck_info, 'Small_Wrapped_TH'), PB('หัก 1%', 'Small_TH'), make_para(format_num(total_wht_1), 'Small_Right_TH')], 
        [PB('ค่าจัดส่ง', 'Small_TH'), make_para(''), PB('หัก 3%', 'Small_TH'), make_para(format_num(total_wht_3), 'Small_Right_TH')]
    ]
    unified_payment_data.extend(shipping_data)

    # Cutting Data (2 rows)
    cutting_data = [
        [PB('ค่าบริการตัด/เจาะ', 'Small_TH'), make_para(format_num(cutting_cost), 'Small_Right_TH'), PB(f'หัก {cutting_wht_type}', 'Small_TH'), make_para(format_num(cutting_wht_amount), 'Small_Right_TH')],
        [PB('หมายเหตุตัด/เจาะ', 'Small_TH'), make_para(cutting_remark, 'Small_Wrapped_TH'), None, None]
    ]
    unified_payment_data.extend(cutting_data)

    # Summary (2 rows)
    summary_data = [
        [PB('ยอดชำระจริง', 'Small_TH'), make_para(''), make_para('วัน................เดือน................ปี.......', 'Small_Center_TH'), None], 
        [PB('Remark*', 'Small_TH'), make_para(header_data.get('remark', ''), 'Small_Wrapped_TH'), None, None]
    ]
    unified_payment_data.extend(summary_data)
    
    # รวมแถวทั้งหมด
    row_heights = [None] * 13 + [0.5*cm, 1.2*cm]

    payment_table_unified = Table(unified_payment_data, colWidths=UNIFIED_COL_WIDTHS, rowHeights=row_heights)
    
    # Update styles
    unified_styles = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        
        ('LINEABOVE', (0, 3), (-1, 3), 1, colors.black),  # After Top
        ('LINEABOVE', (0, 7), (-1, 7), 1, colors.black),  # After Mid
        ('LINEABOVE', (0, 11), (-1, 11), 1, colors.black), # After Shipping
        ('LINEABOVE', (0, 13), (-1, 13), 1, colors.black), # After Cutting (New)

        ('SPAN', (1,12), (3,12)), # Span for Cutting Remark
        
        ('SPAN', (2, 13), (3, 13)), # Summary Date
        ('SPAN', (1, 14), (3, 14)), # Summary Remark
        ('VALIGN', (0, 14), (-1, 14), 'TOP'), 
        ('BACKGROUND', (0, 13), (0, 13), colors.HexColor("#DDEBF7")), 
    ]
    payment_table_unified.setStyle(TableStyle(unified_styles))

    story.append(payment_table_unified)
    return story

def generate_multi_po_pdf(so_header_data, all_po_data):
    register_thai_fonts()
    documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
    if not os.path.exists(documents_path): documents_path = os.path.join(os.path.expanduser('~'), 'Desktop')

    default_filename = f"ALL_POs_for_SO_{so_header_data.get('so_number', '')}.pdf"
    save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile=default_filename, initialdir=documents_path)
    if not save_path: return

    doc = BaseDocTemplate(save_path, pagesize=A4, leftMargin=1.0*cm, rightMargin=1.0*cm, topMargin=1.2*cm, bottomMargin=1.2*cm)
    gap = 0.5 * cm
    col_width = (doc.width - gap) / 2
    
    left_frame = Frame(doc.leftMargin, doc.bottomMargin, col_width, doc.height, id='left_col')
    right_frame = Frame(doc.leftMargin + col_width + gap, doc.bottomMargin, col_width, doc.height, id='right_col')
    doc.addPageTemplates([PageTemplate(id='TwoCol', frames=[left_frame, right_frame])])

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Normal_TH', fontName='THSarabunNew', fontSize=10, leading=12))
    styles.add(ParagraphStyle(name='Bold_TH', fontName='THSarabunNew-Bold', fontSize=11, leading=13))

    def P(text, style='Normal_TH'): return Paragraph(str(text), styles[style])
    def PB(text, style='Bold_TH'): return Paragraph(str(text), styles[style])
    def format_num(value):
        try: return f"{float(value):,.2f}" if float(value) != 0 else "0.00"
        except: return "0.00"

    try:
        story = []
        for i, po_data in enumerate(all_po_data):
            if i > 0: story.append(PageBreak())
            frame_width = col_width
            story.extend(_build_left_column(so_header_data, styles, P, PB, format_num, width=frame_width))
            story.append(FrameBreak())
            story.extend(_build_right_column(po_data['header'], po_data['items'], po_data.get('payments', []), styles, P, PB, format_num, width=frame_width))
        doc.build(story)
        messagebox.showinfo("สำเร็จ", f"สร้างเอกสารรวมเรียบร้อย:\n{save_path}")
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการสร้าง PDF:\n{str(e)}")
        print(traceback.format_exc())

# --- [🔥 NEW] ฟังก์ชันสร้างใบปะหน้าค่าขนส่งโดยเฉพาะ ---
# --- [🔥 NEW FUNCTION] สร้างใบสรุปค่าขนส่ง (Transport Fee) ---

def register_thai_fonts():
    """ ลงทะเบียนฟอนต์ไทย (รองรับทั้ง Dev และ .exe) """
    try:
        # 1. ลองหาฟอนต์จาก _MEIPASS (ตอน Pack เป็น .exe)
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        
        font_path = os.path.join(base_path, "THSarabunNew.ttf")
        font_bold_path = os.path.join(base_path, "THSarabunNew Bold.ttf")
        
        # 2. ถ้าหาไม่เจอ ลองหาใน resources/
        if not os.path.exists(font_path):
            font_path = os.path.join(base_path, "resources", "THSarabunNew.ttf")
        if not os.path.exists(font_bold_path):
            font_bold_path = os.path.join(base_path, "resources", "THSarabunNew Bold.ttf")
        
        # 3. ลงทะเบียนฟอนต์
        pdfmetrics.registerFont(TTFont('THSarabunNew', font_path))
        pdfmetrics.registerFont(TTFont('THSarabunNew-Bold', font_bold_path))
        
        # 4. สร้างความสัมพันธ์ (Mapping)
        registerFontFamily('THSarabunNew', 
                          normal='THSarabunNew', 
                          bold='THSarabunNew-Bold')
        
        print(f"✅ Font loaded: {font_path}")  # Debug
        
    except Exception as e:
        print(f"❌ Font Error: {e}")
        import traceback
        traceback.print_exc()


def generate_transport_fee_pdf(so_header_data, transport_data_list):
    """
    สร้างใบสรุปค่าขนส่ง (Transportation Expense Record)
    Version: แก้ไขแล้ว - แสดงข้อมูลถูกต้อง + วันที่จ่ายช่องว่าง + หมายเหตุไม่รวมทะเบียน
    + ชื่อบริษัทผู้จัดส่งตรงกับระบบ
    """
    register_thai_fonts()
    
    so_number = so_header_data.get('so_number', 'Unknown')
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf", 
        initialfile=f"Transport_Fee_{so_number}.pdf"
    )
    if not save_path: 
        return

    styles = getSampleStyleSheet()
    st_header = ParagraphStyle('H', parent=styles['Normal'], fontName='THSarabunNew', fontSize=18, alignment=TA_CENTER)
    st_norm = ParagraphStyle('N', parent=styles['Normal'], fontName='THSarabunNew', fontSize=14, alignment=TA_LEFT)
    st_val_center = ParagraphStyle('VC', parent=styles['Normal'], fontName='THSarabunNew-Bold', fontSize=11, alignment=TA_CENTER)
    st_val_left = ParagraphStyle('VL', parent=styles['Normal'], fontName='THSarabunNew-Bold', fontSize=11, alignment=TA_LEFT)

    def P(text, style=st_val_center): 
        val = str(text).strip()
        if not val or val.lower() in ['none', 'null', 'nan', 'false', '0', '0.0']:
            return Paragraph("-", style)
        return Paragraph(val, style)
    
    def fmt_num(v): 
        try:
            return f"{float(v):,.2f}"
        except:
            return "0.00"

    doc = BaseDocTemplate(save_path, pagesize=A4, leftMargin=1*cm, rightMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')
    doc.addPageTemplates([PageTemplate(id='OneCol', frames=frame)])
    
    story = []
    story.append(Paragraph(f"<b>ใบสรุปค่าขนส่ง (Transportation Expense Record)</b>", st_header))
    story.append(Spacer(1, 0.5*cm))

    curr_date = datetime.now().strftime('%d/%m/%Y')
    info_table = Table([
        [Paragraph(f"<b>SO Number:</b> {so_number}", st_norm), Paragraph(f"<b>วันที่พิมพ์:</b> {curr_date}", st_norm)],
        [Paragraph(f"<b>ลูกค้า:</b> {so_header_data.get('customer_name', '-')}", st_norm), Paragraph(f"<b>ผู้จัดทำ/Sale:</b> {so_header_data.get('sale_name', '-')}", st_norm)]
    ], colWidths=[12*cm, 7*cm])
    info_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(info_table)
    story.append(Spacer(1, 0.5*cm))

    grand_total_cost = 0
    grand_total_paid = 0
    blue = colors.HexColor("#DDEBF7")

    for raw in transport_data_list:
        configs = [
            {
                'title': 'ค่าย้ายของ (เข้าโกดัง)',
                'cost': float(raw.get('stock_cost') or 0),
                'driver': raw.get('stock_driver', '-'),
                'plate': raw.get('stock_plate', '-'),
                'note': raw.get('stock_notes', '-'),  
                'vat': str(raw.get('stock_vat') or 'CASH').upper(),
                'date': raw.get('stock_date'),
                'wht': str(raw.get('stock_wht') or 'ไม่มีหัก'),
                'supplier_company': raw.get('stock_supplier', '-')  # ✅ ใช้ชื่อบริษัทจริง
            },
            {
                'title': 'ค่าขนส่ง (ส่งหน้างาน)',
                'cost': float(raw.get('site_cost') or 0),
                'driver': raw.get('site_driver', '-'),
                'plate': raw.get('site_plate', '-'),
                'note': raw.get('site_notes', '-'),  
                'vat': str(raw.get('site_vat') or 'CASH').upper(),
                'date': raw.get('site_date'),
                'wht': str(raw.get('site_wht') or 'ไม่มีหัก'),
                'supplier_company': raw.get('site_supplier', '-')  # ✅ ใช้ชื่อบริษัทจริง
            }
        ]

        for cfg in configs:
            cost = cfg['cost']
            is_vat = (cfg['vat'] == 'VAT')
            vat_amt = cost * 0.07 if is_vat else 0
            wht_str = cfg['wht']
            wht_rate = 0.01 if '1' in wht_str else (0.03 if '3' in wht_str else 0)
            net_paid = cost + vat_amt - (cost * wht_rate)

            grand_total_cost += cost
            grand_total_paid += net_paid

            # หมายเหตุไม่รวมทะเบียน
            note_str = cfg['note']
            full_remark = note_str if note_str != '-' else "-"

            # วันที่จัดส่ง
            s_date = cfg['date']
            s_date_str = "-"
            if s_date:
                try: 
                    if isinstance(s_date, str): 
                        s_date_str = datetime.strptime(s_date[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
                    else: 
                        s_date_str = s_date.strftime("%d/%m/%Y")
                except: 
                    pass

            # Table Top
            t_top = Table([
                [
                    Paragraph("PO NUMBER", st_val_center),
                    Paragraph("ชื่อบริษัทผู้จัดส่ง", st_val_center),
                    Paragraph("ผู้จัดส่ง/คนขับ", st_val_center),
                    Paragraph("ทะเบียนรถ", st_val_center)
                ],
                [
                    P(raw.get('po_number')),
                    P(cfg['supplier_company']),  # ✅ แสดงชื่อบริษัท
                    P(cfg['driver']),
                    P(cfg['plate'])
                ]
            ], colWidths=[3*cm, 6*cm, 5*cm, 5*cm], rowHeights=[0.7*cm, 0.9*cm])
            t_top.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('BACKGROUND', (0,0), (-1,0), blue),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER')
            ]))

            # Table Bottom
            chk_v = "[ / ] vat" if is_vat else "[   ] vat"
            chk_w1 = "[ / ] 1%" if wht_rate == 0.01 else "[   ] 1%"
            chk_w3 = "[ / ] 3%" if wht_rate == 0.03 else "[   ] 3%"

            t_bot = Table([
                [
                    Paragraph("วันที่จัดส่ง", st_val_center),
                    Paragraph("ค่าจัดส่ง", st_val_center),
                    Paragraph("vat", st_val_center),
                    Paragraph("หัก ณ ที่จ่าย", st_val_center),
                    Paragraph("ชำระจริง", st_val_center),
                    Paragraph("วันที่จ่าย", st_val_center)
                ],
                [
                    P(s_date_str),
                    P(fmt_num(cost)),
                    Paragraph(chk_v, st_val_center),
                    Paragraph(f"{chk_w1}<br/>{chk_w3}", st_val_center),
                    P(fmt_num(net_paid)),
                    Paragraph("", st_val_center)
                ],
                [
                    Paragraph("ประเภทรายการ", st_norm),
                    Paragraph(cfg['title'], st_val_left),
                    '', '', '', ''
                ],
                [
                    Paragraph("หมายเหตุ:", st_norm),
                    Paragraph(full_remark, st_val_left),
                    '', '', '', ''
                ]
            ], colWidths=[3*cm, 3.5*cm, 2*cm, 3*cm, 4.5*cm, 3*cm], rowHeights=[0.7*cm, 1.2*cm, 0.7*cm, 0.7*cm])
            t_bot.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('BACKGROUND', (0,0), (-1,0), blue),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('SPAN', (1,2), (-1,2)), ('ALIGN', (1,2), (-1,2), 'LEFT'),
                ('SPAN', (1,3), (-1,3)), ('ALIGN', (1,3), (-1,3), 'LEFT')
            ]))

            story.extend([t_top, t_bot, Spacer(1, 0.5*cm)])

    # Total
    t_total = Table([
        [
            Paragraph("<b>ยอดรวมค่าขนส่ง</b>", st_norm),
            P(fmt_num(grand_total_cost)),
            '', '', P(fmt_num(grand_total_paid))
        ]
    ], colWidths=[3*cm, 3.5*cm, 2*cm, 3*cm, 7.5*cm], rowHeights=[0.8*cm])

    t_total.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),           # รอบด้าน
        ('LINEBEFORE', (1,0), (1,0), 0.5, colors.black),    # เส้นระหว่างช่องยอดรวม
        ('LINEBEFORE', (4,0), (4,0), 0.5, colors.black),    # เส้นระหว่างช่องชำระจริง
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'RIGHT'),
        ('RIGHTPADDING', (0,0), (0,0), 10)
    ]))
    story.append(Spacer(1, 0.2*cm))
    story.append(t_total)

    try:
        doc.build(story)
        messagebox.showinfo("สำเร็จ", f"บันทึกใบสรุปค่าขนส่งเรียบร้อยแล้วที่:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"สร้าง PDF ไม่สำเร็จ: {e}")
        print(traceback.format_exc())