import os
import traceback
import utils
from tkinter import filedialog, messagebox
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, FrameBreak, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)
from reportlab.platypus.frames import Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
import sys
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def register_thai_fonts():
    """Registers Thai fonts for ReportLab."""
    try:
        font_path = resource_path("resources/THSarabunNew.ttf")
        font_bold_path = resource_path("resources/THSarabunNew Bold.ttf")
        pdfmetrics.registerFont(TTFont('THSarabunNew', font_path))
        pdfmetrics.registerFont(TTFont('THSarabunNew-Bold', font_bold_path))
        pdfmetrics.registerFontFamily('THSarabunNew', normal='THSarabunNew', bold='THSarabunNew-Bold')
    except Exception:
        try:
            pdfmetrics.registerFont(TTFont('THSarabunNew', 'THSarabunNew.ttf'))
            pdfmetrics.registerFont(TTFont('THSarabunNew-Bold', 'THSarabunNew Bold.ttf'))
            pdfmetrics.registerFontFamily('THSarabunNew', normal='THSarabunNew', bold='THSarabunNew-Bold')
        except Exception as fallback_e:
            print(f"ERROR: Could not register fonts. Error: {fallback_e}")


def _build_left_column(header_data, styles, P, PB, format_num, width):
    """
    สร้างคอลัมน์ซ้าย (SELL AUDITOR) - (เวอร์ชันปรับปรุงความยืดหยุ่น)
    """
    story = []
    def safe_add_style(styles, style):
        if style.name not in styles.byName: styles.add(style)

    safe_add_style(styles, ParagraphStyle(name='Small_TH', fontName='THSarabunNew', fontSize=10, leading=12))
    safe_add_style(styles, ParagraphStyle(name='Tiny_TH', fontName='THSarabunNew', fontSize=9, leading=11))
    safe_add_style(styles, ParagraphStyle(name='Small_Center_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=1))
    safe_add_style(styles, ParagraphStyle(name='Small_Right_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=2))

    def PS(text, style='Small_TH'): return Paragraph(str(text) if text is not None else '', styles[style])
    def PSafe(text, style='Small_TH', max_length=25):
        text_str = str(text)
        if len(text_str) > max_length: text_str = text_str[:max_length-3] + "..."
        return Paragraph(text_str, styles[style])
    
    c1, c3, c4 = 2.2*cm, 1.0*cm, 2.0*cm 
    c2 = width - (c1 + c3 + c4)

    combined_header_data = [
        [PB('ขาย', 'Small_TH'), PS('SELL AUDITOR', 'Small_Center_TH'), PB('ผู้ตรวจ..............', 'Small_TH'), None],
        [PB('SO NUMBER', 'Small_TH'), PSafe(header_data.get('so_number', ''), 'Small_TH'), PB('แผนก', 'Small_TH'), PS(header_data.get('department', ''))],
        [PB('Sale Name', 'Small_TH'), PSafe(header_data.get('sale_name', ''), 'Small_TH'), PB('วันที่', 'Small_TH'), PS(str(header_data.get('bill_date', '')))],
        [PB('Customer Name', 'Small_TH'), PSafe(header_data.get('customer_name', ''), 'Small_TH', 40), None, None],
    ]
    
    # --- จุดที่แก้ไข: ปล่อยให้ความสูงแถวเป็นอัตโนมัติ ---
    header_table = Table(combined_header_data, colWidths=[c1, c2, c3, c4]) 
    header_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('SPAN', (2,0), (3,0)), ('SPAN', (1,3), (-1,3)),
        ('BACKGROUND', (0,1), (0,3), colors.lemonchiffon), ('BACKGROUND', (2,1), (2,2), colors.lemonchiffon)
    ]))
    story.append(header_table)

    # ... ส่วนที่เหลือของฟังก์ชันนี้เหมือนเดิม ...
    title_style = TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,-1), colors.lemonchiffon), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')])
    story.append(Table([[PB('SELLING RECORD', style='Small_Center_TH')]], colWidths=[width], rowHeights=[0.5*cm], style=title_style))

    sales_before_vat = utils.convert_to_float(header_data.get('sales_service_amount', 0))
    sales_vat = utils.convert_to_float(header_data.get('product_vat_7', 0))
    shipping_before_vat = utils.convert_to_float(header_data.get('shipping_cost', 0))
    shipping_vat = shipping_before_vat * 0.07 if header_data.get('shipping_vat_option') == 'VAT' else 0.0
    
    selling_data = [
        [PB('ยอดขายสินค้าก่อน VAT', 'Small_TH'), PS(format_num(sales_before_vat), 'Small_Right_TH'), PS('☑ Vat 7%', 'Tiny_TH'), PS('☑ ไม่เอาVat', 'Tiny_TH')],
        [PB('ค่าตัดเหล็ก', 'Small_TH'), PS(format_num(header_data.get('cutting_drilling_fee', 0)), 'Small_Right_TH'), PS('☑ Vat 7%', 'Tiny_TH'), PS('☑ ไม่เอาVat', 'Tiny_TH')],
        [PB('Vat 7% ค่าสินค้า', 'Small_TH'), PS(format_num(sales_vat), 'Small_Right_TH'), PB('ภาษีถูกหัก ด ที่จ่าย', 'Small_TH'), PS(format_num(header_data.get('wht_3_percent', 0)), 'Small_Right_TH')],
        [PB('ค่าจัดส่งก่อน Vat', 'Small_TH'), PS(format_num(shipping_before_vat), 'Small_Right_TH'), PB('ค่าธรรมเนียมโอน', 'Small_TH'), PS(format_num(header_data.get('transfer_fee', 0)), 'Small_Right_TH')],
        [PB('Vat 7% ค่าจัดส่ง', 'Small_TH'), PS(format_num(shipping_vat), 'Small_Right_TH'), PB('ส่วนลด', 'Small_TH'), PS(format_num(header_data.get('coupons', 0)), 'Small_Right_TH')],
        [PB('ยอดขายรวมทั้งสิ้น', 'Small_TH'), PS(format_num(header_data.get('so_grand_total', 0)), 'Small_Right_TH'), PB('ค่าการตลาด', 'Small_TH'), PS(format_num(header_data.get('marketing_fee', 0)), 'Small_Right_TH')],
    ]
    selling_table = Table(selling_data, colWidths=[width*0.29, width*0.16, width*0.29, width*0.26], rowHeights=[0.5*cm]*6)
    selling_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3)]))
    story.append(selling_table)
    story.append(Table([[PB('ยอดชำระค่าสินค้า/บริการ', style='Small_Center_TH')]], colWidths=[width], rowHeights=[0.5*cm], style=title_style))
    bank_data = [[PB('บัญชีที่ลูกค้าโอน', 'Small_TH'), PS('☑ ธ.กสิกรไทย', 'Small_TH'), PS('☑ ธ.ทหารไทย(ออมทรัพย์)', 'Small_TH')], [None, PS('☑ ธ.ทหารไทย(กระแสรายวัน)', 'Small_TH'), PS('☑ บัญชีกรรมการ', 'Small_TH')]]
    bank_table = Table(bank_data, colWidths=[width*0.29, width*0.35, width*0.36], rowHeights=[0.5*cm]*2)
    bank_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('SPAN', (0,0), (0,1))]))
    story.append(bank_table)
    combined_payment_data = [
        [PS('☑ มัดจำ 1 ชำระเต็ม'), PS(format_num(header_data.get('payment1_amount',0)), 'Small_Right_TH'), PS('☑ ชำระเงินแล้ว'), PS('....../....../......', 'Small_Center_TH')],
        [PS('☑ มัดจำ 2'), PS(format_num(header_data.get('payment2_amount',0)), 'Small_Right_TH'), PS('☑ ชำระเงินแล้ว'), PS('....../....../......', 'Small_Center_TH')],
        [PB('ยอดค้างชำระ', 'Small_TH'), PS(format_num(header_data.get('balance_due',0)), 'Small_Right_TH'), PS('☑ ชำระเงินแล้ว'), PS('....../....../......', 'Small_Center_TH')],
        [PB('ยอดชำระรวม VAT', 'Small_TH'), PS(format_num(header_data.get('total_payment_amount',0)), 'Small_Right_TH'), PS('☑ ชำระเงินแล้ว'), PS('....../....../......', 'Small_Center_TH')],
        [PB('เลขที่ใบกำกับภาษี', 'Small_TH'), PS('', 'Small_TH'), PB('วันที่ออกเอกสาร', 'Small_TH'), PS(str(header_data.get('bill_date', '')), 'Small_TH')],
        [PB('Remark*', 'Small_TH'), PSafe(header_data.get('remark', ''), 'Small_TH', 80), None, None]
    ]
    payment_final_table = Table(combined_payment_data, colWidths=[width*0.29, width*0.19, width*0.25, width*0.27], rowHeights=[0.5*cm]*5 + [1.0*cm])
    payment_final_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3), ('VALIGN', (0,-1), (-1,-1), 'TOP'), ('SPAN', (1,-1), (-1,-1))]))
    story.append(payment_final_table)
    return story

def _build_right_column(header_data, items_data, payments_data, styles, P, PB, format_num, width):
    """
    สร้างคอลัมน์ขวา (เวอร์ชันสุดท้าย ปรับ Header และ Logic จำนวนแถว)
    """
    story = []
    
    def safe_add_style(styles, style):
        if style.name not in styles.byName: styles.add(style)
    
    safe_add_style(styles, ParagraphStyle(name='Small_TH', fontName='THSarabunNew', fontSize=10, leading=12))
    safe_add_style(styles, ParagraphStyle(name='Small_Center_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=1))
    safe_add_style(styles, ParagraphStyle(name='Small_Right_TH', fontName='THSarabunNew', fontSize=10, leading=12, alignment=2))
    safe_add_style(styles, ParagraphStyle(name='Small_Wrapped_TH', fontName='THSarabunNew', fontSize=9, leading=11, wordWrap='CJK'))
    safe_add_style(styles, ParagraphStyle(name='Header_Bold_TH', fontName='THSarabunNew-Bold', fontSize=11, leading=13, alignment=1))
    safe_add_style(styles, ParagraphStyle(name='Product_Name_TH', fontName='THSarabunNew', fontSize=9, leading=11, wordWrap='CJK'))
    
    # --- จุดที่แก้ไข 1.1: เพิ่ม Style ตัวอักษรขนาดเล็กพิเศษสำหรับหัวตาราง ---
    safe_add_style(styles, ParagraphStyle(name='Tiny_Center_TH', fontName='THSarabunNew', fontSize=8, leading=10, alignment=1))

    def make_para(text, style='Small_TH'):
        return Paragraph(str(text) if text is not None else '', styles[style])

    # --- 1. ตารางส่วนหัว (Header) ---
    header_widths = [1.8*cm, 1.8*cm, 1.5*cm, 2.0*cm, 1.5*cm, 1.5*cm]
    header_scale = width / sum(header_widths)
    HEADER_COL_WIDTHS = [w * header_scale for w in header_widths]
    header_data_grid = [
        [PB('ลำดับ', 'Small_TH'), None, make_para('COST AUDITOR', 'Header_Bold_TH'), None, PB('แผนก', 'Small_TH'), make_para(header_data.get('department', ''))],
        [PB('ชื่อ', 'Small_TH'), None, make_para(header_data.get('user_name', '')), None, PB('RR Number', 'Small_TH'), make_para(header_data.get('rr_number', ''))],
        [PB('PO NUMBER', 'Small_TH'), None, make_para(header_data.get('po_number', '')), None, None, None],
        [PB('Supplier Name', 'Small_TH'), make_para(header_data.get('supplier_name', ''), 'Small_Wrapped_TH'), PB('REMARK', 'Small_TH'), make_para(header_data.get('remark', ''), 'Small_Wrapped_TH'), None, None]
    ]
    header_table = Table(header_data_grid, colWidths=HEADER_COL_WIDTHS)
    header_table.setStyle(TableStyle([ 
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
        ('LEFTPADDING', (0,0), (-1,-1), 2), ('SPAN', (0,0), (1,0)), ('SPAN', (2,0), (3,0)), 
        ('SPAN', (4,0), (5,0)), ('SPAN', (0,1), (1,1)), ('SPAN', (2,1), (3,1)), ('SPAN', (4,1), (5,1)),
        ('SPAN', (0,2), (1,2)), ('SPAN', (2,2), (5,2)), ('SPAN', (1,3), (2,3)), ('SPAN', (4,3), (5,3)),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#DDEBF7")), ('BACKGROUND', (0,1), (1,1), colors.HexColor("#DDEBF7")), 
        ('BACKGROUND', (4,1), (5,1), colors.HexColor("#DDEBF7")), ('BACKGROUND', (0,2), (1,2), colors.HexColor("#DDEBF7")),
        ('BACKGROUND', (0,3), (0,3), colors.HexColor("#DDEBF7")), ('BACKGROUND', (3,3), (3,3), colors.HexColor("#DDEBF7")),
    ]))
    story.append(header_table)

    # --- 2. ตารางรายการสินค้า (Items) ---
    item_widths = [1.0*cm, 1.5*cm, 4.3*cm, 1.2*cm, 2.0*cm, 2.0*cm]
    item_scale = width / sum(item_widths)
    ITEM_COL_WIDTHS = [w * item_scale for w in item_widths]
    item_rows = []
    for i, item in enumerate(items_data, 1):
        item_rows.append([make_para(str(i), 'Small_Center_TH'), make_para(item.get('status', ''), 'Small_Center_TH'), make_para(item.get('product_name', ''), 'Product_Name_TH'), make_para(f"{item.get('quantity', 0):.2f}", 'Small_Right_TH'), make_para(format_num(item.get('unit_price', 0)), 'Small_Right_TH'), make_para(format_num(item.get('total_price', 0)), 'Small_Right_TH'),])
    
    # --- จุดที่แก้ไข 2: เปลี่ยน Logic ให้เติมแถวจนครบ 5 แถวเสมอ ---
    while len(item_rows) < 5:
        item_rows.append([''] * 6)
    
    item_row_heights = [0.6*cm, 0.6*cm] + [None] * len(item_rows)
    
    # --- จุดที่แก้ไข 1.2: ใช้ Style ตัวอักษรขนาดเล็กพิเศษกับหัวตารางที่มีปัญหา ---
    item_header_row = [
        make_para("ลำดับ", 'Tiny_Center_TH'), 
        make_para("สถานะ", 'Small_Center_TH'), 
        make_para("ชื่อสินค้า", 'Small_Center_TH'), 
        make_para("จำนวน", 'Tiny_Center_TH'), 
        make_para("ราคา", 'Small_Center_TH'), 
        make_para("รวม", 'Small_Center_TH')
    ]
    
    full_item_rows = [[make_para("PURCHASED RECORD", 'Header_Bold_TH')], item_header_row] + item_rows
    item_table = Table(full_item_rows, colWidths=ITEM_COL_WIDTHS, rowHeights=item_row_heights)
    item_table.setStyle(TableStyle([ ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('SPAN', (0,0), (-1,0)), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#DDEBF7")), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('ALIGN', (0,1), (-1,1), 'CENTER'), ]))
    story.append(item_table)
    
    # --- 3. ตารางการชำระเงิน (Payment) ---
    deposit_amount = 0.0; full_payment_amount = 0.0; cn_refund_amount = 0.0; latest_deposit_date = None
    full_payment_date = None; cn_refund_date = None; display_bank_name = ""; display_account_number = ""
    for payment in payments_data:
        p_type = payment.get('payment_type'); amount = payment.get('amount', 0); p_date = payment.get('payment_date')
        if p_type in ["Payment 1", "Payment 2"]: deposit_amount += amount
        elif p_type == "Full Payment": full_payment_amount = amount; full_payment_date = p_date
        elif p_type == "CN Refund": cn_refund_amount = amount; cn_refund_date = p_date
        if not display_bank_name and payment.get('bank_name'):
            display_bank_name = payment.get('bank_name'); display_account_number = payment.get('bank_account_number')
    grand_total = header_data.get('grand_total', 0) or 0.0; balance_due = grand_total - (deposit_amount + full_payment_amount)

    payment_widths = [2.6*cm, 2.6*cm, 2.0*cm, 2.6*cm]
    payment_scale = width / sum(payment_widths)
    PAYMENT_COL_WIDTHS = [w * payment_scale for w in payment_widths]
    payment_data_top = [[PB('เลขที่บัญชี', 'Small_TH'), make_para(display_account_number), PB('รวมต้นทุน', 'Small_TH'), make_para(format_num(header_data.get('total_cost', 0)), 'Small_Right_TH')], [PB('ธนาคาร', 'Small_TH'), make_para(display_bank_name), PB('Vat 7%', 'Small_TH'), make_para(format_num(header_data.get('vat_7_percent_amount', 0)), 'Small_Right_TH')], [PB('ประเภท', 'Small_TH'), make_para(header_data.get('bank_account_type', ''), 'Small_TH'), PB('รวมทั้งสิ้น', 'Small_TH'), make_para(format_num(grand_total), 'Small_Right_TH')]]
    payment_table_top = Table(payment_data_top, colWidths=PAYMENT_COL_WIDTHS)
    payment_table_top.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3)]))
    story.append(payment_table_top)

    payment2_widths = [2.0*cm, 2.2*cm, 1.8*cm, 3.8*cm]
    payment2_scale = width / sum(payment2_widths)
    PAYMENT2_COL_WIDTHS = [w * payment2_scale for w in payment2_widths]
    payment_data_mid = [[PB('มัดจำ', 'Small_TH'), make_para(format_num(deposit_amount), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para(str(latest_deposit_date) if latest_deposit_date else '', 'Small_Center_TH')], [PB('ยอดค้าง', 'Small_TH'), make_para(format_num(balance_due), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para('', 'Small_Center_TH')], [PB('ชำระเต็ม', 'Small_TH'), make_para(format_num(full_payment_amount), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para(str(full_payment_date) if full_payment_date else '', 'Small_Center_TH')], [PB('CN/คืนส่วนต่าง', 'Small_TH'), make_para(format_num(cn_refund_amount), 'Small_Right_TH'), PB('วันที่', 'Small_TH'), make_para(str(cn_refund_date) if cn_refund_date else '', 'Small_Center_TH')]]
    payment_table_mid = Table(payment_data_mid, colWidths=PAYMENT2_COL_WIDTHS)
    payment_table_mid.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('LINEABOVE', (0,0), (-1,0), 1, colors.black)]))
    story.append(payment_table_mid)
    
    shipping_widths = [2.8*cm, 2.4*cm, 2.0*cm, 2.6*cm]
    shipping_scale = width / sum(shipping_widths)
    SHIPPING_COL_WIDTHS = [w * shipping_scale for w in shipping_widths]
    shipping_cost = header_data.get('shipping_to_stock_cost', 0) + header_data.get('shipping_to_site_cost', 0)
    shipping_data = [[PB('ค่าจัดส่งรับจ้าง', 'Small_TH'), make_para(format_num(shipping_cost), 'Small_Right_TH'), PB('วันที่จัดส่ง', 'Small_TH'), make_para('..../../..', 'Small_Center_TH')], [PB('ชื่อบริษัทจัดส่ง', 'Small_TH'), make_para(header_data.get('shipper_1', '')), PB('รอบส่ง', 'Small_TH'), make_para(header_data.get('delivery_round', ''), 'Small_TH')], [PB('ประเภทรถ', 'Small_TH'), make_para(''), PB('หัก 1%', 'Small_TH'), make_para(format_num(header_data.get('shipping_wht_1_percent_amount',0)), 'Small_Right_TH')], [PB('ค่าจัดส่ง', 'Small_TH'), make_para(''), PB('หัก 3%', 'Small_TH'), make_para(format_num(header_data.get('shipping_wht_3_percent_amount',0)), 'Small_Right_TH')]]
    shipping_table = Table(shipping_data, colWidths=SHIPPING_COL_WIDTHS)
    shipping_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('LINEABOVE', (0,0), (-1,0), 1, colors.black), ('SPAN', (1,3), (1,3)),]))
    story.append(shipping_table)
    
    summary_data = [[PB('ยอดชำระจริง', 'Small_TH'), make_para(''), make_para('วัน................เดือน................ปี.......', 'Small_Center_TH')], [PB('Remark*', 'Small_TH'), make_para(header_data.get('remark', ''), 'Small_Wrapped_TH'), None]]
    summary_table = Table(summary_data, colWidths=[width * 0.3, width * 0.3, width * 0.4], rowHeights=[0.5*cm, 1.2*cm])
    summary_table.setStyle(TableStyle([ ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 3), ('BACKGROUND', (0,0), (0,0), colors.HexColor("#DDEBF7")), ('SPAN', (1,1), (2,1)), ('VALIGN', (0,1), (-1,1), 'TOP'), ('LINEABOVE', (0,0), (-1,0), 1, colors.black), ]))
    story.append(summary_table)
    
    return story
    
def generate_multi_po_pdf(so_header_data, all_po_data):
    """
    (เวอร์ชันแก้ไขสมบูรณ์) ใช้ PageTemplate แบบ 2 คอลัมน์ (2 Frames)
    เพื่อแก้ปัญหา LayoutError อย่างถาวร
    """
    register_thai_fonts()
    
    documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
    if not os.path.exists(documents_path):
        documents_path = os.path.join(os.path.expanduser('~'), 'Desktop')

    default_filename = f"ALL_POs_for_SO_{so_header_data.get('so_number', '')}.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
        initialfile=default_filename, initialdir=documents_path
    )
    if not save_path: return

    doc = BaseDocTemplate(save_path, pagesize=A4, 
                          leftMargin=1.0*cm, rightMargin=1.0*cm,
                          topMargin=1.2*cm, bottomMargin=1.2*cm)

    # --- จุดที่แก้ไข: สร้าง Page Template ที่มี 2 Frames (2 คอลัมน์) ---
    gap = 0.5 * cm
    col_width = (doc.width - gap) / 2
    
    left_frame = Frame(doc.leftMargin, doc.bottomMargin, col_width, doc.height, id='left_col')
    right_frame = Frame(doc.leftMargin + col_width + gap, doc.bottomMargin, col_width, doc.height, id='right_col')

    two_column_template = PageTemplate(id='TwoCol', frames=[left_frame, right_frame])
    doc.addPageTemplates([two_column_template])
    # --- สิ้นสุดการแก้ไข ---

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Normal_TH', fontName='THSarabunNew', fontSize=10, leading=12))
    styles.add(ParagraphStyle(name='Bold_TH', fontName='THSarabunNew-Bold', fontSize=11, leading=13))

    def P(text, style='Normal_TH'): return Paragraph(str(text), styles[style])
    def PB(text, style='Bold_TH'): return Paragraph(str(text), styles[style])
    def format_num(value):
        try:
            val = float(value)
            return f"{val:,.2f}" if val != 0 else "0.00"
        except (ValueError, TypeError):
            return str(value) if value is not None else "0.00"

    try:
        story = []
        for i, po_data in enumerate(all_po_data):
            if i > 0:
                story.append(PageBreak())
            
            frame_width = col_width

            left_content = _build_left_column(so_header_data, styles, P, PB, format_num, width=frame_width)
            right_content = _build_right_column(
                po_data['header'], 
                po_data['items'],
                po_data.get('payments', []),
                styles, P, PB, format_num,
                width=frame_width
            )
            
            # --- จุดที่แก้ไข: เพิ่มเนื้อหาลง story โดยตรง ไม่ใช้ Master Table ---
            story.extend(left_content)
            story.append(FrameBreak())
            story.extend(right_content)
            # --- สิ้นสุดการแก้ไข ---
        
        doc.build(story)
        messagebox.showinfo("สำเร็จ", f"สร้างเอกสารรวมเรียบร้อย:\n{save_path}")
    
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการสร้าง PDF:\n{str(e)}")
        print(traceback.format_exc())
