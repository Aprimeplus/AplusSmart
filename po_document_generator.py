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
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- ทำการลงทะเบียนฟอนต์ทั้งหมดที่ใช้ในโปรเจกต์ ---
try:
    # ใช้ resource_path เพื่อให้หาไฟล์เจอไม่ว่าจะรันจาก .py หรือ .exe
    font_path = resource_path("THSarabunNew.ttf")
    font_bold_path = resource_path("THSarabunNew Bold.ttf")
    # (ถ้ามีฟอนต์ Italic หรือ BoldItalic ก็เพิ่มได้เลย)
    # font_italic_path = resource_path("THSarabunNew Italic.ttf")
    # font_bold_italic_path = resource_path("THSarabunNew BoldItalic.ttf")

    # ลงทะเบียนฟอนต์กับ reportlab
    pdfmetrics.registerFont(TTFont('THSarabunNew', font_path))
    pdfmetrics.registerFont(TTFont('THSarabunNew-Bold', font_bold_path))
    # pdfmetrics.registerFont(TTFont('THSarabunNew-Italic', font_italic_path))
    # pdfmetrics.registerFont(TTFont('THSarabunNew-BoldItalic', font_bold_italic_path))

    print("Fonts registered successfully for PDF generation.")
except Exception as e:
    print(f"ERROR: Could not register fonts for PDF generation: {e}")

def register_thai_fonts():
    """Registers Thai fonts for ReportLab."""
    font_dir = 'resources'
    if not os.path.isdir(font_dir):
        os.makedirs(font_dir)
    
    fonts = {
        'THSarabunNew': 'THSarabunNew.ttf',
        'THSarabunNew-Bold': 'THSarabunNew Bold.ttf',
    }
    for name, filename in fonts.items():
        font_path = os.path.join(font_dir, filename)
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont(name, font_path))
    pdfmetrics.registerFontFamily('THSarabunNew', normal='THSarabunNew', bold='THSarabunNew-Bold')

def _build_left_column(header_data, styles, P, PB, format_num):
    """
    สร้างคอลัมน์ด้านซ้ายของเอกสาร PDF (SELL AUDITOR)
    -- โค้ดเวอร์ชันสุดท้าย ปรับปรุงตามข้อเสนอแนะทั้งหมด --
    """
    story = []
    
    # 1. SO Info Table
    so_info_data = [
        [PB("A+Smart-SA"), PB("SELL AUDITOR"), ""],
        [PB("SO NUMBER"), P(header_data.get('so_number', '')), P(str(header_data.get('bill_date', '')))],
        [PB("Sale Name"), P(header_data.get('sale_name', '')), P(f"{header_data.get('commission_month', '')}/{header_data.get('commission_year', '')}")],
        [PB("Customer Name"), P(header_data.get('customer_name', '')), P(header_data.get('credit_term', ''))]
    ]
    so_info_table = Table(so_info_data, colWidths=[3*cm, 3.25*cm, 3.25*cm], rowHeights=0.7*cm)
    so_info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('SPAN', (1,0), (2,0)), 
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,0), colors.lemonchiffon),
    ]))
    story.append(so_info_table)
    
    # 2. SELLING RECORD Table
    sales_amount = utils.convert_to_float(header_data.get('sales_service_amount', 0))
    card_fee = utils.convert_to_float(header_data.get('credit_card_fee', 0))
    cutting_fee = utils.convert_to_float(header_data.get('cutting_drilling_fee', 0))
    other_fee = utils.convert_to_float(header_data.get('other_service_fee', 0))
    shipping_cost = utils.convert_to_float(header_data.get('shipping_cost', 0))
    sales_vat = sales_amount * 0.07 if header_data.get('sales_service_vat_option') == 'VAT' else 0.0
    card_fee_vat = card_fee * 0.07 if header_data.get('credit_card_fee_vat_option') == 'VAT' else 0.0
    cutting_vat = cutting_fee * 0.07 if header_data.get('cutting_drilling_fee_vat_option') == 'VAT' else 0.0
    other_vat = other_fee * 0.07 if header_data.get('other_service_fee_vat_option') == 'VAT' else 0.0
    shipping_vat = shipping_cost * 0.07 if header_data.get('shipping_vat_option') == 'VAT' else 0.0
    selling_data = [
        [PB("สรุปยอดขายสินค้า/บริการ"), "", PB("ค่าใช้จ่ายอื่นๆ"), ""],
        [PB("รายการ"), PB("ยอดการขาย"), PB("รายการ"), PB("ยอดเงิน")],
        [P("ยอดขายสินค้า/บริการ"), P(format_num(sales_amount)), P("ค่าธรรมเนียมบัตร"), P(format_num(card_fee))],
        [P("Vat 7% ยอดขายสินค้า/บริการ"), P(format_num(sales_vat)), P("Vat 7% ค่าธรรมเนียมบัตร"), P(format_num(card_fee_vat))],
        [P("ค่าตัด/เจาะเหล็ก"), P(format_num(cutting_fee)), P("ค่าธรรมเนียมโอน (หากมี)"), P(format_num(header_data.get('transfer_fee', 0)))],
        [P("Vat 7% ค่าตัด/เจาะเหล็ก"), P(format_num(cutting_vat)), P("ภาษี หัก ณ ที่จ่าย (หากมี)"), P(format_num(header_data.get('wht_3_percent', 0)))],
        [P("ค่าบริการอื่นๆ"), P(format_num(other_fee)), P("ค่าการตลาด"), P(format_num(header_data.get('marketing_fee', 0)))],
        [P("Vat 7% ค่าบริการอื่นๆ"), P(format_num(other_vat)), P("ค่านายหน้า"), P(format_num(header_data.get('brokerage_fee', 0)))],
        [P("ค่าจัดส่ง/ค่าย้าย"), P(format_num(shipping_cost)), P("คูปอง"), P(format_num(header_data.get('coupons', 0)))],
        [P("Vat 7% ค่าจัดส่ง"), P(format_num(shipping_vat)), P("ของแถม"), P(format_num(header_data.get('giveaways', 0)))],
    ]
    selling_table = Table(selling_data, colWidths=[3.5*cm, 1.7*cm, 2.8*cm, 1.5*cm], rowHeights=0.6*cm)
    selling_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('ALIGN', (1,2), (1,-1), 'RIGHT'), ('ALIGN', (3,2), (3,-1), 'RIGHT'),
        ('SPAN', (0,0), (1,0)), ('SPAN', (2,0), (3,0)), 
        ('ALIGN', (0,0), (-1,1), 'CENTER'), ('BACKGROUND', (0,0), (-1,1), colors.lemonchiffon),
    ]))
    story.append(selling_table)

    # 3. Payment Table
    payment_data = [
        [PB("การชำระเงิน"), PB("ยอดเงิน"), PB("วันที่ชำระ"), PB("ธนาคารที่ชำระ")],
        [P("มัดจำ 1 / ชำระเต็ม"), P(format_num(header_data.get('payment1_amount', 0))), P(""), P("ไม่เลือก")],
        [P("มัดจำ 2"), P(format_num(header_data.get('payment2_amount', 0))), P(""), P("ไม่เลือก")],
        [P("ยอดโอนชำระรวม VAT"), P(format_num(header_data.get('total_payment_amount', 0))), P("ตรวจสอบยอด"), P("Delivery Note")],
        [P("ยอดค้างชำระ"), P(format_num(header_data.get('balance_due', 0))), "", ""]
    ]
    payment_table = Table(payment_data, colWidths=[3.2*cm, 1.8*cm, 1.5*cm, 3.0*cm], rowHeights=0.6*cm)
    payment_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('ALIGN', (1,1), (1,-1), 'RIGHT'),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('SPAN', (2, 4), (3, 4)), 
        ('ALIGN', (2,3), (2,3), 'CENTER'),
    ]))
    story.append(payment_table)

    # 4. Cash Payment Table
    pickup_location = header_data.get('pickup_location', '') or 'None'
    relocation_cost = format_num(header_data.get('relocation_cost', 0))
    date_to_warehouse = str(header_data.get('date_to_warehouse', '')) if header_data.get('date_to_warehouse') else 'None'
    date_to_customer = str(header_data.get('date_to_customer', '')) if header_data.get('date_to_customer') else 'None'
    pickup_registration = header_data.get('pickup_registration', '') or 'None'
    
    delivery_note_inner_data = [
        [P("Location เข้ารับ:"), P(f"{pickup_location}")],
        [P("ค่าย้าย:"), P(f"{relocation_cost}")],
        [P("วันที่ย้ายเข้าคลัง:"), P(f"{date_to_warehouse}")],
        [P("วันที่จัดส่งลูกค้า:"), P(f"{date_to_customer}")],
        [P("ทะเบียนเข้ารับ:"), P(f"{pickup_registration}")],
    ]
    
    inner_table_colWidths = [1.5*cm, 1.5*cm] 
    inner_table_row_height = 0.6*cm
    
    delivery_note_inner_table = Table(delivery_note_inner_data, 
                                      colWidths=inner_table_colWidths, 
                                      rowHeights=[inner_table_row_height]*len(delivery_note_inner_data))
    
    delivery_note_inner_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('FONTSIZE', (0,0), (-1,-1), 8), ('FONTNAME', (0,0), (-1,-1), 'THSarabunNew'),
    ]))

    check_balance_text = P("ตรวจสอบยอด")
    
    cash_data = [
        [PB("การชำระเงิน (CASH)"), PB("ยอดเงิน"), PB("วันที่ชำระ"), PB("รายละเอียดการจัดส่ง")],
        [P("ยอดที่ชำระจริงเงินสด"), P(format_num(header_data.get('cash_actual_payment', 0))), "", ""], 
        [P("ยอดค่าสินค้าเงินสด"), P(format_num(header_data.get('cash_product_input', 0))), "", ""], 
        [P("ยอดรวมค่าบริการเงินสด"), P(format_num(header_data.get('cash_service_total', 0))), check_balance_text, ""], 
        [P("ยอดที่ต้องชำระเงินสด"), P(format_num(header_data.get('cash_required_total', 0))), "", ""], 
    ]
    
    cash_data[1][3] = delivery_note_inner_table 
    
    cash_table = Table(cash_data, colWidths=[3.2*cm, 1.8*cm, 1.5*cm, 3.0*cm], 
                       rowHeights=[0.6*cm, 0.75*cm, 0.75*cm, 0.75*cm, 0.75*cm]) 

    cash_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('BACKGROUND', (0,0), (2,0), colors.lightgrey),
        ('BACKGROUND', (0,1), (0,-1), colors.HexColor("#FFFFE0")), 
        ('ALIGN', (1,1), (1,-1), 'RIGHT'), ('ALIGN', (2,3), (2,3), 'CENTER'),
        ('SPAN', (3, 1), (3, -1)), ('VALIGN', (3, 0), (3, -1), 'TOP'),
        ('LEFTPADDING', (3,0), (3,-1), 0), ('RIGHTPADDING', (3,0), (3,-1), 0),
        ('TOPPADDING', (3,0), (3,-1), 0), ('BOTTOMPADDING', (3,0), (3,-1), 0),
    ]))
    story.append(cash_table)
    
    # 5. Remark Table
    remark_text = header_data.get('remark', '')
    remark_content = Paragraph(remark_text.replace('\n', '<br/>'), styles['Normal_TH'])
    
    remark_table = Table([[PB("หมายเหตุ (Remark)*")], [remark_content]], 
                          colWidths=[9.5*cm], rowHeights=[0.6*cm, 0.6*cm])
    
    remark_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('TOPPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(remark_table)
    
    return story

def _build_right_column(header_data, items_data, styles, P, PB, format_num):
    """
    สร้างคอลัมน์ด้านขวาของเอกสาร PDF (PURCHASE COST AUDITOR)
    """
    story = []

    # PO Info Table
    po_info_data = [
        [PB("A+SMART-PU"), PB("PURCHASE COST AUDITOR"), "", ""],
        [PB("PO/ST NUMBER"), PB("แผนก"), PB("RR Number"), PB("ผู้จัดทำ")],
        [P(header_data.get('po_number','')), P(header_data.get('department','')), P(header_data.get('rr_number','')), P(header_data.get('user_name',''))],
        [PB("Supplier Name"), P(header_data.get('supplier_name','')), "", P(header_data.get('credit_term',''))]
    ]
    po_info_table = Table(po_info_data, colWidths=[2.5*cm, 2.25*cm, 2.25*cm, 2.5*cm], rowHeights=0.7*cm)
    po_info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('SPAN', (1,0), (3,0)), ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E5E7EB")), 
        ('SPAN', (1,3), (2,3)), ('BACKGROUND', (0,1), (3,1), colors.HexColor("#F3F4F6")),
        ('BACKGROUND', (0,3), (0,3), colors.HexColor("#F3F4F6")), 
        ('BACKGROUND', (3,3), (3,3), colors.HexColor("#F3F4F6")),
    ]))
    story.append(po_info_table)
    story.append(Spacer(1, 0.2*cm))

    # Purchase Request Detail Table
    item_rows = [
        [PB("PURCHASE REQUEST DETAIL"), "", "", PB("PO/ST Type"), P(header_data.get('po_mode', '')), ""],
        [PB("No."), PB("สถานะ"), PB("Product Name"), PB("จำนวน"), PB("ราคาทุน"), PB("รวม")]
    ]
    for i, item in enumerate(items_data):
        row = [
            P(str(i+1)), 
            P(item.get('status', 'ST/TD')), 
            Paragraph(item.get('product_name', ''), styles['Normal_TH']),
            P(format_num(item.get('quantity', 0))), 
            P(format_num(item.get('unit_price', 0))),
            P(format_num(item.get('total_price', 0)))
        ]
        item_rows.append(row)
    
    item_col_widths = [0.8*cm, 1.5*cm, 3.0*cm, 1.4*cm, 1.4*cm, 1.4*cm]
    item_table = Table(item_rows, colWidths=item_col_widths)
    item_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), 
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), 
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('ALIGN', (3,1), (-1,-1), 'RIGHT'),
        ('SPAN', (0,0), (2,0)),
        ('SPAN', (4,0), (5,0)), 
        ('BACKGROUND', (0,0), (-1,1), colors.HexColor("#F3F4F6")), 
        ('ALIGN', (0,0), (-1,1), 'CENTER'),
    ]))
    story.append(item_table)
    story.append(Spacer(1, 0.2*cm))
    
    # Payment Summary Table
    payment_summary_data = [
        [PB("การชำระซัพพลายเออร์รวม VAT"), "", "", PB("ยอดรวมทั้งสิ้น"), ""],
        [P("มัดจำ"), P(format_num(header_data.get('deposit_amount', 0))), P("สถานะ: จ/คป"), P("หัก ณ ที่จ่าย 3%"), P(format_num(header_data.get('wht_3_percent_po', 0)))],
        [P("ยอดค้าง"), P(format_num(header_data.get('balance_due_po', 0))), P(str(header_data.get('deposit_date',''))), P("VAT 7%"), P(format_num(header_data.get('vat_7_percent_po', 0)))],
        [P("ชำระเต็ม"), P(format_num(header_data.get('full_payment_amount', 0))), P(str(header_data.get('full_payment_date',''))), P("ยอดต้นทุนรวม VAT"), P(format_num(header_data.get('grand_total_vat_po', 0)))],
        [P("CN/คืนส่วนต่าง"), P(format_num(header_data.get('cn_refund_amount', 0))), P(str(header_data.get('cn_refund_date',''))), P("ยอดต้นทุนชำระ"), P(format_num(header_data.get('net_payable_po', 0)))],
    ]
    payment_summary_table = Table(payment_summary_data, colWidths=[2*cm, 1.5*cm, 2*cm, 2.2*cm, 1.8*cm], rowHeights=0.6*cm)
    payment_summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('SPAN', (0,0), (2,0)), ('SPAN', (3,0), (4,0)),
        ('ALIGN', (0,0), (-1,0), 'CENTER'), ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F3F4F6")),
        ('ALIGN', (1,1), (1,-1), 'RIGHT'), ('ALIGN', (4,1), (4,-1), 'RIGHT'),
    ]))
    story.append(payment_summary_table)
    story.append(Spacer(1, 0.2*cm))

    # Shipping and Approval Table
    shipping_approval_data = [
        [PB("ค่าจัดส่ง"), "", "", "", PB("ผู้จัดทำและอนุมัติ"), ""],
        [P("ค่าจัดส่ง 1"), P(format_num(header_data.get('shipping_cost_1',0))), P("VAT/CASH"), P(header_data.get('shipping_vat_type_1','')), P("ผู้จัดทำ"), P(header_data.get('creator_user',''))],
        [P("ผู้จัดส่ง"), P(header_data.get('shipper_1', '')), P("หมายเหตุ"), "", P("ผู้อนุมัติ 1"), P(header_data.get('approver_1',''))],
        [P("ค่าจัดส่ง 2"), P(format_num(header_data.get('shipping_cost_2',0))), P("VAT/CASH"), P(header_data.get('shipping_vat_type_2','')), P("ผู้อนุมัติ 2"), P(header_data.get('approver_2',''))],
        [P("ผู้จัดส่ง"), P(header_data.get('shipper_2', '')), P("หมายเหตุ"), "", P("ผู้อนุมัติ 3"), P(header_data.get('approver_3',''))],
    ]
    shipping_approval_table = Table(shipping_approval_data, colWidths=[1.5*cm, 1.5*cm, 1.5*cm, 1*cm, 2*cm, 2*cm], rowHeights=0.6*cm)
    shipping_approval_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('SPAN', (0,0), (3,0)),
        ('SPAN', (4,0), (5,0)), ('SPAN', (2,2), (3,2)),
        ('SPAN', (2,4), (3,4)), ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F3F4F6")),
        ('ALIGN', (1,1), (1,4), 'RIGHT'),
        ('ALIGN', (1,2), (1,2), 'LEFT'),
        ('ALIGN', (1,4), (1,4), 'LEFT'),
    ]))
    story.append(shipping_approval_table)
    
    return story
    
def generate_multi_po_pdf(so_header_data, all_po_data):
    """สร้าง PDF ไฟล์เดียวที่มีหลายหน้าสำหรับทุก PO ที่เกี่ยวข้องกับ SO"""
    register_thai_fonts()
    
    documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
    if not os.path.exists(documents_path):
        documents_path = os.path.join(os.path.expanduser('~'), 'Desktop')

    default_filename = f"ALL_POs_for_SO_{so_header_data.get('so_number', '')}.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile=default_filename,
        initialdir=documents_path
    )
    if not save_path: return

    doc = BaseDocTemplate(save_path, pagesize=A4, leftMargin=1.0*cm, rightMargin=1.0*cm, topMargin=1.0*cm, bottomMargin=1.0*cm)
    
    frame_width = doc.width / 2.0 - (0.25 * cm)
    frame_height = doc.height
    left_frame = Frame(doc.leftMargin, doc.bottomMargin, frame_width, frame_height, id='left_col')
    right_frame = Frame(doc.leftMargin + frame_width + (0.5 * cm), doc.bottomMargin, frame_width, frame_height, id='right_col')
    two_column_template = PageTemplate(id='TwoColumn', frames=[left_frame, right_frame])
    doc.addPageTemplates([two_column_template])

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Normal_TH', fontName='THSarabunNew', fontSize=8, leading=11))
    styles.add(ParagraphStyle(name='Bold_TH', fontName='THSarabunNew-Bold', fontSize=9, leading=12))
    styles.add(ParagraphStyle(name='Wrapped_TH', fontName='THSarabunNew', fontSize=8, leading=9))

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
        is_first_po = True

        for po_data in all_po_data:
            if not is_first_po:
                story.append(PageBreak())

            left_content = _build_left_column(so_header_data, styles, P, PB, format_num)
            right_content = _build_right_column(po_data['header'], po_data['items'], styles, P, PB, format_num)
            
            story.extend(left_content)
            story.append(FrameBreak())
            story.extend(right_content)

            is_first_po = False
        
        doc.build(story)
        messagebox.showinfo("สำเร็จ", f"สร้างเอกสารรวมเรียบร้อย:\n{save_path}")
    
    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาดในการสร้าง PDF:\n{str(e)}")
        print(traceback.format_exc())