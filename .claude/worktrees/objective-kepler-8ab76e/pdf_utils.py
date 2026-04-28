# pdf_utils.py

import os
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4, landscape

def register_thai_fonts():
    """ลงทะเบียนฟอนต์ภาษาไทยสำหรับ ReportLab"""
    # ตรวจสอบว่าไฟล์ฟอนต์อยู่ในโฟลเดอร์ resources
    font_dir = 'resources'
    if not os.path.isdir(font_dir):
        print(f"Error: ไม่พบโฟลเดอร์ '{font_dir}'")
        return

    # สร้าง Dictionary ของฟอนต์จากไฟล์ที่ผู้ใช้ให้มา
    fonts = {
        'THSarabunNew': 'THSarabunNew.ttf',
        'THSarabunNew-Bold': 'THSarabunNew Bold.ttf',
        'THSarabunNew-Italic': 'THSarabunNew Italic.ttf',
        'THSarabunNew-BoldItalic': 'THSarabunNew BoldItalic.ttf'
    }

    # ลงทะเบียนฟอนต์แต่ละตัว
    for name, filename in fonts.items():
        font_path = os.path.join(font_dir, filename)
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont(name, font_path))
        else:
            print(f"Warning: ไม่พบไฟล์ฟอนต์ '{filename}' ที่ '{font_dir}'")
    
    # ลงทะเบียนตระกูลฟอนต์
    pdfmetrics.registerFontFamily(
        'THSarabunNew',
        normal='THSarabunNew',
        bold='THSarabunNew-Bold',
        italic='THSarabunNew-Italic',
        boldItalic='THSarabunNew-BoldItalic'
    )

def export_approved_pos_to_pdf(parent_window, pg_engine):
    """Export ข้อมูล PO ที่อนุมัติแล้วไปยังไฟล์ PDF"""
    
    # 1. ลงทะเบียนฟอนต์
    register_thai_fonts()

    # 2. ถามที่จัดเก็บไฟล์
    default_filename = f"Approved_PO_Report_{datetime.now().strftime('%Y%m%d')}.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="บันทึกรายงาน PDF",
        initialfile=default_filename,
        parent=parent_window
    )
    if not save_path:
        return

    try:
        # 3. ดึงข้อมูลจากฐานข้อมูล (ใช้ Query เดียวกับ Excel)
        query = """
            SELECT so.so_number, po.po_number, po.po_mode, po.rr_number, po.timestamp AS po_date,
                   po.supplier_name, po.grand_total AS po_total_payable, po.user_key AS po_creator_key,
                   so.bill_date AS so_date, so.customer_name, u.sale_name,
                   (COALESCE(so.sales_service_amount, 0) + COALESCE(so.product_vat_7, 0)) AS so_grand_total
            FROM purchase_orders po
            LEFT JOIN commissions so ON po.so_number = so.so_number
            LEFT JOIN sales_users u ON so.sale_key = u.sale_key
            WHERE po.status = 'Approved' ORDER BY po.timestamp DESC;
        """
        df = pd.read_sql_query(query, pg_engine)

        if df.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูล PO ที่อนุมัติแล้วสำหรับสร้างรายงาน PDF", parent=parent_window)
            return

        # 4. เตรียมข้อมูลสำหรับตารางใน PDF
        header_map = {
            'so_number': 'เลขที่ SO', 'po_number': 'เลขที่ PO', 'po_mode': 'ประเภท PO', 'rr_number': 'เลขที่ RR',
            'po_date': 'วันที่สร้าง PO', 'supplier_name': 'ชื่อซัพพลายเออร์', 'po_total_payable': 'ยอดชำระ PO',
            'user_key': 'ผู้สร้าง', 'so_date': 'วันที่เปิด SO', 'customer_name': 'ชื่อลูกค้า',
            'sale_name': 'พนักงานขาย', 'so_grand_total': 'ยอดขาย SO'
        }
        
        df.rename(columns=header_map, inplace=True)
        
        # แปลงข้อมูลวันที่ให้อยู่ในรูปแบบที่อ่านง่าย
        for date_col in ['วันที่สร้าง PO', 'วันที่เปิด SO']:
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col]).dt.strftime('%Y-%m-%d')

        # แปลง DataFrame เป็น List of Lists
        data_for_table = [df.columns.tolist()] + df.values.tolist()

        # 5. สร้างเอกสาร PDF
        doc = SimpleDocTemplate(save_path, pagesize=landscape(A4), topMargin=0.5*inch, bottomMargin=0.5*inch)
        story = []
        styles = getSampleStyleSheet()

        # 5.1 เพิ่มโลโก้และหัวข้อ
        logo_path = 'resources/company_logo.png'
        if os.path.exists(logo_path):
            im = Image(logo_path, width=0.8*inch, height=0.8*inch)
            im.hAlign = 'LEFT'
            story.append(im)

        title_style = ParagraphStyle(name='Title', parent=styles['h1'], fontName='THSarabunNew-Bold', fontSize=20, alignment=1)
        title = Paragraph("รายงานใบสั่งซื้อที่อนุมัติแล้ว", title_style)
        story.append(title)
        story.append(Spacer(1, 0.25 * inch))

        # 5.2 สร้างตารางและกำหนดสไตล์
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F46E5")), # สีหัวตาราง
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'THSarabunNew-Bold'), # ฟอนต์หัวตาราง
            ('FONTNAME', (0, 1), (-1, -1), 'THSarabunNew'),    # ฟอนต์เนื้อหา
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
        
        # สร้าง Table object
        table = Table(data_for_table, hAlign='CENTER')
        table.setStyle(table_style)
        story.append(table)

        # 6. สร้างไฟล์ PDF
        doc.build(story)
        messagebox.showinfo("สำเร็จ", f"สร้างไฟล์ PDF เรียบร้อยแล้วที่:\n{save_path}", parent=parent_window)

    except Exception as e:
        messagebox.showerror("ผิดพลาด", f"ไม่สามารถสร้างไฟล์ PDF ได้: {e}", parent=parent_window)