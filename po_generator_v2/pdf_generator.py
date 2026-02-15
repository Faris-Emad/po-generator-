# -*- coding: utf-8 -*-
"""
مولد ملفات PDF مع دعم العربية
PDF Generator with Arabic Support
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from arabic_reshaper import reshape
from bidi.algorithm import get_display
import io


def prepare_arabic_text(text):
    """
    تحضير النص العربي للعرض الصحيح في PDF
    Prepare Arabic text for correct display in PDF
    """
    if not text:
        return ''
    reshaped_text = reshape(str(text))
    bidi_text = get_display(reshaped_text)
    return bidi_text


def create_pdf_from_order(order_data):
    """
    إنشاء ملف PDF من بيانات الطلب
    Create PDF from order data
    """
    buffer = io.BytesIO()
    
    # إعداد المستند
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )
    
    # العناصر التي سيتم إضافتها للـ PDF
    elements = []
    
    # الأنماط
    styles = getSampleStyleSheet()
    
    # نمط العنوان الرئيسي
    title_style = ParagraphStyle(
        'ArabicTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        textColor=colors.HexColor('#2B3E50'),
        alignment=TA_CENTER,
        spaceAfter=10
    )
    
    # نمط العنوان الفرعي
    subtitle_style = ParagraphStyle(
        'ArabicSubtitle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        textColor=colors.HexColor('#2B3E50'),
        alignment=TA_CENTER,
        spaceAfter=15
    )
    
    # نمط النص العادي
    normal_style = ParagraphStyle(
        'ArabicNormal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        alignment=TA_RIGHT,
        spaceAfter=8
    )
    
    # === رأس الشركة ===
    company_name = prepare_arabic_text('شركة الأمانة للتوريدات العامة')
    elements.append(Paragraph(company_name, title_style))
    
    po_title = prepare_arabic_text('طلب توريد - Purchase Order')
    elements.append(Paragraph(po_title, subtitle_style))
    elements.append(Spacer(1, 10*mm))
    
    # === بيانات الطلب الأساسية ===
    basic_data = [
        [prepare_arabic_text('رقم الطلب:'), order_data.get('po_number', ''), 
         prepare_arabic_text('التاريخ:'), order_data.get('po_date', '')],
        [prepare_arabic_text('الرقم الضريبي:'), order_data.get('company_tax_id', ''), 
         prepare_arabic_text('السجل التجاري:'), order_data.get('commercial_reg', '')]
    ]
    
    basic_table = Table(basic_data, colWidths=[40*mm, 45*mm, 40*mm, 45*mm])
    basic_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E8F0F8')),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E8F0F8')),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    elements.append(basic_table)
    elements.append(Spacer(1, 8*mm))
    
    # === بيانات المورد ===
    supplier = order_data.get('supplier', {})
    
    supplier_header = prepare_arabic_text('بيانات المورد - Supplier Information')
    elements.append(Paragraph(supplier_header, subtitle_style))
    
    supplier_data = [
        [prepare_arabic_text('اسم المورد:'), prepare_arabic_text(supplier.get('name', '')), 
         prepare_arabic_text('الرقم الضريبي:'), supplier.get('tax_id', '')],
        [prepare_arabic_text('التليفون:'), supplier.get('phone', ''), 
         prepare_arabic_text('البريد الإلكتروني:'), supplier.get('email', '')],
        [prepare_arabic_text('العنوان:'), prepare_arabic_text(supplier.get('address', '')), '', '']
    ]
    
    supplier_table = Table(supplier_data, colWidths=[40*mm, 45*mm, 40*mm, 45*mm])
    supplier_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F2F2F2')),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, 1), 'Helvetica-Bold'),
        ('SPAN', (1, 2), (3, 2)),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    elements.append(supplier_table)
    elements.append(Spacer(1, 8*mm))
    
    # === جدول الأصناف ===
    items_header = prepare_arabic_text('أصناف الطلب - Order Items')
    elements.append(Paragraph(items_header, subtitle_style))
    
    # رأس الجدول
    items_data = [[
        prepare_arabic_text('م'),
        prepare_arabic_text('كود الصنف'),
        prepare_arabic_text('اسم الصنف'),
        prepare_arabic_text('الكمية'),
        prepare_arabic_text('السعر'),
        prepare_arabic_text('الإجمالي')
    ]]
    
    # بيانات الأصناف
    items = order_data.get('items', [])
    subtotal = 0.0
    
    for idx, item in enumerate(items, start=1):
        item_total = item.get('total_price', 0)
        subtotal += item_total
        
        items_data.append([
            str(idx),
            item.get('product_code', ''),
            prepare_arabic_text(item.get('product_name', '')),
            f"{item.get('quantity', 0):.2f}",
            f"{item.get('unit_price', 0):.2f}",
            f"{item_total:.2f}"
        ])
    
    items_table = Table(items_data, colWidths=[15*mm, 25*mm, 60*mm, 20*mm, 25*mm, 25*mm])
    items_table.setStyle(TableStyle([
        # رأس الجدول
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B3E50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        # البيانات
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # رقم الصنف
        ('ALIGN', (1, 1), (1, -1), 'CENTER'),  # الكود
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),   # الاسم
        ('ALIGN', (3, 1), (-1, -1), 'CENTER'), # الأرقام
        # التنسيق العام
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        # صفوف متبادلة
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F2F2')])
    ]))
    elements.append(items_table)
    elements.append(Spacer(1, 5*mm))
    
    # === الإجماليات ===
    tax_amount = subtotal * 0.14
    final_total = subtotal + tax_amount
    
    totals_data = [
        [prepare_arabic_text('المجموع الفرعي (قبل الضريبة):'), f"{subtotal:,.2f} ج.م"],
        [prepare_arabic_text('ضريبة القيمة المضافة (14%):'), f"{tax_amount:,.2f} ج.م"],
        [prepare_arabic_text('الإجمالي النهائي (شامل الضريبة):'), f"{final_total:,.2f} ج.م"]
    ]
    
    totals_table = Table(totals_data, colWidths=[120*mm, 50*mm])
    totals_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 1), colors.HexColor('#D9E2F3')),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#2B3E50')),
        ('TEXTCOLOR', (0, 0), (-1, 1), colors.black),
        ('TEXTCOLOR', (0, 2), (-1, 2), colors.white),
        ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 10),
        ('FONTSIZE', (0, 2), (-1, 2), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey)
    ]))
    elements.append(totals_table)
    
    # === شروط التوريد ===
    if order_data.get('delivery_period') or order_data.get('delivery_location') or order_data.get('payment_terms'):
        elements.append(Spacer(1, 8*mm))
        terms_header = prepare_arabic_text('شروط التوريد - Delivery Terms')
        elements.append(Paragraph(terms_header, subtitle_style))
        
        terms_data = []
        if order_data.get('delivery_period'):
            terms_data.append([prepare_arabic_text('مدة التوريد:'), 
                             prepare_arabic_text(order_data.get('delivery_period', ''))])
        
        if order_data.get('delivery_location'):
            terms_data.append([prepare_arabic_text('مكان التسليم:'), 
                             prepare_arabic_text(order_data.get('delivery_location', ''))])
        
        if order_data.get('payment_terms'):
            terms_data.append([prepare_arabic_text('شروط الدفع:'), 
                             prepare_arabic_text(order_data.get('payment_terms', ''))])
        
        if terms_data:
            terms_table = Table(terms_data, colWidths=[40*mm, 130*mm])
            terms_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F2F2F2')),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
            ]))
            elements.append(terms_table)
    
    # === ملاحظات ===
    if order_data.get('notes'):
        elements.append(Spacer(1, 8*mm))
        notes_header = prepare_arabic_text('ملاحظات - Notes')
        elements.append(Paragraph(notes_header, subtitle_style))
        
        notes_data = [[prepare_arabic_text(order_data.get('notes', ''))]]
        notes_table = Table(notes_data, colWidths=[170*mm])
        notes_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F9F9F9')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        elements.append(notes_table)
    
    # === التوقيعات ===
    elements.append(Spacer(1, 15*mm))
    signatures_data = [[
        prepare_arabic_text('توقيع المورد\n___________________'),
        prepare_arabic_text('توقيع الشركة\n___________________')
    ]]
    
    signatures_table = Table(signatures_data, colWidths=[85*mm, 85*mm])
    signatures_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 10)
    ]))
    elements.append(signatures_table)
    
    # بناء المستند
    doc.build(elements)
    
    buffer.seek(0)
    return buffer
