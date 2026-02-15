# -*- coding: utf-8 -*-
"""
إعداد قاعدة البيانات SQLite
Database setup for SQLite
"""

from models import db, Supplier, Product, Order, OrderItem
from datetime import datetime


def init_database(app):
    """تهيئة قاعدة البيانات - Initialize database"""
    with app.app_context():
        db.create_all()
        # إضافة بيانات تجريبية إذا كانت قاعدة البيانات فارغة
        add_sample_data()


def add_sample_data():
    """إضافة بيانات تجريبية - Add sample data"""
    
    # التحقق من وجود موردين
    if Supplier.query.count() == 0:
        # إضافة موردين تجريبيين
        suppliers = [
            Supplier(
                name='شركة النور للتوريدات',
                tax_id='123-456-789',
                phone='01012345678',
                email='info@alnoor.com',
                address='15 شارع الجمهورية، القاهرة'
            ),
            Supplier(
                name='مؤسسة الأمل التجارية',
                tax_id='987-654-321',
                phone='01098765432',
                email='contact@alamal.com',
                address='22 شارع النيل، الجيزة'
            ),
            Supplier(
                name='الشركة المصرية للمستلزمات',
                tax_id='456-789-123',
                phone='01123456789',
                email='sales@egy-supplies.com',
                address='10 ميدان التحرير، القاهرة'
            )
        ]
        for supplier in suppliers:
            db.session.add(supplier)
        db.session.commit()
    
    # التحقق من وجود أصناف
    if Product.query.count() == 0:
        # إضافة أصناف تجريبية
        products = [
            Product(
                code='A001',
                name='ورق A4 - 80 جرام',
                description='رزمة ورق طباعة A4، وزن 80 جرام، 500 ورقة',
                default_price=150.00
            ),
            Product(
                code='A002',
                name='أقلام حبر جاف أزرق',
                description='علبة 50 قلم حبر جاف، لون أزرق',
                default_price=75.00
            ),
            Product(
                code='A003',
                name='دباسة معدنية كبيرة',
                description='دباسة معدنية، سعة 50 ورقة',
                default_price=45.00
            ),
            Product(
                code='B001',
                name='طابعة ليزر HP LaserJet',
                description='طابعة ليزر أحادية اللون، سرعة 30 صفحة/دقيقة',
                default_price=4500.00
            ),
            Product(
                code='B002',
                name='حاسب محمول Dell',
                description='حاسب محمول Dell، معالج i5، رام 8GB، هارد 512 SSD',
                default_price=18000.00
            ),
            Product(
                code='C001',
                name='خزانة ملفات معدنية',
                description='خزانة ملفات معدنية 4 أدراج، ارتفاع 120 سم',
                default_price=2500.00
            ),
            Product(
                code='C002',
                name='مكتب خشبي تنفيذي',
                description='مكتب خشبي تنفيذي، مقاس 160x80 سم',
                default_price=3500.00
            ),
            Product(
                code='D001',
                name='حبر طابعة HP أسود',
                description='خرطوشة حبر أصلي HP، لون أسود',
                default_price=650.00
            ),
            Product(
                code='D002',
                name='مسطرة معدنية 50 سم',
                description='مسطرة معدنية مدرجة، طول 50 سم',
                default_price=25.00
            ),
            Product(
                code='E001',
                name='دفتر سجل 200 ورقة',
                description='دفتر سجل مسطر، 200 ورقة، غلاف سميك',
                default_price=35.00
            )
        ]
        for product in products:
            db.session.add(product)
        db.session.commit()


def get_next_po_number():
    """الحصول على رقم الطلب التالي - Get next PO number"""
    current_year = datetime.now().year
    
    # البحث عن آخر طلب في السنة الحالية
    last_order = Order.query.filter(
        Order.po_number.like(f'PO-{current_year}-%')
    ).order_by(Order.po_number.desc()).first()
    
    if last_order:
        # استخراج الرقم التسلسلي من آخر طلب
        try:
            last_number = int(last_order.po_number.split('-')[-1])
            next_number = last_number + 1
        except:
            next_number = 1
    else:
        next_number = 1
    
    # تنسيق الرقم: PO-2026-001
    return f'PO-{current_year}-{next_number:03d}'
