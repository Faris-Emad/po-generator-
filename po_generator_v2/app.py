# -*- coding: utf-8 -*-
"""
التطبيق الرئيسي لنظام طلبات التوريد v2
Main Flask Application for Purchase Order System v2
"""

from flask import Flask, render_template, request, jsonify, send_file
from models import db, Supplier, Product, Order, OrderItem
from database import init_database, get_next_po_number
from excel_generator import create_professional_excel
from pdf_generator import create_pdf_from_order
from datetime import datetime
import io
import os

app = Flask(__name__)

# إعدادات التطبيق - Application Configuration
app.config['SECRET_KEY'] = 'your-secret-key-here-change-in-production'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///po_system.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# تهيئة قاعدة البيانات - Initialize Database
db.init_app(app)
with app.app_context():
    init_database(app)


@app.route('/')
def index():
    """الصفحة الرئيسية - Main page for creating new orders"""
    next_po_number = get_next_po_number()
    return render_template('index.html', next_po_number=next_po_number)


@app.route('/orders')
def orders():
    """صفحة الطلبات السابقة - Previous orders page"""
    return render_template('orders.html')


# === API للموردين - Suppliers API ===

@app.route('/api/suppliers', methods=['GET'])
def get_suppliers():
    """الحصول على جميع الموردين - Get all suppliers"""
    suppliers = Supplier.query.order_by(Supplier.name).all()
    return jsonify([s.to_dict() for s in suppliers])


@app.route('/api/suppliers', methods=['POST'])
def create_supplier():
    """إضافة مورد جديد - Create new supplier"""
    data = request.json
    
    supplier = Supplier(
        name=data.get('name'),
        tax_id=data.get('tax_id'),
        phone=data.get('phone'),
        email=data.get('email'),
        address=data.get('address')
    )
    
    db.session.add(supplier)
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'تم إضافة المورد بنجاح',
        'supplier': supplier.to_dict()
    })


@app.route('/api/suppliers/<int:supplier_id>', methods=['GET'])
def get_supplier(supplier_id):
    """الحصول على بيانات مورد معين - Get specific supplier"""
    supplier = Supplier.query.get_or_404(supplier_id)
    return jsonify(supplier.to_dict())


# === API للأصناف - Products API ===

@app.route('/api/products', methods=['GET'])
def get_products():
    """الحصول على جميع الأصناف - Get all products"""
    products = Product.query.order_by(Product.name).all()
    return jsonify([p.to_dict() for p in products])


@app.route('/api/products', methods=['POST'])
def create_product():
    """إضافة صنف جديد - Create new product"""
    data = request.json
    
    # التحقق من عدم وجود نفس الكود
    existing = Product.query.filter_by(code=data.get('code')).first()
    if existing:
        return jsonify({
            'success': False,
            'message': 'هذا الكود موجود بالفعل'
        }), 400
    
    product = Product(
        code=data.get('code'),
        name=data.get('name'),
        description=data.get('description'),
        default_price=float(data.get('default_price', 0))
    )
    
    db.session.add(product)
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'تم إضافة الصنف بنجاح',
        'product': product.to_dict()
    })


@app.route('/api/products/<int:product_id>', methods=['GET'])
def get_product(product_id):
    """الحصول على بيانات صنف معين - Get specific product"""
    product = Product.query.get_or_404(product_id)
    return jsonify(product.to_dict())


# === API للطلبات - Orders API ===

@app.route('/api/orders', methods=['GET'])
def get_orders():
    """الحصول على جميع الطلبات - Get all orders"""
    # الفلترة حسب المعايير المطلوبة
    search = request.args.get('search', '').strip()
    status = request.args.get('status', '').strip()
    
    query = Order.query
    
    if search:
        query = query.filter(
            (Order.po_number.contains(search)) |
            (Order.supplier.has(Supplier.name.contains(search)))
        )
    
    if status:
        query = query.filter_by(status=status)
    
    orders = query.order_by(Order.created_at.desc()).all()
    return jsonify([o.to_dict() for o in orders])


@app.route('/api/orders', methods=['POST'])
def create_order():
    """إنشاء طلب جديد - Create new order"""
    data = request.json
    
    try:
        # إنشاء الطلب
        order = Order(
            po_number=data.get('po_number'),
            po_date=datetime.strptime(data.get('po_date'), '%Y-%m-%d').date(),
            company_tax_id=data.get('company_tax_id'),
            company_commercial_reg=data.get('commercial_reg'),
            supplier_id=data.get('supplier_id'),
            delivery_period=data.get('delivery_period'),
            delivery_location=data.get('delivery_location'),
            payment_terms=data.get('payment_terms'),
            notes=data.get('notes'),
            status='مؤكد'
        )
        
        # إضافة الأصناف
        subtotal = 0.0
        for idx, item_data in enumerate(data.get('items', []), start=1):
            quantity = float(item_data.get('quantity', 0))
            unit_price = float(item_data.get('unit_price', 0))
            total_price = quantity * unit_price
            subtotal += total_price
            
            item = OrderItem(
                product_code=item_data.get('product_code'),
                product_name=item_data.get('product_name'),
                description=item_data.get('description'),
                quantity=quantity,
                unit_price=unit_price,
                total_price=total_price,
                item_order=idx
            )
            order.items.append(item)
        
        # حساب الإجماليات
        tax_amount = subtotal * 0.14
        order.subtotal = subtotal
        order.tax_amount = tax_amount
        order.total = subtotal + tax_amount
        
        db.session.add(order)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'تم إنشاء الطلب بنجاح',
            'order': order.to_dict()
        })
    
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'message': f'حدث خطأ: {str(e)}'
        }), 500


@app.route('/api/orders/<int:order_id>', methods=['GET'])
def get_order(order_id):
    """الحصول على بيانات طلب معين - Get specific order"""
    order = Order.query.get_or_404(order_id)
    return jsonify(order.to_dict())


@app.route('/api/orders/<int:order_id>', methods=['DELETE'])
def delete_order(order_id):
    """حذف طلب - Delete order"""
    order = Order.query.get_or_404(order_id)
    db.session.delete(order)
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'تم حذف الطلب بنجاح'
    })


# === API لتوليد الملفات - File Generation API ===

@app.route('/api/orders/<int:order_id>/excel', methods=['GET'])
def generate_excel(order_id):
    """توليد ملف Excel للطلب - Generate Excel for order"""
    order = Order.query.get_or_404(order_id)
    order_dict = order.to_dict()
    
    wb = create_professional_excel(order_dict)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f'طلب_توريد_{order.po_number}.xlsx'
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/orders/<int:order_id>/pdf', methods=['GET'])
def generate_pdf(order_id):
    """توليد ملف PDF للطلب - Generate PDF for order"""
    order = Order.query.get_or_404(order_id)
    order_dict = order.to_dict()
    
    pdf_buffer = create_pdf_from_order(order_dict)
    
    filename = f'طلب_توريد_{order.po_number}.pdf'
    
    return send_file(
        pdf_buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/next-po-number', methods=['GET'])
def next_po_number():
    """الحصول على رقم الطلب التالي - Get next PO number"""
    return jsonify({
        'po_number': get_next_po_number()
    })


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
