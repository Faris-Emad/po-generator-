# -*- coding: utf-8 -*-
"""
موديلات قاعدة البيانات لنظام طلبات التوريد
Database models for Purchase Order System
"""

from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


class Supplier(db.Model):
    """موديل الموردين - Suppliers Model"""
    __tablename__ = 'suppliers'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    tax_id = db.Column(db.String(100))
    phone = db.Column(db.String(50))
    email = db.Column(db.String(100))
    address = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # علاقات - Relationships
    orders = db.relationship('Order', backref='supplier', lazy=True)
    
    def to_dict(self):
        """تحويل الموديل إلى قاموس - Convert model to dictionary"""
        return {
            'id': self.id,
            'name': self.name,
            'tax_id': self.tax_id,
            'phone': self.phone,
            'email': self.email,
            'address': self.address
        }


class Product(db.Model):
    """موديل الأصناف - Products Model"""
    __tablename__ = 'products'
    
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(100), unique=True, nullable=False)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    default_price = db.Column(db.Float, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def to_dict(self):
        """تحويل الموديل إلى قاموس - Convert model to dictionary"""
        return {
            'id': self.id,
            'code': self.code,
            'name': self.name,
            'description': self.description,
            'default_price': self.default_price
        }


class Order(db.Model):
    """موديل الطلبات - Orders Model"""
    __tablename__ = 'orders'
    
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(50), unique=True, nullable=False)
    po_date = db.Column(db.Date, nullable=False)
    
    # بيانات الشركة - Company Information
    company_tax_id = db.Column(db.String(100))
    company_commercial_reg = db.Column(db.String(100))
    
    # بيانات المورد - Supplier Information
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'))
    
    # شروط التوريد - Delivery Terms
    delivery_period = db.Column(db.String(200))
    delivery_location = db.Column(db.Text)
    payment_terms = db.Column(db.Text)
    notes = db.Column(db.Text)
    
    # الإجماليات - Totals
    subtotal = db.Column(db.Float, default=0.0)
    tax_amount = db.Column(db.Float, default=0.0)
    total = db.Column(db.Float, default=0.0)
    
    # حالة الطلب - Order Status
    status = db.Column(db.String(50), default='مسودة')  # مسودة، مؤكد، ملغي
    
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # علاقات - Relationships
    items = db.relationship('OrderItem', backref='order', lazy=True, cascade='all, delete-orphan')
    
    def to_dict(self):
        """تحويل الموديل إلى قاموس - Convert model to dictionary"""
        return {
            'id': self.id,
            'po_number': self.po_number,
            'po_date': self.po_date.strftime('%Y-%m-%d'),
            'company_tax_id': self.company_tax_id,
            'company_commercial_reg': self.company_commercial_reg,
            'supplier': self.supplier.to_dict() if self.supplier else None,
            'delivery_period': self.delivery_period,
            'delivery_location': self.delivery_location,
            'payment_terms': self.payment_terms,
            'notes': self.notes,
            'subtotal': self.subtotal,
            'tax_amount': self.tax_amount,
            'total': self.total,
            'status': self.status,
            'items': [item.to_dict() for item in self.items]
        }


class OrderItem(db.Model):
    """موديل أصناف الطلب - Order Items Model"""
    __tablename__ = 'order_items'
    
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    
    # بيانات الصنف - Item Information
    product_code = db.Column(db.String(100))
    product_name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    quantity = db.Column(db.Float, nullable=False)
    unit_price = db.Column(db.Float, nullable=False)
    total_price = db.Column(db.Float, nullable=False)
    
    item_order = db.Column(db.Integer)  # ترتيب الصنف في الطلب
    
    def to_dict(self):
        """تحويل الموديل إلى قاموس - Convert model to dictionary"""
        return {
            'id': self.id,
            'product_code': self.product_code,
            'product_name': self.product_name,
            'description': self.description,
            'quantity': self.quantity,
            'unit_price': self.unit_price,
            'total_price': self.total_price
        }
