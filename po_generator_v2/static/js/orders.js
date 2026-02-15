/**
 * صفحة الطلبات السابقة
 * Previous Orders Page
 */

let allOrders = [];

// === تحميل الطلبات عند فتح الصفحة ===
document.addEventListener('DOMContentLoaded', function() {
    loadOrders();
    
    // إضافة مستمعين للبحث والفلترة
    document.getElementById('searchInput').addEventListener('input', filterOrders);
    document.getElementById('statusFilter').addEventListener('change', filterOrders);
});

// === تحميل جميع الطلبات ===
async function loadOrders() {
    try {
        const search = document.getElementById('searchInput').value;
        const status = document.getElementById('statusFilter').value;
        
        let url = '/api/orders';
        const params = new URLSearchParams();
        if (search) params.append('search', search);
        if (status) params.append('status', status);
        
        if (params.toString()) {
            url += '?' + params.toString();
        }
        
        allOrders = await AppHelpers.apiRequest(url);
        displayOrders(allOrders);
    } catch (error) {
        AppHelpers.showToast('خطأ في تحميل الطلبات', 'error');
        document.getElementById('ordersTableBody').innerHTML = `
            <tr>
                <td colspan="6" class="text-center text-danger py-4">
                    <i class="fas fa-exclamation-circle fa-3x mb-3"></i>
                    <p>حدث خطأ في تحميل الطلبات</p>
                </td>
            </tr>
        `;
    }
}

// === عرض الطلبات في الجدول ===
function displayOrders(orders) {
    const tbody = document.getElementById('ordersTableBody');
    const noOrdersMsg = document.getElementById('noOrdersMessage');
    
    if (orders.length === 0) {
        tbody.innerHTML = '';
        noOrdersMsg.classList.remove('d-none');
        return;
    }
    
    noOrdersMsg.classList.add('d-none');
    
    tbody.innerHTML = orders.map(order => `
        <tr>
            <td class="text-center fw-bold">${order.po_number}</td>
            <td class="text-center">${AppHelpers.formatDate(order.po_date)}</td>
            <td>${order.supplier ? order.supplier.name : '-'}</td>
            <td class="text-center text-currency fw-bold">${AppHelpers.formatCurrency(order.total)}</td>
            <td class="text-center">
                <span class="badge-status ${order.status}">${order.status}</span>
            </td>
            <td class="text-center">
                <div class="btn-group" role="group">
                    <button class="btn btn-primary btn-sm" onclick="downloadExcel(${order.id})" title="تحميل Excel">
                        <i class="fas fa-file-excel me-1"></i>
                        Excel
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="downloadPDF(${order.id})" title="تحميل PDF">
                        <i class="fas fa-file-pdf me-1"></i>
                        PDF
                    </button>
                    <button class="btn btn-info btn-sm" onclick="printOrder(${order.id})" title="طباعة">
                        <i class="fas fa-print me-1"></i>
                        طباعة
                    </button>
                    <button class="btn btn-secondary btn-sm" onclick="viewOrderDetails(${order.id})" title="التفاصيل">
                        <i class="fas fa-eye me-1"></i>
                        عرض
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteOrder(${order.id})" title="حذف">
                        <i class="fas fa-trash me-1"></i>
                        حذف
                    </button>
                </div>
            </td>
        </tr>
    `).join('');
}

// === فلترة الطلبات ===
function filterOrders() {
    loadOrders();
}

// === تحميل Excel ===
async function downloadExcel(orderId) {
    try {
        AppHelpers.showLoading('جاري إنشاء ملف Excel...');
        
        const response = await fetch(`/api/orders/${orderId}/excel`);
        if (!response.ok) throw new Error('فشل في تحميل الملف');
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // الحصول على اسم الملف من الـ header
        const contentDisposition = response.headers.get('content-disposition');
        let filename = `order_${orderId}.xlsx`;
        if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename\*?=['"]?(?:UTF-\d['"]*)?([^;\r\n"']*)['"]?;?/);
            if (filenameMatch && filenameMatch[1]) {
                filename = decodeURIComponent(filenameMatch[1]);
            }
        }
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        AppHelpers.hideLoading();
        AppHelpers.showToast('تم تحميل ملف Excel بنجاح', 'success');
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast('فشل في تحميل ملف Excel', 'error');
    }
}

// === تحميل PDF ===
async function downloadPDF(orderId) {
    try {
        AppHelpers.showLoading('جاري إنشاء ملف PDF...');
        
        const response = await fetch(`/api/orders/${orderId}/pdf`);
        if (!response.ok) throw new Error('فشل في تحميل الملف');
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // الحصول على اسم الملف من الـ header
        const contentDisposition = response.headers.get('content-disposition');
        let filename = `order_${orderId}.pdf`;
        if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename\*?=['"]?(?:UTF-\d['"]*)?([^;\r\n"']*)['"]?;?/);
            if (filenameMatch && filenameMatch[1]) {
                filename = decodeURIComponent(filenameMatch[1]);
            }
        }
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        AppHelpers.hideLoading();
        AppHelpers.showToast('تم تحميل ملف PDF بنجاح', 'success');
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast('فشل في تحميل ملف PDF', 'error');
    }
}

// === طباعة الطلب ===
async function printOrder(orderId) {
    try {
        AppHelpers.showLoading('جاري تحضير الطلب للطباعة...');
        
        const order = await AppHelpers.apiRequest(`/api/orders/${orderId}`);
        
        // إنشاء محتوى الطباعة
        const printContent = createPrintContent(order);
        
        // وضع المحتوى في منطقة الطباعة
        const printArea = document.getElementById('printArea');
        printArea.innerHTML = printContent;
        printArea.style.display = 'block';
        
        AppHelpers.hideLoading();
        
        // طباعة
        window.print();
        
        // إخفاء منطقة الطباعة
        printArea.style.display = 'none';
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast('فشل في طباعة الطلب', 'error');
    }
}

// === إنشاء محتوى الطباعة ===
function createPrintContent(order) {
    let itemsHtml = '';
    let subtotal = 0;
    
    order.items.forEach((item, index) => {
        subtotal += item.total_price;
        itemsHtml += `
            <tr>
                <td>${index + 1}</td>
                <td>${item.product_code || '-'}</td>
                <td>${item.product_name}</td>
                <td>${item.quantity}</td>
                <td>${item.unit_price.toFixed(2)}</td>
                <td>${item.total_price.toFixed(2)}</td>
            </tr>
        `;
    });
    
    const taxRate = order.tax_rate || 14;
    const tax = subtotal * (taxRate / 100);
    const total = subtotal + tax;
    
    return `
        <div class="print-title">
            شركة الأمانة للتوريدات العامة<br>
            طلب توريد - Purchase Order
        </div>
        
        <div class="print-section">
            <h4>بيانات الطلب</h4>
            <table class="print-table">
                <tr>
                    <th>رقم الطلب</th>
                    <td>${order.po_number}</td>
                    <th>التاريخ</th>
                    <td>${order.po_date}</td>
                </tr>
                <tr>
                    <th>الرقم الضريبي</th>
                    <td>${order.company_tax_id || '-'}</td>
                    <th>السجل التجاري</th>
                    <td>${order.company_commercial_reg || '-'}</td>
                </tr>
            </table>
        </div>
        
        <div class="print-section">
            <h4>بيانات المورد</h4>
            <table class="print-table">
                <tr>
                    <th>اسم المورد</th>
                    <td>${order.supplier.name}</td>
                    <th>الرقم الضريبي</th>
                    <td>${order.supplier.tax_id || '-'}</td>
                </tr>
                <tr>
                    <th>التليفون</th>
                    <td>${order.supplier.phone || '-'}</td>
                    <th>البريد الإلكتروني</th>
                    <td>${order.supplier.email || '-'}</td>
                </tr>
                <tr>
                    <th>العنوان</th>
                    <td colspan="3">${order.supplier.address || '-'}</td>
                </tr>
            </table>
        </div>
        
        <div class="print-section">
            <h4>أصناف الطلب</h4>
            <table class="print-table">
                <thead>
                    <tr>
                        <th>م</th>
                        <th>الكود</th>
                        <th>الاسم</th>
                        <th>الكمية</th>
                        <th>السعر</th>
                        <th>الإجمالي</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsHtml}
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="5" style="text-align: left; font-weight: bold;">المجموع الفرعي:</td>
                        <td style="font-weight: bold;">${subtotal.toFixed(2)} ج.م</td>
                    </tr>
                    <tr>
                        <td colspan="5" style="text-align: left; font-weight: bold;">ضريبة القيمة المضافة (${taxRate}%):</td>
                        <td style="font-weight: bold;">${tax.toFixed(2)} ج.م</td>
                    </tr>
                    <tr>
                        <td colspan="5" style="text-align: left; font-weight: bold; font-size: 1.2em;">الإجمالي النهائي:</td>
                        <td style="font-weight: bold; font-size: 1.2em;">${total.toFixed(2)} ج.م</td>
                    </tr>
                </tfoot>
            </table>
        </div>
        
        ${order.delivery_period || order.delivery_location || order.payment_terms ? `
        <div class="print-section">
            <h4>شروط التوريد</h4>
            ${order.delivery_period ? `<p><strong>مدة التوريد:</strong> ${order.delivery_period}</p>` : ''}
            ${order.delivery_location ? `<p><strong>مكان التسليم:</strong> ${order.delivery_location}</p>` : ''}
            ${order.payment_terms ? `<p><strong>شروط الدفع:</strong> ${order.payment_terms}</p>` : ''}
        </div>
        ` : ''}
        
        ${order.notes ? `
        <div class="print-section">
            <h4>ملاحظات</h4>
            <p>${order.notes}</p>
        </div>
        ` : ''}
        
        <div class="print-section" style="margin-top: 50px;">
            <table style="width: 100%; border: none;">
                <tr>
                    <td style="width: 50%; text-align: center; border: none;">
                        <p><strong>توقيع المورد</strong></p>
                        <p>___________________</p>
                    </td>
                    <td style="width: 50%; text-align: center; border: none;">
                        <p><strong>توقيع الشركة</strong></p>
                        <p>___________________</p>
                    </td>
                </tr>
            </table>
        </div>
    `;
}

// === عرض تفاصيل الطلب ===
async function viewOrderDetails(orderId) {
    try {
        AppHelpers.showLoading('جاري تحميل التفاصيل...');
        
        const order = await AppHelpers.apiRequest(`/api/orders/${orderId}`);
        
        // إنشاء محتوى التفاصيل
        let itemsHtml = '';
        order.items.forEach((item, index) => {
            itemsHtml += `
                <tr>
                    <td>${index + 1}</td>
                    <td>${item.product_code || '-'}</td>
                    <td>${item.product_name}</td>
                    <td>${item.description || '-'}</td>
                    <td>${item.quantity}</td>
                    <td>${AppHelpers.formatCurrency(item.unit_price)}</td>
                    <td>${AppHelpers.formatCurrency(item.total_price)}</td>
                </tr>
            `;
        });
        
        const detailsHtml = `
            <div class="row g-3">
                <!-- بيانات الطلب -->
                <div class="col-12">
                    <h5 class="border-bottom pb-2">بيانات الطلب</h5>
                    <div class="row">
                        <div class="col-md-3"><strong>رقم الطلب:</strong> ${order.po_number}</div>
                        <div class="col-md-3"><strong>التاريخ:</strong> ${AppHelpers.formatDate(order.po_date)}</div>
                        <div class="col-md-3"><strong>الحالة:</strong> <span class="badge-status ${order.status}">${order.status}</span></div>
                        <div class="col-md-3"><strong>الإجمالي:</strong> ${AppHelpers.formatCurrency(order.total)}</div>
                    </div>
                </div>
                
                <!-- بيانات المورد -->
                <div class="col-12">
                    <h5 class="border-bottom pb-2">بيانات المورد</h5>
                    <div class="row">
                        <div class="col-md-6"><strong>الاسم:</strong> ${order.supplier.name}</div>
                        <div class="col-md-6"><strong>الرقم الضريبي:</strong> ${order.supplier.tax_id || '-'}</div>
                        <div class="col-md-6"><strong>التليفون:</strong> ${order.supplier.phone || '-'}</div>
                        <div class="col-md-6"><strong>البريد الإلكتروني:</strong> ${order.supplier.email || '-'}</div>
                        <div class="col-md-12"><strong>العنوان:</strong> ${order.supplier.address || '-'}</div>
                    </div>
                </div>
                
                <!-- الأصناف -->
                <div class="col-12">
                    <h5 class="border-bottom pb-2">أصناف الطلب</h5>
                    <div class="table-responsive">
                        <table class="table table-bordered table-sm">
                            <thead class="table-dark">
                                <tr>
                                    <th>م</th>
                                    <th>الكود</th>
                                    <th>الاسم</th>
                                    <th>الوصف</th>
                                    <th>الكمية</th>
                                    <th>السعر</th>
                                    <th>الإجمالي</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${itemsHtml}
                            </tbody>
                            <tfoot class="table-light">
                                <tr>
                                    <td colspan="6" class="text-end fw-bold">المجموع الفرعي:</td>
                                    <td class="fw-bold">${AppHelpers.formatCurrency(order.subtotal)}</td>
                                </tr>
                                <tr>
                                    <td colspan="6" class="text-end fw-bold">ضريبة القيمة المضافة (${order.tax_rate || 14}%):</td>
                                    <td class="fw-bold">${AppHelpers.formatCurrency(order.tax_amount)}</td>
                                </tr>
                                <tr class="table-primary">
                                    <td colspan="6" class="text-end fw-bold">الإجمالي النهائي:</td>
                                    <td class="fw-bold">${AppHelpers.formatCurrency(order.total)}</td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
                
                <!-- شروط التوريد -->
                ${order.delivery_period || order.delivery_location || order.payment_terms ? `
                <div class="col-12">
                    <h5 class="border-bottom pb-2">شروط التوريد</h5>
                    ${order.delivery_period ? `<p><strong>مدة التوريد:</strong> ${order.delivery_period}</p>` : ''}
                    ${order.delivery_location ? `<p><strong>مكان التسليم:</strong> ${order.delivery_location}</p>` : ''}
                    ${order.payment_terms ? `<p><strong>شروط الدفع:</strong> ${order.payment_terms}</p>` : ''}
                </div>
                ` : ''}
                
                <!-- ملاحظات -->
                ${order.notes ? `
                <div class="col-12">
                    <h5 class="border-bottom pb-2">ملاحظات</h5>
                    <p>${order.notes}</p>
                </div>
                ` : ''}
            </div>
        `;
        
        document.getElementById('orderDetailsBody').innerHTML = detailsHtml;
        
        AppHelpers.hideLoading();
        
        const modal = new bootstrap.Modal(document.getElementById('orderDetailsModal'));
        modal.show();
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast('فشل في تحميل التفاصيل', 'error');
    }
}

// === حذف الطلب ===
async function deleteOrder(orderId) {
    AppHelpers.showConfirm('هل أنت متأكد من حذف هذا الطلب؟ لا يمكن التراجع عن هذا الإجراء!', async () => {
        try {
            AppHelpers.showLoading('جاري حذف الطلب...');
            
            await AppHelpers.apiRequest(`/api/orders/${orderId}`, {
                method: 'DELETE'
            });
            
            AppHelpers.hideLoading();
            AppHelpers.showToast('تم حذف الطلب بنجاح', 'success');
            
            // إعادة تحميل الطلبات
            loadOrders();
        } catch (error) {
            AppHelpers.hideLoading();
            AppHelpers.showToast('فشل في حذف الطلب', 'error');
        }
    });
}
