/**
 * صفحة إنشاء طلب جديد
 * New Order Creation Page
 */

let itemCounter = 0;
let suppliers = [];
let products = [];

// === تحميل البيانات عند فتح الصفحة ===
document.addEventListener('DOMContentLoaded', async function() {
    await loadSuppliers();
    await loadProducts();
    loadFormData();
    addItemRow(); // إضافة صف واحد على الأقل
    
    // حفظ تلقائي كل 30 ثانية
    setInterval(autoSaveFormData, 30000);
});

// === تحميل قائمة الموردين ===
async function loadSuppliers() {
    try {
        suppliers = await AppHelpers.apiRequest('/api/suppliers');
        const select = document.getElementById('supplierId');
        select.innerHTML = '<option value="">-- اختر المورد --</option>';
        suppliers.forEach(supplier => {
            select.innerHTML += `<option value="${supplier.id}">${supplier.name}</option>`;
        });
    } catch (error) {
        AppHelpers.showToast('خطأ في تحميل الموردين', 'error');
    }
}

// === تحميل قائمة الأصناف ===
async function loadProducts() {
    try {
        products = await AppHelpers.apiRequest('/api/products');
    } catch (error) {
        AppHelpers.showToast('خطأ في تحميل الأصناف', 'error');
    }
}

// === تحميل بيانات المورد عند اختياره ===
function loadSupplierData() {
    const supplierId = document.getElementById('supplierId').value;
    if (!supplierId) {
        document.getElementById('supplierTaxId').value = '';
        document.getElementById('supplierPhone').value = '';
        document.getElementById('supplierEmail').value = '';
        document.getElementById('supplierAddress').value = '';
        return;
    }
    
    const supplier = suppliers.find(s => s.id == supplierId);
    if (supplier) {
        document.getElementById('supplierTaxId').value = supplier.tax_id || '';
        document.getElementById('supplierPhone').value = supplier.phone || '';
        document.getElementById('supplierEmail').value = supplier.email || '';
        document.getElementById('supplierAddress').value = supplier.address || '';
    }
}

// === إضافة صف صنف جديد ===
function addItemRow() {
    itemCounter++;
    const tbody = document.getElementById('itemsTableBody');
    
    // إنشاء select للأصناف
    let productsSelect = '<option value="">-- اختر الصنف --</option>';
    products.forEach(product => {
        productsSelect += `<option value="${product.id}" 
            data-code="${product.code}" 
            data-name="${product.name}" 
            data-description="${product.description || ''}" 
            data-price="${product.default_price}">
            ${product.code} - ${product.name}
        </option>`;
    });
    
    const row = document.createElement('tr');
    row.className = 'item-row';
    row.innerHTML = `
        <td class="text-center">${itemCounter}</td>
        <td>
            <select class="form-select product-select" onchange="loadProductData(this)">
                ${productsSelect}
            </select>
        </td>
        <td><input type="text" class="form-control item-code" placeholder="الكود"></td>
        <td><input type="text" class="form-control item-name" placeholder="الاسم" required></td>
        <td><input type="text" class="form-control item-description" placeholder="الوصف"></td>
        <td><input type="number" class="form-control item-quantity" placeholder="0" step="0.01" min="0" value="1" onchange="calculateTotals()"></td>
        <td><input type="number" class="form-control item-price" placeholder="0.00" step="0.01" min="0" value="0" onchange="calculateTotals()"></td>
        <td class="text-center fw-bold item-total">0.00</td>
        <td class="text-center">
            <button type="button" class="btn btn-danger btn-sm" onclick="removeItemRow(this)">
                <i class="fas fa-trash"></i>
            </button>
        </td>
    `;
    
    tbody.appendChild(row);
}

// === تحميل بيانات الصنف عند اختياره ===
function loadProductData(select) {
    const row = select.closest('tr');
    const selectedOption = select.options[select.selectedIndex];
    
    if (selectedOption.value) {
        row.querySelector('.item-code').value = selectedOption.dataset.code || '';
        row.querySelector('.item-name').value = selectedOption.dataset.name || '';
        row.querySelector('.item-description').value = selectedOption.dataset.description || '';
        row.querySelector('.item-price').value = selectedOption.dataset.price || '0';
        calculateTotals();
    }
}

// === حذف صف صنف ===
function removeItemRow(button) {
    AppHelpers.showConfirm('هل أنت متأكد من حذف هذا الصنف؟', () => {
        button.closest('tr').remove();
        updateItemNumbers();
        calculateTotals();
        AppHelpers.showToast('تم حذف الصنف', 'info');
    });
}

// === تحديث أرقام الأصناف بعد الحذف ===
function updateItemNumbers() {
    const rows = document.querySelectorAll('#itemsTableBody tr');
    rows.forEach((row, index) => {
        row.querySelector('td:first-child').textContent = index + 1;
    });
    itemCounter = rows.length;
}

// === حساب الإجماليات ===
function calculateTotals() {
    let subtotal = 0;
    
    document.querySelectorAll('.item-row').forEach(row => {
        const quantity = parseFloat(row.querySelector('.item-quantity').value) || 0;
        const price = parseFloat(row.querySelector('.item-price').value) || 0;
        const total = quantity * price;
        
        row.querySelector('.item-total').textContent = total.toFixed(2);
        subtotal += total;
    });
    
    const taxRateValue = parseFloat(document.getElementById('taxRate').value);
    const taxRate = isNaN(taxRateValue) ? 14 : taxRateValue;
    const tax = subtotal * (taxRate / 100);
    const finalTotal = subtotal + tax;
    
    document.getElementById('subtotal').textContent = AppHelpers.formatCurrency(subtotal);
    document.getElementById('taxAmount').textContent = AppHelpers.formatCurrency(tax);
    document.getElementById('totalAmount').textContent = AppHelpers.formatCurrency(finalTotal);
    
    // تحديث نص الضريبة
    document.getElementById('taxLabel').textContent = `ضريبة القيمة المضافة (${taxRate}%):`;
}

// === حفظ البيانات في localStorage ===
function autoSaveFormData() {
    const formData = collectFormData();
    AppHelpers.saveToLocalStorage('orderFormDraft', formData);
}

// === تحميل البيانات من localStorage ===
function loadFormData() {
    const savedData = AppHelpers.loadFromLocalStorage('orderFormDraft');
    if (savedData && confirm('تم العثور على بيانات محفوظة. هل تريد استعادتها؟')) {
        restoreFormData(savedData);
    }
}

// === جمع بيانات النموذج ===
function collectFormData() {
    const items = [];
    document.querySelectorAll('.item-row').forEach(row => {
        items.push({
            product_code: row.querySelector('.item-code').value,
            product_name: row.querySelector('.item-name').value,
            description: row.querySelector('.item-description').value,
            quantity: row.querySelector('.item-quantity').value,
            unit_price: row.querySelector('.item-price').value
        });
    });
    
    return {
        po_number: document.getElementById('poNumber').value,
        po_date: document.getElementById('poDate').value,
        company_tax_id: document.getElementById('companyTaxId').value,
        commercial_reg: document.getElementById('commercialReg').value,
        supplier_id: document.getElementById('supplierId').value,
        delivery_period: document.getElementById('deliveryPeriod').value,
        delivery_location: document.getElementById('deliveryLocation').value,
        payment_terms: document.getElementById('paymentTerms').value,
        notes: document.getElementById('notes').value,
        tax_rate: document.getElementById('taxRate').value,
        items: items
    };
}

// === استعادة بيانات النموذج ===
function restoreFormData(data) {
    document.getElementById('poDate').value = data.po_date || '';
    document.getElementById('companyTaxId').value = data.company_tax_id || '';
    document.getElementById('commercialReg').value = data.commercial_reg || '';
    document.getElementById('supplierId').value = data.supplier_id || '';
    document.getElementById('deliveryPeriod').value = data.delivery_period || '';
    document.getElementById('deliveryLocation').value = data.delivery_location || '';
    document.getElementById('paymentTerms').value = data.payment_terms || '';
    document.getElementById('notes').value = data.notes || '';
    
    if (data.tax_rate !== undefined && data.tax_rate !== null) {
        document.getElementById('taxRate').value = data.tax_rate;
    }
    
    loadSupplierData();
    
    // استعادة الأصناف
    if (data.items && data.items.length > 0) {
        document.getElementById('itemsTableBody').innerHTML = '';
        itemCounter = 0;
        data.items.forEach(item => {
            addItemRow();
            const lastRow = document.querySelector('#itemsTableBody tr:last-child');
            lastRow.querySelector('.item-code').value = item.product_code || '';
            lastRow.querySelector('.item-name').value = item.product_name || '';
            lastRow.querySelector('.item-description').value = item.description || '';
            lastRow.querySelector('.item-quantity').value = item.quantity || '';
            lastRow.querySelector('.item-price').value = item.unit_price || '';
        });
        calculateTotals();
    }
}

// === مسح النموذج ===
function clearForm() {
    AppHelpers.showConfirm('هل أنت متأكد من مسح جميع البيانات؟', () => {
        document.getElementById('orderForm').reset();
        document.getElementById('itemsTableBody').innerHTML = '';
        itemCounter = 0;
        addItemRow();
        calculateTotals();
        AppHelpers.removeFromLocalStorage('orderFormDraft');
        AppHelpers.showToast('تم مسح النموذج بنجاح', 'info');
    });
}

// === عرض نافذة إضافة مورد ===
function showAddSupplierModal() {
    const modal = new bootstrap.Modal(document.getElementById('addSupplierModal'));
    modal.show();
}

// === حفظ مورد جديد ===
async function saveNewSupplier() {
    const data = {
        name: document.getElementById('newSupplierName').value,
        tax_id: document.getElementById('newSupplierTaxId').value,
        phone: document.getElementById('newSupplierPhone').value,
        email: document.getElementById('newSupplierEmail').value,
        address: document.getElementById('newSupplierAddress').value
    };
    
    if (!data.name) {
        AppHelpers.showToast('يجب إدخال اسم المورد', 'error');
        return;
    }
    
    try {
        AppHelpers.showLoading('جاري حفظ المورد...');
        const result = await AppHelpers.apiRequest('/api/suppliers', {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        AppHelpers.hideLoading();
        AppHelpers.showToast(result.message, 'success');
        
        // إعادة تحميل قائمة الموردين
        await loadSuppliers();
        
        // اختيار المورد الجديد
        document.getElementById('supplierId').value = result.supplier.id;
        loadSupplierData();
        
        // إغلاق النافذة ومسح الحقول
        bootstrap.Modal.getInstance(document.getElementById('addSupplierModal')).hide();
        document.getElementById('addSupplierForm').reset();
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast(error.message, 'error');
    }
}

// === عرض نافذة إضافة صنف ===
function showAddProductModal() {
    const modal = new bootstrap.Modal(document.getElementById('addProductModal'));
    modal.show();
}

// === حفظ صنف جديد ===
async function saveNewProduct() {
    const data = {
        code: document.getElementById('newProductCode').value,
        name: document.getElementById('newProductName').value,
        description: document.getElementById('newProductDescription').value,
        default_price: document.getElementById('newProductPrice').value
    };
    
    if (!data.code || !data.name) {
        AppHelpers.showToast('يجب إدخال الكود واسم الصنف', 'error');
        return;
    }
    
    try {
        AppHelpers.showLoading('جاري حفظ الصنف...');
        const result = await AppHelpers.apiRequest('/api/products', {
            method: 'POST',
            body: JSON.stringify(data)
        });
        
        AppHelpers.hideLoading();
        AppHelpers.showToast(result.message, 'success');
        
        // إعادة تحميل قائمة الأصناف
        await loadProducts();
        
        // إعادة بناء صفوف الأصناف لتحديث القوائم
        const rows = document.querySelectorAll('.item-row');
        rows.forEach(row => {
            const select = row.querySelector('.product-select');
            let productsSelect = '<option value="">-- اختر الصنف --</option>';
            products.forEach(product => {
                productsSelect += `<option value="${product.id}" 
                    data-code="${product.code}" 
                    data-name="${product.name}" 
                    data-description="${product.description || ''}" 
                    data-price="${product.default_price}">
                    ${product.code} - ${product.name}
                </option>`;
            });
            select.innerHTML = productsSelect;
        });
        
        // إغلاق النافذة ومسح الحقول
        bootstrap.Modal.getInstance(document.getElementById('addProductModal')).hide();
        document.getElementById('addProductForm').reset();
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast(error.message, 'error');
    }
}

// === إرسال النموذج ===
document.getElementById('orderForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    // التحقق من وجود أصناف
    const items = document.querySelectorAll('.item-row');
    if (items.length === 0) {
        AppHelpers.showToast('يجب إضافة صنف واحد على الأقل', 'error');
        return;
    }
    
    // التحقق من اختيار مورد
    if (!document.getElementById('supplierId').value) {
        AppHelpers.showToast('يجب اختيار المورد', 'error');
        return;
    }
    
    const formData = collectFormData();
    
    try {
        AppHelpers.showLoading('جاري حفظ الطلب...');
        const result = await AppHelpers.apiRequest('/api/orders', {
            method: 'POST',
            body: JSON.stringify(formData)
        });
        
        AppHelpers.hideLoading();
        AppHelpers.showToast('تم حفظ الطلب بنجاح! ✅', 'success');
        
        // مسح البيانات المحفوظة
        AppHelpers.removeFromLocalStorage('orderFormDraft');
        
        // الانتقال لصفحة الطلبات بعد ثانيتين
        setTimeout(() => {
            window.location.href = '/orders';
        }, 2000);
        
    } catch (error) {
        AppHelpers.hideLoading();
        AppHelpers.showToast(error.message, 'error');
    }
});
