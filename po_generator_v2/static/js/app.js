/**
 * ملف JavaScript الرئيسي للتطبيق
 * Main JavaScript file for the application
 */

// === نظام الإشعارات - Toast Notifications System ===
function showToast(message, type = 'success') {
    const toastEl = document.getElementById('toastNotification');
    const toastBody = toastEl.querySelector('.toast-body');
    const toastMessage = toastEl.querySelector('.toast-message');
    const toastIcon = toastEl.querySelector('.toast-icon');
    
    // تحديد الأيقونة واللون حسب النوع
    let icon = '';
    let bgClass = '';
    
    switch(type) {
        case 'success':
            icon = 'fas fa-check-circle';
            bgClass = 'bg-success text-white';
            break;
        case 'error':
            icon = 'fas fa-times-circle';
            bgClass = 'bg-danger text-white';
            break;
        case 'warning':
            icon = 'fas fa-exclamation-triangle';
            bgClass = 'bg-warning text-dark';
            break;
        case 'info':
            icon = 'fas fa-info-circle';
            bgClass = 'bg-info text-white';
            break;
    }
    
    // تطبيق الأيقونة والرسالة
    toastIcon.className = `toast-icon me-2 ${icon}`;
    toastMessage.textContent = message;
    
    // إزالة جميع الكلاسات السابقة وإضافة الكلاسات الجديدة
    toastEl.className = `toast align-items-center border-0 ${bgClass}`;
    toastBody.className = 'toast-body fw-bold';
    
    // عرض الإشعار
    const toast = new bootstrap.Toast(toastEl, {
        autohide: true,
        delay: 4000
    });
    toast.show();
}

// === نظام التأكيد - Confirmation System ===
function showConfirm(message, onConfirm) {
    if (confirm(message)) {
        onConfirm();
    }
}

// === نظام التحميل - Loading System ===
function showLoading(message = 'جاري المعالجة...') {
    // إزالة أي overlay موجود
    hideLoading();
    
    const overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.className = 'loading-overlay';
    overlay.innerHTML = `
        <div class="loading-spinner">
            <div class="spinner-border text-light" role="status"></div>
            <p>${message}</p>
        </div>
    `;
    document.body.appendChild(overlay);
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.remove();
    }
}

// === مساعدات API - API Helpers ===
async function apiRequest(url, options = {}) {
    try {
        const response = await fetch(url, {
            headers: {
                'Content-Type': 'application/json',
                ...options.headers
            },
            ...options
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.message || 'حدث خطأ في الطلب');
        }
        
        return await response.json();
    } catch (error) {
        console.error('API Error:', error);
        throw error;
    }
}

// === تنسيق العملة - Currency Formatting ===
function formatCurrency(amount) {
    return new Intl.NumberFormat('ar-EG', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(amount) + ' ج.م';
}

// === تنسيق التاريخ - Date Formatting ===
function formatDate(dateString) {
    const date = new Date(dateString);
    return new Intl.DateTimeFormat('ar-EG', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    }).format(date);
}

// === LocalStorage Helpers ===
function saveToLocalStorage(key, data) {
    try {
        localStorage.setItem(key, JSON.stringify(data));
    } catch (error) {
        console.error('Error saving to localStorage:', error);
    }
}

function loadFromLocalStorage(key) {
    try {
        const data = localStorage.getItem(key);
        return data ? JSON.parse(data) : null;
    } catch (error) {
        console.error('Error loading from localStorage:', error);
        return null;
    }
}

function removeFromLocalStorage(key) {
    try {
        localStorage.removeItem(key);
    } catch (error) {
        console.error('Error removing from localStorage:', error);
    }
}

// === تهيئة عند تحميل الصفحة - Initialize on page load ===
document.addEventListener('DOMContentLoaded', function() {
    // تعيين التاريخ الحالي تلقائياً
    const dateInputs = document.querySelectorAll('input[type="date"]');
    dateInputs.forEach(input => {
        if (!input.value) {
            input.valueAsDate = new Date();
        }
    });
});

// === Export functions for use in other scripts ===
window.AppHelpers = {
    showToast,
    showConfirm,
    showLoading,
    hideLoading,
    apiRequest,
    formatCurrency,
    formatDate,
    saveToLocalStorage,
    loadFromLocalStorage,
    removeFromLocalStorage
};
