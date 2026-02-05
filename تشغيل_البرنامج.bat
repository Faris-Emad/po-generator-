@echo off
chcp 65001 >nul
title نظام طلبات التوريد - شركة الأمانة
color 0A

echo.
echo ╔═══════════════════════════════════════════════════════════╗
echo ║                                                           ║
echo ║       🏢 نظام إنشاء طلبات التوريد - شركة الأمانة       ║
echo ║                                                           ║
echo ╚═══════════════════════════════════════════════════════════╝
echo.
echo [*] جاري تشغيل البرنامج...
echo.

python "امانة_طلبات_التوريد.py"

if errorlevel 1 (
    echo.
    echo ❌ خطأ: تأكد من تثبيت Python و Flask و openpyxl
    echo.
    echo للتثبيت، شغل الأمر التالي:
    echo pip install flask openpyxl
    echo.
    pause
)
