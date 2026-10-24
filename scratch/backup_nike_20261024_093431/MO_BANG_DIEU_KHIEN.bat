@echo off
title BANG DIEU KHIEN VA TRO LY SFC CHING LUH (COPILOT COCKPIT)
chcp 65001 >nul
echo ===============================================================================
echo       🚀 ĐANG KHỞI ĐỘNG BẢNG ĐIỀU KHIỂN SFC CHING LUH COPILOT...
echo ===============================================================================
echo.
echo [*] Kiểm tra kết nối Chrome cổng 9222...
echo [*] Đang khởi chạy Dashboard Server tại http://localhost:5050...
echo.

start "SFC Cockpit Server" /min python sfc_dashboard_server.py

timeout /t 2 /nobreak >nul

start http://localhost:5050

echo [THANH CONG] Đã mở Bảng điều khiển SFC trên trình duyệt!
echo Bạn có thể thu nhỏ cửa sổ này lại.
echo.
pause
