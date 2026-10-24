@echo off
setlocal
chcp 65001 >nul
title BANG DIEU KHIEN SFC CHING LUH
cd /d "%~dp0"
echo ============================================================
echo                 BANG DIEU KHIEN SFC
echo ============================================================
echo.
echo   1. Mở Dashboard
echo   2. Mở dashboard và DOM
echo.
choice /c 12 /n /m "Chọn 1 hoặc 2: "
if errorlevel 2 goto with_dom
call "%~dp0MO_SERVER.bat" dashboard
exit /b %errorlevel%

:with_dom
call "%~dp0MO_SERVER.bat" dom
exit /b %errorlevel%
