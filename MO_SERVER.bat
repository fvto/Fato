@echo off
setlocal
chcp 65001 >nul
title SFC DASHBOARD SERVER
cd /d "%~dp0"
set "MQAA_MODE=%~1"
if not defined MQAA_MODE set "MQAA_MODE=server"
if exist "%~dp0.venv\Scripts\python.exe" goto venv_python
python -c "import sys" >nul 2>&1
if not errorlevel 1 goto system_python
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 goto py_launcher
echo [LOI] Khong tim thay Python 3. Hay cai Python roi thu lai.
pause
exit /b 1

:venv_python
"%~dp0.venv\Scripts\python.exe" "%~dp0dashboard_launcher.py" "%MQAA_MODE%"
goto finished

:system_python
python "%~dp0dashboard_launcher.py" "%MQAA_MODE%"
goto finished

:py_launcher
py -3 "%~dp0dashboard_launcher.py" "%MQAA_MODE%"

:finished
set "MQAA_RESULT=%errorlevel%"
if not "%MQAA_RESULT%"=="0" pause
if "%MQAA_MODE%"=="server" if "%MQAA_RESULT%"=="0" pause
exit /b %MQAA_RESULT%
