@echo off
chcp 65001 > nul
title 설비 QR 이력관리 시스템
cd /d "%~dp0"

set PYTHON=C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe
set QR_SERVER_HOST=desktop-v8l9jm0.tail1d2229.ts.net

echo.
echo ========================================
echo   설비 QR 이력관리 시스템
echo ========================================
echo.
echo   접속 주소: http://desktop-v8l9jm0.tail1d2229.ts.net:5001
echo.

:: 포트 5001 사용 중인 프로세스 종료
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5001" 2^>nul') do (
    taskkill /F /PID %%a >nul 2>&1
)

%PYTHON% app.py

pause
