@echo off
title Portal Karyawan - Scan Agent
cd /d "%~dp0"
echo ============================================
echo   Portal Karyawan - Scan Agent Lokal
echo   http://127.0.0.1:3000
echo ============================================
echo.

rem Versi siap pakai (tidak perlu Python)
if exist "%~dp0ScanAgent.exe" (
    "%~dp0ScanAgent.exe"
    if errorlevel 1 pause
    exit /b
)
if exist "%~dp0scan-agent-dist\ScanAgent.exe" (
    "%~dp0scan-agent-dist\ScanAgent.exe"
    if errorlevel 1 pause
    exit /b
)

rem Cadangan: jalankan dari source Python
where python >nul 2>&1
if errorlevel 1 (
    echo Python tidak ditemukan dan ScanAgent.exe tidak ada.
    echo Salin ScanAgent.exe dari folder scan-agent-dist ke folder ini,
    echo atau instal Python 3.10+ lalu jalankan:
    echo   pip install -r requirements-scan-agent.txt
    pause
    exit /b 1
)

python -c "import flask, flask_cors, win32com, fitz" >nul 2>&1
if errorlevel 1 (
    echo Dependensi Scan Agent belum lengkap.
    echo Jalankan perintah berikut, lalu buka file ini lagi:
    echo   pip install -r requirements-scan-agent.txt
    pause
    exit /b 1
)

python scan_agent.py
if errorlevel 1 (
    echo.
    echo Scan Agent berhenti. Periksa pesan error di atas.
    pause
)
