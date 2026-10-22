@echo off
title Portal Karyawan - Scan Agent
cd /d "%~dp0"
echo ============================================
echo   Portal Karyawan - Scan Agent Lokal
echo   http://127.0.0.1:3000
echo ============================================
echo.
"%~dp0ScanAgent.exe"
if errorlevel 1 pause
