@echo off
title Portal Karyawan - K.O.M.A.N.D.O
cd /d "%~dp0"

REM ============================================================
REM  Konfigurasi database MySQL
REM  MYSQL_USER / MYSQL_PASSWORD adalah akun MySQL untuk aplikasi,
REM  BUKAN akun login portal. Ganti nilainya sesuai user di phpMyAdmin.
REM ============================================================
set MYSQL_HOST=127.0.0.1
set MYSQL_PORT=3306
set MYSQL_USER=satsibertni
set MYSQL_PASSWORD=satsiber2026
set MYSQL_DB=portal_karyawan

echo ============================================================
echo   Menjalankan Portal Karyawan...
echo   Setelah muncul "Running on http://127.0.0.1:5000",
echo   buka browser ke alamat tersebut.
echo   Untuk menghentikan: tekan Ctrl+C di jendela ini.
echo ============================================================
echo.

python app.py

echo.
echo ============================================================
echo   Aplikasi berhenti. Tekan tombol apa saja untuk menutup.
echo ============================================================
pause >nul
