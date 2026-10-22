@echo off
setlocal
title Install Dependensi Portal Karyawan
cd /d "%~dp0"

echo ============================================================
echo   INSTALL DEPENDENSI PORTAL KARYAWAN K.O.M.A.N.D.O
echo ============================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [GAGAL] Python tidak ditemukan.
    echo.
    echo Install Python 3.12 terlebih dahulu dari:
    echo   https://www.python.org/downloads/
    echo.
    echo Saat instalasi, centang "Add python.exe to PATH".
    pause
    exit /b 1
)

echo Python terdeteksi:
python --version
echo.

echo [1/3] Memperbarui pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [PERINGATAN] Gagal memperbarui pip. Mencoba lanjut...
)

echo.
echo [2/3] Memasang dependensi dari requirements.txt...
echo Proses ini bisa memakan waktu beberapa menit dan mengunduh banyak data.
echo.
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo [GAGAL] Sebagian dependensi gagal dipasang.
    echo Periksa koneksi internet, lalu jalankan file ini lagi.
    pause
    exit /b 1
)

echo.
echo [3/3] Mengunduh data NLTK...
python -c "import nltk; nltk.download('punkt', quiet=True); nltk.download('punkt_tab', quiet=True)"
if errorlevel 1 (
    echo [PERINGATAN] Data NLTK belum berhasil diunduh.
    echo Fitur Ringkas Dokumen mungkin perlu koneksi internet saat pertama dipakai.
)

echo.
echo Memeriksa import utama...
python -c "import flask, flask_cors, pymysql, requests, docx, pdf2docx, docx2pdf, pymupdf, sumy, Sastrawi, nltk, numpy, rapidocr_onnxruntime, win32com.client, firebase_admin; print('Semua dependensi utama OK')"
if errorlevel 1 (
    echo.
    echo [GAGAL] Ada dependensi yang belum bisa diimpor.
    echo Jalankan lagi INSTALL-DEPENDENSI.bat atau kirim pesan error di atas.
    pause
    exit /b 1
)

echo.
powershell -NoProfile -Command "try { $w = New-Object -ComObject Word.Application; $w.Quit(); exit 0 } catch { exit 1 }"
if errorlevel 1 (
    echo [PERINGATAN] Microsoft Word tidak terdeteksi.
    echo Fitur Word ke PDF dan cetak PDF disposisi membutuhkan Microsoft Word.
)

if not exist "C:\xampp\mysql\bin\mysql.exe" (
    echo [PERINGATAN] XAMPP/MySQL tidak ditemukan di C:\xampp.
    echo Pastikan MySQL/MariaDB sudah terpasang dan berjalan.
)

echo.
echo ============================================================
echo   SELESAI
echo   Jalankan JALANKAN-APP.bat untuk membuka portal.
echo ============================================================
pause
