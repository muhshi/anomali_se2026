@echo off
echo ==================================================
echo   Menjalankan Penarik Data Anomali SE2026
echo ==================================================

if not exist config.json (
    echo [ERROR] config.json tidak ditemukan!
    echo Silakan copy config.example.json menjadi config.json dan isi kredensial database Anda.
    pause
    exit /b 1
)

echo Memulai web dashboard di latar belakang...
start /B python web_anomali.py

echo Menunggu 3 detik...
timeout /t 3 /nobreak >nul

echo Memulai penarik data anomali...
python tarik_anomali.py

echo.
echo Selesai. Tekan tombol apa saja untuk keluar.
pause >nul
