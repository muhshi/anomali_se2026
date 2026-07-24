@echo off
setlocal enabledelayedexpansion

echo ====================================================
echo   AUTO-INSTALLER ^& RUNNER - BPS FASIH ANOMALI BOT
echo ====================================================

:: Buat folder data jika belum ada
if not exist "data\" (
    mkdir "data"
    echo.
    echo Folder 'data' telah dibuat!
    echo Silakan masukkan file Excel .xlsx ke dalam folder 'data' lalu jalankan lagi file ini.
    echo.
    pause
    exit /b
)

:: Cek apakah file excel ada di folder data
set count=0
for %%x in (data\*.xlsx) do set /a count+=1
if %count%==0 (
    echo.
    echo TIDAK ADA FILE EXCEL DITEMUKAN!
    echo Silakan masukkan file Excel anomali .xlsx ke dalam folder 'data' lalu jalankan ulang.
    echo.
    pause
    exit /b
)

:: Cek apakah Python terinstall
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Python belum terinstall atau belum masuk PATH!
    echo Mendownload installer Python 3.11...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -OutFile 'python_installer.exe'"
    if exist python_installer.exe (
        echo Menginstall Python... Mohon tunggu, proses ini memakan waktu 1-2 menit.
        start /wait python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_doc=0
        echo Instalasi selesai. 
        echo Mohon TUTUP JENDELA INI dan JALANKAN ULANG run.bat agar Windows mengenali Python.
        pause
        exit /b
    ) else (
        echo Gagal mendownload Python. Silakan install Python dari python.org secara manual.
        pause
        exit /b
    )
)

echo Python terdeteksi.

:: Cek Virtual Environment
if not exist "venv\Scripts\activate.bat" (
    echo.
    echo Membuat Virtual Environment (venv) agar tidak bentrok dengan program lain...
    python -m venv venv
)

echo.
echo Mengaktifkan Virtual Environment...
call venv\Scripts\activate.bat

echo Memeriksa dan menginstall library yang dibutuhkan (Pandas, Openpyxl, Playwright)...
python -m pip install --upgrade pip >nul 2>&1
pip install pandas openpyxl playwright >nul 2>&1

echo Mendownload engine browser Playwright (Chromium) jika belum ada...
playwright install chromium >nul 2>&1

echo.
echo ====================================================
echo Menjalankan Script Otomatisasi...
echo ====================================================
python reject_anomali.py

echo.
pause
