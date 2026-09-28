@echo off
title Bot Otomatisasi Reject Anomali Fasih-SM BPS

echo ====================================================
echo   AUTO-INSTALLER ^& RUNNER - BPS FASIH ANOMALI BOT
echo ====================================================
echo.

:: 1. Buat folder data dan subfolder jika belum ada
if not exist "data" mkdir "data"
if not exist "data\Anomali" mkdir "data\Anomali"
if not exist "data\BKU Ditautkan" mkdir "data\BKU Ditautkan"

:: 2. Pindahkan file .xlsx dari folder utama ke folder data/subfolder jika ada
for %%f in (*anomali*.xlsx) do (
    echo [Info] Memindahkan file Anomali: %%f ke data\Anomali\
    move /y "%%f" "data\Anomali\" >nul
)
for %%f in (*bku*.xlsx) do (
    echo [Info] Memindahkan file BKU: %%f ke data\BKU Ditautkan\
    move /y "%%f" "data\BKU Ditautkan\" >nul
)
for %%f in (*.xlsx) do (
    if not "%%f"=="laporan_status_ditemukan.xlsx" (
        echo [Info] Menemukan file Excel: %%f
        move /y "%%f" "data\" >nul
    )
)

:: 3. Cek ketersediaan file Excel di folder data (termasuk subfolder)
set EXCEL_FOUND=0
for /r data %%f in (*.xlsx) do (
    echo %%~nxf | findstr /i /v "laporan" >nul && set EXCEL_FOUND=1
)

if "%EXCEL_FOUND%"=="1" goto EXCEL_OK

echo [Peringatan] TIDAK ADA FILE EXCEL DITEMUKAN di folder 'data' (atau subfolder 'Anomali' / 'BKU Ditautkan')!
echo Silakan masukkan file Excel anomali / BKU (.xlsx) ke dalam folder 'data'.
echo.
echo Tekan ENTER setelah Anda menaruh file Excel di folder 'data'...
pause >nul

set EXCEL_FOUND=0
for /r data %%f in (*.xlsx) do (
    echo %%~nxf | findstr /i /v "laporan" >nul && set EXCEL_FOUND=1
)

if "%EXCEL_FOUND%"=="1" goto EXCEL_OK

echo.
echo Masih belum ada file Excel di folder 'data'. Program dihentikan.
echo Tekan sembarang tombol untuk keluar...
pause >nul
exit /b

:EXCEL_OK
echo [Info] File Excel data ditemukan.

:: 4. Cek Python
set "PY_CMD="

python --version >nul 2>&1
if %errorlevel% equ 0 set "PY_CMD=python"

if "%PY_CMD%"=="" (
    py --version >nul 2>&1
    if %errorlevel% equ 0 set "PY_CMD=py"
)

if "%PY_CMD%"=="" (
    if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PY_CMD=%LocalAppData%\Programs\Python\Python311\python.exe"
)

if "%PY_CMD%"=="" (
    if exist "%LocalAppData%\Programs\Python\Python310\python.exe" set "PY_CMD=%LocalAppData%\Programs\Python\Python310\python.exe"
)

if "%PY_CMD%"=="" (
    if exist "C:\laragon\bin\python\python-3.10\python.exe" set "PY_CMD=C:\laragon\bin\python\python-3.10\python.exe"
)

if not "%PY_CMD%"=="" goto PYTHON_OK

echo.
echo [Info] Python tidak ditemukan di PATH. Mendownload Python 3.11...
powershell -Command "[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -OutFile 'python_installer.exe'"

if not exist python_installer.exe goto PY_INSTALL_FAIL

echo Menginstall Python... Mohon tunggu 1-2 menit.
start /wait python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_doc=0
del python_installer.exe >nul 2>&1

if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PY_CMD=%LocalAppData%\Programs\Python\Python311\python.exe"

python --version >nul 2>&1
if %errorlevel% equ 0 set "PY_CMD=python"

if not "%PY_CMD%"=="" goto PYTHON_OK

:PY_INSTALL_FAIL
echo.
echo [ERROR] Python tidak terdeteksi dan instalasi gagal.
echo Silakan install Python 3.11/3.12 dari https://www.python.org/ secara manual.
echo Tekan sembarang tombol untuk keluar...
pause >nul
exit /b

:PYTHON_OK
echo [Info] Menggunakan Python: %PY_CMD%

:: 5. Virtual Environment
if not exist "venv\Scripts\python.exe" goto CREATE_VENV

"venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if %errorlevel% equ 0 goto VENV_OK

echo [Peringatan] Virtual environment venv lama tidak valid atau berasal dari PC lain.
echo Menghapus venv dan membuat ulang khusus untuk PC ini...
rd /s /q "venv" >nul 2>&1

:CREATE_VENV
echo.
echo Membuat Virtual Environment venv...
"%PY_CMD%" -m venv venv

:VENV_OK
echo Mengaktifkan Virtual Environment...
call venv\Scripts\activate.bat

echo.
echo Memeriksa dan menginstall library (openpyxl, playwright, pandas)...
python -m pip install --upgrade pip
python -m pip install openpyxl playwright pandas
python -m playwright install chromium

echo.
echo ====================================================
echo PILIH BOT YANG INGIN DIJALANKAN:
echo ====================================================
echo [1] Bot Reject Anomali (reject_anomali.py)
echo [2] Bot Edit Anomali by Admin (edit_anomali_admin.py) [DEFAULT]
echo [3] Bot Ubah Status Ditemukan ^& Submit Paksa (ubah_status_ditemukan.py)
echo ====================================================
set "BOT_CHOICE=2"
set /p BOT_CHOICE="Pilih nomor [1/2/3] (Default: 2): "

if "%BOT_CHOICE%"=="1" (
    echo.
    echo Menjalankan Bot Otomatisasi Reject Anomali...
    python reject_anomali.py
) else if "%BOT_CHOICE%"=="3" (
    echo.
    echo Menjalankan Bot Ubah Status Ditemukan ^& Submit Paksa...
    python ubah_status_ditemukan.py
) else (
    echo.
    echo Menjalankan Bot Edit Anomali by Admin...
    python edit_anomali_admin.py
)

echo.
echo ====================================================
echo Selesai. Tekan sembarang tombol untuk keluar.
echo ====================================================
pause >nul
