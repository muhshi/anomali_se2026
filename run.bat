@echo off
title Bot Otomatisasi Reject Anomali Fasih-SM BPS

echo ====================================================
echo   AUTO-INSTALLER ^& RUNNER - BPS FASIH ANOMALI BOT
echo ====================================================
echo.

:: 1. Buat folder data jika belum ada
if not exist "data" mkdir "data"

:: 2. Pindahkan file .xlsx dari folder utama ke folder data jika ada
for %%f in (*.xlsx) do (
    echo [Info] Menemukan file Excel: %%f
    move /y "%%f" "data\" >nul
)

:: 3. Cek ketersediaan file Excel di folder data
set EXCEL_FOUND=0
for %%f in (data\*.xlsx) do set EXCEL_FOUND=1

if "%EXCEL_FOUND%"=="1" goto EXCEL_OK

echo [Peringatan] TIDAK ADA FILE EXCEL DITEMUKAN di folder 'data'!
echo Silakan masukkan file Excel anomali (.xlsx) ke dalam folder 'data'.
echo.
echo Tekan ENTER setelah Anda menaruh file Excel di folder 'data'...
pause >nul

set EXCEL_FOUND=0
for %%f in (data\*.xlsx) do set EXCEL_FOUND=1

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
if exist "venv\Scripts\activate.bat" goto VENV_OK

echo.
echo Membuat Virtual Environment (venv)...
"%PY_CMD%" -m venv venv

:VENV_OK
echo Mengaktifkan Virtual Environment...
call venv\Scripts\activate.bat

echo.
echo Memeriksa dan menginstall library (pandas, openpyxl, playwright)...
python -m pip install --upgrade pip
pip install pandas openpyxl playwright
playwright install chromium

echo.
echo ====================================================
echo Menjalankan Bot Otomatisasi Reject Anomali...
echo ====================================================
echo.
python reject_anomali.py

echo.
echo ====================================================
echo Selesai. Tekan sembarang tombol untuk keluar.
echo ====================================================
pause >nul
