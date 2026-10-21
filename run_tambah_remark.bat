@echo off
title Bot Tambah Remark 27.a Nilai Produksi - Fasih-SM BPS
color 0a

echo ====================================================
echo   AUTO-INSTALLER ^& RUNNER - TAMBAH REMARK 27.a
echo ====================================================
echo.

:: 1. Buat folder data jika belum ada
if not exist "data" mkdir "data"

:: 2. Cek Python
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

:: 3. Virtual Environment
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
echo Memeriksa dan menginstall library (openpyxl, playwright)...
python -m pip install --upgrade pip
python -m pip install openpyxl playwright
python -m playwright install chromium

echo.
echo ====================================================
echo Menjalankan Bot Tambah Remark 27.a (#0)...
echo ====================================================
echo.
python tambah_remark_27a.py

echo.
echo ====================================================
echo Selesai. Tekan sembarang tombol untuk keluar.
echo ====================================================
pause >nul
