@echo off
setlocal enabledelayedexpansion

title Bot Otomatisasi Reject Anomali Fasih-SM BPS

echo ====================================================
echo   AUTO-INSTALLER ^& RUNNER - BPS FASIH ANOMALI BOT
echo ====================================================
echo.

:: 1. Buat folder data jika belum ada
if not exist "data\" (
    mkdir "data"
    echo [Info] Folder 'data' telah dibuat!
)

:: 2. Cek ketersediaan file Excel
set count=0
for %%x in (data\*.xlsx) do set /a count+=1
if %count%==0 (
    echo [Peringatan] TIDAK ADA FILE EXCEL DITEMUKAN di folder 'data'!
    echo Silakan masukkan file Excel anomali (.xlsx) ke dalam folder 'data'.
    echo.
    echo Tekan sembarang tombol setelah Anda menaruh file Excel di folder 'data'...
    pause >nul
)

:: Re-check file Excel setelah pause
set count=0
for %%x in (data\*.xlsx) do set /a count+=1
if %count%==0 (
    echo.
    echo Masih belum ada file Excel di folder 'data'. Silakan jalankan ulang run.bat setelah file disiapkan.
    goto FINISH
)

:: 3. Cek Python
set "PYTHON_EXE="

python --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_EXE=python"
) else (
    py --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON_EXE=py"
    ) else (
        if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
            set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"
            set "PATH=%LocalAppData%\Programs\Python\Python311;%LocalAppData%\Programs\Python\Python311\Scripts;!PATH!"
        ) else if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
            set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
            set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;!PATH!"
        )
    )
)

if "%PYTHON_EXE%"=="" (
    echo.
    echo [Info] Python belum terinstall di komputer ini.
    echo Mendownload installer Python 3.11...
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -OutFile 'python_installer.exe'"
    
    if exist python_installer.exe (
        echo Menginstall Python... Mohon tunggu 1-2 menit.
        start /wait python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_doc=0
        del python_installer.exe >nul 2>&1
        
        if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
            set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"
            set "PATH=%LocalAppData%\Programs\Python\Python311;%LocalAppData%\Programs\Python\Python311\Scripts;!PATH!"
            echo [Sukses] Instalasi Python selesai dan berhasil dikonfigurasi!
        ) else (
            python --version >nul 2>&1
            if %errorlevel% equ 0 (
                set "PYTHON_EXE=python"
                echo [Sukses] Instalasi Python selesai!
            )
        )
    )
    
    if "%PYTHON_EXE%"=="" (
        echo.
        echo [ERROR] Gagal mendownload/menginstall Python secara otomatis.
        echo Silakan install Python dari https://www.python.org/downloads/ secara manual.
        goto FINISH
    )
)

echo [Info] Menggunakan Python: %PYTHON_EXE%

:: 4. Virtual Environment Setup
if not exist "venv\Scripts\activate.bat" (
    echo.
    echo Membuat Virtual Environment (venv)...
    "%PYTHON_EXE%" -m venv venv
)

echo Mengaktifkan Virtual Environment...
call venv\Scripts\activate.bat

echo.
echo Memeriksa dan menginstall dependensi (Pandas, Openpyxl, Playwright)...
python -m pip install --upgrade pip
pip install pandas openpyxl playwright

echo.
echo Mendownload browser Playwright (Chromium)...
playwright install chromium

echo.
echo ====================================================
echo Menjalankan Bot Otomatisasi Reject Anomali...
echo ====================================================
echo.
python reject_anomali.py

:FINISH
echo.
echo ====================================================
echo Program selesai atau dihentikan.
echo Tekan sembarang tombol untuk menutup jendela ini...
echo ====================================================
pause >nul

