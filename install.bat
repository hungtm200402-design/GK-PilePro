@echo off
cd /d "%~dp0"

echo ============================================
echo  TOOL KL V23.1 - CAI DAT LAN DAU
echo ============================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Dang tao moi truong .venv...
    py -3.12 -m venv .venv
    if errorlevel 1 (
        echo Khong thay Python 3.12 bang py -3.12, thu bang python...
        python -m venv .venv
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo LOI: Khong tao duoc .venv.
    echo Hay cai Python 3.12 roi chay lai install.bat.
    pause
    exit /b 1
)

echo.
echo Dang nang cap pip...
.\.venv\Scripts\python.exe -m pip install --upgrade pip

echo.
echo Dang cai thu vien can thiet...
.\.venv\Scripts\python.exe -m pip install --no-cache-dir --force-reinstall pillow python-dotenv openpyxl google-genai

echo.
echo ============================================
echo  CAI DAT XONG
echo  Lan sau chi can bam run.bat
echo ============================================
pause
