@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Chua co .venv. Hay chay install.bat truoc.
    pause
    exit /b 1
)

.\.venv\Scripts\python.exe -m py_compile app.py
if errorlevel 1 (
    echo.
    echo TEST LOI: app.py bi loi compile.
    pause
    exit /b 1
)

echo.
echo TEST OK: app.py compile duoc.
pause
