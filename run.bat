@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Chua co moi truong .venv.
    echo Hay bam install.bat lan dau truoc.
    pause
    exit /b 1
)

if not exist "app.py" (
    echo LOI: Khong thay app.py trong thu muc nay.
    pause
    exit /b 1
)

.\.venv\Scripts\python.exe app.py
pause
