@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Chua co .venv. Hay chay install.bat truoc.
    pause
    exit /b 1
)

.\.venv\Scripts\python.exe -m compileall -q app.py gk_pilepro presence_server.py tools pyi_hooks
if errorlevel 1 (
    echo.
    echo TEST LOI: source Python bi loi compile.
    pause
    exit /b 1
)

echo.
echo TEST OK: source Python compile duoc.
pause

