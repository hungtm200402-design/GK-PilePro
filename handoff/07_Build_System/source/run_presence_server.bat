@echo off
setlocal
cd /d "%~dp0"
"%~dp0\.venv\Scripts\python.exe" "%~dp0presence_server.py" --host 0.0.0.0 --port 8765
endlocal
