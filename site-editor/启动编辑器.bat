@echo off
rem Homepage Editor launcher (ASCII only - avoid cmd encoding issues)
cd /d "%~dp0"
set "PY=C:\Users\zhouy\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=python"
echo ============================================
echo   Homepage Editor is starting...
echo   Browser will open http://127.0.0.1:8765
echo   Keep this window open while editing.
echo   (Close it to stop the editor)
echo ============================================
start "" /min cmd /c "timeout /t 2 >nul & start http://127.0.0.1:8765/"
"%PY%" editor_server.py
pause
