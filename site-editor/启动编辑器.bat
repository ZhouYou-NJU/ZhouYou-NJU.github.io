@echo off
rem Homepage Editor launcher (ASCII only - avoid cmd encoding issues)
cd /d "%~dp0"
set "PY=C:\Users\zhouy\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=python"
echo ============================================
echo   Homepage Editor is starting...
echo   The browser will open automatically.
echo   Keep this window open while editing.
echo   (Close it to stop the editor)
echo ============================================
"%PY%" editor_server.py
pause
