@echo off
chcp 65001 >nul
echo ============================================
echo  WellyBox Downloader - Setup ^& Launch
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

echo Installing dependencies...
pip install -r "%~dp0requirements.txt" --quiet

echo.
echo Launching app...
python "%~dp0wellybox_app.py"
pause
