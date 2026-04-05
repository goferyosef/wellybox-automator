@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

set "INSTALL_DIR=%USERPROFILE%\WellyBoxApp"
set "DESKTOP=%USERPROFILE%\Desktop"
set "SHORTCUT_NAME=WellyBox Automator"
set "SOURCE_DIR=%~dp0"

echo.
echo ============================================================
echo   WellyBox Automator - Installer
echo ============================================================
echo.

REM ── 1. Check Python ───────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found on this computer.
    echo.
    echo Please install Python 3.10 or newer from https://python.org
    echo IMPORTANT: During installation, check "Add Python to PATH"
    echo Then run this installer again.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%V in ('python --version 2^>^&1') do echo [OK] %%V found.

REM ── 2. Backup existing config ─────────────────────────────────
if exist "%INSTALL_DIR%\wellybox_config.json" (
    copy /Y "%INSTALL_DIR%\wellybox_config.json" "%TEMP%\wb_config_backup.json" >nul
    echo [OK] Existing settings backed up.
)

REM ── 3. Remove previous version files ─────────────────────────
if exist "%INSTALL_DIR%\wellybox_app.py" (
    del /F /Q "%INSTALL_DIR%\wellybox_app.py" >nul
    echo [OK] Previous version removed.
)

REM ── 4. Install to %USERPROFILE%\WellyBoxApp ──────────────────
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /Y "%SOURCE_DIR%wellybox_app.py"   "%INSTALL_DIR%\wellybox_app.py"   >nul
copy /Y "%SOURCE_DIR%requirements.txt"  "%INSTALL_DIR%\requirements.txt"  >nul
echo [OK] Files installed to %INSTALL_DIR%

REM ── 5. Restore settings ───────────────────────────────────────
if exist "%TEMP%\wb_config_backup.json" (
    copy /Y "%TEMP%\wb_config_backup.json" "%INSTALL_DIR%\wellybox_config.json" >nul
    del "%TEMP%\wb_config_backup.json" >nul
    echo [OK] Previous settings restored.
)

REM ── 6. Install Python dependencies ───────────────────────────
echo.
echo Installing required packages (this may take a minute)...
pip install -r "%INSTALL_DIR%\requirements.txt" --quiet --upgrade
if errorlevel 1 (
    echo [WARN] Some packages may not have installed correctly.
    echo        The app may still work — try running it first.
) else (
    echo [OK] All packages installed.
)

REM ── 7. Find pythonw.exe for the shortcut ─────────────────────
set "PYTHONW="
for /f "delims=" %%P in ('python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))" 2^>nul') do set "PYTHONW=%%P"
if not exist "%PYTHONW%" (
    REM Fall back to python.exe if pythonw not found
    for /f "delims=" %%P in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHONW=%%P"
)
echo [OK] Python executable: %PYTHONW%

REM ── 8. Remove old desktop shortcuts ──────────────────────────
echo.
for %%N in ("WellyBox Automator" "WellyBox Downloader" "WellyBox" "WellyBox App") do (
    if exist "%DESKTOP%\%%~N.lnk" (
        del /F /Q "%DESKTOP%\%%~N.lnk"
        echo [OK] Removed old shortcut: %%~N
    )
)

REM ── 9. Create new desktop shortcut ───────────────────────────
set "PS_TMP=%TEMP%\wb_shortcut.ps1"
(
    echo $ws = New-Object -ComObject WScript.Shell
    echo $s  = $ws.CreateShortcut('%DESKTOP%\%SHORTCUT_NAME%.lnk'^)
    echo $s.TargetPath       = '%PYTHONW%'
    echo $s.Arguments        = '"%INSTALL_DIR%\wellybox_app.py"'
    echo $s.WorkingDirectory = '%INSTALL_DIR%'
    echo $s.Description      = 'WellyBox Automator'
    echo $s.Save(^)
) > "%PS_TMP%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_TMP%" >nul 2>&1
del "%PS_TMP%" >nul

if exist "%DESKTOP%\%SHORTCUT_NAME%.lnk" (
    echo [OK] Desktop shortcut created: "%SHORTCUT_NAME%"
) else (
    echo [WARN] Could not create desktop shortcut automatically.
    echo        You can run the app manually from: %INSTALL_DIR%\wellybox_app.py
)

echo.
echo ============================================================
echo   Installation complete!
echo   Double-click "%SHORTCUT_NAME%" on the desktop to start.
echo ============================================================
echo.
pause
