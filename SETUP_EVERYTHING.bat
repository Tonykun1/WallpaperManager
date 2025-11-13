@echo off
chcp 65001 >nul
title Wallpaper Engine Manager - Installation
color 0A

echo ═══════════════════════════════════════════════════════
echo   🎨 Wallpaper Engine Manager - Auto Installation
echo ═══════════════════════════════════════════════════════
echo.

REM Check if Python is installed
echo [1/4] Checking if Python is installed...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python not installed!
    echo.
    echo 📥 Please install Python from:
    echo    https://www.python.org/downloads/
    echo.
    echo ⚠️  Important! Check "Add Python to PATH" during installation!
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b 1
) else (
    for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
    echo ✅ Found: %PYTHON_VERSION%
)
echo.

REM Check if pip works
echo [2/4] Checking if pip works...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ pip not working!
    echo.
    echo Trying to fix...
    python -m ensurepip --default-pip
    python -m pip install --upgrade pip
)
echo ✅ pip working!
echo.

REM Install required packages
echo [3/4] Installing required packages...
echo    (This may take a minute or two...)
echo.

pip install --quiet --upgrade pip
pip install --quiet psutil pywin32

if %errorlevel% neq 0 (
    echo.
    echo ❌ Error installing packages!
    echo.
    echo Trying again with older versions...
    pip install psutil==5.9.0 pywin32==305
)

echo ✅ Packages installed successfully!
echo.

REM Run installation script
echo [4/4] Running installation script...
echo.
python install.py

echo.
echo ═══════════════════════════════════════════════════════
echo   ✅ Installation complete!
echo ═══════════════════════════════════════════════════════
echo.
pause