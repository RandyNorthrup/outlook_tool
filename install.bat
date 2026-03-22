@echo off
echo ========================================
echo  Outlook PST/OST Tool - Setup
echo ========================================
echo.

REM Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Install dependencies
echo Installing dependencies...
pip install libpff-python
if %errorlevel% neq 0 (
    echo.
    echo WARNING: pip install failed for libpff-python.
    echo.
    echo Alternative install methods:
    echo   1. conda install -c conda-forge libpff-python
    echo   2. See https://github.com/libyal/libpff for manual build
    echo.
    echo The app will still launch but cannot open PST/OST files
    echo until libpff-python is installed.
    echo.
)

echo.
echo Setup complete. Run the app with:
echo   python main.py
echo.
pause
