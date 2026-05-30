@echo off
REM Outlook PST/OST Tool - setup for Windows.
REM Creates a local .venv and installs the dependencies from requirements.txt.
cd /d "%~dp0"

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

echo Using:
python --version
echo.

REM Create the virtual environment if it does not exist yet.
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment in .venv ...
    python -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip

echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo WARNING: pip install failed.
    echo libpff-python is the only dependency that may need extra steps:
    echo   1. conda install -c conda-forge libpff-python
    echo   2. See https://github.com/libyal/libpff for a manual build
    echo xhtml2pdf and pikepdf ^(PDF export^) install as normal wheels.
    echo.
)

echo.
echo Setup complete. Launch the app with:  run.bat
echo.
pause
