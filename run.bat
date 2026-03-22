@echo off
cd /d "%~dp0"
python main.py
if %errorlevel% neq 0 (
    echo.
    echo Failed to start. Make sure Python is installed and in PATH.
    pause
)
