@echo off
REM Outlook PST/OST Tool - launcher for Windows.
REM Uses the local .venv if present, otherwise the system Python.
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py
) else (
    python main.py
)

if %errorlevel% neq 0 (
    echo.
    echo Failed to start. Run install.bat first, and make sure Python is in PATH.
    pause
)
