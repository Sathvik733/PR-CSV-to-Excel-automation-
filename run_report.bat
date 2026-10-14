@echo off
cd /d "%~dp0"
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing project dependencies...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Failed to install dependencies. Check your internet connection and Python installation.
        pause
        exit /b 1
    )
)
python procurement_report.py
if errorlevel 1 (
    pause
    exit /b 1
)
pause
