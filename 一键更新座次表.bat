@echo off
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
chcp 65001 >nul
echo ==================================================
echo     Wedding Seating Chart Updater
echo ==================================================
echo.

cd /d "%~dp0"

REM Check Python - try py first, then python
py --version >nul 2>&1
if %errorlevel% neq 0 (
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: Python not found. Please install Python first.
        echo Download from: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    ) else (
        set PY_CMD=python
    )
) else (
    set PY_CMD=py
)

REM Install dependencies if needed
%PY_CMD% -c "import pandas; import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing dependencies...
    %PY_CMD% -m pip install pandas openpyxl -q
    echo Dependencies installed.
    echo.
)

REM Run the update script
echo Reading Excel and generating HTML...
echo.
%PY_CMD% update_seating.py

echo.
echo Press any key to exit...
pause >nul
