@echo off
setlocal

cd /d "%~dp0"

echo ========================================
echo TopCV Export TUI - Setup and Run
echo ========================================
echo.

set "PYTHON_CMD=python"
python --version >nul 2>&1
if errorlevel 1 (
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python is not installed or not in PATH.
        echo Please install Python 3, then run this file again.
        echo.
        pause
        exit /b 1
    ) else (
        set "PYTHON_CMD=py -3"
    )
)

if not exist ".venv\Scripts\activate.bat" (
    echo [SETUP] Creating virtual environment...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        echo.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment.
    echo.
    pause
    exit /b 1
)

echo [SETUP] Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [WARN] Failed to upgrade pip. Continuing...
)

if exist "requirements.txt" (
    echo [SETUP] Installing Python dependencies...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies from requirements.txt.
        echo.
        pause
        exit /b 1
    )
) else (
    echo [WARN] requirements.txt not found. Skipping dependency install.
)

echo.
echo [RUN] Starting TopCV Export TUI...
echo.
python "scripts\topcv_export_tui.py"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] Script exited with error code %EXIT_CODE%.
)

echo.
pause
exit /b %EXIT_CODE%
