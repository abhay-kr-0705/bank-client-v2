@echo off
title BankTech Valuation OCR and Excel Generator
cd /d "%~dp0"

echo ================================================================
echo   BANKTECH VALUATION OCR AND EXCEL REPORT GENERATOR
echo ================================================================
echo.

:: 1. Check Python installation
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] CRITICAL ERROR: Python is not installed or not in system PATH!
    echo [*] Please install Python 3.10, 3.11, or 3.12 from python.org
    echo [*] Ensure you check the box: Add python.exe to PATH during installation!
    echo.
    pause
    exit /b 1
)

:: 2. Check if venv exists and activate it if present
if exist "venv\Scripts\activate.bat" (
    echo [*] Activating virtual environment venv...
    call venv\Scripts\activate.bat
)

:: 3. Test if core libraries are installed
python -c "import fastapi, uvicorn, openpyxl, fitz, rapidocr" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [*] First-time setup: Core dependencies are missing.
    echo [*] Creating isolated virtual environment and installing packages...
    echo [*] Please wait 1 to 2 minutes...
    echo.
    if not exist "venv\Scripts\activate.bat" (
        python -m venv venv
        if exist "venv\Scripts\activate.bat" (
            call venv\Scripts\activate.bat
        )
    )
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [!] Dependency installation encountered an issue.
        pause
        exit /b 1
    )
    echo.
    echo [+] Dependencies installed successfully!
    echo.
)

:: 4. Ensure required runtime folders exist
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs
if not exist "uploads\temp" mkdir uploads\temp

:: 5. Launch application
echo [*] Starting local server and opening web interface...
echo [*] Web URL: http://127.0.0.1:8000
echo.
python run.py

echo.
echo [*] Server process ended.
pause
