@echo off
title BankTech Client Live Preview (Cloudflare Tunnel)
cd /d "%~dp0"
color 0A
set PYTHONUNBUFFERED=1

echo ================================================================
echo   BANKTECH LIVE CLIENT PREVIEW VIA CLOUDFLARE TUNNEL
echo ================================================================
echo.

:: 1. Check Python installation
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] CRITICAL ERROR: Python is not installed or not in your system PATH!
    echo [*] Please download and install Python (3.10, 3.11, or 3.12) from:
    echo     https://www.python.org/downloads/
    echo [*] IMPORTANT: Ensure you check the box "Add python.exe to PATH" during install!
    echo.
    pause
    exit /b 1
)

:: 2. Check or create Virtual Environment
if not exist "venv\Scripts\activate.bat" (
    echo [*] First-time setup detected on this PC!
    echo [*] Creating isolated virtual environment (venv)...
    python -m venv venv
)

:: 3. Activate venv if available
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

:: 4. Verify core dependencies
python -c "import fastapi, uvicorn, openpyxl, fitz, rapidocr" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [*] Installing required dependencies automatically...
    echo [*] Please wait 1-2 minutes...
    echo.
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [!] Failed to install dependencies. Please check your internet connection.
        pause
        exit /b 1
    )
    echo [+] All dependencies installed successfully!
    echo.
)

:: 5. Ensure required runtime folders exist
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs
if not exist "uploads\temp" mkdir uploads\temp

:: 6. Launch Cloudflare tunnel and local server
echo [*] Launching application and creating live Cloudflare preview link...
python share.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Preview stopped unexpectedly.
    pause
)
