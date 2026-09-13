@echo off
title BankTech Client Live Preview via Cloudflare Tunnel
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
    echo [*] Please install Python 3.10, 3.11, or 3.12 from python.org
    echo [*] IMPORTANT: Ensure you check the box: Add python.exe to PATH during installation!
    echo.
    pause
    exit /b 1
)

:: 2. Check or create Virtual Environment
if exist "venv\Scripts\activate.bat" (
    echo [*] Activating virtual environment venv...
    call venv\Scripts\activate.bat
)

:: 3. Test if core dependencies are installed
python -c "import fastapi, uvicorn, openpyxl, fitz, rapidocr" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [*] First-time setup: Installing required dependencies automatically...
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
        echo [!] Failed to install dependencies. Please check your internet connection.
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

:: 5. Launch Cloudflare tunnel and local server
echo [*] Launching application and creating live Cloudflare preview link...
python share.py

echo.
echo [*] Live preview session closed.
pause
