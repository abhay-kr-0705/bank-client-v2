@echo off
title BankTech Setup & Dependency Installer
cd /d "%~dp0"

echo ================================================================
echo   BANKTECH VALUATION OCR -- AUTOMATIC SETUP & INSTALLER
echo ================================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] ERROR: Python is not detected in your PATH.
    echo [*] Please install Python (version 3.10, 3.11, or 3.12 recommended) from python.org
    echo [*] CRITICAL: Ensure you check "Add python.exe to PATH" during installation!
    echo.
    pause
    exit /b 1
)

echo [+] Python found:
python --version
echo.

:: Create virtual environment if not already present
if not exist "venv" (
    echo [*] Creating isolated virtual environment (venv)...
    python -m venv venv
    if %ERRORLEVEL% NEQ 0 (
        echo [!] Failed to create venv. Proceeding with global python...
    ) else (
        echo [+] Virtual environment created successfully.
    )
)

:: Activate virtual environment
if exist "venv\Scripts\activate.bat" (
    echo [*] Activating virtual environment...
    call venv\Scripts\activate.bat
)

:: Upgrade pip and install dependencies
echo.
echo [*] Installing required Python libraries from requirements.txt...
echo [*] This includes FastAPI, PyMuPDF, OpenPyXL, RapidOCR, ONNX Runtime, and Pillow.
echo.
python -m pip install --upgrade pip
pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Warning: Some dependencies may have failed to install.
    echo     Please review the error messages above.
    pause
    exit /b 1
)

:: Create runtime folders
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs
if not exist "uploads\temp" mkdir uploads\temp

echo.
echo ================================================================
echo   [+] SETUP COMPLETED SUCCESSFULLY!
echo ================================================================
echo.
echo You can now launch the application anytime by double-clicking:
echo   --^> start.bat (Runs locally at http://127.0.0.1:8000)
echo   --^> share.bat (Runs locally + generates Cloudflare live link)
echo.
pause
