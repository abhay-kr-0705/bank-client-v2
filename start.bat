@echo off
title BankTech Valuation OCR & Excel Generator
cd /d "%~dp0"

echo ================================================================
echo   BANKTECH VALUATION OCR & EXCEL REPORT GENERATOR
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

:: 4. Verify core dependencies (FastAPI, Uvicorn, RapidOCR, PyMuPDF, OpenPyXL)
python -c "import fastapi, uvicorn, openpyxl, fitz, rapidocr" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [*] First-time launch: Installing all required dependencies automatically...
    echo [*] Installing FastAPI, RapidOCR, ONNX Runtime, OpenPyXL, PyMuPDF...
    echo [*] Please wait 1-2 minutes (internet connection required for first-time install)...
    echo.
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [!] Failed to install some dependencies. Please check your internet connection.
        pause
        exit /b 1
    )
    echo.
    echo [+] All dependencies installed successfully!
    echo.
)

:: 5. Ensure required runtime folders exist
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs
if not exist "uploads\temp" mkdir uploads\temp

:: 6. Launch application
echo [*] Starting server & opening web interface...
python run.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Server stopped unexpectedly.
    pause
)
