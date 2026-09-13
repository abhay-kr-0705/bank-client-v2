@echo off
title BankTech Valuation OCR & Excel Generator
cd /d "%~dp0"

echo ================================================================
echo   Starting BankTech Valuation OCR & Excel Generator...
echo ================================================================

:: Activate virtual environment if present
if exist "venv\Scripts\activate.bat" (
    echo [*] Using virtual environment (venv)...
    call venv\Scripts\activate.bat
)

python run.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Server failed to start.
    echo [*] If dependencies are missing, please double-click setup.bat first.
    echo.
)
pause
