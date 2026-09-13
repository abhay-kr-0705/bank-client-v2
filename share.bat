@echo off
title BankTech Client Live Preview (Cloudflare Tunnel)
cd /d "%~dp0"
color 0A
set PYTHONUNBUFFERED=1
echo ================================================================
echo Starting BankTech Live Client Preview via Cloudflare Tunnel...
echo ================================================================

:: Activate virtual environment if present
if exist "venv\Scripts\activate.bat" (
    echo [*] Using virtual environment (venv)...
    call venv\Scripts\activate.bat
)

python share.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Preview failed to start.
    echo [*] If dependencies are missing, please double-click setup.bat first.
    echo.
)
pause
