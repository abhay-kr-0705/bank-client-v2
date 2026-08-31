@echo off
title BankTech Cloudflare Public Tunnel
echo ================================================================
echo   Creating Public HTTPS Link for BankTech Valuation OCR...
echo ================================================================
echo.

if exist "D:\DOWNLOADS\cloudflared-windows-amd64.exe" (
    "D:\DOWNLOADS\cloudflared-windows-amd64.exe" tunnel --url http://127.0.0.1:8000
) else if exist "cloudflared.exe" (
    cloudflared.exe tunnel --url http://127.0.0.1:8000
) else (
    echo Cloudflared executable not found in D:\DOWNLOADS\
    pause
)
