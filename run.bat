@echo off
cd /d "%~dp0"
title Smart Farm Delivery System

set "PY_EXE=python"
if exist "C:\Users\jeema\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe" (
    set "PY_EXE=C:\Users\jeema\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe"
)

echo ====================================================================
echo   [SmartFarm] Gangbuk-gu Delivery Ledger System
echo   Starting server and Cloudflare Mobile Tunnel...
echo   Web browser will open automatically at http://127.0.0.1:8000
echo.
echo   Press Ctrl+C or close this window to stop the server.
echo ====================================================================
echo.

"%PY_EXE%" app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo --------------------------------------------------------------------
    echo [ERROR] Server terminated with error code %ERRORLEVEL%.
    echo --------------------------------------------------------------------
    pause
)
