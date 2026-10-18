@echo off
title FileOpener v2.0 - Universal Extractor ^& Converter
echo ============================================
echo   FileOpener v2.0 - Universal File Tool
echo ============================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python ma la helin! Degso Python.org
    pause
    exit /b 1
)

REM Install dependencies if needed
echo [*] Checking libraries...
python -c "import tkinterdnd2" >nul 2>&1
if errorlevel 1 (
    echo [*] Installing required libraries...
    python -m pip install py7zr rarfile Pillow pandas openpyxl python-docx fpdf2 pypdf tkinterdnd2 -q
    echo [*] Libraries installed!
)

echo [*] Launching FileOpener...
python "%~dp0extractor.py" %*
if errorlevel 1 (
    echo.
    echo [ERROR] FileOpener ayaa khalad ku dhacay!
    pause
)
