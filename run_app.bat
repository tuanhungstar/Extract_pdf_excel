@echo off
title AI PDF-to-Excel Converter
echo [INFO] Initializing portable Python environment...
if not exist "python-embed\python.exe" (
    echo [ERROR] Python environment not found! Please run 'setup_env.bat' first.
    pause
    exit /b 1
)

echo [INFO] Launching PyQt6 application...
start "" "python-embed\pythonw.exe" main.py
exit /b 0
