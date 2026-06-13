@echo off
cd /d "%~dp0"
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Could not launch. Make sure Python is installed and added to PATH.
    pause
)
