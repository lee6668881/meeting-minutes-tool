@echo off
cd /d %~dp0
echo Starting Meeting Minutes Agent...
echo.
py main.py
if errorlevel 1 (
    echo.
    echo ========================================
    echo [Error] Program exited with errors.
    echo Please check the log above.
    echo ========================================
    pause
)
