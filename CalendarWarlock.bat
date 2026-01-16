@echo off
:: CalendarWarlock Launcher
:: This batch file launches the CalendarWarlock PowerShell GUI application

:: Change to the directory where this batch file is located
cd /d "%~dp0"

:: Launch PowerShell with the launcher script
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Start-CalendarWarlock.ps1"

:: Pause if there was an error
if %errorlevel% neq 0 (
    echo.
    echo An error occurred. Press any key to exit...
    pause >nul
)
