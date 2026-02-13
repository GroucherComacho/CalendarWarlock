@echo off
:: CalendarWarlock Launcher
:: This batch file launches the CalendarWarlock PowerShell GUI application

:: Change to the directory where this batch file is located
cd /d "%~dp0"

:: Launch PowerShell with the launcher script (hidden window for seamless GUI launch)
:: Using RemoteSigned instead of Bypass for security (MEDIUM-006)
start "" /B powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "Start-CalendarWarlock.ps1"
