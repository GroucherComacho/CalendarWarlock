@echo off
REM ============================================================
REM CalendarWarlock MSI Installer Build Script
REM Requires: WiX Toolset v3.x installed
REM ============================================================

echo.
echo ========================================
echo  CalendarWarlock MSI Builder
echo ========================================
echo.

REM Set WiX path - using short variable name and careful quoting
set "WIXBIN=C:\Program Files (x86)\WiX Toolset v3.14\bin"

REM Verify WiX is installed
if not exist "%WIXBIN%\candle.exe" goto :nowix

echo [1/3] WiX Toolset found
echo       %WIXBIN%

REM Navigate to installer directory
cd /d "%~dp0"

REM Clean previous build artifacts
if exist "*.wixobj" del /q *.wixobj
if exist "*.wixpdb" del /q *.wixpdb
if exist "CalendarWarlock.msi" del /q CalendarWarlock.msi

echo [2/3] Compiling Product.wxs...

REM Compile WiX source to object file
"%WIXBIN%\candle.exe" -nologo Product.wxs -out Product.wixobj
if errorlevel 1 goto :compilefail

echo [3/3] Linking CalendarWarlock.msi...

REM Link object file to MSI (with WixUI extension for the minimal UI)
"%WIXBIN%\light.exe" -nologo -ext WixUIExtension -out CalendarWarlock.msi Product.wixobj -spdb
if errorlevel 1 goto :linkfail

REM Clean up intermediate files
del /q *.wixobj 2>nul

echo.
echo ========================================
echo  BUILD SUCCESSFUL!
echo ========================================
echo.
echo Output: %cd%\CalendarWarlock.msi
echo.
goto :eof

:nowix
echo ERROR: WiX Toolset not found at:
echo        %WIXBIN%
echo.
echo Please install WiX Toolset v3.14 or update the WIXBIN variable
goto :failed

:compilefail
echo ERROR: Compilation failed!
goto :failed

:linkfail
echo ERROR: Linking failed!
goto :failed

:failed
echo.
echo ========================================
echo  BUILD FAILED
echo ========================================
echo.
exit /b 1
