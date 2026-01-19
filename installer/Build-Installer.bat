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

REM Navigate to installer directory first
cd /d "%~dp0"

REM Check for WiX in standard location
set WIXDIR=C:\Program Files (x86)\WiX Toolset v3.14\bin
if exist "%WIXDIR%\candle.exe" goto :found

REM Try alternate location
set WIXDIR=C:\Program Files\WiX Toolset v3.14\bin
if exist "%WIXDIR%\candle.exe" goto :found

echo ERROR: WiX Toolset not found.
echo Please install WiX Toolset v3.14
goto :failed

:found
echo [1/3] WiX Toolset found

REM Clean previous build artifacts
del /q *.wixobj 2>nul
del /q *.wixpdb 2>nul
del /q CalendarWarlock.msi 2>nul

echo [2/3] Compiling Product.wxs...
call "%WIXDIR%\candle.exe" -nologo Product.wxs -out Product.wixobj
if errorlevel 1 goto :compilefail

echo [3/3] Linking CalendarWarlock.msi...
call "%WIXDIR%\light.exe" -nologo -ext WixUIExtension -out CalendarWarlock.msi Product.wixobj -spdb
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
goto :done

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
pause
exit /b 1

:done
pause
