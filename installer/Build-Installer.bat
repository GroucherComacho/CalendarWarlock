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

REM Use Windows environment variable for Program Files (x86)
set "WIXBIN=%ProgramFiles(x86)%\WiX Toolset v3.14\bin"

REM Check if WiX exists
if not exist "%WIXBIN%\candle.exe" goto :nowix

echo [1/3] WiX Toolset found

REM Clean previous build artifacts
del /q *.wixobj 2>nul
del /q *.wixpdb 2>nul
del /q CalendarWarlock.msi 2>nul

echo [2/3] Compiling Product.wxs...
"%WIXBIN%\candle.exe" -nologo Product.wxs -out Product.wixobj
if errorlevel 1 goto :compilefail

echo [3/3] Linking CalendarWarlock.msi...
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
goto :done

:nowix
echo ERROR: WiX Toolset not found.
echo Looked in: %WIXBIN%
echo Please install WiX Toolset v3.14
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
pause
exit /b 1

:done
pause
