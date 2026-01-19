@echo off
REM ============================================================
REM CalendarWarlock MSI Installer Build Script
REM Requires: WiX Toolset v3.x installed
REM ============================================================

setlocal EnableDelayedExpansion

echo.
echo ========================================
echo  CalendarWarlock MSI Builder
echo ========================================
echo.

REM Set WiX path
set "WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.14\bin"

REM Verify WiX is installed
if not exist "%WIX_PATH%\candle.exe" (
    echo ERROR: WiX Toolset not found at %WIX_PATH%
    echo Please install WiX Toolset v3.14 or update the WIX_PATH variable
    goto :error
)

echo [1/3] WiX Toolset found at: %WIX_PATH%

REM Navigate to installer directory
cd /d "%~dp0"

REM Clean previous build artifacts
if exist "*.wixobj" del /q *.wixobj
if exist "*.wixpdb" del /q *.wixpdb
if exist "CalendarWarlock.msi" del /q CalendarWarlock.msi

echo [2/3] Compiling Product.wxs...

REM Compile WiX source to object file
"%WIX_PATH%\candle.exe" -nologo Product.wxs -out Product.wixobj
if errorlevel 1 (
    echo ERROR: Compilation failed!
    goto :error
)

echo [3/3] Linking CalendarWarlock.msi...

REM Link object file to MSI (with WixUI extension for the minimal UI)
"%WIX_PATH%\light.exe" -nologo -ext WixUIExtension -out CalendarWarlock.msi Product.wixobj -spdb
if errorlevel 1 (
    echo ERROR: Linking failed!
    goto :error
)

REM Clean up intermediate files
del /q *.wixobj 2>nul

echo.
echo ========================================
echo  BUILD SUCCESSFUL!
echo ========================================
echo.
echo Output: %cd%\CalendarWarlock.msi
echo.

goto :end

:error
echo.
echo ========================================
echo  BUILD FAILED
echo ========================================
echo.
exit /b 1

:end
endlocal
