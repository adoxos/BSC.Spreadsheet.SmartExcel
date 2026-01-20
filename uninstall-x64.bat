@echo off
REM SmartExcel Add-in Uninstallation Script (x64)
REM This script removes the SmartExcel add-in files from the local user's add-ins folder

echo ====================================
echo SmartExcel Add-in Uninstaller (x64)
echo ====================================
echo.

REM Prompt for organization name
set /p "ORG_NAME=Enter your organization name (same as used during installation, or press Enter for default 'Add-ins'): "

REM If no organization name provided, use default
if "%ORG_NAME%"=="" (
    set "TARGET_DIR=%LOCALAPPDATA%\Add-ins"
) else (
    set "TARGET_DIR=%LOCALAPPDATA%\%ORG_NAME%\Add-ins"
)

echo.
echo This will remove SmartExcel add-in files from:
echo %TARGET_DIR%
echo.
set /p CONFIRM="Are you sure you want to continue? (Y/N): "

if /i not "%CONFIRM%"=="Y" (
    echo Uninstallation cancelled.
    pause
    exit /b 0
)

echo.
echo Removing SmartExcel add-in files...

REM Delete the files
if exist "%TARGET_DIR%\SmartExcel64.xll" (
    del /F /Q "%TARGET_DIR%\SmartExcel64.xll"
    echo - Removed SmartExcel64.xll
)

if exist "%TARGET_DIR%\SmartExcel.dll" (
    del /F /Q "%TARGET_DIR%\SmartExcel.dll"
    echo - Removed SmartExcel.dll
)

if exist "%TARGET_DIR%\SmartExcel64.dna" (
    del /F /Q "%TARGET_DIR%\SmartExcel64.dna"
    echo - Removed SmartExcel64.dna
)

if exist "%TARGET_DIR%\Microsoft.Office.Interop.Excel.dll" (
    del /F /Q "%TARGET_DIR%\Microsoft.Office.Interop.Excel.dll"
    echo - Removed Microsoft.Office.Interop.Excel.dll
)

if exist "%TARGET_DIR%\SmartExcel.deps.json" (
    del /F /Q "%TARGET_DIR%\SmartExcel.deps.json"
    echo - Removed SmartExcel.deps.json
)

if exist "%TARGET_DIR%\SmartExcel.runtimeconfig.json" (
    del /F /Q "%TARGET_DIR%\SmartExcel.runtimeconfig.json"
    echo - Removed SmartExcel.runtimeconfig.json
)

echo.
echo ====================================
echo Uninstallation completed!
echo ====================================
echo.
echo Note: If the add-in is still loaded in Excel, you may need to:
echo 1. Close all Excel instances
echo 2. Open Excel
echo 3. Go to File ^> Options ^> Add-ins
echo 4. Select "Excel Add-ins" from the Manage dropdown and click "Go..."
echo 5. Uncheck SmartExcel and click OK
echo.
pause
