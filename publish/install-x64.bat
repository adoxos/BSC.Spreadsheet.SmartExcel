@echo off
REM SmartExcel Add-in Installation Script (x64)
REM This script copies the SmartExcel add-in files to the local user's add-ins folder

echo ====================================
echo SmartExcel Add-in Installer (x64)
echo ====================================
echo.

REM Define the target directory
set "TARGET_DIR=%LOCALAPPDATA%\Statkraft AS\Add-ins"

REM Define the source directory (same directory as this script)
set "SOURCE_DIR=%~dp0"

REM Check if source files exist
if not exist "%SOURCE_DIR%SmartExcel64.xll" (
    echo ERROR: SmartExcel64.xll not found in %SOURCE_DIR%
    echo Please ensure this script is run from the publish folder.
    pause
    exit /b 1
)

REM Create target directory if it doesn't exist
echo Creating target directory...
if not exist "%TARGET_DIR%" (
    mkdir "%TARGET_DIR%"
    if errorlevel 1 (
        echo ERROR: Failed to create directory: %TARGET_DIR%
        pause
        exit /b 1
    )
    echo Directory created: %TARGET_DIR%
) else (
    echo Directory already exists: %TARGET_DIR%
)

echo.
echo Copying SmartExcel add-in files...

REM Copy the main XLL file (64-bit)
copy /Y "%SOURCE_DIR%SmartExcel64.xll" "%TARGET_DIR%\"
if errorlevel 1 (
    echo ERROR: Failed to copy SmartExcel64.xll
    pause
    exit /b 1
)
echo - Copied SmartExcel64.xll

REM Copy the DLL file
copy /Y "%SOURCE_DIR%SmartExcel.dll" "%TARGET_DIR%\"
if errorlevel 1 (
    echo ERROR: Failed to copy SmartExcel.dll
    pause
    exit /b 1
)
echo - Copied SmartExcel.dll

REM Copy the DNA file
copy /Y "%SOURCE_DIR%SmartExcel64.dna" "%TARGET_DIR%\"
if errorlevel 1 (
    echo ERROR: Failed to copy SmartExcel64.dna
    pause
    exit /b 1
)
echo - Copied SmartExcel64.dna

REM Copy dependencies if they exist
if exist "%SOURCE_DIR%Microsoft.Office.Interop.Excel.dll" (
    copy /Y "%SOURCE_DIR%Microsoft.Office.Interop.Excel.dll" "%TARGET_DIR%\"
    echo - Copied Microsoft.Office.Interop.Excel.dll
)

if exist "%SOURCE_DIR%SmartExcel.deps.json" (
    copy /Y "%SOURCE_DIR%SmartExcel.deps.json" "%TARGET_DIR%\"
    echo - Copied SmartExcel.deps.json
)

if exist "%SOURCE_DIR%SmartExcel.runtimeconfig.json" (
    copy /Y "%SOURCE_DIR%SmartExcel.runtimeconfig.json" "%TARGET_DIR%\"
    echo - Copied SmartExcel.runtimeconfig.json
)

echo.
echo ====================================
echo Installation completed successfully!
echo ====================================
echo.
echo Files installed to: %TARGET_DIR%
echo.
echo To use the add-in in Excel:
echo 1. Open Excel
echo 2. Go to File ^> Options ^> Add-ins
echo 3. Select "Excel Add-ins" from the Manage dropdown and click "Go..."
echo 4. Click "Browse..." and navigate to: %TARGET_DIR%
echo 5. Select SmartExcel64.xll and click OK
echo 6. Make sure the checkbox next to SmartExcel is checked
echo.
echo The "SmartExcel" tab should now appear in your Excel ribbon.
echo.
pause
