@echo off
setlocal

echo Getting application version...

for /f "tokens=2 delims== " %%A in (
    'findstr /b "APP_VERSION =" import_course_assignments.py'
) do (
    set VERSION=%%~A
)

if not defined VERSION (
    echo.
    echo ERROR: Could not determine application version.
    exit /b 1
)

echo Building SNHU Assignment Importer v%VERSION%...

echo.
echo Cleaning previous build...

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist snhu-assignment-importer.spec del snhu-assignment-importer.spec

echo.
echo Building application...

.venv\Scripts\python.exe -m PyInstaller ^
    --name snhu-assignment-importer ^
    --add-data ".venv\Lib\site-packages\playwright\driver\package\.local-browsers;playwright\driver\package\.local-browsers" ^
    import_course_assignments.py

if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    exit /b 1
)

echo.
echo Copying Excel template...

copy /y snhu_template.xlsx dist\snhu-assignment-importer\snhu_template.xlsx >nul

if errorlevel 1 (
    echo.
    echo FAILED TO COPY EXCEL TEMPLATE.
    exit /b 1
)

echo.
echo Creating release package...

powershell -NoProfile -Command ^
    "Compress-Archive -Path 'dist\snhu-assignment-importer\*' -DestinationPath 'dist\snhu-assignment-importer-v%VERSION%-windows.zip' -Force"

if errorlevel 1 (
    echo.
    echo FAILED TO CREATE RELEASE PACKAGE.
    exit /b 1
)

echo.
echo Build completed successfully.
echo.
echo Application:
echo dist\snhu-assignment-importer
echo.
echo Release package:
echo dist\snhu-assignment-importer-v%VERSION%-windows.zip

endlocal