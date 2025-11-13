@echo off
echo ===============================================
echo BOQ-Tools Executable Builder
echo ===============================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher
    pause
    exit /b 1
)

:: Show Python version
echo Python version:
python --version
echo.

:: Run the build script
echo Starting build process...
echo.
python build_exe.py

:: Check if build was successful
if %errorlevel% equ 0 (
    echo.
    echo ===============================================
    echo BUILD COMPLETED SUCCESSFULLY!
    echo ===============================================
    echo.
    echo The executable has been created in the 'dist' folder
    echo You can now distribute BOQ-Tools.exe
    echo.
    if exist "dist\BOQ-Tools.exe" (
        echo Executable location: dist\BOQ-Tools.exe
        for %%I in ("dist\BOQ-Tools.exe") do echo File size: %%~zI bytes
    )
    echo.
    echo Press any key to open the dist folder...
    pause >nul
    explorer dist
) else (
    echo.
    echo ===============================================
    echo BUILD FAILED!
    echo ===============================================
    echo.
    echo Check the build.log file for error details
    echo.
    pause
) 