@echo off
echo ===================================================
echo             BUILDING CONVERT_MD EXECUTABLE
echo ===================================================
echo.

rem Check if PyInstaller is available
pyinstaller --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] PyInstaller is not installed or not in PATH.
    echo Please install it using: pip install pyinstaller
    pause
    exit /b 1
)

echo Packaging convert_md.py with bundled binaries...
pyinstaller --onefile ^
    --add-data "bin/pandoc.exe;bin" ^
    --add-data "bin/tidy.exe;bin" ^
    --add-data "bin/typst.exe;bin" ^
    convert_md.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo [SUCCESS] Build completed! Executable is located in dist\convert_md.exe
    rem Copy to root directory for convenience
    copy /y dist\convert_md.exe .
    echo Copied convert_md.exe to root directory.
) else (
    echo.
    echo [ERROR] Build failed.
)

pause
