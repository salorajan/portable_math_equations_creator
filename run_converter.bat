@echo off
setlocal enabledelayedexpansion
title Equation Converter CLI

echo ===================================================
echo             EQUATION CONVERTER UTILITY
echo ===================================================
echo.

rem Check if a file was dragged and dropped or passed as an argument
set "INPUT_FILE=%~1"

if "%INPUT_FILE%"=="" (
    set /p "INPUT_FILE=Enter the path of the file to convert: "
)

rem Remove surrounding quotes if any
set "INPUT_FILE=!INPUT_FILE:"=!"

if not exist "!INPUT_FILE!" (
    echo [ERROR] Input file "!INPUT_FILE!" does not exist.
    pause
    exit /b 1
)

echo.
echo Input file selected: !INPUT_FILE!
echo.
echo Select the target output format:
echo  [1] MS Word (.docx)
echo  [2] HTML (.html)
echo  [3] PDF (.pdf)
echo  [4] Markdown (.md)
echo.
set /p "FORMAT_CHOICE=Enter choice (1-4): "

if "!FORMAT_CHOICE!"=="1" (
    set "FORMAT=word"
) else if "!FORMAT_CHOICE!"=="2" (
    set "FORMAT=html"
) else if "!FORMAT_CHOICE!"=="3" (
    set "FORMAT=pdf"
) else if "!FORMAT_CHOICE!"=="4" (
    set "FORMAT=md"
) else (
    echo [ERROR] Invalid choice.
    pause
    exit /b 1
)

rem Resolve execution runner (prioritize bin folder EXE, then root level EXE, then python scripts)
if exist "bin\convert_md.exe" (
    set "RUNNER=bin\convert_md.exe"
) else if exist "convert_md.exe" (
    set "RUNNER=convert_md.exe"
) else if exist "dist\convert_md.exe" (
    set "RUNNER=dist\convert_md.exe"
) else if exist "references\convert_md.py" (
    set "RUNNER=python references\convert_md.py"
) else (
    set "RUNNER=python convert_md.py"
)

echo.
echo Running conversion: !RUNNER! "!INPUT_FILE!" !FORMAT!
echo.

!RUNNER! "!INPUT_FILE!" !FORMAT!

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUCCESS] Conversion completed successfully!
) else (
    echo.
    echo [ERROR] Conversion failed.
)

pause
