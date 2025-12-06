@echo off
setlocal

:: Wrapper for AdvancedRenamer.ps1
:: Usage: rename_tool.bat [excel_file] [options]

set SCRIPT_DIR=%~dp0
set PS_SCRIPT="%SCRIPT_DIR%AdvancedRenamer.ps1"

:: Check if PowerShell is available
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] PowerShell is not found in PATH.
    pause
    exit /b 1
)

:: Forward arguments
:: Mapping common flags to PS params
:: Usage: rename_tool.bat input.csv /preview /undo

:: We will just pass all args to PowerShell. 

if "%~1"=="/?" (
    echo Usage: rename_tool.bat [inputfile.csv] [options]
    echo.
    echo Options:
    echo   -targetdir "path"    : Files directory
    echo   -map "0,1"           : Columns for Old,New
    echo   -conflict [skip|overwrite|autonumber]
    echo   -subfolders          : Search subfolders
    echo   -dryrun              : Preview only
    echo   -undo                : Undo last run
    echo.
    echo Example:
    echo   rename_tool.bat files.csv -dryrun
    exit /b 0
)

powershell -NoProfile -ExecutionPolicy Bypass -File %PS_SCRIPT% %*

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Script execution failed.
    pause
)
