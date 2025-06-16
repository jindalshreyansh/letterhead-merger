@echo off
setlocal
cls

REM =================================================================
REM  Build Script for PDF Letterhead Merger (Final Reliable Version)
REM =================================================================
REM  - This script uses the standard, most reliable PyInstaller workflow.
REM  - It creates a '_dist' folder for the final .exe.
REM  - It AUTOMATICALLY cleans up all temporary build files.
REM =================================================================

REM --- Define the final output directory ---
set DIST_DIR=_dist

echo [1/4] Preparing the build environment...

REM --- Clean up all possible old build artifacts ---
if exist "build" ( rmdir /s /q "build" )
if exist "dist" ( rmdir /s /q "dist" )
if exist "%DIST_DIR%" ( rmdir /s /q "%DIST_DIR%" )
if exist "*.spec" ( del /q "*.spec" )
echo      ... Old build artifacts removed.

REM --- Check if the virtual environment exists ---
if not exist ".venv\Scripts\activate.bat" (
    echo.
    echo ERROR: Virtual environment not found at '.\venv\Scripts\activate.bat'.
    echo Please create it first using: python -m venv .venv
    goto :error
)

echo.
echo [2/4] Activating virtual environment...
call .venv\Scripts\activate.bat

echo.
echo [3/4] Running PyInstaller to create the executable...
REM --- This is the core command to build the application ---
REM --- We let PyInstaller create 'build' and '.spec' in the root,
REM --- where it expects them, which solves all pathing issues.
pyinstaller main.py ^
    --name "PDF Letterhead Merger" ^
    --onefile ^
    --windowed ^
    --clean ^
    --icon="icon.ico" ^
    --add-data "icon.ico;." ^
    --distpath "%DIST_DIR%"

REM --- Check if PyInstaller was successful ---
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: PyInstaller failed to build the application.
    echo Please review the errors above.
    goto :error
)

echo.
echo [4/4] Cleaning up temporary build files...

REM --- This is the key to a clean directory: remove the temp files AFTER success ---
if exist "build" ( rmdir /s /q "build" )
if exist "*.spec" ( del /q "*.spec" )
echo      ... Temporary files removed.

REM --- Deactivate the virtual environment ---
call .venv\Scripts\deactivate.bat

echo.
echo =================================================================
echo  SUCCESS!
echo  Your standalone application is located in the output directory:
echo  .\%DIST_DIR%\PDF Letterhead Merger.exe
echo =================================================================
goto :end

:error
echo.
echo === BUILD FAILED ===
REM Also clean up failed build attempts
if exist "build" ( rmdir /s /q "build" )
if exist "*.spec" ( del /q "*.spec" )
pause
exit /b 1

:end
pause