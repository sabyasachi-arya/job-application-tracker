@echo off
title Job Tracker - Build
echo.
echo  ============================================
echo   Job Application Tracker  -  Build to EXE
echo  ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not found.
    echo  Install from https://python.org  (tick "Add Python to PATH")
    pause & exit /b 1
)

echo  [1/3] Installing dependencies...
pip install openpyxl pyinstaller --quiet --upgrade

echo  [2/3] Building EXE (1-2 minutes)...
pyinstaller --onefile --windowed --name "JobApplicationTracker" app.py

if errorlevel 1 (echo  BUILD FAILED & pause & exit /b 1)

echo  [3/3] Copying Excel file to dist\...
if exist "Job-application-tracksheet.xlsx" (
    copy /Y "Job-application-tracksheet.xlsx" "dist\Job-application-tracksheet.xlsx" >nul
    echo  Excel file copied.
)

echo.
echo  ============================================
echo   DONE!
echo  ============================================
echo.
echo   dist\JobApplicationTracker.exe
echo   dist\Job-application-tracksheet.xlsx
echo.
echo  Keep both files in the same folder.
echo  App auto-connects on first launch. No import/export ever.
echo.
pause
