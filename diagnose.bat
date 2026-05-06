@echo off
title Diagnostics
echo ============================================
echo  Job Tracker - Diagnostic Check
echo ============================================
echo.

echo [CHECK 1] Python version:
python --version
if errorlevel 1 (
    echo   FAIL: Python not found or not in PATH
    echo   Fix: Reinstall Python from python.org
    echo        During install tick "Add Python to PATH"
    goto end
)
echo   OK

echo.
echo [CHECK 2] pip version:
pip --version
if errorlevel 1 (
    echo   FAIL: pip not found
    goto end
)
echo   OK

echo.
echo [CHECK 3] Installing openpyxl...
pip install openpyxl
echo   Exit code: %errorlevel%

echo.
echo [CHECK 4] Installing pyinstaller...
pip install pyinstaller
echo   Exit code: %errorlevel%

echo.
echo [CHECK 5] Checking app.py exists:
if exist app.py (echo   OK - app.py found) else (echo   FAIL - app.py not found in this folder)

echo.
echo [CHECK 6] Running PyInstaller with full output:
pyinstaller --onefile --windowed --name "JobApplicationTracker" app.py
echo.
echo PyInstaller exit code: %errorlevel%

echo.
if exist "dist\JobApplicationTracker.exe" (
    echo SUCCESS - EXE was created at dist\JobApplicationTracker.exe
    if exist "Job-application-tracksheet.xlsx" (
        copy /Y "Job-application-tracksheet.xlsx" "dist\Job-application-tracksheet.xlsx" >nul
        echo Excel file copied to dist\ folder
    )
) else (
    echo FAIL - EXE was NOT created. Check errors above.
)

:end
echo.
echo ============================================
echo  Done. Read the output above for any errors.
echo ============================================
echo.
pause
