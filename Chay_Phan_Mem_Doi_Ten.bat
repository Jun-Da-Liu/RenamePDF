@echo off
cd /d "%~dp0"
if exist "%~dp0PhanMemDoiTenPDF.exe" (
    start "" "%~dp0PhanMemDoiTenPDF.exe"
    exit /b 0
)
echo Starting PDF Rename GUI...
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -X utf8 app_gui.py
) else (
    python -X utf8 app_gui.py
)
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to start application.
    pause
)
