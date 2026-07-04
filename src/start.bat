@echo off
echo Starting Namvaran Excel Tracker...
echo.

:: Find and show local IP
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a
    goto :found
)
:found
set IP=%IP: =%
echo Server running at: http://%IP%:5000
echo.
echo Press Ctrl+C to stop.
echo.

python app.py
if errorlevel 1 (
    echo.
    echo ERROR: Could not start. Make sure Python is installed.
)
pause
