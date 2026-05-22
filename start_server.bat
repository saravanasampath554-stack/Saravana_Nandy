@echo off
echo ============================================
echo  Measurement Dashboard - Production Server
echo ============================================
echo.

cd /d "%~dp0backend"

:: Get the machine's IP address
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    set IP=%%a
)
set IP=%IP: =%

echo Starting server...
echo.
echo Share this link with your team:
echo   http://%IP%:5000
echo.
echo Press Ctrl+C to stop the server.
echo ============================================

py app.py --prod --port 5000

pause
