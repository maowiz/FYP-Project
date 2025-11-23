REM test_run.bat
REM Quick launcher for testing without face authentication
REM This will kill any existing process on port 8765 first

@echo off
echo ============================================
echo    VOICE ASSISTANT - TEST MODE
echo ============================================
echo Face authentication: DISABLED
echo System will start immediately
echo ============================================
echo.

echo Checking for existing instances...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8765') do (
    echo Killing process %%a on port 8765...
    taskkill /F /PID %%a >nul 2>&1
)

echo Starting in TEST MODE...
set TEST_MODE=1
python main.py

pause
