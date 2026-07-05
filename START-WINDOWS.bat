@echo off
:: Change to the folder where this .bat file lives
cd /d "%~dp0"

title AutoLead Showcaser
echo.
echo  ========================================
echo   AutoLead Showcaser - Starting Server
echo  ========================================
echo.

:: Check if Node.js is installed
node --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo  ERROR: Node.js is not installed!
    echo  Please download from: https://nodejs.org
    echo.
    pause
    exit /b 1
)

:: Install dependencies if node_modules missing
IF NOT EXIST "node_modules" (
    echo  Installing dependencies, please wait...
    npm install
    echo.
)

:: Get LAN IP
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set RAW=%%a
    goto :found
)
:found
:: Trim leading space
set IP=%RAW: =%

echo.
echo  ========================================
echo   Server is RUNNING!
echo  ========================================
echo.
echo   Admin Panel : http://%IP%:3000/admin
echo   Agent Panel : http://%IP%:3000/agent
echo.
echo   Keep this window open while working.
echo   Press Ctrl+C to stop the server.
echo  ========================================
echo.

node server.js
pause
