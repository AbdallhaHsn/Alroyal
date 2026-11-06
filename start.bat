@echo off
title Royal System Launcher
echo ===================================================
echo Starting RoyalDb Web System (Backend & Frontend)...
echo ===================================================

start "Royal API Server" cmd /k "cd /d %~dp0server && node index.js"
start "Royal Web App" cmd /k "cd /d %~dp0client && npm run dev"

echo.
echo Both servers have been launched!
echo Access the application at: http://localhost:3000
echo.
pause

