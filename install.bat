@echo off
title Installing Excel Highlighter Add-in...
echo ===================================================
echo   Installing Excel Highlighter Add-in...
echo ===================================================
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"

echo.
echo Press any key to exit...
pause > nul
