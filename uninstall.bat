@echo off
title Uninstalling Excel Highlighter Add-in...
echo ===================================================
echo   Uninstalling Excel Highlighter Add-in...
echo ===================================================
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"

echo.
echo Press any key to exit...
pause > nul
