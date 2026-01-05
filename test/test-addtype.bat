@echo off
REM Test script to verify Add-Type works with -ExecutionPolicy Bypass
REM Double-click this file or run from command prompt

echo.
echo Running Add-Type compatibility test...
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0test-addtype.ps1"

echo.
echo Press any key to close...
pause >nul
