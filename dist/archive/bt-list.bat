@echo off
REM List all Bluetooth audio devices

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0bt-headphones.ps1" list

echo.
pause
