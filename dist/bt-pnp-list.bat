@echo off
REM List Bluetooth devices (no Add-Type required)

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0bt-pnp.ps1" list

echo.
pause
