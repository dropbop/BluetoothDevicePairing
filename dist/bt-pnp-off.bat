@echo off
REM Disconnect Bluetooth audio device (disable all matching entries)
REM Edit "OpenRun" below to match your device name

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0bt-pnp.ps1" disconnect "OpenRun"

echo.
pause
