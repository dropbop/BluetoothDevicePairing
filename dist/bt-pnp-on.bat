@echo off
REM Connect Bluetooth audio device (enable all matching entries)
REM Edit "OpenRun" below to match your device name

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0bt-pnp.ps1" connect "OpenRun"

echo.
pause
