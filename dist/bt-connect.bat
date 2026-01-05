@echo off
REM Attempt to connect Bluetooth device via Settings UI automation
REM Edit "OpenRun" to match your device name

echo.
echo Attempting to connect Bluetooth device...
echo DO NOT touch keyboard/mouse until complete.
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0bt-settings.ps1" connect "OpenRun"

pause
